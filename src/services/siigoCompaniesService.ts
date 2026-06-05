import { ObjectId } from "mongodb";
import { getDb } from "./database.js";
import { encryptSecret, decryptSecret } from "./secretCrypto.js";
import { runWithSiigoCompany, authenticateWithSiigo, type SiigoContext } from "./siigoService.js";

const COMPANIES = "siigoCompanies";

const baseUrl = () => process.env.SIIGO_API_BASE_URL || "https://api.siigo.com";
const partnerId = () => process.env.SIIGO_PARTNER_ID || "SentiidoAI";

export interface SiigoCompanyPublic {
  id: string;
  name: string;
  username: string;
  /** NIT de la empresa (sin dígito de verificación), para validar tokens DIAN. */
  nit?: string;
  ownerUserId?: string;
  sharedWith?: string[];
}

/** Normaliza un NIT para comparación: descarta el dígito de verificación y deja solo dígitos. */
export function normalizeNit(value: unknown): string {
  return String(value ?? "").split("-")[0].replace(/\D/g, "");
}

function toPublic(doc: any): SiigoCompanyPublic {
  return {
    id: doc._id.toString(),
    name: doc.name,
    username: doc.username,
    nit: doc.nit || "",
    ownerUserId: doc.ownerUserId ? String(doc.ownerUserId) : undefined,
    sharedWith: Array.isArray(doc.sharedWith) ? doc.sharedWith.map(String) : [],
  };
}

/** _id de un usuario admin para asignar como dueño por defecto (seed, fallback). */
async function resolveMainAdminId(): Promise<string | null> {
  const db = getDb();
  const email = (process.env.ADMIN_EMAIL || "").trim().toLowerCase();
  let admin = email ? await db.collection<any>("users").findOne({ email }) : null;
  if (!admin) admin = await db.collection<any>("users").findOne({ is_admin: true }, { sort: { created_at: 1 } });
  return admin ? admin._id.toString() : null;
}

/** Todas las empresas (uso interno / API key). Para usuarios usar listCompaniesForUser. */
export async function listCompanies(): Promise<SiigoCompanyPublic[]> {
  const docs = await getDb().collection<any>(COMPANIES).find({}).sort({ name: 1 }).toArray();
  return docs.map(toPublic);
}

/** Empresas visibles para un usuario: las que posee o las que le comparten. */
export async function listCompaniesForUser(userId: string): Promise<SiigoCompanyPublic[]> {
  const docs = await getDb()
    .collection<any>(COMPANIES)
    .find({ $or: [{ ownerUserId: userId }, { sharedWith: userId }] })
    .sort({ name: 1 })
    .toArray();
  return docs.map(toPublic);
}

/** True si el usuario es dueño de la empresa o la tiene compartida. */
export async function userCanAccessCompany(companyId: string, userId: string): Promise<boolean> {
  let oid: ObjectId;
  try {
    oid = new ObjectId(companyId);
  } catch {
    return false;
  }
  const doc = await getDb()
    .collection<any>(COMPANIES)
    .findOne({ _id: oid, $or: [{ ownerUserId: userId }, { sharedWith: userId }] }, { projection: { _id: 1 } });
  return !!doc;
}

/** True si el usuario es el dueño de la empresa. */
export async function userOwnsCompany(companyId: string, userId: string): Promise<boolean> {
  let oid: ObjectId;
  try {
    oid = new ObjectId(companyId);
  } catch {
    return false;
  }
  const doc = await getDb().collection<any>(COMPANIES).findOne({ _id: oid, ownerUserId: userId }, { projection: { _id: 1 } });
  return !!doc;
}

/** Credenciales completas (descifradas) de una empresa, listas para el contexto Siigo. */
export async function getCompanyContext(id: string): Promise<SiigoContext | null> {
  let oid: ObjectId;
  try {
    oid = new ObjectId(id);
  } catch {
    return null;
  }
  const doc = await getDb().collection<any>(COMPANIES).findOne({ _id: oid });
  if (!doc) return null;
  return {
    companyId: doc._id.toString(),
    baseUrl: baseUrl(),
    partnerId: partnerId(),
    username: doc.username,
    accessKey: decryptSecret(doc.accessKeyEnc),
    nit: doc.nit || "",
  };
}

/**
 * Define/actualiza el NIT de una empresa. Se usa como fuente de verdad para
 * validar que un token DIAN pertenezca a esta empresa. Guarda el NIT normalizado
 * (solo dígitos, sin DV).
 */
export async function setCompanyNit(companyId: string, nit: string): Promise<string> {
  let oid: ObjectId;
  try {
    oid = new ObjectId(companyId);
  } catch {
    throw new Error("Empresa inválida.");
  }
  const clean = normalizeNit(nit);
  if (!clean) throw new Error("El NIT debe contener dígitos.");
  const res = await getDb().collection<any>(COMPANIES).updateOne({ _id: oid }, { $set: { nit: clean } });
  if (res.matchedCount === 0) throw new Error("Empresa no encontrada.");
  return clean;
}

/**
 * Crea una empresa, validando las credenciales contra Siigo antes de guardar.
 * `ownerUserId` queda como dueño; si no se pasa, cae al admin principal.
 */
export async function createCompany(
  name: string,
  username: string,
  accessKey: string,
  ownerUserId?: string,
  nit?: string
): Promise<SiigoCompanyPublic> {
  const cleanName = String(name || "").trim();
  const cleanUser = String(username || "").trim();
  const cleanKey = String(accessKey || "").trim();
  const cleanNit = normalizeNit(nit);
  if (!cleanName || !cleanUser || !cleanKey) {
    throw new Error("Nombre, usuario y access key son obligatorios.");
  }

  const db = getDb();
  const dup = await db.collection<any>(COMPANIES).findOne({ username: cleanUser });
  if (dup) throw new Error(`Ya existe una empresa con el usuario ${cleanUser}.`);

  // Validar credenciales autenticando contra Siigo
  const probe: SiigoContext = {
    companyId: "__probe__",
    baseUrl: baseUrl(),
    partnerId: partnerId(),
    username: cleanUser,
    accessKey: cleanKey,
  };
  try {
    await runWithSiigoCompany(probe, () => authenticateWithSiigo());
  } catch (e) {
    throw new Error(
      `Las credenciales no autenticaron con Siigo: ${e instanceof Error ? e.message : "error"}`
    );
  }

  const owner = ownerUserId || (await resolveMainAdminId()) || "";
  const res = await db.collection<any>(COMPANIES).insertOne({
    name: cleanName,
    username: cleanUser,
    accessKeyEnc: encryptSecret(cleanKey),
    nit: cleanNit,
    ownerUserId: owner,
    sharedWith: [],
    createdAt: new Date().toISOString(),
  });
  return { id: res.insertedId.toString(), name: cleanName, username: cleanUser, nit: cleanNit, ownerUserId: owner, sharedWith: [] };
}

export async function deleteCompany(id: string): Promise<boolean> {
  let oid: ObjectId;
  try {
    oid = new ObjectId(id);
  } catch {
    return false;
  }
  const res = await getDb().collection<any>(COMPANIES).deleteOne({ _id: oid });
  return res.deletedCount > 0;
}

/** Comparte una empresa con otro usuario (lo agrega a sharedWith). */
export async function shareCompany(companyId: string, targetUserId: string): Promise<boolean> {
  let oid: ObjectId;
  try {
    oid = new ObjectId(companyId);
  } catch {
    return false;
  }
  const clean = String(targetUserId || "").trim();
  if (!clean) return false;
  const res = await getDb().collection<any>(COMPANIES).updateOne({ _id: oid }, { $addToSet: { sharedWith: clean } });
  return res.matchedCount > 0;
}

/** Deja de compartir una empresa con un usuario. */
export async function unshareCompany(companyId: string, targetUserId: string): Promise<boolean> {
  let oid: ObjectId;
  try {
    oid = new ObjectId(companyId);
  } catch {
    return false;
  }
  const res = await getDb().collection<any>(COMPANIES).updateOne({ _id: oid }, { $pull: { sharedWith: String(targetUserId) } } as any);
  return res.matchedCount > 0;
}

/** Transfiere la propiedad de una empresa a otro usuario. */
export async function transferOwner(companyId: string, newOwnerUserId: string): Promise<boolean> {
  let oid: ObjectId;
  try {
    oid = new ObjectId(companyId);
  } catch {
    return false;
  }
  const clean = String(newOwnerUserId || "").trim();
  if (!clean) return false;
  const res = await getDb()
    .collection<any>(COMPANIES)
    .updateOne({ _id: oid }, { $set: { ownerUserId: clean }, $pull: { sharedWith: clean } } as any);
  return res.matchedCount > 0;
}

/** Siembra la empresa del .env como primera empresa si todavía no existe. */
export async function seedSiigoCompanyFromEnv(): Promise<void> {
  const username = String(process.env.SIIGO_USERNAME || "").trim();
  const accessKey = String(process.env.SIIGO_ACCESS_KEY || "").trim();
  if (!username || !accessKey) return;
  const db = getDb();
  const existing = await db.collection<any>(COMPANIES).findOne({ username });
  if (existing) {
    // Empresa sembrada antes de tener dueño: la dejamos consistente con el modelo.
    if (!existing.ownerUserId) {
      const owner = await resolveMainAdminId();
      if (owner) {
        await db.collection<any>(COMPANIES).updateOne(
          { _id: existing._id },
          { $set: { ownerUserId: owner }, ...(Array.isArray(existing.sharedWith) ? {} : { $setOnInsert: {} }) }
        );
        if (!Array.isArray(existing.sharedWith)) {
          await db.collection<any>(COMPANIES).updateOne({ _id: existing._id }, { $set: { sharedWith: [] } });
        }
      }
    }
    return;
  }
  const owner = (await resolveMainAdminId()) || "";
  await db.collection<any>(COMPANIES).insertOne({
    name: process.env.SIIGO_SEED_COMPANY_NAME || "Fundación Sentiido",
    username,
    accessKeyEnc: encryptSecret(accessKey),
    ownerUserId: owner,
    sharedWith: [],
    createdAt: new Date().toISOString(),
    seeded: true,
  });
  console.log(`[Siigo] Empresa sembrada desde .env: ${username}`);
}
