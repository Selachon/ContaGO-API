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
}

function toPublic(doc: any): SiigoCompanyPublic {
  return { id: doc._id.toString(), name: doc.name, username: doc.username };
}

export async function listCompanies(): Promise<SiigoCompanyPublic[]> {
  const docs = await getDb().collection<any>(COMPANIES).find({}).sort({ name: 1 }).toArray();
  return docs.map(toPublic);
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
  };
}

/** Crea una empresa, validando las credenciales contra Siigo antes de guardar. */
export async function createCompany(
  name: string,
  username: string,
  accessKey: string
): Promise<SiigoCompanyPublic> {
  const cleanName = String(name || "").trim();
  const cleanUser = String(username || "").trim();
  const cleanKey = String(accessKey || "").trim();
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

  const res = await db.collection<any>(COMPANIES).insertOne({
    name: cleanName,
    username: cleanUser,
    accessKeyEnc: encryptSecret(cleanKey),
    createdAt: new Date().toISOString(),
  });
  return { id: res.insertedId.toString(), name: cleanName, username: cleanUser };
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

/** Siembra la empresa del .env como primera empresa si todavía no existe. */
export async function seedSiigoCompanyFromEnv(): Promise<void> {
  const username = String(process.env.SIIGO_USERNAME || "").trim();
  const accessKey = String(process.env.SIIGO_ACCESS_KEY || "").trim();
  if (!username || !accessKey) return;
  const db = getDb();
  const existing = await db.collection<any>(COMPANIES).findOne({ username });
  if (existing) return;
  await db.collection<any>(COMPANIES).insertOne({
    name: process.env.SIIGO_SEED_COMPANY_NAME || "Fundación Sentiido",
    username,
    accessKeyEnc: encryptSecret(accessKey),
    createdAt: new Date().toISOString(),
    seeded: true,
  });
  console.log(`[Siigo] Empresa sembrada desde .env: ${username}`);
}
