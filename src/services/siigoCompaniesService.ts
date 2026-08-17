import { ObjectId } from "mongodb";
import { getDb } from "./database.js";
import { encryptSecret, decryptSecret } from "./secretCrypto.js";
import { runWithSiigoCompany, authenticateWithSiigo, type SiigoContext } from "./siigoService.js";

const COMPANIES = "siigoCompanies";

const baseUrl = () => process.env.SIIGO_API_BASE_URL || "https://api.siigo.com";
const partnerId = () => process.env.SIIGO_PARTNER_ID || "SentiidoAI";

export interface CompanySettings {
  /** Si true, la causación usa productos Siigo en lugar de cuentas contables PUC. */
  useProducts?: boolean;
  /** ID de bodega Siigo a usar cuando useProducts=true. */
  warehouseId?: string | null;
  /** Código de producto Siigo predeterminado para todos los ítems. */
  defaultProductCode?: string | null;
  /** ID de forma de pago Siigo predeterminada para todas las facturas. */
  defaultPaymentTypeId?: string | null;
  /** Si true, el centro de costo no es obligatorio ni se muestra. */
  skipCostCenter?: boolean;
  /** Si true, los ítems nunca se condensan por impuesto; siempre van separados. */
  noCondenseItems?: boolean;
  /** Mapa de nombre de impuesto extra → código de cuenta PUC para causación (ej. { ICL: "239530", IBUA: "239590" }). */
  extraTaxAccounts?: Record<string, string>;
}

export interface PaymentRecord {
  paidAt: string;
  period: string; // "YYYY-MM" — mes al que corresponde el pago
  amount?: number;
  method?: string;
  invoiceRef?: string;
}

export interface CompanySubscription {
  licenseStartDate?: string;
  licenseEndDate?: string;
  paymentAmount?: number;
  paymentMethod?: string;
  invoiceRef?: string;
  paymentStatus?: "pending" | "paid";
  paidAt?: string;
  paymentHistory?: PaymentRecord[];
}

export interface SiigoCompanyPublic {
  id: string;
  name: string;
  username: string;
  nit?: string;
  ownerUserId?: string;
  sharedWith?: string[];
  settings?: CompanySettings;
  toolIds: string[];
  /** Suscripción legada a nivel de empresa (pre-multi-módulo). */
  subscription?: CompanySubscription;
  /** Suscripción por módulo: subscriptions[toolId] = datos de ese módulo. */
  subscriptions: Record<string, CompanySubscription>;
}

export interface SubscriptionRow {
  companyId: string;
  companyName: string;
  toolId: string;
  userId: string;
  userName: string;
  userEmail: string;
  licenseStartDate?: string;
  licenseEndDate?: string;
  paymentAmount?: number;
  paymentMethod?: string;
  invoiceRef?: string;
  paymentStatus?: "pending" | "paid";
  paidAt?: string;
  paymentHistory?: PaymentRecord[];
}

/** Normaliza un NIT para comparación: descarta el dígito de verificación y deja solo dígitos. */
export function normalizeNit(value: unknown): string {
  return String(value ?? "").split("-")[0].replace(/\D/g, "");
}

function resolveToolIds(doc: any): string[] {
  if (Array.isArray(doc.toolIds)) return doc.toolIds;
  // compatibilidad con documentos viejos que tenían toolId singular
  if (doc.toolId) return [doc.toolId];
  return [];
}

function toPublic(doc: any): SiigoCompanyPublic {
  // Suscripción legada a nivel empresa
  const sub: CompanySubscription = {};
  if (doc.licenseStartDate) sub.licenseStartDate = doc.licenseStartDate;
  if (doc.licenseEndDate) sub.licenseEndDate = doc.licenseEndDate;
  if (doc.paymentAmount != null) sub.paymentAmount = doc.paymentAmount;
  if (doc.paymentMethod) sub.paymentMethod = doc.paymentMethod;
  if (doc.invoiceRef) sub.invoiceRef = doc.invoiceRef;
  return {
    id: doc._id.toString(),
    name: doc.name,
    username: doc.username,
    nit: doc.nit || "",
    ownerUserId: doc.ownerUserId ? String(doc.ownerUserId) : undefined,
    sharedWith: Array.isArray(doc.sharedWith) ? doc.sharedWith.map(String) : [],
    settings: doc.settings || {},
    toolIds: resolveToolIds(doc),
    subscription: Object.keys(sub).length ? sub : undefined,
    subscriptions: typeof doc.subscriptions === "object" && doc.subscriptions !== null ? doc.subscriptions : {},
  };
}

export async function updateToolSubscription(
  companyId: string,
  toolId: string,
  data: CompanySubscription & { period?: string }
): Promise<SiigoCompanyPublic> {
  let oid: ObjectId;
  try { oid = new ObjectId(companyId); } catch { throw new Error("Empresa inválida."); }
  const patch: Record<string, unknown> = {};
  if (data.licenseStartDate !== undefined) patch[`subscriptions.${toolId}.licenseStartDate`] = data.licenseStartDate || null;
  if (data.licenseEndDate !== undefined) patch[`subscriptions.${toolId}.licenseEndDate`] = data.licenseEndDate || null;
  if (data.paymentAmount !== undefined) patch[`subscriptions.${toolId}.paymentAmount`] = data.paymentAmount ?? null;
  if (data.paymentMethod !== undefined) patch[`subscriptions.${toolId}.paymentMethod`] = data.paymentMethod || null;
  if (data.invoiceRef !== undefined) patch[`subscriptions.${toolId}.invoiceRef`] = data.invoiceRef || null;
  if (data.paymentStatus !== undefined) patch[`subscriptions.${toolId}.paymentStatus`] = data.paymentStatus;
  if (data.paidAt !== undefined) patch[`subscriptions.${toolId}.paidAt`] = data.paidAt || null;
  if (Object.keys(patch).length === 0) throw new Error("Sin campos a actualizar.");

  const db = getDb().collection<any>(COMPANIES);

  // Si hay pago nuevo, agregar al historial (evitando duplicados por periodo)
  if (data.paidAt && data.period) {
    const record: PaymentRecord = {
      paidAt: data.paidAt,
      period: data.period,
      ...(data.paymentAmount != null && { amount: data.paymentAmount }),
      ...(data.paymentMethod && { method: data.paymentMethod }),
      ...(data.invoiceRef && { invoiceRef: data.invoiceRef }),
    };
    // Quitar entrada previa del mismo periodo antes de insertar
    await db.updateOne({ _id: oid }, {
      $pull: { [`subscriptions.${toolId}.paymentHistory`]: { period: data.period } } as any
    });
    await db.updateOne({ _id: oid }, {
      $set: patch,
      $push: { [`subscriptions.${toolId}.paymentHistory`]: record } as any
    });
  } else {
    const res = await db.updateOne({ _id: oid }, { $set: patch });
    if (res.matchedCount === 0) throw new Error("Empresa no encontrada.");
  }

  const doc = await getDb().collection<any>(COMPANIES).findOne({ _id: oid });
  return toPublic(doc!);
}

export async function listAllSubscriptionRows(): Promise<SubscriptionRow[]> {
  const db = getDb();
  // Empresas que tienen toolIds
  const companies = await db.collection<any>(COMPANIES)
    .find({ $or: [{ toolIds: { $exists: true, $not: { $size: 0 } } }, { toolId: { $exists: true } }] })
    .toArray();

  // Reunir ownerUserIds únicos para hacer una sola consulta de usuarios
  const userIds = [...new Set(companies.map((c: any) => c.ownerUserId).filter(Boolean))];
  const users = userIds.length
    ? await db.collection<any>("users").find({ _id: { $in: userIds.map((id: string) => { try { return new ObjectId(id); } catch { return null; } }).filter(Boolean) } }).toArray()
    : [];
  const userMap = new Map(users.map((u: any) => [u._id.toString(), u]));

  const rows: SubscriptionRow[] = [];
  for (const company of companies) {
    const toolIds = resolveToolIds(company);
    const owner = userMap.get(String(company.ownerUserId || ""));
    const subs: Record<string, any> = company.subscriptions || {};
    // Fallback legado: si la empresa tiene licenseEndDate a nivel raíz, lo usamos para herramientas sin datos propios
    const legacySub: CompanySubscription = {};
    if (company.licenseStartDate) legacySub.licenseStartDate = company.licenseStartDate;
    if (company.licenseEndDate) legacySub.licenseEndDate = company.licenseEndDate;
    if (company.paymentAmount != null) legacySub.paymentAmount = company.paymentAmount;
    if (company.paymentMethod) legacySub.paymentMethod = company.paymentMethod;
    if (company.invoiceRef) legacySub.invoiceRef = company.invoiceRef;

    for (const toolId of toolIds) {
      const sub: CompanySubscription = subs[toolId] || (Object.keys(legacySub).length ? legacySub : {});
      rows.push({
        companyId: company._id.toString(),
        companyName: company.name,
        toolId,
        userId: String(company.ownerUserId || ""),
        userName: owner?.name || owner?.email || "—",
        userEmail: owner?.email || "—",
        ...sub,
      });
    }
  }
  return rows;
}

export async function removePaymentFromHistory(
  companyId: string,
  toolId: string,
  period: string
): Promise<SiigoCompanyPublic> {
  let oid: ObjectId;
  try { oid = new ObjectId(companyId); } catch { throw new Error("Empresa inválida."); }
  await getDb().collection<any>(COMPANIES).updateOne(
    { _id: oid },
    { $pull: { [`subscriptions.${toolId}.paymentHistory`]: { period } } } as any
  );
  const doc = await getDb().collection<any>(COMPANIES).findOne({ _id: oid });
  if (!doc) throw new Error("Empresa no encontrada.");
  return toPublic(doc);
}

export async function updateCompanySubscription(
  companyId: string,
  data: CompanySubscription & { toolIds?: string[] }
): Promise<SiigoCompanyPublic> {
  let oid: ObjectId;
  try { oid = new ObjectId(companyId); } catch { throw new Error("Empresa inválida."); }
  const patch: Record<string, unknown> = {};
  if (data.toolIds !== undefined) {
    patch.toolIds = data.toolIds;
    patch.toolId = null; // limpiar campo legado
  }
  if (data.licenseStartDate !== undefined) patch.licenseStartDate = data.licenseStartDate || null;
  if (data.licenseEndDate !== undefined) patch.licenseEndDate = data.licenseEndDate || null;
  if (data.paymentAmount !== undefined) patch.paymentAmount = data.paymentAmount ?? null;
  if (data.paymentMethod !== undefined) patch.paymentMethod = data.paymentMethod || null;
  if (data.invoiceRef !== undefined) patch.invoiceRef = data.invoiceRef || null;
  if (Object.keys(patch).length === 0) throw new Error("Sin campos a actualizar.");
  const res = await getDb().collection<any>(COMPANIES).updateOne({ _id: oid }, { $set: patch });
  if (res.matchedCount === 0) throw new Error("Empresa no encontrada.");
  const doc = await getDb().collection<any>(COMPANIES).findOne({ _id: oid });
  return toPublic(doc!);
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

/** Cuántas empresas es DUEÑO un usuario (para hacer cumplir el cupo del plan). */
export async function countCompaniesOwnedBy(userId: string): Promise<number> {
  return getDb().collection<any>(COMPANIES).countDocuments({ ownerUserId: userId });
}

export async function countCompaniesOwnedByForTool(userId: string, toolId: string): Promise<number> {
  // Cuenta empresas que tienen este toolId en toolIds (array nuevo) O en toolId (campo legado)
  return getDb().collection<any>(COMPANIES).countDocuments({
    ownerUserId: userId,
    $or: [{ toolIds: toolId }, { toolId }],
  });
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
export async function updateCompanySettings(companyId: string, settings: CompanySettings): Promise<CompanySettings> {
  let oid: ObjectId;
  try { oid = new ObjectId(companyId); } catch { throw new Error("Empresa inválida."); }
  const patch: Record<string, unknown> = {};
  if (settings.useProducts !== undefined) patch["settings.useProducts"] = Boolean(settings.useProducts);
  if (settings.warehouseId !== undefined) patch["settings.warehouseId"] = settings.warehouseId || null;
  if (settings.defaultProductCode !== undefined) patch["settings.defaultProductCode"] = settings.defaultProductCode || null;
  if (settings.defaultPaymentTypeId !== undefined) patch["settings.defaultPaymentTypeId"] = settings.defaultPaymentTypeId || null;
  if (settings.skipCostCenter !== undefined) patch["settings.skipCostCenter"] = Boolean(settings.skipCostCenter);
  if (settings.noCondenseItems !== undefined) patch["settings.noCondenseItems"] = Boolean(settings.noCondenseItems);
  if (settings.extraTaxAccounts !== undefined) patch["settings.extraTaxAccounts"] = settings.extraTaxAccounts || {};
  if (Object.keys(patch).length === 0) throw new Error("Sin campos a actualizar.");
  const res = await getDb().collection<any>(COMPANIES).updateOne({ _id: oid }, { $set: patch });
  if (res.matchedCount === 0) throw new Error("Empresa no encontrada.");
  const doc = await getDb().collection<any>(COMPANIES).findOne({ _id: oid }, { projection: { settings: 1 } });
  return doc?.settings || {};
}

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
  nit?: string,
  toolIds?: string[]
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
  const doc: Record<string, unknown> = {
    name: cleanName,
    username: cleanUser,
    accessKeyEnc: encryptSecret(cleanKey),
    nit: cleanNit,
    ownerUserId: owner,
    sharedWith: [],
    createdAt: new Date().toISOString(),
  };
  if (toolIds?.length) doc.toolIds = toolIds;
  const res = await db.collection<any>(COMPANIES).insertOne(doc);
  return { id: res.insertedId.toString(), name: cleanName, username: cleanUser, nit: cleanNit, ownerUserId: owner, sharedWith: [], toolIds: toolIds || [], subscriptions: {} };
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

/** Liga una conexión de Google Drive (de un usuario) a la empresa, para archivar soportes. */
export async function setCompanyDrive(companyId: string, ownerUserId: string, connectionId: string): Promise<boolean> {
  let oid: ObjectId;
  try { oid = new ObjectId(companyId); } catch { return false; }
  const res = await getDb().collection<any>(COMPANIES).updateOne(
    { _id: oid },
    { $set: { driveOwnerUserId: ownerUserId, driveConnectionId: connectionId } },
  );
  return res.matchedCount > 0;
}

/** Drive ligado a la empresa (usuario dueño + conexión) + su NIT. */
export async function getCompanyDrive(companyId: string): Promise<{ ownerUserId: string; connectionId: string; nit: string } | null> {
  let oid: ObjectId;
  try { oid = new ObjectId(companyId); } catch { return null; }
  const doc = await getDb().collection<any>(COMPANIES).findOne({ _id: oid }, { projection: { driveOwnerUserId: 1, driveConnectionId: 1, nit: 1 } });
  if (!doc?.driveOwnerUserId || !doc?.driveConnectionId) return null;
  return { ownerUserId: doc.driveOwnerUserId, connectionId: doc.driveConnectionId, nit: doc.nit || "" };
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
