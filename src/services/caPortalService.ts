import { getDb } from "./database.js";

const CONFIG_COL = "caPortalConfig";
const CONFIG_ID = "castro-arroyave-caja";

export interface CaPortalConfig {
  allowedEmails: string[];
  companyId: string; // MongoDB ObjectId del siigoCompany de Castro Arroyave
  updatedAt?: string;
}

export async function getCaPortalConfig(): Promise<CaPortalConfig | null> {
  const doc = await getDb().collection<any>(CONFIG_COL).findOne({ _id: CONFIG_ID });
  if (!doc) return null;
  return {
    allowedEmails: Array.isArray(doc.allowedEmails) ? doc.allowedEmails : [],
    companyId: doc.companyId || "",
    updatedAt: doc.updatedAt,
  };
}

export async function saveCaPortalConfig(patch: Partial<CaPortalConfig>): Promise<void> {
  const set: Record<string, unknown> = { updatedAt: new Date().toISOString() };
  if (patch.allowedEmails !== undefined) set.allowedEmails = patch.allowedEmails.map((e) => e.toLowerCase().trim());
  if (patch.companyId !== undefined) set.companyId = patch.companyId.trim();
  await getDb()
    .collection<any>(CONFIG_COL)
    .updateOne({ _id: CONFIG_ID }, { $set: set }, { upsert: true });
}

export function isEmailAllowed(email: string, config: CaPortalConfig): boolean {
  return config.allowedEmails.includes(email.toLowerCase().trim());
}
