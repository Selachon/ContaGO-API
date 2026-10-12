/**
 * Catálogos Siigo cacheados por tenant.
 *
 * Se sincronizan bajo demanda y quedan almacenados localmente para que el ERP
 * siga operando (parametrizar, filtrar, mostrar nombres) aunque Siigo esté
 * temporalmente caído. Subcentros de costo: Siigo NO expone un recurso propio,
 * así que se mantienen únicamente en el ERP (documentado en INTEGRATIONS.md).
 */
import { COL, col, fromDoc, fromDocs, now, scoped, toId, upsertDoc } from "./db.js";
import {
  listCostCenters,
  listPaymentTypes,
  listPurchaseDocumentTypes,
  listPurchaseSupportDocumentTypes,
  listPaymentReceiptDocumentTypes,
  listTaxes,
} from "../services/siigoService.js";
import { getAccountsCatalog } from "../services/siigoSuggestionsService.js";
import type { ErpCostCenter } from "./types.js";
import { randomUUID } from "node:crypto";

export type CatalogKind =
  | "costCenters"
  | "taxes"
  | "paymentTypesFC"
  | "paymentTypesRC"
  | "purchaseDocumentTypes"
  | "supportDocumentTypes"
  | "paymentReceiptDocumentTypes"
  | "accounts";

export interface CatalogEntry {
  externalId: string | number;
  code?: string;
  name: string;
  active?: boolean;
  extra?: Record<string, unknown>;
}

export interface CatalogRecord {
  id: string;
  tenantId: string;
  kind: CatalogKind;
  entries: CatalogEntry[];
  lastSyncedAt: string;
  error?: string;
}

function asArray(value: unknown): any[] {
  if (Array.isArray(value)) return value;
  if (value && typeof value === "object") {
    const results = (value as any).results;
    if (Array.isArray(results)) return results;
  }
  return [];
}

function mapCostCenters(raw: unknown): CatalogEntry[] {
  return asArray(raw).map((c) => ({
    externalId: c.id,
    code: String(c.code ?? c.id ?? ""),
    name: String(c.name ?? c.code ?? c.id ?? ""),
    active: c.active !== false,
  }));
}

function mapTaxes(raw: unknown): CatalogEntry[] {
  return asArray(raw).map((t) => ({
    externalId: t.id,
    code: String(t.id ?? ""),
    name: String(t.name ?? ""),
    active: t.active !== false,
    extra: { type: t.type, percentage: t.percentage },
  }));
}

function mapDocumentTypes(raw: unknown): CatalogEntry[] {
  return asArray(raw).map((d) => ({
    externalId: d.id,
    code: String(d.code ?? ""),
    name: String(d.name ?? d.description ?? d.code ?? ""),
    active: d.active !== false,
    extra: { consecutive: d.consecutive, type: d.type, electronic_type: d.electronic_type },
  }));
}

function mapPaymentTypes(raw: unknown): CatalogEntry[] {
  return asArray(raw).map((p) => ({
    externalId: p.id,
    code: String(p.id ?? ""),
    name: String(p.name ?? ""),
    active: p.active !== false,
    extra: { type: p.type, due_date: p.due_date },
  }));
}

/**
 * Plan de cuentas (PUC).
 *
 * LIMITACIÓN DOCUMENTADA: la API de Siigo no expone un recurso de plan de
 * cuentas — `/v1/accounts` responde "Resource not found" y `/v1/account-groups`
 * devuelve grupos de productos/servicios, no el PUC. Se usa el plan cargado
 * localmente en ContaGO (`siigoSuggestionsService`), que es la misma fuente que
 * ya alimenta la herramienta de contabilización en producción.
 */
async function fetchAccounts(): Promise<CatalogEntry[]> {
  const accounts = await getAccountsCatalog();
  return accounts.map((a) => ({
    externalId: a.code,
    code: a.code,
    name: a.name,
    active: true,
    extra: { relatedTo: a.relatedTo, maturity: a.maturity },
  }));
}

const FETCHERS: Record<CatalogKind, () => Promise<CatalogEntry[]>> = {
  costCenters: async () => mapCostCenters(await listCostCenters()),
  taxes: async () => mapTaxes(await listTaxes()),
  paymentTypesFC: async () => mapPaymentTypes(await listPaymentTypes(undefined, "FC")),
  paymentTypesRC: async () => mapPaymentTypes(await listPaymentTypes(undefined, "RC")),
  purchaseDocumentTypes: async () => mapDocumentTypes(await listPurchaseDocumentTypes()),
  supportDocumentTypes: async () => mapDocumentTypes(await listPurchaseSupportDocumentTypes()),
  paymentReceiptDocumentTypes: async () => mapDocumentTypes(await listPaymentReceiptDocumentTypes()),
  accounts: fetchAccounts,
};

export const CATALOG_KINDS = Object.keys(FETCHERS) as CatalogKind[];

export async function getCatalog(tenantId: string, kind: CatalogKind): Promise<CatalogRecord | null> {
  const doc = await col(COL.catalogs).findOne(scoped(tenantId, { kind }));
  return doc ? fromDoc<CatalogRecord>(doc) : null;
}

export async function getAllCatalogs(tenantId: string): Promise<Record<string, CatalogRecord>> {
  const docs = await col(COL.catalogs).find(scoped(tenantId)).toArray();
  const out: Record<string, CatalogRecord> = {};
  for (const d of fromDocs<CatalogRecord>(docs)) out[d.kind] = d;
  return out;
}

/** Sincroniza un catálogo. El contexto Siigo del tenant ya debe estar activo. */
export async function syncCatalog(tenantId: string, kind: CatalogKind): Promise<CatalogRecord> {
  let entries: CatalogEntry[] = [];
  let error: string | undefined;
  try {
    entries = await FETCHERS[kind]();
  } catch (err) {
    error = err instanceof Error ? err.message : String(err);
  }

  const set: Record<string, unknown> = { lastSyncedAt: now(), error };
  if (!error) set.entries = entries;

  await col(COL.catalogs).updateOne(
    scoped(tenantId, { kind }) as any,
    // Si la sincronización falló no se pisan las entradas previas: el catálogo
    // cacheado sigue sirviendo y solo se registra el error.
    upsertDoc(set, { tenantId, kind, entries: [] }) as any,
    { upsert: true }
  );
  const saved = await getCatalog(tenantId, kind);
  return saved!;
}

export interface SyncSummary {
  kind: CatalogKind;
  count: number;
  error?: string;
  lastSyncedAt: string;
}

export async function syncAllCatalogs(tenantId: string): Promise<SyncSummary[]> {
  const results: SyncSummary[] = [];
  for (const kind of CATALOG_KINDS) {
    const rec = await syncCatalog(tenantId, kind);
    results.push({ kind, count: rec.entries?.length || 0, error: rec.error, lastSyncedAt: rec.lastSyncedAt });
  }
  // Espeja los centros de costo al catálogo local (que además guarda subcentros).
  await mirrorCostCenters(tenantId);
  return results;
}

/** Copia los centros de costo de Siigo a la colección local sin perder subcentros. */
export async function mirrorCostCenters(tenantId: string): Promise<void> {
  const cat = await getCatalog(tenantId, "costCenters");
  if (!cat?.entries?.length) return;
  for (const entry of cat.entries) {
    await col(COL.costCenters).updateOne(
      scoped(tenantId, { siigoId: Number(entry.externalId) }) as any,
      upsertDoc(
        { code: entry.code || "", name: entry.name, active: entry.active !== false, lastSyncedAt: now() },
        { tenantId, siigoId: Number(entry.externalId), subCenters: [] }
      ) as any,
      { upsert: true }
    );
  }
}

// ─── Centros y subcentros de costo (dimensiones del ERP) ─────────────────

export async function listCostCentersLocal(tenantId: string): Promise<ErpCostCenter[]> {
  const docs = await col(COL.costCenters).find(scoped(tenantId)).sort({ name: 1 }).toArray();
  return fromDocs<ErpCostCenter>(docs);
}

export async function createCostCenter(
  tenantId: string,
  input: { code: string; name: string; siigoId?: number }
): Promise<ErpCostCenter> {
  const doc = {
    tenantId,
    siigoId: input.siigoId,
    code: String(input.code || "").trim(),
    name: String(input.name || "").trim(),
    active: true,
    subCenters: [] as ErpCostCenter["subCenters"],
  };
  const res = await col(COL.costCenters).insertOne(doc as any);
  return { id: String(res.insertedId), ...doc };
}

export async function updateCostCenter(
  tenantId: string,
  id: string,
  patch: Partial<ErpCostCenter>
): Promise<ErpCostCenter | null> {
  const set: Record<string, unknown> = {};
  for (const key of ["code", "name", "active", "siigoId"] as const) {
    if (patch[key] !== undefined) set[key] = patch[key];
  }
  const res = await col(COL.costCenters).findOneAndUpdate(
    scoped(tenantId, { _id: toId(id) }) as any,
    { $set: set },
    { returnDocument: "after" }
  );
  return res ? fromDoc<ErpCostCenter>(res) : null;
}

export async function addSubCostCenter(
  tenantId: string,
  costCenterId: string,
  input: { code: string; name: string }
): Promise<ErpCostCenter | null> {
  const sub = { id: randomUUID(), code: String(input.code || "").trim(), name: String(input.name || "").trim(), active: true };
  const res = await col(COL.costCenters).findOneAndUpdate(
    scoped(tenantId, { _id: toId(costCenterId) }) as any,
    { $push: { subCenters: sub } as any },
    { returnDocument: "after" }
  );
  return res ? fromDoc<ErpCostCenter>(res) : null;
}

export async function removeSubCostCenter(
  tenantId: string,
  costCenterId: string,
  subId: string
): Promise<ErpCostCenter | null> {
  const res = await col(COL.costCenters).findOneAndUpdate(
    scoped(tenantId, { _id: toId(costCenterId) }) as any,
    { $pull: { subCenters: { id: subId } } as any },
    { returnDocument: "after" }
  );
  return res ? fromDoc<ErpCostCenter>(res) : null;
}

export async function deleteCostCenter(tenantId: string, id: string): Promise<boolean> {
  const res = await col(COL.costCenters).deleteOne(scoped(tenantId, { _id: toId(id) }) as any);
  return res.deletedCount > 0;
}

/** Resuelve el nombre legible de un centro de costo (para mostrar en la UI). */
export async function costCenterName(tenantId: string, siigoId?: number): Promise<string | undefined> {
  if (siigoId == null) return undefined;
  const doc = await col(COL.costCenters).findOne(scoped(tenantId, { siigoId: Number(siigoId) }));
  return doc?.name;
}
