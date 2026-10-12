/**
 * Terceros/proveedores del ERP.
 *
 * El ERP mantiene su propio catálogo (para operar aunque Siigo esté caído) y lo
 * mapea contra los terceros de Siigo. Si un tercero no existe en Siigo, el error
 * NO se oculta: se expone como acción "Resolver" en la UI.
 */
import { COL, col, fromDoc, fromDocs, now, scoped, toId, upsertDoc } from "./db.js";
import type { ErpDocument, ErpSupplier } from "./types.js";
import { ErpHttpError } from "./context.js";

export function normalizeNit(value: unknown): string {
  return String(value ?? "").replace(/\D/g, "");
}

/** Dígito de verificación DIAN. */
export function computeCheckDigit(nit: string): string {
  const weights = [3, 7, 13, 17, 19, 23, 29, 37, 41, 43, 47, 53, 59, 67, 71];
  const digits = normalizeNit(nit).split("").reverse();
  let total = 0;
  digits.forEach((d, i) => {
    total += Number(d) * (weights[i] ?? 0);
  });
  const mod = total % 11;
  return String(mod > 1 ? 11 - mod : mod);
}

export async function listSuppliers(
  tenantId: string,
  opts: { search?: string; limit?: number; active?: boolean } = {}
): Promise<ErpSupplier[]> {
  const filter: Record<string, unknown> = {};
  if (opts.active != null) filter.active = opts.active;
  if (opts.search) {
    const rx = new RegExp(String(opts.search).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i");
    filter.$or = [{ name: rx }, { identification: rx }];
  }
  const docs = await col(COL.suppliers)
    .find(scoped(tenantId, filter))
    .sort({ name: 1 })
    .limit(Math.min(opts.limit ?? 500, 1000))
    .toArray();
  return fromDocs<ErpSupplier>(docs);
}

export async function getSupplierByNit(tenantId: string, nit: string): Promise<ErpSupplier | null> {
  const doc = await col(COL.suppliers).findOne(scoped(tenantId, { identification: normalizeNit(nit) }));
  return doc ? fromDoc<ErpSupplier>(doc) : null;
}

export async function getSupplier(tenantId: string, id: string): Promise<ErpSupplier | null> {
  try {
    const doc = await col(COL.suppliers).findOne(scoped(tenantId, { _id: toId(id) }));
    return doc ? fromDoc<ErpSupplier>(doc) : null;
  } catch {
    return null;
  }
}

export interface UpsertSupplierInput {
  identification: string;
  name: string;
  identificationType?: string;
  email?: string;
  phone?: string;
  siigoId?: string;
  siigoBranchOffice?: number;
  defaultAccountCode?: string;
  defaultCostCenterId?: number;
  defaultSubCostCenterId?: string;
  paymentTermsDays?: number;
  defaultPaymentTypeId?: number;
  active?: boolean;
}

export async function upsertSupplier(tenantId: string, input: UpsertSupplierInput): Promise<ErpSupplier> {
  const identification = normalizeNit(input.identification);
  if (!identification) throw new ErpHttpError("La identificación del tercero es obligatoria.", 400, "invalid_nit");

  const set: Record<string, unknown> = { updatedAt: now() };
  for (const key of [
    "name", "email", "phone", "siigoId", "siigoBranchOffice", "defaultAccountCode",
    "defaultCostCenterId", "defaultSubCostCenterId", "paymentTermsDays", "defaultPaymentTypeId",
    "identificationType", "active",
  ] as const) {
    if (input[key] !== undefined) set[key] = input[key];
  }

  await col(COL.suppliers).updateOne(
    scoped(tenantId, { identification }) as any,
    upsertDoc(set, {
      tenantId,
      identification,
      identificationType: identification.length > 9 ? "31" : "13",
      checkDigit: computeCheckDigit(identification),
      active: true,
      documentCount: 0,
      createdAt: now(),
    }) as any,
    { upsert: true }
  );
  const saved = await getSupplierByNit(tenantId, identification);
  return saved!;
}

/** Crea/actualiza el tercero a partir de un documento recibido. */
export async function upsertSupplierFromDocument(tenantId: string, doc: ErpDocument): Promise<void> {
  if (!doc.supplierNit) return;
  await col(COL.suppliers).updateOne(
    scoped(tenantId, { identification: doc.supplierNit }) as any,
    upsertDoc(
      { name: doc.supplierName, updatedAt: now() },
      {
        tenantId,
        identification: doc.supplierNit,
        identificationType: doc.supplierNit.length > 9 ? "31" : "13",
        checkDigit: computeCheckDigit(doc.supplierNit),
        siigoBranchOffice: doc.supplierBranchOffice ?? 0,
        active: true,
        createdAt: now(),
      },
      { $inc: { documentCount: 1 } }
    ) as any,
    { upsert: true }
  );
}

/** Guarda la última parametrización usada — base de "Aplicar configuración habitual". */
export async function rememberAccounting(tenantId: string, nit: string, accounting: Record<string, unknown>): Promise<void> {
  await col(COL.suppliers).updateOne(
    scoped(tenantId, { identification: normalizeNit(nit) }) as any,
    { $set: { lastAccounting: accounting, updatedAt: now() } }
  );
}

export async function updateSupplier(tenantId: string, id: string, patch: Partial<ErpSupplier>): Promise<ErpSupplier | null> {
  const set: Record<string, unknown> = { updatedAt: now() };
  for (const key of [
    "name", "email", "phone", "siigoId", "siigoBranchOffice", "defaultAccountCode",
    "defaultCostCenterId", "defaultSubCostCenterId", "paymentTermsDays", "defaultPaymentTypeId", "active",
  ] as const) {
    if (patch[key] !== undefined) set[key] = patch[key];
  }
  const res = await col(COL.suppliers).findOneAndUpdate(
    scoped(tenantId, { _id: toId(id) }) as any,
    { $set: set },
    { returnDocument: "after" }
  );
  return res ? fromDoc<ErpSupplier>(res) : null;
}
