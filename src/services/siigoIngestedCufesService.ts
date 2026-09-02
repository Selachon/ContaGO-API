import { getDb } from "./database.js";

const COLLECTION = "siigoIngestedCufes";

export type DianInvoiceStatus = "pending" | "caused" | "ignored";

export interface DianInvoiceRecord {
  companyId: string;
  cufe: string;
  docnum: string;
  trackId?: string;
  status: DianInvoiceStatus;
  supplierNit?: string;
  supplierName?: string;
  issueDate?: string;
  total?: number;
  ingestedAt: string;   // primera vez que se vio (= fetchedAt inicial)
  fetchedAt?: string;   // última descarga
  causedAt?: string;
  siigoId?: string;
}

/**
 * Devuelve el conjunto de CUFEs ya conocidos para una empresa (cualquier status).
 * Se usa para no re-descargar del portal DIAN lo que ya está en la tabla.
 */
export async function getIngestedCufes(companyId: string): Promise<Set<string>> {
  if (!companyId) return new Set();
  const docs = await getDb()
    .collection<any>(COLLECTION)
    .find({ companyId }, { projection: { cufe: 1, _id: 0 } })
    .toArray();
  return new Set(docs.map((d) => d.cufe).filter(Boolean));
}

/** Conjunto de trackIds ya registrados (para saltar re-descarga por trackId). */
export async function getIngestedTrackIds(companyId: string): Promise<Set<string>> {
  if (!companyId) return new Set();
  const docs = await getDb()
    .collection<any>(COLLECTION)
    .find({ companyId }, { projection: { trackId: 1, _id: 0 } })
    .toArray();
  return new Set(docs.map((d) => (d.trackId || "").trim()).filter(Boolean));
}

/**
 * Upsert masivo de facturas descargadas de DIAN.
 * - Nuevas → status='pending' + todos los campos.
 * - Existentes → actualiza datos de factura pero NO toca status (si ya está caused/ignored se respeta).
 */
export async function upsertDianInvoices(
  companyId: string,
  invoices: Array<{
    cufe: string;
    docnum?: string;
    trackId?: string;
    supplierNit?: string;
    supplierName?: string;
    issueDate?: string;
    total?: number;
    fetchedAt?: string;
  }>
): Promise<void> {
  if (!companyId || invoices.length === 0) return;
  const now = new Date().toISOString();
  const ops = invoices
    .filter((r) => r.cufe)
    .map((r) => ({
      updateOne: {
        filter: { companyId, cufe: r.cufe },
        update: {
          $set: {
            docnum: r.docnum || "",
            ...(r.trackId !== undefined && { trackId: r.trackId }),
            supplierNit: r.supplierNit || "",
            supplierName: r.supplierName || "",
            issueDate: r.issueDate || "",
            total: r.total ?? 0,
            fetchedAt: r.fetchedAt || now,
          },
          $setOnInsert: { companyId, cufe: r.cufe, status: "pending", ingestedAt: now },
        },
        upsert: true,
      },
    }));
  if (ops.length === 0) return;
  await getDb().collection<any>(COLLECTION).bulkWrite(ops, { ordered: false });
}

/** Marca una factura como causada en Siigo. */
export async function markCausedInSiigo(
  companyId: string,
  cufe: string,
  siigoId?: string
): Promise<void> {
  if (!companyId || !cufe) return;
  await getDb()
    .collection<any>(COLLECTION)
    .updateOne(
      { companyId, cufe },
      { $set: { status: "caused", causedAt: new Date().toISOString(), ...(siigoId ? { siigoId } : {}) } }
    );
}

/** Marca una factura como ignorada (no volver a traer del portal). */
export async function markIgnoredInDian(companyId: string, cufe: string): Promise<void> {
  if (!companyId || !cufe) return;
  const now = new Date().toISOString();
  await getDb()
    .collection<any>(COLLECTION)
    .updateOne(
      { companyId, cufe },
      { $set: { status: "ignored" }, $setOnInsert: { companyId, cufe, docnum: "", ingestedAt: now } },
      { upsert: true }
    );
}

/** Vuelve a marcar una factura como pendiente (el contador quiere re-causarla). */
export async function markPendingInDian(companyId: string, cufe: string): Promise<void> {
  if (!companyId || !cufe) return;
  await getDb()
    .collection<any>(COLLECTION)
    .updateOne({ companyId, cufe }, { $set: { status: "pending" }, $unset: { causedAt: "", siigoId: "" } });
}

/**
 * Lista las facturas DIAN de una empresa con filtro opcional de status.
 * Retorna las más recientes primero.
 */
export async function listDianInvoices(
  companyId: string,
  status?: DianInvoiceStatus | "all"
): Promise<DianInvoiceRecord[]> {
  if (!companyId) return [];
  const filter: any = { companyId };
  if (status && status !== "all") filter.status = status;
  const docs = await getDb()
    .collection<any>(COLLECTION)
    .find(filter, { projection: { _id: 0 } })
    .sort({ fetchedAt: -1, ingestedAt: -1 })
    .limit(2000)
    .toArray();
  return docs.map((d) => ({
    ...d,
    status: d.status || "ignored", // registros legacy sin status = bloqueados
  }));
}

/**
 * Registro masivo para bloqueo explícito (compat con llamadores legacy).
 * Ahora equivale a marcar como 'ignored'.
 */
export async function recordIngestedCufes(
  companyId: string,
  records: { cufe: string; docnum?: string; trackId?: string }[]
): Promise<void> {
  if (!companyId || records.length === 0) return;
  const now = new Date().toISOString();
  const ops = records
    .filter((r) => r.cufe || r.trackId || r.docnum)
    .map((r) => {
      const filter = r.cufe
        ? { companyId, cufe: r.cufe }
        : r.trackId
        ? { companyId, trackId: r.trackId }
        : { companyId, docnum: r.docnum };
      return {
        updateOne: {
          filter,
          update: {
            $set: { cufe: r.cufe || "", docnum: r.docnum || "", trackId: r.trackId || "", status: "ignored" },
            $setOnInsert: { companyId, ingestedAt: now },
          },
          upsert: true,
        },
      };
    });
  if (ops.length === 0) return;
  await getDb().collection<any>(COLLECTION).bulkWrite(ops, { ordered: false });
}
