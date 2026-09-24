import { getDb } from "./database.js";
import { SIIGO_CUFE_PREFIX, nitKey, sameInvoiceNumber } from "./siigoCausadasPlan.js";

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
  /** Nombre del comprobante en Siigo (p.ej. "FC-1-123"). */
  siigoName?: string;
}

/**
 * Devuelve el conjunto de CUFEs ya conocidos para una empresa (cualquier status).
 * Se usa para no re-descargar del portal DIAN lo que ya está en la tabla.
 */
export async function getIngestedCufes(companyId: string): Promise<Set<string>> {
  if (!companyId) return new Set();
  // Solo bloquea re-descarga para facturas ya causadas o ignoradas explícitamente.
  // Las 'pending' se vuelven a descargar en cada sesión para que aparezcan en la tabla.
  const docs = await getDb()
    .collection<any>(COLLECTION)
    .find({ companyId, status: { $in: ["caused", "ignored"] } }, { projection: { cufe: 1, _id: 0 } })
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
  await adoptSyntheticCaused(companyId, invoices).catch((e) => console.warn("[Siigo Ingest] adoptSyntheticCaused falló:", e instanceof Error ? e.message : e));
}

/** Datos opcionales de la factura para completar/crear el registro al marcarla causada. */
export interface CausedMeta {
  docnum?: string;
  supplierNit?: string;
  supplierName?: string;
  issueDate?: string;
  total?: number;
  siigoName?: string;
}

/**
 * Marca una factura como causada en Siigo.
 *
 * Hace UPSERT: antes era un update simple y, si la factura no estaba en el registro
 * (subida como ZIP/XML, o registro que no se alcanzó a escribir), no pasaba nada y
 * la factura causada nunca aparecía en la pestaña "Causadas".
 */
export async function markCausedInSiigo(
  companyId: string,
  cufe: string,
  siigoId?: string,
  meta: CausedMeta = {}
): Promise<void> {
  if (!companyId || !cufe) return;
  const now = new Date().toISOString();
  const set: Record<string, unknown> = { status: "caused", causedAt: now };
  if (siigoId) set.siigoId = siigoId;
  if (meta.siigoName) set.siigoName = meta.siigoName;
  if (meta.docnum) set.docnum = meta.docnum;
  if (meta.supplierNit) set.supplierNit = meta.supplierNit;
  if (meta.supplierName) set.supplierName = meta.supplierName;
  if (meta.issueDate) set.issueDate = meta.issueDate;
  if (meta.total) set.total = meta.total;
  const onInsert: Record<string, unknown> = { companyId, cufe, ingestedAt: now, fetchedAt: now };
  if (!("docnum" in set)) onInsert.docnum = "";
  await getDb()
    .collection<any>(COLLECTION)
    .updateOne({ companyId, cufe }, { $set: set, $setOnInsert: onInsert }, { upsert: true });
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
 *
 * Las CAUSADAS se devuelven siempre completas (sin tope): antes todo el listado se
 * cortaba en 2000 por fetchedAt y, en empresas con muchas pendientes que se
 * re-descargan cada sesión, las causadas más antiguas se caían de la pestaña.
 */
export async function listDianInvoices(
  companyId: string,
  status?: DianInvoiceStatus | "all"
): Promise<DianInvoiceRecord[]> {
  if (!companyId) return [];
  const col = getDb().collection<any>(COLLECTION);
  const sort = { fetchedAt: -1, ingestedAt: -1 } as const;
  const noId = { projection: { _id: 0 } };
  let docs: any[];
  if (status && status !== "all") {
    const q = col.find({ companyId, status }, noId).sort(sort);
    docs = await (status === "caused" ? q : q.limit(2000)).toArray();
  } else {
    const [caused, rest] = await Promise.all([
      col.find({ companyId, status: "caused" }, noId).sort(sort).toArray(),
      col.find({ companyId, status: { $ne: "caused" } }, noId).sort(sort).limit(2000).toArray(),
    ]);
    docs = [...caused, ...rest].sort((a, b) => String(b.fetchedAt || b.ingestedAt || "").localeCompare(String(a.fetchedAt || a.ingestedAt || "")));
  }
  return docs.map((d) => ({
    ...d,
    status: d.status || "ignored", // registros legacy sin status = bloqueados
  }));
}

/** Llaves `nit|dígitos` de las facturas ya causadas (registro DIAN + compras de Siigo sincronizadas). */
export async function getCausedInvoiceIndex(companyId: string): Promise<Array<{ nit: string; docnum: string }>> {
  if (!companyId) return [];
  const docs = await getDb()
    .collection<any>(COLLECTION)
    .find({ companyId, status: "caused" }, { projection: { supplierNit: 1, docnum: 1, _id: 0 } })
    .toArray();
  return docs.map((d) => ({ nit: nitKey(d.supplierNit), docnum: String(d.docnum || "") })).filter((d) => d.nit && d.docnum);
}

/**
 * Cuando llega de la DIAN una factura que ya existe como causada por compra de
 * Siigo (registro sintético `siigo:<id>`), el registro real hereda el estado y el
 * sintético se elimina — así no queda duplicada ni vuelve a aparecer como pendiente.
 */
export async function adoptSyntheticCaused(
  companyId: string,
  invoices: Array<{ cufe: string; docnum?: string; supplierNit?: string }>
): Promise<number> {
  if (!companyId || invoices.length === 0) return 0;
  const col = getDb().collection<any>(COLLECTION);
  const synth = await col
    .find({ companyId, status: "caused", cufe: { $regex: `^${SIIGO_CUFE_PREFIX}` } })
    .toArray();
  if (synth.length === 0) return 0;
  let adopted = 0;
  for (const inv of invoices) {
    if (!inv.cufe || inv.cufe.startsWith(SIIGO_CUFE_PREFIX)) continue;
    const nit = nitKey(inv.supplierNit);
    if (!nit) continue;
    const match = synth.find((s) => nitKey(s.supplierNit) === nit && sameInvoiceNumber(s.docnum, inv.docnum));
    if (!match) continue;
    await col.updateOne(
      { companyId, cufe: inv.cufe },
      { $set: { status: "caused", causedAt: match.causedAt, siigoId: match.siigoId, ...(match.siigoName ? { siigoName: match.siigoName } : {}) } }
    );
    await col.deleteOne({ companyId, cufe: match.cufe });
    synth.splice(synth.indexOf(match), 1);
    adopted++;
  }
  return adopted;
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
