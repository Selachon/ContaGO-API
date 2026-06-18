import { getDb } from "./database.js";

const COLLECTION = "siigoIngestedCufes";

/**
 * Registro persistente de CUFEs ya descargados por la herramienta "Contabilizar
 * Gastos Siigo" por empresa. Sirve para no re-descargar/re-procesar documentos
 * que ya se trajeron en consultas anteriores (ahorro de tiempo y ancho de banda).
 */

/** Devuelve el conjunto de CUFEs ya registrados para una empresa. */
export async function getIngestedCufes(companyId: string): Promise<Set<string>> {
  if (!companyId) return new Set();
  const docs = await getDb()
    .collection<any>(COLLECTION)
    .find({ companyId }, { projection: { cufe: 1, _id: 0 } })
    .toArray();
  return new Set(docs.map((d) => d.cufe).filter(Boolean));
}

/**
 * Conjunto de trackIds (identificador de descarga de la DIAN) ya registrados.
 * La tabla /Document/Received NO muestra el CUFE, pero SÍ el trackId (1‑a‑1 con
 * el documento), así que sirve para saltar la re-descarga. La identidad/criterio
 * de negocio sigue siendo el CUFE (que leemos del XML y guardamos).
 */
export async function getIngestedTrackIds(companyId: string): Promise<Set<string>> {
  if (!companyId) return new Set();
  const docs = await getDb()
    .collection<any>(COLLECTION)
    .find({ companyId }, { projection: { trackId: 1, _id: 0 } })
    .toArray();
  return new Set(docs.map((d) => (d.trackId || "").trim()).filter(Boolean));
}

/**
 * Registra (idempotente) los documentos descargados con éxito para una empresa.
 * El CRITERIO es el CUFE (llave única del documento, leído del XML). Se indexa
 * por CUFE cuando existe; si no, por trackId. Guarda también docnum y trackId.
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
          update: { $set: { cufe: r.cufe || "", docnum: r.docnum || "", trackId: r.trackId || "" }, $setOnInsert: { companyId, ingestedAt: now } },
          upsert: true,
        },
      };
    });
  if (ops.length === 0) return;
  await getDb().collection<any>(COLLECTION).bulkWrite(ops, { ordered: false });
}
