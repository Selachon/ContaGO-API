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
  return new Set(docs.map((d) => d.cufe));
}

/** Registra (idempotente) los CUFEs descargados con éxito para una empresa. */
export async function recordIngestedCufes(
  companyId: string,
  records: { cufe: string; docnum?: string }[]
): Promise<void> {
  if (!companyId || records.length === 0) return;
  const now = new Date().toISOString();
  const ops = records
    .filter((r) => r.cufe)
    .map((r) => ({
      updateOne: {
        filter: { companyId, cufe: r.cufe },
        update: { $setOnInsert: { companyId, cufe: r.cufe, docnum: r.docnum || "", ingestedAt: now } },
        upsert: true,
      },
    }));
  if (ops.length === 0) return;
  await getDb().collection<any>(COLLECTION).bulkWrite(ops, { ordered: false });
}
