/**
 * Sincroniza el registro de facturas DIAN de una empresa con las compras que ya
 * existen en Siigo (fuente de verdad), para que TODA factura causada aparezca en
 * la pestaña "Causadas" de su mes, sin importar por dónde se causó.
 *
 * - Incremental: tras la primera sincronización completa (12 meses) solo pide las
 *   compras creadas desde la última (con 3 días de traslape); una completa a la semana.
 * - Una sola ejecución a la vez por empresa; si hay una en curso, se reutiliza.
 * - Nunca lanza hacia quien la consulta salvo `syncCausadasNow`: si Siigo falla, el
 *   registro queda como estaba.
 */
import { getDb } from "./database.js";
import { fetchSiigoPurchases } from "./siigoAccountingService.js";
import { planCausadasSync, planSiigoDocBackfill } from "./siigoCausadasPlan.js";
import type { DianInvoiceRecord } from "./siigoIngestedCufesService.js";

const REGISTRY = "siigoIngestedCufes";
const META = "siigoCausadasSync";
const DAY = 24 * 60 * 60 * 1000;
const FULL_EVERY_MS = 7 * DAY;
const OVERLAP_MS = 3 * DAY;
const MONTHS_BACK = 12;

export interface CausadasSyncResult {
  companyId: string;
  mode: "full" | "incremental";
  since: string;
  purchases: number;
  updated: number;
  inserted: number;
  skipped: number;
  ms: number;
}

const inflight = new Map<string, Promise<CausadasSyncResult>>();

const isoDay = (ms: number): string => new Date(ms).toISOString().slice(0, 10);

async function runSync(companyId: string, force: boolean): Promise<CausadasSyncResult> {
  const t0 = Date.now();
  const db = getDb();
  const meta = await db.collection<any>(META).findOne({ companyId });
  const now = Date.now();
  const twelveMonthsAgo = (() => { const d = new Date(now); d.setMonth(d.getMonth() - MONTHS_BACK); return d.getTime(); })();

  const lastFull = meta?.lastFullAt ? Date.parse(meta.lastFullAt) : 0;
  const lastSync = meta?.lastSyncAt ? Date.parse(meta.lastSyncAt) : 0;
  const full = force || !lastFull || now - lastFull > FULL_EVERY_MS || !lastSync;
  const sinceMs = full ? twelveMonthsAgo : Math.max(twelveMonthsAgo, lastSync - OVERLAP_MS);
  const since = isoDay(sinceMs);

  const purchases = await fetchSiigoPurchases(since);

  const existing = (await db
    .collection<any>(REGISTRY)
    .find({ companyId }, { projection: { _id: 0 } })
    .toArray()) as DianInvoiceRecord[];
  const nowIso = new Date(now).toISOString();
  const plan = planCausadasSync(companyId, existing, purchases, nowIso);

  const ops: any[] = [
    ...plan.updates.map((u) => ({ updateOne: { filter: { companyId, cufe: u.cufe }, update: { $set: u.set } } })),
    ...plan.inserts.map((r) => {
      const { companyId: _c, cufe: _k, ...rest } = r;
      return { updateOne: { filter: { companyId, cufe: r.cufe }, update: { $setOnInsert: { companyId, cufe: r.cufe, ...rest } }, upsert: true } };
    }),
  ];
  if (ops.length > 0) await db.collection<any>(REGISTRY).bulkWrite(ops, { ordered: false });

  await db.collection<any>(META).updateOne(
    { companyId },
    { $set: { companyId, lastSyncAt: nowIso, ...(full ? { lastFullAt: nowIso } : {}), lastResult: { purchases: purchases.length, updated: plan.updated, inserted: plan.inserted, skipped: plan.skipped }, lastError: null } },
    { upsert: true },
  );

  const result: CausadasSyncResult = {
    companyId, mode: full ? "full" : "incremental", since, purchases: purchases.length,
    updated: plan.updated, inserted: plan.inserted, skipped: plan.skipped, ms: Date.now() - t0,
  };
  console.log(`[CausadasSync] ${companyId.slice(-6)} ${result.mode} desde ${since}: ${result.purchases} compras · +${result.inserted} nuevas · ${result.updated} marcadas causadas · ${result.skipped} ya reflejadas (${result.ms} ms)`);
  return result;
}

/** ¿Hay una sincronización en curso para esta empresa? */
export function isCausadasSyncRunning(companyId: string): boolean {
  return inflight.has(companyId);
}

/** Sincroniza ahora (o reutiliza la ejecución en curso). Lanza si Siigo/Mongo fallan. */
export function syncCausadasNow(companyId: string, force = false): Promise<CausadasSyncResult> {
  const running = inflight.get(companyId);
  if (running) return running;
  const p = runSync(companyId, force).catch(async (err) => {
    await getDb().collection<any>(META)
      .updateOne({ companyId }, { $set: { companyId, lastError: { at: new Date().toISOString(), message: err instanceof Error ? err.message : String(err) } } }, { upsert: true })
      .catch(() => undefined);
    throw err;
  }).finally(() => inflight.delete(companyId));
  inflight.set(companyId, p);
  return p;
}

/**
 * Sincroniza solo si la última vez fue hace más de `maxAgeMs` (por defecto 10 min).
 * Devuelve null si no hacía falta. NUNCA lanza: la lectura del listado no debe
 * depender de que Siigo responda.
 */
export async function syncCausadasIfStale(companyId: string, maxAgeMs = 10 * 60_000): Promise<CausadasSyncResult | null> {
  try {
    if (!inflight.has(companyId)) {
      const meta = await getDb().collection<any>(META).findOne({ companyId });
      if (meta?.lastSyncAt && Date.now() - Date.parse(meta.lastSyncAt) < maxAgeMs) return null;
    }
    return await syncCausadasNow(companyId);
  } catch (err) {
    console.warn(`[CausadasSync] ${companyId.slice(-6)} no se pudo sincronizar con Siigo:`, err instanceof Error ? err.message : err);
    return null;
  }
}

export interface DocBackfillResult { purchases: number; updated: number; notFound: number; since: string; ms: number }

/**
 * Completa los datos del comprobante de Siigo (nombre, consecutivo, fecha, total) en las
 * facturas causadas de la empresa que solo guardaban el id. Es SOLO lectura en Siigo y
 * SOLO enriquece el registro propio: no cambia estados ni crea facturas. Sirve para el
 * histórico anterior a que el comprobante se guardara al causar.
 */
export async function backfillSiigoDocs(companyId: string, months = MONTHS_BACK): Promise<DocBackfillResult> {
  const t0 = Date.now();
  const db = getDb();
  const since = (() => { const d = new Date(); d.setMonth(d.getMonth() - Math.min(Math.max(1, months), 36)); return d.toISOString().slice(0, 10); })();
  const existing = (await db.collection<any>(REGISTRY).find({ companyId, status: "caused" }, { projection: { _id: 0 } }).toArray()) as DianInvoiceRecord[];
  if (!existing.some((d) => d.siigoId && !d.siigoName)) return { purchases: 0, updated: 0, notFound: 0, since, ms: Date.now() - t0 };
  const purchases = await fetchSiigoPurchases(since);
  const plan = planSiigoDocBackfill(existing, purchases);
  if (plan.updates.length > 0) {
    await db.collection<any>(REGISTRY).bulkWrite(
      plan.updates.map((u) => ({ updateOne: { filter: { companyId, cufe: u.cufe }, update: { $set: u.set } } })),
      { ordered: false },
    );
  }
  console.log(`[CausadasSync] ${companyId.slice(-6)} backfill de comprobantes desde ${since}: ${plan.updates.length} completados · ${plan.notFound} sin comprobante en Siigo (${Date.now() - t0} ms)`);
  return { purchases: purchases.length, updated: plan.updates.length, notFound: plan.notFound, since, ms: Date.now() - t0 };
}
