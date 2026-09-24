/**
 * Control de jobs DIAN frente a reinicios y jobs "huérfanos".
 *
 * Cada ruta DIAN guarda sus jobs en un Map en memoria. Eso tenía dos problemas:
 *
 *  1. Un reinicio (cada redeploy de Railway) borraba los jobs: el usuario veía
 *     "Job no encontrado" sin saber qué pasó, y los archivos temporales
 *     quedaban en disco.
 *  2. Los jobs que nadie consulta (el usuario cerró la pestaña) o que se
 *     quedaron sin avanzar seguían ocupando navegador, memoria y disco hasta
 *     el TTL de 2-3 h; acumulados saturaban el servidor y empezaban los errores.
 *
 * Este módulo, sin cambiar la lógica de cada ruta:
 *  - Guarda una foto liviana de cada job en Mongo (`dianJobs`) con latido, para
 *    poder decir "se interrumpió por un reinicio" en vez de un 404 mudo.
 *  - Corre un chequeo concurrente (timer propio, no bloquea a los jobs) que
 *      · cancela jobs abandonados (sin consultas del cliente),
 *      · detiene jobs colgados (sin avance),
 *      · marca como interrumpidos los jobs cuyo dueño dejó de latir (otra
 *        instancia muerta), y
 *      · borra de disco los restos de jobs que ya no existen.
 */
import fs from "fs";
import path from "path";
import { randomUUID } from "crypto";
import { fileURLToPath } from "url";
import type { NextFunction, Request, Response } from "express";

// ── Tipos ────────────────────────────────────────────────────────────────────

/** Forma mínima que necesitamos de un job de cualquier ruta DIAN. */
export interface GuardedJob {
  status: string;
  progress?: { step?: string; current?: number; total?: number; detalle?: string };
  userId?: string;
  createdAt?: number;
  error?: string;
  /** Lo pone el guardián al abortar el job; las rutas lo respetan vía isJobAborted. */
  aborted?: boolean;
  tempDir?: string;
  excelPath?: string;
  outputPath?: string;
  zipPath?: string;
  filesZipPath?: string;
}

export interface JobSnapshot {
  _id: string; // `${tool}:${jobId}`
  tool: string;
  jobId: string;
  userId?: string;
  status: string;
  step?: string;
  current?: number;
  total?: number;
  error?: string;
  createdAt: number;
  updatedAt: number;
  heartbeatAt: number;
  instanceId: string;
  expireAt: Date;
}

export interface JobStore {
  upsertMany(docs: JobSnapshot[]): Promise<void>;
  /** Marca "interrumpido" los jobs activos de OTRAS instancias sin latido desde `olderThan`. */
  markStaleInterrupted(olderThan: number, selfInstanceId: string, msg: string, now: number): Promise<number>;
  /** Marca "interrumpido" los jobs activos de esta instancia (apagado ordenado). */
  markInstanceInterrupted(instanceId: string, msg: string, now: number): Promise<number>;
  find(id: string): Promise<JobSnapshot | null>;
}

export const INTERRUPTED_MSG =
  "El servidor se reinició mientras se procesaba tu solicitud y el trabajo se interrumpió. " +
  "Vuelve a iniciar la generación (no se pierde nada en la DIAN).";
export const GONE_MSG =
  "El resultado ya no está disponible porque el servidor se reinició. Vuelve a iniciar la generación.";

// ── Configuración ────────────────────────────────────────────────────────────

const num = (v: string | undefined, def: number): number => {
  const n = Number(v);
  return Number.isFinite(n) && n >= 0 ? n : def;
};

export const guardConfig = {
  /** Sin consultas del cliente durante este tiempo => job abandonado. 0 desactiva. */
  abandonMs: () => num(process.env.JOB_ABANDON_MS, 10 * 60_000),
  /** Sin cambio de progreso durante este tiempo => job colgado. 0 desactiva. */
  stallMs: () => num(process.env.JOB_STALL_MS, 15 * 60_000),
  /** Cada cuánto corre el chequeo de huérfanos. */
  watchIntervalMs: () => num(process.env.JOB_WATCH_INTERVAL_MS, 30_000),
  /** Cada cuánto se persiste el latido. */
  heartbeatMs: () => num(process.env.JOB_HEARTBEAT_MS, 15_000),
  /** Sin latido durante este tiempo => la instancia dueña murió. */
  staleHeartbeatMs: () => num(process.env.JOB_STALE_HEARTBEAT_MS, 60_000),
  /** Edad mínima para borrar restos en disco que ningún job referencia. */
  orphanFileAgeMs: () => num(process.env.JOB_ORPHAN_FILE_AGE_MS, 4 * 60 * 60_000),
};

const SNAPSHOT_TTL_MS = 24 * 60 * 60_000;
const TERMINAL = new Set(["completed", "error", "cancelled", "interrupted"]);
const isActive = (s: string): boolean => s === "pending" || s === "processing";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOWNLOADS_DIR = path.join(__dirname, "../../downloads");

// ── Estado del guardián ──────────────────────────────────────────────────────

export const instanceId = `${process.env.RAILWAY_DEPLOYMENT_ID || "local"}-${randomUUID().slice(0, 8)}`;

interface ToolEntry { map: Map<string, GuardedJob> }
const tools = new Map<string, ToolEntry>();
/** Última consulta del cliente por job. Va aparte del job para no tocar sus tipos. */
const lastPoll = new Map<string, number>();
interface Track { sig: string; sigSince: number; seen: number }
const track = new Map<string, Track>();
const persistedTerminal = new Set<string>();

let store: JobStore | null = null;
export function setJobStore(s: JobStore | null): void { store = s; }

const keyOf = (tool: string, jobId: string): string => `${tool}:${jobId}`;

/** Registra el Map de jobs de una ruta para que el guardián lo vigile. */
export function attachJobs(tool: string, map: Map<string, GuardedJob>): void {
  tools.set(tool, { map });
}

/** Las rutas lo usan en su isJobCancelled: cancelado por el usuario O abortado por el guardián. */
export function isJobAborted(job: GuardedJob | undefined | null): boolean {
  return !!job && (job.status === "cancelled" || job.aborted === true);
}

// ── Chequeo de huérfanos (abandonados / colgados) ────────────────────────────

function cleanupJobFiles(job: GuardedJob): void {
  for (const dir of [job.tempDir]) {
    if (dir) { try { fs.rmSync(dir, { recursive: true, force: true }); } catch { /* ya no está */ } }
  }
  for (const f of [job.excelPath, job.outputPath, job.zipPath, job.filesZipPath]) {
    if (f) { try { fs.unlinkSync(f); } catch { /* ya no está */ } }
  }
}

const jobSignature = (job: GuardedJob): string =>
  `${job.status}|${job.progress?.step ?? ""}|${job.progress?.current ?? 0}|${job.progress?.total ?? 0}`;

/** Esperando cupo de navegador: no es un cuelgue, no cuenta como "sin avance". */
const isQueued = (job: GuardedJob): boolean => /^en cola/i.test(job.progress?.step ?? "");

function abortJob(job: GuardedJob, status: "cancelled" | "error", step: string, message: string): void {
  job.aborted = true;
  job.status = status;
  if (status === "error") job.error = message;
  job.progress = { ...(job.progress || {}), step, detalle: message };
  cleanupJobFiles(job);
}

export interface OrphanCheckResult { active: number; abandoned: string[]; stalled: string[] }

/**
 * Un chequeo sobre todos los jobs de todas las rutas. Es síncrono y barato
 * (recorre Maps en memoria); lo dispara un timer aparte, así corre mientras los
 * jobs siguen trabajando.
 */
export function runOrphanCheck(now: number = Date.now()): OrphanCheckResult {
  const result: OrphanCheckResult = { active: 0, abandoned: [], stalled: [] };
  const abandonMs = guardConfig.abandonMs();
  const stallMs = guardConfig.stallMs();

  for (const [tool, { map }] of tools) {
    for (const [jobId, job] of map) {
      const key = keyOf(tool, jobId);
      if (!isActive(job.status)) { track.delete(key); lastPoll.delete(key); continue; }
      result.active++;

      let t = track.get(key);
      if (!t) { t = { sig: "", sigSince: now, seen: now }; track.set(key, t); }
      const sig = jobSignature(job);
      if (sig !== t.sig) { t.sig = sig; t.sigSince = now; }

      const lastActivity = Math.max(lastPoll.get(key) ?? 0, job.createdAt ?? t.seen);
      if (abandonMs > 0 && now - lastActivity > abandonMs) {
        abortJob(job, "cancelled", "Cancelado por inactividad",
          `El cliente dejó de consultar el estado hace más de ${Math.round(abandonMs / 60000)} min; el job se canceló para liberar recursos.`);
        result.abandoned.push(key);
        continue;
      }
      if (stallMs > 0 && !isQueued(job) && now - t.sigSince > stallMs) {
        abortJob(job, "error", "Detenido por falta de avance",
          `El proceso no reportó avance en ${Math.round(stallMs / 60000)} min y se detuvo. Vuelve a intentarlo.`);
        result.stalled.push(key);
      }
    }
  }

  for (const k of [...lastPoll.keys()]) {
    const [tool, ...rest] = k.split(":");
    if (!tools.get(tool)?.map.has(rest.join(":"))) lastPoll.delete(k);
  }
  if (result.abandoned.length || result.stalled.length) {
    console.warn(`[JobGuard] huérfanos: abandonados=[${result.abandoned.join(", ")}] colgados=[${result.stalled.join(", ")}]`);
  }
  return result;
}

// ── Restos en disco ──────────────────────────────────────────────────────────

/**
 * Borra de `dir` archivos/carpetas viejos que ningún job vivo referencia (restos
 * de jobs que murieron con un reinicio o un error). Nunca toca lo que un job en
 * memoria referencia ni algo cuyo nombre contenga el id de un job vivo.
 */
export function sweepOrphanFiles(dir: string = DOWNLOADS_DIR, minAgeMs: number = guardConfig.orphanFileAgeMs(), now: number = Date.now()): number {
  let entries: string[];
  try { entries = fs.readdirSync(dir); } catch { return 0; }

  const referenced = new Set<string>();
  const liveIds: string[] = [];
  for (const { map } of tools.values()) {
    for (const [jobId, job] of map) {
      liveIds.push(jobId);
      for (const p of [job.tempDir, job.excelPath, job.outputPath, job.zipPath, job.filesZipPath]) {
        if (p) referenced.add(path.resolve(p));
      }
    }
  }

  let removed = 0;
  for (const name of entries) {
    const full = path.join(dir, name);
    if (referenced.has(path.resolve(full))) continue;
    if (liveIds.some((id) => name.includes(id))) continue;
    try {
      if (now - fs.statSync(full).mtimeMs < minAgeMs) continue;
      fs.rmSync(full, { recursive: true, force: true });
      removed++;
    } catch { /* en uso o ya borrado */ }
  }
  if (removed > 0) console.warn(`[JobGuard] restos en disco: ${removed} elemento(s) huérfano(s) borrado(s) de ${dir}.`);
  return removed;
}

// ── Persistencia (latido) ────────────────────────────────────────────────────

function snapshotOf(tool: string, jobId: string, job: GuardedJob, now: number): JobSnapshot {
  return {
    _id: keyOf(tool, jobId), tool, jobId,
    userId: job.userId,
    status: job.status,
    step: job.progress?.step,
    current: job.progress?.current,
    total: job.progress?.total,
    error: job.error,
    createdAt: job.createdAt ?? now,
    updatedAt: now,
    heartbeatAt: now,
    instanceId,
    expireAt: new Date(now + SNAPSHOT_TTL_MS),
  };
}

/** Persiste el latido de los jobs activos y, una sola vez, el estado final de los terminados. */
export async function persistHeartbeat(now: number = Date.now()): Promise<void> {
  if (!store) return;
  const docs: JobSnapshot[] = [];
  for (const [tool, { map }] of tools) {
    for (const [jobId, job] of map) {
      const key = keyOf(tool, jobId);
      if (isActive(job.status)) {
        persistedTerminal.delete(key);
        docs.push(snapshotOf(tool, jobId, job, now));
      } else if (!persistedTerminal.has(key)) {
        persistedTerminal.add(key);
        docs.push(snapshotOf(tool, jobId, job, now));
      }
    }
  }
  if (docs.length === 0) return;
  try { await store.upsertMany(docs); }
  catch (err) { console.warn("[JobGuard] no se pudo persistir el latido:", (err as Error)?.message); }
}

/** Marca como interrumpidos los jobs cuyo dueño (otra instancia) dejó de latir. */
export async function reapDeadInstanceJobs(now: number = Date.now()): Promise<number> {
  if (!store) return 0;
  try {
    const n = await store.markStaleInterrupted(now - guardConfig.staleHeartbeatMs(), instanceId, INTERRUPTED_MSG, now);
    if (n > 0) console.warn(`[JobGuard] ${n} job(s) de una instancia caída marcados como interrumpidos.`);
    return n;
  } catch (err) {
    console.warn("[JobGuard] no se pudo revisar jobs de instancias caídas:", (err as Error)?.message);
    return 0;
  }
}

/** Cantidad de jobs activos en esta instancia. */
export function activeJobCount(): number {
  let n = 0;
  for (const { map } of tools.values()) for (const job of map.values()) if (isActive(job.status)) n++;
  return n;
}

/**
 * Apagado ordenado: espera (hasta `drainMs`) a que terminen los jobs en curso y
 * deja constancia inmediata de los que no alcancen, para que el cliente reciba
 * un mensaje claro en vez de un 404 cuando consulte a la instancia nueva.
 */
export async function onShutdown(drainMs: number = num(process.env.SHUTDOWN_DRAIN_MS, 0)): Promise<void> {
  const deadline = Date.now() + drainMs;
  while (activeJobCount() > 0 && Date.now() < deadline) {
    await new Promise((r) => setTimeout(r, 1000));
  }
  const pending = activeJobCount();
  if (pending === 0 || !store) return;
  console.warn(`[JobGuard] apagado con ${pending} job(s) en curso: se marcan como interrumpidos.`);
  try {
    const now = Date.now();
    for (const { map } of tools.values()) {
      for (const job of map.values()) {
        if (isActive(job.status)) { job.status = "error"; job.error = INTERRUPTED_MSG; job.aborted = true; }
      }
    }
    await persistHeartbeat(now);
    await store.markInstanceInterrupted(instanceId, INTERRUPTED_MSG, now);
  } catch (err) {
    console.warn("[JobGuard] no se pudo registrar la interrupción por apagado:", (err as Error)?.message);
  }
}

// ── Middleware por ruta ──────────────────────────────────────────────────────

const POLL_PATH = /^\/(?:job-status|job-download|job-download-zip|download|download-files)\/([A-Za-z0-9_-]+)/;

export interface JobGuardOptions {
  /** Rutas de consulta del cliente (estado y resultado/descarga); el grupo 1 debe ser el jobId. */
  pollPath?: RegExp;
  /** De las rutas de consulta, cuáles son de "estado" (responden JSON de progreso). El resto responde 410. */
  isStatusPath?: (path: string) => boolean;
  /** Cuerpo de la respuesta de estado para un job perdido, en el formato que espera el frontend de esa ruta. */
  statusBody?: (message: string, snap: JobSnapshot) => unknown;
  /** Cuerpo de la respuesta 410 para las consultas de resultado/descarga. */
  goneBody?: (message: string) => unknown;
}

const defaultStatusBody = (message: string, snap: JobSnapshot): unknown => ({
  status: "error", error: message, interrupted: true,
  progress: { step: "Interrumpido", current: snap.current ?? 0, total: snap.total ?? 0 },
});
const defaultGoneBody = (message: string): unknown => ({ status: "error", detalle: message });

/**
 * Se monta DESPUÉS de la autenticación. Para las consultas de estado/descarga:
 *  - si el job existe en memoria: registra que el cliente sigue ahí (evita que
 *    lo cancelen por abandono);
 *  - si NO existe (típicamente tras un reinicio): responde con lo que quedó en
 *    Mongo en vez del "Job no encontrado" mudo; si tampoco hay registro, deja
 *    seguir para que la ruta responda 404 como siempre.
 */
export function jobGuardMiddleware(tool: string, opts: JobGuardOptions = {}) {
  const pollPath = opts.pollPath ?? POLL_PATH;
  const isStatusPath = opts.isStatusPath ?? ((p: string) => p.startsWith("/job-status/"));
  const statusBody = opts.statusBody ?? defaultStatusBody;
  const goneBody = opts.goneBody ?? defaultGoneBody;
  return async (req: Request, res: Response, next: NextFunction): Promise<void> => {
    try {
      if (req.method !== "GET") { next(); return; }
      const m = pollPath.exec(req.path);
      if (!m) { next(); return; }
      const jobId = m[1];
      const user = (req as Request & { user?: { userId?: string; isAdmin?: boolean } }).user;
      const job = tools.get(tool)?.map.get(jobId);

      if (job) {
        if (user && (job.userId === user.userId || user.isAdmin)) lastPoll.set(keyOf(tool, jobId), Date.now());
        next();
        return;
      }
      if (!store || !user) { next(); return; }

      const snap = await store.find(keyOf(tool, jobId));
      if (!snap || (snap.userId !== user.userId && !user.isAdmin)) { next(); return; }

      const message = snap.status === "interrupted" || isActive(snap.status)
        ? INTERRUPTED_MSG
        : snap.status === "error" ? (snap.error || GONE_MSG) : GONE_MSG;
      if (isStatusPath(req.path)) {
        res.json(statusBody(message, snap));
      } else {
        res.status(410).json(goneBody(message));
      }
    } catch (err) {
      console.warn("[JobGuard] middleware:", (err as Error)?.message);
      next();
    }
  };
}

// ── Timers ───────────────────────────────────────────────────────────────────

const timers: NodeJS.Timeout[] = [];

/** Arranca latido, reaper de instancias caídas, chequeo de huérfanos y barrido de disco. */
export function startJobGuard(jobStore?: JobStore): void {
  if (timers.length > 0) return;
  if (jobStore) store = jobStore;

  // Al arrancar, restos de jobs de la ejecución anterior (ningún job en memoria los reclama).
  try { sweepOrphanFiles(DOWNLOADS_DIR, 30 * 60_000); } catch { /* best-effort */ }

  const every = (ms: number, fn: () => void | Promise<void>) => {
    if (ms <= 0) return;
    const t = setInterval(() => { Promise.resolve().then(fn).catch((e) => console.warn("[JobGuard] tick:", e?.message || e)); }, ms);
    t.unref();
    timers.push(t);
  };
  every(guardConfig.watchIntervalMs(), () => { runOrphanCheck(); });
  every(guardConfig.heartbeatMs(), async () => { await persistHeartbeat(); await reapDeadInstanceJobs(); });
  every(10 * 60_000, () => { sweepOrphanFiles(); });
  console.log(`[JobGuard] activo (instancia ${instanceId}): abandono=${guardConfig.abandonMs() / 60000}min colgado=${guardConfig.stallMs() / 60000}min`);
}

export function stopJobGuard(): void {
  for (const t of timers) clearInterval(t);
  timers.length = 0;
}

/** Solo para pruebas. */
export function _resetJobGuardForTests(): void {
  stopJobGuard();
  tools.clear(); lastPoll.clear(); track.clear(); persistedTerminal.clear(); store = null;
}
