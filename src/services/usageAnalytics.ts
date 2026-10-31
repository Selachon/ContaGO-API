/**
 * Analítica de uso del portal: captura de eventos (entradas y operaciones de herramientas).
 *
 * Principios:
 *  - NUNCA bloquea ni rompe la petición: registrar es "dispara y olvida" y todo error se traga.
 *  - NUNCA se guarda contenido de archivos, cuerpos de petición ni la IP; solo metadatos:
 *    quién, cuándo, qué herramienta, cuántas facturas, cuánto tardó, si salió bien y la causa del error.
 *  - La ubicación se guarda como municipio aproximado, calculado en el servidor (ver geoIp.ts).
 *
 * Dos fuentes de operaciones:
 *  - "request": operaciones síncronas (la respuesta HTTP ES la operación). Solo las rutas de la lista
 *    blanca OPERATIONS cuentan; guardar configuración o consultar listados no es "usar" la herramienta.
 *  - "job": operaciones asíncronas (descargas DIAN y similares). La petición solo las inicia, así que la
 *    duración real se mide cuando el job llega a un estado final (ver setJobFinishedHandler en jobGuard).
 */
import type { Request, Response } from "express";
import type { Collection } from "mongodb";
import { getDb } from "./database.js";
import { resolveGeo, type Geo } from "./geoIp.js";
import { setJobFinishedHandler, type GuardedJob } from "./jobGuard.js";

export type UsageKind = "login" | "use";

export interface UsageEvent {
  ts: Date;
  userId: string;
  kind: UsageKind;
  /** id de la herramienta del portal (toolsCatalog) */
  tool?: string;
  /** nombre corto de la operación (p. ej. "reconcile", "download") */
  op?: string;
  src?: "request" | "job";
  /** facturas/documentos procesados, cuando se conocen */
  docs?: number;
  ok?: boolean;
  /** causa del error, ya resumida (no el cuerpo de la respuesta) */
  err?: string;
  /** duración en milisegundos */
  durMs?: number;
  geo?: Geo;
}

export interface EventSink {
  insert(ev: UsageEvent): Promise<void>;
}

const COLLECTION = "usageEvents";
/** Se conservan 14 meses: alcanza para comparar contra el mismo mes del año anterior sin guardar de más. */
const RETENTION_DAYS = 425;

let sink: EventSink | null = null;

function mongoSink(): EventSink {
  return {
    async insert(ev) {
      await getDb().collection(COLLECTION).insertOne({ ...ev });
    },
  };
}

/** Solo para pruebas. */
export function _setSinkForTests(s: EventSink | null): void { sink = s; }

export async function initUsageAnalytics(): Promise<void> {
  sink = mongoSink();
  const col: Collection = getDb().collection(COLLECTION);
  const safe = async (keys: Record<string, 1 | -1>, opts?: Parameters<Collection["createIndex"]>[1]) => {
    try { await col.createIndex(keys, opts); } catch (err) {
      console.warn("[Usage] índice omitido:", JSON.stringify(keys), (err as Error)?.message);
    }
  };
  await safe({ ts: -1 });
  await safe({ userId: 1, ts: -1 });
  await safe({ tool: 1, ts: -1 });
  await safe({ ts: 1 }, { expireAfterSeconds: RETENTION_DAYS * 86_400, name: "ts_ttl" });
  setJobFinishedHandler(onJobFinished);
}

/** Dispara y olvida: jamás lanza ni espera. */
export function recordEvent(ev: UsageEvent): void {
  const s = sink;
  if (!s) return;
  try {
    s.insert(ev).catch((err) => console.warn("[Usage] no se pudo registrar el evento:", (err as Error)?.message));
  } catch (err) {
    console.warn("[Usage] no se pudo registrar el evento:", (err as Error)?.message);
  }
}

// ── Entradas ────────────────────────────────────────────────────────────────

export function recordLogin(req: Request, userId: string): void {
  recordEvent({ ts: new Date(), userId, kind: "login", op: "login", geo: resolveGeo(req.ip || req.socket?.remoteAddress) || undefined });
}

/**
 * La sesión dura 7 días, así que quien vuelve al portal casi nunca vuelve a hacer login. Cada vez que el
 * portal verifica la sesión (/auth/me) cuenta como una "entrada" de regreso, como máximo una cada 6 horas.
 */
const SESSION_GAP_MS = 6 * 3_600_000;
const lastTouch = new Map<string, number>();
export function recordSessionReturn(req: Request, userId: string, now: number = Date.now()): void {
  const prev = lastTouch.get(userId);
  if (prev !== undefined && now - prev < SESSION_GAP_MS) return;
  lastTouch.set(userId, now);
  if (lastTouch.size > 5000) for (const [k, t] of lastTouch) if (now - t >= SESSION_GAP_MS) lastTouch.delete(k);
  recordEvent({ ts: new Date(now), userId, kind: "login", op: "return", geo: resolveGeo(req.ip || req.socket?.remoteAddress) || undefined });
}

/** Solo para pruebas. */
export function _resetSessionTouchForTests(): void { lastTouch.clear(); }

// ── Operaciones síncronas (por petición) ────────────────────────────────────

interface OperationRule { method: string; path: RegExp; tool?: string; op: string }

/**
 * Qué peticiones cuentan como "usar" cada herramienta. `tool` reemplaza la herramienta del gate cuando el
 * router se comparte (p. ej. el router de Siigo sirve a "Contabilizar gastos" y a "Egresos").
 * Las herramientas de descarga DIAN no están aquí: se miden al terminar el job.
 */
export const OPERATIONS: Record<string, OperationRule[]> = {
  "conciliacion-bancaria": [{ method: "POST", path: /^\/reconcile$/, op: "reconcile" }],
  "siigo-xml-accounting": [
    { method: "POST", path: /^\/egresos\/(submit|income\/voucher|income\/journal)$/, tool: "siigo-egresos", op: "egreso" },
    { method: "POST", path: /^\/accounting\/(submit|process-xml|process-batch)$/, op: "causar" },
  ],
  "contabilizacion-dian-siigo": [{ method: "POST", path: /^\/(compras|ventas)$/, op: "contabilizar" }],
  "documentos-soporte": [{ method: "POST", path: /^\/$/, op: "crear" }],
  "facturacion-electronica": [{ method: "POST", path: /^\/empresas\/[^/]+\/(facturas|notas-credito|documentos-soporte)\/[^/]+\/emitir$/, op: "emitir" }],
  "rut-consulta": [{ method: "POST", path: /^\/(consultar|bulk)$/, op: "consultar" }],
};

export function classifyOperation(gateTool: string, method: string, path: string): { tool: string; op: string } | null {
  const rules = OPERATIONS[gateTool];
  if (!rules) return null;
  const m = method.toUpperCase();
  for (const r of rules) {
    if (r.method === m && r.path.test(path)) return { tool: r.tool || gateTool, op: r.op };
  }
  return null;
}

/** Resume el error sin copiar el cuerpo de la respuesta (puede traer datos del cliente). */
export function summarizeError(status: number): string {
  if (status === 401 || status === 403) return "Sin autorización";
  if (status === 404) return "No encontrado";
  if (status === 408 || status === 504) return "Tiempo de espera agotado";
  if (status === 413) return "Archivo demasiado grande";
  if (status === 429) return "Demasiadas solicitudes";
  if (status >= 500) return "Error del servidor";
  return "Solicitud inválida";
}

/** Llamado por requireToolAccess cuando el acceso ya fue concedido. No hace nada si la petición no es una operación. */
export function trackToolOperation(req: Request, res: Response, gateTool: string, userId: string): void {
  try {
    const op = classifyOperation(gateTool, req.method, req.path);
    if (!op) return;
    const started = Date.now();
    res.once("finish", () => {
      const ok = res.statusCode < 400;
      recordEvent({
        ts: new Date(), userId, kind: "use", tool: op.tool, op: op.op, src: "request", ok,
        ...(ok ? {} : { err: summarizeError(res.statusCode) }),
        durMs: Date.now() - started,
        geo: resolveGeo(req.ip || req.socket?.remoteAddress) || undefined,
      });
    });
  } catch { /* nunca afectar la petición */ }
}

// ── Operaciones asíncronas (jobs) ───────────────────────────────────────────

/** id del job en jobGuard → id de la herramienta en el portal. Los que no estén aquí no se registran. */
export const JOB_TOOLS: Record<string, { tool: string; op: string }> = {
  "dian-cufe": { tool: "dian-cufe-downloader", op: "descarga" },
  "dian-recibidos": { tool: "dian-recibidos", op: "descarga" },
  "dian-mass-download": { tool: "dian-mass-download", op: "descarga" },
  "dian": { tool: "dian-downloader", op: "descarga" },
  "dian-excel": { tool: "dian-excel-exporter", op: "excel" },
  "dian-third-parties": { tool: "dian-third-parties-excel", op: "terceros" },
  "ci2-dian-traer": { tool: "causacion-interfaces-v2", op: "traer" },
  "siigo-dian-ingest": { tool: "siigo-xml-accounting", op: "ingesta" },
};

function summarizeJobError(msg: string | undefined): string {
  const m = (msg || "").toLowerCase();
  if (!m) return "Error desconocido";
  if (m.includes("bloque")) return "La DIAN bloqueó la solicitud";
  if (m.includes("token") && (m.includes("venc") || m.includes("expir") || m.includes("inv"))) return "Token DIAN vencido";
  if (m.includes("tiempo") || m.includes("timeout") || m.includes("timed out")) return "Tiempo de espera agotado";
  if (m.includes("reinici") || m.includes("interrump")) return "Interrumpido por reinicio";
  if (m.includes("drive")) return "Error con Google Drive";
  if (m.includes("credencial")) return "Credenciales inválidas";
  return "Error en el proceso";
}

export function onJobFinished(e: { tool: string; jobId: string; job: GuardedJob; durMs: number }): void {
  const map = JOB_TOOLS[e.tool];
  const userId = e.job.userId;
  if (!map || !userId) return;
  // Cancelado por el usuario: no es éxito ni falla del servicio; no ensucia ninguna de las dos métricas.
  if (e.job.status === "cancelled") return;
  const ok = e.job.status === "completed";
  recordEvent({
    ts: new Date(), userId, kind: "use", tool: map.tool, op: map.op, src: "job", ok,
    docs: ok ? Math.max(0, Math.round(Number(e.job.progress?.total) || 0)) : 0,
    ...(ok ? {} : { err: summarizeJobError(e.job.error) }),
    durMs: e.durMs,
  });
}

// ── Lectura para el panel de admin ──────────────────────────────────────────

export interface CompactEvent {
  t: number; u: string; k: UsageKind; tool?: string; op?: string; docs?: number; ok?: boolean; err?: string; dur?: number;
  city?: string; region?: string; country?: string;
}

/** Eventos en formato compacto para el panel (se piden por rango, con tope para no saturar la respuesta). */
export async function listEvents(days: number, limit = 200_000): Promise<CompactEvent[]> {
  const since = new Date(Date.now() - days * 86_400_000);
  const rows = await getDb().collection(COLLECTION).find({ ts: { $gte: since } }).sort({ ts: 1 }).limit(limit).toArray();
  return rows.map((r) => ({
    t: new Date(r.ts as Date).getTime(), u: String(r.userId), k: r.kind as UsageKind,
    tool: r.tool, op: r.op, docs: r.docs, ok: r.ok, err: r.err, dur: r.durMs,
    city: r.geo?.city, region: r.geo?.region, country: r.geo?.country,
  }));
}

export interface ClientRow {
  id: string; name: string; email: string; status: string; created: number;
  tools: string[]; licenseEnd?: number; paid?: number; nitsAllowed?: number;
  /** municipio más frecuente según la IP (aproximado) */
  city?: string; region?: string;
}

interface UserDoc {
  _id: { toString(): string }; name?: string; email?: string; status?: string; created_at?: string; is_admin?: boolean;
  purchasedTools?: string[]; licenseEndDate?: string; paymentAmount?: number; companiesInPlan?: number;
}

/** Clientes (sin administradores) con su municipio más frecuente. Los administradores quedan fuera: sus pruebas no son uso de clientes. */
export function summarizeClients(users: UserDoc[], events: CompactEvent[]): ClientRow[] {
  const byUser = new Map<string, Map<string, { n: number; region?: string }>>();
  for (const e of events) {
    if (!e.city) continue;
    const m = byUser.get(e.u) || new Map();
    const cur = m.get(e.city) || { n: 0, region: e.region };
    cur.n++;
    m.set(e.city, cur);
    byUser.set(e.u, m);
  }
  return users.filter((u) => !u.is_admin).map((u) => {
    const id = u._id.toString();
    const top = [...(byUser.get(id) || [])].sort((a, b) => b[1].n - a[1].n)[0];
    const created = u.created_at ? Date.parse(u.created_at) : NaN;
    const end = u.licenseEndDate ? Date.parse(u.licenseEndDate) : NaN;
    return {
      id, name: u.name || u.email || id, email: u.email || "", status: u.status || "active",
      created: Number.isFinite(created) ? created : 0, tools: u.purchasedTools || [],
      ...(Number.isFinite(end) ? { licenseEnd: end } : {}),
      ...(typeof u.paymentAmount === "number" ? { paid: u.paymentAmount } : {}),
      ...(typeof u.companiesInPlan === "number" ? { nitsAllowed: u.companiesInPlan } : {}),
      ...(top ? { city: top[0], region: top[1].region } : {}),
    };
  });
}

export async function analyticsPayload(days: number): Promise<{ generatedAt: number; days: number; clients: ClientRow[]; events: CompactEvent[] }> {
  const users = (await getDb().collection("users").find({ is_admin: { $ne: true } }).toArray()) as unknown as UserDoc[];
  const all = await listEvents(days);
  const clients = summarizeClients(users, all);
  const ids = new Set(clients.map((c) => c.id));
  const events = all.filter((e) => ids.has(e.u));
  return { generatedAt: Date.now(), days, clients, events };
}
