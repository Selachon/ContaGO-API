import { getDb } from "./database.js";

/** Días de gracia tras el vencimiento antes de suspender el uso de la herramienta. */
export const GRACE_DAYS = Number(process.env.SUBSCRIPTION_GRACE_DAYS ?? 5);
/** Interruptor de emergencia: ENFORCE_SUBSCRIPTION_EXPIRY=false desactiva la suspensión sin tocar código. */
export const ENFORCE_EXPIRY = process.env.ENFORCE_SUBSCRIPTION_EXPIRY !== "false";

// Herramientas que se cobran por empresa Siigo: su vencimiento sale de las suscripciones de la empresa
const COMPANY_TOOLS = new Set(["siigo-xml-accounting", "siigo-egresos", "conciliacion-bancaria"]);

const DATE_RE = /^\d{4}-\d{2}-\d{2}/;
const monthsBetween = (a: string, b: string) => (Number(b.slice(0, 4)) - Number(a.slice(0, 4))) * 12 + Number(b.slice(5, 7)) - Number(a.slice(5, 7));

export function addMonthsIso(d: string, n: number): string {
  const [y, m, day] = d.slice(0, 10).split("-").map(Number);
  const t = m - 1 + n;
  const yy = y + Math.floor(t / 12), mm = ((t % 12) + 12) % 12;
  const dd = Math.min(day, new Date(Date.UTC(yy, mm + 1, 0)).getUTCDate());
  return `${yy}-${String(mm + 1).padStart(2, "0")}-${String(dd).padStart(2, "0")}`;
}

export function addDaysIso(d: string, n: number): string {
  const t = new Date(`${d.slice(0, 10)}T00:00:00Z`);
  t.setUTCDate(t.getUTCDate() + n);
  return t.toISOString().slice(0, 10);
}

// Cobro mensual: cada pago cubre un mes desde el día de inicio; el plan llega hasta el último mes pagado.
// licenseEndDate no se actualiza con cada pago, así que se toma el mayor entre ambos.
function resolveEnd(start: string | undefined, end: string | undefined, history: any[] | undefined): string | undefined {
  const last = (history || []).filter((h) => h?.paidAt && h.period).map((h) => String(h.period)).sort().pop();
  if (last && start && DATE_RE.test(start)) {
    const m = addMonthsIso(start, monthsBetween(start.slice(0, 7), last) + 1);
    if (!end || m > end) return m;
  }
  return end;
}

/** Vencimiento de una herramienta del cliente (null = sin fecha: no vence). Misma lógica que la hoja de Suscripciones. */
export function toolEnd(user: any, toolId: string, companies: any[], depth = 0): string | null {
  const ts = user.toolSubscriptions?.[toolId] || {};
  // Obsequio ligado a otra herramienta: sigue sus fechas de renovación
  if (ts.followsTool && depth === 0) {
    const target = toolEnd(user, ts.followsTool, companies, depth + 1);
    if (target) return target;
  }
  if (COMPANY_TOOLS.has(toolId)) {
    const ends: string[] = [];
    for (const co of companies) {
      const s = co.subscriptions?.[toolId];
      if (!s || s.status === "cancelled") continue;
      const e = resolveEnd(s.licenseStartDate, s.licenseEndDate, s.paymentHistory);
      if (e) ends.push(e);
    }
    // Las ventas del libro (Añadir venta) no actualizan la suscripción de la empresa: también cuentan
    const saleEnd = (ts.payments || []).filter((p: any) => p.type !== "mejora" && p.type !== "devolucion" && p.periodEnd).map((p: any) => String(p.periodEnd)).sort().pop();
    if (saleEnd) ends.push(saleEnd);
    // Con varias empresas se toma la más lejana: el acceso es del usuario, no de una empresa
    if (ends.length) return ends.sort().pop()!;
  }
  const sold = (ts.payments || []).filter((p: any) => p.type !== "mejora" && p.type !== "devolucion" && p.periodEnd);
  const lastSaleEnd: string | undefined = sold.map((p: any) => p.periodEnd as string).sort().pop();
  const stored = [ts.licenseEndDate, lastSaleEnd].filter(Boolean).sort().pop() ?? user.licenseEndDate;
  const start = ts.licenseStartDate ?? user.licenseStartDate;
  return resolveEnd(start, stored, ts.paymentHistory) ?? null;
}

export interface SuspendedTool { toolId: string; end: string; suspendedFrom: string }

/** Herramientas cuyo vencimiento + días de gracia ya pasó. Los administradores y las cuentas DEMO no se suspenden por aquí. */
export function suspendedToolsOf(user: any, companies: any[], today = new Date().toISOString().slice(0, 10)): SuspendedTool[] {
  if (!user || user.is_admin || user.role === "ADMIN" || user.role === "DEMO" || user.demo) return [];
  // Excepción concedida por un administrador: conserva el acceso hasta esa fecha (inclusive)
  if (user.suspensionExemptUntil && today <= String(user.suspensionExemptUntil).slice(0, 10)) return [];
  const out: SuspendedTool[] = [];
  for (const toolId of user.purchasedTools || []) {
    // Plan cerrado sin fecha de vencimiento: se corta desde el aviso del cliente
    const closedAt = user.toolSubscriptions?.[toolId]?.closedAt ? String(user.toolSubscriptions[toolId].cancelledAt || user.toolSubscriptions[toolId].closedAt).slice(0, 10) : null;
    const end = toolEnd(user, toolId, companies) ?? closedAt;
    if (!end) continue;
    const suspendedFrom = addDaysIso(end, GRACE_DAYS + 1); // el último día de acceso es vencimiento + gracia
    if (today >= suspendedFrom) out.push({ toolId, end: end.slice(0, 10), suspendedFrom });
  }
  return out;
}

// Caché corta por usuario: la regla se consulta en cada petición a una herramienta
const cache = new Map<string, { at: number; list: SuspendedTool[] }>();
const TTL_MS = 60_000;

export function clearSuspensionCache(userId?: string): void {
  if (userId) cache.delete(userId); else cache.clear();
}

export async function getSuspendedTools(userId: string): Promise<SuspendedTool[]> {
  if (!ENFORCE_EXPIRY) return [];
  const hit = cache.get(userId);
  if (hit && Date.now() - hit.at < TTL_MS) return hit.list;
  const { ObjectId } = await import("mongodb");
  let oid: any;
  try { oid = new ObjectId(userId); } catch { return []; }
  const db = getDb();
  const user = await db.collection<any>("users").findOne({ _id: oid }, { projection: { password_hash: 0, google_drive: 0, google_drives: 0 } });
  if (!user) return [];
  const companies = await db.collection<any>("siigoCompanies").find({ ownerUserId: userId }, { projection: { subscriptions: 1 } }).toArray();
  const list = suspendedToolsOf(user, companies);
  cache.set(userId, { at: Date.now(), list });
  return list;
}
