import { ObjectId } from "mongodb";
import { randomUUID } from "crypto";
import { getDb } from "./database.js";
import type { ToolPayment } from "./adminService.js";
import { rutAvailability } from "./rutStorage.js";

/** Comisión del vendedor sobre el total de la venta (dato informativo). */
export const COMMISSION_RATE = 0.15;

const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
const TOOL_ID_RE = /^[a-z0-9-]+$/;

const users = () => getDb().collection<any>("users");

export interface FacturacionRow {
  id: string; // userId|toolId|paymentId
  userId: string;
  toolId: string;
  paymentId: string;
  clientName: string;
  clientId?: string;
  userEmail: string;
  type: ToolPayment["type"];
  paidAt: string;
  dateEstimated: boolean;
  amount: number;
  gift: boolean;
  commission: number;
  method?: string;
  invoiceRef?: string;
  invoicedAt?: string;
  periodStart?: string;
  periodEnd?: string;
  periodicityMonths?: number;
  alignToBase?: boolean;
  addCompanies?: number;
  addNits?: string[];
  planAfter?: ToolPayment["planAfter"];
  imported: boolean;
  note?: string;
  /** Devolución: pide nota crédito en lugar de factura; `original` es la venta que se devuelve. */
  refund: boolean;
  original?: { paidAt: string; amount: number; invoiceRef?: string; type: ToolPayment["type"] };
}

/** Datos de facturación del cliente (se muestran al facturar, no en la tabla). */
export interface FacturacionClient {
  userId: string;
  name: string;
  userEmail: string;
  phone?: string;
  isCompany?: boolean;
  idType?: string;
  id?: string;
  billName?: string;
  billEmail?: string;
  city?: string;
  address?: string;
  hasRut: boolean;
}

function clientOf(u: any): { name: string; id?: string } {
  const b = u.billing;
  const name = b?.isCompany ? b.companyName : [b?.firstName, b?.lastName].filter(Boolean).join(" ");
  return { name: name || u.name || u.email, id: b?.id };
}

/** Todas las ventas del libro de pagos (activación, renovación y mejora) de los clientes, con los datos de facturación de cada cliente. */
export async function listLedgerRows(): Promise<{ rows: FacturacionRow[]; clients: Record<string, FacturacionClient> }> {
  const docs = await users()
    .find({ toolSubscriptions: { $exists: true }, role: { $ne: "DEMO" } }, { projection: { email: 1, name: 1, billing: 1, toolSubscriptions: 1 } })
    .toArray();
  const rows: FacturacionRow[] = [];
  const clients: Record<string, FacturacionClient> = {};
  const ruts: { id: string; filename?: string }[] = [];
  for (const u of docs) {
    const client = clientOf(u);
    const b = u.billing;
    clients[String(u._id)] = {
      userId: String(u._id), name: u.name || u.email, userEmail: u.email, phone: u.phone,
      isCompany: b?.isCompany, idType: b?.idType, id: b?.id, billName: client.name, billEmail: b?.email,
      city: b?.city, address: b?.address, hasRut: false,
    };
    if (b?.rutFile) ruts.push({ id: String(u._id), filename: b.rutFile.filename });
    for (const [toolId, sub] of Object.entries<any>(u.toolSubscriptions || {})) {
      const all = (sub?.payments || []) as ToolPayment[];
      for (const p of all) {
        const amount = Number(p.amount) || 0;
        const orig = p.type === "devolucion" && p.refundOf ? all.find(x => x.id === p.refundOf) : undefined;
        const gift = p.method === "Obsequio";
        rows.push({
          id: `${u._id}|${toolId}|${p.id}`, userId: String(u._id), toolId, paymentId: p.id,
          clientName: client.name, clientId: client.id, userEmail: u.email,
          type: p.type, paidAt: p.paidAt, dateEstimated: !!p.dateEstimated, amount, gift,
          commission: gift ? 0 : Math.round(amount * COMMISSION_RATE),
          method: p.method, invoiceRef: p.invoiceRef, invoicedAt: p.invoicedAt,
          periodStart: p.periodStart, periodEnd: p.periodEnd, periodicityMonths: p.periodicityMonths, alignToBase: p.alignToBase,
          addCompanies: p.addCompanies, addNits: p.addNits, planAfter: p.planAfter,
          imported: !!p.imported, note: p.note, refund: p.type === "devolucion",
          ...(orig && { original: { paidAt: orig.paidAt, amount: Number(orig.amount) || 0, invoiceRef: orig.invoiceRef, type: orig.type } }),
        });
      }
    }
  }
  const available = ruts.length ? await rutAvailability(ruts) : new Set<string>();
  for (const id of available) if (clients[id]) clients[id].hasRut = true;
  return { rows: rows.sort((a, b) => b.paidAt.localeCompare(a.paidAt)), clients };
}

export interface InvoiceTarget { userId: string; toolId: string; paymentId: string }

/**
 * Relaciona una factura (emitida afuera) con una o varias ventas; vacío = quitar la factura.
 * La referencia queda en la venta, así se ve también en la ficha del usuario.
 */
export async function setInvoice(targets: InvoiceTarget[], invoiceRef: string, invoicedAt?: string): Promise<number> {
  const ref = String(invoiceRef || "").trim();
  if (!targets.length) throw new Error("Seleccione al menos una venta.");
  if (targets.length > 200) throw new Error("Demasiadas ventas en una sola operación.");
  if (ref && invoicedAt && !DATE_RE.test(invoicedAt)) throw new Error("Fecha de factura inválida (use AAAA-MM-DD).");
  const date = ref ? (invoicedAt || new Date().toISOString().slice(0, 10)) : undefined;
  if (!ref) throw new Error("Las facturas relacionadas no se pueden quitar: la venta queda bloqueada al facturarse.");
  // Antes de escribir nada: ninguna de las ventas puede estar ya facturada
  for (const t of targets) {
    if (!ObjectId.isValid(t.userId) || !TOOL_ID_RE.test(t.toolId) || !t.paymentId) throw new Error("Venta inválida.");
    const u = await users().findOne({ _id: new ObjectId(t.userId) }, { projection: { toolSubscriptions: 1 } });
    const p = (u?.toolSubscriptions?.[t.toolId]?.payments || []).find((x: any) => x.id === t.paymentId);
    if (p?.invoiceRef) throw new Error("Una de las ventas ya está facturada y no se puede modificar. Solo se cambia por código interno.");
  }
  let changed = 0;
  for (const t of targets) {
    const key = `toolSubscriptions.${t.toolId}.payments`;
    const $set: Record<string, unknown> = { updated_at: new Date().toISOString() };
    const $unset: Record<string, ""> = {};
    if (ref) { $set[`${key}.$[p].invoiceRef`] = ref; $set[`${key}.$[p].invoicedAt`] = date; }
    else { $unset[`${key}.$[p].invoiceRef`] = ""; $unset[`${key}.$[p].invoicedAt`] = ""; }
    const res = await users().updateOne(
      { _id: new ObjectId(t.userId), [key]: { $elemMatch: { id: t.paymentId, invoiceRef: { $exists: false } } } },
      { $set, ...(Object.keys($unset).length ? { $unset } : {}) } as any,
      { arrayFilters: [{ "p.id": t.paymentId }] }
    );
    if (res.matchedCount === 0) continue;
    changed++;
    // La factura de la última activación/renovación se refleja también en la suscripción (otros módulos la leen ahí)
    const u = await users().findOne({ _id: new ObjectId(t.userId) }, { projection: { toolSubscriptions: 1 } });
    const pays: any[] = (u?.toolSubscriptions?.[t.toolId]?.payments || []).filter((p: any) => p.type !== "mejora" && p.invoiceRef);
    pays.sort((a, b) => String(b.paidAt).localeCompare(String(a.paidAt)));
    const base = `toolSubscriptions.${t.toolId}.invoiceRef`;
    await users().updateOne({ _id: new ObjectId(t.userId) }, pays[0] ? { $set: { [base]: pays[0].invoiceRef } } : { $unset: { [base]: "" } } as any);
  }
  return changed;
}

// ─── Importación de pagos anteriores al libro ────────────────────────────────

const monthsBetween = (a: string, b: string) => Math.round((new Date(b).getTime() - new Date(a).getTime()) / 86_400_000 / 30.44);

function addMonthsIso(d: string, n: number): string {
  const [y, m, day] = d.slice(0, 10).split("-").map(Number);
  const t = m - 1 + n;
  const yy = y + Math.floor(t / 12), mm = ((t % 12) + 12) % 12;
  const dd = Math.min(day, new Date(Date.UTC(yy, mm + 1, 0)).getUTCDate());
  return `${yy}-${String(mm + 1).padStart(2, "0")}-${String(dd).padStart(2, "0")}`;
}

export interface LegacyImportPlan {
  users: number;
  entries: number;
  estimatedDate: number;
  skippedNoDate: Array<{ user: string; toolId: string }>;
  skippedNoValue: Array<{ user: string; toolId: string }>;
  writes: Array<{ userId: string; toolId: string; payments: ToolPayment[] }>;
}

/** Arma (sin escribir) los pagos anteriores que aún no están en el libro. Sin fecha de pago se usa el inicio de la suscripción. */
export async function planLegacyImport(): Promise<LegacyImportPlan> {
  const docs = await users().find({ is_admin: { $ne: true }, role: { $ne: "DEMO" } }).toArray();
  const plan: LegacyImportPlan = { users: 0, entries: 0, estimatedDate: 0, skippedNoDate: [], skippedNoValue: [], writes: [] };
  const now = new Date().toISOString();
  const touched = new Set<string>();

  for (const u of docs) {
    const tools: string[] = [...new Set<string>([...(u.purchasedTools || []), ...Object.keys(u.toolSubscriptions || {})])];
    // Los datos sueltos del usuario solo se atribuyen a una herramienta si es la única
    const single = tools.length === 1;
    for (const toolId of tools) {
      if (!TOOL_ID_RE.test(toolId)) continue;
      const sub: any = u.toolSubscriptions?.[toolId] || {};
      const ledger: ToolPayment[] = sub.payments || [];
      const hasLedger = ledger.length > 0;
      // Meses mensuales ya cubiertos por el libro (para no duplicarlos al reimportar)
      const covered = new Set(ledger.map(p => p.periodStart?.slice(0, 7)).filter(Boolean));
      const label = u.email;
      const method = sub.paymentMethod ?? (single ? u.paymentMethod : undefined);
      const out: ToolPayment[] = [];
      const history: any[] = [...(sub.paymentHistory || [])].filter(h => h?.paidAt && !covered.has(h.period)).sort((a, b) => String(a.period).localeCompare(String(b.period)));

      if (history.length) {
        // Servicio mensual: cada mes pagado es una venta
        const startDay = (sub.licenseStartDate || "").slice(8, 10) || "01";
        history.forEach((h, i) => {
          const amount = Number(h.amount ?? sub.paymentAmount ?? (single ? u.paymentAmount : undefined));
          if (isNaN(amount)) { plan.skippedNoValue.push({ user: label, toolId }); return; }
          const periodStart = `${h.period}-${startDay}`;
          const inv = (sub.invoiceHistory || []).find((x: any) => x.period === h.period);
          out.push({
            id: randomUUID(), type: i === 0 && !hasLedger ? "activacion" : "renovacion", paidAt: String(h.paidAt).slice(0, 10), amount,
            ...(method && { method }), ...(inv?.invoiceRef && { invoiceRef: inv.invoiceRef }), ...(inv?.invoicedAt && { invoicedAt: inv.invoicedAt }),
            periodicityMonths: 1, periodStart, periodEnd: addMonthsIso(periodStart, 1),
            note: "Importado de datos anteriores", createdAt: now, imported: true,
          });
        });
      } else if (!hasLedger) {
        const start: string | undefined = sub.licenseStartDate ?? (single ? u.licenseStartDate : undefined);
        const end: string | undefined = sub.licenseEndDate ?? (single ? u.licenseEndDate : undefined);
        const raw = sub.paymentAmount ?? (single ? u.paymentAmount : undefined);
        if (raw == null) {
          if (sub.licenseStartDate || sub.paidAt) plan.skippedNoValue.push({ user: label, toolId });
          continue;
        }
        const paidAt: string | undefined = sub.paidAt ? String(sub.paidAt).slice(0, 10) : start;
        if (!paidAt || !DATE_RE.test(paidAt)) { plan.skippedNoDate.push({ user: label, toolId }); continue; }
        const estimated = !sub.paidAt;
        const months = start && end ? monthsBetween(start, end) : 0;
        const invoiceRef = sub.invoiceRef ?? (single ? u.invoiceRef : undefined);
        out.push({
          id: randomUUID(), type: "activacion", paidAt, amount: Number(raw),
          ...(method && { method }), ...(invoiceRef && { invoiceRef }),
          ...(months > 0 && { periodicityMonths: months }), ...(start && { periodStart: start }), ...(end && { periodEnd: end }),
          note: `Importado de datos anteriores${estimated ? " (fecha estimada: inicio de la suscripción)" : ""}`,
          createdAt: now, imported: true, ...(estimated && { dateEstimated: true }),
        });
        if (estimated) plan.estimatedDate++;
      }
      if (out.length) {
        plan.writes.push({ userId: String(u._id), toolId, payments: out });
        plan.entries += out.length;
        touched.add(String(u._id));
      }
    }
  }
  plan.users = touched.size;
  return plan;
}

export async function runLegacyImport(): Promise<LegacyImportPlan> {
  const plan = await planLegacyImport();
  for (const w of plan.writes) {
    await users().updateOne(
      { _id: new ObjectId(w.userId) },
      { $push: { [`toolSubscriptions.${w.toolId}.payments`]: { $each: w.payments } }, $set: { updated_at: new Date().toISOString() } } as any
    );
  }
  return plan;
}
