/**
 * Creación de INGRESOS en Siigo desde la herramienta de Causación de Egresos.
 *  - Recibo de Caja (RC, /v1/vouchers): Abono a deuda (cruza FV) y Anticipo.
 *  - Avanzado: como Siigo eliminó el RC "Detailed", el método de cuentas
 *    contables manuales (DB/CR) se crea como Comprobante Contable (CC, /v1/journals).
 *
 * Limitación: la API no tiene reporte de cuentas por cobrar; para el cruce de
 * facturas de venta se listan las facturas del cliente con /v1/invoices.
 */
import { createVoucher, createJournalVoucher, listInvoices, SiigoError } from "./siigoService.js";

const round2 = (n: unknown): number => Math.round((Number(n) || 0) * 100) / 100;
const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;

// ─── Listar facturas de venta de un cliente (para el cruce DebtPayment) ───

export interface CustomerInvoice {
  id: string;
  name: string;
  prefix: string;
  consecutive: number;
  quote: number;
  date: string;
  total: number;
  balance: number | null;
}

/** Deriva prefijo+consecutivo del nombre de la factura ("FV-1-123" → FV-1 / 123). */
function splitName(name: string): { prefix: string; consecutive: number } {
  const i = String(name).lastIndexOf("-");
  if (i < 0) return { prefix: String(name), consecutive: 0 };
  return { prefix: name.slice(0, i), consecutive: Number(name.slice(i + 1).replace(/\D/g, "")) || 0 };
}

export async function listCustomerInvoices(nit: string): Promise<CustomerInvoice[]> {
  const id = String(nit || "").replace(/\D/g, "");
  if (!id) return [];
  const out: CustomerInvoice[] = [];
  for (let page = 1; page <= 5; page++) {
    let raw: any;
    try {
      raw = await listInvoices({ customer_identification: id, page, page_size: 100 });
    } catch (e) {
      // Siigo returns 404 when the customer has no invoices instead of an empty array.
      if (e instanceof SiigoError && e.status === 404) break;
      throw e;
    }
    const rows: any[] = Array.isArray(raw?.results) ? raw.results : Array.isArray(raw) ? raw : [];
    if (rows.length === 0) break;
    for (const inv of rows) {
      const name = String(inv?.name || "");
      const { prefix, consecutive } = splitName(name);
      out.push({
        id: String(inv?.id ?? ""),
        name,
        prefix,
        consecutive,
        quote: 1,
        date: String(inv?.date ?? ""),
        total: Number(inv?.total) || 0,
        balance: inv?.balance != null ? Number(inv.balance) : null,
      });
    }
    if (rows.length < 100) break;
  }
  // Más recientes primero.
  return out.sort((a, b) => (a.date < b.date ? 1 : -1)).slice(0, 100);
}

// ─── Recibo de Caja (RC) ──────────────────────────────────────────────────

export type VoucherType = "DebtPayment" | "AdvancePayment" | "Detailed";

export interface VoucherDebtItem {
  due: { prefix: string; consecutive: number; quote: number; date?: string };
  value: number;
}

export interface VoucherDetailedItem {
  account: { code: string };
  customer: { identification: string; branch_office?: number };
  description?: string;
  value: number;
}

export interface VoucherPayload {
  document: { id: number };
  type: VoucherType;
  date: string;
  customer: { identification: string; branch_office?: number };
  items?: Array<VoucherDebtItem | VoucherDetailedItem>;
  payment: { id: number; value: number };
  observations?: string;
}

export function validateVoucher(p: Partial<VoucherPayload> | undefined): string[] {
  const e: string[] = [];
  if (!p || typeof p !== "object") return ["Payload vacío."];
  if (!p.document?.id || !Number.isFinite(Number(p.document.id))) e.push("Falta el tipo de comprobante (document.id).");
  if (!["DebtPayment", "AdvancePayment", "Detailed"].includes(p.type as string)) e.push("Tipo de recibo de caja inválido.");
  if (!p.date || !ISO_DATE.test(String(p.date))) e.push("Fecha inválida (AAAA-MM-DD).");
  if (!p.customer?.identification) e.push("Falta el cliente (customer.identification).");
  if (!p.payment?.id || !Number.isFinite(Number(p.payment.id))) e.push("Falta el medio de pago (payment.id).");
  if (!p.payment || !(Number(p.payment.value) > 0)) e.push("El valor debe ser mayor a 0.");
  if (p.type === "DebtPayment") {
    if (!Array.isArray(p.items) || p.items.length === 0) e.push("Un abono a deuda requiere al menos una factura cruzada.");
    else p.items.forEach((it: any, i) => {
      if (!it.due?.prefix) e.push(`Factura ${i + 1}: falta el prefijo.`);
      if (!Number.isFinite(Number(it.due?.consecutive))) e.push(`Factura ${i + 1}: falta el consecutivo.`);
      if (!(Number(it.value) > 0)) e.push(`Factura ${i + 1}: valor inválido.`);
    });
  }
  if (p.type === "Detailed") {
    if (!Array.isArray(p.items) || p.items.length === 0) e.push("El RC avanzado requiere al menos una partida.");
    else p.items.forEach((it: any, i) => {
      if (!it.account?.code) e.push(`Partida ${i + 1}: falta la cuenta contable.`);
      if (!it.customer?.identification) e.push(`Partida ${i + 1}: falta el tercero.`);
      if (!(Number(it.value) > 0)) e.push(`Partida ${i + 1}: valor inválido.`);
    });
  }
  return e;
}

export async function submitVoucher(p: VoucherPayload): Promise<unknown> {
  const clean: VoucherPayload = {
    ...p,
    items: p.items?.map((it) => ({ ...it, value: round2(it.value) })),
    payment: { ...p.payment, value: round2(p.payment.value) },
  };
  return createVoucher(clean);
}

// ─── Avanzado = Comprobante Contable (CC) ─────────────────────────────────

export interface JournalItem {
  account: { code: string; movement: "Debit" | "Credit" };
  customer: { identification: string; branch_office?: number };
  due?: { prefix: string; consecutive: number; quote: number; date?: string };
  description?: string;
  value: number;
}

export interface JournalPayload {
  document: { id: number };
  date: string;
  items: JournalItem[];
  observations?: string;
}

export function validateIncomeJournal(p: Partial<JournalPayload> | undefined): string[] {
  const e: string[] = [];
  if (!p || typeof p !== "object") return ["Payload vacío."];
  if (!p.document?.id || !Number.isFinite(Number(p.document.id))) e.push("Falta el tipo de comprobante contable (document.id).");
  if (!p.date || !ISO_DATE.test(String(p.date))) e.push("Fecha inválida (AAAA-MM-DD).");
  if (!Array.isArray(p.items) || p.items.length === 0) { e.push("El comprobante requiere al menos una partida."); return e; }
  let deb = 0, cred = 0;
  p.items.forEach((it, i) => {
    if (!it.account?.code) e.push(`Partida ${i + 1}: falta la cuenta contable.`);
    if (it.account?.movement !== "Debit" && it.account?.movement !== "Credit") e.push(`Partida ${i + 1}: movimiento debe ser Debit o Credit.`);
    if (!it.customer?.identification) e.push(`Partida ${i + 1}: falta el tercero (customer.identification).`);
    if (!(Number(it.value) > 0)) e.push(`Partida ${i + 1}: valor inválido.`);
    if (it.account?.movement === "Debit") deb += Number(it.value) || 0;
    if (it.account?.movement === "Credit") cred += Number(it.value) || 0;
  });
  if (e.length === 0 && Math.abs(deb - cred) > 1) e.push(`No está balanceado: débitos ${deb} ≠ créditos ${cred}.`);
  return e;
}

export async function submitIncomeJournal(p: JournalPayload): Promise<unknown> {
  const clean: JournalPayload = { ...p, items: p.items.map((it) => ({ ...it, value: round2(it.value) })) };
  return createJournalVoucher(clean);
}
