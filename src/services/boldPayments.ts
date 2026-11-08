import crypto from "crypto";
import { ObjectId } from "mongodb";
import { getDb } from "./database.js";
import { listCompaniesForUser, updateToolSubscription } from "./siigoCompaniesService.js";
import { addMonthsIso } from "./subscriptionExpiry.js";
import { addToolPayment, updateUser } from "./adminService.js";

const ORDERS = "boldOrders";

/** Herramientas que se pueden pagar por Bold: precio mensual por empresa (COP). El monto SIEMPRE sale de aquí, nunca del cliente. */
export const BOLD_TOOLS: Record<string, { name: string; monthlyPrice: number }> = {
  "siigo-xml-accounting": { name: "Contabilización Gastos Siigo Nube", monthlyPrice: 90_000 },
};
export const BOLD_MONTHS = [1, 3, 6, 12];
/** Máximo de empresas que se pueden comprar de una vez al activar la herramienta por la web. */
export const BOLD_MAX_COMPANIES = 10;
export const BOLD_PAYMENT_METHOD = "Link pago Bold";

const apiKey = () => process.env.BOLD_API_KEY || "";
const secretKey = () => process.env.BOLD_SECRET_KEY || "";
/** En modo de pruebas Bold firma el webhook con llave vacía. */
const webhookSecret = () => (process.env.BOLD_ENV === "test" ? "" : secretKey());
export const boldConfigured = () => !!apiKey() && !!secretKey();

const today = () => new Date().toLocaleDateString("en-CA", { timeZone: "America/Bogota" });
const monthOf = (d: string) => d.slice(0, 7);
const addMonthsPeriod = (period: string, n: number) => addMonthsIso(`${period}-01`, n).slice(0, 7);

export type BoldOrderStatus = "pending" | "paid" | "rejected";

interface BoldOrder {
  orderId: string;
  userId: string;
  toolId: string;
  months: number;
  amount: number;
  /** renewal: paga meses de empresas ya registradas. activation: el cliente aún no tiene la herramienta; compra N empresas y se activa en su libro de ventas. */
  kind: "renewal" | "activation";
  companiesCount: number;
  /** companyId → periodos (AAAA-MM) que cubre este pago */
  items: Array<{ companyId: string; companyName: string; periods: string[] }>;
  status: BoldOrderStatus;
  createdAt: string;
  paidAt?: string;
  paymentId?: string;
  boldStatus?: string;
}

const orders = () => getDb().collection<BoldOrder>(ORDERS);

/** Firma de integridad del botón: SHA-256 hex de {orderId}{monto}{divisa}{llaveSecreta}. */
export function integritySignature(orderId: string, amount: number, currency = "COP"): string {
  return crypto.createHash("sha256").update(`${orderId}${amount}${currency}${secretKey()}`).digest("hex");
}

/** Webhook: HMAC-SHA256 (hex) del body crudo en Base64, con la llave secreta. */
export function verifyWebhookSignature(rawBody: Buffer, signature: string | undefined): boolean {
  if (!signature) return false;
  const expected = crypto.createHmac("sha256", webhookSecret()).update(rawBody.toString("base64")).digest("hex");
  const a = Buffer.from(expected), b = Buffer.from(signature.trim().toLowerCase());
  return a.length === b.length && crypto.timingSafeEqual(a, b);
}

/** Primer periodo a pagar de una empresa: el siguiente al último pagado, sin quedarse en el pasado. */
function nextPeriod(history: Array<{ period?: string; paidAt?: string }> | undefined): string {
  const last = (history || []).filter((h) => h?.paidAt && h.period).map((h) => String(h.period)).sort().pop();
  const now = monthOf(today());
  const after = last ? addMonthsPeriod(last, 1) : now;
  return after > now ? after : now;
}

export interface PayableCompany { id: string; name: string; nit?: string; status: string; nextPeriod: string; paidUntil: string | null }

/** Empresas del usuario que pueden pagarse para una herramienta, con hasta cuándo están pagadas. */
/** ¿El usuario ya tiene la herramienta (renovación) o es una activación nueva? */
export async function hasTool(userId: string, toolId: string): Promise<boolean> {
  const u = await getDb().collection<any>("users").findOne({ _id: new ObjectId(userId) }, { projection: { purchasedTools: 1 } });
  return (u?.purchasedTools || []).includes(toolId);
}

export async function payableCompanies(userId: string, toolId: string): Promise<PayableCompany[]> {
  if (!BOLD_TOOLS[toolId]) throw new Error("Esta herramienta no admite pago en línea.");
  const companies = await listCompaniesForUser(userId, toolId);
  return companies
    .filter((c) => c.ownerUserId === userId && c.subscriptions?.[toolId]?.status !== "cancelled")
    .map((c) => {
      const sub = c.subscriptions?.[toolId];
      const last = (sub?.paymentHistory || []).filter((h) => h?.paidAt && h.period).map((h) => h.period).sort().pop();
      return {
        id: c.id, name: c.name, nit: c.nit, status: sub?.status || "active",
        nextPeriod: nextPeriod(sub?.paymentHistory),
        paidUntil: last ? addMonthsPeriod(last, 1) : null,
      };
    });
}

export async function createOrder(userId: string, input: { toolId: string; companyIds?: string[]; companies?: number; months: number }) {
  if (!boldConfigured()) throw new Error("El pago en línea no está disponible por ahora.");
  const tool = BOLD_TOOLS[input.toolId];
  if (!tool) throw new Error("Esta herramienta no admite pago en línea.");
  const months = Number(input.months);
  if (!BOLD_MONTHS.includes(months)) throw new Error("Cantidad de meses inválida.");
  const owned = await hasTool(userId, input.toolId);
  let items: BoldOrder["items"] = [];
  let count: number;
  if (owned) {
    const ids = [...new Set((input.companyIds || []).map(String))];
    if (!ids.length) throw new Error("Selecciona al menos una empresa.");
    const payable = await payableCompanies(userId, input.toolId);
    items = ids.map((id) => {
      const c = payable.find((p) => p.id === id);
      if (!c) throw new Error("Una de las empresas no existe o no te pertenece.");
      return { companyId: c.id, companyName: c.name, periods: Array.from({ length: months }, (_, i) => addMonthsPeriod(c.nextPeriod, i)) };
    });
    count = items.length;
  } else {
    count = Math.floor(Number(input.companies));
    if (!(count >= 1 && count <= BOLD_MAX_COMPANIES)) throw new Error(`Indica entre 1 y ${BOLD_MAX_COMPANIES} empresas.`);
  }
  const amount = tool.monthlyPrice * months * count;
  const orderId = `CG-${Date.now().toString(36)}-${crypto.randomBytes(4).toString("hex")}`;
  await orders().insertOne({ orderId, userId, toolId: input.toolId, months, amount, items, kind: owned ? "renewal" : "activation", companiesCount: count, status: "pending", createdAt: new Date().toISOString() } as BoldOrder);

  const names = owned && items.length === 1 ? items[0].companyName : `${count} ${count === 1 ? "empresa" : "empresas"}`;
  const description = `${tool.name} · ${months} ${months === 1 ? "mes" : "meses"} · ${names}`.slice(0, 100);
  return {
    orderId, amount, currency: "COP", description,
    integritySignature: integritySignature(orderId, amount),
    apiKey: apiKey(), items, companiesCount: count, kind: owned ? "renewal" : "activation",
  };
}

export async function getOrder(orderId: string, userId: string) {
  const o = await orders().findOne({ orderId, userId });
  return o ? { orderId: o.orderId, status: o.status, amount: o.amount, months: o.months, kind: o.kind, items: o.items, paidAt: o.paidAt } : null;
}

/** Aplica el pago una sola vez: reclama la orden de forma atómica y registra cada mes en cada empresa. */
export async function settleOrder(orderId: string, evt: { paymentId?: string; amount?: number; approved: boolean; boldStatus?: string }): Promise<"paid" | "rejected" | "ignored"> {
  const order = await orders().findOne({ orderId });
  if (!order) return "ignored";
  if (!evt.approved) {
    await orders().updateOne({ orderId, status: "pending" }, { $set: { status: "rejected", boldStatus: evt.boldStatus } });
    return "rejected";
  }
  if (evt.amount != null && Number(evt.amount) !== order.amount) {
    console.error(`[bold] monto no coincide en ${orderId}: esperado ${order.amount}, recibido ${evt.amount}`);
    return "ignored";
  }
  // "rejected" también puede reabrirse: Bold permite reintentar el pago de la misma orden
  const claimed = await orders().findOneAndUpdate(
    { orderId, status: { $in: ["pending", "rejected"] } },
    { $set: { status: "paid", paidAt: new Date().toISOString(), paymentId: evt.paymentId, boldStatus: evt.boldStatus } },
  );
  if (!claimed) return "paid"; // ya aplicada (webhook duplicado o consulta previa)
  try {
    if (order.kind === "activation") {
      await settleActivation(order, evt.paymentId || orderId);
      return "paid";
    }
    const perCompany = BOLD_TOOLS[order.toolId].monthlyPrice;
    for (const item of order.items) {
      for (const period of item.periods) {
        await updateToolSubscription(item.companyId, order.toolId, {
          paidAt: today(), period, paymentAmount: perCompany, paymentMethod: BOLD_PAYMENT_METHOD,
          invoiceRef: evt.paymentId || orderId, paymentStatus: "paid",
        });
      }
      // Empresa sin fecha de inicio: el primer pago la fija (el vencimiento se calcula desde ahí)
      const doc = await getDb().collection<any>("siigoCompanies").findOne({ _id: new ObjectId(item.companyId) }, { projection: { subscriptions: 1 } });
      if (!doc?.subscriptions?.[order.toolId]?.licenseStartDate) {
        await updateToolSubscription(item.companyId, order.toolId, { licenseStartDate: `${item.periods[0]}-${today().slice(8)}` });
      }
    }
    return "paid";
  } catch (err) {
    // Reabrir para que el reintento del webhook (o la consulta) vuelva a aplicarla; los registros por periodo son idempotentes
    await orders().updateOne({ orderId }, { $set: { status: "pending" } });
    throw err;
  }
}

/** Consulta de respaldo (webhook puede tardar hasta 10 min): estado de la venta en Bold por orderId. */
export async function pullOrderStatus(orderId: string): Promise<"APPROVED" | "REJECTED" | "PENDING" | "NONE"> {
  const res = await fetch(`https://payments.api.bold.co/v2/payment-voucher/${encodeURIComponent(orderId)}`, {
    headers: { Authorization: `x-api-key ${apiKey()}` },
  });
  if (!res.ok) return "NONE";
  const j: any = await res.json().catch(() => ({}));
  const s = String(j.payment_status || "");
  if (s === "APPROVED") return "APPROVED";
  if (["REJECTED", "FAILED", "VOIDED"].includes(s)) return "REJECTED";
  if (s === "NO_TRANSACTION_FOUND" || !s) return "NONE";
  return "PENDING";
}

/** Activación por la web: queda como venta en el libro del servicio (igual que la registra el admin) y fija el cupo de empresas comprado. */
async function settleActivation(order: BoldOrder, ref: string) {
  const rec = await getDb().collection<any>("users").findOne({ _id: new ObjectId(order.userId) }, { projection: { toolSubscriptions: 1 } });
  const sub = rec?.toolSubscriptions?.[order.toolId] || {};
  const alreadyHasPlan = !!sub.licenseEndDate && sub.status !== "cancelled";
  await addToolPayment(order.userId, order.toolId, {
    type: alreadyHasPlan ? "renovacion" : "activacion",
    paidAt: today(), amount: order.amount, method: BOLD_PAYMENT_METHOD, invoiceRef: ref,
    periodicityMonths: order.months,
    note: `Pago en línea (Bold) · ${order.companiesCount} ${order.companiesCount === 1 ? "empresa" : "empresas"}`,
  });
  await updateUser(order.userId, { [`toolCompanyLimits.${order.toolId}`]: order.companiesCount });
}
