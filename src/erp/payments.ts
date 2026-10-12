/**
 * Cuentas por pagar, programación, pago efectivo y egreso en Siigo.
 *
 * Modelo (§37): Payable (el propio documento con `balance`) → PaymentSchedule
 * (programar ≠ pagar) → Payment (pago real) → PaymentAllocation (a qué
 * documentos se imputó) → Recibo de pago en Siigo.
 *
 * Reutiliza el motor de egresos ya probado (`siigoEgresosService`).
 */
import {
  submitEgreso,
  validateEgresoPayload,
  findExistingEgresoInSiigo,
  type EgresoItem,
  type EgresoPayload,
} from "../services/siigoEgresosService.js";
import { SiigoError } from "../services/siigoService.js";
import { COL, col, fromDoc, fromDocs, now, scoped, toId, today } from "./db.js";
import { audit, finishOperation, startOperation } from "./audit.js";
import { gte, round2, subtract, sum } from "./money.js";
import { documentLabel, getDocument, patchDocument, requireDocument, transitionDocument } from "./documents.js";
import { translateSiigoError, type FriendlyError } from "./posting.js";
import { notify } from "./notifications.js";
import { ErpHttpError, type ErpRequestContext } from "./context.js";
import type { ErpBankAccount, ErpDocument, ErpPayment, ErpPaymentAllocation, ErpPaymentSchedule } from "./types.js";

// ─── Cuentas por pagar ───────────────────────────────────────────────────

export interface PayableRow {
  documentId: string;
  supplierNit: string;
  supplierName: string;
  label: string;
  type: string;
  issueDate: string;
  dueDate?: string;
  estimatedPaymentDate?: string;
  total: number;
  balance: number;
  paid: number;
  daysToDue: number | null;
  costCenterId?: number;
  status: ErpDocument["status"];
  siigoName?: string;
  siigoPrefix?: string;
  siigoConsecutive?: number;
  agingBucket: AgingBucket;
}

export type AgingBucket = "NOT_DUE" | "D_1_30" | "D_31_60" | "D_61_90" | "D_90_PLUS";

export const AGING_LABELS: Record<AgingBucket, string> = {
  NOT_DUE: "Por vencer",
  D_1_30: "1-30 días",
  D_31_60: "31-60 días",
  D_61_90: "61-90 días",
  D_90_PLUS: "+90 días",
};

function daysBetween(from: string, to: string): number {
  const a = Date.parse(`${from}T00:00:00Z`);
  const b = Date.parse(`${to}T00:00:00Z`);
  if (Number.isNaN(a) || Number.isNaN(b)) return 0;
  return Math.round((b - a) / 86_400_000);
}

export function agingBucket(dueDate: string | undefined, reference = today()): AgingBucket {
  if (!dueDate) return "NOT_DUE";
  const overdue = daysBetween(dueDate, reference);
  if (overdue <= 0) return "NOT_DUE";
  if (overdue <= 30) return "D_1_30";
  if (overdue <= 60) return "D_31_60";
  if (overdue <= 90) return "D_61_90";
  return "D_90_PLUS";
}

/** Descompone el nombre del comprobante Siigo ("FC-1-482") en prefijo y consecutivo. */
export function splitSiigoName(name?: string): { prefix?: string; consecutive?: number } {
  if (!name) return {};
  const parts = String(name).split("-");
  if (parts.length < 2) return {};
  const consecutive = Number(parts[parts.length - 1]);
  const prefix = parts.slice(0, -1).join("-");
  return { prefix, consecutive: Number.isFinite(consecutive) ? consecutive : undefined };
}

function toPayableRow(doc: ErpDocument): PayableRow {
  const { prefix, consecutive } = splitSiigoName(doc.siigo?.siigoName);
  return {
    documentId: doc.id,
    supplierNit: doc.supplierNit,
    supplierName: doc.supplierName,
    label: documentLabel(doc),
    type: doc.type,
    issueDate: doc.issueDate,
    dueDate: doc.dueDate,
    estimatedPaymentDate: doc.estimatedPaymentDate,
    total: doc.total,
    balance: doc.balance,
    paid: subtract(doc.total, doc.balance),
    daysToDue: doc.dueDate ? daysBetween(today(), doc.dueDate) : null,
    costCenterId: doc.costCenterId ?? doc.accounting?.costCenterId,
    status: doc.status,
    siigoName: doc.siigo?.siigoName,
    siigoPrefix: prefix,
    siigoConsecutive: consecutive,
    agingBucket: agingBucket(doc.dueDate),
  };
}

export interface PayablesFilters {
  search?: string;
  supplierNit?: string;
  costCenterId?: number;
  bucket?: AgingBucket;
  dueTo?: string;
  includeScheduled?: boolean;
}

export interface PayablesSummary {
  total: number;
  overdue: number;
  due7: number;
  due30: number;
  scheduled: number;
  byBucket: Record<AgingBucket, { count: number; amount: number }>;
}

export async function listPayables(
  tenantId: string,
  filters: PayablesFilters = {}
): Promise<{ rows: PayableRow[]; summary: PayablesSummary }> {
  const q: Record<string, unknown> = {
    status: { $in: ["PAYABLE", "PAYMENT_SCHEDULED", "PARTIALLY_PAID"] },
    balance: { $gt: 0 },
  };
  if (filters.supplierNit) q.supplierNit = String(filters.supplierNit).replace(/\D/g, "");
  if (filters.costCenterId != null) q.costCenterId = Number(filters.costCenterId);
  if (filters.dueTo) q.dueDate = { $lte: filters.dueTo };
  if (filters.includeScheduled === false) q.status = { $in: ["PAYABLE", "PARTIALLY_PAID"] };
  if (filters.search) {
    const rx = new RegExp(String(filters.search).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i");
    q.$or = [{ supplierName: rx }, { supplierNit: rx }, { documentNumber: rx }, { "siigo.siigoName": rx }];
  }

  const docs = await col(COL.documents).find(scoped(tenantId, q)).sort({ dueDate: 1 }).limit(2000).toArray();
  let rows = fromDocs<ErpDocument>(docs).map(toPayableRow);
  if (filters.bucket) rows = rows.filter((r) => r.agingBucket === filters.bucket);

  const ref = today();
  const in7 = addDays(ref, 7);
  const in30 = addDays(ref, 30);

  const byBucket = Object.keys(AGING_LABELS).reduce((acc, k) => {
    acc[k as AgingBucket] = { count: 0, amount: 0 };
    return acc;
  }, {} as PayablesSummary["byBucket"]);
  for (const r of rows) {
    byBucket[r.agingBucket].count += 1;
    byBucket[r.agingBucket].amount = round2(byBucket[r.agingBucket].amount + r.balance);
  }

  const summary: PayablesSummary = {
    total: sum(rows.map((r) => r.balance)),
    overdue: sum(rows.filter((r) => r.dueDate && r.dueDate < ref).map((r) => r.balance)),
    due7: sum(rows.filter((r) => r.dueDate && r.dueDate >= ref && r.dueDate <= in7).map((r) => r.balance)),
    due30: sum(rows.filter((r) => r.dueDate && r.dueDate >= ref && r.dueDate <= in30).map((r) => r.balance)),
    scheduled: sum(rows.filter((r) => r.status === "PAYMENT_SCHEDULED").map((r) => r.balance)),
    byBucket,
  };

  return { rows, summary };
}

function addDays(dateStr: string, days: number): string {
  const d = new Date(`${dateStr}T12:00:00Z`);
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

// ─── Cuentas bancarias ───────────────────────────────────────────────────

export async function listBankAccounts(tenantId: string): Promise<ErpBankAccount[]> {
  const docs = await col(COL.bankAccounts).find(scoped(tenantId)).sort({ bank: 1 }).toArray();
  return fromDocs<ErpBankAccount>(docs);
}

export function maskAccountNumber(value: string): string {
  const digits = String(value || "").replace(/\D/g, "");
  return digits.length <= 4 ? digits : `••••${digits.slice(-4)}`;
}

export async function createBankAccount(tenantId: string, input: Partial<ErpBankAccount> & { accountNumber?: string }): Promise<ErpBankAccount> {
  const doc = {
    tenantId,
    bank: String(input.bank || "").trim(),
    accountName: String(input.accountName || "").trim(),
    accountType: (input.accountType || "CORRIENTE") as ErpBankAccount["accountType"],
    // Solo se guarda enmascarado: el ERP no necesita el número completo.
    maskedAccountNumber: input.maskedAccountNumber || maskAccountNumber(input.accountNumber || ""),
    siigoAccountCode: input.siigoAccountCode,
    siigoPaymentTypeId: input.siigoPaymentTypeId,
    currency: input.currency || "COP",
    active: input.active !== false,
    createdAt: now(),
    updatedAt: now(),
  };
  const res = await col(COL.bankAccounts).insertOne(doc as any);
  return { id: String(res.insertedId), ...doc };
}

export async function updateBankAccount(tenantId: string, id: string, patch: Partial<ErpBankAccount>): Promise<ErpBankAccount | null> {
  const set: Record<string, unknown> = { updatedAt: now() };
  for (const key of ["bank", "accountName", "accountType", "maskedAccountNumber", "siigoAccountCode", "siigoPaymentTypeId", "currency", "active"] as const) {
    if (patch[key] !== undefined) set[key] = patch[key];
  }
  const res = await col(COL.bankAccounts).findOneAndUpdate(
    scoped(tenantId, { _id: toId(id) }) as any,
    { $set: set },
    { returnDocument: "after" }
  );
  return res ? fromDoc<ErpBankAccount>(res) : null;
}

// ─── Programación de pagos ───────────────────────────────────────────────

export interface ScheduleInput {
  scheduledDate: string;
  bankAccountId?: string;
  paymentTypeId?: number;
  observations?: string;
  allocations: Array<{ documentId: string; value?: number }>;
}

async function resolveAllocations(
  tenantId: string,
  input: Array<{ documentId: string; value?: number }>
): Promise<{ allocations: ErpPaymentAllocation[]; documents: ErpDocument[] }> {
  const allocations: ErpPaymentAllocation[] = [];
  const documents: ErpDocument[] = [];

  for (const item of input) {
    const doc = await getDocument(tenantId, item.documentId);
    if (!doc) throw new ErpHttpError(`Documento ${item.documentId} no encontrado.`, 404, "document_not_found");
    if (doc.balance <= 0) {
      throw new ErpHttpError(`${documentLabel(doc)} ya no tiene saldo pendiente.`, 409, "no_balance");
    }
    const value = round2(item.value != null ? Number(item.value) : doc.balance);
    if (value <= 0) throw new ErpHttpError(`El valor imputado a ${documentLabel(doc)} debe ser mayor que cero.`, 400, "invalid_value");
    if (!gte(doc.balance, value)) {
      throw new ErpHttpError(
        `El valor imputado a ${documentLabel(doc)} (${value.toLocaleString("es-CO")}) supera su saldo (${doc.balance.toLocaleString("es-CO")}).`,
        400,
        "value_exceeds_balance"
      );
    }
    const { prefix, consecutive } = splitSiigoName(doc.siigo?.siigoName);
    allocations.push({
      documentId: doc.id,
      documentLabel: documentLabel(doc),
      supplierNit: doc.supplierNit,
      siigoPrefix: prefix,
      siigoConsecutive: consecutive,
      value,
    });
    documents.push(doc);
  }
  return { allocations, documents };
}

export async function createSchedule(ctx: ErpRequestContext, input: ScheduleInput): Promise<ErpPaymentSchedule> {
  if (!input.scheduledDate) throw new ErpHttpError("La fecha programada es obligatoria.", 400, "date_required");
  if (!input.allocations?.length) throw new ErpHttpError("Selecciona al menos una cuenta por pagar.", 400, "no_allocations");

  const { allocations, documents } = await resolveAllocations(ctx.tenant.id, input.allocations);
  const bank = input.bankAccountId
    ? (await col(COL.bankAccounts).findOne(scoped(ctx.tenant.id, { _id: toId(input.bankAccountId) })))
    : null;

  const doc = {
    tenantId: ctx.tenant.id,
    scheduledDate: input.scheduledDate.slice(0, 10),
    bankAccountId: input.bankAccountId,
    bankAccountLabel: bank ? `${bank.bank} ${bank.maskedAccountNumber}` : undefined,
    paymentTypeId: input.paymentTypeId,
    allocations,
    total: sum(allocations.map((a) => a.value)),
    status: "SCHEDULED" as const,
    observations: input.observations,
    createdById: ctx.actorId,
    createdAt: now(),
    updatedAt: now(),
  };
  const res = await col(COL.schedules).insertOne(doc as any);
  const schedule: ErpPaymentSchedule = { id: String(res.insertedId), ...doc };

  for (const d of documents) {
    if (d.status === "PAYABLE" || d.status === "PARTIALLY_PAID") {
      await transitionDocument(ctx, d.id, "SCHEDULE_PAYMENT", {
        system: true,
        extraPatch: { estimatedPaymentDate: schedule.scheduledDate },
      });
    }
  }

  await audit({
    tenantId: ctx.tenant.id,
    action: "PAYMENT_SCHEDULED",
    entityType: "paymentSchedule",
    entityId: schedule.id,
    actorId: ctx.actorId,
    actorName: ctx.actorName,
    after: { total: schedule.total, date: schedule.scheduledDate, documents: allocations.length },
    summary: `${allocations.length} documento(s) programados para el ${schedule.scheduledDate} por ${schedule.total.toLocaleString("es-CO")}.`,
  });

  await notify(ctx.tenant.id, {
    type: "PAYMENT_SCHEDULED",
    title: "Pago programado",
    body: `${allocations.length} documento(s) para el ${schedule.scheduledDate}`,
    entityType: "paymentSchedule",
    entityId: schedule.id,
    targetRole: "TREASURY",
  });

  return schedule;
}

export async function listSchedules(tenantId: string, opts: { status?: string; from?: string; to?: string } = {}): Promise<ErpPaymentSchedule[]> {
  const q: Record<string, unknown> = {};
  if (opts.status) q.status = opts.status;
  if (opts.from || opts.to) {
    q.scheduledDate = { ...(opts.from ? { $gte: opts.from } : {}), ...(opts.to ? { $lte: opts.to } : {}) };
  }
  const docs = await col(COL.schedules).find(scoped(tenantId, q)).sort({ scheduledDate: 1 }).limit(500).toArray();
  return fromDocs<ErpPaymentSchedule>(docs);
}

export async function cancelSchedule(ctx: ErpRequestContext, scheduleId: string): Promise<ErpPaymentSchedule> {
  const raw = await col(COL.schedules).findOne(scoped(ctx.tenant.id, { _id: toId(scheduleId) }));
  if (!raw) throw new ErpHttpError("Programación no encontrada.", 404, "schedule_not_found");
  const schedule = fromDoc<ErpPaymentSchedule>(raw);
  if (schedule.status === "EXECUTED") {
    throw new ErpHttpError("Esta programación ya se ejecutó; no puede cancelarse.", 409, "already_executed");
  }

  await col(COL.schedules).updateOne(
    scoped(ctx.tenant.id, { _id: toId(scheduleId) }) as any,
    { $set: { status: "CANCELLED", updatedAt: now() } }
  );

  for (const a of schedule.allocations) {
    const doc = await getDocument(ctx.tenant.id, a.documentId);
    if (doc?.status === "PAYMENT_SCHEDULED") {
      await transitionDocument(ctx, doc.id, "UNSCHEDULE_PAYMENT", { system: true });
    }
  }

  await audit({
    tenantId: ctx.tenant.id,
    action: "PAYMENT_SCHEDULE_CANCELLED",
    entityType: "paymentSchedule",
    entityId: scheduleId,
    actorId: ctx.actorId,
    actorName: ctx.actorName,
    summary: `Programación del ${schedule.scheduledDate} cancelada.`,
  });

  return { ...schedule, status: "CANCELLED" };
}

// ─── Pago efectivo ───────────────────────────────────────────────────────

export interface PaymentInput {
  paidAt: string;
  bankAccountId?: string;
  paymentTypeId?: number;
  reference?: string;
  observations?: string;
  scheduleId?: string;
  allocations: Array<{ documentId: string; value?: number }>;
}

/**
 * Registra un pago realmente ejecutado. La fecha efectiva SIEMPRE la aporta el
 * usuario: nunca se infiere de la fecha programada (§36).
 */
export async function recordPayment(ctx: ErpRequestContext, input: PaymentInput): Promise<ErpPayment> {
  if (!input.paidAt) throw new ErpHttpError("La fecha efectiva del pago es obligatoria.", 400, "date_required");
  if (!input.allocations?.length) throw new ErpHttpError("Indica a qué documentos se imputa el pago.", 400, "no_allocations");

  const { allocations, documents } = await resolveAllocations(ctx.tenant.id, input.allocations);
  const bank = input.bankAccountId
    ? (await col(COL.bankAccounts).findOne(scoped(ctx.tenant.id, { _id: toId(input.bankAccountId) })))
    : null;

  const doc = {
    tenantId: ctx.tenant.id,
    scheduleId: input.scheduleId,
    paidAt: input.paidAt.slice(0, 10),
    value: sum(allocations.map((a) => a.value)),
    bankAccountId: input.bankAccountId,
    bankAccountLabel: bank ? `${bank.bank} ${bank.maskedAccountNumber}` : undefined,
    paymentTypeId: input.paymentTypeId ?? bank?.siigoPaymentTypeId,
    reference: input.reference,
    observations: input.observations,
    allocations,
    files: [],
    siigo: { status: "PENDING" as const, attempts: 0 },
    createdById: ctx.actorId,
    createdAt: now(),
    updatedAt: now(),
  };
  const res = await col(COL.payments).insertOne(doc as any);
  const payment: ErpPayment = { id: String(res.insertedId), ...doc };

  // Imputación: baja el saldo de cada documento y avanza su estado.
  for (const alloc of allocations) {
    const target = documents.find((d) => d.id === alloc.documentId)!;
    const newBalance = subtract(target.balance, alloc.value);
    await patchDocument(ctx.tenant.id, target.id, { balance: newBalance });
    await transitionDocument(ctx, target.id, newBalance <= 0 ? "FULL_PAYMENT" : "PARTIAL_PAYMENT", { system: true });
  }

  if (input.scheduleId) {
    await col(COL.schedules).updateOne(
      scoped(ctx.tenant.id, { _id: toId(input.scheduleId) }) as any,
      { $set: { status: "EXECUTED", updatedAt: now() } }
    );
  }

  await audit({
    tenantId: ctx.tenant.id,
    action: "PAYMENT_RECORDED",
    entityType: "payment",
    entityId: payment.id,
    actorId: ctx.actorId,
    actorName: ctx.actorName,
    after: { value: payment.value, paidAt: payment.paidAt, bank: payment.bankAccountLabel },
    summary: `Pago de ${payment.value.toLocaleString("es-CO")} registrado por ${ctx.actorName} (${allocations.length} documento(s)).`,
  });

  return payment;
}

export async function listPayments(
  tenantId: string,
  opts: { from?: string; to?: string; supplierNit?: string; bankAccountId?: string; limit?: number } = {}
): Promise<ErpPayment[]> {
  const q: Record<string, unknown> = {};
  if (opts.from || opts.to) {
    q.paidAt = { ...(opts.from ? { $gte: opts.from } : {}), ...(opts.to ? { $lte: opts.to } : {}) };
  }
  if (opts.bankAccountId) q.bankAccountId = opts.bankAccountId;
  if (opts.supplierNit) q["allocations.supplierNit"] = String(opts.supplierNit).replace(/\D/g, "");
  const docs = await col(COL.payments)
    .find(scoped(tenantId, q))
    .sort({ paidAt: -1 })
    .limit(Math.min(opts.limit ?? 500, 1000))
    .toArray();
  return fromDocs<ErpPayment>(docs);
}

export async function getPayment(tenantId: string, id: string): Promise<ErpPayment | null> {
  try {
    const doc = await col(COL.payments).findOne(scoped(tenantId, { _id: toId(id) }));
    return doc ? fromDoc<ErpPayment>(doc) : null;
  } catch {
    return null;
  }
}

// ─── Egreso en Siigo (recibo de pago) ────────────────────────────────────

export interface EgresoResult {
  payment: ErpPayment;
  ok: boolean;
  error?: FriendlyError;
}

/**
 * Crea el recibo de pago (egreso) en Siigo para un pago ya registrado.
 * Idempotente: si ya existe recibo asociado no vuelve a enviarlo, y ante
 * incertidumbre busca el egreso en Siigo antes de reintentar.
 */
export async function postPaymentToSiigo(ctx: ErpRequestContext, paymentId: string): Promise<EgresoResult> {
  const tenant = ctx.tenant;
  const payment = await getPayment(tenant.id, paymentId);
  if (!payment) throw new ErpHttpError("Pago no encontrado.", 404, "payment_not_found");

  if (payment.siigo?.status === "SUCCESS" && payment.siigo.siigoId) {
    return { payment, ok: true };
  }

  const documentTypeId = tenant.settings?.defaultPaymentReceiptDocumentTypeId;
  if (!documentTypeId) {
    throw new ErpHttpError(
      "Falta configurar el tipo de comprobante de recibo de pago (RP) en Configuración → Contabilidad.",
      422,
      "missing_receipt_document_type"
    );
  }

  // Un recibo de pago en Siigo es por tercero: se agrupa por NIT.
  const bySupplier = new Map<string, ErpPaymentAllocation[]>();
  for (const a of payment.allocations) {
    const list = bySupplier.get(a.supplierNit) || [];
    list.push(a);
    bySupplier.set(a.supplierNit, list);
  }
  if (bySupplier.size > 1) {
    throw new ErpHttpError(
      "Este pago cubre documentos de varios terceros. Registra un egreso por tercero para que Siigo lo acepte.",
      422,
      "multiple_suppliers"
    );
  }

  const [supplierNit, allocations] = [...bySupplier.entries()][0];
  const missingDue = allocations.filter((a) => !a.siigoPrefix || a.siigoConsecutive == null);
  if (missingDue.length) {
    throw new ErpHttpError(
      `No podemos cruzar ${missingDue.map((m) => m.documentLabel).join(", ")} con Siigo: falta el comprobante de causación.`,
      422,
      "missing_siigo_document"
    );
  }

  const items: EgresoItem[] = allocations.map((a) => ({
    due: { prefix: a.siigoPrefix!, consecutive: a.siigoConsecutive!, quote: 1 },
    value: a.value,
    description: a.documentLabel,
  }));

  const egreso: EgresoPayload = {
    document: { id: Number(documentTypeId) },
    type: "DebtPayment",
    date: payment.paidAt,
    supplier: { identification: supplierNit, branch_office: 0 },
    items,
    payment: {
      id: Number(payment.paymentTypeId || tenant.settings?.defaultPaymentTypeId || 0),
      value: payment.value,
    },
    observations: [payment.reference, payment.observations].filter(Boolean).join(" · ").slice(0, 300) || undefined,
  };

  const errors = validateEgresoPayload(egreso);
  if (errors.length) {
    throw new ErpHttpError("El egreso no está listo para enviarse a Siigo.", 422, "validation_failed", { issues: errors });
  }

  // Reserva PROCESSING para evitar dobles envíos.
  const reserved = await col(COL.payments).findOneAndUpdate(
    scoped(tenant.id, { _id: toId(paymentId), "siigo.status": { $nin: ["PROCESSING", "SUCCESS"] } }) as any,
    {
      $set: {
        "siigo.status": "PROCESSING",
        "siigo.postingKey": `${tenant.id}:${paymentId}:PAYMENT`,
        "siigo.lastAttemptAt": now(),
        updatedAt: now(),
      },
      $inc: { "siigo.attempts": 1 },
    },
    { returnDocument: "after" }
  );
  if (!reserved) throw new ErpHttpError("Este egreso ya se está enviando a Siigo.", 409, "posting_in_progress");

  const opId = await startOperation({
    tenantId: tenant.id,
    integration: "SIIGO",
    operation: "createPaymentReceipt",
    entityType: "payment",
    entityId: paymentId,
    attempt: (payment.siigo?.attempts || 0) + 1,
    requestSummary: egreso,
  });

  try {
    const result: any = await submitEgreso(egreso);
    const siigoId = String(result?.id ?? "");
    if (!siigoId) throw new SiigoError("Siigo respondió sin identificador de recibo.", 502, "no_receipt_id", result);

    await finishOperation(opId, { status: "SUCCESS", externalId: siigoId, responseSummary: result });
    const updated = await applyEgresoSuccess(ctx, paymentId, result);
    return { payment: updated, ok: true };
  } catch (err) {
    const friendly = translateSiigoError(err);
    const uncertain = friendly.action === "RECONCILE";

    // Ante incertidumbre, buscar si el egreso quedó creado antes de reintentar.
    if (uncertain) {
      try {
        const matches = await findExistingEgresoInSiigo(payment.value, payment.paidAt);
        const existing = matches.find(
          (m) => String(m.thirdParty ?? "").replace(/\D/g, "") === supplierNit
        ) || matches[0];
        if (existing) {
          await finishOperation(opId, { status: "SUCCESS", externalId: String((existing as any).id || "") });
          const updated = await applyEgresoSuccess(ctx, paymentId, existing);
          return { payment: updated, ok: true };
        }
      } catch { /* si la reconciliación falla, se marca UNKNOWN */ }
    }

    await finishOperation(opId, {
      status: uncertain ? "UNKNOWN" : "FAILED",
      errorCode: friendly.code,
      errorMessage: friendly.message,
    });

    await col(COL.payments).updateOne(
      scoped(tenant.id, { _id: toId(paymentId) }) as any,
      {
        $set: {
          "siigo.status": uncertain ? "UNKNOWN" : "FAILED",
          "siigo.errorCode": friendly.code,
          "siigo.friendlyError": friendly.message,
          "siigo.errorMessage": err instanceof Error ? err.message : String(err),
          updatedAt: now(),
        },
      }
    );

    for (const a of payment.allocations) {
      const target = await getDocument(tenant.id, a.documentId);
      if (target?.status === "PAID") {
        await transitionDocument(ctx, target.id, "PAYMENT_POSTING_PENDING", { system: true });
      }
    }

    await audit({
      tenantId: tenant.id,
      action: "PAYMENT_POSTING_FAILED",
      entityType: "payment",
      entityId: paymentId,
      actorId: ctx.actorId,
      actorName: ctx.actorName,
      summary: `Error creando el egreso en Siigo: ${friendly.message}`,
    });

    const refreshed = (await getPayment(tenant.id, paymentId))!;
    return { payment: refreshed, ok: false, error: friendly };
  }
}

async function applyEgresoSuccess(ctx: ErpRequestContext, paymentId: string, result: any): Promise<ErpPayment> {
  const tenant = ctx.tenant;
  await col(COL.payments).updateOne(
    scoped(tenant.id, { _id: toId(paymentId) }) as any,
    {
      $set: {
        "siigo.status": "SUCCESS",
        "siigo.siigoId": String(result?.id ?? ""),
        "siigo.siigoName": String(result?.name ?? ""),
        "siigo.siigoNumber": Number(result?.number ?? 0) || undefined,
        "siigo.postedAt": now(),
        "siigo.errorCode": undefined,
        "siigo.friendlyError": undefined,
        updatedAt: now(),
      },
    }
  );

  const payment = (await getPayment(tenant.id, paymentId))!;

  for (const a of payment.allocations) {
    const doc = await getDocument(tenant.id, a.documentId);
    if (!doc) continue;
    if (doc.balance <= 0 && ["PAID", "PAYMENT_POSTING_PENDING"].includes(doc.status)) {
      await transitionDocument(ctx, doc.id, "PAYMENT_POSTED", { system: true });
    }
  }

  await audit({
    tenantId: tenant.id,
    action: "PAYMENT_POSTED",
    entityType: "payment",
    entityId: paymentId,
    actorId: ctx.actorId,
    actorName: ctx.actorName,
    after: { siigoName: payment.siigo.siigoName },
    summary: `Egreso creado en Siigo. ${payment.siigo.siigoName || ""}`,
  });

  return payment;
}
