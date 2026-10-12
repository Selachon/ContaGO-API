/**
 * Agregados del dashboard. Todas las cifras salen de queries agregadas en
 * MongoDB — nunca se traen miles de documentos al navegador.
 */
import { COL, col, fromDocs, scoped, today } from "./db.js";
import { round2, sum } from "./money.js";
import { documentLabel } from "./documents.js";
import { listPayables } from "./payments.js";
import { cashFlowOutlook } from "./cashflow.js";
import { listAudit } from "./audit.js";
import type { ErpDocument, ErpPayment } from "./types.js";
import { STATUS_LABELS } from "./workflow.js";

function addDays(dateStr: string, days: number): string {
  const d = new Date(`${dateStr}T12:00:00Z`);
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

function monthStart(offset = 0): string {
  const d = new Date();
  d.setUTCDate(1);
  d.setUTCMonth(d.getUTCMonth() + offset);
  return d.toISOString().slice(0, 10);
}

export interface AttentionItem {
  kind: "APPROVAL" | "ACCOUNTING" | "SIIGO_ERROR" | "DUE_SOON" | "OVERDUE" | "DRIVE" | "PAYMENT_POSTING";
  severity: "info" | "warning" | "danger";
  title: string;
  detail?: string;
  count?: number;
  href?: string;
}

export interface DashboardData {
  cards: {
    pendingDocuments: number;
    pendingApproval: number;
    pendingAccounting: number;
    siigoErrors: number;
    payableTotal: number;
    payableOverdue: number;
    dueNext7: number;
    paidThisMonth: number;
  };
  attention: AttentionItem[];
  byStatus: Array<{ status: string; label: string; count: number; amount: number }>;
  paymentsByMonth: Array<{ month: string; amount: number }>;
  expensesBySupplier: Array<{ supplierNit: string; supplierName: string; amount: number }>;
  expensesByCostCenter: Array<{ costCenterId: number | null; label: string; amount: number }>;
  upcoming: Array<{ documentId: string; label: string; supplierName: string; date: string; value: number; overdue: boolean }>;
  cashflow: { series: Array<{ date: string; projected: number; actual: number }>; projected: number; actual: number };
  recentActivity: Array<{ id: string; action: string; summary?: string; actorName?: string; createdAt: string }>;
}

export async function getDashboard(tenantId: string): Promise<DashboardData> {
  const ref = today();
  const in7 = addDays(ref, 7);
  const thisMonth = monthStart(0);

  const documents = col(COL.documents);

  const [
    pendingDocuments,
    pendingApproval,
    pendingAccounting,
    siigoErrors,
    payables,
    statusAgg,
    paymentsAgg,
    supplierAgg,
    costCenterAgg,
    paidMonthAgg,
    outlook,
    activity,
    upcomingDocs,
  ] = await Promise.all([
    documents.countDocuments(scoped(tenantId, { status: { $in: ["RECEIVED", "PENDING_REVIEW", "PENDING_APPROVAL"] } })),
    documents.countDocuments(scoped(tenantId, { status: "PENDING_APPROVAL" })),
    documents.countDocuments(scoped(tenantId, { status: { $in: ["PENDING_ACCOUNTING", "ACCOUNTING_READY"] } })),
    documents.countDocuments(scoped(tenantId, { status: "POSTING_ERROR" })),
    listPayables(tenantId),
    documents
      .aggregate([
        { $match: scoped(tenantId) },
        { $group: { _id: "$status", count: { $sum: 1 }, amount: { $sum: "$total" } } },
      ])
      .toArray(),
    col(COL.payments)
      .aggregate([
        { $match: scoped(tenantId) },
        { $group: { _id: { $substr: ["$paidAt", 0, 7] }, amount: { $sum: "$value" } } },
        { $sort: { _id: 1 } },
        { $limit: 24 },
      ])
      .toArray(),
    documents
      .aggregate([
        { $match: scoped(tenantId, { status: { $nin: ["VOIDED", "REJECTED"] } }) },
        { $group: { _id: { nit: "$supplierNit", name: "$supplierName" }, amount: { $sum: "$total" } } },
        { $sort: { amount: -1 } },
        { $limit: 10 },
      ])
      .toArray(),
    documents
      .aggregate([
        { $match: scoped(tenantId, { status: { $nin: ["VOIDED", "REJECTED"] } }) },
        { $group: { _id: "$costCenterId", amount: { $sum: "$total" } } },
        { $sort: { amount: -1 } },
        { $limit: 10 },
      ])
      .toArray(),
    col(COL.payments)
      .aggregate([
        { $match: scoped(tenantId, { paidAt: { $gte: thisMonth } }) },
        { $group: { _id: null, amount: { $sum: "$value" } } },
      ])
      .toArray(),
    cashFlowOutlook(tenantId, 30),
    listAudit(tenantId, { limit: 12 }),
    documents
      .find(scoped(tenantId, { status: { $in: ["PAYABLE", "PAYMENT_SCHEDULED", "PARTIALLY_PAID"] }, balance: { $gt: 0 } }))
      .sort({ dueDate: 1 })
      .limit(8)
      .toArray(),
  ]);

  const attention: AttentionItem[] = [];
  if (pendingApproval) {
    attention.push({
      kind: "APPROVAL",
      severity: "warning",
      title: `${pendingApproval} documento${pendingApproval > 1 ? "s" : ""} pendiente${pendingApproval > 1 ? "s" : ""} de aprobación`,
      count: pendingApproval,
      href: "/aprobaciones",
    });
  }
  if (pendingAccounting) {
    attention.push({
      kind: "ACCOUNTING",
      severity: "info",
      title: `${pendingAccounting} documento${pendingAccounting > 1 ? "s" : ""} por causar`,
      count: pendingAccounting,
      href: "/contabilidad/por-causar",
    });
  }
  if (siigoErrors) {
    attention.push({
      kind: "SIIGO_ERROR",
      severity: "danger",
      title: `${siigoErrors} causación${siigoErrors > 1 ? "es" : ""} falló en Siigo`,
      count: siigoErrors,
      href: "/contabilidad/errores",
    });
  }

  const dueSoon = payables.rows.filter((r) => r.dueDate && r.dueDate >= ref && r.dueDate <= in7);
  if (dueSoon.length) {
    attention.push({
      kind: "DUE_SOON",
      severity: "warning",
      title: `${dueSoon.length} factura${dueSoon.length > 1 ? "s" : ""} vence${dueSoon.length > 1 ? "n" : ""} en los próximos 7 días`,
      detail: `${round2(sum(dueSoon.map((d) => d.balance))).toLocaleString("es-CO")} COP`,
      count: dueSoon.length,
      href: "/cuentas-por-pagar",
    });
  }
  const overdue = payables.rows.filter((r) => r.dueDate && r.dueDate < ref);
  if (overdue.length) {
    attention.push({
      kind: "OVERDUE",
      severity: "danger",
      title: `${overdue.length} factura${overdue.length > 1 ? "s" : ""} vencida${overdue.length > 1 ? "s" : ""}`,
      detail: `${round2(sum(overdue.map((d) => d.balance))).toLocaleString("es-CO")} COP`,
      count: overdue.length,
      href: "/cuentas-por-pagar?bucket=overdue",
    });
  }

  const pendingEgresos = await col(COL.payments).countDocuments(
    scoped(tenantId, { "siigo.status": { $in: ["PENDING", "FAILED", "UNKNOWN"] } })
  );
  if (pendingEgresos) {
    attention.push({
      kind: "PAYMENT_POSTING",
      severity: "warning",
      title: `${pendingEgresos} pago${pendingEgresos > 1 ? "s" : ""} sin egreso en Siigo`,
      count: pendingEgresos,
      href: "/pagos",
    });
  }

  const upcoming = fromDocs<ErpDocument>(upcomingDocs).map((d) => ({
    documentId: d.id,
    label: documentLabel(d),
    supplierName: d.supplierName,
    date: d.dueDate || d.estimatedPaymentDate || d.issueDate,
    value: d.balance,
    overdue: !!(d.dueDate && d.dueDate < ref),
  }));

  return {
    cards: {
      pendingDocuments,
      pendingApproval,
      pendingAccounting,
      siigoErrors,
      payableTotal: payables.summary.total,
      payableOverdue: payables.summary.overdue,
      dueNext7: payables.summary.due7,
      paidThisMonth: round2(paidMonthAgg[0]?.amount || 0),
    },
    attention,
    byStatus: statusAgg.map((s: any) => ({
      status: s._id,
      label: STATUS_LABELS[s._id as keyof typeof STATUS_LABELS] || s._id,
      count: s.count,
      amount: round2(s.amount || 0),
    })),
    paymentsByMonth: paymentsAgg.map((p: any) => ({ month: p._id, amount: round2(p.amount || 0) })),
    expensesBySupplier: supplierAgg.map((s: any) => ({
      supplierNit: s._id?.nit || "",
      supplierName: s._id?.name || "Sin nombre",
      amount: round2(s.amount || 0),
    })),
    expensesByCostCenter: costCenterAgg.map((c: any) => ({
      costCenterId: c._id ?? null,
      label: c._id != null ? `Centro ${c._id}` : "Sin centro",
      amount: round2(c.amount || 0),
    })),
    upcoming,
    cashflow: {
      series: outlook.series,
      projected: outlook.totals.projected,
      actual: outlook.totals.actual,
    },
    recentActivity: activity.map((a) => ({
      id: a.id,
      action: a.action,
      summary: a.summary,
      actorName: a.actorName,
      createdAt: a.createdAt,
    })),
  };
}

/** Vista específica del equipo contable (§42). */
export async function getAccountingDashboard(tenantId: string) {
  const ref = today();
  const documents = col(COL.documents);
  const [toReview, toPost, withError, postedToday, postedMonth, queue] = await Promise.all([
    documents.countDocuments(scoped(tenantId, { status: { $in: ["RECEIVED", "PENDING_REVIEW"] } })),
    documents.countDocuments(scoped(tenantId, { status: { $in: ["PENDING_ACCOUNTING", "ACCOUNTING_READY"] } })),
    documents.countDocuments(scoped(tenantId, { status: "POSTING_ERROR" })),
    documents.countDocuments(scoped(tenantId, { "siigo.postedAt": { $gte: `${ref}T00:00:00.000Z` } })),
    documents.countDocuments(scoped(tenantId, { "siigo.postedAt": { $gte: `${monthStart(0)}T00:00:00.000Z` } })),
    documents
      .find(scoped(tenantId, { status: { $in: ["PENDING_ACCOUNTING", "ACCOUNTING_READY"] } }))
      .sort({ issueDate: 1 })
      .limit(50)
      .toArray(),
  ]);

  const rows = fromDocs<ErpDocument>(queue).map((d) => {
    const days = Math.max(0, Math.round((Date.parse(`${ref}T00:00:00Z`) - Date.parse(`${d.issueDate}T00:00:00Z`)) / 86_400_000));
    return {
      documentId: d.id,
      label: documentLabel(d),
      supplierName: d.supplierName,
      supplierNit: d.supplierNit,
      total: d.total,
      daysPending: days,
      priority: days > 15 ? "ALTA" : days > 7 ? "MEDIA" : "NORMAL",
      status: d.status,
      hasAccounting: !!d.accounting?.accountCode,
      assignedToId: d.assignedToId,
    };
  });

  return { cards: { toReview, toPost, withError, postedToday, postedMonth }, queue: rows };
}

/** Vista específica de tesorería (§43). */
export async function getTreasuryDashboard(tenantId: string) {
  const ref = today();
  const in7 = addDays(ref, 7);
  const [payables, schedulesToday, paidToday, payments] = await Promise.all([
    listPayables(tenantId),
    col(COL.schedules).find(scoped(tenantId, { status: "SCHEDULED", scheduledDate: ref })).toArray(),
    col(COL.payments)
      .aggregate([{ $match: scoped(tenantId, { paidAt: ref }) }, { $group: { _id: null, amount: { $sum: "$value" } } }])
      .toArray(),
    col(COL.payments).find(scoped(tenantId)).sort({ paidAt: -1 }).limit(20).toArray(),
  ]);

  const upcomingSchedules = await col(COL.schedules)
    .find(scoped(tenantId, { status: "SCHEDULED", scheduledDate: { $gte: ref, $lte: addDays(ref, 45) } }))
    .sort({ scheduledDate: 1 })
    .limit(100)
    .toArray();

  return {
    cards: {
      payableTotal: payables.summary.total,
      overdue: payables.summary.overdue,
      dueThisWeek: payables.summary.due7,
      scheduledToday: sum(fromDocs<any>(schedulesToday).map((s) => s.total)),
      paidToday: round2(paidToday[0]?.amount || 0),
    },
    calendar: fromDocs<any>(upcomingSchedules).map((s) => ({
      id: s.id,
      date: s.scheduledDate,
      total: s.total,
      count: s.allocations?.length || 0,
      bankAccountLabel: s.bankAccountLabel,
      status: s.status,
    })),
    aging: payables.summary.byBucket,
    recentPayments: fromDocs<ErpPayment>(payments).map((p) => ({
      id: p.id,
      paidAt: p.paidAt,
      value: p.value,
      bankAccountLabel: p.bankAccountLabel,
      siigoStatus: p.siigo?.status,
      siigoName: p.siigo?.siigoName,
      documents: p.allocations.length,
    })),
  };
}
