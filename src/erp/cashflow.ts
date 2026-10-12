/**
 * Flujo de caja. Distingue explícitamente PROYECTADO de REAL (§39):
 *   - Proyectado = programaciones de pago + fecha estimada de las CxP abiertas.
 *   - Real       = pagos efectivamente registrados.
 * Los dos conceptos nunca se suman en la misma serie.
 */
import { COL, col, fromDocs, scoped, today } from "./db.js";
import { round2, sum } from "./money.js";
import { documentLabel } from "./documents.js";
import type { ErpDocument, ErpPayment, ErpPaymentSchedule } from "./types.js";

export type CashFlowKind = "PROJECTED" | "ACTUAL";

export interface CashFlowRow {
  id: string;
  kind: CashFlowKind;
  date: string;
  supplierNit: string;
  supplierName: string;
  concept: string;
  documentId?: string;
  documentLabel?: string;
  costCenterId?: number;
  subCostCenterId?: string;
  bankAccountId?: string;
  bankAccountLabel?: string;
  value: number;
  origin: "PAYMENT" | "SCHEDULE" | "PAYABLE";
  siigoName?: string;
}

export interface CashFlowFilters {
  from?: string;
  to?: string;
  costCenterId?: number;
  subCostCenterId?: string;
  bankAccountId?: string;
  supplierNit?: string;
  kind?: CashFlowKind;
}

export interface CashFlowResult {
  rows: CashFlowRow[];
  totals: { projected: number; actual: number };
  /** Serie diaria para la gráfica: proyectado y real por fecha. */
  series: Array<{ date: string; projected: number; actual: number }>;
  byCostCenter: Array<{ costCenterId: number | null; label: string; projected: number; actual: number }>;
  byBank: Array<{ bankAccountId: string | null; label: string; actual: number }>;
  bySupplier: Array<{ supplierNit: string; supplierName: string; actual: number; projected: number }>;
}

function inRange(date: string, from?: string, to?: string): boolean {
  if (from && date < from) return false;
  if (to && date > to) return false;
  return true;
}

export async function getCashFlow(tenantId: string, f: CashFlowFilters = {}): Promise<CashFlowResult> {
  const rows: CashFlowRow[] = [];

  // ── REAL: pagos ejecutados ────────────────────────────────────────────
  if (f.kind !== "PROJECTED") {
    const q: Record<string, unknown> = {};
    if (f.from || f.to) q.paidAt = { ...(f.from ? { $gte: f.from } : {}), ...(f.to ? { $lte: f.to } : {}) };
    if (f.bankAccountId) q.bankAccountId = f.bankAccountId;
    const payments = fromDocs<ErpPayment>(
      await col(COL.payments).find(scoped(tenantId, q)).sort({ paidAt: -1 }).limit(3000).toArray()
    );

    // Se necesita el centro de costo del documento imputado.
    const docIds = new Set<string>();
    for (const p of payments) for (const a of p.allocations) docIds.add(a.documentId);
    const docMap = await loadDocuments(tenantId, [...docIds]);

    for (const p of payments) {
      for (const a of p.allocations) {
        const doc = docMap.get(a.documentId);
        if (f.supplierNit && a.supplierNit !== String(f.supplierNit).replace(/\D/g, "")) continue;
        const ccId = doc?.costCenterId ?? doc?.accounting?.costCenterId;
        if (f.costCenterId != null && Number(ccId) !== Number(f.costCenterId)) continue;
        if (f.subCostCenterId && (doc?.subCostCenterId ?? doc?.accounting?.subCostCenterId) !== f.subCostCenterId) continue;

        rows.push({
          id: `${p.id}:${a.documentId}`,
          kind: "ACTUAL",
          date: p.paidAt,
          supplierNit: a.supplierNit,
          supplierName: doc?.supplierName || a.supplierNit,
          concept: doc?.description || a.documentLabel,
          documentId: a.documentId,
          documentLabel: a.documentLabel,
          costCenterId: ccId,
          subCostCenterId: doc?.subCostCenterId ?? doc?.accounting?.subCostCenterId,
          bankAccountId: p.bankAccountId,
          bankAccountLabel: p.bankAccountLabel,
          value: a.value,
          origin: "PAYMENT",
          siigoName: p.siigo?.siigoName,
        });
      }
    }
  }

  // ── PROYECTADO: programaciones + CxP con fecha estimada ───────────────
  if (f.kind !== "ACTUAL") {
    const schedules = fromDocs<ErpPaymentSchedule>(
      await col(COL.schedules).find(scoped(tenantId, { status: "SCHEDULED" })).limit(2000).toArray()
    );
    const scheduledDocIds = new Set<string>();
    for (const s of schedules) for (const a of s.allocations) scheduledDocIds.add(a.documentId);
    const scheduleDocs = await loadDocuments(tenantId, [...scheduledDocIds]);

    for (const s of schedules) {
      if (!inRange(s.scheduledDate, f.from, f.to)) continue;
      for (const a of s.allocations) {
        const doc = scheduleDocs.get(a.documentId);
        if (f.supplierNit && a.supplierNit !== String(f.supplierNit).replace(/\D/g, "")) continue;
        const ccId = doc?.costCenterId ?? doc?.accounting?.costCenterId;
        if (f.costCenterId != null && Number(ccId) !== Number(f.costCenterId)) continue;
        rows.push({
          id: `sch:${s.id}:${a.documentId}`,
          kind: "PROJECTED",
          date: s.scheduledDate,
          supplierNit: a.supplierNit,
          supplierName: doc?.supplierName || a.supplierNit,
          concept: doc?.description || a.documentLabel,
          documentId: a.documentId,
          documentLabel: a.documentLabel,
          costCenterId: ccId,
          subCostCenterId: doc?.subCostCenterId,
          bankAccountId: s.bankAccountId,
          bankAccountLabel: s.bankAccountLabel,
          value: a.value,
          origin: "SCHEDULE",
        });
      }
    }

    // CxP abiertas sin programación: se proyectan por fecha estimada de pago.
    const openQ: Record<string, unknown> = {
      status: { $in: ["PAYABLE", "PARTIALLY_PAID"] },
      balance: { $gt: 0 },
    };
    if (f.supplierNit) openQ.supplierNit = String(f.supplierNit).replace(/\D/g, "");
    if (f.costCenterId != null) openQ.costCenterId = Number(f.costCenterId);
    const open = fromDocs<ErpDocument>(
      await col(COL.documents).find(scoped(tenantId, openQ)).limit(3000).toArray()
    );
    for (const d of open) {
      const date = d.estimatedPaymentDate || d.dueDate;
      if (!date || !inRange(date, f.from, f.to)) continue;
      rows.push({
        id: `pay:${d.id}`,
        kind: "PROJECTED",
        date,
        supplierNit: d.supplierNit,
        supplierName: d.supplierName,
        concept: d.description || documentLabel(d),
        documentId: d.id,
        documentLabel: documentLabel(d),
        costCenterId: d.costCenterId ?? d.accounting?.costCenterId,
        subCostCenterId: d.subCostCenterId,
        value: d.balance,
        origin: "PAYABLE",
      });
    }
  }

  rows.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));

  const projectedRows = rows.filter((r) => r.kind === "PROJECTED");
  const actualRows = rows.filter((r) => r.kind === "ACTUAL");

  // Serie diaria.
  const byDate = new Map<string, { projected: number; actual: number }>();
  for (const r of rows) {
    const entry = byDate.get(r.date) || { projected: 0, actual: 0 };
    if (r.kind === "PROJECTED") entry.projected = round2(entry.projected + r.value);
    else entry.actual = round2(entry.actual + r.value);
    byDate.set(r.date, entry);
  }

  const byCC = new Map<string, { costCenterId: number | null; label: string; projected: number; actual: number }>();
  for (const r of rows) {
    const key = String(r.costCenterId ?? "none");
    const entry = byCC.get(key) || {
      costCenterId: r.costCenterId ?? null,
      label: r.costCenterId != null ? `Centro ${r.costCenterId}` : "Sin centro",
      projected: 0,
      actual: 0,
    };
    if (r.kind === "PROJECTED") entry.projected = round2(entry.projected + r.value);
    else entry.actual = round2(entry.actual + r.value);
    byCC.set(key, entry);
  }

  const byBankMap = new Map<string, { bankAccountId: string | null; label: string; actual: number }>();
  for (const r of actualRows) {
    const key = r.bankAccountId || "none";
    const entry = byBankMap.get(key) || {
      bankAccountId: r.bankAccountId ?? null,
      label: r.bankAccountLabel || "Sin cuenta",
      actual: 0,
    };
    entry.actual = round2(entry.actual + r.value);
    byBankMap.set(key, entry);
  }

  const bySupplierMap = new Map<string, { supplierNit: string; supplierName: string; actual: number; projected: number }>();
  for (const r of rows) {
    const entry = bySupplierMap.get(r.supplierNit) || {
      supplierNit: r.supplierNit,
      supplierName: r.supplierName,
      actual: 0,
      projected: 0,
    };
    if (r.kind === "PROJECTED") entry.projected = round2(entry.projected + r.value);
    else entry.actual = round2(entry.actual + r.value);
    bySupplierMap.set(r.supplierNit, entry);
  }

  return {
    rows,
    totals: { projected: sum(projectedRows.map((r) => r.value)), actual: sum(actualRows.map((r) => r.value)) },
    series: [...byDate.entries()]
      .map(([date, v]) => ({ date, ...v }))
      .sort((a, b) => (a.date < b.date ? -1 : 1)),
    byCostCenter: [...byCC.values()].sort((a, b) => b.actual + b.projected - (a.actual + a.projected)),
    byBank: [...byBankMap.values()].sort((a, b) => b.actual - a.actual),
    bySupplier: [...bySupplierMap.values()].sort((a, b) => b.actual + b.projected - (a.actual + a.projected)).slice(0, 25),
  };
}

async function loadDocuments(tenantId: string, ids: string[]): Promise<Map<string, ErpDocument>> {
  const map = new Map<string, ErpDocument>();
  if (!ids.length) return map;
  const { ObjectId } = await import("mongodb");
  const objectIds = ids.filter((id) => ObjectId.isValid(id)).map((id) => new ObjectId(id));
  const docs = fromDocs<ErpDocument>(
    await col(COL.documents).find(scoped(tenantId, { _id: { $in: objectIds } })).toArray()
  );
  for (const d of docs) map.set(d.id, d);
  return map;
}

/** Resumen para el widget del dashboard: próximos 30 días. */
export async function cashFlowOutlook(tenantId: string, days = 30) {
  const from = today();
  const to = new Date(Date.parse(`${from}T12:00:00Z`) + days * 86_400_000).toISOString().slice(0, 10);
  const flow = await getCashFlow(tenantId, { from, to });
  return { from, to, ...flow };
}
