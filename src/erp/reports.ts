/**
 * Reportes del ERP. Todos devuelven filas planas listas para exportar a
 * Excel/CSV desde el propio backend (nada de armar el archivo en el navegador).
 */
import ExcelJS from "exceljs";
import { COL, col, fromDocs, scoped } from "./db.js";
import { round2 } from "./money.js";
import { documentLabel } from "./documents.js";
import { listPayables } from "./payments.js";
import { getCashFlow } from "./cashflow.js";
import { STATUS_LABELS } from "./workflow.js";
import type { ErpDocument, ErpPayment } from "./types.js";

export type ReportId =
  | "documents-by-status"
  | "expenses-by-supplier"
  | "expenses-by-cost-center"
  | "payables"
  | "due-dates"
  | "payments"
  | "cashflow";

export const REPORTS: Array<{ id: ReportId; name: string; description: string }> = [
  { id: "documents-by-status", name: "Documentos por estado", description: "Conteo y valor de documentos agrupados por estado del flujo." },
  { id: "expenses-by-supplier", name: "Gastos por proveedor", description: "Total causado por tercero en el período." },
  { id: "expenses-by-cost-center", name: "Gastos por centro de costo", description: "Distribución del gasto por centro y subcentro." },
  { id: "payables", name: "Cuentas por pagar", description: "Saldos abiertos con antigüedad y vencimiento." },
  { id: "due-dates", name: "Vencimientos", description: "Documentos que vencen en el rango seleccionado." },
  { id: "payments", name: "Pagos realizados", description: "Pagos efectivos con imputación y estado del egreso en Siigo." },
  { id: "cashflow", name: "Flujo de caja", description: "Proyectado vs. real, por fecha y centro de costo." },
];

export interface ReportParams {
  from?: string;
  to?: string;
  costCenterId?: number;
  supplierNit?: string;
}

export interface ReportResult {
  id: ReportId;
  name: string;
  columns: Array<{ key: string; label: string; type?: "money" | "date" | "number" | "text" }>;
  rows: Array<Record<string, unknown>>;
  totals?: Record<string, number>;
}

function dateFilter(params: ReportParams, field = "issueDate"): Record<string, unknown> {
  if (!params.from && !params.to) return {};
  return {
    [field]: {
      ...(params.from ? { $gte: params.from } : {}),
      ...(params.to ? { $lte: params.to } : {}),
    },
  };
}

export async function runReport(tenantId: string, id: ReportId, params: ReportParams = {}): Promise<ReportResult> {
  const meta = REPORTS.find((r) => r.id === id);
  if (!meta) throw new Error(`Reporte desconocido: ${id}`);

  switch (id) {
    case "documents-by-status": {
      const agg = await col(COL.documents)
        .aggregate([
          { $match: scoped(tenantId, dateFilter(params)) },
          { $group: { _id: "$status", count: { $sum: 1 }, amount: { $sum: "$total" } } },
          { $sort: { amount: -1 } },
        ])
        .toArray();
      return {
        id, name: meta.name,
        columns: [
          { key: "label", label: "Estado" },
          { key: "count", label: "Documentos", type: "number" },
          { key: "amount", label: "Valor", type: "money" },
        ],
        rows: agg.map((a: any) => ({
          label: STATUS_LABELS[a._id as keyof typeof STATUS_LABELS] || a._id,
          count: a.count,
          amount: round2(a.amount || 0),
        })),
        totals: { amount: round2(agg.reduce((s: number, a: any) => s + (a.amount || 0), 0)) },
      };
    }

    case "expenses-by-supplier": {
      const agg = await col(COL.documents)
        .aggregate([
          { $match: scoped(tenantId, { ...dateFilter(params), status: { $nin: ["VOIDED", "REJECTED"] } }) },
          { $group: { _id: { nit: "$supplierNit", name: "$supplierName" }, count: { $sum: 1 }, amount: { $sum: "$total" } } },
          { $sort: { amount: -1 } },
          { $limit: 500 },
        ])
        .toArray();
      return {
        id, name: meta.name,
        columns: [
          { key: "supplierNit", label: "NIT" },
          { key: "supplierName", label: "Proveedor" },
          { key: "count", label: "Documentos", type: "number" },
          { key: "amount", label: "Total", type: "money" },
        ],
        rows: agg.map((a: any) => ({
          supplierNit: a._id?.nit, supplierName: a._id?.name, count: a.count, amount: round2(a.amount || 0),
        })),
        totals: { amount: round2(agg.reduce((s: number, a: any) => s + (a.amount || 0), 0)) },
      };
    }

    case "expenses-by-cost-center": {
      const agg = await col(COL.documents)
        .aggregate([
          { $match: scoped(tenantId, { ...dateFilter(params), status: { $nin: ["VOIDED", "REJECTED"] } }) },
          {
            $group: {
              _id: { cc: "$costCenterId", sub: "$subCostCenterId" },
              count: { $sum: 1 }, amount: { $sum: "$total" },
            },
          },
          { $sort: { amount: -1 } },
        ])
        .toArray();
      return {
        id, name: meta.name,
        columns: [
          { key: "costCenterId", label: "Centro de costo" },
          { key: "subCostCenterId", label: "Subcentro" },
          { key: "count", label: "Documentos", type: "number" },
          { key: "amount", label: "Total", type: "money" },
        ],
        rows: agg.map((a: any) => ({
          costCenterId: a._id?.cc ?? "Sin centro",
          subCostCenterId: a._id?.sub ?? "—",
          count: a.count,
          amount: round2(a.amount || 0),
        })),
        totals: { amount: round2(agg.reduce((s: number, a: any) => s + (a.amount || 0), 0)) },
      };
    }

    case "payables": {
      const { rows, summary } = await listPayables(tenantId, {
        supplierNit: params.supplierNit,
        costCenterId: params.costCenterId,
      });
      return {
        id, name: meta.name,
        columns: [
          { key: "supplierName", label: "Proveedor" },
          { key: "label", label: "Documento" },
          { key: "issueDate", label: "Fecha", type: "date" },
          { key: "dueDate", label: "Vencimiento", type: "date" },
          { key: "total", label: "Total", type: "money" },
          { key: "balance", label: "Saldo", type: "money" },
          { key: "daysToDue", label: "Días", type: "number" },
          { key: "agingBucket", label: "Antigüedad" },
        ],
        rows: rows as unknown as Array<Record<string, unknown>>,
        totals: { balance: summary.total, overdue: summary.overdue },
      };
    }

    case "due-dates": {
      const docs = fromDocs<ErpDocument>(
        await col(COL.documents)
          .find(scoped(tenantId, { ...dateFilter(params, "dueDate"), status: { $in: ["PAYABLE", "PAYMENT_SCHEDULED", "PARTIALLY_PAID"] } }))
          .sort({ dueDate: 1 })
          .limit(2000)
          .toArray()
      );
      return {
        id, name: meta.name,
        columns: [
          { key: "dueDate", label: "Vencimiento", type: "date" },
          { key: "supplierName", label: "Proveedor" },
          { key: "label", label: "Documento" },
          { key: "balance", label: "Saldo", type: "money" },
          { key: "estimatedPaymentDate", label: "Fecha estimada de pago", type: "date" },
        ],
        rows: docs.map((d) => ({
          dueDate: d.dueDate, supplierName: d.supplierName, label: documentLabel(d),
          balance: d.balance, estimatedPaymentDate: d.estimatedPaymentDate,
        })),
        totals: { balance: round2(docs.reduce((s, d) => s + d.balance, 0)) },
      };
    }

    case "payments": {
      const payments = fromDocs<ErpPayment>(
        await col(COL.payments)
          .find(scoped(tenantId, dateFilter(params, "paidAt")))
          .sort({ paidAt: -1 })
          .limit(2000)
          .toArray()
      );
      const rows = payments.flatMap((p) =>
        p.allocations.map((a) => ({
          paidAt: p.paidAt,
          supplierNit: a.supplierNit,
          documentLabel: a.documentLabel,
          value: a.value,
          bankAccountLabel: p.bankAccountLabel || "—",
          reference: p.reference || "—",
          siigoName: p.siigo?.siigoName || "—",
          siigoStatus: p.siigo?.status || "PENDING",
        }))
      );
      return {
        id, name: meta.name,
        columns: [
          { key: "paidAt", label: "Fecha", type: "date" },
          { key: "supplierNit", label: "NIT" },
          { key: "documentLabel", label: "Documento" },
          { key: "value", label: "Valor", type: "money" },
          { key: "bankAccountLabel", label: "Cuenta" },
          { key: "reference", label: "Referencia" },
          { key: "siigoName", label: "Egreso Siigo" },
          { key: "siigoStatus", label: "Estado egreso" },
        ],
        rows,
        totals: { value: round2(rows.reduce((s, r) => s + (r.value as number), 0)) },
      };
    }

    case "cashflow": {
      const flow = await getCashFlow(tenantId, {
        from: params.from, to: params.to,
        costCenterId: params.costCenterId, supplierNit: params.supplierNit,
      });
      return {
        id, name: meta.name,
        columns: [
          { key: "date", label: "Fecha", type: "date" },
          { key: "kind", label: "Tipo" },
          { key: "supplierName", label: "Proveedor" },
          { key: "concept", label: "Concepto" },
          { key: "costCenterId", label: "Centro" },
          { key: "bankAccountLabel", label: "Cuenta" },
          { key: "value", label: "Valor", type: "money" },
          { key: "origin", label: "Origen" },
        ],
        rows: flow.rows.map((r) => ({
          ...r,
          kind: r.kind === "ACTUAL" ? "Real" : "Proyectado",
          costCenterId: r.costCenterId ?? "—",
          bankAccountLabel: r.bankAccountLabel || "—",
        })),
        totals: { projected: flow.totals.projected, actual: flow.totals.actual },
      };
    }
  }
}

/** Exporta un reporte a Excel usando la misma librería del resto de ContaGO. */
export async function reportToExcel(result: ReportResult): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "ContaGO ERP";
  const ws = wb.addWorksheet(result.name.slice(0, 30));

  ws.columns = result.columns.map((c) => ({
    header: c.label,
    key: c.key,
    width: c.type === "money" ? 18 : c.type === "date" ? 14 : 28,
  }));
  ws.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
  ws.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF8227CC" } };
  ws.getRow(1).alignment = { vertical: "middle" };
  ws.getRow(1).height = 22;

  for (const row of result.rows) ws.addRow(row);

  for (const c of result.columns) {
    if (c.type === "money") ws.getColumn(c.key).numFmt = '"$" #,##0';
    if (c.type === "number") ws.getColumn(c.key).numFmt = "#,##0";
  }
  ws.views = [{ state: "frozen", ySplit: 1 }];
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: result.columns.length } };

  return Buffer.from(await wb.xlsx.writeBuffer());
}

/** Convierte un reporte a CSV (separador `;`, coma decimal — convención local). */
export function reportToCsv(result: ReportResult): string {
  const escape = (v: unknown) => {
    const s = v == null ? "" : String(v);
    return /[";\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const header = result.columns.map((c) => escape(c.label)).join(";");
  const lines = result.rows.map((r) => result.columns.map((c) => escape(r[c.key])).join(";"));
  return [header, ...lines].join("\n");
}
