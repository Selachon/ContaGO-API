/**
 * Reporte Excel de la conciliación bancaria (exceljs).
 * Hojas: Cuadre · Conciliados · Agrupados · Partidas conciliatorias.
 */
import ExcelJS from "exceljs";
import type { ReconciliationResult, RecItem } from "./bankReconciliationService.js";

interface Meta {
  bank?: string;
  accountCode?: string;
  accountName?: string;
  period?: string;
}

const BRAND = "FF7C2DD3";
const GOLD = "FFC9A961";
const HEAD_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: BRAND } } as const;
const MONEY = '#,##0.00;[Red]-#,##0.00';

function styleHeaderRow(row: ExcelJS.Row) {
  row.eachCell((cell) => {
    cell.fill = HEAD_FILL as ExcelJS.Fill;
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.alignment = { vertical: "middle" };
  });
}

function titleRow(ws: ExcelJS.Worksheet, text: string, span: number) {
  const row = ws.addRow([text]);
  ws.mergeCells(row.number, 1, row.number, span);
  row.getCell(1).font = { bold: true, size: 14, color: { argb: BRAND } };
  return row;
}

export async function generateReconciliationExcel(
  result: ReconciliationResult,
  meta: Meta
): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "ContaGO";
  wb.created = new Date();

  // ─── Hoja CUADRE ────────────────────────────────────────────────────────
  const c = result.cuadre;
  const ws = wb.addWorksheet("Cuadre");
  ws.columns = [{ width: 46 }, { width: 20 }];
  titleRow(ws, "Conciliación bancaria", 2);
  ws.addRow([`Banco: ${meta.bank || "-"}`]);
  ws.addRow([`Cuenta: ${meta.accountCode || ""} ${meta.accountName || ""}`.trim()]);
  ws.addRow([`Periodo: ${meta.period || "-"}`]);
  ws.addRow([`Tolerancia: $${result.tolerance.toLocaleString()}`]);
  ws.addRow([]);

  const money = (label: string, value: number | null) => {
    const r = ws.addRow([label, value ?? 0]);
    r.getCell(2).numFmt = MONEY;
    return r;
  };

  money("Saldo final según EXTRACTO", c.closingStatement);
  money("Saldo final según CONTABILIDAD", c.closingLedger);
  const dr = money("Diferencia", c.closingDiff);
  dr.font = { bold: true };
  ws.addRow([]);

  ws.addRow(["Explicación de la diferencia"]).getCell(1).font = { bold: true, color: { argb: GOLD } };
  money("Diferencia de saldo inicial (partida anterior)", c.openingDiff);
  money("(+) Ingresos en extracto no contabilizados", c.partidas.ingresos_no_contabilizados);
  money("(−) Egresos en extracto no contabilizados", -c.partidas.egresos_no_contabilizados);
  money("(−) Ingresos en contabilidad sin extracto", -c.partidas.ingresos_contab_sin_extracto);
  money("(+) Egresos en contabilidad sin extracto", c.partidas.egresos_contab_sin_extracto);
  money("(±) Ajuste por tolerancia (montos casados)", c.toleranceAdjustment);
  const er = money("= Diferencia explicada", c.explained);
  er.font = { bold: true };
  const ur = money("Diferencia SIN EXPLICAR", c.unexplained);
  ur.font = { bold: true, color: { argb: c.balanced ? "FF1B7F3B" : "FFC0392B" } };
  ws.addRow([c.balanced ? "✓ Conciliación cuadrada" : "⚠ Hay diferencias sin explicar"]).getCell(1).font = {
    bold: true,
    color: { argb: c.balanced ? "FF1B7F3B" : "FFC0392B" },
  };
  ws.addRow([]);
  const cnt = result.counts;
  ws.addRow([`Movimientos extracto: ${cnt.statement} · contabilidad: ${cnt.ledger}`]);
  ws.addRow([`Casados 1:1: ${cnt.matched1to1} · Agrupados: ${cnt.grouped} · Manuales: ${cnt.manual}`]);
  ws.addRow([`Sin cruzar — extracto: ${cnt.unmatchedStatement} · contabilidad: ${cnt.unmatchedLedger}`]);

  // ─── Hoja CONCILIADOS (1:1 y manuales) ───────────────────────────────────
  const wsM = wb.addWorksheet("Conciliados");
  wsM.columns = [
    { header: "Tipo", width: 10 },
    { header: "Dir", width: 8 },
    { header: "Fecha extracto", width: 14 },
    { header: "Descripción extracto", width: 42 },
    { header: "Valor extracto", width: 16 },
    { header: "Fecha contab.", width: 14 },
    { header: "Comprobante", width: 14 },
    { header: "Tercero", width: 30 },
    { header: "Valor contab.", width: 16 },
    { header: "Dif.", width: 12 },
    { header: "Ambiguo", width: 10 },
  ];
  styleHeaderRow(wsM.getRow(1));
  for (const m of result.matches.filter((x) => x.type !== "group")) {
    const s = m.statement[0];
    const l = m.ledger[0];
    const r = wsM.addRow([
      m.type,
      m.direction === "in" ? "Ingreso" : "Egreso",
      s?.date || "",
      s?.description || "",
      m.valueStatement,
      l?.date || "",
      l?.voucher || "",
      l?.thirdParty || "",
      m.valueLedger,
      m.residual,
      m.ambiguous ? "Sí" : "",
    ]);
    r.getCell(5).numFmt = MONEY;
    r.getCell(9).numFmt = MONEY;
    r.getCell(10).numFmt = MONEY;
    if (m.ambiguous) r.getCell(11).font = { color: { argb: "FFC0392B" }, bold: true };
  }

  // ─── Hoja AGRUPADOS (subset-sum) ─────────────────────────────────────────
  const wsG = wb.addWorksheet("Agrupados");
  wsG.columns = [
    { header: "Grupo", width: 8 },
    { header: "Lado", width: 12 },
    { header: "Fecha", width: 14 },
    { header: "Descripción / Comprobante", width: 46 },
    { header: "Valor", width: 16 },
  ];
  styleHeaderRow(wsG.getRow(1));
  let gi = 0;
  for (const m of result.matches.filter((x) => x.type === "group")) {
    gi++;
    for (const s of m.statement) {
      const r = wsG.addRow([gi, "Extracto", s.date, s.description, s.value]);
      r.getCell(5).numFmt = MONEY;
    }
    for (const l of m.ledger) {
      const r = wsG.addRow([gi, "Contabilidad", l.date, `${l.voucher || ""} ${l.description}`.trim(), l.value]);
      r.getCell(5).numFmt = MONEY;
    }
    const tot = wsG.addRow([
      gi,
      "TOTAL",
      "",
      `Extracto ${m.valueStatement.toLocaleString()} ↔ Contab. ${m.valueLedger.toLocaleString()} (dif ${m.residual})`,
      m.valueStatement,
    ]);
    tot.font = { bold: true };
    tot.getCell(5).numFmt = MONEY;
    wsG.addRow([]);
  }

  // ─── Hoja PARTIDAS CONCILIATORIAS ────────────────────────────────────────
  const wsP = wb.addWorksheet("Partidas conciliatorias");
  wsP.columns = [
    { header: "Categoría", width: 36 },
    { header: "Fecha", width: 14 },
    { header: "Descripción", width: 46 },
    { header: "Tercero / Comprobante", width: 34 },
    { header: "Valor", width: 16 },
  ];
  styleHeaderRow(wsP.getRow(1));
  const labels: Record<string, string> = {
    ingresos_no_contabilizados: "Ingresos no contabilizados (en extracto)",
    egresos_no_contabilizados: "Egresos no contabilizados (en extracto)",
    ingresos_contab_sin_extracto: "Ingresos en contabilidad sin extracto",
    egresos_contab_sin_extracto: "Egresos en contabilidad sin extracto",
  };
  for (const [key, label] of Object.entries(labels)) {
    const items = (result.partidas as Record<string, RecItem[]>)[key] || [];
    if (items.length === 0) continue;
    const hdr = wsP.addRow([label]);
    hdr.getCell(1).font = { bold: true, color: { argb: BRAND } };
    for (const it of items) {
      const r = wsP.addRow([
        "",
        it.date,
        it.description,
        it.thirdParty || it.voucher || "",
        it.value,
      ]);
      r.getCell(5).numFmt = MONEY;
    }
    const subtotal = items.reduce((a, x) => a + x.value, 0);
    const st = wsP.addRow(["", "", "", "Subtotal", subtotal]);
    st.font = { bold: true };
    st.getCell(5).numFmt = MONEY;
    wsP.addRow([]);
  }

  const out = await wb.xlsx.writeBuffer();
  return Buffer.from(out);
}
