/**
 * Genera el papel de trabajo de conciliación bancaria (3 hojas):
 *  1. Portada    — cuadre + partidas conciliatorias, estilo plantilla contable
 *  2. Auxiliar   — movimientos del auxiliar contable (SIIGO)
 *  3. Extracto   — movimientos del extracto bancario
 */
import ExcelJS from "exceljs";
import type { ReconciliationResult, RecItem } from "./bankReconciliationService.js";

export interface GenMeta {
  bank?: string;
  accountCode?: string;
  accountName?: string;
  period?: string;
}

export interface GenMovement {
  date: string;
  description: string;
  value: number;
  direction: "in" | "out";
  kind?: string;
  balance?: number | null;
}

export interface GenEntry {
  date: string;
  voucher?: string;
  nit?: string;
  thirdParty?: string;
  description: string;
  value: number;
  direction: "in" | "out";
}

// ─── Colores ──────────────────────────────────────────────────────────────────
const GRAY_DARK  = "FFA6A6A6"; // encabezado banco
const GRAY_MED   = "FFD9D9D9"; // cabecera de sección
const GRAY_LITE  = "FFF2F2F2"; // cabecera de columnas
const AMBER      = "FFFFC000"; // filas de total
const MONEY_FMT  = '#,##0.00;[Red]-#,##0.00';
const DATE_FMT   = 'dd/mm/yyyy';

type ARGB = string;

function fillSolid(cell: ExcelJS.Cell, argb: ARGB) {
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb } };
}
function boldFont(cell: ExcelJS.Cell, color: ARGB = "FF000000") {
  cell.font = { ...(cell.font as object ?? {}), bold: true, color: { argb: color } };
}
function moneyFmt(cell: ExcelJS.Cell) {
  cell.numFmt = MONEY_FMT;
  cell.alignment = { ...(cell.alignment ?? {}), horizontal: "right", vertical: "middle" };
}
function center(cell: ExcelJS.Cell) {
  cell.alignment = { ...(cell.alignment ?? {}), horizontal: "center", vertical: "middle" };
}
function thinBorder(cell: ExcelJS.Cell) {
  const s = { style: "thin" as const };
  cell.border = { top: s, left: s, bottom: s, right: s };
}

// Aplica borde fino a todas las celdas de una fila entre columnas c1 y c2.
function borderRow(row: ExcelJS.Row, c1: number, c2: number) {
  for (let c = c1; c <= c2; c++) thinBorder(row.getCell(c));
}

const ROW_H = 15;

// ─── Hoja PORTADA ────────────────────────────────────────────────────────────
function buildPortada(
  wb: ExcelJS.Workbook,
  result: ReconciliationResult,
  meta: GenMeta,
) {
  const ws = wb.addWorksheet("Portada");
  ws.views = [{ showGridLines: false }];

  // A=spacer B C D E | F=operador G=etiqueta H=valor
  ws.columns = [
    { width: 2 },   // A
    { width: 12 },  // B
    { width: 14 },  // C
    { width: 28 },  // D
    { width: 14 },  // E
    { width: 7 },   // F
    { width: 50 },  // G
    { width: 20 },  // H
  ];

  const c  = result.cuadre;
  const bankName   = (meta.bank || "").toUpperCase();
  const acctCode   = meta.accountCode || "";
  const acctName   = meta.accountName || "";
  const period     = meta.period || "";

  // Partidas conciliatorias (valores positivos)
  const egresosNoContab   = c.partidas.egresos_no_contabilizados;
  const ingresosNoContab  = c.partidas.ingresos_no_contabilizados;
  const egresosContabSin  = c.partidas.egresos_contab_sin_extracto;
  const ingresosContabSin = c.partidas.ingresos_contab_sin_extracto;
  const partidasContabNet = egresosContabSin - ingresosContabSin; // neto en libros sin extracto

  // ─── Bloque de identificación + resumen (filas 1-10) ─────────────────────

  // Fila 1: banco (izq) | "+" | SALDO SEGÚN EXTRACTO | valor
  const r1 = ws.addRow([null, bankName, null, null, null, "+", "SALDO SEGÚN EXTRACTO", c.closingStatement ?? 0]);
  r1.height = 22;
  fillSolid(r1.getCell(2), GRAY_DARK); boldFont(r1.getCell(2), "FFFFFFFF");
  r1.getCell(2).alignment = { horizontal: "center", vertical: "middle" };
  r1.getCell(6).alignment = { horizontal: "center", vertical: "middle" };
  boldFont(r1.getCell(7)); moneyFmt(r1.getCell(8)); boldFont(r1.getCell(8));

  // Fila 2: (merge con fila 1 col B-E) | "-" | SALDO SEGÚN LIBROS | valor
  const r2 = ws.addRow([null, null, null, null, null, "-", "SALDO SEGÚN LIBROS", c.closingLedger ?? 0]);
  r2.height = 22;
  fillSolid(r2.getCell(2), GRAY_DARK);
  r2.getCell(6).alignment = { horizontal: "center", vertical: "middle" };
  moneyFmt(r2.getCell(8)); boldFont(r2.getCell(8));
  ws.mergeCells(1, 2, 2, 5);

  // Fila 3: BANCO + nombre | "=" | DIFERENCIA | valor
  const diffColor = c.balanced ? "FF1B7F3B" : "FFC0392B";
  const r3 = ws.addRow([null, "BANCO", bankName, null, null, "=", "DIFERENCIA", c.closingDiff]);
  r3.height = ROW_H;
  boldFont(r3.getCell(2)); boldFont(r3.getCell(3));
  r3.getCell(6).alignment = { horizontal: "center", vertical: "middle" };
  boldFont(r3.getCell(7)); moneyFmt(r3.getCell(8));
  r3.getCell(8).font = { bold: true, color: { argb: diffColor } };
  ws.mergeCells(3, 3, 3, 5);

  // Fila 4: CTA No + numero | cabecera RESUMEN
  const r4 = ws.addRow([null, "CTA No", acctName, null, null, null, "RESUMEN DETALLE DE LA DIFERENCIA", null]);
  r4.height = ROW_H;
  boldFont(r4.getCell(2)); boldFont(r4.getCell(3));
  fillSolid(r4.getCell(6), GRAY_DARK); fillSolid(r4.getCell(7), GRAY_DARK); fillSolid(r4.getCell(8), GRAY_DARK);
  boldFont(r4.getCell(7), "FFFFFFFF");
  r4.getCell(7).alignment = { horizontal: "center", vertical: "middle" };
  ws.mergeCells(4, 3, 4, 5);
  ws.mergeCells(4, 6, 4, 8);

  // Fila 5: CODIGO | CHEQUES/EGRESOS NO COBRADOS
  const r5 = ws.addRow([null, "CODIGO", acctCode, null, null, "+", "CHEQUES / EGRESOS NO CONTABILIZADOS", egresosNoContab]);
  r5.height = ROW_H;
  boldFont(r5.getCell(2));
  r5.getCell(6).alignment = { horizontal: "center", vertical: "middle" };
  moneyFmt(r5.getCell(8));
  ws.mergeCells(5, 3, 5, 5);

  // Fila 6: FECHA + periodo | CONSIGNACIONES NO CONTABILIZADAS
  const r6 = ws.addRow([null, "FECHA", period, null, null, "+", "CONSIGNACIONES / INGRESOS NO CONTABILIZADOS", ingresosNoContab]);
  r6.height = ROW_H;
  boldFont(r6.getCell(2));
  r6.getCell(6).alignment = { horizontal: "center", vertical: "middle" };
  moneyFmt(r6.getCell(8));

  // Fila 7: (merge B6:B7, C6:E7) | NOTAS DEBITO
  const r7 = ws.addRow([null, null, null, null, null, "+", "NOTAS DÉBITO NO CONTABILIZADAS", 0]);
  r7.height = ROW_H;
  r7.getCell(6).alignment = { horizontal: "center", vertical: "middle" };
  moneyFmt(r7.getCell(8));
  ws.mergeCells(6, 2, 7, 2);
  ws.mergeCells(6, 3, 7, 5);

  // Fila 8: OBSERVACIONES | NOTAS CREDITO
  const r8 = ws.addRow([null, "OBSERVACIONES:", null, null, null, "-", "NOTAS CRÉDITO NO CONTABILIZADAS", 0]);
  r8.height = ROW_H;
  r8.getCell(6).alignment = { horizontal: "center", vertical: "middle" };
  moneyFmt(r8.getCell(8));

  // Fila 9: (merge B8:E10) | PARTIDAS NO REFLEJADAS
  const r9 = ws.addRow([null, null, null, null, null, "+/-", "PARTIDAS NO REFLEJADAS EN EXTRACTO", partidasContabNet]);
  r9.height = ROW_H;
  r9.getCell(6).alignment = { horizontal: "center", vertical: "middle" };
  moneyFmt(r9.getCell(8));

  // Fila 10: TOTAL DIFERENCIA CONCILIADA
  const r10 = ws.addRow([null, null, null, null, null, "=", "TOTAL DIFERENCIA CONCILIADA", c.explained]);
  r10.height = ROW_H;
  r10.getCell(6).alignment = { horizontal: "center", vertical: "middle" };
  boldFont(r10.getCell(7)); moneyFmt(r10.getCell(8)); boldFont(r10.getCell(8));
  ws.mergeCells(8, 2, 10, 5);

  // ─── Tablas de detalle (filas 11+) ────────────────────────────────────────

  let rn = 11;

  const addSectionHeader = (leftLabel: string, rightLabel: string) => {
    const r = ws.addRow([null, leftLabel, null, null, null, rightLabel, null, null]);
    r.height = ROW_H;
    fillSolid(r.getCell(2), GRAY_MED); boldFont(r.getCell(2));
    fillSolid(r.getCell(3), GRAY_MED);
    fillSolid(r.getCell(4), GRAY_MED);
    fillSolid(r.getCell(5), GRAY_MED);
    fillSolid(r.getCell(6), GRAY_MED); boldFont(r.getCell(6));
    fillSolid(r.getCell(7), GRAY_MED); boldFont(r.getCell(7));
    fillSolid(r.getCell(8), GRAY_MED);
    ws.mergeCells(rn, 2, rn, 5);
    ws.mergeCells(rn, 6, rn, 8);
    rn++;
  };

  const addColHeaders = (left: string[], right: string[]) => {
    const r = ws.addRow([null, ...left, ...right]);
    r.height = ROW_H;
    for (let c = 2; c <= 8; c++) {
      fillSolid(r.getCell(c), GRAY_LITE);
      boldFont(r.getCell(c));
      r.getCell(c).alignment = { horizontal: "center", vertical: "middle" };
      thinBorder(r.getCell(c));
    }
    rn++;
  };

  const addDataRows = (leftItems: RecItem[], rightItems: RecItem[]) => {
    const len = Math.max(leftItems.length, rightItems.length, 1);
    for (let i = 0; i < len; i++) {
      const l = leftItems[i];
      const r = rightItems[i];
      const row = ws.addRow([
        null,
        l?.date || null,
        l?.voucher || l?.thirdParty || null,
        l?.description?.slice(0, 45) || null,
        l?.value || null,
        r?.date || null,
        r?.description?.slice(0, 55) || null,
        r?.value || null,
      ]);
      row.height = ROW_H;
      if (l?.date) row.getCell(2).numFmt = DATE_FMT;
      if (l?.value) moneyFmt(row.getCell(5));
      if (r?.date) row.getCell(6).numFmt = DATE_FMT;
      if (r?.value) moneyFmt(row.getCell(8));
      for (let c = 2; c <= 8; c++) thinBorder(row.getCell(c));
      rn++;
    }
  };

  const addSubtotalRow = (leftTotal: number, rightTotal: number) => {
    const r = ws.addRow([null, null, null, "TOTAL", leftTotal, null, "TOTAL", rightTotal]);
    r.height = ROW_H;
    boldFont(r.getCell(4)); moneyFmt(r.getCell(5)); boldFont(r.getCell(5));
    boldFont(r.getCell(7)); moneyFmt(r.getCell(8)); boldFont(r.getCell(8));
    fillSolid(r.getCell(5), AMBER);
    fillSolid(r.getCell(8), AMBER);
    for (let c = 2; c <= 8; c++) thinBorder(r.getCell(c));
    rn++;
  };

  // Sección 1: egresos extracto sin libros | ingresos extracto sin libros
  addSectionHeader("EGRESOS EN EXTRACTO NO CONTABILIZADOS", "INGRESOS EN EXTRACTO NO CONTABILIZADOS");
  addColHeaders(["FECHA", "COMPROBANTE", "DETALLE", "VALOR"], ["FECHA", "DETALLE", "VALOR"]);
  const eg1 = result.partidas.egresos_no_contabilizados;
  const ing1 = result.partidas.ingresos_no_contabilizados;
  addDataRows(eg1, ing1);
  addSubtotalRow(egresosNoContab, ingresosNoContab);

  // Sección 2: egresos libros sin extracto | ingresos libros sin extracto
  ws.addRow([]); rn++;
  addSectionHeader("EGRESOS EN LIBROS SIN REFLEJO EN EXTRACTO", "INGRESOS EN LIBROS SIN REFLEJO EN EXTRACTO");
  addColHeaders(["FECHA", "COMPROBANTE", "DETALLE", "VALOR"], ["FECHA", "COMPROBANTE", "VALOR"]);
  const eg2 = result.partidas.egresos_contab_sin_extracto;
  const ing2 = result.partidas.ingresos_contab_sin_extracto;
  addDataRows(eg2, ing2);
  addSubtotalRow(egresosContabSin, ingresosContabSin);

  // Fila de firmas
  ws.addRow([]); rn++;
  const sigRow = ws.addRow([null, "PREPARADO POR:", null, "ELABORADO POR:", null, "REVISADO POR:", null, "CONTABILIZADO POR:"]);
  sigRow.height = 36;
  for (let col of [2, 4, 6, 8]) {
    fillSolid(sigRow.getCell(col), GRAY_MED);
    boldFont(sigRow.getCell(col));
    center(sigRow.getCell(col));
    thinBorder(sigRow.getCell(col));
  }
  ws.mergeCells(rn, 2, rn, 3);
  ws.mergeCells(rn, 4, rn, 5);
  ws.mergeCells(rn, 6, rn, 7);
}

// ─── Hoja AUXILIAR CONTABLE ───────────────────────────────────────────────────
function buildAuxiliar(wb: ExcelJS.Workbook, entries: GenEntry[], meta: GenMeta) {
  const ws = wb.addWorksheet("Auxiliar contable");
  ws.columns = [
    { header: "Fecha",        width: 13, key: "date"       },
    { header: "Comprobante",  width: 14, key: "voucher"    },
    { header: "NIT",          width: 14, key: "nit"        },
    { header: "Tercero",      width: 36, key: "third"      },
    { header: "Descripción",  width: 46, key: "desc"       },
    { header: "Débito",       width: 16, key: "debit"      },
    { header: "Crédito",      width: 16, key: "credit"     },
  ];

  // Título
  const tRow = ws.insertRow(1, [`Auxiliar contable — ${meta.accountName || ""} · ${meta.period || ""}`]);
  ws.mergeCells(1, 1, 1, 7);
  tRow.height = 18;
  tRow.getCell(1).font = { bold: true, size: 12 };
  tRow.getCell(1).alignment = { horizontal: "left", vertical: "middle" };

  // Reemplazar la fila de encabezados (ahora en fila 2)
  styleHeaderRow(ws.getRow(2));

  // Datos
  for (const e of entries) {
    const r = ws.addRow([
      e.date,
      e.voucher || "",
      e.nit || "",
      e.thirdParty || "",
      e.description,
      e.direction === "out" ? e.value : 0,
      e.direction === "in"  ? e.value : 0,
    ]);
    r.getCell(1).numFmt = DATE_FMT;
    r.getCell(6).numFmt = MONEY_FMT;
    r.getCell(7).numFmt = MONEY_FMT;
    r.height = ROW_H;
  }

  // Totales
  const lastRow = ws.lastRow!.number;
  const totRow = ws.addRow([
    "", "", "", "", "TOTAL",
    { formula: `SUM(F3:F${lastRow})` },
    { formula: `SUM(G3:G${lastRow})` },
  ]);
  totRow.height = ROW_H;
  boldFont(totRow.getCell(5));
  totRow.getCell(6).numFmt = MONEY_FMT; boldFont(totRow.getCell(6)); fillSolid(totRow.getCell(6), AMBER);
  totRow.getCell(7).numFmt = MONEY_FMT; boldFont(totRow.getCell(7)); fillSolid(totRow.getCell(7), AMBER);
}

// ─── Hoja MOVIMIENTO DEL BANCO ────────────────────────────────────────────────
function buildExtracto(wb: ExcelJS.Workbook, movements: GenMovement[], meta: GenMeta) {
  const ws = wb.addWorksheet("Movimiento del banco");
  ws.columns = [
    { header: "Fecha",       width: 13, key: "date"  },
    { header: "Descripción", width: 52, key: "desc"  },
    { header: "Tipo",        width: 10, key: "kind"  },
    { header: "Ingreso",     width: 16, key: "in"    },
    { header: "Egreso",      width: 16, key: "out"   },
    { header: "Saldo",       width: 18, key: "bal"   },
  ];

  // Título + resumen de saldos
  const bank = (meta.bank || "").toUpperCase();
  const tRow = ws.insertRow(1, [`Extracto bancario — ${bank} · ${meta.accountName || ""} · ${meta.period || ""}`]);
  ws.mergeCells(1, 1, 1, 6);
  tRow.height = 18;
  tRow.getCell(1).font = { bold: true, size: 12 };
  tRow.getCell(1).alignment = { horizontal: "left", vertical: "middle" };

  styleHeaderRow(ws.getRow(2));

  for (const m of movements) {
    const r = ws.addRow([
      m.date,
      m.description,
      m.kind === "bank_fee" ? "Gasto banco" : m.direction === "in" ? "Ingreso" : "Egreso",
      m.direction === "in"  ? m.value : 0,
      m.direction === "out" ? m.value : 0,
      m.balance ?? null,
    ]);
    r.getCell(1).numFmt = DATE_FMT;
    r.getCell(4).numFmt = MONEY_FMT;
    r.getCell(5).numFmt = MONEY_FMT;
    if (m.balance != null) r.getCell(6).numFmt = MONEY_FMT;
    r.height = ROW_H;
  }

  const lastRow = ws.lastRow!.number;
  const totRow = ws.addRow([
    "", "", "TOTAL",
    { formula: `SUM(D3:D${lastRow})` },
    { formula: `SUM(E3:E${lastRow})` },
    "",
  ]);
  totRow.height = ROW_H;
  boldFont(totRow.getCell(3));
  totRow.getCell(4).numFmt = MONEY_FMT; boldFont(totRow.getCell(4)); fillSolid(totRow.getCell(4), AMBER);
  totRow.getCell(5).numFmt = MONEY_FMT; boldFont(totRow.getCell(5)); fillSolid(totRow.getCell(5), AMBER);
}

function styleHeaderRow(row: ExcelJS.Row) {
  const BRAND = "FF7C2DD3";
  row.height = ROW_H;
  row.eachCell((cell) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BRAND } };
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.alignment = { horizontal: "center", vertical: "middle" };
  });
}

// ─── Entrada pública ──────────────────────────────────────────────────────────
export async function generateReconciliationExcel(
  result: ReconciliationResult,
  meta: GenMeta,
  rawStatement?: GenMovement[],
  rawLedger?: GenEntry[],
): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "ContaGO";
  wb.created = new Date();

  buildPortada(wb, result, meta);
  buildAuxiliar(wb, rawLedger ?? [], meta);
  buildExtracto(wb, rawStatement ?? [], meta);

  const out = await wb.xlsx.writeBuffer();
  return Buffer.from(out);
}
