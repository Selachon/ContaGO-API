import JSZip from "jszip";
import fs from "fs";
import type { InvoiceData, TaxDetail } from "../types/dianExcel.js";

// ─── helpers ────────────────────────────────────────────────────────────────

function colLetter(n: number): string {
  let result = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    result = String.fromCharCode(65 + r) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

function xmlEsc(s: string): string {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function strCell(col: string, row: number, value: string): string {
  const escaped = xmlEsc(value);
  return `<c r="${col}${row}" t="inlineStr"><is><t>${escaped}</t></is></c>`;
}

function numCell(col: string, row: number, value: number): string {
  const v = isFinite(value) ? value : 0;
  return `<c r="${col}${row}"><v>${v}</v></c>`;
}

function getTaxAmt(taxes: TaxDetail[], name: string): number {
  return taxes?.find((t) => t.taxName === name)?.amount ?? 0;
}

function getTaxPct(taxes: TaxDetail[], name: string): number {
  return taxes?.find((t) => t.taxName === name)?.percent ?? 0;
}

// ─── row builders ────────────────────────────────────────────────────────────

function buildSheet1Row(inv: Partial<InvoiceData>, id: number, rowNum: number, withDriveLink: boolean): string {
  const taxes = inv.taxes ?? [];
  const baseCells = [
    numCell("A", rowNum, id),
    strCell("B", rowNum, inv.documentType ?? ""),
    strCell("C", rowNum, inv.docNumber ?? ""),
    strCell("D", rowNum, inv.issuerNit ?? ""),
    strCell("E", rowNum, inv.issuerName ?? ""),
    strCell("F", rowNum, inv.issueDate ?? ""),
    strCell("G", rowNum, inv.concepts ?? ""),
    strCell("H", rowNum, inv.paymentMethod ?? ""),
    numCell("I", rowNum, inv.subtotal ?? 0),
    numCell("J", rowNum, inv.discount ?? 0),
    numCell("K", rowNum, inv.surcharge ?? 0),
    numCell("L", rowNum, getTaxAmt(taxes, "IVA")),
    numCell("M", rowNum, getTaxAmt(taxes, "INC")),
    numCell("N", rowNum, getTaxAmt(taxes, "Bolsas")),
    numCell("O", rowNum, getTaxAmt(taxes, "ICUI")),
    numCell("P", rowNum, getTaxAmt(taxes, "IC")),
    numCell("Q", rowNum, getTaxAmt(taxes, "ICL")),
    numCell("R", rowNum, getTaxAmt(taxes, "IC Porcentual")),
    numCell("S", rowNum, getTaxAmt(taxes, "IBUA")),
    numCell("T", rowNum, getTaxAmt(taxes, "ADV")),
    numCell("U", rowNum, inv.total ?? 0),
    strCell("V", rowNum, inv.notes ?? ""),
    strCell("W", rowNum, inv.cufe ?? ""),
  ];
  if (withDriveLink) baseCells.push(strCell("X", rowNum, inv.driveUrl ?? ""));
  const cols = withDriveLink ? 24 : 23;
  return `<row r="${rowNum}" spans="1:${cols}">${baseCells.join("")}</row>`;
}

function buildSheet2Row(
  inv: Partial<InvoiceData>,
  lineIdx: number,
  itemNum: number,
  rowNum: number
): string {
  const line = inv.lineItems![lineIdx];
  const taxes = line.taxes ?? [];
  const cells = [
    numCell("A", rowNum, itemNum),
    strCell("B", rowNum, inv.docNumber ?? ""),
    strCell("C", rowNum, inv.documentType ?? ""),
    strCell("D", rowNum, line.description ?? ""),
    numCell("E", rowNum, line.totalUnitPrice ?? 0),
    numCell("F", rowNum, line.quantity ?? 0),
    numCell("G", rowNum, line.discount ?? 0),
    numCell("H", rowNum, line.surcharge ?? 0),
    numCell("I", rowNum, getTaxAmt(taxes, "IVA")),
    numCell("J", rowNum, getTaxPct(taxes, "IVA")),
    numCell("K", rowNum, getTaxAmt(taxes, "INC")),
    numCell("L", rowNum, getTaxPct(taxes, "INC")),
    numCell("M", rowNum, getTaxAmt(taxes, "Bolsas")),
    numCell("N", rowNum, getTaxPct(taxes, "Bolsas")),
    numCell("O", rowNum, getTaxAmt(taxes, "ICUI")),
    numCell("P", rowNum, getTaxPct(taxes, "ICUI")),
    numCell("Q", rowNum, getTaxAmt(taxes, "IC")),
    numCell("R", rowNum, getTaxPct(taxes, "IC")),
    numCell("S", rowNum, getTaxAmt(taxes, "IC Porcentual")),
    numCell("T", rowNum, getTaxPct(taxes, "IC Porcentual")),
    numCell("U", rowNum, getTaxAmt(taxes, "ICL")),
    numCell("V", rowNum, getTaxPct(taxes, "ICL")),
    numCell("W", rowNum, getTaxAmt(taxes, "IBUA")),
    numCell("X", rowNum, getTaxPct(taxes, "IBUA")),
    numCell("Y", rowNum, getTaxAmt(taxes, "ADV")),
    numCell("Z", rowNum, getTaxPct(taxes, "ADV")),
    numCell(colLetter(27), rowNum, line.unitPrice ?? 0),
  ].join("");
  return `<row r="${rowNum}" spans="1:27">${cells}</row>`;
}

function buildSheet3Row(inv: Partial<InvoiceData>, rowNum: number): string {
  const cells = [
    strCell("A", rowNum, inv.issuerNit ?? ""),
    strCell("B", rowNum, inv.issuerName ?? ""),
    strCell("C", rowNum, inv.issuerCommercialName ?? ""),
    strCell("D", rowNum, inv.issuerTaxResponsibility ?? ""),
    strCell("E", rowNum, inv.issuerCountry ?? ""),
    strCell("F", rowNum, inv.issuerDepartment ?? ""),
    strCell("G", rowNum, inv.issuerCity ?? ""),
    strCell("H", rowNum, inv.issuerAddress ?? ""),
    strCell("I", rowNum, inv.issuerPhone ?? ""),
    strCell("J", rowNum, inv.issuerEmail ?? ""),
  ].join("");
  return `<row r="${rowNum}" spans="1:10">${cells}</row>`;
}

// ─── sheet patcher ──────────────────────────────────────────────────────────

/**
 * Reemplaza el contenido de <sheetData> manteniendo las filas 1 y 2 (logo y
 * encabezados) e inyectando las filas de datos desde la fila 3 en adelante.
 * También actualiza <dimension ref="..."> para reflejar el rango real.
 */
function patchSheetData(
  sheetXml: string,
  dataRows: string[],
  lastCol: string
): string {
  // Extraer filas 1 y 2 del sheetData original
  const row1Match = sheetXml.match(/<row r="1"[^>]*>[\s\S]*?<\/row>/);
  const row2Match = sheetXml.match(/<row r="2"[^>]*>[\s\S]*?<\/row>/);
  const row1 = row1Match ? row1Match[0] : "";
  const row2 = row2Match ? row2Match[0] : "";

  const totalRows = 2 + dataRows.length;
  const newSheetData = `<sheetData>${row1}${row2}${dataRows.join("")}</sheetData>`;

  let result = sheetXml.replace(/<sheetData>[\s\S]*?<\/sheetData>/, newSheetData);

  // Actualizar dimensión
  result = result.replace(
    /<dimension ref="[^"]*"/,
    `<dimension ref="A1:${lastCol}${totalRows}"`
  );

  return result;
}

// ─── table ref patcher ──────────────────────────────────────────────────────

function patchTableRef(tableXml: string, lastCol: string, lastRow: number): string {
  return tableXml.replace(
    /(<table\b[^>]*\bref=")[^"]*(")/,
    `$1A2:${lastCol}${lastRow}$2`
  );
}

// ─── main export ────────────────────────────────────────────────────────────

/**
 * Genera un Excel a partir del template de la plantilla DIAN-ContaGO y los
 * datos parseados de cada factura. Devuelve el buffer XLSX listo para enviar.
 */
export async function generateCufeExcel(
  invoices: Partial<InvoiceData>[],
  templatePath: string,
  includeDriveLinks: boolean = false
): Promise<Buffer> {
  const templateBuffer = fs.readFileSync(templatePath);
  const zip = await JSZip.loadAsync(templateBuffer);

  // ── Sheet 1: Facturas DIAN ──────────────────────────────────────────────
  const s1LastCol = includeDriveLinks ? "X" : "W";
  const s1Rows: string[] = [];
  invoices.forEach((inv, i) => {
    s1Rows.push(buildSheet1Row(inv, i + 1, i + 3, includeDriveLinks));
  });

  const s1Xml = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
  const s1LastRow = 2 + s1Rows.length;
  zip.file("xl/worksheets/sheet1.xml", patchSheetData(s1Xml, s1Rows, s1LastCol));
  const t1Xml = await zip.file("xl/tables/table1.xml")!.async("string");
  zip.file("xl/tables/table1.xml", patchTableRef(t1Xml, s1LastCol, s1LastRow));

  // ── Sheet 2: Detallado ─────────────────────────────────────────────────
  const s2Rows: string[] = [];
  let s2Row = 3;
  for (const inv of invoices) {
    const lines = inv.lineItems ?? [];
    lines.forEach((_, lineIdx) => {
      s2Rows.push(buildSheet2Row(inv, lineIdx, lineIdx + 1, s2Row));
      s2Row++;
    });
  }

  const s2Xml = await zip.file("xl/worksheets/sheet2.xml")!.async("string");
  const s2LastRow = 2 + s2Rows.length;
  zip.file(
    "xl/worksheets/sheet2.xml",
    patchSheetData(s2Xml, s2Rows, "AA")
  );
  const t2Xml = await zip.file("xl/tables/table2.xml")!.async("string");
  zip.file("xl/tables/table2.xml", patchTableRef(t2Xml, "AA", s2LastRow));

  // ── Sheet 3: Datos de terceros ─────────────────────────────────────────
  const seenNits = new Set<string>();
  const s3Rows: string[] = [];
  let s3Row = 3;
  for (const inv of invoices) {
    const nit = inv.issuerNit ?? "";
    if (!nit || seenNits.has(nit)) continue;
    seenNits.add(nit);
    s3Rows.push(buildSheet3Row(inv, s3Row));
    s3Row++;
  }

  const s3Xml = await zip.file("xl/worksheets/sheet3.xml")!.async("string");
  const s3LastRow = 2 + s3Rows.length;
  zip.file(
    "xl/worksheets/sheet3.xml",
    patchSheetData(s3Xml, s3Rows, "J")
  );
  const t3Xml = await zip.file("xl/tables/table3.xml")!.async("string");
  zip.file("xl/tables/table3.xml", patchTableRef(t3Xml, "J", s3LastRow));

  // ── Generar buffer final ───────────────────────────────────────────────
  const outputBuffer = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });

  return outputBuffer;
}
