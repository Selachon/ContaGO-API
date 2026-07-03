import ExcelJS from "exceljs";
import type { InvoiceData, TaxDetail } from "../types/dianExcel.js";

// ── Helpers ──────────────────────────────────────────────────────────────────

function sortInvoicesByDate(invoices: InvoiceData[]): InvoiceData[] {
  return [...invoices].sort((a, b) => {
    const da = a.issueDateISO || "9999-12-31";
    const db = b.issueDateISO || "9999-12-31";
    return da.localeCompare(db);
  });
}

function normalizeNit(nit: string | null | undefined): string {
  const raw = (nit || "").trim();
  if (!raw) return "N/A";
  return raw.replace(/[^0-9A-Za-z]/g, "").toUpperCase() || "N/A";
}

function uppercaseBusinessName(value: string | null | undefined): string {
  const raw = String(value || "").trim();
  if (!raw || raw === "N/A") return raw || "";
  return raw.toLocaleUpperCase("es-CO");
}

/**
 * Convierte un string de fecha (ISO o DD/MM/YYYY) en un objeto Date de JS
 * que ExcelJS pueda escribir como fecha nativa de Excel.
 */
function parseExcelDate(dateStr: string | undefined, dateISO: string | undefined): Date | string {
  try {
    // 1. Intentar con ISO (YYYY-MM-DD)
    if (dateISO && /^\d{4}-\d{2}-\d{2}$/.test(dateISO)) {
      const d = new Date(dateISO + "T12:00:00");
      if (!isNaN(d.getTime())) return d;
    }

    // 2. Fallback al string si tiene formato ISO
    if (dateStr && /^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
      const d = new Date(dateStr + "T12:00:00");
      if (!isNaN(d.getTime())) return d;
    }

    // 3. Fallback para formato colombiano DD/MM/YYYY
    if (dateStr && /^\d{2}\/\d{2}\/\d{4}$/.test(dateStr)) {
      const [day, month, year] = dateStr.split("/");
      const d = new Date(`${year}-${month}-${day}T12:00:00`);
      if (!isNaN(d.getTime())) return d;
    }
  } catch (err) {
    console.error(`[Excel] Error parseando fecha: ${dateStr} / ${dateISO}`, err);
  }
  return dateStr || "";
}

const NUM_FMT = "#,##0.00";
const PCT_FMT = "0.00%";
const DATE_FMT = "yyyy-mm-dd";

// Columnas que llevan formato numérico #,##0.00
const CURRENCY_HEADERS = new Set([
  "Subtotal", "Descuento", "Recargo",
  "IVA", "% IVA", "INC", "% INC", "Bolsas", "% Bolsas",
  "ICUI", "% ICUI", "IC", "Otros Impuestos", "IC Porcentual", "% IC Porcentual",
  "ICL", "IBUA", "% IBUA", "ADV",
  "Total",
  "Cantidad", "Base del impuesto", "Descuento detalle", "Recargo detalle", "Precio unitario (incluye impuestos)",
  "Bases de IVAS al 19%", "IVAS del 19%", "Bases IVAS 5%", "IVAS 5%", "Bases sin IVA", "IVA 0%",
]);

const DATE_HEADERS = new Set(["Fecha"]);

/**
 * Calcula los índices de columnas (1-based) que requieren formatos específicos.
 */
function computeFormatCols(headers: string[]): { currencyCols: number[]; percentCols: number[]; dateCols: number[] } {
  const currencyCols: number[] = [];
  const dateCols: number[] = [];
  headers.forEach((h, i) => {
    if (CURRENCY_HEADERS.has(h)) currencyCols.push(i + 1);
    if (DATE_HEADERS.has(h)) dateCols.push(i + 1);
  });
  return { currencyCols, percentCols: [], dateCols };
}

/**
 * Aplica formatos a las celdas de una fila según los índices calculados.
 */
function applyFormats(row: ExcelJS.Row, currencyCols: number[], percentCols: number[], dateCols: number[]): void {
  for (const ci of currencyCols) {
    const cell = row.getCell(ci);
    cell.numFmt = NUM_FMT;
  }
  for (const ci of percentCols) {
    const cell = row.getCell(ci);
    cell.numFmt = PCT_FMT;
  }
  for (const ci of dateCols) {
    const cell = row.getCell(ci);
    cell.numFmt = DATE_FMT;
  }
}

const BRAND_VIOLET = "FF7C2DD3"; // #7C2DD3

function freezeHeaderRows(ws: ExcelJS.Worksheet): void {
  ws.views = [{ state: "frozen" as const, xSplit: 0, ySplit: 4, topLeftCell: "A5", activeCell: "A5" }];
}

function applyCompanyHeader(ws: ExcelJS.Worksheet, name: string, nit: string): void {
  const row1 = ws.getRow(1);
  row1.height = 25;
  const cell1 = row1.getCell(1);
  cell1.value = uppercaseBusinessName(name) || "N/A";
  cell1.font = { bold: true, size: 14, color: { argb: BRAND_VIOLET } };
  cell1.alignment = { vertical: "middle" };

  const row2 = ws.getRow(2);
  row2.height = 20;
  const cell2 = row2.getCell(1);
  cell2.value = nit ? `NIT: ${nit}` : "NIT: N/A";
  cell2.font = { bold: true, size: 11, color: { argb: "FF44546A" } };
  cell2.alignment = { vertical: "middle" };

  ws.getRow(3).height = 8; // Spacer
}

function applyHeaderRow(row: ExcelJS.Row): void {
  row.height = 28;
  row.eachCell((cell) => {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: BRAND_VIOLET },
    };
    cell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 11 };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = {
      bottom: { style: "thin", color: { argb: "FFFFFFFF" } },
      right: { style: "thin", color: { argb: "FFFFFFFF" } },
    };
  });
}

function autoFitColumns(ws: ExcelJS.Worksheet, minWidth = 8, maxWidth = 70): void {
  const colMaxLens: number[] = [];

  ws.eachRow({ includeEmpty: false }, (row) => {
    if (row.number <= 2) return;
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      let len: number;
      const v = cell.value;
      if (v === null || v === undefined) {
        len = 0;
      } else if (v instanceof Date) {
        len = 12;
      } else if (typeof v === "object" && "text" in v) {
        len = String((v as ExcelJS.CellHyperlinkValue).text).length;
      } else if (typeof v === "number") {
        len = String(Math.round(Math.abs(v))).length + 5;
      } else {
        len = String(v).length;
      }
      if (!colMaxLens[colNumber] || len > colMaxLens[colNumber]) {
        colMaxLens[colNumber] = len;
      }
    });
  });

  colMaxLens.forEach((maxLen, colNumber) => {
    if (colNumber > 0) {
      ws.getColumn(colNumber).width = Math.min(Math.max(maxLen + 2, minWidth), maxWidth);
    }
  });
}

// ── Sheet 1: Facturas DIAN ────────────────────────────────────────────────────

function buildSheet1(
  ws: ExcelJS.Worksheet,
  invoices: InvoiceData[],
  includeDriveColumn: boolean,
  companyName: string = "",
  companyNit: string = ""
): void {
  applyCompanyHeader(ws, companyName, companyNit);

  const baseHeaders = [
    "No.", "Tipo documento", "Número factura",
    "NIT Emisor", "Razón Social Emisor",
    "NIT Receptor", "Razón Social Receptor",
    "Fecha", "Concepto", "Forma de pago",
    "Subtotal", "Descuento", "Recargo",
    "IVA", "INC", "Bolsas", "ICUI", "IC", "Otros Impuestos", "ICL", "IC Porcentual", "IBUA", "ADV",
    "Total", "Observaciones",
  ];
  if (includeDriveColumn) baseHeaders.push("Enlace Drive");
  baseHeaders.push("CUFE");

  const headerRow = ws.getRow(4);
  headerRow.values = baseHeaders;
  applyHeaderRow(headerRow);
  freezeHeaderRows(ws);

  const { currencyCols, percentCols, dateCols } = computeFormatCols(baseHeaders);

  let rowNum = 5;
  for (const inv of invoices) {
    const td = Object.fromEntries((inv.taxes || []).map((t) => [t.taxName, t]));

    const rowData: (any)[] = [
      rowNum - 4,
      inv.documentType || "",
      (inv.docNumber || inv.trackId || "").trim(),
      inv.issuerNit || "",
      uppercaseBusinessName(inv.issuerName),
      inv.receiverNit || "",
      uppercaseBusinessName(inv.receiverName),
      parseExcelDate(inv.issueDate, inv.issueDateISO),
      inv.concepts || "",
      inv.paymentMethod || "N/A",
      typeof inv.subtotal === "number" ? inv.subtotal : 0,
      inv.discount || 0,
      inv.surcharge || 0,
      td["IVA"]?.amount ?? 0,
      td["INC"]?.amount ?? 0,
      td["Bolsas"]?.amount ?? 0,
      td["ICUI"]?.amount ?? 0,
      td["IC"]?.amount ?? 0,
      td["Otros Impuestos"]?.amount ?? 0,
      td["ICL"]?.amount ?? 0,
      td["IC Porcentual"]?.amount ?? 0,
      td["IBUA"]?.amount ?? 0,
      td["ADV"]?.amount ?? 0,
      typeof inv.total === "number" ? inv.total : 0,
      inv.notes || "",
    ];

    if (includeDriveColumn) {
      if (inv.driveUrl && !inv.driveUrl.startsWith("ERROR")) {
        rowData.push({ text: "Ver factura", hyperlink: inv.driveUrl });
      } else {
        rowData.push("");
      }
    }
    rowData.push(inv.cufe || "");

    const row = ws.getRow(rowNum);
    row.values = rowData;
    applyFormats(row, currencyCols, percentCols, dateCols);
    rowNum++;
  }
}

// ── Sheet 2: Detallado ────────────────────────────────────────────────────────

function buildSheet2(ws: ExcelJS.Worksheet, invoices: InvoiceData[], companyName: string = "", companyNit: string = ""): void {
  applyCompanyHeader(ws, companyName, companyNit);

  const headers = [
    "Item", "Número Factura", "Tipo documento",
    "NIT Emisor", "Razón Social Emisor", "NIT Receptor", "Razón Social Receptor", "Fecha",
    "Concepto",
    "Cantidad", "Base del impuesto", "Descuento detalle", "Recargo detalle",
    "IVA", "% IVA", "INC", "% INC", "Bolsas", "% Bolsas",
    "ICUI", "% ICUI", "IC", "Otros Impuestos",
    "IC Porcentual", "% IC Porcentual",
    "ICL",
    "IBUA", "% IBUA",
    "ADV",
    "Precio unitario (incluye impuestos)",
    "CUFE",
  ];

  const headerRow = ws.getRow(4);
  headerRow.values = headers;
  applyHeaderRow(headerRow);
  freezeHeaderRows(ws);

  const { currencyCols, percentCols, dateCols } = computeFormatCols(headers);

  let rowNum = 5;
  for (const inv of invoices) {
    const invDocNumber = (inv.docNumber || inv.trackId || "").trim();

    // Build common invoice fields shared by every row of this invoice
    const invCommon = (itemIdx: number, concepto: string) => [
      itemIdx,                                           // A  Item (sequential)
      invDocNumber,                                      // B  Número Factura
      inv.documentType || "",                            // C  Tipo documento
      inv.issuerNit || "",                               // D  NIT Emisor
      uppercaseBusinessName(inv.issuerName),             // E  Razón Social Emisor
      inv.receiverNit || "",                             // F  NIT Receptor
      uppercaseBusinessName(inv.receiverName),           // G  Razón Social Receptor
      parseExcelDate(inv.issueDate, inv.issueDateISO),  // H  Fecha
      concepto,                                          // I  Concepto
    ];

    const buildTaxCols = (td: Record<string, TaxDetail>) => [
      td["IVA"]?.amount ?? 0,                     // N
      (td["IVA"]?.percent ?? 0) / 100,            // O
      td["INC"]?.amount ?? 0,                     // P
      (td["INC"]?.percent ?? 0) / 100,            // Q
      td["Bolsas"]?.amount ?? 0,                  // R
      (td["Bolsas"]?.percent ?? 0) / 100,         // S
      td["ICUI"]?.amount ?? 0,                    // T
      (td["ICUI"]?.percent ?? 0) / 100,           // U
      td["IC"]?.amount ?? 0,                      // V
      td["Otros Impuestos"]?.amount ?? 0,         // W
      td["IC Porcentual"]?.amount ?? 0,           // X
      (td["IC Porcentual"]?.percent ?? 0) / 100,  // Y
      td["ICL"]?.amount ?? 0,                     // AA
      td["IBUA"]?.amount ?? 0,                    // AB
      (td["IBUA"]?.percent ?? 0) / 100,           // AC
      td["ADV"]?.amount ?? 0,                     // AD
    ];

    // Track which tax names are already covered by line items
    const lineItemTaxNames = new Set<string>();
    for (const li of inv.lineItems || []) {
      for (const tx of li.taxes || []) if (tx.amount > 0) lineItemTaxNames.add(tx.taxName);
    }

    let itemIdx = 1;

    // ── Line items ──
    for (const li of inv.lineItems || []) {
      const td = Object.fromEntries((li.taxes || []).map((t) => [t.taxName, t]));
      const totalTax = (li.taxes || []).reduce((s, t) => s + t.amount, 0);

      const rowData: any[] = [
        ...invCommon(itemIdx++, li.description || ""),
        li.quantity,             // J  Cantidad
        li.totalUnitPrice,       // K  Base del impuesto
        li.discount,             // L  Descuento detalle
        li.surcharge,            // M  Recargo detalle
        ...buildTaxCols(td),
        li.totalUnitPrice + totalTax, // AC Precio unitario
        inv.cufe || "",               // AD CUFE
      ];

      const row = ws.getRow(rowNum);
      row.values = rowData;
      applyFormats(row, currencyCols, percentCols, dateCols);
      rowNum++;
    }

    // ── Global taxes (invoice-level, not attributed to any specific line item) ──
    for (const tax of inv.taxes || []) {
      if (tax.amount <= 0 || lineItemTaxNames.has(tax.taxName)) continue;
      const td: Record<string, TaxDetail> = { [tax.taxName]: tax };

      const rowData: any[] = [
        ...invCommon(itemIdx++, tax.taxName),
        0,   // J  Cantidad
        0,   // K  Base del impuesto
        0,   // L  Descuento
        0,   // M  Recargo
        ...buildTaxCols(td),
        tax.amount, // AC Precio unitario = monto del impuesto
        inv.cufe || "",
      ];

      const row = ws.getRow(rowNum);
      row.values = rowData;
      applyFormats(row, currencyCols, percentCols, dateCols);
      rowNum++;
    }
  }
}

// ── Sheet IVA: Reporte Auxiliar IVA ──────────────────────────────────────────

// Impuestos que no son IVA ni retenciones — se excluyen de las columnas extra
const IVA_EXCLUDED_TAXES = new Set(["IVA", "ReteIVA", "ReteRenta", "ReteICA"]);

function colLetter(n: number): string {
  let s = "";
  while (n > 0) {
    n--;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

const EXTRA_TAX_PRIORITY = ["INC", "IC", "Otros Impuestos", "IC Porcentual", "Bolsas", "ICUI", "ICL", "IBUA", "ADV"];

function collectExtraTaxNames(invoices: InvoiceData[]): string[] {
  const seen = new Set<string>();
  for (const inv of invoices) {
    for (const li of inv.lineItems || []) {
      for (const tax of li.taxes || []) {
        if (!IVA_EXCLUDED_TAXES.has(tax.taxName)) seen.add(tax.taxName);
      }
    }
  }
  const priority = EXTRA_TAX_PRIORITY.filter(n => seen.has(n));
  const rest = [...seen].filter(n => !EXTRA_TAX_PRIORITY.includes(n));
  return [...priority, ...rest];
}

function buildSheetIVA(ws: ExcelJS.Worksheet, invoices: InvoiceData[], companyName: string = "", companyNit: string = ""): void {
  applyCompanyHeader(ws, companyName, companyNit);

  const extraTaxNames = collectExtraTaxNames(invoices);
  const BASE_COLS = 14; // columnas fijas antes de los impuestos extra
  const totalCols = BASE_COLS + extraTaxNames.length;
  const lastCol = colLetter(totalCols);

  // Disclaimer Row
  const disclaimerRow = ws.getRow(4);
  disclaimerRow.height = 30;
  const disclaimerCell = disclaimerRow.getCell(1);
  disclaimerCell.value = "AVISO LEGAL: Esta herramienta proporciona un reporte auxiliar basado en los datos extraídos. ContaGO NO liquida impuestos. Es responsabilidad exclusiva del usuario revisar y validar esta información antes de cualquier presentación ante autoridades tributarias.";
  disclaimerCell.font = { italic: true, size: 10, color: { argb: "FFFF0000" } };
  disclaimerCell.alignment = { vertical: "middle", wrapText: true };
  ws.mergeCells(`A4:${lastCol}4`);

  const headers = [
    "No.", "Tipo documento", "Número factura",
    "NIT Emisor", "Razón Social Emisor",
    "NIT Receptor", "Razón Social Receptor",
    "Fecha",
    "Bases de IVAS al 19%", "IVAS del 19%",
    "Bases IVAS 5%", "IVAS 5%",
    "Bases sin IVA", "IVA 0%",
    ...extraTaxNames,
  ];

  const headerRow = ws.getRow(5);
  headerRow.values = headers;
  applyHeaderRow(headerRow);

  // Custom freeze for this sheet since we added a row
  ws.views = [{ state: "frozen" as const, xSplit: 0, ySplit: 5, topLeftCell: "A6", activeCell: "A6" }];

  const { currencyCols, percentCols, dateCols } = computeFormatCols(headers);
  // Extra tax columns are always currency regardless of whether the name is in CURRENCY_HEADERS
  for (let i = BASE_COLS + 1; i <= totalCols; i++) {
    if (!currencyCols.includes(i)) currencyCols.push(i);
  }

  let rowNum = 6;
  for (const inv of invoices) {
    let base19 = 0;
    let iva19 = 0;
    let base5 = 0;
    let iva5 = 0;
    let baseSinIVA = 0;
    const extraTaxTotals = new Map<string, number>(extraTaxNames.map(n => [n, 0]));

    for (const li of inv.lineItems || []) {
      const ivaTax = (li.taxes || []).find(t => t.taxName === "IVA");
      const base = li.totalUnitPrice || 0;

      if (!ivaTax) {
        baseSinIVA += base;
      } else {
        const pct = ivaTax.percent;
        if (pct === 19) {
          base19 += base;
          iva19 += ivaTax.amount || 0;
        } else if (pct === 5) {
          base5 += base;
          iva5 += ivaTax.amount || 0;
        } else {
          // Tanto el IVA 0% como cualquier otro porcentaje diferente a 19 o 5
          // se suman a la base sin IVA, ya que no generan IVA a reportar en esas columnas.
          baseSinIVA += base;
        }
      }

      for (const tax of li.taxes || []) {
        if (extraTaxTotals.has(tax.taxName)) {
          extraTaxTotals.set(tax.taxName, (extraTaxTotals.get(tax.taxName) || 0) + (tax.amount || 0));
        }
      }
    }

    const rowData: (any)[] = [
      rowNum - 5,
      inv.documentType || "",
      (inv.docNumber || inv.trackId || "").trim(),
      inv.issuerNit || "",
      uppercaseBusinessName(inv.issuerName),
      inv.receiverNit || "",
      uppercaseBusinessName(inv.receiverName),
      parseExcelDate(inv.issueDate, inv.issueDateISO),
      base19,
      iva19,
      base5,
      iva5,
      baseSinIVA,
      0, // IVA 0% siempre en ceros según requerimiento
      ...extraTaxNames.map(n => extraTaxTotals.get(n) || 0),
    ];

    const row = ws.getRow(rowNum);
    row.values = rowData;
    applyFormats(row, currencyCols, percentCols, dateCols);
    rowNum++;
  }
}

// ── DV Calculation ───────────────────────────────────────────────────────────

function calcularDV(nit: string): string {
  const clean = nit.replace(/[\s,.\-]/g, "");
  if (!clean || isNaN(Number(clean))) return "";
  const vpri = [0, 3, 7, 13, 17, 19, 23, 29, 37, 41, 43, 47, 53, 59, 67, 71];
  const z = clean.length;
  let x = 0;
  for (let i = 0; i < z; i++) {
    x += parseInt(clean[i], 10) * vpri[z - i];
  }
  const y = x % 11;
  return String(y > 1 ? 11 - y : y);
}

// ── Sheet 3: Datos de terceros ────────────────────────────────────────────────

function buildSheet3(ws: ExcelJS.Worksheet, invoices: InvoiceData[], companyName: string = "", companyNit: string = ""): void {
  applyCompanyHeader(ws, companyName, companyNit);

  const headers = [
    "NIT", "DV", "Razón Social", "Nombre Comercial", "Resp. Tributaria",
    "País", "Departamento", "Ciudad", "Dirección", "Teléfono", "Correo",
  ];

  const headerRow = ws.getRow(4);
  headerRow.values = headers;
  applyHeaderRow(headerRow);
  freezeHeaderRows(ws);

  type ThirdPartyRow = {
    nit: string; name: string; commercial: string; taxResp: string;
    country: string; dept: string; city: string; addr: string;
    phone: string; email: string;
  };

  const byNit = new Map<string, ThirdPartyRow>();

  function addParty(party: ThirdPartyRow): void {
    const key = normalizeNit(party.nit);
    if (!byNit.has(key)) byNit.set(key, party);
  }

  for (const inv of invoices) {
    addParty({
      nit: inv.issuerNit,
      name: inv.issuerName,
      commercial: inv.issuerCommercialName || "N/A",
      taxResp: inv.issuerTaxResponsibility || "N/A",
      country: inv.issuerCountry || "N/A",
      dept: inv.issuerDepartment || "N/A",
      city: inv.issuerCity || "N/A",
      addr: inv.issuerAddress || "N/A",
      phone: inv.issuerPhone || "N/A",
      email: inv.issuerEmail || "N/A",
    });
    addParty({
      nit: inv.receiverNit,
      name: inv.receiverName,
      commercial: inv.receiverCommercialName || "N/A",
      taxResp: inv.receiverTaxResponsibility || "N/A",
      country: inv.receiverCountry || "N/A",
      dept: inv.receiverDepartment || "N/A",
      city: inv.receiverCity || "N/A",
      addr: inv.receiverAddress || "N/A",
      phone: inv.receiverPhone || "N/A",
      email: inv.receiverEmail || "N/A",
    });
  }

  let rowNum = 5;
  for (const p of byNit.values()) {
    const nitDigits = (p.nit || "").replace(/\D/g, "");
    const nitValue: string | number = nitDigits ? parseInt(nitDigits, 10) : (p.nit || "");
    const dvStr = nitDigits ? calcularDV(p.nit) : "";
    const dvValue: string | number = dvStr !== "" ? parseInt(dvStr, 10) : "";
    const cleanPhone = (p.phone || "").replace(/\|/g, "").trim();
    const firstEmail = (p.email || "").split(",")[0].trim();

    const row = ws.getRow(rowNum);
    row.values = [
      nitValue, dvValue, uppercaseBusinessName(p.name), p.commercial, p.taxResp,
      p.country, p.dept, p.city, p.addr,
      cleanPhone, firstEmail,
    ];
    if (nitDigits) {
      row.getCell(1).numFmt = "0";
      row.getCell(2).numFmt = "0";
    }
    rowNum++;
  }
}

// ── Public exports ────────────────────────────────────────────────────────────

export async function generateExcelFile(
  invoices: InvoiceData[],
  outputPath: string,
  includeDriveColumn: boolean,
  isSentDocuments: boolean = false,
  companyName: string = "",
  companyNit: string = ""
): Promise<void> {
  const sorted = sortInvoicesByDate(invoices);

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "ContaGO";
  workbook.created = new Date();

  const ws1 = workbook.addWorksheet("Facturas DIAN");
  const ws2 = workbook.addWorksheet("Detallado");
  const ws3 = workbook.addWorksheet("Reporte Auxiliar IVA");
  const ws4 = workbook.addWorksheet("Datos de terceros");

  buildSheet1(ws1, sorted, includeDriveColumn, companyName, companyNit);
  autoFitColumns(ws1);
  buildSheet2(ws2, sorted, companyName, companyNit);
  autoFitColumns(ws2);
  buildSheetIVA(ws3, sorted, companyName, companyNit);
  autoFitColumns(ws3);
  buildSheet3(ws4, sorted, companyName, companyNit);
  autoFitColumns(ws4);

  await workbook.xlsx.writeFile(outputPath);
}

export function generateExcelFilename(
  startDate?: string,
  endDate?: string,
  prefix: string = "Facturas DIAN"
): string {
  const fmt = (d: string): string => {
    const [year, month, day] = d.split("-");
    const months = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
    return `${months[parseInt(month) - 1]} ${parseInt(day)} ${year}`;
  };
  const start = startDate ? fmt(startDate) : "Inicio";
  const end = endDate ? fmt(endDate) : "Fin";
  return `${prefix} ${start} - ${end}.xlsx`;
}

export async function generateThirdPartiesExcelFile(
  invoices: Partial<InvoiceData>[],
  outputPath: string,
  isSentDocuments: boolean = false,
  companyName: string = "",
  companyNit: string = ""
): Promise<void> {
  const sorted = sortInvoicesByDate(invoices as InvoiceData[]);

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "ContaGO";
  workbook.created = new Date();

  const ws3 = workbook.addWorksheet("Datos de terceros");
  buildSheet3(ws3, sorted, companyName, companyNit);
  autoFitColumns(ws3);

  await workbook.xlsx.writeFile(outputPath);
}
