import ExcelJS from "exceljs";
import type { InvoiceData } from "../types/dianExcel.js";

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

const HEADER_FILL: ExcelJS.Fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FF1F3864" },
};

const HEADER_FONT: Partial<ExcelJS.Font> = {
  bold: true,
  color: { argb: "FFFFFFFF" },
  size: 11,
};

function styleHeader(cell: ExcelJS.Cell): void {
  cell.fill = HEADER_FILL;
  cell.font = HEADER_FONT;
  cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
}

function applyHeaderRow(row: ExcelJS.Row): void {
  row.eachCell({ includeEmpty: true }, (cell) => styleHeader(cell));
  row.height = 30;
}

const NUM_FMT = "#,##0.00";
const PCT_FMT = "0.00%";

// All numeric columns use #,##0.00 — user can convert to % manually in Excel as needed.
// "%" columns store the raw percentage value (e.g. 19, not 0.19) so the number is readable.
const CURRENCY_HEADERS = new Set([
  "Subtotal", "Descuento", "Recargo",
  "IVA", "% IVA", "INC", "% INC", "Bolsas", "% Bolsas",
  "ICUI", "% ICUI", "IC", "IC Porcentual", "% IC Porcentual",
  "ICL", "IBUA", "% IBUA", "ADV",
  "Total",
  "Cantidad", "Base del impuesto", "Descuento detalle", "Recargo detalle",
  "Precio unitario (incluye impuestos)",
]);

// Computes format column indices from the actual headers array so indices can never drift.
function computeFormatCols(headers: string[]): { currencyCols: number[]; percentCols: number[] } {
  const currencyCols: number[] = [];
  headers.forEach((h, i) => {
    if (CURRENCY_HEADERS.has(h)) currencyCols.push(i + 1);
  });
  return { currencyCols, percentCols: [] }; // no percent cell format — all numeric as #,##0.00
}

// Applies numFmt directly to cells — more reliable than column-level style in ExcelJS.
function applyFormats(row: ExcelJS.Row, currencyCols: number[], percentCols: number[]): void {
  for (const ci of currencyCols) row.getCell(ci).numFmt = NUM_FMT;
  for (const ci of percentCols) row.getCell(ci).numFmt = PCT_FMT;
}

// Freezes the first two rows (decorative spacer + header) so they stay visible on scroll.
function freezeHeaderRows(ws: ExcelJS.Worksheet): void {
  ws.views = [{ state: "frozen" as const, xSplit: 0, ySplit: 2, topLeftCell: "A3", activeCell: "A3" }];
}

// Auto-fits column widths based on actual cell content after data is written.
function autoFitColumns(ws: ExcelJS.Worksheet, minWidth = 8, maxWidth = 70): void {
  const colMaxLens: number[] = [];

  ws.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      let len: number;
      const v = cell.value;
      if (v === null || v === undefined) {
        len = 0;
      } else if (typeof v === "object" && v !== null && "text" in v) {
        len = String((v as ExcelJS.CellHyperlinkValue).text).length;
      } else if (typeof v === "number") {
        // Estimate formatted length (commas + 2 decimal places)
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
//
// A: No.               B: Tipo documento      C: Número factura
// D: NIT Emisor        E: Razón Social Emisor
// F: NIT Receptor      G: Razón Social Receptor
// H: Fecha             I: Concepto            J: Forma de pago
// K: Subtotal          L: Descuento           M: Recargo
// N: IVA               O: INC                 P: Bolsas
// Q: ICUI              R: IC                  S: ICL
// T: IC Porcentual     U: IBUA                V: ADV
// W: Total             [X: Enlace Drive]      Y/X: CUFE

function buildSheet1(
  ws: ExcelJS.Worksheet,
  invoices: InvoiceData[],
  includeDriveColumn: boolean
): void {
  ws.getRow(1).height = 5;

  const baseHeaders = [
    "No.", "Tipo documento", "Número factura",
    "NIT Emisor", "Razón Social Emisor",
    "NIT Receptor", "Razón Social Receptor",
    "Fecha", "Concepto", "Forma de pago",
    "Subtotal", "Descuento", "Recargo",
    "IVA", "INC", "Bolsas", "ICUI", "IC", "ICL", "IC Porcentual", "IBUA", "ADV",
    "Total",
  ];
  if (includeDriveColumn) baseHeaders.push("Enlace Drive");
  baseHeaders.push("CUFE");

  const headerRow = ws.getRow(2);
  headerRow.values = ["", ...baseHeaders];
  applyHeaderRow(headerRow);
  freezeHeaderRows(ws);

  const { currencyCols, percentCols } = computeFormatCols(baseHeaders);

  let rowNum = 3;
  for (const inv of invoices) {
    const td = Object.fromEntries((inv.taxes || []).map((t) => [t.taxName, t]));

    const rowData: (string | number | ExcelJS.CellHyperlinkValue)[] = [
      rowNum - 2,
      inv.documentType || "",
      (inv.docNumber || inv.trackId || "").trim(),
      inv.issuerNit || "",
      inv.issuerName || "",
      inv.receiverNit || "",
      inv.receiverName || "",
      inv.issueDate || "",
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
      td["ICL"]?.amount ?? 0,
      td["IC Porcentual"]?.amount ?? 0,
      td["IBUA"]?.amount ?? 0,
      td["ADV"]?.amount ?? 0,
      typeof inv.total === "number" ? inv.total : 0,
    ];

    if (includeDriveColumn) {
      if (inv.driveUrl && !inv.driveUrl.startsWith("ERROR")) {
        rowData.push({ text: "Ver factura", hyperlink: inv.driveUrl } as ExcelJS.CellHyperlinkValue);
      } else {
        rowData.push("");
      }
    }
    rowData.push(inv.cufe || "");

    const row = ws.getRow(rowNum);
    row.values = ["", ...rowData];
    applyFormats(row, currencyCols, percentCols);
    rowNum++;
  }
}

// ── Sheet 2: Detallado ────────────────────────────────────────────────────────
//
// A: Item              B: Número Factura      C: Tipo documento    D: Concepto
// E: Cantidad          F: Base del impuesto   G: Descuento detalle H: Recargo detalle
// I: IVA               J: % IVA               K: INC               L: % INC
// M: Bolsas            N: % Bolsas            O: ICUI              P: % ICUI
// Q: IC                (sin % IC — impuesto por grados de alcohol, sin % fijo)
// R: IC Porcentual     S: % IC Porcentual
// T: ICL               (sin % ICL — ídem)
// U: IBUA              V: % IBUA
// W: ADV               (sin % ADV)
// X: Precio unitario (incluye impuestos)

function buildSheet2(ws: ExcelJS.Worksheet, invoices: InvoiceData[]): void {
  ws.getRow(1).height = 5;

  const headers = [
    "Item", "Número Factura", "Tipo documento", "Concepto",
    "Cantidad", "Base del impuesto", "Descuento detalle", "Recargo detalle",
    "IVA", "% IVA", "INC", "% INC", "Bolsas", "% Bolsas",
    "ICUI", "% ICUI", "IC",
    "IC Porcentual", "% IC Porcentual",
    "ICL",
    "IBUA", "% IBUA",
    "ADV",
    "Precio unitario (incluye impuestos)",
  ];

  const headerRow = ws.getRow(2);
  headerRow.values = ["", ...headers];
  applyHeaderRow(headerRow);
  freezeHeaderRows(ws);

  const { currencyCols, percentCols } = computeFormatCols(headers);

  let rowNum = 3;
  for (const inv of invoices) {
    const invDocNumber = (inv.docNumber || inv.trackId || "").trim();
    for (const li of inv.lineItems || []) {
      const td = Object.fromEntries((li.taxes || []).map((t) => [t.taxName, t]));
      const totalTax = (li.taxes || []).reduce((s, t) => s + t.amount, 0);

      const rowData: (string | number)[] = [
        li.lineNumber,           // A  Item
        invDocNumber,            // B  Número Factura
        inv.documentType || "",  // C  Tipo documento
        li.description || "",    // D  Concepto
        li.quantity,             // E  Cantidad
        li.totalUnitPrice,                   // F  Base del impuesto
        li.discount,                         // G  Descuento detalle
        li.surcharge,                        // H  Recargo detalle
        td["IVA"]?.amount ?? 0,                     // I  IVA
        (td["IVA"]?.percent ?? 0) / 100,            // J  % IVA (0.19 → usuario convierte a %)
        td["INC"]?.amount ?? 0,                     // K  INC
        (td["INC"]?.percent ?? 0) / 100,            // L  % INC
        td["Bolsas"]?.amount ?? 0,                  // M  Bolsas
        (td["Bolsas"]?.percent ?? 0) / 100,         // N  % Bolsas
        td["ICUI"]?.amount ?? 0,                    // O  ICUI
        (td["ICUI"]?.percent ?? 0) / 100,           // P  % ICUI
        td["IC"]?.amount ?? 0,                      // Q  IC (sin % IC)
        td["IC Porcentual"]?.amount ?? 0,           // R  IC Porcentual
        (td["IC Porcentual"]?.percent ?? 0) / 100,  // S  % IC Porcentual
        td["ICL"]?.amount ?? 0,                     // T  ICL (sin % ICL)
        td["IBUA"]?.amount ?? 0,                    // U  IBUA
        (td["IBUA"]?.percent ?? 0) / 100,           // V  % IBUA
        td["ADV"]?.amount ?? 0,                     // W  ADV (sin % ADV)
        li.totalUnitPrice + totalTax,        // X  Precio unitario (incluye impuestos)
      ];

      const row = ws.getRow(rowNum);
      row.values = ["", ...rowData];
      applyFormats(row, currencyCols, percentCols);
      rowNum++;
    }
  }
}

// ── Sheet 3: Datos de terceros ────────────────────────────────────────────────

function buildSheet3(ws: ExcelJS.Worksheet, invoices: InvoiceData[]): void {
  ws.getRow(1).height = 5;

  const headers = [
    "NIT", "Razón Social", "Nombre Comercial", "Resp. Tributaria",
    "País", "Departamento", "Ciudad", "Dirección", "Teléfono", "Correo",
  ];

  const headerRow = ws.getRow(2);
  headerRow.values = ["", ...headers];
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

  let rowNum = 3;
  for (const p of byNit.values()) {
    ws.getRow(rowNum).values = [
      "",
      p.nit, p.name, p.commercial, p.taxResp,
      p.country, p.dept, p.city, p.addr,
      p.phone, p.email,
    ];
    rowNum++;
  }
}

// ── Public exports ────────────────────────────────────────────────────────────

export async function generateExcelFile(
  invoices: InvoiceData[],
  outputPath: string,
  includeDriveColumn: boolean,
  isSentDocuments: boolean = false // kept for API compatibility, no longer affects columns
): Promise<void> {
  const sorted = sortInvoicesByDate(invoices);

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "ContaGO";
  workbook.created = new Date();

  const ws1 = workbook.addWorksheet("Facturas DIAN");
  const ws2 = workbook.addWorksheet("Detallado");
  const ws3 = workbook.addWorksheet("Datos de terceros");

  buildSheet1(ws1, sorted, includeDriveColumn);
  autoFitColumns(ws1);
  buildSheet2(ws2, sorted);
  autoFitColumns(ws2);
  buildSheet3(ws3, sorted);
  autoFitColumns(ws3);

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
  isSentDocuments: boolean = false // kept for API compatibility
): Promise<void> {
  const sorted = sortInvoicesByDate(invoices as InvoiceData[]);

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "ContaGO";
  workbook.created = new Date();

  const ws3 = workbook.addWorksheet("Datos de terceros");
  buildSheet3(ws3, sorted);
  autoFitColumns(ws3);

  await workbook.xlsx.writeFile(outputPath);
}
