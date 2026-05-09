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

// Style helpers applied directly to cells / columns
const HEADER_FILL: ExcelJS.Fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FF1F3864" }, // dark blue
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

function setCurrencyFmt(col: ExcelJS.Column | Partial<ExcelJS.Column>): void {
  col.numFmt = "#,##0.00";
}

function setPercentFmt(col: ExcelJS.Column | Partial<ExcelJS.Column>): void {
  col.numFmt = "0.00%";
}

// ── Sheet 1: Facturas DIAN ────────────────────────────────────────────────────

/**
 * Sheet 1 columns (always both issuer and receiver):
 * A  No.          B  Tipo documento     C  Número factura
 * D  NIT Emisor   E  Razón Social Emisor
 * F  NIT Receptor G  Razón Social Receptor
 * H  Fecha        I  Concepto           J  Forma de pago
 * K  Subtotal     L  Descuento          M  Recargo
 * N  IVA          O  INC                P  Bolsas
 * Q  ICUI         R  IC                 S  ICL
 * T  IC Porcentual U  IBUA              V  ADV
 * W  Total        [X  Enlace Drive — solo si includeDriveColumn]
 * Y/X  CUFE
 */
function buildSheet1(
  ws: ExcelJS.Worksheet,
  invoices: InvoiceData[],
  includeDriveColumn: boolean
): void {
  ws.getRow(1).height = 5; // decorative spacer row (empty)

  // Build header labels
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
  headerRow.values = ["", ...baseHeaders]; // ExcelJS row values are 1-indexed via array pos [1..n]
  applyHeaderRow(headerRow);

  // Column widths and formats
  const colWidths: number[] = [
    6,   // A No.
    22,  // B Tipo documento
    28,  // C Número factura
    18,  // D NIT Emisor
    35,  // E Razón Social Emisor
    18,  // F NIT Receptor
    35,  // G Razón Social Receptor
    14,  // H Fecha
    40,  // I Concepto
    20,  // J Forma de pago
    16,  // K Subtotal
    14,  // L Descuento
    14,  // M Recargo
    14,  // N IVA
    14,  // O INC
    14,  // P Bolsas
    14,  // Q ICUI
    14,  // R IC
    14,  // S ICL
    18,  // T IC Porcentual
    14,  // U IBUA
    14,  // V ADV
    16,  // W Total
  ];
  if (includeDriveColumn) colWidths.push(20); // X Enlace Drive
  colWidths.push(80); // CUFE (last)

  colWidths.forEach((w, i) => {
    ws.getColumn(i + 1).width = w;
  });

  // Currency columns: K(11) L(12) M(13) N(14) O(15) P(16) Q(17) R(18) S(19) T(20) U(21) V(22) W(23)
  const currencyCols = [11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23];
  for (const ci of currencyCols) {
    setCurrencyFmt(ws.getColumn(ci));
  }

  // Data rows start at row 3
  let rowNum = 3;
  for (const inv of invoices) {
    const td = Object.fromEntries((inv.taxes || []).map((t) => [t.taxName, t]));

    const rowData: (string | number | ExcelJS.CellHyperlinkValue)[] = [
      rowNum - 2, // No.
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
    rowNum++;
  }
}

// ── Sheet 2: Detallado ────────────────────────────────────────────────────────

/**
 * Sheet 2 columns (line items):
 * A  Item            B  Número Factura    C  Tipo documento  D  Concepto
 * E  Cantidad        F  Base del impuesto G  Descuento detalle H Recargo detalle
 * I  IVA             J  % IVA             K  INC             L  % INC
 * M  Bolsas          N  % Bolsas         O  ICUI            P  % ICUI
 * Q  IC              R  % IC             S  IC Porcentual   T  % IC Porcentual
 * U  ICL             V  % ICL            W  IBUA            X  % IBUA
 * Y  ADV             Z  % ADV            AA Precio unitario (incluye impuestos)
 */
function buildSheet2(ws: ExcelJS.Worksheet, invoices: InvoiceData[]): void {
  ws.getRow(1).height = 5;

  const headers = [
    "Item", "Número Factura", "Tipo documento", "Concepto",
    "Cantidad", "Base del impuesto", "Descuento detalle", "Recargo detalle",
    "IVA", "% IVA", "INC", "% INC", "Bolsas", "% Bolsas",
    "ICUI", "% ICUI", "IC", "% IC",
    "IC Porcentual", "% IC Porcentual", "ICL", "% ICL",
    "IBUA", "% IBUA", "ADV", "% ADV",
    "Precio unitario (incluye impuestos)",
  ];

  const headerRow = ws.getRow(2);
  headerRow.values = ["", ...headers];
  applyHeaderRow(headerRow);

  const colWidths = [
    6, 28, 22, 40,         // A B C D
    10, 16, 14, 14,        // E F G H
    14, 10, 14, 10,        // I J K L
    14, 10, 14, 10,        // M N O P
    14, 10, 18, 12,        // Q R S T
    14, 10, 14, 10,        // U V W X
    14, 10,                // Y Z
    26,                    // AA
  ];
  colWidths.forEach((w, i) => {
    ws.getColumn(i + 1).width = w;
  });

  // Currency columns: F(6) G(7) H(8) I(9) K(11) M(13) O(15) Q(17) S(19) U(21) W(23) Y(25) AA(27)
  const currencyCols = [6, 7, 8, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27];
  for (const ci of currencyCols) {
    setCurrencyFmt(ws.getColumn(ci));
  }

  // Percent columns: J(10) L(12) N(14) P(16) R(18) T(20) V(22) X(24) Z(26)
  const percentCols = [10, 12, 14, 16, 18, 20, 22, 24, 26];
  for (const ci of percentCols) {
    setPercentFmt(ws.getColumn(ci));
  }

  let rowNum = 3;
  for (const inv of invoices) {
    const invDocNumber = (inv.docNumber || inv.trackId || "").trim();
    for (const li of inv.lineItems || []) {
      const td = Object.fromEntries((li.taxes || []).map((t) => [t.taxName, t]));
      const totalTax = (li.taxes || []).reduce((s, t) => s + t.amount, 0);

      const rowData: (string | number)[] = [
        li.lineNumber,
        invDocNumber,
        inv.documentType || "",
        li.description || "",
        li.quantity,
        li.totalUnitPrice,
        li.discount,
        li.surcharge,
        td["IVA"]?.amount ?? 0,
        (td["IVA"]?.percent ?? 0) / 100,
        td["INC"]?.amount ?? 0,
        (td["INC"]?.percent ?? 0) / 100,
        td["Bolsas"]?.amount ?? 0,
        (td["Bolsas"]?.percent ?? 0) / 100,
        td["ICUI"]?.amount ?? 0,
        (td["ICUI"]?.percent ?? 0) / 100,
        td["IC"]?.amount ?? 0,
        (td["IC"]?.percent ?? 0) / 100,
        td["IC Porcentual"]?.amount ?? 0,
        (td["IC Porcentual"]?.percent ?? 0) / 100,
        td["ICL"]?.amount ?? 0,
        (td["ICL"]?.percent ?? 0) / 100,
        td["IBUA"]?.amount ?? 0,
        (td["IBUA"]?.percent ?? 0) / 100,
        td["ADV"]?.amount ?? 0,
        (td["ADV"]?.percent ?? 0) / 100,
        li.totalUnitPrice + totalTax,
      ];

      const row = ws.getRow(rowNum);
      row.values = ["", ...rowData];
      rowNum++;
    }
  }
}

// ── Sheet 3: Datos de terceros ────────────────────────────────────────────────

/**
 * Sheet 3 columns: NIT, Razón Social, Nombre Comercial, Resp. Tributaria,
 *                  País, Departamento, Ciudad, Dirección, Teléfono, Correo
 * Lists ALL NITs appearing in invoices (both issuers and receivers).
 */
function buildSheet3(ws: ExcelJS.Worksheet, invoices: InvoiceData[]): void {
  ws.getRow(1).height = 5;

  const headers = [
    "NIT", "Razón Social", "Nombre Comercial", "Resp. Tributaria",
    "País", "Departamento", "Ciudad", "Dirección", "Teléfono", "Correo",
  ];

  const headerRow = ws.getRow(2);
  headerRow.values = ["", ...headers];
  applyHeaderRow(headerRow);

  const colWidths = [20, 40, 35, 30, 16, 22, 22, 40, 18, 35];
  colWidths.forEach((w, i) => {
    ws.getColumn(i + 1).width = w;
  });

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
  buildSheet2(ws2, sorted);
  buildSheet3(ws3, sorted);

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

  await workbook.xlsx.writeFile(outputPath);
}
