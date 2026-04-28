import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";
import type { InvoiceData } from "../types/dianExcel.js";

/**
 * Convierte un numero de columna (1-indexed) a letra de Excel
 * Ej: 1 -> A, 26 -> Z, 27 -> AA
 */
function getExcelColumnLetter(colNum: number): string {
  let result = "";
  let num = colNum;
  while (num > 0) {
    const remainder = (num - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    num = Math.floor((num - 1) / 26);
  }
  return result;
}

/**
 * Recolecta todos los tipos de impuestos unicos de todas las facturas
 * IVA siempre se incluye aunque ninguna factura lo tenga
 */
function collectAllTaxTypes(invoices: InvoiceData[]): string[] {
  const taxTypesSet = new Set<string>();

  // IVA siempre presente
  taxTypesSet.add("IVA");

  for (const invoice of invoices) {
    // Impuestos a nivel de factura
    for (const tax of invoice.taxes || []) {
      if (tax.taxName && tax.taxName !== "IVA") {
        taxTypesSet.add(tax.taxName);
      }
    }

    // Impuestos a nivel de linea
    for (const line of invoice.lineItems || []) {
      for (const tax of line.taxes || []) {
        if (tax.taxName && tax.taxName !== "IVA") {
          taxTypesSet.add(tax.taxName);
        }
      }
    }
  }

  // Ordenar: IVA primero, luego INC, luego Bolsas, luego el resto alfabeticamente
  const priority = ["IVA", "INC", "Bolsas", "ICUI", "IC"];
  const sortedTaxTypes = Array.from(taxTypesSet).sort((a, b) => {
    const aIdx = priority.indexOf(a);
    const bIdx = priority.indexOf(b);
    if (aIdx !== -1 && bIdx !== -1) return aIdx - bIdx;
    if (aIdx !== -1) return -1;
    if (bIdx !== -1) return 1;
    return a.localeCompare(b);
  });

  return sortedTaxTypes;
}

/**
 * Ordena las facturas por fecha de emision en orden ascendente
 */
function sortInvoicesByDate(invoices: InvoiceData[]): InvoiceData[] {
  return [...invoices].sort((a, b) => {
    const dateA = a.issueDateISO || "9999-12-31";
    const dateB = b.issueDateISO || "9999-12-31";
    return dateA.localeCompare(dateB);
  });
}

function resolveTemplatePath(): string {
  const configuredPath = process.env.DIAN_EXCEL_TEMPLATE_PATH?.trim() || "templates/dian-excel-template.xlsx";

  const candidates = new Set<string>();

  if (path.isAbsolute(configuredPath)) {
    candidates.add(configuredPath);
  } else {
    candidates.add(path.join(process.cwd(), configuredPath));
  }

  candidates.add(path.join(process.cwd(), "templates", "dian-excel-template.xlsx"));
  candidates.add(path.join(path.dirname(new URL(import.meta.url).pathname), "../../templates/dian-excel-template.xlsx"));
  candidates.add("/app/templates/dian-excel-template.xlsx");

  for (const candidate of candidates) {
    const normalized = path.normalize(candidate);
    if (fs.existsSync(normalized)) {
      return normalized;
    }
  }

  throw new Error(
    `No se encontró la plantilla Excel. Revisa DIAN_EXCEL_TEMPLATE_PATH. cwd=${process.cwd()} configurado=${configuredPath}`
  );
}

function getTemplateHeaderRow(): number {
  const raw = Number(process.env.DIAN_EXCEL_TEMPLATE_HEADER_ROW || "2");
  if (!Number.isFinite(raw) || raw < 1) return 1;
  return Math.floor(raw);
}

function normalizeHeader(value: string): string {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "")
    .trim();
}

function cellText(value: ExcelJS.CellValue | undefined | null): string {
  if (value == null) return "";
  if (typeof value === "string") return value.trim();
  if (typeof value === "number" || typeof value === "boolean") return String(value).trim();
  if (typeof value === "object" && "text" in value && typeof value.text === "string") {
    return value.text.trim();
  }
  if (typeof value === "object" && "richText" in value && Array.isArray(value.richText)) {
    return value.richText.map((part: { text?: string }) => part.text || "").join("").trim();
  }
  return "";
}

function buildHeaderMap(worksheet: ExcelJS.Worksheet, headerRowIndex: number): Map<string, number> {
  const map = new Map<string, number>();
  const row = worksheet.getRow(headerRowIndex);
  row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    const text = cellText(cell.value);
    if (text) map.set(normalizeHeader(text), colNumber);
  });
  return map;
}

function ensureHeaders(
  worksheet: ExcelJS.Worksheet,
  headerRowIndex: number,
  headers: string[],
  getWidth?: (header: string) => number
): Map<string, number> {
  const headerRow = worksheet.getRow(headerRowIndex);
  const map = buildHeaderMap(worksheet, headerRowIndex);

  for (const header of headers) {
    const key = normalizeHeader(header);
    if (map.has(key)) continue;
    const colNumber = worksheet.columnCount + 1;
    headerRow.getCell(colNumber).value = header;
    if (getWidth) {
      worksheet.getColumn(colNumber).width = getWidth(header);
    }
    map.set(key, colNumber);
  }

  return map;
}

function clearDataRows(worksheet: ExcelJS.Worksheet, headerRowIndex: number): void {
  const rowsToDelete = Math.max(0, worksheet.rowCount - headerRowIndex);
  if (rowsToDelete > 0) {
    worksheet.spliceRows(headerRowIndex + 1, rowsToDelete);
  }
}

function getCol(headerMap: Map<string, number>, header: string): number {
  return headerMap.get(normalizeHeader(header)) || 1;
}

function findSheet(workbook: ExcelJS.Workbook, ...names: string[]): ExcelJS.Worksheet | undefined {
  const wanted = names.map((name) => normalizeHeader(name));
  return workbook.worksheets.find((sheet) => wanted.includes(normalizeHeader(sheet.name)));
}

function updateSheetTableRange(
  worksheet: ExcelJS.Worksheet,
  headerRowIndex: number,
  dataRowCount: number,
  requestedName?: string
): void {
  const tables = (worksheet.model as { tables?: Array<Record<string, any>> } | undefined)?.tables;
  if (!tables || tables.length === 0) return;

  const table = tables[0];
  const ref = table.tableRef;
  if (!ref) return;

  const match = ref.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
  if (!match) return;

  const startCol = match[1].toUpperCase();
  const endCol = match[3].toUpperCase();
  const endRow = Math.max(headerRowIndex + 1, headerRowIndex + Math.max(1, dataRowCount));
  table.tableRef = `${startCol}${headerRowIndex}:${endCol}${endRow}`;

  if (requestedName) {
    const safeName = requestedName.replace(/[^A-Za-z0-9_]/g, "_");
    table.name = safeName;
    table.displayName = safeName;
  }
}

function pickThirdPartyValue(
  headerLabel: string,
  index: number,
  invoice: InvoiceData,
  isSentDocuments: boolean
): ExcelJS.CellValue {
  const key = normalizeHeader(headerLabel);
  const party = isSentDocuments
    ? {
        nit: invoice.receiverNit,
        name: invoice.receiverName,
        email: invoice.receiverEmail || "N/A",
        phone: invoice.receiverPhone || "N/A",
        address: invoice.receiverAddress || "N/A",
        city: invoice.receiverCity || "N/A",
        department: invoice.receiverDepartment || "N/A",
        country: invoice.receiverCountry || "N/A",
        commercialName: invoice.receiverCommercialName || "N/A",
        taxpayerType: invoice.receiverTaxpayerType || "N/A",
        fiscalRegime: invoice.receiverFiscalRegime || "N/A",
        taxResponsibility: invoice.receiverTaxResponsibility || "N/A",
        economicActivity: invoice.receiverEconomicActivity || "N/A",
      }
    : {
        nit: invoice.issuerNit,
        name: invoice.issuerName,
        email: invoice.issuerEmail || "N/A",
        phone: invoice.issuerPhone || "N/A",
        address: invoice.issuerAddress || "N/A",
        city: invoice.issuerCity || "N/A",
        department: invoice.issuerDepartment || "N/A",
        country: invoice.issuerCountry || "N/A",
        commercialName: invoice.issuerCommercialName || "N/A",
        taxpayerType: invoice.issuerTaxpayerType || "N/A",
        fiscalRegime: invoice.issuerFiscalRegime || "N/A",
        taxResponsibility: invoice.issuerTaxResponsibility || "N/A",
        economicActivity: invoice.issuerEconomicActivity || "N/A",
      };

  if (["id", "item", "consecutivo", "indice"].some((token) => key === token || key.startsWith(token))) {
    return index + 1;
  }
  if (key.includes("numerofactura") || key.includes("nrofactura") || key.includes("numfactura")) return invoice.docNumber;
  if (key.includes("fechaemision") || key.includes("fechafactura")) return invoice.issueDate;
  if (key.includes("tipodocumento") && key.includes("factura")) return invoice.documentType;
  if (key.includes("cufe")) return invoice.cufe;
  if (key.includes("nit")) return party.nit;
  if (key.includes("razonsocial") || key.includes("nombretercero") || key.includes("nombreemisor") || key.includes("nombrevendedor") || key.includes("nombrereceptor")) return party.name;
  if (key.includes("nombrecomercial")) return party.commercialName;
  if (key.includes("tipodecontribuyente") || key.includes("tipocontribuyente")) return party.taxpayerType;
  if (key.includes("regimenfiscal")) return party.fiscalRegime;
  if (key.includes("responsabilidadtributaria")) return party.taxResponsibility;
  if (key.includes("actividadeconomica")) return party.economicActivity;
  if (key.includes("correo") || key.includes("email")) return party.email;
  if (key.includes("telefono") || key.includes("celular") || key.includes("movil")) return party.phone;
  if (key.includes("direccion")) return party.address;
  if (key.includes("ciudad") || key.includes("municipio") || key.includes("municipo")) return party.city;
  if (key.includes("departamento") || key.includes("provincia")) return party.department;
  if (key.includes("pais")) return party.country;
  if (key.includes("tipotercero") || key.includes("roltercero")) {
    return isSentDocuments ? "Receptor / Comprador" : "Emisor / Vendedor";
  }

  return "";
}

/**
 * Genera archivo Excel con los datos de las facturas
 * @param invoices - Lista de facturas a exportar
 * @param outputPath - Ruta donde guardar el archivo
 * @param includeDriveColumn - Si incluir columna de enlace a Drive
 * @param isSentDocuments - Si son documentos emitidos (muestra receptor en lugar de emisor)
 */
export async function generateExcelFile(
  invoices: InvoiceData[],
  outputPath: string,
  includeDriveColumn: boolean,
  isSentDocuments: boolean = false
): Promise<void> {
  const templatePath = resolveTemplatePath();

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);

  workbook.creator = "ContaGO";
  workbook.created = workbook.created || new Date();
  workbook.modified = new Date();
  workbook.lastPrinted = new Date();
  workbook.calcProperties = { fullCalcOnLoad: true };

  const sortedInvoices = sortInvoicesByDate(invoices);
  const allTaxTypes = collectAllTaxTypes(sortedInvoices);
  const headerRowIndex = getTemplateHeaderRow();

  const requestedMainSheetName = isSentDocuments ? "Facturas Emitidas" : "Facturas DIAN";
  const worksheet =
    findSheet(workbook, requestedMainSheetName) ||
    findSheet(workbook, "Facturas DIAN") ||
    workbook.worksheets[0];

  if (!worksheet) {
    throw new Error("La plantilla no contiene una hoja para el resumen de facturas.");
  }

  clearDataRows(worksheet, headerRowIndex);

  const partyLabel = isSentDocuments ? "Receptor" : "Emisor";
  const mainHeaders = [
    "ID",
    "Tipo de documento",
    "Numero Factura",
    `NIT ${partyLabel}`,
    `Razon Social ${partyLabel}`,
    "Fecha de emision",
    "Concepto",
    "Forma de pago",
    "Subtotal antes de impuestos",
    "Descuento detalle",
    "Recargo detalle",
    ...allTaxTypes.map((taxType) => `Valor ${taxType}`),
    "Descuento Global (-)",
    "Recargo Global (+)",
    "Valor total",
    ...(includeDriveColumn ? ["Enlace factura"] : []),
    "CUFE",
  ];

  const mainHeaderMap = ensureHeaders(worksheet, headerRowIndex, mainHeaders, (header) => {
    if (header === "ID") return 6;
    if (header === "Tipo de documento") return 18;
    if (header === "Numero Factura") return 20;
    if (header.startsWith("NIT ")) return 14;
    if (header.startsWith("Razon Social ")) return 40;
    if (header === "Fecha de emision") return 16;
    if (header === "Concepto") return 55;
    if (header === "Forma de pago") return 18;
    if (header === "Subtotal antes de impuestos") return 22;
    if (header.includes("Descuento")) return 18;
    if (header.includes("Recargo")) return 16;
    if (header.startsWith("Valor ")) return 14;
    if (header === "Enlace factura") return 45;
    if (header === "CUFE") return 100;
    return 14;
  });

  const currencyFmt = '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)';

  let mainRowNumber = headerRowIndex + 1;
  sortedInvoices.forEach((invoice, index) => {
    const partyNit = isSentDocuments ? invoice.receiverNit : invoice.issuerNit;
    const partyName = isSentDocuments ? invoice.receiverName : invoice.issuerName;

    const row = worksheet.getRow(mainRowNumber++);
    row.getCell(getCol(mainHeaderMap, "ID")).value = index + 1;
    row.getCell(getCol(mainHeaderMap, "Tipo de documento")).value = invoice.documentType;
    row.getCell(getCol(mainHeaderMap, "Numero Factura")).value = invoice.docNumber;
    row.getCell(getCol(mainHeaderMap, `NIT ${partyLabel}`)).value = partyNit;
    row.getCell(getCol(mainHeaderMap, `Razon Social ${partyLabel}`)).value = partyName;
    row.getCell(getCol(mainHeaderMap, "Fecha de emision")).value = invoice.issueDate;
    row.getCell(getCol(mainHeaderMap, "Concepto")).value = invoice.concepts;
    row.getCell(getCol(mainHeaderMap, "Forma de pago")).value = invoice.paymentMethod || "N/A";

    const subtotalCell = row.getCell(getCol(mainHeaderMap, "Subtotal antes de impuestos"));
    subtotalCell.value = typeof invoice.subtotal === "number" ? invoice.subtotal : 0;
    subtotalCell.numFmt = currencyFmt;

    const discountCell = row.getCell(getCol(mainHeaderMap, "Descuento detalle"));
    discountCell.value = invoice.discount || 0;
    discountCell.numFmt = currencyFmt;

    const surchargeCell = row.getCell(getCol(mainHeaderMap, "Recargo detalle"));
    surchargeCell.value = invoice.surcharge || 0;
    surchargeCell.numFmt = currencyFmt;

    for (const taxType of allTaxTypes) {
      const tax = (invoice.taxes || []).find((t) => t.taxName === taxType);
      const taxCell = row.getCell(getCol(mainHeaderMap, `Valor ${taxType}`));
      taxCell.value = tax ? tax.amount : 0;
      taxCell.numFmt = currencyFmt;
    }

    const globalDiscountCell = row.getCell(getCol(mainHeaderMap, "Descuento Global (-)"));
    globalDiscountCell.value = 0;
    globalDiscountCell.numFmt = currencyFmt;

    const globalSurchargeCell = row.getCell(getCol(mainHeaderMap, "Recargo Global (+)"));
    globalSurchargeCell.value = 0;
    globalSurchargeCell.numFmt = currencyFmt;

    const totalCell = row.getCell(getCol(mainHeaderMap, "Valor total"));
    totalCell.value = typeof invoice.total === "number" ? invoice.total : 0;
    totalCell.numFmt = currencyFmt;

    row.getCell(getCol(mainHeaderMap, "CUFE")).value = invoice.cufe;

    if (includeDriveColumn && invoice.driveUrl && !invoice.driveUrl.includes("ERROR")) {
      const driveCell = row.getCell(getCol(mainHeaderMap, "Enlace factura"));
      driveCell.value = { text: "Ver factura", hyperlink: invoice.driveUrl };
      driveCell.font = { color: { argb: "FF0066CC" }, underline: true };
    }

    if (invoice.error || invoice.cufe === "N/A") {
      row.eachCell((cell) => {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFFF3CD" },
        };
      });
    }

    row.alignment = { vertical: "middle", wrapText: true };
  });

  const lastRow = worksheet.rowCount;
  const cufeColIndex = getCol(mainHeaderMap, "CUFE");
  const cufeColLetter = getExcelColumnLetter(cufeColIndex);
  const dataStartRow = headerRowIndex + 1;

  if (lastRow >= dataStartRow) {
    worksheet.addConditionalFormatting({
      ref: `${cufeColLetter}${dataStartRow}:${cufeColLetter}${lastRow}`,
      rules: [
        {
          type: "expression",
          formulae: [`AND(${cufeColLetter}${dataStartRow}<>"N/A",COUNTIF($${cufeColLetter}$${dataStartRow}:$${cufeColLetter}$${lastRow},${cufeColLetter}${dataStartRow})>1)`],
          priority: 1,
          style: {
            fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFF6B6B" } },
            font: { color: { argb: "FFFFFFFF" }, bold: true },
          },
        },
      ],
    });
  }

  worksheet.autoFilter = {
    from: { row: headerRowIndex, column: 1 },
    to: { row: Math.max(headerRowIndex, lastRow), column: worksheet.columnCount },
  };
  updateSheetTableRange(worksheet, headerRowIndex, sortedInvoices.length);

  const detailedSheet =
    findSheet(workbook, "Detallado") ||
    findSheet(workbook, "Detalle") ||
    workbook.worksheets[1];

  if (!detailedSheet) {
    throw new Error("La plantilla no contiene una hoja para el detalle de conceptos.");
  }

  clearDataRows(detailedSheet, headerRowIndex);

  const detailedHeaders = [
    "Item",
    "Numero Factura",
    "Concepto",
    "Cantidad",
    "Base del impuesto",
    "Descuento detalle",
    "Recargo detalle",
    ...allTaxTypes.flatMap((taxType) => [taxType, `% ${taxType}`]),
    "Precio unitario (incluye impuestos)",
  ];

  const detailedHeaderMap = ensureHeaders(detailedSheet, headerRowIndex, detailedHeaders, (header) => {
    if (header === "Item") return 8;
    if (header === "Numero Factura") return 20;
    if (header === "Concepto") return 55;
    if (header === "Cantidad") return 12;
    if (header === "Base del impuesto") return 18;
    if (header.includes("Descuento")) return 16;
    if (header.includes("Recargo")) return 16;
    if (header.startsWith("% ")) return 10;
    if (header === "Precio unitario (incluye impuestos)") return 28;
    return 14;
  });

  const docNumberToRow = new Map<string, number>();
  sortedInvoices.forEach((invoice, index) => {
    docNumberToRow.set(invoice.docNumber, index + dataStartRow);
  });

  const percentFmt = '0.00"%"';

  let detailedRowsCount = 0;
  sortedInvoices.forEach((invoice) => {
    const mainSheetRow = docNumberToRow.get(invoice.docNumber) || dataStartRow;

    (invoice.lineItems || []).forEach((lineItem) => {
      const row = detailedSheet.getRow(detailedSheet.rowCount + 1);
      detailedRowsCount += 1;

      row.getCell(getCol(detailedHeaderMap, "Item")).value = lineItem.lineNumber;
      row.getCell(getCol(detailedHeaderMap, "Concepto")).value = lineItem.description;
      row.getCell(getCol(detailedHeaderMap, "Cantidad")).value = lineItem.quantity;

      const baseCell = row.getCell(getCol(detailedHeaderMap, "Base del impuesto"));
      baseCell.value = lineItem.totalUnitPrice;
      baseCell.numFmt = currencyFmt;

      const discountCell = row.getCell(getCol(detailedHeaderMap, "Descuento detalle"));
      discountCell.value = lineItem.discount;
      discountCell.numFmt = currencyFmt;

      const surchargeCell = row.getCell(getCol(detailedHeaderMap, "Recargo detalle"));
      surchargeCell.value = lineItem.surcharge;
      surchargeCell.numFmt = currencyFmt;

      for (const taxType of allTaxTypes) {
        const tax = (lineItem.taxes || []).find((t) => t.taxName === taxType);
        const amountCell = row.getCell(getCol(detailedHeaderMap, taxType));
        amountCell.value = tax ? tax.amount : 0;
        amountCell.numFmt = currencyFmt;

        const percentCell = row.getCell(getCol(detailedHeaderMap, `% ${taxType}`));
        percentCell.value = tax ? tax.percent : 0;
        percentCell.numFmt = percentFmt;
      }

      const totalTaxAmount = (lineItem.taxes || []).reduce((sum, t) => sum + t.amount, 0);
      const unitPriceCell = row.getCell(getCol(detailedHeaderMap, "Precio unitario (incluye impuestos)"));
      unitPriceCell.value = lineItem.totalUnitPrice + totalTaxAmount;
      unitPriceCell.numFmt = currencyFmt;

      const docNumberCell = row.getCell(getCol(detailedHeaderMap, "Numero Factura"));
      docNumberCell.value = {
        text: invoice.docNumber,
        hyperlink: `#'${worksheet.name}'!C${mainSheetRow}`,
      };
      docNumberCell.font = { color: { argb: "FF0066CC" }, underline: true };

      row.alignment = { vertical: "middle", wrapText: true };
    });
  });

  const detailedLastRow = detailedSheet.rowCount;
  detailedSheet.autoFilter = {
    from: { row: headerRowIndex, column: 1 },
    to: { row: Math.max(headerRowIndex, detailedLastRow), column: detailedSheet.columnCount },
  };
  updateSheetTableRange(detailedSheet, headerRowIndex, detailedRowsCount);

  const thirdSheet = findSheet(workbook, "Datos de terceros", "Datos terceros", "Terceros");
  if (thirdSheet) {
    clearDataRows(thirdSheet, headerRowIndex);
    const headerRow = thirdSheet.getRow(headerRowIndex);
    const headersByColumn = new Map<number, string>();

    headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      const label = cellText(cell.value);
      if (label) headersByColumn.set(colNumber, label);
    });

    sortedInvoices.forEach((invoice, index) => {
      const row = thirdSheet.getRow(headerRowIndex + 1 + index);
      for (const [col, label] of headersByColumn.entries()) {
        row.getCell(col).value = pickThirdPartyValue(label, index, invoice, isSentDocuments);
      }
      row.alignment = { vertical: "middle", wrapText: true };
    });

    const thirdLastRow = thirdSheet.rowCount;
    thirdSheet.autoFilter = {
      from: { row: headerRowIndex, column: 1 },
      to: { row: Math.max(headerRowIndex, thirdLastRow), column: thirdSheet.columnCount },
    };
    updateSheetTableRange(thirdSheet, headerRowIndex, sortedInvoices.length, "Datos de terceros");
  }

  await workbook.xlsx.writeFile(outputPath, {
    useStyles: true,
    useSharedStrings: true,
  });
}

/**
 * Genera nombre de archivo Excel basado en rango de fechas
 * @param startDate - Fecha inicio (YYYY-MM-DD)
 * @param endDate - Fecha fin (YYYY-MM-DD)
 * @param prefix - Prefijo del nombre de archivo (default: "Facturas DIAN")
 */
export function generateExcelFilename(startDate?: string, endDate?: string, prefix: string = "Facturas DIAN"): string {
  const formatDate = (date: string): string => {
    const [year, month, day] = date.split("-");
    const months = [
      "Ene", "Feb", "Mar", "Abr", "May", "Jun",
      "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"
    ];
    return `${months[parseInt(month) - 1]} ${parseInt(day)} ${year}`;
  };

  const start = startDate ? formatDate(startDate) : "Inicio";
  const end = endDate ? formatDate(endDate) : "Fin";

  return `${prefix} ${start} - ${end}.xlsx`;
}
