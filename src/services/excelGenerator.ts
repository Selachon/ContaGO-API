import ExcelJS from "exceljs";
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
  const workbook = new ExcelJS.Workbook();
  
  // Propiedades del workbook para mejor compatibilidad con Excel y Sheets
  workbook.creator = "ContaGO";
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.lastPrinted = new Date();
  
  // Establecer propiedades de calculo para compatibilidad
  workbook.calcProperties = {
    fullCalcOnLoad: true,
  };

  // Ordenar facturas por fecha de emision ascendente
  const sortedInvoices = sortInvoicesByDate(invoices);

  // Recolectar todos los tipos de impuestos
  const allTaxTypes = collectAllTaxTypes(sortedInvoices);

  // Nombre de la hoja segun el tipo de documentos
  const sheetName = isSentDocuments ? "Facturas Emitidas" : "Facturas DIAN";
  
  const worksheet = workbook.addWorksheet(sheetName, {
    views: [{ state: "frozen", ySplit: 1, xSplit: 0 }], // Congelar primera fila
    properties: {
      defaultColWidth: 12,
      defaultRowHeight: 15,
      tabColor: { argb: "FF4472C4" },
    },
  });

  // Definir columnas - Para emitidos mostramos el receptor, para recibidos el emisor
  const partyLabel = isSentDocuments ? "Receptor" : "Emisor";
  
  const columns: Partial<ExcelJS.Column>[] = [
    { header: "ID", key: "id", width: 6 },
    { header: "Tipo de documento", key: "documentType", width: 18 },
    { header: "Numero Factura", key: "docNumber", width: 20 },
    { header: `NIT ${partyLabel}`, key: "partyNit", width: 14 },
    { header: `Razon Social ${partyLabel}`, key: "partyName", width: 40 },
    { header: "Fecha de emision", key: "issueDate", width: 16 },
    { header: "Concepto", key: "concepts", width: 55 },
    { header: "Forma de pago", key: "paymentMethod", width: 18 },
    { header: "Subtotal antes de impuestos", key: "subtotal", width: 22 },
    { header: "Descuento detalle", key: "discount", width: 16 },
    { header: "Recargo detalle", key: "surcharge", width: 16 },
  ];

  // Agregar columnas dinamicas de impuestos (Valor X)
  for (const taxType of allTaxTypes) {
    columns.push({
      header: `Valor ${taxType}`,
      key: `tax_${taxType}`,
      width: 14,
    });
  }

  // Columnas finales
  columns.push(
    { header: "Descuento Global (-)", key: "globalDiscount", width: 18 },
    { header: "Recargo Global (+)", key: "globalSurcharge", width: 16 },
    { header: "Valor total", key: "total", width: 15 }
  );

  if (includeDriveColumn) {
    columns.push({ header: "Enlace factura", key: "driveUrl", width: 45 });
  }

  columns.push({ header: "CUFE", key: "cufe", width: 100 });

  worksheet.columns = columns;

  // Estilos del header
  const headerRow = worksheet.getRow(1);
  headerRow.font = { bold: true, size: 11, color: { argb: "FF000000" } };
  headerRow.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9D9D9" },
  };
  headerRow.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  headerRow.height = 25;

  // Formato numerico compatible con Excel y Sheets
  const currencyFmt = '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)';

  // Agregar filas de datos (ya ordenadas por fecha)
  sortedInvoices.forEach((invoice, index) => {
    // Para emitidos, mostramos datos del receptor; para recibidos, del emisor
    const partyNit = isSentDocuments ? invoice.receiverNit : invoice.issuerNit;
    const partyName = isSentDocuments ? invoice.receiverName : invoice.issuerName;
    
    const rowData: Record<string, unknown> = {
      id: index + 1, // ID basado en orden por fecha
      documentType: invoice.documentType,
      docNumber: invoice.docNumber,
      partyNit,
      partyName,
      issueDate: invoice.issueDate,
      concepts: invoice.concepts,
      paymentMethod: invoice.paymentMethod || "N/A",
      subtotal: invoice.subtotal,
      discount: invoice.discount || 0,
      surcharge: invoice.surcharge || 0,
      globalDiscount: 0,
      globalSurcharge: 0,
      total: invoice.total,
      cufe: invoice.cufe,
    };

    // Agregar valores de impuestos dinamicos
    for (const taxType of allTaxTypes) {
      const tax = (invoice.taxes || []).find(t => t.taxName === taxType);
      rowData[`tax_${taxType}`] = tax ? tax.amount : 0;
    }

    if (includeDriveColumn) {
      rowData.driveUrl = invoice.driveUrl || "";
    }

    const row = worksheet.addRow(rowData);

    // Formato de celdas numericas
    const subtotalCell = row.getCell("subtotal");
    subtotalCell.numFmt = currencyFmt;
    subtotalCell.value = typeof invoice.subtotal === "number" ? invoice.subtotal : 0;

    row.getCell("discount").numFmt = currencyFmt;
    row.getCell("surcharge").numFmt = currencyFmt;

    // Formato para columnas de impuestos dinamicos
    for (const taxType of allTaxTypes) {
      const taxCell = row.getCell(`tax_${taxType}`);
      taxCell.numFmt = currencyFmt;
    }

    row.getCell("globalDiscount").numFmt = currencyFmt;
    row.getCell("globalSurcharge").numFmt = currencyFmt;

    const totalCell = row.getCell("total");
    totalCell.numFmt = currencyFmt;
    totalCell.value = typeof invoice.total === "number" ? invoice.total : 0;

    // Hipervinculo en columna Drive
    if (includeDriveColumn && invoice.driveUrl && !invoice.driveUrl.includes("ERROR")) {
      const driveCell = row.getCell("driveUrl");
      driveCell.value = {
        text: "Ver factura",
        hyperlink: invoice.driveUrl,
      };
      driveCell.font = { color: { argb: "FF0066CC" }, underline: true };
    }

    // Marcar errores
    if (invoice.error || invoice.cufe === "N/A") {
      row.eachCell((cell) => {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFFF3CD" }, // Amarillo suave
        };
      });
    }

    // Alignment
    row.alignment = { vertical: "middle", wrapText: true };
  });

  // Formato condicional para CUFEs duplicados (excluyendo "N/A")
  const lastRow = worksheet.rowCount;
  const cufeColIndex = columns.length;
  const cufeColLetter = getExcelColumnLetter(cufeColIndex);

  if (lastRow > 1) {
    worksheet.addConditionalFormatting({
      ref: `${cufeColLetter}2:${cufeColLetter}${lastRow}`,
      rules: [
        {
          type: "expression",
          formulae: [`AND(${cufeColLetter}2<>"N/A",COUNTIF($${cufeColLetter}$2:$${cufeColLetter}$${lastRow},${cufeColLetter}2)>1)`],
          priority: 1,
          style: {
            fill: {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FFFF6B6B" }, // Rojo
            },
            font: {
              color: { argb: "FFFFFFFF" }, // Blanco
              bold: true,
            },
          },
        },
      ],
    });
  }

  // Agregar filtros automaticos
  worksheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: lastRow, column: columns.length },
  };

  // Bordes para todas las celdas con datos
  for (let row = 1; row <= lastRow; row++) {
    for (let col = 1; col <= columns.length; col++) {
      const cell = worksheet.getCell(row, col);
      cell.border = {
        top: { style: "thin", color: { argb: "FFD0D0D0" } },
        left: { style: "thin", color: { argb: "FFD0D0D0" } },
        bottom: { style: "thin", color: { argb: "FFD0D0D0" } },
        right: { style: "thin", color: { argb: "FFD0D0D0" } },
      };
    }
  }

  // Hoja 2: Detallado por concepto/linea
  const detailedSheet = workbook.addWorksheet("Detallado", {
    views: [{ state: "frozen", ySplit: 1, xSplit: 0 }],
    properties: {
      defaultColWidth: 12,
      defaultRowHeight: 15,
      tabColor: { argb: "FF70AD47" },
    },
  });

  // Columnas base para la hoja detallada
  const detailedColumns: Partial<ExcelJS.Column>[] = [
    { header: "Item", key: "lineNumber", width: 8 },
    { header: "Numero Factura", key: "docNumber", width: 20 },
    { header: "Concepto", key: "description", width: 55 },
    { header: "Cantidad", key: "quantity", width: 12 },
    { header: "Base del impuesto", key: "totalUnitPrice", width: 18 },
    { header: "Descuento detalle", key: "discount", width: 16 },
    { header: "Recargo detalle", key: "surcharge", width: 16 },
  ];

  // Agregar columnas dinamicas de impuestos (Valor y %)
  for (const taxType of allTaxTypes) {
    detailedColumns.push(
      { header: taxType, key: `tax_${taxType}_amount`, width: 14 },
      { header: `% ${taxType}`, key: `tax_${taxType}_percent`, width: 10 }
    );
  }

  // Columna final
  detailedColumns.push({ header: "Precio unitario (incluye impuestos)", key: "unitPriceWithTax", width: 28 });

  detailedSheet.columns = detailedColumns;

  const detailedHeaderRow = detailedSheet.getRow(1);
  detailedHeaderRow.font = { bold: true, size: 11, color: { argb: "FF000000" } };
  detailedHeaderRow.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9D9D9" },
  };
  detailedHeaderRow.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  detailedHeaderRow.height = 25;

  // Crear un mapa de docNumber a su fila en la hoja principal (despues de ordenar)
  const docNumberToRow = new Map<string, number>();
  sortedInvoices.forEach((invoice, index) => {
    docNumberToRow.set(invoice.docNumber, index + 2);
  });

  const percentFmt = '0.00"%"';

  sortedInvoices.forEach((invoice) => {
    const mainSheetRow = docNumberToRow.get(invoice.docNumber) || 2;

    (invoice.lineItems || []).forEach((lineItem) => {
      const rowData: Record<string, unknown> = {
        lineNumber: lineItem.lineNumber,
        docNumber: invoice.docNumber,
        description: lineItem.description,
        quantity: lineItem.quantity,
        totalUnitPrice: lineItem.totalUnitPrice,
        discount: lineItem.discount,
        surcharge: lineItem.surcharge,
      };

      // Agregar valores de impuestos dinamicos por linea
      for (const taxType of allTaxTypes) {
        const tax = (lineItem.taxes || []).find(t => t.taxName === taxType);
        rowData[`tax_${taxType}_amount`] = tax ? tax.amount : 0;
        rowData[`tax_${taxType}_percent`] = tax ? tax.percent : 0;
      }

      // Calcular precio unitario con impuestos
      const totalTaxAmount = (lineItem.taxes || []).reduce((sum, t) => sum + t.amount, 0);
      rowData.unitPriceWithTax = lineItem.totalUnitPrice + totalTaxAmount;

      const row = detailedSheet.addRow(rowData);

      // Hipervinculo al numero de factura
      const docNumberCell = row.getCell("docNumber");
      docNumberCell.value = {
        text: invoice.docNumber,
        hyperlink: `#'${sheetName}'!C${mainSheetRow}`,
      };
      docNumberCell.font = { color: { argb: "FF0066CC" }, underline: true };

      // Formato de celdas numericas
      row.getCell("totalUnitPrice").numFmt = currencyFmt;
      row.getCell("discount").numFmt = currencyFmt;
      row.getCell("surcharge").numFmt = currencyFmt;
      row.getCell("unitPriceWithTax").numFmt = currencyFmt;

      // Formato para columnas de impuestos dinamicos
      for (const taxType of allTaxTypes) {
        const amountCell = row.getCell(`tax_${taxType}_amount`);
        amountCell.numFmt = currencyFmt;
        const percentCell = row.getCell(`tax_${taxType}_percent`);
        percentCell.numFmt = percentFmt;
      }

      row.alignment = { vertical: "middle", wrapText: true };
    });
  });

  const detailedLastRow = detailedSheet.rowCount;
  detailedSheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: detailedLastRow, column: detailedSheet.columnCount },
  };

  for (let row = 1; row <= detailedLastRow; row++) {
    for (let col = 1; col <= detailedSheet.columnCount; col++) {
      const cell = detailedSheet.getCell(row, col);
      cell.border = {
        top: { style: "thin", color: { argb: "FFD0D0D0" } },
        left: { style: "thin", color: { argb: "FFD0D0D0" } },
        bottom: { style: "thin", color: { argb: "FFD0D0D0" } },
        right: { style: "thin", color: { argb: "FFD0D0D0" } },
      };
    }
  }

  // Guardar archivo con opciones para mejor compatibilidad
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
