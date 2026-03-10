import ExcelJS from "exceljs";
import type { InvoiceData } from "../types/dianExcel.js";

/**
 * Convierte un número de columna (1-indexed) a letra de Excel
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
 * Recolecta todos los tipos de impuestos únicos de todas las facturas
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

    // Impuestos a nivel de línea
    for (const line of invoice.lineItems || []) {
      for (const tax of line.taxes || []) {
        if (tax.taxName && tax.taxName !== "IVA") {
          taxTypesSet.add(tax.taxName);
        }
      }
    }
  }

  // Ordenar: IVA primero, luego INC, luego Bolsas, luego el resto alfabéticamente
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
 * Ordena las facturas por fecha de emisión en orden ascendente
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
 */
export async function generateExcelFile(
  invoices: InvoiceData[],
  outputPath: string,
  includeDriveColumn: boolean
): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "ContaGO";
  workbook.created = new Date();

  // Ordenar facturas por fecha de emisión ascendente
  const sortedInvoices = sortInvoicesByDate(invoices);

  // Recolectar todos los tipos de impuestos
  const allTaxTypes = collectAllTaxTypes(sortedInvoices);

  const worksheet = workbook.addWorksheet("Facturas DIAN", {
    views: [{ state: "frozen", ySplit: 1 }], // Congelar primera fila
  });

  // Definir columnas base según la nueva plantilla
  const columns: Partial<ExcelJS.Column>[] = [
    { header: "ID", key: "id", width: 6 },
    { header: "Tipo de documento", key: "documentType", width: 18 },
    { header: "Número Factura", key: "docNumber", width: 20 },
    { header: "NIT Emisor", key: "issuerNit", width: 14 },
    { header: "Razón Social Emisor", key: "issuerName", width: 40 },
    { header: "Fecha de emisión", key: "issueDate", width: 16 },
    { header: "Concepto", key: "concepts", width: 55 },
    { header: "Subtotal antes de impuestos", key: "subtotal", width: 22 },
    { header: "Descuento detalle", key: "discount", width: 16 },
    { header: "Recargo detalle", key: "surcharge", width: 16 },
  ];

  // Agregar columnas dinámicas de impuestos (Valor X)
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

  // Agregar filas de datos (ya ordenadas por fecha)
  sortedInvoices.forEach((invoice, index) => {
    const rowData: Record<string, unknown> = {
      id: index + 1, // ID basado en orden por fecha
      documentType: invoice.documentType,
      docNumber: invoice.docNumber,
      issuerNit: invoice.issuerNit,
      issuerName: invoice.issuerName,
      issueDate: invoice.issueDate,
      concepts: invoice.concepts,
      subtotal: invoice.subtotal,
      discount: invoice.discount || 0,
      surcharge: invoice.surcharge || 0,
      globalDiscount: 0, // Por ahora 0, se puede expandir si se necesita
      globalSurcharge: 0,
      total: invoice.total,
      cufe: invoice.cufe,
    };

    // Agregar valores de impuestos dinámicos
    for (const taxType of allTaxTypes) {
      const tax = (invoice.taxes || []).find(t => t.taxName === taxType);
      rowData[`tax_${taxType}`] = tax ? tax.amount : 0;
    }

    if (includeDriveColumn) {
      rowData.driveUrl = invoice.driveUrl || "";
    }

    const row = worksheet.addRow(rowData);

    // Formato de celdas numéricas
    const subtotalCell = row.getCell("subtotal");
    subtotalCell.numFmt = '"$"#,##0.00';

    row.getCell("discount").numFmt = '"$"#,##0.00';
    row.getCell("surcharge").numFmt = '"$"#,##0.00';

    // Formato para columnas de impuestos dinámicos
    for (const taxType of allTaxTypes) {
      const taxCell = row.getCell(`tax_${taxType}`);
      taxCell.numFmt = '"$"#,##0.00';
    }

    row.getCell("globalDiscount").numFmt = '"$"#,##0.00';
    row.getCell("globalSurcharge").numFmt = '"$"#,##0.00';

    const totalCell = row.getCell("total");
    totalCell.numFmt = '"$"#,##0.00';

    // Hipervínculo en columna Drive
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
  // Columna CUFE es la última columna
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

  // Agregar filtros automáticos
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

  // Hoja 2: Detallado por concepto/línea
  const detailedSheet = workbook.addWorksheet("Detallado", {
    views: [{ state: "frozen", ySplit: 1 }],
  });

  // Columnas base para la hoja detallada
  const detailedColumns: Partial<ExcelJS.Column>[] = [
    { header: "Item", key: "lineNumber", width: 8 },
    { header: "Número Factura", key: "docNumber", width: 20 },
    { header: "Concepto", key: "description", width: 55 },
    { header: "Cantidad", key: "quantity", width: 12 },
    { header: "Base del impuesto", key: "totalUnitPrice", width: 18 },
    { header: "Descuento detalle", key: "discount", width: 16 },
    { header: "Recargo detalle", key: "surcharge", width: 16 },
  ];

  // Agregar columnas dinámicas de impuestos (Valor y %)
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

  // Crear un mapa de docNumber a su fila en la hoja principal (después de ordenar)
  const docNumberToRow = new Map<string, number>();
  sortedInvoices.forEach((invoice, index) => {
    docNumberToRow.set(invoice.docNumber, index + 2);
  });

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

      // Agregar valores de impuestos dinámicos por línea
      for (const taxType of allTaxTypes) {
        const tax = (lineItem.taxes || []).find(t => t.taxName === taxType);
        rowData[`tax_${taxType}_amount`] = tax ? tax.amount : 0;
        rowData[`tax_${taxType}_percent`] = tax ? tax.percent : 0;
      }

      // Calcular precio unitario con impuestos
      const totalTaxAmount = (lineItem.taxes || []).reduce((sum, t) => sum + t.amount, 0);
      rowData.unitPriceWithTax = lineItem.totalUnitPrice + totalTaxAmount;

      const row = detailedSheet.addRow(rowData);

      // Hipervínculo al número de factura
      const docNumberCell = row.getCell("docNumber");
      docNumberCell.value = {
        text: invoice.docNumber,
        hyperlink: `#'Facturas DIAN'!C${mainSheetRow}`,
      };
      docNumberCell.font = { color: { argb: "FF0066CC" }, underline: true };

      // Formato de celdas numéricas
      row.getCell("totalUnitPrice").numFmt = '"$"#,##0.00';
      row.getCell("discount").numFmt = '"$"#,##0.00';
      row.getCell("surcharge").numFmt = '"$"#,##0.00';
      row.getCell("unitPriceWithTax").numFmt = '"$"#,##0.00';

      // Formato para columnas de impuestos dinámicos
      for (const taxType of allTaxTypes) {
        const amountCell = row.getCell(`tax_${taxType}_amount`);
        amountCell.numFmt = '"$"#,##0.00';
        const percentCell = row.getCell(`tax_${taxType}_percent`);
        percentCell.numFmt = '0.00';
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

  // Guardar archivo
  await workbook.xlsx.writeFile(outputPath);
}

/**
 * Genera nombre de archivo Excel basado en rango de fechas
 */
export function generateExcelFilename(startDate?: string, endDate?: string): string {
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

  return `Facturas DIAN ${start} - ${end}.xlsx`;
}
