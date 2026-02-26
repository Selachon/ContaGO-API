import ExcelJS from "exceljs";
import type { InvoiceData } from "../types/dianExcel.js";

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

  const worksheet = workbook.addWorksheet("Facturas DIAN", {
    views: [{ state: "frozen", ySplit: 1 }], // Congelar primera fila
  });

  // Definir columnas
  const columns: Partial<ExcelJS.Column>[] = [
    { header: "ID", key: "id", width: 6 },
    { header: "Número Factura", key: "docNumber", width: 20 },
    { header: "NIT Emisor", key: "issuerNit", width: 14 },
    { header: "Razón Social Emisor", key: "issuerName", width: 40 },
    { header: "Fecha de emisión", key: "issueDate", width: 16 },
    { header: "Valor antes", key: "subtotal", width: 15 },
    { header: "Valor IVA (si aplica)", key: "iva", width: 18 },
    { header: "Valor total", key: "total", width: 15 },
    { header: "Concepto", key: "concepts", width: 55 },
  ];

  if (includeDriveColumn) {
    columns.push({ header: "Adjunte factura", key: "driveUrl", width: 45 });
  }

  columns.push(
    { header: "Tipo de documento", key: "documentType", width: 20 },
    { header: "CUFE", key: "cufe", width: 100 }
  );

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

  // Agregar filas de datos
  invoices.forEach((invoice, index) => {
    const rowData: Record<string, unknown> = {
      id: index + 1,
      docNumber: invoice.docNumber,
      issuerNit: invoice.issuerNit,
      issuerName: invoice.issuerName,
      issueDate: invoice.issueDate,
      subtotal: invoice.subtotal,
      iva: invoice.iva,
      total: invoice.total,
      concepts: invoice.concepts,
      documentType: invoice.documentType,
      cufe: invoice.cufe,
    };

    if (includeDriveColumn) {
      rowData.driveUrl = invoice.driveUrl || "";
    }

    const row = worksheet.addRow(rowData);

    // Formato de celdas numéricas
    const subtotalCell = row.getCell("subtotal");
    subtotalCell.numFmt = '"$"#,##0.00';

    const ivaCell = row.getCell("iva");
    ivaCell.numFmt = '"$"#,##0.00';

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
  // Columna CUFE: con Drive es L, sin Drive es K
  const cufeColLetter = includeDriveColumn ? "L" : "K";
  const lastRow = worksheet.rowCount;

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

  detailedSheet.columns = [
    { header: "Item", key: "lineNumber", width: 8 },
    { header: "Número Factura", key: "docNumber", width: 20 },
    { header: "Descripción", key: "description", width: 55 },
    { header: "Cantidad", key: "quantity", width: 12 },
    { header: "Precio unitario", key: "unitPrice", width: 16 },
    { header: "Descuento detalle", key: "discount", width: 18 },
    { header: "Recargo detalle", key: "surcharge", width: 16 },
    { header: "IVA", key: "ivaAmount", width: 14 },
    { header: "%", key: "ivaPercent", width: 10 },
    { header: "INC", key: "incAmount", width: 14 },
    { header: "%", key: "incPercent", width: 10 },
    { header: "Precio unitario de venta", key: "totalUnitPrice", width: 24 },
  ];

  const detailedHeaderRow = detailedSheet.getRow(1);
  detailedHeaderRow.font = { bold: true, size: 11, color: { argb: "FF000000" } };
  detailedHeaderRow.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9D9D9" },
  };
  detailedHeaderRow.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  detailedHeaderRow.height = 25;

  invoices.forEach((invoice, invoiceIndex) => {
    const mainSheetRow = invoiceIndex + 2;

    (invoice.lineItems || []).forEach((lineItem) => {
      const row = detailedSheet.addRow({
        lineNumber: lineItem.lineNumber,
        docNumber: invoice.docNumber,
        description: lineItem.description,
        quantity: lineItem.quantity,
        unitPrice: lineItem.unitPrice,
        discount: lineItem.discount,
        surcharge: lineItem.surcharge,
        ivaAmount: lineItem.ivaAmount,
        ivaPercent: lineItem.ivaPercent,
        incAmount: lineItem.incAmount,
        incPercent: lineItem.incPercent,
        totalUnitPrice: lineItem.totalUnitPrice,
      });

      const docNumberCell = row.getCell("docNumber");
      docNumberCell.value = {
        text: invoice.docNumber,
        hyperlink: `#'Facturas DIAN'!B${mainSheetRow}`,
      };
      docNumberCell.font = { color: { argb: "FF0066CC" }, underline: true };

      row.getCell("unitPrice").numFmt = '"$"#,##0.00';
      row.getCell("discount").numFmt = '"$"#,##0.00';
      row.getCell("surcharge").numFmt = '"$"#,##0.00';
      row.getCell("ivaAmount").numFmt = '"$"#,##0.00';
      row.getCell("incAmount").numFmt = '"$"#,##0.00';
      row.getCell("totalUnitPrice").numFmt = '"$"#,##0.00';
      row.getCell("ivaPercent").numFmt = '0.00';
      row.getCell("incPercent").numFmt = '0.00';

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
