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
    { header: "EMPRESA/PN", key: "entityType", width: 12 },
    { header: "Fecha de emisión", key: "issueDate", width: 16 },
    { header: "Nombres Completos o Razón social", key: "entityName", width: 45 },
    { header: "Valor antes", key: "subtotal", width: 15 },
    { header: "Valor IVA (si aplica)", key: "iva", width: 18 },
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
      entityType: invoice.entityType,
      issueDate: invoice.issueDate,
      entityName: invoice.entityName,
      subtotal: invoice.subtotal,
      iva: invoice.iva,
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
  const cufeColLetter = includeDriveColumn ? "J" : "I";
  const lastRow = worksheet.rowCount;

  if (lastRow > 1) {
    // ExcelJS no soporta duplicateValues directamente, usamos fórmula COUNTIF
    // La fórmula excluye "N/A" de la detección de duplicados
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
