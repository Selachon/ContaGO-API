import fs from "fs";
import path from "path";
import ExcelJS from "exceljs";

const projectRoot = process.cwd();
const templatesDir = path.join(projectRoot, "templates");
const outputPath = path.join(templatesDir, "dian-excel-template.xlsx");

function ensureDir(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
}

function styleHeaderRow(row) {
  row.font = { bold: true, size: 11, color: { argb: "FF000000" } };
  row.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFD9D9D9" },
  };
  row.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  row.height = 25;
}

function addBrandRow(sheet, title) {
  sheet.mergeCells(1, 1, 1, Math.max(6, sheet.columnCount || 6));
  const titleCell = sheet.getCell(1, 1);
  titleCell.value = title;
  titleCell.font = { bold: true, size: 16, color: { argb: "FF1F2937" } };
  titleCell.alignment = { vertical: "middle", horizontal: "left" };
  sheet.getRow(1).height = 32;
}

function setHeadersOnRowTwo(sheet) {
  const row = sheet.getRow(2);
  sheet.columns.forEach((column, index) => {
    row.getCell(index + 1).value = column.header || "";
  });
}

async function buildTemplate() {
  ensureDir(templatesDir);

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "ContaGO";
  workbook.created = new Date();
  workbook.modified = new Date();

  const mainSheet = workbook.addWorksheet("Facturas DIAN", {
    views: [{ state: "frozen", ySplit: 2, xSplit: 0 }],
    properties: {
      defaultColWidth: 12,
      defaultRowHeight: 15,
      tabColor: { argb: "FF4472C4" },
    },
  });

  mainSheet.columns = [
    { header: "ID", key: "id", width: 6 },
    { header: "Tipo de documento", key: "documentType", width: 18 },
    { header: "Numero Factura", key: "docNumber", width: 20 },
    { header: "NIT Emisor", key: "partyNit", width: 14 },
    { header: "Razon Social Emisor", key: "partyName", width: 40 },
    { header: "Fecha de emision", key: "issueDate", width: 16 },
    { header: "Concepto", key: "concepts", width: 55 },
    { header: "Forma de pago", key: "paymentMethod", width: 18 },
    { header: "Subtotal antes de impuestos", key: "subtotal", width: 22 },
    { header: "Descuento detalle", key: "discount", width: 16 },
    { header: "Recargo detalle", key: "surcharge", width: 16 },
    { header: "Valor IVA", key: "tax_IVA", width: 14 },
    { header: "Valor INC", key: "tax_INC", width: 14 },
    { header: "Valor Bolsas", key: "tax_Bolsas", width: 14 },
    { header: "Valor ICUI", key: "tax_ICUI", width: 14 },
    { header: "Valor IC", key: "tax_IC", width: 14 },
    { header: "Descuento Global (-)", key: "globalDiscount", width: 18 },
    { header: "Recargo Global (+)", key: "globalSurcharge", width: 16 },
    { header: "Valor total", key: "total", width: 15 },
    { header: "Enlace factura", key: "driveUrl", width: 45 },
    { header: "CUFE", key: "cufe", width: 100 },
  ];
  addBrandRow(mainSheet, "ContaGO - Exportador DIAN");
  setHeadersOnRowTwo(mainSheet);
  styleHeaderRow(mainSheet.getRow(2));

  mainSheet.addRow({
    id: 1,
    documentType: "Factura Electronica de Venta",
    docNumber: "FV-1001",
    partyNit: "900123456",
    partyName: "Empresa Ejemplo SAS",
    issueDate: "2026-04-27",
    concepts: "Servicio contable mensual",
    paymentMethod: "Contado",
    subtotal: 1000000,
    discount: 0,
    surcharge: 0,
    tax_IVA: 190000,
    tax_INC: 0,
    tax_Bolsas: 0,
    tax_ICUI: 0,
    tax_IC: 0,
    globalDiscount: 0,
    globalSurcharge: 0,
    total: 1190000,
    driveUrl: "https://drive.google.com/file/d/EJEMPLO/view",
    cufe: "e4f91de5-EXAMPLE-CUFE",
  });

  const detailSheet = workbook.addWorksheet("Detallado", {
    views: [{ state: "frozen", ySplit: 2, xSplit: 0 }],
    properties: {
      defaultColWidth: 12,
      defaultRowHeight: 15,
      tabColor: { argb: "FF70AD47" },
    },
  });

  detailSheet.columns = [
    { header: "Item", key: "lineNumber", width: 8 },
    { header: "Numero Factura", key: "docNumber", width: 20 },
    { header: "Concepto", key: "description", width: 55 },
    { header: "Cantidad", key: "quantity", width: 12 },
    { header: "Base del impuesto", key: "totalUnitPrice", width: 18 },
    { header: "Descuento detalle", key: "discount", width: 16 },
    { header: "Recargo detalle", key: "surcharge", width: 16 },
    { header: "IVA", key: "tax_IVA_amount", width: 14 },
    { header: "% IVA", key: "tax_IVA_percent", width: 10 },
    { header: "INC", key: "tax_INC_amount", width: 14 },
    { header: "% INC", key: "tax_INC_percent", width: 10 },
    { header: "Bolsas", key: "tax_Bolsas_amount", width: 14 },
    { header: "% Bolsas", key: "tax_Bolsas_percent", width: 10 },
    { header: "ICUI", key: "tax_ICUI_amount", width: 14 },
    { header: "% ICUI", key: "tax_ICUI_percent", width: 10 },
    { header: "IC", key: "tax_IC_amount", width: 14 },
    { header: "% IC", key: "tax_IC_percent", width: 10 },
    { header: "Precio unitario (incluye impuestos)", key: "unitPriceWithTax", width: 28 },
  ];
  addBrandRow(detailSheet, "ContaGO - Detallado de Conceptos");
  setHeadersOnRowTwo(detailSheet);
  styleHeaderRow(detailSheet.getRow(2));

  detailSheet.addRow({
    lineNumber: 1,
    docNumber: "FV-1001",
    description: "Servicio contable mensual",
    quantity: 1,
    totalUnitPrice: 1000000,
    discount: 0,
    surcharge: 0,
    tax_IVA_amount: 190000,
    tax_IVA_percent: 19,
    tax_INC_amount: 0,
    tax_INC_percent: 0,
    tax_Bolsas_amount: 0,
    tax_Bolsas_percent: 0,
    tax_ICUI_amount: 0,
    tax_ICUI_percent: 0,
    tax_IC_amount: 0,
    tax_IC_percent: 0,
    unitPriceWithTax: 1190000,
  });

  const thirdSheet = workbook.addWorksheet("Datos de terceros", {
    views: [{ state: "frozen", ySplit: 2, xSplit: 0 }],
    properties: {
      defaultColWidth: 18,
      defaultRowHeight: 15,
      tabColor: { argb: "FFF59E0B" },
    },
  });

  thirdSheet.columns = [
    { header: "ID", key: "id", width: 8 },
    { header: "Numero Factura", key: "docNumber", width: 20 },
    { header: "Tipo tercero", key: "thirdType", width: 20 },
    { header: "NIT", key: "nit", width: 16 },
    { header: "Razon Social", key: "name", width: 40 },
    { header: "Correo", key: "email", width: 32 },
    { header: "Telefono", key: "phone", width: 20 },
    { header: "Direccion", key: "address", width: 45 },
    { header: "Ciudad", key: "city", width: 20 },
    { header: "Departamento", key: "department", width: 20 },
    { header: "Pais", key: "country", width: 20 },
    { header: "Fecha de emision", key: "issueDate", width: 18 },
    { header: "CUFE", key: "cufe", width: 80 },
  ];
  addBrandRow(thirdSheet, "ContaGO - Datos de terceros");
  setHeadersOnRowTwo(thirdSheet);
  styleHeaderRow(thirdSheet.getRow(2));

  thirdSheet.addRow({
    id: 1,
    docNumber: "FV-1001",
    thirdType: "Emisor / Vendedor",
    nit: "900123456",
    name: "Empresa Ejemplo SAS",
    email: "facturacion@ejemplo.com",
    phone: "6011234567",
    address: "Calle 100 # 10-20",
    city: "Bogota",
    department: "Cundinamarca",
    country: "Colombia",
    issueDate: "2026-04-27",
    cufe: "e4f91de5-EXAMPLE-CUFE",
  });

  await workbook.xlsx.writeFile(outputPath, {
    useStyles: true,
    useSharedStrings: true,
  });

  console.log(`Template generado en: ${outputPath}`);
}

buildTemplate().catch((err) => {
  console.error("No se pudo generar la plantilla:", err);
  process.exit(1);
});
