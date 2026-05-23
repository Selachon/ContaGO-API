import ExcelJS from "exceljs";
import {
  listInvoices,
  listPurchases,
  listPurchaseSupportDocuments,
  listPaymentReceipts,
  getPurchaseById,
  getPurchaseSupportDocumentById,
  getPaymentReceiptById,
  getInvoiceById,
} from "./siigoService.js";

export interface CrossRefItem {
  originalName: string;
  status: "pending" | "found" | "not_found" | "error";
  type?: string;
  siigoId?: string;
  siigoName?: string;
  date?: string;
  supplierId?: string;
  supplierName?: string;
  totalValue?: number;
  observations?: string;
  relatedDoc?: string;
  notes?: string;
}

export interface BulkProcessResult {
  excelBuffer: Buffer;
  summary: {
    total: number;
    found: number;
    notFound: number;
    errors: number;
  };
}

async function delay(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function findDocumentData(name: string): Promise<Partial<CrossRefItem> | null> {
  const cleanName = name.trim();
  if (!cleanName) return null;

  const upperName = cleanName.toUpperCase();
  console.log(`[CrossRef] Searching for "${cleanName}"...`);

  const searchPlan: Array<{ type: string; action: () => Promise<any> }> = [];

  if (upperName.startsWith("FV") || upperName.startsWith("FE")) {
    searchPlan.push({ type: "Factura Venta", action: () => listInvoices({ name: cleanName, page: 1, page_size: 10 }) });
  } else if (upperName.startsWith("DS")) {
    searchPlan.push({ type: "Documento Soporte", action: () => listPurchaseSupportDocuments({ name: cleanName, page: 1, page_size: 10 }) });
  } else if (upperName.startsWith("FC") || upperName.startsWith("PC")) {
    searchPlan.push({ type: "Factura Compra", action: () => listPurchases({ name: cleanName, page: 1, page_size: 10 }) });
  } else if (upperName.startsWith("RP") || upperName.startsWith("RC") || upperName.startsWith("EG")) {
    searchPlan.push({ type: "Recibo Pago", action: () => listPaymentReceipts({ name: cleanName, page: 1, page_size: 10 }) });
  } else {
    searchPlan.push({ type: "Factura Venta", action: () => listInvoices({ name: cleanName, page: 1, page_size: 10 }) });
    searchPlan.push({ type: "Documento Soporte", action: () => listPurchaseSupportDocuments({ name: cleanName, page: 1, page_size: 10 }) });
    searchPlan.push({ type: "Factura Compra", action: () => listPurchases({ name: cleanName, page: 1, page_size: 10 }) });
    searchPlan.push({ type: "Recibo Pago", action: () => listPaymentReceipts({ name: cleanName, page: 1, page_size: 10 }) });
  }

  for (const step of searchPlan) {
    try {
      const raw = await step.action();
      const results = (raw as any)?.results || [];
      const match = results.find((doc: any) => {
        const docName = String(doc.name || "").toUpperCase();
        const docAltName = `${String(doc.prefix || "").toUpperCase()}-${doc.number}`;
        return docName === upperName || docAltName === upperName;
      });

      if (match) {
        console.log(`[CrossRef] Found ${step.type} match for ${cleanName}`);
        let fullDetail = match;

        try {
          if (step.type === "Documento Soporte") fullDetail = await getPurchaseSupportDocumentById(match.id);
          else if (step.type === "Factura Compra") fullDetail = await getPurchaseById(match.id);
          else if (step.type === "Factura Venta") fullDetail = await getInvoiceById(match.id);
          else if (step.type === "Recibo Pago") fullDetail = await getPaymentReceiptById(match.id);
        } catch (e) {
          console.warn(`[CrossRef] Could not get full detail for ${cleanName}.`);
        }

        const supplier = fullDetail.supplier || fullDetail.customer || fullDetail.third_party;
        
        let related = "";
        if (step.type === "Recibo Pago" && fullDetail.items) {
          related = fullDetail.items
            .map((it: any) => it.due ? `${it.due.prefix}-${it.due.consecutive}` : "")
            .filter(Boolean)
            .join(", ");
        }

        return {
          status: "found" as const,
          type: step.type,
          siigoId: match.id,
          siigoName: match.name,
          date: fullDetail.date || match.date,
          supplierId: String(supplier?.identification || ""),
          supplierName: String(supplier?.name || (Array.isArray(supplier?.name) ? supplier.name.join(" ") : "")),
          totalValue: Number(fullDetail.total || 0),
          observations: fullDetail.observations || "",
          relatedDoc: related
        };
      }
    } catch (err) {
      if ((err as any)?.status === 429) throw err;
    }
    await delay(100);
  }

  return null;
}

export async function processBulkPdfExcel(excelBuffer: Buffer): Promise<BulkProcessResult> {
  console.log("[CrossRef] Starting processing...");
  const inputWorkbook = new ExcelJS.Workbook();
  await inputWorkbook.xlsx.load(excelBuffer as any);

  const sheet = inputWorkbook.worksheets[0];
  if (!sheet) throw new Error("El Excel no contiene hojas");

  const documentNames: string[] = [];
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const value = row.getCell(1).value;
    if (value) {
      const name = typeof value === "object" && value && "result" in value ? String((value as any).result) : String(value);
      if (name.trim()) documentNames.push(name.trim());
    }
  });

  console.log(`[CrossRef] Processing ${documentNames.length} documents...`);
  const items: CrossRefItem[] = documentNames.map((name) => ({ originalName: name, status: "pending" }));

  let foundCount = 0;
  let notFoundCount = 0;
  let errorCount = 0;

  for (const item of items) {
    try {
      const data = await findDocumentData(item.originalName);
      if (data) {
        Object.assign(item, data);
        foundCount++;
      } else {
        item.status = "not_found";
        notFoundCount++;
      }
      await delay(200); 
    } catch (err) {
      console.error(`[CrossRef] Error processing ${item.originalName}:`, err);
      item.status = "error";
      item.notes = (err as any)?.message || String(err);
      errorCount++;
      if ((err as any)?.status === 429) await delay(10000);
    }
  }

  const outputWorkbook = new ExcelJS.Workbook();
  const outSheet = outputWorkbook.addWorksheet("Reporte Siigo");

  outSheet.columns = [
    { header: "Documento en Excel", key: "originalName", width: 25 },
    { header: "Estado en Siigo", key: "status", width: 15 },
    { header: "Tipo", key: "type", width: 20 },
    { header: "Nombre en Siigo", key: "siigoName", width: 20 },
    { header: "Fecha", key: "date", width: 15 },
    { header: "NIT Proveedor/Cliente", key: "supplierId", width: 20 },
    { header: "Nombre Proveedor/Cliente", key: "supplierName", width: 30 },
    { header: "Valor Total", key: "totalValue", width: 15 },
    { header: "Documentos Relacionados (si es Pago)", key: "relatedDoc", width: 30 },
    { header: "Observaciones", key: "observations", width: 40 },
    { header: "Notas Proceso", key: "notes", width: 30 },
  ];

  outSheet.getRow(1).font = { bold: true };
  outSheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE0E0E0" } };

  items.forEach((item) => {
    const row = outSheet.addRow({
      originalName: item.originalName,
      status: item.status === "found" ? "ENCONTRADO" : (item.status === "not_found" ? "NO EXISTE" : "ERROR"),
      type: item.type,
      siigoName: item.siigoName,
      date: item.date,
      supplierId: item.supplierId,
      supplierName: item.supplierName,
      totalValue: item.totalValue,
      relatedDoc: item.relatedDoc,
      observations: item.observations,
      notes: item.notes,
    });
    row.getCell("totalValue").numFmt = '"$"#,##0.00';
  });

  const buffer = await outputWorkbook.xlsx.writeBuffer();
  return {
    excelBuffer: Buffer.from(buffer),
    summary: {
      total: items.length,
      found: foundCount,
      notFound: notFoundCount,
      errors: errorCount,
    },
  };
}


