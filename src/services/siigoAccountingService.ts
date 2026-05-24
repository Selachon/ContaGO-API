import JSZip from "jszip";
import { extractInvoiceDataFromXml } from "./xmlParser.js";
import {
  createPurchase,
  createPurchaseSupportDocument,
  createCreditNote,
  listPurchases,
} from "./siigoService.js";
import { getSupplierProfile, type SupplierProfile } from "./siigoSuggestionsService.js";

/** Una línea del XML lista para volcar como ítem de causación. */
export interface XmlCausacionItem {
  description: string; // Concepto (hoja "Detallado")
  base: number; // Base del impuesto / valor unitario
  ivaPercent: number; // 0 si la línea no tiene IVA
  incPercent: number; // 0 si la línea no tiene Impoconsumo
}

/** Datos del XML mapeados a los campos del formulario de causación. */
export interface XmlCausacionData {
  documentType: string;
  isCreditNote: boolean;
  date: string; // YYYY-MM-DD ("" si no se pudo leer)
  supplierNit: string;
  supplierName: string;
  providerInvoicePrefix: string;
  providerInvoiceNumber: string;
  docNumberRaw: string;
  cufe: string;
  items: XmlCausacionItem[];
  totals: { subtotal: number; iva: number; total: number };
}

/**
 * Separa el número de factura del proveedor en prefijo (letras) + número (dígitos).
 * Si no hay letras, el prefijo por defecto es "FC". Siigo exige number numérico.
 */
function splitDocNumber(raw: string): { prefix: string; number: string } {
  const clean = String(raw || "").trim().replace(/[\s\-_.]/g, "");
  const m = clean.match(/^([A-Za-z]*)(\d+)$/);
  if (m) {
    return { prefix: m[1] ? m[1].toUpperCase() : "FC", number: m[2] };
  }
  const letters = clean.replace(/[^A-Za-z]/g, "");
  const digits = clean.replace(/\D/g, "");
  return { prefix: letters ? letters.toUpperCase() : "FC", number: digits };
}

/**
 * Lee un XML de factura/documento DIAN y lo traduce a los campos de la
 * herramienta de causación. Reutiliza el mismo parser que "Exportar Excel DIAN".
 */
export async function processXmlForAccounting(xmlBuffer: Buffer): Promise<XmlCausacionData> {
  console.log(`[SiigoAccounting] Procesando XML (${xmlBuffer.length} bytes)`);

  const xmlData = (await extractInvoiceDataFromXml(xmlBuffer, { id: "xml", docnum: "" })) as any;

  const lineItems: any[] = Array.isArray(xmlData.lineItems) ? xmlData.lineItems : [];
  const docTypeStr = String(xmlData.documentType || "").trim();
  const NON_INVOICE_LABELS = new Set([
    "Application Response",
    "Documento Adjunto DIAN",
    "Nómina Individual Electrónica",
    "Nómina Individual de Ajuste",
    "Aviso de Despacho",
    "Aviso de Recepción",
  ]);
  if (NON_INVOICE_LABELS.has(docTypeStr)) {
    throw new Error(`No es una factura procesable: el XML es de tipo "${docTypeStr}" (no es una factura ni una nota crédito de compra).`);
  }
  if (!xmlData.issuerNit && lineItems.length === 0) {
    throw new Error(
      docTypeStr
        ? `XML de tipo "${docTypeStr}" sin contenido procesable (sin NIT del emisor ni líneas).`
        : "El XML no parece ser una factura o documento procesable (no se identificó el emisor ni líneas)."
    );
  }

  const { prefix, number } = splitDocNumber(xmlData.docNumber);
  const isCreditNote = /cr[eé]dito/i.test(String(xmlData.documentType || ""));
  const validDate =
    xmlData.issueDateISO && xmlData.issueDateISO !== "9999-12-31" ? xmlData.issueDateISO : "";

  const items: XmlCausacionItem[] = lineItems.map((li) => ({
    description: li.description || "",
    base:
      Number(li.totalUnitPrice) ||
      (Number(li.quantity) || 0) * (Number(li.unitPrice) || 0),
    ivaPercent: Number(li.ivaPercent) || 0,
    incPercent: Number(li.incPercent) || 0,
  }));

  console.log(
    `[SiigoAccounting] XML ${xmlData.docNumber || "?"} — NIT ${xmlData.issuerNit}, ${items.length} ítem(s)`
  );

  return {
    documentType: xmlData.documentType || "Factura Electrónica",
    isCreditNote,
    date: validDate,
    supplierNit: xmlData.issuerNit || "",
    supplierName: xmlData.issuerName || "",
    providerInvoicePrefix: prefix,
    providerInvoiceNumber: number,
    docNumberRaw: xmlData.docNumber || "",
    cufe: xmlData.cufe || "",
    items,
    totals: {
      subtotal: Number(xmlData.subtotal) || 0,
      iva: Number(xmlData.iva) || 0,
      total: Number(xmlData.total) || 0,
    },
  };
}

// ─── Procesamiento por lote (varios XML o un ZIP) ────────────────────
export interface BatchItem {
  fileName: string;
  ok: boolean;
  message?: string;
  xml?: XmlCausacionData;
  profile?: SupplierProfile | null;
  alreadyCausada?: boolean;
  pdfBase64?: string;
  pdfName?: string;
}

const onlyDigits = (s: unknown) => String(s ?? "").replace(/\D/g, "");
const nitKey = (s: unknown) => String(s ?? "").split("-")[0].replace(/\D/g, "");

/**
 * Recolecta las facturas de compra ya registradas en Siigo en los últimos 12
 * meses. Llave: NIT del proveedor + número de la factura del proveedor.
 */
async function fetchCausadasKeys(): Promise<Set<string>> {
  const start = new Date();
  start.setMonth(start.getMonth() - 12);
  const createdStart = start.toISOString().slice(0, 10);
  const keys = new Set<string>();
  for (let page = 1; page <= 100; page++) {
    const resp = (await listPurchases({ created_start: createdStart, page, page_size: 100 })) as any;
    const results: any[] = Array.isArray(resp?.results) ? resp.results : [];
    if (results.length === 0) break;
    for (const p of results) {
      const nit = nitKey(p?.supplier?.identification);
      const num = onlyDigits(p?.provider_invoice?.number ?? p?.provider_invoice_number);
      if (nit && num) keys.add(`${nit}|${num}`);
    }
  }
  return keys;
}

export async function processXmlBatch(
  files: { name: string; buffer: Buffer }[]
): Promise<BatchItem[]> {
  // Expandir: descomprimir los ZIP y recolectar todos los XML + PDFs
  const xmlFiles: { name: string; buffer: Buffer }[] = [];
  const pdfMap = new Map<string, { name: string; buffer: Buffer }>();
  const stripExt = (n: string) => n.replace(/\.[^.]+$/, "");
  const baseName = (p: string) => (p.split("/").pop() || p);
  for (const f of files) {
    if (/\.zip$/i.test(f.name)) {
      const zip = await JSZip.loadAsync(f.buffer);
      for (const entry of Object.values(zip.files)) {
        if (entry.dir) continue;
        const n = baseName(entry.name);
        if (/\.xml$/i.test(n)) {
          const buf = await entry.async("nodebuffer");
          xmlFiles.push({ name: n, buffer: buf });
        } else if (/\.pdf$/i.test(n)) {
          const buf = await entry.async("nodebuffer");
          pdfMap.set(stripExt(n).toLowerCase(), { name: n, buffer: buf });
        }
      }
    } else if (/\.xml$/i.test(f.name)) {
      xmlFiles.push(f);
    } else if (/\.pdf$/i.test(f.name)) {
      pdfMap.set(stripExt(baseName(f.name)).toLowerCase(), f);
    }
  }

  console.log(`[SiigoAccounting] Lote: ${files.length} archivo(s) → ${xmlFiles.length} XML, ${pdfMap.size} PDF`);

  // Empareja un PDF al XML por: 1) basename exacto, 2) PDF cuyo nombre contenga
  // el número de factura del proveedor (con o sin prefijo).
  const findPdf = (xmlName: string, xml?: XmlCausacionData) => {
    const baseXml = stripExt(xmlName).toLowerCase();
    const direct = pdfMap.get(baseXml);
    if (direct) return direct;
    if (!xml) return null;
    const num = onlyDigits(xml.providerInvoiceNumber);
    const pref = (xml.providerInvoicePrefix || "").toUpperCase();
    if (!num) return null;
    for (const [key, entry] of pdfMap) {
      const u = key.toUpperCase();
      if (pref && u.includes(`${pref}${num}`)) return entry;
      if (u.includes(num)) return entry;
    }
    return null;
  };

  const results: BatchItem[] = [];
  for (const xf of xmlFiles) {
    try {
      const xml = await processXmlForAccounting(xf.buffer);
      let profile: SupplierProfile | null = null;
      try {
        profile = await getSupplierProfile(xml.supplierNit);
      } catch {
        /* sin perfil disponible */
      }
      const pdf = findPdf(xf.name, xml);
      results.push({
        fileName: xf.name,
        ok: true,
        xml,
        profile,
        ...(pdf ? { pdfBase64: pdf.buffer.toString("base64"), pdfName: pdf.name } : {}),
      });
    } catch (e) {
      results.push({
        fileName: xf.name,
        ok: false,
        message: e instanceof Error ? e.message : "Error procesando el XML",
      });
    }
  }

  // Marcar las que ya están causadas en Siigo (NIT + número de factura)
  try {
    const causadas = await fetchCausadasKeys();
    console.log(`[SiigoAccounting] Duplicados: ${causadas.size} compras en los últimos 12 meses`);
    for (const item of results) {
      if (!item.ok || !item.xml) continue;
      const nit = nitKey(item.xml.supplierNit);
      const num = onlyDigits(item.xml.providerInvoiceNumber);
      if (nit && num && causadas.has(`${nit}|${num}`)) item.alreadyCausada = true;
    }
  } catch (e) {
    console.warn("[SiigoAccounting] No se pudo consultar duplicados:", e instanceof Error ? e.message : e);
  }

  return results;
}

export async function submitToSiigo(type: "FC" | "DS" | "NC", payload: unknown) {
  if (type === "FC") return createPurchase(payload);
  if (type === "DS") return createPurchaseSupportDocument(payload);
  if (type === "NC") return createCreditNote(payload);
  throw new Error("Tipo de documento no soportado");
}
