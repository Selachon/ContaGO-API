import JSZip from "jszip";
import { extractInvoiceDataFromXml } from "./xmlParser.js";
import {
  createPurchase,
  createPurchaseSupportDocument,
  createCreditNote,
  listPurchases,
  listCustomers,
  getCurrentSiigoCompanyId,
  getDocumentType,
  SiigoError,
} from "./siigoService.js";
import { getSupplierProfile, type SupplierProfile } from "./siigoSuggestionsService.js";

/** Impuesto extra (ICL, IBUA, IC, Bolsas…) de una línea del XML. */
export interface XmlExtraTax {
  taxName: string; // "ICL", "IBUA", "IC Porcentual", etc.
  amount: number;  // valor monetario ya calculado
  percent: number; // tasa porcentual (0 si es específico)
}

/** Una línea del XML lista para volcar como ítem de causación. */
export interface XmlCausacionItem {
  description: string;
  base: number;
  ivaPercent: number;
  incPercent: number;
  /** Impuestos adicionales distintos de IVA/Retefte que generan líneas separadas. */
  extraTaxes?: XmlExtraTax[];
  /**
   * Obsequio/bonificación: el proveedor regala el producto (base a pagar = 0) pero
   * traslada el IVA. La base se ignora y el IVA se causa como impuesto extra en una
   * cuenta PUC parametrizable. La UI omite la línea normal (precio 0) de estos ítems.
   */
  isGift?: boolean;
  /**
   * Recargo a nivel documento (ChargeTotalAmount): se aplica DESPUÉS del total y no
   * afecta impuestos. Se causa como ítem aparte en una cuenta PUC parametrizable, sin
   * IVA. La UI omite la línea normal (igual que isGift).
   */
  isSurcharge?: boolean;
}

/** Borrador de tercero (proveedor) prellenado desde el XML para crearlo en Siigo. */
export interface SupplierDraft {
  type: "Supplier";
  person_type: "Person" | "Company";
  id_type: string; // "31" NIT, "13" cédula
  identification: string;
  name: string[];
  commercial_name?: string;
  vat_responsible: boolean;
  fiscal_responsibilities: Array<{ code: string }>;
  address: {
    address: string;
    city: { country_code: string; state_code: string; city_code: string };
  };
  phones: Array<{ number: string }>;
  contacts: Array<{ first_name: string; last_name?: string; email?: string }>;
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
  supplierDraft: SupplierDraft;
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
 * Construye un borrador de tercero (proveedor) a partir de los datos del emisor
 * del XML, listo para que el usuario lo revise/edite y se cree vía POST /v1/customers.
 */
function buildSupplierDraft(xml: any): SupplierDraft {
  const clean = (v: unknown) => {
    const s = String(v ?? "").trim();
    return s && s.toUpperCase() !== "N/A" ? s : "";
  };
  const nit = clean(xml.issuerNit).split("-")[0].replace(/\D/g, "");
  // AdditionalAccountID (issuerTaxpayerType): "1" persona jurídica, "2" persona natural.
  const isCompany = clean(xml.issuerTaxpayerType) !== "2";
  const name = clean(xml.issuerName);
  const email = clean(xml.issuerEmail);
  const phone = clean(xml.issuerPhone).replace(/\D/g, "");
  const cityCode = clean(xml.issuerCityCode).replace(/\D/g, "");
  const stateCode = clean(xml.issuerStateCode).replace(/\D/g, "") || (cityCode.length >= 2 ? cityCode.slice(0, 2) : "");
  let countryCode = clean(xml.issuerCountryCode) || "CO";
  if (countryCode.toUpperCase() === "CO") countryCode = "Co"; // Siigo guarda Colombia como "Co"

  return {
    type: "Supplier",
    person_type: isCompany ? "Company" : "Person",
    id_type: isCompany ? "31" : "13",
    identification: nit,
    name: [name],
    commercial_name: clean(xml.issuerCommercialName) || undefined,
    vat_responsible: false,
    fiscal_responsibilities: [{ code: "R-99-PN" }],
    address: {
      address: clean(xml.issuerAddress) || "N/A",
      city: { country_code: countryCode, state_code: stateCode, city_code: cityCode },
    },
    phones: phone ? [{ number: phone }] : [],
    contacts: [{ first_name: name || "Contacto", ...(email ? { email } : {}) }],
  };
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
    throw new NonInvoiceXmlError(docTypeStr);
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

  // Impuestos que NO son IVA/Retefte y se envían como ítems separados a Siigo
  const IVA_LIKE = new Set(["IVA", "Retefuente", "ReteICA", "ReteIVA"]);
  // Nombre del "impuesto extra" bajo el cual se causa el IVA de obsequios (cuenta PUC parametrizable).
  const GIFT_TAX_NAME = "IVA Obsequio";
  const items: XmlCausacionItem[] = lineItems.map((li) => {
    const lineExt = Number(li.totalUnitPrice) || 0;
    const ivaAmt = Number(li.ivaAmount) || 0;
    const incAmt = Number(li.incAmount) || 0;
    // Obsequio/bonificación: base a pagar = 0 (LineExtensionAmount) pero con IVA/INC > 0.
    // Se ignora la base; el impuesto se enruta como ítem extra hacia su cuenta PUC.
    const isGift = lineExt === 0 && (ivaAmt > 0 || incAmt > 0);

    const extraTaxes: XmlExtraTax[] = (li.taxes || [])
      .filter((t: any) => !IVA_LIKE.has(t.taxName) && Number(t.amount) > 0)
      .map((t: any) => ({ taxName: t.taxName, amount: Number(t.amount) || 0, percent: Number(t.percent) || 0 }));

    if (isGift) {
      if (ivaAmt > 0) extraTaxes.push({ taxName: GIFT_TAX_NAME, amount: ivaAmt, percent: Number(li.ivaPercent) || 0 });
      if (incAmt > 0) extraTaxes.push({ taxName: GIFT_TAX_NAME, amount: incAmt, percent: Number(li.incPercent) || 0 });
    }

    // Base = base gravable del IVA (TaxableAmount), que ya viene neta de descuentos de
    // línea. Si el XML no la trae, se cae a LineExtensionAmount o cantidad×precio.
    const taxBase = Number(li.taxableBase) || 0;
    return {
      description: li.description || "",
      base: isGift ? 0 : (taxBase > 0 ? taxBase : (lineExt || (Number(li.quantity) || 0) * (Number(li.unitPrice) || 0))),
      ivaPercent: isGift ? 0 : (Number(li.ivaPercent) || 0),
      incPercent: isGift ? 0 : (Number(li.incPercent) || 0),
      extraTaxes: extraTaxes.length ? extraTaxes : undefined,
      isGift: isGift || undefined,
    };
  });

  // Recargo a nivel documento (ChargeTotalAmount): se aplica después del total, sin IVA.
  // Se causa como ítem aparte hacia una cuenta PUC parametrizable "Recargo".
  const surcharge = Number(xmlData.surcharge) || 0;
  if (surcharge > 0) {
    items.push({
      description: "RECARGO",
      base: 0,
      ivaPercent: 0,
      incPercent: 0,
      extraTaxes: [{ taxName: "Recargo", amount: Math.round(surcharge * 100) / 100, percent: 0 }],
      isSurcharge: true,
    });
  }

  console.log(
    `[SiigoAccounting] XML ${xmlData.docNumber || "?"} — NIT ${xmlData.issuerNit}, ${items.length} ítem(s)` +
      (surcharge > 0 ? ` (incluye recargo ${surcharge})` : "")
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
    supplierDraft: buildSupplierDraft(xmlData),
  };
}

// ─── Procesamiento por lote (varios XML o un ZIP) ────────────────────
export interface BatchItem {
  fileName: string;
  ok: boolean;
  message?: string;
  /** XML auxiliar conocido (Application Response, nómina, etc.): se ignora sin tratarlo como error. */
  ignored?: boolean;
  xml?: XmlCausacionData;
  profile?: SupplierProfile | null;
  alreadyCausada?: boolean;
  pdfBase64?: string;
  pdfName?: string;
}

/** Error para XMLs que no son facturas (acuses, nómina, etc.). No es un fallo real. */
export class NonInvoiceXmlError extends Error {
  docType: string;
  constructor(docType: string) {
    super(`Documento auxiliar de tipo "${docType}" (no es una factura).`);
    this.name = "NonInvoiceXmlError";
    this.docType = docType;
  }
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
      // Los XMLs auxiliares (Application Response, nómina, etc.) son ruido esperado
      // dentro de los ZIP de la DIAN: se marcan como "ignored", no como error.
      if (e instanceof NonInvoiceXmlError) {
        results.push({ fileName: xf.name, ok: false, ignored: true, message: e.message });
      } else {
        results.push({
          fileName: xf.name,
          ok: false,
          message: e instanceof Error ? e.message : "Error procesando el XML",
        });
      }
    }
  }

  // La verificación de duplicados contra Siigo fue removida: la herramienta
  // permite subir todas las facturas y deja al usuario decidir cuáles causar.

  return results;
}

// ─── Índice de terceros para autocompletar (nombre / NIT) ─────────────
export interface CustomerIndexEntry {
  id: string;
  identification: string;
  name: string;
  branch_office: number;
}

const customerIndexCache = new Map<string, { at: number; data: CustomerIndexEntry[] }>();
const CUSTOMER_INDEX_TTL_MS = 10 * 60 * 1000;

/** Borra la caché del índice (p. ej. tras crear un tercero nuevo). */
export function invalidateCustomersIndex(): void {
  customerIndexCache.clear();
}

/**
 * Trae todos los terceros de la empresa (paginado) en forma reducida para
 * autocompletar en el cliente. Cacheado por empresa con TTL para no golpear la API.
 */
export async function getCustomersIndex(): Promise<CustomerIndexEntry[]> {
  const key = getCurrentSiigoCompanyId() || "env";
  const cached = customerIndexCache.get(key);
  if (cached && Date.now() - cached.at < CUSTOMER_INDEX_TTL_MS) return cached.data;

  const out: CustomerIndexEntry[] = [];
  for (let page = 1; page <= 200; page++) {
    const resp = (await listCustomers({ page, page_size: 100 })) as any;
    const results: any[] = Array.isArray(resp?.results) ? resp.results : [];
    if (!results.length) break;
    for (const c of results) {
      out.push({
        id: String(c?.id ?? ""),
        identification: String(c?.identification ?? ""),
        name: Array.isArray(c?.name) ? c.name.filter(Boolean).join(" ") : String(c?.name ?? c?.commercial_name ?? ""),
        branch_office: Number(c?.branch_office) || 0,
      });
    }
    const total = Number(resp?.pagination?.total_results) || 0;
    if (results.length < 100 || (total && out.length >= total)) break;
  }
  customerIndexCache.set(key, { at: Date.now(), data: out });
  return out;
}

function isNumberRequiredError(details: unknown): boolean {
  const errs = ((details as any)?.errors ?? (details as any)?.Errors);
  if (!Array.isArray(errs)) return false;
  return errs.some((e: any) => {
    const code = String(e?.code ?? e?.Code ?? "").toLowerCase();
    const params = Array.isArray(e?.params ?? e?.Params) ? (e?.params ?? e?.Params) : [];
    return code === "parameter_required" && params.includes("number");
  });
}

/**
 * Si el error de Siigo es `invalid_total_payments`, devuelve el total de compra
 * que Siigo calculó (lo informa en el mensaje: "The total purchase calculated is 97999.94").
 * Ese es exactamente el valor que `payments` debe sumar. Devuelve null si no aplica.
 */
function extractCalculatedPurchaseTotal(details: unknown): number | null {
  const errs = ((details as any)?.errors ?? (details as any)?.Errors);
  if (!Array.isArray(errs)) return null;
  for (const e of errs as Array<Record<string, unknown>>) {
    const code = String(e?.code ?? e?.Code ?? "").toLowerCase();
    if (code !== "invalid_total_payments") continue;
    const message = String(e?.message ?? e?.Message ?? "");
    // Siigo responde en formato en-US: punto decimal, coma como separador de miles.
    const m = message.match(/calculated is\s*([\d,]+\.?\d*)/i);
    if (m) {
      const num = Number(m[1].replace(/,/g, ""));
      if (Number.isFinite(num)) return num;
    }
  }
  return null;
}

const round2 = (n: number): number => Math.round((Number(n) || 0) * 100) / 100;

/**
 * Devuelve una copia del payload con `payments` ajustado para que sume `total`.
 * Con un solo pago, fija su valor; con varios, ajusta el último por la diferencia.
 */
function applyCorrectedPaymentTotal(payload: unknown, total: number): unknown {
  const p = payload as any;
  const payments = Array.isArray(p?.payments) ? p.payments.map((x: any) => ({ ...x })) : [];
  if (payments.length === 0) return payload;
  if (payments.length === 1) {
    payments[0].value = round2(total);
  } else {
    const sum = round2(payments.reduce((a: number, x: any) => a + (Number(x.value) || 0), 0));
    const last = payments[payments.length - 1];
    last.value = round2((Number(last.value) || 0) + (round2(total) - sum));
  }
  return { ...p, payments };
}

/**
 * Ejecuta `fn(payload)` y, si Siigo rechaza por `invalid_total_payments`,
 * reenvía una sola vez con el total exacto que Siigo informó.
 */
async function createWithTotalRetry(
  fn: (body: unknown) => Promise<unknown>,
  payload: unknown
): Promise<unknown> {
  try {
    return await fn(payload);
  } catch (err) {
    if (err instanceof SiigoError) {
      const total = extractCalculatedPurchaseTotal(err.details);
      if (total != null) {
        console.log(`[SiigoAccounting] invalid_total_payments: reintentando con total exacto de Siigo=${total}`);
        return await fn(applyCorrectedPaymentTotal(payload, total));
      }
    }
    throw err;
  }
}

export async function submitToSiigo(type: "FC" | "DS" | "NC", payload: unknown) {
  if (type === "FC") {
    const create = (body: unknown) => createWithTotalRetry(createPurchase, body);
    try {
      return await create(payload);
    } catch (err) {
      if (err instanceof SiigoError && isNumberRequiredError(err.details)) {
        const docId = (payload as any)?.document?.id;
        if (docId) {
          try {
            const docType = await getDocumentType(Number(docId)) as any;
            const nextNumber = docType?.consecutive ?? docType?.next_consecutive ?? docType?.consecutive_number;
            if (nextNumber != null) {
              console.log(`[SiigoAccounting] Reintentando con number=${nextNumber} (tipo documento manual)`);
              return await create({ ...(payload as any), number: Number(nextNumber) });
            }
          } catch { /* si falla el fallback, propagar error original */ }
        }
      }
      throw err;
    }
  }
  if (type === "DS") return createWithTotalRetry(createPurchaseSupportDocument, payload);
  if (type === "NC") return createCreditNote(payload);
  throw new Error("Tipo de documento no soportado");
}

/**
 * Si los detalles de un error de Siigo corresponden a "el proveedor no existe"
 * (code invalid_reference sobre supplier.identification), devuelve el NIT
 * mencionado (o "" si no se pudo extraer). Si no es ese error, devuelve null.
 */
export function supplierNotFoundNit(details: unknown): string | null {
  const rec = details as Record<string, unknown> | null;
  const errs = (rec?.errors ?? rec?.Errors) as unknown;
  if (!Array.isArray(errs)) return null;
  for (const e of errs as Array<Record<string, unknown>>) {
    const code = String(e?.code ?? e?.Code ?? "").toLowerCase();
    const message = String(e?.message ?? e?.Message ?? "");
    const params = (Array.isArray(e?.params ?? e?.Params) ? (e?.params ?? e?.Params) : []) as unknown[];
    const aboutSupplier =
      params.some((p) => String(p).toLowerCase().includes("supplier")) ||
      /supplier|proveedor/i.test(message);
    if (code === "invalid_reference" && aboutSupplier) {
      const m = message.match(/(\d{4,})/);
      return m ? m[1] : "";
    }
  }
  return null;
}

/**
 * Si el error de Siigo es `parameter_required` sobre `items[N].tax.id` (la cuenta de
 * esa línea está configurada en Siigo para manejar impuesto y se envió sin él),
 * devuelve el índice N de la línea. Si no aplica, devuelve null.
 */
export function taxRequiredItemIndex(details: unknown): number | null {
  const rec = details as Record<string, unknown> | null;
  const errs = (rec?.errors ?? rec?.Errors) as unknown;
  if (!Array.isArray(errs)) return null;
  for (const e of errs as Array<Record<string, unknown>>) {
    const code = String(e?.code ?? e?.Code ?? "").toLowerCase();
    if (code !== "parameter_required") continue;
    const params = (Array.isArray(e?.params ?? e?.Params) ? (e?.params ?? e?.Params) : []) as unknown[];
    for (const p of params) {
      const m = String(p).match(/items\[(\d+)\]\.tax(?:es)?\.id/i);
      if (m) return Number(m[1]);
    }
  }
  return null;
}
