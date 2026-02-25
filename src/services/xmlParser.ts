import { XMLParser } from "fast-xml-parser";
import type { InvoiceData } from "../types/dianExcel.js";

interface DocInfo {
  id: string;
  nit: string;
  docnum: string;
}

// Namespaces comunes en facturas electrónicas DIAN (UBL 2.1)
const NAMESPACES = {
  cbc: "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
  cac: "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
  fe: "http://www.dian.gov.co/contratos/facturaelectronica/v1",
};

/**
 * Extrae datos estructurados de un XML de factura electrónica DIAN (UBL 2.1)
 */
export async function extractInvoiceDataFromXml(
  xmlBuffer: Buffer,
  docInfo: DocInfo
): Promise<Partial<InvoiceData>> {
  try {
    const xmlString = xmlBuffer.toString("utf-8");

    const parser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: "@_",
      removeNSPrefix: true, // Elimina prefijos de namespace para facilitar acceso
      parseAttributeValue: true,
      trimValues: true,
    });

    const parsed = parser.parse(xmlString);

    // El documento puede ser Invoice (factura) o CreditNote (nota crédito)
    const invoice = parsed.Invoice || parsed.CreditNote;
    if (!invoice) {
      throw new Error("No se encontró Invoice o CreditNote en el XML");
    }

    const isNotaCredito = !!parsed.CreditNote;

    // Extraer datos principales
    const issueDate = extractIssueDate(invoice);
    const supplierParty = invoice.AccountingSupplierParty?.Party;
    const entityName = extractEntityName(supplierParty);
    const entityType = detectEntityType(supplierParty);
    const { subtotal, iva } = extractTotals(invoice);
    const concepts = extractConcepts(invoice);
    const cufe = extractCUFE(invoice);

    return {
      entityType,
      issueDate,
      entityName,
      subtotal,
      iva,
      concepts,
      documentType: isNotaCredito ? "Nota Crédito" : "Factura Electrónica",
      cufe,
      trackId: docInfo.id,
      nit: docInfo.nit,
      docNumber: docInfo.docnum,
    };
  } catch (err) {
    console.error(`Error parseando XML ${docInfo.docnum}:`, err);
    return {
      entityType: "N/A",
      issueDate: "N/A",
      entityName: "N/A",
      subtotal: 0,
      iva: 0,
      concepts: "ERROR: No se pudo leer el XML",
      documentType: "N/A",
      cufe: "N/A",
      trackId: docInfo.id,
      nit: docInfo.nit,
      docNumber: docInfo.docnum,
      error: (err as Error).message,
    };
  }
}

/**
 * Extrae fecha de emisión en formato DD/MM/YYYY
 */
function extractIssueDate(invoice: any): string {
  try {
    const issueDate = invoice.IssueDate;
    if (!issueDate) return "N/A";

    // Formato ISO: YYYY-MM-DD
    if (typeof issueDate === "string" && /^\d{4}-\d{2}-\d{2}/.test(issueDate)) {
      const [year, month, day] = issueDate.split("-");
      return `${day}/${month}/${year}`;
    }

    return String(issueDate);
  } catch {
    return "N/A";
  }
}

/**
 * Extrae razón social del emisor (AccountingSupplierParty)
 */
function extractEntityName(party: any): string {
  if (!party) return "N/A";

  try {
    // Intentar PartyLegalEntity > RegistrationName (más confiable)
    const legalEntity = party.PartyLegalEntity;
    if (legalEntity) {
      const regName = Array.isArray(legalEntity)
        ? legalEntity[0]?.RegistrationName
        : legalEntity.RegistrationName;
      if (regName) {
        return cleanText(getText(regName));
      }
    }

    // Fallback: PartyName > Name
    const partyName = party.PartyName;
    if (partyName) {
      const name = Array.isArray(partyName)
        ? partyName[0]?.Name
        : partyName.Name;
      if (name) {
        return cleanText(getText(name));
      }
    }

    return "N/A";
  } catch {
    return "N/A";
  }
}

/**
 * Detecta si es empresa o persona natural basado en datos del emisor
 */
function detectEntityType(party: any): "EMPRESA" | "PN" | "N/A" {
  if (!party) return "N/A";

  try {
    // En UBL DIAN, el tipo de persona está en PartyTaxScheme o en extensiones
    const partyTaxScheme = party.PartyTaxScheme;

    if (partyTaxScheme) {
      const taxScheme = Array.isArray(partyTaxScheme)
        ? partyTaxScheme[0]
        : partyTaxScheme;

      // TaxLevelCode puede indicar régimen (útil pero no determinante)
      // CompanyID suele tener atributos con tipo de documento

      const companyId = taxScheme?.CompanyID;
      if (companyId) {
        // Atributo schemeName puede ser "NIT" (empresa) o "CC" (persona natural)
        const schemeName = companyId["@_schemeName"] || "";
        if (/NIT/i.test(schemeName)) return "EMPRESA";
        if (/CC|CE|TI|PA/i.test(schemeName)) return "PN";
      }
    }

    // Fallback: analizar razón social por sufijos de empresa
    const name = extractEntityName(party);
    if (/\b(S\.?A\.?S\.?|S\.?A\.?|LTDA\.?|E\.?U\.?|SAS|LIMITADA)\b/i.test(name)) {
      return "EMPRESA";
    }

    return "EMPRESA"; // Default
  } catch {
    return "N/A";
  }
}

/**
 * Extrae subtotal e IVA de los totales monetarios
 */
function extractTotals(invoice: any): { subtotal: number; iva: number } {
  let subtotal = 0;
  let iva = 0;

  try {
    // LegalMonetaryTotal > LineExtensionAmount (subtotal antes de impuestos)
    const legalMonetaryTotal = invoice.LegalMonetaryTotal;
    if (legalMonetaryTotal) {
      const lineExtension = legalMonetaryTotal.LineExtensionAmount;
      subtotal = parseAmount(lineExtension);
    }

    // TaxTotal > TaxAmount (IVA)
    const taxTotal = invoice.TaxTotal;
    if (taxTotal) {
      const taxTotalArr = Array.isArray(taxTotal) ? taxTotal : [taxTotal];

      for (const tax of taxTotalArr) {
        // TaxSubtotal puede tener múltiples impuestos
        const taxSubtotal = tax.TaxSubtotal;
        if (taxSubtotal) {
          const subtotals = Array.isArray(taxSubtotal) ? taxSubtotal : [taxSubtotal];

          for (const sub of subtotals) {
            // TaxCategory > TaxScheme > ID indica tipo de impuesto
            const taxSchemeId = sub.TaxCategory?.TaxScheme?.ID;
            const taxId = getText(taxSchemeId);

            // "01" = IVA en nomenclatura DIAN
            if (taxId === "01" || /IVA/i.test(taxId)) {
              iva += parseAmount(sub.TaxAmount);
            }
          }
        }

        // Si no hay TaxSubtotal, usar TaxAmount directo
        if (!taxSubtotal && tax.TaxAmount) {
          iva += parseAmount(tax.TaxAmount);
        }
      }
    }
  } catch (err) {
    console.error("Error extrayendo totales:", err);
  }

  return { subtotal, iva };
}

/**
 * Extrae descripciones de los productos/servicios facturados
 */
function extractConcepts(invoice: any): string {
  try {
    // InvoiceLine o CreditNoteLine contienen los ítems
    const lines = invoice.InvoiceLine || invoice.CreditNoteLine;
    if (!lines) return "N/A";

    const linesArr = Array.isArray(lines) ? lines : [lines];
    const descriptions: string[] = [];

    for (const line of linesArr) {
      // Item > Description o Item > Name
      const item = line.Item;
      if (item) {
        let desc = getText(item.Description) || getText(item.Name);
        if (desc && desc !== "N/A") {
          desc = cleanDescription(desc);
          if (!descriptions.some((d) => d.toLowerCase() === desc.toLowerCase())) {
            descriptions.push(desc);
          }
        }
      }
    }

    return formatConcepts(descriptions);
  } catch {
    return "N/A";
  }
}

/**
 * Extrae CUFE/CUDE del documento
 */
function extractCUFE(invoice: any): string {
  try {
    // UUID está en UBLExtensions > ExtensionContent > DianExtensions > InvoiceControl > UUID
    // O directamente en el campo UUID del documento
    const uuid = invoice.UUID;
    if (uuid) {
      const cufe = getText(uuid);
      if (cufe && /^[a-f0-9]{64,96}$/i.test(cufe)) {
        return cufe.toLowerCase();
      }
    }

    // Buscar en extensiones DIAN
    const extensions = invoice.UBLExtensions?.UBLExtension;
    if (extensions) {
      const extArr = Array.isArray(extensions) ? extensions : [extensions];
      for (const ext of extArr) {
        const dianExt = ext.ExtensionContent?.DianExtensions;
        if (dianExt) {
          const invoiceControl = dianExt.InvoiceControl;
          if (invoiceControl?.UUID) {
            const cufe = getText(invoiceControl.UUID);
            if (cufe && /^[a-f0-9]{64,96}$/i.test(cufe)) {
              return cufe.toLowerCase();
            }
          }
        }
      }
    }

    return "N/A";
  } catch {
    return "N/A";
  }
}

// ============================================
// Utilidades
// ============================================

/**
 * Obtiene texto de un nodo XML (puede ser string o objeto con #text)
 */
function getText(node: any): string {
  if (!node) return "";
  if (typeof node === "string") return node;
  if (typeof node === "number") return String(node);
  if (node["#text"]) return String(node["#text"]);
  return "";
}

/**
 * Parsea montos desde XML (pueden tener atributos de moneda)
 */
function parseAmount(node: any): number {
  const text = getText(node);
  if (!text) return 0;

  // Limpiar y parsear
  const cleaned = text.replace(/[^\d.,\-]/g, "");

  // Determinar formato
  const lastComma = cleaned.lastIndexOf(",");
  const lastDot = cleaned.lastIndexOf(".");

  let normalized = cleaned;
  if (lastComma > lastDot) {
    // Formato europeo: 1.234,56
    normalized = cleaned.replace(/\./g, "").replace(",", ".");
  } else if (lastDot > lastComma) {
    // Formato americano: 1,234.56
    normalized = cleaned.replace(/,/g, "");
  }

  const result = parseFloat(normalized);
  return isNaN(result) ? 0 : result;
}

/**
 * Limpia texto de espacios extra y caracteres especiales
 */
function cleanText(text: string): string {
  return text.replace(/\s+/g, " ").trim().substring(0, 100);
}

/**
 * Limpia y formatea descripción de producto
 */
function cleanDescription(text: string): string {
  let cleaned = text.replace(/\s+/g, " ").trim();

  // Capitalizar primera letra
  if (cleaned.length > 0) {
    cleaned = cleaned.charAt(0).toUpperCase() + cleaned.slice(1).toLowerCase();
  }

  return cleaned.substring(0, 80);
}

/**
 * Formatea lista de conceptos para el Excel
 */
function formatConcepts(concepts: string[]): string {
  if (concepts.length === 0) return "N/A";

  // Eliminar duplicados
  const unique = [...new Set(concepts.map((c) => c.toLowerCase()))].map(
    (c) => concepts.find((orig) => orig.toLowerCase() === c)!
  );

  if (unique.length === 1) return unique[0];
  if (unique.length === 2) return unique.join(", ");

  return `${unique.slice(0, 2).join(", ")}... (+${unique.length - 2} items mas)`;
}
