import { XMLParser } from "fast-xml-parser";
import type { InvoiceData, InvoiceLineItem } from "../types/dianExcel.js";

interface DocInfo {
  id: string;
  docnum: string;
  docType?: string; // Tipo de documento de la tabla DIAN (ej: "Documento soporte", "Factura electrónica")
}

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
      removeNSPrefix: true,
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
    
    // Determinar el tipo de documento:
    // 1. Si viene de la tabla DIAN (docType), usarlo directamente
    // 2. Si no, deducirlo del XML (Nota Crédito vs Factura Electrónica)
    const documentType = docInfo.docType || (isNotaCredito ? "Nota Crédito" : "Factura Electrónica");

    // Extraer datos del emisor (AccountingSupplierParty)
    const supplierParty = invoice.AccountingSupplierParty?.Party;
    const issuerNit = extractPartyNit(supplierParty);
    const issuerName = extractPartyName(supplierParty);

    // Extraer datos del receptor (AccountingCustomerParty)
    const customerParty = invoice.AccountingCustomerParty?.Party;
    const receiverNit = extractPartyNit(customerParty);
    const receiverName = extractPartyName(customerParty);

    // Extraer otros datos
    const issueDate = extractIssueDate(invoice);
    const { subtotal, iva, total } = extractTotals(invoice);
    const concepts = extractConcepts(invoice);
    const lineItems = extractLineItems(invoice);
    const cufe = extractCUFE(invoice);
    const paymentMethod = extractPaymentMethod(invoice);

    return {
      issuerNit,
      issuerName,
      receiverNit,
      receiverName,
      issueDate,
      subtotal,
      iva,
      total,
      concepts,
      lineItems,
      documentType,
      cufe,
      paymentMethod,
      trackId: docInfo.id,
      docNumber: docInfo.docnum,
    };
  } catch (err) {
    console.error(`Error parseando XML ${docInfo.docnum}:`, err);
    return {
      issuerNit: "N/A",
      issuerName: "N/A",
      receiverNit: "N/A",
      receiverName: "N/A",
      issueDate: "N/A",
      subtotal: 0,
      iva: 0,
      total: 0,
      concepts: "ERROR: No se pudo leer el XML",
      lineItems: [],
      documentType: "N/A",
      cufe: "N/A",
      paymentMethod: "N/A",
      trackId: docInfo.id,
      docNumber: docInfo.docnum,
      error: (err as Error).message,
    };
  }
}

/**
 * Extrae NIT de un Party (PartyTaxScheme > CompanyID)
 */
function extractPartyNit(party: any): string {
  if (!party) return "N/A";

  try {
    // Ruta principal: PartyTaxScheme > CompanyID
    const partyTaxScheme = party.PartyTaxScheme;
    if (partyTaxScheme) {
      const taxScheme = Array.isArray(partyTaxScheme)
        ? partyTaxScheme[0]
        : partyTaxScheme;
      const companyId = taxScheme?.CompanyID;
      if (companyId) {
        return String(getText(companyId));
      }
    }

    // Fallback: PartyLegalEntity > CompanyID
    const legalEntity = party.PartyLegalEntity;
    if (legalEntity) {
      const entity = Array.isArray(legalEntity) ? legalEntity[0] : legalEntity;
      const companyId = entity?.CompanyID;
      if (companyId) {
        return String(getText(companyId));
      }
    }

    return "N/A";
  } catch {
    return "N/A";
  }
}

/**
 * Extrae razón social de un Party (PartyLegalEntity > RegistrationName)
 * Nota: PartyName > Name contiene el nombre comercial, NO la razón social
 */
function extractPartyName(party: any): string {
  if (!party) return "N/A";

  try {
    // Ruta principal: PartyLegalEntity > RegistrationName (Razón Social)
    const legalEntity = party.PartyLegalEntity;
    if (legalEntity) {
      const entity = Array.isArray(legalEntity) ? legalEntity[0] : legalEntity;
      const regName = entity?.RegistrationName;
      if (regName) {
        return cleanText(getText(regName));
      }
    }

    // Fallback: PartyTaxScheme > RegistrationName
    const partyTaxScheme = party.PartyTaxScheme;
    if (partyTaxScheme) {
      const taxScheme = Array.isArray(partyTaxScheme)
        ? partyTaxScheme[0]
        : partyTaxScheme;
      const regName = taxScheme?.RegistrationName;
      if (regName) {
        return cleanText(getText(regName));
      }
    }

    // Último fallback: PartyName > Name (nombre comercial, solo si no hay razón social)
    const partyName = party.PartyName;
    if (partyName) {
      const name = Array.isArray(partyName) ? partyName[0]?.Name : partyName.Name;
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
 * Extrae subtotal e IVA de los totales monetarios
 */
function extractTotals(invoice: any): { subtotal: number; iva: number; total: number } {
  let subtotal = 0;
  let iva = 0;
  let total = 0;

  try {
    // LegalMonetaryTotal > LineExtensionAmount (subtotal antes de impuestos)
    const legalMonetaryTotal = invoice.LegalMonetaryTotal;
    if (legalMonetaryTotal) {
      const lineExtension = legalMonetaryTotal.LineExtensionAmount;
      subtotal = parseAmount(lineExtension);
      total = parseAmount(legalMonetaryTotal.PayableAmount);
    }

    // TaxTotal > TaxAmount (IVA)
    const taxTotal = invoice.TaxTotal;
    if (taxTotal) {
      const taxTotalArr = Array.isArray(taxTotal) ? taxTotal : [taxTotal];

      for (const tax of taxTotalArr) {
        const taxSubtotal = tax.TaxSubtotal;
        if (taxSubtotal) {
          const subtotals = Array.isArray(taxSubtotal) ? taxSubtotal : [taxSubtotal];

          for (const sub of subtotals) {
            const taxSchemeId = sub.TaxCategory?.TaxScheme?.ID;
            const taxId = String(getText(taxSchemeId));

            // "01" o "1" = IVA en nomenclatura DIAN
            if (taxId === "01" || taxId === "1" || /IVA/i.test(taxId)) {
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

  return { subtotal, iva, total };
}

/**
 * Extrae descripciones de los 2 primeros items (InvoiceLine con ID 1 y 2)
 */
function extractConcepts(invoice: any): string {
  try {
    const lineItems = extractLineItems(invoice);
    const descriptions = lineItems
      .slice(0, 2)
      .map((item) => item.description)
      .filter((desc) => !!desc && desc !== "N/A");

    if (descriptions.length === 0) return "N/A";
    return descriptions.join(", ");
  } catch {
    return "N/A";
  }
}

/**
 * Extrae líneas detalladas de la factura para la hoja "Detallado"
 */
function extractLineItems(invoice: any): InvoiceLineItem[] {
  try {
    const lines = invoice.InvoiceLine || invoice.CreditNoteLine;
    if (!lines) return [];

    const linesArr = Array.isArray(lines) ? lines : [lines];

    return linesArr.map((line: any, index: number) => {
      const lineNumber = parseAmount(line.ID) || index + 1;
      const quantity = parseAmount(line.InvoicedQuantity);
      const unitPrice = parseAmount(line.Price?.PriceAmount);
      const lineExtensionAmount = parseAmount(line.LineExtensionAmount);

      let description = "";
      const itemDesc = line.Item?.Description;
      if (Array.isArray(itemDesc)) {
        description = itemDesc.map((d) => getText(d)).filter(Boolean).join(" ");
      } else {
        description = getText(itemDesc);
      }
      description = cleanDescription(description.replace(/^\|+/, ""));

      let discount = 0;
      let surcharge = 0;
      const allowanceCharge = line.AllowanceCharge;
      if (allowanceCharge) {
        const allowances = Array.isArray(allowanceCharge) ? allowanceCharge : [allowanceCharge];
        for (const allowance of allowances) {
          const amount = parseAmount(allowance.Amount);
          const isCharge = String(getText(allowance.ChargeIndicator)).toLowerCase() === "true";
          if (isCharge) surcharge += amount;
          else discount += amount;
        }
      }

      let ivaAmount = 0;
      let ivaPercent = 0;
      let incAmount = 0;
      let incPercent = 0;

      const taxTotals = line.TaxTotal ? (Array.isArray(line.TaxTotal) ? line.TaxTotal : [line.TaxTotal]) : [];
      for (const taxTotal of taxTotals) {
        const subtotals = taxTotal.TaxSubtotal
          ? (Array.isArray(taxTotal.TaxSubtotal) ? taxTotal.TaxSubtotal : [taxTotal.TaxSubtotal])
          : [];

        for (const subtotal of subtotals) {
          const taxId = String(getText(subtotal.TaxCategory?.TaxScheme?.ID));
          const taxName = String(getText(subtotal.TaxCategory?.TaxScheme?.Name)).toUpperCase();
          const taxAmount = parseAmount(subtotal.TaxAmount);
          const taxPercent = parseAmount(subtotal.TaxCategory?.Percent);

          if (taxId === "01" || taxId === "1" || taxName.includes("IVA")) {
            ivaAmount += taxAmount;
            ivaPercent = taxPercent || ivaPercent;
          } else if (taxId === "04" || taxId === "4" || taxName.includes("INC")) {
            incAmount += taxAmount;
            incPercent = taxPercent || incPercent;
          }
        }
      }

      return {
        lineNumber,
        description: description || "N/A",
        quantity,
        unitPrice,
        discount,
        surcharge,
        ivaAmount,
        ivaPercent,
        incAmount,
        incPercent,
        totalUnitPrice: lineExtensionAmount,
      };
    });
  } catch {
    return [];
  }
}

/**
 * Extrae forma de pago del documento (PaymentMeans > PaymentMeansCode)
 * Códigos DIAN/UN-EDIFACT
 */
function extractPaymentMethod(invoice: any): string {
  try {
    const paymentMeans = invoice.PaymentMeans;
    if (!paymentMeans) return "N/A";

    const paymentArr = Array.isArray(paymentMeans) ? paymentMeans : [paymentMeans];
    const firstPayment = paymentArr[0];

    if (!firstPayment) return "N/A";

    const paymentCode = getText(firstPayment.PaymentMeansCode);
    
    // Mapeo de códigos DIAN/UN-EDIFACT a descripciones
    const paymentCodeMap: Record<string, string> = {
      "1": "Contado",
      "2": "Crédito",
      "3": "Débito",
      "10": "Efectivo",
      "20": "Cheque",
      "30": "Transferencia crédito",
      "31": "Transferencia débito",
      "32": "Transferencia bancaria",
      "42": "Pago a cuenta bancaria",
      "47": "Transferencia entre cuentas",
      "48": "Tarjeta crédito",
      "49": "Cargo automático",
      "50": "Giro postal",
      "60": "Pagaré",
      "70": "Retiro de caja",
      "ZZZ": "Otro",
      "41": "Transferencia electrónica",
    };

    if (paymentCode && paymentCodeMap[paymentCode]) {
      return paymentCodeMap[paymentCode];
    }

    if (paymentCode) {
      return `Código ${paymentCode}`;
    }

    return "N/A";
  } catch {
    return "N/A";
  }
}

/**
 * Extrae CUFE/CUDE del documento
 */
function extractCUFE(invoice: any): string {
  try {
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
 * Obtiene texto de un nodo XML (puede ser string, número, o objeto con #text)
 */
function getText(node: any): string {
  if (node === null || node === undefined) return "";
  if (typeof node === "string") return node;
  if (typeof node === "number") return String(node);
  if (typeof node === "object" && "#text" in node) return String(node["#text"]);
  return "";
}

/**
 * Parsea montos desde XML (pueden tener atributos de moneda o ser números directos)
 */
function parseAmount(node: any): number {
  if (typeof node === "number") return node;

  if (node && typeof node === "object" && "#text" in node) {
    if (typeof node["#text"] === "number") return node["#text"];
  }

  const text = getText(node);
  if (!text) return 0;

  const cleaned = text.replace(/[^\d.,\-]/g, "");
  const lastComma = cleaned.lastIndexOf(",");
  const lastDot = cleaned.lastIndexOf(".");

  let normalized = cleaned;
  if (lastComma > lastDot) {
    normalized = cleaned.replace(/\./g, "").replace(",", ".");
  } else if (lastDot > lastComma) {
    normalized = cleaned.replace(/,/g, "");
  }

  const result = parseFloat(normalized);
  return isNaN(result) ? 0 : result;
}

/**
 * Limpia texto de espacios extra
 */
function cleanText(text: string): string {
  return text.replace(/\s+/g, " ").trim();
}

/**
 * Limpia y formatea descripción de producto
 */
function cleanDescription(text: string): string {
  let cleaned = text.replace(/\s+/g, " ").trim();

  // Capitalizar primera letra, resto mantener
  if (cleaned.length > 0) {
    cleaned = cleaned.charAt(0).toUpperCase() + cleaned.slice(1);
  }

  return cleaned.substring(0, 100);
}
