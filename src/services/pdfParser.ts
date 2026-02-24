import pdf from "pdf-parse";
import type { InvoiceData } from "../types/dianExcel.js";

interface DocInfo {
  id: string;
  nit: string;
  docnum: string;
}

/**
 * Extrae datos estructurados de un PDF de factura DIAN
 */
export async function extractInvoiceData(
  pdfBuffer: Buffer,
  docInfo: DocInfo
): Promise<Partial<InvoiceData>> {
  try {
    const data = await pdf(pdfBuffer);
    const text = data.text;

    return {
      entityType: detectEntityType(text),
      issueDate: extractIssueDate(text),
      entityName: extractEntityName(text),
      subtotal: extractSubtotal(text),
      iva: extractIVA(text),
      concepts: extractConcepts(text),
      documentType: detectDocumentType(text),
      cufe: extractCUFE(text),
      trackId: docInfo.id,
      nit: docInfo.nit,
      docNumber: docInfo.docnum,
    };
  } catch (err) {
    console.error(`Error parseando PDF ${docInfo.docnum}:`, err);
    return {
      entityType: "N/A",
      issueDate: "N/A",
      entityName: "N/A",
      subtotal: 0,
      iva: 0,
      concepts: "ERROR: No se pudo leer el PDF",
      documentType: "N/A",
      cufe: "N/A",
      trackId: docInfo.id,
      nit: docInfo.nit,
      docNumber: docInfo.docnum,
      error: (err as Error).message,
    };
  }
}

// Detecta si es empresa o persona natural
function detectEntityType(text: string): "EMPRESA" | "PN" | "N/A" {
  const upperText = text.toUpperCase();

  // Indicadores de persona natural
  if (/PERSONA\s*NATURAL|R[EÉ]GIMEN\s*SIMPLIFICADO|NO\s*RESPONSABLE\s*DE\s*IVA/i.test(upperText)) {
    return "PN";
  }

  // Indicadores de empresa
  if (/S\.?A\.?S\.?|S\.?A\.?|LTDA\.?|LIMITADA|SOCIEDAD|CIA\.?|COMPA[NÑ][IÍ]A|E\.?U\.?|CORPORACI[OÓ]N/i.test(upperText)) {
    return "EMPRESA";
  }

  // Si tiene NIT con DV, probablemente empresa
  if (/NIT[:\s]*\d{9,10}[-\s]*\d/i.test(upperText)) {
    return "EMPRESA";
  }

  return "EMPRESA"; // Default
}

// Extrae fecha de emisión
function extractIssueDate(text: string): string {
  const patterns = [
    /Fecha\s+(?:de\s+)?[Ee]misi[oó]n\s*:?\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})/i,
    /Fecha\s+(?:de\s+)?[Gg]eneraci[oó]n\s*:?\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})/i,
    /Fecha\s*:?\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})/i,
    /(\d{1,2}[\/\-\.][A-Za-z]{3,}[\/\-\.]\d{2,4})/i, // 15/Ene/2024
    /(\d{4}[\/\-\.]\d{1,2}[\/\-\.]\d{1,2})/i, // 2024-01-15
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) {
      return formatDate(match[1]);
    }
  }

  return "N/A";
}

// Extrae razón social o nombre
function extractEntityName(text: string): string {
  const patterns = [
    /Raz[oó]n\s+[Ss]ocial\s*:?\s*([^\n\r]{3,80})/i,
    /Nombre\s+(?:del\s+)?[Ee]misor\s*:?\s*([^\n\r]{3,80})/i,
    /Emisor\s*:?\s*([^\n\r]{3,80})/i,
    /(?:Adquiriente|Cliente)\s*:?\s*([^\n\r]{3,80})/i,
    /Nombre\s*:?\s*([^\n\r]{3,80})/i,
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) {
      const name = match[1].trim();
      // Limpiar caracteres no deseados y limitar longitud
      const cleaned = name
        .replace(/[\t\r]+/g, " ")
        .replace(/\s{2,}/g, " ")
        .substring(0, 80);
      
      if (cleaned.length >= 3) {
        return cleaned;
      }
    }
  }

  return "N/A";
}

// Extrae subtotal/valor bruto
function extractSubtotal(text: string): number {
  const patterns = [
    /(?:Sub\s*-?\s*total|Valor\s+(?:Bruto|antes\s+de\s+IVA)|Base\s+Gravable|Total\s+Bruto)\s*:?\s*\$?\s*([\d.,]+)/i,
    /Valor\s+sin\s+(?:IVA|impuestos?)\s*:?\s*\$?\s*([\d.,]+)/i,
    /Base\s+imponible\s*:?\s*\$?\s*([\d.,]+)/i,
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) {
      return parseAmount(match[1]);
    }
  }

  // Fallback: buscar "Total" y restarle el IVA
  const totalMatch = text.match(/Total\s+(?:a\s+pagar|Factura)?\s*:?\s*\$?\s*([\d.,]+)/i);
  if (totalMatch) {
    const total = parseAmount(totalMatch[1]);
    const iva = extractIVA(text);
    if (iva > 0 && total > iva) {
      return total - iva;
    }
    return total;
  }

  return 0;
}

// Extrae valor del IVA
function extractIVA(text: string): number {
  const patterns = [
    /IVA\s*(?:\d{1,2}\s*%?)?\s*:?\s*\$?\s*([\d.,]+)/i,
    /Impuesto\s+(?:al\s+)?(?:Valor\s+)?(?:Agregado|IVA)\s*:?\s*\$?\s*([\d.,]+)/i,
    /Total\s+IVA\s*:?\s*\$?\s*([\d.,]+)/i,
    /Valor\s+(?:del\s+)?IVA\s*:?\s*\$?\s*([\d.,]+)/i,
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) {
      const value = parseAmount(match[1]);
      // IVA no debería ser mayor que el 19% de un valor razonable
      if (value > 0) {
        return value;
      }
    }
  }

  return 0;
}

// Extrae conceptos/descripción de items
function extractConcepts(text: string): string {
  // Buscar sección de productos/servicios/detalles
  const sectionPatterns = [
    /Detalles?\s+(?:de\s+)?Productos?([\s\S]+?)(?:Sub\s*-?\s*total|Total|Observaciones|Forma\s+de\s+pago)/i,
    /Descripci[oó]n\s+(?:de\s+)?(?:bienes|servicios|productos?)([\s\S]+?)(?:Sub\s*-?\s*total|Total)/i,
    /(?:Items?|L[ií]neas?|Detalle)([\s\S]+?)(?:Sub\s*-?\s*total|Total|Observaciones)/i,
  ];

  let detailsSection = "";
  for (const pattern of sectionPatterns) {
    const match = text.match(pattern);
    if (match) {
      detailsSection = match[1];
      break;
    }
  }

  if (!detailsSection) {
    // Fallback: buscar descripciones directamente
    const descMatch = text.match(/Descripci[oó]n\s*:?\s*([^\n\r]{5,100})/gi);
    if (descMatch && descMatch.length > 0) {
      const concepts = descMatch
        .map(d => d.replace(/Descripci[oó]n\s*:?\s*/i, "").trim())
        .filter(d => d.length > 3);
      
      return formatConcepts(concepts);
    }
    return "N/A";
  }

  // Extraer descripciones de la sección
  const descPattern = /Descripci[oó]n\s*:?\s*([^\n\r]{5,100})/gi;
  const concepts: string[] = [];
  let match;

  while ((match = descPattern.exec(detailsSection)) !== null) {
    const desc = match[1].trim();
    if (desc.length > 3 && !concepts.includes(desc)) {
      concepts.push(desc);
    }
  }

  // Si no encontramos con "Descripción:", buscar líneas que parezcan items
  if (concepts.length === 0) {
    const lines = detailsSection.split(/[\n\r]+/);
    for (const line of lines) {
      const cleaned = line.trim();
      // Líneas que parecen descripciones (no son números, ni muy cortas)
      if (
        cleaned.length > 10 &&
        cleaned.length < 100 &&
        !/^[\d\s.,\$%]+$/.test(cleaned) &&
        !/^(Cantidad|Precio|Valor|Total|IVA|Código|Und)/i.test(cleaned)
      ) {
        concepts.push(cleaned);
        if (concepts.length >= 3) break;
      }
    }
  }

  return formatConcepts(concepts);
}

// Formatea conceptos para el Excel
function formatConcepts(concepts: string[]): string {
  if (concepts.length === 0) return "N/A";
  
  if (concepts.length <= 2) {
    return concepts.join(", ");
  }

  // Más de 2 conceptos: mostrar primeros 2 + nota
  const first2 = concepts.slice(0, 2).join(", ");
  return `${first2} (+${concepts.length - 2} ítems más. Ver factura adjunta)`;
}

// Detecta tipo de documento
function detectDocumentType(text: string): "Factura Electrónica" | "Nota Crédito" | "N/A" {
  const upperText = text.toUpperCase();

  if (/NOTA\s*(?:DE\s*)?CR[EÉ]DITO|NC[-\s]?\d/i.test(upperText)) {
    return "Nota Crédito";
  }

  if (/FACTURA\s*(?:ELECTR[OÓ]NICA|DE\s*VENTA)?|FE[-\s]?\d|FVFE/i.test(upperText)) {
    return "Factura Electrónica";
  }

  return "Factura Electrónica"; // Default
}

// Extrae CUFE (código único)
function extractCUFE(text: string): string {
  // CUFE es típicamente una cadena hexadecimal de 96 caracteres
  const patterns = [
    /CUFE\s*:?\s*([a-f0-9]{96})/i,
    /C[oó]digo\s+[UÚ]nico\s*:?\s*([a-f0-9]{96})/i,
    /UUID\s*:?\s*([a-f0-9]{96})/i,
    // Buscar cualquier cadena de 96 hex chars que no sea parte de otra cosa
    /\b([a-f0-9]{96})\b/i,
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) {
      return match[1].toLowerCase();
    }
  }

  // Buscar cadenas más cortas que podrían ser CUFE parciales o diferentes formatos
  const shortPattern = /CUFE\s*:?\s*([a-f0-9-]{32,})/i;
  const shortMatch = text.match(shortPattern);
  if (shortMatch) {
    return shortMatch[1].replace(/-/g, "").toLowerCase();
  }

  return "N/A";
}

// Helpers
function parseAmount(str: string): number {
  if (!str) return 0;
  
  // Determinar formato (punto como miles o como decimal)
  const hasComma = str.includes(",");
  const hasDot = str.includes(".");
  
  let cleaned = str;
  
  if (hasComma && hasDot) {
    // Formato: 1.234.567,89 o 1,234,567.89
    if (str.lastIndexOf(",") > str.lastIndexOf(".")) {
      // Coma es decimal: 1.234.567,89
      cleaned = str.replace(/\./g, "").replace(",", ".");
    } else {
      // Punto es decimal: 1,234,567.89
      cleaned = str.replace(/,/g, "");
    }
  } else if (hasComma) {
    // Solo comas: podría ser 1,234 (miles) o 1234,56 (decimal)
    const parts = str.split(",");
    if (parts.length === 2 && parts[1].length <= 2) {
      // Es decimal: 1234,56
      cleaned = str.replace(",", ".");
    } else {
      // Son miles: 1,234,567
      cleaned = str.replace(/,/g, "");
    }
  } else if (hasDot) {
    // Solo puntos: podría ser 1.234 (miles) o 1234.56 (decimal)
    const parts = str.split(".");
    if (parts.length === 2 && parts[1].length <= 2) {
      // Es decimal, dejarlo así
    } else {
      // Son miles: 1.234.567
      cleaned = str.replace(/\./g, "");
    }
  }

  const result = parseFloat(cleaned);
  return isNaN(result) ? 0 : result;
}

function formatDate(dateStr: string): string {
  if (!dateStr) return "N/A";

  // Mapeo de meses en español
  const months: Record<string, string> = {
    ene: "01", enero: "01",
    feb: "02", febrero: "02",
    mar: "03", marzo: "03",
    abr: "04", abril: "04",
    may: "05", mayo: "05",
    jun: "06", junio: "06",
    jul: "07", julio: "07",
    ago: "08", agosto: "08",
    sep: "09", sept: "09", septiembre: "09",
    oct: "10", octubre: "10",
    nov: "11", noviembre: "11",
    dic: "12", diciembre: "12",
  };

  let cleaned = dateStr.toLowerCase().trim();

  // Reemplazar meses en texto por números
  for (const [name, num] of Object.entries(months)) {
    cleaned = cleaned.replace(new RegExp(name, "i"), num);
  }

  // Normalizar separadores
  cleaned = cleaned.replace(/[\.\-]/g, "/");

  const parts = cleaned.split("/").filter(Boolean);
  if (parts.length !== 3) return dateStr;

  let day: string, month: string, year: string;

  if (parts[0].length === 4) {
    // Formato YYYY/MM/DD
    [year, month, day] = parts;
  } else if (parts[2].length === 4) {
    // Formato DD/MM/YYYY
    [day, month, year] = parts;
  } else {
    // Formato DD/MM/YY
    [day, month, year] = parts;
    if (year.length === 2) {
      year = parseInt(year) > 50 ? `19${year}` : `20${year}`;
    }
  }

  return `${day.padStart(2, "0")}/${month.padStart(2, "0")}/${year}`;
}
