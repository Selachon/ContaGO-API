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

// Detecta si es empresa o persona natural basado en el emisor
function detectEntityType(text: string): "EMPRESA" | "PN" | "N/A" {
  // Buscar en la sección del emisor
  const emisorSection = text.match(/Datos del Emisor[\s\S]*?Datos del Adquiriente/i)?.[0] || text;
  
  // Indicadores de persona natural
  if (/Persona\s*Natural/i.test(emisorSection)) {
    return "PN";
  }
  
  // Indicadores de empresa (tipo de contribuyente)
  if (/Persona\s*Jur[ií]dica/i.test(emisorSection)) {
    return "EMPRESA";
  }

  // Fallback: buscar sufijos de empresa en razón social
  const razonSocial = text.match(/Raz[oó]n\s+Social\s*:\s*([^\n]+)/i)?.[1] || "";
  if (/\b(S\.?A\.?S\.?|S\.?A\.?|LTDA\.?|E\.?U\.?|SAS|LIMITADA)\b/i.test(razonSocial)) {
    return "EMPRESA";
  }

  return "EMPRESA"; // Default
}

// Extrae fecha de emisión
function extractIssueDate(text: string): string {
  // Buscar específicamente "Fecha de Emisión:"
  const match = text.match(/Fecha\s+de\s+Emisi[oó]n\s*:\s*(\d{1,2}\/\d{1,2}\/\d{4})/i);
  if (match) {
    return match[1];
  }
  
  // Fallback
  const fallback = text.match(/(\d{1,2}\/\d{1,2}\/\d{4})/);
  return fallback ? fallback[1] : "N/A";
}

// Extrae razón social del EMISOR (vendedor)
function extractEntityName(text: string): string {
  // Buscar la primera "Razón Social:" que aparece (es del emisor)
  const match = text.match(/Raz[oó]n\s+Social\s*:\s*([^\n]+)/i);
  if (match) {
    const name = match[1].trim();
    // Limpiar y limitar
    const cleaned = name
      .replace(/Nombre\s+Comercial.*/i, "") // Quitar si sigue "Nombre Comercial"
      .replace(/[\t\r]+/g, " ")
      .replace(/\s{2,}/g, " ")
      .trim()
      .substring(0, 80);
    
    if (cleaned.length >= 3) {
      return cleaned;
    }
  }
  return "N/A";
}

// Extrae subtotal de la sección con valores reales (la que tiene COP)
function extractSubtotal(text: string): number {
  // Buscar en la sección que tiene "COP" (la tabla con valores reales)
  // El patrón es: MONEDACOP ... Subtotal<valor>
  const copSection = text.match(/MONEDACOP[\s\S]*?Total\s+factura/i)?.[0] || "";
  
  if (copSection) {
    // Buscar Subtotal seguido de un número
    const subtotalMatch = copSection.match(/Subtotal([\d.,]+)/i);
    if (subtotalMatch) {
      return parseAmount(subtotalMatch[1]);
    }
  }
  
  // Fallback: buscar "Total Bruto Factura" con valor
  const brutoMatch = text.match(/Total\s+Bruto\s+Factura([\d.,]+)/i);
  if (brutoMatch) {
    const value = parseAmount(brutoMatch[1]);
    if (value > 0) return value;
  }
  
  return 0;
}

// Extrae IVA de la sección con valores reales (la que tiene COP)
function extractIVA(text: string): number {
  // Buscar en la sección que tiene "COP" (la tabla con valores reales)
  const copSection = text.match(/MONEDACOP[\s\S]*?Total\s+factura/i)?.[0] || "";
  
  if (copSection) {
    // Buscar IVA seguido de un número (no 0,00)
    const ivaMatch = copSection.match(/\bIVA([\d.,]+)/i);
    if (ivaMatch) {
      const value = parseAmount(ivaMatch[1]);
      if (value > 0) return value;
    }
  }
  
  // Fallback: buscar "Total impuesto" con valor
  const impuestoMatch = text.match(/Total\s+impuesto\s*\(=\)([\d.,]+)/i);
  if (impuestoMatch) {
    const value = parseAmount(impuestoMatch[1]);
    if (value > 0) return value;
  }
  
  return 0;
}

// Extrae descripciones de la tabla "Detalles de Productos"
function extractConcepts(text: string): string {
  // Buscar la sección "Detalles de Productos" hasta "Hoja X de Y" o "Notas Finales" o "Datos Totales"
  const detailsMatch = text.match(/Detalles\s+de\s+Productos([\s\S]*?)(?:Hoja\s+\d|Notas\s+Finales|Datos\s+Totales|\$\$)/i);
  
  if (!detailsMatch) {
    return "N/A";
  }
  
  const detailsSection = detailsMatch[1];
  const descriptions: string[] = [];
  
  // Primero, intentar extraer de líneas que tienen todo junto
  // Patrón: número + código + DESCRIPCIÓN + unidad + números
  // Ejemplo: "101030701LIBROS DE COMERCIOZZ2,0024.200,00"
  const inlinePattern = /\d{1,2}[0-9]{6,12}([A-ZÁÉÍÓÚÑ][A-Za-záéíóúñÁÉÍÓÚÑ\s]+?)(ZZ|UN|KG|LT|MT|EA|94)\d/gi;
  let match;
  while ((match = inlinePattern.exec(detailsSection)) !== null) {
    let desc = match[1].trim();
    if (desc.length >= 3) {
      desc = cleanDescription(desc);
      if (!descriptions.some(d => d.toLowerCase() === desc.toLowerCase())) {
        descriptions.push(desc);
      }
    }
  }
  
  // Segundo, buscar descripciones en líneas separadas
  // Agrupar líneas consecutivas que son texto (no números ni códigos)
  const lines = detailsSection.split(/[\n\r]+/).map(l => l.trim()).filter(l => l.length > 0);
  let textBuffer: string[] = [];
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    
    // Ignorar headers de la tabla
    // Líneas cortas que son headers exactos
    if (/^(Nro\.?|C[oó]digo|Descripci[oó]n|U\/M|Cantidad|Precio|Descuento|Recargo|IVA|INC|IMPUESTOS|venta|unitario|detalle)$/i.test(line)) {
      continue;
    }
    // Líneas que son combinaciones de headers (ej: "Nro.CódigoDescripciónU/M...")
    if (/^(Nro\.?)?C[oó]digo.*Descripci[oó]n.*U\/M/i.test(line) ||
        /Precio\s*unitario/i.test(line) && /Cantidad/i.test(line) ||
        /^Recargo\s+detalle/i.test(line) ||
        /^unitario\s+de$/i.test(line) ||
        /^IMPUESTOS$/i.test(line)) {
      continue;
    }
    
    // Si es un código (solo números o alfanumérico corto), procesar buffer anterior
    const isCode = /^[0-9]{3,10}$/.test(line) || /^[A-Z0-9]{6,15}$/i.test(line);
    // Si es línea de datos (muchos números)
    const isDataLine = /^(ZZ|UN|KG|LT|MT|EA|94)?\d/.test(line) && (line.match(/\d/g) || []).length > 5;
    
    if (isCode || isDataLine) {
      // Procesar texto acumulado
      if (textBuffer.length > 0) {
        let desc = textBuffer.join(" ");
        desc = cleanDescription(desc);
        if (desc.length >= 3 && !descriptions.some(d => d.toLowerCase() === desc.toLowerCase())) {
          descriptions.push(desc);
        }
        textBuffer = [];
      }
      continue;
    }
    
    // Si es texto con letras, acumularlo
    if (/[a-záéíóúñ]{2,}/i.test(line) && line.length >= 2) {
      textBuffer.push(line);
    }
  }
  
  // Procesar último buffer - pero solo si tiene contenido significativo
  if (textBuffer.length > 0) {
    let desc = textBuffer.join(" ");
    desc = cleanDescription(desc);
    // Solo agregar si tiene al menos 5 caracteres (evita fragmentos sueltos como "pelería")
    // Y no es similar a una descripción existente
    if (desc.length >= 5 && !descriptions.some(d => isSimilarDescription(d, desc))) {
      descriptions.push(desc);
    }
  }
  
  return formatConcepts(descriptions);
}

// Compara si dos descripciones son similares (una contiene a la otra o son casi iguales)
function isSimilarDescription(a: string, b: string): boolean {
  const cleanA = a.toLowerCase().replace(/\s/g, "");
  const cleanB = b.toLowerCase().replace(/\s/g, "");
  
  // Exactamente iguales
  if (cleanA === cleanB) return true;
  
  // Una contiene a la otra
  if (cleanA.includes(cleanB) || cleanB.includes(cleanA)) return true;
  
  // Muy similares (diferencia de pocos caracteres)
  if (Math.abs(cleanA.length - cleanB.length) <= 3) {
    let diff = 0;
    const minLen = Math.min(cleanA.length, cleanB.length);
    for (let i = 0; i < minLen; i++) {
      if (cleanA[i] !== cleanB[i]) diff++;
    }
    diff += Math.abs(cleanA.length - cleanB.length);
    return diff <= 3;
  }
  
  return false;
}

// Limpia una descripción de texto duplicado y formatea
function cleanDescription(text: string): string {
  let cleaned = text.replace(/\s+/g, " ").trim();
  
  // PASO 1: Unir palabras cortadas por el PDF
  // El PDF corta palabras así: "Carr o" o "pa pelería"
  cleaned = joinFragments(cleaned);
  
  // PASO 2: Detectar y eliminar duplicación
  // El PDF de DIAN duplica el texto: "Servicio Parqueadero Carro Servicio Parqueadero Carro"
  const words = cleaned.split(" ");
  
  for (let splitPoint = Math.floor(words.length / 2); splitPoint >= 1; splitPoint--) {
    const firstPart = words.slice(0, splitPoint).join("").toLowerCase();
    const secondPart = words.slice(splitPoint, splitPoint * 2).join("").toLowerCase();
    
    if (firstPart.length >= 3 && areSimilar(firstPart, secondPart)) {
      cleaned = words.slice(0, splitPoint).join(" ");
      break;
    }
  }
  
  // PASO 3: Capitalizar primera letra, resto en minúsculas
  if (cleaned.length > 0) {
    cleaned = cleaned.charAt(0).toUpperCase() + cleaned.slice(1).toLowerCase();
  }
  
  return cleaned;
}

// Une fragmentos de palabras cortadas por el PDF
function joinFragments(text: string): string {
  const words = text.split(/\s+/);
  const result: string[] = [];
  const prepositions = ["de", "del", "la", "el", "los", "las", "en", "con", "para", "por", "al", "un", "una", "y", "a", "e"];
  
  for (let i = 0; i < words.length; i++) {
    const word = words[i];
    const nextWord = words[i + 1];
    
    if (!nextWord) {
      result.push(word);
      continue;
    }
    
    // Check if next word is a short fragment (1-3 lowercase letters)
    if (nextWord.length <= 3 && /^[a-záéíóúñ]+$/.test(nextWord)) {
      const isPreposition = prepositions.includes(nextWord.toLowerCase());
      const endsInConsonant = /[bcdfghjklmnpqrstvwxyz]$/i.test(word);
      
      // Single letter 'o' after consonant is likely a fragment, not preposition
      if (nextWord.toLowerCase() === "o" && endsInConsonant) {
        result.push(word + nextWord);
        i++;
        continue;
      }
      
      // If not a preposition, join it
      if (!isPreposition) {
        result.push(word + nextWord);
        i++;
        continue;
      }
    }
    
    result.push(word);
  }
  
  return result.join(" ");
}

// Compara dos strings ignorando pequeñas diferencias (espacios, mayúsculas)
function areSimilar(a: string, b: string): boolean {
  // Eliminar espacios y convertir a minúsculas para comparar
  const cleanA = a.replace(/\s/g, "").toLowerCase();
  const cleanB = b.replace(/\s/g, "").toLowerCase();
  
  if (cleanA === cleanB) return true;
  
  // Permitir diferencia de hasta 3 caracteres (para manejar letras cortadas)
  if (Math.abs(cleanA.length - cleanB.length) <= 3) {
    let differences = 0;
    const minLen = Math.min(cleanA.length, cleanB.length);
    for (let i = 0; i < minLen; i++) {
      if (cleanA[i] !== cleanB[i]) differences++;
    }
    differences += Math.abs(cleanA.length - cleanB.length);
    return differences <= 3;
  }
  
  return false;
}

// Formatea conceptos para el Excel
function formatConcepts(concepts: string[]): string {
  if (concepts.length === 0) return "N/A";
  
  // Limpiar conceptos duplicados o muy similares
  const unique = concepts.filter((c, i) => 
    concepts.findIndex(x => x.toLowerCase() === c.toLowerCase()) === i
  );
  
  if (unique.length === 1) {
    return unique[0];
  }
  
  if (unique.length === 2) {
    return unique.join(", ");
  }

  // Más de 2 conceptos: mostrar primeros 2 + nota
  const first2 = unique.slice(0, 2).join(", ");
  return `${first2}... (+${unique.length - 2} ítems más)`;
}

// Detecta tipo de documento basado en el título del PDF
function detectDocumentType(text: string): "Factura Electrónica" | "Nota Crédito" | "N/A" {
  // Buscar en las primeras líneas del documento (el título)
  const firstPart = text.substring(0, 500).toUpperCase();
  
  // Nota Crédito tiene prioridad (es más específico)
  if (/NOTA\s*(DE\s*)?CR[EÉ]DITO/i.test(firstPart)) {
    return "Nota Crédito";
  }
  
  // Factura Electrónica
  if (/FACTURA\s*ELECTR[OÓ]NICA/i.test(firstPart)) {
    return "Factura Electrónica";
  }
  
  // Fallback: buscar en todo el texto
  if (/NOTA\s*(DE\s*)?CR[EÉ]DITO/i.test(text)) {
    return "Nota Crédito";
  }
  
  return "Factura Electrónica"; // Default
}

// Extrae CUFE (código único de 96 caracteres hexadecimales)
function extractCUFE(text: string): string {
  // Buscar después de "CUFE :" o "Código Único de Factura"
  const patterns = [
    /C[oó]digo\s+[UÚ]nico\s+de\s+Factura\s*[-:]?\s*(?:CUFE\s*:?)?\s*([a-f0-9]{96})/i,
    /CUFE\s*:?\s*([a-f0-9]{96})/i,
    /\b([a-f0-9]{96})\b/i,
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) {
      return match[1].toLowerCase();
    }
  }

  return "N/A";
}

// Elimina texto duplicado que aparece en algunos PDFs de DIAN
function removeDuplicateText(text: string): string {
  // Normalizar espacios primero
  let cleaned = text.replace(/\s+/g, " ").trim();
  
  // Estrategia 1: Buscar repetición exacta de la mitad del texto
  const words = cleaned.split(" ");
  const len = words.length;
  
  if (len >= 4) {
    // Probar diferentes puntos de corte
    for (let splitPoint = Math.floor(len / 2); splitPoint >= 2; splitPoint--) {
      const firstPart = words.slice(0, splitPoint).join(" ").toLowerCase();
      const secondPart = words.slice(splitPoint, splitPoint * 2).join(" ").toLowerCase();
      
      if (firstPart === secondPart) {
        cleaned = words.slice(0, splitPoint).join(" ");
        break;
      }
    }
  }
  
  // Estrategia 2: Buscar patrón de palabras cortadas y repetidas
  // Ej: "Parqueadero Carr o Servicio Parqueadero Ca rro" -> "Servicio Parqueadero Carro"
  // Esto ocurre cuando el PDF tiene texto en columnas que se mezcla
  
  // Unir palabras cortadas (letra sola seguida de palabra)
  cleaned = cleaned.replace(/\b([A-Za-záéíóúñ])\s+([a-záéíóúñ])/gi, "$1$2");
  
  // Si aún hay duplicación parcial, intentar detectarla
  const cleanedWords = cleaned.split(" ");
  if (cleanedWords.length >= 4) {
    // Buscar si la segunda mitad contiene las mismas palabras
    const half = Math.ceil(cleanedWords.length / 2);
    const firstHalf = cleanedWords.slice(0, half);
    const secondHalf = cleanedWords.slice(half);
    
    // Contar palabras en común
    const commonWords = firstHalf.filter(w => 
      secondHalf.some(w2 => w2.toLowerCase().includes(w.toLowerCase()) || w.toLowerCase().includes(w2.toLowerCase()))
    );
    
    // Si más del 50% de las palabras son comunes, probablemente hay duplicación
    if (commonWords.length >= firstHalf.length * 0.5) {
      // Tomar la parte más larga/completa
      const firstStr = firstHalf.join(" ");
      const secondStr = secondHalf.join(" ");
      cleaned = firstStr.length >= secondStr.length ? firstStr : secondStr;
    }
  }
  
  return cleaned.trim();
}

// Parsea montos en formato colombiano (1.234.567,89 o 1234567.89)
function parseAmount(str: string): number {
  if (!str) return 0;
  
  // Limpiar espacios
  let cleaned = str.trim();
  
  // Determinar formato basado en la posición de punto y coma
  const lastComma = cleaned.lastIndexOf(",");
  const lastDot = cleaned.lastIndexOf(".");
  
  if (lastComma > lastDot) {
    // Formato europeo/colombiano: 1.234.567,89
    // Coma es decimal, puntos son miles
    cleaned = cleaned.replace(/\./g, "").replace(",", ".");
  } else if (lastDot > lastComma) {
    // Formato americano: 1,234,567.89
    // Punto es decimal, comas son miles
    cleaned = cleaned.replace(/,/g, "");
  } else if (lastComma !== -1) {
    // Solo comas: verificar si es decimal o miles
    const parts = cleaned.split(",");
    if (parts.length === 2 && parts[1].length <= 2) {
      // Es decimal: 1234,56
      cleaned = cleaned.replace(",", ".");
    } else {
      // Son miles: 1,234,567
      cleaned = cleaned.replace(/,/g, "");
    }
  } else if (lastDot !== -1) {
    // Solo puntos: verificar si es decimal o miles
    const parts = cleaned.split(".");
    if (parts.length === 2 && parts[1].length <= 2) {
      // Es decimal, dejarlo así
    } else {
      // Son miles: 1.234.567
      cleaned = cleaned.replace(/\./g, "");
    }
  }

  const result = parseFloat(cleaned);
  return isNaN(result) ? 0 : result;
}
