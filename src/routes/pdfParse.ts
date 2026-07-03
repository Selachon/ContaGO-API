import { Router, type Request, type Response } from "express";
import multer from "multer";
import pdfParse from "pdf-parse";
import path from "path";
import fs from "fs";
import os from "os";
import { requireAuth } from "../middleware/auth.js";
import { generateExcelFile } from "../services/excelGenerator.js";
import type { InvoiceData, TaxDetail, InvoiceLineItem } from "../types/dianExcel.js";
import { parseLineItemsPositional } from "../services/pdfPositionalParser.js";
import { parseDianListing, type DianTaxRow } from "../services/dianListingParser.js";
import { uploadInvoiceFilesToDrive } from "../services/googleDrive.js";
import { getUserGoogleDriveById, updateUserDriveTokens } from "../services/database.js";
import { encryptToken } from "../utils/encryption.js";

const router = Router();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024, files: 5001 },
});

function parseCOP(s: string | null | undefined): number {
  if (!s) return 0;
  // Remove thousand separators (.) and convert decimal comma to dot
  return parseFloat(s.trim().replace(/\./g, "").replace(",", ".")) || 0;
}

function toISO(ddMMyyyy: string): string {
  const m = ddMMyyyy.match(/(\d{2})\/(\d{2})\/(\d{4})/);
  return m ? `${m[3]}-${m[2]}-${m[1]}` : "";
}

// Common U/M (unit of measure) codes used in Colombian e-invoices
const UM_CODES = "NIU|EA|UND|UNI|KGM|GLO|PK|PAR|C62|BX|PAC|SET|ACT|SRV|HUR|DAY|MON|ANO|LTR|MLT|GRM|MTR|CMT|HLT|TNE|DZN|PCK|CAR|BOX|ZZ|NAR|XUN|CCT|MMT|MGM|CGM|BO|BG|BJ|ST|PR|CJ|UN";
// Match U/M code immediately followed by a digit or comma (start of Cantidad)
// Cannot use \b because the code is glued to description (no word boundary before it)
const UM_RE = new RegExp(`(${UM_CODES})(?=\\d|,)`);

// Fix font kerning in numbers: "2. 095, 00" → "2.095,00"
function collapseNumericSpaces(s: string): string {
  return s.replace(/(\d)\.\s+(\d)/g, "$1.$2").replace(/(\d),\s+(\d)/g, "$1,$2");
}

// IVA/INC percentages use dot as decimal separator (not the Colombian thousands dot)
function parsePct(s: string): number {
  return parseFloat(s?.trim() || "0") || 0;
}

function parseLineItems(sectionText: string): InvoiceLineItem[] {
  // Repair split data lines: some PDFs put the DIAN double-space separator and
  // the unit price on their own lines instead of inline.
  // e.g.  "EA4,00\n15.564,71…1719.00\n  \n62.258,82"
  //    →  "EA4,00\n15.564,71…1719.00  62.258,82"
  const repaired = collapseNumericSpaces(sectionText)
    .replace(/\n {2,}\n(\d[\d.,]*)/g, "  $1");
  const lines = repaired.split("\n").map((l) => l.trim()).filter(Boolean);
  const rows: string[] = [];
  let current: string | null = null;

  for (const line of lines) {
    const sN = /^\d/.test(line);
    const pN = /^\d+$/.test(line);          // pure digits only (e.g. "1", "11001512")
    const hasUM = UM_RE.test(line);
    const hasDbl = /  /.test(line);
    // A "complete" item is a single line that already has text description + numeric data:
    // starts with digit, has a letter (description text), and has U/M or double-space
    const isComplete = sN && /[A-Za-záéíóúüñÁÉÍÓÚÜÑ]/.test(line) && (hasUM || hasDbl);

    if (current === null) {
      if (pN || isComplete) {
        current = line;
        if (isComplete) { rows.push(current); current = null; }
      }
    } else {
      if (isComplete) {
        // New complete single-line item encountered; push previous, push this one too
        rows.push(current);
        rows.push(line);
        current = null;
      } else if (hasDbl || (hasUM && hasDbl)) {
        // Completion/data line: must have double-space (separator before unit price).
        // A line with only hasUM (e.g. "EA4,00") but no double-space means the price
        // data follows on the next line — treat as description continuation instead.
        current += line;
        rows.push(current);
        current = null;
      } else if (pN) {
        // Pure digits: new item number OR fragment of the previous description line.
        // Continuation cases:
        //   (a) current ends with a letter  → "MALBORO X" + "10" → "MALBORO X10"
        //   (b) single digit (0-9)          → "CIGARRILLO MUSTANG X1" + "0" → "X10"
        //       Single-digit item numbers always start a NEW block (not mid-description),
        //       so a lone digit here is always a description suffix, not an item number.
        const isSingleDigit = line.length === 1;
        if (/[A-Za-záéíóúüñÁÉÍÓÚÜÑ]$/.test(current.trimEnd()) || isSingleDigit) {
          current += line;
        } else {
          rows.push(current);
          current = line;
        }
      } else {
        // Description continuation
        current += " " + line;
      }
    }
  }
  if (current !== null) rows.push(current);

  const items: InvoiceLineItem[] = [];
  // Colombian COP number: 1-3 digits + optional (.NNN)+ then ,NN
  const COP_RE = /\d{1,3}(?:\.\d{3})*,\d{2}/g;
  // Decimal-dot percentage: digits + dot + exactly 2 decimal digits (not followed by digit or comma)
  // e.g. "19.00", "5.00" but NOT "85.528" (has 3 decimal digits)
  const PCT_RE = /\d+\.\d{2}(?![\d,])/g;

  for (const row of rows) {
    // Line extension: always the last Colombian COP number after double-space
    const dblIdx = row.lastIndexOf("  ");
    if (dblIdx < 0) continue;
    const lineExtension = parseCOP(row.slice(dblIdx).trim().replace(/\s/g, ""));
    if (lineExtension <= 0) continue;
    const beforeDbl = row.slice(0, dblIdx);

    let nro = 0, description = "", suffix = "";

    const umMatch = UM_RE.exec(row);
    if (umMatch) {
      // Path A: U/M code found (PDF 1, PDF 2 style) — use as anchor
      const umIdx = umMatch.index!;
      const prefix = beforeDbl.slice(0, umIdx) || row.slice(0, umIdx);
      suffix = row.slice(umIdx + umMatch[0].length, dblIdx);

      const nroM = prefix.match(/^(\d+)/);
      const rawNroStr = nroM?.[1] || "0";
      const nro0 = rawNroStr.length >= 5 ? 1 : parseInt(rawNroStr.length > 3 ? rawNroStr.slice(0, -4) || rawNroStr[0] : rawNroStr);
      nro = nro0;
      const afterNro = prefix.slice(rawNroStr.length);
      const codigoM = afterNro.match(/^(\d{3,8})/);
      const codigo = codigoM?.[1] || "";
      description = afterNro.slice(codigo.length).trim();
    } else {
      // Path B: no standard U/M — anchor on first Colombian COP number
      const copAll = [...beforeDbl.matchAll(COP_RE)];
      if (!copAll.length) continue;
      const firstCopIdx = copAll[0].index!;
      const descPrefix = row.slice(0, firstCopIdx);
      const nroM = descPrefix.match(/^(\d+)/);
      const rawNroStr = nroM?.[1] || "0";
      nro = rawNroStr.length >= 5 ? parseInt(rawNroStr[0]) : parseInt(rawNroStr);
      const afterNro = descPrefix.slice(rawNroStr.length);
      const codigoM = afterNro.match(/^(\d{3,8})/);
      const codigo = codigoM?.[1] || "";
      description = afterNro.slice(codigo.length).trim();
      suffix = row.slice(firstCopIdx, dblIdx);
    }

    if (!description) continue;

    // COP numbers in beforeDbl: qty, PVP_ref, discount, surcharge, IVA_amount
    const copNums = [...beforeDbl.matchAll(COP_RE)].map((m) => parseCOP(m[0]));
    if (!copNums.length) continue;

    const cantidad = copNums[0];
    if (cantidad <= 0) continue;

    // Percentage values: remove COP numbers first, then match decimal-dot numbers.
    // This avoids "85.528,575.00" being greedily matched as "575.00" = 575%.
    const withoutCOP = beforeDbl.replace(COP_RE, " ");
    const pctNums = [...withoutCOP.matchAll(PCT_RE)].map((m) => parseFloat(m[0]));
    const ivaPercent = pctNums[pctNums.length - 1] ?? 0;
    const incPercent = pctNums.length >= 2 ? pctNums[pctNums.length - 2] : 0;

    // IVA amount: last COP number before the double-space
    const ivaAmount = copNums.length >= 2 ? copNums[copNums.length - 1] : 0;

    const unitPrice = cantidad > 0 ? lineExtension / cantidad : lineExtension;

    const taxes: TaxDetail[] = [];
    if (ivaPercent > 0 || ivaAmount > 0) {
      taxes.push({ taxId: "01", taxName: "IVA", amount: ivaAmount, percent: ivaPercent });
    }
    if (incPercent > 0) {
      taxes.push({ taxId: "04", taxName: "INC", amount: 0, percent: incPercent });
    }

    items.push({
      lineNumber: nro,
      description,
      quantity: cantidad,
      unitPrice,
      discount: 0,
      surcharge: 0,
      taxes,
      ivaAmount,
      ivaPercent,
      incAmount: 0,
      incPercent,
      totalUnitPrice: lineExtension,
      taxableBase: lineExtension,
    });
  }

  return items;
}

function parsePdfText(rawText: string, filename: string): InvoiceData {
  // Normalize: remove $$$$$ aesthetic artifacts (2+ consecutive $), keep single $ (e.g. "COP $34.111,00")
  const t = rawText.replace(/\$\$+/g, " ").replace(/\r/g, "\n");

  // ── CUFE: prefer the value embedded in the PDF text (authoritative even if the file was
  //    mis-named due to a concurrent-download race), fall back to filename.
  const cufeFromText = t.match(/(?:CUFE|CUDE)\s*:?\s*\n?\s*([a-f0-9A-F]{80,})/i)?.[1]?.trim();
  const cufe = cufeFromText || filename.replace(/\.pdf$/i, "").trim();

  // ── Document type ─────────────────────────────────────────────────────────────
  let documentType = "Factura electrónica";
  if (/NOTA\s*CR[ÉE]DITO/i.test(t)) documentType = "Nota Crédito";
  else if (/NOTA\s*D[ÉE]BITO/i.test(t)) documentType = "Nota Débito";
  else if (/DOCUMENTO\s*SOPORTE/i.test(t)) documentType = "Documento soporte";
  else if (/DOCUMENTO\s+EQUIVALENTE\s+POS/i.test(t)) documentType = "Documento equivalente POS";

  const isPOS = documentType === "Documento equivalente POS";

  // ── Invoice number — runs directly into "Forma de pago:" on same line ─────────
  let docNumber = t.match(/Número de Factura:\s*([^\n]+?)(?:Forma de pago:|\n|$)/i)?.[1]?.trim() || "";
  if (!docNumber) docNumber = t.match(/Número de documento:\s*([^\n]+?)(?:\n|$)/i)?.[1]?.trim() || "";

  // ── Issue date ────────────────────────────────────────────────────────────────
  let issueDate = t.match(/Fecha de Emisión:\s*(\d{2}\/\d{2}\/\d{4})/i)?.[1] || "";
  if (!issueDate) {
    const posDate = t.match(/Fecha y hora de expedición:\s*(\d{4})-(\d{2})-(\d{2})/i);
    if (posDate) issueDate = `${posDate[3]}/${posDate[2]}/${posDate[1]}`;
  }
  let issueDateISO = toISO(issueDate);

  // ── Payment method ────────────────────────────────────────────────────────────
  const paymentMethod = t.match(/Forma de pago:\s*([^\n]+?)(?:\s{2,}|\n|Fecha|$)/i)?.[1]?.trim() || "";

  // ── Emisor / Receptor sections ────────────────────────────────────────────────
  // Note: Notas Crédito use lowercase "emisor / vendedor"; use case-insensitive search
  const tl = t.toLowerCase();
  const emisorStart   = tl.indexOf("datos del emisor");
  const vendedorStart = tl.indexOf("datos del vendedor");
  const receptorStart = tl.indexOf("datos del adquiriente");
  const detallesStart = tl.indexOf("detalles de productos");

  // POS uses "Datos del vendedor" instead of "Datos del Emisor / Vendedor"
  const efectiveEmisorStart = emisorStart >= 0 ? emisorStart : vendedorStart;
  let emisorBlock = efectiveEmisorStart >= 0
    ? t.slice(efectiveEmisorStart, receptorStart > efectiveEmisorStart ? receptorStart : undefined)
    : "";
  const receptorBlock = receptorStart >= 0
    ? t.slice(receptorStart, detallesStart > receptorStart ? detallesStart : undefined)
    : "";

  // Stop pattern: any known DIAN label (right- or left-column) or end-of-line.
  // Using a lookahead instead of consuming the next label avoids the i-flag bug
  // where [a-z]+ with /i also matches uppercase letters and stops too early.
  const STOP = `(?=País:|Departamento:|Municipio|Dirección:|Teléfono|Correo:|Nit del|Nombre Comercial:|Tipo de|Régimen|Responsabilidad|Actividad|Número|\\n|$)`;
  const field = (block: string, label: string) =>
    block.match(new RegExp(`${label}\\s*(.+?)${STOP}`, "i"))?.[1]?.trim() || "";

  // Razón Social can run into "Nombre Comercial:" on the same line
  let issuerName = (
    emisorBlock.match(/Razón [Ss]ocial:\s*([^\n]+?)(?:Nombre Comercial:|Tipo de documento:|Número de documento:|\n|$)/i)?.[1]?.trim() || ""
  );
  let issuerNit = field(emisorBlock, "Nit del Emisor:");
  if (!issuerNit) issuerNit = field(emisorBlock, "Número de documento:");
  const issuerCommercialName     = field(emisorBlock,    "Nombre Comercial:");
  const issuerTaxpayerType       = field(emisorBlock,    "Tipo de Contribuyente:");
  const issuerFiscalRegime       = field(emisorBlock,    "Régimen Fiscal:");
  const issuerTaxResponsibility  = field(emisorBlock,    "Responsabilidad tributaria:");
  const issuerEconomicActivity   = field(emisorBlock,    "Actividad Económica:");
  const issuerCountry            = field(emisorBlock,    "País:");
  const issuerDepartment         = field(emisorBlock,    "Departamento:");
  const issuerCity               = field(emisorBlock,    "Municipio / Ciudad:");
  const issuerAddress            = field(emisorBlock,    "Dirección:");
  const issuerPhone              = field(emisorBlock,    "Teléfono / Móvil:");
  const issuerEmail              = field(emisorBlock,    "Correo:");

  let receiverName = receptorBlock.match(/Nombre o Razón Social:\s*([^\n]+)/i)?.[1]?.trim() || "";
  if (!receiverName) receiverName = receptorBlock.match(/Razón social:\s*([^\n]+?)(?:NIT del|Tipo de|\n|$)/i)?.[1]?.trim() || "";
  let receiverNit = field(receptorBlock, "Número Documento:");
  if (!receiverNit) receiverNit = receptorBlock.match(/NIT del adquiriente:\s*([^\n]+?)(?:\n|$)/i)?.[1]?.trim() || "";
  const receiverTaxpayerType     = field(receptorBlock,  "Tipo de Contribuyente:");
  const receiverFiscalRegime     = field(receptorBlock,  "Régimen [Ff]iscal:");
  const receiverTaxResponsibility= field(receptorBlock,  "Responsabilidad tributaria:");
  const receiverCountry          = field(receptorBlock,  "País:");
  const receiverDepartment       = field(receptorBlock,  "Departamento:");
  const receiverCity             = field(receptorBlock,  "Municipio / Ciudad:");
  const receiverAddress          = field(receptorBlock,  "Dirección:");
  const receiverPhone            = field(receptorBlock,  "Teléfono / Móvil:");
  const receiverEmail            = field(receptorBlock,  "Correo:");

  // ── Totals: the PDF renders two copies of the summary table; the second (after
  // "MONEDACOP") has the actual values.  Some fonts space it as "MO NEDA CO P".
  // Anchor on the last occurrence of "Subtotal<digit>" (template copy has "Subtotal\n").
  const subtotalMatches = [...t.matchAll(/Subtotal([\d.,]+)/g)];
  const lastSubtotalMatch = subtotalMatches[subtotalMatches.length - 1];

  let subtotal = 0, iva = 0, inc = 0, bolsas = 0, icui = 0, ibua = 0, icl = 0, adv = 0, icOtros = 0, discount = 0, surcharge = 0, total = 0;

  if (lastSubtotalMatch) {
    const tb = t.slice(lastSubtotalMatch.index!);
    subtotal  = parseCOP(lastSubtotalMatch[1]);
    // Font kerning can insert spaces mid-word ("De scue nt o") — use flexible patterns
    discount  = parseCOP(tb.match(/D\s*e\s*s\s*c\s*u\s*e\s*n\s*t\s*o\s+d\s*e\s+t\s*a\s*l\s*l\s*e\s*([\d.,]+)/i)?.[1]
                      ?? tb.match(/Descuento detalle([\d.,]+)/i)?.[1]);
    surcharge = parseCOP(tb.match(/R\s*e\s*c\s*a\s*r\s*g\s*o\s+d\s*e\s+t\s*a\s*l\s*l\s*e\s*([\d.,]+)/i)?.[1]
                      ?? tb.match(/Recargo detalle([\d.,]+)/i)?.[1]);
    if (isPOS) {
      // POS IVA uses US format (comma=thousands, dot=decimal): "28,739.50"
      const posIvaStr = tb.match(/Total\s+IVA\s*([\d,]+\.\d{2})/i)?.[1];
      if (posIvaStr) iva = parseFloat(posIvaStr.replace(/,/g, "")) || 0;
      // "Total documento" spans multiple lines in POS PDFs
      const posTotalStr = tb.match(/Total\s+documento[\s\S]{1,30}COP\s*\$([\d.,]+)/i)?.[1];
      if (posTotalStr) total = parseCOP(posTotalStr);
    } else {
      // "IVA" can appear as "IV A" due to font kerning — flexible spacing
      iva     = parseCOP(tb.match(/\nI\s*V\s*A\s*([\d.,]+)/i)?.[1]);
      inc     = parseCOP(tb.match(/\nINC\s*([\d.,]+)/i)?.[1]);
      bolsas  = parseCOP(tb.match(/Bolsas\s*([\d.,]+)/i)?.[1]);
      icui    = parseCOP(tb.match(/ICUI\s*([\d.,]+)/i)?.[1]);
      ibua    = parseCOP(tb.match(/IBUA\s*([\d.,]+)/i)?.[1]);
      icl     = parseCOP(tb.match(/\nICL\s*([\d.,]+)/i)?.[1]);
      adv     = parseCOP(tb.match(/\nADV\s*([\d.,]+)/i)?.[1]);
      // IC aparece bajo el rótulo "Otros impuestos" en el PDF de la Solución Gratuita DIAN
      icOtros = parseCOP(tb.match(/Otros impuestos\s*([\d.,]+)/i)?.[1]);
      total   = parseCOP(tb.match(/Total factura\s*\(=\)[^\n]*COP\s*\$([\d.,]+)/i)?.[1]);
    }
  }

  // ── Taxes array ───────────────────────────────────────────────────────────────
  const taxes: TaxDetail[] = [];
  if (iva > 0)    taxes.push({ taxId: "01", taxName: "IVA",    amount: iva,    percent: subtotal > 0 ? Math.round(iva / subtotal * 100) : 0 });
  if (inc > 0)    taxes.push({ taxId: "04", taxName: "INC",    amount: inc,    percent: 0 });
  if (bolsas > 0) taxes.push({ taxId: "22", taxName: "Bolsas", amount: bolsas, percent: 0 });
  if (icui > 0)    taxes.push({ taxId: "35", taxName: "ICUI",   amount: icui,    percent: 0 });
  if (ibua > 0)    taxes.push({ taxId: "54", taxName: "IBUA",   amount: ibua,    percent: 0 });
  if (icl > 0)     taxes.push({ taxId: "08", taxName: "ICL",    amount: icl,     percent: 0 });
  if (adv > 0)     taxes.push({ taxId: "10", taxName: "ADV",    amount: adv,     percent: 0 });
  if (icOtros > 0) taxes.push({ taxId: "03", taxName: "Otros Impuestos", amount: icOtros, percent: 0 });

  // ── Line items: extract "Detalles de Productos" section ─────────────────────
  // Section ends at "Notas Finales", "Valores informativos", or the second
  // "Datos Totales" block (whichever comes first after detallesStart).
  let lineItems: InvoiceLineItem[] = [];
  if (detallesStart >= 0) {
    // "Hoja X de Y" can appear immediately after the header (multi-page PDFs) or after the items —
    // it is NOT a reliable end-of-section marker. Only "Notas Finales" and "Datos Totales" mark the end.
    const sectionEnd = Math.min(
      ...[tl.indexOf("notas finales", detallesStart), tl.indexOf("datos totales", detallesStart)]
        .filter((i) => i > detallesStart)
    );
    const itemsSection = t.slice(
      detallesStart + "Detalles de Productos".length,
      sectionEnd < Infinity ? sectionEnd : undefined
    );
    lineItems = parseLineItems(itemsSection);
  }

  return {
    cufe,
    documentType,
    docNumber,
    trackId: cufe,
    zipFilename: filename,
    issueDate,
    issueDateISO,
    paymentMethod,
    issuerNit,
    issuerName,
    issuerCommercialName,
    issuerTaxpayerType,
    issuerFiscalRegime,
    issuerTaxResponsibility,
    issuerEconomicActivity,
    issuerCountry,
    issuerDepartment,
    issuerCity,
    issuerAddress,
    issuerPhone,
    issuerEmail,
    receiverNit,
    receiverName,
    receiverTaxpayerType,
    receiverFiscalRegime,
    receiverTaxResponsibility,
    receiverCountry,
    receiverDepartment,
    receiverCity,
    receiverAddress,
    receiverPhone,
    receiverEmail,
    subtotal,
    discount,
    surcharge,
    iva,
    total,
    taxes,
    concepts: "",
    lineItems,
    isDocumentoSoporte: documentType === "Documento soporte",
  };
}

// Aplica los impuestos del listado DIAN a una factura, reemplazando los extraídos del PDF.
function mergeDianListingTaxes(inv: InvoiceData, lt: DianTaxRow): void {
  const subtotal = inv.subtotal || 0;
  const taxes: TaxDetail[] = [];
  if (lt.iva > 0)            taxes.push({ taxId: "01", taxName: "IVA",             amount: lt.iva,            percent: subtotal > 0 ? Math.round(lt.iva / subtotal * 100) : 0 });
  if (lt.ica > 0)            taxes.push({ taxId: "07", taxName: "ICA",             amount: lt.ica,            percent: 0 });
  if (lt.ic > 0)             taxes.push({ taxId: "03", taxName: "IC",              amount: lt.ic,             percent: 0 });
  if (lt.inc > 0)            taxes.push({ taxId: "04", taxName: "INC",             amount: lt.inc,            percent: 0 });
  if (lt.bolsas > 0)         taxes.push({ taxId: "22", taxName: "Bolsas",          amount: lt.bolsas,         percent: 0 });
  if (lt.icui > 0)           taxes.push({ taxId: "35", taxName: "ICUI",            amount: lt.icui,           percent: 0 });
  if (lt.icl > 0)            taxes.push({ taxId: "08", taxName: "ICL",             amount: lt.icl,            percent: 0 });
  if (lt.ibua > 0)           taxes.push({ taxId: "54", taxName: "IBUA",            amount: lt.ibua,           percent: 0 });
  if (lt.inCarbono > 0)      taxes.push({ taxId: "09", taxName: "IN Carbono",      amount: lt.inCarbono,      percent: 0 });
  if (lt.inCombustibles > 0) taxes.push({ taxId: "10", taxName: "IN Combustibles", amount: lt.inCombustibles, percent: 0 });
  if (lt.icDatos > 0)        taxes.push({ taxId: "02", taxName: "IC Datos",        amount: lt.icDatos,        percent: 0 });
  if (lt.inpp > 0)           taxes.push({ taxId: "12", taxName: "INPP",            amount: lt.inpp,           percent: 0 });
  if (lt.timbre > 0)         taxes.push({ taxId: "11", taxName: "Timbre",          amount: lt.timbre,         percent: 0 });
  inv.taxes = taxes;
}

// POST /api/pdf-parse/to-excel
// Accepts multipart/form-data with field "pdfs" (multiple PDF files) and optional "listado" (DIAN listing Excel)
// Returns an XLSX file
router.post(
  "/to-excel",
  requireAuth,
  (req: Request, res: Response, next: Function) => {
    upload.fields([{ name: "pdfs", maxCount: 5000 }, { name: "listado", maxCount: 1 }])(req, res, (err: any) => {
      if (err) {
        console.error("[/api/pdf-parse/to-excel] Multer error:", err);
        return res.status(400).json({ error: `Error procesando archivos: ${err.message || err.code}` });
      }
      next();
    });
  },
  async (req: Request, res: Response) => {
  try {
  const fieldFiles = req.files as Record<string, Express.Multer.File[]> | undefined;
  const pdfFiles  = fieldFiles?.["pdfs"]    || [];
  const listadoFiles = fieldFiles?.["listado"] || [];

  if (!pdfFiles.length) {
    res.status(400).json({ error: "No se recibieron PDFs." });
    return;
  }

  // Parse DIAN listing if provided
  let listadoMap: Map<string, DianTaxRow> | null = null;
  if (listadoFiles.length) {
    try {
      listadoMap = await parseDianListing(listadoFiles[0].buffer);
    } catch (e) {
      // Non-fatal: proceed without listing
    }
  }

  const rawInvoices: InvoiceData[] = [];
  const errors: string[] = [];

  for (const file of pdfFiles) {
    try {
      const parsed = await pdfParse(file.buffer);
      const invoice = parsePdfText(parsed.text, file.originalname);
      // Replace text-flow line items with positional parse (preserves table structure)
      try {
        invoice.lineItems = await parseLineItemsPositional(file.buffer);
      } catch (e) {
        // Fallback to text-based line items on pdfjs error
        errors.push(`${file.originalname} (positional): ${(e as Error).message}`);
      }
      // Merge DIAN listing taxes if available
      if (listadoMap) {
        const lt = listadoMap.get(invoice.cufe);
        if (lt) mergeDianListingTaxes(invoice, lt);
      }
      rawInvoices.push(invoice);
    } catch (e) {
      errors.push(`${file.originalname}: ${(e as Error).message}`);
    }
  }

  // Deduplicate by CUFE: when a race condition caused a PDF to be saved under the
  // wrong filename, both the mis-named file and the correct file parse to the same
  // CUFE. Keep only the last occurrence (the correctly-named file arrives last).
  const seen = new Map<string, InvoiceData>();
  for (const inv of rawInvoices) {
    const key = inv.cufe || inv.docNumber || inv.zipFilename;
    seen.set(key, inv);
  }
  const invoices = [...seen.values()];

  if (!invoices.length) {
    res.status(422).json({ error: "No se pudo extraer datos de ningún PDF.", errors });
    return;
  }

  // Detect company NIT (most frequent receptor NIT)
  const nitCounts = new Map<string, number>();
  for (const inv of invoices) {
    if (inv.receiverNit) nitCounts.set(inv.receiverNit, (nitCounts.get(inv.receiverNit) || 0) + 1);
  }
  const companyNit = [...nitCounts.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] || "";
  const companyName = invoices.find((i) => i.receiverNit === companyNit)?.receiverName || "";

  // ── Optional Drive upload ────────────────────────────────────────────────────
  const driveConnectionId = typeof req.body?.driveConnectionId === "string" ? req.body.driveConnectionId.trim() : "";
  const includeDriveLinks = req.body?.includeDriveLinks === "true" || req.body?.includeDriveLinks === true;
  let driveUploaded = 0;

  if (driveConnectionId) {
    const userId = req.user!.userId;
    const driveConfig = await getUserGoogleDriveById(userId, driveConnectionId).catch(() => null);
    if (driveConfig) {
      const onTokenRefresh = async (newToken: string, expiry: number) => {
        await updateUserDriveTokens(userId, encryptToken(newToken), new Date(expiry).toISOString(), driveConnectionId);
      };

      for (const inv of invoices) {
        try {
          const lt = listadoMap?.get(inv.cufe);
          const isReceived = lt ? /recibida/i.test(lt.grupo) : true;
          const direction = isReceived ? "received" : "sent";

          // Own company: receptor for received, emisor for emitted
          const ownNit  = isReceived ? (lt?.nitReceptor  || inv.receiverNit || companyNit) : (lt?.nitEmisor  || inv.issuerNit || companyNit);
          const ownName = isReceived ? (lt?.nombreReceptor || inv.receiverName || companyName) : (lt?.nombreEmisor || inv.issuerName || companyName);

          // Third-party NIT goes in the filename for easy identification
          const thirdPartyNit = isReceived ? (lt?.nitEmisor || inv.issuerNit || "SinNIT") : (lt?.nitReceptor || inv.receiverNit || "SinNIT");
          const folio         = lt?.folio || inv.docNumber || inv.cufe.slice(0, 12);
          const safeStr       = (s: string) => s.replace(/[\\/:*?"<>|]/g, "").trim();
          const docNumber     = safeStr(`${thirdPartyNit} - ${folio}`);

          const issueDate = lt?.fechaEmision || inv.issueDate || new Date().toISOString().slice(0, 10);

          // Find the PDF buffer for this invoice (match by original filename or cufe)
          const pdfFile = pdfFiles.find((f) => {
            const stem = f.originalname.replace(/\.pdf$/i, "");
            return stem === inv.cufe || stem === inv.zipFilename || stem === inv.docNumber;
          });
          const pdfBuffer = pdfFile?.buffer ?? null;

          const result = await uploadInvoiceFilesToDrive(
            pdfBuffer, null, docNumber, ownNit, issueDate,
            driveConfig, userId, onTokenRefresh, direction, ownName
          );
          if (result.pdfUrl) {
            inv.driveUrl = result.pdfUrl;
            driveUploaded++;
          }
        } catch (e) {
          // Non-fatal: Drive upload failure doesn't block Excel generation
        }
      }
    }
  }

  const tmpPath = path.join(os.tmpdir(), `contago-pdf-excel-${Date.now()}.xlsx`);
  try {
    await generateExcelFile(invoices, tmpPath, includeDriveLinks && driveUploaded > 0, false, companyName, companyNit);
    const buf = fs.readFileSync(tmpPath);
    const dateStr = new Date().toISOString().slice(0, 10);
    const cleanName = (s: string) => s.replace(/[^a-zA-Z0-9À-ÿ\s.\-]/g, "").trim().slice(0, 60);
    const nitTxt  = companyNit  ? `${companyNit} - `  : "";
    const nomTxt  = companyName ? `${cleanName(companyName)} - ` : "";
    const xlsxName = `${nitTxt}${nomTxt}Facturas Recibidas DIAN ${dateStr}.xlsx`;
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${encodeURIComponent(xlsxName)}"`);
    if (errors.length) res.setHeader("X-Parse-Errors", errors.join(" | "));
    if (listadoMap) res.setHeader("X-Listado-Matched", String([...listadoMap.keys()].filter(k => rawInvoices.some(i => i.cufe === k)).length));
    if (driveConnectionId) res.setHeader("X-Drive-Uploaded", String(driveUploaded));
    res.send(buf);
  } finally {
    try { fs.unlinkSync(tmpPath); } catch {}
  }
  } catch (err) {
    console.error("[/api/pdf-parse/to-excel] Unhandled error:", err);
    if (!res.headersSent) {
      res.status(500).json({ error: `Error interno: ${(err as Error).message || err}` });
    }
  }
});

export default router;
