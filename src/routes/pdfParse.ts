import { Router, type Request, type Response } from "express";
import multer from "multer";
import pdfParse from "pdf-parse";
import path from "path";
import fs from "fs";
import os from "os";
import { requireAuth } from "../middleware/auth.js";
import { generateExcelFile } from "../services/excelGenerator.js";
import type { InvoiceData, TaxDetail, InvoiceLineItem } from "../types/dianExcel.js";

const router = Router();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024, files: 600 },
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
const UM_CODES = "NIU|EA|UND|UNI|KGM|GLO|PK|PAR|C62|BX|PAC|SET|ACT|SRV|HUR|DAY|MON|ANO|LTR|MLT|GRM|MTR|CMT|HLT|TNE|DZN|PCK|CAR|BOX|ZZ|NAR|XUN|CCT|MMT|MGM|CGM";
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
  const t = collapseNumericSpaces(sectionText);
  const lines = t.split("\n").map((l) => l.trim()).filter(Boolean);
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
      } else if (hasUM || hasDbl) {
        // Completion/data line for the current multi-line item
        current += line;
        rows.push(current);
        current = null;
      } else if (pN) {
        // Pure digits = new Nro starting a new multi-line item
        rows.push(current);
        current = line;
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

  // ── CUFE from filename (our downloader saves <cufe>.pdf) ─────────────────────
  const cufe = filename.replace(/\.pdf$/i, "").trim() ||
    t.match(/Código Único de Factura[\s\S]{0,30}CUFE\s*:?\s*\n?\s*([a-f0-9A-F]{80,})/i)?.[1]?.trim() || "";

  // ── Document type ─────────────────────────────────────────────────────────────
  let documentType = "Factura electrónica";
  if (/NOTA\s*CR[ÉE]DITO/i.test(t)) documentType = "Nota Crédito";
  else if (/NOTA\s*D[ÉE]BITO/i.test(t)) documentType = "Nota Débito";
  else if (/DOCUMENTO\s*SOPORTE/i.test(t)) documentType = "Documento soporte";

  // ── Invoice number — runs directly into "Forma de pago:" on same line ─────────
  const docNumber = t.match(/Número de Factura:\s*([^\n]+?)(?:Forma de pago:|\n|$)/i)?.[1]?.trim() || "";

  // ── Issue date ────────────────────────────────────────────────────────────────
  const issueDate = t.match(/Fecha de Emisión:\s*(\d{2}\/\d{2}\/\d{4})/i)?.[1] || "";
  const issueDateISO = toISO(issueDate);

  // ── Payment method ────────────────────────────────────────────────────────────
  const paymentMethod = t.match(/Forma de pago:\s*([^\n]+?)(?:\s{2,}|\n|Fecha|$)/i)?.[1]?.trim() || "";

  // ── Emisor / Receptor sections ────────────────────────────────────────────────
  const emisorStart  = t.indexOf("Datos del Emisor");
  const receptorStart = t.indexOf("Datos del Adquiriente");
  const detallesStart = t.indexOf("Detalles de Productos");

  const emisorBlock   = emisorStart >= 0
    ? t.slice(emisorStart, receptorStart > emisorStart ? receptorStart : undefined)
    : "";
  const receptorBlock = receptorStart >= 0
    ? t.slice(receptorStart, detallesStart > receptorStart ? detallesStart : undefined)
    : "";

  // Razón Social can run into "Nombre Comercial:" on the same line
  const issuerName = (
    emisorBlock.match(/Razón Social:\s*([^\n]+?)(?:Nombre Comercial:|\n|$)/i)?.[1]?.trim() || ""
  );
  const issuerNit  = emisorBlock.match(/Nit del Emisor:\s*([\d-]+)/i)?.[1]?.trim() || "";

  const receiverName = receptorBlock.match(/Nombre o Razón Social:\s*([^\n]+)/i)?.[1]?.trim() || "";
  const receiverNit  = receptorBlock.match(/Número Documento:\s*([\d-]+)/i)?.[1]?.trim() || "";

  // ── Totals: the PDF renders two copies of the summary table; the second (after
  // "MONEDACOP") has the actual values.  Some fonts space it as "MO NEDA CO P".
  // Anchor on the last occurrence of "Subtotal<digit>" (template copy has "Subtotal\n").
  const subtotalMatches = [...t.matchAll(/Subtotal([\d.,]+)/g)];
  const lastSubtotalMatch = subtotalMatches[subtotalMatches.length - 1];

  let subtotal = 0, iva = 0, inc = 0, bolsas = 0, icui = 0, discount = 0, surcharge = 0, total = 0;

  if (lastSubtotalMatch) {
    const tb = t.slice(lastSubtotalMatch.index!);
    subtotal  = parseCOP(lastSubtotalMatch[1]);
    discount  = parseCOP(tb.match(/Descuento detalle([\d.,]+)/i)?.[1]);
    surcharge = parseCOP(tb.match(/Recargo detalle([\d.,]+)/i)?.[1]);
    // "IVA" can appear as "IV A" due to font kerning — use flexible spacing
    iva       = parseCOP(tb.match(/\nI\s*V\s*A\s*([\d.,]+)/i)?.[1]);
    inc       = parseCOP(tb.match(/\nINC([\d.,]+)/i)?.[1]);
    bolsas    = parseCOP(tb.match(/Bolsas([\d.,]+)/i)?.[1]);
    icui      = parseCOP(tb.match(/ICUI([\d.,]+)/i)?.[1]);
    total     = parseCOP(tb.match(/Total factura\s*\(=\)[^\n]*COP\s*\$([\d.,]+)/i)?.[1]);
  }

  // ── Taxes array ───────────────────────────────────────────────────────────────
  const taxes: TaxDetail[] = [];
  if (iva > 0)    taxes.push({ taxId: "01", taxName: "IVA",    amount: iva,    percent: subtotal > 0 ? Math.round(iva    / subtotal * 100) : 0 });
  if (inc > 0)    taxes.push({ taxId: "04", taxName: "INC",    amount: inc,    percent: 0 });
  if (bolsas > 0) taxes.push({ taxId: "22", taxName: "Bolsas", amount: bolsas, percent: 0 });
  if (icui > 0)   taxes.push({ taxId: "35", taxName: "ICUI",   amount: icui,   percent: 0 });

  // ── Line items: extract "Detalles de Productos" section ─────────────────────
  // Section ends at "Notas Finales", "Valores informativos", or the second
  // "Datos Totales" block (whichever comes first after detallesStart).
  let lineItems: InvoiceLineItem[] = [];
  if (detallesStart >= 0) {
    const sectionEnd = Math.min(
      ...[t.indexOf("Notas Finales", detallesStart), t.indexOf("Datos Totales", detallesStart), t.indexOf("Hoja 1 de", detallesStart)]
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
    receiverNit,
    receiverName,
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

// POST /api/pdf-parse/to-excel
// Accepts multipart/form-data with field "pdfs" (multiple PDF files)
// Returns an XLSX file
router.post("/to-excel", requireAuth, upload.array("pdfs", 600), async (req: Request, res: Response) => {
  const files = req.files as Express.Multer.File[] | undefined;
  if (!files?.length) {
    res.status(400).json({ error: "No se recibieron PDFs." });
    return;
  }

  const invoices: InvoiceData[] = [];
  const errors: string[] = [];

  for (const file of files) {
    try {
      const parsed = await pdfParse(file.buffer);
      invoices.push(parsePdfText(parsed.text, file.originalname));
    } catch (e) {
      errors.push(`${file.originalname}: ${(e as Error).message}`);
    }
  }

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

  const tmpPath = path.join(os.tmpdir(), `contago-pdf-excel-${Date.now()}.xlsx`);
  try {
    await generateExcelFile(invoices, tmpPath, false, false, companyName, companyNit);
    const buf = fs.readFileSync(tmpPath);
    const dateStr = new Date().toISOString().slice(0, 10);
    const xlsxName = `Facturas PDF DIAN ${dateStr}.xlsx`;
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${encodeURIComponent(xlsxName)}"`);
    if (errors.length) res.setHeader("X-Parse-Errors", errors.join(" | "));
    res.send(buf);
  } finally {
    try { fs.unlinkSync(tmpPath); } catch {}
  }
});

export default router;
