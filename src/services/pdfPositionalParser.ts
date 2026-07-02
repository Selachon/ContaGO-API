/**
 * Positional line-item extractor for DIAN "Solución Gratuita" PDFs.
 *
 * Uses pdfjs-dist to get each text fragment with its (x, y) coordinates,
 * then reconstructs table rows by grouping items that share the same Y.
 *
 * Column X boundaries (595pt wide A4, DIAN fixed layout):
 *   Nro x<58 | Código ~55 | Descripción 75-168 | U/M ~168
 *   Cantidad ~200 | IVA 408-480 | INC 480-530 | LineExt 520+
 */

import type { InvoiceLineItem, TaxDetail } from "../types/dianExcel.js";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
let _pdfjs: any = null;
async function getPdfjs() {
  if (!_pdfjs) _pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  return _pdfjs;
}

// ── Column boundaries (DIAN fixed layout) ────────────────────────────────────
const X = {
  nroMax:   58,   // item number: x < 58
  codigoX:  54,   // Código starts at x ≈ 55
  descMin:  75,   // description text
  descMax: 168,
  umMin:   160,   // U/M column (anchor: row has U/M)
  qtyMin:  190, qtyMax: 240,
  ivaMin:  408, ivaMax: 480,  // IVA amount + % (may be merged)
  incMin:  480, incMax: 530,  // INC amount + %
  extMin:  520, extMax: 590,  // precio unitario de venta (line extension)
  rowTol:    5,   // Y tolerance to group items into same row (pt)
  descTol:  25,   // max Y distance for description overflow rows
};

type PosItem = { text: string; x: number; y: number };
type Row     = { y: number; items: PosItem[] };

// ── Number parsers ───────────────────────────────────────────────────────────
const COP_RE = /\d{1,3}(?:\.\d{3})*,\d{2}/g;  // Colombian: 1.234,56
const PCT_RE = /\d+\.\d{2}(?![\d,])/g;          // Percent:   19.00

function parseCOPs(s: string): number[] {
  return [...s.matchAll(COP_RE)].map((m) =>
    parseFloat(m[0].replace(/\./g, "").replace(",", "."))
  );
}
function parsePCTs(s: string): number[] {
  const withoutCOP = s.replace(new RegExp(COP_RE.source, "g"), " ");
  return [...withoutCOP.matchAll(PCT_RE)].map((m) => parseFloat(m[0]));
}

// ── pdfjs-dist: extract text items with (x, y) coordinates ──────────────────
export async function extractPositionedItems(buffer: Buffer): Promise<PosItem[]> {
  const pdfjs = await getPdfjs();
  const pdf = await pdfjs.getDocument({
    data: new Uint8Array(buffer),
    useWorkerFetch: false,
    isEvalSupported: false,
    useSystemFonts: true,
  }).promise;

  const items: PosItem[] = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page    = await pdf.getPage(p);
    const vp      = page.getViewport({ scale: 1 });
    const content = await page.getTextContent();
    for (const it of content.items) {
      if (!("str" in it) || !it.str.trim()) continue;
      items.push({
        text: it.str.trim(),
        x:   Math.round(it.transform[4]),
        // Flip Y (PDF origin is bottom-left) and add page offset
        y:   Math.round(vp.height - it.transform[5]) + (p - 1) * 2000,
      });
    }
  }
  return items.sort((a, b) => a.y - b.y || a.x - b.x);
}

// ── Row grouping ─────────────────────────────────────────────────────────────
function groupRows(items: PosItem[]): Row[] {
  const rows: Row[] = [];
  for (const it of items) {
    const row = rows.find((r) => Math.abs(r.y - it.y) <= X.rowTol);
    if (row) row.items.push(it);
    else     rows.push({ y: it.y, items: [it] });
  }
  rows.sort((a, b) => a.y - b.y);
  for (const r of rows) r.items.sort((a, b) => a.x - b.x);
  return rows;
}

function inRange(row: Row, xMin: number, xMax: number): string {
  return row.items
    .filter((it) => it.x >= xMin && it.x < xMax)
    .map((it) => it.text)
    .join(" ")
    .trim();
}

// A data row has a numeric item number at the left margin AND a U/M code
function isDataRow(row: Row): boolean {
  return (
    row.items.some((it) => it.x < X.nroMax && /^\d+$/.test(it.text)) &&
    row.items.some((it) => it.x >= X.umMin && it.x < 220)
  );
}

// ── Public: parse line items from a PDF buffer ────────────────────────────────
export async function parseLineItemsPositional(buffer: Buffer): Promise<InvoiceLineItem[]> {
  const posItems = await extractPositionedItems(buffer);

  // Restrict to the "Detalles de Productos" table area
  const detalles = posItems.find((it) =>
    it.text.toLowerCase().includes("detalles de productos")
  );
  const notas = posItems.find(
    (it) =>
      it.text.toLowerCase().includes("notas finales") ||
      it.text.toLowerCase().includes("datos totales")
  );
  const tableItems = detalles
    ? posItems.filter(
        (it) =>
          it.y > detalles.y + 5 && (!notas || it.y < notas.y)
      )
    : posItems;

  const rows = groupRows(tableItems);
  return parseRows(rows);
}

// ── Core row parser ───────────────────────────────────────────────────────────
function parseRows(rows: Row[]): InvoiceLineItem[] {
  // Step 1: find all data row indices
  const dataIdxs = rows.map((r, i) => (isDataRow(r) ? i : -1)).filter((i) => i >= 0);

  // Step 2: assign each non-data row to its NEAREST data row (by Y distance).
  // This prevents post-overflow lines of item N from being picked up as
  // pre-overflow of item N+1.
  const assigned = new Map<number, number>(); // rowIdx → dataRowIdx
  for (let i = 0; i < rows.length; i++) {
    if (isDataRow(rows[i])) continue;
    const t = inRange(rows[i], X.descMin, X.descMax);
    if (!t) continue;
    // Skip column-header rows (the row right after "Detalles de Productos" header)
    if (/descripci[oó]n/i.test(t)) continue;
    let nearest = -1, nearestDist = Infinity;
    for (const di of dataIdxs) {
      const d = Math.abs(rows[i].y - rows[di].y);
      if (d < nearestDist) { nearestDist = d; nearest = di; }
    }
    if (nearest >= 0 && nearestDist <= X.descTol) assigned.set(i, nearest);
  }

  const result: InvoiceLineItem[] = [];
  let lineNum = 1;

  for (const di of dataIdxs) {
    const row = rows[di];

    // Collect overflow rows assigned to this data row, split by before/after
    type DescFrag = { y: number; t: string };
    const beforeFrags: DescFrag[] = [];
    const afterFrags:  DescFrag[] = [];
    for (const [rowIdx, dataIdx] of assigned) {
      if (dataIdx !== di) continue;
      const t = inRange(rows[rowIdx], X.descMin, X.descMax);
      if (!t) continue;
      if (rows[rowIdx].y < row.y) beforeFrags.push({ y: rows[rowIdx].y, t });
      else                         afterFrags.push({ y: rows[rowIdx].y, t });
    }
    beforeFrags.sort((a, b) => a.y - b.y);
    afterFrags.sort((a, b)  => a.y - b.y);

    // In-row description (x 80-168, after código)
    let inRowDesc = inRange(row, 80, X.descMax);
    // Handle Código+Descripción merged in one text item at x≈55
    // e.g. "0010127 CERVEZA PILSEN EN LATA"
    if (!inRowDesc) {
      const codigoCell = inRange(row, X.codigoX, 80);
      const merged = codigoCell.match(/^\d{5,8}\s+(.+)$/);
      if (merged) inRowDesc = merged[1];
    }

    // Join fragments: preserve word boundaries but collapse digit-digit seams
    // (e.g. "LIGHT 3" + "30 BOTELLA" → "LIGHT 330 BOTELLA" — PDF split a number).
    const parts = [...beforeFrags.map((p) => p.t), inRowDesc, ...afterFrags.map((p) => p.t)]
      .filter(Boolean);
    let description = "";
    for (const part of parts) {
      if (!description) { description = part; continue; }
      const digitSeam = /\d$/.test(description) && /^\d/.test(part);
      description = digitSeam
        ? description + part
        : (description.trimEnd() + " " + part.trimStart()).trim();
    }
    description = description.trim();

    if (!description) continue;

    // ── Numeric columns ───────────────────────────────────────────────────
    const qtyText = inRange(row, X.qtyMin,  X.qtyMax);
    const ivaText = inRange(row, X.ivaMin,  X.ivaMax);
    const incText = inRange(row, X.incMin,  X.incMax);
    const extText = inRange(row, X.extMin,  X.extMax);

    const qtyCOPs = parseCOPs(qtyText);
    const ivaCOPs = parseCOPs(ivaText);
    const ivaPCTs = parsePCTs(ivaText);
    const incCOPs = parseCOPs(incText);
    const incPCTs = parsePCTs(incText);
    const extCOPs = parseCOPs(extText);

    const cantidad       = qtyCOPs[0] || 0;
    if (cantidad <= 0) continue;

    const ivaAmount      = ivaCOPs[0] || 0;
    const ivaPercent     = ivaPCTs[0] || 0;  // e.g. 19.00 (not 0.19), matches excelGenerator convention
    const incAmount      = incCOPs[0] || 0;
    const incPercent     = incPCTs[0] || 0;
    const totalUnitPrice = extCOPs.at(-1) || 0;  // precio unitario de venta = line extension
    const unitPrice      = cantidad > 0 ? totalUnitPrice / cantidad : totalUnitPrice;

    const taxes: TaxDetail[] = [];
    if (ivaPercent > 0 || ivaAmount > 0) {
      taxes.push({ taxId: "01", taxName: "IVA", amount: ivaAmount, percent: ivaPercent });
    }
    if (incPercent > 0 || incAmount > 0) {
      taxes.push({ taxId: "04", taxName: "INC", amount: incAmount, percent: incPercent });
    }

    result.push({
      lineNumber:    lineNum++,
      description,
      quantity:      cantidad,
      unitPrice,
      discount:      0,
      surcharge:     0,
      taxes,
      ivaAmount,
      ivaPercent,
      incAmount,
      incPercent,
      totalUnitPrice,
      taxableBase:   totalUnitPrice,
    });
  }

  return result;
}
