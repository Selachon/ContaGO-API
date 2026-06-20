/**
 * Parser de extractos bancarios en PDF (Bancolombia / Itaú), con contraseña.
 *
 * Decisión de diseño clave: el extracto es un sistema contable cerrado, así que
 * NO confiamos ciegamente en el parseo: se VERIFICA matemáticamente.
 *  - La dirección (ingreso/egreso) se decide por el delta del SALDO corriente
 *    (agnóstico al banco: Bancolombia trae valor firmado, Itaú columnas separadas).
 *  - Motor de cuadre: (a) continuidad del saldo (cada fila: |saldo−saldo_prev| == valor)
 *    y (b) totales de control declarados en el propio extracto (abonos/cargos).
 *  Si algo no cuadra, se reporta en `reconciliation.issues` y la UI lo bloquea.
 */
import { createRequire } from "module";

const require = createRequire(import.meta.url);
// pdf.js empaquetado por pdf-parse (soporta contraseña vía getDocument({password})).
// eslint-disable-next-line @typescript-eslint/no-var-requires
const pdfjsLib = require("pdf-parse/lib/pdf.js/v2.0.550/build/pdf.js");

export type MovKind = "ingreso" | "egreso" | "bank_fee";

export interface BankMovement {
  date: string;
  description: string;
  value: number;
  balance: number;
  kind: MovKind;
  direction: "in" | "out";
}

export interface Reconciliation {
  ok: boolean;
  balanceChainOk: boolean;
  totalsOk: boolean;
  openingBalance: number | null;
  closingBalance: number | null;
  declared: { credits: number | null; debits: number | null };
  parsed: { credits: number; debits: number; count: number };
  issues: string[];
}

export interface ParsedStatement {
  bank: string;
  movements: BankMovement[];
  reconciliation: Reconciliation;
}

export class StatementError extends Error {
  status: number;
  code: string;
  constructor(message: string, status = 422, code = "statement_error") {
    super(message);
    this.name = "StatementError";
    this.status = status;
    this.code = code;
  }
}

const TOL = 1; // tolerancia en pesos para el cuadre
const num = (s: unknown): number => Number(String(s ?? "").replace(/[^\d.-]/g, ""));

const FEE_RX =
  /4x1000|4 x 1000|gmf|gravamen|rendimientos financ|cobro transf|manejo portal|cobro serv disp|davipl|impto gobierno|impuesto trans|impuesto al valor agregado|\biva\b|comisi|comtransferencia|cuota (de )?(plan|manejo)|inter[eé]s|intereses|seguro de vida|^servicio\b/i;

// ─── Extracción de texto del PDF (líneas reconstruidas por coordenada Y) ──
async function extractLines(buffer: Buffer, password?: string): Promise<string[][]> {
  const data = new Uint8Array(buffer);
  let doc: any;
  try {
    doc = await pdfjsLib.getDocument({ data, password }).promise;
  } catch (e: any) {
    const name = String(e?.name || e?.message || "");
    if (/password/i.test(name)) {
      throw new StatementError(
        password ? "La contraseña del PDF es incorrecta." : "El PDF está protegido: ingresa la contraseña.",
        422,
        password ? "wrong_password" : "password_required"
      );
    }
    throw new StatementError("No se pudo leer el PDF (¿archivo dañado?).", 422, "pdf_invalido");
  }
  const pages: string[][] = [];
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    const content = await page.getTextContent();
    const rows = new Map<number, { x: number; s: string }[]>();
    for (const it of content.items as any[]) {
      const y = Math.round(it.transform[5]);
      if (!rows.has(y)) rows.set(y, []);
      rows.get(y)!.push({ x: it.transform[4], s: it.str });
    }
    const lines = [...rows.keys()]
      .sort((a, b) => b - a)
      .map((y) =>
        rows
          .get(y)!
          .sort((a, b) => a.x - b.x)
          .map((i) => i.s)
          .join(" ")
          .replace(/\s+/g, " ")
          .trim()
      )
      .filter(Boolean);
    pages.push(lines);
  }
  return pages;
}

function detectBank(pages: string[][]): string {
  const allText = pages.flat().join(" ").toLowerCase();
  const head = allText; // reutilizamos allText para los checks de encabezado también
  // Davivienda primero: extractos Davivienda siempre traen "davivienda.com" en el pie,
  // pero sus movimientos pueden mencionar "BANCOLOMBIA" en transferencias entrantes,
  // lo que dispararía una detección falsa si se verifica bancolombia antes.
  if (allText.includes("davivienda")) return "davivienda";
  const head40 = pages.flat().slice(0, 40).join(" ").toLowerCase();
  if (head40.includes("bancolombia")) return "bancolombia";
  if (head40.includes("itaú") || head40.includes("itau")) return "itau";
  // Occidente se detecta en el encabezado (no en el cuerpo) para no confundir con
  // statements de otro banco que solo mencionen "Banco de Occidente" en un movimiento.
  if (head40.includes("banco de occidente")) return "occidente";
  // Banco de Bogotá: el nombre va en el pie (URL bancodebogota.com), no en el
  // encabezado. Es marcador específico y los demás bancos ya se descartaron arriba.
  if (allText.includes("bancodebogota.com")) return "bancobogota";
  return "desconocido";
}

interface RawMov {
  date: string;
  description: string;
  value: number;
  balance: number;
}

// ─── Bancolombia: "D/MM DESCRIPCIÓN VALOR(±) SALDO" ───────────────────────
function parseBancolombia(pages: string[][]): { raws: RawMov[]; opening: number | null; declared: { credits: number | null; debits: number | null }; closing: number | null } {
  const all = pages.flat();
  let year = "";
  const hasta = all.find((l) => /HASTA:/i.test(l));
  if (hasta) { const m = hasta.match(/HASTA:\s*(\d{4})/); if (m) year = m[1]; }

  const findMoney = (rx: RegExp): number | null => {
    const line = all.find((l) => rx.test(l));
    if (!line) return null;
    const m = line.match(/([\d,]+\.\d{2})/);
    return m ? num(m[1]) : null;
  };
  const opening = findMoney(/SALDO ANTERIOR/i);
  const closing = findMoney(/SALDO ACTUAL/i);
  const credits = findMoney(/TOTAL ABONOS/i);
  const debits = findMoney(/TOTAL CARGOS/i);

  const raws: RawMov[] = [];
  const rowRx = /^(\d{1,2}\/\d{2})\s+(.+?)\s+(-?[\d,]+\.\d{2})\s+([\d,]+\.\d{2})$/;
  for (const line of all) {
    const m = line.match(rowRx);
    if (!m) continue;
    const [, dm, desc, valueStr, balanceStr] = m;
    const [d, mo] = dm.split("/");
    raws.push({
      date: `${year || new Date().getFullYear()}-${mo.padStart(2, "0")}-${d.padStart(2, "0")}`,
      description: desc.trim(),
      value: Math.abs(num(valueStr)),
      balance: num(balanceStr),
    });
  }
  return { raws, opening, declared: { credits, debits }, closing };
}

// ─── Itaú: columnas separadas; texto desordenado; fila a veces partida ────
function parseItau(pages: string[][]): { raws: RawMov[]; opening: number | null; declared: { credits: number | null; debits: number | null }; closing: number | null } {
  const all = pages.flat();
  let year = "", month = "";
  const per = all.find((l) => /\d{2}\/\d{2}\/\d{4}\s+AL\s+\d{2}\/\d{2}\/\d{4}/i.test(l));
  if (per) { const m = per.match(/(\d{2})\/(\d{2})\/(\d{4})\s+AL/i); if (m) { month = m[2]; year = m[3]; } }

  let opening: number | null = null;
  const sal = all.find((l) => /Saldo al \d{2}\/\d{2}\/\d{4}/i.test(l));
  if (sal) { const m = sal.match(/Saldo al \d{2}\/\d{2}\/\d{4}\s+([\d,]+\.\d{2})/i); if (m) opening = num(m[1]); }

  // Resumen "Cuenta corriente <n> <saldoIni> <créditos> <débitos> <saldoFin>"
  let credits: number | null = null, debits: number | null = null, closing: number | null = null;
  const resumen = all.find((l) => /Cuenta corriente\s+[\d-]+\s+[\d,]+\.\d{2}\s+[\d,]+\.\d{2}\s+[\d,]+\.\d{2}\s+[\d,]+\.\d{2}/i.test(l));
  if (resumen) {
    const nums = resumen.match(/[\d,]+\.\d{2}/g) || [];
    if (nums.length >= 4) { credits = num(nums[1]); debits = num(nums[2]); closing = num(nums[3]); }
  }

  const NOISE = /^(A\.|Itaú|Itau|Nit\.|0{3,}|_|Pasan|Tiro|Viene|Página|P á gina|Cuenta corriente|CASTRO|NIT|Movimientos|Descripci|Día|Oficina|Número|transacci|docum|\d{2}\/\d{2}\/\d{4}$)/i;
  const raws: RawMov[] = [];
  let pendingDesc = "";
  const push = (dd: string, desc: string, valueStr: string, balanceStr: string) => {
    raws.push({
      date: `${year || new Date().getFullYear()}-${(month || "01").padStart(2, "0")}-${dd.padStart(2, "0")}`,
      description: desc.replace(/\bCalle 73\b/i, "").replace(/\s+/g, " ").trim(),
      value: num(valueStr),
      balance: num(balanceStr),
    });
  };
  for (const raw of all) {
    const line = raw.trim();
    if (!line || NOISE.test(line)) continue;
    let m = line.match(/^(\d{2})\s+\d+\s+(.+?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})$/);
    if (m) { push(m[1], m[2], m[3], m[4]); pendingDesc = ""; continue; }
    m = line.match(/^(\d{2})\s+\d+\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})$/);
    if (m && pendingDesc) { push(m[1], pendingDesc, m[2], m[3]); pendingDesc = ""; continue; }
    if (!/^\d/.test(line) && /[A-Za-z]/.test(line)) pendingDesc = line.replace(/\bCalle 73\b/i, "").trim();
  }
  return { raws, opening, declared: { credits, debits }, closing };
}

// ─── Banco de Occidente: extracto CONSOLIDADO; solo la CUENTA CORRIENTE ────
// Tabla "DIA TRANSACCIÓN IDENT DEBITOS CREDITOS SALDO". Hay que acotar a la
// sección "Extracto - CUENTA CORRIENTE" (el PDF trae también tarjetas de crédito).
function parseOccidente(pages: string[][]): { raws: RawMov[]; opening: number | null; declared: { credits: number | null; debits: number | null }; closing: number | null } {
  const all = pages.flat();
  // Acotar a la sección de cuenta corriente (desde su título hasta la de tarjetas).
  const startIdx = all.findIndex((l) => /Extracto\s*-\s*CUENTA CORRIENTE/i.test(l));
  let endIdx = all.length;
  if (startIdx >= 0) {
    const rel = all.slice(startIdx + 1).findIndex((l) => /MASTERCARD|TARJETA No\.|CREDENCIAL/i.test(l));
    if (rel >= 0) endIdx = startIdx + 1 + rel;
  }
  const sec = startIdx >= 0 ? all.slice(startIdx, endIdx) : all;

  let year = "", month = "";
  const corte = sec.find((l) => /FECHA DE CORTE:\s*\d{2}\/\d{2}\/\d{4}/i.test(l));
  if (corte) { const m = corte.match(/FECHA DE CORTE:\s*\d{2}\/(\d{2})\/(\d{4})/i); if (m) { month = m[1]; year = m[2]; } }

  const findMoney = (rx: RegExp): number | null => {
    const line = sec.find((l) => rx.test(l));
    const m = line?.match(rx);
    return m && m[1] ? num(m[1]) : null;
  };
  const opening = findMoney(/SALDO ANTERIOR\s+([\d,]+\.\d{2})/i);
  const closing = findMoney(/SALDO ACTUAL\s+([\d,]+\.\d{2})/i);
  const credits = findMoney(/\bCREDITOS\s+([\d,]+\.\d{2})/i);
  const debits = findMoney(/\bDEBITOS\s+([\d,]+\.\d{2})/i);

  // Tabla: desde el encabezado "DIA TRANSACCI" hasta el texto legal/fin.
  const hdr = sec.findIndex((l) => /DIA\s+TRANSACCI/i.test(l));
  const tailRel = hdr >= 0 ? sec.slice(hdr + 1).findIndex((l) => /En caso de mora|Hoja \d+ de/i.test(l)) : -1;
  const tableLines = hdr >= 0 ? sec.slice(hdr + 1, tailRel >= 0 ? hdr + 1 + tailRel : undefined) : sec;

  const raws: RawMov[] = [];
  const rowRx = /^(\d{1,2})\s+(.+?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})$/;
  for (const line of tableLines) {
    const m = line.match(rowRx);
    if (!m) continue;
    const [, dd, descIdent, debStr, credStr, balStr] = m;
    const deb = num(debStr), cred = num(credStr);
    const value = cred > 0 ? cred : deb;
    // Quitar el identificador final (A700827 / 0000000) de la descripción.
    const description = descIdent.replace(/\s+(?:[A-Z]\d{4,}|0{4,})$/i, "").replace(/\s+/g, " ").trim();
    raws.push({
      date: `${year || new Date().getFullYear()}-${(month || "01").padStart(2, "0")}-${dd.padStart(2, "0")}`,
      description,
      value: Math.abs(value),
      balance: num(balStr),
    });
  }
  return { raws, opening, declared: { credits, debits }, closing };
}

// ─── Banco de Bogotá: "DD/MM CODTRANS DESC ... VALOR(±) SALDO" ─────────────
// Resumen separa Cargos/IVA/GMF/Retención/Intereses; los débitos reales son su suma.
function parseBancoBogota(pages: string[][]): { raws: RawMov[]; opening: number | null; declared: { credits: number | null; debits: number | null }; closing: number | null } {
  const all = pages.flat();
  // Año: del rango del periodo (o primer 20xx del encabezado).
  let year = "";
  const yhead = all.slice(0, 25).join(" ").match(/\b(20\d{2})\b/);
  if (yhead) year = yhead[1];

  const findMoney = (rx: RegExp): number | null => {
    const line = all.find((l) => rx.test(l));
    const m = line?.match(rx);
    return m && m[1] ? num(m[1]) : null;
  };
  const opening = findMoney(/Saldo Inicial:\s*(-?[\d,]+\.\d{2})/i);
  const closing = findMoney(/Saldo Final:\s*(-?[\d,]+\.\d{2})/i);
  const credits = findMoney(/Abonos:\s*(-?[\d,]+\.\d{2})/i);
  // Débitos declarados = Cargos + IVA + GMF + Retención + Intereses (vienen separados).
  const sumAbs = (rx: RegExp): number => {
    const line = all.find((l) => rx.test(l));
    const m = line?.match(rx);
    return m && m[1] ? Math.abs(num(m[1])) : 0;
  };
  const debits =
    sumAbs(/Cargos:\s*(-?[\d,]+\.\d{2})/i) +
    sumAbs(/Total IVA:\s*(-?[\d,]+\.\d{2})/i) +
    sumAbs(/GMF:\s*(-?[\d,]+\.\d{2})/i) +
    sumAbs(/Retencion:\s*(-?[\d,]+\.\d{2})/i) +
    sumAbs(/Total Intereses:\s*(-?[\d,]+\.\d{2})/i);

  // Filas de movimiento: DD/MM ... VALOR(±) SALDO. Las líneas de continuación
  // (descripción partida, "efectuada el...", "DAVIVIENDA el...") no casan.
  const raws: RawMov[] = [];
  const rowRx = /^(\d{2})\/(\d{2})\s+(.+?)\s+(-?[\d,]+\.\d{2})\s+([\d,]+\.\d{2})$/;
  for (const line of all) {
    const m = line.match(rowRx);
    if (!m) continue;
    const [, dd, mm, desc, valueStr, balanceStr] = m;
    raws.push({
      date: `${year || new Date().getFullYear()}-${mm}-${dd}`,
      description: desc.replace(/\s+(Bogota|Calle 104 Av 19|Gcia Oper Trans)\b/gi, " ").replace(/\s+\d{6,}\s*$/, "").replace(/\s+/g, " ").trim(),
      value: Math.abs(num(valueStr)),
      balance: num(balanceStr),
    });
  }
  return { raws, opening, declared: { credits, debits: debits || null }, closing };
}

// ─── Davivienda: cuenta de ahorros/corriente ──────────────────────────────
// Formato por fila: "DD MM $ MONTO +/- DOCNUM DESCRIPCION [refs...]"
// Sin columna de saldo corriente: se computa desde Saldo Anterior.
// Los totales de control vienen en el encabezado (Más Créditos / Menos Débitos).
function parseDavivienda(pages: string[][]): { raws: RawMov[]; opening: number | null; declared: { credits: number | null; debits: number | null }; closing: number | null } {
  const all = pages.flat();

  // Año desde "INFORME DEL MES: MARZO /2026"
  let year = "";
  const informe = all.find((l) => /INFORME DEL MES:/i.test(l));
  if (informe) { const m = informe.match(/\/(\d{4})/); if (m) year = m[1]; }
  if (!year) { const m = all.slice(0, 30).join(" ").match(/\b(20\d{2})\b/); if (m) year = m[1]; }

  // Totales del encabezado (cada label en su propia línea, valor en la siguiente).
  const findNextMoney = (rx: RegExp): number | null => {
    const idx = all.findIndex((l) => rx.test(l));
    if (idx < 0) return null;
    const next = all[idx + 1];
    const m = next?.match(/\$([\d,]+\.\d{2})/);
    return m ? num(m[1]) : null;
  };
  const opening = findNextMoney(/^Saldo Anterior$/i);
  const credits  = findNextMoney(/^Más Créditos$/i);
  const debits   = findNextMoney(/^Menos Débitos$/i);
  const closing  = findNextMoney(/^Nuevo Saldo$/i);

  // Fila de movimiento: "DD MM $ MONTO +/- DOCNUM DESCRIPCION [refs...]"
  const rowRx = /^(\d{2})\s+(\d{2})\s+\$\s*([\d,]+\.\d{2})\s+([+-])\s+\d{1,6}\s+(.+)/;
  let runBal = opening ?? 0;
  const raws: RawMov[] = [];

  for (const line of all) {
    const m = line.match(rowRx);
    if (!m) continue;
    const [, dd, mm, amountStr, sign, descRaw] = m;
    const value = num(amountStr);
    // Quitar números de referencia largos (NIT, orden de pago); conservar sufijo textual
    // (destinatario CIC069, AMELIA, etc.) que viene después del NIT 9016750463.
    // Deduplicar tokens consecutivos idénticos que el PDF a veces duplica.
    const description = descRaw
      .replace(/\s+\d{7,}/g, "")
      .replace(/\b(\w+)\s+\1\b/gi, "$1")
      .replace(/\s+/g, " ")
      .trim();
    if (sign === "+") runBal += value; else runBal -= value;
    raws.push({
      date: `${year || new Date().getFullYear()}-${mm}-${dd}`,
      description,
      value,
      balance: Math.round(runBal * 100) / 100,
    });
  }
  return { raws, opening, declared: { credits, debits }, closing };
}

// ─── Clasificación + motor de cuadre ─────────────────────────────────────
function classifyAndReconcile(
  raws: RawMov[],
  opening: number | null,
  declared: { credits: number | null; debits: number | null },
  closing: number | null
): { movements: BankMovement[]; reconciliation: Reconciliation } {
  const movements: BankMovement[] = [];
  const issues: string[] = [];
  let prev = opening;
  let credits = 0, debits = 0;
  let chainOk = true;

  raws.forEach((r, i) => {
    const delta = prev == null ? null : r.balance - prev;
    const direction: "in" | "out" = delta == null ? "out" : delta >= 0 ? "in" : "out";
    // Verificación de continuidad: el valor debe igualar |delta|.
    if (delta != null && Math.abs(Math.abs(delta) - r.value) > TOL) {
      chainOk = false;
      issues.push(
        `Fila ${i + 1} (${r.date} ${r.description.slice(0, 30)}): el valor ${r.value} no cuadra con el cambio de saldo ${Math.abs(delta).toFixed(2)}.`
      );
    }
    const kind: MovKind = FEE_RX.test(r.description) ? "bank_fee" : direction === "in" ? "ingreso" : "egreso";
    if (direction === "in") credits += r.value; else debits += r.value;
    movements.push({ date: r.date, description: r.description, value: r.value, balance: r.balance, kind, direction });
    prev = r.balance;
  });

  // Cuadre contra totales de control declarados.
  let totalsOk = true;
  if (declared.credits != null && Math.abs(declared.credits - credits) > TOL) {
    totalsOk = false;
    issues.push(`Total abonos no cuadra: extracto $${declared.credits.toLocaleString()} vs. parseado $${credits.toLocaleString()}.`);
  }
  if (declared.debits != null && Math.abs(declared.debits - debits) > TOL) {
    totalsOk = false;
    issues.push(`Total cargos no cuadra: extracto $${declared.debits.toLocaleString()} vs. parseado $${debits.toLocaleString()}.`);
  }
  // Cuadre del saldo final (si se conoce).
  if (closing != null && prev != null && Math.abs(closing - prev) > TOL) {
    issues.push(`Saldo final no cuadra: extracto $${closing.toLocaleString()} vs. calculado $${prev.toLocaleString()}.`);
    totalsOk = false;
  }
  if (raws.length === 0) issues.push("No se detectaron movimientos en el extracto.");

  return {
    movements,
    reconciliation: {
      ok: chainOk && totalsOk && raws.length > 0,
      balanceChainOk: chainOk,
      totalsOk,
      openingBalance: opening,
      closingBalance: prev,
      declared,
      parsed: { credits, debits, count: raws.length },
      issues,
    },
  };
}

/** Parsea un extracto bancario en PDF y devuelve movimientos + cuadre. */
export async function parseBankPdf(buffer: Buffer, password?: string): Promise<ParsedStatement> {
  const pages = await extractLines(buffer, password);
  const bank = detectBank(pages);
  if (bank === "desconocido") {
    throw new StatementError(
      "No reconozco el banco de este extracto (por ahora: Bancolombia, Itaú, Occidente, Banco de Bogotá, Davivienda). Sube el Excel de movimientos.",
      422,
      "bank_unsupported"
    );
  }
  const parsed =
    bank === "bancolombia" ? parseBancolombia(pages)
    : bank === "occidente" ? parseOccidente(pages)
    : bank === "bancobogota" ? parseBancoBogota(pages)
    : bank === "davivienda" ? parseDavivienda(pages)
    : parseItau(pages);
  const { movements, reconciliation } = classifyAndReconcile(parsed.raws, parsed.opening, parsed.declared, parsed.closing);
  return { bank, movements, reconciliation };
}
