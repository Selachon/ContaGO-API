/**
 * Parser del "Movimiento auxiliar por cuenta contable" de SIIGO (Excel).
 *
 * Es el libro auxiliar de una cuenta contable (típicamente el banco) en un
 * periodo. Para conciliación bancaria lo que importa por cada renglón es:
 *   - fecha, comprobante (RP-1-5104…), tercero/NIT, descripción
 *   - Débito / Crédito  →  en una cuenta de banco (activo, naturaleza débito):
 *        Débito  = entra plata  → "in"  (ingreso)
 *        Crédito = sale plata   → "out" (egreso)
 *
 * El archivo trae cabeceras de reporte (título, empresa, NIT, periodo), una
 * fila de encabezados de columna, una fila resumen con el SALDO INICIAL, los
 * movimientos, y una fila "Total general". Mapeamos columnas por nombre de
 * encabezado (no por posición) para tolerar variaciones de export.
 */
import ExcelJS from "exceljs";

export interface LedgerEntry {
  date: string; // YYYY-MM-DD (vacío si no se pudo parsear)
  voucher: string; // Comprobante, p.ej. "RP-1-5104"
  seq: string; // Secuencia
  nit: string; // Identificación del tercero
  thirdParty: string; // Nombre del tercero
  description: string;
  debit: number;
  credit: number;
  value: number; // = debit || credit (el que sea > 0)
  direction: "in" | "out"; // débito → in, crédito → out
}

export interface ParsedLedger {
  accountCode: string;
  accountName: string;
  period: string;
  opening: number | null;
  closing: number | null;
  totals: { debits: number; credits: number };
  entries: LedgerEntry[];
  check: { ok: boolean; expectedClosing: number | null; issues: string[] };
}

export class LedgerError extends Error {
  status: number;
  code: string;
  constructor(message: string, status = 422, code = "ledger_error") {
    super(message);
    this.name = "LedgerError";
    this.status = status;
    this.code = code;
  }
}

const TOL = 1; // tolerancia en pesos para el autochequeo de saldos

// Convierte cualquier celda (número, texto con separadores, etc.) a número.
const num = (v: unknown): number => {
  if (v == null || v === "") return 0;
  if (typeof v === "number") return v;
  const s = String(v).trim();
  // Formato colombiano puede venir como "1.234.567,89" o "1,234,567.89".
  // Heurística: si hay coma decimal (coma seguida de 1-2 dígitos al final), tratar punto como miles.
  let norm = s.replace(/[^\d.,-]/g, "");
  if (/,\d{1,2}$/.test(norm)) {
    norm = norm.replace(/\./g, "").replace(",", ".");
  } else {
    norm = norm.replace(/,/g, "");
  }
  const n = Number(norm);
  return Number.isFinite(n) ? n : 0;
};

const norm = (s: unknown): string =>
  String(s ?? "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/\s+/g, " ")
    .trim();

// Lee una celda como string limpio.
const str = (v: unknown): string => {
  if (v == null) return "";
  if (typeof v === "object") {
    // ExcelJS puede devolver { text } para rich text o { result } para fórmulas.
    const o = v as Record<string, unknown>;
    if (typeof o.text === "string") return o.text.trim();
    if (typeof o.result !== "undefined") return String(o.result).trim();
    if (o.richText && Array.isArray(o.richText)) {
      return (o.richText as { text: string }[]).map((r) => r.text).join("").trim();
    }
  }
  return String(v).trim();
};

// Convierte "22/04/2026" o un Date a "YYYY-MM-DD".
const toIsoDate = (v: unknown): string => {
  if (v instanceof Date) {
    return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, "0")}-${String(v.getDate()).padStart(2, "0")}`;
  }
  const s = str(v);
  const m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  const iso = s.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;
  return "";
};

interface ColMap {
  codigo?: number;
  cuenta?: number;
  comprobante?: number;
  secuencia?: number;
  fecha?: number;
  identificacion?: number;
  tercero?: number;
  descripcion?: number;
  detalle?: number;
  saldoInicial?: number;
  debito?: number;
  credito?: number;
}

// Localiza la fila de encabezados (la que contiene "débito" y "crédito") y mapea columnas.
function findHeader(rows: unknown[][]): { headerRow: number; map: ColMap } {
  for (let i = 0; i < rows.length; i++) {
    const cells = rows[i].map(norm);
    const hasDebito = cells.findIndex((c) => c === "debito" || c === "debitos");
    const hasCredito = cells.findIndex((c) => c === "credito" || c === "creditos");
    if (hasDebito >= 0 && hasCredito >= 0) {
      const find = (...names: string[]) =>
        cells.findIndex((c) => names.some((n) => c === n || c.includes(n)));
      const map: ColMap = {
        codigo: find("codigo contable", "codigo"),
        cuenta: find("cuenta contable", "cuenta"),
        comprobante: find("comprobante"),
        secuencia: find("secuencia"),
        fecha: find("fecha elaboracion", "fecha"),
        identificacion: find("identificacion", "nit"),
        tercero: find("nombre del tercero", "tercero", "nombre"),
        descripcion: find("descripcion"),
        detalle: find("detalle"),
        saldoInicial: find("saldo inicial"),
        debito: hasDebito,
        credito: hasCredito,
      };
      // En el formato "MOVIMIENTO CUENTAS - GENERAL" hay dos columnas DESCRIPCION:
      // la primera es el nombre de cuenta (col 2) y la segunda es la descripción
      // del movimiento (col 8, después de COMPROBANTE). Usamos la segunda.
      const allDescr = cells.reduce<number[]>((a, c, i) => {
        if (c === "descripcion") a.push(i);
        return a;
      }, []);
      if (allDescr.length > 1 && map.comprobante != null) {
        const afterComp = allDescr.find((i) => i > map.comprobante!);
        if (afterComp != null) map.descripcion = afterComp;
      }
      // Limpia índices -1.
      for (const k of Object.keys(map) as (keyof ColMap)[]) {
        if (map[k] != null && map[k]! < 0) delete map[k];
      }
      return { headerRow: i, map };
    }
  }
  throw new LedgerError(
    "No reconozco este Excel como un Movimiento auxiliar por cuenta contable de SIIGO (no encontré las columnas Débito/Crédito).",
    422,
    "auxiliar_unrecognized"
  );
}

/** Parsea el libro auxiliar de SIIGO y devuelve movimientos normalizados + cuadre interno. */
export async function parseAuxiliaryLedger(buffer: Buffer): Promise<ParsedLedger> {
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.load(buffer as unknown as ArrayBuffer);
  } catch {
    throw new LedgerError("No se pudo leer el Excel del auxiliar (¿archivo dañado?).", 422, "excel_invalido");
  }
  const ws = wb.worksheets[0];
  if (!ws) throw new LedgerError("El Excel del auxiliar está vacío.", 422, "excel_vacio");

  // Materializa la matriz de celdas (1-based de ExcelJS → 0-based).
  const rows: unknown[][] = [];
  ws.eachRow({ includeEmpty: true }, (row) => {
    const arr: unknown[] = [];
    row.eachCell({ includeEmpty: true }, (cell) => arr.push(cell.value));
    rows.push(arr);
  });

  // Metadatos de cabecera (antes de los encabezados de columna). El texto suele
  // venir repetido en todas las columnas por celdas combinadas → tomamos el
  // valor único de la primera celda de cada fila.
  const headLines = rows.slice(0, 12).map((r) => str(r.find((c) => str(c)) ?? ""));
  // "De: JUN 1/2026 A: JUN 30/2026" o "De JUN 1/2026 A JUN 30/2026"
  const period = headLines.find((l) => /^de[:\s]+\w+.*\d{4}/i.test(l)) || "";

  const { headerRow, map } = findHeader(rows);

  const get = (row: unknown[], col?: number): unknown => (col == null ? undefined : row[col]);

  let opening: number | null = null;
  let accountCode = "";
  let accountName = "";
  const entries: LedgerEntry[] = [];
  let totalDebits = 0;
  let totalCredits = 0;
  let closing: number | null = null;
  const issues: string[] = [];

  for (let i = headerRow + 1; i < rows.length; i++) {
    const row = rows[i];
    const codigoCell = str(get(row, map.codigo));
    const cuentaCell = str(get(row, map.cuenta));
    const comprobante = str(get(row, map.comprobante));
    const fecha = toIsoDate(get(row, map.fecha));

    // Fila resumen de apertura: "Cuenta contable: 11100503Bancolombia" con Saldo inicial.
    if (/cuenta contable:/i.test(codigoCell)) {
      const si = num(get(row, map.saldoInicial));
      if (si || opening == null) opening = si;
      const mc = codigoCell.match(/(\d{4,})/);
      if (mc) accountCode = mc[1];
      accountName = codigoCell.replace(/cuenta contable:\s*/i, "").replace(/^\d+/, "").trim();
      continue;
    }

    // Fila de totales / cierre.
    if (/total general|total$/i.test(codigoCell) || /total general/i.test(str(get(row, map.cuenta)))) {
      const d = num(get(row, map.debito));
      const c = num(get(row, map.credito));
      if (d) totalDebits = d;
      if (c) totalCredits = c;
      continue;
    }

    // Fila de "Procesado en:" u otras notas al pie.
    if (/procesado en:/i.test(codigoCell)) continue;

    // Movimiento real: requiere comprobante o fecha y algún valor.
    const debit = num(get(row, map.debito));
    const credit = num(get(row, map.credito));
    if (!comprobante && !fecha) continue;
    if (debit === 0 && credit === 0) continue;

    if (!accountCode && /^\d{4,}$/.test(codigoCell)) accountCode = codigoCell;
    if (!accountCode && cuentaCell) {
      // Formato "MOVIMIENTO CUENTAS - GENERAL": primera col = "11100505   BCO OCCIDENTE..."
      const mc = cuentaCell.match(/^(\d{5,})\s+(.+)/);
      if (mc) { accountCode = mc[1]; accountName = mc[2].trim(); }
    }
    if (!accountName && cuentaCell) accountName = cuentaCell;
    // En el formato "MOVIMIENTO CUENTAS - GENERAL" el saldo inicial se repite
    // en cada fila de movimiento — lo tomamos de la primera fila de datos.
    if (opening == null && map.saldoInicial != null) {
      const si = num(get(row, map.saldoInicial));
      if (si !== 0) opening = si;
    }

    const direction: "in" | "out" = debit > 0 ? "in" : "out";
    entries.push({
      date: fecha,
      voucher: comprobante,
      seq: str(get(row, map.secuencia)),
      nit: str(get(row, map.identificacion)),
      thirdParty: str(get(row, map.tercero)),
      description: str(get(row, map.descripcion)) || str(get(row, map.detalle)),
      debit,
      credit,
      value: debit > 0 ? debit : credit,
      direction,
    });
  }

  // Totales parseados (por si la fila "Total general" no apareció).
  const sumDebits = entries.reduce((a, e) => a + e.debit, 0);
  const sumCredits = entries.reduce((a, e) => a + e.credit, 0);
  if (!totalDebits) totalDebits = sumDebits;
  if (!totalCredits) totalCredits = sumCredits;

  // Cierre esperado por la ecuación contable.
  const expectedClosing = opening != null ? opening + sumDebits - sumCredits : null;
  closing = expectedClosing;

  // Autochequeos.
  if (entries.length === 0) issues.push("No se detectaron movimientos en el auxiliar.");
  if (Math.abs(totalDebits - sumDebits) > TOL) {
    issues.push(
      `Total débitos no cuadra: reporte $${totalDebits.toLocaleString()} vs. sumado $${sumDebits.toLocaleString()}.`
    );
  }
  if (Math.abs(totalCredits - sumCredits) > TOL) {
    issues.push(
      `Total créditos no cuadra: reporte $${totalCredits.toLocaleString()} vs. sumado $${sumCredits.toLocaleString()}.`
    );
  }

  return {
    accountCode,
    accountName,
    period,
    opening,
    closing,
    totals: { debits: sumDebits, credits: sumCredits },
    entries,
    check: { ok: issues.length === 0, expectedClosing, issues },
  };
}
