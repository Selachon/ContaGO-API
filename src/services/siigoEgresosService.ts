/**
 * Causación de Egresos / Recibos de pago en Siigo a partir de un Excel de
 * movimientos bancarios.
 *
 * Dos responsabilidades:
 *  1. Leer el Excel del extracto bancario y devolver sus hojas como tablas
 *     (columnas + filas), para que el frontend mapee qué columna es fecha,
 *     valor, descripción y (opcional) NIT del beneficiario.
 *  2. Validar y enviar a Siigo el payload del recibo de pago/egreso (RP)
 *     construido por el frontend, vía POST /v1/payment-receipts.
 *
 * El banco no está fijado: el parser es genérico (detección de encabezado +
 * mapeo de columnas en la UI), no asume un formato concreto.
 */
import ExcelJS from "exceljs";
import { createPaymentReceipt, getAccountsPayable, listVouchers, listPaymentReceipts, listJournals } from "./siigoService.js";

export interface BankSheet {
  name: string;
  columns: string[];
  rows: Record<string, unknown>[];
}

export interface ParsedBankFile {
  sheets: BankSheet[];
}

export class EgresosError extends Error {
  status: number;
  code: string;
  constructor(message: string, status = 400, code = "egresos_error") {
    super(message);
    this.name = "EgresosError";
    this.status = status;
    this.code = code;
  }
}

// ─── Lectura del Excel bancario ──────────────────────────────────────────

/** Extrae el valor primitivo de una celda exceljs (fórmulas, rich text, fechas). */
function cellPrimitive(value: ExcelJS.CellValue): unknown {
  if (value === null || value === undefined) return null;
  if (value instanceof Date) return value;
  if (typeof value === "object") {
    const obj = value as unknown as Record<string, unknown>;
    if ("result" in obj) return cellPrimitive(obj.result as ExcelJS.CellValue);
    if ("richText" in obj && Array.isArray(obj.richText)) {
      return (obj.richText as Array<{ text?: string }>).map((t) => t.text ?? "").join("");
    }
    if ("text" in obj) return obj.text;
    if ("error" in obj) return null;
    return null;
  }
  return value;
}

/** Convierte una celda Date a yyyy-MM-dd; deja el resto como está. */
function normalizeCell(value: unknown): unknown {
  if (value instanceof Date) {
    const y = value.getUTCFullYear();
    const m = String(value.getUTCMonth() + 1).padStart(2, "0");
    const d = String(value.getUTCDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  return value === undefined ? null : value;
}

function readRawRows(ws: ExcelJS.Worksheet): unknown[][] {
  const rows: unknown[][] = [];
  const nCols = ws.columnCount;
  for (let r = 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const cells: unknown[] = [];
    for (let c = 1; c <= nCols; c++) {
      cells.push(cellPrimitive(row.getCell(c).value));
    }
    rows.push(cells);
  }
  return rows;
}

/**
 * Detecta la fila de encabezado: dentro de las primeras 25 filas, la que tenga
 * más celdas de texto no vacías (los extractos suelen traer filas de relleno
 * con totales/títulos antes del encabezado real).
 */
function detectHeaderRow(rows: unknown[][]): number {
  const limit = Math.min(25, rows.length);
  let bestIdx = 0;
  let bestScore = -1;
  for (let i = 0; i < limit; i++) {
    const score = rows[i].filter((v) => typeof v === "string" && v.trim() !== "").length;
    if (score > bestScore) {
      bestScore = score;
      bestIdx = i;
    }
  }
  return bestScore <= 0 ? 0 : bestIdx;
}

function buildSheet(name: string, rows: unknown[][]): BankSheet {
  if (rows.length === 0) return { name, columns: [], rows: [] };

  const headerIdx = detectHeaderRow(rows);
  const header = rows[headerIdx] ?? [];
  const columns: string[] = [];
  const indices: number[] = [];
  header.forEach((h, idx) => {
    let label = h === null || h === undefined ? "" : String(h).trim();
    if (label === "") return;
    // Evita columnas duplicadas (mismo título): sufija con su índice.
    if (columns.includes(label)) label = `${label} (${idx + 1})`;
    columns.push(label);
    indices.push(idx);
  });

  const dataRows: Record<string, unknown>[] = [];
  for (let r = headerIdx + 1; r < rows.length; r++) {
    const raw = rows[r];
    // Salta filas totalmente vacías.
    if (!raw.some((v) => v !== null && v !== undefined && String(v).trim() !== "")) continue;
    const obj: Record<string, unknown> = {};
    columns.forEach((col, k) => {
      obj[col] = normalizeCell(raw[indices[k]]);
    });
    dataRows.push(obj);
  }

  return { name, columns, rows: dataRows };
}

/** Lee el Excel del extracto y devuelve todas las hojas como tablas. */
export async function parseBankExcel(buffer: Buffer): Promise<ParsedBankFile> {
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.load(buffer as unknown as ExcelJS.Buffer);
  } catch {
    throw new EgresosError(
      "No se pudo leer el archivo Excel (formato inválido o dañado).",
      422,
      "excel_invalido"
    );
  }

  const sheets: BankSheet[] = [];
  wb.eachSheet((ws) => {
    sheets.push(buildSheet(ws.name, readRawRows(ws)));
  });

  if (sheets.length === 0 || sheets.every((s) => s.rows.length === 0)) {
    throw new EgresosError("El archivo no contiene movimientos.", 422, "excel_vacio");
  }

  return { sheets };
}

// ─── Construcción/validación del payload del recibo de pago/egreso ────────

export type EgresoType = "DebtPayment" | "AdvancePayment" | "Detailed";

export interface EgresoDue {
  prefix: string;
  consecutive: number;
  quote: number;
  date?: string;
}

export interface EgresoAccount {
  code: string;
  movement: "Debit" | "Credit";
}

export interface EgresoItem {
  due?: EgresoDue;
  account?: EgresoAccount;
  /** Impuesto(s) de la cuenta (método avanzado): algunas cuentas en Siigo lo exigen. */
  taxes?: Array<{ id: number }>;
  description?: string;
  value: number;
}

export interface EgresoCurrency {
  code: string;
  exchange_rate: number;
}

export interface EgresoPayload {
  document: { id: number };
  type: EgresoType;
  date: string;
  supplier: { identification: string; branch_office?: number };
  currency?: EgresoCurrency;
  items?: EgresoItem[];
  payment?: { id: number; value: number };
  cost_center?: number;
  observations?: string;
}

const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;

/**
 * Valida la forma del payload antes de enviarlo a Siigo. Devuelve la lista de
 * errores (vacía = válido). No replica todas las reglas de producción de Siigo
 * (esas las verifica la API), solo lo que evita POSTs obviamente inválidos.
 */
export function validateEgresoPayload(p: Partial<EgresoPayload> | undefined): string[] {
  const errors: string[] = [];
  if (!p || typeof p !== "object") {
    return ["Payload vacío."];
  }

  if (!p.document?.id || !Number.isFinite(Number(p.document.id))) {
    errors.push("Falta el tipo de comprobante (document.id).");
  }
  if (!p.type || !["DebtPayment", "AdvancePayment", "Detailed"].includes(p.type)) {
    errors.push("Tipo de recibo inválido (type).");
  }
  if (!p.date || !ISO_DATE.test(String(p.date))) {
    errors.push("Fecha inválida (debe ser AAAA-MM-DD).");
  }
  if (!p.supplier?.identification || String(p.supplier.identification).trim() === "") {
    errors.push("Falta la identificación del beneficiario (supplier.identification).");
  }
  // El medio de pago aplica a DebtPayment/AdvancePayment; el Avanzado (Detailed)
  // arma el banco como una cuenta más, así que NO lleva bloque payment.
  if (p.type !== "Detailed") {
    if (!p.payment || !p.payment.id || !Number.isFinite(Number(p.payment.id))) {
      errors.push("Falta el medio de pago (payment.id).");
    }
    if (!p.payment || !Number.isFinite(Number(p.payment.value)) || Number(p.payment.value) <= 0) {
      errors.push("El valor del pago debe ser un número mayor a 0 (payment.value).");
    }
  }

  if (p.type === "DebtPayment") {
    if (!Array.isArray(p.items) || p.items.length === 0) {
      errors.push("Un abono a deuda requiere al menos una factura cruzada (items).");
    } else {
      p.items.forEach((it, i) => {
        if (!it.due?.prefix) errors.push(`Item ${i + 1}: falta el prefijo de la factura (due.prefix).`);
        if (!Number.isFinite(Number(it.due?.consecutive)))
          errors.push(`Item ${i + 1}: falta el consecutivo de la factura (due.consecutive).`);
        if (!Number.isFinite(Number(it.value)) || Number(it.value) <= 0)
          errors.push(`Item ${i + 1}: valor inválido.`);
      });

      // El neto de las facturas cruzadas debe coincidir con el valor del pago.
      const itemsTotal = p.items.reduce((s, it) => s + (Number(it.value) || 0), 0);
      const payTotal = Number(p.payment?.value) || 0;
      if (Math.abs(itemsTotal - payTotal) > 1) {
        errors.push(
          `La suma de las facturas cruzadas (${itemsTotal}) no coincide con el valor del pago (${payTotal}).`
        );
      }
    }
  }

  // Avanzado: partidas contables manuales; débitos deben igualar créditos.
  if (p.type === "Detailed") {
    if (!Array.isArray(p.items) || p.items.length === 0) {
      errors.push("El recibo avanzado requiere al menos una partida contable (items).");
    } else {
      let deb = 0, cred = 0;
      p.items.forEach((it, i) => {
        const code = it.account?.code;
        const mov = it.account?.movement;
        if (!code || String(code).trim() === "") errors.push(`Partida ${i + 1}: falta la cuenta contable (account.code).`);
        if (mov !== "Debit" && mov !== "Credit") errors.push(`Partida ${i + 1}: el movimiento debe ser Debit o Credit.`);
        if (!Number.isFinite(Number(it.value)) || Number(it.value) <= 0) errors.push(`Partida ${i + 1}: valor inválido.`);
        if (mov === "Debit") deb += Number(it.value) || 0;
        if (mov === "Credit") cred += Number(it.value) || 0;
      });
      if (errors.length === 0 && Math.abs(deb - cred) > 1) {
        errors.push(`El comprobante no está balanceado: débitos ${deb} ≠ créditos ${cred}.`);
      }
    }
  }

  return errors;
}

/** Redondeo a 2 decimales (evita artefactos de punto flotante que Siigo rechaza). */
const round2 = (n: unknown): number => Math.round((Number(n) || 0) * 100) / 100;

/** Envía el recibo de pago/egreso a Siigo. POST a producción. */
export async function submitEgreso(payload: EgresoPayload): Promise<unknown> {
  // Red de seguridad: redondea todos los valores a 2 decimales antes del POST.
  const clean: EgresoPayload = {
    ...payload,
    items: payload.items?.map((it) => ({ ...it, value: round2(it.value) })),
    payment: payload.payment ? { ...payload.payment, value: round2(payload.payment.value) } : undefined,
  };
  return createPaymentReceipt(clean);
}

// ─── Sugerencia de tercero por valor (a partir de cuentas por pagar) ──────

export interface SuggestedDue {
  prefix: string;
  consecutive: string | number;
  quote: number;
  date: string;
  balance: number;
}

export interface SupplierSuggestion {
  provider: { identification: string; name: string; branch_office: number };
  /** "due": una factura con ese saldo exacto. "providerTotal": el total del proveedor coincide. */
  matchType: "due" | "providerTotal";
  dues: SuggestedDue[];
  total: number;
}

/**
 * Busca en TODO el reporte de cuentas por pagar (sin filtrar proveedor) los
 * vencimientos cuyo saldo coincide con `value` (±tolerance), y los proveedores
 * cuyo saldo total coincide. Sirve para sugerir de quién es un pago a partir del
 * monto del movimiento bancario.
 */
export async function suggestSuppliersByValue(value: number, tolerance = 1): Promise<SupplierSuggestion[]> {
  if (!Number.isFinite(value) || value <= 0) return [];
  const raw = (await getAccountsPayable({})) as unknown;
  const rows: Array<Record<string, any>> = Array.isArray(raw)
    ? (raw as Array<Record<string, any>>)
    : Array.isArray((raw as any)?.results)
      ? (raw as any).results
      : [];

  const byProvider = new Map<string, { provider: SupplierSuggestion["provider"]; dues: SuggestedDue[] }>();
  for (const r of rows) {
    const id = String(r?.provider?.identification ?? "");
    if (!id) continue;
    const branch = Number(r?.provider?.branch_office) || 0;
    const key = `${id}-${branch}`;
    if (!byProvider.has(key)) {
      byProvider.set(key, {
        provider: { identification: id, name: String(r?.provider?.name ?? ""), branch_office: branch },
        dues: [],
      });
    }
    byProvider.get(key)!.dues.push({
      prefix: r?.due?.prefix ?? "",
      consecutive: r?.due?.consecutive ?? "",
      quote: Number(r?.due?.quote) || 1,
      date: r?.due?.date ?? "",
      balance: Number(r?.due?.balance) || 0,
    });
  }

  const out: SupplierSuggestion[] = [];
  for (const { provider, dues } of byProvider.values()) {
    const matchedDues = dues.filter((d) => Math.abs(d.balance - value) <= tolerance);
    if (matchedDues.length > 0) {
      out.push({ provider, matchType: "due", dues: matchedDues, total: matchedDues.reduce((s, d) => s + d.balance, 0) });
    }
    if (dues.length > 1) {
      const sum = dues.reduce((s, d) => s + d.balance, 0);
      if (Math.abs(sum - value) <= tolerance) {
        out.push({ provider, matchType: "providerTotal", dues, total: sum });
      }
    }
  }
  // Prioriza coincidencia por factura individual.
  out.sort((a, b) => (a.matchType === b.matchType ? 0 : a.matchType === "due" ? -1 : 1));
  return out.slice(0, 25);
}

// ─── Conciliación de ingresos: ¿ya existe el Recibo de Caja en Siigo? ─────

export interface VoucherMatch {
  id: string;
  name: string;
  number: string | number;
  date: string;
  value: number;
  customer: string;
}

function voucherValue(v: any): number {
  if (v?.payment && Number.isFinite(Number(v.payment.value))) return Number(v.payment.value);
  if (Array.isArray(v?.payments)) return v.payments.reduce((s: number, p: any) => s + (Number(p.value) || 0), 0);
  return 0;
}

/**
 * Busca Recibos de Caja (RC) ya creados en Siigo cuyo valor coincide con `value`
 * (±tolerance), dentro de una ventana de fechas de creación. Sirve para conciliar
 * un ingreso del extracto contra el RC que el contador ya causó en Siigo.
 */
export async function findMatchingVouchers(
  value: number,
  tolerance = 1,
  createdStart?: string,
  createdEnd?: string
): Promise<VoucherMatch[]> {
  if (!Number.isFinite(value) || value <= 0) return [];
  const matches: VoucherMatch[] = [];
  for (let page = 1; page <= 5; page++) {
    const query: Record<string, unknown> = { page, page_size: 100 };
    if (createdStart) query.created_start = createdStart;
    if (createdEnd) query.created_end = createdEnd;
    const raw = (await listVouchers(query)) as any;
    const rows: any[] = Array.isArray(raw?.results) ? raw.results : Array.isArray(raw) ? raw : [];
    if (rows.length === 0) break;
    for (const v of rows) {
      const val = voucherValue(v);
      if (Math.abs(val - value) <= tolerance) {
        matches.push({
          id: String(v?.id ?? ""),
          name: String(v?.name ?? ""),
          number: v?.number ?? "",
          date: String(v?.date ?? ""),
          value: val,
          customer: String(v?.customer?.identification ?? v?.customer?.id ?? ""),
        });
      }
    }
    if (rows.length < 100) break;
  }
  return matches.slice(0, 25);
}

// ─── Anti-duplicado contra Siigo: ¿ya existe un egreso (RP o CC) por ese valor? ──

export interface SiigoEgresoMatch {
  kind: "RP" | "CC";
  id: string;
  name: string;
  date: string;
  value: number;
  thirdParty: string;
}

const within = (a: string, b: string, days: number): boolean => {
  if (!a || !b) return true; // si falta alguna fecha, no descartar por fecha
  const da = Date.parse(a), db = Date.parse(b);
  if (Number.isNaN(da) || Number.isNaN(db)) return true;
  return Math.abs(da - db) <= days * 86400000;
};

/**
 * Busca en Siigo si ya existe un egreso por `value` (±tolerance), tanto como
 * Recibo de pago/egreso (RP) como Comprobante Contable (CC). Sirve para avisar
 * antes de causar y evitar duplicados de pagos hechos por fuera de la herramienta.
 */
export async function findExistingEgresoInSiigo(
  value: number,
  date: string,
  tolerance = 1,
  createdStart?: string,
  createdEnd?: string,
  dateWindowDays = 8
): Promise<SiigoEgresoMatch[]> {
  if (!Number.isFinite(value) || value <= 0) return [];
  const out: SiigoEgresoMatch[] = [];
  const q = (page: number): Record<string, unknown> => {
    const query: Record<string, unknown> = { page, page_size: 100 };
    if (createdStart) query.created_start = createdStart;
    if (createdEnd) query.created_end = createdEnd;
    return query;
  };

  // 1) Recibos de pago/egreso (RP) — match por payment.value
  try {
    for (let page = 1; page <= 5; page++) {
      const raw = (await listPaymentReceipts(q(page))) as any;
      const rows: any[] = Array.isArray(raw?.results) ? raw.results : Array.isArray(raw) ? raw : [];
      if (rows.length === 0) break;
      for (const r of rows) {
        const val = voucherValue(r);
        if (Math.abs(val - value) <= tolerance && within(r?.date, date, dateWindowDays)) {
          out.push({ kind: "RP", id: String(r?.id ?? ""), name: String(r?.name ?? ""), date: String(r?.date ?? ""), value: val, thirdParty: String(r?.supplier?.identification ?? "") });
        }
      }
      if (rows.length < 100) break;
    }
  } catch { /* la API puede limitar; seguimos con CC */ }

  // 2) Comprobantes contables (CC) — match si alguna partida tiene ese valor
  try {
    for (let page = 1; page <= 5; page++) {
      const raw = (await listJournals(q(page))) as any;
      const rows: any[] = Array.isArray(raw?.results) ? raw.results : Array.isArray(raw) ? raw : [];
      if (rows.length === 0) break;
      for (const j of rows) {
        const items: any[] = Array.isArray(j?.items) ? j.items : [];
        const hit = items.some((it) => Math.abs((Number(it?.value) || 0) - value) <= tolerance);
        if (hit && within(j?.date, date, dateWindowDays)) {
          out.push({ kind: "CC", id: String(j?.id ?? ""), name: String(j?.name ?? ""), date: String(j?.date ?? ""), value, thirdParty: "" });
        }
      }
      if (rows.length < 100) break;
    }
  } catch { /* journals puede no soportar created_start; ignoramos */ }

  return out.slice(0, 25);
}

/**
 * Versión por lote: trae los RP y CC de Siigo UNA sola vez (paginado) y empareja
 * todos los movimientos pasados, en vez de una consulta por fila. Devuelve un mapa
 * movementId → coincidencias.
 */
export async function findExistingForValues(
  items: Array<{ id: string; value: number; date: string }>,
  tolerance = 1,
  createdStart?: string,
  createdEnd?: string,
  dateWindowDays = 8
): Promise<Record<string, SiigoEgresoMatch[]>> {
  const q = (page: number): Record<string, unknown> => {
    const query: Record<string, unknown> = { page, page_size: 100 };
    if (createdStart) query.created_start = createdStart;
    if (createdEnd) query.created_end = createdEnd;
    return query;
  };

  // Descarga RP y CC una sola vez.
  const rps: any[] = [];
  try {
    for (let page = 1; page <= 8; page++) {
      const raw = (await listPaymentReceipts(q(page))) as any;
      const rows: any[] = Array.isArray(raw?.results) ? raw.results : Array.isArray(raw) ? raw : [];
      if (rows.length === 0) break;
      rps.push(...rows);
      if (rows.length < 100) break;
    }
  } catch { /* sigue con CC */ }

  const ccs: any[] = [];
  try {
    for (let page = 1; page <= 8; page++) {
      const raw = (await listJournals(q(page))) as any;
      const rows: any[] = Array.isArray(raw?.results) ? raw.results : Array.isArray(raw) ? raw : [];
      if (rows.length === 0) break;
      ccs.push(...rows);
      if (rows.length < 100) break;
    }
  } catch { /* journals puede no soportar created_start */ }

  const result: Record<string, SiigoEgresoMatch[]> = {};
  for (const it of items) {
    if (!Number.isFinite(it.value) || it.value <= 0) continue;
    const matches: SiigoEgresoMatch[] = [];
    for (const r of rps) {
      const val = voucherValue(r);
      if (Math.abs(val - it.value) <= tolerance && within(r?.date, it.date, dateWindowDays)) {
        matches.push({ kind: "RP", id: String(r?.id ?? ""), name: String(r?.name ?? ""), date: String(r?.date ?? ""), value: val, thirdParty: String(r?.supplier?.identification ?? "") });
      }
    }
    for (const j of ccs) {
      const jitems: any[] = Array.isArray(j?.items) ? j.items : [];
      if (jitems.some((x) => Math.abs((Number(x?.value) || 0) - it.value) <= tolerance) && within(j?.date, it.date, dateWindowDays)) {
        matches.push({ kind: "CC", id: String(j?.id ?? ""), name: String(j?.name ?? ""), date: String(j?.date ?? ""), value: it.value, thirdParty: "" });
      }
    }
    if (matches.length) result[it.id] = matches.slice(0, 5);
  }
  return result;
}
