/**
 * Dinero. Los cálculos financieros se hacen en centavos (enteros) y solo se
 * exponen como número decimal de 2 posiciones al persistir/mostrar.
 * Estrategia de redondeo: half-up sobre centavos (la usada por Siigo/DIAN).
 */

export function toCents(value: number | string | undefined | null): number {
  const n = typeof value === "string" ? parseAmount(value) : Number(value);
  if (!Number.isFinite(n)) return 0;
  return Math.round(n * 100);
}

/**
 * Interpreta un importe escrito por una persona. Admite formato colombiano
 * ("1.250.000,50") y anglosajón ("1,250,000.50"): manda el último separador,
 * salvo que sea un separador de miles (exactamente 3 dígitos detrás).
 */
export function parseAmount(raw: string): number {
  const cleaned = String(raw).replace(/[^\d.,-]/g, "").trim();
  if (!cleaned) return NaN;

  const lastDot = cleaned.lastIndexOf(".");
  const lastComma = cleaned.lastIndexOf(",");
  const decimalSep = lastDot > lastComma ? "." : lastComma > -1 ? "," : "";

  if (!decimalSep) return Number(cleaned);

  const decimals = cleaned.length - cleaned.lastIndexOf(decimalSep) - 1;
  // "1.250.000" o "1,250,000": el separador final agrupa miles, no decimales.
  const groupsOnly = decimals === 3 && (cleaned.match(new RegExp(`\\${decimalSep}`, "g")) || []).length >= 1
    && !(lastDot > -1 && lastComma > -1);

  if (groupsOnly) return Number(cleaned.replace(/[.,]/g, ""));

  const intPart = cleaned.slice(0, cleaned.lastIndexOf(decimalSep)).replace(/[.,]/g, "");
  const decPart = cleaned.slice(cleaned.lastIndexOf(decimalSep) + 1);
  return Number(`${intPart}.${decPart}`);
}

export function fromCents(cents: number): number {
  return Math.round(cents) / 100;
}

/** Redondeo a 2 decimales half-up, seguro frente a errores de coma flotante. */
export function round2(value: number): number {
  if (!Number.isFinite(value)) return 0;
  return fromCents(toCents(value));
}

export function sum(values: Array<number | undefined | null>): number {
  return fromCents(values.reduce<number>((acc, v) => acc + toCents(v ?? 0), 0));
}

export function subtract(a: number, b: number): number {
  return fromCents(toCents(a) - toCents(b));
}

export function multiply(value: number, factor: number): number {
  return fromCents(Math.round(toCents(value) * factor));
}

/** Porcentaje aplicado a una base (retenciones). */
export function percentage(base: number, rate: number): number {
  return fromCents(Math.round((toCents(base) * rate) / 100));
}

export function isZero(value: number): boolean {
  return toCents(value) === 0;
}

export function gte(a: number, b: number): boolean {
  return toCents(a) >= toCents(b);
}

/** Formato colombiano: $ 1.250.000 */
export function formatCOP(value: number): string {
  return new Intl.NumberFormat("es-CO", {
    style: "currency",
    currency: "COP",
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  }).format(value || 0);
}
