/**
 * Funciones de normalización del motor de causación.
 *
 * Portadas 1:1 desde los scripts Python (script_compras.py / script_ventas.py /
 * script_terceros.py). Se mantiene el comportamiento exacto para garantizar
 * paridad con las salidas históricas.
 */

/**
 * Normaliza texto: minúsculas, sin tildes, sin caracteres de combinación,
 * espacios colapsados. Equivale a `normalizar_texto` de los scripts.
 */
export function normalizarTexto(valor: unknown): string {
  let t = String(valor ?? "").trim().toLowerCase();
  // Algunos Excel de SIIGO traen el carácter de reemplazo U+FFFD donde iba "ú".
  t = t.replace(/�/g, "u");
  // Descompone (NFKD) y elimina los diacríticos combinantes (U+0300–U+036F).
  t = t.normalize("NFKD").replace(/[̀-ͯ]/g, "");
  return t.split(/\s+/).filter(Boolean).join(" ");
}

/**
 * Convierte un valor de celda a número. Cualquier valor no numérico → 0.
 * Equivale a `pd.to_numeric(col, errors="coerce").fillna(0)`.
 */
export function toNumero(valor: unknown): number {
  if (valor === null || valor === undefined || valor === "") return 0;
  if (typeof valor === "number") return Number.isFinite(valor) ? valor : 0;
  if (typeof valor === "boolean") return valor ? 1 : 0;
  if (valor instanceof Date) return 0;
  const n = Number(String(valor).trim());
  return Number.isFinite(n) ? n : 0;
}

/**
 * Redondea a entero con redondeo bancario (round-half-to-even), igual que
 * `int(round(x, 0))` de Python. NO equivale a `Math.round` (que redondea .5
 * siempre hacia arriba).
 */
export function redondear(valor: unknown): number {
  const n = Number(valor);
  if (!Number.isFinite(n)) return 0;
  const piso = Math.floor(n);
  const resto = n - piso;
  if (resto < 0.5) return piso;
  if (resto > 0.5) return piso + 1;
  // Empate exacto en .5 → al par más cercano.
  return piso % 2 === 0 ? piso : piso + 1;
}

/** Redondea a 4 decimales (para comparar tarifas sin ruido de coma flotante). */
export function redondear4(valor: number): number {
  return Math.round(valor * 10000) / 10000;
}

/**
 * Normaliza una tarifa porcentual: limpia "%" y comas, y si viene como
 * fracción (0 < t ≤ 1) la lleva a porcentaje. Equivale a
 * `normalizar_tarifa_porcentaje` de script_compras.py.
 */
export function normalizarTarifaPorcentaje(valor: unknown): number {
  if (valor === null || valor === undefined) return 0;
  const limpio = String(valor).trim().replace(/%/g, "").replace(/,/g, ".");
  const n = Number(limpio);
  if (!Number.isFinite(n)) return 0;
  if (n > 0 && n <= 1) return n * 100;
  return n;
}

/** Limpia un NIT: descarta "nan", recorta el sufijo ".0" de lecturas float. */
export function limpiarNit(valor: unknown): string {
  if (valor === null || valor === undefined) return "";
  let t = String(valor).trim();
  if (t.toLowerCase() === "nan") return "";
  if (t.endsWith(".0")) t = t.slice(0, -2);
  return t;
}

/**
 * Limpia un código de cuenta contable: si trae " - Nombre", conserva solo el
 * código; descarta "nan" y el sufijo ".0".
 */
export function limpiarCuenta(valor: unknown): string {
  if (valor === null || valor === undefined) return "";
  let t = String(valor).trim();
  if (t.toLowerCase() === "nan") return "";
  if (t.includes("-")) t = t.split("-")[0].trim();
  if (t.endsWith(".0")) t = t.slice(0, -2);
  return t;
}

/** Abrevia el tipo de documento DIAN para las descripciones contables. */
export function abreviarTipoDocumento(tipoDocumento: unknown): string {
  const t = String(tipoDocumento ?? "").trim().toUpperCase();
  if (t.includes("DOCUMENTO EQUIVALENTE POS")) return "DEPOS";
  if (t.includes("FACTURA ELECTRÓNICA") || t.includes("FACTURA ELECTRONICA")) return "FE";
  if (t.includes("NOTA DE CRÉDITO") || t.includes("NOTA DE CREDITO")) return "NC";
  return t.slice(0, 10);
}

/** Indica si un tipo de documento es una nota crédito (criterio de compras). */
export function esNotaCredito(tipoDocumento: unknown): boolean {
  const t = normalizarTexto(tipoDocumento);
  return t.includes("nota de credito") || t.includes("nota credito") || t.startsWith("nc");
}

/**
 * Indica si un tipo de documento es una nota crédito (criterio de ventas).
 * A diferencia de compras, NO considera el prefijo "nc".
 */
export function esNotaCreditoVentas(tipoDocumento: unknown): boolean {
  const t = normalizarTexto(tipoDocumento);
  return t.includes("nota de credito") || t.includes("nota credito");
}

/** Indica si un tipo de documento es un documento soporte. */
export function esDocumentoSoporte(tipoDocumento: unknown): boolean {
  return normalizarTexto(tipoDocumento).includes("documento soporte");
}
