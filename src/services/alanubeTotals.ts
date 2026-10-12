/** Cálculo de ítems y totales para el payload de factura de Alanube (e-provider Alegra). */

export interface InvoiceTaxInput {
  /** Código de tributo DIAN (01 IVA, 04 INC, 03 ICA…). */
  taxCode: string;
  /** Tarifa en porcentaje (ej. 19, 5, 0). */
  percentage: number;
}

export interface InvoiceItemInput {
  description: string;
  price: number;
  quantity: number;
  /** Código de unidad DIAN (ej. "94" unidad). */
  unitCode: string;
  /** Descuento porcentual 0-100. */
  discountPct?: number;
  taxes?: InvoiceTaxInput[];
  /** Código estándar del bien/servicio (la DIAN notifica FAZ09 si falta). */
  standardCode?: { identificationId: string; id: "001" | "010" | "020" | "999" };
  code?: string;
}

export interface BuiltTax {
  taxCode: string;
  taxAmount: number;
  taxPercentage: string;
  taxableAmount: number;
}

export interface BuiltItem {
  description: string;
  price: number;
  quantity: number;
  unitCode: string;
  subtotal: number;
  taxAmount: number;
  total: number;
  taxes: BuiltTax[];
  discount?: number;
  discountAmount?: number;
  standardCode?: InvoiceItemInput["standardCode"];
  code?: string;
}

export interface BuiltTotals {
  grossTotal: number;
  taxableTotal: number;
  taxTotal: number;
  discountTotal: number;
  chargeTotal: number;
  advanceTotal: number;
  payableTotal: number;
  currencyCode: string;
}

/** Redondeo comercial a 2 decimales (half-up, tolerante al error binario). */
export function round2(n: number): number {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

export function buildItem(input: InvoiceItemInput): BuiltItem {
  if (!(input.quantity > 0)) throw new Error(`Cantidad inválida en "${input.description}"`);
  if (!(input.price >= 0)) throw new Error(`Precio inválido en "${input.description}"`);
  const pct = input.discountPct ?? 0;
  if (pct < 0 || pct > 100) throw new Error(`Descuento fuera de rango en "${input.description}"`);

  const gross = round2(input.quantity * input.price);
  const discountAmount = round2((gross * pct) / 100);
  const subtotal = round2(gross - discountAmount);

  const taxes: BuiltTax[] = (input.taxes ?? []).map((t) => ({
    taxCode: t.taxCode,
    taxAmount: round2((subtotal * t.percentage) / 100),
    taxPercentage: t.percentage.toFixed(2),
    taxableAmount: subtotal,
  }));
  const taxAmount = round2(taxes.reduce((s, t) => s + t.taxAmount, 0));

  const item: BuiltItem = {
    description: input.description,
    price: input.price,
    quantity: input.quantity,
    unitCode: input.unitCode,
    subtotal,
    taxAmount,
    total: round2(subtotal + taxAmount),
    taxes,
  };
  if (pct > 0) {
    item.discount = pct;
    item.discountAmount = discountAmount;
  }
  if (input.standardCode) item.standardCode = input.standardCode;
  if (input.code) item.code = input.code;
  return item;
}

/**
 * Totales del documento a partir de los ítems ya construidos.
 * Verificado contra el sandbox: el descuento de ítem ya está dentro de `subtotal` de cada línea, así que
 * `grossTotal` = suma de subtotales y `discountTotal` solo cuenta descuentos a nivel de documento (FAU02/FAU08/FAU14).
 * Las retenciones no restan del valor a pagar del documento electrónico.
 */
export function buildTotals(items: BuiltItem[], opts: { advanceTotal?: number; currencyCode?: string } = {}): BuiltTotals {
  const grossTotal = round2(items.reduce((s, i) => s + i.subtotal, 0));
  const taxTotal = round2(items.reduce((s, i) => s + i.taxAmount, 0));
  const discountTotal = 0;
  const advanceTotal = round2(opts.advanceTotal ?? 0);
  return {
    grossTotal,
    taxableTotal: grossTotal,
    taxTotal,
    discountTotal,
    chargeTotal: 0,
    advanceTotal,
    payableTotal: round2(grossTotal + taxTotal - discountTotal - advanceTotal),
    currencyCode: opts.currencyCode ?? "COP",
  };
}
