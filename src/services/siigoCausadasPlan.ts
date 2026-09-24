/**
 * Lógica PURA (sin base de datos ni red) para reconciliar el registro de facturas
 * DIAN de una empresa (`siigoIngestedCufes`) con las facturas de compra que ya
 * existen en Siigo. Vive aparte para poder probarla sin Mongo.
 *
 * Regla del producto: toda factura que ya esté causada en Siigo debe aparecer en
 * la pestaña "Causadas" de su mes, se haya causado desde ContaGO, desde otro
 * archivo/flujo o directamente en Siigo.
 */
import type { DianInvoiceRecord } from "./siigoIngestedCufesService.js";

/** Compra de Siigo reducida a lo que necesitamos. */
export interface SiigoPurchaseLite {
  id: string;
  name: string;      // "FC-1-123"
  nit: string;       // NIT del proveedor sin DV, solo dígitos
  prefix: string;    // prefijo de la factura del proveedor
  number: string;    // número de la factura del proveedor
  date: string;      // YYYY-MM-DD
  created: string;   // ISO de creación en Siigo
  total: number;
}

/** Prefijo de los registros creados a partir de una compra de Siigo sin CUFE conocido. */
export const SIIGO_CUFE_PREFIX = "siigo:";

const digits = (s: unknown): string => String(s ?? "").replace(/\D/g, "");
export const nitKey = (s: unknown): string => String(s ?? "").split("-")[0].replace(/\D/g, "");

/** Convierte un elemento de GET /v1/purchases en `SiigoPurchaseLite` (null si no trae id). */
export function normalizePurchase(p: any): SiigoPurchaseLite | null {
  const id = p?.id != null ? String(p.id) : "";
  if (!id) return null;
  const pi = p?.provider_invoice ?? {};
  return {
    id,
    name: String(p?.name ?? ""),
    nit: nitKey(p?.supplier?.identification),
    prefix: String(pi?.prefix ?? ""),
    number: String(pi?.number ?? p?.provider_invoice_number ?? ""),
    date: String(p?.date ?? "").slice(0, 10),
    created: String(p?.created ?? ""),
    total: Number(p?.total) || 0,
  };
}

/**
 * ¿Son el mismo número de factura? Mismos dígitos, o uno termina en el otro
 * (Siigo limita el número del proveedor a 11 dígitos, así que el folio completo
 * de la DIAN puede ser más largo). Mínimo 4 dígitos para no unir facturas distintas.
 */
export function sameInvoiceNumber(a: unknown, b: unknown): boolean {
  const x = digits(a), y = digits(b);
  if (!x || !y) return false;
  if (x === y) return true;
  const [s, l] = x.length <= y.length ? [x, y] : [y, x];
  return s.length >= 4 && l.endsWith(s);
}

const alnum = (s: unknown): string => String(s ?? "").replace(/[^A-Za-z0-9]/g, "").toUpperCase();
const letters = (s: unknown): string => String(s ?? "").replace(/[^A-Za-z]/g, "").toUpperCase();

/** 3 = mismo folio completo, 2 = mismo prefijo, 1 = solo coinciden los dígitos. */
function matchScore(docnum: string, prefix: string, number: string): number {
  if (alnum(docnum) === alnum(`${prefix}${number}`)) return 3;
  if (letters(docnum) === letters(prefix)) return 2;
  return 1;
}

/**
 * ¿Esta factura recién descargada de la DIAN ya está causada? (mismo NIT + mismo folio
 * contra el índice de causadas). Las notas crédito nunca se consideran duplicadas de una
 * factura aunque compartan dígitos. Sin prefijo en el registro se acepta por dígitos.
 */
export function matchesCaused(
  index: Array<{ nit: string; docnum: string }>,
  x: { supplierNit?: string; docNumberRaw?: string; providerInvoicePrefix?: string; providerInvoiceNumber?: string; isCreditNote?: boolean },
): boolean {
  if (x.isCreditNote) return false;
  const nit = nitKey(x.supplierNit);
  if (!nit) return false;
  const mine = [alnum(x.docNumberRaw), alnum(`${x.providerInvoicePrefix ?? ""}${x.providerInvoiceNumber ?? ""}`)].filter(Boolean);
  if (mine.length === 0) return false;
  return index.some((c) => {
    if (c.nit !== nit) return false;
    const theirs = alnum(c.docnum);
    if (mine.includes(theirs)) return true;
    // El registro de Siigo puede traer solo el número: entonces se compara por dígitos.
    return !letters(c.docnum) && sameInvoiceNumber(c.docnum, x.providerInvoiceNumber || x.docNumberRaw);
  });
}

export interface CausadasSyncPlan {
  /** Registros existentes que pasan a "caused" (o completan su siigoId). */
  updates: Array<{ cufe: string; set: Record<string, unknown> }>;
  /** Compras de Siigo sin registro: se crean como "caused" con cufe sintético `siigo:<id>`. */
  inserts: DianInvoiceRecord[];
  updated: number;
  inserted: number;
  /** Compras ya reflejadas (mismo siigoId o duplicado de una ya causada). */
  skipped: number;
}

/**
 * Cruza las compras de Siigo con el registro existente de la empresa.
 * - Si el registro (mismo NIT + mismo número) no está causado → se marca causado.
 * - Si no existe → se crea causado (cufe sintético), para que aparezca igual.
 * - Idempotente: correrlo dos veces con los mismos datos no cambia nada la segunda vez.
 */
export function planCausadasSync(
  companyId: string,
  existing: DianInvoiceRecord[],
  purchases: SiigoPurchaseLite[],
  nowIso: string,
): CausadasSyncPlan {
  const plan: CausadasSyncPlan = { updates: [], inserts: [], updated: 0, inserted: 0, skipped: 0 };

  const byCufe = new Set(existing.map((d) => d.cufe));
  const linkedSiigoIds = new Set(existing.map((d) => (d.siigoId ? String(d.siigoId) : "")).filter(Boolean));
  const byNit = new Map<string, DianInvoiceRecord[]>();
  const nameByNit = new Map<string, string>();
  for (const d of existing) {
    const k = nitKey(d.supplierNit);
    if (!k) continue;
    if (d.supplierName && !nameByNit.has(k)) nameByNit.set(k, d.supplierName);
    if (d.cufe.startsWith(SIIGO_CUFE_PREFIX)) continue; // los sintéticos no participan como "factura DIAN"
    const list = byNit.get(k) ?? [];
    list.push(d);
    byNit.set(k, list);
  }

  for (const p of purchases) {
    if (linkedSiigoIds.has(p.id) || byCufe.has(`${SIIGO_CUFE_PREFIX}${p.id}`)) { plan.skipped++; continue; }

    const candidates = p.nit && p.number
      ? (byNit.get(p.nit) ?? []).filter((d) => sameInvoiceNumber(d.docnum, p.number))
      : [];
    const target = candidates
      .filter((d) => d.status !== "caused" || !d.siigoId)
      .map((d) => ({ d, score: matchScore(d.docnum, p.prefix, p.number) }))
      .sort((a, b) => b.score - a.score)[0]?.d;
    if (target) {
      plan.updates.push({
        cufe: target.cufe,
        set: { status: "caused", siigoId: p.id, siigoName: p.name, causedAt: p.created || nowIso },
      });
      target.status = "caused";       // no reutilizar el mismo registro con otra compra
      target.siigoId = p.id;
      linkedSiigoIds.add(p.id);
      plan.updated++;
      continue;
    }
    if (candidates.length > 0) { plan.skipped++; continue; } // otra compra de Siigo para una factura ya causada

    if (!p.nit && !p.number && !p.name) { plan.skipped++; continue; }
    plan.inserts.push({
      companyId,
      cufe: `${SIIGO_CUFE_PREFIX}${p.id}`,
      docnum: `${p.prefix}${p.number}` || p.name,
      status: "caused",
      supplierNit: p.nit,
      supplierName: nameByNit.get(p.nit) || "",
      issueDate: p.date,
      total: p.total,
      ingestedAt: nowIso,
      fetchedAt: nowIso,
      causedAt: p.created || nowIso,
      siigoId: p.id,
      siigoName: p.name,
    });
    plan.inserted++;
  }
  return plan;
}
