/**
 * Herencia del proyecto: al causar una factura se etiqueta con un proyecto (cajaErpTags, kind "invoice", por CUFE o
 * NIT|nº doc). Cuando esa factura se paga en Egresos, el movimiento bancario hereda ese proyecto en vez de pedirlo otra vez.
 *
 * El puente es el comprobante de compra que se creó en Siigo al causar (consecutivo + nombre guardados en
 * siigoIngestedCufes) frente al vencimiento que Siigo devuelve en cuentas por pagar (prefijo + consecutivo).
 */

export interface CausadaRef {
  cufe?: string;
  docnum?: string;
  supplierNit?: string;
  siigoName?: string;
  siigoNumber?: string;
}

export interface DueRef {
  prefix?: string;
  consecutive?: string | number;
}

export interface HerenciaProyecto {
  /** Proyecto único de las facturas cruzadas; null si ninguna tiene o hay varios distintos. */
  proyectoId: string | null;
  /** Hay facturas con proyectos distintos: no se hereda nada y el usuario elige. */
  conflicto: boolean;
  porFactura: Array<{ prefix: string; consecutive: string; proyectoId: string | null }>;
}

const digits = (s: unknown): string => String(s ?? "").replace(/\D/g, "");
const alnum = (s: unknown): string => String(s ?? "").replace(/[^A-Za-z0-9]/g, "").toUpperCase();

/** Mismo NIT con o sin dígito de verificación. */
function sameNit(a: unknown, b: unknown): boolean {
  const x = digits(a), y = digits(b);
  if (!x || !y) return false;
  if (x === y) return true;
  return (x.length >= 9 && x.slice(0, -1) === y) || (y.length >= 9 && y.slice(0, -1) === x);
}

const sameNumber = (a: unknown, b: unknown): boolean => {
  const x = Number(digits(a)), y = Number(digits(b));
  return x > 0 && y > 0 && x === y;
};

/** Claves con las que Causación pudo etiquetar la factura: CUFE, o NIT|prefijo+número (documento soporte). */
function tagKeys(c: CausadaRef): string[] {
  const keys: string[] = [];
  if (c.cufe) keys.push(String(c.cufe).trim());
  const nit = String(c.supplierNit ?? "").trim();
  if (nit && c.docnum) {
    const doc = String(c.docnum).trim();
    keys.push(`${nit}|${doc}`, `${digits(nit)}|${doc}`);
    if (digits(nit).length >= 9) keys.push(`${digits(nit).slice(0, -1)}|${doc}`);
  }
  return keys;
}

export function proyectoDeFacturas(
  causadas: CausadaRef[],
  tags: Record<string, string>,
  supplierNit: string,
  dues: DueRef[],
): HerenciaProyecto {
  const delProveedor = causadas.filter((c) => sameNit(c.supplierNit, supplierNit));
  const porFactura = dues.map((d) => {
    const prefix = String(d.prefix ?? "");
    const consecutive = String(d.consecutive ?? "");
    let cands = delProveedor.filter((c) => sameNumber(c.siigoNumber, consecutive));
    if (cands.length > 1 && prefix) {
      const p = alnum(prefix);
      const conPrefijo = cands.filter((c) => alnum(c.siigoName).includes(p));
      if (conPrefijo.length) cands = conPrefijo;
    }
    let proyectoId: string | null = null;
    for (const c of cands) {
      const hit = tagKeys(c).map((k) => tags[k]).find((v) => !!v);
      if (hit) { proyectoId = hit; break; }
    }
    return { prefix, consecutive, proyectoId };
  });
  const distintos = [...new Set(porFactura.map((f) => f.proyectoId).filter((p): p is string => !!p))];
  return { proyectoId: distintos.length === 1 ? distintos[0] : null, conflicto: distintos.length > 1, porFactura };
}
