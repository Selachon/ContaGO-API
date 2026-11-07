/**
 * Sugerencias por proveedor a partir del balance de prueba por tercero (Excel de Siigo): cuenta de gasto con más
 * débito y, por cada retención, la cuenta con más crédito y su tarifa (leída del nombre de la cuenta).
 */
import { robustReadExcel } from "../siigoSuggestionsService.js";

export interface PerfilBalance {
  nit: string;
  nombre: string;
  cuentaGasto: string;
  /** Cuenta por pagar (22 o 23) con más crédito del tercero: define su forma de pago. */
  cuentaPorPagar: string;
  retefuente: { cuenta: string; tarifa: number } | null;
  reteica: { cuenta: string; tarifa: number } | null;
  reteiva: { cuenta: string; tarifa: number } | null;
}

const numero = (s: string) => {
  let t = String(s ?? "").replace(/[^\d,.\-]/g, "");
  if (!t) return 0;
  if (t.includes(",") && t.includes(".")) t = t.indexOf(",") < t.indexOf(".") ? t.replace(/,/g, "") : t.replace(/\./g, "").replace(",", ".");
  else if (t.includes(",")) t = t.replace(",", ".");
  const n = parseFloat(t);
  return Number.isFinite(n) ? n : 0;
};
const nitDe = (s: string) => String(s ?? "").split("-")[0].replace(/\D/g, "");

/** Tarifa que dice el nombre de la cuenta de retención ("… 2.5%", "… 11 X MIL"). */
function tarifaDe(nombre: string): number | null {
  const s = String(nombre);
  const p = s.match(/(\d+(?:[.,]\d+)?)\s*%/);
  if (p) return parseFloat(p[1].replace(",", "."));
  const nums = [...s.matchAll(/(\d+(?:[.,]\d+)?)/g)].map((x) => parseFloat(x[1].replace(",", "."))).filter((n) => n > 0 && n < 40);
  return nums.length ? nums[nums.length - 1] : null;
}

export async function perfilesDesdeBalance(buffer: Buffer): Promise<PerfilBalance[]> {
  const rows = await robustReadExcel(buffer);
  const h = rows.findIndex((r) => r.some((c) => String(c || "").toLowerCase().includes("transaccional")));
  if (h < 0) throw new Error("No se encontró el encabezado del balance: sube el «Balance de prueba por tercero» de Siigo.");
  const cab = rows[h].map((c) => String(c || "").trim().toLowerCase());
  const col = (f: (x: string) => boolean) => cab.findIndex(f);
  const c = {
    trans: col((x) => x.includes("transaccional")),
    codigo: col((x) => x.includes("código") || x.includes("codigo")),
    cuenta: col((x) => x.includes("nombre cuenta")),
    nit: col((x) => x.includes("identificaci")),
    tercero: col((x) => x.includes("nombre tercero")),
    deb: col((x) => x.includes("débito") || x.includes("debito")),
    cre: col((x) => x.includes("crédito") || x.includes("credito")),
  };
  if (c.codigo < 0 || c.nit < 0 || c.deb < 0 || c.cre < 0) throw new Error("El balance no trae las columnas de código, identificación, débito y crédito.");
  const cel = (r: string[], i: number) => (i >= 0 ? String(r[i] ?? "").trim() : "");
  type Acc = { codigo: string; nombre: string; deb: number; cre: number };
  const por = new Map<string, { nombre: string; gasto: Map<string, Acc>; rf: Acc[]; ri: Acc[]; rc: Acc[]; cxp: Map<string, Acc> }>();
  for (const r of rows.slice(h + 1)) {
    if (!r || !/^s[ií]$/i.test(cel(r, c.trans))) continue;
    const nit = nitDe(cel(r, c.nit));
    const codigo = cel(r, c.codigo);
    if (!nit || !codigo) continue;
    const acc: Acc = { codigo, nombre: cel(r, c.cuenta), deb: numero(cel(r, c.deb)), cre: numero(cel(r, c.cre)) };
    let p = por.get(nit);
    if (!p) por.set(nit, (p = { nombre: cel(r, c.tercero), gasto: new Map(), rf: [], ri: [], rc: [], cxp: new Map() }));
    if (/^[567]/.test(codigo)) {
      const e = p.gasto.get(codigo);
      if (e) e.deb += acc.deb; else p.gasto.set(codigo, acc);
    } else if (codigo.startsWith("2365")) p.rf.push(acc);
    else if (codigo.startsWith("2367")) p.ri.push(acc);
    else if (codigo.startsWith("2368")) p.rc.push(acc);
    else if (/^2[23]/.test(codigo) && !/^23(65|67|68)/.test(codigo)) {
      const e = p.cxp.get(codigo);
      if (e) e.cre += acc.cre; else p.cxp.set(codigo, acc);
    }
  }
  const ret = (l: Acc[]) => {
    const top = l.slice().sort((a, b) => b.cre - a.cre)[0];
    const tarifa = top ? tarifaDe(top.nombre) : null;
    return top && tarifa ? { cuenta: top.codigo, tarifa } : null;
  };
  return [...por.entries()].map(([nit, p]) => ({
    nit,
    nombre: p.nombre,
    cuentaGasto: [...p.gasto.values()].sort((a, b) => b.deb - a.deb)[0]?.codigo ?? "",
    cuentaPorPagar: [...p.cxp.values()].sort((a, b) => b.cre - a.cre)[0]?.codigo ?? "",
    retefuente: ret(p.rf),
    reteica: ret(p.rc),
    reteiva: ret(p.ri),
  }));
}
