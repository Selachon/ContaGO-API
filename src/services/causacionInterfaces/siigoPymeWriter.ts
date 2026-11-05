/**
 * Interfaz de Siigo Pyme: "Modelo para la importación de movimiento contable - modelo general" (.xlsx, 129 columnas).
 * Toma las mismas líneas contables que la de Siigo; los datos propios del modelo (IVA de la línea, bases, proveedor,
 * cruce, vencimiento…) viajan en columnas auxiliares de cada línea.
 */
import * as XLSX from "xlsx";
import { COLUMNAS_PYME } from "./siigoPymePlantilla.js";
import { toNumero } from "./normalize.js";
import type { FilaSiigo } from "./types.js";

export interface PymeConfig {
  empresa: string;
  vendedor: number;
}

const MESES = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"];
const MARCA_PROVEEDOR = "'001";
const CRUCE = "P-001";

const texto = (v: unknown) => (v === null || v === undefined ? "" : String(v));
const pad = (v: string, ancho: number) => v.slice(0, ancho).padEnd(ancho, " ");
const num = (v: unknown) => {
  const n = toNumero(v);
  return Number.isFinite(n) ? n : 0;
};
const fechaPartes = (ddmmyyyy: string): [number, number, number] => {
  const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(ddmmyyyy || "");
  return m ? [Number(m[3]), Number(m[2]), Number(m[1])] : [0, 0, 0];
};

/** "P-1" → letra y código del comprobante; null si no tiene esa forma. */
export function partirComprobantePyme(tipo: string): { letra: string; codigo: number } | null {
  const m = /^([A-Za-z]+)[-\s]*0*(\d+)$/.exec(String(tipo ?? "").trim());
  return m ? { letra: m[1].toUpperCase(), codigo: Number(m[2]) } : null;
}

/** Cuenta de 10 dígitos (el modelo de Siigo Pyme rellena con ceros a la derecha). */
export function cuentaPyme(cuenta: unknown): number {
  const d = texto(cuenta).replace(/\D/g, "");
  if (!d || d.length > 10) throw new Error(`La cuenta «${texto(cuenta)}» no cabe en el modelo de Siigo Pyme (máximo 10 dígitos).`);
  return Number(d.padEnd(10, "0"));
}

/** "100/1" o "100-1" → centro y subcentro. */
function centro(v: unknown): [number, number] {
  const m = /^\s*(\d+)(?:\D+(\d+))?\s*$/.exec(texto(v));
  return m ? [Number(m[1]), Number(m[2] ?? 0)] : [0, 0];
}

export function escribirInterfazSiigoPyme(filas: FilaSiigo[], cfg: PymeConfig, ahora: Date = new Date()): Buffer {
  const ncol = COLUMNAS_PYME.length;
  const vacia = (): (string | number)[] => COLUMNAS_PYME.map(([, t, w]) => (t === "n" ? 0 : " ".repeat(w)));
  const secuencias = new Map<string, number>();
  let desde = "";
  let hasta = "";
  const llave = (d: [number, number, number]) => `${String(d[0]).padStart(4, "0")}${String(d[1]).padStart(2, "0")}${String(d[2]).padStart(2, "0")}`;
  const aaaammdd = ahora.getFullYear() * 10000 + (ahora.getMonth() + 1) * 100 + ahora.getDate();
  const hhmmss = ahora.getHours() * 10000 + ahora.getMinutes() * 100 + ahora.getSeconds();

  const datos = filas.map((l) => {
    const c = partirComprobantePyme(texto(l["Tipo de comprobante"]));
    if (!c) throw new Error(`El comprobante «${texto(l["Tipo de comprobante"])}» no sirve para Siigo Pyme: usa la letra y el código, por ejemplo P-1.`);
    const numero = num(l["Consecutivo comprobante"]);
    const clave = `${c.letra}-${c.codigo}-${numero}`;
    const sec = (secuencias.get(clave) ?? 0) + 1;
    secuencias.set(clave, sec);
    const [anio, mes, dia] = fechaPartes(texto(l["Fecha de elaboración"]));
    const k = llave([anio, mes, dia]);
    if (!desde || k < desde) desde = k;
    if (!hasta || k > hasta) hasta = k;

    const f = vacia();
    const deb = num(l["Débito"]);
    const cre = num(l["Crédito"]);
    const rol = texto(l["Rol"]);
    const [cc, sub] = centro(l["Código centro/subcentro de costos"]);
    const baseRet = l["Base retención"] === undefined || l["Base retención"] === "" ? null : Math.round(num(l["Base retención"]));
    const [anioP, mesP, diaP] = fechaPartes(texto(l["Fecha doc proveedor"]));
    const docProv = Math.trunc(num(l["Doc proveedor"]));

    f[0] = c.letra;
    f[1] = c.codigo;
    f[2] = numero;
    f[3] = cuentaPyme(l["Código cuenta contable"]);
    f[4] = deb !== 0 ? "D" : "C";
    f[5] = Math.abs(deb !== 0 ? deb : cre);
    f[6] = anio;
    f[7] = mes;
    f[8] = dia;
    f[9] = cfg.vendedor;
    f[12] = sec;
    f[13] = cc;
    f[14] = sub;
    f[15] = Math.trunc(num(l["Identificación tercero"]));
    f[17] = pad(texto(l["Descripción"]), 50);
    f[19] = "N";
    f[37] = aaaammdd;
    f[38] = hhmmss;
    if (baseRet !== null) { f[18] = baseRet; f[68] = baseRet; }
    if (rol === "base") {
      f[66] = num(l["IVA %"]);
      f[67] = num(l["IVA valor"]);
      f[70] = texto(l["Gravada"]) || " ";
    }
    if (l["Base ReteIVA"] !== undefined && l["Base ReteIVA"] !== "") f[69] = num(l["Base ReteIVA"]);
    f[90] = docProv;
    f[91] = pad(MARCA_PROVEEDOR, 11);
    f[92] = anioP;
    f[93] = mesP;
    f[94] = diaP;
    if (rol === "cxp") {
      const [anioV, mesV, diaV] = fechaPartes(texto(l["Vence"]));
      f[21] = num(l["Forma pago Pyme"]);
      f[103] = CRUCE;
      f[104] = docProv;
      f[105] = 1;
      f[106] = anioV;
      f[107] = mesV;
      f[108] = diaV;
    }
    return f;
  });

  const rango = (k: string) => {
    if (!k) return "";
    return `${MESES[Number(k.slice(4, 6)) - 1] ?? ""} ${Number(k.slice(6, 8))}/${k.slice(0, 4)}`;
  };
  const aoa: unknown[][] = [
    [cfg.empresa.slice(0, 50)],
    ["MODELO PARA LA IMPORTACION DE MOVIMIENTO CONTABLE - MODELO GENERAL"],
    [`De :  ${rango(desde)}   A :  ${rango(hasta)}`],
    [],
    COLUMNAS_PYME.map(([h]) => h),
    ...datos,
  ];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: ncol - 1 } });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Hoja1");
  return XLSX.write(wb, { bookType: "xlsx", type: "buffer" }) as Buffer;
}
