/**
 * Motor de la herramienta "Causación Interfaces v2" (admin).
 *
 * A diferencia del motor Python/TS de `causacionInterfaces` (que parte de los
 * Excel de la DIAN + parametrización), éste parte de las facturas ya leídas del
 * XML y de las cuentas/retenciones que el usuario asignó a cada una. La lógica
 * de líneas (gasto, IVA, INC, otros, obsequios, retenciones, cuadre contra la
 * cuenta por pagar), el redondeo y el formato del archivo son los mismos que
 * `script_compras.py` para que la interfaz sea importable igual en Siigo.
 *
 * Modo "venta" (facturas emitidas): el tercero es el cliente, ingresos e IVA van
 * al crédito, las retenciones que nos practican al débito, y la contrapartida es
 * la cuenta por cobrar. El comprobante sale del prefijo de la factura (el
 * consecutivo es su número), no de un contador.
 */
import type { XmlCausacionItem } from "./siigoAccountingService.js";
import { COLUMNAS_SIIGO, COLUMNAS_INFORME, type FilaSiigo, type FilaInforme, type CeldaValor } from "./causacionInterfaces/types.js";
import { dividirEnPartes } from "./causacionInterfaces/compras.js";
import { extraerPrefijoConsecutivo } from "./causacionInterfaces/ventas.js";
import { abreviarTipoDocumento, limpiarCuenta, redondear, redondear4, toNumero } from "./causacionInterfaces/normalize.js";

/** Cuentas y retenciones efectivas de una factura (sugeridas + ediciones del usuario). */
export interface Ci2Asignacion {
  cuentaGasto: string;
  cuentaPorPagar: string;
  cuentaIngresoObsequios: string;
  cuentaOtrosImpuestos: string;
  cuentaDescuento: string;
  cuentaRecargo: string;
  /** Ventas: cuenta de cada impuesto extra (Bolsas, ICUI…) por nombre, y del INC sin fila en la tabla ("INC 8%"). */
  cuentasImpuestos?: Record<string, string>;
  centroCostos: string;
  tarifaRetefuente: number;
  tarifaReteica: number; // por mil
  tarifaReteiva: number;
  baseMinimaReteicaEspecial: number;
  /** Cuenta de gasto por ítem (índice en `items`) cuando difiere de `cuentaGasto`. */
  cuentasItems?: Record<string, string>;
}

export interface Ci2FacturaInput {
  id: string;
  /** dd/mm/yyyy, como la espera Siigo. */
  fecha: string;
  documentType: string;
  isCreditNote: boolean;
  docNumber: string;
  nit: string;
  proveedor: string;
  items: XmlCausacionItem[];
  totalIva: number;
  asignacion: Ci2Asignacion;
}

export interface Ci2Impuesto {
  codigo: CeldaValor;
  cuenta: string;
  baseMinima: number;
}

/** Fila de la tabla maestra de impuestos tal como se guarda en la empresa. */
type FilaImpuesto = Record<string, unknown>;

export interface Ci2Opciones {
  modo?: "compra" | "venta";
  impuestos: FilaImpuesto[];
  tipoCompras: string;
  tipoComprasNc: string;
  /** Consecutivo inicial por tipo de comprobante. */
  consecutivoInicial: Record<string, number>;
  /** Ventas: tipo de comprobante por prefijo de la factura (la nota crédito cae al de factura si no tiene el suyo). */
  tipoPorPrefijo?: { factura: Record<string, string>; notaCredito: Record<string, string> };
}

export interface Ci2Resultado {
  salida: FilaSiigo[];
  partes: FilaSiigo[][];
  informe: FilaInforme[];
  advertencias: string[];
  /** Facturas que no se pudieron armar (no consumen consecutivo). */
  errores: { id: string; docNumber: string; proveedor: string; mensaje: string }[];
  /** Comprobante asignado a cada factura generada. */
  asignados: { id: string; tipo: string; consecutivo: number }[];
  /** Siguiente consecutivo por tipo, para recordar en la empresa. */
  siguientes: Record<string, number>;
  resumen: { documentos: number; notasCredito: number; lineas: number; partes: number };
}

class Ci2Error extends Error {}

const texto = (v: unknown) => (v === null || v === undefined ? "" : String(v));

function buscarImpuesto(impuestos: FilaImpuesto[], tipo: string, tarifa: number, columna = "Compras"): Ci2Impuesto {
  const tipoU = tipo.toUpperCase().trim();
  const objetivo = redondear4(tarifa);
  const fila = impuestos.find((r) => {
    if (redondear4(toNumero(r["Tarifa"])) !== objetivo) return false;
    if (tipoU === "INC") return texto(r["Nombre"]).toUpperCase().includes("IMPOCONSUMO");
    return texto(r["Tipo de impuesto"]).toUpperCase().includes(tipoU);
  });
  if (!fila) throw new Ci2Error(`No se encontró el impuesto ${tipoU} con tarifa ${tarifa} en la tabla de impuestos de la empresa.`);
  // En ventas las notas crédito usan la columna de devoluciones si la tabla la tiene y la fila la llena.
  const cuenta = limpiarCuenta(fila[columna]) || (columna === "Devolución ventas" ? limpiarCuenta(fila["Ventas"]) : "");
  return { codigo: (fila["Código"] ?? "") as CeldaValor, cuenta, baseMinima: toNumero(fila["Base"]) };
}

function exigirCuenta(valor: string, campo: string, motivo: string): string {
  const c = limpiarCuenta(valor);
  if (c === "" || c === "0") throw new Ci2Error(`${motivo} Falta asignar "${campo}".`);
  return c;
}

/** Clave estable del nombre de un impuesto (sirve como llave en Mongo). */
export const claveImpuesto = (nombre: string) => String(nombre ?? "").replace(/[.$]/g, "").trim();
const sumarNombre = (m: Map<string, number>, nombre: string, v: number) => m.set(claveImpuesto(nombre), (m.get(claveImpuesto(nombre)) ?? 0) + v);
/**
 * Descuento a nivel documento que realmente se causa. En las facturas con obsequios el descuento que trae
 * la DIAN es el IVA de los obsequios (el comprador no lo paga) y ya queda neteado con la línea de OBSEQUIO:
 * causarlo otra vez duplicaría ese valor.
 */
export function descuentoNeto(items: XmlCausacionItem[]): number {
  const ag = agregar(items);
  const obsequios = [...ag.obsequios.values()].reduce((x, y) => x + y, 0);
  const neto = ag.descuento - (obsequios > 0 ? obsequios : 0);
  return Math.abs(neto) <= 1 ? 0 : Math.max(neto, 0);
}

/** Nombres de los impuestos extra de la factura (sin el INC ya modelado por tarifa). */
export function nombresImpuestosExtra(items: XmlCausacionItem[]): string[] {
  return [...agregar(items).otros.keys()];
}

const sumar = (m: Map<number, number>, k: number, v: number) => m.set(k, (m.get(k) ?? 0) + v);

interface Agregados {
  baseGasto: number;
  /** Base por ítem (índice) para repartir el gasto entre cuentas. */
  basesItem: Map<number, number>;
  iva: Map<number, number>;
  inc: Map<number, number>;
  /** Impuestos extra por nombre (en compras se causan juntos en una sola cuenta). */
  otros: Map<string, number>;
  obsequios: Map<number, number>;
  descuento: number;
  recargo: number;
}

function agregar(items: XmlCausacionItem[]): Agregados {
  const a: Agregados = { baseGasto: 0, basesItem: new Map(), iva: new Map(), inc: new Map(), otros: new Map(), obsequios: new Map(), descuento: 0, recargo: 0 };
  for (const [idx, it] of items.entries()) {
    const extras = it.extraTaxes || [];
    if (it.isSurcharge) {
      for (const e of extras) {
        if (e.taxName === "Descuento") a.descuento += -e.amount;
        else if (e.taxName === "Recargo") a.recargo += e.amount;
      }
      continue;
    }
    if (it.isGift) {
      for (const e of extras) {
        if (e.taxName === "IVA Obsequio") sumar(a.obsequios, redondear4(e.percent), e.amount);
        else sumarNombre(a.otros, e.taxName, e.amount);
      }
      continue;
    }
    a.baseGasto += it.base;
    a.basesItem.set(idx, it.base);
    if (it.ivaPercent > 0) sumar(a.iva, redondear4(it.ivaPercent), (it.base * it.ivaPercent) / 100);
    if (it.incPercent > 0) sumar(a.inc, redondear4(it.incPercent), (it.base * it.incPercent) / 100);
    for (const e of extras) if (!(e.taxName === "INC" && it.incPercent > 0)) sumarNombre(a.otros, e.taxName, e.amount);
  }
  return a;
}

/** Arma las líneas e informe de UNA factura; lanza Ci2Error si falta algo. */
function armarFactura(f: Ci2FacturaInput, tipo: string, consecutivo: number, opts: Ci2Opciones) {
  const as = f.asignacion;
  const ag = agregar(f.items);
  const esNc = f.isCreditNote;
  const venta = opts.modo === "venta";
  const natBase: "D" | "C" = venta ? "C" : "D";
  // Retenciones: en compras las practicamos (pasivo, crédito); en ventas nos las practican (activo, débito).
  const natRet: "D" | "C" = venta ? "D" : "C";
  const contra = (n: "D" | "C"): "D" | "C" => (n === "D" ? "C" : "D");
  const colCuenta = venta ? (esNc && opts.impuestos.some((r) => "Devolución ventas" in r) ? "Devolución ventas" : "Ventas") : "Compras";
  const impuesto = (tipoImp: string, tarifa: number) => buscarImpuesto(opts.impuestos, tipoImp, tarifa, colCuenta);
  const tipoDocAbrev = esNc ? "NC" : abreviarTipoDocumento(f.documentType);
  const descripcion = `${tipoDocAbrev} ${f.docNumber} ${f.proveedor}`;
  const lineas: FilaSiigo[] = [];

  const crear = (cuenta: unknown, debito: number, credito: number, desc: string, codImp: CeldaValor = "", centro = "") => {
    const l: FilaSiigo = {};
    for (const c of COLUMNAS_SIIGO) l[c] = "";
    l["Tipo de comprobante"] = tipo;
    l["Consecutivo comprobante"] = consecutivo;
    l["Fecha de elaboración"] = f.fecha;
    l["Código cuenta contable"] = limpiarCuenta(cuenta);
    l["Identificación tercero"] = f.nit;
    l["Código impuesto"] = codImp;
    l["Descripción"] = desc.trim().slice(0, 50);
    l["Código centro/subcentro de costos"] = centro;
    l["Débito"] = debito;
    l["Crédito"] = credito;
    lineas.push(l);
  };
  const agregarLinea = (cuenta: unknown, valor: number, nat: "D" | "C", desc: string, codImp: CeldaValor = "", centro = "") => {
    const v = redondear(valor);
    if (v === 0) return;
    const debito = nat === "D" ? (esNc ? 0 : v) : esNc ? v : 0;
    const credito = nat === "D" ? (esNc ? v : 0) : esNc ? 0 : v;
    crear(cuenta, debito, credito, desc, codImp, centro);
  };

  const baseGasto = ag.baseGasto;
  const tarifasIva = [...ag.iva.keys()].sort((x, y) => y - x);
  const ivaTotal = tarifasIva.reduce((s, t) => s + redondear(ag.iva.get(t)!), 0);

  if (baseGasto > 0) {
    // Un renglón de gasto por cuenta distinta (en orden de aparición); los ítems sin cuenta propia van a la del proveedor.
    const porCuenta = new Map<string, number>();
    for (const [idx, base] of ag.basesItem) {
      const propia = limpiarCuenta(as.cuentasItems?.[String(idx)] ?? "");
      const cuenta = exigirCuenta(propia || as.cuentaGasto, venta ? "Cuenta de ingreso" : "Cuenta de gasto", venta ? "Hay base de ingreso." : "Hay base de gasto.");
      porCuenta.set(cuenta, (porCuenta.get(cuenta) ?? 0) + base);
    }
    for (const [cuenta, base] of porCuenta) agregarLinea(cuenta, base, natBase, descripcion, "", as.centroCostos);
  }
  for (const t of tarifasIva) {
    if (t <= 0) continue;
    const imp = impuesto("IVA", t);
    agregarLinea(imp.cuenta, ag.iva.get(t)!, natBase, `IVA${t} ${descripcion}`, imp.codigo);
  }
  for (const t of [...ag.inc.keys()].sort((x, y) => y - x)) {
    // En ventas, si la tabla no tiene el impoconsumo de esa tarifa, se pide su cuenta aparte ("INC 8%").
    const cuentaAparte = venta ? limpiarCuenta(as.cuentasImpuestos?.[`INC ${t}%`] ?? "") : "";
    const imp = cuentaAparte ? { cuenta: cuentaAparte, codigo: "" as CeldaValor, baseMinima: 0 } : impuesto("INC", t);
    agregarLinea(imp.cuenta, ag.inc.get(t)!, natBase, `IMPOCONSUMO ${descripcion}`, imp.codigo);
  }
  if (venta) {
    // Cada impuesto se declara en un formulario distinto: una línea y una cuenta por impuesto.
    for (const [nombre, valor] of ag.otros) {
      if (redondear(valor) === 0) continue;
      const cuenta = exigirCuenta(as.cuentasImpuestos?.[nombre] ?? "", `Cuenta de ${nombre}`, `Hay ${nombre}.`);
      agregarLinea(cuenta, valor, natBase, `${nombre} ${descripcion}`);
    }
  } else {
    const otros = [...ag.otros.values()].reduce((x, y) => x + y, 0);
    if (redondear(otros) !== 0) {
      agregarLinea(exigirCuenta(as.cuentaOtrosImpuestos, "Cuenta de otros impuestos", "Hay otros impuestos."), otros, natBase, `OTROS IMP ${descripcion}`);
    }
  }
  let ivaObsequios = 0;
  for (const t of [...ag.obsequios.keys()].sort((x, y) => y - x)) {
    const v = ag.obsequios.get(t)!;
    ivaObsequios += v;
    const imp = impuesto("IVA", t);
    agregarLinea(imp.cuenta, v, natBase, `IVA${t} OBSEQUIO ${descripcion}`, imp.codigo);
  }
  if (redondear(ivaObsequios) > 0) {
    agregarLinea(exigirCuenta(as.cuentaIngresoObsequios, "Cuenta de ingreso por obsequios", "Hay obsequios."), ivaObsequios, contra(natBase), `OBSEQUIO ${descripcion}`);
  }
  if (redondear(ag.recargo) > 0) {
    agregarLinea(exigirCuenta(as.cuentaRecargo, "Cuenta de recargo", "Hay recargo a nivel documento."), ag.recargo, natBase, `RECARGO ${descripcion}`);
  }
  const descuento = descuentoNeto(f.items);
  if (redondear(descuento) > 0) {
    agregarLinea(exigirCuenta(as.cuentaDescuento, "Cuenta de descuento", "Hay descuento a nivel documento."), descuento, contra(natBase), `DESCUENTO ${descripcion}`);
  }

  // Retenciones (misma base y fórmulas que el motor actual).
  let baseMinRf = 0, baseMinRica = 0, baseMinRiva = 0, rf = 0, rica = 0, riva = 0;
  if (as.tarifaRetefuente > 0) {
    const imp = impuesto("RETEFUENTE", as.tarifaRetefuente);
    baseMinRf = imp.baseMinima;
    if (baseGasto >= baseMinRf) {
      rf = (baseGasto * as.tarifaRetefuente) / 100;
      agregarLinea(imp.cuenta, rf, natRet, `RF ${descripcion}`, imp.codigo);
    }
  }
  if (as.tarifaReteica > 0) {
    const imp = impuesto("RETEICA", as.tarifaReteica);
    baseMinRica = as.baseMinimaReteicaEspecial || imp.baseMinima;
    if (baseGasto >= baseMinRica) {
      rica = (baseGasto * as.tarifaReteica) / 1000;
      agregarLinea(imp.cuenta, rica, natRet, `RICA ${descripcion}`, imp.codigo);
    }
  }
  if (as.tarifaReteiva > 0 && ivaTotal > 0) {
    const imp = impuesto("RETEIVA", as.tarifaReteiva);
    baseMinRiva = imp.baseMinima;
    if (ivaTotal >= baseMinRiva) {
      riva = (ivaTotal * as.tarifaReteiva) / 100;
      agregarLinea(imp.cuenta, riva, natRet, `RIVA ${descripcion}`, imp.codigo);
    }
  }

  // Cuadre contra la cuenta por pagar (forma de pago).
  let deb = 0, cre = 0;
  for (const l of lineas) { deb += toNumero(l["Débito"]); cre += toNumero(l["Crédito"]); }
  const dif = deb - cre;
  if (venta) {
    // Las líneas ya llevan el sentido de la nota crédito: la cuenta por cobrar es lo que falta para cuadrar.
    const cuenta = exigirCuenta(as.cuentaPorPagar, "Cuenta por cobrar", "Falta la contrapartida.");
    const v = redondear(Math.abs(dif));
    if (v !== 0) crear(cuenta, dif < 0 ? v : 0, dif < 0 ? 0 : v, descripcion);
  } else {
    agregarLinea(exigirCuenta(as.cuentaPorPagar, "Cuenta por pagar / forma de pago", "Falta la contrapartida."), esNc ? dif * -1 : dif, "C", descripcion);
  }

  const inf: FilaInforme = {};
  for (const c of COLUMNAS_INFORME) inf[c] = "";
  inf["Tipo documento"] = tipoDocAbrev;
  inf["Número factura"] = f.docNumber;
  inf["Fecha"] = f.fecha;
  inf["NIT proveedor"] = f.nit;
  inf["Proveedor"] = f.proveedor;
  inf["Base gasto"] = redondear(baseGasto);
  inf["IVA 19"] = redondear(ag.iva.get(19) ?? 0);
  inf["IVA 5"] = redondear(ag.iva.get(5) ?? 0);
  inf["IVA obsequios"] = redondear(ivaObsequios);
  inf["IVA total base ReteIVA"] = redondear(ivaTotal);
  inf["Tarifa retefuente"] = as.tarifaRetefuente;
  inf["Base mínima retefuente"] = redondear(baseMinRf);
  inf["Base usada retefuente"] = redondear(baseGasto);
  inf["Retefuente aplicada"] = redondear(rf);
  inf["Tarifa ReteICA"] = as.tarifaReteica;
  inf["Base mínima ReteICA"] = redondear(baseMinRica);
  inf["Base usada ReteICA"] = redondear(baseGasto);
  inf["ReteICA aplicada"] = redondear(rica);
  inf["Tarifa ReteIVA"] = as.tarifaReteiva;
  inf["Base mínima ReteIVA"] = redondear(baseMinRiva);
  inf["Base usada ReteIVA"] = redondear(ivaTotal);
  inf["ReteIVA aplicada"] = redondear(riva);
  inf["Total retenciones"] = redondear(rf) + redondear(rica) + redondear(riva);
  inf["Tipo comprobante Siigo"] = tipo;
  inf["Consecutivo Siigo"] = consecutivo;

  const aviso = Math.abs(ivaTotal + redondear(ivaObsequios) - f.totalIva) > 1
    ? `${descripcion}: IVA calculado ${ivaTotal + redondear(ivaObsequios)} difiere del IVA del documento ${f.totalIva}.`
    : "";
  return { lineas, inf, aviso };
}

/** Genera la interfaz Siigo + informe de retenciones para las facturas indicadas, en el orden dado. */
export function generarInterfaz(facturas: Ci2FacturaInput[], opts: Ci2Opciones): Ci2Resultado {
  const salida: FilaSiigo[] = [];
  const informe: FilaInforme[] = [];
  const advertencias: string[] = [];
  const errores: Ci2Resultado["errores"] = [];
  const asignados: Ci2Resultado["asignados"] = [];
  const contador: Record<string, number> = { ...opts.consecutivoInicial };
  let notasCredito = 0;

  const venta = opts.modo === "venta";
  for (const f of facturas) {
    let tipo: string;
    let consecutivo: number;
    if (venta) {
      // El comprobante sale del prefijo de la factura y su consecutivo es el número de la factura.
      const { prefijo, consecutivo: num } = extraerPrefijoConsecutivo(f.docNumber);
      const mapas = opts.tipoPorPrefijo || { factura: {}, notaCredito: {} };
      tipo = (f.isCreditNote ? mapas.notaCredito[prefijo] : "") || mapas.factura[prefijo] || "";
      consecutivo = num;
      if (!tipo) {
        errores.push({ id: f.id, docNumber: f.docNumber, proveedor: f.proveedor, mensaje: `No hay tipo de comprobante para el prefijo "${prefijo}".` });
        continue;
      }
      if (!(consecutivo > 0)) {
        errores.push({ id: f.id, docNumber: f.docNumber, proveedor: f.proveedor, mensaje: "No se pudo sacar el consecutivo del número de la factura." });
        continue;
      }
    } else {
      tipo = f.isCreditNote ? opts.tipoComprasNc : opts.tipoCompras;
      if (!tipo) {
        errores.push({ id: f.id, docNumber: f.docNumber, proveedor: f.proveedor, mensaje: "La empresa no tiene configurado el tipo de comprobante (Configuración → Comprobantes)." });
        continue;
      }
      consecutivo = contador[tipo] ?? 1;
    }
    try {
      const r = armarFactura(f, tipo, consecutivo, opts);
      salida.push(...r.lineas);
      informe.push(r.inf);
      if (r.aviso) advertencias.push(r.aviso);
      asignados.push({ id: f.id, tipo, consecutivo });
      if (!venta) contador[tipo] = consecutivo + 1;
      if (f.isCreditNote) notasCredito += 1;
    } catch (e) {
      if (!(e instanceof Ci2Error)) throw e;
      errores.push({ id: f.id, docNumber: f.docNumber, proveedor: f.proveedor, mensaje: e.message });
    }
  }

  const balances = new Map<string, number>();
  for (const l of salida) {
    const k = `${l["Tipo de comprobante"]}-${l["Consecutivo comprobante"]}`;
    balances.set(k, (balances.get(k) ?? 0) + toNumero(l["Débito"]) - toNumero(l["Crédito"]));
  }
  const descuadrados = [...balances.entries()].filter(([, d]) => d !== 0);
  if (descuadrados.length) {
    advertencias.push(`Hay ${descuadrados.length} comprobante(s) descuadrado(s): ${descuadrados.slice(0, 10).map(([k, d]) => `${k} (dif. ${d})`).join(", ")}. Revisa antes de importar a Siigo.`);
  }

  const partes = dividirEnPartes(salida);
  const siguientes: Record<string, number> = {};
  if (!venta) for (const a of asignados) siguientes[a.tipo] = contador[a.tipo];
  return {
    salida, partes, informe, advertencias, errores, asignados, siguientes,
    resumen: { documentos: asignados.length, notasCredito, lineas: salida.length, partes: partes.length },
  };
}
