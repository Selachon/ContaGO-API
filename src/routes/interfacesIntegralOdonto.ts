/**
 * Interfaces Integral Odonto Avanzada (herramienta admin).
 *
 * Copia de Causación Interfaces 2.0 sobre la que se hacen los ajustes especiales de esa empresa.
 * Comparte servicios, empresas y parámetros; las rutas y los jobs son propios.
 *
 * Dos módulos con el mismo flujo: compras (facturas recibidas, por mes) y ventas
 * (facturas emitidas, desde el listado de la DIAN). Las rutas que sirven a ambos
 * toman `tipo=venta` en la query o el cuerpo; sin él, compras.
 *
 * Trae las facturas recibidas de la DIAN con el token (NIT del token = NIT de la
 * empresa), las lista con cuentas/retenciones sugeridas desde los parámetros por
 * proveedor, y en vez de causarlas en Siigo genera la interfaz de importación
 * (con consecutivos y tabla de impuestos de la empresa).
 *
 *   POST /api/interfaces-integral-odonto/empresas/:id/traer              → {jobId}  (token + mes)
 *   POST /api/interfaces-integral-odonto/empresas/:id/traer-ventas       → {jobId}  (token + listado de emitidos)
 *   GET  /api/interfaces-integral-odonto/traer/:jobId                    → estado del job
 *   GET  /api/interfaces-integral-odonto/empresas/:id/facturas           → lista con sugerencias
 *   PATCH /api/interfaces-integral-odonto/empresas/:id/facturas/:fid     → edición manual
 *   POST /api/interfaces-integral-odonto/empresas/:id/facturas/estado    → descartar / restaurar
 *   PUT  /api/interfaces-integral-odonto/empresas/:id/proveedores/:nit   → guardar como parámetro del proveedor
 *   POST /api/interfaces-integral-odonto/empresas/:id/generar            → crea la interfaz
 *   GET  /api/interfaces-integral-odonto/empresas/:id/interfaces         → historial
 *   GET  /api/interfaces-integral-odonto/interfaces/:iid/archivos/:n     → descarga
 *   GET  /api/interfaces-integral-odonto/interfaces/:iid/zip             → todos los documentos en un ZIP
 *   POST /api/interfaces-integral-odonto/interfaces/:iid/revertir        → libera las facturas
 */
import { Router, type NextFunction, type Request, type Response } from "express";
import multer from "multer";
import { cuentaIvaDeRama } from "../services/ci2Motor.js";
import { perfilesDesdeBalance } from "../services/causacionInterfaces/balanceTercero.js";
import { v4 as uuidv4 } from "uuid";
import { requireToolAccess } from "../middleware/requireToolAccess.js";
import { exigirEmpresaHabilitada } from "../services/empresasNitGuard.js";
import { getUserById } from "../services/database.js";
import { requireAuth } from "../middleware/auth.js";
import { ingestFromDian, ingestNewByDateRange, type IngestResult } from "../services/siigoDianIngestService.js";
import { attachJobs, isJobAborted } from "../services/jobGuard.js";
import type { ProgressData } from "../types/dian.js";
import {
  detectarConsecutivosDesdeAuxiliar,
  getComprobantesEmpresa,
  getSistemaContable,
  getConfigWorldOffice,
  getProductosWo,
  getConfigSiigoPyme,
  importarFormasPagoPyme,
  importarCentrosPyme,
  generarTablaImpuestosPyme,
  getCuentasConCentroPyme,
  type ConfigSiigoPyme,
  getConfigInventario,
  getContextoCi2,
  getConfigIvaRama,
  getConfigPagos,
  getCodigosPuc,
  setConfigPagos,
  getCuentasConVencimiento,
  type ConfigPagos,
  marcarBalanceTercero,
  setConfigIvaRama,
  type ConfigInventario,
  setConfigInventario,
  setProductosWo,
  getEmpresaNit,
  importarPuc,
  getTabla,
  guardarConsecutivos,
  iniciarTablaConPlantilla,
  listEmpresasForUser,
  setTabla,
  userCanAccessEmpresa,
} from "../services/contabilizacionEmpresasService.js";
import JSZip from "jszip";
import fs from "fs";
import os from "os";
import path from "path";
import { extractInvoiceDataFromXml } from "../services/xmlParser.js";
import { generateExcelFile, generateExcelFilename } from "../services/excelGenerator.js";
import { escribirTercerosXlsm } from "../services/causacionInterfaces/tercerosXlsm.js";
import { filaTerceroDesdeFactura } from "../services/causacionInterfaces/terceros.js";
import { PLANTILLA_TERCEROS_BASE64 } from "../services/causacionInterfaces/plantillaTercerosBase.js";
import type { InvoiceData } from "../types/dianExcel.js";
import type { FilaTercero } from "../services/causacionInterfaces/types.js";
import { categoriaDocumento, conceptoSugeridoPyme, cuentaBaseInventario, cuentaIvaMayor, descuentoNeto, generarInterfaz, nombresImpuestosExtra, type Ci2Asignacion, type Ci2FacturaInput } from "../services/ci2Motor.js";
import { extractCufesFromExcel, resolveExcelBuffer } from "../services/dianExcelCufes.js";
import {
  cufesRegistrados,
  editarAsignacion,
  getPdfFactura,
  getCuentasGlobales,
  setCuentasGlobales,
  CAMPOS_GLOBALES,
  type Ci2CuentasGlobales,
  type Ci2Tipo,
  getPrefijos,
  getComprobantesSiigo,
  setComprobantesSiigo,
  esDuenoEmpresa,
  tablaImpuestosDeSiigo,
  guardarLoginSiigo,
  leerLoginSiigo,
  metaLoginSiigo,
  olvidarLoginSiigo,
  getTiposDocumento,
  getAutorretencion,
  setTiposDocumento,
  setPrefijos,
  ensureCi2Indexes,
  facturasPorIds,
  getArchivoInterfaz,
  getInterfaz,
  guardarFacturasTraidas,
  guardarInterfaz,
  facturasConArchivos,
  listarFacturas,
  listarInterfaces,
  marcarEstado,
  marcarRevision,
  revertirInterfaz,
  type Ci2Factura,
} from "../services/ci2Store.js";
import { escribirInformeRetenciones, escribirInterfazSiigo } from "../services/causacionInterfaces/excelWriter.js";
import { escribirInterfazWo } from "../services/causacionInterfaces/worldOfficeWriter.js";
import { escribirInterfazSiigoPyme } from "../services/causacionInterfaces/siigoPymeWriter.js";
import { escribirInterfazSyscafe } from "../services/causacionInterfaces/syscafeWriter.js";
import { limpiarCuenta, toNumero } from "../services/causacionInterfaces/normalize.js";
import { tryExtractDianUrlFromSafeLink } from "../middleware/validateDianUrl.js";
import { rateLimit } from "../middleware/rateLimit.js";
import { fetchSiigoComprobantes, fetchSiigoParametrizacion, SiigoWebError } from "../services/siigoWebSyncService.js";

export const ODONTO_TOOL_ID = "interfaces-integral-odonto";

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 40 * 1024 * 1024, files: 1 } });

const router = Router();
router.use(requireAuth);
router.use(requireToolAccess(ODONTO_TOOL_ID));
// Los índices se crean en la primera petición (al importar el módulo Mongo aún no está conectado).
let indicesListos: Promise<void> | null = null;
router.use((_req: Request, _res: Response, next: NextFunction) => {
  indicesListos ??= ensureCi2Indexes().catch((e) => {
    indicesListos = null;
    console.error("[CI2] índices:", e);
  });
  indicesListos.then(() => next());
});

const limpiarMapa = (m: unknown): Record<string, string> =>
  m && typeof m === "object" ? Object.fromEntries(Object.entries(m as Record<string, unknown>).map(([k, v]) => [k.trim().toUpperCase(), String(v ?? "").trim()]).filter(([k, v]) => k && v)) : {};
const nitDigitos = (v: unknown) => String(v ?? "").split("-")[0].replace(/\D/g, "");
const err = (res: Response, e: unknown, code = 400) =>
  res.status(code).json({ status: "error", mensaje: e instanceof Error ? e.message : String(e) });

async function exigirEmpresa(req: Request, res: Response): Promise<string | null> {
  const id = req.params.id;
  if (!(await userCanAccessEmpresa(id, req.user!.userId, req.user!.isAdmin))) {
    res.status(403).json({ status: "error", mensaje: "No tienes acceso a esta empresa." });
    return null;
  }
  try { await exigirEmpresaHabilitada(req.user!.userId, !!req.user!.isAdmin, id); } catch (e) { err(res, e, 403); return null; }
  return id;
}

// ── Parámetros por proveedor (tabla paramCompras de la empresa) ─────────────

const COL_NIT = "NIT Emisor";
const COLS_NIT_VENTAS = ["NIT Receptor", "NIT Emisor"];
const ASIG_VACIA: Ci2Asignacion = {
  cuentaGasto: "", cuentaPorPagar: "", cuentaIngresoObsequios: "", cuentaOtrosImpuestos: "", cuentaDescuento: "",
  cuentaRecargo: "", centroCostos: "", tarifaRetefuente: 0, tarifaReteica: 0, tarifaReteiva: 0, baseMinimaReteicaEspecial: 0,
};
/** Campo de la asignación → columna de paramCompras (o de paramVentas, que nombra distinto las cuentas). */
const COLUMNA_PARAM: Partial<Record<keyof Ci2Asignacion, string>> = {
  cuentaGasto: "Cuenta_gasto",
  cuentaPorPagar: "Cuenta_por_pagar",
  cuentaIngresoObsequios: "Cuenta_ingreso_obsequios",
  cuentaOtrosImpuestos: "Cuenta_otros_impuestos",
  centroCostos: "Centro_costos",
  tarifaRetefuente: "Tarifa_retefuente",
  tarifaReteica: "Tarifa_reteica",
  tarifaReteiva: "Tarifa_reteiva",
  baseMinimaReteicaEspecial: "Base_minima_reteica_especial",
  // Lo que el usuario decide por factura y casi siempre repite con el mismo proveedor.
  formaPago: "Forma_pago",
  formaPagoCodigo: "Forma_pago_codigo",
  comprobante: "Comprobante",
  tratamientoIva: "Tratamiento_iva",
  cuentaRetefuente: "Cuenta_retefuente",
  cuentaReteica: "Cuenta_reteica",
  cuentaReteiva: "Cuenta_reteiva",
  destino: "Destino",
  cuentasInventario: "Cuentas_inventario",
  cuentasIvaMayor: "Cuentas_iva_mayor",
  cuentasInc: "Cuentas_inc",
};
/** Campos del parámetro que se guardan como texto tal cual (sin limpiar como cuenta) o como JSON ({tarifa: cuenta}). */
const PARAM_TEXTO = new Set<keyof Ci2Asignacion>(["formaPago", "formaPagoCodigo", "comprobante", "tratamientoIva", "destino"]);
const PARAM_JSON = new Set<keyof Ci2Asignacion>(["cuentasInventario", "cuentasIvaMayor", "cuentasInc"]);
const COLUMNA_PARAM_VENTAS: Partial<Record<keyof Ci2Asignacion, string>> = {
  ...COLUMNA_PARAM,
  cuentaGasto: "Cuenta_ingreso",
  cuentaPorPagar: "Cuenta_por_cobrar",
  cuentaIngresoObsequios: "Cuenta_devolucion_ingreso",
};
// Integral Odonto: la cuenta por pagar/cobrar sale de la forma de pago de cada factura, y las de obsequios y otros impuestos son generales: no son del proveedor.
const NO_POR_PROVEEDOR: (keyof Ci2Asignacion)[] = ["cuentaPorPagar", "cuentaIngresoObsequios", "cuentaOtrosImpuestos"];
const quitarGenerales = (m: Partial<Record<keyof Ci2Asignacion, string>>) => Object.fromEntries(Object.entries(m).filter(([k]) => !NO_POR_PROVEEDOR.includes(k as keyof Ci2Asignacion))) as Partial<Record<keyof Ci2Asignacion, string>>;
const sinFormaPago = (m: Partial<Record<keyof Ci2Asignacion, string>>) => Object.fromEntries(Object.entries(m).filter(([k]) => k !== "formaPagoCodigo")) as Partial<Record<keyof Ci2Asignacion, string>>;
// En ventas la forma de pago es predeterminada (Parámetros) y no se aprende por cliente.
const columnasParam = (tipo: Ci2Tipo) => (tipo === "venta" ? sinFormaPago(quitarGenerales(COLUMNA_PARAM_VENTAS)) : quitarGenerales(COLUMNA_PARAM));
const slotParam = (tipo: Ci2Tipo) => (tipo === "venta" ? ("paramVentas" as const) : ("paramCompras" as const));
/** Compras por defecto; ventas solo cuando se pide expresamente (query `tipo` o cuerpo). */
const tipoDe = (req: Request): Ci2Tipo => ((req.query?.tipo ?? req.body?.tipo) === "venta" ? "venta" : "compra");

async function sugerenciasPorNit(empresaId: string, tipo: Ci2Tipo = "compra"): Promise<Map<string, Ci2Asignacion>> {
  const tabla = await getTabla(empresaId, slotParam(tipo));
  const mapa = new Map<string, Ci2Asignacion>();
  const colNit = tipo === "venta" ? COLS_NIT_VENTAS.find((c) => tabla?.columns.includes(c)) || COLS_NIT_VENTAS[0] : COL_NIT;
  for (const row of tabla?.rows || []) {
    const nit = nitDigitos(row[colNit]);
    if (!nit || mapa.has(nit)) continue;
    const a = { ...ASIG_VACIA };
    for (const [campo, col] of Object.entries(columnasParam(tipo))) {
      const v = row[col as string];
      const k = campo as keyof Ci2Asignacion;
      if (typeof ASIG_VACIA[k] === "number") (a as any)[k] = toNumero(v);
      else if (PARAM_TEXTO.has(k)) { const t = String(v ?? "").trim(); if (t) (a as any)[k] = t; }
      else if (PARAM_JSON.has(k)) {
        try { const o = typeof v === "string" && v.trim() ? JSON.parse(v) : null; if (o && typeof o === "object" && Object.keys(o).length) (a as any)[k] = o; } catch { /* celda editada a mano con formato inválido: se ignora */ }
      }
      else (a as any)[k] = limpiarCuenta(v);
    }
    mapa.set(nit, a);
  }
  return mapa;
}

/** Retenciones configuradas que el motor no causa porque la base de la factura no alcanza la base mínima de la tabla. */
function retencionesNoAplican(items: any[], ef: any, op: ReturnType<typeof opcionesRetencion>) {
  const normales = (items || []).filter((i) => !i.isGift && !i.isSurcharge);
  const base = normales.reduce((t, i) => t + (i.base || 0), 0);
  const iva = normales.reduce((t, i) => t + ((i.base || 0) * (i.ivaPercent || 0)) / 100, 0);
  const baseMin = (lista: { tarifa: number; baseMinima: number }[], tarifa: number) => lista.find((o) => o.tarifa === tarifa)?.baseMinima ?? 0;
  const libre = (k: string) => !ef.retForzadas?.includes(k);
  return {
    tarifaRetefuente: libre("tarifaRetefuente") && ef.tarifaRetefuente > 0 && base < baseMin(op.retefuente, ef.tarifaRetefuente),
    tarifaReteica: libre("tarifaReteica") && ef.tarifaReteica > 0 && base < (Number(ef.baseMinimaReteicaEspecial) || baseMin(op.reteica, ef.tarifaReteica)),
    tarifaReteiva: libre("tarifaReteiva") && ef.tarifaReteiva > 0 && iva < baseMin(op.reteiva, ef.tarifaReteiva),
  };
}

/** Tarifas de retención disponibles en la tabla de impuestos de la empresa. */
function opcionesRetencion(rows: Record<string, unknown>[], porCuenta = false) {
  const por = (tipo: string) => {
    const vistos = new Set<string>();
    const out: { tarifa: number; nombre: string; baseMinima: number; cuenta?: string }[] = [];
    for (const r of rows) {
      if (!String(r["Tipo de impuesto"] ?? "").toUpperCase().includes(tipo)) continue;
      const tarifa = toNumero(r["Tarifa"]);
      const cuenta = limpiarCuenta(r["Compras"]);
      const llave = porCuenta ? cuenta : String(tarifa);
      if (!(tarifa > 0) || vistos.has(llave)) continue;
      vistos.add(llave);
      // Siigo Pyme: varias cuentas comparten tarifa, así que cada opción es una cuenta.
      out.push({ tarifa, nombre: String(r["Nombre"] ?? "").trim(), baseMinima: toNumero(r["Base"]), ...(porCuenta ? { cuenta } : {}) });
    }
    return out.sort((a, b) => a.tarifa - b.tarifa || (a.cuenta ?? "").localeCompare(b.cuenta ?? ""));
  };
  return { retefuente: por("RETEFUENTE"), reteica: por("RETEICA"), reteiva: por("RETEIVA") };
}

/** Cuentas de impoconsumo de la tabla de impuestos por tarifa ("8": [cuentas]); solo importan las tarifas con más de una. */
function opcionesInc(rows: Record<string, unknown>[], venta: boolean): Record<string, string[]> {
  const col = venta ? "Ventas" : "Compras";
  const por: Record<string, string[]> = {};
  for (const r of rows) {
    if (!String(r["Nombre"] ?? "").toUpperCase().includes("IMPOCONSUMO")) continue;
    const cuenta = limpiarCuenta(r[col]);
    const t = String(Math.round(toNumero(r["Tarifa"]) * 10000) / 10000);
    if (!cuenta || t === "0") continue;
    (por[t] ||= []).includes(cuenta) || por[t].push(cuenta);
  }
  return Object.fromEntries(Object.entries(por).filter(([, l]) => l.length > 1));
}

/** Sugerencia del proveedor + lo editado en la factura; las cuentas globales de la empresa rellenan lo que siga vacío. */
function efectiva(f: Ci2Factura, sug?: Ci2Asignacion, glob: Ci2CuentasGlobales = {}, pyme: ConfigSiigoPyme | null = null, pagos: ConfigPagos | null = null): Ci2Asignacion {
  const a: Ci2Asignacion = { ...ASIG_VACIA, ...(sug || {}), ...f.asignacion };
  if (pyme) {
    // Siigo Pyme: la cuenta por pagar es la que Siigo Pyme relaciona con la forma de pago elegida en la factura.
    const forma = pyme.formasPago.find((x) => String(x.codigo) === String(a.formaPagoCodigo ?? "").trim());
    if (forma?.cuenta) a.cuentaPorPagar = forma.cuenta;
  }
  if (pagos && !pyme) {
    // Ventas: la forma de pago es la de la factura o, si no se eligió, la predeterminada; nunca la sugerencia del cliente.
    if (f.tipo === "venta") a.formaPagoCodigo = String(f.asignacion.formaPagoCodigo ?? "").trim() || pagos.defectoVentas || "";
    // La forma de pago se elige por nombre; su cuenta es la cuenta por pagar (compras) o por cobrar (ventas) de la factura.
    const forma = (f.tipo === "venta" ? pagos.ventas : pagos.compras).find((x) => x.nombre === String(a.formaPagoCodigo ?? "").trim());
    if (forma?.cuenta) a.cuentaPorPagar = forma.cuenta;
    // Ventas: una sola cuenta de ingreso para todas las facturas.
    if (f.tipo === "venta" && pagos.cuentaIngreso) { a.cuentaGasto = pagos.cuentaIngreso; a.cuentasItems = {}; }
  }
  a.retForzadas = (["tarifaRetefuente", "tarifaReteica", "tarifaReteiva"] as const).filter((k) => k in f.asignacion && Number((f.asignacion as any)[k]) > 0);
  for (const k of CAMPOS_GLOBALES) if (!limpiarCuenta(a[k])) a[k] = glob[k] || "";
  a.cuentasImpuestos = { ...(glob.cuentasImpuestos || {}) };
  return a;
}

/** Parámetros de Siigo Pyme de la empresa, o null si no es de Siigo Pyme. */
async function pymeDe(empresaId: string): Promise<ConfigSiigoPyme | null> {
  return (await getSistemaContable(empresaId)) === "siigopyme" ? getConfigSiigoPyme(empresaId) : null;
}

/** Tarifas de impoconsumo que existen en la tabla de impuestos de la empresa. */
const tarifasIncDe = (rows: Record<string, any>[] = []) =>
  new Set(rows.filter((r) => String(r["Nombre"] ?? "").toUpperCase().includes("IMPOCONSUMO")).map((r) => Math.round(Number(r["Tarifa"]) * 10000) / 10000));

/** IVA mayor valor que se sugiere por tarifa (misma rama que la cuenta de la base), para mostrarlo ya elegido. */
function ivaMayorSugerido(a: Ci2Asignacion, cfgInv: ConfigInventario | null): Record<string, string> {
  const out: Record<string, string> = {};
  if (!cfgInv || a.destino !== "inventario") return out;
  for (const [t, posibles] of Object.entries(cfgInv.ivaMayorValor)) {
    const c = cuentaIvaMayor(a, posibles, t, cuentaBaseInventario(a, cfgInv.base[t] || [], t));
    if (c) out[t] = c;
  }
  return out;
}

/** Cuenta de IVA mayor valor que corresponde a cada cuenta de gasto de la factura (para mostrarla ya elegida). */
function ivaRamaDe(a: Ci2Asignacion, items: Ci2Factura["items"], ivaRama: { cuentas: string[] }): Record<string, { sugerida: string; elegida: string }> {
  const out: Record<string, { sugerida: string; elegida: string }> = {};
  if (!ivaRama.cuentas.length) return out;
  items.forEach((it, i) => {
    if (it.isGift || it.isSurcharge || !(it.base > 0) || !(it.ivaPercent > 0)) return;
    const cuenta = limpiarCuenta(a.cuentasItems?.[String(i)] ?? "") || limpiarCuenta(a.cuentaGasto);
    if (!cuenta) return;
    const sugerida = cuentaIvaDeRama(cuenta, ivaRama.cuentas);
    out[cuenta] = { sugerida, elegida: limpiarCuenta(a.cuentasIvaRama?.[cuenta] ?? "") || sugerida };
  });
  return out;
}

function faltantesDe(a: Ci2Asignacion, items: Ci2Factura["items"], tipo: Ci2Tipo = "compra", tarifasInc: Set<number> = new Set(), worldOffice = false, cfgInv: ConfigInventario | null = null, pyme = false, ivaRama: { cuentas: string[]; inc: string } | null = null, formas: { nombre: string; cuenta: string }[] = []): string[] {
  const clasifica = !!cfgInv;
  const out: string[] = [];
  // Siigo Pyme: cada factura se asigna a mano a su forma de pago (catálogo de la empresa) y a su centro de costo.
  if (pyme && tipo === "compra") {
    if (!/^\d+$/.test(String(a.formaPagoCodigo ?? "").trim())) out.push("formaPagoCodigo");
    if (!/^\d+(\D+\d+)?$/.test(String(a.centroCostos ?? "").trim())) out.push("centroCostos");
    if (items.some((i) => !i.isGift && !i.isSurcharge && i.base > 0 && i.ivaPercent > 0) && !["compras", "servicios", "asumido", "mayor"].includes(String(a.tratamientoIva ?? ""))) out.push("tratamientoIva");
    if (a.tratamientoIva === "mayor" && !/^\d{4,}/.test(String(a.cuentaIvaMayorValor ?? "").trim()) && items.some((i) => !i.isGift && !i.isSurcharge && i.base > 0 && i.ivaPercent > 0)) out.push("cuentaIvaMayorValor");
  }
  if (clasifica && tipo === "compra" && a.destino !== "inventario" && a.destino !== "gasto") out.push("destino");
  // World Office: el producto sale de la tarifa de IVA (maestro de productos), no de la cuenta del proveedor.
  if (worldOffice) return out;
  // La cuenta del proveedor solo hace falta si algún ítem con base no tiene cuenta propia.
  const sinCuentaPropia = !(clasifica && a.destino === "inventario") && items.some((it, i) => !it.isGift && !it.isSurcharge && it.base > 0 && !limpiarCuenta(a.cuentasItems?.[String(i)] ?? ""));
  if (sinCuentaPropia && !limpiarCuenta(a.cuentaGasto)) out.push("cuentaGasto");
  if (!limpiarCuenta(a.cuentaPorPagar)) out.push("cuentaPorPagar");
  // Cuentas adicionales que el motor exige según lo que traiga la factura (descuento/recargo del documento, obsequios, otros impuestos).
  const extra = (it: Ci2Factura["items"][number]) => it.extraTaxes || [];
  // Descuento y recargo son un ítem de la factura: basta la cuenta elegida en el ítem o la general de la empresa.
  const cuentaDelItem = (tipo: string) => items.some((it, i) => it.isSurcharge && extra(it).some((e) => e.taxName === tipo && e.amount !== 0) && !!limpiarCuenta(a.cuentasItems?.[String(i)] ?? ""));
  if (descuentoNeto(items) > 0 && !limpiarCuenta(a.cuentaDescuento) && !cuentaDelItem("Descuento")) out.push("cuentaDescuento");
  if (items.some((it) => it.isSurcharge && extra(it).some((e) => e.taxName === "Recargo" && e.amount !== 0)) && !limpiarCuenta(a.cuentaRecargo) && !cuentaDelItem("Recargo")) out.push("cuentaRecargo");
  if (items.some((it) => it.isGift && extra(it).some((e) => e.taxName === "IVA Obsequio" && e.amount !== 0)) && !limpiarCuenta(a.cuentaIngresoObsequios)) out.push("cuentaIngresoObsequios");
  const nombres = nombresImpuestosExtra(items);
  // El impoconsumo cuya tarifa no está en la tabla de impuestos pide su propia cuenta ("INC 4%"), en compras y en ventas.
  const incs = new Set(items.filter((it) => !it.isGift && !it.isSurcharge && it.incPercent > 0).map((it) => Math.round(it.incPercent * 10000) / 10000));
  const inv = clasifica && a.destino === "inventario";
  if (inv) {
    // Cada tarifa de IVA con varias cuentas de inventario posibles pide elegir una.
    const tarifas = new Set(items.filter((it) => !it.isGift && !it.isSurcharge && it.base > 0).map((it) => String(it.ivaPercent > 0 ? Math.round(it.ivaPercent * 10000) / 10000 : 0)));
    for (const t of tarifas) {
      const posibles = cfgInv!.base[t] || [];
      if (posibles.length > 1 && !posibles.includes(limpiarCuenta(a.cuentasInventario?.[t] ?? ""))) out.push(`inv:${t}`);
      const iva = cfgInv!.ivaMayorValor[t] || [];
      if (Number(t) > 0 && iva.length > 1 && !cuentaIvaMayor(a, iva, t, cuentaBaseInventario(a, posibles, t))) out.push(`iva:${t}`);
    }
  }
  for (const t of incs) if (!inv && !(ivaRama && ivaRama.inc) && !tarifasInc.has(t) && !limpiarCuenta(a.cuentasImpuestos?.[`INC ${t}%`] ?? "")) out.push(`imp:INC ${t}%`);
  // Forma de pago por nombre (Odonto): debe elegirse una que tenga cuenta.
  if (formas.length && !pyme && !out.includes("formaPagoCodigo")) {
    const forma = formas.find((x) => x.nombre === String(a.formaPagoCodigo ?? "").trim());
    if (!forma || !forma.cuenta) out.push("formaPagoCodigo");
  }
  // IVA mayor valor por rama: la cuenta de gasto de cada ítem con IVA debe tener su cuenta de IVA en la lista de la empresa.
  if (ivaRama && ivaRama.cuentas.length && !inv && tipo === "compra") {
    const sinRama = items.some((it, i) => !it.isGift && !it.isSurcharge && it.base > 0 && it.ivaPercent > 0 && (() => {
      const cuenta = limpiarCuenta(a.cuentasItems?.[String(i)] ?? "") || limpiarCuenta(a.cuentaGasto);
      return !!cuenta && !limpiarCuenta(a.cuentasIvaRama?.[cuenta] ?? "") && !cuentaIvaDeRama(cuenta, ivaRama.cuentas);
    })());
    if (sinRama) out.push("ivaRama");
  }
  if (tipo === "venta") {
    // Cada impuesto pide su propia cuenta.
    for (const n of nombres) if (!limpiarCuenta(a.cuentasImpuestos?.[n] ?? "")) out.push(`imp:${n}`);
  } else if (!inv && !limpiarCuenta(a.cuentaOtrosImpuestos)) {
    // Sin cuenta general de otros impuestos, cada impuesto de la factura pide su propia cuenta (queda recordada para la empresa).
    for (const n of nombres) if (!limpiarCuenta(a.cuentasImpuestos?.[n] ?? "")) out.push(`imp:${n}`);
  }
  return out;
}

/** dd/mm/yyyy desde yyyy-mm-dd (formato de fecha de la interfaz Siigo). */
function fechaSiigo(iso: string): string {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || "");
  return m ? `${m[3]}/${m[2]}/${m[1]}` : iso || "";
}

// ── Traer facturas de la DIAN ───────────────────────────────────────────────

interface TraerJob {
  status: "processing" | "completed" | "error" | "cancelled";
  progress: ProgressData;
  userId: string;
  empresaId: string;
  createdAt: number;
  error?: string;
  result?: { nuevas: number; stats: IngestResult["stats"]; failures: IngestResult["failures"]; ignorados: number; errores: { fileName: string; message: string }[]; omitidos?: number };
}
const traerJobs = new Map<string, TraerJob>();
attachJobs("odonto-dian-traer", traerJobs as any);
setInterval(() => {
  const now = Date.now();
  for (const [id, j] of traerJobs) if (now - j.createdAt > 3 * 60 * 60 * 1000) traerJobs.delete(id);
}, 60_000);

/** El token debe ser de la MISMA empresa: NIT del parámetro `rk` = NIT de la empresa. Responde el error y devuelve false si no. */
async function tokenDeLaEmpresa(res: Response, empresaId: string, tokenUrl: string): Promise<boolean> {
  if (!/catalogo-vpfe\.dian\.gov\.co\/User\/AuthToken/i.test(tokenUrl)) {
    err(res, new Error("Token DIAN inválido. Pega la URL completa de ingreso (/User/AuthToken?...)."));
    return false;
  }
  const nitEmpresa = nitDigitos(await getEmpresaNit(empresaId));
  if (!nitEmpresa) {
    err(res, new Error("La empresa no tiene NIT. Defínelo en la configuración de la empresa."));
    return false;
  }
  let nitToken = "";
  try { nitToken = nitDigitos(new URL(tokenUrl).searchParams.get("rk") || ""); } catch { /* URL inválida */ }
  if (!nitToken) {
    err(res, new Error("No se pudo determinar el NIT del token (parámetro rk). Verifica que la URL esté completa."));
    return false;
  }
  if (nitToken !== nitEmpresa) {
    res.status(403).json({ status: "error", mensaje: `El token pertenece al NIT ${nitToken}, pero la empresa es NIT ${nitEmpresa}. No se permite traer facturas de otra empresa.` });
    return false;
  }
  return true;
}

/** Ventana de meses que se pueden traer: los últimos 12, o desde el mes habilitado al usuario (admin sin límite). */
async function mesPermitido(req: Request, month: string): Promise<boolean> {
  if (req.user?.isAdmin) return true;
  const now = new Date();
  const ventana = new Date(now.getFullYear(), now.getMonth() - 11, 1);
  const desdeVentana = `${ventana.getFullYear()}-${String(ventana.getMonth() + 1).padStart(2, "0")}`;
  const user = req.user?.userId ? await getUserById(req.user.userId) : null;
  const habilitado = user?.ingestFromMonth?.["causacion-interfaces-v2"];
  const piso = habilitado && /^\d{4}-\d{2}$/.test(habilitado) && habilitado < desdeVentana ? habilitado : desdeVentana;
  return month >= piso;
}

router.post("/empresas/:id/traer", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    const tokenUrl = tryExtractDianUrlFromSafeLink(String(req.body?.tokenUrl || "").trim());
    const month = String(req.body?.month || "").trim();
    if (!/^\d{4}-\d{2}$/.test(month)) return err(res, new Error("Selecciona el mes a traer."));
    if (!(await mesPermitido(req, month))) return err(res, new Error("Ese mes está fuera de los últimos 12 meses de tu plan. Escríbenos para habilitarlo."));

    if (!(await tokenDeLaEmpresa(res, empresaId, tokenUrl))) return;

    const [y, m] = month.split("-").map(Number);
    const fechaInicio = `${month}-01`;
    const fechaFin = `${month}-${String(new Date(y, m, 0).getDate()).padStart(2, "0")}`;
    const skipCufes = await cufesRegistrados(empresaId);

    const jobId = uuidv4();
    traerJobs.set(jobId, { status: "processing", progress: { step: "Iniciando...", current: 0, total: 0 }, userId: req.user!.userId, empresaId, createdAt: Date.now() });

    ingestNewByDateRange({
      companyId: `ci2:${empresaId}`,
      tokenUrl,
      fechaInicio,
      fechaFin,
      maxDocuments: 200,
      skipCufes,
      onProgress: (p) => { const j = traerJobs.get(jobId); if (j && j.status === "processing") j.progress = p; },
      isCancelled: () => isJobAborted(traerJobs.get(jobId) as any),
    })
      .then(async (result) => {
        const j = traerJobs.get(jobId);
        if (!j || isJobAborted(j as any)) return;
        const nuevas = await guardarFacturasTraidas(empresaId, result.items);
        j.status = "completed";
        j.result = {
          nuevas,
          stats: result.stats,
          failures: result.failures,
          ignorados: result.items.filter((i) => i.ignored).length,
          errores: result.items.filter((i) => !i.ok && !i.ignored).map((i) => ({ fileName: i.fileName, message: i.message || "" })),
        };
        j.progress = { step: "Completado", current: result.stats.downloaded, total: result.stats.listed };
      })
      .catch((e) => {
        const j = traerJobs.get(jobId);
        if (!j || isJobAborted(j as any)) return;
        j.status = "error";
        j.error = e instanceof Error ? e.message : "Error trayendo facturas de la DIAN";
      });

    res.json({ status: "ok", jobId, rango: { fechaInicio, fechaFin } });
  } catch (e) {
    err(res, e);
  }
});

// Ventas: el usuario sube el listado exportado de la DIAN (emitidos) y el token; se descargan XML y PDF por CUFE
// con el mismo mecanismo que "Descarga Masiva + Excel DIAN".
router.post("/empresas/:id/traer-ventas", upload.single("listado"), async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    const tokenUrl = tryExtractDianUrlFromSafeLink(String(req.body?.tokenUrl || "").trim());
    const archivo = req.file as Express.Multer.File | undefined;
    if (!archivo) return err(res, new Error("Sube el listado exportado desde la DIAN (.xlsx o .zip)."));
    if (!(await tokenDeLaEmpresa(res, empresaId, tokenUrl))) return;

    let listado;
    try {
      listado = await extractCufesFromExcel(await resolveExcelBuffer(archivo));
    } catch (e) {
      return err(res, new Error(`No se pudo leer el listado: ${e instanceof Error ? e.message : e}`));
    }
    if (listado.mixedDirections) return err(res, new Error("El listado mezcla documentos emitidos y recibidos. Exporta solo los emitidos."));
    if (listado.detectedDirection === "received") return err(res, new Error("El listado es de documentos RECIBIDOS. Para ventas exporta el de emitidos."));
    const cufes = [...new Set(listado.cufes)];
    if (!cufes.length) return err(res, new Error("El listado no contiene documentos válidos."));

    const skipCufes = await cufesRegistrados(empresaId);
    const nitEmpresa = nitDigitos(await getEmpresaNit(empresaId));
    const jobId = uuidv4();
    traerJobs.set(jobId, { status: "processing", progress: { step: "Iniciando...", current: 0, total: cufes.length }, userId: req.user!.userId, empresaId, createdAt: Date.now() });

    ingestFromDian({
      companyId: `ci2:${empresaId}`,
      tokenUrl,
      grupo: "Emitidos",
      providedCufes: cufes,
      skipCufes,
      // Por tandas: lo ya guardado se omite, así que repetir «Traer» continúa donde quedó y cada tanda se guarda completa.
      maxDocuments: 300,
      onProgress: (p) => { const j = traerJobs.get(jobId); if (j && j.status === "processing") j.progress = p; },
      isCancelled: () => isJobAborted(traerJobs.get(jobId) as any),
    })
      .then(async (result) => {
        const j = traerJobs.get(jobId);
        if (!j || isJobAborted(j as any)) return;
        // Solo lo emitido por la empresa: el documento soporte (donde la empresa es la compradora) no es una venta.
        const propias = result.items.filter((i) => !(i.ok && i.xml) || nitDigitos(i.xml.supplierNit) === nitEmpresa);
        const nuevas = await guardarFacturasTraidas(empresaId, propias, "venta");
        j.status = "completed";
        j.result = {
          nuevas,
          stats: result.stats,
          failures: result.failures,
          ignorados: result.items.filter((i) => i.ignored).length,
          errores: result.items.filter((i) => !i.ok && !i.ignored).map((i) => ({ fileName: i.fileName, message: i.message || "" })),
          omitidos: result.items.length - propias.length,
        };
        j.progress = { step: "Completado", current: result.stats.downloaded, total: result.stats.listed };
      })
      .catch((e) => {
        const j = traerJobs.get(jobId);
        if (!j || isJobAborted(j as any)) return;
        j.status = "error";
        j.error = e instanceof Error ? e.message : "Error trayendo facturas de la DIAN";
      });

    res.json({ status: "ok", jobId, documentos: cufes.length });
  } catch (e) {
    err(res, e);
  }
});

router.get("/traer/:jobId", (req: Request, res: Response) => {
  const j = traerJobs.get(req.params.jobId);
  if (!j) return res.status(404).json({ status: "error", mensaje: "Job no encontrado (¿se reinició el servidor?)." });
  res.json({ status: "ok", estado: j.status, progreso: j.progress, error: j.error, resultado: j.result });
});

router.post("/traer/:jobId/cancelar", (req: Request, res: Response) => {
  const j = traerJobs.get(req.params.jobId);
  if (j && j.status === "processing") j.status = "cancelled";
  res.json({ status: "ok" });
});

// ── Facturas ────────────────────────────────────────────────────────────────

router.get("/empresas/:id/facturas", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    const estado = ["pendiente", "generada", "descartada"].includes(String(req.query.estado)) ? (String(req.query.estado) as any) : undefined;
    const tipo = tipoDe(req);
    const [facturas, sugs, impuestos, glob, ctx] = await Promise.all([listarFacturas(empresaId, estado, tipo), sugerenciasPorNit(empresaId, tipo), getTabla(empresaId, "impuestos"), getCuentasGlobales(empresaId, tipo), getContextoCi2(empresaId)]);
    const tarifasInc = tarifasIncDe(impuestos?.rows);
    const worldOffice = ctx.sistema === "worldoffice";
    const pymeCfg = ctx.pyme;
    const siigoPyme = ctx.sistema === "siigopyme";
    const cfgInv = ctx.inventario;
    const opciones = opcionesRetencion(impuestos?.rows || [], siigoPyme);
    res.json({
      status: "ok",
      cuentasGlobales: glob,
      opcionesRetencion: opciones,
      opcionesInc: opcionesInc((impuestos?.rows || []) as any, tipo === "venta"),
      formasPago: tipo === "venta" ? ctx.pagos.ventas : ctx.pagos.compras,
      cuentaIngresoUnica: tipo === "venta" ? ctx.pagos.cuentaIngreso : "",
      facturas: facturas.map((f) => {
        const sug = sugs.get(f.nit);
        const ef = efectiva(f, sug, glob, pymeCfg, ctx.pagos);
        return { ...f, retNoAplica: retencionesNoAplican(f.items, ef, opciones), categoria: categoriaDocumento(f.documentType, f.isCreditNote), sugerencia: sug || null, efectiva: ef, conceptoSugerido: conceptoSugeridoPyme(f.docNumber, f.items, f.proveedor), faltantes: faltantesDe(ef, f.items, tipo, tarifasInc, worldOffice, cfgInv.activo && !f.isCreditNote ? cfgInv : null, siigoPyme, ctx.ivaRama, tipo === "venta" ? ctx.pagos.ventas : ctx.pagos.compras), ivaRama: ivaRamaDe(ef, f.items, ctx.ivaRama), ivaMayorSugerido: ivaMayorSugerido(ef, cfgInv.activo ? cfgInv : null), sinParametros: !sug };
      }),
    });
  } catch (e) {
    err(res, e, 500);
  }
});

router.get("/empresas/:id/facturas/:fid/pdf", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    const pdf = await getPdfFactura(empresaId, req.params.fid);
    if (!pdf) return res.status(404).json({ status: "error", mensaje: "Esta factura no tiene PDF guardado." });
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `inline; filename="${encodeURIComponent(pdf.nombre)}"`);
    res.send(pdf.data);
  } catch (e) {
    err(res, e, 500);
  }
});

router.patch("/empresas/:id/facturas/:fid", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    let f = await editarAsignacion(empresaId, req.params.fid, req.body || {});
    if (!f) return res.status(404).json({ status: "error", mensaje: "Factura no encontrada o ya incluida en una interfaz." });
    // Cuentas y retenciones editadas en una factura de un proveedor nuevo pasan a ser su parámetro (así las demás facturas lo sugieren); en uno que ya tiene parámetros solo cambian esa factura.
    const editados = (ks: readonly (keyof Ci2Asignacion)[]) => ks.filter((k) => req.body && k in req.body && req.body[k] !== null && k in f!.asignacion);
    // Descuento, recargo y otros impuestos son cuentas de la empresa, no del proveedor.
    const tipo = f.tipo;
    const aGlobal = editados(CAMPOS_GLOBALES) as (keyof Ci2CuentasGlobales)[];
    if (aGlobal.length) {
      await setCuentasGlobales(empresaId, Object.fromEntries(aGlobal.map((k) => [k, (f!.asignacion as any)[k]])), tipo);
      f = (await editarAsignacion(empresaId, req.params.fid, Object.fromEntries(aGlobal.map((k) => [k, null])))) || f;
    }
    const aParametro = editados((Object.keys(columnasParam(tipo)) as (keyof Ci2Asignacion)[]).filter((k) => !(CAMPOS_GLOBALES as readonly string[]).includes(k)));
    // El proveedor aprende de cada edición: se guarda lo que su fila aún no tiene (nuevo o con campos vacíos); lo que ya tiene no se pisa, ese cambio aplica solo a esta factura.
    const sugActual = (await sugerenciasPorNit(empresaId, tipo)).get(f.nit);
    const vacio = (v: unknown) => v === undefined || v === null || v === "" || v === 0 || (typeof v === "object" && !Object.keys(v as object).length);
    const aGuardar: Partial<Ci2Asignacion> = {};
    for (const k of aParametro) {
      const nuevo = (f.asignacion as any)[k];
      if (vacio(nuevo) || (k === "tratamientoIva" && nuevo === "mayor")) continue; // el IVA mayor valor es esporádico: no se aprende
      const previo = (sugActual as any)?.[k];
      if (PARAM_JSON.has(k)) {
        const faltan = Object.fromEntries(Object.entries(nuevo as Record<string, string>).filter(([t]) => !(previo && previo[t])));
        if (Object.keys(faltan).length) (aGuardar as any)[k] = { ...(previo || {}), ...faltan };
      } else if (vacio(previo)) (aGuardar as any)[k] = nuevo;
    }
    const guardoParametro = !!f.nit && Object.keys(aGuardar).length > 0;
    if (guardoParametro) await guardarParametroProveedor(empresaId, tipo, f.nit, f.proveedor, aGuardar);
    const [sugs, glob, impuestos, ctx] = await Promise.all([sugerenciasPorNit(empresaId, tipo), getCuentasGlobales(empresaId, tipo), getTabla(empresaId, "impuestos"), getContextoCi2(empresaId)]);
    const sug = sugs.get(f.nit);
    const ef = efectiva(f, sug, glob, ctx.pyme, ctx.pagos);
    const tarifasInc = tarifasIncDe(impuestos?.rows);
    const siigoPyme = ctx.sistema === "siigopyme";
    const cfgInv = ctx.inventario.activo && !f.isCreditNote ? ctx.inventario : null;
    // Si lo editado quedó como parámetro del proveedor o de la empresa, las demás facturas cambian de sugerencia: el portal las recarga.
    const afectaOtras = aGlobal.length > 0 || guardoParametro;
    res.json({ status: "ok", afectaOtras, factura: { ...f, retNoAplica: retencionesNoAplican(f.items, ef, opcionesRetencion(((impuestos?.rows || []) as any), siigoPyme)), sugerencia: sug || null, efectiva: ef, conceptoSugerido: conceptoSugeridoPyme(f.docNumber, f.items, f.proveedor), faltantes: faltantesDe(ef, f.items, tipo, tarifasInc, ctx.sistema === "worldoffice", cfgInv, siigoPyme, ctx.ivaRama, tipo === "venta" ? ctx.pagos.ventas : ctx.pagos.compras), ivaRama: ivaRamaDe(ef, f.items, ctx.ivaRama), ivaMayorSugerido: ivaMayorSugerido(ef, cfgInv), sinParametros: !sug } });
  } catch (e) {
    err(res, e);
  }
});

router.post("/empresas/:id/facturas/revision", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    const ids = Array.isArray(req.body?.ids) ? req.body.ids.map(String) : [];
    res.json({ status: "ok", modificadas: await marcarRevision(empresaId, ids, !!req.body?.revisada) });
  } catch (e) {
    err(res, e);
  }
});

router.post("/empresas/:id/facturas/estado", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    const estado = req.body?.estado;
    if (estado !== "pendiente" && estado !== "descartada") return err(res, new Error("Estado inválido."));
    const ids = Array.isArray(req.body?.ids) ? req.body.ids.map(String) : [];
    res.json({ status: "ok", modificadas: await marcarEstado(empresaId, ids, estado) });
  } catch (e) {
    err(res, e);
  }
});

/** Escribe en la fila del proveedor/cliente (paramCompras o paramVentas) los campos indicados; la crea si no existe. */
async function guardarParametroProveedor(empresaId: string, tipo: Ci2Tipo, nit: string, proveedor: string, campos: Partial<Ci2Asignacion>): Promise<void> {
  const slot = slotParam(tipo);
  const tabla = (await getTabla(empresaId, slot)) || (await iniciarTablaConPlantilla(empresaId, slot));
  const colNit = tipo === "venta" ? COLS_NIT_VENTAS.find((c) => tabla.columns.includes(c)) || COLS_NIT_VENTAS[0] : COL_NIT;
  let fila = tabla.rows.find((r) => nitDigitos(r[colNit]) === nit);
  if (!fila) {
    fila = { [colNit]: nit, Proveedor: proveedor };
    tabla.rows.push(fila);
  }
  for (const [campo, col] of Object.entries(columnasParam(tipo))) {
    if (!(campo in campos)) continue;
    if (!tabla.columns.includes(col as string)) tabla.columns.push(col as string);
    const v = (campos as any)[campo];
    fila[col as string] = PARAM_JSON.has(campo as keyof Ci2Asignacion) ? JSON.stringify(v) : (v as any);
  }
  await setTabla(empresaId, slot, tabla);
}

/** Llena en la tabla de proveedores (compras) lo que esté vacío con lo que dice el balance de prueba por tercero; lo que ya hay no se toca. */
async function aplicarBalanceAProveedores(empresaId: string, buffer: Buffer) {
  const perfiles = await perfilesDesdeBalance(buffer);
  const tabla = (await getTabla(empresaId, "paramCompras")) || (await iniciarTablaConPlantilla(empresaId, "paramCompras"));
  const pyme = (await getSistemaContable(empresaId)) === "siigopyme";
  const impuestos = ((await getTabla(empresaId, "impuestos"))?.rows || []) as Record<string, unknown>[];
  const formasCompra = (await getConfigPagos(empresaId)).compras;
  const opc = opcionesRetencion(impuestos, pyme);
  const cols = columnasParam("compra");
  const vacia = (v: unknown) => v === undefined || v === null || String(v).trim() === "" || Number(v) === 0;
  const poner = (fila: Record<string, any>, campo: keyof Ci2Asignacion, valor: string | number): boolean => {
    const col = cols[campo] as string;
    if (!vacia(fila[col])) return false;
    if (!tabla.columns.includes(col)) tabla.columns.push(col);
    fila[col] = valor;
    return true;
  };
  let creados = 0, actualizados = 0, sinTarifa = 0;
  for (const pf of perfiles) {
    const fila = tabla.rows.find((r) => nitDigitos(r[COL_NIT]) === pf.nit);
    const nueva = !fila;
    const f: Record<string, any> = fila ?? { [COL_NIT]: pf.nit, Proveedor: pf.nombre };
    let cambios = 0;
    if (pf.cuentaGasto && poner(f, "cuentaGasto", pf.cuentaGasto)) cambios++;
    // Forma de pago: la que tiene la cuenta por pagar que más mueve el tercero en el balance.
    const forma = pf.cuentaPorPagar ? formasCompra.find((x) => x.cuenta === pf.cuentaPorPagar) : undefined;
    if (forma && poner(f, "formaPagoCodigo", forma.nombre)) cambios++;
    const rets = [["retefuente", "tarifaRetefuente", "cuentaRetefuente", opc.retefuente], ["reteica", "tarifaReteica", "cuentaReteica", opc.reteica], ["reteiva", "tarifaReteiva", "cuentaReteiva", opc.reteiva]] as const;
    for (const [k, campoTarifa, campoCuenta, lista] of rets) {
      const r = pf[k];
      if (!r) continue;
      // Solo se sugiere una retención que exista en la tabla de impuestos (si no, la interfaz no podría causarla).
      const mismaTarifa = lista.filter((o) => Math.abs(o.tarifa - r.tarifa) < 0.0001);
      const op = (pyme && mismaTarifa.find((o) => "cuenta" in o && limpiarCuenta(o.cuenta) === limpiarCuenta(r.cuenta))) || mismaTarifa[0];
      if (!op) { sinTarifa++; continue; }
      if (poner(f, campoTarifa, op.tarifa)) cambios++;
      if (pyme && "cuenta" in op && op.cuenta) poner(f, campoCuenta, op.cuenta);
    }
    if (!cambios) continue;
    if (nueva) { tabla.rows.push(f); creados++; } else actualizados++;
  }
  await setTabla(empresaId, "paramCompras", tabla);
  await marcarBalanceTercero(empresaId, perfiles.length);
  return { proveedoresEnBalance: perfiles.length, creados, actualizados, retencionesSinTarifaEnTabla: sinTarifa };
}

// Sugerencias por proveedor desde el balance de prueba por tercero subido a mano.
router.post("/empresas/:id/parametros-desde-balance", upload.single("archivo"), async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    if (!req.file) return err(res, new Error("Sube el Excel «Balance de prueba por tercero» de Siigo."));
    res.json({ status: "ok", ...(await aplicarBalanceAProveedores(empresaId, req.file.buffer)) });
  } catch (e) {
    err(res, e);
  }
});

// Guarda la asignación efectiva de una factura como parámetro del proveedor (fila de paramCompras).
router.put("/empresas/:id/proveedores/:nit", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    const nit = nitDigitos(req.params.nit);
    if (!nit) return err(res, new Error("NIT inválido."));
    const campos = (req.body?.campos || {}) as Partial<Ci2Asignacion>;
    await guardarParametroProveedor(empresaId, tipoDe(req), nit, String(req.body?.proveedor || ""), campos);
    res.json({ status: "ok" });
  } catch (e) {
    err(res, e);
  }
});

// Cuentas globales de la empresa (descuento, recargo, otros impuestos).
const cfgInventarioDe = async (empresaId: string, esNc: boolean) => {
  const c = await getConfigInventario(empresaId);
  return c.activo && !esNc ? c : null;
};

// ── Inventario y gasto (compras de Siigo) ───────────────────────────────────

router.get("/empresas/:id/inventario", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    res.json({ status: "ok", inventario: await getConfigInventario(empresaId) });
  } catch (e) {
    err(res, e, 500);
  }
});

router.put("/empresas/:id/inventario", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    res.json({ status: "ok", inventario: await setConfigInventario(empresaId, req.body?.inventario) });
  } catch (e) {
    err(res, e);
  }
});

// ── IVA mayor valor por rama de gasto ───────────────────────────────────────

router.get("/empresas/:id/iva-rama", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    res.json({ status: "ok", ivaRama: await getConfigIvaRama(empresaId) });
  } catch (e) {
    err(res, e, 500);
  }
});

router.put("/empresas/:id/iva-rama", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    // Solo cuentas activas del PUC cargado: una inactiva no puede recibir el IVA, así que se excluye (y se avisa).
    const puc = await getCodigosPuc(empresaId);
    const pedidas: string[] = Array.isArray(req.body?.ivaRama?.cuentas) ? req.body.ivaRama.cuentas.map((c: unknown) => String(c ?? "").replace(/\D/g, "")).filter(Boolean) : [];
    const excluidas = puc.size ? [...new Set(pedidas.filter((c) => !puc.has(c)))] : [];
    const cuentas = pedidas.filter((c) => !excluidas.includes(c));
    res.json({ status: "ok", ivaRama: await setConfigIvaRama(empresaId, { ...req.body?.ivaRama, cuentas }), excluidas });
  } catch (e) {
    err(res, e);
  }
});

// ── Formas de pago (cuenta por pagar / por cobrar) y cuenta única de ingreso de ventas ──

router.get("/empresas/:id/formas-pago", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    res.json({ status: "ok", pagos: await getConfigPagos(empresaId) });
  } catch (e) {
    err(res, e, 500);
  }
});

router.put("/empresas/:id/formas-pago", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    res.json({ status: "ok", pagos: await setConfigPagos(empresaId, req.body?.pagos) });
  } catch (e) {
    err(res, e);
  }
});

// ── Maestro de productos de World Office ────────────────────────────────────

router.get("/empresas/:id/productos-wo", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    res.json({ status: "ok", productos: await getProductosWo(empresaId, req.query.modo === "ventas" ? "ventas" : "compras") });
  } catch (e) {
    err(res, e, 500);
  }
});

router.put("/empresas/:id/productos-wo", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    res.json({ status: "ok", productos: await setProductosWo(empresaId, req.body?.productos, req.query.modo === "ventas" ? "ventas" : "compras") });
  } catch (e) {
    err(res, e);
  }
});

router.get("/empresas/:id/cuentas-globales", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    res.json({ status: "ok", cuentas: await getCuentasGlobales(empresaId, tipoDe(req)) });
  } catch (e) {
    err(res, e, 500);
  }
});

router.put("/empresas/:id/cuentas-globales", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    const cambios: Ci2CuentasGlobales = {};
    for (const k of CAMPOS_GLOBALES) if (k in (req.body || {})) cambios[k] = limpiarCuenta(req.body[k]);
    if (req.body?.cuentasImpuestos && typeof req.body.cuentasImpuestos === "object") {
      cambios.cuentasImpuestos = Object.fromEntries(Object.entries(req.body.cuentasImpuestos).map(([n, c]) => [n, limpiarCuenta(c)]));
    }
    const tipo = tipoDe(req);
    await setCuentasGlobales(empresaId, cambios, tipo);
    res.json({ status: "ok", cuentas: await getCuentasGlobales(empresaId, tipo) });
  } catch (e) {
    err(res, e);
  }
});

// Lee el movimiento auxiliar de Siigo y devuelve el siguiente consecutivo por tipo de comprobante.
router.post("/empresas/:id/formas-pago-pyme", upload.single("archivo"), async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    if ((await getSistemaContable(empresaId)) !== "siigopyme") return err(res, new Error("Esta empresa no es de Siigo Pyme."));
    if (!req.file) return err(res, new Error("Sube el Excel de formas de pago de Siigo Pyme."));
    res.json({ status: "ok", siigoPyme: await importarFormasPagoPyme(empresaId, req.file.buffer) });
  } catch (e) {
    err(res, e);
  }
});

router.post("/empresas/:id/impuestos-desde-puc", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    if ((await getSistemaContable(empresaId)) !== "siigopyme") return err(res, new Error("Esta empresa no es de Siigo Pyme."));
    const tabla = await generarTablaImpuestosPyme(empresaId);
    res.json({ status: "ok", filas: tabla.rows.length });
  } catch (e) {
    err(res, e);
  }
});

router.post("/empresas/:id/centros-pyme", upload.single("archivo"), async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    if ((await getSistemaContable(empresaId)) !== "siigopyme") return err(res, new Error("Esta empresa no es de Siigo Pyme."));
    if (!req.file) return err(res, new Error("Sube el Excel de centros de costo de Siigo Pyme."));
    res.json({ status: "ok", siigoPyme: await importarCentrosPyme(empresaId, req.file.buffer) });
  } catch (e) {
    err(res, e);
  }
});

router.post("/empresas/:id/consecutivos/detectar", upload.single("auxiliar"), async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    const f = req.file as Express.Multer.File | undefined;
    if (!f) return err(res, new Error("Debes subir el movimiento auxiliar de Siigo."));
    const comp = await getComprobantesEmpresa(empresaId);
    // El lote puede usar más comprobantes que los dos de la empresa (documento soporte, POS…): llegan en `tipos`.
    const pedidos = String(req.body?.tipos ?? "").split(",").map((t) => t.trim()).filter(Boolean);
    const tipos = pedidos.length ? pedidos : [comp.tipoCompras, comp.tipoComprasNc].filter(Boolean);
    res.json({ status: "ok", consecutivos: await detectarConsecutivosDesdeAuxiliar(f.buffer, tipos) });
  } catch (e) {
    err(res, e, 500);
  }
});

// Tipo de comprobante recordado por clase de documento, con los de la empresa como valor inicial de FE y NC (compras).
router.get("/empresas/:id/tipos-documento", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    const tipo = tipoDe(req);
    const [guardados, comp] = await Promise.all([getTiposDocumento(empresaId, tipo), getComprobantesEmpresa(empresaId)]);
    const porDefecto: Record<string, string> = tipo === "venta" ? {} : { FE: comp.tipoCompras, NC: comp.tipoComprasNc };
    res.json({ status: "ok", tipos: { ...Object.fromEntries(Object.entries(porDefecto).filter(([, v]) => v)), ...guardados } });
  } catch (e) {
    err(res, e, 500);
  }
});

// Tipo de comprobante recordado por prefijo (ventas).
// ── Tipos de comprobante de Siigo (desplegable): se traen de la web de Siigo con el mismo inicio de sesión ─────────
// Las credenciales llegan en el cuerpo, se usan una vez y no se guardan ni se registran.
function credencialesSiigo(req: Request): { email: string; password: string } | null {
  const email = typeof req.body?.email === "string" ? req.body.email.trim() : "";
  const password = typeof req.body?.password === "string" ? req.body.password : "";
  return email && password ? { email, password } : null;
}
function errorSiigoWeb(res: Response, e: unknown) {
  if (e instanceof SiigoWebError) return res.status(e.code === "bad_credentials" ? 401 : 422).json({ status: "error", codigo: e.code, mensaje: e.message });
  console.error("[CI2] Siigo web falló:", e instanceof Error ? e.message : "error");
  return res.status(500).json({ status: "error", mensaje: "No se pudo conectar con Siigo. Intenta de nuevo." });
}

// Al crear la empresa (aún sin id): se trae el listado por NIT y el cliente lo guarda con la empresa.
router.post("/siigo/comprobantes", rateLimit(6, 15 * 60_000), async (req: Request, res: Response) => {
  const cred = credencialesSiigo(req);
  const nit = nitDigitos(req.body?.nit);
  if (!cred) return err(res, new Error("Escribe el correo y la contraseña de Siigo."));
  if (!nit) return err(res, new Error("Escribe primero el NIT de la empresa."));
  try {
    res.json({ status: "ok", items: await fetchSiigoComprobantes(cred.email, cred.password, { nit, name: String(req.body?.nombre || "") }) });
  } catch (e) { errorSiigoWeb(res, e); }
});

router.get("/empresas/:id/comprobantes-siigo", async (req: Request, res: Response) => {
  const empresaId = await exigirEmpresa(req, res);
  if (!empresaId) return;
  res.json({ status: "ok", ...(await getComprobantesSiigo(empresaId)) });
});

router.put("/empresas/:id/comprobantes-siigo", async (req: Request, res: Response) => {
  const empresaId = await exigirEmpresa(req, res);
  if (!empresaId) return;
  try { res.json({ status: "ok", items: await setComprobantesSiigo(empresaId, req.body?.items) }); } catch (e) { err(res, e); }
});

// ── Inicio de sesión de Siigo guardado (cifrado) por empresa ────────────────────────────────────────────────
// Solo el dueño (o un admin) puede guardarlo u olvidarlo; cualquiera con acceso a la empresa puede usarlo para traer
// la parametrización, pero nunca verlo: la contraseña no sale del servidor.
async function puedeGestionarLogin(req: Request, empresaId: string): Promise<boolean> {
  return !!req.user!.isAdmin || (await esDuenoEmpresa(empresaId, req.user!.userId));
}

router.get("/empresas/:id/siigo-login", async (req: Request, res: Response) => {
  const empresaId = await exigirEmpresa(req, res);
  if (!empresaId) return;
  res.json({ status: "ok", ...(await metaLoginSiigo(empresaId)) });
});

// Se usa al crear la empresa: el login ya se verificó al traer la lista y se guarda con la empresa recién creada.
router.put("/empresas/:id/siigo-login", async (req: Request, res: Response) => {
  const empresaId = await exigirEmpresa(req, res);
  if (!empresaId) return;
  if (!(await puedeGestionarLogin(req, empresaId))) return err(res, new Error("Solo el dueño de la empresa puede guardar el inicio de sesión."), 403);
  const cred = credencialesSiigo(req);
  if (!cred) return err(res, new Error("Escribe el correo y la contraseña de Siigo."));
  try { await guardarLoginSiigo(empresaId, cred.email, cred.password, req.user!.userId); res.json({ status: "ok" }); } catch (e) { err(res, e); }
});

router.delete("/empresas/:id/siigo-login", async (req: Request, res: Response) => {
  const empresaId = await exigirEmpresa(req, res);
  if (!empresaId) return;
  if (!(await puedeGestionarLogin(req, empresaId))) return err(res, new Error("Solo el dueño de la empresa puede olvidar el inicio de sesión."), 403);
  await olvidarLoginSiigo(empresaId);
  res.json({ status: "ok" });
});

/** Credenciales escritas o, si no hay, las guardadas. `guardadas` indica que se usaron las guardadas. */
async function credencialesParaEmpresa(req: Request, empresaId: string): Promise<{ email: string; password: string; guardadas: boolean } | null> {
  const typed = credencialesSiigo(req);
  if (typed) return { ...typed, guardadas: false };
  const saved = await leerLoginSiigo(empresaId);
  return saved ? { ...saved, guardadas: true } : null;
}
function errorSiigoEmpresa(res: Response, e: unknown, guardadas: boolean) {
  if (guardadas && e instanceof SiigoWebError && e.code === "bad_credentials") {
    return res.status(401).json({ status: "error", codigo: "login_guardado_invalido", mensaje: "La contraseña guardada ya no funciona en Siigo. Escríbela de nuevo para actualizarla." });
  }
  return errorSiigoWeb(res, e);
}
/** Guarda el login escrito si el usuario lo pidió y puede hacerlo (solo tras iniciar sesión bien en Siigo). */
async function guardarSiAplica(req: Request, empresaId: string, cred: { email: string; password: string; guardadas: boolean }): Promise<boolean> {
  if (cred.guardadas || req.body?.recordar !== true || !(await puedeGestionarLogin(req, empresaId))) return false;
  await guardarLoginSiigo(empresaId, cred.email, cred.password, req.user!.userId);
  return true;
}

// Empresa ya creada: actualiza solo el listado de tipos de comprobante, con el NIT de la empresa.
router.post("/empresas/:id/comprobantes-siigo/traer", rateLimit(6, 15 * 60_000), async (req: Request, res: Response) => {
  const empresaId = await exigirEmpresa(req, res);
  if (!empresaId) return;
  const cred = await credencialesParaEmpresa(req, empresaId);
  if (!cred) return res.status(400).json({ status: "error", codigo: "sin_login_guardado", mensaje: "Escribe el correo y la contraseña de Siigo." });
  try {
    const items = await fetchSiigoComprobantes(cred.email, cred.password, { nit: await getEmpresaNit(empresaId) });
    const guardado = await guardarSiAplica(req, empresaId, cred);
    res.json({ status: "ok", items: await setComprobantesSiigo(empresaId, items), loginGuardado: guardado });
  } catch (e) { errorSiigoEmpresa(res, e, cred.guardadas); }
});

// Parametrización completa desde Siigo: tipos de comprobante + impuestos + plan de cuentas (PUC), en una sola sesión.
router.post("/empresas/:id/parametrizacion-siigo/traer", rateLimit(6, 15 * 60_000), async (req: Request, res: Response) => {
  const empresaId = await exigirEmpresa(req, res);
  if (!empresaId) return;
  const cred = await credencialesParaEmpresa(req, empresaId);
  if (!cred) return res.status(400).json({ status: "error", codigo: "sin_login_guardado", mensaje: "Escribe el correo y la contraseña de Siigo." });
  try {
    const par = await fetchSiigoParametrizacion(cred.email, cred.password, { nit: await getEmpresaNit(empresaId) }, { balance: true });
    const guardado = await guardarSiAplica(req, empresaId, cred);
    const resultado: Record<string, unknown> = { loginGuardado: guardado };
    const errores: string[] = [];
    if (par.comprobantes) resultado.comprobantes = (await setComprobantesSiigo(empresaId, par.comprobantes)).length;
    else if (par.comprobantesError) errores.push(`Tipos de comprobante: ${par.comprobantesError}`);
    if (par.impuestos) {
      try { resultado.impuestos = (await setTabla(empresaId, "impuestos", tablaImpuestosDeSiigo(par.impuestos))).rows.length; }
      catch (e) { errores.push(`Impuestos: ${e instanceof Error ? e.message : "error"}`); }
    } else if (par.impuestosError) errores.push(`Impuestos: ${par.impuestosError}`);
    // El balance se aplica después de los impuestos: solo se sugieren retenciones que existan en la tabla.
    if (par.balanceBuffer) {
      try { resultado.balance = await aplicarBalanceAProveedores(empresaId, par.balanceBuffer); }
      catch (e) { errores.push(`Balance de prueba: ${e instanceof Error ? e.message : "error"}`); }
    } else if (par.balanceError) errores.push(`Balance de prueba: ${par.balanceError}`);
    if (par.pucBuffer) {
      try { resultado.puc = (await importarPuc(empresaId, par.pucBuffer, "Cuentas contables (Siigo).xlsx")).length; }
      catch (e) { errores.push(`Plan de cuentas: ${e instanceof Error ? e.message : "error"}`); }
    } else if (par.pucError) errores.push(`Plan de cuentas: ${par.pucError}`);
    res.json({ status: "ok", resultado, errores });
  } catch (e) { errorSiigoEmpresa(res, e, cred.guardadas); }
});

router.get("/empresas/:id/prefijos", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    res.json({ status: "ok", prefijos: await getPrefijos(empresaId) });
  } catch (e) {
    err(res, e, 500);
  }
});

/** Archivo de terceros (plantilla .xlsm de Siigo) con los proveedores (compras) o clientes (ventas) de las facturas, sin repetir. El consumidor final no se crea. */
async function armarArchivoTerceros(empresaId: string, ids: string[], tipo: Ci2Tipo): Promise<{ data: Buffer | null; terceros: number; sinDatos: number }> {
  const venta = tipo === "venta";
  const base = await facturasConArchivos(empresaId, ids);
  const filas = new Map<string, FilaTercero>();
  const soloNombre = new Set<string>();
  for (const f of base) {
    const nit = String(f.nit || "").replace(/\D/g, "");
    if (!nit || /^2{6,}$/.test(nit) || filas.has(nit)) continue;
    let fila: FilaTercero | null = null;
    if (f.xml) {
      const d = await extractInvoiceDataFromXml(f.xml, { id: f.id, docnum: f.docNumber });
      fila = venta
        ? filaTerceroDesdeFactura({ nit: d.receiverNit || nit, nombre: d.receiverName && d.receiverName !== "N/A" ? d.receiverName : f.proveedor, nombres: d.receiverFirstNames, apellidos: d.receiverFamilyNames, direccion: d.receiverAddress, codPais: d.receiverCountryCode, codDepartamento: d.receiverStateCode, codCiudad: d.receiverCityCode, telefono: d.receiverPhone, correo: d.receiverEmail })
        : filaTerceroDesdeFactura({ nit: d.issuerNit || nit, nombre: d.issuerName && d.issuerName !== "N/A" ? d.issuerName : f.proveedor, nombres: d.issuerFirstNames, apellidos: d.issuerFamilyNames, direccion: d.issuerAddress, codPais: d.issuerCountryCode, codDepartamento: d.issuerStateCode, codCiudad: d.issuerCityCode, telefono: d.issuerPhone, correo: d.issuerEmail });
    }
    if (!fila) { fila = filaTerceroDesdeFactura({ nit, nombre: f.proveedor }); soloNombre.add(nit); }
    if (fila) filas.set(nit, fila);
  }
  if (!filas.size) return { data: null, terceros: 0, sinDatos: 0 };
  return { data: await escribirTercerosXlsm(Buffer.from(PLANTILLA_TERCEROS_BASE64, "base64"), [...filas.values()]), terceros: filas.size, sinDatos: [...soloNombre].filter((n) => filas.has(n)).length };
}

// ── Descargas de las facturas traídas: Excel detallado (el de la descarga masiva + Excel) y ZIP con los PDF y XML ──

/** Una factura del XML como fila del Excel detallado: los mismos campos y valores por defecto de la descarga masiva + Excel. */
function filaExcelDetallado(d: Partial<InvoiceData>, f: { id: string; docNumber: string; nit: string }): InvoiceData {
  const na = (v: unknown) => (v === undefined || v === null || v === "" ? "N/A" : String(v));
  return {
    issuerNit: na(d.issuerNit || f.nit), issuerName: na(d.issuerName), issuerEmail: na(d.issuerEmail), issuerPhone: na(d.issuerPhone), issuerAddress: na(d.issuerAddress),
    issuerCity: na(d.issuerCity), issuerDepartment: na(d.issuerDepartment), issuerCountry: na(d.issuerCountry), issuerCommercialName: na(d.issuerCommercialName),
    issuerTaxpayerType: na(d.issuerTaxpayerType), issuerFiscalRegime: na(d.issuerFiscalRegime), issuerTaxResponsibility: na(d.issuerTaxResponsibility), issuerEconomicActivity: na(d.issuerEconomicActivity),
    receiverNit: na(d.receiverNit), receiverName: na(d.receiverName), receiverEmail: na(d.receiverEmail), receiverPhone: na(d.receiverPhone), receiverAddress: na(d.receiverAddress),
    receiverCity: na(d.receiverCity), receiverDepartment: na(d.receiverDepartment), receiverCountry: na(d.receiverCountry), receiverCommercialName: na(d.receiverCommercialName),
    receiverTaxpayerType: na(d.receiverTaxpayerType), receiverFiscalRegime: na(d.receiverFiscalRegime), receiverTaxResponsibility: na(d.receiverTaxResponsibility), receiverEconomicActivity: na(d.receiverEconomicActivity),
    issueDate: d.issueDate || "N/A", issueDateISO: d.issueDateISO || "9999-12-31", paymentMethod: d.paymentMethod || "N/A",
    subtotal: d.subtotal || 0, iva: d.iva || 0, total: d.total || 0, taxes: d.taxes || [], discount: d.discount || 0, surcharge: d.surcharge || 0,
    concepts: d.concepts || "N/A", lineItems: d.lineItems || [], documentType: d.documentType || "Factura Electrónica", isDocumentoSoporte: d.isDocumentoSoporte || false,
    cufe: d.cufe || "N/A", notes: d.notes || "", trackId: f.id, docNumber: d.docNumber || f.docNumber, zipFilename: `${d.issuerNit || f.nit} - ${f.docNumber}.zip`,
  } as InvoiceData;
}

const idsDeCuerpo = (req: Request): string[] => (Array.isArray(req.body?.ids) ? req.body.ids.map(String).slice(0, 2000) : []);

router.post("/empresas/:id/facturas/excel", async (req: Request, res: Response) => {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "ci2-excel-"));
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    const ids = idsDeCuerpo(req);
    if (!ids.length) return err(res, new Error("No hay facturas para el Excel."));
    const base = (await facturasConArchivos(empresaId, ids)).filter((f) => f.xml);
    if (!base.length) return err(res, new Error("Estas facturas no tienen el XML guardado. Vuelve a traerlas de la DIAN para generar el Excel."));
    const facturas: InvoiceData[] = [];
    for (const f of base) facturas.push(filaExcelDetallado(await extractInvoiceDataFromXml(f.xml!, { id: f.id, docnum: f.docNumber }), f));
    const empresa = (await listEmpresasForUser(req.user!.userId, req.user!.isAdmin)).find((e) => e.id === empresaId);
    const ruta = path.join(tmp, "facturas.xlsx");
    await generateExcelFile(facturas, ruta, false, false, empresa?.nombre || "", empresa?.nit || "");
    const fechas = base.map((f) => f.fecha).filter(Boolean).sort();
    const nombre = generateExcelFilename(fechas[0], fechas[fechas.length - 1], "Facturas DIAN");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${nombre.replace(/[^\w.\- ]/g, "_")}"`);
    res.setHeader("X-Facturas-Sin-Xml", String(ids.length - base.length));
    res.send(fs.readFileSync(ruta));
  } catch (e) {
    err(res, e, 500);
  } finally {
    fs.rmSync(tmp, { recursive: true, force: true });
  }
});

router.post("/empresas/:id/facturas/zip", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    const ids = idsDeCuerpo(req);
    if (!ids.length) return err(res, new Error("No hay facturas para descargar."));
    const base = await facturasConArchivos(empresaId, ids);
    const zip = new JSZip();
    let n = 0;
    for (const f of base) {
      const nombre = `${String(f.proveedor || f.nit).replace(/[^\w ]/g, "").trim().slice(0, 30)} ${String(f.docNumber).replace(/[^\w-]/g, "")}`.trim();
      if (f.xml) { zip.file(`XML/${nombre}.xml`, f.xml); n++; }
      if (f.pdf) { zip.file(`PDF/${nombre}.pdf`, f.pdf); n++; }
    }
    if (!n) return err(res, new Error("Estas facturas no tienen PDF ni XML guardados. Vuelve a traerlas de la DIAN."));
    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", 'attachment; filename="Facturas_PDF_XML.zip"');
    res.send(await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  } catch (e) {
    err(res, e, 500);
  }
});

// ── Generar interfaz ────────────────────────────────────────────────────────

router.post("/empresas/:id/generar", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    const ids: string[] = Array.isArray(req.body?.ids) ? req.body.ids.map(String) : [];
    if (!ids.length) return err(res, new Error("Selecciona al menos una factura."));

    const tipo = tipoDe(req);
    const venta = tipo === "venta";
    const [facturas, sugs, impuestos, comp, empresas, glob, prefijosGuardados, tiposGuardados, autorretencion] = await Promise.all([
      facturasPorIds(empresaId, ids),
      sugerenciasPorNit(empresaId, tipo),
      getTabla(empresaId, "impuestos"),
      getComprobantesEmpresa(empresaId),
      listEmpresasForUser(req.user!.userId, req.user!.isAdmin),
      getCuentasGlobales(empresaId, tipo),
      getPrefijos(empresaId),
      getTiposDocumento(empresaId, tipo),
      getAutorretencion(empresaId),
    ]);
    const sistemaContable = await getSistemaContable(empresaId);
    const cfgInventario = await getConfigInventario(empresaId);
    const cfgIvaRama = await getConfigIvaRama(empresaId);
    const cfgPagos = await getConfigPagos(empresaId);
    const worldoffice = sistemaContable === "worldoffice";
    const siigoPyme = sistemaContable === "siigopyme";
    const vencimientos = !worldoffice && !siigoPyme ? await getCuentasConVencimiento(empresaId) : undefined;
    const configPyme = siigoPyme ? await getConfigSiigoPyme(empresaId) : null;
    const pymeCfg = await pymeDe(empresaId);
    if (siigoPyme && venta) return err(res, new Error("Siigo Pyme solo está habilitado para compras por ahora."));
    const configWo = worldoffice ? await getConfigWorldOffice(empresaId) : null;
    if (worldoffice) {
      if (!configWo!.empresa || !configWo!.terceroInterno || !configWo!.bodega) return err(res, new Error("Faltan los parámetros de World Office de la empresa (nombre en World Office, Tercero Interno y Bodega)."));
      if (venta && !configWo!.tipoVentas) return err(res, new Error("Falta el Tipo Documento de ventas de World Office (Parámetros → datos de la empresa en World Office)."));
    } else if (!impuestos?.rows?.length) return err(res, new Error("La empresa no tiene tabla de impuestos cargada (Configuración)."));
    const pendientes = facturas.filter((f) => f.estado === "pendiente" && f.tipo === tipo);
    if (!pendientes.length) return err(res, new Error("Ninguna de las facturas seleccionadas está pendiente."));
    // En compras solo pasan a la interfaz las facturas con la revisión oficial del usuario.
    const sinRevisar = venta ? [] : pendientes.filter((f) => !f.revisada); // en ventas no se exige
    if (sinRevisar.length) return err(res, new Error(`${sinRevisar.length} factura(s) no tienen la revisión oficial (casilla «Revisada»): ${sinRevisar.slice(0, 5).map((f) => f.docNumber).join(", ")}${sinRevisar.length > 5 ? "…" : ""}.`));

    // Consecutivo inicial: lo recordado en la empresa, o el que el usuario fije para este lote.
    const recordados = empresas.find((e) => e.id === empresaId)?.consecutivos || {};
    const consecutivoInicial: Record<string, number> = { ...recordados };
    for (const [t, n] of Object.entries((req.body?.consecutivos || {}) as Record<string, unknown>)) {
      const v = Math.trunc(Number(n));
      if (Number.isFinite(v) && v > 0) consecutivoInicial[t] = v;
    }

    // Ventas: tipo de comprobante por prefijo (lo guardado, más lo que el usuario indique en este lote).
    const prefijos = {
      factura: { ...prefijosGuardados.factura, ...limpiarMapa(req.body?.prefijos?.factura) },
      notaCredito: { ...prefijosGuardados.notaCredito, ...limpiarMapa(req.body?.prefijos?.notaCredito) },
    };

    // Tipo de comprobante por clase de documento (lo guardado, más lo que el usuario indique en este lote).
    const tipoPorDocumento = { ...tiposGuardados, ...limpiarMapa(req.body?.tiposDocumento) };

    const entradas: Ci2FacturaInput[] = pendientes.map((f) => ({
      id: f.id,
      fecha: fechaSiigo(f.fecha),
      documentType: f.documentType,
      isCreditNote: f.isCreditNote,
      docNumber: f.docNumber,
      nit: f.nit,
      proveedor: f.proveedor,
      items: f.items,
      totalIva: f.totals.iva,
      formaPago: f.formaPago,
      fechaVencimiento: f.fechaVencimiento,
      asignacion: efectiva(f, sugs.get(f.nit), glob, pymeCfg, cfgPagos),
    }));

    const r = generarInterfaz(entradas, {
      modo: tipo,
      descripcionConcepto: !worldoffice && !siigoPyme,
      inventario: !venta && !worldoffice && cfgInventario.activo ? cfgInventario : undefined,
      ivaRama: !venta && !worldoffice && cfgIvaRama.cuentas.length ? cfgIvaRama : undefined,
      vencimientos,
      worldoffice,
      pyme: siigoPyme ? { diasVencimiento: configPyme!.diasVencimiento, cuentaIvaAsumido: configPyme!.cuentaIvaAsumido, formasPago: configPyme!.formasPago } : undefined,
      productosWo: worldoffice ? await getProductosWo(empresaId, venta ? "ventas" : "compras") : undefined,
      tipoVentasWo: worldoffice && venta ? configWo!.tipoVentas : undefined,
      productosEspecialesWo: worldoffice ? (venta ? { otros: configWo!.productoOtrosImpuestosVentas, descuento: configWo!.productoDescuentoVentas, recargo: configWo!.productoRecargoVentas } : { otros: configWo!.productoOtrosImpuestos, descuento: configWo!.productoDescuento, recargo: configWo!.productoRecargo }) : undefined,
      tipoPorPrefijo: prefijos,
      tipoPorDocumento,
      autorretencion: venta ? autorretencion : undefined,
      impuestos: (impuestos?.rows || []) as Record<string, unknown>[],
      tipoCompras: comp.tipoCompras,
      tipoComprasNc: comp.tipoComprasNc,
      consecutivoInicial,
    });
    // Todo o nada: si alguna factura tiene error no se genera ninguna interfaz ni se consume consecutivo.
    if (r.errores.length) {
      return res.status(422).json({
        status: "error",
        mensaje: `${r.errores.length} factura(s) tienen errores. No se generó la interfaz: corrígelas o quítalas de la selección y vuelve a intentar.`,
        errores: r.errores,
      });
    }

    const sufijo = venta ? "Ventas" : "Compras";
    // Mismas líneas contables; cambia solo el archivo según el programa contable de la empresa.
    const syscafe = sistemaContable === "syscafe";
    const sistema = syscafe ? "Syscafe" : siigoPyme ? "SiigoPyme" : "Siigo";
    const ext = syscafe ? "xls" : "xlsx";
    const escribir = async (filas: typeof r.salida) => (syscafe ? escribirInterfazSyscafe(filas) : await escribirInterfazSiigo(filas));
    let archivos: { nombre: string; tipo: "interfaz" | "informe" | "terceros"; data: Buffer }[];
    if (worldoffice) {
      // World Office carga dos archivos: encabezados de los documentos y su contenido.
      const wo = escribirInterfazWo(r.wo || [], configWo!, new Date(), venta);
      archivos = [
        { nombre: `Encabezados_WorldOffice_${sufijo}.xls`, tipo: "interfaz", data: wo.encabezados },
        { nombre: `Movimientos_WorldOffice_${sufijo}.xls`, tipo: "interfaz", data: wo.movimientos },
      ];
    } else if (siigoPyme) {
      const empresaNombre = empresas.find((e) => e.id === empresaId)?.nombre || "";
      const cuentasConCentro = await getCuentasConCentroPyme(empresaId);
      archivos = [
        { nombre: `MovimientoContable_SiigoPyme_${sufijo}.xlsx`, tipo: "interfaz", data: escribirInterfazSiigoPyme(r.salida, { empresa: empresaNombre, vendedor: configPyme!.vendedor, cuentasConCentro: cuentasConCentro.size ? cuentasConCentro : undefined }) },
        { nombre: `Informe_Retenciones_${sufijo}.xlsx`, tipo: "informe", data: await escribirInformeRetenciones(r.informe) },
      ];
    } else {
      archivos = [
        { nombre: `Salida_${sistema}_${sufijo}.${ext}`, tipo: "interfaz", data: await escribir(r.salida) },
        { nombre: `Informe_Retenciones_${sufijo}.xlsx`, tipo: "informe", data: await escribirInformeRetenciones(r.informe) },
      ];
      if (r.partes.length > 1) {
        for (let i = 0; i < r.partes.length; i++) {
          archivos.push({ nombre: `Salida_${sistema}_${sufijo}_Parte_${String(i + 1).padStart(3, "0")}.${ext}`, tipo: "interfaz", data: await escribir(r.partes[i]) });
        }
      }
    }

    // Terceros de las facturas de esta interfaz (proveedores en compras, clientes en ventas), en la plantilla de importación de Siigo: se crean primero y luego se sube la de comprobantes.
    if (sistemaContable === "siigo") {
      try {
        const t = await armarArchivoTerceros(empresaId, r.asignados.map((a) => a.id), tipo);
        if (t.data) archivos.push({ nombre: `Terceros_Siigo_${sufijo}.xlsm`, tipo: "terceros", data: t.data });
        if (t.sinDatos) r.advertencias.push(`${t.sinDatos} tercero(s) van solo con NIT y nombre (sin dirección ni ciudad): sus facturas no tienen el XML guardado.`);
      } catch (e) {
        r.advertencias.push(`No se pudo armar el archivo de terceros: ${e instanceof Error ? e.message : "error"}`);
      }
    }

    const usados = Object.fromEntries(Object.keys(r.siguientes).map((t) => [t, consecutivoInicial[t] ?? 1]));
    const interfazId = await guardarInterfaz(
      { empresaId, userId: req.user!.userId, tipo, createdAt: new Date(), resumen: r.resumen, facturaIds: r.asignados.map((a) => a.id), inicial: usados, siguientes: r.siguientes, advertencias: r.advertencias },
      archivos,
      r.asignados
    );
    await setTiposDocumento(empresaId, tipo, limpiarMapa(req.body?.tiposDocumento));
    if (venta) await setPrefijos(empresaId, prefijos);
    else await guardarConsecutivos(empresaId, r.siguientes);

    res.json({
      status: "ok",
      interfazId,
      resumen: r.resumen,
      advertencias: r.advertencias,
      errores: r.errores,
      asignados: r.asignados,
      siguientes: r.siguientes,
      archivos: archivos.map((a) => ({ nombre: a.nombre, tipo: a.tipo, url: `/api/interfaces-integral-odonto/interfaces/${interfazId}/archivos/${encodeURIComponent(a.nombre)}` })),
    });
  } catch (e) {
    err(res, e, 500);
  }
});

// ── Historial ───────────────────────────────────────────────────────────────

router.get("/empresas/:id/interfaces", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    const lista = await listarInterfaces(empresaId);
    res.json({
      status: "ok",
      interfaces: lista.map((i) => ({
        ...i,
        archivos: i.archivos.map((a) => ({ ...a, url: `/api/interfaces-integral-odonto/interfaces/${i.id}/archivos/${encodeURIComponent(a.nombre)}` })),
      })),
    });
  } catch (e) {
    err(res, e, 500);
  }
});

async function interfazAutorizada(req: Request, res: Response) {
  const i = await getInterfaz(req.params.iid);
  if (!i) { res.status(404).json({ status: "error", mensaje: "Interfaz no encontrada." }); return null; }
  if (!(await userCanAccessEmpresa(i.empresaId, req.user!.userId, req.user!.isAdmin))) {
    res.status(403).json({ status: "error", mensaje: "No tienes acceso a esta empresa." });
    return null;
  }
  try { await exigirEmpresaHabilitada(req.user!.userId, !!req.user!.isAdmin, i.empresaId); } catch (e) { err(res, e, 403); return null; }
  return i;
}

// Archivo de terceros de una interfaz ya creada (por si se generó antes de existir este archivo o se necesita de nuevo).
router.post("/interfaces/:iid/terceros", async (req: Request, res: Response) => {
  try {
    const i = await interfazAutorizada(req, res);
    if (!i) return;
    const t = await armarArchivoTerceros(i.empresaId, i.facturaIds, i.tipo === "venta" ? "venta" : "compra");
    if (!t.data) return err(res, new Error("Esta interfaz no tiene terceros para crear."));
    res.setHeader("Content-Type", "application/vnd.ms-excel.sheet.macroEnabled.12");
    res.setHeader("Content-Disposition", `attachment; filename="Terceros_Siigo_${i.tipo === "venta" ? "Ventas" : "Compras"}.xlsm"`);
    res.setHeader("X-Terceros", String(t.terceros));
    res.send(t.data);
  } catch (e) {
    err(res, e, 500);
  }
});

router.get("/interfaces/:iid/archivos/:nombre", async (req: Request, res: Response) => {
  try {
    if (!(await interfazAutorizada(req, res))) return;
    const a = await getArchivoInterfaz(req.params.iid, req.params.nombre);
    if (!a) return res.status(404).json({ status: "error", mensaje: "Archivo no encontrado." });
    res.setHeader("Content-Type", a.nombre.endsWith(".xls") ? "application/vnd.ms-excel" : a.nombre.endsWith(".xlsm") ? "application/vnd.ms-excel.sheet.macroEnabled.12" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${a.nombre.replace(/[^\w.\-]/g, "_")}"`);
    res.send(a.data);
  } catch (e) {
    err(res, e, 500);
  }
});

// Todos los documentos de la interfaz (interfaz Siigo + informe de retenciones) en un solo ZIP.
router.get("/interfaces/:iid/zip", async (req: Request, res: Response) => {
  try {
    const i = await interfazAutorizada(req, res);
    if (!i) return;
    const zip = new JSZip();
    for (const a of i.archivos) {
      const doc = await getArchivoInterfaz(i.id, a.nombre);
      if (doc) zip.file(doc.nombre, doc.data);
    }
    const nombre = `Interfaz_${i.tipo === "venta" ? "Ventas" : "Compras"}_${new Date(i.createdAt).toISOString().slice(0, 10)}.zip`;
    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", `attachment; filename="${nombre}"`);
    res.send(await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  } catch (e) {
    err(res, e, 500);
  }
});

// Revierte una interfaz: libera sus facturas. Si fue la última de su tipo, devuelve también el consecutivo.
router.post("/interfaces/:iid/revertir", async (req: Request, res: Response) => {
  try {
    const i = await interfazAutorizada(req, res);
    if (!i) return;
    if (i.revertida) return err(res, new Error("La interfaz ya fue revertida."));
    const liberadas = await revertirInterfaz(i.id);
    const empresas = await listEmpresasForUser(req.user!.userId, req.user!.isAdmin);
    const actuales = empresas.find((e) => e.id === i.empresaId)?.consecutivos || {};
    const devueltos: Record<string, number> = {};
    for (const [tipo, sig] of Object.entries(i.siguientes)) {
      if (actuales[tipo] === sig && i.inicial[tipo] !== undefined) devueltos[tipo] = i.inicial[tipo];
    }
    if (Object.keys(devueltos).length) await guardarConsecutivos(i.empresaId, devueltos);
    res.json({ status: "ok", liberadas, consecutivosDevueltos: devueltos });
  } catch (e) {
    err(res, e);
  }
});

export default router;
