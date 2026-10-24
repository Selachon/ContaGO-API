/**
 * Causación Interfaces 2.0.
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
 *   POST /api/causacion-interfaces-v2/empresas/:id/traer              → {jobId}  (token + mes)
 *   POST /api/causacion-interfaces-v2/empresas/:id/traer-ventas       → {jobId}  (token + listado de emitidos)
 *   GET  /api/causacion-interfaces-v2/traer/:jobId                    → estado del job
 *   GET  /api/causacion-interfaces-v2/empresas/:id/facturas           → lista con sugerencias
 *   PATCH /api/causacion-interfaces-v2/empresas/:id/facturas/:fid     → edición manual
 *   POST /api/causacion-interfaces-v2/empresas/:id/facturas/estado    → descartar / restaurar
 *   PUT  /api/causacion-interfaces-v2/empresas/:id/proveedores/:nit   → guardar como parámetro del proveedor
 *   POST /api/causacion-interfaces-v2/empresas/:id/generar            → crea la interfaz
 *   GET  /api/causacion-interfaces-v2/empresas/:id/interfaces         → historial
 *   GET  /api/causacion-interfaces-v2/interfaces/:iid/archivos/:n     → descarga
 *   GET  /api/causacion-interfaces-v2/interfaces/:iid/zip             → todos los documentos en un ZIP
 *   POST /api/causacion-interfaces-v2/interfaces/:iid/revertir        → libera las facturas
 */
import { Router, type NextFunction, type Request, type Response } from "express";
import multer from "multer";
import { v4 as uuidv4 } from "uuid";
import { requireToolAccess } from "../middleware/requireToolAccess.js";
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
  getConfigInventario,
  type ConfigInventario,
  setConfigInventario,
  setProductosWo,
  getEmpresaNit,
  getTabla,
  guardarConsecutivos,
  iniciarTablaConPlantilla,
  listEmpresasForUser,
  setTabla,
  userCanAccessEmpresa,
} from "../services/contabilizacionEmpresasService.js";
import JSZip from "jszip";
import { categoriaDocumento, cuentaBaseInventario, cuentaIvaMayor, descuentoNeto, generarInterfaz, nombresImpuestosExtra, type Ci2Asignacion, type Ci2FacturaInput } from "../services/ci2Motor.js";
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
  listarFacturas,
  listarInterfaces,
  marcarEstado,
  revertirInterfaz,
  type Ci2Factura,
} from "../services/ci2Store.js";
import { escribirInformeRetenciones, escribirInterfazSiigo } from "../services/causacionInterfaces/excelWriter.js";
import { escribirInterfazWo } from "../services/causacionInterfaces/worldOfficeWriter.js";
import { escribirInterfazSyscafe } from "../services/causacionInterfaces/syscafeWriter.js";
import { limpiarCuenta, toNumero } from "../services/causacionInterfaces/normalize.js";
import { tryExtractDianUrlFromSafeLink } from "../middleware/validateDianUrl.js";

export const CI2_TOOL_ID = "causacion-interfaces-v2";

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 40 * 1024 * 1024, files: 1 } });

const router = Router();
router.use(requireAuth);
router.use(requireToolAccess(CI2_TOOL_ID));
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
};
const COLUMNA_PARAM_VENTAS: Partial<Record<keyof Ci2Asignacion, string>> = {
  ...COLUMNA_PARAM,
  cuentaGasto: "Cuenta_ingreso",
  cuentaPorPagar: "Cuenta_por_cobrar",
  cuentaIngresoObsequios: "Cuenta_devolucion_ingreso",
};
const columnasParam = (tipo: Ci2Tipo) => (tipo === "venta" ? COLUMNA_PARAM_VENTAS : COLUMNA_PARAM);
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
function opcionesRetencion(rows: Record<string, unknown>[]) {
  const por = (tipo: string) => {
    const vistos = new Set<number>();
    const out: { tarifa: number; nombre: string; baseMinima: number }[] = [];
    for (const r of rows) {
      if (!String(r["Tipo de impuesto"] ?? "").toUpperCase().includes(tipo)) continue;
      const tarifa = toNumero(r["Tarifa"]);
      if (!(tarifa > 0) || vistos.has(tarifa)) continue;
      vistos.add(tarifa);
      out.push({ tarifa, nombre: String(r["Nombre"] ?? "").trim(), baseMinima: toNumero(r["Base"]) });
    }
    return out.sort((a, b) => a.tarifa - b.tarifa);
  };
  return { retefuente: por("RETEFUENTE"), reteica: por("RETEICA"), reteiva: por("RETEIVA") };
}

/** Sugerencia del proveedor + lo editado en la factura; las cuentas globales de la empresa rellenan lo que siga vacío. */
function efectiva(f: Ci2Factura, sug?: Ci2Asignacion, glob: Ci2CuentasGlobales = {}): Ci2Asignacion {
  const a: Ci2Asignacion = { ...ASIG_VACIA, ...(sug || {}), ...f.asignacion };
  a.retForzadas = (["tarifaRetefuente", "tarifaReteica", "tarifaReteiva"] as const).filter((k) => k in f.asignacion && Number((f.asignacion as any)[k]) > 0);
  for (const k of CAMPOS_GLOBALES) if (!limpiarCuenta(a[k])) a[k] = glob[k] || "";
  a.cuentasImpuestos = { ...(glob.cuentasImpuestos || {}) };
  return a;
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

function faltantesDe(a: Ci2Asignacion, items: Ci2Factura["items"], tipo: Ci2Tipo = "compra", tarifasInc: Set<number> = new Set(), worldOffice = false, cfgInv: ConfigInventario | null = null): string[] {
  const clasifica = !!cfgInv;
  const out: string[] = [];
  if (clasifica && tipo === "compra" && a.destino !== "inventario" && a.destino !== "gasto") out.push("destino");
  // World Office: el producto sale de la tarifa de IVA (maestro de productos), no de la cuenta del proveedor.
  if (worldOffice) return out;
  // La cuenta del proveedor solo hace falta si algún ítem con base no tiene cuenta propia.
  const sinCuentaPropia = !(clasifica && a.destino === "inventario") && items.some((it, i) => !it.isGift && !it.isSurcharge && it.base > 0 && !limpiarCuenta(a.cuentasItems?.[String(i)] ?? ""));
  if (sinCuentaPropia && !limpiarCuenta(a.cuentaGasto)) out.push("cuentaGasto");
  if (!limpiarCuenta(a.cuentaPorPagar)) out.push("cuentaPorPagar");
  // Cuentas adicionales que el motor exige según lo que traiga la factura (descuento/recargo del documento, obsequios, otros impuestos).
  const extra = (it: Ci2Factura["items"][number]) => it.extraTaxes || [];
  if (descuentoNeto(items) > 0 && !limpiarCuenta(a.cuentaDescuento)) out.push("cuentaDescuento");
  if (items.some((it) => it.isSurcharge && extra(it).some((e) => e.taxName === "Recargo" && e.amount !== 0)) && !limpiarCuenta(a.cuentaRecargo)) out.push("cuentaRecargo");
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
  for (const t of incs) if (!inv && !tarifasInc.has(t) && !limpiarCuenta(a.cuentasImpuestos?.[`INC ${t}%`] ?? "")) out.push(`imp:INC ${t}%`);
  if (tipo === "venta") {
    // Cada impuesto pide su propia cuenta.
    for (const n of nombres) if (!limpiarCuenta(a.cuentasImpuestos?.[n] ?? "")) out.push(`imp:${n}`);
  } else if (!inv && nombres.length && !limpiarCuenta(a.cuentaOtrosImpuestos)) out.push("cuentaOtrosImpuestos");
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
attachJobs("ci2-dian-traer", traerJobs as any);
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

router.post("/empresas/:id/traer", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    const tokenUrl = tryExtractDianUrlFromSafeLink(String(req.body?.tokenUrl || "").trim());
    const month = String(req.body?.month || "").trim();
    if (!/^\d{4}-\d{2}$/.test(month)) return err(res, new Error("Selecciona el mes a traer."));

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
    const [facturas, sugs, impuestos, glob] = await Promise.all([listarFacturas(empresaId, estado, tipo), sugerenciasPorNit(empresaId, tipo), getTabla(empresaId, "impuestos"), getCuentasGlobales(empresaId, tipo)]);
    const tarifasInc = tarifasIncDe(impuestos?.rows);
    const worldOffice = (await getSistemaContable(empresaId)) === "worldoffice";
    const cfgInv = await getConfigInventario(empresaId);
    const opciones = opcionesRetencion(impuestos?.rows || []);
    res.json({
      status: "ok",
      cuentasGlobales: glob,
      opcionesRetencion: opciones,
      facturas: facturas.map((f) => {
        const sug = sugs.get(f.nit);
        const ef = efectiva(f, sug, glob);
        return { ...f, retNoAplica: retencionesNoAplican(f.items, ef, opciones), categoria: categoriaDocumento(f.documentType, f.isCreditNote), sugerencia: sug || null, efectiva: ef, faltantes: faltantesDe(ef, f.items, tipo, tarifasInc, worldOffice, cfgInv.activo && !f.isCreditNote ? cfgInv : null), ivaMayorSugerido: ivaMayorSugerido(ef, cfgInv.activo ? cfgInv : null), sinParametros: !sug };
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
    // Solo un proveedor sin parámetros guarda lo editado; si ya tiene fila, el cambio aplica únicamente a esta factura.
    const proveedorNuevo = !(await sugerenciasPorNit(empresaId, tipo)).has(f.nit);
    if (aParametro.length && f.nit && proveedorNuevo) {
      await guardarParametroProveedor(empresaId, tipo, f.nit, f.proveedor, Object.fromEntries(aParametro.map((k) => [k, (f!.asignacion as any)[k]])));
    }
    const sug = (await sugerenciasPorNit(empresaId, tipo)).get(f.nit);
    const ef = efectiva(f, sug, await getCuentasGlobales(empresaId, tipo));
    const tarifasInc = tarifasIncDe((await getTabla(empresaId, "impuestos"))?.rows);
    res.json({ status: "ok", factura: { ...f, retNoAplica: retencionesNoAplican(f.items, ef, opcionesRetencion(((await getTabla(empresaId, "impuestos"))?.rows || []) as any)), sugerencia: sug || null, efectiva: ef, faltantes: faltantesDe(ef, f.items, tipo, tarifasInc, (await getSistemaContable(empresaId)) === "worldoffice", await cfgInventarioDe(empresaId, f.isCreditNote)), ivaMayorSugerido: ivaMayorSugerido(ef, await cfgInventarioDe(empresaId, f.isCreditNote)), sinParametros: !sug } });
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
    fila[col as string] = (campos as any)[campo] as any;
  }
  await setTabla(empresaId, slot, tabla);
}

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

// ── Maestro de productos de World Office ────────────────────────────────────

router.get("/empresas/:id/productos-wo", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    res.json({ status: "ok", productos: await getProductosWo(empresaId) });
  } catch (e) {
    err(res, e, 500);
  }
});

router.put("/empresas/:id/productos-wo", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    res.json({ status: "ok", productos: await setProductosWo(empresaId, req.body?.productos) });
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
router.get("/empresas/:id/prefijos", async (req: Request, res: Response) => {
  try {
    const empresaId = await exigirEmpresa(req, res);
    if (!empresaId) return;
    res.json({ status: "ok", prefijos: await getPrefijos(empresaId) });
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
    const worldoffice = sistemaContable === "worldoffice";
    const configWo = worldoffice ? await getConfigWorldOffice(empresaId) : null;
    if (worldoffice) {
      if (venta) return err(res, new Error("World Office solo está habilitado para compras por ahora."));
      if (!configWo!.empresa || !configWo!.terceroInterno || !configWo!.bodega) return err(res, new Error("Faltan los parámetros de World Office de la empresa (nombre en World Office, Tercero Interno y Bodega)."));
    } else if (!impuestos?.rows?.length) return err(res, new Error("La empresa no tiene tabla de impuestos cargada (Configuración)."));
    const pendientes = facturas.filter((f) => f.estado === "pendiente" && f.tipo === tipo);
    if (!pendientes.length) return err(res, new Error("Ninguna de las facturas seleccionadas está pendiente."));

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
      asignacion: efectiva(f, sugs.get(f.nit), glob),
    }));

    const r = generarInterfaz(entradas, {
      modo: tipo,
      inventario: !venta && !worldoffice && cfgInventario.activo ? cfgInventario : undefined,
      worldoffice,
      productosWo: worldoffice ? await getProductosWo(empresaId) : undefined,
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
    const sistema = syscafe ? "Syscafe" : "Siigo";
    const ext = syscafe ? "xls" : "xlsx";
    const escribir = async (filas: typeof r.salida) => (syscafe ? escribirInterfazSyscafe(filas) : await escribirInterfazSiigo(filas));
    let archivos: { nombre: string; tipo: "interfaz" | "informe"; data: Buffer }[];
    if (worldoffice) {
      // World Office carga dos archivos: encabezados de los documentos y su contenido.
      const wo = escribirInterfazWo(r.wo || [], configWo!);
      archivos = [
        { nombre: `Encabezados_WorldOffice_${sufijo}.xls`, tipo: "interfaz", data: wo.encabezados },
        { nombre: `Movimientos_WorldOffice_${sufijo}.xls`, tipo: "interfaz", data: wo.movimientos },
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
      archivos: archivos.map((a) => ({ nombre: a.nombre, tipo: a.tipo, url: `/api/causacion-interfaces-v2/interfaces/${interfazId}/archivos/${encodeURIComponent(a.nombre)}` })),
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
        archivos: i.archivos.map((a) => ({ ...a, url: `/api/causacion-interfaces-v2/interfaces/${i.id}/archivos/${encodeURIComponent(a.nombre)}` })),
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
  return i;
}

router.get("/interfaces/:iid/archivos/:nombre", async (req: Request, res: Response) => {
  try {
    if (!(await interfazAutorizada(req, res))) return;
    const a = await getArchivoInterfaz(req.params.iid, req.params.nombre);
    if (!a) return res.status(404).json({ status: "error", mensaje: "Archivo no encontrado." });
    res.setHeader("Content-Type", a.nombre.endsWith(".xls") ? "application/vnd.ms-excel" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
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
