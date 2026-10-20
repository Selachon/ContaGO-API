/**
 * Persistencia de "Causación Interfaces 2.0" (admin).
 *
 * Colecciones propias (la parametrización de la empresa —tabla de impuestos,
 * parámetros por proveedor, comprobantes y consecutivos— se comparte con
 * "Causación Siigo (Interfaces)" porque es la misma contabilidad):
 *   - ci2Facturas:    facturas traídas de la DIAN + asignaciones del usuario.
 *   - ci2Interfaces:  interfaces generadas (con sus archivos) para descargar/revertir.
 */
import { ObjectId, type Binary } from "mongodb";
import { getDb } from "./database.js";
import type { BatchItem, XmlCausacionItem } from "./siigoAccountingService.js";
import { claveImpuesto, type Ci2Asignacion } from "./ci2Motor.js";

const FACTURAS = "ci2Facturas";
const INTERFACES = "ci2Interfaces";

export type Ci2Estado = "pendiente" | "generada" | "descartada";
/** Compras = facturas recibidas (proveedor); ventas = facturas emitidas (cliente). Lo guardado antes de existir ventas es compra. */
export type Ci2Tipo = "compra" | "venta";

const filtroTipo = (tipo: Ci2Tipo) => (tipo === "venta" ? { tipo: "venta" } : { tipo: { $ne: "venta" } });

export interface Ci2Factura {
  id: string;
  empresaId: string;
  tipo: Ci2Tipo;
  cufe: string;
  fecha: string; // yyyy-mm-dd
  documentType: string;
  isCreditNote: boolean;
  docNumber: string;
  /** NIT y nombre de la contraparte: el proveedor en compras, el cliente en ventas. */
  nit: string;
  proveedor: string;
  items: XmlCausacionItem[];
  totals: { subtotal: number; iva: number; total: number };
  estado: Ci2Estado;
  tienePdf: boolean;
  /** Solo lo que el usuario editó a mano; el resto sale de los parámetros del proveedor. */
  asignacion: Partial<Ci2Asignacion>;
  interfazId?: string;
  comprobante?: { tipo: string; consecutivo: number };
}

export interface Ci2InterfazArchivo {
  nombre: string;
  tipo: "interfaz" | "informe";
  data: Buffer;
}

export interface Ci2InterfazDoc {
  id: string;
  empresaId: string;
  userId: string;
  tipo?: Ci2Tipo;
  createdAt: Date;
  resumen: { documentos: number; notasCredito: number; lineas: number; partes: number };
  facturaIds: string[];
  /** Consecutivo con el que arrancó y el siguiente que quedó, por tipo. */
  inicial: Record<string, number>;
  siguientes: Record<string, number>;
  advertencias: string[];
  revertida?: boolean;
  archivos: { nombre: string; tipo: "interfaz" | "informe" }[];
}

const oid = (id: string): ObjectId | null => (ObjectId.isValid(id) ? new ObjectId(id) : null);

const SIN_PDF = { pdf: 0 };
const colFacturas = () => getDb().collection<any>(FACTURAS);
const colInterfaces = () => getDb().collection<any>(INTERFACES);

export async function ensureCi2Indexes(): Promise<void> {
  await colFacturas().createIndex({ empresaId: 1, cufe: 1 }, { unique: true }).catch(() => undefined);
  await colFacturas().createIndex({ empresaId: 1, estado: 1, fecha: -1 }).catch(() => undefined);
  await colInterfaces().createIndex({ empresaId: 1, createdAt: -1 }).catch(() => undefined);
}

function toFactura(d: any): Ci2Factura {
  return {
    id: String(d._id),
    empresaId: d.empresaId,
    tipo: d.tipo === "venta" ? "venta" : "compra",
    cufe: d.cufe,
    fecha: d.fecha,
    documentType: d.documentType,
    isCreditNote: !!d.isCreditNote,
    docNumber: d.docNumber,
    nit: d.nit,
    proveedor: d.proveedor,
    items: d.items || [],
    totals: d.totals || { subtotal: 0, iva: 0, total: 0 },
    estado: d.estado,
    tienePdf: !!d.tienePdf,
    asignacion: d.asignacion || {},
    interfazId: d.interfazId,
    comprobante: d.comprobante,
  };
}

/** Guarda las facturas recién traídas; las que ya existen (por CUFE) no se tocan. Devuelve cuántas son nuevas. */
export async function guardarFacturasTraidas(empresaId: string, items: BatchItem[], tipo: Ci2Tipo = "compra"): Promise<number> {
  const ops: any[] = [];
  for (const it of items) {
    if (!it.ok || !it.xml) continue;
    const x = it.xml;
    const nitContraparte = tipo === "venta" ? x.receiverNit : x.supplierNit;
    const cufe = x.cufe || `f:${it.fileName}`;
    ops.push({
        updateOne: {
          filter: { empresaId, cufe },
          update: {
            $setOnInsert: {
              empresaId,
              tipo,
              cufe,
              fileName: it.fileName,
              fecha: x.date,
              documentType: x.documentType,
              isCreditNote: x.isCreditNote,
              docNumber: x.docNumberRaw,
              nit: String(nitContraparte || "").split("-")[0].replace(/\D/g, ""),
              proveedor: tipo === "venta" ? x.receiverName || "" : x.supplierName,
              items: x.items,
              totals: x.totals,
              estado: "pendiente" as Ci2Estado,
              asignacion: {},
              fetchedAt: new Date(),
            },
          },
          upsert: true,
        },
    });
    // El PDF se completa también en facturas ya guardadas que no lo tenían.
    ops.push({
      updateOne: {
        filter: { empresaId, cufe, tienePdf: { $ne: true } },
        update: { $set: { pdfIntentado: true, ...(it.pdfBase64 ? { pdf: Buffer.from(it.pdfBase64, "base64"), pdfName: it.pdfName || "factura.pdf", tienePdf: true } : {}) } },
      },
    });
  }
  if (!ops.length) return 0;
  const r = await colFacturas().bulkWrite(ops, { ordered: true });
  return r.upsertedCount;
}

export async function listarFacturas(empresaId: string, estado?: Ci2Estado, tipo: Ci2Tipo = "compra"): Promise<Ci2Factura[]> {
  const filter: any = { empresaId, ...filtroTipo(tipo) };
  if (estado) filter.estado = estado;
  const docs = await colFacturas().find(filter, { projection: SIN_PDF }).sort({ fecha: 1, docNumber: 1 }).toArray();
  return docs.map(toFactura);
}

export async function facturasPorIds(empresaId: string, ids: string[]): Promise<Ci2Factura[]> {
  const oids = ids.map(oid).filter((x): x is ObjectId => !!x);
  const docs = await colFacturas().find({ empresaId, _id: { $in: oids } }, { projection: SIN_PDF }).toArray();
  const porId = new Map(docs.map((d) => [String(d._id), toFactura(d)]));
  // Conserva el orden pedido por el usuario.
  return ids.map((i) => porId.get(i)).filter((f): f is Ci2Factura => !!f);
}

/** CUFEs ya registrados de la empresa que no hace falta volver a descargar (los que aún no se intentó traer con su PDF sí se reintentan). */
export async function cufesRegistrados(empresaId: string): Promise<string[]> {
  const docs = await colFacturas().find({ empresaId, $or: [{ tienePdf: true }, { pdfIntentado: true }] }, { projection: { cufe: 1 } }).toArray();
  return docs.map((d) => d.cufe as string).filter((c) => c && !c.startsWith("f:"));
}

const CAMPOS_ASIGNACION: (keyof Ci2Asignacion)[] = [
  "cuentaGasto", "cuentaPorPagar", "cuentaIngresoObsequios", "cuentaOtrosImpuestos", "cuentaDescuento",
  "cuentaRecargo", "centroCostos", "tarifaRetefuente", "tarifaReteica", "tarifaReteiva", "baseMinimaReteicaEspecial", "cuentasItems",
];
const CAMPOS_TEXTO = new Set<string>(["cuentaGasto", "cuentaPorPagar", "cuentaIngresoObsequios", "cuentaOtrosImpuestos", "cuentaDescuento", "cuentaRecargo", "centroCostos"]);

/** Aplica ediciones manuales a una factura pendiente. `null` en un campo = volver a la sugerencia. */
export async function editarAsignacion(empresaId: string, id: string, cambios: Record<string, unknown>): Promise<Ci2Factura | null> {
  const _id = oid(id);
  if (!_id) return null;
  const set: any = {};
  const unset: any = {};
  for (const k of CAMPOS_ASIGNACION) {
    if (!(k in cambios)) continue;
    const v = cambios[k];
    if (v === null) unset[`asignacion.${k}`] = "";
    else if (k === "cuentasItems") {
      // { índiceÍtem: cuenta }; un valor vacío quita la cuenta propia de ese ítem.
      const limpio: Record<string, string> = {};
      for (const [i, c] of Object.entries((v && typeof v === "object" ? v : {}) as Record<string, unknown>)) {
        const cuenta = String(c ?? "").trim();
        if (/^\d+$/.test(i) && cuenta) limpio[i] = cuenta;
      }
      if (Object.keys(limpio).length) set["asignacion.cuentasItems"] = limpio;
      else unset["asignacion.cuentasItems"] = "";
    }
    else if (CAMPOS_TEXTO.has(k)) set[`asignacion.${k}`] = String(v).trim();
    else {
      const n = Number(String(v).replace(",", "."));
      if (!Number.isFinite(n) || n < 0) throw new Error(`Valor inválido en ${k}.`);
      set[`asignacion.${k}`] = n;
    }
  }
  const upd: any = {};
  if (Object.keys(set).length) upd.$set = set;
  if (Object.keys(unset).length) upd.$unset = unset;
  if (!Object.keys(upd).length) return toFactura(await colFacturas().findOne({ _id, empresaId }, { projection: SIN_PDF }));
  const r = await colFacturas().findOneAndUpdate({ _id, empresaId, estado: { $ne: "generada" } }, upd, { returnDocument: "after", projection: SIN_PDF });
  return r ? toFactura(r) : null;
}

/** Cambia el estado entre pendiente y descartada (las generadas solo se liberan revirtiendo la interfaz). */
export async function getPdfFactura(empresaId: string, id: string): Promise<{ nombre: string; data: Buffer } | null> {
  const _id = oid(id);
  if (!_id) return null;
  const d = await colFacturas().findOne({ _id, empresaId }, { projection: { pdf: 1, pdfName: 1 } });
  if (!d?.pdf) return null;
  const raw = d.pdf as Binary | Buffer;
  return { nombre: d.pdfName || "factura.pdf", data: Buffer.isBuffer(raw) ? raw : Buffer.from((raw as Binary).buffer) };
}

export async function marcarEstado(empresaId: string, ids: string[], estado: "pendiente" | "descartada"): Promise<number> {
  const oids = ids.map(oid).filter((x): x is ObjectId => !!x);
  const r = await colFacturas().updateMany({ empresaId, _id: { $in: oids }, estado: { $ne: "generada" } }, { $set: { estado } });
  return r.modifiedCount;
}

export async function guardarInterfaz(
  doc: Omit<Ci2InterfazDoc, "id" | "archivos">,
  archivos: Ci2InterfazArchivo[],
  asignados: { id: string; tipo: string; consecutivo: number }[]
): Promise<string> {
  const ins = await colInterfaces().insertOne({
    ...doc,
    archivos: archivos.map((a) => ({ nombre: a.nombre, tipo: a.tipo, data: a.data })),
  });
  const interfazId = String(ins.insertedId);
  await Promise.all(
    asignados.map((a) => {
      const _id = oid(a.id);
      return _id
        ? colFacturas().updateOne(
            { _id, empresaId: doc.empresaId },
            { $set: { estado: "generada", interfazId, comprobante: { tipo: a.tipo, consecutivo: a.consecutivo } } }
          )
        : Promise.resolve();
    })
  );
  return interfazId;
}

function toInterfaz(d: any): Ci2InterfazDoc {
  return {
    id: String(d._id),
    empresaId: d.empresaId,
    userId: d.userId,
    tipo: d.tipo === "venta" ? "venta" : "compra",
    createdAt: d.createdAt,
    resumen: d.resumen,
    facturaIds: d.facturaIds || [],
    inicial: d.inicial || {},
    siguientes: d.siguientes || {},
    advertencias: d.advertencias || [],
    revertida: !!d.revertida,
    archivos: (d.archivos || []).map((a: any) => ({ nombre: a.nombre, tipo: a.tipo })),
  };
}

export async function listarInterfaces(empresaId: string, limite = 50): Promise<Ci2InterfazDoc[]> {
  const docs = await colInterfaces().find({ empresaId }, { projection: { "archivos.data": 0 } }).sort({ createdAt: -1 }).limit(limite).toArray();
  return docs.map(toInterfaz);
}

export async function getInterfaz(id: string): Promise<Ci2InterfazDoc | null> {
  const _id = oid(id);
  if (!_id) return null;
  const d = await colInterfaces().findOne({ _id }, { projection: { "archivos.data": 0 } });
  return d ? toInterfaz(d) : null;
}

export async function getArchivoInterfaz(id: string, nombre: string): Promise<{ nombre: string; data: Buffer } | null> {
  const _id = oid(id);
  if (!_id) return null;
  const d = await colInterfaces().findOne({ _id });
  const a = (d?.archivos || []).find((x: any) => x.nombre === nombre);
  if (!a) return null;
  const raw = a.data as Binary | Buffer;
  return { nombre: a.nombre, data: Buffer.isBuffer(raw) ? raw : Buffer.from((raw as Binary).buffer) };
}

/** Libera las facturas de una interfaz y la marca revertida. */
export async function revertirInterfaz(id: string): Promise<number> {
  const _id = oid(id);
  if (!_id) return 0;
  const r = await colFacturas().updateMany(
    { interfazId: id },
    { $set: { estado: "pendiente" }, $unset: { interfazId: "", comprobante: "" } }
  );
  await colInterfaces().updateOne({ _id }, { $set: { revertida: true } });
  return r.modifiedCount;
}

// ── Cuentas globales de la empresa (no dependen del proveedor) ──────────────

export type Ci2CuentasGlobales = { cuentaDescuento?: string; cuentaRecargo?: string; cuentaOtrosImpuestos?: string; cuentasImpuestos?: Record<string, string> };
export const CAMPOS_GLOBALES = ["cuentaDescuento", "cuentaRecargo", "cuentaOtrosImpuestos"] as const;

/** Campo de la empresa donde se guardan las cuentas generales de cada módulo. */
const campoGlobales = (tipo: Ci2Tipo) => (tipo === "venta" ? "ci2CuentasVentas" : "ci2Cuentas");

export async function getCuentasGlobales(empresaId: string, tipo: Ci2Tipo = "compra"): Promise<Ci2CuentasGlobales> {
  const _id = oid(empresaId);
  if (!_id) return {};
  const campo = campoGlobales(tipo);
  const d = await getDb().collection<any>("contabilizacionEmpresas").findOne({ _id }, { projection: { [campo]: 1 } });
  return (d?.[campo] as Ci2CuentasGlobales) || {};
}

export async function setCuentasGlobales(empresaId: string, cambios: Ci2CuentasGlobales, tipo: Ci2Tipo = "compra"): Promise<void> {
  const _id = oid(empresaId);
  if (!_id) throw new Error("Empresa inválida.");
  const campo = campoGlobales(tipo);
  const set: any = {};
  for (const k of CAMPOS_GLOBALES) if (k in cambios) set[`${campo}.${k}`] = String(cambios[k] ?? "").trim();
  // Cuenta por impuesto (ventas): se mezcla por nombre sin tocar las demás.
  for (const [nombre, cuenta] of Object.entries(cambios.cuentasImpuestos || {})) {
    const clave = claveImpuesto(nombre);
    if (clave) set[`${campo}.cuentasImpuestos.${clave}`] = String(cuenta ?? "").trim();
  }
  if (Object.keys(set).length) await getDb().collection<any>("contabilizacionEmpresas").updateOne({ _id }, { $set: set });
}

// ── Autorretención de ventas (por empresa; la usan ambas herramientas) ──────

export interface Autorretencion {
  habilitada: boolean;
  /** % sobre la base de la factura antes de impuestos. */
  porcentaje: number;
  /** Activo 1355… (débito). */
  cuentaActivo: string;
  /** Pasivo 2365… (crédito). */
  cuentaPasivo: string;
}

const AUTORRETENCION_VACIA: Autorretencion = { habilitada: false, porcentaje: 0, cuentaActivo: "", cuentaPasivo: "" };

export async function getAutorretencion(empresaId: string): Promise<Autorretencion> {
  const _id = oid(empresaId);
  if (!_id) return { ...AUTORRETENCION_VACIA };
  const d = await getDb().collection<any>("contabilizacionEmpresas").findOne({ _id }, { projection: { autorretencionVentas: 1 } });
  return { ...AUTORRETENCION_VACIA, ...(d?.autorretencionVentas || {}) };
}

/** Valida y guarda; si está habilitada exige porcentaje y las dos cuentas. */
export async function setAutorretencion(empresaId: string, a: Partial<Autorretencion>): Promise<Autorretencion> {
  const _id = oid(empresaId);
  if (!_id) throw new Error("Empresa inválida.");
  const cuenta = (v: unknown) => String(v ?? "").trim().split(/\s|-/)[0].replace(/\D/g, "");
  const limpia: Autorretencion = {
    habilitada: !!a.habilitada,
    porcentaje: Math.round((Number(a.porcentaje) || 0) * 10000) / 10000,
    cuentaActivo: cuenta(a.cuentaActivo),
    cuentaPasivo: cuenta(a.cuentaPasivo),
  };
  if (limpia.habilitada) {
    if (!(limpia.porcentaje > 0 && limpia.porcentaje <= 100)) throw new Error("Indica el porcentaje de la autorretención (mayor que 0 y hasta 100).");
    if (!limpia.cuentaActivo) throw new Error("Falta la cuenta de activo de la autorretención (1355…).");
    if (!limpia.cuentaPasivo) throw new Error("Falta la cuenta de pasivo de la autorretención (2365…).");
  }
  await getDb().collection<any>("contabilizacionEmpresas").updateOne({ _id }, { $set: { autorretencionVentas: limpia } });
  return limpia;
}

// ── Tipo de comprobante por clase de documento (FE, NC, DS, DEPOS…) ─────────

export async function getTiposDocumento(empresaId: string, tipo: Ci2Tipo = "compra"): Promise<Record<string, string>> {
  const _id = oid(empresaId);
  if (!_id) return {};
  const d = await getDb().collection<any>("contabilizacionEmpresas").findOne({ _id }, { projection: { ci2TiposDocumento: 1 } });
  return d?.ci2TiposDocumento?.[tipo === "venta" ? "venta" : "compra"] || {};
}

/** Mezcla por clave: lo no enviado queda como estaba. */
export async function setTiposDocumento(empresaId: string, tipo: Ci2Tipo, nuevos: Record<string, string>): Promise<void> {
  const _id = oid(empresaId);
  if (!_id) throw new Error("Empresa inválida.");
  const set: Record<string, string> = {};
  for (const [k, v] of Object.entries(nuevos)) {
    const clave = k.replace(/[.$]/g, "").trim();
    const valor = String(v ?? "").trim();
    if (clave && valor) set[`ci2TiposDocumento.${tipo === "venta" ? "venta" : "compra"}.${clave}`] = valor;
  }
  if (Object.keys(set).length) await getDb().collection<any>("contabilizacionEmpresas").updateOne({ _id }, { $set: set });
}

// ── Tipo de comprobante por prefijo de la factura emitida (ventas) ──────────

export interface Ci2Prefijos {
  factura: Record<string, string>;
  notaCredito: Record<string, string>;
}

export async function getPrefijos(empresaId: string): Promise<Ci2Prefijos> {
  const _id = oid(empresaId);
  if (!_id) return { factura: {}, notaCredito: {} };
  const d = await getDb().collection<any>("contabilizacionEmpresas").findOne({ _id }, { projection: { ci2Prefijos: 1 } });
  return { factura: d?.ci2Prefijos?.factura || {}, notaCredito: d?.ci2Prefijos?.notaCredito || {} };
}

export async function setPrefijos(empresaId: string, nuevos: Partial<Ci2Prefijos>): Promise<void> {
  const _id = oid(empresaId);
  if (!_id) throw new Error("Empresa inválida.");
  const actual = await getPrefijos(empresaId);
  const limpiar = (m: Record<string, string> | undefined) =>
    Object.fromEntries(Object.entries(m || {}).map(([k, v]) => [k.trim().toUpperCase(), String(v ?? "").trim()]).filter(([k, v]) => k && v));
  await getDb().collection<any>("contabilizacionEmpresas").updateOne({ _id }, {
    $set: { ci2Prefijos: { factura: { ...actual.factura, ...limpiar(nuevos.factura) }, notaCredito: { ...actual.notaCredito, ...limpiar(nuevos.notaCredito) } } },
  });
}
