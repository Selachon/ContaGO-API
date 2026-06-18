/**
 * Almacén multiempresa (tenant) para la herramienta de Contabilización DIAN→Siigo.
 *
 * La parametrización de cada empresa (su contenido difiere por empresa y NUNCA
 * se comparte ni se hardcodea) se IMPORTA UNA VEZ desde Excel a nuestra base de
 * datos y a partir de ahí se EDITA en el portal (sin volver a cargar archivos):
 *   - Parametrización Compras  → tabla (hoja "Proveedores")
 *   - Parametrización Ventas   → tabla (hoja "Proveedores")
 *   - Tabla maestro impuestos  → tabla (primera hoja; mapea el plan de cuentas)
 *
 * La plantilla de terceros (.xlsm con macros/geografía) SÍ sigue por archivo,
 * porque no es una tabla editable.
 *
 * Al correr el motor, `materializarConfig` reconstruye .xlsx idénticos desde las
 * tablas de la BD para pasárselos al motor sin alterar su lógica.
 *
 * Metadatos y tablas viven en Mongo (`contabilizacionEmpresas`); la plantilla en
 * disco bajo CONTAGO_DATA_DIR/contabilizacion/empresas/<empresaId>/.
 */
import { ObjectId } from "mongodb";
import fs from "fs";
import path from "path";
import { getDb } from "./database.js";
import { leerHoja, escribirTabla, sanearTabla, type Tabla } from "./contabilizacionTablasIO.js";
import { limpiarNit } from "./contabilizacionDianTerceros.js";

const EMPRESAS = "contabilizacionEmpresas";

export type ObsequiosMode = "error" | "contabilizar";
/** Tablas de parametrización editables (viven en la BD). */
export type TablaSlot = "paramCompras" | "paramVentas" | "impuestos";

/** Hoja que debe leerse al importar / escribirse al materializar cada tabla. */
const SHEET_REQUERIDA: Record<TablaSlot, string | undefined> = {
  paramCompras: "Proveedores",
  paramVentas: "Proveedores",
  impuestos: undefined, // primera hoja (como pandas sheet_name=0)
};

const PLANTILLA_FILENAME = "Plantilla_Terceros_Siigo.xlsm";
const TABLA_SLOTS: TablaSlot[] = ["paramCompras", "paramVentas", "impuestos"];

/** Tipos de comprobante de compras (parámetro inicial, normalmente estable). */
export interface Comprobantes {
  /** Tipo de comprobante para facturas de compra (p.ej. "CCOMP"). */
  tipoCompras: string;
  /** Tipo de comprobante para notas crédito de compra (p.ej. "NCOMP"). */
  tipoComprasNc: string;
}

/** Siguiente consecutivo recordado por tipo de comprobante (memoria). */
export type Consecutivos = Record<string, number>;

export interface EmpresaPublic {
  id: string;
  nombre: string;
  nit: string;
  obsequiosMode: ObsequiosMode;
  comprobantes: Comprobantes;
  /** Siguiente consecutivo recordado por tipo de comprobante. */
  consecutivos: Consecutivos;
  /** Qué tablas ya fueron importadas, con su conteo de filas. */
  tablas: Record<TablaSlot, { cargada: boolean; filas: number; columnas: number }>;
  /** True si la plantilla de terceros (.xlsm) está cargada. */
  plantillaTerceros: boolean;
  /** True si tiene las 3 tablas mínimas para compras/ventas. */
  listaParaContabilizar: boolean;
}

export interface ConfigEmpresa {
  id: string;
  nombre: string;
  nit: string;
  obsequiosMode: ObsequiosMode;
  comprobantes: Comprobantes;
  consecutivos: Consecutivos;
  /** Rutas absolutas a .xlsx materializados / archivo de plantilla. */
  paramCompras?: string;
  paramVentas?: string;
  impuestos?: string;
  plantillaTerceros?: string;
}

function normComprobantes(c: any): Comprobantes {
  return {
    tipoCompras: String(c?.tipoCompras ?? "").trim(),
    tipoComprasNc: String(c?.tipoComprasNc ?? "").trim(),
  };
}

/** Raíz de datos persistentes. En Railway conviene montar un volumen aquí. */
export function dataRoot(): string {
  return process.env.CONTAGO_DATA_DIR || path.join(process.cwd(), "data");
}

function empresaDir(empresaId: string): string {
  return path.join(dataRoot(), "contabilizacion", "empresas", empresaId);
}

function normalizeNit(value: unknown): string {
  return String(value ?? "").split("-")[0].replace(/\D/g, "");
}

function toObjectId(id: string): ObjectId | null {
  try {
    return new ObjectId(id);
  } catch {
    return null;
  }
}

function resumenTablas(doc: any): EmpresaPublic["tablas"] {
  const t = doc.tables || {};
  const uno = (slot: TablaSlot) => {
    const tab = t[slot] as Tabla | undefined;
    return {
      cargada: !!tab && Array.isArray(tab.columns) && tab.columns.length > 0,
      filas: tab?.rows?.length || 0,
      columnas: tab?.columns?.length || 0,
    };
  };
  return { paramCompras: uno("paramCompras"), paramVentas: uno("paramVentas"), impuestos: uno("impuestos") };
}

function toPublic(doc: any): EmpresaPublic {
  const tablas = resumenTablas(doc);
  return {
    id: doc._id.toString(),
    nombre: doc.nombre,
    nit: doc.nit || "",
    obsequiosMode: doc.obsequiosMode === "contabilizar" ? "contabilizar" : "error",
    comprobantes: normComprobantes(doc.comprobantes),
    consecutivos: (doc.consecutivos && typeof doc.consecutivos === "object" ? doc.consecutivos : {}) as Consecutivos,
    tablas,
    plantillaTerceros: !!doc.files?.plantillaTerceros,
    listaParaContabilizar: tablas.paramCompras.cargada && tablas.paramVentas.cargada && tablas.impuestos.cargada,
  };
}

/** Empresas visibles para un usuario (las que posee o le comparten). Admin ve todas. */
export async function listEmpresasForUser(userId: string, isAdmin: boolean): Promise<EmpresaPublic[]> {
  const filter = isAdmin ? {} : { $or: [{ ownerUserId: userId }, { sharedWith: userId }] };
  const docs = await getDb().collection<any>(EMPRESAS).find(filter).sort({ nombre: 1 }).toArray();
  return docs.map(toPublic);
}

/** True si el usuario puede usar la empresa (dueño, compartida o admin). */
export async function userCanAccessEmpresa(empresaId: string, userId: string, isAdmin: boolean): Promise<boolean> {
  const oid = toObjectId(empresaId);
  if (!oid) return false;
  if (isAdmin) {
    return !!(await getDb().collection<any>(EMPRESAS).findOne({ _id: oid }, { projection: { _id: 1 } }));
  }
  const doc = await getDb()
    .collection<any>(EMPRESAS)
    .findOne({ _id: oid, $or: [{ ownerUserId: userId }, { sharedWith: userId }] }, { projection: { _id: 1 } });
  return !!doc;
}

/** Crea una empresa para el tenant indicado. */
export async function createEmpresa(
  nombre: string,
  nit: string,
  obsequiosMode: ObsequiosMode,
  ownerUserId: string,
  comprobantes?: Partial<Comprobantes>
): Promise<EmpresaPublic> {
  const cleanNombre = String(nombre || "").trim();
  if (!cleanNombre) throw new Error("El nombre de la empresa es obligatorio.");
  const mode: ObsequiosMode = obsequiosMode === "contabilizar" ? "contabilizar" : "error";

  const db = getDb();
  const now = new Date();
  const res = await db.collection<any>(EMPRESAS).insertOne({
    nombre: cleanNombre,
    nit: normalizeNit(nit),
    obsequiosMode: mode,
    comprobantes: normComprobantes(comprobantes),
    consecutivos: {},
    ownerUserId,
    sharedWith: [],
    tables: {},
    files: {},
    createdAt: now,
    updatedAt: now,
  });
  const doc = await db.collection<any>(EMPRESAS).findOne({ _id: res.insertedId });
  return toPublic(doc);
}

/** Actualiza nombre / NIT / obsequiosMode / tipos de comprobante / consecutivos. */
export async function updateEmpresa(
  empresaId: string,
  patch: {
    nombre?: string;
    nit?: string;
    obsequiosMode?: ObsequiosMode;
    comprobantes?: Partial<Comprobantes>;
    consecutivos?: Consecutivos;
  }
): Promise<EmpresaPublic> {
  const oid = toObjectId(empresaId);
  if (!oid) throw new Error("Empresa inválida.");
  const set: any = { updatedAt: new Date() };
  if (patch.nombre !== undefined) set.nombre = String(patch.nombre).trim();
  if (patch.nit !== undefined) set.nit = normalizeNit(patch.nit);
  if (patch.obsequiosMode !== undefined) {
    set.obsequiosMode = patch.obsequiosMode === "contabilizar" ? "contabilizar" : "error";
  }
  if (patch.comprobantes !== undefined) set.comprobantes = normComprobantes(patch.comprobantes);
  if (patch.consecutivos !== undefined && patch.consecutivos && typeof patch.consecutivos === "object") {
    const limpio: Consecutivos = {};
    for (const [k, v] of Object.entries(patch.consecutivos)) {
      const n = Number(v);
      if (k && Number.isFinite(n)) limpio[k] = Math.trunc(n);
    }
    set.consecutivos = limpio;
  }
  const db = getDb();
  const res = await db.collection<any>(EMPRESAS).updateOne({ _id: oid }, { $set: set });
  if (res.matchedCount === 0) throw new Error("Empresa no encontrada.");
  const doc = await db.collection<any>(EMPRESAS).findOne({ _id: oid });
  return toPublic(doc);
}

/**
 * Recuerda el siguiente consecutivo por tipo de comprobante (memoria). Solo
 * actualiza las claves provistas; conserva el resto.
 */
export async function guardarConsecutivos(empresaId: string, siguientes: Consecutivos): Promise<void> {
  const oid = toObjectId(empresaId);
  if (!oid) throw new Error("Empresa inválida.");
  const set: any = { updatedAt: new Date() };
  for (const [tipo, n] of Object.entries(siguientes)) {
    if (tipo && Number.isFinite(Number(n))) set[`consecutivos.${tipo}`] = Math.trunc(Number(n));
  }
  await getDb().collection<any>(EMPRESAS).updateOne({ _id: oid }, { $set: set });
}

/**
 * Importa (primera vez o reemplazo) una tabla de parametrización desde un Excel a
 * la BD. Lee la hoja que el motor espera y guarda {sheetName, columns, rows}.
 */
export async function importarTabla(empresaId: string, slot: TablaSlot, buffer: Buffer): Promise<Tabla> {
  const oid = toObjectId(empresaId);
  if (!oid) throw new Error("Empresa inválida.");
  const tabla = await leerHoja(buffer, SHEET_REQUERIDA[slot]);
  await getDb()
    .collection<any>(EMPRESAS)
    .updateOne({ _id: oid }, { $set: { [`tables.${slot}`]: tabla, updatedAt: new Date() } });
  return tabla;
}

/** Devuelve la tabla de parametrización (para editar en el portal) o null. */
export async function getTabla(empresaId: string, slot: TablaSlot): Promise<Tabla | null> {
  const oid = toObjectId(empresaId);
  if (!oid) throw new Error("Empresa inválida.");
  const doc = await getDb().collection<any>(EMPRESAS).findOne({ _id: oid }, { projection: { [`tables.${slot}`]: 1 } });
  return (doc?.tables?.[slot] as Tabla) || null;
}

/** Guarda la tabla editada desde el portal (sin volver a cargar archivo). */
export async function setTabla(empresaId: string, slot: TablaSlot, input: unknown): Promise<Tabla> {
  const oid = toObjectId(empresaId);
  if (!oid) throw new Error("Empresa inválida.");
  const tabla = sanearTabla(input);
  // Conserva el nombre de hoja que el motor espera para esta tabla.
  if (SHEET_REQUERIDA[slot]) tabla.sheetName = SHEET_REQUERIDA[slot] as string;
  const res = await getDb()
    .collection<any>(EMPRESAS)
    .updateOne({ _id: oid }, { $set: { [`tables.${slot}`]: tabla, updatedAt: new Date() } });
  if (res.matchedCount === 0) throw new Error("Empresa no encontrada.");
  return tabla;
}

function normalizarCol(s: string): string {
  return String(s).trim().toLowerCase().normalize("NFKD").replace(/[̀-ͯ]/g, "").replace(/\s+/g, " ");
}

/** Columna que contiene el NIT del tercero en la tabla de parametrización. */
function columnaNit(columns: string[]): string | null {
  const exact = columns.find((c) => normalizarCol(c) === "nit emisor");
  if (exact) return exact;
  const contiene = columns.find((c) => normalizarCol(c).includes("nit"));
  return contiene || null;
}

/** Columna de nombre/razón social del tercero (para pre-rellenar), si existe. */
function columnaNombre(columns: string[], nitCol: string | null): string | null {
  return (
    columns.find(
      (c) =>
        c !== nitCol &&
        /(razon|raz n|proveedor|nombre|tercero)/.test(normalizarCol(c))
    ) || null
  );
}

export interface PrecreacionResultado {
  slot: TablaSlot;
  nitColumn: string | null;
  nameColumn: string | null;
  /** NITs efectivamente pre-creados (no existían en la tabla). */
  creados: string[];
  /** NITs solicitados (faltantes), para que el front los resalte. */
  nits: string[];
}

/**
 * Pre-crea en la tabla de parametrización los terceros faltantes con la info
 * disponible (NIT y, si la tabla tiene columna de nombre, la razón social del
 * DIAN). El resto de columnas quedan vacías para que el usuario las complete.
 */
export async function precrearTerceros(
  empresaId: string,
  slot: TablaSlot,
  nits: string[],
  nombres: Map<string, string>
): Promise<PrecreacionResultado> {
  const tabla = await getTabla(empresaId, slot);
  const faltantes = nits.map((n) => limpiarNit(n)).filter(Boolean);
  if (!tabla || !tabla.columns?.length) {
    return { slot, nitColumn: null, nameColumn: null, creados: [], nits: faltantes };
  }

  const nitCol = columnaNit(tabla.columns);
  const nameCol = columnaNombre(tabla.columns, nitCol);
  if (!nitCol) {
    return { slot, nitColumn: null, nameColumn: nameCol, creados: [], nits: faltantes };
  }

  const existentes = new Set(tabla.rows.map((r) => limpiarNit(r[nitCol] as unknown)));
  const creados: string[] = [];
  for (const nit of faltantes) {
    if (existentes.has(nit)) continue;
    const fila: Record<string, string | number | boolean | null> = {};
    for (const c of tabla.columns) fila[c] = "";
    fila[nitCol] = nit;
    if (nameCol) fila[nameCol] = nombres.get(nit) || "";
    tabla.rows.push(fila);
    existentes.add(nit);
    creados.push(nit);
  }

  if (creados.length) await setTabla(empresaId, slot, tabla);
  return { slot, nitColumn: nitCol, nameColumn: nameCol, creados, nits: faltantes };
}

/** Primera columna de la tabla cuyo nombre normalizado está entre los alias. */
function columnaPorAlias(columns: string[], alias: string[]): string | null {
  const set = new Set(alias.map((a) => normalizarCol(a)));
  return columns.find((c) => set.has(normalizarCol(c))) || null;
}

/** Cuentas SIEMPRE requeridas por fila (acepta los alias de compras y ventas). */
const ALIAS_CUENTA_GASTO = ["Cuenta_gasto", "Cuenta_ingreso"];
const ALIAS_CUENTA_PAGAR = ["Cuenta_por_pagar", "Cuenta_por_cobrar"];

export interface ProveedoresIncompletosResultado {
  slot: TablaSlot;
  nitColumn: string | null;
  nameColumn: string | null;
  /** NITs presentes en la tabla pero con alguna cuenta obligatoria vacía. */
  incompletos: string[];
}

const vacio = (v: unknown): boolean => String(v ?? "").trim() === "";

/**
 * Detecta terceros que SÍ están en la tabla (e involucrados en el DIAN) pero con
 * cuentas obligatorias (gasto / por pagar) vacías. Solo evalúa los NITs presentes
 * en `dianNits` para no bloquear por filas ajenas al archivo procesado.
 */
export async function proveedoresIncompletos(
  empresaId: string,
  slot: TablaSlot,
  dianNits: Set<string>
): Promise<ProveedoresIncompletosResultado> {
  const tabla = await getTabla(empresaId, slot);
  if (!tabla || !tabla.columns?.length) {
    return { slot, nitColumn: null, nameColumn: null, incompletos: [] };
  }
  const nitCol = columnaNit(tabla.columns);
  const nameCol = columnaNombre(tabla.columns, nitCol);
  const gastoCol = columnaPorAlias(tabla.columns, ALIAS_CUENTA_GASTO);
  const pagarCol = columnaPorAlias(tabla.columns, ALIAS_CUENTA_PAGAR);
  if (!nitCol || (!gastoCol && !pagarCol)) {
    return { slot, nitColumn: nitCol, nameColumn: nameCol, incompletos: [] };
  }

  const incompletos: string[] = [];
  for (const row of tabla.rows) {
    const nit = limpiarNit(row[nitCol] as unknown);
    if (!nit || !dianNits.has(nit)) continue;
    const faltaGasto = gastoCol ? vacio(row[gastoCol]) : false;
    const faltaPagar = pagarCol ? vacio(row[pagarCol]) : false;
    if (faltaGasto || faltaPagar) incompletos.push(nit);
  }
  return { slot, nitColumn: nitCol, nameColumn: nameCol, incompletos };
}

/** Guarda (o reemplaza) la plantilla de terceros (.xlsm) en disco. */
export async function savePlantillaTerceros(empresaId: string, buffer: Buffer): Promise<void> {
  const oid = toObjectId(empresaId);
  if (!oid) throw new Error("Empresa inválida.");
  const dir = empresaDir(empresaId);
  fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(path.join(dir, PLANTILLA_FILENAME), buffer);
  await getDb()
    .collection<any>(EMPRESAS)
    .updateOne({ _id: oid }, { $set: { "files.plantillaTerceros": PLANTILLA_FILENAME, updatedAt: new Date() } });
}

/**
 * Materializa la configuración de la empresa en `destDir`: reconstruye los .xlsx
 * de las tablas (compras/ventas/impuestos) desde la BD y resuelve la plantilla de
 * terceros en disco. Devuelve rutas absolutas listas para el motor. No se comparte
 * nada entre empresas: solo se materializa lo que esta empresa tiene.
 */
export async function materializarConfig(empresaId: string, destDir: string): Promise<ConfigEmpresa> {
  const oid = toObjectId(empresaId);
  if (!oid) throw new Error("Empresa inválida.");
  const doc = await getDb().collection<any>(EMPRESAS).findOne({ _id: oid });
  if (!doc) throw new Error("Empresa no encontrada.");

  fs.mkdirSync(destDir, { recursive: true });
  const filenames: Record<TablaSlot, string> = {
    paramCompras: "Parametrizacion_Siigo_Compras.xlsx",
    paramVentas: "Parametrizacion_Siigo_Ventas.xlsx",
    impuestos: "Tabla_maestro_impuestos.xlsx",
  };

  const out: ConfigEmpresa = {
    id: doc._id.toString(),
    nombre: doc.nombre,
    nit: doc.nit || "",
    obsequiosMode: doc.obsequiosMode === "contabilizar" ? "contabilizar" : "error",
    comprobantes: normComprobantes(doc.comprobantes),
    consecutivos: (doc.consecutivos && typeof doc.consecutivos === "object" ? doc.consecutivos : {}) as Consecutivos,
  };

  for (const slot of TABLA_SLOTS) {
    const tabla = doc.tables?.[slot] as Tabla | undefined;
    if (tabla && tabla.columns?.length) {
      // Fuerza el nombre de hoja que el motor lee (Proveedores en compras/ventas).
      const sheetName = SHEET_REQUERIDA[slot] || tabla.sheetName || "Hoja1";
      const dest = path.join(destDir, filenames[slot]);
      await escribirTabla({ ...tabla, sheetName }, dest);
      out[slot] = dest;
    }
  }

  // Plantilla de terceros (archivo en disco; no es tabla editable).
  const plantillaName = doc.files?.plantillaTerceros;
  if (plantillaName) {
    const p = path.join(empresaDir(empresaId), plantillaName);
    if (fs.existsSync(p)) out.plantillaTerceros = p;
  }

  return out;
}
