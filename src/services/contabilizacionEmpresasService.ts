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
import JSZip from "jszip";
import fs from "fs";
import path from "path";
import { getDb } from "./database.js";
import { leerHoja, leerHojaConEncabezadoDinamico, escribirTabla, sanearTabla, type Tabla } from "./contabilizacionTablasIO.js";
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

export interface CuentaPuc {
  codigo: string;
  nombre: string;
}

export interface PucResumen {
  cargado: boolean;
  cuentas: number;
}

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
  /** PUC propio de la empresa para ayudar a seleccionar cuentas. */
  puc: PucResumen;
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
    puc: {
      cargado: Array.isArray(doc.puc?.cuentas) && doc.puc.cuentas.length > 0,
      cuentas: Array.isArray(doc.puc?.cuentas) ? doc.puc.cuentas.length : 0,
    },
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

function limpiarCodigoCuenta(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "number" && Number.isFinite(value)) return String(Math.trunc(value));
  let text = String(value).trim();
  if (text.endsWith(".0")) text = text.slice(0, -2);
  return text;
}

function columnaPucCodigo(columns: string[]): string | null {
  const aliases = new Set([
    "codigo",
    "codigo cuenta",
    "codigo contable",
    "cuenta",
    "cuenta contable",
    "cod cuenta",
    "puc",
  ]);
  return columns.find((c) => aliases.has(normalizarCol(c))) || columns.find((c) => normalizarCol(c).includes("codigo")) || columns[0] || null;
}

function columnaPucNombre(columns: string[], codeCol: string | null): string | null {
  const aliases = ["nombre", "nombre cuenta", "descripcion", "descripción", "detalle", "cuenta"];
  const exact = columns.find((c) => c !== codeCol && aliases.includes(normalizarCol(c)));
  if (exact) return exact;
  return columns.find((c) => c !== codeCol && /(nombre|descripcion|descripci n|detalle|cuenta)/.test(normalizarCol(c))) || columns.find((c) => c !== codeCol) || null;
}

function columnaPucNivelAgrupacion(columns: string[]): string | null {
  return columns.find((c) => normalizarCol(c) === "nivel agrupacion") ||
    columns.find((c) => normalizarCol(c).includes("nivel") && normalizarCol(c).includes("agrupacion")) ||
    null;
}

function columnaPucActivo(columns: string[]): string | null {
  return columns.find((c) => normalizarCol(c) === "activo") ||
    columns.find((c) => normalizarCol(c).includes("activo")) ||
    null;
}

function esPucTransaccional(value: unknown): boolean {
  return normalizarCol(String(value ?? "")).includes("transaccional");
}

function esPucActivo(value: unknown): boolean {
  const n = normalizarCol(String(value ?? ""));
  return n === "si" || n === "s" || n === "true" || n === "1";
}


function decodeXmlPuc(value: string): string {
  return value
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

function colLettersToIndexPuc(letters: string): number {
  let index = 0;
  for (const ch of letters) index = index * 26 + (ch.charCodeAt(0) - 64);
  return index;
}

function parseSharedStringsPuc(xml: string): string[] {
  if (!xml) return [];
  const strings: string[] = [];
  const siMatches = xml.match(/<(?:\w+:)?si[\s\S]*?<\/(?:\w+:)?si>/g) || [];
  for (const si of siMatches) {
    const textParts = Array.from(si.matchAll(/<(?:\w+:)?t[^>]*>([\s\S]*?)<\/(?:\w+:)?t>/g)).map((m) => decodeXmlPuc(m[1]));
    strings.push(textParts.join(""));
  }
  return strings;
}

function parseSheetRowsPuc(xml: string, sharedStrings: string[]): string[][] {
  const rows: string[][] = [];
  const rowMatches = xml.match(/<(?:\w+:)?row[^>]*>[\s\S]*?<\/(?:\w+:)?row>/g) || [];
  for (const rowXml of rowMatches) {
    const cells = Array.from(rowXml.matchAll(/<(?:\w+:)?c([^>]*)>([\s\S]*?)<\/(?:\w+:)?c>/g));
    const rowValues: string[] = [];
    let autoColIndex = 0;
    for (const cellMatch of cells) {
      const attrs = cellMatch[1];
      const body = cellMatch[2];
      const refMatch = attrs.match(/r="([A-Z]+)\d+"/);
      let colIndex: number;
      if (refMatch) {
        colIndex = colLettersToIndexPuc(refMatch[1]);
        autoColIndex = colIndex;
      } else {
        colIndex = ++autoColIndex;
      }
      while (rowValues.length < colIndex) rowValues.push("");
      const type = (attrs.match(/t="([^"]+)"/) || [])[1] || "";
      const valueMatch = body.match(/<(?:\w+:)?v>([\s\S]*?)<\/(?:\w+:)?v>/);
      const inlineMatch = body.match(/<(?:\w+:)?t[^>]*>([\s\S]*?)<\/(?:\w+:)?t>/);
      let value = "";
      if (type === "s" && valueMatch) {
        const idx = Number(valueMatch[1]);
        value = Number.isFinite(idx) && sharedStrings[idx] !== undefined ? sharedStrings[idx] : "";
      } else if (inlineMatch) {
        value = decodeXmlPuc(inlineMatch[1]);
      } else if (valueMatch) {
        value = decodeXmlPuc(valueMatch[1]);
      }
      rowValues[colIndex - 1] = value.trim();
    }
    if (rowValues.some((v) => v !== "")) rows.push(rowValues);
  }
  return rows;
}

function tablaDesdeFilasPuc(sheetName: string, rows: string[][]): Tabla {
  if (!rows.length) throw new Error("La primera hoja del PUC está vacía.");
  let headerIdx = -1;
  const limite = Math.min(30, rows.length);
  for (let i = 0; i < limite; i++) {
    const cols = rows[i].map((v) => String(v ?? "").trim()).filter(Boolean);
    const codeCol = columnaPucCodigo(cols);
    if (cols.length && codeCol && columnaPucNombre(cols, codeCol)) {
      headerIdx = i;
      break;
    }
  }
  if (headerIdx < 0) headerIdx = 0;

  const rawHeader = rows[headerIdx] || [];
  const columns: string[] = [];
  const colIndex: number[] = [];
  rawHeader.forEach((h, idx) => {
    let name = String(h ?? "").trim();
    if (!name) return;
    if (columns.includes(name)) name = `${name}_${idx + 1}`;
    columns.push(name);
    colIndex.push(idx);
  });
  if (!columns.length) throw new Error("No se encontró una fila de encabezados en el PUC.");

  const outRows: Tabla["rows"] = [];
  for (let r = headerIdx + 1; r < rows.length; r++) {
    const source = rows[r];
    const obj: Record<string, string> = {};
    let algo = false;
    columns.forEach((col, i) => {
      const value = String(source[colIndex[i]] ?? "").trim();
      obj[col] = value;
      if (value) algo = true;
    });
    if (algo) outRows.push(obj);
  }
  return { sheetName, columns, rows: outRows };
}

async function leerPucXlsxZip(buffer: Buffer): Promise<Tabla> {
  const zip = await JSZip.loadAsync(buffer);
  const sharedStringsXml = await zip.file("xl/sharedStrings.xml")?.async("string");
  const sheetPath = Object.keys(zip.files).find((n) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(n) && !zip.files[n].dir);
  if (!sheetPath) throw new Error("El .xlsx no contiene hojas XML legibles.");
  const sheetXml = await zip.file(sheetPath)?.async("string");
  if (!sheetXml) throw new Error("No se pudo leer la primera hoja XML del .xlsx.");
  const sharedStrings = parseSharedStringsPuc(sharedStringsXml || "");
  const rows = parseSheetRowsPuc(sheetXml, sharedStrings);
  return tablaDesdeFilasPuc(sheetPath.replace(/^xl\/worksheets\//i, "").replace(/\.xml$/i, ""), rows);
}

export async function importarPuc(empresaId: string, buffer: Buffer, filename = "PUC.xlsx"): Promise<CuentaPuc[]> {
  const oid = toObjectId(empresaId);
  if (!oid) throw new Error("Empresa inválida.");
  let tabla: Tabla;
  try {
    tabla = await leerHojaConEncabezadoDinamico(buffer, (columns) => {
      const codeCol = columnaPucCodigo(columns);
      const nameCol = columnaPucNombre(columns, codeCol);
      return !!codeCol && !!nameCol;
    });
  } catch (excelErr) {
    try {
      tabla = await leerPucXlsxZip(buffer);
    } catch (zipErr) {
      throw new Error(
        `No se pudo leer el PUC como Excel .xlsx. Detalle exceljs: ${excelErr instanceof Error ? excelErr.message : String(excelErr)}. Detalle XML: ${zipErr instanceof Error ? zipErr.message : String(zipErr)}`
      );
    }
  }
  const codeCol = columnaPucCodigo(tabla.columns);
  const nameCol = columnaPucNombre(tabla.columns, codeCol);
  const nivelCol = columnaPucNivelAgrupacion(tabla.columns);
  const activoCol = columnaPucActivo(tabla.columns);
  if (!codeCol) throw new Error("No se encontró la columna de código de cuenta en el PUC.");
  if (!nivelCol) throw new Error('No se encontró la columna "Nivel agrupación" en el PUC.');
  if (!activoCol) throw new Error('No se encontró la columna "Activo" en el PUC.');

  const rowsFiltradas = tabla.rows.filter((row) => esPucTransaccional(row[nivelCol]) && esPucActivo(row[activoCol]));
  const byCode = new Map<string, CuentaPuc>();
  for (const row of rowsFiltradas) {
    const codigo = limpiarCodigoCuenta(row[codeCol]);
    if (!codigo) continue;
    const nombre = nameCol ? String(row[nameCol] ?? "").trim() : "";
    if (!byCode.has(codigo)) byCode.set(codigo, { codigo, nombre });
  }
  const cuentas = [...byCode.values()].sort((a, b) => a.codigo.localeCompare(b.codigo, "es", { numeric: true }));
  if (!cuentas.length) throw new Error('No se encontraron cuentas transaccionales activas en el PUC (Nivel agrupación = "Transaccional" y Activo = "Sí").');

  const res = await getDb().collection<any>(EMPRESAS).updateOne(
    { _id: oid },
    { $set: { puc: { cuentas, filename, updatedAt: new Date() }, updatedAt: new Date() } }
  );
  if (res.matchedCount === 0) throw new Error("Empresa no encontrada.");
  return cuentas;
}

export async function getComprobantesEmpresa(empresaId: string): Promise<Comprobantes> {
  const oid = toObjectId(empresaId);
  if (!oid) throw new Error("Empresa inválida.");
  const doc = await getDb().collection<any>(EMPRESAS).findOne({ _id: oid }, { projection: { comprobantes: 1 } });
  if (!doc) throw new Error("Empresa no encontrada.");
  return normComprobantes(doc.comprobantes);
}

/**
 * Lee un movimiento auxiliar de SIIGO (.xlsx) y devuelve el siguiente consecutivo
 * para cada tipo de comprobante indicado. Busca entradas "CC-{tipo}-{n}" en la
 * columna "Comprobante" y retorna max(n) + 1 (o 1 si no hay ninguna).
 */
export async function detectarConsecutivosDesdeAuxiliar(
  buffer: Buffer,
  tipos: string[]
): Promise<Record<string, number>> {
  const tiposValidos = tipos.filter(Boolean);
  if (!tiposValidos.length) return {};
  const tabla = await leerHojaConEncabezadoDinamico(buffer, (cols) =>
    cols.some((c) => normalizarCol(c) === "comprobante")
  );
  const compCol = tabla.columns.find((c) => normalizarCol(c) === "comprobante");
  if (!compCol) return Object.fromEntries(tiposValidos.map((t) => [t, 1]));
  const maxPorTipo: Record<string, number> = Object.fromEntries(tiposValidos.map((t) => [t, 0]));
  for (const row of tabla.rows) {
    const val = String(row[compCol] ?? "").trim();
    const m = /^CC-(\w+)-(\d+)$/i.exec(val);
    if (m && m[1] in maxPorTipo) {
      maxPorTipo[m[1]] = Math.max(maxPorTipo[m[1]], parseInt(m[2], 10));
    }
  }
  return Object.fromEntries(tiposValidos.map((t) => [t, (maxPorTipo[t] || 0) + 1]));
}

export async function getPuc(empresaId: string): Promise<CuentaPuc[]> {
  const oid = toObjectId(empresaId);
  if (!oid) throw new Error("Empresa inválida.");
  const doc = await getDb().collection<any>(EMPRESAS).findOne({ _id: oid }, { projection: { "puc.cuentas": 1 } });
  const cuentas = doc?.puc?.cuentas;
  return Array.isArray(cuentas) ? cuentas.map((c: any) => ({ codigo: String(c.codigo || ""), nombre: String(c.nombre || "") })).filter((c: CuentaPuc) => c.codigo) : [];
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
        /(razon|raz n|proveedor|cliente|nombre|tercero)/.test(normalizarCol(c))
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
  let actualizado = false;
  for (const nit of faltantes) {
    if (existentes.has(nit)) {
      if (nameCol) {
        const filaExistente = tabla.rows.find((r) => limpiarNit(r[nitCol] as unknown) === nit);
        const nombre = nombres.get(nit) || "";
        if (filaExistente && nombre && String(filaExistente[nameCol] ?? "").trim() === "") {
          filaExistente[nameCol] = nombre;
          actualizado = true;
        }
      }
      continue;
    }
    const fila: Record<string, string | number | boolean | null> = {};
    for (const c of tabla.columns) fila[c] = "";
    fila[nitCol] = nit;
    if (nameCol) fila[nameCol] = nombres.get(nit) || "";
    tabla.rows.push(fila);
    existentes.add(nit);
    creados.push(nit);
  }

  if (creados.length || actualizado) await setTabla(empresaId, slot, tabla);
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
