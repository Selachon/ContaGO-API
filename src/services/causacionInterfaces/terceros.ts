/**
 * Proceso de TERCEROS del motor "Causación vía Interfaces".
 *
 * Portado 1:1 desde `script_terceros.py`. Lee la hoja "Datos de terceros" de
 * los archivos DIAN, normaliza cada tercero (dígito de verificación, tipo de
 * identificación, nombres/apellidos, geografía, teléfono, correo) y produce las
 * filas listas para volcar en la plantilla de importación de Siigo.
 */
import {
  CausacionInterfacesError,
  COLUMNAS_TERCEROS,
  type TercerosInput,
  type TercerosResultado,
  type FilaTercero,
} from "./types.js";
import {
  cargarLibro,
  leerHojaConHeaderDinamico,
  type Fila,
  type Tabla,
} from "./excelReader.js";
import { normalizarTexto } from "./normalize.js";

const HOJA_TERCEROS_DIAN = "Datos de terceros";
const HOJA_PAISES = "Paises";

/** Convierte a texto en mayúsculas; cadena vacía si nulo. */
function aMayusculas(valor: unknown): string {
  if (valor === null || valor === undefined) return "";
  return String(valor).trim().toUpperCase();
}

/** Limpia un NIT dejando solo dígitos (criterio del proceso de terceros). */
function limpiarNitTerceros(valor: unknown): string {
  if (valor === null || valor === undefined) return "";
  let s = String(valor).trim();
  if (s.toLowerCase() === "nan") return "";
  s = s.replace(/ /g, "").replace(/-/g, "");
  if (s.endsWith(".0")) s = s.slice(0, -2);
  return s.replace(/[^0-9]/g, "");
}

/** Clasifica el tipo de identificación: 31 (empresa/NIT) o 13 (persona natural). */
function tipoIdentificacionDesdeNit(nit: string): string {
  if (nit.length === 9) {
    const n = Number(nit);
    if (Number.isFinite(n) && n >= 70000000 && n <= 999000000) return "31";
  }
  return "13";
}

/** Dígito de verificación del NIT según el algoritmo de la DIAN. */
export function calcularDigitoVerificacion(nit: string): string {
  const limpio = String(nit).replace(/[^0-9]/g, "");
  if (!limpio) return "";
  const pesos = [3, 7, 13, 17, 19, 23, 29, 37, 41, 43, 47, 53, 59, 67, 71];
  const invertido = limpio.split("").reverse();
  let total = 0;
  for (let i = 0; i < invertido.length; i++) {
    if (i >= pesos.length) return "";
    total += parseInt(invertido[i], 10) * pesos[i];
  }
  const residuo = total % 11;
  if (residuo === 0 || residuo === 1) return String(residuo);
  return String(11 - residuo);
}

// Nombres de pila frecuentes, usados para separar nombres de apellidos.
const NOMBRES_COMUNES = new Set([
  "jose", "antonio", "juan", "manuel", "francisco", "luis", "javier", "miguel", "carlos", "angel",
  "jesus", "david", "daniel", "pedro", "alejandro", "maria", "alberto", "pablo", "fernando", "rafael",
  "jorge", "ramon", "sergio", "enrique", "andres", "diego", "adrian", "vicente", "victor", "alvaro",
  "ignacio", "raul", "eduardo", "ivan", "oscar", "ruben", "joaquin", "santiago", "mario", "roberto",
  "gabriel", "marcos", "alfonso", "jaime", "ricardo", "hugo", "julio", "emilio", "martin", "salvador",
  "guillermo", "mohamed", "nicolas", "tomas", "jordi", "julian", "gonzalo", "agustin", "cristian", "cesar",
  "marc", "felix", "joan", "josep", "samuel", "sebastian", "lucas", "hector", "felipe", "ismael",
  "alfredo", "domingo", "aitor", "alex", "mariano", "rodrigo", "mateo", "alexander", "iker", "marco",
  "xavier", "esteban", "arturo", "gregorio", "lorenzo", "dario", "borja", "albert", "aaron", "joel",
  "isaac", "eugenio", "cristobal", "eric", "jonathan", "christian", "mohammed", "pau", "german", "omar",
  "carmen", "ana", "isabel", "dolores", "pilar", "teresa", "rosa", "josefa", "cristina", "laura",
  "angeles", "elena", "antonia", "lucia", "marta", "francisca", "mercedes", "luisa", "concepcion", "rosario",
  "paula", "sara", "raquel", "rocio", "eva", "patricia", "beatriz", "victoria", "juana", "manuela",
  "julia", "andrea", "belen", "alba", "silvia", "esther", "irene", "nuria", "encarnacion", "montserrat",
  "sandra", "angela", "monica", "alicia", "inmaculada", "yolanda", "mar", "sonia", "marina", "sofia",
  "susana", "margarita", "claudia", "natalia", "carolina", "ines", "alejandra", "daniela", "carla", "veronica",
  "amparo", "gloria", "lourdes", "nieves", "luz", "soledad", "noelia", "lorena", "fatima", "begona",
  "blanca", "olga", "nerea", "miriam", "clara", "consuelo", "asuncion", "milagros", "esperanza", "martina",
  "lidia", "catalina", "adriana", "celia", "anna", "aurora", "magdalena", "emilia", "elisa", "vanesa",
  "ainhoa", "virginia", "eugenia", "diana", "gema", "alexandra", "valeria", "tatiana", "leonor",
]);

/** Separa la razón social de una persona en nombres y apellidos. */
function dividirNombrePersona(razonSocial: string): { nombres: string; apellidos: string } {
  const partes = String(razonSocial).trim().split(/\s+/).filter(Boolean);
  if (partes.length === 0) return { nombres: "", apellidos: "" };
  if (partes.length === 1) return { nombres: partes[0], apellidos: "" };
  if (partes.length === 2) return { nombres: partes[0], apellidos: partes[1] };
  if (partes.length === 3) {
    return { nombres: `${partes[0]} ${partes[1]}`, apellidos: partes[2] };
  }

  const puntuar = (tokens: string[]): number =>
    tokens.filter((x) => NOMBRES_COMUNES.has(normalizarTexto(x))).length;

  const mitad = Math.floor(partes.length / 2);
  const bloqueIzq = partes.slice(0, mitad);
  const bloqueDer = partes.slice(mitad);

  // Si los nombres parecen venir al final (p. ej. CASTRO ABUCHAIBE TATIANA LEONOR), invierte.
  if (puntuar(bloqueDer) > puntuar(bloqueIzq)) {
    return { nombres: bloqueDer.join(" "), apellidos: bloqueIzq.join(" ") };
  }
  // Caso usual en Colombia: NOMBRES + APELLIDOS.
  return {
    nombres: partes.slice(0, 2).join(" "),
    apellidos: partes.slice(-2).join(" "),
  };
}

/** Normaliza un teléfono a máximo 10 dígitos, quitando indicativos. */
function limpiarTelefono(valor: unknown): string {
  if (valor === null || valor === undefined) return "";
  let s = String(valor).replace(/[^0-9]/g, "");
  if (s.length > 10 && s.startsWith("57")) s = s.slice(2);
  const indicativosCiudad = new Set(["601", "602", "603", "604", "605", "606", "607", "608"]);
  while (s.length > 10 && indicativosCiudad.has(s.slice(0, 3))) s = s.slice(3);
  while (s.length > 10 && s.startsWith("0")) s = s.slice(1);
  if (s.length > 10) s = s.slice(-10);
  return s;
}

/** Extrae el primer correo electrónico válido de una celda. */
function limpiarCorreo(valor: unknown): string {
  if (valor === null || valor === undefined) return "";
  const s = String(valor).trim().toLowerCase();
  if (!s) return "";
  const patron = /^[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}$/;
  for (const candidato of s.split(/[;,\s]+/)) {
    if (patron.test(candidato.trim())) return candidato.trim();
  }
  return "";
}

/** Normaliza un texto geográfico para cruzarlo con la hoja de países. */
function normalizarGeoTexto(valor: unknown): string {
  let t = normalizarTexto(valor);
  t = t.replace(/[^a-z0-9\s]/g, " ");
  t = t.split(/\s+/).filter(Boolean).join(" ");
  if (t.includes("bogota")) return "bogota";
  if (t.includes("cartagena")) return "cartagena";
  const reemplazos: Record<string, string> = {
    "bogota d c": "bogota",
    "bogota dc": "bogota",
    "bogota distrito capital": "bogota",
    "bogota d.c": "bogota",
    "bogota d. c.": "bogota",
    "cartagena de indias": "cartagena",
    "distrito capital": "bogota",
  };
  return reemplazos[t] ?? t;
}

/** Índices geográficos construidos desde la hoja "Paises" de la plantilla. */
interface DiccionarioGeo {
  /** (pais|dep|ciudad) → [codDep, codCiudad, codPais] */
  idxCiudad: Map<string, [string, string, string]>;
  /** (pais|dep) → [codDep, codPais] */
  idxDep: Map<string, [string, string]>;
  /** (pais|ciudad) → [codDep, codCiudad, codPais] */
  idxCiudadPais: Map<string, [string, string, string]>;
}

function textoCelda(valor: unknown): string {
  return valor === null || valor === undefined ? "" : String(valor).trim();
}

/** Construye los índices geográficos a partir de la hoja "Paises". */
function construirDiccionarioGeo(wbPlantilla: import("exceljs").Workbook): DiccionarioGeo {
  const paises = leerHojaConHeaderDinamico(wbPlantilla, HOJA_PAISES, [
    "estado / departamento",
    "codigo ciudad",
  ]);

  const colsNorm = new Map<string, string>();
  for (const c of paises.columnas) colsNorm.set(normalizarTexto(c), c);
  const cPais = colsNorm.get("pais");
  const cDep = colsNorm.get("estado / departamento");
  const cCiudad = colsNorm.get("ciudad");
  const cCodPais = colsNorm.get("codigo pais");
  const cCodDep = colsNorm.get("codigo estado / departamento");
  const cCodCiudad = colsNorm.get("codigo ciudad");

  const faltan = [
    ["País", cPais],
    ["Estado / Departamento", cDep],
    ["Ciudad", cCiudad],
    ["Código país", cCodPais],
    ["Código Estado / Departamento", cCodDep],
    ["Código ciudad", cCodCiudad],
  ].filter(([, v]) => v === undefined).map(([k]) => k);
  if (faltan.length > 0) {
    throw new CausacionInterfacesError(
      `Faltan columnas en la hoja "Paises" de la plantilla: ${faltan.join(", ")}.`,
      422,
      "paises_columnas_faltantes",
      { columnas: faltan }
    );
  }

  const idxCiudad = new Map<string, [string, string, string]>();
  const idxDep = new Map<string, [string, string]>();
  const idxCiudadPais = new Map<string, [string, string, string]>();

  for (const r of paises.filas) {
    const kPais = normalizarGeoTexto(r[cPais as string]);
    const kDep = normalizarGeoTexto(r[cDep as string]);
    const kCiudad = normalizarGeoTexto(r[cCiudad as string]);
    const codPais = textoCelda(r[cCodPais as string]);
    const codDep = textoCelda(r[cCodDep as string]);
    const codCiudad = textoCelda(r[cCodCiudad as string]);

    const kc = `${kPais}|${kDep}|${kCiudad}`;
    const kd = `${kPais}|${kDep}`;
    const kcp = `${kPais}|${kCiudad}`;
    if (!idxCiudad.has(kc)) idxCiudad.set(kc, [codDep, codCiudad, codPais]);
    if (!idxDep.has(kd)) idxDep.set(kd, [codDep, codPais]);
    if (!idxCiudadPais.has(kcp)) idxCiudadPais.set(kcp, [codDep, codCiudad, codPais]);
  }

  return { idxCiudad, idxDep, idxCiudadPais };
}

/** Lee la hoja "Datos de terceros" de un archivo DIAN. */
function leerTercerosDesdeExcel(wb: import("exceljs").Workbook): Tabla {
  return leerHojaConHeaderDinamico(wb, HOJA_TERCEROS_DIAN, ["nit", "razon social", "ciudad"]);
}

/** Transforma las filas DIAN de terceros en filas para la plantilla de Siigo. */
function mapearTerceros(filas: Fila[], columnas: string[], geo: DiccionarioGeo): FilaTercero[] {
  const colsNorm = new Map<string, string>();
  for (const c of columnas) colsNorm.set(normalizarTexto(c), c);
  const cNit = colsNorm.get("nit");
  const cRazon = colsNorm.get("razon social");
  const cDir = colsNorm.get("direccion");
  const cPais = colsNorm.get("pais");
  const cDep = colsNorm.get("departamento");
  const cCiudad = colsNorm.get("ciudad");
  const cTel = colsNorm.get("telefono");
  const cMail = colsNorm.get("correo");

  if (!cNit || !cRazon) {
    throw new CausacionInterfacesError(
      'La hoja "Datos de terceros" debe incluir las columnas NIT y Razón social.',
      422,
      "terceros_columnas_faltantes"
    );
  }

  const salida: FilaTercero[] = [];
  for (const r of filas) {
    const nit = limpiarNitTerceros(r[cNit]);
    if (!nit) continue;
    const razon = aMayusculas(r[cRazon]);
    if (!razon) continue;

    const tipoId = tipoIdentificacionDesdeNit(nit);
    const tipo = tipoId === "31" ? "Empresa" : "Es persona";

    let razonSocial = "";
    let nombres = "";
    let apellidos = "";
    if (tipoId === "31") {
      razonSocial = aMayusculas(razon);
    } else {
      const partido = dividirNombrePersona(razon);
      nombres = aMayusculas(partido.nombres);
      apellidos = aMayusculas(partido.apellidos);
    }

    let pais = "Colombia";
    if (cPais && r[cPais] !== null && r[cPais] !== undefined) {
      pais = String(r[cPais]).trim() || "Colombia";
    }
    const dep = cDep && r[cDep] !== null && r[cDep] !== undefined ? String(r[cDep]).trim() : "";
    const ciudad =
      cCiudad && r[cCiudad] !== null && r[cCiudad] !== undefined ? String(r[cCiudad]).trim() : "";

    const kPais = normalizarGeoTexto(pais);
    const kDep = normalizarGeoTexto(dep);
    const kCiudad = normalizarGeoTexto(ciudad);

    let codPais = kPais === "colombia" ? "COL" : "";
    let codDep = "";
    let codCiudad = "";

    const porCiudad = geo.idxCiudad.get(`${kPais}|${kDep}|${kCiudad}`);
    const porCiudadPais = geo.idxCiudadPais.get(`${kPais}|${kCiudad}`);
    const porDep = geo.idxDep.get(`${kPais}|${kDep}`);
    if (porCiudad) {
      [codDep, codCiudad] = [porCiudad[0], porCiudad[1]];
      if (porCiudad[2]) codPais = porCiudad[2];
    } else if (porCiudadPais) {
      [codDep, codCiudad] = [porCiudadPais[0], porCiudadPais[1]];
      if (porCiudadPais[2]) codPais = porCiudadPais[2];
    } else if (porDep) {
      codDep = porDep[0];
      if (porDep[1]) codPais = porDep[1];
    }

    const direccion =
      cDir && r[cDir] !== null && r[cDir] !== undefined ? String(r[cDir]).trim() : "";
    const telefono = cTel ? limpiarTelefono(r[cTel]) : "";
    const correo = cMail ? limpiarCorreo(r[cMail]) : "";

    salida.push({
      "Identificación (Obligatorio)": nit,
      "Dígito de verificación": calcularDigitoVerificacion(nit),
      "Tipo identificación (Obligatorio)": tipoId,
      "Tipo (Obligatorio)": tipo,
      "Razón social (Obligatorio)": razonSocial,
      "Nombres del tercero (Obligatorio)": nombres,
      "Apellidos del tercero (Obligatorio)": apellidos,
      Dirección: direccion,
      "Código país": codPais,
      "Código departamento/estado": codDep,
      "Código ciudad": codCiudad,
      "Teléfono principal": telefono,
      "Tipo de régimen IVA": tipoId === "31" ? "2" : "0",
      "Correo electrónico contacto principal": correo,
      Clientes: "",
      Estado: "Activo",
    });
  }

  // Elimina duplicados por identificación, conservando el primero.
  const vistos = new Set<string>();
  return salida.filter((t) => {
    const id = t["Identificación (Obligatorio)"];
    if (vistos.has(id)) return false;
    vistos.add(id);
    return true;
  });
}

/** Ejecuta el proceso de terceros: produce las filas para la plantilla de Siigo. */
export async function runTerceros(input: TercerosInput): Promise<TercerosResultado> {
  if (!input.dian || input.dian.length === 0) {
    throw new CausacionInterfacesError(
      "No se proporcionó ningún archivo DIAN para extraer terceros.",
      400,
      "sin_archivos_dian"
    );
  }

  const wbPlantilla = await cargarLibro(input.plantilla);
  const geo = construirDiccionarioGeo(wbPlantilla);

  const filasDian: Fila[] = [];
  const columnas = new Set<string>();
  for (const buffer of input.dian) {
    const wb = await cargarLibro(buffer);
    if (!wb.getWorksheet(HOJA_TERCEROS_DIAN)) continue;
    const tabla = leerTercerosDesdeExcel(wb);
    for (const c of tabla.columnas) columnas.add(c);
    filasDian.push(...tabla.filas);
  }

  if (filasDian.length === 0) {
    throw new CausacionInterfacesError(
      'Ninguno de los archivos DIAN contiene la hoja "Datos de terceros".',
      422,
      "hoja_terceros_no_encontrada"
    );
  }

  const terceros = mapearTerceros(filasDian, [...columnas], geo);
  const empresas = terceros.filter((t) => t["Tipo identificación (Obligatorio)"] === "31").length;

  return {
    terceros,
    resumen: {
      terceros: terceros.length,
      empresas,
      personas: terceros.length - empresas,
    },
  };
}

/** Columnas que el motor rellena (reexportado para conveniencia). */
export { COLUMNAS_TERCEROS };
