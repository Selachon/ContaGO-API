/**
 * Archivo de terceros de Siigo Pyme («MODELO TERCEROS», 76 columnas), con los datos que ya trae cada factura del proveedor.
 * La estructura y los valores por defecto salen de un archivo real de la Fundación Corazón Verde: texto con relleno de
 * espacios al ancho de cada campo, país 169 (Colombia) y códigos de una letra para el tipo de identificación y la clase.
 */
import ExcelJS from "exceljs";
import { calcularDigitoVerificacion } from "./terceros.js";
import type { FilaTercero } from "./types.js";

const ENCABEZADOS = ["IDENTIFICACIÓN  (OBLIGATORIO)", "SUCURSAL  (OBLIGATORIO)", "DIGITO DE VERIFICACIÓN", "NOMBRE", "RAZÓN SOCIAL", "PRIMER NOMBRE", "SEGUNDO NOMBRE", "PRIMER APELLIDO", "SEGUNDO APELLIDO", "NÚMERO DE IDENTIFICACIÓN DEL EXTRANJERO", "CÓDIGO IDENTIFICACIÓN FISCAL", "NOMBRE DEL CONTACTO", "DIRECCIÓN", "PAÍS", "CIUDAD", "ACTIVO", "TELÉFONO 1", "TELÉFONO 2", "TELÉFONO 3", "TELÉFONO 4", "TELÉFONO CELULAR", "FAX", "APARTADO AÉREO", "SEXO", "AÑO DE CUMPLEAÑOS", "MES DE CUMPLEAÑOS", "DÍA DE CUMPLEAÑOS", "TIPO DE PERSONA", "CORREO ELECTRÓNICO", "CONTACTO DE FACTURACIÓN", "CORREO ELECT. CONTACTO DE FACTURACIÓN", "TIPO DE IDENTIFICACIÓN", "CLASIFICACIÓN - CLASE DE TERCERO", "DECLARANTE", "AUTO RETENEDOR", "AGENTE RETENEDOR", "BENEFICIO DIAN RETEIVA COMPRAS", "TARIFA DIFERENCIAL RETE IVA VENTAS", "PORCENTAJE DIFERENCIAL RETE IVA VENTAS", "TARIFA DIFERENCIAL RETE IVA COMPRAS", "PORCENTAJE DIFERENCIAL RETE IVA COMPRAS", "CUPO DE CRÉDITO", "LISTA DE PRECIO", "FORMA DE PAGO", "CALIFICACIÓN", "TIPO CONTRIBUYENTE", "CÓDIGO ACTIVIDAD ECONÓMICA", "VENDEDOR", "COBRADOR", "PORCENTAJE DESCUENTO EN VENTAS", "PERÍODO DE PAGO", "OBSERVACIÓN", "DÍAS OPTIMISTA", "DÍAS PESIMISTA", "CÓDIGO TIPO DE EMPRESA", "CÓDIGO DE BANCO", "CÓDIGO INTERNO", "CÓDIGO OFICINA", "TIPO DE CUENTA", "NÚMERO DE CUENTA", "NIT DEL TITULAR DE LA CUENTA", "DÍGITO DE VERIFICACIÓN TITULAR DE LA CUENTA", "NOMBRE DEL TITULAR DE LA CUENTA", "PAÍS DE LA CUENTA", "CIUDAD DE LA CUENTA", "SIGLAS DEPARTAMENTO DE LA CUENTA", "APLICA RETENCIÓN ICA FACTURA DE VENTA/DEVOLUCIÓN", "APLICA RETENCIÓN ICA FACTURA DE COMPRA/DEVOLUCIÓN", "ACEPTA ENVÍO FACTURA POR MEDIO ELECTRÓNICO", "NOMBRE COMERCIAL", "CODIGO POSTAL", "RESPONSABILIDAD FISCAL", "AÑO APERTURA", "MES APERTURA", "DIA APERTURA", "TRIBUTOS"];

const pad = (v: unknown, n: number) => String(v ?? "").replace(/\s+/g, " ").trim().slice(0, n).padEnd(n, " ");
const primero = (t: string) => { const p = String(t || "").trim().split(/\s+/).filter(Boolean); return [p[0] ?? "", p.slice(1).join(" ")] as const; };
const num = (v: unknown) => { const n = Number(String(v ?? "").replace(/\D/g, "")); return Number.isFinite(n) ? n : 0; };

/** Ciudades de Colombia en Siigo Pyme (listado «Modelo ciudades» de la Fundación Corazón Verde): código DANE de 5 dígitos → código de ciudad de Siigo Pyme. */
const CIUDADES_PYME: Record<string, number> = {
  "11001": 1,
  "05001": 2,
  "08001": 3,
  "13001": 4,
  "15001": 5,
  "17001": 6,
  "18001": 7,
  "19001": 8,
  "20001": 9,
  "23001": 10,
  "27001": 11,
  "41001": 12,
  "44001": 13,
  "47001": 14,
  "50001": 15,
  "52001": 16,
  "54001": 17,
  "63001": 18,
  "66001": 19,
  "68001": 20,
  "70001": 21,
  "73001": 22,
  "76001": 23,
  "81001": 24,
  "85001": 25,
  "86001": 26,
  "88001": 27,
  "91001": 28,
  "94001": 29,
  "95001": 30,
  "97001": 31,
  "99001": 32,
  "05088": 50,
  "05172": 67,
  "50266": 78,
  "25797": 79,
  "05360": 90,
  "05380": 95,
  "20011": 100,
  "10911": 109,
  "11311": 113,
  "05615": 116,
  "05631": 118,
  "05129": 129,
  "08141": 130,
  "05154": 154,
  "08573": 169,
  "08758": 175,
  "25200": 200,
  "50226": 226,
  "15097": 229,
  "15238": 238,
  "54244": 244,
  "25245": 245,
  "81140": 248,
  "26840": 268,
  "25269": 269,
  "68276": 276,
  "05282": 282,
  "15299": 299,
  "15759": 319,
  "17380": 351,
  "35680": 356,
  "17486": 357,
  "36486": 364,
  "37786": 377,
  "40086": 400,
  "19698": 415,
  "19785": 420,
  "19807": 421,
  "23417": 459,
  "25035": 479,
  "25175": 492,
  "25214": 497,
  "25286": 505,
  "25290": 507,
  "25307": 512,
  "52007": 520,
  "25377": 524,
  "25430": 531,
  "25473": 534,
  "25535": 545,
  "25754": 563,
  "86568": 568,
  "25572": 572,
  "25799": 573,
  "25898": 591,
  "25899": 592,
  "27073": 596,
  "60673": 606,
  "60773": 607,
  "41615": 615,
  "61615": 616,
  "41298": 632,
  "08634": 634,
  "41551": 646,
  "41668": 649,
  "41770": 651,
  "44430": 667,
  "68630": 686,
  "50006": 702,
  "25718": 718,
  "76736": 736,
  "19743": 743,
  "75836": 758,
  "05313": 767,
  "76913": 769,
  "25785": 785,
  "25793": 793,
  "54518": 817,
  "19845": 845,
  "66682": 854,
  "68081": 861,
  "68132": 865,
  "52835": 870,
  "17873": 873,
  "54874": 874,
  "86885": 885,
  "68307": 890,
  "68406": 902,
  "68547": 915,
  "68773": 934,
  "68855": 937,
  "70771": 964,
  "73449": 999,
  "73585": 1000,
  "76147": 1024,
  "76233": 1025,
  "76622": 1042,
  "76834": 1047,
  "76892": 1052,
  "86320": 1079,
  "13430": 3430,
  "54498": 4498,
  "25126": 5126,
  "15572": 5572,
  "66170": 6170,
  "76248": 6248,
  "47053": 7053,
  "08638": 8638,
  "19548": 9548,
};
/** Si la ciudad del XML no está en el listado de Siigo Pyme se usa Bogotá (1), la de la mayoría de los terceros. */
const CIUDAD_POR_DEFECTO = 1;
export const ciudadPyme = (codigoDane: unknown): number | null => CIUDADES_PYME[String(codigoDane ?? "").replace(/\D/g, "").slice(0, 5)] ?? null;
const PAIS_COLOMBIA = 169;

/** Una fila del archivo a partir de la fila de terceros que ya arma el proceso para Siigo. */
export function filaTerceroPyme(t: FilaTercero, hoy: Date): (string | number)[] {
  const nit = String(t["Identificación (Obligatorio)"] || "").replace(/\D/g, "");
  const empresa = t["Tipo (Obligatorio)"] === "Empresa";
  const [nombre1, nombre2] = primero(t["Nombres del tercero (Obligatorio)"]);
  const [apellido1, apellido2] = primero(t["Apellidos del tercero (Obligatorio)"]);
  const razon = empresa ? t["Razón social (Obligatorio)"] : [apellido1, apellido2, nombre1, nombre2].filter(Boolean).join(" ");
  const correo = String(t["Correo electrónico contacto principal"] || "").toUpperCase();
  const dv = t["Dígito de verificación"] || calcularDigitoVerificacion(nit);
  return [
    num(nit), 0, num(dv), pad(razon, 100), empresa ? "S" : "N",
    pad(empresa ? "" : nombre1, 25), pad(empresa ? "" : nombre2, 25), pad(empresa ? "" : apellido1, 25), pad(empresa ? "" : apellido2, 25),
    pad("", 20), " ", pad(razon, 50), pad(t["Dirección"], 100), PAIS_COLOMBIA, ciudadPyme(t["Código ciudad"]) ?? CIUDAD_POR_DEFECTO, "A",
    num(t["Teléfono principal"]), 0, 0, 0, 0, 0, 0, "E", 0, 0, 0, empresa ? 2 : 1,
    pad(correo, 100), pad(razon, 50), pad(correo, 90), empresa ? "N" : "C", "P", "N", "N", "N", " ", "N", 0, "N", 0,
    0, 0, 201, 1, 4, 0, 0, 0, 0, 0, pad("", 30), 0, 0, 0, 0, 0, 0, " ", pad("", 33), 0, 0, pad("", 80), 0, 0, pad("", 4),
    "N", "N", " ", pad(razon, 100), 110111, "R-99-PN", hoy.getFullYear(), hoy.getMonth() + 1, hoy.getDate(), "'01",
  ];
}

export async function escribirTercerosPyme(filas: FilaTercero[], empresaNombre: string, hoy: Date = new Date()): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Hoja1");
  ws.addRow([String(empresaNombre || "").toUpperCase()]);
  ws.addRow(["MODELO TERCEROS"]);
  ws.addRow([]);
  ws.addRow([]);
  ws.addRow(ENCABEZADOS);
  for (const f of filas) ws.addRow(filaTerceroPyme(f, hoy));
  return Buffer.from(await wb.xlsx.writeBuffer());
}

/** Identificaciones (sin dígito de verificación) de un «Modelo terceros» exportado de Siigo Pyme: son los terceros que ya existen. */
export async function leerTercerosExistentesPyme(buffer: Buffer): Promise<string[]> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer as any);
  const ws = wb.worksheets[0];
  if (!ws) throw new Error("El archivo no tiene hojas.");
  const norm = (v: unknown) => String(v ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();
  let cabecera = 0;
  for (let r = 1; r <= Math.min(15, ws.rowCount); r++) {
    if (norm(ws.getRow(r).getCell(1).value).includes("IDENTIFICACION")) { cabecera = r; break; }
  }
  if (!cabecera) throw new Error("No encontré la columna «IDENTIFICACIÓN» en las primeras filas: sube el «Modelo terceros» exportado de Siigo Pyme.");
  const nits = new Set<string>();
  for (let r = cabecera + 1; r <= ws.rowCount; r++) {
    const v = ws.getRow(r).getCell(1).value;
    const digitos = String(typeof v === "object" && v !== null && "result" in (v as object) ? (v as { result: unknown }).result : v ?? "").replace(/\D/g, "");
    if (digitos) nits.add(digitos.replace(/^0+/, "") || "0");
  }
  if (!nits.size) throw new Error("El archivo no trae terceros.");
  return [...nits];
}
