const MONTHS_ES = [
  "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
  "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
];

export type ZipFileKind = "PDF" | "XML" | "ZIP";

/**
 * Carpeta base dentro del ZIP de entrega, común a "Descarga Masiva DIAN" y
 * "Descarga masiva + Excel DIAN": <año>/<MM. Mes>/<Emitidas|Recibidas>.
 * Acepta fecha ISO (YYYY-MM-DD) o dd/mm/aaaa; si no se puede leer, SinAño/SinMes.
 */
export function zipBaseFolder(issueDate: string | undefined, direction: "sent" | "received"): string {
  let year = "SinAño";
  let month = "SinMes";
  const d = issueDate || "";
  let y = "", m = 0;
  if (/^\d{4}-\d{2}-\d{2}/.test(d)) {
    const p = d.split("-");
    y = p[0]; m = parseInt(p[1], 10);
  } else if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(d)) {
    const p = d.split("/");
    y = p[2].slice(0, 4); m = parseInt(p[1], 10);
  }
  if (y && m >= 1 && m <= 12) {
    year = y;
    month = `${String(m).padStart(2, "0")}. ${MONTHS_ES[m - 1]}`;
  }
  return `${year}/${month}/${direction === "sent" ? "Emitidas" : "Recibidas"}`;
}

export function zipEntryPath(
  issueDate: string | undefined,
  direction: "sent" | "received",
  kind: ZipFileKind,
  fileName: string,
): string {
  return `${zipBaseFolder(issueDate, direction)}/${kind}/${fileName}`;
}

const ES_MONTHS_SHORT = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];

function toISODate(d: string | undefined): string {
  const v = d || "";
  if (/^\d{4}-\d{2}-\d{2}/.test(v)) return v.slice(0, 10);
  const m = v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return m ? `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}` : "";
}

function fmtDateES(iso: string): string {
  const [y, m, d] = iso.split("-");
  return `${ES_MONTHS_SHORT[parseInt(m, 10) - 1] || m} ${d} ${y}`;
}

/**
 * Nombre del ZIP de entrega, igual al de "Descarga masiva + Excel DIAN":
 * "<NIT> - Facturas <Recibidas|Emitidas> DIAN <Mes dd aaaa - Mes dd aaaa>.zip".
 * El rango sale de las fechas de emisión de los documentos descargados.
 */
export function buildDeliveryZipName(nit: string, direction: "sent" | "received", issueDates: Array<string | undefined>): string {
  const isos = issueDates.map(toISODate).filter(Boolean).sort();
  const first = isos[0], last = isos[isos.length - 1];
  const range = !first ? new Date().toISOString().slice(0, 10)
    : first === last ? fmtDateES(first)
    : `${fmtDateES(first)} - ${fmtDateES(last)}`;
  const safeNit = (nit || "").replace(/[^a-zA-Z0-9]/g, "") || "SinNIT";
  return `${safeNit} - Facturas ${direction === "sent" ? "Emitidas" : "Recibidas"} DIAN ${range}.zip`;
}
