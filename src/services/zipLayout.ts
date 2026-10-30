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
