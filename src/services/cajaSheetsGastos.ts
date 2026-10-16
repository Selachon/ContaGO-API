/**
 * Pestaña "Gastos Oficina" del Sheet de las socias: egresos del proyecto Oficina
 * por categoría (filas) y por mes (columnas). Cada categoría es una fila con su
 * total mensual; debajo, plegado, el detalle de los movimientos con el valor en
 * la columna de su mes. La escribe el portal (una vía) en cada sincronización.
 */
import { getDb } from "./database.js";
import { getSheetsClient } from "./cajaSheetsAuth.js";
import type { SyncProyectoRow } from "./cajaSheetsSync.js";

const COL = "cajaErpSheets";
export const GASTOS_SHEET = "Gastos Oficina";
const MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];

// Orden de presentación (las que no tengan movimientos no se muestran).
const ORDEN = [
  "Nómina oficina", "Seguridad social", "Honorarios contabilidad", "Arriendo", "Seguros", "Impuestos",
  "Gastos bancarios", "Cámara de comercio", "Licencias software", "Cesantías", "Préstamos Sylvia – CC SAS",
  "Viajes", "Sin clasificar", "Seguridad social base", "Seguridad social adicional",
];
const SIN = "Sin clasificar";

// Secciones del gasto de oficina. Por ahora solo Sylvia tiene partes propias; María Claudia
// aparecerá cuando se le asignen conceptos.
const SEC_COMUN = "Gastos comunes";
const SEC_SYLVIA = "Sylvia";
const SECCIONES = [SEC_COMUN, SEC_SYLVIA, "María Claudia"];
// Costos mensuales fijos (tabla auxiliar al pie). Valores de referencia tomados del último pago normal de cada concepto.
const FIJOS_COMUNES: [string, number, string][] = [
  ["Nómina oficina", 3807946, "Nómina de sep-26"],
  ["Seguridad social (sin la parte de Sylvia)", 1447500, "Planilla de sep-26 menos la parte de Sylvia"],
  ["Arriendo", 1838594, "Canon mensual jun/jul-26"],
  ["Honorarios contabilidad", 1607010, "Cuota mensual"],
  ["Seguros (Metlife, 2 pólizas)", 967035, "485.660 + 481.375"],
];
const FIJOS_SYLVIA: [string, number, string][] = [
  ["Seguridad social base", 430000, "Fija todos los meses"],
  ["Seguridad social adicional", 528800, "Desde ago-26"],
];
const SS_SYLVIA_BASE = 430000;
const SS_SYLVIA_ADICIONAL = 528800;
const SS_SYLVIA_ADICIONAL_DESDE = "2026-08"; // mes en que subió la planilla de seguridad social

const norm = (s: string) => s.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");

/** Unifica las categorías (escritas de varias formas) en las de las socias. */
export function categoriaGasto(categoria: string, descripcion: string, valor: number): string {
  const t = norm(`${categoria} ${descripcion}`);
  if (/nomina|rp-1-182\b|prestamos? sylvia/.test(t)) return "Nómina oficina";
  if (/seguridad social|seg social|planilla/.test(t)) return "Seguridad social";
  if (/honorarios contab|contago sas/.test(t)) return "Honorarios contabilidad";
  if (/arriendo/.test(t)) return "Arriendo";
  if (/seguros|metlife/.test(t) || (/debito a solicitud de terceros/.test(t) && Math.round(valor) === 485660)) return "Seguros";
  if (/impuesto|\bica\b/.test(t)) return "Impuestos";
  if (/gastos? bancari/.test(t)) return "Gastos bancarios";
  if (/camara/.test(t)) return "Cámara de comercio";
  if (/licencia/.test(t)) return "Licencias software";
  if (/cesant/.test(t)) return "Cesantías";
  if (/viaje/.test(t)) return "Viajes";
  return SIN;
}

const txt = (s: unknown) => { const v = String(s ?? ""); return /^[=+\-@]/.test(v) ? `'${v}` : v; };
const colL = (i: number) => { let n = i, s = ""; do { s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) - 1; } while (n >= 0); return s; };

export async function syncGastosOficina(companyId: string, rows: SyncProyectoRow[]): Promise<{ categorias: number; movimientos: number }> {
  const d = await getDb().collection<any>(COL).findOne({ companyId });
  const spreadsheetId: string = d?.spreadsheetId;
  if (!spreadsheetId) throw new Error("Primero indica el enlace del Google Sheet.");
  const sheets = await getSheetsClient(companyId);
  if (!sheets) throw new Error("Google Sheets no está autorizado.");

  const oficina = rows.find((r) => r.proyecto.esTransversal || /oficina/i.test(r.proyecto.nombre));
  const egresos = (oficina?.items || []).filter((i) => i.direction === "out" && Number(i.valor) > 0);

  // Meses desde el primero hasta el último con gasto (contiguos).
  const keys = egresos.map((e) => String(e.fecha || "").slice(0, 7)).filter((k) => /^\d{4}-\d{2}$/.test(k)).sort();
  const months: string[] = [];
  if (keys.length) {
    let [y, m] = keys[0].split("-").map(Number);
    const [y1, m1] = keys[keys.length - 1].split("-").map(Number);
    while (y < y1 || (y === y1 && m <= m1)) { months.push(`${y}-${String(m).padStart(2, "0")}`); if (++m > 12) { m = 1; y++; } }
  }
  const mLabel = (k: string) => `${MESES[Number(k.slice(5, 7)) - 1]}-${k.slice(2, 4)}`;
  const nM = months.length;
  const cMonth = (i: number) => colL(1 + i);          // B.. = meses
  const cTot = colL(1 + nM);                          // columna Total
  const lastCol = 1 + nM;                             // índice (0-based) de Total

  // Reparte cada egreso en partes por socia. Hoy solo se separa la seguridad social de Sylvia
  // (base fija + adicional desde que subió la planilla); lo demás es gasto común.
  type Pieza = { fecha: unknown; descripcion: string; valor: number };
  const bySec = new Map<string, Map<string, Pieza[]>>();
  const add = (sec: string, cat: string, p: Pieza) => {
    if (!bySec.has(sec)) bySec.set(sec, new Map());
    const m = bySec.get(sec)!;
    if (!m.has(cat)) m.set(cat, []);
    m.get(cat)!.push(p);
  };
  for (const e of egresos) {
    const c = categoriaGasto(e.categoria || "", e.descripcion || "", Number(e.valor));
    let resto = Number(e.valor);
    if (c === "Seguridad social") {
      const base = Math.min(resto, SS_SYLVIA_BASE);
      if (base > 0) { add(SEC_SYLVIA, "Seguridad social base", { fecha: e.fecha, descripcion: `${e.descripcion || ""} (parte de Sylvia)`, valor: base }); resto -= base; }
      if (String(e.fecha).slice(0, 7) >= SS_SYLVIA_ADICIONAL_DESDE) {
        const ad = Math.min(resto, SS_SYLVIA_ADICIONAL);
        if (ad > 0) { add(SEC_SYLVIA, "Seguridad social adicional", { fecha: e.fecha, descripcion: `${e.descripcion || ""} (parte de Sylvia)`, valor: ad }); resto -= ad; }
      }
    }
    if (resto > 0) add(SEC_COMUN, c, { fecha: e.fecha, descripcion: e.descripcion || "", valor: resto });
  }

  // ── Contenido ──
  const out: any[][] = [];
  const set = (row: number, vals: any[]) => { out[row - 1] = vals; };
  const HEAD = 4, TOT = 5;
  let cur = TOT + 1;
  type G = { c: string; h: number; d0: number; d1: number; items: Pieza[] };
  type S = { name: string; h: number; groups: G[] };
  const secs: S[] = [];
  const groups: G[] = [];
  for (const name of SECCIONES) {
    const m = bySec.get(name);
    if (!m) continue;
    const orden = [...ORDEN.filter((c) => m.has(c)), ...[...m.keys()].filter((c) => !ORDEN.includes(c))];
    const sec: S = { name, h: cur++, groups: [] };
    for (const c of orden) {
      const items = [...m.get(c)!].sort((a, b) => String(a.fecha).localeCompare(String(b.fecha)));
      const h = cur, d0 = h + 1, d1 = h + items.length;
      const g = { c, h, d0, d1, items };
      sec.groups.push(g); groups.push(g);
      cur = d1 + 1;
    }
    secs.push(sec);
  }
  // Tabla auxiliar de costos mensuales fijos (al pie)
  const FJ = cur + 2;
  const fjRows: { h: number; d0: number; d1: number; total: number }[] = [];
  set(FJ, ["COSTOS MENSUALES FIJOS", "Valor mensual", "Base"]);
  let fr = FJ + 1;
  for (const [titulo, lista] of [["Gastos comunes", FIJOS_COMUNES], ["Sylvia", FIJOS_SYLVIA]] as [string, [string, number, string][]][]) {
    const h = fr++, d0 = fr, d1 = fr + lista.length - 1;
    set(h, [titulo, `=SUM(B${d0}:B${d1})`]);
    lista.forEach(([n, v, b], i) => set(d0 + i, [n, v, b]));
    fr = d1 + 1;
    fjRows.push({ h, d0, d1, total: 0 });
  }
  const FJ_TOT = fr;
  set(FJ_TOT, ["TOTAL FIJOS AL MES", `=${fjRows.map((x) => `B${x.h}`).join("+")}`]);
  const lastRow = FJ_TOT;

  set(1, ["Gastos de oficina por categoría y mes"]);
  set(2, ["Egresos del proyecto Oficina, según el portal. Usa el + del margen izquierdo para ver el detalle de cada categoría."]);
  set(HEAD, ["Categoría", ...months.map(mLabel), "Total"]);
  const sumRefs = (col: string, hs: number[]) => (hs.length ? `=SUM(${hs.map((h) => `${col}${h}`).join(";")})` : "=0");
  set(TOT, ["TOTAL GASTOS", ...months.map((_, i) => sumRefs(cMonth(i), secs.map((x) => x.h))), `=SUM(B${TOT}:${colL(nM)}${TOT})`]);
  for (const sec of secs) {
    set(sec.h, [sec.name.toUpperCase(), ...months.map((_, i) => sumRefs(cMonth(i), sec.groups.map((g) => g.h))), `=SUM(B${sec.h}:${colL(nM)}${sec.h})`]);
  }
  for (const g of groups) {
    set(g.h, [g.c, ...months.map((_, i) => `=SUM(${cMonth(i)}${g.d0}:${cMonth(i)}${g.d1})`), `=SUM(B${g.h}:${colL(nM)}${g.h})`]);
    g.items.forEach((it, i) => {
      const row = new Array(lastCol + 1).fill("");
      const mi = months.indexOf(String(it.fecha).slice(0, 7));
      row[0] = txt(`${String(it.fecha).slice(0, 10)} · ${it.descripcion || ""}`);
      if (mi >= 0) row[1 + mi] = Math.round(it.valor);
      set(g.d0 + i, row);
    });
  }
  for (let i = 0; i < lastRow; i++) if (!out[i]) out[i] = [];

  // ── Pestaña: crear si falta, limpiar grupos/formato ──
  const getMeta = () => sheets.spreadsheets.get({ spreadsheetId, fields: "sheets(properties(sheetId,title,gridProperties),rowGroups)" });
  let meta = await getMeta();
  if (!meta.data.sheets?.some((s) => s.properties?.title === GASTOS_SHEET)) {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [{ addSheet: { properties: { title: GASTOS_SHEET, gridProperties: { hideGridlines: true } } } }] } });
    meta = await getMeta();
  }
  const sh = meta.data.sheets!.find((s) => s.properties?.title === GASTOS_SHEET)!;
  const sheetId = sh.properties!.sheetId!;
  const oldRows = sh.properties?.gridProperties?.rowCount || 1000;
  await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [
    ...(sh.rowGroups || []).map((g) => ({ deleteDimensionGroup: { range: g.range } })),
    { updateDimensionProperties: { range: { sheetId, dimension: "ROWS", startIndex: 0, endIndex: oldRows }, properties: { hiddenByUser: false }, fields: "hiddenByUser" } },
    { repeatCell: { range: { sheetId, startRowIndex: 0, endRowIndex: oldRows, startColumnIndex: 0, endColumnIndex: 30 }, cell: {}, fields: "userEnteredFormat" } },
  ] } });
  await sheets.spreadsheets.values.clear({ spreadsheetId, range: `${GASTOS_SHEET}!A:AD` });
  await sheets.spreadsheets.values.update({ spreadsheetId, range: `${GASTOS_SHEET}!A1`, valueInputOption: "USER_ENTERED", requestBody: { values: out } });

  // ── Formato ──
  const NAVY = { red: 0.118, green: 0.227, blue: 0.373 };
  const WHITE = { red: 1, green: 1, blue: 1 };
  const BAND = { red: 0.9, green: 0.93, blue: 0.97 };
  const GREY = { red: 0.95, green: 0.95, blue: 0.95 };
  const MUTED = { red: 0.4, green: 0.4, blue: 0.4 };
  const money = { type: "NUMBER", pattern: "#,##0;[Red]-#,##0;" };
  const rng = (r0: number, r1: number, c0: number, c1: number) => ({ sheetId, startRowIndex: r0 - 1, endRowIndex: r1, startColumnIndex: c0, endColumnIndex: c1 });
  const fmt = (range: any, f: any, fields: string) => ({ repeatCell: { range, cell: { userEnteredFormat: f }, fields: `userEnteredFormat(${fields})` } });
  const W = lastCol + 1;
  const reqs: any[] = [
    { updateSheetProperties: { properties: { sheetId, gridProperties: { hideGridlines: true, frozenRowCount: HEAD, frozenColumnCount: 1, rowCount: Math.max(1000, lastRow + 50), rowGroupControlAfter: false } }, fields: "gridProperties(hideGridlines,frozenRowCount,frozenColumnCount,rowCount,rowGroupControlAfter)" } },
    fmt(rng(1, lastRow, 0, W), { textFormat: { fontSize: 12 }, verticalAlignment: "MIDDLE" }, "textFormat,verticalAlignment"),
    fmt(rng(1, 1, 0, 1), { textFormat: { bold: true, fontSize: 18, foregroundColor: NAVY } }, "textFormat"),
    fmt(rng(2, 2, 0, 1), { textFormat: { italic: true, fontSize: 11, foregroundColor: MUTED } }, "textFormat"),
    fmt(rng(HEAD, HEAD, 0, W), { backgroundColor: NAVY, textFormat: { bold: true, fontSize: 12, foregroundColor: WHITE } }, "backgroundColor,textFormat"),
    fmt(rng(HEAD, HEAD, 1, W), { horizontalAlignment: "RIGHT" }, "horizontalAlignment"),
    fmt(rng(TOT, lastRow, 1, W), { numberFormat: money }, "numberFormat"),
    fmt(rng(TOT, TOT, 0, W), { backgroundColor: GREY, textFormat: { bold: true, fontSize: 13 } }, "backgroundColor,textFormat"),
    fmt(rng(HEAD, lastRow, W - 1, W), { textFormat: { bold: true } }, "textFormat"),
    { updateDimensionProperties: { range: { sheetId, dimension: "COLUMNS", startIndex: 0, endIndex: 1 }, properties: { pixelSize: 460 }, fields: "pixelSize" } },
    { updateDimensionProperties: { range: { sheetId, dimension: "COLUMNS", startIndex: 1, endIndex: Math.max(2, W) }, properties: { pixelSize: 105 }, fields: "pixelSize" } },
    { updateDimensionProperties: { range: { sheetId, dimension: "COLUMNS", startIndex: W - 1, endIndex: W }, properties: { pixelSize: 125 }, fields: "pixelSize" } },
  ];
  reqs.push(
    fmt(rng(FJ, FJ, 0, W), { backgroundColor: NAVY, textFormat: { bold: true, fontSize: 12, foregroundColor: WHITE } }, "backgroundColor,textFormat"),
    fmt(rng(FJ, FJ, 1, 2), { horizontalAlignment: "RIGHT" }, "horizontalAlignment"),
    fmt(rng(FJ + 1, FJ_TOT, 2, 3), { numberFormat: { type: "TEXT" }, textFormat: { fontSize: 11, italic: true, foregroundColor: MUTED }, horizontalAlignment: "LEFT" }, "numberFormat,textFormat,horizontalAlignment"),
    fmt(rng(FJ + 1, FJ_TOT, 1, 2), { textFormat: { bold: false } }, "textFormat"),
    fmt(rng(FJ_TOT, FJ_TOT, 0, 2), { backgroundColor: GREY, textFormat: { bold: true, fontSize: 13 } }, "backgroundColor,textFormat"),
    ...fjRows.map((x) => fmt(rng(x.h, x.h, 0, 2), { backgroundColor: BAND, textFormat: { bold: true, fontSize: 12 } }, "backgroundColor,textFormat")),
  );
  for (const sec of secs) {
    reqs.push(fmt(rng(sec.h, sec.h, 0, W), { backgroundColor: NAVY, textFormat: { bold: true, fontSize: 12, foregroundColor: WHITE } }, "backgroundColor,textFormat"));
  }
  for (const g of groups) {
    reqs.push(
      fmt(rng(g.h, g.h, 0, W), { backgroundColor: BAND, textFormat: { bold: true, fontSize: 12 } }, "backgroundColor,textFormat"),
      fmt(rng(g.d0, g.d1, 0, W), { textFormat: { fontSize: 11, foregroundColor: MUTED } }, "textFormat"),
      { addDimensionGroup: { range: { sheetId, dimension: "ROWS", startIndex: g.d0 - 1, endIndex: g.d1 } } },
    );
  }
  await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: reqs } });
  if (groups.length) {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: groups.map((g) => ({
      updateDimensionGroup: { dimensionGroup: { range: { sheetId, dimension: "ROWS", startIndex: g.d0 - 1, endIndex: g.d1 }, depth: 1, collapsed: true }, fields: "collapsed" },
    })) } });
  }
  return { categorias: groups.length, movimientos: egresos.length };
}
