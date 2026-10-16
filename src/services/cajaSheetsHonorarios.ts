/**
 * Pestaña "Honorarios Socias" del Sheet: honorarios entregados a Sylvia y a María Claudia,
 * por proyecto (filas) con una columna por socia. Debajo de cada proyecto, plegado, el
 * detalle de los pagos. La escribe el portal (una vía) en cada sincronización.
 */
import { getDb } from "./database.js";
import { getSheetsClient } from "./cajaSheetsAuth.js";
import type { SyncProyectoRow } from "./cajaSheetsSync.js";

const COL = "cajaErpSheets";
export const HONORARIOS_SHEET = "Honorarios Socias";

const norm = (s: string) => s.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
const txt = (s: unknown) => { const v = String(s ?? ""); return /^[=+\-@]/.test(v) ? `'${v}` : v; };

/** Socia a la que corresponde un pago de honorarios (null si no es de ellas). */
export function sociaHonorario(categoria: string, descripcion: string): "Sylvia" | "María Claudia" | null {
  const t = norm(`${categoria} ${descripcion}`);
  if (!/honorarios/.test(t) || /contabilidad|contago/.test(t)) return null;
  if (/sylvia|silvia|calderon/.test(t)) return "Sylvia";
  if (/maria claudia|m\. ?claudia|bahamon/.test(t)) return "María Claudia";
  return null;
}

export async function syncHonorariosSocias(companyId: string, rows: SyncProyectoRow[]): Promise<{ proyectos: number; pagos: number }> {
  const d = await getDb().collection<any>(COL).findOne({ companyId });
  const spreadsheetId: string = d?.spreadsheetId;
  if (!spreadsheetId) throw new Error("Primero indica el enlace del Google Sheet.");
  const sheets = await getSheetsClient(companyId);
  if (!sheets) throw new Error("Google Sheets no está autorizado.");

  type P = { fecha: unknown; descripcion: string; valor: number; socia: "Sylvia" | "María Claudia" };
  const proys = rows
    .map((r) => ({
      nombre: r.proyecto.nombre,
      pagos: (r.items || [])
        .filter((i) => i.direction === "out" && Number(i.valor) > 0)
        .map((i): P | null => {
          const socia = sociaHonorario(i.categoria || "", i.descripcion || "");
          return socia ? { fecha: i.fecha, descripcion: i.descripcion || "", valor: Number(i.valor), socia } : null;
        })
        .filter((x): x is P => !!x)
        .sort((a, b) => String(a.fecha).localeCompare(String(b.fecha))),
    }))
    .filter((p) => p.pagos.length);

  // ── Contenido ──
  const out: any[][] = [];
  const set = (row: number, vals: any[]) => { out[row - 1] = vals; };
  const HEAD = 4, TOT = 5;
  let cur = TOT + 1;
  const groups = proys.map((p) => { const h = cur, d0 = h + 1, d1 = h + p.pagos.length; cur = d1 + 1; return { ...p, h, d0, d1 }; });
  const lastRow = cur - 1;

  set(1, ["Honorarios entregados a las socias"]);
  set(2, ["Por proyecto, según el portal. Usa el + del margen izquierdo para ver cada pago."]);
  set(HEAD, ["Proyecto", "Sylvia", "María Claudia", "Total"]);
  const sumH = (c: string) => (groups.length ? `=SUM(${groups.map((g) => `${c}${g.h}`).join(";")})` : "=0");
  set(TOT, ["TOTAL HONORARIOS", sumH("B"), sumH("C"), "=B5+C5"]);
  for (const g of groups) {
    set(g.h, [g.nombre, `=SUM(B${g.d0}:B${g.d1})`, `=SUM(C${g.d0}:C${g.d1})`, `=B${g.h}+C${g.h}`]);
    g.pagos.forEach((p, i) => {
      const row: (string | number)[] = ["", "", "", ""];
      row[0] = txt(`${String(p.fecha).slice(0, 10)} · ${p.descripcion}`);
      row[p.socia === "Sylvia" ? 1 : 2] = Math.round(p.valor);
      set(g.d0 + i, row);
    });
  }
  for (let i = 0; i < lastRow; i++) if (!out[i]) out[i] = [];

  // ── Pestaña: crear si falta, limpiar grupos/formato ──
  const getMeta = () => sheets.spreadsheets.get({ spreadsheetId, fields: "sheets(properties(sheetId,title,gridProperties),rowGroups)" });
  let meta = await getMeta();
  if (!meta.data.sheets?.some((s) => s.properties?.title === HONORARIOS_SHEET)) {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [{ addSheet: { properties: { title: HONORARIOS_SHEET, gridProperties: { hideGridlines: true } } } }] } });
    meta = await getMeta();
  }
  const sh = meta.data.sheets!.find((s) => s.properties?.title === HONORARIOS_SHEET)!;
  const sheetId = sh.properties!.sheetId!;
  const oldRows = sh.properties?.gridProperties?.rowCount || 1000;
  await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [
    ...(sh.rowGroups || []).map((g) => ({ deleteDimensionGroup: { range: g.range } })),
    { updateDimensionProperties: { range: { sheetId, dimension: "ROWS", startIndex: 0, endIndex: oldRows }, properties: { hiddenByUser: false }, fields: "hiddenByUser" } },
    { repeatCell: { range: { sheetId, startRowIndex: 0, endRowIndex: oldRows, startColumnIndex: 0, endColumnIndex: 10 }, cell: {}, fields: "userEnteredFormat" } },
  ] } });
  await sheets.spreadsheets.values.clear({ spreadsheetId, range: `${HONORARIOS_SHEET}!A:J` });
  await sheets.spreadsheets.values.update({ spreadsheetId, range: `${HONORARIOS_SHEET}!A1`, valueInputOption: "USER_ENTERED", requestBody: { values: out } });

  // ── Formato ──
  const NAVY = { red: 0.118, green: 0.227, blue: 0.373 };
  const WHITE = { red: 1, green: 1, blue: 1 };
  const BAND = { red: 0.9, green: 0.93, blue: 0.97 };
  const GREY = { red: 0.95, green: 0.95, blue: 0.95 };
  const MUTED = { red: 0.4, green: 0.4, blue: 0.4 };
  const money = { type: "NUMBER", pattern: "#,##0;[Red]-#,##0;" };
  const rng = (r0: number, r1: number, c0: number, c1: number) => ({ sheetId, startRowIndex: r0 - 1, endRowIndex: r1, startColumnIndex: c0, endColumnIndex: c1 });
  const fmt = (range: any, f: any, fields: string) => ({ repeatCell: { range, cell: { userEnteredFormat: f }, fields: `userEnteredFormat(${fields})` } });
  const reqs: any[] = [
    { updateSheetProperties: { properties: { sheetId, gridProperties: { hideGridlines: true, frozenRowCount: HEAD, frozenColumnCount: 1, rowCount: Math.max(1000, lastRow + 50), rowGroupControlAfter: false } }, fields: "gridProperties(hideGridlines,frozenRowCount,frozenColumnCount,rowCount,rowGroupControlAfter)" } },
    fmt(rng(1, lastRow, 0, 4), { textFormat: { fontSize: 12 }, verticalAlignment: "MIDDLE" }, "textFormat,verticalAlignment"),
    fmt(rng(1, 1, 0, 1), { textFormat: { bold: true, fontSize: 18, foregroundColor: NAVY } }, "textFormat"),
    fmt(rng(2, 2, 0, 1), { textFormat: { italic: true, fontSize: 11, foregroundColor: MUTED } }, "textFormat"),
    fmt(rng(HEAD, HEAD, 0, 4), { backgroundColor: NAVY, textFormat: { bold: true, fontSize: 12, foregroundColor: WHITE } }, "backgroundColor,textFormat"),
    fmt(rng(HEAD, HEAD, 1, 4), { horizontalAlignment: "RIGHT" }, "horizontalAlignment"),
    fmt(rng(TOT, lastRow, 1, 4), { numberFormat: money }, "numberFormat"),
    fmt(rng(TOT, TOT, 0, 4), { backgroundColor: GREY, textFormat: { bold: true, fontSize: 13 } }, "backgroundColor,textFormat"),
    fmt(rng(HEAD, lastRow, 3, 4), { textFormat: { bold: true } }, "textFormat"),
    { updateDimensionProperties: { range: { sheetId, dimension: "COLUMNS", startIndex: 0, endIndex: 1 }, properties: { pixelSize: 460 }, fields: "pixelSize" } },
    { updateDimensionProperties: { range: { sheetId, dimension: "COLUMNS", startIndex: 1, endIndex: 4 }, properties: { pixelSize: 135 }, fields: "pixelSize" } },
  ];
  for (const g of groups) {
    reqs.push(
      fmt(rng(g.h, g.h, 0, 4), { backgroundColor: BAND, textFormat: { bold: true, fontSize: 12 } }, "backgroundColor,textFormat"),
      fmt(rng(g.d0, g.d1, 0, 4), { textFormat: { fontSize: 11, foregroundColor: MUTED } }, "textFormat"),
      { addDimensionGroup: { range: { sheetId, dimension: "ROWS", startIndex: g.d0 - 1, endIndex: g.d1 } } },
    );
  }
  await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: reqs } });
  if (groups.length) {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: groups.map((g) => ({
      updateDimensionGroup: { dimensionGroup: { range: { sheetId, dimension: "ROWS", startIndex: g.d0 - 1, endIndex: g.d1 }, depth: 1, collapsed: true }, fields: "collapsed" },
    })) } });
  }
  return { proyectos: groups.length, pagos: groups.reduce((n, g) => n + g.pagos.length, 0) };
}
