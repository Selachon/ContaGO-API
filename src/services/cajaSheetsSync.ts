/**
 * Sincroniza la Caja por proyecto (portal = fuente de verdad) hacia UNA pestaña
 * propia ("Caja Portal") del Google Sheet de las socias, en una sola vía. El
 * portal reescribe esa pestaña completa en cada sincronización y no toca ninguna
 * otra hoja del archivo. Los bloques de proyecciones del final los escriben las
 * socias a mano: el portal los lee antes de reescribir y los vuelve a colocar.
 *
 * Diseño deliberadamente simple: arriba las respuestas clave; debajo cada
 * proyecto con su saldo, y su detalle en un grupo de filas plegado ("+" del margen).
 */
import { getDb } from "./database.js";
import { getSheetsClient } from "./cajaSheetsAuth.js";

const COL = "cajaErpSheets";
export const SYNC_SHEET = "Caja Portal";
const ING_MARK = "INGRESOS PROYECTADOS";
const EGR_MARK = "EGRESOS PROYECTADOS";
const ING_TOT = "Total ingresos proyectados";
const EGR_TOT = "Total egresos proyectados";
const MIN_INPUT_ROWS = 8;

export interface SyncProyectoRow {
  proyecto: { id: string; nombre: string; esTransversal?: boolean; saldoInicial?: number };
  items: Array<{ fecha?: string; descripcion?: string; valor: number; direction: "in" | "out"; origen?: string; categoria?: string }>;
}
export interface SyncBanco { nombre: string; numero: string; saldo: number; fechaSaldo: string }

// Evita que Sheets interprete como fórmula un texto que empieza por = + - @
const txt = (s: unknown) => { const v = String(s ?? ""); return /^[=+\-@]/.test(v) ? `'${v}` : v; };
const iso = (s?: string) => String(s || "").slice(0, 10);

export async function syncCajaToSheet(companyId: string, rows: SyncProyectoRow[], bancos: SyncBanco[]): Promise<{ url: string; movimientos: number }> {
  const d = await getDb().collection<any>(COL).findOne({ companyId });
  const spreadsheetId: string = d?.spreadsheetId;
  if (!spreadsheetId) throw new Error("Primero indica el enlace del Google Sheet.");
  const sheets = await getSheetsClient(companyId);
  if (!sheets) throw new Error("Google Sheets no está autorizado.");

  // ── 1. Asegura las pestañas y limpia grupos, filas ocultas y formato de la vez anterior ──
  const getMeta = () => sheets.spreadsheets.get({ spreadsheetId, fields: "sheets(properties(sheetId,title,gridProperties),rowGroups)" });
  let meta = await getMeta();
  const find = (t: string) => meta.data.sheets?.find((s) => s.properties?.title === t);
  if (!find(SYNC_SHEET)) {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [{ addSheet: { properties: { title: SYNC_SHEET, gridProperties: { hideGridlines: true } } } }] } });
    meta = await getMeta();
  }
  meta = await getMeta();
  const sh = find(SYNC_SHEET)!;
  const sheetId = sh.properties!.sheetId!;
  const oldRows = sh.properties?.gridProperties?.rowCount || 1000;
  await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: [
    ...(sh.rowGroups || []).map((g) => ({ deleteDimensionGroup: { range: g.range } })),
    { updateDimensionProperties: { range: { sheetId, dimension: "ROWS", startIndex: 0, endIndex: oldRows }, properties: { hiddenByUser: false }, fields: "hiddenByUser" } },
    { repeatCell: { range: { sheetId, startRowIndex: 0, endRowIndex: oldRows, startColumnIndex: 0, endColumnIndex: 10 }, cell: {}, fields: "userEnteredFormat" } },
  ] } });
  // Proyecciones escritas por las socias (concepto en A, valor en D): se leen tal cual (con fórmulas) y se reponen.
  const prev = (await sheets.spreadsheets.values.get({ spreadsheetId, range: `${SYNC_SHEET}!A1:D3000`, valueRenderOption: "FORMULA" })).data.values || [];
  const grab = (mark: string, totalMark: string): Array<[any, any]> => {
    const i0 = prev.findIndex((r) => r[0] === mark);
    if (i0 < 0) return [];
    const res: Array<[any, any]> = [];
    for (let i = i0 + 1; i < prev.length && !String(prev[i]?.[0] ?? "").startsWith(totalMark); i++) {
      const c = prev[i]?.[0], v = prev[i]?.[3];
      if ((c !== undefined && c !== "") || (v !== undefined && v !== "")) res.push([c ?? "", v ?? ""]);
    }
    return res;
  };
  const savedIng = grab(ING_MARK, ING_TOT);
  const savedEgr = grab(EGR_MARK, EGR_TOT);
  await sheets.spreadsheets.values.clear({ spreadsheetId, range: `${SYNC_SHEET}!A:Z` });

  // ── 2. Contenido (filas 1-based). Columnas: A texto/fecha · B concepto · C valor/fecha saldo · D saldo ──
  const out: any[][] = [];
  const set = (row: number, vals: any[]) => { out[row - 1] = vals; };
  // Solo proyectos con actividad: sin movimientos ni saldo inicial no aportan nada a la hoja.
  rows = rows.filter((r) => r.items.length > 0 || (Number(r.proyecto.saldoInicial) || 0) !== 0);
  const esOficina = (p: SyncProyectoRow) => !!p.proyecto.esTransversal || /oficina/i.test(p.proyecto.nombre);

  const bk = bancos.length ? bancos : [{ nombre: "—", numero: "", saldo: 0, fechaSaldo: "" }];
  const rFecha = 3, rBanco = 4;
  const bankFirst = rBanco + 1, bankLast = rBanco + bk.length;
  const rReserva = bankLast + 1, rDeficit = rReserva + 1;
  const rProy = rDeficit + 2;              // ingresos proyectados · rProy+1 egresos proyectados · rProy+2 déficit con proyecciones
  const rProyDef = rProy + 2;
  const projHead = rProyDef + 2;

  let cur = projHead + 1;
  type Pl = { h: number; d0: number; d1: number; r: SyncProyectoRow };
  const plan: Pl[] = [];
  for (const r of rows) {
    const h = cur, d0 = h + 1, d1 = d0 + r.items.length; // d0 = "Saldo inicial"; movimientos d0+1..d1
    plan.push({ h, d0, d1, r });
    cur = d1 + 1;
  }
  const projLast = cur - 1;

  // Bloques de proyecciones (al final): encabezado + filas de captura + total
  const nIng = Math.max(MIN_INPUT_ROWS, savedIng.length + 3);
  const nEgr = Math.max(MIN_INPUT_ROWS, savedEgr.length + 3);
  const rProyTitle = projLast + 2;
  const ingHead = rProyTitle + 1, ingFirst = ingHead + 1, ingLast = ingFirst + nIng - 1, ingTot = ingLast + 1;
  const egrHead = ingTot + 2, egrFirst = egrHead + 1, egrLast = egrFirst + nEgr - 1, egrTot = egrLast + 1;
  const lastRow = egrTot;

  set(1, ["Caja por proyecto"]);
  set(rFecha, ["Fecha última actualización", "", "", `=MAX(C${bankFirst}:C${bankLast})`]);
  set(rBanco, ["¿Cuánto tenemos en el banco?", "", "", `=SUM(D${bankFirst}:D${bankLast})`]);
  bk.forEach((b, i) => set(bankFirst + i, [txt(b.nombre), "", iso(b.fechaSaldo), Math.round(Number(b.saldo) || 0)]));
  const reservaRefs = plan.filter((p) => !esOficina(p.r)).map((p) => `D${p.h}`);
  set(rReserva, ["¿Cuánto se debe tener reservado de proyectos?", "", "", reservaRefs.length ? `=SUM(${reservaRefs.join(";")})` : "=0"]);
  set(rDeficit, ["¿Cuál es el déficit de caja?", "", "", `=D${rBanco}-D${rReserva}`]);
  set(rProy, ["Ingresos proyectados (ver final de la hoja)", "", "", `=D${ingTot}`]);
  set(rProy + 1, ["Egresos proyectados (ver final de la hoja)", "", "", `=-D${egrTot}`]);
  set(rProyDef, ["Déficit de caja con proyecciones", "", "", `=D${rDeficit}+D${rProy}+D${rProy + 1}`]);
  set(projHead, ["Proyecto", "", "", "Saldo"]);

  let nMov = 0;
  for (const p of plan) {
    const ini = Math.round(Number(p.r.proyecto.saldoInicial) || 0);
    set(p.h, [txt(p.r.proyecto.nombre), "", "", `=D${p.d0}+SUM(C${p.d0 + 1}:C${p.d1})`]);
    set(p.d0, ["", "Saldo inicial", "", ini]);
    let saldo = ini;
    [...p.r.items].sort((a, b) => iso(a.fecha).localeCompare(iso(b.fecha))).forEach((it, i) => {
      const v = Math.round(Number(it.valor) || 0) * (it.direction === "in" ? 1 : -1);
      saldo += v;
      set(p.d0 + 1 + i, [iso(it.fecha), txt(it.categoria && !String(it.descripcion || "").includes(it.categoria) ? `${it.descripcion || ""} · ${it.categoria}` : it.descripcion), v, saldo]);
      nMov++;
    });
  }
  set(rProyTitle, ["PROYECCIONES — simulación, no afecta el saldo real. Escribe aquí el concepto y el valor (siempre positivo)."]);
  set(ingHead, [ING_MARK, "", "", "Valor"]);
  savedIng.forEach(([c, v], i) => set(ingFirst + i, [c, "", "", v]));
  set(ingTot, [ING_TOT, "", "", `=SUM(D${ingFirst}:D${ingLast})`]);
  set(egrHead, [EGR_MARK, "", "", "Valor"]);
  savedEgr.forEach(([c, v], i) => set(egrFirst + i, [c, "", "", v]));
  set(egrTot, [EGR_TOT, "", "", `=SUM(D${egrFirst}:D${egrLast})`]);
  for (let i = 0; i < lastRow; i++) if (!out[i]) out[i] = [];
  await sheets.spreadsheets.values.update({ spreadsheetId, range: `${SYNC_SHEET}!A1`, valueInputOption: "USER_ENTERED", requestBody: { values: out } });

  // ── 3. Formato ──
  const NAVY = { red: 0.118, green: 0.227, blue: 0.373 };
  const WHITE = { red: 1, green: 1, blue: 1 };
  const BAND = { red: 0.9, green: 0.93, blue: 0.97 };
  const AMBER = { red: 0.71, green: 0.33, blue: 0.04 };
  const AMBER_BG = { red: 1, green: 0.95, blue: 0.8 };
  const MUTED = { red: 0.4, green: 0.4, blue: 0.4 };
  const money = { type: "NUMBER", pattern: "#,##0;[Red]-#,##0" };
  const rng = (r0: number, r1: number, c0: number, c1: number) => ({ sheetId, startRowIndex: r0 - 1, endRowIndex: r1, startColumnIndex: c0, endColumnIndex: c1 });
  const fmt = (range: any, f: any, fields: string) => ({ repeatCell: { range, cell: { userEnteredFormat: f }, fields: `userEnteredFormat(${fields})` } });
  const width = (c0: number, c1: number, px: number) => ({ updateDimensionProperties: { range: { sheetId, dimension: "COLUMNS", startIndex: c0, endIndex: c1 }, properties: { pixelSize: px }, fields: "pixelSize" } });
  const bold12 = { textFormat: { bold: true, fontSize: 12 } };

  const reqs: any[] = [
    { updateSheetProperties: { properties: { sheetId, gridProperties: { hideGridlines: true, frozenRowCount: 0, rowCount: Math.max(1000, lastRow + 100), rowGroupControlAfter: false } }, fields: "gridProperties(hideGridlines,frozenRowCount,rowCount,rowGroupControlAfter)" } },
    fmt(rng(1, lastRow, 0, 4), { textFormat: { fontSize: 12 }, verticalAlignment: "MIDDLE" }, "textFormat,verticalAlignment"),
    fmt(rng(1, 1, 0, 1), { textFormat: { bold: true, fontSize: 18, foregroundColor: NAVY } }, "textFormat"),
    // respuestas clave
    fmt(rng(rFecha, rDeficit, 0, 4), { textFormat: { bold: true, fontSize: 13 } }, "textFormat"),
    fmt(rng(rFecha, rDeficit, 3, 4), { numberFormat: money, horizontalAlignment: "RIGHT" }, "numberFormat,horizontalAlignment"),
    fmt(rng(rFecha, rFecha, 3, 4), { numberFormat: { type: "DATE", pattern: "dd/mm/yyyy" } }, "numberFormat"),
    fmt(rng(rDeficit, rDeficit, 0, 4), { backgroundColor: BAND }, "backgroundColor"),
    // cuentas (detalle del banco)
    fmt(rng(bankFirst, bankLast, 0, 4), { textFormat: { fontSize: 11, foregroundColor: MUTED } }, "textFormat"),
    fmt(rng(bankFirst, bankLast, 2, 3), { numberFormat: { type: "DATE", pattern: "dd/mm/yyyy" } }, "numberFormat"),
    fmt(rng(bankFirst, bankLast, 3, 4), { numberFormat: money }, "numberFormat"),
    // proyecciones (ámbar)
    fmt(rng(rProy, rProyDef, 0, 4), { backgroundColor: AMBER_BG, textFormat: { bold: true, fontSize: 13, foregroundColor: AMBER } }, "backgroundColor,textFormat"),
    fmt(rng(rProy, rProyDef, 3, 4), { numberFormat: money, horizontalAlignment: "RIGHT" }, "numberFormat,horizontalAlignment"),
    // tabla de proyectos
    // bloques de captura del final
    fmt(rng(rProyTitle, rProyTitle, 0, 4), { textFormat: { bold: true, fontSize: 13, foregroundColor: AMBER } }, "textFormat"),
    ...[[ingHead, ingFirst, ingLast, ingTot], [egrHead, egrFirst, egrLast, egrTot]].flatMap(([h, f, l, t]) => [
      fmt(rng(h, h, 0, 4), { backgroundColor: AMBER, textFormat: { bold: true, fontSize: 12, foregroundColor: WHITE } }, "backgroundColor,textFormat"),
      fmt(rng(h, h, 3, 4), { horizontalAlignment: "RIGHT" }, "horizontalAlignment"),
      fmt(rng(f, l, 0, 4), { backgroundColor: AMBER_BG, textFormat: { fontSize: 12 } }, "backgroundColor,textFormat"),
      fmt(rng(f, t, 3, 4), { numberFormat: money }, "numberFormat"),
      fmt(rng(t, t, 0, 4), { backgroundColor: BAND, textFormat: { bold: true, fontSize: 12 } }, "backgroundColor,textFormat"),
    ]),
    fmt(rng(projHead, projHead, 0, 4), { backgroundColor: NAVY, textFormat: { bold: true, fontSize: 12, foregroundColor: WHITE } }, "backgroundColor,textFormat"),
    fmt(rng(projHead, projHead, 3, 4), { horizontalAlignment: "RIGHT" }, "horizontalAlignment"),
    width(0, 1, 400), width(1, 2, 430), width(2, 3, 130), width(3, 4, 150),
    { addDimensionGroup: { range: { sheetId, dimension: "ROWS", startIndex: bankFirst - 1, endIndex: bankLast } } },
  ];
  if (plan.length) {
    reqs.push(fmt(rng(projHead + 1, projLast, 2, 4), { numberFormat: money }, "numberFormat"));
    reqs.push(fmt(rng(projHead + 1, projLast, 0, 1), { numberFormat: { type: "DATE", pattern: "yyyy-mm-dd" } }, "numberFormat"));
  }
  for (const p of plan) {
    reqs.push(
      fmt(rng(p.h, p.h, 0, 4), { backgroundColor: BAND, ...bold12 }, "backgroundColor,textFormat"),
      fmt(rng(p.h, p.h, 3, 4), { horizontalAlignment: "RIGHT" }, "horizontalAlignment"),
      fmt(rng(p.d0, p.d1, 0, 4), { textFormat: { fontSize: 11, foregroundColor: MUTED } }, "textFormat"),
      { addDimensionGroup: { range: { sheetId, dimension: "ROWS", startIndex: p.d0 - 1, endIndex: p.d1 } } },
    );
  }
  await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: reqs } });

  // Todo plegado: se ve solo cada proyecto con su saldo; el "+" del margen abre el detalle.
  const ranges = [{ s: bankFirst - 1, e: bankLast }, ...plan.map((p) => ({ s: p.d0 - 1, e: p.d1 }))];
  await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: ranges.map((g) => ({
    updateDimensionGroup: { dimensionGroup: { range: { sheetId, dimension: "ROWS", startIndex: g.s, endIndex: g.e }, depth: 1, collapsed: true }, fields: "collapsed" },
  })) } });

  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${sheetId}`;
  await getDb().collection<any>(COL).updateOne({ companyId }, { $set: { syncEnabled: true, syncSheetGid: sheetId, lastSyncAt: new Date().toISOString(), lastSyncError: "" } });
  return { url, movimientos: nMov };
}

export async function markSyncError(companyId: string, message: string) {
  await getDb().collection<any>(COL).updateOne({ companyId }, { $set: { lastSyncError: message, lastSyncErrorAt: new Date().toISOString() } });
}
