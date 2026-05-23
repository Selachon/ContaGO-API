import ExcelJS from "exceljs";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { getDb } from "./database.js";
import { request, getCurrentSiigoCompanyId } from "./siigoService.js";

// Colecciones MongoDB
const PROFILES = "siigoSupplierProfiles";
const CONFIG = "siigoConfig";
const PAYMENT_MAP_ID = "siigoPaymentMap";
const ACCOUNTS_ID = "siigoAccountsCatalog";
const PLAN_ID = "siigoPlanCuentas";

// Aislamiento por empresa: cada documento se etiqueta con la empresa activa.
function currentCompany(): string {
  return getCurrentSiigoCompanyId() || "env";
}
function cfgId(type: string): string {
  return `${currentCompany()}:${type}`;
}

export interface RetentionHint {
  accountName: string;
  rate: number | null;
}

export interface AccountCatalogEntry {
  code: string;
  name: string;
}

export interface SupplierProfile {
  nit: string;
  name: string;
  gastoCode: string | null;
  gastoName: string | null;
  gastoAccounts: AccountCatalogEntry[];
  retefuente: RetentionHint | null;
  reteiva: RetentionHint | null;
  reteica: RetentionHint | null;
  paymentName: string | null;
  updatedAt: string;
}

// ─── Helpers ─────────────────────────────────────────────────────────
function cellVal(cell: ExcelJS.Cell): string {
  let v: unknown = cell.value;
  if (v && typeof v === "object") {
    const o = v as Record<string, unknown>;
    v = o.result ?? o.text ?? (Array.isArray(o.richText) ? o.richText.map((t: any) => t.text).join("") : "");
  }
  return v == null ? "" : String(v).trim();
}

const num = (s: string) => Number(String(s).replace(/[^0-9.\-]/g, "")) || 0;
const normNit = (s: string) => String(s || "").split("-")[0].replace(/[.\s]/g, "").trim();

// Tarifa desde el nombre de la cuenta de retención ("Honorarios 11%", "Reteica 8.66")
function parseRate(name: string): number | null {
  const s = String(name);
  const m = s.match(/(\d+(?:[.,]\d+)?)\s*%/);
  if (m) return parseFloat(m[1].replace(",", "."));
  const nums = [...s.matchAll(/(\d+(?:[.,]\d+)?)/g)].map((x) => parseFloat(x[1].replace(",", ".")));
  const small = nums.filter((n) => n > 0 && n < 40);
  return small.length ? small[small.length - 1] : null;
}

async function firstSheet(buffer: Buffer): Promise<ExcelJS.Worksheet> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer as unknown as ArrayBuffer);
  const ws = wb.worksheets[0];
  if (!ws) throw new Error("El Excel no contiene hojas");
  return ws;
}

// ─── Traductor: cuenta contable → forma de pago ──────────────────────
function parsePaymentMapSheet(ws: ExcelJS.Worksheet): Record<string, string> {
  const map: Record<string, string> = {};
  for (let r = 2; r <= ws.rowCount; r++) {
    const R = ws.getRow(r);
    const nombre = cellVal(R.getCell(3));
    const relacion = cellVal(R.getCell(4));
    const cuenta = cellVal(R.getCell(5));
    if (!relacion.includes("Proveedor")) continue;
    const code = (cuenta.match(/^\d+/) || [""])[0];
    if (!code || !nombre) continue;
    // Solo pasivos 21/22/23 — los bancos 11/12 salen por el recibo de pago posterior
    if (!/^2[123]/.test(code)) continue;
    if (map[code]) {
      // Conflicto (ej. 23359501): preferir la que NO sea "Causación"
      if (/causaci/i.test(map[code]) && !/causaci/i.test(nombre)) map[code] = nombre;
    } else {
      map[code] = nombre;
    }
  }
  return map;
}

// ─── Balance de prueba por tercero → perfiles ────────────────────────
interface AccRow {
  code: string;
  name: string;
  debito: number;
  credito: number;
  paymentName?: string;
}

function buildFromBalance(
  rows: string[][],
  paymentMap: Record<string, string>
): { profiles: SupplierProfile[]; accountsCatalog: AccountCatalogEntry[] } {
  // La fila de encabezado varía entre el export manual (fila 8) y el de la API (fila 5).
  const headerIdx = rows.findIndex((r) =>
    r.some((c) => String(c || "").trim().toLowerCase() === "transaccional")
  );
  if (headerIdx < 0) {
    throw new Error("No se encontró el encabezado del balance (columna 'Transaccional').");
  }
  const header = rows[headerIdx].map((c) => String(c || "").trim().toLowerCase());
  const col = {
    transaccional: header.findIndex((h) => /transaccional/.test(h)),
    code: header.findIndex((h) => /c[oó]digo/.test(h)),
    name: header.findIndex((h) => /nombre cuenta/.test(h)),
    nit: header.findIndex((h) => /identificaci/.test(h)),
    tercero: header.findIndex((h) => /nombre tercero/.test(h)),
    debito: header.findIndex((h) => /d[eé]bito/.test(h)),
    credito: header.findIndex((h) => /cr[eé]dito/.test(h)),
  };
  if (col.transaccional < 0 || col.code < 0 || col.nit < 0) {
    throw new Error("El balance no tiene las columnas esperadas (Transaccional, Código, Identificación).");
  }
  const cell = (r: string[], i: number) => (i >= 0 ? String(r[i] ?? "").trim() : "");

  const raw = new Map<
    string,
    { nit: string; name: string; gasto: AccRow[]; refte: AccRow[]; reiva: AccRow[]; reica: AccRow[]; pago: AccRow[] }
  >();
  const catalog = new Map<string, string>(); // código → nombre (cuentas transaccionales)

  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r) continue;
    const trans = cell(r, col.transaccional).toLowerCase();
    if (trans !== "sí" && trans !== "si") continue; // solo transaccionales
    const code = cell(r, col.code);
    const accName = cell(r, col.name);
    const nit = normNit(cell(r, col.nit));
    const tercero = cell(r, col.tercero);
    if (code) catalog.set(code, accName);
    if (!nit) continue;

    const acc: AccRow = {
      code,
      name: accName,
      debito: num(cell(r, col.debito)),
      credito: num(cell(r, col.credito)),
    };

    let p = raw.get(nit);
    if (!p) {
      p = { nit, name: tercero, gasto: [], refte: [], reiva: [], reica: [], pago: [] };
      raw.set(nit, p);
    }

    if (/^[567]/.test(code)) p.gasto.push(acc);
    else if (code.startsWith("2365")) p.refte.push(acc);
    else if (code.startsWith("2367")) p.reiva.push(acc);
    else if (code.startsWith("2368")) p.reica.push(acc);
    if (paymentMap[code]) p.pago.push({ ...acc, paymentName: paymentMap[code] });
  }

  const top = (arr: AccRow[], key: "debito" | "credito"): AccRow | null =>
    arr.length ? arr.slice().sort((a, b) => b[key] - a[key])[0] : null;

  const now = new Date().toISOString();
  const out: SupplierProfile[] = [];
  for (const p of raw.values()) {
    // Cuentas de gasto del proveedor: deduplicadas por código, sumando débito, mayor primero
    const byCode = new Map<string, { code: string; name: string; debito: number }>();
    for (const g of p.gasto) {
      const e = byCode.get(g.code);
      if (e) e.debito += g.debito;
      else byCode.set(g.code, { code: g.code, name: g.name, debito: g.debito });
    }
    const gastoAccounts = [...byCode.values()]
      .sort((a, b) => b.debito - a.debito)
      .map(({ code, name }) => ({ code, name }));

    const rf = top(p.refte, "credito");
    const ri = top(p.reiva, "credito");
    const rc = top(p.reica, "credito");
    const pg = top(p.pago, "credito");
    out.push({
      nit: p.nit,
      name: p.name,
      gastoCode: gastoAccounts[0]?.code ?? null,
      gastoName: gastoAccounts[0]?.name ?? null,
      gastoAccounts,
      retefuente: rf ? { accountName: rf.name, rate: parseRate(rf.name) } : null,
      reteiva: ri ? { accountName: ri.name, rate: parseRate(ri.name) } : null,
      reteica: rc ? { accountName: rc.name, rate: parseRate(rc.name) } : null,
      paymentName: pg?.paymentName ?? null,
      updatedAt: now,
    });
  }

  const accountsCatalog: AccountCatalogEntry[] = [...catalog.entries()]
    .map(([code, name]) => ({ code, name }))
    .sort((a, b) => a.code.localeCompare(b.code));

  return { profiles: out, accountsCatalog };
}

// ─── Almacenamiento ──────────────────────────────────────────────────
export async function savePaymentMap(buffer: Buffer): Promise<{ count: number }> {
  const ws = await firstSheet(buffer);
  const map = parsePaymentMapSheet(ws);
  if (Object.keys(map).length === 0) {
    throw new Error("No se encontraron formas de pago de proveedores en el Excel.");
  }
  await getDb()
    .collection<any>(CONFIG)
    .updateOne(
      { _id: cfgId(PAYMENT_MAP_ID) },
      { $set: { map, updatedAt: new Date().toISOString() } },
      { upsert: true }
    );
  return { count: Object.keys(map).length };
}

async function getStoredPaymentMap(): Promise<Record<string, string>> {
  const doc = await getDb().collection<any>(CONFIG).findOne({ _id: cfgId(PAYMENT_MAP_ID) });
  return (doc?.map as Record<string, string>) || {};
}

interface RebuildResult {
  profiles: number;
  withGasto: number;
  withPayment: number;
  accounts: number;
}

// Construye y guarda los perfiles a partir de las filas del balance.
async function storeProfiles(rows: string[][]): Promise<RebuildResult> {
  const paymentMap = await getStoredPaymentMap();
  const { profiles, accountsCatalog } = buildFromBalance(rows, paymentMap);
  if (profiles.length === 0) {
    throw new Error("No se encontraron terceros en el balance. ¿Es un balance de prueba por tercero?");
  }
  const db = getDb();
  const cid = currentCompany();
  const col = db.collection<any>(PROFILES);
  // Borra los perfiles de esta empresa (y los heredados sin empresa, si los hay).
  await col.deleteMany({ $or: [{ companyId: cid }, { companyId: { $exists: false } }] });
  await col.insertMany(profiles.map((p) => ({ _id: `${cid}:${p.nit}`, companyId: cid, ...p })));
  await db.collection<any>(CONFIG).updateOne(
    { _id: cfgId(ACCOUNTS_ID) },
    { $set: { accounts: accountsCatalog, updatedAt: new Date().toISOString() } },
    { upsert: true }
  );
  return {
    profiles: profiles.length,
    withGasto: profiles.filter((p) => p.gastoCode).length,
    withPayment: profiles.filter((p) => p.paymentName).length,
    accounts: accountsCatalog.length,
  };
}

// Balance subido como archivo
export async function rebuildProfilesFromBalance(buffer: Buffer): Promise<RebuildResult> {
  return storeProfiles(await readSheetRows(buffer));
}

// Genera el balance de prueba por tercero directamente desde la API de Siigo.
async function fetchBalanceFromSiigo(year: number, monthStart: number, monthEnd: number): Promise<Buffer> {
  const resp = (await request("/v1/test-balance-report-by-thirdparty", {
    method: "POST",
    body: { year, month_start: monthStart, month_end: monthEnd, includes_tax_difference: false },
  })) as { file_url?: string };
  if (!resp?.file_url) throw new Error("Siigo no devolvió la URL del balance.");
  const r = await fetch(resp.file_url);
  if (!r.ok) throw new Error(`No se pudo descargar el balance de Siigo (HTTP ${r.status}).`);
  return Buffer.from(await r.arrayBuffer());
}

// Balance traído directamente desde Siigo (sin subir archivo)
export async function rebuildProfilesFromSiigoReport(
  year: number,
  monthStart: number,
  monthEnd: number
): Promise<RebuildResult> {
  const buffer = await fetchBalanceFromSiigo(year, monthStart, monthEnd);
  return storeProfiles(await readSheetRows(buffer));
}

// Lee un .xlsx tolerando prefijos de namespace (x:) que ExcelJS no soporta.
// Devuelve las filas como arreglos de strings.
async function readSheetRows(buffer: Buffer): Promise<string[][]> {
  const zip = await JSZip.loadAsync(buffer);
  // parseTagValue:false → no convertir texto numérico a número (códigos/NIT se mantienen como string).
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "@_",
    removeNSPrefix: true,
    parseTagValue: false,
  });

  let shared: string[] = [];
  const ssF = zip.file("xl/sharedStrings.xml");
  if (ssF) {
    const ss: any = parser.parse(await ssF.async("string"));
    let si = ss?.sst?.si ?? [];
    if (!Array.isArray(si)) si = [si];
    shared = si.map((s: any) => {
      if (s == null) return "";
      if (s.t != null && typeof s.t !== "object") return String(s.t);
      if (s.t && typeof s.t === "object") return String(s.t["#text"] ?? "");
      if (s.r) {
        const r = Array.isArray(s.r) ? s.r : [s.r];
        return r.map((x: any) => (typeof x.t === "string" ? x.t : x.t?.["#text"] || "")).join("");
      }
      return "";
    });
  }

  const sheetPath = Object.keys(zip.files).find((n) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(n));
  if (!sheetPath) throw new Error("El Excel no contiene hojas");
  const sheet: any = parser.parse(await zip.file(sheetPath)!.async("string"));
  let xrows = sheet?.worksheet?.sheetData?.row ?? [];
  if (!Array.isArray(xrows)) xrows = [xrows];

  const cellText = (c: any): string => {
    if (c == null) return "";
    const t = c["@_t"];
    let v = c.v;
    if (v && typeof v === "object") v = v["#text"];
    if (t === "s") return shared[Number(v)] ?? "";
    if (t === "inlineStr") return c.is?.t?.["#text"] ?? c.is?.t ?? "";
    return v == null ? "" : String(v);
  };

  const out: string[][] = [];
  for (const row of xrows) {
    let cs = row?.c ?? [];
    if (!Array.isArray(cs)) cs = [cs];
    const arr: string[] = [];
    let auto = 0;
    for (const c of cs) {
      const ref = c?.["@_r"];
      const m = ref ? String(ref).match(/^([A-Z]+)/) : null;
      let idx: number;
      if (m) {
        let n = 0;
        for (const ch of m[1]) n = n * 26 + (ch.charCodeAt(0) - 64);
        idx = n - 1;
      } else {
        idx = auto;
      }
      auto = idx + 1;
      arr[idx] = cellText(c);
    }
    out.push(arr);
  }
  return out;
}

/**
 * Plan de cuentas exportado de Siigo. Columnas: Código | Nombre | Categoría |
 * Clase | Relación con | Vencimientos | Diferencia fiscal | Activo | Nivel agrupación.
 * Se toman solo las cuentas de nivel "Transaccional" y activas.
 */
export async function savePlanCuentas(buffer: Buffer): Promise<{ count: number }> {
  const rows = await readSheetRows(buffer);
  const accounts: AccountCatalogEntry[] = [];
  const seen = new Set<string>();
  for (const r of rows) {
    const code = String(r[0] ?? "").trim();
    const name = String(r[1] ?? "").trim();
    const activo = String(r[7] ?? "").trim().toLowerCase();
    const nivel = String(r[8] ?? "").trim().toLowerCase();
    if (!/^\d/.test(code)) continue;
    if (nivel !== "transaccional") continue;
    if (activo && activo !== "sí" && activo !== "si") continue;
    if (seen.has(code)) continue;
    seen.add(code);
    accounts.push({ code, name });
  }
  if (accounts.length === 0) {
    throw new Error("No se encontraron cuentas transaccionales. ¿Es el plan de cuentas de Siigo?");
  }
  accounts.sort((a, b) => a.code.localeCompare(b.code));
  await getDb()
    .collection<any>(CONFIG)
    .updateOne({ _id: cfgId(PLAN_ID) }, { $set: { accounts, updatedAt: new Date().toISOString() } }, { upsert: true });
  return { count: accounts.length };
}

// Catálogo de cuentas: el plan completo si se subió; si no, el derivado del balance.
export async function getAccountsCatalog(): Promise<AccountCatalogEntry[]> {
  const db = getDb();
  const plan = await db.collection<any>(CONFIG).findOne({ _id: cfgId(PLAN_ID) });
  if (Array.isArray(plan?.accounts) && plan.accounts.length) return plan.accounts;
  const bal = await db.collection<any>(CONFIG).findOne({ _id: cfgId(ACCOUNTS_ID) });
  return (bal?.accounts as AccountCatalogEntry[]) || [];
}

export async function getSupplierProfile(nit: string): Promise<SupplierProfile | null> {
  const key = normNit(nit);
  if (!key) return null;
  const doc = await getDb().collection<any>(PROFILES).findOne({ _id: `${currentCompany()}:${key}` });
  if (!doc) return null;
  const { _id, ...rest } = doc;
  return rest as SupplierProfile;
}

export async function getSuggestionsStatus(): Promise<{
  profileCount: number;
  profilesUpdatedAt: string | null;
  paymentMapCount: number;
  paymentMapUpdatedAt: string | null;
}> {
  const db = getDb();
  const cid = currentCompany();
  const profileCount = await db.collection(PROFILES).countDocuments({ companyId: cid });
  const anyProfile = await db.collection<any>(PROFILES).findOne({ companyId: cid }, { sort: { updatedAt: -1 } });
  const cfg = await db.collection<any>(CONFIG).findOne({ _id: cfgId(PAYMENT_MAP_ID) });
  return {
    profileCount,
    profilesUpdatedAt: anyProfile?.updatedAt || null,
    paymentMapCount: cfg?.map ? Object.keys(cfg.map).length : 0,
    paymentMapUpdatedAt: cfg?.updatedAt || null,
  };
}
