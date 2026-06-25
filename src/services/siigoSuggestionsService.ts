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
  /** Valor de la columna "Relación con" del PUC de Siigo (ej. "Impuesto", "Activo fijo", "Proveedor"). */
  relatedTo?: string;
  /** Valor de la columna "Maneja vencimientos" del PUC (ej. "Con vencimiento en proveedores", "Con vencimiento en cartera"). */
  maturity?: string;
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
  paymentCode: string | null; // ID numérico de Siigo (ej: 113)
  cxpCode: string | null; // Cuenta por pagar dominante en el balance (21xx)
  /**
   * Variante de IVA inferida del balance según la cuenta de IVA descontable que
   * mueve el proveedor: "servicios" (cuenta "...por servicios"), "bienes"
   * ("...por compras"), o null si no hay señal (p. ej. el IVA se contabiliza bajo
   * DIAN, no bajo el proveedor). Determina el default IVA 19% Servicios vs IVA 19%.
   */
  ivaKind: "servicios" | "bienes" | null;
  updatedAt: string;
}

// ─── Lector de Excel Tolerante ───────────────────────────────────────
async function robustReadExcel(buffer: Buffer): Promise<string[][]> {
  const zip = await JSZip.loadAsync(buffer);
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "@_",
    removeNSPrefix: true,
    parseTagValue: false,
  });

  // 1. Cargar Shared Strings
  let shared: string[] = [];
  const ssFile = zip.file("xl/sharedStrings.xml");
  if (ssFile) {
    const xml = await ssFile.async("string");
    const data: any = parser.parse(xml);
    let si = data?.sst?.si ?? [];
    if (!Array.isArray(si)) si = [si];
    shared = si.map((s: any) => {
      if (s == null) return "";
      if (s.t != null) return typeof s.t === "object" ? String(s.t["#text"] ?? "") : String(s.t);
      if (s.r) {
        const parts = Array.isArray(s.r) ? s.r : [s.r];
        return parts.map((p: any) => (typeof p.t === "string" ? p.t : p.t?.["#text"] || "")).join("");
      }
      return "";
    });
  }

  // 2. Cargar la primera hoja
  const sheetPath = Object.keys(zip.files).find((n) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(n));
  if (!sheetPath) throw new Error("El archivo Excel no contiene hojas de datos.");
  
  const sheetXml = await zip.file(sheetPath)!.async("string");
  const sheetData: any = parser.parse(sheetXml);
  let xrows = sheetData?.worksheet?.sheetData?.row ?? [];
  if (!Array.isArray(xrows)) xrows = [xrows];

  const getCellVal = (c: any): string => {
    if (c == null) return "";
    const type = c["@_t"];
    let v = c.v;
    if (v && typeof v === "object") v = v["#text"];
    if (type === "s") return shared[Number(v)] ?? "";
    if (type === "inlineStr") return c.is?.t?.["#text"] ?? c.is?.t ?? "";
    return v == null ? "" : String(v).trim();
  };

  const rows: string[][] = [];
  for (const row of xrows) {
    let cells = row?.c ?? [];
    if (!Array.isArray(cells)) cells = [cells];
    const rowArr: string[] = [];
    let nextColIdx = 0;
    for (const c of cells) {
      const ref = c?.["@_r"];
      let colIdx = nextColIdx;
      if (ref) {
        const match = String(ref).match(/^([A-Z]+)/);
        if (match) {
          let n = 0;
          for (const ch of match[1]) n = n * 26 + (ch.charCodeAt(0) - 64);
          colIdx = n - 1;
        }
      }
      rowArr[colIdx] = getCellVal(c);
      nextColIdx = colIdx + 1;
    }
    rows.push(rowArr);
  }
  return rows;
}

// ─── Helpers ─────────────────────────────────────────────────────────
const num = (s: string) => {
  if (!s) return 0;
  let clean = String(s).replace(/[^\d,.\-]/g, "");
  if (clean.includes(",") && clean.includes(".")) {
    const firstComma = clean.indexOf(",");
    const firstDot = clean.indexOf(".");
    if (firstComma < firstDot) clean = clean.replace(/,/g, "");
    else clean = clean.replace(/\./g, "").replace(",", ".");
  } else if (clean.includes(",")) {
    if (clean.split(",").pop()!.length <= 2) clean = clean.replace(",", ".");
    else clean = clean.replace(/,/g, "");
  }
  return parseFloat(clean) || 0;
};

const normNit = (s: string) => String(s || "").split("-")[0].replace(/[.\s]/g, "").trim();

function parseRate(name: string): number | null {
  const s = String(name);
  const m = s.match(/(\d+(?:[.,]\d+)?)\s*%/);
  if (m) return parseFloat(m[1].replace(",", "."));
  const nums = [...s.matchAll(/(\d+(?:[.,]\d+)?)/g)].map((x) => parseFloat(x[1].replace(",", ".")));
  const small = nums.filter((n) => n > 0 && n < 40);
  return small.length ? small[small.length - 1] : null;
}

// ─── Traductor: cuenta contable → forma de pago ──────────────────────
interface PaymentMapEntry {
  name: string;
  code: string;
}

function parsePaymentMapRows(rows: string[][]): Record<string, PaymentMapEntry> {
  const map: Record<string, PaymentMapEntry> = {};
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const codeMedio = String(row[1] || "").trim(); // Columna 2 (B) - Código del medio
    const nombre = String(row[2] || "").trim(); // Columna 3 (C) - Nombre del medio
    const cuenta = String(row[4] || ""); // Columna 5 (E)
    
    const codeAcc = (cuenta.match(/^\d+/) || [""])[0];
    if (!codeAcc || !nombre) continue;
    
    // Solo pasivos 21/22/23
    if (!/^2[123]/.test(codeAcc)) continue;
    
    if (map[codeAcc]) {
      const currentIsCausacion = /causaci/i.test(map[codeAcc].name);
      const nextIsCausacion = /causaci/i.test(nombre);
      if (currentIsCausacion && !nextIsCausacion) map[codeAcc] = { name: nombre, code: codeMedio };
    } else {
      map[codeAcc] = { name: nombre, code: codeMedio };
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
}

function buildFromBalance(
  rows: string[][],
  paymentMap: Record<string, PaymentMapEntry>
): { profiles: SupplierProfile[]; accountsCatalog: AccountCatalogEntry[] } {
  const headerIdx = rows.findIndex((r) => r.some((c) => String(c || "").trim().toLowerCase().includes("transaccional")));
  if (headerIdx < 0) throw new Error("No se encontró el encabezado del balance.");

  const header = rows[headerIdx].map((c) => String(c || "").trim().toLowerCase());
  const col = {
    transaccional: header.findIndex((h) => h.includes("transaccional")),
    code: header.findIndex((h) => h.includes("código") || h.includes("codigo")),
    name: header.findIndex((h) => h.includes("nombre cuenta")),
    nit: header.findIndex((h) => h.includes("identificaci")),
    tercero: header.findIndex((h) => h.includes("nombre tercero")),
    debito: header.findIndex((h) => h.includes("débito") || h.includes("debito")),
    credito: header.findIndex((h) => h.includes("crédito") || h.includes("credito")),
  };

  const cell = (r: string[], i: number) => (i >= 0 ? String(r[i] ?? "").trim() : "");
  const raw = new Map<string, { nit: string; name: string; gasto: AccRow[]; refte: AccRow[]; reiva: AccRow[]; reica: AccRow[]; ivades: AccRow[]; pago: (AccRow & { paymentName: string; paymentCode: string })[] }>();
  const catalog = new Map<string, string>();

  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r) continue;
    if (cell(r, col.transaccional).toLowerCase() !== "sí" && cell(r, col.transaccional).toLowerCase() !== "si") continue;
    
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
    if (!p) { p = { nit, name: tercero, gasto: [], refte: [], reiva: [], reica: [], ivades: [], pago: [] }; raw.set(nit, p); }

    if (/^[567]/.test(code)) p.gasto.push(acc);
    else if (code.startsWith("2365")) p.refte.push(acc);
    else if (code.startsWith("2367")) p.reiva.push(acc);
    else if (code.startsWith("2368")) p.reica.push(acc);
    // IVA descontable (2408..) cuyo nombre describe el tipo: "...por compras" vs "...por servicios".
    else if (code.startsWith("2408") && /descont/i.test(accName)) p.ivades.push(acc);
    
    if (/^2[123]/.test(code) && paymentMap[code]) {
      p.pago.push({ ...acc, paymentName: paymentMap[code].name, paymentCode: paymentMap[code].code });
    }
  }

  const top = <T extends AccRow>(arr: T[], key: keyof AccRow): T | null => arr.length ? arr.slice().sort((a, b) => (b[key] as number) - (a[key] as number))[0] : null;
  const now = new Date().toISOString();
  const out: SupplierProfile[] = [];
  for (const p of raw.values()) {
    const byCode = new Map<string, { code: string; name: string; debito: number }>();
    for (const g of p.gasto) {
      const e = byCode.get(g.code);
      if (e) e.debito += g.debito;
      else byCode.set(g.code, { code: g.code, name: g.name, debito: g.debito });
    }
    const gastoAccounts = [...byCode.values()].sort((a, b) => b.debito - a.debito).map(({ code, name }) => ({ code, name }));
    const rf = top(p.refte, "credito");
    const ri = top(p.reiva, "credito");
    const rc = top(p.reica, "credito");
    const pg = top(p.pago, "credito");
    // Cuenta de IVA descontable con mayor movimiento → su nombre define la variante.
    const ivaTop = p.ivades.slice().sort(
      (a, b) => (Math.abs(b.debito) + Math.abs(b.credito)) - (Math.abs(a.debito) + Math.abs(a.credito))
    )[0] || null;
    let ivaKind: "servicios" | "bienes" | null = null;
    if (ivaTop) {
      if (/servicio/i.test(ivaTop.name)) ivaKind = "servicios";
      else if (/compra|bien/i.test(ivaTop.name)) ivaKind = "bienes";
    }
    out.push({
      nit: p.nit, name: p.name, gastoCode: gastoAccounts[0]?.code ?? null, gastoName: gastoAccounts[0]?.name ?? null, gastoAccounts,
      retefuente: rf ? { accountName: rf.name, rate: parseRate(rf.name) } : null,
      reteiva: ri ? { accountName: ri.name, rate: parseRate(ri.name) } : null,
      reteica: rc ? { accountName: rc.name, rate: parseRate(rc.name) } : null,
      paymentName: pg?.paymentName ?? null, paymentCode: pg?.paymentCode ?? null, cxpCode: pg?.code ?? null, ivaKind, updatedAt: now,
    });
  }
  return { profiles: out, accountsCatalog: [...catalog.entries()].map(([code, name]) => ({ code, name })).sort((a, b) => a.code.localeCompare(b.code)) };
}

// ─── Almacenamiento ──────────────────────────────────────────────────
export async function savePaymentMap(buffer: Buffer): Promise<{ count: number }> {
  const rows = await robustReadExcel(buffer);
  const map = parsePaymentMapRows(rows);
  if (Object.keys(map).length === 0) throw new Error("No se encontraron formas de pago en el Excel.");
  await getDb().collection<any>(CONFIG).updateOne({ _id: cfgId(PAYMENT_MAP_ID) }, { $set: { map, updatedAt: new Date().toISOString() } }, { upsert: true });
  return { count: Object.keys(map).length };
}

async function getStoredPaymentMap(): Promise<Record<string, PaymentMapEntry>> {
  const doc = await getDb().collection<any>(CONFIG).findOne({ _id: cfgId(PAYMENT_MAP_ID) });
  return (doc?.map as Record<string, PaymentMapEntry>) || {};
}

interface RebuildResult { profiles: number; withGasto: number; withPayment: number; withIvaServicios?: number; withIvaBienes?: number; accounts: number; }

async function storeProfiles(rows: string[][]): Promise<RebuildResult> {
  const paymentMap = await getStoredPaymentMap();
  const { profiles, accountsCatalog } = buildFromBalance(rows, paymentMap);
  if (profiles.length === 0) throw new Error("No se encontraron terceros en el balance.");
  const db = getDb();
  const cid = currentCompany();
  const col = db.collection<any>(PROFILES);
  await col.deleteMany({ $or: [{ companyId: cid }, { companyId: { $exists: false } }] });
  await col.insertMany(profiles.map((p) => ({ _id: `${cid}:${p.nit}`, companyId: cid, ...p })));
  await db.collection<any>(CONFIG).updateOne({ _id: cfgId(ACCOUNTS_ID) }, { $set: { accounts: accountsCatalog, updatedAt: new Date().toISOString() } }, { upsert: true });
  return {
    profiles: profiles.length,
    withGasto: profiles.filter((p) => p.gastoCode).length,
    withPayment: profiles.filter((p) => p.paymentName).length,
    withIvaServicios: profiles.filter((p) => p.ivaKind === "servicios").length,
    withIvaBienes: profiles.filter((p) => p.ivaKind === "bienes").length,
    accounts: accountsCatalog.length,
  };
}

export async function rebuildProfilesFromBalance(buffer: Buffer): Promise<RebuildResult> {
  return storeProfiles(await robustReadExcel(buffer));
}

async function fetchBalanceFromSiigo(year: number, monthStart: number, monthEnd: number): Promise<Buffer> {
  const resp = (await request("/v1/test-balance-report-by-thirdparty", { method: "POST", body: { year, month_start: monthStart, month_end: monthEnd, includes_tax_difference: false }, })) as { file_url?: string };
  if (!resp?.file_url) throw new Error("Siigo no devolvió la URL del balance.");
  const r = await fetch(resp.file_url);
  if (!r.ok) throw new Error(`No se pudo descargar el balance de Siigo (HTTP ${r.status}).`);
  return Buffer.from(await r.arrayBuffer());
}

export async function rebuildProfilesFromSiigoReport(year: number, monthStart: number, monthEnd: number): Promise<RebuildResult> {
  return storeProfiles(await robustReadExcel(await fetchBalanceFromSiigo(year, monthStart, monthEnd)));
}

export async function savePlanCuentas(buffer: Buffer): Promise<{ count: number }> {
  const rows = await robustReadExcel(buffer);

  // Detectar columnas por nombre de header (la primera fila con texto reconocible)
  let headerRow = -1;
  let colCode = 0, colName = 1, colMaturity = -1, colRelatedTo = -1, colActivo = 7, colNivel = 8;
  for (let i = 0; i < Math.min(5, rows.length); i++) {
    const r = rows[i];
    const found = r.findIndex((c) => /^c[oó]digo/i.test(String(c ?? "").trim()));
    if (found >= 0) {
      headerRow = i;
      colCode = found;
      r.forEach((c, j) => {
        const v = String(c ?? "").trim().toLowerCase();
        if (/^nombre/i.test(v)) colName = j;
        else if (/vencim/i.test(v)) colMaturity = j;
        else if (/relaci/i.test(v)) colRelatedTo = j;
        else if (/^activo/i.test(v)) colActivo = j;
        else if (/^nivel/i.test(v)) colNivel = j;
      });
      break;
    }
  }

  const accounts: AccountCatalogEntry[] = [];
  const seen = new Set<string>();
  for (let i = headerRow + 1; i < rows.length; i++) {
    const r = rows[i];
    const code = String(r[colCode] ?? "").trim();
    const name = String(r[colName] ?? "").trim();
    const activo = String(r[colActivo] ?? "").trim().toLowerCase();
    const nivel = String(r[colNivel] ?? "").trim().toLowerCase();
    if (!/^\d/.test(code)) continue;
    if (nivel !== "transaccional") continue;
    if (activo && activo !== "sí" && activo !== "si") continue;
    if (seen.has(code)) continue;
    seen.add(code);
    const entry: AccountCatalogEntry = { code, name };
    if (colRelatedTo >= 0) {
      const v = String(r[colRelatedTo] ?? "").trim();
      if (v) entry.relatedTo = v;
    }
    if (colMaturity >= 0) {
      const v = String(r[colMaturity] ?? "").trim();
      if (v) entry.maturity = v;
    }
    accounts.push(entry);
  }
  if (accounts.length === 0) throw new Error("No se encontraron cuentas transaccionales.");
  accounts.sort((a, b) => a.code.localeCompare(b.code));
  await getDb().collection<any>(CONFIG).updateOne({ _id: cfgId(PLAN_ID) }, { $set: { accounts, updatedAt: new Date().toISOString() } }, { upsert: true });
  return { count: accounts.length };
}

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

export async function getSuggestionsStatus(): Promise<{ profileCount: number; profilesUpdatedAt: string | null; paymentMapCount: number; paymentMapUpdatedAt: string | null; }> {
  const db = getDb();
  const cid = currentCompany();
  const profileCount = await db.collection(PROFILES).countDocuments({ companyId: cid });
  const anyProfile = await db.collection<any>(PROFILES).findOne({ companyId: cid }, { sort: { updatedAt: -1 } });
  const cfg = await db.collection<any>(CONFIG).findOne({ _id: cfgId(PAYMENT_MAP_ID) });
  return { profileCount, profilesUpdatedAt: anyProfile?.updatedAt || null, paymentMapCount: cfg?.map ? Object.keys(cfg.map).length : 0, paymentMapUpdatedAt: cfg?.updatedAt || null, };
}
