import fs from "fs";
import os from "os";
import path from "path";
import type { Browser, Page } from "puppeteer";
import { launchBrowserWithRetry, resolveExecutablePath, closeBrowserSafely, REAL_USER_AGENT } from "./dianScraper.js";

/**
 * Entra a Siigo Nube con el correo y la contraseña que el usuario escribe en ContaGO (se usan una sola vez, no se guardan)
 * y trae lo que la API pública no entrega: el Excel del plan de cuentas, la tabla de Medios de pago y el balance de
 * prueba por tercero. Si la cuenta es de Contador, primero entra al cliente (por NIT) desde "Mis clientes".
 *
 * Siigo pinta su interfaz con componentes que usan shadow DOM: los selectores normales no los ven, por eso todo el
 * recorrido de la página pasa por `deep()`, que entra a cada shadowRoot.
 */

const SIIGO_WEB = "https://siigonube.siigo.com";
const PUC_ROUTE = "#/accountingparameter/1570?TabID=1647";
const PAYMENT_MEANS_ROUTE = "#/paymentmeans/1893";
const BALANCE_ROUTE = "#/reports/2000/5441?TabID=1612&pTabID=1612&rReport=1";
const LOGIN_TIMEOUT_MS = 60_000;
const DOWNLOAD_TIMEOUT_MS = 60_000;
const SEARCH_TIMEOUT_MS = 90_000;

export type SiigoWebErrorCode =
  | "bad_credentials"
  | "mfa_required"
  | "login_failed"
  | "client_not_found"
  | "puc_failed"
  | "payment_failed"
  | "balance_failed"
  | "comprobantes_failed"
  | "impuestos_failed";

export class SiigoWebError extends Error {
  constructor(public code: SiigoWebErrorCode, message: string) {
    super(message);
  }
}

export interface SiigoWebArtifacts {
  pucBuffer: Buffer | null;
  pucError: string | null;
  paymentText: string | null;
  paymentError: string | null;
  balanceBuffer: Buffer | null;
  balanceError: string | null;
}

export interface SiigoWebTarget {
  nit?: string;
  name?: string;
}

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

// ─── Utilidades de página ───────────────────────────────────────────────────

const DEEP_PRELUDE =
  'const all=(root)=>{const out=[];const walk=(r)=>{for(const e of Array.from(r.querySelectorAll("*"))){out.push(e);if(e.shadowRoot)walk(e.shadowRoot);}};walk(root);return out;};' +
  'const txt=(e)=>((e.innerText!==undefined?e.innerText:e.textContent)||"").replace(/\\s+/g," ").trim();' +
  'const vis=(e)=>{const r=e.getBoundingClientRect();return r.width>0&&r.height>0;};';

/** Ejecuta `body` (cuerpo de función) dentro de la página con helpers que recorren también los shadow roots. */
export function deep<T = unknown>(page: Page, body: string, arg?: unknown): Promise<T> {
  return page.evaluate(new Function("arg", DEEP_PRELUDE + body) as (a: unknown) => T, arg);
}

/** Clic en el elemento más específico cuyo texto coincide con el patrón. */
async function deepClickText(page: Page, pattern: RegExp, maxLen = 80): Promise<boolean> {
  return deep<boolean>(
    page,
    'const re=new RegExp(arg.src,arg.flags);' +
      'const el=all(document).filter((e)=>vis(e)&&re.test(txt(e))&&txt(e).length<arg.maxLen).sort((a,b)=>a.querySelectorAll("*").length-b.querySelectorAll("*").length)[0];' +
      'if(!el)return false;el.click();return true;',
    { src: pattern.source, flags: pattern.flags, maxLen }
  );
}

/** Carpeta de diagnóstico local (data/ está en .gitignore). Guarda captura y texto visible, nunca credenciales. */
function diagDir(): string {
  const dir = path.join(process.cwd(), "data", "siigo-web-diag");
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

export async function saveDiagnostic(page: Page, step: string, screenshot = true): Promise<string | null> {
  try {
    const base = path.join(diagDir(), `${new Date().toISOString().replace(/[:.]/g, "-")}__${step}`);
    if (screenshot) await page.screenshot({ path: `${base}.png`, fullPage: true }).catch(() => undefined);
    const text = await page.evaluate(() => document.body?.innerText?.slice(0, 6000) || "").catch(() => "");
    // Solo la ruta: el callback de inicio de sesión trae códigos de autenticación en la URL y no deben guardarse.
    const u = page.url();
    const route = /openid-callback|[?&]code=/i.test(u) ? "(callback de inicio de sesión)" : u.replace(/^https?:\/\/[^/]+/, "");
    fs.writeFileSync(`${base}.txt`, `RUTA: ${route}\n\n${text}`);
    return base;
  } catch {
    return null;
  }
}

async function waitForDownload(dir: string, ext: RegExp): Promise<string | null> {
  const deadline = Date.now() + DOWNLOAD_TIMEOUT_MS;
  while (Date.now() < deadline) {
    const hit = fs.readdirSync(dir).find((f) => !f.endsWith(".crdownload") && ext.test(f));
    if (hit) {
      await sleep(500); // deja terminar de escribir
      return path.join(dir, hit);
    }
    await sleep(700);
  }
  return null;
}

async function allowDownloads(page: Page, dir: string): Promise<void> {
  const cdp = await page.createCDPSession();
  await cdp.send("Browser.setDownloadBehavior", { behavior: "allow", downloadPath: dir });
}

/**
 * Abre una pantalla de la SPA de Siigo cambiando el hash. Si la app todavía está arrancando, ella misma redirige al
 * inicio (#/dashboard) y la ruta pedida se pierde; por eso se verifica que la ruta se quede y, si no, se reintenta.
 */
export async function goToRoute(page: Page, route: string): Promise<void> {
  const wanted = route.replace(/^#\//, "").split("?")[0]; // p. ej. "paymentmeans/1893"
  for (let attempt = 0; attempt < 4; attempt++) {
    await page.evaluate((r) => { location.hash = r; }, route);
    await sleep(6000);
    if (page.url().includes(wanted)) return;
    await sleep(3000);
  }
}

// ─── Inicio de sesión y cliente ─────────────────────────────────────────────

export async function login(page: Page, email: string, password: string): Promise<void> {
  await page.goto(SIIGO_WEB, { waitUntil: "domcontentloaded", timeout: 45_000 });
  // La página trae dos formularios; el visible es el de Siigo ("Continuar"), el otro (#signInName) queda oculto.
  await page.waitForSelector("#siigoSignInName", { visible: true, timeout: 30_000 });
  await page.type("#siigoSignInName", email, { delay: 15 });
  await page.type("#siigoPassword", password, { delay: 15 });
  await page.click("#siigoNext");

  const deadline = Date.now() + LOGIN_TIMEOUT_MS;
  while (Date.now() < deadline) {
    await sleep(700);
    if (page.url().startsWith(SIIGO_WEB)) return;
    const state = await page
      .evaluate(() => {
        const body = document.body?.innerText || "";
        return {
          rejected: /inv[aá]lid|incorrect|no coincid|no es correct/i.test(body),
          mfa: !!document.querySelector("#verificationCode, input[name*='code' i], input[id*='verif' i]") || /c[oó]digo de verificaci[oó]n|verification code/i.test(body),
        };
      })
      .catch(() => ({ rejected: false, mfa: false }));
    if (state.mfa) throw new SiigoWebError("mfa_required", "Esta cuenta de Siigo pide verificación en dos pasos; no se puede automatizar.");
    if (state.rejected) throw new SiigoWebError("bad_credentials", "Siigo rechazó el correo o la contraseña.");
  }
  throw new SiigoWebError("login_failed", "Siigo no terminó de iniciar sesión a tiempo.");
}

/**
 * Cuentas de Contador: aterrizan en el Portal Contador. Se abre "Mis clientes", se busca el cliente por NIT y se pulsa
 * "Ingresar", que abre la empresa en una pestaña nueva. Cuentas de empresa: ya están dentro y se devuelve la misma página.
 */
export async function enterCompany(browser: Browser, page: Page, target: SiigoWebTarget): Promise<Page> {
  await sleep(3000); // el SPA termina de decidir a dónde aterriza
  if (!/#\/portal\//i.test(page.url())) {
    await sleep(6000); // cuenta de empresa directa: el SPA aún carga el inicio y redirige si se navega demasiado pronto
    return page;
  }

  // La tarjeta "Mis clientes" (título + descripción) tarda unos segundos en pintarse.
  let inClients = false;
  for (let i = 0; i < 20 && !inClients; i++) {
    inClients = await deepClickText(page, /^mis clientes/i, 200);
    if (!inClients) await sleep(1000);
  }
  if (!inClients) {
    await saveDiagnostic(page, "portal-sin-mis-clientes");
    throw new SiigoWebError("client_not_found", 'No encontré "Mis clientes" en el Portal Contador.');
  }
  await sleep(5000);

  const query = target.nit ? target.nit.replace(/\D/g, "") : target.name || "";
  if (query) {
    const focused = await deep<boolean>(
      page,
      'const i=all(document).find((e)=>e.tagName==="INPUT"&&vis(e)&&/buscar cliente/i.test(e.placeholder||""));if(!i)return false;i.focus();return true;'
    );
    if (focused) {
      await page.keyboard.type(query, { delay: 30 });
      await sleep(3500);
    }
  }

  const rows = await deep<number>(page, 'return all(document).filter((e)=>e.tagName==="TR"&&vis(e)&&/ingresar/i.test(txt(e))).length;');
  if (!rows) {
    await saveDiagnostic(page, "cliente-no-encontrado");
    throw new SiigoWebError("client_not_found", `No encontré el cliente ${target.nit || target.name || ""} en Mis clientes de esta cuenta de Siigo.`);
  }

  const clicked = await deep<boolean>(
    page,
    'const tr=all(document).find((e)=>e.tagName==="TR"&&vis(e)&&/ingresar/i.test(txt(e)));if(!tr)return false;' +
      'const b=all(tr).find((e)=>e.tagName==="BUTTON"&&/^ingresar$/i.test(txt(e)))||all(tr).find((e)=>/^ingresar$/i.test(txt(e)));if(!b)return false;b.click();return true;'
  );
  if (!clicked) {
    await saveDiagnostic(page, "ingresar-sin-boton");
    throw new SiigoWebError("client_not_found", 'No encontré el botón "Ingresar" del cliente.');
  }

  // "Ingresar" abre la empresa en una pestaña nueva o, si el navegador bloquea el pop-up, en la misma pestaña.
  const inCompany = (p: Page) => /siigonube\.siigo\.com\/#\/(?!portal)/.test(p.url());
  let company: Page | null = null;
  for (let i = 0; i < 60 && !company; i++) {
    await sleep(1000);
    const pages = await browser.pages();
    company = pages.find((p) => p !== page && inCompany(p)) || (inCompany(page) ? page : null);
  }
  if (!company) {
    await saveDiagnostic(page, "ingresar-sin-pestana");
    throw new SiigoWebError("client_not_found", "Siigo no abrió la empresa del cliente.");
  }
  await company.setViewport({ width: 1366, height: 900 });
  await sleep(6000);
  return company;
}

// ─── Plan de cuentas ────────────────────────────────────────────────────────

async function downloadPuc(page: Page, dir: string): Promise<Buffer> {
  await allowDownloads(page, dir);
  await goToRoute(page, PUC_ROUTE);

  let clicked = false;
  for (let i = 0; i < 20 && !clicked; i++) {
    clicked = await deepClickText(page, /^descargar\s+excel$/i);
    if (!clicked) await sleep(1000);
  }
  if (!clicked) {
    await saveDiagnostic(page, "puc-sin-boton");
    throw new SiigoWebError("puc_failed", 'No encontré el botón "Descargar Excel" del plan de cuentas.');
  }
  const file = await waitForDownload(dir, /cuentas contables.*\.xlsx?$/i);
  if (!file) {
    await saveDiagnostic(page, "puc-sin-descarga");
    throw new SiigoWebError("puc_failed", "Siigo no generó el Excel del plan de cuentas a tiempo.");
  }
  return fs.readFileSync(file);
}

// ─── Medios de pago (se lee de la pantalla, sin descargar) ──────────────────

async function readPaymentMeans(page: Page): Promise<string> {
  await goToRoute(page, PAYMENT_MEANS_ROUTE);
  const extract = () =>
    deep<string>(
      page,
      'return all(document).filter((e)=>e.tagName==="TR"||e.getAttribute("role")==="row")' +
        '.map((r)=>Array.from(r.querySelectorAll("th,td,[role=gridcell],[role=cell],[role=columnheader]")).map((c)=>txt(c)).join("\\t"))' +
        '.filter((l)=>l.replace(/\\t/g,"").trim()).join("\\n");'
    );

  // La tabla se pinta después de cargar la pantalla: se espera hasta ~25 s a que aparezcan encabezado y filas.
  for (let i = 0; i < 25; i++) {
    const text = await extract();
    if (/cuenta contable/i.test(text) && text.split("\n").length >= 3) return text;
    await sleep(1000);
  }
  await saveDiagnostic(page, "medios-de-pago");
  throw new SiigoWebError("payment_failed", "No pude leer la tabla de Medios de pago de Siigo.");
}

// ─── Balance de prueba por tercero ──────────────────────────────────────────

/** Año presente completo (enero a diciembre), sin cuenta de diferencia fiscal; busca y exporta con el ícono de Excel. */
async function downloadBalance(page: Page, dir: string): Promise<Buffer> {
  await allowDownloads(page, dir);
  await goToRoute(page, BALANCE_ROUTE);

  const fail = async (step: string, msg: string): Promise<never> => {
    await saveDiagnostic(page, step);
    throw new SiigoWebError("balance_failed", msg);
  };

  let ready = "";
  for (let i = 0; i < 15 && ready !== "ok"; i++) {
    ready = await deep<string>(
      page,
      'const sels=Array.from(new Set(all(document).filter((e)=>e.tagName==="SELECT"))).filter(vis);' +
        'const months=sels.filter((s)=>s.options[0]&&/enero/i.test(s.options[0].text));' +
        'const fiscal=sels.find((s)=>Array.from(s.options).some((o)=>/^s[ií]$/i.test(o.text.trim()))&&Array.from(s.options).some((o)=>/^no$/i.test(o.text.trim())));' +
        'if(months.length<2||!fiscal)return "esperando";' +
        'const set=(s,v)=>{s.value=String(v);s.dispatchEvent(new Event("input",{bubbles:true}));s.dispatchEvent(new Event("change",{bubbles:true}));};' +
        'set(months[0],1);set(months[1],12);set(fiscal,0);return "ok";'
    );
    if (ready !== "ok") await sleep(1000);
  }
  if (ready !== "ok") await fail("balance-sin-filtros", "No encontré los filtros del balance de terceros.");
  await sleep(800);

  const searched = await deep<boolean>(
    page,
    'const b=all(document).find((e)=>e.tagName==="BUTTON"&&/^buscar$/i.test(txt(e))&&vis(e));if(!b)return false;b.click();return true;'
  );
  if (!searched) await fail("balance-sin-buscar", 'No encontré el botón "Buscar" del balance de terceros.');

  // "Cuando termine de buscar" la tabla trae filas de datos.
  let done = false;
  for (let i = 0; i < SEARCH_TIMEOUT_MS / 1500 && !done; i++) {
    await sleep(1500);
    done = await deep<boolean>(page, 'return all(document).filter((e)=>e.tagName==="TR").length>40;');
  }
  if (!done) await fail("balance-sin-resultados", "La búsqueda del balance no devolvió resultados a tiempo.");

  const exported = await deep<boolean>(
    page,
    'const el=all(document).find((e)=>e.id==="export-excel-button"||String(e.className||"").includes("export-excel-button"));if(!el)return false;el.click();return true;'
  );
  if (!exported) await fail("balance-sin-exportar", "No encontré el botón de exportar datos del balance.");

  const file = await waitForDownload(dir, /balance.*\.xlsx?$/i);
  if (!file) await fail("balance-sin-descarga", "Siigo no generó el Excel del balance a tiempo.");
  return fs.readFileSync(file as string);
}

// ─── Tipos de comprobante contable ──────────────────────────────────────────

const COMPROBANTES_ROUTE = "#/reports/2000/5416?TabID=1399";

export interface SiigoComprobanteTipo {
  codigo: string;
  titulo: string;
}

/** Lee el listado "Tipos de comprobante contable" (código + título) directo de la pantalla, sin descargar. */
async function readComprobantes(page: Page): Promise<SiigoComprobanteTipo[]> {
  await goToRoute(page, COMPROBANTES_ROUTE);
  const extract = () =>
    deep<string[]>(
      page,
      'return all(document).filter((e)=>e.tagName==="TR"||e.getAttribute("role")==="row")' +
        '.map((r)=>Array.from(r.querySelectorAll("th,td,[role=gridcell],[role=cell]")).map((c)=>txt(c)).join("\t")).filter(Boolean);'
    );
  for (let i = 0; i < 25; i++) {
    const items: SiigoComprobanteTipo[] = [];
    const seen = new Set<string>();
    for (const line of await extract()) {
      const [codigo, titulo] = line.split("\t");
      if (/^\d+$/.test(codigo || "") && titulo && !seen.has(codigo)) {
        seen.add(codigo);
        items.push({ codigo, titulo });
      }
    }
    if (items.length) return items.sort((a, b) => Number(a.codigo) - Number(b.codigo));
    await sleep(1000);
  }
  await saveDiagnostic(page, "comprobantes");
  throw new SiigoWebError("comprobantes_failed", "No pude leer los tipos de comprobante de Siigo.");
}

// ─── Sesión y orquestadores ─────────────────────────────────────────────────

/** Abre un navegador, inicia sesión, entra a la empresa (por NIT si es cuenta de Contador) y ejecuta `fn`. */
async function withCompanySession<T>(email: string, password: string, target: SiigoWebTarget, fn: (company: Page, dir: string) => Promise<T>): Promise<T> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "siigo-web-"));
  const browser = await launchBrowserWithRetry(resolveExecutablePath(), () => undefined);
  try {
    const page = await browser.newPage();
    await page.setUserAgent(REAL_USER_AGENT);
    await page.setExtraHTTPHeaders({ "Accept-Language": "es-CO,es;q=0.9" });
    await page.evaluateOnNewDocument(() => {
      try { Object.defineProperty(navigator, "webdriver", { get: () => undefined }); } catch { /* noop */ }
    });
    await page.setViewport({ width: 1366, height: 900 });

    try {
      await login(page, email, password);
    } catch (e) {
      if (!(e instanceof SiigoWebError) || e.code === "login_failed") await saveDiagnostic(page, "login", false);
      throw e;
    }
    const company = await enterCompany(browser, page, target);
    return await fn(company, dir);
  } finally {
    await closeBrowserSafely(browser).catch(() => undefined);
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

export function fetchSiigoComprobantes(email: string, password: string, target: SiigoWebTarget = {}): Promise<SiigoComprobanteTipo[]> {
  return withCompanySession(email, password, target, (company) => readComprobantes(company));
}

export interface SiigoTablaImpuestos {
  columns: string[];
  rows: string[][];
}

const IMPUESTOS_ROUTE = "#/accountingparameter/1575";

/**
 * Lee la tabla de Impuestos (la que se copiaba y pegaba en Excel): encabezado "En uso, Código, Nombre, Tipo de impuesto,
 * Por valor, Tarifa, Ventas, Compras, Devolución ventas, Devolución compras" y una fila por impuesto. La pantalla trae
 * además otras tablas (resoluciones, autorretenciones): solo se toma la que tiene "Tipo de impuesto" y "Compras".
 */
async function readImpuestos(page: Page): Promise<SiigoTablaImpuestos> {
  await goToRoute(page, IMPUESTOS_ROUTE);
  const extract = () =>
    deep<string[][]>(
      page,
      'return all(document).filter((e)=>e.tagName==="TR"||e.getAttribute("role")==="row")' +
        '.map((r)=>Array.from(r.querySelectorAll("th,td,[role=gridcell],[role=cell],[role=columnheader]")).map((c)=>txt(c)));'
    );
  for (let i = 0; i < 25; i++) {
    const rows = await extract();
    const h = rows.findIndex((r) => r.some((c) => /^tipo de impuesto$/i.test(c)) && r.some((c) => /^compras$/i.test(c)));
    if (h >= 0) {
      const header = rows[h];
      const keep = header.map((c, idx) => (c.trim() ? idx : -1)).filter((idx) => idx >= 0);
      const iCodigo = header.findIndex((c) => /^c[oó]digo$/i.test(c.trim()));
      const data: string[][] = [];
      for (const r of rows.slice(h + 1)) {
        if (r.some((c) => /^en uso$/i.test(c.trim()))) break; // empieza otra tabla (autorretenciones)
        if (iCodigo >= 0 && (r[iCodigo] || "").trim()) data.push(keep.map((idx) => (r[idx] ?? "").trim()));
      }
      if (data.length) return { columns: keep.map((idx) => header[idx].trim()), rows: data };
    }
    await sleep(1000);
  }
  await saveDiagnostic(page, "impuestos");
  throw new SiigoWebError("impuestos_failed", "No pude leer la tabla de impuestos de Siigo.");
}

export interface SiigoParametrizacion {
  comprobantes: SiigoComprobanteTipo[] | null;
  comprobantesError: string | null;
  impuestos: SiigoTablaImpuestos | null;
  impuestosError: string | null;
  pucBuffer: Buffer | null;
  pucError: string | null;
  /** Solo si se pidió (`opciones.balance`): balance de prueba por tercero del año en curso. */
  balanceBuffer?: Buffer | null;
  balanceError?: string | null;
}

/** Parametrización de Causación Interfaces: tipos de comprobante + impuestos + plan de cuentas (el plan va al final). */
export function fetchSiigoParametrizacion(email: string, password: string, target: SiigoWebTarget = {}, opciones: { balance?: boolean } = {}): Promise<SiigoParametrizacion> {
  return withCompanySession(email, password, target, async (company, dir) => {
    const out: SiigoParametrizacion = { comprobantes: null, comprobantesError: null, impuestos: null, impuestosError: null, pucBuffer: null, pucError: null };
    try { out.comprobantes = await readComprobantes(company); }
    catch (e) { out.comprobantesError = e instanceof Error ? e.message : String(e); }
    try { out.impuestos = await readImpuestos(company); }
    catch (e) { out.impuestosError = e instanceof Error ? e.message : String(e); }
    if (opciones.balance) {
      try { out.balanceBuffer = await downloadBalance(company, dir); }
      catch (e) { out.balanceError = e instanceof Error ? e.message : String(e); }
    }
    // Tras la descarga del plan, Siigo deja la SPA sin pintar las pantallas siguientes: va al final.
    try { out.pucBuffer = await downloadPuc(company, dir); }
    catch (e) { out.pucError = e instanceof Error ? e.message : String(e); }
    return out;
  });
}

export function fetchSiigoWebArtifacts(email: string, password: string, target: SiigoWebTarget = {}): Promise<SiigoWebArtifacts> {
  return withCompanySession(email, password, target, async (company, dir) => {
    const out: SiigoWebArtifacts = { pucBuffer: null, pucError: null, paymentText: null, paymentError: null, balanceBuffer: null, balanceError: null };

    // El plan de cuentas va al final: tras su descarga Siigo deja la SPA sin pintar las pantallas siguientes.
    try { out.paymentText = await readPaymentMeans(company); }
    catch (e) { out.paymentError = e instanceof Error ? e.message : String(e); }

    try { out.balanceBuffer = await downloadBalance(company, dir); }
    catch (e) { out.balanceError = e instanceof Error ? e.message : String(e); }

    try { out.pucBuffer = await downloadPuc(company, dir); }
    catch (e) { out.pucError = e instanceof Error ? e.message : String(e); }
    return out;
  });
}
