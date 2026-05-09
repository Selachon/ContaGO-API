import puppeteer, { Browser, Page } from "puppeteer";
import fs from "fs";

const DIAN_RUT_URL = "https://muisca.dian.gov.co/WebRutMuisca/DefConsultaEstadoRUT.faces";
const DIAN_SITEKEY = "0x4AAAAAAAg1YFKr1lxPdUIL"; // Cloudflare Turnstile sitekey de DIAN MUISCA

// ── CapSolver (Cloudflare Turnstile via direct HTTP) ─────────────────────────
async function resolverTurnstile(pageUrl: string): Promise<string | null> {
  const apiKey = process.env.CAPSOLVER_API_KEY;
  if (!apiKey) return null;

  try {
    // Create task
    const createRes = await fetch("https://api.capsolver.com/createTask", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        clientKey: apiKey,
        task: {
          type: "AntiTurnstileTaskProxyLess",
          websiteURL: pageUrl,
          websiteKey: DIAN_SITEKEY,
        },
      }),
    });
    const createData = await createRes.json() as { taskId?: string; errorId?: number; errorDescription?: string };
    if (createData.errorId || !createData.taskId) {
      console.warn("[RUT] CapSolver createTask error:", createData.errorDescription);
      return null;
    }

    // Poll for result (up to 60s)
    for (let i = 0; i < 30; i++) {
      await new Promise((r) => setTimeout(r, 2000));
      const resultRes = await fetch("https://api.capsolver.com/getTaskResult", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ clientKey: apiKey, taskId: createData.taskId }),
      });
      const resultData = await resultRes.json() as { status?: string; solution?: { token?: string }; errorId?: number };
      if (resultData.status === "ready") return resultData.solution?.token || null;
      if (resultData.errorId) { console.warn("[RUT] CapSolver task error:", resultData); return null; }
    }
    console.warn("[RUT] CapSolver timeout after 60s");
    return null;
  } catch (err) {
    console.warn("[RUT] CapSolver error:", (err as Error).message);
    return null;
  }
}

// ── Verification digit ────────────────────────────────────────────────────────
const DV_PRIMES = [3, 7, 13, 17, 19, 23, 29, 37, 41, 43, 47, 53, 59, 67, 71];

export function calcularDV(nit: string): string {
  const digits = nit.replace(/\D/g, "").padStart(15, "0").split("").map(Number);
  const sum = digits.reduce((acc, d, i) => acc + d * DV_PRIMES[i], 0);
  const rem = sum % 11;
  return (rem < 2 ? rem : 11 - rem).toString();
}

export function normalizarNIT(raw: string): { nit: string; dv: string } {
  // Accept: "901234567", "901234567-8", "901234567 8"
  const clean = raw.replace(/\D/g, "");
  if (clean.length > 10) {
    const nit = clean.slice(0, -1);
    return { nit, dv: calcularDV(nit) };
  }
  return { nit: clean, dv: calcularDV(clean) };
}

// ── Result types ──────────────────────────────────────────────────────────────
export interface RutResult {
  nit: string;
  dv: string;
  primerApellido: string;
  segundoApellido: string;
  primerNombre: string;
  otrosNombres: string;
  razonSocial: string;
  estado: string;
  fechaInscripcion: string;
  fechaActualizacion: string;
  responsabilidades: string[];
  error?: string;
}

// ── In-memory cache (24h TTL) ─────────────────────────────────────────────────
interface CacheEntry {
  result: RutResult;
  expiresAt: number;
}
const cache = new Map<string, CacheEntry>();
const CACHE_TTL_MS = 24 * 60 * 60 * 1000;

function getCached(nit: string): RutResult | null {
  const entry = cache.get(nit);
  if (!entry) return null;
  if (Date.now() > entry.expiresAt) { cache.delete(nit); return null; }
  return entry.result;
}

function setCache(nit: string, result: RutResult): void {
  cache.set(nit, { result, expiresAt: Date.now() + CACHE_TTL_MS });
}

// Clean expired entries periodically
setInterval(() => {
  const now = Date.now();
  for (const [k, v] of cache) {
    if (now > v.expiresAt) cache.delete(k);
  }
}, 60 * 60 * 1000);

// ── Browser singleton ─────────────────────────────────────────────────────────
let rutBrowser: Browser | null = null;
let browserInitTime = 0;
const BROWSER_MAX_AGE_MS = 4 * 60 * 60 * 1000; // restart every 4h to avoid leaks

const BROWSER_ARGS = [
  "--no-sandbox",
  "--disable-setuid-sandbox",
  "--disable-dev-shm-usage",
  // NOT --disable-gpu: makes fingerprint more bot-like
  "--disable-extensions",
  "--disable-background-networking",
  "--no-first-run",
  "--window-size=1280,900",
];

function resolveChromePath(): string | undefined {
  if (!process.env.PUPPETEER_CACHE_DIR) {
    process.env.PUPPETEER_CACHE_DIR = `${process.cwd()}/.cache/puppeteer`;
  }
  const candidates = [
    process.env.PUPPETEER_EXECUTABLE_PATH,
    "/usr/bin/chromium-browser",
    "/usr/bin/chromium",
    "/usr/bin/google-chrome",
    "/usr/bin/google-chrome-stable",
  ];
  for (const c of candidates) {
    if (c && fs.existsSync(c)) return c;
  }
  return undefined;
}

async function getBrowser(): Promise<Browser> {
  const tooOld = Date.now() - browserInitTime > BROWSER_MAX_AGE_MS;

  if (rutBrowser && !tooOld) {
    try {
      // Quick liveness check
      const pages = await rutBrowser.pages();
      if (pages !== null) return rutBrowser;
    } catch {
      rutBrowser = null;
    }
  }

  if (rutBrowser) {
    try { await rutBrowser.close(); } catch {}
    rutBrowser = null;
  }

  rutBrowser = await puppeteer.launch({
    headless: true,
    args: BROWSER_ARGS,
    executablePath: resolveChromePath(),
    timeout: 60000,
  });
  browserInitTime = Date.now();
  return rutBrowser;
}

// ── DOM extraction script (plain JS string — avoids esbuild __name helper leak) ─
// Uses direct element IDs confirmed from live DIAN HTML inspection.
// All fields in format: id="vistaConsultaEstadoRUT:formConsultaEstadoRUT:<field>"
const EXTRACT_SCRIPT = `(function() {
  var P = "vistaConsultaEstadoRUT:formConsultaEstadoRUT:";
  function g(id) { var el = document.getElementById(P + id); return el ? (el.textContent || "").trim() : ""; }

  // Error detection
  var bodyLower = ((document.body && document.body.textContent) || "").toLowerCase();
  if (bodyLower.indexOf("error validando token") !== -1 || bodyLower.indexOf("error validando captcha") !== -1) {
    return { error: "CAPTCHA requerido — la DIAN requiere verificación de navegador" };
  }
  var tblMsg = document.getElementById("tblMensajes");
  if (tblMsg) {
    var mt = (tblMsg.textContent || "").toLowerCase();
    if (mt.indexOf("no se encontr") !== -1 || mt.indexOf("no existe") !== -1) {
      return { error: "NIT no registrado en el RUT de la DIAN" };
    }
    if (mt.indexOf("token") !== -1 || mt.indexOf("captcha") !== -1) {
      return { error: "CAPTCHA requerido — la DIAN requiere verificación de navegador" };
    }
  }

  // Direct ID lookup (confirmed IDs from live DIAN page)
  var data = {
    primerApellido:  g("primerApellido"),
    segundoApellido: g("segundoApellido"),
    primerNombre:    g("primerNombre"),
    otrosNombres:    g("otrosNombres"),
    razonSocial:     g("razonSocial") || g("denominacionRazonSocial"),
    estado:          g("estado"),
    dvPage:          g("dv"),
    fechaActualizacion: "",
    fechaInscripcion:   "",
    error: null
  };

  // Date fields have no IDs — scan adjacent td pairs for JSF label keys
  var tds = [].slice.call(document.querySelectorAll("td.fondoTituloLeftAjustado, td.fondoTituloLeft"));
  for (var i = 0; i < tds.length; i++) {
    var lbl = (tds[i].textContent || "").trim().toLowerCase()
      .replace(/^\\?\\?\\?label_/, "").replace(/\\?\\?\\?$/, "").replace(/_/g, " ");
    var next = tds[i].nextElementSibling;
    var val = next && !next.querySelector("input") ? (next.textContent || "").trim() : "";
    if (!val) continue;
    if (lbl === "fec actual" || lbl.indexOf("actualizacion") !== -1) data.fechaActualizacion = val;
    if (lbl.indexOf("inscripcion") !== -1 || lbl === "fec inscripcion") data.fechaInscripcion = val;
  }

  // Validate we actually have result data
  if (!data.estado && !data.primerApellido && !data.razonSocial) {
    return { error: "No se encontraron datos — posible CAPTCHA no resuelto o NIT inválido" };
  }

  data.error = null;
  return data;
})()`;

// ── Core query function ───────────────────────────────────────────────────────
export async function consultarRUT(raw: string): Promise<RutResult> {
  const { nit, dv } = normalizarNIT(raw);

  if (!nit || nit.length < 5) {
    return errorResult(nit, dv, "NIT inválido: debe tener al menos 5 dígitos");
  }

  const cached = getCached(nit);
  if (cached) return cached;

  const browser = await getBrowser();
  const page = await browser.newPage();

  try {
    page.setDefaultTimeout(30000);
    page.setDefaultNavigationTimeout(30000);

    await page.setUserAgent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36");
    await page.evaluateOnNewDocument(() => {
      Object.defineProperty(navigator, "webdriver", { get: () => undefined });
      (window as unknown as Record<string, unknown>).chrome = { runtime: {} };
      Object.defineProperty(navigator, "plugins", { get: () => [1, 2, 3] });
      Object.defineProperty(navigator, "languages", { get: () => ["es-CO", "es", "en"] });
    });

    await page.goto(DIAN_RUT_URL, { waitUntil: "networkidle2", timeout: 45000 });

    // Fill NIT field — real DIAN selector (discovered from live page inspection)
    // The form prefix is "vistaConsultaEstadoRUT:formConsultaEstadoRUT"
    const PREFIX = "vistaConsultaEstadoRUT:formConsultaEstadoRUT";
    const nitSelectors = [
      `[id="${PREFIX}:numNit"]`,
      `input[name="${PREFIX}:numNit"]`,
      'input[id*="numNit"]',
      '[id="form:numNit"]',
    ];
    let nitFilled = false;
    for (const sel of nitSelectors) {
      try {
        await page.waitForSelector(sel, { timeout: 5000 });
        await page.click(sel, { clickCount: 3 });
        await page.type(sel, nit);
        nitFilled = true;
        break;
      } catch {}
    }
    if (!nitFilled) throw new Error("No se encontró el campo NIT en la página de DIAN");

    // Resolver CAPTCHA Turnstile con CapSolver (o esperar auto-resolución)
    const captchaToken = await resolverTurnstile(DIAN_RUT_URL);
    if (captchaToken) {
      // Inyectar token en el campo que DIAN lee
      await page.evaluate((token) => {
        // cf-turnstile-response (creado por el widget)
        let el = document.querySelector<HTMLInputElement>('[name="cf-turnstile-response"]');
        if (!el) {
          el = document.createElement("input");
          el.type = "hidden";
          el.name = "cf-turnstile-response";
          document.body.appendChild(el);
        }
        el.value = token;
        // hddToken también lo lee el JS de DIAN
        const hdd = document.querySelector<HTMLInputElement>('[id*="hddToken"]');
        if (hdd) hdd.value = token;
      }, captchaToken);
    } else {
      // Sin CapSolver, esperar 5s por si Turnstile se auto-resuelve
      await new Promise((r) => setTimeout(r, 5000));
    }

    // Click search button — DIAN uses type="image" (submits .x and .y coords)
    const btnSelectors = [
      `[name="${PREFIX}:btnBuscar"]`,
      `input[name*="btnBuscar"]`,
      `[id="${PREFIX}:btnBuscar"]`,
      'input[type="image"][name*="Buscar"]',
      'input[type="submit"]',
    ];
    let clicked = false;
    for (const sel of btnSelectors) {
      try {
        await page.click(sel);
        clicked = true;
        break;
      } catch {}
    }
    if (!clicked) throw new Error("No se encontró el botón de búsqueda en la página de DIAN");

    // Wait for results to appear
    await Promise.race([
      page.waitForNavigation({ waitUntil: "networkidle2", timeout: 20000 }),
      new Promise<void>((resolve) => setTimeout(resolve, 10000)),
    ]).catch(() => {});

    // Small stabilization delay
    await new Promise((r) => setTimeout(r, 800));

    // Extract data from the page (string eval avoids esbuild __name helper issue)
    const raw = await page.evaluate(EXTRACT_SCRIPT) as Record<string, string | null>;

    const result: RutResult = {
      nit,
      dv: raw.dvPage || dv,   // prefer DIAN's DV over locally calculated
      primerApellido: raw.primerApellido || "",
      segundoApellido: raw.segundoApellido || "",
      primerNombre: raw.primerNombre || "",
      otrosNombres: raw.otrosNombres || "",
      razonSocial: raw.razonSocial || "",
      estado: raw.estado || "",
      fechaInscripcion: raw.fechaInscripcion || "",
      fechaActualizacion: raw.fechaActualizacion || "",
      responsabilidades: [],
      error: raw.error || undefined,
    };

    // Derive display name
    if (!result.razonSocial && (result.primerNombre || result.primerApellido)) {
      result.razonSocial = [
        result.primerNombre,
        result.otrosNombres,
        result.primerApellido,
        result.segundoApellido,
      ].filter(Boolean).join(" ");
    }

    if (!result.error) {
      setCache(nit, result);
    }
    return result;

  } catch (err) {
    const msg = (err as Error).message || "Error desconocido";
    return errorResult(nit, dv, `Error consultando DIAN: ${msg}`);
  } finally {
    try { await page.close(); } catch {}
  }
}

function errorResult(nit: string, dv: string, error: string): RutResult {
  return {
    nit, dv,
    primerApellido: "", segundoApellido: "", primerNombre: "", otrosNombres: "",
    razonSocial: "", estado: "", fechaInscripcion: "", fechaActualizacion: "",
    responsabilidades: [], error,
  };
}

// ── Bulk job tracker ──────────────────────────────────────────────────────────
export interface BulkJob {
  status: "running" | "completed" | "error";
  results: RutResult[];
  current: number;
  total: number;
  error?: string;
  createdAt: number;
}

export const bulkJobs = new Map<string, BulkJob>();

// Clean old bulk jobs after 1h
setInterval(() => {
  const cutoff = Date.now() - 60 * 60 * 1000;
  for (const [id, job] of bulkJobs) {
    if (job.createdAt < cutoff) bulkJobs.delete(id);
  }
}, 15 * 60 * 1000);

export async function processBulkJob(jobId: string, nits: string[]): Promise<void> {
  const job = bulkJobs.get(jobId);
  if (!job) return;

  for (let i = 0; i < nits.length; i++) {
    job.current = i + 1;
    try {
      const result = await consultarRUT(nits[i]);
      job.results.push(result);
    } catch (err) {
      const { nit, dv } = normalizarNIT(nits[i]);
      job.results.push(errorResult(nit, dv, (err as Error).message));
    }
    // Small delay between requests to not overload DIAN
    if (i < nits.length - 1) await new Promise((r) => setTimeout(r, 600));
  }

  job.status = "completed";
}
