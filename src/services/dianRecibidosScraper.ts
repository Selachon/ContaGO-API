/**
 * Descarga documentos recibidos desde gratis-vpfe.dian.gov.co.
 *
 * Flujo:
 *  1. Abrir TOKEN_URL (catalogo-vpfe) → esperar redirect → copiar cookies.
 *  2. Nueva página endurecida → copiar cookies → /User/RedirectToBiller.
 *  3. /Document/Received → llenar fechas → Buscar.
 *  4. Llamada AJAX directa a /Document/GetReceivedDocuments → obtener total.
 *  5. Segunda llamada AJAX con length=total → lista completa.
 *  6. Por cada documento: GET XML (y opcionalmente PDF) con fetch + cookies.
 */

import puppeteer, { type Browser, type Page } from "puppeteer";
import { resolveExecutablePath, closeBrowserSafely } from "./dianScraper.js";

const UA =
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36";

export interface RecibidoDocument {
  transactionId: string;
  docNumber: string;
  docType: string;
  issueDate: string;
  senderName: string;
  total: string;
}

export interface DownloadedFile {
  name: string;
  buffer: Buffer;
  mimeType: string;
}

export interface DownloadProgress {
  step: string;
  current?: number;
  total?: number;
  pct?: number;
}

export type ProgressCallback = (p: DownloadProgress) => void;

function delay(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}

async function hardenPage(page: Page): Promise<void> {
  await page.setUserAgent(UA);
  await page.setExtraHTTPHeaders({ "Accept-Language": "es-CO,es;q=0.9" });
  await page.evaluateOnNewDocument(() => {
    try {
      Object.defineProperty(navigator, "webdriver", { get: () => undefined });
    } catch {}
  });
}

/** Construye body de DataTables para GetReceivedDocuments o GetSentDocuments */
function buildDtParams(start: number, length: number, draw: number, direction: "received" | "sent" = "received"): string {
  const cols = direction === "sent"
    ? ["rowCheck", "docTypeName", "docNumber", "issueDate", "receiverCode", "receiverName", "total", "paymentMeans", "eventStatus", "actions"]
    : ["rowCheck", "docTypeName", "docNumber", "issueDate", "senderCode", "senderName", "total", "paymentMeans", "eventStatus", "actions"];
  const p = new URLSearchParams();
  p.set("draw", String(draw));
  cols.forEach((col, i) => {
    p.set(`columns[${i}][data]`, col);
    p.set(`columns[${i}][name]`, "");
    p.set(`columns[${i}][searchable]`, "true");
    p.set(`columns[${i}][orderable]`, "false");
    p.set(`columns[${i}][search][value]`, "");
    p.set(`columns[${i}][search][regex]`, "false");
  });
  p.set("order[0][column]", "3");
  p.set("order[0][dir]", "desc");
  p.set("start", String(start));
  p.set("length", String(length));
  p.set("search[value]", "");
  p.set("search[regex]", "false");
  return p.toString();
}

/** Lanza Chromium con los args mínimos necesarios */
async function launchBrowser(): Promise<Browser> {
  const executablePath = resolveExecutablePath() ?? undefined;
  return puppeteer.launch({
    headless: true,
    args: [
      "--no-sandbox",
      "--disable-setuid-sandbox",
      "--disable-blink-features=AutomationControlled",
      "--disable-dev-shm-usage",
      "--disable-gpu",
      "--disable-extensions",
      "--no-first-run",
    ],
    executablePath,
  });
}

/** Limita a max 20 descargas por minuto de forma global por instancia. */
class RateLimiter {
  private count = 0;
  private windowStart = Date.now();
  constructor(private readonly max: number) {}
  async throttle(): Promise<void> {
    const now = Date.now();
    if (now - this.windowStart >= 60_000) {
      this.count = 0;
      this.windowStart = now;
    }
    if (this.count >= this.max) {
      const wait = 60_000 - (now - this.windowStart);
      if (wait > 0) await delay(wait);
      this.count = 0;
      this.windowStart = Date.now();
    }
    this.count++;
  }
}

/**
 * Autentica vía tokenUrl y devuelve la página de gratis-vpfe lista para
 * interactuar (ya posicionada en /Document/Received o /Document/Sent con el filtro aplicado).
 */
/**
 * Llena From/To y hace clic en "Buscar" en la página YA autenticada de
 * gratis-vpfe (Document/Received o /Sent). Reutilizable para re-consultar
 * sub-rangos de fecha dentro de la MISMA sesión (sin reabrir el token), porque
 * `GetReceivedDocuments`/`GetSentDocuments` no pagina de verdad más allá de 150
 * filas — la única forma de traer más de 150 documentos es acotar el rango.
 */
export async function applyReceivedDateFilter(page: Page, from: string, to: string): Promise<void> {
  await page.evaluate((f: string, t: string) => {
    const $ = (window as any).$;
    const fromEl = document.getElementById("From") as HTMLInputElement | null;
    const toEl = document.getElementById("To") as HTMLInputElement | null;
    const rangeEl = document.getElementById("idDateRange") as HTMLInputElement | null;
    if (fromEl) { fromEl.value = f; if ($) $(fromEl).trigger("change"); }
    if (toEl) { toEl.value = t; if ($) $(toEl).trigger("change"); }
    if (rangeEl) { rangeEl.value = `${f} a ${t}`; }
  }, from, to);
  await delay(300);
  const navDone = page.waitForNavigation({ waitUntil: "networkidle2", timeout: 20_000 }).catch(() => null);
  await page.evaluate(() => {
    const btn = Array.from(document.querySelectorAll("button")).find(
      (b) => (b.textContent || "").trim().toLowerCase().includes("buscar")
    );
    if (btn) btn.click();
  });
  await navDone;
  await delay(1500);
}

export async function authenticateAndNavigate(
  tokenUrl: string,
  from: string,
  to: string,
  progress: ProgressCallback,
  direction: "received" | "sent" = "received"
): Promise<{ browser: Browser; page: Page }> {
  progress({ step: "Iniciando navegador..." });
  const browser = await launchBrowser();

  try {
    // ── Página de autenticación (catalogo-vpfe) ──────────────────────────
    const authPage = await browser.newPage();
    await hardenPage(authPage);
    authPage.setDefaultTimeout(60_000);

    progress({ step: "Autenticando con DIAN..." });
    await authPage.goto(tokenUrl, { waitUntil: "domcontentloaded" });

    // Esperar que el AuthToken redirija
    const ts = Date.now();
    while (Date.now() - ts < 30_000) {
      if (!/\/User\/AuthToken/i.test(authPage.url())) break;
      await delay(1000);
    }
    if (/\/User\/AuthToken/i.test(authPage.url())) {
      throw new Error("El token no redirigió — puede estar expirado o ser de otra IP.");
    }
    await delay(2000);
    const cookies = await authPage.cookies();

    // ── Página endurecida para gratis-vpfe ──────────────────────────────
    const page = await browser.newPage();
    await hardenPage(page);
    page.setDefaultTimeout(60_000);
    page.setDefaultNavigationTimeout(60_000);
    await page.setViewport({ width: 1440, height: 900 });
    if (cookies.length) await page.setCookie(...(cookies as Parameters<typeof page.setCookie>));

    progress({ step: "Navegando al portal DIAN..." });
    await page.goto("https://catalogo-vpfe.dian.gov.co/User/RedirectToBiller", {
      waitUntil: "networkidle2",
      timeout: 30_000,
    }).catch(() => {});
    await delay(2000);

    // ── Formulario de documentos recibidos / emitidos ────────────────────
    const docPath = direction === "sent" ? "Sent" : "Received";
    progress({ step: direction === "sent" ? "Abriendo Documentos Emitidos..." : "Abriendo Documentos Recibidos..." });
    await page.goto(`https://gratis-vpfe.dian.gov.co/Document/${docPath}`, {
      waitUntil: "networkidle2",
      timeout: 30_000,
    }).catch(() => {});
    await delay(2000);

    progress({ step: "Buscando documentos..." });
    await applyReceivedDateFilter(page, from, to);

    return { browser, page };
  } catch (err) {
    await closeBrowserSafely(browser).catch(() => undefined);
    throw err;
  }
}

/** Llama al endpoint DataTables paginando hasta obtener todos los registros */
export async function fetchDocumentList(page: Page, direction: "received" | "sent" = "received"): Promise<RecibidoDocument[]> {
  const endpoint = direction === "sent" ? "/Document/GetSentDocuments" : "/Document/GetReceivedDocuments";
  const PAGE_SIZE = 150; // límite real del servidor DIAN

  const fetchPage = async (start: number, length: number, draw: number): Promise<{ data: any[]; recordsTotal: number; recordsFiltered: number }> => {
    return page.evaluate(async (params: string, url: string) => {
      const resp = await fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
          "X-Requested-With": "XMLHttpRequest",
        },
        body: params,
      });
      const json = await resp.json();
      return { data: json.data as any[], recordsTotal: json.recordsTotal as number, recordsFiltered: json.recordsFiltered as number };
    }, buildDtParams(start, length, draw, direction), endpoint);
  };

  // Primera página: primeros registros. NO confiamos en recordsFiltered/recordsTotal
  // para decidir cuándo parar: el servidor DIAN los reporta mal cuando hay más de
  // PAGE_SIZE documentos en el rango, lo que truncaba silenciosamente meses con
  // >150 documentos (se perdían los más antiguos del rango, al final del orden
  // DESC por fecha). Pero TAMPOCO podemos confiar en "la página vino llena
  // (==PAGE_SIZE) → sigo pidiendo": si el servidor ignora `start` y siempre
  // devuelve la MISMA primera página, eso deja el loop girando indefinidamente
  // con páginas idénticas. Criterio robusto: seguir mientras la página traiga
  // AL MENOS UNA fila nueva (por id, `DT_RowId`); si una página no aporta nada
  // nuevo, el servidor no está paginando de verdad y hay que parar ahí.
  const rowId = (row: any): string => row?.DT_RowId || row?.DT_RowData?.pkey || "";
  const first = await fetchPage(0, PAGE_SIZE, 1);
  const allRows: any[] = [];
  const seenIds = new Set<string>();
  for (const r of first.data || []) {
    const id = rowId(r);
    if (id && seenIds.has(id)) continue;
    if (id) seenIds.add(id);
    allRows.push(r);
  }
  console.log(`[dianRecibidos] página 1: ${allRows.length} fila(s)`);
  if (allRows.length === 0) return [];

  let draw = 2;
  let start = PAGE_SIZE;
  let lastPageFull = (first.data || []).length === PAGE_SIZE;
  const MAX_PAGES = 30; // salvaguarda: 30 * 150 = 4,500 documentos por mes
  while (lastPageFull && draw <= MAX_PAGES) {
    await new Promise<void>((r) => setTimeout(r, 400)); // pausa entre páginas
    const page_ = await fetchPage(start, PAGE_SIZE, draw);
    const batch = page_.data || [];
    lastPageFull = batch.length === PAGE_SIZE;
    let newCount = 0;
    for (const r of batch) {
      const id = rowId(r);
      if (id && seenIds.has(id)) continue;
      if (id) seenIds.add(id);
      allRows.push(r);
      newCount++;
    }
    console.log(`[dianRecibidos] página ${draw} (start=${start}): ${batch.length} fila(s), ${newCount} nueva(s)`);
    draw++;
    if (newCount === 0) break; // el servidor repitió la página: no pagina de verdad, paramos
    start += PAGE_SIZE;
  }

  const parseRow = (row: any): RecibidoDocument => {
    const txId: string = row.DT_RowId || row.DT_RowData?.pkey || "";
    const dateM = (row.docDate || "").match(/>([\d/]+)</);
    const senderM = (row.docReceiverName || row.docSenderName || "").match(/>([^<]+)</);
    const totalM = (row.docTotalAmount || "").match(/>([\s\S]+?)</);
    return {
      transactionId: txId,
      docNumber: row.docNumber || row.DT_RowSerie || "",
      docType: row.docTypeName || "",
      issueDate: dateM ? dateM[1] : (row.DT_RowDate || ""),
      senderName: senderM ? senderM[1].trim() : "",
      total: totalM ? totalM[1].trim() : "",
    };
  };

  return allRows.map(parseRow);
}

/** Descarga un XML y opcionalmente el PDF para un transactionId dado */
export async function downloadXmlFile(
  page: Page,
  txId: string,
  docNumber: string,
  includePdf: boolean
): Promise<DownloadedFile[]> {
  const cookies = await page.cookies();
  const cookieHeader = cookies.map((c) => `${c.name}=${c.value}`).join("; ");
  const base = "https://gratis-vpfe.dian.gov.co";

  const xmlUrl = `${base}/Document/DownloadXml?transactionId=${txId}&type=2`;
  const safeName = docNumber.replace(/[^a-zA-Z0-9_\-]/g, "_");

  const xmlResp = await fetch(xmlUrl, {
    headers: { "User-Agent": UA, Cookie: cookieHeader },
  });
  if (!xmlResp.ok) {
    throw new Error(`XML ${docNumber}: HTTP ${xmlResp.status}`);
  }
  const xmlBuf = Buffer.from(await xmlResp.arrayBuffer());
  const files: DownloadedFile[] = [
    { name: `${safeName}.xml`, buffer: xmlBuf, mimeType: "application/xml" },
  ];

  if (includePdf) {
    const pdfUrl = `${base}/IoFacturo/Print/PrintStoragePdf?transactionId=${txId}&viewMode=attachment`;
    const pdfResp = await fetch(pdfUrl, {
      headers: { "User-Agent": UA, Cookie: cookieHeader },
    });
    if (pdfResp.ok) {
      const pdfBuf = Buffer.from(await pdfResp.arrayBuffer());
      if (pdfBuf[0] === 0x25 && pdfBuf[1] === 0x50) {
        files.push({ name: `${safeName}.pdf`, buffer: pdfBuf, mimeType: "application/pdf" });
      }
    }
  }

  return files;
}

/** Punto de entrada principal */
export async function downloadReceivedDocuments(opts: {
  tokenUrl: string;
  from: string;   // dd/mm/yyyy
  to: string;     // dd/mm/yyyy
  includePdf?: boolean;
  concurrency?: number;
  rateLimit?: number;  // max docs/min (default: 20)
  documentDirection?: "received" | "sent";
  progress?: ProgressCallback;
  isCancelled?: () => boolean;
}): Promise<{ files: DownloadedFile[]; docs: RecibidoDocument[] }> {
  const {
    tokenUrl,
    from,
    to,
    includePdf = false,
    concurrency = 3,
    rateLimit: maxPerMin = 20,
    progress = () => {},
    isCancelled = () => false,
  } = opts;

  let browser: Browser | null = null;
  try {
    const dir = opts.documentDirection ?? "received";
    const { browser: b, page } = await authenticateAndNavigate(tokenUrl, from, to, progress, dir);
    browser = b;

    progress({ step: "Listando documentos..." });
    const docs = await fetchDocumentList(page, dir);

    if (docs.length === 0) {
      return { files: [], docs: [] };
    }

    progress({ step: `Descargando ${docs.length} documentos...`, current: 0, total: docs.length, pct: 0 });

    const allFiles: DownloadedFile[] = [];
    let done = 0;
    const limiter = new RateLimiter(maxPerMin);

    // Descargas con concurrencia limitada y rate limiting
    const queue = [...docs];
    const workers = Array.from({ length: Math.min(concurrency, docs.length) }, async () => {
      while (queue.length > 0) {
        if (isCancelled()) return;
        const doc = queue.shift();
        if (!doc) break;
        if (!doc.transactionId) { done++; continue; }
        await limiter.throttle();
        if (isCancelled()) return;
        try {
          const downloaded = await downloadXmlFile(page, doc.transactionId, doc.docNumber, includePdf ?? false);
          allFiles.push(...downloaded);
        } catch (err) {
          console.warn(`[dianRecibidos] Error descargando ${doc.docNumber}:`, err);
        }
        done++;
        const pct = Math.round((done / docs.length) * 100);
        progress({ step: `Descargando... (${done}/${docs.length})`, current: done, total: docs.length, pct });
      }
    });
    await Promise.all(workers);

    return { files: allFiles, docs };
  } finally {
    if (browser) await closeBrowserSafely(browser).catch(() => undefined);
  }
}
