/**
 * Experimento: ¿entrar al token por segunda vez (sesión B) invalida las
 * cookies de la sesión A que ya estaba descargando?
 *
 * Flujo:
 *  1. Sesión A: entra al token, captura cookies, descarga 1 doc → debe OK.
 *  2. Sesión B: entra al MISMO token en otro navegador, captura cookies,
 *     descarga 1 doc → ¿OK?
 *  3. Sesión A: reintenta descargar OTRO doc con sus cookies viejas → si
 *     falla con 403/login, la reentrada invalida sesiones previas.
 *
 * Uso: npx tsx scripts/test-dian-session-invalidation.ts <tokenUrl> <start> <end>
 */
import puppeteer, { type Browser, type Page } from "puppeteer";
import { extractDocumentIdsByCufe, throttledDianDownload, REAL_USER_AGENT } from "../src/services/dianScraper.js";

const tokenUrl = process.argv[2];
const startDate = process.argv[3] || "2026-05-01";
const endDate = process.argv[4] || "2026-05-31";
if (!tokenUrl) { console.error("Falta tokenUrl"); process.exit(1); }

const log = (m: string) => console.log(`[sess ${new Date().toISOString().slice(11, 19)}] ${m}`);

async function tryDownload(label: string, trackId: string, cookies: Record<string, string>): Promise<boolean> {
  const cookieHeader = Object.entries(cookies).map(([k, v]) => `${k}=${v}`).join("; ");
  const url = `https://catalogo-vpfe.dian.gov.co/Document/DownloadZipFiles?trackId=${trackId}`;
  try {
    // 1 solo intento, sin reintentos: queremos ver el estado crudo de la sesión
    const buf = await throttledDianDownload(url, cookieHeader, { maxRetries: 1, timeoutMs: 60000 });
    log(`${label}: OK (${buf.length} bytes)`);
    return true;
  } catch (e) {
    log(`${label}: FALLÓ → ${e instanceof Error ? e.message : e}`);
    return false;
  }
}

async function main() {
  // Obtener lista de docs + cookies de la sesión A vía el flujo normal
  log("Sesión A: entrando al token y resolviendo docs...");
  const { documents, cookies: cookiesA } = await extractDocumentIdsByCufe(
    tokenUrl, startDate, endDate, undefined, "received", undefined, undefined, undefined, 6,
  );
  log(`Sesión A: ${documents.length} docs resueltos, cookies: ${Object.keys(cookiesA).join(",")}`);
  if (documents.length < 4) { log("Necesito >=4 docs"); process.exit(1); }

  const okA1 = await tryDownload("A-descarga-1 (sesión A fresca)", documents[0].id, cookiesA);

  // Sesión B: segundo navegador entrando al MISMO token
  log("Sesión B: entrando al MISMO token en navegador nuevo...");
  const browser: Browser = await puppeteer.launch({
    headless: true,
    args: ["--no-sandbox", "--disable-setuid-sandbox", "--disable-blink-features=AutomationControlled", "--no-zygote"],
  });
  let cookiesB: Record<string, string> = {};
  try {
    const page: Page = await browser.newPage();
    await page.setUserAgent(REAL_USER_AGENT);
    await page.evaluateOnNewDocument(() => { Object.defineProperty(navigator, "webdriver", { get: () => undefined }); });
    await page.goto(tokenUrl, { waitUntil: "domcontentloaded", timeout: 60000 });
    await new Promise((r) => setTimeout(r, 1500));
    log(`Sesión B: URL tras token = ${page.url()}`);
    const arr = await page.cookies();
    cookiesB = Object.fromEntries(arr.map((c) => [c.name, c.value]));
    log(`Sesión B: cookies: ${Object.keys(cookiesB).join(",")}`);
  } finally {
    await browser.close().catch(() => {});
  }

  const okB1 = await tryDownload("B-descarga-1 (sesión B fresca)", documents[1].id, cookiesB);

  // Punto clave: ¿siguen vivas las cookies de A?
  const okA2 = await tryDownload("A-descarga-2 (cookies A tras entrada de B)", documents[2].id, cookiesA);
  await new Promise((r) => setTimeout(r, 3000));
  const okA3 = await tryDownload("A-descarga-3 (cookies A, 3s después)", documents[3].id, cookiesA);

  console.log("RESULT " + JSON.stringify({ okA1, okB1, okA2_trasB: okA2, okA3_trasB: okA3 }));
  process.exit(0);
}

main().catch((e) => { console.error("FATAL", e?.stack || e); process.exit(1); });
