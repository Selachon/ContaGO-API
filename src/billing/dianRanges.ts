import { launchBrowserWithRetry, resolveExecutablePath, closeBrowserSafely } from "../services/dianScraper.js";
import { assertTokenMatchesNit } from "./dianToken.js";

/**
 * Rangos de numeración de PRUEBA que el portal de habilitación de la DIAN muestra para el modo de operación con Alegra
 * (Registro y habilitación → Facturación electrónica → Configurar modos de operación → Rangos de prueba).
 * Frágil por naturaleza (lee pantallas de la DIAN): por eso la entrada manual siempre queda como alternativa.
 */

export interface DianRange {
  prefix: string; resolutionNumber: string; resolutionDate?: string; technicalKey: string;
  minNumber: number; maxNumber: number; startDate: string; endDate: string;
}

const isoDate = (d: string) => { const m = /^(\d{2})-(\d{2})-(\d{4})$/.exec(d.trim()); return m ? `${m[3]}-${m[2]}-${m[1]}` : d.trim(); };

/** Filas de la tabla «Prefijo | Nº resolución | Fecha resolución | Llave técnica | Desde | Hasta | Fecha desde | Fecha hasta» → rangos. */
export function parseRangeRows(rows: string[][]): DianRange[] {
  const out: DianRange[] = [];
  for (const r of rows) {
    const c = r.map(x => x.trim());
    if (c.length < 8 || !c[0] || !/^\d+$/.test(c[4]) || !/^\d+$/.test(c[5])) continue; // encabezados o filas vacías
    out.push({ prefix: c[0], resolutionNumber: c[1], resolutionDate: c[2], technicalKey: c[3], minNumber: Number(c[4]), maxNumber: Number(c[5]), startDate: isoDate(c[6]), endDate: isoDate(c[7]) });
  }
  return out;
}

export class DianRangesError extends Error { constructor(message: string, readonly code: string) { super(message); } }

/** Entra al portal con el token (no se guarda ni se imprime) y devuelve los rangos de prueba del modo de operación de Alegra. */
export async function fetchDianTestRanges(tokenUrl: string, nit: string): Promise<DianRange[]> {
  assertTokenMatchesNit(tokenUrl, nit);
  const browser = await launchBrowserWithRetry(resolveExecutablePath(), () => {});
  try {
    const page = await browser.newPage();
    await page.setUserAgent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36");
    await page.evaluateOnNewDocument(() => { Object.defineProperty(navigator, "webdriver", { get: () => undefined }); });
    page.setDefaultNavigationTimeout(60_000);
    await page.goto(tokenUrl, { waitUntil: "networkidle2" });
    if (/\/User\/Login/i.test(page.url())) throw new DianRangesError("El enlace de la DIAN expiró (dura 60 minutos). Genera uno nuevo.", "TOKEN_EXPIRED");

    await page.goto("https://catalogo-vpfe-hab.dian.gov.co/Contributor/Check", { waitUntil: "networkidle2" });
    const modesPath = await page.$$eval("a", as => as.map(a => a.getAttribute("href") || "").find(h => /ConfigureOperationModes/i.test(h)) || "");
    if (!modesPath) throw new DianRangesError("No encontré «Configurar modos de operación» en la DIAN. ¿Ya te registraste como facturador electrónico?", "NO_MODES");
    await page.goto(`https://catalogo-vpfe-hab.dian.gov.co${modesPath}`, { waitUntil: "networkidle2" });

    // Filas de modos de operación: solo las del proveedor tecnológico Alegra
    const links: string[] = await page.$$eval("tr", trs => trs
      .filter(tr => /alegra/i.test(tr.textContent || ""))
      .flatMap(tr => [...tr.querySelectorAll("a")].map(a => a.getAttribute("href") || "").filter(h => /NumberRange\/ListTests/i.test(h))));
    const unique = [...new Set(links)];
    if (!unique.length) throw new DianRangesError("En la DIAN no aparece asociado el proveedor «Soluciones Alegra S.A.S». Asóciaselo en «Configurar modos de operación» y vuelve a intentar.", "NOT_ASSOCIATED");

    const ranges: DianRange[] = [];
    for (const path of unique) {
      await page.goto(`https://catalogo-vpfe-hab.dian.gov.co${path}`, { waitUntil: "networkidle2" });
      const rows = await page.$$eval("table tr", trs => trs.map(tr => [...tr.querySelectorAll("td")].map(td => (td.textContent || "").trim())));
      ranges.push(...parseRangeRows(rows));
    }
    const seen = new Set<string>();
    return ranges.filter(r => { const k = `${r.prefix}|${r.resolutionNumber}`; if (seen.has(k)) return false; seen.add(k); return true; });
  } finally {
    await closeBrowserSafely(browser);
  }
}
