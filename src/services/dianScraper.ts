import puppeteer, { Browser, Page, Cookie } from "puppeteer";
import fs from "fs";
import type { DocumentInfo, ProgressData } from "../types/dian.js";

// Progress tracker compartido
export const progressTracker: Map<string, ProgressData> = new Map();

interface ExtractionResult {
  documents: DocumentInfo[];
  cookies: Record<string, string>;
}

/**
 * Extrae IDs de documentos de la DIAN usando Puppeteer
 * Traducido de Python/Selenium a TypeScript/Puppeteer
 */
export async function extractDocumentIds(
  tokenUrl: string,
  startDate: string | undefined,
  endDate: string | undefined,
  progressUid?: string
): Promise<ExtractionResult> {
  const updateProgress = (data: Partial<ProgressData>) => {
    if (progressUid) {
      const current = progressTracker.get(progressUid) || { step: "", current: 0, total: 0 };
      progressTracker.set(progressUid, { ...current, ...data });
    }
  };

  updateProgress({ step: "Iniciando navegador...", current: 0, total: 0 });

  let browser: Browser | null = null;

  try {
    const executablePath = resolveExecutablePath();

    browser = await puppeteer.launch({
      headless: true,
      args: [
        "--no-sandbox",
        "--disable-setuid-sandbox",
        "--disable-dev-shm-usage",
        "--disable-gpu",
      ],
      executablePath: executablePath || undefined,
    });

    const page = await browser.newPage();
    await page.setViewport({ width: 1280, height: 800 });

    // Timeout generoso para conexiones lentas
    page.setDefaultTimeout(60000);

    // 1) Acceso inicial con token
    updateProgress({ step: "Accediendo con token..." });
    await page.goto(tokenUrl, { waitUntil: "networkidle2" });
    await delay(1000);

    // 2) Ir a la sección Recibidos
    updateProgress({ step: "Navegando a documentos recibidos..." });
    await page.goto("https://catalogo-vpfe.dian.gov.co/Document/Received", {
      waitUntil: "networkidle2",
    });
    await delay(600);

    updateProgress({ step: "Extrayendo lista (iniciando)...", current: 0, total: 0 });

    // 3) Aplicar filtros de fecha si vienen
    if (startDate && endDate) {
      updateProgress({ step: "Aplicando rango de fechas..." });
      await applyDateFilter(page, startDate, endDate);
    }

    // 4) Esperar a que carguen resultados
    updateProgress({ step: "Cargando resultados..." });
    await waitForTableLoad(page);

    // 5) Cambiar a mostrar 100 registros por página
    updateProgress({ step: "Optimizando vista (100 registros)..." });
    await setPageLength(page, 100);
    await delay(1100);

    // 6) Extracción con paginación
    const allDocuments: DocumentInfo[] = [];
    const seenIds = new Set<string>();
    let pageIndex = 0;

    while (true) {
      pageIndex++;
      updateProgress({
        step: `Extrayendo lista (página ${pageIndex})...`,
        current: allDocuments.length,
        total: Math.max(allDocuments.length + 5, 10),
      });

      await delay(250);

      // Verificar si tabla vacía
      const isEmpty = await page.$("tr.dataTables_empty");
      if (isEmpty) break;

      // Extraer documentos de la página actual
      const newDocs = await extractDocsFromPage(page, seenIds);
      allDocuments.push(...newDocs);

      updateProgress({
        current: allDocuments.length,
        total: allDocuments.length,
      });

      console.log(`Página ${pageIndex} - nuevos IDs: ${newDocs.length}, acumulados: ${allDocuments.length}`);

      // Intentar ir a la siguiente página
      const hasNext = await goToNextPage(page);
      if (!hasNext) break;

      // Esperar a que cambie la tabla
      await waitForTableChange(page, seenIds);
    }

    // Obtener cookies para las descargas
    const cookies = await page.cookies();
    const cookieMap: Record<string, string> = {};
    for (const c of cookies) {
      cookieMap[c.name] = c.value;
    }

    updateProgress({
      step: "Lista extraída",
      current: allDocuments.length,
      total: allDocuments.length,
    });

    console.log(`Total documentos encontrados: ${allDocuments.length}`);

    return { documents: allDocuments, cookies: cookieMap };
  } finally {
    if (browser) {
      await browser.close();
    }
  }
}

function resolveExecutablePath(): string | null {
  if (!process.env.PUPPETEER_CACHE_DIR) {
    process.env.PUPPETEER_CACHE_DIR = `${process.cwd()}/.cache/puppeteer`;
  }

  const envPath = process.env.PUPPETEER_EXECUTABLE_PATH;
  if (envPath && !fs.existsSync(envPath)) {
    // Limpiar env si apunta a un binario que no existe
    delete process.env.PUPPETEER_EXECUTABLE_PATH;
  }

  const candidates = [
    envPath,
    "/usr/bin/chromium-browser",
    "/usr/bin/chromium",
    "/usr/bin/google-chrome",
    "/usr/bin/google-chrome-stable",
  ].filter(Boolean) as string[];

  for (const candidate of candidates) {
    if (fs.existsSync(candidate)) {
      // Asegurar que Puppeteer no intente usar un path inválido del env
      process.env.PUPPETEER_EXECUTABLE_PATH = candidate;
      return candidate;
    }
  }

  // Si no hay binario del sistema, borrar para que Puppeteer use el descargado
  delete process.env.PUPPETEER_EXECUTABLE_PATH;

  return null;
}

async function applyDateFilter(page: Page, startDate: string, endDate: string): Promise<void> {
  try {
    // Intentar encontrar el input de rango de fechas
    const rangeInput = await page.$(
      "#dashboard-report-range input, " +
      "input[placeholder*='Rango'], " +
      "input[aria-label*='Rango'], " +
      "input[placeholder*='Fecha'], " +
      "input[aria-label*='Fecha'], " +
      "input[name*='date'], " +
      "input[name*='fecha'], " +
      "input[id*='range']"
    );

    if (rangeInput) {
      const sDate = startDate.replace(/-/g, "/");
      const eDate = endDate.replace(/-/g, "/");
      const rangoCompleto = `${sDate} - ${eDate}`;

      await rangeInput.click({ clickCount: 3 }); // Seleccionar todo
      await rangeInput.type(rangoCompleto);
      await page.keyboard.press("Enter");
      await delay(450);
    } else {
      console.log("No se encontró input de rango - continuando sin escribir fechas.");
    }

    // Intentar click en botón 'Buscar'
    const clicked = await page.evaluate(() => {
      const buttons = Array.from(document.querySelectorAll("button"));
      const btn = buttons.find((b) => b.textContent?.trim().includes("Buscar"));
      if (btn) {
        (btn as HTMLElement).click();
        return true;
      }
      return false;
    });
    if (clicked) {
      await delay(450);
    }
  } catch (err) {
    console.error("Error aplicando rango de fechas:", err);
  }
}

async function waitForTableLoad(page: Page): Promise<void> {
  try {
    // Esperar a que desaparezca el overlay de procesamiento
    try {
      await page.waitForSelector("#tableDocuments_processing", { visible: true, timeout: 5000 });
      await page.waitForSelector("#tableDocuments_processing", { hidden: true, timeout: 20000 });
    } catch {
      // Overlay no detectado, esperar filas directamente
      await page.waitForSelector("table#tableDocuments tbody tr", { timeout: 6000 }).catch(() => {});
    }
  } catch (err) {
    console.log("Espera resultados:", err);
  }
}

async function setPageLength(page: Page, length: number): Promise<void> {
  try {
    // Intentar encontrar el select de longitud de página
    const selectHandle = await page.$(
      "select[name='tableDocuments_length'], " +
      "#tableDocuments_length select, " +
      "select[name*='length']"
    );

    if (selectHandle) {
      await selectHandle.select(length.toString());
      await delay(500);
    } else {
      // Fallback: usar DataTables API directamente
      await page.evaluate((len) => {
        try {
          const $ = (window as any).$;
          if ($ && $.fn && $.fn.dataTable) {
            const dt = $("#tableDocuments").DataTable();
            dt.page.len(len).draw();
          }
        } catch {}
      }, length);
    }
  } catch (err) {
    console.error("Error cambiando a 100 registros:", err);
  }
}

async function extractDocsFromPage(page: Page, seenIds: Set<string>): Promise<DocumentInfo[]> {
  const docs: DocumentInfo[] = [];

  const items = await page.evaluate(() => {
    const results: Array<{ id: string; docnum: string; nit: string }> = [];
    
    const buttons = document.querySelectorAll(
      ".download-document, .download-support-document, .download-eventos, " +
      ".download-equivalente-document, .download-individual-payroll, " +
      "a[href*='DownloadZipFiles'], a[data-trackid], button[data-trackid]"
    );

    for (const btn of buttons) {
      let trackId = btn.id || btn.getAttribute("data-trackid") || btn.getAttribute("data-id");
      
      if (!trackId) {
        const href = btn.getAttribute("href") || "";
        const match = href.match(/trackId=([A-Za-z0-9-]+)/);
        if (match) trackId = match[1];
      }

      if (!trackId) continue;

      // Obtener info de la fila
      const row = btn.closest("tr");
      const tds = row?.querySelectorAll("td") || [];
      const docnum = tds[4]?.textContent?.trim() || "";
      const nit = tds[6]?.textContent?.trim() || "";

      results.push({ id: trackId, docnum, nit });
    }

    return results;
  });

  for (const item of items) {
    if (!seenIds.has(item.id)) {
      seenIds.add(item.id);
      docs.push(item);
    }
  }

  return docs;
}

async function goToNextPage(page: Page): Promise<boolean> {
  try {
    const nextBtn = await page.$(
      "#tableDocuments_next, .paginate_button.next, a.next"
    );

    if (!nextBtn) return false;

    const isDisabled = await page.evaluate((el) => {
      return el?.classList.contains("disabled") || false;
    }, nextBtn);

    if (isDisabled) return false;

    await nextBtn.click();
    return true;
  } catch {
    return false;
  }
}

async function waitForTableChange(page: Page, seenIds: Set<string>): Promise<void> {
  const startTime = Date.now();
  const timeout = 6000;

  while (Date.now() - startTime < timeout) {
    const hasNewIds = await page.evaluate((seen) => {
      const buttons = document.querySelectorAll(
        ".download-document, .download-support-document, .download-eventos, " +
        ".download-equivalente-document, .download-individual-payroll"
      );

      for (const btn of buttons) {
        const tid = btn.id || btn.getAttribute("data-trackid") || btn.getAttribute("data-id");
        if (tid && !seen.includes(tid)) {
          return true;
        }
      }
      return false;
    }, Array.from(seenIds));

    if (hasNewIds) return;
    await delay(250);
  }

  await delay(200);
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
