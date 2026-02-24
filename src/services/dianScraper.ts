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

    // 5) Cambiar a mostrar 50 registros por página (100 causa datos incompletos en DIAN)
    updateProgress({ step: "Optimizando vista (50 registros)..." });
    await setPageLength(page, 50);
    
    // Esperar a que la tabla se actualice con 50 registros
    try {
      await page.waitForSelector("#tableDocuments_processing", { visible: true, timeout: 3000 });
      await page.waitForSelector("#tableDocuments_processing", { hidden: true, timeout: 20000 });
    } catch {
      // Si no aparece el processing, esperar más tiempo
      await delay(2000);
    }
    
    // Esperar a que la tabla tenga todas las filas esperadas
    await waitForFullTableLoad(page, 50);

    // 6) Extracción con paginación
    const allDocuments: DocumentInfo[] = [];
    const seenIds = new Set<string>();
    let pageIndex = 0;

    // Obtener total esperado de la info de paginación
    const expectedTotal = await page.evaluate(() => {
      const info = document.querySelector("#tableDocuments_info, .dataTables_info");
      const text = info?.textContent || "";
      // Buscar patrón "de X registros" o "of X entries"
      const match = text.match(/de\s+([\d,.]+)\s+registros|of\s+([\d,.]+)\s+entries/i);
      if (match) {
        const num = (match[1] || match[2]).replace(/[,.]/g, "");
        return parseInt(num, 10) || 0;
      }
      return 0;
    });
    console.log(`Total esperado según paginación: ${expectedTotal}`);

    while (true) {
      pageIndex++;
      updateProgress({
        step: `Extrayendo lista (página ${pageIndex})...`,
        current: allDocuments.length,
        total: expectedTotal || Math.max(allDocuments.length + 5, 10),
      });

      await delay(500);

      // Contar filas visibles antes de extraer
      const visibleRows = await page.evaluate(() => {
        return document.querySelectorAll("#tableDocuments tbody tr:not(.dataTables_empty)").length;
      });
      console.log(`Página ${pageIndex} - filas visibles: ${visibleRows}`);

      // Verificar si tabla vacía
      const isEmpty = await page.$("tr.dataTables_empty");
      if (isEmpty) {
        console.log(`Página ${pageIndex} - tabla vacía, terminando`);
        break;
      }

      // Extraer documentos de la página actual
      const newDocs = await extractDocsFromPage(page, seenIds);
      allDocuments.push(...newDocs);

      updateProgress({
        current: allDocuments.length,
        total: expectedTotal || allDocuments.length,
      });

      console.log(`Página ${pageIndex} - extraídos: ${newDocs.length}, acumulados: ${allDocuments.length}/${expectedTotal}`);

      // Verificar si ya tenemos todos
      if (expectedTotal > 0 && allDocuments.length >= expectedTotal) {
        console.log(`Alcanzado el total esperado (${expectedTotal}), terminando`);
        break;
      }

      // Intentar ir a la siguiente página
      const hasNext = await goToNextPage(page);
      if (!hasNext) {
        console.log(`No hay más páginas disponibles, terminando`);
        break;
      }

      // Esperar a que cambie la tabla y tenga todas las filas
      await waitForTableChange(page, seenIds);
      await waitForFullTableLoad(page, 50);
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

/**
 * Espera a que la tabla tenga el número esperado de filas según la info de paginación.
 * Esto es crítico porque DIAN a veces devuelve datos parciales en las respuestas AJAX.
 */
async function waitForFullTableLoad(page: Page, pageLength: number): Promise<void> {
  const maxWait = 15000;
  const startTime = Date.now();
  
  while (Date.now() - startTime < maxWait) {
    const { expectedRows, actualRows } = await page.evaluate(() => {
      const info = document.querySelector("#tableDocuments_info, .dataTables_info");
      const text = info?.textContent || "";
      
      // Buscar patrón "del X al Y de Z registros" 
      // Ejemplo: "Mostrando del 1 al 50 de 172 registros"
      const match = text.match(/del\s+(\d+)\s+al\s+(\d+)\s+de\s+([\d,.]+)/i);
      
      let expected = 0;
      if (match) {
        const from = parseInt(match[1], 10);
        const to = parseInt(match[2], 10);
        expected = to - from + 1;
      }
      
      const rows = document.querySelectorAll("#tableDocuments tbody tr:not(.dataTables_empty)");
      return { expectedRows: expected, actualRows: rows.length };
    });
    
    if (expectedRows > 0 && actualRows >= expectedRows) {
      console.log(`Tabla completamente cargada: ${actualRows}/${expectedRows} filas`);
      return;
    }
    
    // También verificar si tenemos el máximo de filas esperado (pageLength)
    if (actualRows >= pageLength) {
      console.log(`Tabla cargada con ${actualRows} filas (máximo por página)`);
      return;
    }
    
    console.log(`Esperando filas: ${actualRows}/${expectedRows || pageLength}`);
    await delay(500);
  }
  
  console.log("Timeout esperando carga completa de tabla, continuando...");
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
    
    // Obtener todas las filas de la tabla (excluyendo la fila vacía)
    const rows = document.querySelectorAll("#tableDocuments tbody tr:not(.dataTables_empty)");
    
    for (const row of rows) {
      let trackId: string | null = null;
      
      // Método 1: Buscar en botones/links de descarga dentro de la fila
      const downloadElements = row.querySelectorAll(
        ".download-document, .download-support-document, .download-eventos, " +
        ".download-equivalente-document, .download-individual-payroll, " +
        "a[href*='DownloadZipFiles'], a[href*='trackId'], " +
        "[data-trackid], [data-id], [id^='doc-'], [id*='track']"
      );
      
      for (const el of downloadElements) {
        trackId = el.id || 
                  el.getAttribute("data-trackid") || 
                  el.getAttribute("data-id") ||
                  el.getAttribute("data-track-id");
        
        if (!trackId) {
          const href = el.getAttribute("href") || "";
          const match = href.match(/trackId=([A-Za-z0-9-]+)/i);
          if (match) trackId = match[1];
        }
        
        if (trackId) break;
      }
      
      // Método 2: Buscar en cualquier elemento de la fila con atributos de tracking
      if (!trackId) {
        const anyWithTrack = row.querySelector("[data-trackid], [data-id], [data-track-id]");
        if (anyWithTrack) {
          trackId = anyWithTrack.getAttribute("data-trackid") || 
                    anyWithTrack.getAttribute("data-id") ||
                    anyWithTrack.getAttribute("data-track-id");
        }
      }
      
      // Método 3: Buscar en la propia fila
      if (!trackId) {
        trackId = row.getAttribute("data-trackid") || 
                  row.getAttribute("data-id") ||
                  row.id;
      }
      
      // Método 4: Buscar en cualquier href dentro de la fila
      if (!trackId) {
        const links = row.querySelectorAll("a[href]");
        for (const link of links) {
          const href = link.getAttribute("href") || "";
          const match = href.match(/trackId=([A-Za-z0-9-]+)/i);
          if (match) {
            trackId = match[1];
            break;
          }
        }
      }
      
      if (!trackId) continue;
      
      // Obtener info de las celdas
      const tds = row.querySelectorAll("td");
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

    // Guardar el número de página actual antes de hacer click
    const currentPageNum = await page.evaluate(() => {
      const active = document.querySelector("#tableDocuments_paginate .paginate_button.current, .paginate_button.active");
      return active?.textContent?.trim() || "0";
    });

    await nextBtn.click();
    
    // Esperar a que cambie el número de página activa
    const startTime = Date.now();
    while (Date.now() - startTime < 10000) {
      const newPageNum = await page.evaluate(() => {
        const active = document.querySelector("#tableDocuments_paginate .paginate_button.current, .paginate_button.active");
        return active?.textContent?.trim() || "0";
      });
      
      if (newPageNum !== currentPageNum) {
        break;
      }
      await delay(200);
    }
    
    // Esperar a que el processing indicator desaparezca
    try {
      await page.waitForSelector("#tableDocuments_processing", { hidden: true, timeout: 15000 });
    } catch {
      // Continuar de todos modos
    }
    
    // Espera adicional para que la tabla se estabilice
    await delay(800);
    
    return true;
  } catch {
    return false;
  }
}

async function waitForTableChange(page: Page, seenIds: Set<string>): Promise<void> {
  // Primero esperar a que el processing indicator desaparezca
  try {
    await page.waitForSelector("#tableDocuments_processing", { hidden: true, timeout: 15000 });
  } catch {
    // Si no hay processing indicator, continuar
  }
  
  // Esperar un momento para que la tabla se estabilice
  await delay(800);
  
  // Luego verificar que hay nuevos IDs
  const startTime = Date.now();
  const timeout = 10000;

  while (Date.now() - startTime < timeout) {
    const hasNewIds = await page.evaluate((seen) => {
      const rows = document.querySelectorAll("#tableDocuments tbody tr:not(.dataTables_empty)");
      
      for (const row of rows) {
        const elements = row.querySelectorAll(
          ".download-document, .download-support-document, .download-eventos, " +
          ".download-equivalente-document, .download-individual-payroll, " +
          "[data-trackid], [data-id], a[href*='trackId']"
        );
        
        for (const el of elements) {
          let tid = el.id || el.getAttribute("data-trackid") || el.getAttribute("data-id");
          
          if (!tid) {
            const href = el.getAttribute("href") || "";
            const match = href.match(/trackId=([A-Za-z0-9-]+)/i);
            if (match) tid = match[1];
          }
          
          if (tid && !seen.includes(tid)) {
            return true;
          }
        }
      }
      return false;
    }, Array.from(seenIds));

    if (hasNewIds) {
      // Esperar un poco más para asegurar que toda la tabla cargó
      await delay(500);
      return;
    }
    await delay(300);
  }

  // Timeout alcanzado, esperar un poco más por si acaso
  await delay(500);
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
