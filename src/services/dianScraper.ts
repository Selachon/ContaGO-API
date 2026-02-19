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
    let consecutiveEmptyPages = 0;
    const MAX_EMPTY_PAGES = 2;

    // Debug: capturar info de la tabla
    const tableInfo = await page.evaluate(() => {
      const table = document.querySelector("#tableDocuments, table.dataTable, table");
      const rows = document.querySelectorAll("#tableDocuments tbody tr, table.dataTable tbody tr");
      const emptyRow = document.querySelector("tr.dataTables_empty, td.dataTables_empty");
      const infoText = document.querySelector(".dataTables_info, #tableDocuments_info")?.textContent || "";
      
      return {
        tableExists: !!table,
        rowCount: rows.length,
        hasEmptyIndicator: !!emptyRow,
        emptyText: emptyRow?.textContent || "",
        infoText,
        firstRowHtml: rows[0]?.innerHTML?.substring(0, 200) || "N/A"
      };
    });
    console.log("Estado inicial de tabla:", JSON.stringify(tableInfo));

    while (true) {
      pageIndex++;
      updateProgress({
        step: `Extrayendo lista (página ${pageIndex})...`,
        current: allDocuments.length,
        total: Math.max(allDocuments.length + 5, 10),
      });

      await delay(300);

      // Verificar si tabla vacía con múltiples selectores
      const isEmpty = await page.evaluate(() => {
        const emptyRow = document.querySelector("tr.dataTables_empty, td.dataTables_empty");
        const noDataText = document.body.innerText.includes("No hay datos disponibles") ||
                          document.body.innerText.includes("No data available") ||
                          document.body.innerText.includes("No se encontraron");
        const rows = document.querySelectorAll("#tableDocuments tbody tr:not(.dataTables_empty)");
        return (!!emptyRow || noDataText) && rows.length === 0;
      });
      
      if (isEmpty) {
        console.log(`Página ${pageIndex}: tabla vacía detectada`);
        break;
      }

      // Extraer documentos de la página actual
      const newDocs = await extractDocsFromPage(page, seenIds);
      
      if (newDocs.length === 0) {
        consecutiveEmptyPages++;
        console.log(`Página ${pageIndex}: 0 nuevos docs (consecutivos vacíos: ${consecutiveEmptyPages})`);
        if (consecutiveEmptyPages >= MAX_EMPTY_PAGES) {
          console.log("Deteniendo: demasiadas páginas consecutivas sin documentos nuevos");
          break;
        }
      } else {
        consecutiveEmptyPages = 0;
        allDocuments.push(...newDocs);
      }

      updateProgress({
        current: allDocuments.length,
        total: allDocuments.length,
      });

      console.log(`Página ${pageIndex} - nuevos IDs: ${newDocs.length}, acumulados: ${allDocuments.length}`);

      // Intentar ir a la siguiente página
      const hasNext = await goToNextPage(page);
      if (!hasNext) {
        console.log("No hay más páginas");
        break;
      }

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
    // Convertir YYYY-MM-DD a DD/MM/YYYY (formato que usa la DIAN)
    const formatToDian = (dateStr: string): string => {
      const [year, month, day] = dateStr.split("-");
      return `${day}/${month}/${year}`;
    };

    const sDate = formatToDian(startDate);
    const eDate = formatToDian(endDate);
    const rangoCompleto = `${sDate} - ${eDate}`;

    console.log(`Aplicando filtro de fechas: ${rangoCompleto}`);

    // Intentar encontrar el input de rango de fechas con selectores más específicos
    const rangeInput = await page.$(
      "#dashboard-report-range input, " +
      "#dashboard-report-range, " +
      "input[placeholder*='Rango'], " +
      "input[aria-label*='Rango'], " +
      ".daterangepicker-input, " +
      "input.form-control[name*='date'], " +
      "input[id*='range'], " +
      "input[id*='date']"
    );

    if (rangeInput) {
      console.log("Input de rango encontrado, escribiendo fechas...");
      
      // Limpiar el campo primero
      await rangeInput.click({ clickCount: 3 });
      await page.keyboard.press("Backspace");
      await delay(100);
      
      // Escribir el rango
      await rangeInput.type(rangoCompleto, { delay: 50 });
      await page.keyboard.press("Enter");
      await delay(800);
    } else {
      console.log("No se encontró input de rango directo. Intentando con daterangepicker...");
      
      // Intentar hacer click en el div del daterangepicker para abrirlo
      const dateRangeDiv = await page.$("#dashboard-report-range");
      if (dateRangeDiv) {
        await dateRangeDiv.click();
        await delay(500);
        
        // Buscar inputs de fecha inicio y fin dentro del picker
        const startInput = await page.$(".daterangepicker .drp-calendar.left input, .daterangepicker input[name='daterangepicker_start']");
        const endInput = await page.$(".daterangepicker .drp-calendar.right input, .daterangepicker input[name='daterangepicker_end']");
        
        if (startInput && endInput) {
          await startInput.click({ clickCount: 3 });
          await startInput.type(sDate);
          await endInput.click({ clickCount: 3 });
          await endInput.type(eDate);
          
          // Click en Apply/Aplicar
          const applyBtn = await page.$(".daterangepicker .applyBtn, .daterangepicker button.btn-primary");
          if (applyBtn) {
            await applyBtn.click();
            await delay(500);
          }
        }
      } else {
        console.log("No se encontró daterangepicker div.");
      }
    }

    // Intentar click en botón 'Buscar' o 'Filtrar'
    const clicked = await page.evaluate(() => {
      const buttons = Array.from(document.querySelectorAll("button, input[type='submit'], a.btn"));
      const btn = buttons.find((b) => {
        const text = b.textContent?.trim().toLowerCase() || "";
        return text.includes("buscar") || text.includes("filtrar") || text.includes("aplicar");
      });
      if (btn) {
        (btn as HTMLElement).click();
        return true;
      }
      return false;
    });
    
    if (clicked) {
      console.log("Botón de búsqueda clickeado");
      await delay(800);
    } else {
      console.log("No se encontró botón de búsqueda explícito");
    }

    // Esperar a que la tabla se actualice
    await delay(1000);
    
  } catch (err) {
    console.error("Error aplicando rango de fechas:", err);
  }
}

async function waitForTableLoad(page: Page): Promise<void> {
  try {
    console.log("Esperando carga de tabla...");
    
    // Esperar a que desaparezca el overlay de procesamiento
    try {
      const processingVisible = await page.$("#tableDocuments_processing:not([style*='display: none'])");
      if (processingVisible) {
        console.log("Overlay de procesamiento detectado, esperando...");
        await page.waitForSelector("#tableDocuments_processing", { hidden: true, timeout: 30000 });
        console.log("Overlay de procesamiento oculto");
      }
    } catch {
      // Overlay no detectado
    }
    
    // Esperar a que aparezcan filas en la tabla
    try {
      await page.waitForSelector(
        "table#tableDocuments tbody tr, table.dataTable tbody tr, #tableDocuments tbody tr",
        { timeout: 15000 }
      );
      console.log("Filas de tabla detectadas");
    } catch {
      console.log("No se detectaron filas en la tabla después de esperar");
    }
    
    // Esperar un poco más para asegurar que los datos estén completos
    await delay(500);
    
  } catch (err) {
    console.log("Error en waitForTableLoad:", err);
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
