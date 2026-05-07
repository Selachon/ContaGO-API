import puppeteer, { Browser, Page, Cookie } from "puppeteer";
import fs from "fs";
import JSZip from "jszip";
import type { DocumentInfo, ProgressData, DocumentDirection } from "../types/dian.js";

// Estado de progreso compartido entre scraper y rutas de consulta.
export const progressTracker: Map<string, ProgressData> = new Map();

interface ExtractionResult {
  documents: DocumentInfo[];
  cookies: Record<string, string>;
}

interface ListingRecord {
  cufe: string;
  docnum: string;
}

interface ExtractOptions {
  tokenUrl: string;
  startDate?: string;
  endDate?: string;
  progressUid?: string;
  /** Tipo de documentos: "received" (recibidos) o "sent" (emitidos). Default: "received" */
  documentDirection?: DocumentDirection;
}

const BROWSER_LAUNCH_TIMEOUT_MS = Number(process.env.PUPPETEER_LAUNCH_TIMEOUT_MS || 120000);
const BROWSER_LAUNCH_RETRIES = Number(process.env.PUPPETEER_LAUNCH_RETRIES || 3);

/**
 * Extrae ids de documentos DIAN y cookies de sesion para descargas posteriores.
 */
export async function extractDocumentIds(
  tokenUrl: string,
  startDate: string | undefined,
  endDate: string | undefined,
  progressUid?: string,
  documentDirection: DocumentDirection = "received"
): Promise<ExtractionResult> {
  const direction = documentDirection || "received";
  const isSent = direction === "sent";
  const directionLabel = isSent ? "emitidos" : "recibidos";
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

    browser = await launchBrowserWithRetry(executablePath, updateProgress);

    const page = await browser.newPage();
    await page.setViewport({ width: 1280, height: 800 });

    // Timeout alto para ambientes con latencia variable (2 minutos).
    page.setDefaultTimeout(120000);
    page.setDefaultNavigationTimeout(120000);

    // 1) Acceso inicial con token_url (con reintentos).
    updateProgress({ step: "Accediendo con token..." });
    await navigateWithRetry(page, tokenUrl, 3);
    await delay(1000);

    // Verificar si el token expiró (redirige a página de login)
    const currentUrl = page.url();
    if (isLoginPage(currentUrl)) {
      throw new Error("TOKEN_EXPIRED: El token ha expirado. Por favor, genera un nuevo token desde el portal DIAN.");
    }

    // 2) Navegar al listado de documentos (con reintentos).
    const documentUrl = isSent 
      ? "https://catalogo-vpfe.dian.gov.co/Document/Sent"
      : "https://catalogo-vpfe.dian.gov.co/Document/Received";
    updateProgress({ step: `Navegando a documentos ${directionLabel}...` });
    await navigateWithRetry(page, documentUrl, 3);
    await delay(600);

    // Verificar nuevamente si redirigió a login (sesión inválida)
    const urlAfterNav = page.url();
    if (isLoginPage(urlAfterNav)) {
      throw new Error("TOKEN_EXPIRED: La sesión ha expirado. Por favor, genera un nuevo token desde el portal DIAN.");
    }

    updateProgress({ step: "Extrayendo lista (iniciando)...", current: 0, total: 0 });

    // 3) Aplicar filtro solo cuando hay rango completo.
    if (startDate && endDate) {
      updateProgress({ step: "Aplicando rango de fechas..." });
      await applyDateFilter(page, startDate, endDate);
    }

    // 4) Esperar render inicial de la tabla.
    updateProgress({ step: "Cargando resultados..." });
    await waitForTableLoad(page);

    // 5) Forzar mayor tamaño de página para reducir paginación y omisiones.
    updateProgress({ step: "Ajustando paginación..." });
    await setPageLength(page, 100);
    await waitForFullTableLoad(page, 100);

    // 6) Extraer con paginacion deduplicando por trackId.
    const allDocuments: DocumentInfo[] = [];
    const seenIds = new Set<string>();
    let pageIndex = 0;

    // Total reportado por DataTables para estimar progreso.
    const expectedTotal = await page.evaluate(() => {
      const info = document.querySelector("#tableDocuments_info, .dataTables_info, .dt-info");
      const text = info?.textContent || "";

      // Extrae todos los numeros del texto (ej: "Mostrando registros del 1 al 50 de 8000").
      const nums = Array.from(text.matchAll(/\d[\d.,]*/g)).map((m) =>
        parseInt((m[0] || "").replace(/[.,]/g, ""), 10)
      ).filter((n) => Number.isFinite(n));

      if (nums.length >= 3) {
        return nums[nums.length - 1] || 0;
      }

      return 0;
    });
    console.log(`Total esperado según paginación: ${expectedTotal}`);

    while (true) {
      pageIndex++;
      
      // Mensaje descriptivo que muestra cantidad encontrada
      const progressMsg = allDocuments.length > 0
        ? `Extrayendo documentos... ${allDocuments.length} encontrados`
        : `Extrayendo documentos (página ${pageIndex})...`;
      
      updateProgress({
        step: progressMsg,
        current: allDocuments.length,
        total: expectedTotal || Math.max(allDocuments.length + 50, 100),
      });

      await delay(500);

      // Diagnostico rapido de filas visibles en esta pagina.
      const visibleRows = await page.evaluate(() => {
        return document.querySelectorAll("#tableDocuments tbody tr:not(.dataTables_empty)").length;
      });
      console.log(`Página ${pageIndex} - filas visibles: ${visibleRows}`);

      // Evita terminar prematuramente en estados intermedios de recarga.
      if (visibleRows === 0) {
        await waitForFullTableLoad(page, 100);
        const rowsAfterWait = await page.evaluate(() => {
          return document.querySelectorAll("#tableDocuments tbody tr:not(.dataTables_empty)").length;
        });

        if (rowsAfterWait === 0) {
          const hasNextWhenEmpty = await goToNextPage(page);
          if (!hasNextWhenEmpty) {
            console.log(`Página ${pageIndex} - tabla vacía sin más páginas, terminando`);
            break;
          }

          await waitForTableChange(page, seenIds);
          await waitForFullTableLoad(page, 100);
          continue;
        }
      }

      // Extrae ids desde selectores alternos por variaciones de DIAN.
      const newDocs = await extractDocsFromPage(page, seenIds, isSent);
      allDocuments.push(...newDocs);

      // Actualizar progreso inmediatamente después de extraer
      const updatedMsg = `Extrayendo documentos... ${allDocuments.length} encontrados`;
      updateProgress({
        step: updatedMsg,
        current: allDocuments.length,
        total: expectedTotal || allDocuments.length,
      });

      console.log(`Página ${pageIndex} - extraídos: ${newDocs.length}, acumulados: ${allDocuments.length}/${expectedTotal}`);

      // Finaliza si se alcanzo el total esperado.
      if (expectedTotal > 0 && allDocuments.length >= expectedTotal) {
        console.log(`Alcanzado el total esperado (${expectedTotal}), terminando`);
        break;
      }

      // Avanza de pagina si hay boton Next habilitado.
      const hasNext = await goToNextPage(page);
      if (!hasNext) {
        console.log(`No hay más páginas disponibles, terminando`);
        break;
      }

      // Espera cambio real de pagina antes de continuar.
      await waitForTableChange(page, seenIds);
      await waitForFullTableLoad(page, 100);
    }

    // 7) Reconciliar con listado maestro por CUFE/numero de documento.
    const listedRecords = await extractListingRecordsFromDownloadTab(page, direction, startDate, endDate);
    if (listedRecords.length > 0) {
      const byId = new Map(allDocuments.map((d) => [d.id, d]));
      const byDocNum = new Map<string, DocumentInfo>();
      for (const d of allDocuments) {
        const key = (d.docnum || "").trim();
        if (key && !byDocNum.has(key)) byDocNum.set(key, d);
      }

      const missing = listedRecords.filter((r) => {
        const doc = byDocNum.get(r.docnum);
        return !doc;
      });

      if (missing.length > 0) {
        console.log(`Reconciliación DIAN: faltan ${missing.length} registros del listado maestro. Buscando por CUFE/docnum...`);
        updateProgress({
          step: `Reconciliando faltantes por CUFE (${missing.length})...`,
          current: allDocuments.length,
          total: Math.max(expectedTotal || allDocuments.length + missing.length, allDocuments.length + missing.length),
        });

        let rescued = 0;
        for (const rec of missing) {
          const found = await findDocumentByUniqueCodeOrDocnum(page, rec.cufe, rec.docnum, seenIds, isSent);
          if (found) {
            if (!byId.has(found.id)) {
              byId.set(found.id, found);
              allDocuments.push(found);
              rescued++;
            }
            if (!byDocNum.has(found.docnum)) {
              byDocNum.set(found.docnum, found);
            }
          }
        }

        console.log(`Reconciliación DIAN: recuperados ${rescued}/${missing.length} faltantes.`);

        const unresolved = missing.filter((r) => !byDocNum.has(r.docnum));
        if (unresolved.length > 0) {
          const sample = unresolved.slice(0, 5).map((r) => r.docnum).join(", ");
          const msg = `INCOMPLETE_EXTRACTION: Faltan ${unresolved.length} documentos del listado DIAN después de reconciliación por CUFE. Ejemplos: ${sample}`;
          updateProgress({
            step: "Error de reconciliación",
            detalle: msg,
          });
          throw new Error(msg);
        }
      }
    }

    // Reutiliza cookies del navegador para las descargas HTTP.
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

async function launchBrowserWithRetry(
  executablePath: string | null,
  updateProgress: (data: Partial<ProgressData>) => void
): Promise<Browser> {
  let lastError: Error | null = null;

  for (let attempt = 1; attempt <= BROWSER_LAUNCH_RETRIES; attempt++) {
    try {
      return await puppeteer.launch({
        headless: true,
        timeout: BROWSER_LAUNCH_TIMEOUT_MS,
        protocolTimeout: BROWSER_LAUNCH_TIMEOUT_MS,
        args: [
          "--no-sandbox",
          "--disable-setuid-sandbox",
          "--disable-dev-shm-usage",
          "--disable-gpu",
          "--disable-extensions",
          "--disable-background-networking",
          "--disable-sync",
          "--no-first-run",
        ],
        executablePath: executablePath || undefined,
      });
    } catch (err) {
      lastError = err as Error;
      const isLastAttempt = attempt >= BROWSER_LAUNCH_RETRIES;
      const message = lastError.message || String(lastError);

      console.warn(`Fallo iniciando navegador (intento ${attempt}/${BROWSER_LAUNCH_RETRIES}): ${message}`);

      if (isLastAttempt) {
        break;
      }

      updateProgress({
        step: `Reintentando inicio de navegador (${attempt}/${BROWSER_LAUNCH_RETRIES})...`,
      });

      await delay(1500 * attempt);
    }
  }

  throw new Error(
    `No fue posible iniciar Chromium tras ${BROWSER_LAUNCH_RETRIES} intentos. Último error: ${lastError?.message || "desconocido"}`
  );
}

function resolveExecutablePath(): string | null {
  if (!process.env.PUPPETEER_CACHE_DIR) {
    process.env.PUPPETEER_CACHE_DIR = `${process.cwd()}/.cache/puppeteer`;
  }

  const envPath = process.env.PUPPETEER_EXECUTABLE_PATH;
  if (envPath && !fs.existsSync(envPath)) {
    // Evita que Puppeteer falle por un binario inexistente en env.
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
      // Fija un ejecutable valido para todos los launches siguientes.
      process.env.PUPPETEER_EXECUTABLE_PATH = candidate;
      return candidate;
    }
  }

  // Si no hay binario del sistema, usar el gestionado por Puppeteer.
  delete process.env.PUPPETEER_EXECUTABLE_PATH;

  return null;
}

async function applyDateFilter(page: Page, startDate: string, endDate: string): Promise<void> {
  try {
    // Seleccion flexible para cambios menores en el DOM de DIAN.
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

    // Click en Buscar con manejo de navegacion que puede invalidar el contexto.
    try {
      const clickPromise = page.evaluate(() => {
        const buttons = Array.from(document.querySelectorAll("button"));
        const btn = buttons.find((b) => b.textContent?.trim().includes("Buscar"));
        if (btn) {
          (btn as HTMLElement).click();
          return true;
        }
        return false;
      });
      
      // Timeout corto para no bloquear si no existe boton compatible.
      const clicked = await Promise.race([
        clickPromise,
        delay(2000).then(() => false)
      ]);
      
      if (clicked) {
        // La accion puede disparar recarga parcial o completa.
        await delay(500);
        await waitForTableLoad(page);
      }
    } catch (navErr: unknown) {
      // "context was destroyed" es esperado cuando hay navegacion.
      const errMsg = navErr instanceof Error ? navErr.message : String(navErr);
      if (errMsg.includes("context was destroyed") || errMsg.includes("navigation")) {
        console.log("Navegación detectada después de aplicar filtro de fechas - esperando recarga...");
        await delay(1000);
        await waitForTableLoad(page);
      } else {
        throw navErr;
      }
    }
  } catch (err) {
    console.error("Error aplicando rango de fechas:", err);
    // El filtro es opcional; se prioriza retornar datos disponibles.
  }
}

async function waitForTableLoad(page: Page): Promise<void> {
  try {
    // DataTables usa este overlay durante recargas (v1: #..._processing, v2: .dt-processing).
    try {
      await page.waitForSelector("#tableDocuments_processing, .dt-processing", { visible: true, timeout: 5000 });
      await page.waitForSelector("#tableDocuments_processing, .dt-processing", { hidden: true, timeout: 20000 });
    } catch {
      // Fallback cuando el overlay no se renderiza.
      await page.waitForSelector("table#tableDocuments tbody tr", { timeout: 6000 }).catch(() => {});
    }
  } catch (err) {
    console.log("Espera resultados:", err);
  }
}

/**
 * Espera a que la tabla alcance las filas esperadas segun paginacion.
 * DIAN puede entregar respuestas AJAX parciales; este guard reduce faltantes.
 */
async function waitForFullTableLoad(page: Page, pageLength: number): Promise<void> {
  const maxWait = 15000;
  const startTime = Date.now();
  
  while (Date.now() - startTime < maxWait) {
    const { expectedRows, actualRows } = await page.evaluate(() => {
      const info = document.querySelector("#tableDocuments_info, .dataTables_info, .dt-info");
      const text = info?.textContent || "";
      
      // Ejemplo esperado: "Mostrando del 1 al 50 de 1.172 registros".
      const match = text.match(/del\s+([\d.,]+)\s+al\s+([\d.,]+)/i);
      
      let expected = 0;
      if (match) {
        const from = parseInt(match[1].replace(/[.,]/g, ""), 10);
        const to = parseInt(match[2].replace(/[.,]/g, ""), 10);
        expected = to - from + 1;
      }
      
      const rows = document.querySelectorAll("#tableDocuments tbody tr:not(.dataTables_empty)");
      return { expectedRows: expected, actualRows: rows.length };
    });
    
    if (expectedRows > 0 && actualRows >= expectedRows) {
      console.log(`Tabla completamente cargada: ${actualRows}/${expectedRows} filas`);
      return;
    }
    
    // Fallback: aceptar pagina llena aunque no haya texto parseable.
    if (pageLength > 0 && actualRows >= pageLength) {
      console.log(`Tabla cargada con ${actualRows} filas (máximo por página)`);
      return;
    }
    
    console.log(`Esperando filas: ${actualRows}/${expectedRows || (pageLength > 0 ? pageLength : "?")}`);
    await delay(500);
  }
  
  console.log("Timeout esperando carga completa de tabla, continuando...");
}

async function setPageLength(page: Page, length: number): Promise<void> {
  try {
    // Primero intenta el select nativo de DataTables.
    const selectHandle = await page.$(
      "select[name='tableDocuments_length'], " +
      "#tableDocuments_length select, " +
      "select[name*='length'], " +
      ".dt-length select"
    );

    if (selectHandle) {
      await selectHandle.select(length.toString());
      await delay(500);
    } else {
      // Fallback cuando el select no esta visible o cambia de id.
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
    // Errores de contexto destruido son comunes cuando la página recarga
    const errMsg = err instanceof Error ? err.message : String(err);
    if (errMsg.includes("context") || errMsg.includes("Context") || errMsg.includes("detached")) {
      console.log("Contexto cambiado durante setPageLength, esperando estabilización...");
      await delay(2000);
    } else {
      console.warn("Error cambiando longitud de página:", errMsg.substring(0, 100));
    }
  }
}

// Tipos de documento que deben ser ignorados siempre (no son facturas reales)
const IGNORED_DOC_TYPES = [
  "application response",
  "applicationresponse",
  "app response",
  "respuesta de aplicación",
  "respuesta aplicación",
];

// Tipos de documento adicionales que deben ser ignorados solo para documentos emitidos
const IGNORED_DOC_TYPES_SENT_ONLY = [
  "nomina individual",
  "nómina individual",
];

function shouldIgnoreDocType(docType: string, isSentDocuments: boolean = false): boolean {
  const normalized = docType.toLowerCase().trim();
  
  // Siempre ignorar estos tipos
  if (IGNORED_DOC_TYPES.some(ignored => normalized.includes(ignored))) {
    return true;
  }
  
  // Para documentos emitidos, también ignorar nóminas
  if (isSentDocuments && IGNORED_DOC_TYPES_SENT_ONLY.some(ignored => normalized.includes(ignored))) {
    return true;
  }
  
  return false;
}

async function extractDocsFromPage(page: Page, seenIds: Set<string>, isSentDocuments: boolean = false): Promise<DocumentInfo[]> {
  const docs: DocumentInfo[] = [];

  const items = await page.evaluate((isSentDocumentsInPage) => {
    const results: Array<{
      id: string;
      docnum: string;
      nit: string;
      cufe: string;
      docType: string;
      documentTypeId?: string;
      fechaValidacion?: string;
      fechaGeneracion?: string;
    }> = [];
    
    // Excluye la fila placeholder de DataTables.
    const rows = document.querySelectorAll("#tableDocuments tbody tr:not(.dataTables_empty)");
    
    for (const row of rows) {
      let trackId: string | null = null;
      let documentTypeId: string | undefined;
      let fechaValidacion: string | undefined;
      let fechaGeneracion: string | undefined;
      
      // Metodo 1: botones/enlaces de descarga en la fila.
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
        
        // Extraer atributos adicionales para documentos equivalentes POS
        if (el.classList.contains("download-equivalente-document")) {
          documentTypeId = el.getAttribute("documentypeid") || el.getAttribute("documenttypeid") || undefined;
          // emissiondate es la fecha de validación, generationdate es la fecha de generación
          fechaValidacion = el.getAttribute("emissiondate") || undefined;
          fechaGeneracion = el.getAttribute("generationdate") || undefined;
        }
        
        if (!trackId) {
          const href = el.getAttribute("href") || "";
          const match = href.match(/trackId=([A-Za-z0-9-]+)/i);
          if (match) trackId = match[1];
        }
        
        if (trackId) break;
      }
      
      // Metodo 2: atributos de tracking en otros elementos de la fila.
      if (!trackId) {
        const anyWithTrack = row.querySelector("[data-trackid], [data-id], [data-track-id]");
        if (anyWithTrack) {
          trackId = anyWithTrack.getAttribute("data-trackid") || 
                    anyWithTrack.getAttribute("data-id") ||
                    anyWithTrack.getAttribute("data-track-id");
        }
      }
      
      // Metodo 3: atributos en el tr.
      if (!trackId) {
        trackId = row.getAttribute("data-trackid") || 
                  row.getAttribute("data-id") ||
                  row.id;
      }
      
      // Metodo 4: parseo de trackId desde hrefs internos.
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
      
      // Columnas de la tabla DIAN (verificado):
      // 0: checkbox, 1: Recepción, 2: Fecha, 3: Prefijo, 4: Nº documento,
      // 5: Tipo, 6: NIT Emisor, 7: Emisor, 8: NIT Receptor, 9: Receptor, etc.
      const tds = row.querySelectorAll("td");
      const docnum = tds[4]?.textContent?.trim() || "";
      const docType = tds[5]?.textContent?.trim() || "";
      const preferredNitIndex = isSentDocumentsInPage ? 8 : 6;
      const fallbackNitIndex = isSentDocumentsInPage ? 6 : 8;
      const nit =
        tds[preferredNitIndex]?.textContent?.trim() ||
        tds[fallbackNitIndex]?.textContent?.trim() ||
        "";

      const cufe = (tds[13]?.textContent?.trim() || tds[12]?.textContent?.trim() || "").toLowerCase();

      results.push({ id: trackId, docnum, nit, cufe, docType, documentTypeId, fechaValidacion, fechaGeneracion });
    }

    return results;
  }, isSentDocuments);

  let skippedCount = 0;
  
  for (const item of items) {
    if (!seenIds.has(item.id)) {
      // Filtrar Application Response y otros tipos no relevantes
      if (shouldIgnoreDocType(item.docType, isSentDocuments)) {
        skippedCount++;
        continue;
      }
      
      seenIds.add(item.id);
      docs.push({
        id: item.id,
        docnum: item.docnum,
        nit: item.nit,
        cufe: item.cufe,
        docType: item.docType,
        documentTypeId: item.documentTypeId,
        fechaValidacion: item.fechaValidacion,
        fechaGeneracion: item.fechaGeneracion,
      });
    }
  }
  
  if (skippedCount > 0) {
    console.log(`  -> Omitidos ${skippedCount} documentos tipo "Application Response"`);
  }

  return docs;
}

async function extractListingRecordsFromDownloadTab(
  page: Page,
  direction: DocumentDirection,
  startDate?: string,
  endDate?: string
): Promise<ListingRecord[]> {
  try {
    const baseUrl = "https://catalogo-vpfe.dian.gov.co";
    await navigateWithRetry(page, `${baseUrl}/Document/Export`, 2);
    await delay(800);

    const formData = await page.evaluate(() => {
      const token = (document.querySelector("input[name='__RequestVerificationToken']") as HTMLInputElement | null)?.value || "";
      const type = (document.querySelector("input[name='Type']") as HTMLInputElement | null)?.value || "0";
      const amountAdmin = (document.querySelector("input[name='AmountAdmin']") as HTMLInputElement | null)?.value || "100000";
      return { token, type, amountAdmin };
    });

    if (!formData.token) return [];

    const cookieHeader = (await page.cookies()).map((c) => `${c.name}=${c.value}`).join("; ");
    const body = new URLSearchParams();
    body.set("__RequestVerificationToken", formData.token);
    body.set("Type", formData.type || "0");
    body.set("AmountAdmin", formData.amountAdmin || "100000");
    body.set("ReceiverCode", "");
    body.set("GroupCode", direction === "sent" ? "1" : "2");
    if (startDate) body.set("StartDate", toDianExportDate(startDate));
    if (endDate) body.set("EndDate", toDianExportDate(endDate, true));

    await fetch(`${baseUrl}/Document/Export`, {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        Cookie: cookieHeader,
        Referer: `${baseUrl}/Document/Export`,
      },
      body: body.toString(),
    });

    await delay(1500);
    const exportHtml = await fetch(`${baseUrl}/Document/Export`, {
      method: "GET",
      headers: {
        Cookie: cookieHeader,
        Referer: `${baseUrl}/Document/Export`,
      },
    }).then((r) => r.text());

    const links = Array.from(exportHtml.matchAll(/href="(\/Document\/DownloadExportedZipFile\?pk=[^"]+)"/g))
      .map((m) => m[1].replace(/&amp;/g, "&"));
    if (links.length === 0) return [];

    const latestLink = links[0];
    const downloaded = await fetch(`${baseUrl}${latestLink}`, {
      method: "GET",
      headers: { Cookie: cookieHeader, Referer: `${baseUrl}/Document/Export` },
    }).then((r) => r.arrayBuffer());
    const zipBuffer = Buffer.from(new Uint8Array(downloaded));

    return await parseListingRecordsFromExportZip(zipBuffer, direction);
  } catch {
    return [];
  }
}

function toDianExportDate(dateISO: string, endOfDay: boolean = false): string {
  const [year, month, day] = dateISO.split("-");
  if (!year || !month || !day) return dateISO;
  return endOfDay
    ? `${Number(month)}/${Number(day)}/${year} 11:59:59 PM`
    : `${Number(month)}/${Number(day)}/${year} 12:00:00 AM`;
}

async function parseListingRecordsFromExportZip(zipBuffer: Buffer, direction: DocumentDirection): Promise<ListingRecord[]> {
  const zip = await JSZip.loadAsync(zipBuffer);
  const xlsxName = Object.keys(zip.files).find((n) => n.toLowerCase().endsWith(".xlsx") && !zip.files[n].dir);
  if (!xlsxName) return [];

  const xlsxBuffer = await zip.files[xlsxName].async("nodebuffer");
  const xlsx = await JSZip.loadAsync(xlsxBuffer);

  const sharedStringsXml = await xlsx.file("xl/sharedStrings.xml")?.async("string");
  const sheetXml = await xlsx.file("xl/worksheets/sheet1.xml")?.async("string");
  if (!sheetXml) return [];

  const sharedStrings = parseSharedStrings(sharedStringsXml || "");
  const rows = parseSheetRows(sheetXml, sharedStrings);
  if (rows.length < 2) return [];

  const headers = rows[0].map((h) => h.trim().toLowerCase());
  const cufeIdx = headers.findIndex((h) => h.includes("cufe") || h.includes("cude") || h.includes("código único") || h.includes("codigo unico"));
  const folioIdx = headers.findIndex((h) => h === "folio" || h.includes("número") || h.includes("numero"));
  const groupIdx = headers.findIndex((h) => h === "grupo");

  if (cufeIdx < 0 || folioIdx < 0) return [];

  const expectedGroup = direction === "sent" ? "emitido" : "recibido";
  const out: ListingRecord[] = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const group = (groupIdx >= 0 ? row[groupIdx] : "").toLowerCase();
    if (group && !group.includes(expectedGroup)) continue;

    const cufe = (row[cufeIdx] || "").trim().toLowerCase();
    const docnum = (row[folioIdx] || "").trim();
    if (!docnum) continue;
    out.push({ cufe, docnum });
  }

  const dedup = new Map<string, ListingRecord>();
  for (const rec of out) {
    const key = `${rec.docnum}::${rec.cufe}`;
    if (!dedup.has(key)) dedup.set(key, rec);
  }
  return Array.from(dedup.values());
}

function parseSharedStrings(xml: string): string[] {
  if (!xml) return [];
  const strings: string[] = [];
  const siMatches = xml.match(/<si[\s\S]*?<\/si>/g) || [];
  for (const si of siMatches) {
    const textParts = Array.from(si.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)).map((m) => decodeXml(m[1]));
    strings.push(textParts.join(""));
  }
  return strings;
}

function parseSheetRows(xml: string, sharedStrings: string[]): string[][] {
  const rows: string[][] = [];
  const rowMatches = xml.match(/<row[^>]*>[\s\S]*?<\/row>/g) || [];

  for (const rowXml of rowMatches) {
    const cells = Array.from(rowXml.matchAll(/<c([^>]*)>([\s\S]*?)<\/c>/g));
    const rowValues: string[] = [];

    for (const cellMatch of cells) {
      const attrs = cellMatch[1];
      const body = cellMatch[2];
      const ref = (attrs.match(/r="([A-Z]+)\d+"/) || [])[1] || "A";
      const colIndex = colLettersToIndex(ref);
      while (rowValues.length < colIndex) rowValues.push("");

      const type = (attrs.match(/t="([^"]+)"/) || [])[1] || "";
      const valueMatch = body.match(/<v>([\s\S]*?)<\/v>/);
      const inlineMatch = body.match(/<t[^>]*>([\s\S]*?)<\/t>/);

      let value = "";
      if (type === "s" && valueMatch) {
        const idx = Number(valueMatch[1]);
        value = Number.isFinite(idx) && sharedStrings[idx] !== undefined ? sharedStrings[idx] : "";
      } else if (inlineMatch) {
        value = decodeXml(inlineMatch[1]);
      } else if (valueMatch) {
        value = decodeXml(valueMatch[1]);
      }

      rowValues[colIndex - 1] = value.trim();
    }

    if (rowValues.some((v) => v !== "")) rows.push(rowValues);
  }

  return rows;
}

function colLettersToIndex(letters: string): number {
  let index = 0;
  for (const ch of letters) {
    index = index * 26 + (ch.charCodeAt(0) - 64);
  }
  return index;
}

function decodeXml(value: string): string {
  return value
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

async function findDocumentByUniqueCodeOrDocnum(
  page: Page,
  cufe: string,
  docnum: string,
  seenIds: Set<string>,
  isSentDocuments: boolean
): Promise<DocumentInfo | null> {
  const candidates = [cufe, docnum].filter(Boolean);
  for (const term of candidates) {
    try {
      await page.evaluate((value) => {
        const input = document.querySelector(
          "#tableDocuments_filter input, .dataTables_filter input, .dt-search input, input[type='search']"
        ) as HTMLInputElement | null;
        if (!input) return;
        input.value = value;
        input.dispatchEvent(new Event("input", { bubbles: true }));
        input.dispatchEvent(new Event("keyup", { bubbles: true }));
        input.dispatchEvent(new Event("change", { bubbles: true }));
      }, term);

      await delay(900);
      await waitForTableLoad(page);

      const found = await extractDocsFromPage(page, seenIds, isSentDocuments);
      if (found.length > 0) {
        const exact = found.find((d) => (docnum && d.docnum === docnum) || (cufe && d.cufe === cufe));
        return exact || found[0];
      }
    } catch {
      // continuar con siguiente candidato
    }
  }

  return null;
}

async function goToNextPage(page: Page): Promise<boolean> {
  try {
    const nextBtn = await page.$(
      "#tableDocuments_next, .paginate_button.next, a.next, .dt-paging-button.next"
    );

    if (!nextBtn) return false;

    const isDisabled = await page.evaluate((el) => {
      const selfDisabled = el?.classList.contains("disabled") ||
        el?.getAttribute("aria-disabled") === "true" ||
        (el as HTMLButtonElement | null)?.disabled === true;
      const parentDisabled = el?.parentElement?.classList.contains("disabled") || false;
      return Boolean(selfDisabled || parentDisabled);
    }, nextBtn);

    if (isDisabled) return false;

    // Permite confirmar que realmente avanzo de pagina.
    const currentPageNum = await page.evaluate(() => {
      const active = document.querySelector(
        "#tableDocuments_paginate .paginate_button.current, .paginate_button.active, .dt-paging-button.current"
      );
      return active?.textContent?.trim() || "0";
    });

    await nextBtn.click();
    
    // Espera activa hasta detectar nuevo numero de pagina.
    const startTime = Date.now();
    while (Date.now() - startTime < 10000) {
      const newPageNum = await page.evaluate(() => {
        const active = document.querySelector(
          "#tableDocuments_paginate .paginate_button.current, .paginate_button.active, .dt-paging-button.current"
        );
        return active?.textContent?.trim() || "0";
      });
      
      if (newPageNum !== currentPageNum) {
        break;
      }
      await delay(200);
    }
    
    // Espera fin de recarga si el indicador aparece.
    try {
      await page.waitForSelector("#tableDocuments_processing, .dt-processing", { hidden: true, timeout: 15000 });
    } catch {
      // Si no aparece, seguir con espera de estabilizacion.
    }
    
    // Margen corto para evitar lecturas sobre DOM intermedio.
    await delay(800);
    
    return true;
  } catch {
    return false;
  }
}

async function waitForTableChange(page: Page, seenIds: Set<string>): Promise<void> {
  // Espera fin de recarga cuando DataTables lo reporta.
  try {
    await page.waitForSelector("#tableDocuments_processing, .dt-processing", { hidden: true, timeout: 15000 });
  } catch {
    // Si no hay indicador, continuar con verificacion por ids.
  }
  
  // Margen corto antes de inspeccionar filas.
  await delay(800);
  
  // Verifica que los ids visibles no sean la misma pagina previa.
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
      // Margen adicional para reducir lecturas parciales.
      await delay(500);
      return;
    }
    await delay(300);
  }

  // Ultimo margen de seguridad antes de continuar.
  await delay(500);
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * Verifica si la URL actual es la página de login de DIAN.
 * Esto indica que el token expiró o la sesión es inválida.
 */
function isLoginPage(url: string): boolean {
  const loginIndicators = [
    "/Account/Login",
    "/User/Login", 
    "/auth/login",
    "login.microsoftonline.com",
    "login.live.com",
  ];
  
  const lowerUrl = url.toLowerCase();
  return loginIndicators.some(indicator => lowerUrl.includes(indicator.toLowerCase()));
}

/**
 * Navega a una URL con reintentos y diferentes estrategias de espera.
 * Estrategia progresiva: domcontentloaded → load → networkidle2
 */
async function navigateWithRetry(page: Page, url: string, maxRetries: number = 3): Promise<void> {
  let lastError: Error | null = null;
  
  // Estrategias de espera progresivas (de más rápido a más robusto)
  const strategies: Array<{ waitUntil: "domcontentloaded" | "load" | "networkidle2"; timeout: number }> = [
    { waitUntil: "domcontentloaded", timeout: 60000 },
    { waitUntil: "load", timeout: 90000 },
    { waitUntil: "networkidle2", timeout: 120000 },
  ];
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    const strategy = strategies[Math.min(attempt - 1, strategies.length - 1)];
    
    try {
      console.log(`Navegación intento ${attempt}/${maxRetries} (${strategy.waitUntil}, ${strategy.timeout/1000}s): ${url.substring(0, 60)}...`);
      
      await page.goto(url, { 
        waitUntil: strategy.waitUntil,
        timeout: strategy.timeout,
      });
      
      // Espera adicional para que scripts asíncronos carguen
      await delay(1500);
      
      console.log(`Navegación exitosa en intento ${attempt}`);
      return;
      
    } catch (err) {
      lastError = err instanceof Error ? err : new Error(String(err));
      const isTimeout = lastError.message.includes("timeout") || lastError.message.includes("Timeout");
      const isContextDestroyed = lastError.message.includes("context") || lastError.message.includes("Context");
      
      console.log(`Navegación falló intento ${attempt}: ${lastError.message.substring(0, 100)}`);
      
      // Si es error de contexto destruido, la página puede haber cargado parcialmente
      if (isContextDestroyed) {
        console.log("Contexto destruido - verificando si la página cargó...");
        await delay(2000);
        
        // Intentar verificar si estamos en una página válida
        try {
          const currentUrl = page.url();
          if (currentUrl && !currentUrl.includes("about:blank")) {
            console.log(`Página cargada parcialmente: ${currentUrl.substring(0, 60)}`);
            return; // Considerar como éxito parcial
          }
        } catch {
          // Ignorar errores de verificación
        }
      }
      
      // Si es timeout, esperar más antes de reintentar
      if (isTimeout) {
        console.log(`Timeout detectado, esperando ${3 + attempt}s antes de reintentar...`);
        await delay(3000 + attempt * 1000);
      } else {
        await delay(2000);
      }
    }
  }
  
  // Si todos los intentos fallaron, lanzar error descriptivo
  const errorMsg = lastError?.message || "Error desconocido";
  if (errorMsg.includes("timeout") || errorMsg.includes("Timeout")) {
    throw new Error(`NAVIGATION_TIMEOUT: No se pudo conectar con DIAN después de ${maxRetries} intentos. El servidor puede estar lento o no disponible. Por favor, intenta de nuevo en unos minutos.`);
  }
  
  throw lastError || new Error(`Navegación falló después de ${maxRetries} intentos`);
}
