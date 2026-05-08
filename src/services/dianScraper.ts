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

type OnDocumentFound = (ctx: {
  doc: DocumentInfo;
  index: number;
  total: number;
  cookies: Record<string, string>;
}) => Promise<void> | void;

interface PrecheckResult {
  ok: true;
  listedCount: number;
}

interface ListingRecord {
  cufe: string;
  docnum: string;
}

function normalizeCufe(value: string): string {
  return (value || "").replace(/[^A-Fa-f0-9]/g, "").trim();
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
    await hardenPageRuntime(page);
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

/**
 * Nuevo flujo principal: obtiene CUFEs desde Descarga de listados y luego
 * consulta documento por documento en la bandeja (recibidos/emitidos).
 * Mantiene precisión alta y evita pérdidas por paginación larga.
 */
export async function extractDocumentIdsByCufe(
  tokenUrl: string,
  startDate: string | undefined,
  endDate: string | undefined,
  progressUid?: string,
  documentDirection: DocumentDirection = "received",
  onProgress?: (data: Partial<ProgressData>) => void,
  onDocumentFound?: OnDocumentFound
): Promise<ExtractionResult> {
  const direction = documentDirection || "received";
  const isSent = direction === "sent";
  const directionLabel = isSent ? "emitidos" : "recibidos";
  const updateProgress = (data: Partial<ProgressData>) => {
    onProgress?.(data);
    if (progressUid) {
      const current = progressTracker.get(progressUid) || { step: "", current: 0, total: 0 };
      progressTracker.set(progressUid, { ...current, ...data });
    }
  };

  let browser: Browser | null = null;

  try {
    updateProgress({ step: "Iniciando navegador...", current: 0, total: 0 });
    const executablePath = resolveExecutablePath();
    browser = await launchBrowserWithRetry(executablePath, updateProgress);

    const page = await browser.newPage();
    await hardenPageRuntime(page);
    await page.setViewport({ width: 1280, height: 800 });
    page.setDefaultTimeout(120000);
    page.setDefaultNavigationTimeout(120000);

    updateProgress({ step: "Accediendo con token...", current: 0, total: 1 });
    await navigateWithRetry(page, tokenUrl, 3);
    await delay(1000);

    if (isLoginPage(page.url())) {
      throw new Error("TOKEN_EXPIRED: El token ha expirado. Por favor, genera un nuevo token desde el portal DIAN.");
    }

    updateProgress({ step: "Descargando listado DIAN por CUFE...", current: 0, total: 1 });
    const listedRecords = await extractListingRecordsFromDownloadTab(page, direction, startDate, endDate);
    const cufes = Array.from(new Set(listedRecords.map((r) => normalizeCufe(r.cufe || "")).filter(Boolean)));

    if (cufes.length === 0) {
      throw new Error("No se encontraron CUFEs válidos en Descarga de listados para ese rango.");
    }

    updateProgress({
      step: `Listado listo: ${cufes.length} CUFEs para procesar`,
      current: 0,
      total: cufes.length,
    });

    const documentUrl = isSent
      ? "https://catalogo-vpfe.dian.gov.co/Document/Sent"
      : "https://catalogo-vpfe.dian.gov.co/Document/Received";

    const requestedWorkers = Number(process.env.DIAN_CUFE_WORKERS || 4);
    const workerCount = Math.max(1, Math.min(Number.isFinite(requestedWorkers) ? requestedWorkers : 2, 4));
    updateProgress({ step: `Navegando a documentos ${directionLabel}...`, current: 0, total: cufes.length });

    // Página base (se mantiene para compatibilidad y fallback de cookies).
    await navigateWithRetry(page, documentUrl, 3);
    await delay(600);
    if (startDate && endDate) {
      updateProgress({ step: "Aplicando rango de fechas...", current: 0, total: cufes.length });
      await applyDateFilter(page, startDate, endDate, false);
    }
    await waitForTableLoad(page);

    const baseCookies = await page.cookies();
    const baseCookieMap: Record<string, string> = {};
    for (const c of baseCookies) baseCookieMap[c.name] = c.value;

    const documents: DocumentInfo[] = [];
    const seenIds = new Set<string>();
    const acceptedDocIds = new Set<string>();
    const acceptedCufes = new Set<string>();
    let failures = 0;
    let duplicatesSkipped = 0;

    if (!browser) {
      throw new Error("No se pudo inicializar el navegador para validación CUFE.");
    }
    const activeBrowser = browser;

    let processed = 0;
    let nextIndex = 0;
    let lastProgressAt = 0;

    const workerPages: Page[] = [];
    const workerCookies: Record<string, string>[] = [];

    const initWorkerPage = async (): Promise<{ page: Page; cookies: Record<string, string> }> => {
      const wp = await activeBrowser.newPage();
      await hardenPageRuntime(wp);
      await wp.setViewport({ width: 1280, height: 800 });
      wp.setDefaultTimeout(120000);
      wp.setDefaultNavigationTimeout(120000);

      await navigateWithRetry(wp, tokenUrl, 3);
      await delay(250);
      if (isLoginPage(wp.url())) {
        throw new Error("TOKEN_EXPIRED: El token ha expirado. Por favor, genera un nuevo token desde el portal DIAN.");
      }

      await navigateWithRetry(wp, documentUrl, 3);
      await delay(250);
      if (startDate && endDate) {
        // Ajustar rango sin disparar búsqueda aún; cada CUFE ejecuta su propio
        // submit y evita esperas largas redundantes durante inicialización.
        await applyDateFilter(wp, startDate, endDate, false);
      }
      await waitForTableLoad(wp);

      const wc = await wp.cookies();
      const wcMap: Record<string, string> = {};
      for (const c of wc) wcMap[c.name] = c.value;

      return { page: wp, cookies: wcMap };
    };

    for (let w = 0; w < workerCount; w++) {
      const worker = await initWorkerPage();
      workerPages.push(worker.page);
      workerCookies.push(worker.cookies);
    }

    console.log(`[DIAN CUFE] Workers activos: ${workerCount}`);

    const runWorker = async (workerIdx: number) => {
      let wp = workerPages[workerIdx];
      let wc = workerCookies[workerIdx] || baseCookieMap;
      let processedByWorker = 0;
      const recycleEvery = Math.max(50, Number(process.env.DIAN_CUFE_RECYCLE_EVERY || 250));

      while (true) {
        const i = nextIndex;
        if (i >= cufes.length) break;
        nextIndex++;

        const cufe = cufes[i];
        try {
          const maxAttempts = 2;
          let found: DocumentInfo | null = null;
          for (let attempt = 1; attempt <= maxAttempts; attempt++) {
            try {
              found = await findDocumentByUniqueCodeOrDocnum(wp, cufe, "", seenIds, isSent);
              break;
            } catch (err) {
              const msg = err instanceof Error ? err.message : String(err);
              const recoverable =
                msg.includes("detached Frame") ||
                msg.includes("Execution context was destroyed") ||
                msg.includes("Target closed") ||
                msg.includes("Session closed") ||
                msg.includes("Protocol error");

              if (!recoverable || attempt >= maxAttempts) throw err;

              try { await wp.close(); } catch {}
              const refreshed = await initWorkerPage();
              wp = refreshed.page;
              wc = refreshed.cookies;
              workerPages[workerIdx] = wp;
              workerCookies[workerIdx] = wc;
            }
          }

          if (found?.id) {
            found.cufe = cufe;
            const normalizedFoundCufe = normalizeCufe(found.cufe || cufe);

            // Regla estricta: jamás permitir duplicados en el resultado final.
            // Si se repite ID o CUFE, se omite sin descargar nuevamente.
            if (acceptedDocIds.has(found.id) || (normalizedFoundCufe && acceptedCufes.has(normalizedFoundCufe))) {
              duplicatesSkipped++;
            } else {
              acceptedDocIds.add(found.id);
              if (normalizedFoundCufe) acceptedCufes.add(normalizedFoundCufe);
              documents.push(found);
              if (onDocumentFound) {
                await onDocumentFound({
                  doc: found,
                  index: i + 1,
                  total: cufes.length,
                  cookies: wc,
                });
              }
            }
          } else {
            failures++;
            if (failures <= 10) {
              console.warn(`[DIAN CUFE] Sin resultado para CUFE ${cufe.slice(0, 16)}...`);
            }
          }
        } catch (err) {
          failures++;
          if (failures <= 10) {
            console.warn(`[DIAN CUFE] Error consultando CUFE ${cufe.slice(0, 16)}...`, err);
          }
        } finally {
          processed++;
          processedByWorker++;

          const now = Date.now();
          const shouldFlushProgress =
            processed === cufes.length || now - lastProgressAt >= 400 || processed % 10 === 0;
          if (shouldFlushProgress) {
            lastProgressAt = now;
            updateProgress({
              step: `Validando CUFEs (${processed}/${cufes.length})...`,
              current: processed,
              total: cufes.length,
            });
          }

          if (processedByWorker > 0 && processedByWorker % recycleEvery === 0) {
            try { await wp.close(); } catch {}
            const refreshed = await initWorkerPage();
            wp = refreshed.page;
            wc = refreshed.cookies;
            workerPages[workerIdx] = wp;
            workerCookies[workerIdx] = wc;
          }
        }
      }
    };

    await Promise.all(workerPages.map((_, idx) => runWorker(idx)));
    await Promise.all(workerPages.map(async (wp) => { try { await wp.close(); } catch {} }));

    const finalDocuments = documents;

    console.log(
      `[DIAN CUFE] listado=${cufes.length} encontrados=${finalDocuments.length} fallidos=${failures} duplicados_omitidos=${duplicatesSkipped}`
    );

    if (finalDocuments.length === 0) {
      const msg =
        "[DIAN CUFE] Cero resultados por CUFE. Se detiene el proceso sin fallback legacy (flujo estricto CUFE por CUFE).";
      console.error(msg);
      updateProgress({
        step: "Sin resultados por CUFE (sin fallback legacy)",
        current: 0,
        total: cufes.length,
      });
      throw new Error(msg);
    }

    updateProgress({
      step: `Preparación completada (${finalDocuments.length} IDs listos)` ,
      current: 0,
      total: 0,
    });

    return { documents: finalDocuments, cookies: baseCookieMap };
  } finally {
    if (browser) await browser.close();
  }
}

export async function runDianExtractionPrecheck(
  tokenUrl: string,
  startDate: string | undefined,
  endDate: string | undefined,
  documentDirection: DocumentDirection = "received",
  progressUid?: string
): Promise<PrecheckResult> {
  const direction = documentDirection || "received";
  const isSent = direction === "sent";
  const directionLabel = isSent ? "emitidos" : "recibidos";
  const updateProgress = (data: Partial<ProgressData>) => {
    if (progressUid) {
      const current = progressTracker.get(progressUid) || { step: "", current: 0, total: 0 };
      progressTracker.set(progressUid, { ...current, ...data });
    }
  };

  let browser: Browser | null = null;
  try {
    try {
      updateProgress({ step: "Prevalidando listado DIAN...", current: 0, total: 1 });
      browser = await launchBrowserWithRetry(resolveExecutablePath(), updateProgress);
      const page = await browser.newPage();
      await hardenPageRuntime(page);
      await page.setViewport({ width: 1280, height: 800 });
      page.setDefaultTimeout(120000);
      page.setDefaultNavigationTimeout(120000);

      await navigateWithRetry(page, tokenUrl, 3);
      if (isLoginPage(page.url())) {
        throw new Error("token inválido o expirado durante precheck");
      }

      const listedRecords = await extractListingRecordsFromDownloadTab(page, direction, startDate, endDate);
      if (listedRecords.length === 0) {
        throw new Error("listado DIAN sin registros válidos para el rango");
      }

      const documentUrl = isSent
        ? "https://catalogo-vpfe.dian.gov.co/Document/Sent"
        : "https://catalogo-vpfe.dian.gov.co/Document/Received";
      await navigateWithRetry(page, documentUrl, 2);
      await waitForTableLoad(page);
      await setPageLength(page, 100);
      await waitForFullTableLoad(page, 100);

      const sample = listedRecords[0];
      const found = await findDocumentByUniqueCodeOrDocnum(page, sample.cufe, sample.docnum, new Set<string>(), isSent);
      if (!found?.id) {
        throw new Error(`no se pudo ubicar muestra ${sample.docnum} por CUFE/Folio`);
      }

      updateProgress({
        step: `Prevalidación OK (${listedRecords.length} en listado ${directionLabel})`,
        current: 1,
        total: 1,
      });

      return { ok: true, listedCount: listedRecords.length };
    } catch (precheckErr) {
      const msg = precheckErr instanceof Error ? precheckErr.message : String(precheckErr);
      console.warn(`[DIAN Precheck] Falló validación por listado (${directionLabel}): ${msg}. Continuando flujo normal.`);
      updateProgress({ step: "Prevalidación por listado no disponible, continuando...", current: 1, total: 1 });
      return { ok: true, listedCount: 0 };
    }
  } finally {
    if (browser) await browser.close();
  }
}

async function hardenPageRuntime(page: Page): Promise<void> {
  // En algunos entornos de transpile, page.evaluate puede serializar helpers
  // como __name. Definirlo en runtime evita el fallo "__name is not defined".
  await page.evaluateOnNewDocument(() => {
    const g = globalThis as unknown as { __name?: (fn: unknown, _n?: string) => unknown };
    if (typeof g.__name !== "function") {
      g.__name = (fn) => fn;
    }

    // Asegura identificador global (no solo propiedad) para código estricto
    // que referencia __name como variable libre.
    try {
      (0, eval)("var __name = globalThis.__name;");
    } catch {
      // ignore
    }
  });
}

async function launchBrowserWithRetry(
  executablePath: string | null,
  updateProgress: (data: Partial<ProgressData>) => void
): Promise<Browser> {
  let lastError: Error | null = null;
  const chromiumLibPath = buildChromiumLibPath();

  if (chromiumLibPath) {
    const currentLd = process.env.LD_LIBRARY_PATH || "";
    const parts = [chromiumLibPath, currentLd].filter(Boolean);
    process.env.LD_LIBRARY_PATH = parts.join(":");
  }

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

function buildChromiumLibPath(): string {
  const base = `${process.cwd()}/.chromium-libs`;
  const candidates = [
    `${base}/usr/lib/x86_64-linux-gnu`,
    `${base}/lib/x86_64-linux-gnu`,
    `${base}/usr/lib`,
    `${base}/lib`,
  ];

  const existing = candidates.filter((p) => fs.existsSync(p));
  return existing.join(":");
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

async function applyDateFilter(page: Page, startDate: string, endDate: string, triggerSearch: boolean = true): Promise<void> {
  try {
    // Selector explícito del rango visible (evita caer en hidden StartDate/EndDate).
    const rangeInput = await page.$(
      "#dashboard-report-range, " +
      "input#dashboard-report-range, " +
      "input[type='text'][placeholder*='Rango'], " +
      "input[type='text'][aria-label*='Rango'], " +
      "input[type='text'][placeholder*='Fecha'], " +
      "input[type='text'][aria-label*='Fecha'], " +
      "input[type='text'][id*='range']"
    );

    if (rangeInput) {
      const [sy, sm, sd] = startDate.split("-");
      const [ey, em, ed] = endDate.split("-");
      const sDate = sy && sm && sd ? `${sy}/${sm}/${sd}` : startDate;
      const eDate = ey && em && ed ? `${ey}/${em}/${ed}` : endDate;
      const rangoCompleto = `${sDate} - ${eDate}`;

      await rangeInput.click({ clickCount: 3 });
      await rangeInput.type(rangoCompleto);
      await delay(450);

      // Refuerza los campos que DIAN usa en el POST del formulario.
      // No forzamos campos ocultos; se respeta el comportamiento UI real de DIAN.
    } else {
      console.log("No se encontró input de rango - continuando sin escribir fechas.");
    }

    if (!triggerSearch) return;

    // Click en Buscar del form documents-form y esperar recarga.
    const searchButton =
      (await page.$("form#documents-form button[type='submit']")) ||
      (await page.$("button.btn.btn-success.btn-radian-success")) ||
      (await page.$("button[type='submit']"));

    if (searchButton) {
      try {
        await Promise.all([
          page.waitForNavigation({ waitUntil: "domcontentloaded", timeout: 15000 }).catch(() => null),
          searchButton.click(),
        ]);
      } catch {
        // continuar con espera de tabla
      }
      await delay(500);
      await waitForTableLoad(page);
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

function shouldIgnoreDocType(docType: string, isSentDocuments: boolean = false): boolean {
  const normalized = docType.toLowerCase().trim();
  
  // Siempre ignorar estos tipos
  if (IGNORED_DOC_TYPES.some(ignored => normalized.includes(ignored))) {
    return true;
  }
  
  return false;
}

async function extractDocsFromPage(
  page: Page,
  seenIds: Set<string>,
  isSentDocuments: boolean = false,
  ignoreDocTypeFilter: boolean = true
): Promise<DocumentInfo[]> {
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
        "[data-trackid], [data-id], [id^='doc-'], [id*='track'], " +
        "[onclick*='DownloadZipFiles'], [onclick*='trackId']"
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
          const href = String(el.getAttribute("href") || "");
          const direct = href.match(/trackId=([^&'"\s]+)/i);
          if (direct) trackId = direct[1];
          if (!trackId) {
            const quoted = href.match(/DownloadZipFiles(?:Equivalente)?\s*\(\s*['\"]([^'\"]+)['\"]/i);
            if (quoted) trackId = quoted[1];
          }
          if (!trackId) {
            const generic = href.match(/['\"]([a-f0-9]{32,128})['\"]/i);
            if (generic) trackId = generic[1];
          }
        }
        if (!trackId) {
          const onclick = String(el.getAttribute("onclick") || "");
          const direct = onclick.match(/trackId=([^&'"\s]+)/i);
          if (direct) trackId = direct[1];
          if (!trackId) {
            const quoted = onclick.match(/DownloadZipFiles(?:Equivalente)?\s*\(\s*['\"]([^'\"]+)['\"]/i);
            if (quoted) trackId = quoted[1];
          }
          if (!trackId) {
            const generic = onclick.match(/['\"]([a-f0-9]{32,128})['\"]/i);
            if (generic) trackId = generic[1];
          }
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
          const href = String(link.getAttribute("href") || "");
          const directHref = href.match(/trackId=([^&'"\s]+)/i);
          if (directHref) trackId = directHref[1];
          if (!trackId) {
            const quotedHref = href.match(/DownloadZipFiles(?:Equivalente)?\s*\(\s*['\"]([^'\"]+)['\"]/i);
            if (quotedHref) trackId = quotedHref[1];
          }
          if (!trackId) {
            const genericHref = href.match(/['\"]([a-f0-9]{32,128})['\"]/i);
            if (genericHref) trackId = genericHref[1];
          }

          if (!trackId) {
            const onclick = String(link.getAttribute("onclick") || "");
            const directOnclick = onclick.match(/trackId=([^&'"\s]+)/i);
            if (directOnclick) trackId = directOnclick[1];
            if (!trackId) {
              const quotedOnclick = onclick.match(/DownloadZipFiles(?:Equivalente)?\s*\(\s*['\"]([^'\"]+)['\"]/i);
              if (quotedOnclick) trackId = quotedOnclick[1];
            }
            if (!trackId) {
              const genericOnclick = onclick.match(/['\"]([a-f0-9]{32,128})['\"]/i);
              if (genericOnclick) trackId = genericOnclick[1];
            }
          }
          if (trackId) break;
        }
      }

      // Metodo 5: fallback sobre HTML completo de la fila
      if (!trackId) {
        const raw = String((row as HTMLElement).outerHTML || "");
        const direct = raw.match(/trackId=([^&'"\s]+)/i);
        if (direct) trackId = direct[1];
        if (!trackId) {
          const quoted = raw.match(/DownloadZipFiles(?:Equivalente)?\s*\(\s*['\"]([^'\"]+)['\"]/i);
          if (quoted) trackId = quoted[1];
        }
        if (!trackId) {
          const generic = raw.match(/['\"]([a-f0-9]{32,128})['\"]/i);
          if (generic) trackId = generic[1];
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

      const cufe = tds[13]?.textContent?.trim() || tds[12]?.textContent?.trim() || "";

      results.push({ id: trackId, docnum, nit, cufe, docType, documentTypeId, fechaValidacion, fechaGeneracion });
    }

    return results;
  }, isSentDocuments);

  let skippedCount = 0;
  
  for (const item of items) {
    if (!seenIds.has(item.id)) {
      // Filtrar tipos no relevantes solo cuando se solicita explícitamente.
      if (ignoreDocTypeFilter && shouldIgnoreDocType(item.docType, isSentDocuments)) {
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
  
  if (ignoreDocTypeFilter && skippedCount > 0) {
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

    const cookieHeader = (await page.cookies()).map((c) => `${c.name}=${c.value}`).join("; ");

    const exportPageBefore = await fetch(`${baseUrl}/Document/Export`, {
      method: "GET",
      headers: {
        Cookie: cookieHeader,
        Referer: `${baseUrl}/Document/Export`,
      },
    }).then((r) => r.text());

    const reusableLinks = findReusableExportLinks(exportPageBefore, direction, startDate, endDate);
    if (reusableLinks.length > 0) {
      console.log(`[DIAN Export] Se encontraron ${reusableLinks.length} listados reutilizables para el rango.`);
      for (let i = 0; i < reusableLinks.length; i++) {
        const reusableLink = reusableLinks[i];
        try {
          console.log(`[DIAN Export] Reutilizando listado #${i + 1}: ${reusableLink}`);
          const downloaded = await fetch(`${baseUrl}${reusableLink}`, {
            method: "GET",
            headers: { Cookie: cookieHeader, Referer: `${baseUrl}/Document/Export` },
          }).then((r) => r.arrayBuffer());

          const zipBuffer = Buffer.from(new Uint8Array(downloaded));
          const reusedRecords = await parseListingRecordsFromExportZip(zipBuffer, direction);
          if (reusedRecords.length > 0) {
            console.log(`[DIAN Export] Reutilizado OK #${i + 1}: ${reusedRecords.length} CUFEs (${direction}).`);
            return reusedRecords;
          }

          console.warn(`[DIAN Export] Listado reutilizado #${i + 1} sin CUFEs válidos.`);
        } catch (reuseErr) {
          console.warn(`[DIAN Export] Falló descarga de listado reutilizado #${i + 1}:`, reuseErr);
        }
      }

      console.warn("[DIAN Export] Ningún listado reutilizable fue válido; se intentará regenerar.");
    }

    const existingRks = new Set(parseDownloadLinksFromExportHtml(exportPageBefore).map((l) => l.rk));
    const formData = await page.evaluate(() => {
      const token = (document.querySelector("input[name='__RequestVerificationToken']") as HTMLInputElement | null)?.value || "";
      const type = (document.querySelector("input[name='Type']") as HTMLInputElement | null)?.value || "0";
      const amountAdmin = (document.querySelector("input[name='AmountAdmin']") as HTMLInputElement | null)?.value || "100000";
      return { token, type, amountAdmin };
    });

    if (!formData.token) return [];

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

    // DIAN genera el archivo en segundo plano. Polling hasta 5 minutos.
    const pollTimeoutMs = 300_000;
    const pollEveryMs = 4_000;
    const startedAt = Date.now();

    let selectedLink: string | null = null;
    while (Date.now() - startedAt < pollTimeoutMs) {
      const exportHtml = await fetch(`${baseUrl}/Document/Export`, {
        method: "GET",
        headers: {
          Cookie: cookieHeader,
          Referer: `${baseUrl}/Document/Export`,
        },
      }).then((r) => r.text());

      const links = parseDownloadLinksFromExportHtml(exportHtml);
      const newlyGenerated = links.find((l) => !existingRks.has(l.rk));
      if (newlyGenerated) {
        selectedLink = newlyGenerated.href;
        break;
      }

      await delay(pollEveryMs);
    }

    if (!selectedLink) {
      throw new Error("No se detectó un listado nuevo para el rango solicitado dentro de 5 minutos.");
    }

    const downloaded = await fetch(`${baseUrl}${selectedLink}`, {
      method: "GET",
      headers: { Cookie: cookieHeader, Referer: `${baseUrl}/Document/Export` },
    }).then((r) => r.arrayBuffer());
    const zipBuffer = Buffer.from(new Uint8Array(downloaded));

    return await parseListingRecordsFromExportZip(zipBuffer, direction);
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    if (msg.includes("TOKEN_EXPIRED")) {
      throw err;
    }
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
  // DIAN puede devolver:
  // 1) ZIP contenedor con un .xlsx adentro
  // 2) el .xlsx directamente (que también es un ZIP OOXML)
  // Este bloque soporta ambos formatos.
  let xlsxBuffer: Buffer | null = null;
  const maybeZip = await JSZip.loadAsync(zipBuffer);

  const embeddedXlsxName = Object.keys(maybeZip.files).find(
    (n) => n.toLowerCase().endsWith(".xlsx") && !maybeZip.files[n].dir
  );

  if (embeddedXlsxName) {
    xlsxBuffer = await maybeZip.files[embeddedXlsxName].async("nodebuffer");
  } else if (maybeZip.file("xl/workbook.xml")) {
    // El archivo descargado ya es el .xlsx
    xlsxBuffer = zipBuffer;
  }

  if (!xlsxBuffer) {
    console.warn("[DIAN Export] Formato de archivo no reconocido al parsear listado");
    return [];
  }

  const xlsx = await JSZip.loadAsync(xlsxBuffer);

  const sharedStringsXml = await xlsx.file("xl/sharedStrings.xml")?.async("string");
  const sheetPath = Object.keys(xlsx.files).find(
    (n) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(n) && !xlsx.files[n].dir
  );
  const sheetXml = sheetPath ? await xlsx.file(sheetPath)?.async("string") : null;
  if (!sheetXml) return [];

  const sharedStrings = parseSharedStrings(sharedStringsXml || "");
  const rows = parseSheetRows(sheetXml, sharedStrings);
  if (rows.length < 2) {
    console.warn("[DIAN Export] XLSX sin filas de datos");
    return [];
  }

  const headers = rows[0].map((h) => h.trim().toLowerCase());
  const cufeIdx = headers.findIndex((h) => h.includes("cufe") || h.includes("cude") || h.includes("código único") || h.includes("codigo unico"));
  const folioIdx = headers.findIndex((h) => h === "folio");
  const groupIdx = headers.findIndex((h) => h === "grupo");

  if (cufeIdx < 0 || folioIdx < 0) {
    console.warn("[DIAN Export] Encabezados esperados no encontrados", {
      cufeIdx,
      folioIdx,
      headers: rows[0],
    });
    return [];
  }

  const out: ListingRecord[] = [];

  const expectedGroupText = direction === "sent" ? "emitid" : "recibid";

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const group = (groupIdx >= 0 ? row[groupIdx] : "").toLowerCase();

    // Si viene grupo, filtrar estrictamente por dirección solicitada.
    if (group && !group.includes(expectedGroupText)) continue;

    const cufe = normalizeCufe(row[cufeIdx] || "");

    const docnum = (folioIdx >= 0 ? row[folioIdx] : "").trim();

    if (!cufe) continue;
    if (!docnum) continue;

    out.push({ cufe, docnum });
  }

  const dedup = new Map<string, ListingRecord>();
  for (const rec of out) {
    const key = `${rec.docnum}::${rec.cufe}`;
    if (!dedup.has(key)) dedup.set(key, rec);
  }
  const finalRows = Array.from(dedup.values());
  if (finalRows.length === 0) {
    console.warn("[DIAN Export] No se pudieron mapear CUFEs desde el XLSX", {
      cufeIdx,
      folioIdx,
      groupIdx,
      totalRows: rows.length,
      sampleHeaders: rows[0]?.slice(0, 10),
    });
  }
  return finalRows;
}

function parseSharedStrings(xml: string): string[] {
  if (!xml) return [];
  const strings: string[] = [];
  const siMatches = xml.match(/<(?:\w+:)?si[\s\S]*?<\/(?:\w+:)?si>/g) || [];
  for (const si of siMatches) {
    const textParts = Array.from(si.matchAll(/<(?:\w+:)?t[^>]*>([\s\S]*?)<\/(?:\w+:)?t>/g)).map((m) => decodeXml(m[1]));
    strings.push(textParts.join(""));
  }
  return strings;
}

function parseSheetRows(xml: string, sharedStrings: string[]): string[][] {
  const rows: string[][] = [];
  const rowMatches = xml.match(/<(?:\w+:)?row[^>]*>[\s\S]*?<\/(?:\w+:)?row>/g) || [];

  for (const rowXml of rowMatches) {
    const cells = Array.from(rowXml.matchAll(/<(?:\w+:)?c([^>]*)>([\s\S]*?)<\/(?:\w+:)?c>/g));
    const rowValues: string[] = [];

    for (const cellMatch of cells) {
      const attrs = cellMatch[1];
      const body = cellMatch[2];
      const ref = (attrs.match(/r="([A-Z]+)\d+"/) || [])[1] || "A";
      const colIndex = colLettersToIndex(ref);
      while (rowValues.length < colIndex) rowValues.push("");

      const type = (attrs.match(/t="([^"]+)"/) || [])[1] || "";
      const valueMatch = body.match(/<(?:\w+:)?v>([\s\S]*?)<\/(?:\w+:)?v>/);
      const inlineMatch = body.match(/<(?:\w+:)?t[^>]*>([\s\S]*?)<\/(?:\w+:)?t>/);

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

function parseDownloadLinksFromExportHtml(html: string): Array<{ href: string; rk: string }> {
  const links = Array.from(html.matchAll(/href="(\/Document\/DownloadExportedZipFile\?pk=[^"]+)"/g))
    .map((m) => m[1].replace(/&amp;/g, "&"));

  const parsed: Array<{ href: string; rk: string }> = [];
  for (const href of links) {
    const rkMatch = href.match(/[?&]rk=([^&]+)/i);
    const rk = rkMatch?.[1] || "";
    if (!rk) continue;
    parsed.push({ href, rk });
  }

  return parsed;
}

function findReusableExportLinks(
  html: string,
  direction: DocumentDirection,
  startDate?: string,
  endDate?: string
): string[] {
  if (!startDate || !endDate) return [];
  const requestedStart = parseIsoDate(startDate);
  const requestedEnd = parseIsoDate(endDate);
  if (!requestedStart || !requestedEnd) return [];

    const desiredTypes = direction === "sent" ? ["enviados", "emitidos"] : ["recibidos"];

  const tbodyMatch = html.match(/<table[^>]*id="tableExport"[\s\S]*?<tbody>([\s\S]*?)<\/tbody>/i);
  if (!tbodyMatch) return [];

  const rows = Array.from(tbodyMatch[1].matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/gi)).map((m) => m[1]);
  const matches: string[] = [];
  for (const row of rows) {
    const tds = Array.from(row.matchAll(/<td[^>]*>([\s\S]*?)<\/td>/gi)).map((m) => m[1]);
    if (tds.length < 4) continue;

    const rangeText = stripHtml(tds[2]).toLowerCase();
    const typeText = stripHtml(tds[3]).toLowerCase();

    const parsedRange = parseExportRange(rangeText);
    if (!parsedRange) continue;

    // "Enviados y Recibidos" se puede reutilizar siempre que luego se filtre
    // por columna "Grupo" en el parser del XLSX según dirección solicitada.
    const isBothType = typeText.includes("enviados") && typeText.includes("recibidos");
    const typeMatches = isBothType || desiredTypes.some((dt) => typeText.includes(dt));

    // Reutilizar solo si el rango del listado coincide exactamente.
    const rangeMatches = parsedRange.start === requestedStart && parsedRange.end === requestedEnd;
    if (!typeMatches || !rangeMatches) continue;

    const hrefMatch = row.match(/href="(\/Document\/DownloadExportedZipFile\?pk=[^"]+)"/i);
    if (!hrefMatch) continue;
    matches.push(hrefMatch[1].replace(/&amp;/g, "&"));
  }

  // La tabla de export normalmente lista más nuevo primero; priorizar ese orden
  // mantiene el comportamiento manual (tomar la primera fila visible que coincide).
  return matches;
}

function parseIsoDate(value: string): number | null {
  const m = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  return Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function parseExportRange(text: string): { start: number; end: number } | null {
  const m = text.match(/desde\s+(\d{2})-(\d{2})-(\d{4})\s+hasta\s+(\d{2})-(\d{2})-(\d{4})/i);
  if (!m) return null;
  const start = Date.UTC(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  const end = Date.UTC(Number(m[6]), Number(m[5]) - 1, Number(m[4]));
  return { start, end };
}

function toDianDisplayDate(dateISO: string): string {
  const [year, month, day] = dateISO.split("-");
  if (!year || !month || !day) return dateISO;
  return `${day.padStart(2, "0")}-${month.padStart(2, "0")}-${year}`;
}

function stripHtml(raw: string): string {
  return raw.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
}

async function findDocumentByUniqueCodeOrDocnum(
  page: Page,
  cufe: string,
  docnum: string,
  seenIds: Set<string>,
  isSentDocuments: boolean,
  searchTriggeredByCaller: boolean = false
): Promise<DocumentInfo | null> {
  const candidates = [cufe].filter(Boolean);
  for (const term of candidates) {
    try {
      // 1) Flujo UI directo igual al manual: Código único + Buscar.
      // No tocar otros filtros para no alterar el estado de la búsqueda.

      const uniqueInputSelectors = [
        "#DocumentKey",
        "input[name='DocumentKey']",
      ];
      let searchedByUniqueCode = false;

      for (const sel of uniqueInputSelectors) {
        const input = await page.$(sel);
        if (!input) continue;

        // Escritura robusta del CUFE sin depender de eventos de teclado/click
        // (reduce timeouts Runtime.callFunctionOn en cargas altas).
        await page.$eval(sel, (el, value) => {
          const i = el as HTMLInputElement;
          i.value = String(value || "");
          i.dispatchEvent(new Event("input", { bubbles: true }));
          i.dispatchEvent(new Event("change", { bubbles: true }));
        }, term);

        const inputValue = await page.$eval(sel, (el) => (el as HTMLInputElement).value || "");
        if (normalizeCufe(inputValue) !== normalizeCufe(term)) {
          console.warn("[DIAN CUFE] El campo Código único no conservó el valor esperado", {
            expected: normalizeCufe(term).slice(0, 16),
            actual: normalizeCufe(inputValue).slice(0, 16),
          });
        }

        searchedByUniqueCode = true;
        break;
      }

      if (searchedByUniqueCode && !searchTriggeredByCaller) {
        const submitted = await page.evaluate(() => {
          const form = document.querySelector("form#documents-form") as HTMLFormElement | null;
          if (!form) return false;
          if (typeof form.requestSubmit === "function") {
            form.requestSubmit();
            return true;
          }
          form.submit();
          return true;
        }).catch(() => false);

        if (!submitted) {
          const searchButton =
            (await page.$("form#documents-form button.btn.btn-success.btn-radian-success")) ||
            (await page.$("form#documents-form button[type='submit']")) ||
            (await page.$("button.btn.btn-success.btn-radian-success")) ||
            (await page.$("button[type='submit']"));
          if (searchButton) {
            await searchButton.click();
          } else {
            await page.keyboard.press("Enter");
          }
        }
      }

      if (searchedByUniqueCode) {
        await delay(80);
        await waitForTableLoad(page);

        // Usar un set temporal en búsqueda puntual para no descartar un resultado
        // válido por deduplicación previa de otro CUFE.
        const foundByUniqueCode = await extractDocsFromCurrentPageHtml(page, new Set<string>(), isSentDocuments);
        if (foundByUniqueCode.length > 0) {
          const exact = foundByUniqueCode.find((d) =>
            (docnum && d.docnum === docnum) ||
            (cufe && normalizeCufe(d.cufe || "") === normalizeCufe(cufe))
          );
          const selected = exact || foundByUniqueCode[0];
          if (!seenIds.has(selected.id)) seenIds.add(selected.id);
          return selected;
        }

        const visibleRows = await page.evaluate(() =>
          document.querySelectorAll("#tableDocuments tbody tr:not(.dataTables_empty)").length
        );
        if (visibleRows > 0) {
          // Fallback ultra directo: tomar la primera fila visible exactamente
          // como en uso manual (si hay una fila, descargar esa).
          const firstRowDoc = await page.evaluate((isSentDocumentsInPage) => {
            const row = document.querySelector("#tableDocuments tbody tr:not(.dataTables_empty)") as HTMLTableRowElement | null;
            if (!row) return null;

            let trackId = "";
            const pickFrom = row.querySelectorAll(
              ".download-document, .download-support-document, .download-eventos, .download-equivalente-document, .download-individual-payroll, [data-trackid], [data-id], [id], a[href*='trackId'], a[href*='DownloadZipFiles'], [onclick*='DownloadZipFiles'], [onclick*='trackId']"
            );

            for (const el of pickFrom) {
              const id = (el as HTMLElement).id || "";
              const d1 = el.getAttribute("data-trackid") || "";
              const d2 = el.getAttribute("data-id") || "";
              const href = el.getAttribute("href") || "";
              const onclick = el.getAttribute("onclick") || "";
              const m1 = href.match(/trackId=([^&'"\s]+)/i);
              const m2 = onclick.match(/trackId=([^&'"\s]+)/i);
              const m3 = href.match(/DownloadZipFiles(?:Equivalente)?\s*\(\s*['\"]([^'\"]+)['\"]/i);
              const m4 = onclick.match(/DownloadZipFiles(?:Equivalente)?\s*\(\s*['\"]([^'\"]+)['\"]/i);
              trackId = id || d1 || d2 || m1?.[1] || m2?.[1] || m3?.[1] || m4?.[1] || "";
              if (trackId) break;
            }

            if (!trackId) {
              trackId = row.getAttribute("data-id") || row.getAttribute("data-trackid") || row.id || "";
            }
            if (!trackId) return null;

            const tds = row.querySelectorAll("td");
            const docnum = tds[4]?.textContent?.trim() || "";
            const docType = tds[5]?.textContent?.trim() || "";
            const nit = isSentDocumentsInPage
              ? (tds[8]?.textContent?.trim() || tds[6]?.textContent?.trim() || "")
              : (tds[6]?.textContent?.trim() || tds[8]?.textContent?.trim() || "");
            const cufeValue = tds[13]?.textContent?.trim() || tds[12]?.textContent?.trim() || "";

            return { id: trackId, docnum, nit, cufe: cufeValue, docType };
          }, isSentDocuments);

          if (firstRowDoc?.id) {
            if (!seenIds.has(firstRowDoc.id)) seenIds.add(firstRowDoc.id);
            return firstRowDoc;
          }

          // Fallback final: en muchos casos DIAN usa el CUFE como trackId del botón.
          // Si hay fila visible y no se pudo extraer id por DOM, continuar con CUFE
          // como identificador para no perder el documento.
          const docnumFromFirstRow = await page.evaluate(() => {
            const row = document.querySelector("#tableDocuments tbody tr:not(.dataTables_empty)") as HTMLTableRowElement | null;
            if (!row) return "";
            const tds = row.querySelectorAll("td");
            return tds[4]?.textContent?.trim() || "";
          });

          const cufeAsId = normalizeCufe(cufe || "");
          if (cufeAsId) {
            if (!seenIds.has(cufeAsId)) seenIds.add(cufeAsId);
            return {
              id: cufeAsId,
              docnum: docnumFromFirstRow,
              nit: "",
              cufe,
              docType: "",
            };
          }

          throw new Error(`Resultado visible sin trackId extraíble para CUFE ${cufe.slice(0, 16)}...`);
        }
      }
    } catch (err) {
      // Propagar para no ocultar errores técnicos de extracción/estado.
      throw err;
    }
  }

  return null;
}

async function extractDocsFromCurrentPageHtml(page: Page, seenIds: Set<string>, isSentDocuments: boolean): Promise<DocumentInfo[]> {
  // Reusar el extractor DOM principal (más robusto para trackId y variantes de botones)
  // también en búsqueda CUFE puntual.
  return extractDocsFromPage(page, seenIds, isSentDocuments, false);
}

function extractDocsFromHtml(html: string, seenIds: Set<string>, isSentDocuments: boolean): DocumentInfo[] {
  // Tomar exclusivamente la tabla de documentos DIAN para evitar leer tbody
  // de otros componentes del layout.
  const tableMatch = html.match(/<table[^>]*id=["']tableDocuments["'][^>]*>([\s\S]*?)<\/table>/i);
  if (!tableMatch) return [];

  const tbodyMatch = tableMatch[1].match(/<tbody[^>]*>([\s\S]*?)<\/tbody>/i);
  if (!tbodyMatch) return [];

  const rows = Array.from(tbodyMatch[1].matchAll(/<tr[^>]*>([\s\S]*?)<\/tr>/gi)).map((m) => m[1]);
  const docs: DocumentInfo[] = [];

  for (const row of rows) {
    if (/dataTables_empty/i.test(row)) continue;
    let trackId = "";
    const m1 = row.match(/trackId=([^&'"\s]+)/i);
    const m2 = row.match(/DownloadZipFiles(?:Equivalente)?\s*\(\s*['\"]([^'\"]+)['\"]/i);
    const m3 = row.match(/data-trackid=['\"]([^'\"]+)['\"]/i);
    const m4 = row.match(/data-id=['\"]([^'\"]+)['\"]/i);
    const m5 = row.match(/<button[^>]*\sid=['\"]([^'\"]+)['\"]/i);
    const m6 = row.match(/<tr[^>]*\sdata-id=['\"]([^'\"]+)['\"]/i);
    trackId = m1?.[1] || m2?.[1] || m3?.[1] || m4?.[1] || m5?.[1] || m6?.[1] || "";
    if (!trackId || seenIds.has(trackId)) continue;

    const cells = Array.from(row.matchAll(/<td[^>]*>([\s\S]*?)<\/td>/gi)).map((c) => stripHtml(c[1]));
    const docnum = cells[4] || "";
    const docType = cells[5] || "";
    const nit = isSentDocuments ? (cells[8] || cells[6] || "") : (cells[6] || cells[8] || "");
    const cufe = cells[13] || cells[12] || "";

    if (shouldIgnoreDocType(docType, isSentDocuments)) continue;
    seenIds.add(trackId);
    docs.push({ id: trackId, docnum, nit, cufe, docType });
  }

  return docs;
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
          "[data-trackid], [data-id], a[href*='trackId'], " +
          "a[href*='DownloadZipFiles'], [onclick*='DownloadZipFiles'], [onclick*='trackId']"
        );
        
        for (const el of elements) {
          let tid = el.id || el.getAttribute("data-trackid") || el.getAttribute("data-id");
          
          if (!tid) {
            const href = el.getAttribute("href") || "";
            const match = href.match(/trackId=([A-Za-z0-9-]+)/i);
            if (match) tid = match[1];
          }

          if (!tid) {
            const onclick = el.getAttribute("onclick") || "";
            const matchOnclick = onclick.match(/DownloadZipFiles(?:Equivalente)?\s*\(\s*['\"]([A-Za-z0-9-]+)['\"]/i)
              || onclick.match(/trackId=([A-Za-z0-9-]+)/i);
            if (matchOnclick) tid = matchOnclick[1];
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

      const currentUrl = page.url();
      if (isLoginPage(currentUrl)) {
        const tokenExpiredError = new Error(
          "TOKEN_EXPIRED: El token ha expirado. Por favor, genera un nuevo token desde el portal DIAN."
        );

        // Regla operativa solicitada: si al segundo intento redirige a User/Login,
        // informar expiración de token de inmediato.
        if (attempt >= 2) {
          throw tokenExpiredError;
        }

        lastError = tokenExpiredError;
        console.log(`Redirección a login detectada en intento ${attempt}, reintentando...`);
        await delay(1200);
        continue;
      }
      
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
