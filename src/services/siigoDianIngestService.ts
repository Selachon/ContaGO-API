import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";
import { v4 as uuidv4 } from "uuid";
import type { Page } from "puppeteer";
import {
  acquireDianJobSlot,
  getCufeListing,
  extractDocumentIds,
  extractDocumentIdsByCufe,
  fetchZipToFile,
  downloadZipInPage,
  type CufeDownloadItem,
  type ListingRecord,
} from "./dianScraper.js";
import { processXmlBatch, type BatchItem } from "./siigoAccountingService.js";
import { getIngestedCufes, getIngestedTrackIds, recordIngestedCufes } from "./siigoIngestedCufesService.js";
import type { DocumentDirection, ProgressData } from "../types/dian.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOWNLOADS_DIR = path.join(__dirname, "../../downloads");

if (!fs.existsSync(DOWNLOADS_DIR)) {
  fs.mkdirSync(DOWNLOADS_DIR, { recursive: true });
}

const delay = (ms: number) => new Promise<void>((r) => setTimeout(r, ms));

export type IngestGrupo = "Emitidos" | "Recibidos" | "Todos";

export interface IngestFailure {
  cufe: string;
  error: string;
}

export interface IngestStats {
  /** CUFEs listados (tras aplicar tope/límite demo). */
  listed: number;
  /** CUFEs omitidos por estar ya registrados (descargados antes). */
  alreadyRegistered: number;
  /** Documentos descargados con éxito en esta corrida. */
  downloaded: number;
  /** CUFEs no recuperables tras agotar las rondas de reintento. */
  failed: number;
  /** Número máximo de rondas de reintento usadas en alguna dirección. */
  rounds: number;
}

export interface IngestResult {
  /** Ítems listos para revisar/causar, mismo shape que /accounting/process-batch. */
  items: BatchItem[];
  stats: IngestStats;
  /** CUFEs que NO se pudieron descargar tras agotar reintentos (reporte explícito). */
  failures: IngestFailure[];
}

export interface IngestOptions {
  /** Empresa Siigo asociada; se usa para el registro persistente de CUFEs. */
  companyId: string;
  tokenUrl: string;
  /** Rango opcional para acotar la descarga por CUFE. Si se omite (modo listado),
   *  la descarga busca por CUFE sin filtro de fecha, igual que "Exportar Excel DIAN". */
  fechaInicio?: string;
  fechaFin?: string;
  grupo: IngestGrupo;
  nitReceptor?: string;
  /** CUFEs ya obtenidos de un listado subido por el usuario. Si se provee, se
   *  omite el paso lento `getCufeListing` y se descargan estos CUFEs directamente
   *  (la herramienta es solo Recibidos, así que aplican a la dirección received). */
  providedCufes?: string[];
  /** Tope de documentos (p. ej. límite demo o DIAN_MAX_DOCUMENTS). */
  maxDocuments?: number;
  /** Rondas de reintento por CUFE fallido, además de la pasada inicial. Def: 3. */
  retryRounds?: number;
  /** Si true, re-descarga también los CUFEs ya registrados (ignora el registro). */
  forceRedownload?: boolean;
  /**
   * CUFEs a OMITIR de la ingesta (ya causados en SIIGO / ya presentes en pantalla).
   * Para "Causación + Caja": el dedup correcto es contra el buzón "Facturas"
   * (causadas) + lo que ya está en la tabla de trabajo, NO contra el historial de
   * descargas. Así las facturas no causadas se mantienen/recuperan al re-traer.
   */
  skipCufes?: string[];
  onProgress?: (p: ProgressData) => void;
  isCancelled?: () => boolean;
}

function grupoToDirections(grupo: IngestGrupo): DocumentDirection[] {
  if (grupo === "Emitidos") return ["sent"];
  if (grupo === "Recibidos") return ["received"];
  return ["received", "sent"]; // Todos
}

/**
 * Descarga los CUFEs de una dirección con reintentos: una pasada inicial y hasta
 * `retryRounds` rondas adicionales SOLO sobre los CUFEs que sigan fallando. Esto
 * garantiza que ningún CUFE se "salte" silenciosamente; los que no se recuperen
 * tras agotar las rondas se devuelven explícitamente en `failures`.
 */
async function downloadDirectionWithRetries(
  opts: IngestOptions,
  direction: DocumentDirection,
  cufes: string[],
  tempDir: string
): Promise<{ success: CufeDownloadItem[]; failures: IngestFailure[]; rounds: number }> {
  // Mismo paso a paso que "Exportar Excel DIAN" (dianCufeDownload): entrar al token
  // UNA vez por sesión, resolver los documentos por CUFE y descargar cada ZIP por
  // HTTP con las cookies de esa sesión (sin reabrir el token). Reintenta los que
  // fallen abriendo una sesión nueva. Esto evita el falso "TOKEN_EXPIRED" del
  // flujo anterior (que reabría el token —de un solo uso— en cada ronda).
  const retryRounds = Math.max(0, opts.retryRounds ?? 3);
  const dirLabel = direction === "sent" ? "emitidos" : "recibidos";
  const total = cufes.length;
  const dlResults = new Map<string, CufeDownloadItem>();
  let dlOk = 0;

  // Semáforo de descargas en paralelo con contrapresión (cookies frescas).
  const MAX_DL = Math.max(1, Math.min(Number(process.env.DIAN_DOWNLOAD_WORKERS || 5), 12));
  let dlSlots = MAX_DL;
  const dlQueue: Array<() => void> = [];
  const acquireDl = (): Promise<void> =>
    dlSlots > 0 ? (dlSlots--, Promise.resolve()) : new Promise<void>((res) => dlQueue.push(res));
  const releaseDl = () => { const n = dlQueue.shift(); if (n) n(); else dlSlots++; };

  const downloadDoc = async (
    doc: { id: string; docnum: string; nit?: string; cufe?: string; docType?: string },
    cookies: Record<string, string>,
    page?: Page
  ): Promise<void> => {
    const key = normCufe(doc.cufe || "") || (doc.id || "").toLowerCase();
    const isEquiv = doc.docType?.toLowerCase().includes("equivalente") ?? false;
    const safeNit = (doc.nit || "SinNIT").replace(/[^a-zA-Z0-9_-]/g, "_");
    const safeDoc = (doc.docnum || doc.id.slice(0, 12)).replace(/[^a-zA-Z0-9_-]/g, "_");
    const destPath = path.join(tempDir, `${safeNit} - ${safeDoc}.zip`);
    try {
      // Descarga disparando el botón del portal (corre el reCAPTCHA) y captura el
      // ZIP. Si no hay página, cae al fetch de Node (entornos sin reCAPTCHA).
      if (page) {
        const dlDir = path.join(tempDir, `_dl_${key.slice(0, 16)}`);
        const buf = await downloadZipInPage(page, doc.id, isEquiv, dlDir);
        fs.writeFileSync(destPath, buf);
        try { fs.rmSync(dlDir, { recursive: true, force: true }); } catch { /* noop */ }
      } else {
        const base = isEquiv
          ? `${DIAN_BASE}/Document/DownloadZipFilesEquivalente?trackId=`
          : `${DIAN_BASE}/Document/DownloadZipFiles?trackId=`;
        const cookieHeader = Object.entries(cookies).map(([k, v]) => `${k}=${v}`).join("; ");
        await fetchZipToFile(`${base}${doc.id}`, destPath, cookieHeader);
      }
      dlResults.set(key, { cufe: doc.cufe || "", trackId: doc.id, destPath, docnum: doc.docnum, nit: doc.nit || "", success: true });
      dlOk++;
      opts.onProgress?.({ step: `Descargando documentos (${dirLabel}): ${dlOk}/${total}…`, current: dlOk, total });
    } catch (err) {
      dlResults.set(key, { cufe: doc.cufe || "", trackId: doc.id, destPath: null, docnum: doc.docnum, nit: doc.nit || "", success: false, error: err instanceof Error ? err.message : String(err) });
    }
  };

  let pending: ListingRecord[] = cufes.map((c) => ({ cufe: c, docnum: "" }));
  let round = 0;
  while (pending.length > 0 && round <= retryRounds) {
    if (opts.isCancelled?.()) break;
    const label = round === 0 ? "Descargando documentos" : `Reintento ${round}/${retryRounds}`;
    const promises: Promise<void>[] = [];
    try {
      await extractDocumentIdsByCufe(
        opts.tokenUrl,
        opts.fechaInicio,
        opts.fechaFin,
        undefined,
        direction,
        (p) => opts.onProgress?.({
          // Mientras no se complete ninguna descarga, muestra la fase REAL del
          // scraper (navegando/buscando) para no parecer congelado en 0/N.
          step: dlOk > 0 ? `${label} (${dirLabel}): ${dlOk}/${total}…` : (p?.step || `${label} (${dirLabel})…`),
          current: dlOk,
          total,
        }),
        async ({ doc, cookies, page }) => {
          // Con descarga EN-navegador hay que serializar sobre la página del worker
          // (no se puede fetch en paralelo mientras el worker navega a la siguiente
          // búsqueda). onDocumentFound ya se await en serie, así que esperamos aquí.
          if (page) {
            await downloadDoc(doc, cookies, page);
          } else {
            await acquireDl();
            const dl = (async () => { try { await downloadDoc(doc, cookies); } finally { releaseDl(); } })();
            promises.push(dl);
          }
        },
        pending,
      );
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      console.warn(`[Siigo Ingest] Sesión de descarga falló (${dirLabel}, ronda ${round}): ${msg}`);
      // El token DIAN es de UN SOLO USO: si expiró/lo bloquearon, reabrir la sesión
      // con el mismo token es inútil (vuelve a fallar) y RETIENE el cupo global de
      // jobs (deja a los demás en "turno 1"). Abortamos de inmediato con un error
      // claro para que el usuario solicite un token nuevo. `extractDocumentIdsByCufe`
      // ya reintenta internamente los bloqueos transitorios antes de lanzar esto.
      if (/TOKEN_EXPIRED|DIAN_BLOCKED|expirad/i.test(msg)) {
        await Promise.all(promises);
        if (dlOk === 0) {
          throw new Error("El token de la DIAN expiró o fue bloqueado por el portal. Solicita un token nuevo en la DIAN y vuelve a intentar (no reutilices el mismo token).");
        }
        break; // algo se alcanzó a descargar: corta y devuelve lo logrado
      }
    }
    await Promise.all(promises);

    pending = pending.filter((r) => !dlResults.get(normCufe(r.cufe || ""))?.destPath);
    if (pending.length === 0) break;
    round++;
    if (round <= retryRounds) await delay(2000 * round);
  }

  const success = [...dlResults.values()].filter((r) => r.destPath);
  const failures: IngestFailure[] = pending.map((r) => ({
    cufe: r.cufe,
    error: dlResults.get(normCufe(r.cufe || ""))?.error || "No se pudo descargar tras agotar reintentos",
  }));
  return { success, failures, rounds: Math.min(round, retryRounds) };
}

/**
 * Ingesta automática para "Contabilizar Gastos Siigo": obtiene el listado de
 * CUFEs del portal DIAN, descarga cada documento (con reintentos por CUFE) y los
 * parsea con `processXmlBatch`, devolviendo los mismos ítems que la subida manual
 * de ZIP. DEBE invocarse dentro del contexto de empresa (runWithSiigoCompany)
 * porque `processXmlBatch` consulta Siigo para detectar facturas ya causadas.
 */
export async function ingestFromDian(opts: IngestOptions): Promise<IngestResult> {
  // Misma cola global que las herramientas DIAN: el auto-ingest también scrapea
  // el catálogo y compite por el presupuesto anti-bot de la IP.
  const releaseDianJobSlot = await acquireDianJobSlot((pos) =>
    opts.onProgress?.({ step: `En cola para evitar el bloqueo de DIAN (turno ${pos})...`, current: 0, total: 0 })
  );
  if (opts.isCancelled?.()) {
    releaseDianJobSlot();
    return { items: [], stats: { listed: 0, alreadyRegistered: 0, downloaded: 0, failed: 0, rounds: 0 }, failures: [] };
  }
  const directions = grupoToDirections(opts.grupo);
  const sessionId = uuidv4();
  const tempDir = path.join(DOWNLOADS_DIR, `siigo-ingest-${sessionId}`);
  fs.mkdirSync(tempDir, { recursive: true });

  const allSuccess: CufeDownloadItem[] = [];
  const allFailures: IngestFailure[] = [];
  let listed = 0;
  let alreadyRegistered = 0;
  let maxRounds = 0;

  // Registro persistente de CUFEs ya descargados por esta empresa (para omitir
  // re-descargas). Si forceRedownload, se ignora.
  const knownCufes = opts.forceRedownload ? new Set<string>() : await getIngestedCufes(opts.companyId);
  // CUFEs a omitir por estar ya causados / en pantalla (dedup de "Causación + Caja").
  const skipSet = new Set((opts.skipCufes || []).map((c) => normCufe(c).toLowerCase()).filter(Boolean));

  try {
    for (const direction of directions) {
      if (opts.isCancelled?.()) break;

      const dirLabel = direction === "sent" ? "emitidos" : "recibidos";

      // Modo listado: el usuario ya subió el export de la DIAN; usamos esos CUFEs
      // y nos saltamos la generación del listado en el portal (el paso lento).
      let cufes: string[];
      if (opts.providedCufes) {
        cufes = opts.providedCufes;
        opts.onProgress?.({ step: `Listado recibido: ${cufes.length} CUFE(s) ${dirLabel}.`, current: 0, total: 0 });
      } else {
        opts.onProgress?.({ step: `Obteniendo listado de CUFEs (${dirLabel})...`, current: 0, total: 0 });
        const listing = await getCufeListing(
          opts.tokenUrl,
          opts.fechaInicio,
          opts.fechaFin,
          direction,
          opts.nitReceptor || "",
          (p) =>
            opts.onProgress?.({
              step: p.step || "Listando CUFEs...",
              current: p.current ?? 0,
              total: p.total ?? 0,
            })
        );
        cufes = listing.cufes;
      }

      listed += cufes.length;

      // Omitir: (a) los ya registrados en el historial de descargas (salvo
      // forceRedownload), y (b) los CUFEs en skipCufes (ya causados / ya en
      // pantalla). El criterio (b) es el correcto para "Causación + Caja": solo se
      // omiten los que están en "Facturas" (causados) o ya visibles; el resto se
      // mantiene/recupera al re-traer hasta que se contabilicen.
      let work = cufes.filter((c) => !knownCufes.has(c) && !skipSet.has(normCufe(c).toLowerCase()));
      alreadyRegistered += cufes.length - work.length;

      if (opts.maxDocuments && opts.maxDocuments > 0 && work.length > opts.maxDocuments) {
        work = work.slice(0, opts.maxDocuments);
      }

      if (work.length === 0) continue;

      const { success, failures, rounds } = await downloadDirectionWithRetries(opts, direction, work, tempDir);
      allSuccess.push(...success);
      allFailures.push(...failures);
      maxRounds = Math.max(maxRounds, rounds);
    }

    if (allSuccess.length === 0) {
      // Si todo lo listado ya estaba registrado, no es un error: simplemente no
      // hay nada nuevo que traer.
      if (alreadyRegistered > 0 && allFailures.length === 0) {
        return {
          items: [],
          stats: { listed, alreadyRegistered, downloaded: 0, failed: 0, rounds: maxRounds },
          failures: [],
        };
      }
      const detail = allFailures.length
        ? ` ${allFailures.length} CUFE(s) fallaron tras los reintentos.`
        : " No se encontraron CUFEs nuevos en el rango/grupo indicado.";
      throw new Error(`No se descargó ningún documento.${detail}`);
    }

    opts.onProgress?.({
      step: "Procesando XML para contabilización...",
      current: 0,
      total: allSuccess.length,
    });

    // Cada result.destPath es un ZIP por-documento (XML + PDF). processXmlBatch
    // los desempaca y empareja XML↔PDF internamente.
    const files: { name: string; buffer: Buffer }[] = [];
    for (const r of allSuccess) {
      if (!r.destPath || !fs.existsSync(r.destPath)) continue;
      try {
        files.push({ name: path.basename(r.destPath), buffer: fs.readFileSync(r.destPath) });
      } catch (err) {
        allFailures.push({
          cufe: r.cufe,
          error: `Descargado pero ilegible: ${err instanceof Error ? err.message : String(err)}`,
        });
      }
    }

    const items = await processXmlBatch(files);

    // Registrar los CUFEs efectivamente descargados para no re-descargarlos luego.
    await recordIngestedCufes(
      opts.companyId,
      allSuccess.map((r) => ({ cufe: r.cufe, docnum: r.docnum }))
    ).catch(() => undefined);

    return {
      items,
      stats: {
        listed,
        alreadyRegistered,
        downloaded: allSuccess.length,
        failed: allFailures.length,
        rounds: maxRounds,
      },
      failures: allFailures,
    };
  } finally {
    releaseDianJobSlot();
    try {
      fs.rmSync(tempDir, { recursive: true, force: true });
    } catch {
      /* limpieza best-effort */
    }
  }
}

const DIAN_BASE = "https://catalogo-vpfe.dian.gov.co";
const normCufe = (v: string) => (v || "").replace(/[^A-Fa-f0-9]/g, "").trim();

/**
 * Ingesta en UNA SOLA SESIÓN: abre el token DIAN una única vez, lista los
 * documentos del rango EN ESA MISMA sesión y descarga cada uno con sus cookies
 * (sin reabrir el token). Esto es OBLIGATORIO porque el token de la DIAN es de un
 * solo uso: cualquier reapertura en un navegador nuevo rebota a login. Pensada
 * para "Traer facturas nuevas del mes" (Recibidos, sin subir archivos).
 */
export async function ingestNewByDateRange(opts: {
  companyId: string;
  tokenUrl: string;
  fechaInicio: string;
  fechaFin: string;
  nitReceptor?: string;
  maxDocuments?: number;
  onProgress?: (p: ProgressData) => void;
  isCancelled?: () => boolean;
}): Promise<IngestResult> {
  const releaseDianJobSlot = await acquireDianJobSlot((pos) =>
    opts.onProgress?.({ step: `En cola para evitar el bloqueo de DIAN (turno ${pos})...`, current: 0, total: 0 })
  );
  const sessionId = uuidv4();
  const tempDir = path.join(DOWNLOADS_DIR, `caja-ingest-${sessionId}`);
  fs.mkdirSync(tempDir, { recursive: true });

  // El CRITERIO de "ya importada" es el CUFE (llave única del documento). La tabla
  // /Document/Received no muestra el CUFE, pero sí el trackId (1‑a‑1 con el doc):
  // lo usamos para SALTAR la re-descarga sin perder precisión. El CUFE real lo
  // leemos del XML al descargar y lo guardamos como identidad.
  const knownCufes = await getIngestedCufes(opts.companyId);
  const knownTrackIds = await getIngestedTrackIds(opts.companyId);
  const downloaded: { trackId: string; docnum: string; destPath: string; fileName: string }[] = [];
  const failures: IngestFailure[] = [];
  let alreadyRegistered = 0;
  const max = opts.maxDocuments && opts.maxDocuments > 0 ? opts.maxDocuments : Infinity;

  try {
    // Abre el token UNA vez y lista los documentos del rango scrapeando la TABLA
    // /Document/Received (sin depender de la generación de Export, que es lenta/
    // inestable). Devuelve cada documento con su trackId + las cookies de sesión.
    const result = await extractDocumentIds(
      opts.tokenUrl,
      opts.fechaInicio,
      opts.fechaFin,
      undefined,
      "received",
      true, // skipReconciliation: no usar la pestaña Export
    );
    const cookieHeader = Object.entries(result.cookies).map(([k, v]) => `${k}=${v}`).join("; ");
    const listed = result.documents.length;

    opts.onProgress?.({ step: `Listado: ${listed} documentos. Descargando…`, current: 0, total: listed });

    // Descarga cada documento con las cookies de la sesión (sin reabrir el token).
    for (const doc of result.documents) {
      if (opts.isCancelled?.()) break;
      if (downloaded.length >= max) break;
      const trackId = (doc.id || "").trim();
      // Ya importada: por CUFE (si la tabla lo trajera) o por trackId ya descargado.
      const tableCufe = normCufe(doc.cufe || "");
      const realTableCufe = tableCufe.length >= 40 ? tableCufe : "";
      if ((realTableCufe && knownCufes.has(realTableCufe)) || (trackId && knownTrackIds.has(trackId))) { alreadyRegistered++; continue; }
      const isEquiv = doc.docType?.toLowerCase().includes("equivalente") ?? false;
      const base = isEquiv
        ? `${DIAN_BASE}/Document/DownloadZipFilesEquivalente?trackId=`
        : `${DIAN_BASE}/Document/DownloadZipFiles?trackId=`;
      const safeNit = (doc.nit || "SinNIT").replace(/[^a-zA-Z0-9_-]/g, "_");
      const safeDoc = (doc.docnum || doc.id.slice(0, 12)).replace(/[^a-zA-Z0-9_-]/g, "_");
      const destPath = path.join(tempDir, `${safeNit} - ${safeDoc}.zip`);
      const fileName = path.basename(destPath);
      try {
        await fetchZipToFile(`${base}${doc.id}`, destPath, cookieHeader);
        downloaded.push({ trackId, docnum: doc.docnum, destPath, fileName });
        opts.onProgress?.({ step: `Descargando facturas… (${downloaded.length})`, current: downloaded.length, total: listed });
      } catch (err) {
        failures.push({ cufe: "", error: err instanceof Error ? err.message : String(err) });
      }
    }

    if (downloaded.length === 0) {
      // Nada nuevo (todo ya importado) NO es error.
      return { items: [], stats: { listed, alreadyRegistered, downloaded: 0, failed: failures.length, rounds: 0 }, failures };
    }

    opts.onProgress?.({ step: "Procesando XML para contabilización...", current: 0, total: downloaded.length });
    const files: { name: string; buffer: Buffer }[] = [];
    for (const d of downloaded) {
      if (!fs.existsSync(d.destPath)) continue;
      try { files.push({ name: d.fileName, buffer: fs.readFileSync(d.destPath) }); }
      catch (err) { failures.push({ cufe: "", error: `Descargado pero ilegible: ${err instanceof Error ? err.message : String(err)}` }); }
    }
    const items = await processXmlBatch(files);

    // Registrar por el CUFE REAL leído del XML (criterio único), junto al trackId
    // (para saltar la re-descarga) y el docnum. Se mapea item↔descarga por nombre.
    const cufeByFile = new Map<string, string>();
    for (const it of items) {
      const c = (it.xml as { cufe?: string } | undefined)?.cufe || "";
      if (it.fileName && c) cufeByFile.set(it.fileName, c);
    }
    const recs = downloaded.map((d) => ({
      cufe: cufeByFile.get(d.fileName) || "",
      docnum: d.docnum,
      trackId: d.trackId,
    }));
    await recordIngestedCufes(opts.companyId, recs).catch(() => undefined);

    return {
      items,
      stats: { listed, alreadyRegistered, downloaded: downloaded.length, failed: failures.length, rounds: 0 },
      failures,
    };
  } finally {
    releaseDianJobSlot();
    try { fs.rmSync(tempDir, { recursive: true, force: true }); } catch { /* best-effort */ }
  }
}
