import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";
import { v4 as uuidv4 } from "uuid";
import {
  getCufeListing,
  downloadDocumentsByCufe,
  type CufeDownloadItem,
} from "./dianScraper.js";
import { processXmlBatch, type BatchItem } from "./siigoAccountingService.js";
import { getIngestedCufes, recordIngestedCufes } from "./siigoIngestedCufesService.js";
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
  fechaInicio: string;
  fechaFin: string;
  grupo: IngestGrupo;
  nitReceptor?: string;
  /** Tope de documentos (p. ej. límite demo o DIAN_MAX_DOCUMENTS). */
  maxDocuments?: number;
  /** Rondas de reintento por CUFE fallido, además de la pasada inicial. Def: 3. */
  retryRounds?: number;
  /** Si true, re-descarga también los CUFEs ya registrados (ignora el registro). */
  forceRedownload?: boolean;
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
  const retryRounds = Math.max(0, opts.retryRounds ?? 3);
  const successByCufe = new Map<string, CufeDownloadItem>();
  let lastError = new Map<string, string>();
  let pending = [...cufes];
  let round = 0;

  while (pending.length > 0 && round <= retryRounds) {
    if (opts.isCancelled?.()) break;

    const label = round === 0 ? "Descargando documentos" : `Reintento ${round}/${retryRounds}`;
    const batchTotal = pending.length;

    const { results } = await downloadDocumentsByCufe(
      opts.tokenUrl,
      pending,
      opts.fechaInicio,
      opts.fechaFin,
      direction,
      tempDir,
      (p) =>
        opts.onProgress?.({
          step: `${label} (${direction === "sent" ? "emitidos" : "recibidos"})...`,
          current: p.current ?? 0,
          total: p.total ?? batchTotal,
        }),
      opts.isCancelled
    );

    const stillPending: string[] = [];
    lastError = new Map();
    for (const r of results) {
      if (r.success && r.destPath) {
        successByCufe.set(r.cufe, r);
      } else {
        stillPending.push(r.cufe);
        lastError.set(r.cufe, r.error || "No se pudo descargar el documento");
      }
    }

    pending = stillPending;
    if (pending.length === 0) break;

    round++;
    if (round <= retryRounds && pending.length > 0) {
      // Backoff incremental entre rondas para dar respiro al portal.
      await delay(2000 * round);
    }
  }

  const failures: IngestFailure[] = pending.map((cufe) => ({
    cufe,
    error: lastError.get(cufe) || "No se pudo descargar tras agotar reintentos",
  }));

  return { success: [...successByCufe.values()], failures, rounds: Math.min(round, retryRounds) };
}

/**
 * Ingesta automática para "Contabilizar Gastos Siigo": obtiene el listado de
 * CUFEs del portal DIAN, descarga cada documento (con reintentos por CUFE) y los
 * parsea con `processXmlBatch`, devolviendo los mismos ítems que la subida manual
 * de ZIP. DEBE invocarse dentro del contexto de empresa (runWithSiigoCompany)
 * porque `processXmlBatch` consulta Siigo para detectar facturas ya causadas.
 */
export async function ingestFromDian(opts: IngestOptions): Promise<IngestResult> {
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

  try {
    for (const direction of directions) {
      if (opts.isCancelled?.()) break;

      const dirLabel = direction === "sent" ? "emitidos" : "recibidos";
      opts.onProgress?.({ step: `Obteniendo listado de CUFEs (${dirLabel})...`, current: 0, total: 0 });

      const { cufes } = await getCufeListing(
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

      listed += cufes.length;

      // Omitir los ya registrados (descargados en consultas previas).
      let work = cufes.filter((c) => !knownCufes.has(c));
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
    try {
      fs.rmSync(tempDir, { recursive: true, force: true });
    } catch {
      /* limpieza best-effort */
    }
  }
}
