import type { Page } from "puppeteer";
import {
  acquireDianJobSlot,
  getCufeListing,
  REAL_USER_AGENT,
  closeBrowserSafely,
} from "./dianScraper.js";
import { authenticateAndNavigate, applyReceivedDateFilter, fetchDocumentList, downloadXmlFile, type RecibidoDocument } from "./dianRecibidosScraper.js";
import { processXmlBatch, type BatchItem } from "./siigoAccountingService.js";
import { getIngestedCufes } from "./siigoIngestedCufesService.js";
import type { DocumentDirection, ProgressData } from "../types/dian.js";

const GRATIS_VPFE = "https://gratis-vpfe.dian.gov.co";

const delay = (ms: number) => new Promise<void>((r) => setTimeout(r, ms));
const fmtDdMmYyyy = (ms: number): string => {
  const dt = new Date(ms);
  return `${String(dt.getDate()).padStart(2, "0")}/${String(dt.getMonth() + 1).padStart(2, "0")}/${dt.getFullYear()}`;
};

// ── Helpers de fecha ISO (yyyy-mm-dd), sin librerías externas ──────────────
const parseISODate = (d: string): Date => {
  const [y, m, dd] = d.split("-").map(Number);
  return new Date(y, m - 1, dd);
};
const toISODate = (d: Date): string =>
  `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
const addDaysISO = (iso: string, n: number): string => {
  const d = parseISODate(iso);
  d.setDate(d.getDate() + n);
  return toISODate(d);
};
const midISODate = (fromISO: string, toISO: string): string => {
  const f = parseISODate(fromISO).getTime();
  const t = parseISODate(toISO).getTime();
  return toISODate(new Date(f + Math.floor((t - f) / 2)));
};

/**
 * `GetReceivedDocuments`/`GetSentDocuments` de gratis-vpfe NO pagina de verdad
 * más allá de 150 filas (pedir `start` mayor devuelve la MISMA primera página —
 * confirmado en pruebas: es un límite duro del servidor, no un bug de reporte de
 * totales). La única forma de traer más de 150 documentos en un rango es acotar
 * las fechas: si una consulta devuelve exactamente el tope, se parte el rango
 * en dos mitades y se re-consulta cada una (re-enviando el filtro en la MISMA
 * página/sesión, sin reabrir el token). Se combinan resultados por CUFE/id.
 */
async function fetchDocumentsSplitByRange(
  page: Page,
  fromISO: string,
  toISO: string,
  direction: DocumentDirection = "received"
): Promise<RecibidoDocument[]> {
  const LIMIT = 150; // límite real observado del servidor DIAN (gratis-vpfe)
  const merged = new Map<string, RecibidoDocument>();
  const queue: [string, string][] = [[fromISO, toISO]];
  let guard = 0;
  while (queue.length) {
    if (++guard > 200) {
      console.warn("[Siigo Ingest] Demasiadas particiones de rango — deteniendo por seguridad.");
      break;
    }
    const [f, t] = queue.shift()!;
    await applyReceivedDateFilter(
      page,
      fmtDdMmYyyy(parseISODate(f).getTime()),
      fmtDdMmYyyy(parseISODate(t).getTime())
    );
    const docs = await fetchDocumentList(page, direction === "sent" ? "sent" : "received");
    console.log(`[Siigo Ingest] Rango ${f}..${t}: ${docs.length} documento(s)`);
    if (docs.length >= LIMIT && f !== t) {
      // Posible truncado por el tope del servidor: partir el rango en dos.
      const mid = midISODate(f, t);
      const nextFrom = addDaysISO(mid, 1);
      queue.push([f, mid]);
      if (nextFrom <= t) queue.push([nextFrom, t]);
      continue;
    }
    if (docs.length >= LIMIT && f === t) {
      console.warn(`[Siigo Ingest] Día ${f} trae ${docs.length} documento(s) — puede estar truncado (no se puede partir por debajo de un día).`);
    }
    for (const d of docs) {
      const key = d.transactionId || `${d.docNumber}|${d.issueDate}`;
      if (key) merged.set(key, d);
    }
  }
  return [...merged.values()];
}

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
 * Descarga los CUFEs de una dirección exactamente igual que "Descarga Masiva +
 * Excel DIAN" (`routes/dianCufeDownload.ts` → `authenticateAndNavigate` +
 * fetch directo a gratis-vpfe): se autentica UNA vez (con un rango de fechas
 * amplio y fijo, solo para completar el paso de "Buscar" — el rango real de
 * las facturas es irrelevante aquí, porque `Document/DownloadXml` acepta el
 * CUFE directamente como `transactionId`, sin necesidad de ubicarlo en una
 * tabla filtrada por fecha). Tres pasadas: concurrente → cookies frescas →
 * re-autenticación completa, igual que la herramienta que ya funciona.
 */
async function downloadCufesGratisVpfe(
  tokenUrl: string,
  direction: DocumentDirection,
  cufes: string[],
  onProgress?: (p: ProgressData) => void,
  isCancelled?: () => boolean
): Promise<{ files: { name: string; buffer: Buffer }[]; failures: IngestFailure[] }> {
  const total = cufes.length;
  onProgress?.({ step: "Autenticando en portal DIAN...", current: 0, total });

  const { browser, page } = await authenticateAndNavigate(
    tokenUrl,
    "01/01/2024",
    fmtDdMmYyyy(Date.now()),
    (p) => onProgress?.({ step: p.step || "Autenticando...", current: 0, total }),
    direction === "sent" ? "sent" : "received"
  );

  const files: { name: string; buffer: Buffer }[] = [];
  const processed = new Set<string>();
  const cufeOriginalMap = new Map(cufes.map((c) => [c.toLowerCase(), c]));

  try {
    if (isCancelled?.()) return { files: [], failures: [] };

    const getCookieHeader = async () => {
      const c = await page.cookies();
      return c.map((ck) => `${ck.name}=${ck.value}`).join("; ");
    };

    const CONCURRENCY = Math.max(1, Math.min(Number(process.env.DIAN_DOWNLOAD_WORKERS || 4), 12));
    let dlSlots = CONCURRENCY;
    const dlWaitQueue: Array<() => void> = [];
    const acquireDl = (): Promise<void> =>
      dlSlots > 0 ? (dlSlots--, Promise.resolve()) : new Promise<void>((r) => dlWaitQueue.push(r));
    const releaseDl = () => { const n = dlWaitQueue.shift(); if (n) n(); else dlSlots++; };

    const RATE_MS = 60000 / 50;
    let nextAllowedMs = Date.now();
    const rateAcquire = async () => {
      const wait = nextAllowedMs - Date.now();
      nextAllowedMs = Math.max(nextAllowedMs, Date.now()) + RATE_MS;
      if (wait > 0) await delay(wait);
    };

    let dlOk = 0;
    let cookieHeader = await getCookieHeader();

    const isHtml = (buf: Buffer) => {
      const preview = buf.toString("utf8", 0, 300).trim().toLowerCase();
      return preview.startsWith("<!doctype html") || preview.startsWith("<html") || preview.includes("<title>");
    };

    const fetchOne = async (cufe: string, hdr: string): Promise<void> => {
      const xmlResp = await fetch(`${GRATIS_VPFE}/Document/DownloadXml?transactionId=${cufe}&type=2`, {
        headers: { "User-Agent": REAL_USER_AGENT, Cookie: hdr },
      });
      if (!xmlResp.ok) throw new Error(`HTTP ${xmlResp.status}`);
      const xmlBuf = Buffer.from(await xmlResp.arrayBuffer());
      if (isHtml(xmlBuf)) throw new Error("GRATIS_VPFE_SESSION: sesión inválida o CUFE no encontrado");

      const originalCufe = cufeOriginalMap.get(cufe.toLowerCase()) || cufe;
      const safeName = originalCufe.replace(/[^a-zA-Z0-9_-]/g, "_").slice(0, 44);
      files.push({ name: `${safeName}.xml`, buffer: xmlBuf });

      // PDF: mejor esfuerzo, no bloquea el éxito del XML.
      try {
        const pdfResp = await fetch(`${GRATIS_VPFE}/IoFacturo/Print/PrintStoragePdf?transactionId=${cufe}&viewMode=attachment`, {
          headers: { "User-Agent": REAL_USER_AGENT, Cookie: hdr },
        });
        if (pdfResp.ok) {
          const pdfBuf = Buffer.from(await pdfResp.arrayBuffer());
          if (pdfBuf[0] === 0x25 && pdfBuf[1] === 0x50) files.push({ name: `${safeName}.pdf`, buffer: pdfBuf });
        }
      } catch { /* PDF opcional */ }

      processed.add(cufe.toLowerCase());
      dlOk++;
      onProgress?.({ step: `Descargando documentos: ${dlOk}/${total}…`, current: dlOk, total });
    };

    // Pasada 1: concurrente con las cookies de la sesión inicial.
    onProgress?.({ step: `Descargando ${total} documento(s)…`, current: 0, total });
    await Promise.all(cufes.map(async (cufe) => {
      await acquireDl();
      await rateAcquire();
      try { await fetchOne(cufe, cookieHeader); }
      catch (err) { console.warn(`[Siigo Ingest] P1 error ${cufe.slice(0, 16)}:`, err instanceof Error ? err.message : err); }
      finally { releaseDl(); }
    }));
    if (isCancelled?.()) {
      const failures: IngestFailure[] = cufes.filter((c) => !processed.has(c.toLowerCase())).map((c) => ({ cufe: c, error: "Cancelado" }));
      return { files, failures };
    }

    // Pasada 2: cookies frescas, secuencial.
    let missing = cufes.filter((c) => !processed.has(c.toLowerCase()));
    if (missing.length > 0) {
      onProgress?.({ step: `Recuperando ${missing.length} documento(s)…`, current: dlOk, total });
      cookieHeader = await getCookieHeader();
      for (const cufe of missing) {
        if (isCancelled?.()) break;
        await rateAcquire();
        try { await fetchOne(cufe, cookieHeader); }
        catch (err) { console.warn(`[Siigo Ingest] P2 error ${cufe.slice(0, 16)}:`, err instanceof Error ? err.message : err); }
      }
    }

    // Pasada 3: re-autenticar (sesión nueva) y reintentar los que queden.
    missing = cufes.filter((c) => !processed.has(c.toLowerCase()));
    if (missing.length > 0 && !isCancelled?.()) {
      onProgress?.({ step: `Recuperando ${missing.length} documento(s) (re-autenticando)…`, current: dlOk, total });
      const { browser: browser2, page: page2 } = await authenticateAndNavigate(
        tokenUrl, "01/01/2024", fmtDdMmYyyy(Date.now()), () => {}, direction === "sent" ? "sent" : "received"
      );
      try {
        const cookies2 = await page2.cookies();
        const hdr2 = cookies2.map((c) => `${c.name}=${c.value}`).join("; ");
        for (const cufe of missing) {
          if (isCancelled?.()) break;
          await rateAcquire();
          try { await fetchOne(cufe, hdr2); } catch { /* último intento fallido */ }
        }
      } finally {
        await closeBrowserSafely(browser2).catch(() => {});
      }
    }

    const failures: IngestFailure[] = cufes
      .filter((c) => !processed.has(c.toLowerCase()))
      .map((c) => ({ cufe: cufeOriginalMap.get(c.toLowerCase()) || c, error: "No se pudo descargar tras 3 intentos" }));

    return { files, failures };
  } finally {
    await closeBrowserSafely(browser).catch(() => {});
  }
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

  const allFiles: { name: string; buffer: Buffer }[] = [];
  const succeededCufes: string[] = [];
  const allFailures: IngestFailure[] = [];
  let listed = 0;
  let alreadyRegistered = 0;

  // Registro persistente = SOLO bloqueo explícito del usuario (basurita → "no
  // volver a traer"), NUNCA se auto-puebla por una descarga exitosa. Así la
  // herramienta siempre trae lo que no está bloqueado ni en pantalla, aunque ya
  // se haya traído antes. Si forceRedownload, se ignora incluso el bloqueo.
  const knownCufes = opts.forceRedownload ? new Set<string>() : await getIngestedCufes(opts.companyId);
  // CUFEs a omitir por estar ya causados / en pantalla (dedup contra la tabla de
  // trabajo actual que envía el frontend).
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

      // Omitir: (a) CUFEs bloqueados explícitamente por el usuario, y (b) los que
      // ya están en la tabla de trabajo actual (skipCufes). El resto se trae
      // SIEMPRE, aunque ya se haya descargado antes en una corrida anterior.
      let work = cufes.filter((c) => !knownCufes.has(c) && !skipSet.has(normCufe(c).toLowerCase()));
      alreadyRegistered += cufes.length - work.length;

      if (opts.maxDocuments && opts.maxDocuments > 0 && work.length > opts.maxDocuments) {
        work = work.slice(0, opts.maxDocuments);
      }

      if (work.length === 0) continue;

      const { files, failures } = await downloadCufesGratisVpfe(
        opts.tokenUrl, direction, work, opts.onProgress, opts.isCancelled
      );
      allFiles.push(...files);
      allFailures.push(...failures);
      const failedSet = new Set(failures.map((f) => normCufe(f.cufe).toLowerCase()));
      succeededCufes.push(...work.filter((c) => !failedSet.has(normCufe(c).toLowerCase())));
    }

    if (allFiles.length === 0) {
      // Si todo lo listado ya estaba registrado, no es un error: simplemente no
      // hay nada nuevo que traer.
      if (alreadyRegistered > 0 && allFailures.length === 0) {
        return {
          items: [],
          stats: { listed, alreadyRegistered, downloaded: 0, failed: 0, rounds: 0 },
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
      total: allFiles.length,
    });

    const items = await processXmlBatch(allFiles);

    // NO se registra automáticamente en `siigoIngestedCufes`: ese registro ahora
    // es exclusivamente la lista de bloqueo explícito del usuario (basurita →
    // "no volver a traer"), para que la herramienta siempre traiga lo que no está
    // en pantalla aunque ya se haya traído antes. Ver `dianBlockCufes`.

    return {
      items,
      stats: {
        listed,
        alreadyRegistered,
        downloaded: succeededCufes.length,
        failed: allFailures.length,
        rounds: 0,
      },
      failures: allFailures,
    };
  } finally {
    releaseDianJobSlot();
  }
}

const normCufe = (v: string) => (v || "").replace(/[^A-Fa-f0-9]/g, "").trim();

/**
 * Ingesta por MES, sin listado manual: abre el token DIAN una única vez con el
 * rango del mes, lista TODOS los documentos recibidos en ese rango directamente
 * del portal (`fetchDocumentList`, la misma llamada AJAX que usa "Descarga Masiva
 * + Excel DIAN"), filtra los que ya se trajeron antes (registro persistente por
 * CUFE) y descarga solo los nuevos por CUFE directo a gratis-vpfe — el mismo
 * mecanismo que `downloadCufesGratisVpfe`. Pensada para el selector de mes de
 * "Contabilizar Gastos Siigo" (sin subir archivos).
 */
export async function ingestNewByDateRange(opts: {
  companyId: string;
  tokenUrl: string;
  fechaInicio: string; // ISO yyyy-mm-dd
  fechaFin: string;    // ISO yyyy-mm-dd
  /** @deprecated no usado internamente; se mantiene por compatibilidad con llamadores existentes (cajaErp.ts). */
  nitReceptor?: string;
  maxDocuments?: number;
  forceRedownload?: boolean;
  /** CUFEs a omitir además del registro persistente (ya causados / ya en pantalla). */
  skipCufes?: string[];
  onProgress?: (p: ProgressData) => void;
  isCancelled?: () => boolean;
}): Promise<IngestResult> {
  const releaseDianJobSlot = await acquireDianJobSlot((pos) =>
    opts.onProgress?.({ step: `En cola para evitar el bloqueo de DIAN (turno ${pos})...`, current: 0, total: 0 })
  );

  const knownCufes = opts.forceRedownload ? new Set<string>() : await getIngestedCufes(opts.companyId);
  const skipSet = new Set((opts.skipCufes || []).map((c) => normCufe(c).toLowerCase()).filter(Boolean));
  const failures: IngestFailure[] = [];

  try {
    const from = fmtDdMmYyyy(new Date(`${opts.fechaInicio}T00:00:00`).getTime());
    const to = fmtDdMmYyyy(new Date(`${opts.fechaFin}T00:00:00`).getTime());

    opts.onProgress?.({ step: "Autenticando en portal DIAN...", current: 0, total: 0 });
    const { browser, page } = await authenticateAndNavigate(opts.tokenUrl, from, to, (p) =>
      opts.onProgress?.({ step: p.step || "Autenticando...", current: 0, total: 0 }), "received");

    try {
      if (opts.isCancelled?.()) {
        return { items: [], stats: { listed: 0, alreadyRegistered: 0, downloaded: 0, failed: 0, rounds: 0 }, failures: [] };
      }

      opts.onProgress?.({ step: "Listando documentos del mes...", current: 0, total: 0 });
      const docs = await fetchDocumentsSplitByRange(page, opts.fechaInicio, opts.fechaFin, "received");
      const listed = docs.length;
      console.log(`[Siigo Ingest] Listado ${opts.fechaInicio}..${opts.fechaFin}: ${listed} documento(s) en total`);

      let work = docs.filter((d) => {
        const cufe = normCufe(d.transactionId || "");
        return cufe && !knownCufes.has(cufe) && !skipSet.has(cufe.toLowerCase());
      });
      const alreadyRegistered = listed - work.length;

      if (opts.maxDocuments && opts.maxDocuments > 0 && work.length > opts.maxDocuments) {
        work = work.slice(0, opts.maxDocuments);
      }

      if (work.length === 0) {
        return { items: [], stats: { listed, alreadyRegistered, downloaded: 0, failed: 0, rounds: 0 }, failures: [] };
      }

      // Descarga concurrente con rate limit, sobre la MISMA sesión (sin reabrir el token).
      const files: { name: string; buffer: Buffer }[] = [];
      const succeededCufes: string[] = [];
      let dlOk = 0;
      const total = work.length;
      opts.onProgress?.({ step: `Descargando ${total} documento(s)…`, current: 0, total });

      const CONCURRENCY = Math.max(1, Math.min(Number(process.env.DIAN_DOWNLOAD_WORKERS || 4), 12));
      let dlSlots = CONCURRENCY;
      const dlWaitQueue: Array<() => void> = [];
      const acquireDl = (): Promise<void> =>
        dlSlots > 0 ? (dlSlots--, Promise.resolve()) : new Promise<void>((r) => dlWaitQueue.push(r));
      const releaseDl = () => { const n = dlWaitQueue.shift(); if (n) n(); else dlSlots++; };

      const RATE_MS = 60000 / 50;
      let nextAllowedMs = Date.now();
      const rateAcquire = async () => {
        const wait = nextAllowedMs - Date.now();
        nextAllowedMs = Math.max(nextAllowedMs, Date.now()) + RATE_MS;
        if (wait > 0) await delay(wait);
      };

      await Promise.all(work.map(async (doc) => {
        await acquireDl();
        await rateAcquire();
        try {
          if (opts.isCancelled?.()) return;
          const downloaded = await downloadXmlFile(page, doc.transactionId, doc.docNumber, true);
          for (const f of downloaded) files.push(f);
          succeededCufes.push(normCufe(doc.transactionId));
          dlOk++;
          opts.onProgress?.({ step: `Descargando documentos: ${dlOk}/${total}…`, current: dlOk, total });
        } catch (err) {
          failures.push({ cufe: doc.transactionId, error: err instanceof Error ? err.message : String(err) });
        } finally {
          releaseDl();
        }
      }));

      if (files.length === 0) {
        return { items: [], stats: { listed, alreadyRegistered, downloaded: 0, failed: failures.length, rounds: 0 }, failures };
      }

      opts.onProgress?.({ step: "Procesando XML para contabilización...", current: 0, total: files.length });
      const items = await processXmlBatch(files);

      // NO se registra automáticamente (ver nota en `ingestFromDian`): el registro
      // persistente ahora es solo la lista de bloqueo explícito del usuario.

      return {
        items,
        stats: { listed, alreadyRegistered, downloaded: succeededCufes.length, failed: failures.length, rounds: 0 },
        failures,
      };
    } finally {
      await closeBrowserSafely(browser).catch(() => {});
    }
  } finally {
    releaseDianJobSlot();
  }
}
