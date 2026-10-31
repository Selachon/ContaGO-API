/**
 * Núcleo de descarga de "Descarga Masiva DIAN" (`dian-recibidos`), extraído
 * para que otros consumidores (el ERP) reutilicen EXACTAMENTE el mismo
 * mecanismo — autenticación + descarga directa por CUFE contra
 * gratis-vpfe.dian.gov.co, 4 descargas concurrentes limitadas a 50/min, y
 * las mismas 3 pasadas de recuperación — sin reimplementar nada.
 *
 * `src/routes/dianRecibidosDownload.ts` usa esta misma función para su ZIP;
 * lo único que varía entre consumidores es qué se hace con los buffers
 * descargados (empaquetar en ZIP vs. crear documentos del ERP).
 */
import { closeBrowserSafely, REAL_USER_AGENT } from "./dianScraper.js";
import { authenticateAndNavigate } from "./dianRecibidosScraper.js";
import { extractInvoiceDataFromXml } from "./xmlParser.js";
import type { DocumentDirection } from "../types/dian.js";

export interface DianRecibidosItem {
  cufe: string;
  docName: string;
  /** Fecha de emisión (ISO o dd/mm/aaaa) para ubicar el documento por año/mes; "" si no se pudo leer. */
  issueDate: string;
  xmlBuffer: Buffer;
  pdfBuffer: Buffer | null;
}

export interface DianRecibidosProgress {
  step: string;
  current: number;
  total: number;
  pct?: number;
}

export interface DianRecibidosBatchOptions {
  tokenUrl: string;
  cufes: string[];
  direction: DocumentDirection;
  onProgress?: (p: DianRecibidosProgress) => void;
  isCancelled?: () => boolean;
}

export interface DianRecibidosBatchResult {
  items: DianRecibidosItem[];
  downloaded: number;
  total: number;
}

function fmtDdMmYyyy(ms: number): string {
  const dt = new Date(ms);
  return `${String(dt.getDate()).padStart(2, "0")}/${String(dt.getMonth() + 1).padStart(2, "0")}/${dt.getFullYear()}`;
}

export async function runDianRecibidosBatchDownload(
  opts: DianRecibidosBatchOptions
): Promise<DianRecibidosBatchResult> {
  const { tokenUrl, cufes: downloadCufes, direction } = opts;
  const onProgress = opts.onProgress || (() => {});
  const isCancelled = opts.isCancelled || (() => false);

  const items: DianRecibidosItem[] = [];
  const succeededCufes = new Set<string>();

  onProgress({ step: "Autenticando en portal DIAN...", current: 0, total: downloadCufes.length });

  const { browser: gBrowser, page: gPage } = await authenticateAndNavigate(
    tokenUrl, "01/01/2024", fmtDdMmYyyy(Date.now()),
    (p) => onProgress({ step: p.step || "Autenticando...", current: 0, total: downloadCufes.length }),
    direction,
  );

  try {
    if (isCancelled()) return { items, downloaded: 0, total: downloadCufes.length };

    const getCookieHeader = async () => {
      const c = await gPage.cookies();
      return c.map((c: any) => `${c.name}=${c.value}`).join("; ");
    };

    const CONCURRENCY = 4;
    let dlSlots = CONCURRENCY;
    const dlWaitQueue: Array<() => void> = [];
    const acquireDl = (): Promise<void> =>
      dlSlots > 0 ? (dlSlots--, Promise.resolve()) : new Promise<void>((r) => dlWaitQueue.push(r));
    const releaseDl = () => { const n = dlWaitQueue.shift(); if (n) n(); else dlSlots++; };

    const RATE_MS = 60000 / 50;
    let nextAllowedMs = Date.now();
    const rateAcquire = async (): Promise<void> => {
      const wait = nextAllowedMs - Date.now();
      nextAllowedMs = Math.max(nextAllowedMs, Date.now()) + RATE_MS;
      if (wait > 0) await new Promise<void>((r) => setTimeout(r, wait));
    };

    let dlOk = 0;
    const dlStartMs = Date.now();
    let cookieHeader = await getCookieHeader();

    const processDoc = async (xmlBuf: Buffer, cufe: string): Promise<void> => {
      let docName: string;
      let issueDate = "";
      try {
        const inv = await extractInvoiceDataFromXml(xmlBuf, { id: cufe, docnum: "" });
        docName = (inv.docNumber || cufe.slice(0, 16)).replace(/[^a-zA-Z0-9_\-]/g, "_");
        issueDate = inv.issueDateISO || inv.issueDate || "";
      } catch {
        docName = cufe.slice(0, 16);
      }

      let pdfBuf: Buffer | null = null;
      try {
        const pdfResp = await fetch(
          `https://gratis-vpfe.dian.gov.co/IoFacturo/Print/PrintStoragePdf?transactionId=${cufe}&viewMode=attachment`,
          { headers: { "User-Agent": REAL_USER_AGENT, Cookie: cookieHeader } },
        );
        if (pdfResp.ok) {
          const buf = Buffer.from(await pdfResp.arrayBuffer());
          if (buf[0] === 0x25 && buf[1] === 0x50) pdfBuf = buf;
        }
      } catch {}

      items.push({ cufe, docName, issueDate, xmlBuffer: xmlBuf, pdfBuffer: pdfBuf });
      succeededCufes.add(cufe.toLowerCase());
      dlOk++;
      const elapsedMin = (Date.now() - dlStartMs) / 60000;
      const fpm = elapsedMin > 0 ? Math.round(dlOk / elapsedMin) : 0;
      const pct = Math.round((dlOk / downloadCufes.length) * 100);
      onProgress({
        step: `Descargando ${dlOk}/${downloadCufes.length} · ${fpm} fact/min`,
        current: dlOk, total: downloadCufes.length, pct,
      });
    };

    // ── Pasada 1: descarga directa por CUFE con concurrencia ──────────────
    onProgress({ step: `Descargando ${downloadCufes.length} documentos...`, current: 0, total: downloadCufes.length, pct: 0 });

    const p1Promises = downloadCufes.map(async (cufe) => {
      await acquireDl();
      await rateAcquire();
      try {
        const xmlResp = await fetch(
          `https://gratis-vpfe.dian.gov.co/Document/DownloadXml?transactionId=${cufe}&type=2`,
          { headers: { "User-Agent": REAL_USER_AGENT, Cookie: cookieHeader } },
        );
        if (!xmlResp.ok) return;
        await processDoc(Buffer.from(await xmlResp.arrayBuffer()), cufe);
      } catch (err) {
        console.warn(`[Recibidos] P1 error ${cufe.slice(0, 16)}:`, err instanceof Error ? err.message : err);
      } finally {
        releaseDl();
      }
    });
    await Promise.all(p1Promises);
    if (isCancelled()) return { items, downloaded: dlOk, total: downloadCufes.length };

    // ── Pasada 2: cookies frescas ─────────────────────────────────────────
    let missing = downloadCufes.filter((c) => !succeededCufes.has(c.toLowerCase()));
    if (missing.length > 0) {
      console.log(`[Recibidos] Pasada 2: ${missing.length} faltantes`);
      onProgress({ step: `Recuperando ${missing.length} faltantes...`, current: dlOk, total: downloadCufes.length });
      cookieHeader = await getCookieHeader();
      for (const cufe of missing) {
        if (isCancelled()) break;
        await rateAcquire();
        try {
          const xmlResp = await fetch(
            `https://gratis-vpfe.dian.gov.co/Document/DownloadXml?transactionId=${cufe}&type=2`,
            { headers: { "User-Agent": REAL_USER_AGENT, Cookie: cookieHeader } },
          );
          if (xmlResp.ok) await processDoc(Buffer.from(await xmlResp.arrayBuffer()), cufe);
        } catch {}
      }
    }

    // ── Pasada 3: re-autenticar y reintentar ──────────────────────────────
    missing = downloadCufes.filter((c) => !succeededCufes.has(c.toLowerCase()));
    if (missing.length > 0 && !isCancelled()) {
      console.log(`[Recibidos] Pasada 3: ${missing.length} aún faltantes — re-autenticando`);
      onProgress({ step: `Recuperando ${missing.length} faltantes (re-autenticando)...`, current: dlOk, total: downloadCufes.length });
      const { browser: gBrowser2, page: gPage2 } = await authenticateAndNavigate(
        tokenUrl, "01/01/2024", fmtDdMmYyyy(Date.now()), () => {}, direction,
      );
      try {
        const cookies2 = await gPage2.cookies();
        const hdr2 = cookies2.map((c: any) => `${c.name}=${c.value}`).join("; ");
        for (const cufe of missing) {
          if (isCancelled()) break;
          await rateAcquire();
          try {
            const xmlResp = await fetch(
              `https://gratis-vpfe.dian.gov.co/Document/DownloadXml?transactionId=${cufe}&type=2`,
              { headers: { "User-Agent": REAL_USER_AGENT, Cookie: hdr2 } },
            );
            if (xmlResp.ok) await processDoc(Buffer.from(await xmlResp.arrayBuffer()), cufe);
          } catch {}
        }
      } finally {
        await closeBrowserSafely(gBrowser2).catch(() => {});
      }
    }

    console.log(`[Recibidos] Final: ${dlOk}/${downloadCufes.length} descargados`);
    return { items, downloaded: dlOk, total: downloadCufes.length };
  } finally {
    await closeBrowserSafely(gBrowser).catch(() => {});
  }
}
