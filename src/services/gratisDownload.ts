/**
 * Descarga directa de documentos DIAN por CUFE desde gratis-vpfe
 * (transactionId = CUFE), sin buscar cada factura en el portal.
 *
 * Es el mismo mecanismo que usa "Descarga masiva + Excel" (dianCufeDownload) en
 * producción; aquí se encapsula con reintentos, renovación de sesión y el
 * semáforo global de descargas para reutilizarlo desde el exportador de Excel.
 */
import JSZip from "jszip";
import { acquireDownloadSlot, REAL_USER_AGENT } from "./dianScraper.js";

const GRATIS_BASE = "https://gratis-vpfe.dian.gov.co";
/** Mínimo entre re-logins de la sesión gratis-vpfe (cada uno abre un navegador). */
const GRATIS_REFRESH_MIN_MS = 20_000;

export const gratisXmlUrl = (cufe: string): string => `${GRATIS_BASE}/Document/DownloadXml?transactionId=${cufe}&type=2`;
export const gratisPdfUrl = (cufe: string): string => `${GRATIS_BASE}/IoFacturo/Print/PrintStoragePdf?transactionId=${cufe}&viewMode=attachment`;

/** Sesión gratis-vpfe compartida por todas las descargas del job; se renueva sola si caduca. */
export interface GratisSession {
  header: string;
  /** Re-entra al token en un navegador efímero (una sola vez a la vez y con separación mínima, salvo `force`). */
  refresh(force?: boolean): Promise<void>;
}

export function createGratisSession(
  initialHeader: string,
  fetchFreshHeader: () => Promise<string>,
  minRefreshMs: number = GRATIS_REFRESH_MIN_MS,
): GratisSession {
  let refreshing: Promise<void> | null = null;
  let lastRefresh = Date.now();
  const session: GratisSession = {
    header: initialHeader,
    refresh: async (force = false) => {
      if (refreshing) return refreshing;
      if (!force && Date.now() - lastRefresh < minRefreshMs) return; // ya se renovó hace poco: basta reintentar
      refreshing = (async () => {
        try {
          session.header = await fetchFreshHeader();
          lastRefresh = Date.now();
        } finally {
          refreshing = null;
        }
      })();
      return refreshing;
    },
  };
  return session;
}

/**
 * GET a gratis-vpfe con el semáforo global de descargas, reintentos cortos y
 * renovación de la sesión si la DIAN responde con HTML/login/401/403. Valida que
 * el cuerpo sea realmente un XML o un PDF.
 */
export async function gratisFetch(url: string, session: GratisSession, kind: "xml" | "pdf", maxAttempts = 6): Promise<Buffer> {
  let lastError = "sin respuesta";
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    let needRefresh = false;
    let permanent = false;
    const release = await acquireDownloadSlot();
    try {
      const resp = await fetch(url, {
        headers: { "User-Agent": REAL_USER_AGENT, Cookie: session.header, "Accept-Language": "es-CO,es;q=0.9" },
        signal: AbortSignal.timeout(60_000),
      });
      if (resp.status === 401 || resp.status === 403) {
        needRefresh = true;
        lastError = `GRATIS_VPFE_HTTP_${resp.status}`;
      } else if (resp.status === 429 || resp.status >= 500) {
        lastError = `GRATIS_VPFE_HTTP_${resp.status}`;
      } else if (!resp.ok) {
        lastError = `GRATIS_VPFE_HTTP_${resp.status}`;
        permanent = true; // 404/400: ese CUFE no existe para esta sesión; reintentar no cambia nada
      } else {
        const buf = Buffer.from(await resp.arrayBuffer());
        if (kind === "pdf") {
          if (buf.length > 4 && buf[0] === 0x25 && buf[1] === 0x50) return buf; // %PDF
          lastError = "GRATIS_VPFE_UNEXPECTED: no es un PDF";
        } else {
          const head = buf.toString("utf8", 0, 300).trim().toLowerCase();
          if (head.startsWith("<!doctype html") || head.startsWith("<html") || head.includes("<title>")) {
            needRefresh = true;
            lastError = "GRATIS_VPFE_SESSION: gratis-vpfe devolvió HTML (sesión inválida o CUFE no encontrado)";
          } else if (head.startsWith("<")) {
            return buf;
          } else {
            lastError = `GRATIS_VPFE_UNEXPECTED: respuesta inesperada (${buf.length}b)`;
          }
        }
      }
    } catch (err) {
      lastError = err instanceof Error ? err.message : String(err);
    } finally {
      release(); // no retener el cupo global mientras se espera el reintento
    }
    if (permanent || attempt >= maxAttempts) break;
    if (needRefresh) {
      try { await session.refresh(); }
      catch (refreshErr) { lastError = refreshErr instanceof Error ? refreshErr.message : String(refreshErr); }
    }
    await new Promise((r) => setTimeout(r, 400 * attempt + Math.floor(Math.random() * 300)));
  }
  throw new Error(`${lastError} (${kind} ${url.match(/transactionId=([0-9a-f]{0,16})/i)?.[1] ?? ""}...)`);
}

/**
 * Baja XML (+ PDF si aplica) por CUFE y los empaqueta en un ZIP mínimo para que
 * `processDocumentFromZip` los procese igual que un ZIP de la DIAN. El PDF es de
 * mejor esfuerzo: sin él el Excel sale igual (solo se usa para el ZIP de entrega
 * y Drive); los Documentos Soporte no tienen representación gráfica.
 */
export async function downloadFromGratisAsZip(cufe: string, session: GratisSession): Promise<Buffer> {
  const xmlBuf = await gratisFetch(gratisXmlUrl(cufe), session, "xml");
  let pdfBuf: Buffer | null = null;
  const isSupportDoc = /<(?:\w+:)?ProfileID[^>]*>[^<]*soporte/i.test(xmlBuf.toString("utf8"));
  if (!isSupportDoc) {
    try { pdfBuf = await gratisFetch(gratisPdfUrl(cufe), session, "pdf", 3); } catch { pdfBuf = null; }
  }
  const zip = new JSZip();
  zip.file(`${cufe.slice(0, 16)}.xml`, xmlBuf);
  if (pdfBuf) zip.file(`${cufe.slice(0, 16)}.pdf`, pdfBuf);
  return Buffer.from(await zip.generateAsync({ type: "nodebuffer" }));
}

