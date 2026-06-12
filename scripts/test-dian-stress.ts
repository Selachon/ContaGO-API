/**
 * Harness de estrés para el pipeline DIAN (listado → descarga → parseo XML).
 * Replica el flujo de dianCufeDownload (búsqueda serial + descargas con
 * contrapresión) y reporta completitud por documento para aislar dónde se
 * pierden facturas: listado, resolución, descarga o parseo.
 *
 * Uso:
 *   npx tsx scripts/test-dian-stress.ts <tokenUrl> <start> <end> \
 *     [--direction received|sent] [--limit N] [--workers W] [--label L] [--offset K]
 */
import fs from "fs";
import JSZip from "jszip";
import { extractDocumentIdsByCufe, throttledDianDownload } from "../src/services/dianScraper.js";
import { extractInvoiceDataFromXml } from "../src/services/xmlParser.js";
import type { InvoiceData } from "../src/types/dianExcel.js";

const args = process.argv.slice(2);
const positional = args.filter((a) => !a.startsWith("--"));
const flag = (name: string, def: string): string => {
  const i = args.indexOf(`--${name}`);
  return i >= 0 && args[i + 1] ? args[i + 1] : def;
};

const tokenUrl = positional[0];
const startDate = positional[1] || "2026-05-01";
const endDate = positional[2] || "2026-05-31";
const direction = (flag("direction", "received") as "received" | "sent");
const limit = Number(flag("limit", "0")); // 0 = todos
const offset = Number(flag("offset", "0"));
const workers = Number(flag("workers", process.env.DIAN_DOWNLOAD_WORKERS || "5"));
const label = flag("label", "test");
// Si se pasa --save-rows, guarda cada factura como fila InvoiceData completa (JSONL),
// el mismo shape que construye dianExcel — permite generar el Excel de la herramienta
// sin volver a descargar.
const saveRowsPath = flag("save-rows", "");
const rowsWriter = saveRowsPath ? fs.createWriteStream(saveRowsPath, { flags: "w", encoding: "utf8" }) : null;

if (!tokenUrl) {
  console.error("Falta tokenUrl");
  process.exit(1);
}

const log = (msg: string) => console.log(`[${label} ${new Date().toISOString().slice(11, 19)}] ${msg}`);

interface DocResult {
  cufe: string;
  docnum: string;
  nit: string;
  dlOk: boolean;
  dlAttemptsNote?: string;
  hasData: boolean;
  total?: number;
  error?: string;
  ms?: number;
}

async function main(): Promise<void> {
  const t0 = Date.now();
  const results = new Map<string, DocResult>();
  let seen = 0;
  let processed = 0;

  // Semáforo de contrapresión, igual que MAX_DL en dianCufeDownload
  let slots = workers;
  const waiters: Array<() => void> = [];
  const acquire = () => new Promise<void>((res) => {
    if (slots > 0) { slots--; res(); } else waiters.push(() => { slots--; res(); });
  });
  const release = () => { slots++; const w = waiters.shift(); if (w) w(); };

  const inflight: Promise<void>[] = [];
  type DocInfoLite = { id: string; docnum: string; nit?: string; cufe?: string; docType?: string };
  const failedForSweep: Array<{ doc: DocInfoLite; cookies: Record<string, string> }> = [];

  const handleDoc = async (doc: DocInfoLite, cookies: Record<string, string>) => {
    const key = doc.cufe || doc.id;
    const isEquivalente = doc.docType?.toLowerCase().includes("equivalente") ?? false;
    const baseUrl = isEquivalente
      ? "https://catalogo-vpfe.dian.gov.co/Document/DownloadZipFilesEquivalente?trackId="
      : "https://catalogo-vpfe.dian.gov.co/Document/DownloadZipFiles?trackId=";
    const cookieHeader = Object.entries(cookies).map(([k, v]) => `${k}=${v}`).join("; ");
    const td0 = Date.now();
    const rec: DocResult = { cufe: key, docnum: doc.docnum, nit: doc.nit || "", dlOk: false, hasData: false };
    try {
      const buf = await throttledDianDownload(`${baseUrl}${doc.id}`, cookieHeader);
      rec.dlOk = true;
      const zip = await JSZip.loadAsync(buf);
      const xmlName = Object.keys(zip.files).find((n) => n.toLowerCase().endsWith(".xml"));
      if (!xmlName) throw new Error("ZIP sin XML");
      const xmlBuf = Buffer.from(await zip.files[xmlName].async("nodebuffer"));
      const data = await extractInvoiceDataFromXml(xmlBuf, { id: doc.id, docnum: doc.docnum, docType: doc.docType } as never);
      const d = data as { docNumber?: string; issueDate?: string; issuerNit?: string; total?: number };
      rec.hasData = !!(d && d.docNumber && d.issueDate && d.issuerNit);
      rec.total = Number(d.total ?? NaN);
      if (rowsWriter) {
        // Misma fila que construye dianExcel (mismos defaults "N/A"/0).
        const inv = data as Partial<InvoiceData>;
        const row: InvoiceData = {
          issuerNit: inv.issuerNit || "N/A",
          issuerName: inv.issuerName || "N/A",
          issuerEmail: inv.issuerEmail || "N/A",
          issuerPhone: inv.issuerPhone || "N/A",
          issuerAddress: inv.issuerAddress || "N/A",
          issuerCity: inv.issuerCity || "N/A",
          issuerDepartment: inv.issuerDepartment || "N/A",
          issuerCountry: inv.issuerCountry || "N/A",
          issuerCommercialName: inv.issuerCommercialName || "N/A",
          issuerTaxpayerType: inv.issuerTaxpayerType || "N/A",
          issuerFiscalRegime: inv.issuerFiscalRegime || "N/A",
          issuerTaxResponsibility: inv.issuerTaxResponsibility || "N/A",
          issuerEconomicActivity: inv.issuerEconomicActivity || "N/A",
          receiverNit: inv.receiverNit || "N/A",
          receiverName: inv.receiverName || "N/A",
          receiverEmail: inv.receiverEmail || "N/A",
          receiverPhone: inv.receiverPhone || "N/A",
          receiverAddress: inv.receiverAddress || "N/A",
          receiverCity: inv.receiverCity || "N/A",
          receiverDepartment: inv.receiverDepartment || "N/A",
          receiverCountry: inv.receiverCountry || "N/A",
          receiverCommercialName: inv.receiverCommercialName || "N/A",
          receiverTaxpayerType: inv.receiverTaxpayerType || "N/A",
          receiverFiscalRegime: inv.receiverFiscalRegime || "N/A",
          receiverTaxResponsibility: inv.receiverTaxResponsibility || "N/A",
          receiverEconomicActivity: inv.receiverEconomicActivity || "N/A",
          issueDate: inv.issueDate || "N/A",
          issueDateISO: inv.issueDateISO || "9999-12-31",
          paymentMethod: inv.paymentMethod || "N/A",
          subtotal: inv.subtotal || 0,
          iva: inv.iva || 0,
          total: inv.total || 0,
          taxes: inv.taxes || [],
          discount: inv.discount || 0,
          surcharge: inv.surcharge || 0,
          concepts: inv.concepts || "N/A",
          lineItems: inv.lineItems || [],
          documentType: inv.documentType || "Factura Electrónica",
          isDocumentoSoporte: inv.isDocumentoSoporte || false,
          cufe: inv.cufe || "N/A",
          notes: inv.notes || "",
          trackId: doc.id,
          docNumber: inv.docNumber || doc.docnum || doc.id,
          zipFilename: `${inv.issuerNit || doc.nit} - ${doc.docnum}.zip`,
        };
        rowsWriter.write(JSON.stringify(row) + "\n");
      }
    } catch (err) {
      rec.error = err instanceof Error ? err.message : String(err);
      failedForSweep.push({ doc, cookies });
    }
    rec.ms = Date.now() - td0;
    results.set(key, rec);
    processed++;
    if (processed % 10 === 0 || rec.error) {
      const fails = [...results.values()].filter((r) => !r.dlOk || !r.hasData).length;
      log(`${processed} procesados, ${fails} con problema${rec.error ? ` | ÚLTIMO ERROR ${key.slice(0, 12)}: ${rec.error}` : ""}`);
    }
  };

  log(`Inicio: ${startDate}→${endDate} dir=${direction} limit=${limit || "todos"} offset=${offset} workers=${workers}`);

  const { documents } = await extractDocumentIdsByCufe(
    tokenUrl,
    startDate,
    endDate,
    undefined,
    direction,
    (p) => { if (p.step && /listado|Generando|navegador|token/i.test(p.step)) log(`progreso: ${p.step}`); },
    async ({ doc, cookies }) => {
      seen++;
      if (seen <= offset) return;
      if (limit > 0 && seen > offset + limit) return;
      await acquire();
      const p = (async () => {
        try { await handleDoc(doc, cookies); } finally { release(); }
      })();
      inflight.push(p);
    },
  );

  await Promise.all(inflight);

  // Barrido serial (como el de las herramientas): cooldown + reintento de fallidos.
  for (let sweep = 1; sweep <= 3; sweep++) {
    const pending = [...results.values()].filter((r) => !r.dlOk);
    if (pending.length === 0) break;
    const byKey = new Map(failedForSweep.map((f) => [f.doc.cufe || f.doc.id, f]));
    const cooldown = 20000 * sweep;
    log(`Barrido ${sweep}: ${pending.length} pendientes; cooldown ${cooldown / 1000}s`);
    await new Promise((r) => setTimeout(r, cooldown));
    let recovered = 0;
    for (const rec of pending) {
      const item = byKey.get(rec.cufe);
      if (!item) continue;
      try { await handleDoc(item.doc, item.cookies); recovered++; } catch { /* sigue pendiente */ }
    }
    log(`Barrido ${sweep}: recuperados ${recovered}/${pending.length}`);
    if (recovered === 0 && sweep >= 2) break;
  }

  const all = [...results.values()];
  const dlFail = all.filter((r) => !r.dlOk);
  const noData = all.filter((r) => r.dlOk && !r.hasData);
  const elapsed = (Date.now() - t0) / 1000;
  const summary = {
    label,
    range: `${startDate}..${endDate}`,
    direction,
    listedResolved: documents.length,
    attempted: all.length,
    downloadOk: all.length - dlFail.length,
    downloadFail: dlFail.length,
    parsedNoData: noData.length,
    complete: all.length - dlFail.length - noData.length,
    elapsedSec: Math.round(elapsed),
    ratePerMin: all.length ? Math.round((all.length / elapsed) * 600) / 10 : 0,
    sumTotal: Math.round(all.reduce((s, r) => s + (Number.isFinite(r.total) ? (r.total as number) : 0), 0) * 100) / 100,
    docsConTotalCero: all.filter((r) => r.hasData && !(Number(r.total) > 0)).map((r) => r.docnum),
    failures: [...dlFail, ...noData].map((r) => ({ cufe: r.cufe.slice(0, 20), docnum: r.docnum, error: r.error })),
  };
  if (rowsWriter) await new Promise<void>((res) => rowsWriter.end(res));
  console.log(`RESULT ${JSON.stringify(summary, null, 2)}`);
  process.exit(0);
}

main().catch((e) => {
  console.error(`[${label}] FATAL`, e?.stack || e?.message || e);
  process.exit(1);
});
