/**
 * Puerto de facturación electrónica + adaptadores.
 *
 * El ERP no reimplementa la integración DIAN: llama literalmente a la misma
 * función que usa "Descarga Masiva DIAN" (`runDianRecibidosBatchDownload`,
 * extraída de `dianRecibidosDownload.ts` — autenticación + descarga directa
 * por CUFE contra gratis-vpfe.dian.gov.co, 4 concurrentes a 50/min, con las
 * mismas 3 pasadas de recuperación) y al mismo parser de XML/ZIP que ya usa
 * la carga manual de archivos (`processXmlBatch`). No hay una segunda
 * implementación de la descarga en ningún punto: si algo cambia en
 * `dianRecibidosBatchDownload.ts`, corrige a las dos herramientas a la vez.
 * La única diferencia real es el destino: en vez de un ZIP para el equipo del
 * usuario, cada documento se guarda en Drive y se crea en el ERP.
 */
import { randomUUID } from "node:crypto";
import { runDianRecibidosBatchDownload } from "../services/dianRecibidosBatchDownload.js";
import { processXmlBatch, type BatchItem, type XmlCausacionData } from "../services/siigoAccountingService.js";
import { round2, sum } from "./money.js";
import { upsertExternalDocument } from "./documents.js";
import { audit } from "./audit.js";
import { buildFileRecord, uploadToDrive, DriveNotConnectedError, type DriveFolderCache } from "./drive.js";
import { updateFileStorage } from "./documents.js";
import { notify } from "./notifications.js";
import type { ErpDocument, ErpDocumentFile, ErpDocumentItem } from "./types.js";
import type { ErpTenant } from "./tenants.js";
import type { ErpRequestContext } from "./context.js";

// ─── Puerto ──────────────────────────────────────────────────────────────

export interface ElectronicInvoiceDocument {
  cufe: string;
  supplierNit: string;
  supplierName: string;
  prefix: string;
  number: string;
  date: string;
  subtotal: number;
  taxes: number;
  total: number;
  items: Array<Partial<ErpDocumentItem>>;
  xmlBase64?: string;
  pdfBase64?: string;
  pdfName?: string;
}

export interface ElectronicInvoiceProvider {
  readonly name: string;
  syncDocuments(options: SyncOptions): Promise<ElectronicInvoiceDocument[]>;
}

export interface SyncOptions {
  tenant: ErpTenant;
  tokenUrl?: string;
  /** CUFEs ya presentes en el ERP: se omiten de la descarga. */
  skipCufes?: string[];
  /**
   * CUFEs del Excel exportado desde el propio portal de la DIAN (mismo insumo
   * que "Descarga Masiva"). Es obligatorio: el ERP no lista por navegador.
   */
  providedCufes: string[];
  onProgress?: (message: string, percent?: number) => void;
}

// ─── Mapeo XML → documento ERP ───────────────────────────────────────────

function mapItems(xml: XmlCausacionData): Array<Partial<ErpDocumentItem>> {
  return (xml.items || []).map((it) => {
    const base = round2(Number(it.base) || 0);
    const quantity = Number(it.quantity) || 1;
    const ivaPct = (Number(it.ivaPercent) || 0) + (Number(it.incPercent) || 0);
    const taxAmount = round2((base * ivaPct) / 100);
    return {
      id: randomUUID(),
      description: it.description || "Concepto",
      quantity,
      unitPrice: round2(it.unitPrice != null ? Number(it.unitPrice) : base / (quantity || 1)),
      subtotal: base,
      taxAmount,
      total: round2(base + taxAmount),
    };
  });
}

export function mapXmlToDocument(xml: XmlCausacionData): ElectronicInvoiceDocument {
  const items = mapItems(xml);
  return {
    cufe: xml.cufe,
    supplierNit: String(xml.supplierNit || "").replace(/\D/g, ""),
    supplierName: xml.supplierName || "Proveedor sin nombre",
    prefix: xml.providerInvoicePrefix || "",
    number: xml.providerInvoiceNumber || xml.docNumberRaw || "",
    date: xml.date || "",
    subtotal: round2(xml.totals?.subtotal ?? sum(items.map((i) => i.subtotal))),
    taxes: round2(xml.totals?.iva ?? sum(items.map((i) => i.taxAmount))),
    total: round2(xml.totals?.total ?? 0),
    items,
  };
}

// ─── Adaptador ContaGO/DIAN ──────────────────────────────────────────────

export class ContaGOElectronicInvoiceProvider implements ElectronicInvoiceProvider {
  readonly name = "contago-dian";

  async syncDocuments(options: SyncOptions): Promise<ElectronicInvoiceDocument[]> {
    const { tenant } = options;
    if (!tenant.siigoCompanyId) {
      throw new Error("La empresa no tiene compañía Siigo asociada; conéctala antes de importar desde la DIAN.");
    }
    if (!options.tokenUrl) {
      throw new Error("Falta el enlace del token DIAN para iniciar la descarga.");
    }

    const skip = new Set(options.skipCufes || []);
    const cufes = options.providedCufes.filter((c) => !skip.has(c));
    if (cufes.length === 0) return [];

    // Misma función, mismos parámetros, que usa "Descarga Masiva DIAN" — sin envoltorios.
    const { items: downloaded, downloaded: downloadedCount } = await runDianRecibidosBatchDownload({
      tokenUrl: options.tokenUrl,
      cufes,
      direction: "received",
      onProgress: (p) => options.onProgress?.(p.step || "", p.total ? Math.round((p.current / p.total) * 100) : undefined),
    });

    if (downloaded.length === 0) {
      throw new Error(`No se pudo descargar ningún documento (${cufes.length - downloadedCount} error(es)).`);
    }

    // Mismo parser que la carga manual de XML/ZIP: recibe el XML ya extraído
    // (y su PDF, si lo hay) con nombres emparejados por CUFE.
    const files = downloaded.flatMap((r) => {
      const out: Array<{ name: string; buffer: Buffer }> = [{ name: `${r.cufe}.xml`, buffer: r.xmlBuffer }];
      if (r.pdfBuffer) out.push({ name: `${r.cufe}.pdf`, buffer: r.pdfBuffer });
      return out;
    });
    const items = await processXmlBatch(files);

    return items
      .filter((item: BatchItem) => item.ok && item.xml && !item.ignored)
      .map((item: BatchItem) => {
        const mapped = mapXmlToDocument(item.xml!);
        mapped.pdfBase64 = item.pdfBase64;
        mapped.pdfName = item.pdfName;
        return mapped;
      });
  }
}

/** Proveedor de desarrollo. Nunca se activa en producción por defecto (§86). */
export class MockElectronicInvoiceProvider implements ElectronicInvoiceProvider {
  readonly name = "mock";
  async syncDocuments(): Promise<ElectronicInvoiceDocument[]> {
    return [];
  }
}

export function getInvoiceProvider(): ElectronicInvoiceProvider {
  const configured = (process.env.ELECTRONIC_INVOICE_PROVIDER || "contago").toLowerCase();
  if (configured === "mock") {
    if (process.env.NODE_ENV === "production") {
      throw new Error("ELECTRONIC_INVOICE_PROVIDER=mock no está permitido en producción.");
    }
    return new MockElectronicInvoiceProvider();
  }
  return new ContaGOElectronicInvoiceProvider();
}

// ─── Ingesta hacia el ERP ────────────────────────────────────────────────

export interface IngestSummary {
  fetched: number;
  created: number;
  duplicated: number;
  filesStored: number;
  filesPending: number;
  errors: Array<{ cufe: string; message: string }>;
  documents: ErpDocument[];
}

/**
 * Importa facturas electrónicas al ERP.
 * Deduplica por CUFE, sube los soportes a Drive y deja el documento en el
 * estado inicial del flujo (revisión/aprobación).
 */
export async function ingestElectronicInvoices(
  ctx: ErpRequestContext,
  options: Omit<SyncOptions, "tenant">
): Promise<IngestSummary> {
  const tenant = ctx.tenant;
  const provider = getInvoiceProvider();
  const summary: IngestSummary = {
    fetched: 0,
    created: 0,
    duplicated: 0,
    filesStored: 0,
    filesPending: 0,
    errors: [],
    documents: [],
  };

  const downloadStart = Date.now();
  const invoices = await provider.syncDocuments({ ...options, tenant });
  summary.fetched = invoices.length;
  console.log(`[ERP][DIAN] Descarga: ${invoices.length} documento(s) en ${((Date.now() - downloadStart) / 1000).toFixed(1)}s`);

  // La descarga terminó; ahora viene guardar cada documento (Drive + Mongo).
  // Sin esto el usuario ve el progreso saltar de "descargando" a "listo" con
  // un salto largo y silencioso en medio si hay muchos documentos.
  //
  // Cada documento es independiente (CUFE distinto → registro distinto), así
  // que se guardan con el mismo patrón de concurrencia acotada que ya usa la
  // descarga (semáforo de N a la vez) en vez de uno por uno: subir a Drive es
  // la parte de red más lenta de todo el proceso y hacerla en serie multiplica
  // el tiempo total por el número de documentos.
  let saved = 0;
  const saveStart = Date.now();
  const driveFolderCache: DriveFolderCache = new Map();
  options.onProgress?.(
    invoices.length ? `Guardando ${invoices.length} documento(s)…` : "Sin documentos nuevos para guardar.",
    invoices.length ? 0 : undefined
  );

  const saveOne = async (inv: ElectronicInvoiceDocument): Promise<void> => {
    try {
      const files: ErpDocumentFile[] = [];
      if (inv.pdfBase64) {
        files.push(
          buildFileRecord({
            kind: "PDF",
            name: inv.pdfName || `${inv.prefix}${inv.number}.pdf`,
            mimeType: "application/pdf",
            uploadedBy: ctx.actorId,
          })
        );
      }

      const { document, created } = await upsertExternalDocument(
        tenant.id,
        {
          source: "DIAN",
          type: "ELECTRONIC_INVOICE",
          supplierNit: inv.supplierNit,
          supplierName: inv.supplierName,
          documentPrefix: inv.prefix,
          documentNumber: inv.number,
          cufe: inv.cufe,
          issueDate: inv.date,
          subtotal: inv.subtotal,
          taxes: inv.taxes,
          total: inv.total,
          items: inv.items,
          files,
          status: tenant.settings?.requireApproval === false ? "APPROVED" : "PENDING_APPROVAL",
          metadata: { ingestedBy: provider.name, ingestedAt: new Date().toISOString() },
        },
        ctx.actorId
      );

      if (!created) {
        summary.duplicated += 1;
        return;
      }
      summary.created += 1;
      summary.documents.push(document);

      // Soporte en Drive: si falla, el archivo queda PENDING — nunca se da el
      // proceso documental por terminado en silencio (§17).
      if (inv.pdfBase64 && files[0]) {
        const uploadStart = Date.now();
        try {
          const uploaded = await uploadToDrive(tenant, {
            buffer: Buffer.from(inv.pdfBase64, "base64"),
            name: files[0].name,
            mimeType: "application/pdf",
            category: "ELECTRONIC_INVOICE",
            date: inv.date,
          }, driveFolderCache);
          await updateFileStorage(tenant.id, document.id, files[0].id, {
            storageStatus: "STORED",
            driveFileId: uploaded.driveFileId,
            driveWebViewLink: uploaded.driveWebViewLink,
            size: uploaded.size,
          });
          summary.filesStored += 1;
          console.log(`[ERP][DIAN] Drive OK ${inv.cufe.slice(0, 12)}… en ${((Date.now() - uploadStart) / 1000).toFixed(1)}s`);
        } catch (err) {
          const message = err instanceof DriveNotConnectedError
            ? "Google Drive no está conectado para esta empresa."
            : err instanceof Error ? err.message : String(err);
          await updateFileStorage(tenant.id, document.id, files[0].id, {
            storageStatus: "FAILED",
            storageError: message,
          });
          summary.filesPending += 1;
          console.warn(`[ERP][DIAN] Drive FALLÓ ${inv.cufe.slice(0, 12)}… tras ${((Date.now() - uploadStart) / 1000).toFixed(1)}s: ${message}`);
        }
      }
    } catch (err) {
      summary.errors.push({ cufe: inv.cufe, message: err instanceof Error ? err.message : String(err) });
    } finally {
      saved += 1;
      options.onProgress?.(`Guardando documentos… (${saved}/${invoices.length})`, Math.round((saved / invoices.length) * 100));
    }
  };

  const SAVE_CONCURRENCY = Math.max(1, Math.min(Number(process.env.ERP_SAVE_WORKERS || 5), 10));
  for (let i = 0; i < invoices.length; i += SAVE_CONCURRENCY) {
    await Promise.all(invoices.slice(i, i + SAVE_CONCURRENCY).map(saveOne));
  }
  console.log(`[ERP][DIAN] Guardado: ${invoices.length} documento(s) en ${((Date.now() - saveStart) / 1000).toFixed(1)}s`);

  await audit({
    tenantId: tenant.id,
    action: "DIAN_INGESTED",
    entityType: "tenant",
    entityId: tenant.id,
    actorId: ctx.actorId,
    actorName: ctx.actorName,
    after: { fetched: summary.fetched, created: summary.created, duplicated: summary.duplicated },
    summary: `Importación DIAN: ${summary.created} nuevas, ${summary.duplicated} ya existentes.`,
  });

  if (summary.created) {
    await notify(tenant.id, {
      type: "DOCUMENTS_IMPORTED",
      title: `${summary.created} documento(s) importados desde la DIAN`,
      body: summary.filesPending ? `${summary.filesPending} soporte(s) quedaron pendientes de subir a Drive.` : undefined,
    });
  }

  return summary;
}

/**
 * Importa facturas desde archivos XML/ZIP cargados manualmente.
 * Reutiliza el parser probado de ContaGO.
 */
export async function ingestFromFiles(
  ctx: ErpRequestContext,
  files: Array<{ buffer: Buffer; originalname: string }>
): Promise<IngestSummary> {
  const tenant = ctx.tenant;
  const summary: IngestSummary = {
    fetched: 0, created: 0, duplicated: 0, filesStored: 0, filesPending: 0, errors: [], documents: [],
  };

  const items: BatchItem[] = await processXmlBatch(
    files.map((f) => ({ name: f.originalname, buffer: f.buffer }))
  );
  for (const item of items) {
    if (!item.ok || !item.xml) {
      if (!item.ignored) summary.errors.push({ cufe: item.fileName, message: item.message || "XML inválido" });
      continue;
    }
    summary.fetched += 1;
    const inv = mapXmlToDocument(item.xml);
    try {
      const { document, created } = await upsertExternalDocument(
        tenant.id,
        {
          source: "IMPORT",
          type: "ELECTRONIC_INVOICE",
          supplierNit: inv.supplierNit,
          supplierName: inv.supplierName,
          documentPrefix: inv.prefix,
          documentNumber: inv.number,
          cufe: inv.cufe,
          issueDate: inv.date,
          subtotal: inv.subtotal,
          taxes: inv.taxes,
          total: inv.total,
          items: inv.items,
          status: tenant.settings?.requireApproval === false ? "APPROVED" : "PENDING_APPROVAL",
          metadata: { sourceFile: item.fileName },
        },
        ctx.actorId
      );
      if (created) {
        summary.created += 1;
        summary.documents.push(document);
      } else {
        summary.duplicated += 1;
      }
    } catch (err) {
      summary.errors.push({ cufe: inv.cufe || item.fileName, message: err instanceof Error ? err.message : String(err) });
    }
  }

  return summary;
}
