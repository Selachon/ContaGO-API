/**
 * API de documentos: bandeja, detalle, workflow, archivos y comentarios.
 */
import { Router } from "express";
import multer from "multer";
import { requirePermission } from "../context.js";
import { ErpHttpError } from "../context.js";
import {
  actionsFor, addComment, createDocument, getDocument, listComments, listDocuments,
  requireDocument, transitionDocument, updateAdministrative, attachFile, updateFileStorage,
} from "../documents.js";
import { buildFileRecord, uploadToDrive, CATEGORY_FOLDERS, DriveNotConnectedError } from "../drive.js";
import { listAudit } from "../audit.js";
import { listSuppliers, upsertSupplier } from "../suppliers.js";
import { ingestElectronicInvoices, ingestFromFiles } from "../ingest.js";
import { createJob, failJob, finishJob, getJob, updateProgress } from "../jobs.js";
import { extractCufesFromExcel, resolveExcelBuffer } from "../../routes/dianCufeDownload.js";
import { globalSearch } from "../search.js";
import { COL, col, fromDocs, scoped } from "../db.js";
import { round2, sum } from "../money.js";
import { handler, ok, ctxOf, toArray, toBool, toDate, toNumber } from "./helpers.js";
import type { ErpSavedView } from "../types.js";

/** Límite y tipos permitidos de subida (§68). */
const ALLOWED_MIME = new Set([
  "application/pdf",
  "application/xml", "text/xml",
  "image/jpeg", "image/png",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/zip", "application/x-zip-compressed",
]);

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024, files: 20 },
  fileFilter: (_req, file, cb) => {
    if (!ALLOWED_MIME.has(file.mimetype)) {
      cb(new Error(`Tipo de archivo no permitido: ${file.mimetype}`));
      return;
    }
    cb(null, true);
  },
});

export function createDocumentsRouter(): Router {
  const router = Router();

  // ── Bandeja ────────────────────────────────────────────────────────────
  router.get(
    "/",
    requirePermission("documents.read"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const q = req.query;
      const page = await listDocuments(ctx.tenant.id, {
        status: toArray(q.status),
        type: toArray(q.type),
        source: toArray(q.source),
        supplierNit: q.supplierNit ? String(q.supplierNit) : undefined,
        search: q.search ? String(q.search) : undefined,
        issueFrom: toDate(q.issueFrom),
        issueTo: toDate(q.issueTo),
        dueFrom: toDate(q.dueFrom),
        dueTo: toDate(q.dueTo),
        costCenterId: toNumber(q.costCenterId),
        subCostCenterId: q.subCostCenterId ? String(q.subCostCenterId) : undefined,
        assignedToId: q.assignedToId ? String(q.assignedToId) : undefined,
        posted: toBool(q.posted),
        paid: toBool(q.paid),
        hasError: toBool(q.hasError),
        overdue: toBool(q.overdue),
        page: toNumber(q.page),
        pageSize: toNumber(q.pageSize),
        sortBy: q.sortBy ? String(q.sortBy) : undefined,
        sortDir: q.sortDir === "asc" ? "asc" : "desc",
      });
      return ok(res, page);
    })
  );

  // ── Búsqueda global (Cmd+K) ────────────────────────────────────────────
  router.get(
    "/search",
    requirePermission("documents.read"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const hits = await globalSearch(ctx.tenant.id, String(req.query.q || ""));
      return ok(res, hits);
    })
  );

  // ── Vistas guardadas ───────────────────────────────────────────────────
  router.get(
    "/views",
    requirePermission("documents.read"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const docs = await col(COL.savedViews)
        .find(scoped(ctx.tenant.id, { $or: [{ shared: true }, { ownerId: ctx.actorId }] }))
        .sort({ name: 1 })
        .toArray();
      return ok(res, fromDocs<ErpSavedView>(docs));
    })
  );

  router.post(
    "/views",
    requirePermission("documents.read"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const { name, filters, shared } = req.body || {};
      if (!name) throw new ErpHttpError("La vista necesita un nombre.", 400, "name_required");
      const doc = {
        tenantId: ctx.tenant.id,
        name: String(name).slice(0, 80),
        ownerId: ctx.actorId,
        shared: !!shared,
        filters: filters || {},
        createdAt: new Date().toISOString(),
      };
      const result = await col(COL.savedViews).insertOne(doc as any);
      return ok(res, { id: String(result.insertedId), ...doc }, 201);
    })
  );

  router.delete(
    "/views/:id",
    requirePermission("documents.read"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const { toId } = await import("../db.js");
      await col(COL.savedViews).deleteOne(
        scoped(ctx.tenant.id, { _id: toId(req.params.id), ownerId: ctx.actorId }) as any
      );
      return ok(res, { deleted: true });
    })
  );

  // ── Proveedores (para selects y el formulario de cuentas de cobro) ─────
  router.get(
    "/suppliers",
    requirePermission("documents.read"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const items = await listSuppliers(ctx.tenant.id, {
        search: req.query.search ? String(req.query.search) : undefined,
        limit: toNumber(req.query.limit),
      });
      return ok(res, items);
    })
  );

  router.post(
    "/suppliers",
    requirePermission("documents.create"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const supplier = await upsertSupplier(ctx.tenant.id, req.body || {});
      return ok(res, supplier, 201);
    })
  );

  // ── Alta manual / cuenta de cobro ──────────────────────────────────────
  router.post(
    "/",
    requirePermission("documents.create"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const b = req.body || {};
      if (!b.supplierNit) throw new ErpHttpError("El NIT o cédula del beneficiario es obligatorio.", 400, "supplier_required");
      if (!b.issueDate) throw new ErpHttpError("La fecha del documento es obligatoria.", 400, "date_required");

      const items = Array.isArray(b.items) && b.items.length ? b.items : undefined;
      const subtotal = items ? sum(items.map((i: any) => Number(i.subtotal ?? 0))) : round2(Number(b.subtotal ?? 0));
      const taxes = items ? sum(items.map((i: any) => Number(i.taxAmount ?? 0))) : round2(Number(b.taxes ?? 0));

      const requireApproval = ctx.tenant.settings?.requireApproval !== false;
      const threshold = ctx.tenant.settings?.approvalThreshold ?? 0;
      const total = round2(Number(b.total ?? subtotal + taxes));
      const needsApproval = requireApproval && total >= threshold;

      const doc = await createDocument(ctx, {
        source: b.source === "FORM" ? "FORM" : "MANUAL",
        type: b.type || "COLLECTION_ACCOUNT",
        supplierNit: String(b.supplierNit),
        supplierName: String(b.supplierName || ""),
        documentPrefix: b.documentPrefix,
        documentNumber: b.documentNumber,
        issueDate: b.issueDate,
        dueDate: b.dueDate,
        subtotal,
        taxes,
        total,
        description: b.description,
        items,
        costCenterId: toNumber(b.costCenterId),
        subCostCenterId: b.subCostCenterId,
        estimatedPaymentDate: b.estimatedPaymentDate,
        metadata: b.metadata,
        status: needsApproval ? "PENDING_APPROVAL" : "APPROVED",
      });
      return ok(res, doc, 201);
    })
  );

  // ── Importación DIAN ───────────────────────────────────────────────────
  // Mismo flujo que "Descarga Masiva": el usuario exporta desde el propio
  // portal de la DIAN el Excel con los CUFE (columna B + columna "Grupo") y lo
  // sube aquí junto con el enlace del token. Se salta el listado por navegador
  // —el paso lento— y se descarga cada CUFE directo por HTTP, igual que esa
  // herramienta; la única diferencia es que el destino no es un ZIP para el
  // equipo del usuario, sino el propio documento del ERP.
  //
  // Es un job en segundo plano: la descarga y el guardado en Drive pueden
  // tardar minutos, y una petición HTTP bloqueada durante todo ese tiempo es
  // indistinguible de un portal colgado. Esta ruta solo arranca el trabajo y
  // devuelve un `jobId`; el progreso real se consulta con la ruta de abajo.
  router.post(
    "/import/dian",
    requirePermission("documents.create"),
    upload.single("excel"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      if (!ctx.tenant.features?.dian) {
        throw new ErpHttpError("El módulo DIAN no está activo para esta empresa.", 403, "feature_disabled");
      }
      const b = req.body || {};
      if (!b.tokenUrl) throw new ErpHttpError("Falta el enlace del token DIAN.", 400, "token_required");
      const file = req.file;
      if (!file) {
        throw new ErpHttpError(
          "Adjunta el Excel de CUFEs exportado desde el portal de la DIAN.",
          400,
          "excel_required"
        );
      }

      let parsed: Awaited<ReturnType<typeof extractCufesFromExcel>>;
      try {
        const excelBuffer = await resolveExcelBuffer(file);
        parsed = await extractCufesFromExcel(excelBuffer);
      } catch (err) {
        throw new ErpHttpError(
          `Error leyendo el Excel: ${err instanceof Error ? err.message : String(err)}`,
          400,
          "invalid_excel"
        );
      }
      if (parsed.mixedDirections) {
        throw new ErpHttpError(
          "El listado contiene documentos emitidos y recibidos. Sube solo el grupo de recibidos.",
          400,
          "mixed_directions"
        );
      }
      if (parsed.detectedDirection === "sent") {
        throw new ErpHttpError(
          "Este Excel contiene facturas emitidas. El ERP solo importa facturas recibidas de proveedores.",
          400,
          "wrong_direction"
        );
      }
      if (parsed.cufes.length === 0) {
        throw new ErpHttpError(
          "La columna B del Excel no contiene documentos procesables (nómina electrónica y Application Response no son compatibles).",
          400,
          "no_cufes"
        );
      }

      // CUFEs ya presentes: evita re-descargar lo que el ERP ya tiene.
      const existing = await col(COL.documents)
        .find(scoped(ctx.tenant.id, { cufe: { $type: "string" } }), { projection: { cufe: 1 } })
        .limit(20000)
        .toArray();

      const jobId = createJob(ctx.tenant.id);

      // Deliberadamente sin `await`: el resultado se recoge por polling. Los
      // errores del propio job (incluidos los del arranque) se capturan aquí
      // mismo para que nunca queden como una promesa rechazada sin manejar.
      void ingestElectronicInvoices(ctx, {
        tokenUrl: b.tokenUrl,
        providedCufes: parsed.cufes,
        skipCufes: existing.map((d: any) => d.cufe).filter(Boolean),
        onProgress: (step, percent) => updateProgress(jobId, step, percent),
      })
        .then((summary) => finishJob(jobId, { ...summary, documents: summary.documents.length }))
        .catch((err) => failJob(jobId, err instanceof Error ? err.message : "Error inesperado durante la importación."));

      return ok(res, { jobId, totalCufes: parsed.cufes.length }, 202);
    })
  );

  /** Progreso del job de importación. Nunca devuelve el job de otra empresa. */
  router.get(
    "/import/dian/:jobId",
    requirePermission("documents.create"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const job = getJob(ctx.tenant.id, req.params.jobId);
      if (!job) throw new ErpHttpError("El proceso de importación no existe o ya expiró.", 404, "job_not_found");
      return ok(res, job);
    })
  );

  router.post(
    "/import/files",
    requirePermission("documents.create"),
    upload.array("files", 20),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const files = (req.files as Express.Multer.File[]) || [];
      if (!files.length) throw new ErpHttpError("Adjunta al menos un archivo XML o ZIP.", 400, "no_files");
      const summary = await ingestFromFiles(
        ctx,
        files.map((f) => ({ buffer: f.buffer, originalname: f.originalname }))
      );
      return ok(res, { ...summary, documents: summary.documents.length });
    })
  );

  // ── Detalle ────────────────────────────────────────────────────────────
  router.get(
    "/:id",
    requirePermission("documents.read"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const doc = await requireDocument(ctx.tenant.id, req.params.id);
      const [timeline, comments] = await Promise.all([
        listAudit(ctx.tenant.id, { entityType: "document", entityId: doc.id, limit: 100 }),
        listComments(ctx.tenant.id, doc.id, ctx.isContagoStaff || ctx.member?.role === "CONTAGO_ACCOUNTANT"),
      ]);
      return ok(res, {
        document: doc,
        actions: actionsFor(doc, ctx.permissions),
        timeline,
        comments,
      });
    })
  );

  // ── Datos administrativos ──────────────────────────────────────────────
  router.patch(
    "/:id",
    requirePermission("documents.edit"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const doc = await updateAdministrative(ctx, req.params.id, req.body || {});
      return ok(res, doc);
    })
  );

  // ── Transiciones de estado ─────────────────────────────────────────────
  router.post(
    "/:id/transition",
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const { transition, reason, version } = req.body || {};
      if (!transition) throw new ErpHttpError("Indica la acción a ejecutar.", 400, "transition_required");
      const doc = await transitionDocument(ctx, req.params.id, String(transition), {
        reason,
        version: toNumber(version),
      });
      return ok(res, { document: doc, actions: actionsFor(doc, ctx.permissions) });
    })
  );

  router.post(
    "/:id/approve",
    requirePermission("documents.approve"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const doc = await transitionDocument(ctx, req.params.id, "APPROVE", { version: toNumber(req.body?.version) });
      return ok(res, doc);
    })
  );

  router.post(
    "/:id/reject",
    requirePermission("documents.reject"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const reason = String(req.body?.reason || "").trim();
      if (!reason) throw new ErpHttpError("El motivo del rechazo es obligatorio.", 400, "reason_required");
      const doc = await transitionDocument(ctx, req.params.id, "REJECT", { reason, version: toNumber(req.body?.version) });
      return ok(res, doc);
    })
  );

  // ── Acciones masivas ───────────────────────────────────────────────────
  router.post(
    "/bulk",
    requirePermission("documents.read"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const { ids, action, payload } = req.body || {};
      if (!Array.isArray(ids) || !ids.length) throw new ErpHttpError("Selecciona al menos un documento.", 400, "no_selection");
      if (ids.length > 200) throw new ErpHttpError("Máximo 200 documentos por operación.", 400, "too_many");

      const results: Array<{ id: string; ok: boolean; message?: string }> = [];
      for (const id of ids) {
        try {
          if (action === "approve") await transitionDocument(ctx, id, "APPROVE");
          else if (action === "reject") await transitionDocument(ctx, id, "REJECT", { reason: payload?.reason });
          else if (action === "assign") await updateAdministrative(ctx, id, { assignedToId: payload?.assignedToId });
          else if (action === "estimatedDate") await updateAdministrative(ctx, id, { estimatedPaymentDate: payload?.date });
          else if (action === "costCenter") await updateAdministrative(ctx, id, { costCenterId: toNumber(payload?.costCenterId) });
          else throw new ErpHttpError(`Acción masiva no soportada: ${action}`, 400, "unknown_action");
          results.push({ id, ok: true });
        } catch (err) {
          results.push({ id, ok: false, message: err instanceof Error ? err.message : "Error" });
        }
      }
      return ok(res, {
        processed: results.length,
        succeeded: results.filter((r) => r.ok).length,
        failed: results.filter((r) => !r.ok).length,
        results,
      });
    })
  );

  // ── Archivos ───────────────────────────────────────────────────────────
  router.post(
    "/:id/files",
    requirePermission("documents.edit"),
    upload.array("files", 10),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const doc = await requireDocument(ctx.tenant.id, req.params.id);
      const files = (req.files as Express.Multer.File[]) || [];
      if (!files.length) throw new ErpHttpError("Adjunta al menos un archivo.", 400, "no_files");

      const uploaded: Array<{ id: string; name: string; status: string; error?: string }> = [];
      for (const f of files) {
        const kind = f.mimetype === "application/pdf" ? "PDF" : /xml/.test(f.mimetype) ? "XML" : "ATTACHMENT";
        const record = buildFileRecord({
          kind, name: f.originalname, mimeType: f.mimetype, size: f.size, uploadedBy: ctx.actorId,
        });
        await attachFile(ctx.tenant.id, doc.id, record);

        try {
          const result = await uploadToDrive(ctx.tenant, {
            buffer: f.buffer,
            name: record.name,
            mimeType: f.mimetype,
            category: CATEGORY_FOLDERS[doc.type] ? doc.type : "OTHER",
            date: doc.issueDate,
          });
          await updateFileStorage(ctx.tenant.id, doc.id, record.id, {
            storageStatus: "STORED",
            driveFileId: result.driveFileId,
            driveWebViewLink: result.driveWebViewLink,
            size: result.size,
          });
          uploaded.push({ id: record.id, name: record.name, status: "STORED" });
        } catch (err) {
          const message = err instanceof DriveNotConnectedError
            ? "Google Drive no está conectado para esta empresa."
            : err instanceof Error ? err.message : String(err);
          await updateFileStorage(ctx.tenant.id, doc.id, record.id, { storageStatus: "FAILED", storageError: message });
          uploaded.push({ id: record.id, name: record.name, status: "FAILED", error: message });
        }
      }

      const refreshed = await getDocument(ctx.tenant.id, doc.id);
      return ok(res, { document: refreshed, uploaded }, 201);
    })
  );

  // ── Comentarios ────────────────────────────────────────────────────────
  router.post(
    "/:id/comments",
    requirePermission("documents.read"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const visibility = req.body?.visibility === "INTERNAL" && (ctx.isContagoStaff || ctx.member?.role === "CONTAGO_ACCOUNTANT")
        ? "INTERNAL"
        : "TENANT";
      const comment = await addComment(ctx, req.params.id, String(req.body?.message || ""), visibility);
      return ok(res, comment, 201);
    })
  );

  return router;
}
