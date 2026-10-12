/**
 * API de contabilidad: parametrización, causación en Siigo y centro de errores.
 */
import { Router } from "express";
import { requirePermission, requireSiigo, ErpHttpError } from "../context.js";
import { applySuggestion, previewSuggestion, updateAccounting } from "../accounting.js";
import {
  postDocument, postDocumentsBatch, reconcileDocument, strategyFor, validateForPosting, buildPurchasePayload,
} from "../posting.js";
import { listDocuments, requireDocument, transitionDocument } from "../documents.js";
import { createRule, deleteRule, listRules, updateRule } from "../rules.js";
import { getAccountingDashboard } from "../dashboard.js";
import { getAllCatalogs } from "../catalogs.js";
import { handler, ok, ctxOf, toNumber } from "./helpers.js";

export function createAccountingRouter(): Router {
  const router = Router();

  router.get(
    "/dashboard",
    requirePermission("accounting.read"),
    handler(async (req, res) => ok(res, await getAccountingDashboard(ctxOf(req).tenant.id)))
  );

  /** Cola de trabajo: por causar / causados / con error. */
  router.get(
    "/queue",
    requirePermission("accounting.read"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const scope = String(req.query.scope || "pending");
      const status =
        scope === "posted"
          ? ["POSTED", "PAYABLE", "PAYMENT_SCHEDULED", "PARTIALLY_PAID", "PAID", "PAYMENT_POSTED", "CLOSED"]
          : scope === "errors"
            ? ["POSTING_ERROR"]
            : ["PENDING_ACCOUNTING", "ACCOUNTING_READY"];

      const page = await listDocuments(ctx.tenant.id, {
        status,
        search: req.query.search ? String(req.query.search) : undefined,
        page: toNumber(req.query.page),
        pageSize: toNumber(req.query.pageSize) ?? 50,
        sortBy: scope === "posted" ? "updatedAt" : "issueDate",
        sortDir: scope === "posted" ? "desc" : "asc",
      });
      return ok(res, page);
    })
  );

  router.get(
    "/catalogs",
    requirePermission("accounting.read"),
    handler(async (req, res) => ok(res, await getAllCatalogs(ctxOf(req).tenant.id)))
  );

  // ── Parametrización de un documento ────────────────────────────────────
  router.get(
    "/documents/:id/suggestion",
    requirePermission("accounting.read"),
    handler(async (req, res) => ok(res, await previewSuggestion(ctxOf(req), req.params.id)))
  );

  router.post(
    "/documents/:id/suggestion/apply",
    requirePermission("accounting.edit"),
    handler(async (req, res) => ok(res, await applySuggestion(ctxOf(req), req.params.id)))
  );

  router.put(
    "/documents/:id",
    requirePermission("accounting.edit"),
    handler(async (req, res) => ok(res, await updateAccounting(ctxOf(req), req.params.id, req.body || {})))
  );

  /** Confirma la parametrización y deja el documento listo para causar. */
  router.post(
    "/documents/:id/ready",
    requirePermission("accounting.edit"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      await updateAccounting(ctx, req.params.id, { ...(req.body || {}), confirm: true });
      const doc = await transitionDocument(ctx, req.params.id, "MARK_ACCOUNTING_READY");
      return ok(res, doc);
    })
  );

  /** Validación previa: muestra "Validación completada" o errores concretos. */
  router.get(
    "/documents/:id/validate",
    requirePermission("accounting.read"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const doc = await requireDocument(ctx.tenant.id, req.params.id);
      const issues = await validateForPosting(ctx.tenant, doc);
      return ok(res, {
        strategy: strategyFor(doc, ctx.tenant),
        valid: !issues.some((i) => i.severity === "ERROR"),
        issues,
        // Vista previa del payload — solo para personal de ContaGO.
        payload: ctx.isContagoStaff ? buildPurchasePayload(ctx.tenant, doc) : undefined,
      });
    })
  );

  // ── Causación ──────────────────────────────────────────────────────────
  router.post(
    "/documents/:id/post",
    requirePermission("accounting.post_siigo"),
    requireSiigo,
    handler(async (req, res) => {
      const result = await postDocument(ctxOf(req), req.params.id);
      if (!result.ok) {
        return res.status(422).json({
          ok: false,
          code: result.error?.code || "posting_failed",
          message: result.error?.message || "No pudimos causar el documento.",
          action: result.error?.action,
          hint: result.error?.hint,
          data: { document: result.document },
        });
      }
      return ok(res, result);
    })
  );

  router.post(
    "/post-batch",
    requirePermission("accounting.post_siigo"),
    requireSiigo,
    handler(async (req, res) => {
      const ids = Array.isArray(req.body?.ids) ? req.body.ids.map(String) : [];
      if (!ids.length) throw new ErpHttpError("Selecciona al menos un documento.", 400, "no_selection");
      return ok(res, await postDocumentsBatch(ctxOf(req), ids));
    })
  );

  router.post(
    "/documents/:id/reconcile",
    requirePermission("accounting.post_siigo"),
    requireSiigo,
    handler(async (req, res) => {
      const result = await reconcileDocument(ctxOf(req), req.params.id);
      if (!result.ok) {
        return res.status(404).json({
          ok: false,
          code: "not_found_in_siigo",
          message: result.error?.message,
          data: { document: result.document },
        });
      }
      return ok(res, result);
    })
  );

  // ── Reglas contables ───────────────────────────────────────────────────
  router.get(
    "/rules",
    requirePermission("accounting.read"),
    handler(async (req, res) => ok(res, await listRules(ctxOf(req).tenant.id)))
  );

  router.post(
    "/rules",
    requirePermission("settings.accounting"),
    handler(async (req, res) => ok(res, await createRule(ctxOf(req).tenant.id, req.body || {}), 201))
  );

  router.put(
    "/rules/:id",
    requirePermission("settings.accounting"),
    handler(async (req, res) => {
      const rule = await updateRule(ctxOf(req).tenant.id, req.params.id, req.body || {});
      if (!rule) throw new ErpHttpError("Regla no encontrada.", 404, "rule_not_found");
      return ok(res, rule);
    })
  );

  router.delete(
    "/rules/:id",
    requirePermission("settings.accounting"),
    handler(async (req, res) => ok(res, { deleted: await deleteRule(ctxOf(req).tenant.id, req.params.id) }))
  );

  return router;
}
