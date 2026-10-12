/**
 * API de tesorería: cuentas por pagar, programación, pagos, egresos y flujo de caja.
 */
import { Router } from "express";
import { requirePermission, requireSiigo, ErpHttpError } from "../context.js";
import {
  cancelSchedule, createBankAccount, createSchedule, listBankAccounts, listPayables,
  listPayments, listSchedules, getPayment, postPaymentToSiigo, recordPayment, updateBankAccount,
} from "../payments.js";
import { getCashFlow } from "../cashflow.js";
import { getTreasuryDashboard } from "../dashboard.js";
import { handler, ok, ctxOf, toDate, toNumber } from "./helpers.js";
import type { AgingBucket } from "../payments.js";

export function createTreasuryRouter(): Router {
  const router = Router();

  router.get(
    "/dashboard",
    requirePermission("payables.read"),
    handler(async (req, res) => ok(res, await getTreasuryDashboard(ctxOf(req).tenant.id)))
  );

  // ── Cuentas por pagar ──────────────────────────────────────────────────
  router.get(
    "/payables",
    requirePermission("payables.read"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const data = await listPayables(ctx.tenant.id, {
        search: req.query.search ? String(req.query.search) : undefined,
        supplierNit: req.query.supplierNit ? String(req.query.supplierNit) : undefined,
        costCenterId: toNumber(req.query.costCenterId),
        bucket: req.query.bucket ? (String(req.query.bucket) as AgingBucket) : undefined,
        dueTo: toDate(req.query.dueTo),
        includeScheduled: req.query.includeScheduled !== "false",
      });
      return ok(res, data);
    })
  );

  // ── Programación de pagos ──────────────────────────────────────────────
  router.get(
    "/schedules",
    requirePermission("payables.read"),
    handler(async (req, res) =>
      ok(res, await listSchedules(ctxOf(req).tenant.id, {
        status: req.query.status ? String(req.query.status) : undefined,
        from: toDate(req.query.from),
        to: toDate(req.query.to),
      }))
    )
  );

  router.post(
    "/schedules",
    requirePermission("payables.schedule"),
    handler(async (req, res) => ok(res, await createSchedule(ctxOf(req), req.body || {}), 201))
  );

  router.post(
    "/schedules/:id/cancel",
    requirePermission("payables.schedule"),
    handler(async (req, res) => ok(res, await cancelSchedule(ctxOf(req), req.params.id)))
  );

  // ── Pagos efectivos ────────────────────────────────────────────────────
  router.get(
    "/payments",
    requirePermission("payables.read"),
    handler(async (req, res) =>
      ok(res, await listPayments(ctxOf(req).tenant.id, {
        from: toDate(req.query.from),
        to: toDate(req.query.to),
        supplierNit: req.query.supplierNit ? String(req.query.supplierNit) : undefined,
        bankAccountId: req.query.bankAccountId ? String(req.query.bankAccountId) : undefined,
        limit: toNumber(req.query.limit),
      }))
    )
  );

  router.get(
    "/payments/:id",
    requirePermission("payables.read"),
    handler(async (req, res) => {
      const payment = await getPayment(ctxOf(req).tenant.id, req.params.id);
      if (!payment) throw new ErpHttpError("Pago no encontrado.", 404, "payment_not_found");
      return ok(res, payment);
    })
  );

  router.post(
    "/payments",
    requirePermission("payables.pay"),
    handler(async (req, res) => ok(res, await recordPayment(ctxOf(req), req.body || {}), 201))
  );

  /** Egreso en Siigo (recibo de pago) para un pago ya registrado. */
  router.post(
    "/payments/:id/post-siigo",
    requirePermission("payables.post_siigo"),
    requireSiigo,
    handler(async (req, res) => {
      const result = await postPaymentToSiigo(ctxOf(req), req.params.id);
      if (!result.ok) {
        return res.status(422).json({
          ok: false,
          code: result.error?.code || "egreso_failed",
          message: result.error?.message || "No pudimos crear el egreso en Siigo.",
          action: result.error?.action,
          hint: result.error?.hint,
          data: { payment: result.payment },
        });
      }
      return ok(res, result);
    })
  );

  // ── Cuentas bancarias ──────────────────────────────────────────────────
  router.get(
    "/bank-accounts",
    requirePermission("payables.read"),
    handler(async (req, res) => ok(res, await listBankAccounts(ctxOf(req).tenant.id)))
  );

  router.post(
    "/bank-accounts",
    requirePermission("settings.integrations"),
    handler(async (req, res) => ok(res, await createBankAccount(ctxOf(req).tenant.id, req.body || {}), 201))
  );

  router.put(
    "/bank-accounts/:id",
    requirePermission("settings.integrations"),
    handler(async (req, res) => {
      const account = await updateBankAccount(ctxOf(req).tenant.id, req.params.id, req.body || {});
      if (!account) throw new ErpHttpError("Cuenta bancaria no encontrada.", 404, "bank_account_not_found");
      return ok(res, account);
    })
  );

  // ── Flujo de caja ──────────────────────────────────────────────────────
  router.get(
    "/cashflow",
    requirePermission("cashflow.read"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const kindParam = String(req.query.kind || "");
      return ok(res, await getCashFlow(ctx.tenant.id, {
        from: toDate(req.query.from),
        to: toDate(req.query.to),
        costCenterId: toNumber(req.query.costCenterId),
        subCostCenterId: req.query.subCostCenterId ? String(req.query.subCostCenterId) : undefined,
        bankAccountId: req.query.bankAccountId ? String(req.query.bankAccountId) : undefined,
        supplierNit: req.query.supplierNit ? String(req.query.supplierNit) : undefined,
        kind: kindParam === "PROJECTED" || kindParam === "ACTUAL" ? kindParam : undefined,
      }));
    })
  );

  return router;
}
