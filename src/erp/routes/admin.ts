/**
 * Backoffice ContaGO. Área exclusiva del equipo interno: NO pasa por
 * `withErpTenant` (no está atada a un tenant) y exige `isAdmin`.
 */
import { Router, type NextFunction, type Request, type Response } from "express";
import { requireAuth } from "../../middleware/auth.js";
import { COL, col, scoped } from "../db.js";
import {
  createTenant, getTenantById, inviteMember, listMembers, listTenants, normalizeSlug,
  updateTenant, validateSlug, getTenantSiigoCompany,
} from "../tenants.js";
import { listCompanies } from "../../services/siigoCompaniesService.js";
import { seedDemoTenant } from "../seed.js";
import { audit } from "../audit.js";
import { getSiigoIntegrationHealth } from "../../services/siigoService.js";
import { getBrowserStats } from "../../services/dianScraper.js";
import { activeJobCount } from "../../services/jobGuard.js";
import { handler, ok, toNumber } from "./helpers.js";
import { ErpHttpError } from "../context.js";
import type { ErpRole } from "../permissions.js";

function requireContagoStaff(req: Request, res: Response, next: NextFunction): void {
  if (!req.user?.isAdmin) {
    res.status(403).json({ ok: false, code: "forbidden", message: "Área exclusiva del equipo ContaGO." });
    return;
  }
  next();
}

export function createAdminRouter(): Router {
  const router = Router();
  router.use(requireAuth, requireContagoStaff);

  /** Listado de empresas con métricas operativas. */
  router.get(
    "/tenants",
    handler(async (_req, res) => {
      const tenants = await listTenants();
      const rows = await Promise.all(
        tenants.map(async (t) => {
          const [documents, docsThisMonth, members, siigoCompany, errors] = await Promise.all([
            col(COL.documents).countDocuments({ tenantId: t.id }),
            col(COL.documents).countDocuments({
              tenantId: t.id,
              createdAt: { $gte: new Date(new Date().setUTCDate(1)).toISOString().slice(0, 10) },
            }),
            listMembers(t.id),
            getTenantSiigoCompany(t).catch(() => null),
            col(COL.documents).countDocuments({ tenantId: t.id, status: "POSTING_ERROR" }),
          ]);
          return {
            ...t,
            metrics: { documents, docsThisMonth, users: members.length, errors },
            siigo: siigoCompany ? { connected: true, name: siigoCompany.name } : { connected: false },
          };
        })
      );
      return ok(res, rows);
    })
  );

  router.post(
    "/tenants",
    handler(async (req, res) => {
      const b = req.body || {};
      const slug = normalizeSlug(b.slug || b.name || "");
      const slugError = validateSlug(slug);
      if (slugError) throw new ErpHttpError(slugError, 400, "invalid_slug");

      const tenant = await createTenant({
        name: b.name,
        legalName: b.legalName,
        nit: b.nit,
        dv: b.dv,
        slug,
        siigoCompanyId: b.siigoCompanyId,
        timezone: b.timezone,
        currency: b.currency,
        logoUrl: b.logoUrl,
        features: b.features,
        settings: b.settings,
      });

      if (b.adminEmail) {
        await inviteMember(tenant.id, {
          email: String(b.adminEmail),
          role: "CLIENT_ADMIN",
          invitedById: req.user?.userId,
        });
      }
      if (b.accountantEmail) {
        await inviteMember(tenant.id, {
          email: String(b.accountantEmail),
          role: "CONTAGO_ACCOUNTANT",
          invitedById: req.user?.userId,
        });
      }

      await audit({
        tenantId: tenant.id, action: "TENANT_CREATED", entityType: "tenant", entityId: tenant.id,
        actorId: req.user?.userId, actorName: req.user?.email,
        after: { name: tenant.name, slug: tenant.slug },
        summary: `Empresa ${tenant.name} creada por ${req.user?.email}.`, req,
      });
      return ok(res, tenant, 201);
    })
  );

  router.put(
    "/tenants/:id",
    handler(async (req, res) => {
      const tenant = await updateTenant(req.params.id, req.body || {});
      if (!tenant) throw new ErpHttpError("Empresa no encontrada.", 404, "tenant_not_found");
      await audit({
        tenantId: tenant.id, action: "TENANT_UPDATED", entityType: "tenant", entityId: tenant.id,
        actorId: req.user?.userId, actorName: req.user?.email, after: req.body,
        summary: `Empresa ${tenant.name} actualizada.`, req,
      });
      return ok(res, tenant);
    })
  );

  /** Checklist de onboarding de una empresa. */
  router.get(
    "/tenants/:id/onboarding",
    handler(async (req, res) => {
      const tenant = await getTenantById(req.params.id);
      if (!tenant) throw new ErpHttpError("Empresa no encontrada.", 404, "tenant_not_found");
      const [members, costCenters, catalogs, siigoCompany, documents] = await Promise.all([
        listMembers(tenant.id),
        col(COL.costCenters).countDocuments({ tenantId: tenant.id }),
        col(COL.catalogs).countDocuments({ tenantId: tenant.id }),
        getTenantSiigoCompany(tenant).catch(() => null),
        col(COL.documents).countDocuments({ tenantId: tenant.id }),
      ]);

      const steps = [
        { id: "tenant", label: "Empresa creada", done: true },
        { id: "siigo", label: "Siigo conectado", done: !!siigoCompany },
        { id: "drive", label: "Google Drive conectado", done: !!tenant.settings?.driveRootFolderId },
        { id: "catalogs", label: "Catálogos sincronizados", done: catalogs > 0 },
        { id: "costCenters", label: "Centros de costo", done: costCenters > 0 },
        { id: "users", label: "Usuarios invitados", done: members.length > 0 },
        { id: "documents", label: "Primeros documentos", done: documents > 0 },
        { id: "active", label: "Portal activo", done: tenant.status === "ACTIVE" },
      ];
      return ok(res, { tenant, steps, progress: Math.round((steps.filter((s) => s.done).length / steps.length) * 100) });
    })
  );

  /** Empresas Siigo disponibles para asociar a un tenant. */
  router.get(
    "/siigo-companies",
    handler(async (_req, res) => ok(res, await listCompanies()))
  );

  /** Salud del sistema. Nunca expone secretos. */
  router.get(
    "/health",
    handler(async (_req, res) => {
      const [tenants, documents, errors, pendingEgresos, failedOps] = await Promise.all([
        col(COL.tenants).countDocuments({}),
        col(COL.documents).countDocuments({}),
        col(COL.documents).countDocuments({ status: "POSTING_ERROR" }),
        col(COL.payments).countDocuments({ "siigo.status": { $in: ["FAILED", "UNKNOWN"] } }),
        col(COL.operations).countDocuments({ status: "FAILED" }),
      ]);
      return ok(res, {
        tenants, documents,
        postingErrors: errors,
        pendingEgresos,
        failedOperations: failedOps,
        siigo: getSiigoIntegrationHealth(),
        // Ya no hay cola global de jobs DIAN (se retiró): `jobs.active` son los jobs DIAN
        // en curso en este servidor y `browsers.queued > 0` significa que un job espera
        // cupo de navegador (máx. 2), no que esté fallando.
        dian: { jobs: { active: activeJobCount(), queued: 0 }, browsers: getBrowserStats() },
        uptimeSeconds: Math.round(process.uptime()),
      });
    })
  );

  /** Últimas operaciones de integración de todas las empresas. */
  router.get(
    "/operations",
    handler(async (req, res) => {
      const limit = Math.min(toNumber(req.query.limit) ?? 100, 300);
      const docs = await col(COL.operations).find({}).sort({ startedAt: -1 }).limit(limit).toArray();
      const { fromDocs } = await import("../db.js");
      return ok(res, fromDocs(docs));
    })
  );

  /** Crea/completa la empresa de demostración. */
  router.post(
    "/seed-demo",
    handler(async (req, res) => ok(res, await seedDemoTenant(req.body?.adminEmail || req.user?.email)))
  );

  /** Miembros de una empresa (soporte). */
  router.get(
    "/tenants/:id/members",
    handler(async (req, res) => ok(res, await listMembers(req.params.id)))
  );

  router.post(
    "/tenants/:id/members",
    handler(async (req, res) => {
      const member = await inviteMember(req.params.id, {
        email: String(req.body?.email || ""),
        role: (req.body?.role || "VIEWER") as ErpRole,
        name: req.body?.name,
        invitedById: req.user?.userId,
      });
      return ok(res, member, 201);
    })
  );

  /** Auditoría global (soporte). */
  router.get(
    "/audit",
    handler(async (req, res) => {
      const tenantId = req.query.tenantId ? String(req.query.tenantId) : null;
      const filter = tenantId ? scoped(tenantId) : {};
      const docs = await col(COL.audit).find(filter as any).sort({ createdAt: -1 }).limit(200).toArray();
      const { fromDocs } = await import("../db.js");
      return ok(res, fromDocs(docs));
    })
  );

  return router;
}
