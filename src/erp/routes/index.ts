/**
 * Router raíz del ERP.
 *
 * Orden de montaje (importa: cada capa depende de la anterior):
 *   1. /erp/admin  → backoffice ContaGO (autenticación propia, sin tenant)
 *   2. requireAuth → sesión válida
 *   3. withErpTenant → resuelve empresa + membresía + permisos + contexto Siigo
 *   4. routers de dominio
 */
import { Router, type RequestHandler } from "express";
import { requireAuth } from "../../middleware/auth.js";
import { withErpTenant, slugFromHost } from "../context.js";
import { getTenantBySlug, listMembershipsForEmail, getTenantById } from "../tenants.js";
import { ensureErpIndexes } from "../db.js";
import { buildLoginUrl } from "../auth.js";
import { createDocumentsRouter } from "./documents.js";
import { createAccountingRouter } from "./accounting.js";
import { createTreasuryRouter } from "./treasury.js";
import { createSettingsRouter } from "./settings.js";
import { createAdminRouter } from "./admin.js";
import { erpErrorHandler, handler, ok } from "./helpers.js";

export function createErpRouter(authMiddleware: RequestHandler = requireAuth): Router {
  const router = Router();

  // Backoffice ContaGO: no depende de un tenant.
  router.use("/admin", createAdminRouter());

  // Salud del módulo (público, sin secretos).
  router.get("/health", (_req, res) => {
    res.json({ ok: true, service: "ContaGO ERP", time: new Date().toISOString() });
  });

  /**
   * Resolución de portal previa al login: dice al frontend qué empresa
   * corresponde al hostname. No expone datos sensibles.
   */
  router.get(
    "/portal",
    handler(async (req, res) => {
      await ensureErpIndexes();
      const host = (req.headers["x-forwarded-host"] as string) || req.headers.host;
      const slug = slugFromHost(host) || String(req.query.slug || "");
      const tenant = slug ? await getTenantBySlug(slug) : null;
      if (!tenant) {
        return res.status(404).json({ ok: false, code: "tenant_not_found", message: "Portal no encontrado." });
      }
      return ok(res, {
        slug: tenant.slug,
        name: tenant.name,
        logoUrl: tenant.logoUrl,
        primaryColor: tenant.primaryColor,
        status: tenant.status,
      });
    })
  );

  /**
   * Inicio de sesión con Google para el portal de una empresa.
   * El slug se valida contra la base de datos antes de firmar el state.
   */
  router.get(
    "/auth/google/start",
    handler(async (req, res) => {
      await ensureErpIndexes();
      const host = (req.headers["x-forwarded-host"] as string) || req.headers.host;
      const slug = slugFromHost(host) || String(req.query.slug || "");
      const tenant = slug ? await getTenantBySlug(slug) : null;
      if (!tenant) {
        return res.status(404).json({ ok: false, code: "tenant_not_found", message: "Portal no encontrado." });
      }
      const { url } = buildLoginUrl(tenant.slug);
      if (String(req.query.redirect || "") === "true") return res.redirect(302, url);
      return ok(res, { url });
    })
  );

  // A partir de aquí se exige sesión.
  router.use(authMiddleware);

  /**
   * Empresas del usuario autenticado. Alimenta el selector de app.contago.com.co
   * cuando alguien pertenece a varios portales.
   */
  router.get(
    "/my-tenants",
    handler(async (req, res) => {
      await ensureErpIndexes();
      const email = (req.user?.email || "").toLowerCase();
      const memberships = await listMembershipsForEmail(email);
      const tenants = await Promise.all(
        memberships.map(async (m) => {
          const t = await getTenantById(m.tenantId);
          return t ? { id: t.id, slug: t.slug, name: t.name, logoUrl: t.logoUrl, status: t.status, role: m.role } : null;
        })
      );
      return ok(res, tenants.filter(Boolean));
    })
  );

  // Contexto de empresa para el resto del ERP.
  router.use(withErpTenant);

  router.use("/", createSettingsRouter());
  router.use("/documents", createDocumentsRouter());
  router.use("/accounting", createAccountingRouter());
  router.use("/treasury", createTreasuryRouter());

  router.use(erpErrorHandler);
  return router;
}

export default createErpRouter();
