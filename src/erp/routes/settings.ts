/**
 * API de configuración por empresa: general, usuarios, integraciones,
 * centros de costo, notificaciones y auditoría.
 */
import { Router } from "express";
import { requirePermission, requireSiigo, ErpHttpError } from "../context.js";
import {
  inviteMember, listMembers, memberPermissions, patchTenantSettings, removeMember,
  updateMember, updateTenant,
} from "../tenants.js";
import { ROLE_LABELS, ROLE_PERMISSIONS, PERMISSIONS, type ErpRole } from "../permissions.js";
import {
  addSubCostCenter, createCostCenter, deleteCostCenter, getAllCatalogs, listCostCentersLocal,
  removeSubCostCenter, syncAllCatalogs, updateCostCenter,
} from "../catalogs.js";
import { getDriveStatus, linkTenantDrive } from "../drive.js";
import { buildDriveUrl } from "../auth.js";
import { getTenantSiigoCompany } from "../tenants.js";
import { authenticateWithSiigo, getSiigoIntegrationHealth } from "../../services/siigoService.js";
import { listAudit, listOperations } from "../audit.js";
import { listNotifications, markAllRead, markNotificationRead } from "../notifications.js";
import { REPORTS, reportToCsv, reportToExcel, runReport, type ReportId } from "../reports.js";
import { getDashboard } from "../dashboard.js";
import { audit } from "../audit.js";
import { handler, ok, ctxOf, toDate, toNumber } from "./helpers.js";

export function createSettingsRouter(): Router {
  const router = Router();

  // ── Contexto del portal (lo primero que pide el frontend) ──────────────
  router.get(
    "/me",
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const [drive, siigoCompany] = await Promise.all([
        ctx.tenant.features?.drive ? getDriveStatus(ctx.tenant).catch(() => null) : null,
        getTenantSiigoCompany(ctx.tenant).catch(() => null),
      ]);
      return ok(res, {
        tenant: {
          id: ctx.tenant.id,
          slug: ctx.tenant.slug,
          name: ctx.tenant.name,
          legalName: ctx.tenant.legalName,
          nit: ctx.tenant.nit,
          logoUrl: ctx.tenant.logoUrl,
          timezone: ctx.tenant.timezone,
          currency: ctx.tenant.currency,
          status: ctx.tenant.status,
          features: ctx.tenant.features,
          settings: ctx.tenant.settings,
        },
        user: {
          id: ctx.actorId,
          name: ctx.actorName,
          email: ctx.member?.email || req.user?.email,
          role: ctx.member?.role || "SUPER_ADMIN_CONTAGO",
          roleLabel: ROLE_LABELS[(ctx.member?.role || "SUPER_ADMIN_CONTAGO") as ErpRole],
          permissions: ctx.permissions,
          isContagoStaff: ctx.isContagoStaff,
        },
        integrations: {
          siigo: {
            status: ctx.siigoReady ? "CONNECTED" : "DISCONNECTED",
            company: siigoCompany?.name,
            health: ctx.isContagoStaff ? getSiigoIntegrationHealth() : undefined,
          },
          drive: drive || { status: "DISCONNECTED", scopes: [] },
        },
      });
    })
  );

  router.get(
    "/dashboard",
    requirePermission("documents.read"),
    handler(async (req, res) => ok(res, await getDashboard(ctxOf(req).tenant.id)))
  );

  // ── Empresa ────────────────────────────────────────────────────────────
  router.put(
    "/tenant",
    requirePermission("settings.read"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      if (!ctx.permissions.includes("settings.users") && !ctx.isContagoStaff) {
        throw new ErpHttpError("No tienes permisos para modificar la empresa.", 403, "forbidden");
      }
      const before = ctx.tenant;
      const tenant = await updateTenant(ctx.tenant.id, {
        name: req.body?.name,
        legalName: req.body?.legalName,
        logoUrl: req.body?.logoUrl,
        timezone: req.body?.timezone,
        currency: req.body?.currency,
      });
      await audit({
        tenantId: ctx.tenant.id, action: "TENANT_UPDATED", entityType: "tenant", entityId: ctx.tenant.id,
        actorId: ctx.actorId, actorName: ctx.actorName,
        before: { name: before.name }, after: { name: tenant?.name },
        summary: `Datos de la empresa actualizados por ${ctx.actorName}.`, req,
      });
      return ok(res, tenant);
    })
  );

  router.put(
    "/tenant/settings",
    requirePermission("settings.accounting"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const tenant = await patchTenantSettings(ctx.tenant.id, req.body || {});
      await audit({
        tenantId: ctx.tenant.id, action: "TENANT_SETTINGS_UPDATED", entityType: "tenant", entityId: ctx.tenant.id,
        actorId: ctx.actorId, actorName: ctx.actorName, after: req.body,
        summary: `Configuración actualizada por ${ctx.actorName}.`, req,
      });
      return ok(res, tenant);
    })
  );

  // ── Usuarios y roles ───────────────────────────────────────────────────
  router.get(
    "/members",
    requirePermission("settings.read"),
    handler(async (req, res) => {
      const members = await listMembers(ctxOf(req).tenant.id);
      return ok(res, members.map((m) => ({ ...m, permissions: memberPermissions(m) })));
    })
  );

  router.get(
    "/roles",
    requirePermission("settings.read"),
    handler(async (_req, res) =>
      ok(res, {
        roles: Object.entries(ROLE_LABELS).map(([id, label]) => ({
          id, label, permissions: ROLE_PERMISSIONS[id as ErpRole],
        })),
        permissions: PERMISSIONS,
      })
    )
  );

  router.post(
    "/members",
    requirePermission("settings.users"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const member = await inviteMember(ctx.tenant.id, {
        email: String(req.body?.email || ""),
        role: (req.body?.role || "VIEWER") as ErpRole,
        name: req.body?.name,
        extraPermissions: req.body?.extraPermissions,
        invitedById: ctx.actorId,
      });
      await audit({
        tenantId: ctx.tenant.id, action: "USER_INVITED", entityType: "member", entityId: member.id,
        actorId: ctx.actorId, actorName: ctx.actorName, after: { email: member.email, role: member.role },
        summary: `${ctx.actorName} invitó a ${member.email} como ${ROLE_LABELS[member.role]}.`, req,
      });
      return ok(res, member, 201);
    })
  );

  router.put(
    "/members/:id",
    requirePermission("settings.users"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const member = await updateMember(ctx.tenant.id, req.params.id, req.body || {});
      if (!member) throw new ErpHttpError("Miembro no encontrado.", 404, "member_not_found");
      await audit({
        tenantId: ctx.tenant.id, action: "ROLE_CHANGED", entityType: "member", entityId: member.id,
        actorId: ctx.actorId, actorName: ctx.actorName, after: { role: member.role, status: member.status },
        summary: `${ctx.actorName} actualizó el acceso de ${member.email}.`, req,
      });
      return ok(res, member);
    })
  );

  router.delete(
    "/members/:id",
    requirePermission("settings.users"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const removed = await removeMember(ctx.tenant.id, req.params.id);
      await audit({
        tenantId: ctx.tenant.id, action: "USER_REMOVED", entityType: "member", entityId: req.params.id,
        actorId: ctx.actorId, actorName: ctx.actorName, summary: `${ctx.actorName} retiró un usuario del portal.`, req,
      });
      return ok(res, { deleted: removed });
    })
  );

  // ── Centros y subcentros de costo ──────────────────────────────────────
  router.get(
    "/cost-centers",
    requirePermission("settings.read"),
    handler(async (req, res) => ok(res, await listCostCentersLocal(ctxOf(req).tenant.id)))
  );

  router.post(
    "/cost-centers",
    requirePermission("settings.accounting"),
    handler(async (req, res) =>
      ok(res, await createCostCenter(ctxOf(req).tenant.id, {
        code: String(req.body?.code || ""),
        name: String(req.body?.name || ""),
        siigoId: toNumber(req.body?.siigoId),
      }), 201)
    )
  );

  router.put(
    "/cost-centers/:id",
    requirePermission("settings.accounting"),
    handler(async (req, res) => ok(res, await updateCostCenter(ctxOf(req).tenant.id, req.params.id, req.body || {})))
  );

  router.delete(
    "/cost-centers/:id",
    requirePermission("settings.accounting"),
    handler(async (req, res) => ok(res, { deleted: await deleteCostCenter(ctxOf(req).tenant.id, req.params.id) }))
  );

  router.post(
    "/cost-centers/:id/sub",
    requirePermission("settings.accounting"),
    handler(async (req, res) =>
      ok(res, await addSubCostCenter(ctxOf(req).tenant.id, req.params.id, {
        code: String(req.body?.code || ""),
        name: String(req.body?.name || ""),
      }), 201)
    )
  );

  router.delete(
    "/cost-centers/:id/sub/:subId",
    requirePermission("settings.accounting"),
    handler(async (req, res) =>
      ok(res, await removeSubCostCenter(ctxOf(req).tenant.id, req.params.id, req.params.subId))
    )
  );

  // ── Integraciones ──────────────────────────────────────────────────────
  router.get(
    "/integrations",
    requirePermission("settings.read"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const [drive, siigoCompany, catalogs, operations] = await Promise.all([
        getDriveStatus(ctx.tenant).catch((err) => ({ status: "ERROR" as const, scopes: [], message: String(err?.message || err) })),
        getTenantSiigoCompany(ctx.tenant).catch(() => null),
        getAllCatalogs(ctx.tenant.id),
        listOperations(ctx.tenant.id, 25),
      ]);
      return ok(res, {
        drive,
        siigo: {
          status: ctx.siigoReady ? "CONNECTED" : "DISCONNECTED",
          company: siigoCompany?.name,
          nit: siigoCompany?.nit,
          catalogs: Object.entries(catalogs).map(([kind, rec]) => ({
            kind, count: rec.entries?.length || 0, lastSyncedAt: rec.lastSyncedAt, error: rec.error,
          })),
        },
        dian: { status: ctx.tenant.features?.dian ? "CONNECTED" : "DISCONNECTED" },
        recentOperations: operations,
      });
    })
  );

  router.post(
    "/integrations/siigo/test",
    requirePermission("settings.integrations"),
    requireSiigo,
    handler(async (_req, res) => ok(res, await authenticateWithSiigo()))
  );

  router.post(
    "/integrations/siigo/sync",
    requirePermission("settings.integrations"),
    requireSiigo,
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const results = await syncAllCatalogs(ctx.tenant.id);
      await audit({
        tenantId: ctx.tenant.id, action: "SIIGO_SYNCED", entityType: "integration", entityId: "siigo",
        actorId: ctx.actorId, actorName: ctx.actorName, after: results,
        summary: `Catálogos de Siigo sincronizados por ${ctx.actorName}.`, req,
      });
      return ok(res, results);
    })
  );

  /** URL de autorización incremental de Drive (state firmado, expira en 10 min). */
  router.get(
    "/integrations/drive/auth-url",
    requirePermission("settings.integrations"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      return ok(res, buildDriveUrl(ctx.tenant.slug, ctx.tenant.id, ctx.actorId));
    })
  );

  /** Vincula al tenant una conexión de Drive ya autorizada por el usuario. */
  router.post(
    "/integrations/drive/link",
    requirePermission("settings.integrations"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const connectionId = String(req.body?.connectionId || "");
      if (!connectionId) throw new ErpHttpError("Falta la conexión de Drive a vincular.", 400, "connection_required");
      await linkTenantDrive(ctx.tenant, ctx.actorId, connectionId);
      await audit({
        tenantId: ctx.tenant.id, action: "DRIVE_CONNECTED", entityType: "integration", entityId: "drive",
        actorId: ctx.actorId, actorName: ctx.actorName,
        summary: `${ctx.actorName} conectó Google Drive al portal.`, req,
      });
      return ok(res, await getDriveStatus(ctx.tenant));
    })
  );

  // ── Notificaciones ─────────────────────────────────────────────────────
  router.get(
    "/notifications",
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      return ok(res, await listNotifications(ctx.tenant.id, {
        userId: ctx.actorId,
        role: ctx.member?.role,
        unreadOnly: req.query.unread === "true",
      }));
    })
  );

  router.post(
    "/notifications/:id/read",
    handler(async (req, res) => {
      await markNotificationRead(ctxOf(req).tenant.id, req.params.id);
      return ok(res, { read: true });
    })
  );

  router.post(
    "/notifications/read-all",
    handler(async (req, res) => {
      await markAllRead(ctxOf(req).tenant.id);
      return ok(res, { read: true });
    })
  );

  // ── Auditoría ──────────────────────────────────────────────────────────
  router.get(
    "/audit",
    requirePermission("audit.read"),
    handler(async (req, res) =>
      ok(res, await listAudit(ctxOf(req).tenant.id, {
        entityType: req.query.entityType ? String(req.query.entityType) : undefined,
        entityId: req.query.entityId ? String(req.query.entityId) : undefined,
        action: req.query.action ? String(req.query.action) : undefined,
        actorId: req.query.actorId ? String(req.query.actorId) : undefined,
        limit: toNumber(req.query.limit),
      }))
    )
  );

  // ── Reportes ───────────────────────────────────────────────────────────
  router.get(
    "/reports",
    requirePermission("reports.read"),
    handler(async (_req, res) => ok(res, REPORTS))
  );

  router.get(
    "/reports/:id",
    requirePermission("reports.read"),
    handler(async (req, res) => {
      const ctx = ctxOf(req);
      const id = req.params.id as ReportId;
      if (!REPORTS.some((r) => r.id === id)) throw new ErpHttpError("Reporte no encontrado.", 404, "report_not_found");

      const params = {
        from: toDate(req.query.from),
        to: toDate(req.query.to),
        costCenterId: toNumber(req.query.costCenterId),
        supplierNit: req.query.supplierNit ? String(req.query.supplierNit) : undefined,
      };
      const result = await runReport(ctx.tenant.id, id, params);
      const format = String(req.query.format || "json");

      if (format === "xlsx") {
        const buffer = await reportToExcel(result);
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.setHeader("Content-Disposition", `attachment; filename="${id}-${ctx.tenant.slug}.xlsx"`);
        return res.send(buffer);
      }
      if (format === "csv") {
        res.setHeader("Content-Type", "text/csv; charset=utf-8");
        res.setHeader("Content-Disposition", `attachment; filename="${id}-${ctx.tenant.slug}.csv"`);
        return res.send(`﻿${reportToCsv(result)}`);
      }
      return ok(res, result);
    })
  );

  return router;
}
