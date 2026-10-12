/**
 * Contexto de petición del ERP.
 *
 * Resuelve el tenant a partir del HOSTNAME (empresa.contago.com.co). El header
 * `X-Erp-Tenant` solo se admite como respaldo en desarrollo o para usuarios
 * ContaGO; y aun así SIEMPRE se valida contra la base de datos y contra la
 * membresía del usuario. Nunca se confía en un tenantId enviado por el cliente.
 */
import type { NextFunction, Request, Response } from "express";
import { enterSiigoCompany } from "../services/siigoService.js";
import { ensureErpIndexes } from "./db.js";
import {
  activateMembership,
  getMembership,
  getTenantById,
  getTenantBySlug,
  getTenantSiigoContext,
  memberPermissions,
  type ErpMember,
  type ErpTenant,
} from "./tenants.js";
import { permissionsForRole, type Permission } from "./permissions.js";

export interface ErpRequestContext {
  tenant: ErpTenant;
  member: ErpMember | null;
  permissions: Permission[];
  actorId: string;
  actorName: string;
  isContagoStaff: boolean;
  siigoReady: boolean;
}

declare global {
  namespace Express {
    interface Request {
      erp?: ErpRequestContext;
    }
  }
}

export class ErpHttpError extends Error {
  status: number;
  code: string;
  details?: unknown;
  constructor(message: string, status = 400, code = "erp_error", details?: unknown) {
    super(message);
    this.name = "ErpHttpError";
    this.status = status;
    this.code = code;
    this.details = details;
  }
}

const ROOT_DOMAIN = (process.env.ERP_ROOT_DOMAIN || process.env.ROOT_DOMAIN || "contago.com.co").toLowerCase();

/** Extrae el slug del tenant del hostname, si la petición viene de un subdominio. */
export function slugFromHost(host?: string): string | null {
  if (!host) return null;
  const hostname = host.split(":")[0].toLowerCase();
  if (!hostname.endsWith(`.${ROOT_DOMAIN}`)) return null;
  const sub = hostname.slice(0, -(ROOT_DOMAIN.length + 1));
  if (!sub || sub.includes(".")) return null;
  return sub;
}

/**
 * Middleware principal. Debe montarse DESPUÉS de requireAuth.
 * Resuelve tenant + membresía + permisos y fija el contexto Siigo de la empresa.
 */
export async function withErpTenant(req: Request, res: Response, next: NextFunction): Promise<void> {
  try {
    await ensureErpIndexes();

    const user = req.user;
    if (!user) {
      res.status(401).json({ ok: false, code: "unauthorized", message: "Sesión requerida." });
      return;
    }

    // 1) Hostname (fuente preferida y no manipulable por JS del navegador).
    const forwardedHost = (req.headers["x-forwarded-host"] as string | undefined) || req.headers.host;
    const hostSlug = slugFromHost(forwardedHost);

    // 2) Respaldo explícito: header o query. Se valida igual contra membresía.
    const headerTenant = String(req.header("X-Erp-Tenant") || req.query.tenant || "").trim();

    let tenant: ErpTenant | null = null;
    if (hostSlug) tenant = await getTenantBySlug(hostSlug);
    if (!tenant && headerTenant) {
      tenant = (await getTenantBySlug(headerTenant)) || (await getTenantById(headerTenant));
    }

    if (!tenant) {
      res.status(404).json({
        ok: false,
        code: "tenant_not_found",
        message: "No pudimos identificar la empresa de este portal.",
      });
      return;
    }

    if (tenant.status === "SUSPENDED") {
      res.status(403).json({ ok: false, code: "tenant_suspended", message: "Este portal está suspendido." });
      return;
    }

    // 3) Membresía. El staff de ContaGO (isAdmin) entra como SUPER_ADMIN_CONTAGO.
    const email = (user.email || "").toLowerCase();
    const isContagoStaff = !!user.isAdmin;
    const member = await getMembership(tenant.id, email);

    if (!member && !isContagoStaff) {
      res.status(403).json({
        ok: false,
        code: "no_membership",
        message: `${user.email} no tiene acceso a ${tenant.name}. Solicita una invitación al administrador.`,
      });
      return;
    }
    if (member && member.status === "DISABLED") {
      res.status(403).json({ ok: false, code: "membership_disabled", message: "Tu acceso a este portal fue desactivado." });
      return;
    }

    // El staff de ContaGO conserva acceso completo aunque además tenga una
    // membresía de cliente (p. ej. para soporte). La auditoría registra siempre
    // el usuario real, así que el acceso no es anónimo.
    const permissions = isContagoStaff
      ? permissionsForRole("SUPER_ADMIN_CONTAGO", member?.extraPermissions || [])
      : memberPermissions(member!);

    // 4) Contexto Siigo de ESTA empresa. Si no hay credenciales no se cae al
    //    .env global: se marca siigoReady=false y las rutas que lo exijan fallan.
    let siigoReady = false;
    const siigoCtx = await getTenantSiigoContext(tenant);
    if (siigoCtx) {
      enterSiigoCompany(siigoCtx);
      siigoReady = true;
    }

    req.erp = {
      tenant,
      member,
      permissions,
      actorId: user.userId,
      actorName: member?.name || user.email || user.userId,
      isContagoStaff,
      siigoReady,
    };

    if (member && member.status === "INVITED") {
      void activateMembership(tenant.id, email, user.userId);
    }

    next();
  } catch (err) {
    console.error("[ERP][context] error resolviendo tenant:", err);
    res.status(500).json({ ok: false, code: "context_error", message: "No pudimos preparar el portal." });
  }
}

/** Contexto obligatorio: lanza si el middleware no corrió. */
export function requireErp(req: Request): ErpRequestContext {
  if (!req.erp) throw new ErpHttpError("Contexto de empresa no disponible.", 500, "no_context");
  return req.erp;
}

/** Guard de permiso. Toda autorización crítica ocurre aquí, en el servidor. */
export function requirePermission(permission: Permission) {
  return (req: Request, res: Response, next: NextFunction): void => {
    const ctx = req.erp;
    if (!ctx) {
      res.status(500).json({ ok: false, code: "no_context", message: "Contexto no disponible." });
      return;
    }
    if (!ctx.permissions.includes(permission)) {
      res.status(403).json({
        ok: false,
        code: "forbidden",
        message: "No tienes permisos para realizar esta acción.",
        required: permission,
      });
      return;
    }
    next();
  };
}

/** Exige credenciales Siigo configuradas para el tenant. */
export function requireSiigo(req: Request, res: Response, next: NextFunction): void {
  if (!req.erp?.siigoReady) {
    res.status(409).json({
      ok: false,
      code: "siigo_not_connected",
      message: "Siigo no está conectado para esta empresa. Configúralo en Integraciones.",
    });
    return;
  }
  next();
}
