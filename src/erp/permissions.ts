/**
 * RBAC del ERP. Los permisos son granulares; los roles solo son paquetes de
 * permisos. Nunca se autoriza por nombre de rol en la capa de negocio.
 */

export const PERMISSIONS = [
  "documents.read",
  "documents.create",
  "documents.edit",
  "documents.approve",
  "documents.reject",
  "documents.void",

  "accounting.read",
  "accounting.edit",
  "accounting.post_siigo",

  "payables.read",
  "payables.schedule",
  "payables.pay",
  "payables.post_siigo",

  "cashflow.read",
  "reports.read",

  "settings.read",
  "settings.users",
  "settings.integrations",
  "settings.accounting",

  "audit.read",
  "admin.tenants",
] as const;

export type Permission = (typeof PERMISSIONS)[number];

export type ErpRole =
  | "SUPER_ADMIN_CONTAGO"
  | "CONTAGO_ACCOUNTANT"
  | "CLIENT_ADMIN"
  | "APPROVER"
  | "TREASURY"
  | "VIEWER"
  | "DOCUMENT_SUBMITTER";

export const ROLE_LABELS: Record<ErpRole, string> = {
  SUPER_ADMIN_CONTAGO: "Administrador ContaGO",
  CONTAGO_ACCOUNTANT: "Contador ContaGO",
  CLIENT_ADMIN: "Administrador de la empresa",
  APPROVER: "Aprobador",
  TREASURY: "Tesorería",
  VIEWER: "Consulta",
  DOCUMENT_SUBMITTER: "Radicador de documentos",
};

const ALL: Permission[] = [...PERMISSIONS];

export const ROLE_PERMISSIONS: Record<ErpRole, Permission[]> = {
  SUPER_ADMIN_CONTAGO: ALL,

  CONTAGO_ACCOUNTANT: [
    "documents.read", "documents.create", "documents.edit", "documents.void",
    "accounting.read", "accounting.edit", "accounting.post_siigo",
    "payables.read", "payables.post_siigo",
    "cashflow.read", "reports.read",
    "settings.read", "settings.accounting", "settings.integrations",
    "audit.read",
  ],

  CLIENT_ADMIN: [
    "documents.read", "documents.create", "documents.edit",
    "documents.approve", "documents.reject",
    "accounting.read",
    "payables.read", "payables.schedule",
    "cashflow.read", "reports.read",
    "settings.read", "settings.users", "settings.integrations",
    "audit.read",
  ],

  APPROVER: [
    "documents.read", "documents.approve", "documents.reject",
    "payables.read", "reports.read",
  ],

  TREASURY: [
    "documents.read",
    "payables.read", "payables.schedule", "payables.pay", "payables.post_siigo",
    "cashflow.read", "reports.read",
  ],

  VIEWER: ["documents.read", "payables.read", "cashflow.read", "reports.read"],

  DOCUMENT_SUBMITTER: ["documents.read", "documents.create"],
};

export function permissionsForRole(role: ErpRole, extra: Permission[] = []): Permission[] {
  const base = ROLE_PERMISSIONS[role] ?? [];
  return Array.from(new Set([...base, ...extra]));
}

export function hasPermission(granted: readonly string[], needed: Permission): boolean {
  return granted.includes(needed);
}
