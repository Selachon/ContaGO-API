/**
 * Tenants del ERP.
 *
 * Un tenant = una empresa cliente. Se apoya en la colección `siigoCompanies`
 * existente (que ya guarda credenciales Siigo cifradas y la conexión de Drive)
 * y añade la capa de portal: slug/subdominio, branding, módulos y ajustes.
 */
import { COL, col, fromDoc, fromDocs, now, toId, upsertDoc } from "./db.js";
import { getCompanyContext, getCompanyPublic } from "../services/siigoCompaniesService.js";
import { permissionsForRole, type ErpRole, type Permission } from "./permissions.js";
import type { SiigoContext } from "../services/siigoService.js";

/** Subdominios que nunca pueden asignarse a un cliente. */
export const RESERVED_SLUGS = new Set([
  "www", "api", "app", "admin", "auth", "mail", "erp", "portal", "static",
  "cdn", "assets", "status", "help", "docs", "blog", "contago", "test", "dev",
]);

export interface ErpTenantSettings {
  /** Monto desde el cual un documento exige aprobación explícita. */
  approvalThreshold: number;
  requireApproval: boolean;
  /** Auto-enviar a contabilidad al aprobar. */
  autoSendToAccounting: boolean;
  /** Días de plazo por defecto si el documento no trae vencimiento. */
  defaultPaymentTermsDays: number;
  /** Carpeta raíz de la empresa en Drive. */
  driveRootFolderId?: string;
  driveFolderPattern: string;
  /** Comprobantes Siigo por defecto. */
  defaultPurchaseDocumentTypeId?: number;
  defaultSupportDocumentTypeId?: number;
  defaultPaymentReceiptDocumentTypeId?: number;
  defaultPaymentTypeId?: number;
  defaultPayableAccountCode?: string;
  /** Documentos soporte exigidos al radicar una cuenta de cobro. */
  requiredDocuments: string[];
}

export interface ErpTenant {
  id: string;
  slug: string;
  name: string;
  legalName?: string;
  nit: string;
  dv?: string;
  /** Empresa Siigo asociada (colección siigoCompanies). Fuente de credenciales. */
  siigoCompanyId?: string;
  logoUrl?: string;
  primaryColor?: string;
  timezone: string;
  currency: string;
  status: "ACTIVE" | "SUSPENDED" | "ONBOARDING";
  features: Record<string, boolean>;
  settings: ErpTenantSettings;
  createdAt: string;
  updatedAt: string;
}

export const DEFAULT_FEATURES: Record<string, boolean> = {
  documents: true,
  approvals: true,
  accounting: true,
  payables: true,
  payments: true,
  cashflow: true,
  reports: true,
  drive: true,
  siigo: true,
  dian: true,
};

export const DEFAULT_SETTINGS: ErpTenantSettings = {
  approvalThreshold: 0,
  requireApproval: true,
  autoSendToAccounting: true,
  defaultPaymentTermsDays: 30,
  driveFolderPattern: "{year}/{month} - {category}",
  requiredDocuments: [],
};

export function normalizeSlug(value: string): string {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/[^a-z0-9-]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 40);
}

export function validateSlug(slug: string): string | null {
  if (!slug || slug.length < 3) return "El subdominio debe tener al menos 3 caracteres.";
  if (RESERVED_SLUGS.has(slug)) return `El subdominio "${slug}" está reservado.`;
  if (!/^[a-z0-9][a-z0-9-]*[a-z0-9]$/.test(slug)) return "El subdominio solo admite letras, números y guiones.";
  return null;
}

export async function listTenants(): Promise<ErpTenant[]> {
  const docs = await col(COL.tenants).find({}).sort({ name: 1 }).toArray();
  return fromDocs<ErpTenant>(docs);
}

export async function getTenantById(id: string): Promise<ErpTenant | null> {
  try {
    const doc = await col(COL.tenants).findOne({ _id: toId(id) } as any);
    return doc ? fromDoc<ErpTenant>(doc) : null;
  } catch {
    return null;
  }
}

export async function getTenantBySlug(slug: string): Promise<ErpTenant | null> {
  const doc = await col(COL.tenants).findOne({ slug: normalizeSlug(slug) } as any);
  return doc ? fromDoc<ErpTenant>(doc) : null;
}

export interface CreateTenantInput {
  name: string;
  legalName?: string;
  nit: string;
  dv?: string;
  slug: string;
  siigoCompanyId?: string;
  timezone?: string;
  currency?: string;
  logoUrl?: string;
  features?: Record<string, boolean>;
  settings?: Partial<ErpTenantSettings>;
}

export async function createTenant(input: CreateTenantInput): Promise<ErpTenant> {
  const slug = normalizeSlug(input.slug || input.name);
  const slugError = validateSlug(slug);
  if (slugError) throw new Error(slugError);
  if (await getTenantBySlug(slug)) throw new Error(`Ya existe una empresa con el subdominio "${slug}".`);

  const doc = {
    slug,
    name: String(input.name).trim(),
    legalName: input.legalName?.trim(),
    nit: String(input.nit || "").replace(/\D/g, ""),
    dv: input.dv,
    siigoCompanyId: input.siigoCompanyId,
    logoUrl: input.logoUrl,
    timezone: input.timezone || "America/Bogota",
    currency: input.currency || "COP",
    status: "ONBOARDING" as const,
    features: { ...DEFAULT_FEATURES, ...(input.features || {}) },
    settings: { ...DEFAULT_SETTINGS, ...(input.settings || {}) },
    createdAt: now(),
    updatedAt: now(),
  };
  const res = await col(COL.tenants).insertOne(doc as any);
  return { id: String(res.insertedId), ...doc } as ErpTenant;
}

export async function updateTenant(id: string, patch: Partial<ErpTenant>): Promise<ErpTenant | null> {
  const set: Record<string, unknown> = { updatedAt: now() };
  for (const key of ["name", "legalName", "nit", "dv", "siigoCompanyId", "logoUrl", "timezone", "currency", "status", "features", "settings", "primaryColor"] as const) {
    if (patch[key] !== undefined) set[key] = patch[key];
  }
  const res = await col(COL.tenants).findOneAndUpdate(
    { _id: toId(id) } as any,
    { $set: set },
    { returnDocument: "after" }
  );
  return res ? fromDoc<ErpTenant>(res) : null;
}

/** Fusiona ajustes sin perder claves no enviadas. */
export async function patchTenantSettings(id: string, patch: Partial<ErpTenantSettings>): Promise<ErpTenant | null> {
  const tenant = await getTenantById(id);
  if (!tenant) return null;
  return updateTenant(id, { settings: { ...DEFAULT_SETTINGS, ...tenant.settings, ...patch } });
}

/**
 * Contexto Siigo del tenant. Devuelve null si la empresa aún no tiene
 * credenciales — nunca cae silenciosamente al .env global (§106).
 */
export async function getTenantSiigoContext(tenant: ErpTenant): Promise<SiigoContext | null> {
  if (!tenant.siigoCompanyId) return null;
  return getCompanyContext(tenant.siigoCompanyId);
}

export async function getTenantSiigoCompany(tenant: ErpTenant) {
  if (!tenant.siigoCompanyId) return null;
  return getCompanyPublic(tenant.siigoCompanyId);
}

// ─── Membresías ──────────────────────────────────────────────────────────

export interface ErpMember {
  id: string;
  tenantId: string;
  userId?: string;
  email: string;
  name?: string;
  role: ErpRole;
  extraPermissions: Permission[];
  status: "INVITED" | "ACTIVE" | "DISABLED";
  invitedById?: string;
  invitedAt: string;
  acceptedAt?: string;
  lastAccessAt?: string;
}

export async function listMembers(tenantId: string): Promise<ErpMember[]> {
  const docs = await col(COL.members).find({ tenantId }).sort({ email: 1 }).toArray();
  return fromDocs<ErpMember>(docs);
}

export async function getMembership(tenantId: string, email: string): Promise<ErpMember | null> {
  const doc = await col(COL.members).findOne({ tenantId, email: email.toLowerCase().trim() } as any);
  return doc ? fromDoc<ErpMember>(doc) : null;
}

export async function listMembershipsForEmail(email: string): Promise<ErpMember[]> {
  const docs = await col(COL.members)
    .find({ email: email.toLowerCase().trim(), status: { $ne: "DISABLED" } } as any)
    .toArray();
  return fromDocs<ErpMember>(docs);
}

export async function inviteMember(
  tenantId: string,
  input: { email: string; role: ErpRole; name?: string; extraPermissions?: Permission[]; invitedById?: string }
): Promise<ErpMember> {
  const email = input.email.toLowerCase().trim();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error("Correo inválido.");

  // Reinvitar actualiza rol y permisos; los datos de alta solo se fijan al crear.
  // $set y $setOnInsert no pueden compartir campos: Mongo lo rechaza.
  const set: Record<string, unknown> = {
    role: input.role,
    extraPermissions: input.extraPermissions || [],
  };
  if (input.name !== undefined) set.name = input.name;

  await col(COL.members).updateOne(
    { tenantId, email } as any,
    upsertDoc(set, {
      tenantId,
      email,
      status: "INVITED",
      invitedById: input.invitedById,
      invitedAt: now(),
    }) as any,
    { upsert: true }
  );
  const saved = await getMembership(tenantId, email);
  if (!saved) throw new Error("No pudimos crear la membresía.");
  return saved;
}

export async function updateMember(tenantId: string, memberId: string, patch: Partial<ErpMember>): Promise<ErpMember | null> {
  const set: Record<string, unknown> = {};
  for (const key of ["role", "extraPermissions", "status", "name"] as const) {
    if (patch[key] !== undefined) set[key] = patch[key];
  }
  if (!Object.keys(set).length) return null;
  const res = await col(COL.members).findOneAndUpdate(
    { _id: toId(memberId), tenantId } as any,
    { $set: set },
    { returnDocument: "after" }
  );
  return res ? fromDoc<ErpMember>(res) : null;
}

export async function removeMember(tenantId: string, memberId: string): Promise<boolean> {
  const res = await col(COL.members).deleteOne({ _id: toId(memberId), tenantId } as any);
  return res.deletedCount > 0;
}

/** Marca la membresía como activa la primera vez que el usuario entra. */
export async function activateMembership(tenantId: string, email: string, userId: string): Promise<void> {
  await col(COL.members).updateOne(
    { tenantId, email: email.toLowerCase().trim() } as any,
    { $set: { userId, status: "ACTIVE", lastAccessAt: now(), acceptedAt: now() } }
  );
}

export function memberPermissions(member: ErpMember): Permission[] {
  return permissionsForRole(member.role, member.extraPermissions || []);
}
