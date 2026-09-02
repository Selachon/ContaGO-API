import bcrypt from "bcrypt";
import { createHash, randomBytes } from "crypto";
import { MongoClient, type Db, type Collection, ObjectId } from "mongodb";
import type { DemoAccess, User, UserRole } from "../types/auth.js";
import type { GoogleDriveConfig } from "../types/dianExcel.js";

export const DEMO_TRIAL_HOURS = Number(process.env.DEMO_TRIAL_HOURS || 72);
export const DEMO_TRIAL_LIMIT = Number(process.env.DEMO_TRIAL_LIMIT || 30);
export const DEMO_INVITE_TTL_DAYS = Number(process.env.DEMO_INVITE_TTL_DAYS || 14);

interface DemoInviteRecord {
  _id: ObjectId;
  tokenHash: string;
  nit?: string;
  normalizedNit?: string;
  toolId: string;
  createdBy: string;
  createdAt: string;
  expiresAt: string;
  usedAt?: string;
  usedByUserId?: string;
  status?: "active" | "used" | "expired";
}

interface DemoNitTrialRecord {
  _id: string;
  nit: string;
  normalizedNit: string;
  toolId: string;
  inviteId: string;
  createdBy: string;
  createdAt: string;
  status: "reserved" | "consumed";
  userId?: string;
  consumedAt?: string;
}

export interface DemoInviteListItem {
  id: string;
  nit?: string;
  normalizedNit?: string;
  toolId: string;
  createdBy: string;
  createdAt: string;
  expiresAt: string;
  usedAt?: string;
  usedByUserId?: string;
  status: "active" | "used" | "expired";
}

export interface DemoInviteLookup {
  ok: boolean;
  reason?: "invalid" | "used" | "expired";
  invite?: DemoInviteListItem;
}

interface UserRecord {
  _id: ObjectId;
  email: string;
  name: string;
  password_hash: string;
  is_admin: boolean;
  role?: UserRole;
  demo?: DemoAccess;
  purchasedTools: string[];
  nits: string[];
  status?: "active" | "suspended";
  force_password_change?: boolean;
  created_at: string;
  legacyId?: number;
  google_drive?: GoogleDriveConfig;
  google_drives?: GoogleDriveConfig[];
  selected_google_drive_id?: string | null;
  phone?: string;
  paymentAmount?: number;
  paymentMethod?: string;
  licenseStartDate?: string;
  licenseEndDate?: string;
  companiesInPlan?: number;
  toolCompanyLimits?: Record<string, number>;
  invoiceRef?: string;
  siigoCompanies?: string[];
  // Activación de la extensión de navegador (descargador DIAN).
  // El usuario genera el código UNA vez desde el portal; la extensión lo canjea
  // por un token de larga duración. Un admin puede restablecerlo para regenerar.
  ext_activation?: {
    code: string | null;          // código en claro; null = nunca generado / restablecido
    codeGeneratedAt?: string;     // ISO — cuándo se generó
    activatedAt?: string;         // ISO — flujo legacy (un solo dispositivo, pre-multi-device)
    deviceLabel?: string;         // etiqueta del dispositivo (flujo legacy)
  };
  // Dispositivos registrados vía el flujo multi-device (máx. 2).
  ext_devices?: ExtDevice[];
  // Preferencias del Exportador DIAN que se configuran en el portal y la extensión
  // lee para seguirlas tal cual (Drive sí/no, qué cuenta, incluir enlaces).
  exporter_prefs?: {
    uploadToDrive: boolean;
    includeDriveLinks: boolean;
    driveConnectionId: string | null;
  };
}


let client: MongoClient | null = null;
let db: Db | null = null;

function usersCollection(): Collection<UserRecord> {
  if (!db) {
    throw new Error("MongoDB no está conectado");
  }
  return db.collection<UserRecord>("users");
}

function demoInvitesCollection(): Collection<DemoInviteRecord> {
  if (!db) {
    throw new Error("MongoDB no está conectado");
  }
  return db.collection<DemoInviteRecord>("demo_invites");
}

function demoNitTrialsCollection(): Collection<DemoNitTrialRecord> {
  if (!db) {
    throw new Error("MongoDB no está conectado");
  }
  return db.collection<DemoNitTrialRecord>("demo_nit_trials");
}

async function ensureIndex(
  collection: { collectionName: string; createIndex: (keys: Record<string, 1 | -1>, options?: Parameters<Collection["createIndex"]>[1]) => Promise<string> },
  keys: Record<string, 1 | -1>,
  options: Parameters<Collection["createIndex"]>[1] = {}
): Promise<void> {
  try {
    await collection.createIndex(keys, options);
  } catch (err) {
    const code = (err as { code?: number }).code;
    if (code === 14031) {
      console.warn(
        "[Mongo] Índice omitido por espacio insuficiente en Railway: " + collection.collectionName + " " + JSON.stringify(keys)
      );
      return;
    }
    if (code === 11000) {
      console.warn(
        "[Mongo] Índice ÚNICO no creado por duplicados existentes: " + collection.collectionName + " " + JSON.stringify(keys)
      );
      return;
    }
    throw err;
  }
}

export function getDb(): Db {
  if (!db) {
    throw new Error("MongoDB no está conectado");
  }
  return db;
}

export async function connectMongo(): Promise<void> {
  const uri = process.env.MONGODB_URI;
  const dbName = process.env.MONGODB_DB;

  if (!uri) {
    throw new Error("MONGODB_URI no está definido");
  }

  client = new MongoClient(uri);
  await client.connect();

  db = dbName ? client.db(dbName) : client.db();

  // Índices
  const users = usersCollection();
  await ensureIndex(users, { email: 1 }, { unique: true });
  await ensureIndex(users, { legacyId: 1 });
  await ensureIndex(users, { role: 1 });
  // Búsqueda por código de activación de la extensión (único, solo donde existe).
  await ensureIndex(users, { "ext_activation.code": 1 }, { unique: true, sparse: true });
  const demoInvites = demoInvitesCollection();
  await ensureIndex(demoInvites, { tokenHash: 1 }, { unique: true });
  await ensureIndex(demoInvites, { normalizedNit: 1 });
  await ensureIndex(demoInvites, { createdAt: -1 });

  // Índices del dominio Siigo (aislamiento por usuario/empresa). Tolerantes a
  // falta de espacio (Railway) y a duplicados preexistentes.
  const siigoCompanies = db.collection("siigoCompanies");
  await ensureIndex(siigoCompanies, { ownerUserId: 1 });
  await ensureIndex(siigoCompanies, { sharedWith: 1 });
  await ensureIndex(siigoCompanies, { username: 1 }, { unique: true });
  const siigoProfiles = db.collection("siigoSupplierProfiles");
  await ensureIndex(siigoProfiles, { companyId: 1 });
  const siigoIngestedCufes = db.collection("siigoIngestedCufes");
  await ensureIndex(siigoIngestedCufes, { companyId: 1, cufe: 1 }, { unique: true });
  await ensureIndex(siigoIngestedCufes, { companyId: 1, status: 1, fetchedAt: -1 });
  const siigoEgresoMovs = db.collection("siigoEgresoMovs");
  await ensureIndex(siigoEgresoMovs, { companyId: 1, createdAt: -1 });
  await ensureIndex(siigoEgresoMovs, { companyId: 1, bankAccountId: 1, createdAt: -1 });
  await ensureIndex(siigoEgresoMovs, { companyId: 1, status: 1, fingerprint: 1 });
  const siigoBankAccounts = db.collection("siigoBankAccounts");
  await ensureIndex(siigoBankAccounts, { companyId: 1, createdAt: 1 });
}

function mapUser(record: UserRecord | null): User | null {
  if (!record) return null;
  return {
    id: record._id.toString(),
    email: record.email,
    name: record.name,
    password_hash: record.password_hash,
    is_admin: record.is_admin,
    role: record.role || (record.is_admin ? "ADMIN" : "USER"),
    nits: record.nits || [],
    status: record.status || "active",
    force_password_change: !!record.force_password_change,
    created_at: record.created_at,
    demo: record.demo,
    companiesInPlan: record.companiesInPlan,
    toolCompanyLimits: record.toolCompanyLimits,
    licenseStartDate: record.licenseStartDate,
  };
}

// ============================================
// Demo trial functions
// ============================================

export function normalizeNitForDemo(nit: string): string {
  return String(nit || "").replace(/[^0-9A-Za-z]/g, "").toUpperCase().trim();
}

function hashDemoInviteToken(token: string): string {
  return createHash("sha256").update(token).digest("hex");
}

function addHours(date: Date, hours: number): Date {
  return new Date(date.getTime() + hours * 60 * 60 * 1000);
}

function addDays(date: Date, days: number): Date {
  return new Date(date.getTime() + days * 24 * 60 * 60 * 1000);
}

function mapDemoInvite(record: DemoInviteRecord): DemoInviteListItem {
  const expired = !record.usedAt && new Date(record.expiresAt).getTime() <= Date.now();
  return {
    id: record._id.toString(),
    nit: record.nit,
    normalizedNit: record.normalizedNit,
    toolId: record.toolId,
    createdBy: record.createdBy,
    createdAt: record.createdAt,
    expiresAt: record.expiresAt,
    usedAt: record.usedAt,
    usedByUserId: record.usedByUserId,
    status: record.usedAt ? "used" : expired ? "expired" : "active",
  };
}

export function isDemoTrialExpired(demo?: DemoAccess | null): boolean {
  if (!demo?.expiresAt) return false;
  return new Date(demo.expiresAt).getTime() <= Date.now();
}

export async function createDemoInvite(toolId: string, createdBy: string): Promise<{ token: string; invite: DemoInviteListItem } | null> {
  const cleanToolId = String(toolId || "").trim();
  if (!cleanToolId) return null;

  const inviteId = new ObjectId();
  const token = randomBytes(32).toString("hex");
  const now = new Date();
  const nowIso = now.toISOString();
  const expiresAt = addDays(now, DEMO_INVITE_TTL_DAYS).toISOString();

  const inviteRecord: DemoInviteRecord = {
    _id: inviteId,
    tokenHash: hashDemoInviteToken(token),
    toolId: cleanToolId,
    createdBy,
    createdAt: nowIso,
    expiresAt,
    status: "active",
  };

  await demoInvitesCollection().insertOne(inviteRecord);
  return { token, invite: mapDemoInvite(inviteRecord) };
}

export async function listDemoInvites(limit = 25): Promise<DemoInviteListItem[]> {
  const records = await demoInvitesCollection()
    .find({})
    .sort({ createdAt: -1 })
    .limit(Math.min(Math.max(limit, 1), 100))
    .toArray();
  return records.map(mapDemoInvite);
}

export async function getDemoInviteByToken(token: string): Promise<DemoInviteLookup> {
  const tokenHash = hashDemoInviteToken(String(token || ""));
  const record = await demoInvitesCollection().findOne({ tokenHash });
  if (!record) return { ok: false, reason: "invalid" };

  const invite = mapDemoInvite(record);
  if (invite.status === "used") return { ok: false, reason: "used", invite };
  if (invite.status === "expired") return { ok: false, reason: "expired", invite };
  return { ok: true, invite };
}

export async function consumeDemoInvite(
  token: string,
  nit: string,
  email: string,
  name: string,
  password: string
): Promise<{ ok: true; user: User } | { ok: false; reason: "invalid" | "used" | "expired" | "invalid_nit" | "nit_used" | "email_exists" | "create_failed" }> {
  const lookup = await getDemoInviteByToken(token);
  if (!lookup.ok || !lookup.invite) {
    return { ok: false, reason: lookup.reason || "invalid" };
  }

  const cleanNit = String(nit || "").trim();
  const normalizedNit = normalizeNitForDemo(cleanNit);
  if (!cleanNit || normalizedNit.length < 5) return { ok: false, reason: "invalid_nit" };

  const existing = await getUserByEmail(email);
  if (existing) return { ok: false, reason: "email_exists" };

  const now = new Date();
  const nowIso = now.toISOString();
  const trialRecord: DemoNitTrialRecord = {
    _id: normalizedNit,
    nit: cleanNit,
    normalizedNit,
    toolId: lookup.invite.toolId,
    inviteId: lookup.invite.id,
    createdBy: lookup.invite.createdBy,
    createdAt: nowIso,
    status: "reserved",
  };

  try {
    await demoNitTrialsCollection().insertOne(trialRecord);
  } catch (err) {
    if ((err as { code?: number }).code === 11000) return { ok: false, reason: "nit_used" };
    throw err;
  }

  const claimed = await demoInvitesCollection().findOneAndUpdate(
    { _id: new ObjectId(lookup.invite.id), usedAt: { $exists: false } },
    { $set: { usedAt: nowIso, status: "used", nit: cleanNit, normalizedNit } },
    { returnDocument: "after" }
  );

  if (!claimed) {
    await demoNitTrialsCollection().deleteOne({ _id: normalizedNit, inviteId: lookup.invite.id, status: "reserved" });
    return { ok: false, reason: "used" };
  }

  const trialExpiresAt = addHours(now, DEMO_TRIAL_HOURS).toISOString();
  const demo: DemoAccess = {
    nit: cleanNit,
    normalizedNit,
    toolId: lookup.invite.toolId,
    inviteId: lookup.invite.id,
    startedAt: nowIso,
    expiresAt: trialExpiresAt,
    trialLimit: DEMO_TRIAL_LIMIT,
  };

  const user = await createUser(
    email,
    name,
    password,
    false,
    [cleanNit],
    [lookup.invite.toolId],
    {
      role: "DEMO",
      demo,
      licenseStartDate: nowIso.slice(0, 10),
      licenseEndDate: trialExpiresAt.slice(0, 10),
      companiesInPlan: 1,
    }
  );

  if (!user) {
    await Promise.all([
      demoInvitesCollection().updateOne(
        { _id: new ObjectId(lookup.invite.id) },
        { $unset: { usedAt: "", status: "", nit: "", normalizedNit: "" } }
      ),
      demoNitTrialsCollection().deleteOne({ _id: normalizedNit, inviteId: lookup.invite.id, status: "reserved" }),
    ]);
    return { ok: false, reason: "create_failed" };
  }

  await Promise.all([
    demoInvitesCollection().updateOne(
      { _id: new ObjectId(lookup.invite.id) },
      { $set: { usedByUserId: user.id, status: "used", nit: cleanNit, normalizedNit } }
    ),
    demoNitTrialsCollection().updateOne(
      { _id: normalizedNit, inviteId: lookup.invite.id },
      { $set: { status: "consumed", userId: user.id, consumedAt: nowIso } }
    ),
  ]);

  return { ok: true, user };
}

export async function getUserDemoAccess(userId: string): Promise<DemoAccess | null> {
  try {
    const oid = new ObjectId(userId);
    const record = await usersCollection().findOne({ _id: oid }, { projection: { role: 1, demo: 1, is_admin: 1 } });
    const role = record?.role || (record?.is_admin ? "ADMIN" : "USER");
    if (role !== "DEMO" || !record?.demo) return null;
    return record.demo;
  } catch {
    return null;
  }
}

// ============================================
// User functions
// ============================================

export interface UserExtras {
  role?: UserRole;
  demo?: DemoAccess;
  phone?: string;
  paymentAmount?: number;
  paymentMethod?: string;
  licenseStartDate?: string;
  licenseEndDate?: string;
  companiesInPlan?: number;
  toolCompanyLimits?: Record<string, number>;
  invoiceRef?: string;
  siigoCompanies?: string[];
}

export async function createUser(
  email: string,
  name: string,
  password: string,
  isAdmin = false,
  nits: string[] = [],
  purchasedTools: string[] = [],
  extras: UserExtras = {}
): Promise<User | null> {
  try {
    const hash = await bcrypt.hash(password, 10);
    const record: UserRecord = {
      _id: new ObjectId(),
      email: email.toLowerCase().trim(),
      name: name.trim(),
      password_hash: hash,
      is_admin: isAdmin,
      purchasedTools,
      nits,
      created_at: new Date().toISOString(),
      force_password_change: false,
      ...extras,
      role: isAdmin ? "ADMIN" : (extras.role || "USER"),
    };

    await usersCollection().insertOne(record);
    return mapUser(record);
  } catch (err) {
    console.error("Error creating user:", err);
    return null;
  }
}

export async function updateUserPassword(userId: string, plainPassword: string, forcePasswordChange: boolean): Promise<boolean> {
  try {
    const oid = new ObjectId(userId);
    const hash = await bcrypt.hash(plainPassword, 10);
    const result = await usersCollection().updateOne(
      { _id: oid },
      {
        $set: {
          password_hash: hash,
          force_password_change: forcePasswordChange,
          updated_at: new Date().toISOString(),
        },
      }
    );
    return result.modifiedCount > 0 || result.matchedCount > 0;
  } catch (err) {
    console.error("Error updating user password:", err);
    return false;
  }
}

export async function getUserById(id: string): Promise<User | null> {
  try {
    const oid = new ObjectId(id);
    const record = await usersCollection().findOne({ _id: oid });
    return mapUser(record);
  } catch {
    return null;
  }
}

// Version estricta: distingue "no encontrado" de errores de infraestructura.
export async function getUserByIdStrict(id: string): Promise<User | null> {
  if (!ObjectId.isValid(id)) {
    return null;
  }
  const oid = new ObjectId(id);
  const record = await usersCollection().findOne({ _id: oid });
  return mapUser(record);
}

export async function getUserByEmail(email: string): Promise<User | null> {
  const record = await usersCollection().findOne({ email: email.toLowerCase().trim() });
  return mapUser(record);
}

export async function verifyPassword(user: User, password: string): Promise<boolean> {
  return bcrypt.compare(password, user.password_hash);
}

// ============================================
// Purchase functions
// ============================================

// Maps decommissioned tool IDs to the new ones that replace them.
// When a user has the old ID, the new ID is automatically included in their purchased tools.
export const TOOL_SUCCESSOR: Record<string, string> = {
  "dian-downloader": "dian-mass-download",
  "dian-excel-exporter": "dian-cufe-downloader",
  "dian-mass-download": "dian-recibidos",
};

export async function getUserPurchases(userId: string): Promise<string[]> {
  try {
    const oid = new ObjectId(userId);
    const record = await usersCollection().findOne({ _id: oid }, { projection: { purchasedTools: 1 } });
    const tools: string[] = record?.purchasedTools || [];
    // Expand decommissioned IDs to their successors so existing users keep access
    const expanded = new Set(tools);
    for (const tool of tools) {
      const successor = TOOL_SUCCESSOR[tool];
      if (successor) expanded.add(successor);
    }
    return [...expanded];
  } catch {
    return [];
  }
}

export async function addPurchase(userId: string, toolId: string): Promise<boolean> {
  try {
    const oid = new ObjectId(userId);
    const result = await usersCollection().updateOne(
      { _id: oid },
      { $addToSet: { purchasedTools: toolId } }
    );
    return result.modifiedCount > 0;
  } catch {
    return false;
  }
}

export async function hasPurchase(userId: string, toolId: string, aliases: string[] = []): Promise<boolean> {
  try {
    const oid = new ObjectId(userId);
    const idsToCheck = aliases.length > 0 ? [toolId, ...aliases] : [toolId];
    const record = await usersCollection().findOne({ _id: oid, purchasedTools: { $in: idsToCheck } });
    return !!record;
  } catch {
    return false;
  }
}

// ============================================
// NIT functions
// ============================================

export async function getUserNits(userId: string): Promise<string[]> {
  try {
    const oid = new ObjectId(userId);
    const record = await usersCollection().findOne({ _id: oid }, { projection: { nits: 1 } });
    return record?.nits || [];
  } catch {
    return [];
  }
}

// ============================================
// Empresas Siigo asignadas a un usuario
// ============================================

export async function getUserSiigoCompanies(userId: string): Promise<string[]> {
  try {
    const oid = new ObjectId(userId);
    const record = await usersCollection().findOne({ _id: oid }, { projection: { siigoCompanies: 1 } });
    return record?.siigoCompanies || [];
  } catch {
    return [];
  }
}

export async function setUserSiigoCompanies(userId: string, companyIds: string[]): Promise<boolean> {
  try {
    const oid = new ObjectId(userId);
    const clean = Array.from(new Set(companyIds.filter((c) => typeof c === "string" && c.trim())));
    const result = await usersCollection().updateOne({ _id: oid }, { $set: { siigoCompanies: clean } });
    return result.matchedCount > 0;
  } catch {
    return false;
  }
}

// ============================================
// Seed admin user (env-based, no hardcoded password)
// ============================================

// ============================================
// Google Drive functions
// ============================================

export async function getUserGoogleDrive(userId: string): Promise<GoogleDriveConfig | null> {
  return getUserGoogleDriveById(userId);
}

export async function getUserGoogleDrives(userId: string): Promise<GoogleDriveConfig[]> {
  try {
    const oid = new ObjectId(userId);
    const record = await usersCollection().findOne(
      { _id: oid },
      { projection: { google_drive: 1, google_drives: 1 } }
    );
    if (record?.google_drives && record.google_drives.length > 0) {
      return record.google_drives;
    }
    return record?.google_drive ? [record.google_drive] : [];
  } catch {
    return [];
  }
}

export async function getUserGoogleDriveById(
  userId: string,
  connectionId?: string
): Promise<GoogleDriveConfig | null> {
  try {
    const oid = new ObjectId(userId);
    const record = await usersCollection().findOne(
      { _id: oid },
      { projection: { google_drive: 1, google_drives: 1, selected_google_drive_id: 1 } }
    );

    const drives = (record?.google_drives && record.google_drives.length > 0)
      ? record.google_drives
      : (record?.google_drive ? [record.google_drive] : []);

    if (drives.length === 0) return null;
    if (connectionId) return drives.find((d) => d.connection_id === connectionId) || null;

    const selectedId = record?.selected_google_drive_id;
    if (selectedId) {
      const selected = drives.find((d) => d.connection_id === selectedId);
      if (selected) return selected;
    }
    return drives[0] || null;
  } catch {
    return null;
  }
}

export async function updateUserGoogleDrive(
  userId: string,
  driveConfig: GoogleDriveConfig
): Promise<boolean> {
  try {
    const oid = new ObjectId(userId);
    const existingDrives = await getUserGoogleDrives(userId);
    const nextDrives = existingDrives.filter((d) => d.connection_id !== driveConfig.connection_id);
    nextDrives.push(driveConfig);

    const result = await usersCollection().updateOne(
      { _id: oid },
      {
        $set: {
          google_drives: nextDrives,
          google_drive: driveConfig,
          selected_google_drive_id: driveConfig.connection_id,
        },
      }
    );
    return result.modifiedCount > 0 || result.matchedCount > 0;
  } catch (err) {
    console.error("Error actualizando Google Drive config:", err);
    return false;
  }
}

export async function updateUserDriveTokens(
  userId: string,
  encryptedAccessToken: string,
  tokenExpiry: string,
  connectionId?: string
): Promise<boolean> {
  try {
    const oid = new ObjectId(userId);
    if (connectionId) {
      const drives = await getUserGoogleDrives(userId);
      const idx = drives.findIndex((d) => d.connection_id === connectionId);
      if (idx === -1) return false;
      drives[idx] = {
        ...drives[idx],
        encrypted_access_token: encryptedAccessToken,
        token_expiry: tokenExpiry,
        last_used: new Date().toISOString(),
      };
      const selected = await getUserGoogleDriveById(userId);
      const result = await usersCollection().updateOne(
        { _id: oid },
        {
          $set: {
            google_drives: drives,
            ...(selected?.connection_id === connectionId ? { google_drive: drives[idx] } : {}),
          },
        }
      );
      return result.modifiedCount > 0 || result.matchedCount > 0;
    }

    const result = await usersCollection().updateOne(
      { _id: oid },
      {
        $set: {
          "google_drive.encrypted_access_token": encryptedAccessToken,
          "google_drive.token_expiry": tokenExpiry,
          "google_drive.last_used": new Date().toISOString(),
        },
      }
    );
    return result.modifiedCount > 0;
  } catch {
    return false;
  }
}

export async function updateUserDriveFolder(
  userId: string,
  folderId: string,
  folderName: string
): Promise<boolean> {
  try {
    const oid = new ObjectId(userId);
    const result = await usersCollection().updateOne(
      { _id: oid },
      {
        $set: {
          "google_drive.folder_id": folderId,
          "google_drive.folder_name": folderName,
        },
      }
    );
    return result.modifiedCount > 0;
  } catch {
    return false;
  }
}

export async function removeUserGoogleDrive(userId: string): Promise<boolean> {
  return removeUserGoogleDriveById(userId);
}

export async function setSelectedUserGoogleDrive(userId: string, connectionId: string): Promise<boolean> {
  try {
    const oid = new ObjectId(userId);
    const drives = await getUserGoogleDrives(userId);
    const selected = drives.find((d) => d.connection_id === connectionId);
    if (!selected) return false;
    const result = await usersCollection().updateOne(
      { _id: oid },
      {
        $set: {
          selected_google_drive_id: connectionId,
          google_drive: selected,
        },
      }
    );
    return result.modifiedCount > 0 || result.matchedCount > 0;
  } catch {
    return false;
  }
}

export async function removeUserGoogleDriveById(userId: string, connectionId?: string): Promise<boolean> {
  try {
    const oid = new ObjectId(userId);
    if (connectionId) {
      const drives = await getUserGoogleDrives(userId);
      const nextDrives = drives.filter((d) => d.connection_id !== connectionId);
      const selectedId = nextDrives[0]?.connection_id || null;
      const selectedDrive = nextDrives[0] || null;
      const result = await usersCollection().updateOne(
        { _id: oid },
        {
          $set: {
            google_drives: nextDrives,
            selected_google_drive_id: selectedId,
            ...(selectedDrive ? { google_drive: selectedDrive } : {}),
          },
          ...(selectedDrive ? {} : { $unset: { google_drive: "" } }),
        }
      );
      return result.modifiedCount > 0 || result.matchedCount > 0;
    }

    const result = await usersCollection().updateOne(
      { _id: oid },
      { $unset: { google_drive: "", google_drives: "", selected_google_drive_id: "" } }
    );
    return result.modifiedCount > 0;
  } catch {
    return false;
  }
}

// ============================================
// Seed admin user (env-based, no hardcoded password)
// ============================================

export async function seedAdminUser(): Promise<void> {
  const adminEmail = process.env.ADMIN_EMAIL;
  const adminPassword = process.env.ADMIN_PASSWORD;
  const adminName = process.env.ADMIN_NAME || "Admin";
  const adminNits = process.env.ADMIN_NITS
    ? process.env.ADMIN_NITS.split(",").map((n) => n.trim()).filter(Boolean)
    : [];

  if (!adminEmail || !adminPassword) {
    console.log("ADMIN_EMAIL / ADMIN_PASSWORD not set - skipping admin seed.");
    return;
  }

  const existing = await getUserByEmail(adminEmail);

  if (!existing) {
    console.log(`Creating admin user: ${adminEmail}`);
    await createUser(adminEmail, adminName, adminPassword, true, adminNits);
  } else {
    // Ensure admin flag and sync NITs from env
    const updates: Record<string, unknown> = {};
    if (!existing.is_admin) updates.is_admin = true;
    if (adminNits.length > 0) updates.nits = adminNits;

    if (Object.keys(updates).length > 0) {
      await usersCollection().updateOne(
        { email: adminEmail.toLowerCase().trim() },
        { $set: updates }
      );
      if (updates.is_admin) console.log(`Promoted ${adminEmail} to admin.`);
      if (updates.nits) console.log(`Updated NITs for ${adminEmail}: ${adminNits.join(", ")}`);
    }
  }
}

export async function migrateToolSlugs(): Promise<void> {
  try {
    for (const [oldSlug, newSlug] of Object.entries(TOOL_SUCCESSOR)) {
      await usersCollection().updateMany(
        { purchasedTools: oldSlug },
        [
          {
            $set: {
              purchasedTools: {
                $setUnion: [
                  { $filter: { input: "$purchasedTools", cond: { $ne: ["$$this", oldSlug] } } },
                  [newSlug],
                ],
              },
              updated_at: new Date().toISOString(),
            },
          },
        ]
      );
    }
    console.log("[DB] Tool slug migration completed");
  } catch (err) {
    console.error("[DB] Tool slug migration failed:", err);
  }
}

// ============================================
// Activación de la extensión de navegador (DIAN)
// ============================================

// Código legible y sin caracteres ambiguos (sin 0/O/1/I/L). Formato CTG-XXXX-XXXX.
const EXT_CODE_ALPHABET = "ABCDEFGHJKMNPQRSTUVWXYZ23456789";
function makeExtCode(): string {
  const block = () =>
    Array.from({ length: 4 }, () => EXT_CODE_ALPHABET[Math.floor(Math.random() * EXT_CODE_ALPHABET.length)]).join("");
  return `CTG-${block()}-${block()}`;
}

export interface ExtDevice {
  deviceId: string;
  deviceLabel: string;
  activatedAt: string;
}

export interface ExtActivationStatus {
  generated: boolean;
  activated: boolean;
  code?: string | null;          // solo si hay código pendiente (aún no canjeado)
  codeGeneratedAt?: string;
  activatedAt?: string;
  deviceLabel?: string;
  deviceCount: number;
  maxDevices: number;
  devices: Pick<ExtDevice, "deviceId" | "deviceLabel" | "activatedAt">[];
}

export async function getExtActivationStatus(userId: string): Promise<ExtActivationStatus | null> {
  try {
    const record = await usersCollection().findOne({ _id: new ObjectId(userId) });
    if (!record) return null;
    const ext = record.ext_activation;
    const devices: ExtDevice[] = (record.ext_devices || []).map(d => ({
      deviceId: d.deviceId,
      deviceLabel: d.deviceLabel || "",
      activatedAt: d.activatedAt,
    }));
    // El "activatedAt" legacy cuenta como 1 dispositivo (flujo antiguo).
    const legacyActivated = !!(ext?.activatedAt);
    const deviceCount = devices.length + (legacyActivated ? 1 : 0);
    // Código pendiente = existe y no fue activado aún (flujo legacy setea activatedAt en vez de limpiar).
    const hasPendingCode = !!(ext?.code && !ext?.activatedAt);
    return {
      generated: hasPendingCode,
      activated: deviceCount > 0,
      code: hasPendingCode ? ext!.code : null,
      codeGeneratedAt: ext?.codeGeneratedAt,
      activatedAt: ext?.activatedAt,
      deviceLabel: ext?.deviceLabel,
      deviceCount,
      maxDevices: 4,
      devices,
    };
  } catch {
    return null;
  }
}

// Genera un código de activación. Si ya hay uno pendiente → { already: true }.
// Si el usuario ya tiene 4 dispositivos registrados → { deviceLimit: true }.
export async function generateExtActivationCode(
  userId: string
): Promise<{ ok: boolean; already?: boolean; deviceLimit?: boolean; code?: string; status?: ExtActivationStatus }> {
  try {
    const oid = new ObjectId(userId);
    const record = await usersCollection().findOne({ _id: oid });
    if (!record) return { ok: false };

    // Límite de 4 dispositivos (el activatedAt legacy cuenta como 1).
    const legacyActivated = !!(record.ext_activation?.activatedAt);
    const deviceCount = (record.ext_devices?.length || 0) + (legacyActivated ? 1 : 0);
    if (deviceCount >= 4) return { ok: true, deviceLimit: true };

    // Código ya pendiente (no canjeado aún): devolver sin generar otro.
    if (record.ext_activation?.code && !record.ext_activation?.activatedAt) {
      return { ok: true, already: true, status: (await getExtActivationStatus(userId)) || undefined };
    }

    // Generar código nuevo (reemplaza el objeto ext_activation completo, limpiando el activatedAt legacy).
    for (let attempt = 0; attempt < 5; attempt++) {
      const code = makeExtCode();
      try {
        await usersCollection().updateOne(
          { _id: oid },
          { $set: { ext_activation: { code, codeGeneratedAt: new Date().toISOString() } } }
        );
        return { ok: true, code };
      } catch (e: any) {
        if (e?.code === 11000) continue; // colisión de índice único → otro código
        throw e;
      }
    }
    return { ok: false };
  } catch (err) {
    console.error("Error generando código de activación:", err);
    return { ok: false };
  }
}

// Canjea el código: registra un nuevo dispositivo y borra el código pendiente.
// Devuelve null si el código no existe o ya fue usado. Devuelve { deviceLimit: true }
// si el usuario ya alcanzó el máximo de 4 dispositivos.
export async function activateExtensionByCode(
  code: string,
  deviceLabel?: string
): Promise<{ user: User; deviceId: string } | { deviceLimit: true } | null> {
  try {
    const clean = String(code || "").trim().toUpperCase();
    if (!clean) return null;
    const record = await usersCollection().findOne({ "ext_activation.code": clean });
    if (!record) return null;
    if (record.status === "suspended") return null;
    // Flujo legacy: el código ya fue canjeado (activatedAt set) → un solo uso.
    if (record.ext_activation?.activatedAt) return null;

    // Verificar límite de dispositivos.
    const deviceCount = record.ext_devices?.length || 0;
    if (deviceCount >= 4) return { deviceLimit: true };

    // Registrar nuevo dispositivo y limpiar el código pendiente.
    const deviceId = randomBytes(16).toString("hex");
    const device: ExtDevice = {
      deviceId,
      deviceLabel: (deviceLabel || "").slice(0, 120),
      activatedAt: new Date().toISOString(),
    };
    await usersCollection().updateOne(
      { _id: record._id },
      {
        $push: { ext_devices: device as any },
        $unset: { "ext_activation.code": "", "ext_activation.codeGeneratedAt": "" },
      }
    );
    return { user: mapUser(record)!, deviceId };
  } catch (err) {
    console.error("Error activando extensión:", err);
    return null;
  }
}

// Elimina un dispositivo del registro (llamado desde la extensión al desvincularse).
export async function removeExtDevice(userId: string, deviceId: string): Promise<boolean> {
  try {
    await usersCollection().updateOne(
      { _id: new ObjectId(userId) },
      { $pull: { ext_devices: { deviceId } } as any }
    );
    return true;
  } catch (err) {
    console.error("Error eliminando dispositivo:", err);
    return false;
  }
}

// ── Preferencias del Exportador DIAN (portal ↔ extensión) ─────────────────────
export interface ExporterPrefs {
  uploadToDrive: boolean;
  includeDriveLinks: boolean;
  driveConnectionId: string | null;
}

export async function getExporterPrefs(userId: string): Promise<ExporterPrefs> {
  try {
    const rec = await usersCollection().findOne({ _id: new ObjectId(userId) });
    const p = rec?.exporter_prefs;
    return {
      uploadToDrive: !!p?.uploadToDrive,
      includeDriveLinks: !!p?.includeDriveLinks,
      driveConnectionId: p?.driveConnectionId ?? null,
    };
  } catch {
    return { uploadToDrive: false, includeDriveLinks: false, driveConnectionId: null };
  }
}

export async function setExporterPrefs(userId: string, prefs: Partial<ExporterPrefs>): Promise<ExporterPrefs> {
  const cur = await getExporterPrefs(userId);
  const next: ExporterPrefs = {
    uploadToDrive: prefs.uploadToDrive ?? cur.uploadToDrive,
    includeDriveLinks: prefs.includeDriveLinks ?? cur.includeDriveLinks,
    driveConnectionId: prefs.driveConnectionId !== undefined ? prefs.driveConnectionId : cur.driveConnectionId,
  };
  await usersCollection().updateOne({ _id: new ObjectId(userId) }, { $set: { exporter_prefs: next } });
  return next;
}

// Admin: restablece la activación y borra todos los dispositivos registrados.
export async function resetExtActivation(userId: string): Promise<boolean> {
  try {
    const result = await usersCollection().updateOne(
      { _id: new ObjectId(userId) },
      { $unset: { ext_activation: "" }, $set: { ext_devices: [] } }
    );
    return result.matchedCount > 0;
  } catch (err) {
    console.error("Error restableciendo activación:", err);
    return false;
  }
}

export { db };
