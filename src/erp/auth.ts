/**
 * Autenticación del ERP con Google.
 *
 * Dos conceptos separados (§6):
 *   - erp-login  → identifica a la persona (openid/email/profile).
 *   - erp-drive  → autoriza el almacenamiento documental del TENANT (drive.file).
 *
 * Ambos usan el callback central ya registrado en Google Cloud
 * (`${API_URL}/auth/google/callback`); no se registran callbacks por cliente.
 */
import crypto from "node:crypto";
import jwt from "jsonwebtoken";
import { exchangeCodeForTokens, getAuthUrl, encryptDriveTokens, getGoogleUserEmail } from "../services/googleDrive.js";
import { createUser, getUserByEmail, updateUserGoogleDrive } from "../services/database.js";
import { decodeState, encodeState, portalUrl } from "./oauthState.js";
import { DRIVE_SCOPES, linkTenantDrive } from "./drive.js";
import { activateMembership, getMembership, getTenantById, getTenantBySlug } from "./tenants.js";
import { audit } from "./audit.js";

const LOGIN_SCOPES = ["openid", "email", "profile"];

export interface ErpAuthUrlResult {
  url: string;
}

/** URL para iniciar sesión en el portal de una empresa. */
export function buildLoginUrl(slug: string): ErpAuthUrlResult {
  const state = encodeState({ purpose: "erp-login", slug });
  return { url: getAuthUrl(state, LOGIN_SCOPES) };
}

/** URL para autorizar Drive a nombre del tenant. */
export function buildDriveUrl(slug: string, tenantId: string, userId: string): ErpAuthUrlResult {
  const state = encodeState({ purpose: "erp-drive", slug, tenantId, userId });
  return { url: getAuthUrl(state, DRIVE_SCOPES) };
}

export interface LoginCallbackResult {
  ok: boolean;
  message: string;
  redirectTo: string;
}

/**
 * Cierre del login: valida el state firmado, comprueba que el correo tenga
 * membresía en el tenant y devuelve la URL del portal con un token de sesión.
 */
export async function handleErpLoginCallback(rawState: string, code: string): Promise<LoginCallbackResult> {
  const state = decodeState(rawState);
  if (!state || state.purpose !== "erp-login") {
    return { ok: false, message: "Solicitud de inicio de sesión inválida.", redirectTo: "" };
  }

  // El destino se reconstruye desde la base de datos, jamás desde el navegador.
  const tenant = await getTenantBySlug(state.slug);
  if (!tenant) {
    return { ok: false, message: "El portal solicitado no existe.", redirectTo: "" };
  }
  if (tenant.status === "SUSPENDED") {
    return { ok: false, message: "Este portal está suspendido. Contacta a ContaGO.", redirectTo: "" };
  }

  const { accessToken } = await exchangeCodeForTokens(code);
  const email = (await getGoogleUserEmail(accessToken)).toLowerCase().trim();

  let user = await getUserByEmail(email);
  const membership = await getMembership(tenant.id, email);

  if (!membership && !user?.is_admin) {
    return {
      ok: false,
      message: `${email} no tiene acceso al portal de ${tenant.name}. Solicita una invitación al administrador.`,
      redirectTo: "",
    };
  }

  // Usuario invitado que aún no existe en ContaGO: se crea sin contraseña
  // utilizable (el acceso es exclusivamente por Google).
  if (!user) {
    user = await createUser(
      email,
      membership?.name || email.split("@")[0],
      crypto.randomBytes(32).toString("hex"),
      false,
      [],
      []
    );
    if (!user) {
      return { ok: false, message: "No pudimos crear tu usuario. Contacta a ContaGO.", redirectTo: "" };
    }
  }

  if (membership) await activateMembership(tenant.id, email, user.id);

  const token = jwt.sign(
    { userId: user.id, email: user.email, isAdmin: !!user.is_admin },
    process.env.JWT_SECRET!,
    { expiresIn: "7d" }
  );

  await audit({
    tenantId: tenant.id,
    action: "USER_LOGGED_IN",
    entityType: "member",
    entityId: membership?.id || user.id,
    actorId: user.id,
    actorName: email,
    summary: `${email} inició sesión en el portal.`,
  });

  const url = new URL(portalUrl(tenant.slug, "/auth/callback"));
  url.searchParams.set("token", token);
  return { ok: true, message: "Sesión iniciada.", redirectTo: url.toString() };
}

export interface DriveCallbackResult {
  ok: boolean;
  message: string;
}

/**
 * Cierre de la autorización de Drive: guarda los tokens cifrados en el usuario
 * autorizante y vincula esa conexión a la empresa.
 */
export async function handleErpDriveCallback(rawState: string, code: string): Promise<DriveCallbackResult> {
  const state = decodeState(rawState);
  if (!state || state.purpose !== "erp-drive" || !state.tenantId || !state.userId) {
    return { ok: false, message: "Solicitud de conexión inválida." };
  }

  const tenant = await getTenantById(state.tenantId);
  if (!tenant) return { ok: false, message: "La empresa ya no existe." };

  const { accessToken, refreshToken, expiryDate } = await exchangeCodeForTokens(code);
  const userEmail = await getGoogleUserEmail(accessToken);
  // Los refresh tokens nunca se guardan en texto plano.
  const encrypted = encryptDriveTokens(accessToken, refreshToken);

  const connectionId = crypto.randomUUID();
  const saved = await updateUserGoogleDrive(state.userId, {
    connection_id: connectionId,
    encrypted_access_token: encrypted.encrypted_access_token,
    encrypted_refresh_token: encrypted.encrypted_refresh_token,
    token_expiry: new Date(expiryDate).toISOString(),
    folder_id: "",
    folder_name: `ContaGO ERP - ${tenant.name}`,
    connected_at: new Date().toISOString(),
    last_used: new Date().toISOString(),
    user_email: userEmail,
  });

  if (!saved) return { ok: false, message: "No pudimos guardar la conexión de Drive." };

  await linkTenantDrive(tenant, state.userId, connectionId);

  await audit({
    tenantId: tenant.id,
    action: "DRIVE_CONNECTED",
    entityType: "integration",
    entityId: "drive",
    actorId: state.userId,
    actorName: userEmail,
    after: { account: userEmail },
    summary: `Google Drive conectado con la cuenta ${userEmail}.`,
  });

  return {
    ok: true,
    message: `Google Drive conectado (${userEmail}). Los soportes de ${tenant.name} se guardarán en esa cuenta.`,
  };
}
