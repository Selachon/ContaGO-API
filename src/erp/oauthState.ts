/**
 * `state` firmado para OAuth de Google.
 *
 * Nunca se acepta un returnUrl libre desde el query string: el destino se
 * reconstruye a partir del slug del tenant, validado contra la base de datos.
 * El state va firmado (HMAC-SHA256) y expira, de modo que no puede fabricarse
 * ni reutilizarse.
 */
import crypto from "node:crypto";

const TTL_MS = 10 * 60 * 1000;

export interface ErpOAuthState {
  purpose: "erp-login" | "erp-drive";
  slug: string;
  tenantId?: string;
  userId?: string;
  ts: number;
  nonce: string;
}

function secret(): string {
  const value = process.env.JWT_SECRET;
  if (!value) throw new Error("JWT_SECRET no configurado.");
  return value;
}

function sign(payload: string): string {
  return crypto.createHmac("sha256", secret()).update(payload).digest("base64url");
}

export function encodeState(data: Omit<ErpOAuthState, "ts" | "nonce">): string {
  const state: ErpOAuthState = { ...data, ts: Date.now(), nonce: crypto.randomBytes(12).toString("base64url") };
  const payload = Buffer.from(JSON.stringify(state)).toString("base64url");
  return `${payload}.${sign(payload)}`;
}

export function decodeState(raw: string): ErpOAuthState | null {
  const [payload, signature] = String(raw || "").split(".");
  if (!payload || !signature) return null;

  const expected = sign(payload);
  const a = Buffer.from(signature);
  const b = Buffer.from(expected);
  if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) return null;

  try {
    const state = JSON.parse(Buffer.from(payload, "base64url").toString("utf8")) as ErpOAuthState;
    if (!state?.purpose || !state.slug) return null;
    if (Date.now() - Number(state.ts || 0) > TTL_MS) return null;
    return state;
  } catch {
    return null;
  }
}

const ROOT_DOMAIN = (process.env.ERP_ROOT_DOMAIN || process.env.ROOT_DOMAIN || "contago.com.co").toLowerCase();

/**
 * URL del portal de un tenant. El slug SIEMPRE proviene de la base de datos,
 * nunca del navegador, así que no puede provocar un open redirect.
 */
export function portalUrl(slug: string, path = "/"): string {
  const override = process.env.ERP_PORTAL_URL_TEMPLATE;
  if (override) return override.replace("{slug}", slug).replace(/\/+$/, "") + path;
  const scheme = process.env.NODE_ENV === "production" ? "https" : "http";
  const port = process.env.NODE_ENV === "production" ? "" : `:${process.env.ERP_DEV_PORT || 5174}`;
  return `${scheme}://${slug}.${ROOT_DOMAIN}${port}${path}`;
}
