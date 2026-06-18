/**
 * Buzón DIAN de ContaGO (info@contago.com.co).
 *
 * Las empresas reenvían a este buzón el correo de la DIAN que contiene el
 * "token" (en realidad una URL de ingreso: catalogo-vpfe.dian.gov.co/User/AuthToken).
 * El NIT del titular viaja en el parámetro `rk`, así que un solo buzón compartido
 * sirve a todas las empresas: enrutamos cada token por NIT.
 *
 * Lectura por IMAP con "contraseña de aplicación" (App Password). Esto evita el
 * trámite de verificación de scopes restringidos de Gmail y funciona con
 * cualquier proveedor IMAP; por defecto Gmail (imap.gmail.com:993).
 */
import { ImapFlow } from "imapflow";
import { simpleParser } from "mailparser";
import { encryptToken, decryptToken } from "../utils/encryption.js";
import { getDb } from "./database.js";

const SETTINGS = "systemSettings";
const MAILBOX_ID = "dianMailbox";

export interface DianTokenHit {
  /** URL completa de ingreso a la DIAN (…/User/AuthToken?…). */
  tokenUrl: string;
  /** NIT del titular, normalizado (solo dígitos), tomado del parámetro `rk`. */
  nit: string;
}

export interface DianMailToken extends DianTokenHit {
  uid: number;
  subject: string;
  from: string;
  /** Fecha de recepción (epoch ms). */
  receivedAt: number;
}

const URL_TOKEN_RE = /https?:\/\/catalogo-vpfe\.dian\.gov\.co\/User\/AuthToken\?[^\s"'<>)\]]+/gi;

function normalizeNit(raw: string): string {
  return String(raw || "").replace(/\D/g, "").replace(/^0+/, "");
}

/**
 * Extrae todas las URLs de token DIAN presentes en un texto (cuerpo de correo,
 * sea plano o HTML). Decodifica `&amp;` y entidades comunes para no romper la URL.
 * Devuelve cada token con su NIT (param `rk`). Deduplica por URL.
 */
export function parseDianTokenUrls(text: string): DianTokenHit[] {
  if (!text) return [];
  const decoded = text
    .replace(/&amp;/gi, "&")
    .replace(/&#38;/g, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"');
  const seen = new Set<string>();
  const out: DianTokenHit[] = [];
  const matches = decoded.match(URL_TOKEN_RE) || [];
  for (const m of matches) {
    const url = m.replace(/[.,;]+$/, ""); // limpia puntuación final pegada
    if (seen.has(url)) continue;
    seen.add(url);
    let nit = "";
    try {
      nit = normalizeNit(new URL(url).searchParams.get("rk") || "");
    } catch {
      const rk = /[?&]rk=([^&]+)/i.exec(url);
      nit = rk ? normalizeNit(decodeURIComponent(rk[1])) : "";
    }
    out.push({ tokenUrl: url, nit });
  }
  return out;
}

// ── Persistencia del buzón (singleton cifrado) ────────────────────────────
interface DianMailboxDoc {
  _id: string;
  email: string;
  host: string;
  port: number;
  encrypted_password: string;
  connectedAt: string;
}

/** Guarda/actualiza las credenciales IMAP del buzón (contraseña cifrada). */
export async function saveDianMailbox(input: {
  email: string;
  password: string;
  host?: string;
  port?: number;
}): Promise<void> {
  await getDb().collection<DianMailboxDoc>(SETTINGS).updateOne(
    { _id: MAILBOX_ID },
    {
      $set: {
        email: input.email,
        host: input.host || "imap.gmail.com",
        port: input.port || 993,
        encrypted_password: encryptToken(input.password),
        connectedAt: new Date().toISOString(),
      },
    },
    { upsert: true },
  );
}

export async function getDianMailbox(): Promise<DianMailboxDoc | null> {
  return getDb().collection<DianMailboxDoc>(SETTINGS).findOne({ _id: MAILBOX_ID });
}

export async function getDianMailboxStatus(): Promise<{ connected: boolean; email: string; host: string; connectedAt: string }> {
  const m = await getDianMailbox();
  return { connected: !!m, email: m?.email || "", host: m?.host || "", connectedAt: m?.connectedAt || "" };
}

export async function deleteDianMailbox(): Promise<void> {
  await getDb().collection<DianMailboxDoc>(SETTINGS).deleteOne({ _id: MAILBOX_ID });
}

// ── Lectura por IMAP ──────────────────────────────────────────────────────
async function openClient(m: DianMailboxDoc): Promise<ImapFlow> {
  const client = new ImapFlow({
    host: m.host || "imap.gmail.com",
    port: m.port || 993,
    secure: true,
    auth: { user: m.email, pass: decryptToken(m.encrypted_password) },
    logger: false,
  });
  await client.connect();
  return client;
}

/** Prueba de conexión: valida credenciales abriendo y cerrando la sesión IMAP. */
export async function testDianMailboxConnection(input: { email: string; password: string; host?: string; port?: number }): Promise<void> {
  const client = new ImapFlow({
    host: input.host || "imap.gmail.com",
    port: input.port || 993,
    secure: true,
    auth: { user: input.email, pass: input.password },
    logger: false,
  });
  await client.connect();
  await client.logout();
}

/**
 * Lee del buzón los correos recientes que contengan un token DIAN y devuelve los
 * tokens encontrados (uno por URL). Opcionalmente filtra por NIT.
 */
export async function fetchDianTokensFromMailbox(opts: { sinceDays?: number; nit?: string; max?: number } = {}): Promise<DianMailToken[]> {
  const m = await getDianMailbox();
  if (!m) throw new Error("DIAN_MAILBOX_NOT_CONNECTED");
  const sinceDays = Math.max(1, opts.sinceDays ?? 7);
  const since = new Date(Date.now() - sinceDays * 24 * 3600 * 1000);
  const wantNit = opts.nit ? normalizeNit(opts.nit) : "";
  const out: DianMailToken[] = [];

  const client = await openClient(m);
  const lock = await client.getMailboxLock("INBOX");
  try {
    // Búsqueda IMAP por remitente de la DIAN + fecha. Gmail no indexa el texto
    // dentro de los enlaces (body:"...AuthToken" devuelve 0), así que filtramos
    // por el `from` (100% fiable) y el parser hace el filtro fino: que el cuerpo
    // contenga realmente una URL /User/AuthToken con su `rk`.
    const uids = await client.search({ since, from: "facturacionelectronica@dian.gov.co" }, { uid: true });
    const limited = (uids || []).slice(-(opts.max ?? 30));
    if (limited.length) {
      for await (const msg of client.fetch(limited, { uid: true, source: true, envelope: true, internalDate: true }, { uid: true })) {
        const parsed = await simpleParser(msg.source as Buffer);
        const body = `${parsed.text || ""}\n${parsed.html || ""}`;
        const subject = parsed.subject || "";
        const from = parsed.from?.text || "";
        const receivedAt = new Date(msg.internalDate || parsed.date || Date.now()).getTime();
        for (const hit of parseDianTokenUrls(body)) {
          if (wantNit && hit.nit !== wantNit) continue;
          out.push({ ...hit, uid: msg.uid, subject, from, receivedAt });
        }
      }
    }
  } finally {
    lock.release();
    await client.logout().catch(() => {});
  }

  // Más reciente primero (el token fresco gana).
  out.sort((a, b) => b.receivedAt - a.receivedAt);
  return out;
}

/** Devuelve el token más reciente para un NIT, o null si no hay. */
export async function latestDianTokenForNit(nit: string, sinceDays = 7): Promise<DianMailToken | null> {
  const all = await fetchDianTokensFromMailbox({ nit, sinceDays });
  return all[0] || null;
}
