/**
 * Conexión de Google Sheets de una empresa de Caja (Círculo): el OAuth de la
 * cuenta Google DE LA EMPRESA con permiso `spreadsheets`, guardada por empresa
 * (no ligada a un usuario). Base para sincronizar el portal → Google Sheet.
 * El `state` va firmado (HMAC) y expira para que no pueda fabricarse.
 */
import crypto from "node:crypto";
import { google } from "googleapis";
import { getDb } from "./database.js";
import { getAuthUrl, exchangeCodeForTokens, getGoogleUserEmail, createOAuth2Client } from "./googleDrive.js";
import { encryptToken, decryptToken } from "../utils/encryption.js";

const COL = "cajaErpSheets";
const TTL_MS = 10 * 60 * 1000;
const SHEETS_SCOPES = [
  "https://www.googleapis.com/auth/spreadsheets",
  "https://www.googleapis.com/auth/userinfo.email",
];

const secret = () => {
  const v = process.env.JWT_SECRET;
  if (!v) throw new Error("JWT_SECRET no configurado.");
  return v;
};
const sign = (p: string) => crypto.createHmac("sha256", secret()).update(p).digest("base64url");

export function buildSheetsAuthUrl(companyId: string, userId: string): string {
  const payload = Buffer.from(JSON.stringify({
    purpose: "caja-sheets", companyId, userId, ts: Date.now(), nonce: crypto.randomBytes(12).toString("base64url"),
  })).toString("base64url");
  return getAuthUrl(`${payload}.${sign(payload)}`, SHEETS_SCOPES);
}

export function decodeSheetsState(raw: string): { companyId: string; userId: string } | null {
  const [payload, sig] = String(raw || "").split(".");
  if (!payload || !sig) return null;
  const a = Buffer.from(sig), b = Buffer.from(sign(payload));
  if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) return null;
  try {
    const s = JSON.parse(Buffer.from(payload, "base64url").toString("utf8"));
    if (s.purpose !== "caja-sheets" || Date.now() - Number(s.ts || 0) > TTL_MS) return null;
    return { companyId: String(s.companyId), userId: String(s.userId) };
  } catch { return null; }
}

/** Callback OAuth: intercambia el código y guarda los tokens cifrados de la empresa. */
export async function saveSheetsConnection(companyId: string, userId: string, code: string): Promise<{ email: string }> {
  const { accessToken, refreshToken, expiryDate } = await exchangeCodeForTokens(code);
  const email = await getGoogleUserEmail(accessToken);
  const prev = await getDb().collection<any>(COL).findOne({ companyId });
  await getDb().collection<any>(COL).updateOne(
    { companyId },
    { $set: {
        companyId, email, connectedBy: userId,
        accessToken: encryptToken(accessToken),
        refreshToken: encryptToken(refreshToken || (prev?.refreshToken ? decryptToken(prev.refreshToken) : "")),
        expiry: new Date(expiryDate).toISOString(),
        connectedAt: new Date().toISOString(),
    } },
    { upsert: true },
  );
  return { email };
}

export async function getSheetsStatus(companyId: string) {
  const d = await getDb().collection<any>(COL).findOne({ companyId });
  return {
    connected: !!d?.refreshToken, email: d?.email || "", spreadsheetId: d?.spreadsheetId || "", spreadsheetTitle: d?.spreadsheetTitle || "",
    syncUrl: d?.syncEnabled && d?.spreadsheetId ? `https://docs.google.com/spreadsheets/d/${d.spreadsheetId}/edit#gid=${d.syncSheetGid ?? 0}` : "", lastSyncAt: d?.lastSyncAt || "", lastSyncError: d?.lastSyncError || "",
  };
}

export async function removeSheetsConnection(companyId: string): Promise<void> {
  await getDb().collection<any>(COL).deleteOne({ companyId });
}

/** Cliente de Sheets autenticado con la cuenta de la empresa (refresca el token y lo persiste). */
export async function getSheetsClient(companyId: string) {
  const d = await getDb().collection<any>(COL).findOne({ companyId });
  if (!d?.refreshToken) return null;
  const oauth = createOAuth2Client();
  oauth.setCredentials({
    access_token: d.accessToken ? decryptToken(d.accessToken) : undefined,
    refresh_token: decryptToken(d.refreshToken),
    expiry_date: d.expiry ? new Date(d.expiry).getTime() : undefined,
  });
  oauth.on("tokens", (t) => {
    if (!t.access_token) return;
    void getDb().collection<any>(COL).updateOne({ companyId }, { $set: {
      accessToken: encryptToken(t.access_token), expiry: new Date(t.expiry_date || Date.now() + 3300_000).toISOString(),
    } });
  });
  return google.sheets({ version: "v4", auth: oauth });
}

export function parseSpreadsheetId(input: string): string {
  const m = String(input || "").match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : String(input || "").trim();
}

/** Guarda el Sheet destino tras verificar que la cuenta puede abrirlo. */
export async function setSpreadsheet(companyId: string, input: string): Promise<{ id: string; title: string; hojas: string[] }> {
  const id = parseSpreadsheetId(input);
  if (!id) throw new Error("Falta el enlace o ID del Sheet.");
  const sheets = await getSheetsClient(companyId);
  if (!sheets) throw new Error("Primero autoriza Google Sheets.");
  const r = await sheets.spreadsheets.get({ spreadsheetId: id, fields: "properties.title,sheets.properties.title" });
  const title = r.data.properties?.title || "";
  await getDb().collection<any>(COL).updateOne({ companyId }, { $set: { spreadsheetId: id, spreadsheetTitle: title } });
  return { id, title, hojas: (r.data.sheets || []).map((s) => s.properties?.title || "") };
}
