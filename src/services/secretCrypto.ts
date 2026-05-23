import crypto from "node:crypto";

// Clave de cifrado: idealmente SIIGO_ENC_KEY; si no, se deriva del JWT_SECRET.
const KEY = crypto.scryptSync(
  process.env.SIIGO_ENC_KEY || process.env.JWT_SECRET || "contago-siigo-fallback-key",
  "contago-siigo-salt",
  32
);

/** Cifra un secreto (AES-256-GCM). Formato: iv:tag:ciphertext en base64. */
export function encryptSecret(plain: string): string {
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv("aes-256-gcm", KEY, iv);
  const enc = Buffer.concat([cipher.update(String(plain), "utf8"), cipher.final()]);
  const tag = cipher.getAuthTag();
  return `${iv.toString("base64")}:${tag.toString("base64")}:${enc.toString("base64")}`;
}

/** Descifra un secreto producido por encryptSecret. */
export function decryptSecret(stored: string): string {
  const parts = String(stored).split(":");
  if (parts.length !== 3) throw new Error("Secreto cifrado con formato inválido");
  const [ivB, tagB, encB] = parts;
  const decipher = crypto.createDecipheriv("aes-256-gcm", KEY, Buffer.from(ivB, "base64"));
  decipher.setAuthTag(Buffer.from(tagB, "base64"));
  return Buffer.concat([decipher.update(Buffer.from(encB, "base64")), decipher.final()]).toString("utf8");
}
