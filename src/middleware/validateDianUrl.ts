import { Request, Response, NextFunction } from "express";

const DIAN_ALLOWED_ORIGINS = [
  "https://catalogo-vpfe.dian.gov.co",
  "https://catalogo-vpfe-hab.dian.gov.co", // ambiente de habilitación
];

function tryExtractDianUrlFromSafeLink(rawUrl: string): string {
  let parsed: URL;
  try {
    parsed = new URL(rawUrl);
  } catch {
    return rawUrl;
  }

  const host = parsed.hostname.toLowerCase();
  if (!host.endsWith("safelinks.protection.outlook.com")) {
    return rawUrl;
  }

  const wrappedUrl = parsed.searchParams.get("url");
  if (!wrappedUrl) return rawUrl;

  try {
    const candidate = new URL(wrappedUrl);
    return candidate.toString();
  } catch {
    return rawUrl;
  }
}

/**
 * Valida que token_url apunte exclusivamente al dominio de la DIAN.
 * Bloquea SSRF impidiendo URLs a otros hosts o protocolos no HTTPS.
 */
export function validateDianUrl(req: Request, res: Response, next: NextFunction): void {
  const rawTokenUrl = req.body?.token_url;
  const token_url = typeof rawTokenUrl === "string" ? tryExtractDianUrlFromSafeLink(rawTokenUrl.trim()) : rawTokenUrl;

  if (!token_url || typeof token_url !== "string") {
    res.status(400).json({ status: "error", detalle: "token_url es requerido" });
    return;
  }

  let parsed: URL;
  try {
    parsed = new URL(token_url);
  } catch {
    res.status(400).json({ status: "error", detalle: "token_url no es una URL válida" });
    return;
  }

  if (parsed.protocol !== "https:") {
    res.status(400).json({ status: "error", detalle: "token_url debe usar HTTPS" });
    return;
  }

  const origin = `${parsed.protocol}//${parsed.hostname}`;
  const allowed = DIAN_ALLOWED_ORIGINS.some((o) => origin === o);

  if (!allowed) {
    res.status(400).json({
      status: "error",
      detalle: "token_url debe apuntar al dominio de la DIAN (catalogo-vpfe.dian.gov.co)",
    });
    return;
  }

  // Normaliza token_url para que el resto del flujo use siempre el enlace directo DIAN.
  req.body.token_url = token_url;

  next();
}
