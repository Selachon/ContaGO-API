import { Request, Response, NextFunction } from "express";

const DIAN_ALLOWED_ORIGINS = [
  "https://catalogo-vpfe.dian.gov.co",
  "https://catalogo-vpfe-hab.dian.gov.co", // ambiente de habilitación
];

/**
 * Valida que token_url apunte exclusivamente al dominio de la DIAN.
 * Bloquea SSRF impidiendo URLs a otros hosts o protocolos no HTTPS.
 */
export function validateDianUrl(req: Request, res: Response, next: NextFunction): void {
  const { token_url } = req.body;

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

  next();
}
