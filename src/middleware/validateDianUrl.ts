import { Request, Response, NextFunction } from "express";

const DIAN_ALLOWED_ORIGINS = [
  "https://catalogo-vpfe.dian.gov.co",
];

// catalogo-vpfe-hab es el ambiente de PRUEBAS/certificación de la DIAN (datos ficticios,
// sesión propia): nunca tuvo soporte real — todo el scraper navega con el host de
// producción ya fijo (RedirectToBiller, gratis-vpfe, Document/Received...), así que un
// token de hab autentica pero falla justo después al pasar a producción con otra sesión.
// Encontrado 2026-10-05: un usuario mandó un token de hab (vía un enlace de Outlook
// SafeLinks) y el flujo fallaba en "Autenticando" sin explicación. Se rechaza aquí con un
// mensaje claro en vez de dejarlo fallar más adelante.
const DIAN_HAB_ORIGIN = "https://catalogo-vpfe-hab.dian.gov.co";

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

  if (origin === DIAN_HAB_ORIGIN) {
    res.status(400).json({
      status: "error",
      detalle: "Este enlace es del ambiente de PRUEBAS de la DIAN (catalogo-vpfe-hab.dian.gov.co), no el de producción. " +
        "Pide que generen el token desde https://catalogo-vpfe.dian.gov.co — el de pruebas no tiene las facturas reales y no funciona en esta herramienta.",
    });
    return;
  }

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
