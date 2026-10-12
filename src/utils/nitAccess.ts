import { Request, Response } from "express";
import { getAllowedNits, normalizeNitValue } from "../services/database.js";

/**
 * Valida que el NIT (`rk`) del `token_url` esté habilitado para el usuario en la
 * herramienta indicada. Devuelve `true` si ya se respondió con un error (el
 * llamador debe hacer `return`), `false` si el acceso es válido.
 *
 * Los administradores nunca se restringen.
 */
export async function rejectIfNitNotAllowed(
  req: Request,
  res: Response,
  tokenUrl: string | undefined,
  toolId: string,
): Promise<boolean> {
  if (req.user?.isAdmin) return false;

  const allowed = await getAllowedNits(req.user!.userId, toolId);
  if (allowed.length === 0) {
    res.status(403).json({
      status: "error",
      detalle: "No tienes NITs habilitados para esta herramienta. Contacta al administrador.",
    });
    return true;
  }

  let tokenNit = "";
  try {
    tokenNit = new URL(tokenUrl || "").searchParams.get("rk")?.trim() || "";
  } catch {
    tokenNit = "";
  }

  if (!tokenNit) {
    res.status(400).json({ status: "error", detalle: "El token_url no contiene rk (NIT)." });
    return true;
  }

  if (!allowed.includes(normalizeNitValue(tokenNit))) {
    res.status(403).json({ status: "error", detalle: `No tienes acceso al NIT ${tokenNit}` });
    return true;
  }

  return false;
}

/** ¿El NIT (en cualquier formato) está dentro de la lista permitida (ya normalizada)? */
export function isNitAllowed(allowedNormalized: string[], nit: unknown): boolean {
  return allowedNormalized.includes(normalizeNitValue(nit));
}
