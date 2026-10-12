/** Validación de la URL del token del portal de habilitación DIAN (llega al correo del RUT; es un secreto: nunca se guarda ni se imprime). */

export interface DianTokenInfo { host: string; nit: string; delegated: boolean }

const HAB_HOST = "catalogo-vpfe-hab.dian.gov.co";

/** Comprueba que sea el enlace de habilitación (https, host y ruta exactos) y extrae el NIT (`rk`). Lanza un mensaje claro si no lo es. */
export function parseDianTokenUrl(raw: string): DianTokenInfo {
  let u: URL;
  try { u = new URL(String(raw || "").trim()); } catch { throw new Error("Pega el enlace completo que la DIAN envió a tu correo (empieza por https://catalogo-vpfe-hab.dian.gov.co/User/AuthToken…)."); }
  if (u.protocol !== "https:" || u.hostname !== HAB_HOST || u.pathname !== "/User/AuthToken") {
    throw new Error("Ese no es el enlace de HABILITACIÓN de la DIAN. Debe empezar por https://catalogo-vpfe-hab.dian.gov.co/User/AuthToken (no el de producción).");
  }
  const nit = (u.searchParams.get("rk") || "").replace(/\D/g, "");
  const token = u.searchParams.get("token") || "";
  if (!nit || !/^[0-9a-f-]{20,}$/i.test(token)) throw new Error("El enlace de la DIAN está incompleto: copia la dirección completa del correo, con sus parámetros rk y token.");
  return { host: u.hostname, nit, delegated: (u.searchParams.get("pk") || "").includes("|") };
}

/** El enlace debe ser del mismo NIT de la empresa (si no, la DIAN rechaza con COMPANY_DIAN_MISMATCH). */
export function assertTokenMatchesNit(raw: string, nit: string): DianTokenInfo {
  const info = parseDianTokenUrl(raw);
  if (info.nit !== nit.replace(/\D/g, "")) throw new Error(`El enlace de la DIAN es del NIT ${info.nit}, pero esta empresa es ${nit}. Pide el token del NIT correcto.`);
  return info;
}
