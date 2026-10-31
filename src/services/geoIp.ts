/**
 * Ubicación aproximada (municipio) a partir de la IP, calculada en el servidor con una base local
 * (geoip-lite). No se llama a ningún servicio externo y la IP NO se guarda: solo el resultado.
 *
 * Límites: la geolocalización por IP es aproximada. En Colombia suele resolver bien las
 * ciudades grandes, a veces devuelve la ciudad del proveedor de internet y, para algunos rangos,
 * solo el país. Por eso `city` puede venir vacío: el panel lo muestra como "sin municipio".
 */
import { createRequire } from "module";

export interface Geo {
  country: string;
  region?: string;
  city?: string;
}

// Códigos de departamento de MaxMind (ISO 3166-2:CO)
const CO_REGIONS: Record<string, string> = {
  AMA: "Amazonas", ANT: "Antioquia", ARA: "Arauca", ATL: "Atlántico", BOL: "Bolívar", BOY: "Boyacá", CAL: "Caldas",
  CAQ: "Caquetá", CAS: "Casanare", CAU: "Cauca", CES: "Cesar", CHO: "Chocó", COR: "Córdoba", CUN: "Cundinamarca",
  DC: "Bogotá D.C.", GUA: "Guainía", GUV: "Guaviare", HUI: "Huila", LAG: "La Guajira", MAG: "Magdalena", MET: "Meta",
  NAR: "Nariño", NSA: "Norte de Santander", PUT: "Putumayo", QUI: "Quindío", RIS: "Risaralda", SAP: "San Andrés y Providencia",
  SAN: "Santander", SUC: "Sucre", TOL: "Tolima", VAC: "Valle del Cauca", VAU: "Vaupés", VID: "Vichada",
};

type Lookup = (ip: string) => { country?: string; region?: string; city?: string } | null;
let lookupFn: Lookup | null | undefined; // undefined = aún no se intentó cargar

function loadLookup(): Lookup | null {
  if (lookupFn !== undefined) return lookupFn;
  try {
    const require = createRequire(import.meta.url);
    lookupFn = (require("geoip-lite") as { lookup: Lookup }).lookup;
  } catch (err) {
    console.warn("[Geo] geoip-lite no disponible; no se registrará ubicación:", (err as Error)?.message);
    lookupFn = null;
  }
  return lookupFn;
}

/** Solo para pruebas: inyecta un resolvedor. */
export function _setLookupForTests(fn: Lookup | null | undefined): void { lookupFn = fn; }

/** IPv4 mapeada en IPv6 (::ffff:1.2.3.4) y puertos → IP limpia. */
export function normalizeIp(raw: string | undefined | null): string | null {
  if (!raw) return null;
  let ip = String(raw).trim();
  if (ip.startsWith("::ffff:")) ip = ip.slice(7);
  return ip || null;
}

export function isPrivateIp(ip: string): boolean {
  if (ip === "::1" || ip === "localhost") return true;
  if (/^(10\.|127\.|192\.168\.|169\.254\.)/.test(ip)) return true;
  const m = /^172\.(\d+)\./.exec(ip);
  if (m && Number(m[1]) >= 16 && Number(m[1]) <= 31) return true;
  if (/^100\.(6[4-9]|[7-9]\d|1[01]\d|12[0-7])\./.test(ip)) return true; // CGNAT / Tailscale
  if (/^(fc|fd|fe80)/i.test(ip)) return true;
  return false;
}

export function resolveGeo(rawIp: string | undefined | null): Geo | null {
  const ip = normalizeIp(rawIp);
  if (!ip || isPrivateIp(ip)) return null;
  const lookup = loadLookup();
  if (!lookup) return null;
  try {
    const r = lookup(ip);
    if (!r?.country) return null;
    const geo: Geo = { country: r.country };
    if (r.country === "CO") {
      if (r.region) geo.region = CO_REGIONS[r.region] || r.region;
      if (r.city) geo.city = r.city;
    }
    return geo;
  } catch {
    return null;
  }
}
