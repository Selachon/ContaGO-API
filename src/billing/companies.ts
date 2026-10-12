import { ObjectId } from "mongodb";
import { getDb } from "../services/database.js";
import {
  createCompany as alanubeCreateCompany, getAcquirerInfo, getCompanyById, getEnablementProgress, startEnablement, submitTestSet,
  updateCompanyById, type AlanubeCompanyCreate,
} from "../services/alanubeService.js";
import { nitDv } from "../services/alanubeInvoicing.js";
import { assertTokenMatchesNit } from "./dianToken.js";

/**
 * Empresas emisoras de los clientes que compran el servicio de facturación electrónica. Cada empresa tiene su propia
 * empresa en Alanube, su estado de habilitación en la DIAN y sus resoluciones de numeración.
 * Por ahora todo corre en el ambiente de habilitación (pruebas).
 */

export const BILLING_TOOL_ID = "facturacion-electronica";
const col = () => getDb().collection<any>("billingCompanies");
const users = () => getDb().collection<any>("users");

export interface Actor { userId: string; isAdmin: boolean }

/** Responsabilidades del RUT que acepta la API (se envían separadas por «;»). */
export const REGIME_CODES: Record<string, string> = {
  "O-13": "Gran contribuyente", "O-15": "Autorretenedor", "O-23": "Agente de retención en el impuesto sobre las ventas",
  "O-47": "Régimen simple de tributación", "O-48": "Impuesto sobre las ventas - IVA", "O-49": "No responsable de IVA", "R-99-PN": "No aplica - Otros",
};
export const TAX_CODES: Record<string, string> = { "01": "IVA", "04": "Impuesto al consumo", ZA: "IVA e impuesto nacional al consumo", ZZ: "No aplica" };

export interface CompanyInput {
  nit: string;
  name: string;
  tradeName?: string;
  organizationType: 1 | 2;
  regimeCodes: string[];
  taxCode: string;
  economicActivities: string[];
  email: string;
  phone?: string;
  address: { address: string; city: string; department: string };
}

const clean = (v: unknown, max: number) => String(v ?? "").replace(/\s+/g, " ").trim().slice(0, max);
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;

/** Normaliza y valida lo que llega del formulario; lanza el primer problema con un mensaje claro. */
export function normalizeCompanyInput(raw: any): CompanyInput {
  const nit = clean(raw?.nit, 20).replace(/\D/g, "").replace(/^0+/, "");
  if (nit.length < 5 || nit.length > 12) throw new Error("Escribe el NIT sin puntos, guiones ni dígito de verificación (entre 5 y 12 dígitos).");
  const name = clean(raw?.name, 450);
  if (name.length < 5) throw new Error("Escribe la razón social tal como figura en el RUT (mínimo 5 caracteres): la DIAN la compara de forma literal.");
  const regimeCodes = [...new Set((Array.isArray(raw?.regimeCodes) ? raw.regimeCodes : []).map((c: unknown) => clean(c, 10)))].filter(Boolean) as string[];
  if (!regimeCodes.length) throw new Error("Selecciona al menos una responsabilidad del RUT.");
  const badRegime = regimeCodes.find(c => !REGIME_CODES[c]);
  if (badRegime) throw new Error(`Responsabilidad del RUT no válida: ${badRegime}.`);
  const taxCode = clean(raw?.taxCode, 2) || "ZZ";
  if (!TAX_CODES[taxCode]) throw new Error("Tipo de tributo no válido.");
  const economicActivities = [...new Set((Array.isArray(raw?.economicActivities) ? raw.economicActivities : [raw?.economicActivities]).map((c: unknown) => clean(c, 4)))].filter(Boolean) as string[];
  if (!economicActivities.length || economicActivities.some(c => !/^\d{4}$/.test(c))) throw new Error("Escribe la actividad económica (CIIU) de 4 dígitos, por ejemplo 6201.");
  const email = clean(raw?.email, 120);
  if (!EMAIL_RE.test(email)) throw new Error("Escribe un correo electrónico válido: ahí llegarán las facturas que recibas.");
  const address = clean(raw?.address?.address, 200);
  const city = clean(raw?.address?.city, 5);
  if (!address) throw new Error("Escribe la dirección de la empresa.");
  if (!/^\d{5}$/.test(city)) throw new Error("Selecciona el municipio de la empresa.");
  return {
    nit, name, tradeName: clean(raw?.tradeName, 450) || undefined, organizationType: raw?.organizationType === 2 ? 2 : 1,
    regimeCodes, taxCode, economicActivities: economicActivities.slice(0, 5), email, phone: clean(raw?.phone, 30) || undefined,
    address: { address, city, department: city.slice(0, 2) }, // el código DIAN del municipio empieza por el del departamento
  };
}

/** Cuerpo de `POST /companies` de Alanube (por ahora siempre con el certificado de Alegra). */
export function toAlanubeCompany(i: CompanyInput): AlanubeCompanyCreate {
  return {
    name: i.name, identification: i.nit, dv: nitDv(i.nit), useAlegraCertificate: true, econimicActivities: i.economicActivities,
    organizationType: i.organizationType, regimeCode: i.regimeCodes.join(";"), taxCode: { id: i.taxCode }, email: i.email,
    ...(i.phone && { phone: i.phone }), address: { ...i.address, country: "CO" },
  };
}

const oid = (id: string) => { if (!ObjectId.isValid(id)) throw new Error("Empresa inválida."); return new ObjectId(id); };
const scope = (a: Actor) => (a.isAdmin ? {} : { ownerUserId: a.userId });

let indexReady = false;
async function ensureIndex() {
  if (indexReady) return;
  await col().createIndex({ nit: 1 }, { unique: true });
  indexReady = true;
}

/** Empresas que puede crear el usuario: el administrador sin tope; los demás, lo que diga su plan (por defecto 1). */
async function companyLimit(a: Actor): Promise<number> {
  if (a.isAdmin) return Number.POSITIVE_INFINITY;
  const u = await users().findOne({ _id: oid(a.userId) }, { projection: { toolCompanyLimits: 1 } });
  const n = Number(u?.toolCompanyLimits?.[BILLING_TOOL_ID]);
  return Number.isFinite(n) && n > 0 ? n : 1;
}

/** Estado de avance de la empresa para mostrar el asistente. */
export function readiness(c: any) {
  const created = !!c.alanube?.companyId;
  const authorized = c.governmentStatus?.invoices === "AUTHORIZED";
  const hasRange = (c.resolutions || []).some((r: any) => r.docType === "invoice" && r.active);
  // Documento soporte: además de lo anterior, el cliente asocia a Alegra en la DIAN (a mano) y se guarda su numeración
  const supportReady = created && authorized && (c.resolutions || []).some((r: any) => r.docType === "supportDocument" && r.active);
  return { created, authorized, hasRange, ready: created && authorized && hasRange, supportReady };
}

const publicView = (c: any) => ({ ...c, id: String(c._id), _id: undefined, readiness: readiness(c) });

export async function listCompanies(a: Actor) {
  return (await col().find(scope(a)).sort({ createdAt: -1 }).toArray()).map(publicView);
}

export async function getCompany(a: Actor, id: string) {
  const c = await col().findOne({ _id: oid(id), ...scope(a) });
  if (!c) throw new Error("Empresa no encontrada.");
  return c;
}

/** Nombre y correo que la DIAN tiene registrados para el NIT (ayuda a escribir la razón social exacta). */
export async function lookupNit(rawNit: string) {
  const nit = String(rawNit ?? "").replace(/\D/g, "");
  if (nit.length < 5) throw new Error("Escribe primero el NIT.");
  return { nit, dv: nitDv(nit), info: await getAcquirerInfo("31", nit) };
}

/** Crea la empresa en nuestra base y en Alanube. Si Alanube falla no queda nada guardado. */
export async function createCompany(a: Actor, raw: any) {
  await ensureIndex();
  const input = normalizeCompanyInput(raw);
  if (await col().findOne({ nit: input.nit })) throw new Error("Ese NIT ya está registrado en ContaGO. Si es tu empresa y no la registraste tú, escríbenos para ayudarte.");
  const limit = await companyLimit(a);
  const owned = await col().countDocuments({ ownerUserId: a.userId });
  if (owned >= limit) throw new Error(`Tu plan permite ${limit} empresa${limit === 1 ? "" : "s"} para facturación electrónica. Contacta a ContaGO para ampliarlo.`);

  const created = await alanubeCreateCompany(toAlanubeCompany(input));
  const now = new Date().toISOString();
  const doc = {
    ownerUserId: a.userId, input, nit: input.nit, dv: nitDv(input.nit), name: input.name, environment: "habilitacion",
    alanube: { companyId: created.id, createdAt: now }, governmentStatus: created.governmentStatus || {},
    enablement: {}, resolutions: [] as any[], createdAt: now, updatedAt: now,
  };
  try {
    const r = await col().insertOne(doc);
    return publicView({ ...doc, _id: r.insertedId });
  } catch (e: any) {
    if (e?.code === 11000) throw new Error("Ese NIT ya está registrado en ContaGO.");
    throw e;
  }
}

/** Corrige los datos de la empresa (también en Alanube). El NIT no se cambia: es la identidad de la empresa. */
export async function updateCompany(a: Actor, id: string, raw: any) {
  const c = await getCompany(a, id);
  const input = normalizeCompanyInput({ ...raw, nit: c.nit });
  const body = toAlanubeCompany(input);
  await updateCompanyById(c.alanube.companyId, { name: body.name, econimicActivities: body.econimicActivities, organizationType: body.organizationType, regimeCode: body.regimeCode, taxCode: body.taxCode, email: body.email, phone: body.phone, address: body.address });
  await col().updateOne({ _id: c._id }, { $set: { input, name: input.name, updatedAt: new Date().toISOString() } });
  return publicView(await getCompany(a, id));
}

/** Inicia la habilitación de facturas con el token de la DIAN. El token se usa una vez y no se guarda en ningún sitio. */
export async function startCompanyEnablement(a: Actor, id: string, tokenUrl: string) {
  const c = await getCompany(a, id);
  assertTokenMatchesNit(tokenUrl, c.nit);
  const p = await startEnablement(tokenUrl, c.alanube.companyId, ["invoices"]);
  await col().updateOne({ _id: c._id }, { $set: { "enablement.processId": p.processId, "enablement.invoices": "IN_PROCESS", "enablement.startedAt": new Date().toISOString(), "enablement.error": null, updatedAt: new Date().toISOString() } });
  return { processId: p.processId, status: p.status };
}

const ENABLEMENT_ERRORS: Record<string, string> = {
  TEST_SET_ID_NOT_FOUND: "No encontré el set de pruebas en el portal de la DIAN. Verifica que en «Configurar modos de operación» quedó asociado el software Alegra (proveedor SOLUCIONES ALEGRA S.A.S) y que aparece el «Identificador del set de pruebas». Si ya está ahí, pégalo manualmente abajo.",
};
function enablementErrorText(code?: string | null, message?: string | null) {
  return (code && ENABLEMENT_ERRORS[code]) || message || code || "La habilitación falló.";
}

/** Solo sandbox: el simulador acepta un set de pruebas fijo (ACCEPTED) y deja la empresa habilitada sin pasar por la DIAN. */
export const isSandbox = () => /sandbox/i.test(process.env.ALANUBE_BASE_URL || "sandbox");

export async function sandboxEnableCompany(a: Actor, id: string) {
  if (!isSandbox()) throw new Error("La habilitación en pruebas solo existe en el ambiente de pruebas.");
  const c = await getCompany(a, id);
  const r = await submitTestSet(c.alanube.companyId, "invoices");
  if (r.status !== "ACCEPTED") throw new Error(`Alanube no aceptó el set de pruebas (estado ${r.status}).`);
  await col().updateOne({ _id: c._id }, { $set: { "enablement.invoices": "AUTHORIZED", "enablement.processId": null, "enablement.testSetStatus": r.status, "enablement.mode": "sandbox-fixed-id", "enablement.error": null, updatedAt: new Date().toISOString() } });
  return refreshCompanyEnablement(a, id);
}

/** Consulta el avance de la habilitación y el estado de la empresa en Alanube, y lo deja guardado. */
export async function refreshCompanyEnablement(a: Actor, id: string) {
  const c = await getCompany(a, id);
  const set: Record<string, unknown> = { updatedAt: new Date().toISOString() };
  const processId = c.enablement?.processId;
  if (processId && c.enablement?.invoices === "IN_PROCESS") {
    const p = await getEnablementProgress(processId);
    const inv = p.executionPlan.find(x => x.documentType === "invoices");
    if (p.status === "FINISHED") set["enablement.invoices"] = inv?.status || "AUTHORIZED";
    else if (p.status === "FAILED") { set["enablement.invoices"] = "FAILED"; set["enablement.error"] = enablementErrorText(inv?.errorCode || p.errorCode, inv?.errorMessage || p.errorMessage); }
  }
  const remote = await getCompanyById(c.alanube.companyId);
  set.governmentStatus = remote.governmentStatus || {};
  await col().updateOne({ _id: c._id }, { $set: set });
  return publicView(await getCompany(a, id));
}

export interface ResolutionInput { prefix: string; resolutionNumber: string; minNumber: number; maxNumber: number; startDate: string; endDate: string; technicalKey?: string; /** Primer número que se usará (en producción, el último usado + 1). */ startNumber?: number }

/** Valida una resolución de numeración de facturas (la de pruebas de la DIAN o, más adelante, la real). */
export function normalizeResolution(raw: any): ResolutionInput {
  // Los identificadores fiscales NO se recortan: si son demasiado largos se rechazan (recortarlos en silencio los alteraría)
  const prefix = String(raw?.prefix ?? "").trim().toUpperCase();
  const resolutionNumber = String(raw?.resolutionNumber ?? "").trim();
  const minNumber = Number(raw?.minNumber), maxNumber = Number(raw?.maxNumber);
  const startDate = String(raw?.startDate ?? "").trim(), endDate = String(raw?.endDate ?? "").trim();
  const technicalKey = String(raw?.technicalKey ?? "").trim().toLowerCase();
  if (!/^[A-Z0-9]{1,4}$/.test(prefix)) throw new Error("El prefijo debe tener de 1 a 4 letras o números.");
  if (!/^\d{4,14}$/.test(resolutionNumber)) throw new Error("El número de resolución debe tener entre 4 y 14 dígitos.");
  if (!Number.isInteger(minNumber) || !Number.isInteger(maxNumber) || minNumber < 0 || maxNumber <= minNumber) throw new Error("El rango de numeración no es válido (el «hasta» debe ser mayor que el «desde»).");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(startDate) || !/^\d{4}-\d{2}-\d{2}$/.test(endDate) || endDate < startDate) throw new Error("Las fechas de vigencia no son válidas (AAAA-MM-DD).");
  if (!/^[0-9a-f]{20,64}$/.test(technicalKey)) throw new Error("La llave técnica debe ser una cadena de 20 a 64 caracteres hexadecimales.");
  const sn = raw?.startNumber === undefined || raw?.startNumber === null || raw?.startNumber === "" ? undefined : Number(raw.startNumber);
  if (sn !== undefined && (!Number.isInteger(sn) || sn < minNumber || sn > maxNumber)) throw new Error("El número inicial debe estar dentro del rango de la resolución.");
  return { prefix, resolutionNumber, minNumber, maxNumber, startDate, endDate, technicalKey, ...(sn !== undefined && { startNumber: sn }) };
}

/**
 * Número inicial por defecto en el ambiente de pruebas: el rango de prueba de la DIAN (por ejemplo SETT 1 a 5.000.000) lo comparten
 * todas las pruebas, y empezar por el 1 casi siempre choca con documentos ya procesados («regla 90»). Se sugiere uno al azar.
 */
export function suggestTestStart(min: number, max: number, rnd = Math.random()): number {
  const span = max - min;
  if (span < 200_000) return min;
  return min + 100_000 + Math.floor(rnd * Math.min(span - 200_000, 3_000_000));
}

/** Guarda la resolución de facturas de la empresa (reemplaza una anterior con el mismo prefijo y número). */
export async function saveCompanyResolution(a: Actor, id: string, raw: any) {
  const c = await getCompany(a, id);
  const r = normalizeResolution(raw);
  const others = (c.resolutions || []).filter((x: any) => !(x.prefix === r.prefix && x.resolutionNumber === r.resolutionNumber));
  const entry = { ...r, startNumber: r.startNumber ?? suggestTestStart(r.minNumber, r.maxNumber), docType: "invoice", environment: "habilitacion", active: true, addedAt: new Date().toISOString() };
  // Solo una resolución de facturas activa por ambiente: las demás quedan inactivas (siguen guardadas)
  const resolutions = [...others.map((x: any) => (x.docType === "invoice" ? { ...x, active: false } : x)), entry];
  await col().updateOne({ _id: c._id }, { $set: { resolutions, updatedAt: new Date().toISOString() } });
  return publicView(await getCompany(a, id));
}

/** Numeración de prueba del documento soporte en el sandbox (verificada: la DIAN la acepta). No lleva llave técnica. */
export const SANDBOX_SUPPORT_RESOLUTION = { prefix: "SEDS", resolutionNumber: "18760000001", minNumber: 984000000, maxNumber: 985000000, startDate: "2019-01-19", endDate: "2030-01-19" };

/**
 * Activa el documento soporte de la empresa. El cliente confirma que ya asoció a Alegra en el modo de operación «Documento para
 * no obligados a facturar» de la DIAN (esa asociación NO se puede automatizar: la habilitación inteligente falla en ese paso).
 */
export async function enableCompanySupportDocs(a: Actor, id: string, raw: any) {
  const c = await getCompany(a, id);
  if (!readiness(c).authorized) throw new Error("Primero habilita las facturas de tu empresa en la DIAN.");
  if (raw?.confirmAssociated !== true) throw new Error("Confirma que ya asociaste a Alegra en «Documento para no obligados a facturar».");
  const base = { ...SANDBOX_SUPPORT_RESOLUTION, technicalKey: undefined as string | undefined };
  const startRaw = raw?.startNumber;
  const startNumber = startRaw === undefined || startRaw === "" || startRaw === null ? suggestTestStart(base.minNumber, base.maxNumber) : Number(startRaw);
  if (!Number.isInteger(startNumber) || startNumber < base.minNumber || startNumber > base.maxNumber) throw new Error("El número inicial debe estar dentro del rango de la numeración.");
  const entry = { ...base, startNumber, docType: "supportDocument", environment: "habilitacion", active: true, addedAt: new Date().toISOString() };
  const resolutions = [...(c.resolutions || []).filter((x: any) => x.docType !== "supportDocument"), entry];
  await col().updateOne({ _id: c._id }, { $set: { resolutions, "enablement.supportDocsAssociatedAt": new Date().toISOString(), updatedAt: new Date().toISOString() } });
  return publicView(await getCompany(a, id));
}
