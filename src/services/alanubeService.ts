/** Cliente de Alanube (API e-provider de Alegra) para facturación electrónica DIAN. Un solo token de reseller; la empresa emisora va en `company.id`. */

const baseUrl = () => (process.env.ALANUBE_BASE_URL || "https://sandbox-api.alegra.com/e-provider/col/v1").replace(/\/+$/, "");

/** GovernmentId (TestSetId) genérico del sandbox; en producción lo da la DIAN por NIT. */
export const ALANUBE_SANDBOX_TEST_SET_ID = "a70562e0-631e-4ceb-aa65-36887b57dc17";

export class AlanubeError extends Error {
  constructor(
    message: string,
    readonly status: number,
    readonly errors: { code?: string; message: string }[] = [],
  ) {
    super(message);
    this.name = "AlanubeError";
  }
}

export interface AlanubeCompanyInput {
  name: string;
  identification: string;
  dv: string;
  useAlegraCertificate: boolean;
  certificate?: { name: string; extension: string; content: string; password: string };
  econimicActivities?: string[]; // (sic) así lo escribe la API
  organizationType?: number;
  regimeCode?: string;
  taxCode?: { id: string; name?: string };
  email?: string;
  address?: { address: string; city: string; department: string; country?: string };
}

export interface AlanubeCompany {
  id: string;
  name: string;
  identification: string;
  dv: string;
  type: "main" | "associated";
  useAlegraCertificate: boolean;
  governmentStatus: Record<string, string>;
}

export interface AlanubeInvoiceResult {
  id: string;
  cufe: string;
  fullNumber: string;
  status: string;
  /** ACCEPTED / REJECTED…: un HTTP 201 no implica aceptación de la DIAN. */
  legalStatus: string;
  governmentResponse?: { code?: string; message?: string; errorMessages?: string[] };
  /** Fecha y hora de emisión con zona horaria (ISO). */
  date?: string;
  /** Nombre del XML según la DIAN (ej. fv09019658560852600000001.xml). */
  xmlFileName?: string;
  /** Texto que va dentro del QR de la representación gráfica. */
  qrCodeContent?: string;
  /** Enlace presignado del XML: caduca en ~1 hora, hay que descargarlo al emitir. */
  xmlUrl?: string;
}

function token(): string {
  const t = process.env.ALANUBE_TOKEN;
  if (!t) throw new AlanubeError("ALANUBE_TOKEN no está configurado", 500);
  return t;
}

async function request<T>(method: string, path: string, body?: unknown): Promise<T> {
  const res = await fetch(`${baseUrl()}${path}`, {
    method,
    headers: { Authorization: `Bearer ${token()}`, "Content-Type": "application/json" },
    body: body === undefined ? undefined : JSON.stringify(body),
    signal: AbortSignal.timeout(90_000),
  });
  const text = await res.text();
  let json: any = null;
  try { json = text ? JSON.parse(text) : null; } catch { /* respuesta no JSON */ }
  if (!res.ok) {
    const errors: { code?: string; message: string }[] = Array.isArray(json?.errors) ? json.errors : [];
    const detail = errors.map((e) => (e.code ? `${e.code}: ${e.message}` : e.message)).join(" | ") || json?.message || text.slice(0, 200);
    throw new AlanubeError(`Alanube ${method} ${path} → ${res.status}: ${detail}`, res.status, errors);
  }
  return json as T;
}

export async function createCompany(input: AlanubeCompanyInput): Promise<AlanubeCompany> {
  return (await request<{ company: AlanubeCompany }>("POST", "/companies", input)).company;
}

export async function getCompanyByNit(nit: string): Promise<AlanubeCompany | null> {
  const r = await request<{ companies: AlanubeCompany[] }>("GET", `/companies?identification=${encodeURIComponent(nit)}`);
  return r.companies[0] ?? null;
}

/**
 * Envía el set de pruebas de un tipo de documento. Es idempotente: si ya fue aprobado,
 * la API responde con error y aquí se trata como "ya habilitado".
 */
export async function submitTestSet(
  companyId: string,
  type: "invoices" | "payrolls" | "pos",
  governmentId: string = ALANUBE_SANDBOX_TEST_SET_ID,
): Promise<{ status: string; alreadyApproved: boolean }> {
  try {
    const r = await request<{ testSet: { status: string } }>("POST", "/test-sets", { company: { id: companyId }, type, governmentId });
    return { status: r.testSet.status, alreadyApproved: false };
  } catch (e) {
    if (e instanceof AlanubeError && /already been approved/i.test(e.message)) return { status: "ACCEPTED", alreadyApproved: true };
    throw e;
  }
}

/** Emite una factura. El payload ya debe traer `company.id`, `resolution`, ítems y totales calculados. */
export async function createInvoice(payload: Record<string, unknown>): Promise<AlanubeInvoiceResult> {
  const r = await request<{ invoice: AlanubeInvoiceResult; files?: { xml?: string } }>("POST", "/invoices", payload);
  return { ...r.invoice, xmlUrl: r.files?.xml };
}

export interface AlanubeInvoiceFiles { xml?: string; attachedDocument?: string; zip?: string; applicationResponse?: string }

/**
 * Archivos oficiales de una factura (XML firmado, AttachedDocument y ZIP con el AttachedDocument + el PDF de Alanube).
 * Solo vienen en el detalle (`GET /invoices/{id}`), no en la respuesta del POST, y pueden tardar en aparecer.
 */
export async function getInvoiceFiles(id: string, retries = 3, waitMs = 2000, resource: "invoices" | "credit-notes" | "support-documents" = "invoices"): Promise<AlanubeInvoiceFiles> {
  let files: AlanubeInvoiceFiles = {};
  for (let i = 0; i < retries; i++) {
    files = (await request<{ files?: AlanubeInvoiceFiles }>("GET", `/${resource}/${encodeURIComponent(id)}`)).files || {};
    if (files.zip) return files;
    if (i < retries - 1) await new Promise(r => setTimeout(r, waitMs));
  }
  return files;
}

/** Descarga un archivo por su enlace presignado (el del XML caduca, no confiar en guardarlo). */
export async function downloadFile(url: string): Promise<Buffer> {
  const res = await fetch(url, { signal: AbortSignal.timeout(60_000) });
  if (!res.ok) throw new AlanubeError(`No se pudo descargar el archivo (${res.status})`, res.status);
  return Buffer.from(await res.arrayBuffer());
}

/** true si la DIAN aceptó el documento, con o sin observaciones (ACCEPTED_WITH_OBSERVATIONS). */
export function isInvoiceAccepted(r: Pick<AlanubeInvoiceResult, "legalStatus">): boolean {
  return r.legalStatus === "ACCEPTED" || r.legalStatus === "ACCEPTED_WITH_OBSERVATIONS";
}

/**
 * Emite un documento soporte (adquisición a un no obligado a facturar). Lleva resolución propia (prefijo tipo SEDS, sin llave
 * técnica); la respuesta viene bajo `supportDocument`.
 */
export async function createSupportDocument(payload: Record<string, unknown>): Promise<AlanubeInvoiceResult & { cuds?: string; applicationResponseUrl?: string }> {
  const r = await request<{ supportDocument: AlanubeInvoiceResult & { cuds?: string }; files?: { xml?: string; applicationResponse?: string } }>("POST", "/support-documents", payload);
  // A diferencia de factura y nota crédito, aquí Alanube NO entrega ZIP ni PDF: solo el XML y la respuesta de la DIAN
  return { ...r.supportDocument, xmlUrl: r.files?.xml, applicationResponseUrl: r.files?.applicationResponse };
}

/**
 * Nombre y correo que la DIAN tiene registrados para un documento. Para un número que no existe Alanube responde
 * "Consumidor Final" con HTTP 200: eso se trata como "no encontrado".
 */
/** Resoluciones de facturación vigentes del NIT en la DIAN de producción (solo existe en producción). */
export async function getProductionResolutions(nit: string): Promise<Array<{ prefix: string; resolutionNumber: string; minNumber: number; maxNumber: number; startDate: string; endDate: string; technicalKey?: string }>> {
  const r = await request<{ resolutions?: Array<Record<string, any>> }>("GET", `/resolutions/${encodeURIComponent(nit)}`);
  return (r.resolutions || []).map(x => ({
    prefix: String(x.prefix ?? ""), resolutionNumber: String(x.resolutionNumber ?? ""), minNumber: Number(x.minNumber), maxNumber: Number(x.maxNumber),
    startDate: String(x.startDate ?? ""), endDate: String(x.endDate ?? ""), technicalKey: x.technicalKey ? String(x.technicalKey) : undefined,
  }));
}

export async function getAcquirerInfo(identificationType: string, identificationNumber: string): Promise<{ name: string; email: string } | null> {
  const q = `identificationType=${encodeURIComponent(identificationType)}&identificationNumber=${encodeURIComponent(identificationNumber)}`;
  const r = await request<{ receiverName?: string; receiverEmail?: string }>("GET", `/acquirer-info?${q}`);
  const name = (r.receiverName || "").trim();
  if (!name || /consumidor\s+final/i.test(name)) return null;
  return { name, email: (r.receiverEmail || "").trim() };
}

// ─── Empresas emisoras (una por cliente) y habilitación en la DIAN ──────────────────────────

export interface AlanubeCompanyCreate {
  name: string; identification: string; dv: string; useAlegraCertificate: boolean;
  econimicActivities: string[]; // (sic) así lo escribe la API
  organizationType: number; regimeCode: string; taxCode: { id: string };
  email: string; phone?: string; address: { address: string; city: string; department: string; country: string };
}

export async function getCompanyById(id: string): Promise<AlanubeCompany> {
  return (await request<{ company: AlanubeCompany }>("GET", `/companies/${encodeURIComponent(id)}`)).company;
}

/** Reenvía por correo un documento ya emitido (POST /send-email; tipo 01 = factura, 91 = nota crédito). */
export async function sendDocumentEmail(documentType: string, documentId: string, to: string, replyTo: string): Promise<void> {
  await request("POST", "/send-email", { documentType, documentId, email: { to, replyTo } });
}

/** Edita una empresa ya creada (PATCH /companies/{id}; /company/{id} no existe). */
export async function updateCompanyById(id: string, patch: Partial<AlanubeCompanyCreate>): Promise<AlanubeCompany> {
  return (await request<{ company: AlanubeCompany }>("PATCH", `/companies/${encodeURIComponent(id)}`, patch)).company;
}

export interface EnablementPlanItem { documentType: string; requested: boolean; requiredBy?: string | null; status?: string; errorCode?: string | null; errorMessage?: string | null }
export interface EnablementProcess { processId: string; status: "IN_PROCESS" | "FINISHED" | "FAILED"; errorCode?: string | null; errorMessage?: string | null; executionPlan: EnablementPlanItem[] }

/** Habilitación inteligente: recibe la URL del token del portal de habilitación DIAN (secreto, no se guarda ni se registra). */
export async function startEnablement(tokenDIAN: string, companyId: string, documents: string[]): Promise<EnablementProcess> {
  return request<EnablementProcess>("POST", "/test-sets/intelligent", { tokenDIAN, companyId, documents });
}

export async function getEnablementProgress(processId: string): Promise<EnablementProcess> {
  return request<EnablementProcess>("GET", `/test-sets/intelligent/${encodeURIComponent(processId)}`);
}

export interface DianPlace { code: string; value: string; departmentCode?: string }

let placesCache: { at: number; departments: DianPlace[]; municipalities: DianPlace[]; countries: DianPlace[] } | null = null;

/** Departamentos y municipios de la DIAN (cambian casi nunca: se guardan 24 h en memoria). */
export async function getDianPlaces(): Promise<{ departments: DianPlace[]; municipalities: DianPlace[]; countries: DianPlace[] }> {
  if (placesCache && Date.now() - placesCache.at < 24 * 3600_000) return placesCache;
  const [d, m, c] = await Promise.all([
    request<{ departments: DianPlace[] }>("GET", "/dian/departments"),
    request<{ municipalities: DianPlace[] }>("GET", "/dian/municipalities"),
    request<{ countries: DianPlace[] }>("GET", "/dian/countries"),
  ]);
  placesCache = { at: Date.now(), countries: c.countries, departments: d.departments, municipalities: m.municipalities.map(x => ({ code: x.code, value: x.value, departmentCode: x.departmentCode })) };
  return placesCache;
}

/**
 * Emite una nota crédito (sin resolución: numeración propia por prefijo, que controla el emisor).
 * La respuesta es como la de la factura pero bajo `creditNote` y su identificador fiscal es el CUDE.
 */
export async function createCreditNote(payload: Record<string, unknown>): Promise<AlanubeInvoiceResult & { cude: string }> {
  const r = await request<{ creditNote: AlanubeInvoiceResult & { cude: string }; files?: { xml?: string } }>("POST", "/credit-notes", payload);
  return { ...r.creditNote, xmlUrl: r.files?.xml };
}
