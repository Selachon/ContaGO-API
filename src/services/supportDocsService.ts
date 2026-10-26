import { randomBytes, timingSafeEqual } from "crypto";
import { ObjectId, Binary } from "mongodb";
import { getDb } from "./database.js";
import { AlanubeError, createSupportDocument, downloadFile, getAcquirerInfo, isInvoiceAccepted } from "./alanubeService.js";
import { renderSupportDocPdf } from "./alanubePdf.js";
import { buildItem, buildTotals, round2 } from "./alanubeTotals.js";
import { loadIssuerConfig, nitDv, validatePayment, todayBogota, type AlanubeIssuerConfig } from "./alanubeInvoicing.js";

/**
 * Documentos soporte en adquisiciones a no obligados a facturar (cuentas de cobro) emitidos con Alanube.
 * Cada usuario carga la cuenta de cobro (datos + PDF) y emite con unos clics. Por ahora solo ambiente de
 * habilitación (sandbox): el resultado queda en el propio registro y no afecta ninguna otra colección.
 */

const docs = () => getDb().collection<any>("supportDocs");
const docFiles = () => getDb().collection<any>("supportDocFiles"); // PDF de la cuenta de cobro
const numbers = () => getDb().collection<any>("supportDocNumbers"); // un número por documento: el _id único evita repetirlos
const counters = () => getDb().collection<any>("alanubeCounters");
const alanubeFiles = () => getDb().collection<any>("alanubeInvoiceFiles"); // ZIP oficial, por id de Alanube

/** «historico»: documento soporte ya emitido fuera del sistema (importado de la DIAN); no se edita, no se emite y no tiene archivos. */
export type SupportDocStatus = "borrador" | "emitiendo" | "emitido" | "rechazado" | "error" | "historico";

/** Tipos de documento del proveedor que acepta la API. Un residente en Colombia se identifica solo con NIT (31). */
export const SUPPLIER_ID_TYPES: Record<string, string> = {
  "21": "Tarjeta de extranjería", "22": "Cédula de extranjería", "31": "NIT", "41": "Pasaporte", "42": "Documento de identificación extranjero", "47": "PEP", "50": "NIT de otro país",
};

export interface SupportDocInput {
  supplier: {
    /** 10 residente fiscal en Colombia, 11 no residente. */
    origin: "10" | "11";
    organizationType: 1 | 2; // 1 jurídica, 2 natural
    idType: string;
    /** Tipo de documento elegido en el formulario (CC, NIT, CE…): solo informativo; la DIAN recibe `idType`. */
    docType?: string;
    id: string;
    name: string;
    email?: string;
    phone?: string;
    address: string;
    /** Código DIAN del municipio (5 dígitos): solo para residentes. */
    cityCode?: string;
    /** País (ISO de 2 letras, distinto de CO) y ciudad en texto libre: solo para no residentes. */
    country?: string;
    cityName?: string;
    postalCode?: string;
  };
  /** Número y fecha de la cuenta de cobro del proveedor (referencia, va en las notas). */
  bill?: { number?: string; date?: string };
  concept: string;
  quantity: number;
  unitPrice: number;
  /** Varias líneas (Contable): si vienen, mandan sobre `concept`/`quantity`/`unitPrice`, que repiten la primera. */
  lines?: { description: string; quantity: number; unitPrice: number }[];
  issueDate: string;
  /** Retención en la fuente (ReteRenta) en %. */
  reteRentaPct?: number;
  payment?: { form?: "1" | "2"; days?: number; method?: string };
  notes?: string[];
}

export interface SupportDocResolution {
  resolutionNumber: string; prefix: string; minNumber: number; maxNumber: number; startDate: string; endDate: string;
}

/** Resolución del documento soporte (no lleva llave técnica). Habilitación verificada: 18760000001, SEDS 984000000–985000000, 2019-01-19 a 2030-01-19. */
export function loadSupportResolution(): SupportDocResolution {
  const e = (n: string, d?: string) => process.env[n] || d || (() => { throw new Error(`Falta la variable ${n} para emitir documentos soporte.`); })();
  return {
    resolutionNumber: e("ALANUBE_DSE_RES_NUMBER", "18760000001"), prefix: e("ALANUBE_DSE_RES_PREFIX", "SEDS"),
    minNumber: Number(e("ALANUBE_DSE_RES_MIN", "984000000")), maxNumber: Number(e("ALANUBE_DSE_RES_MAX", "985000000")),
    startDate: e("ALANUBE_DSE_RES_START", "2019-01-19"), endDate: e("ALANUBE_DSE_RES_END", "2030-01-19"),
  };
}

/** Tipo de documento del formulario → código DIAN que usa la consulta de nombre registrado. */
const LOOKUP_CODE: Record<string, string> = { CC: "13", TI: "12", NIT: "31", CE: "22", PA: "41", TE: "21", DE: "42", PEP: "47", NITX: "50" };

/** Busca en la DIAN el nombre y el correo registrados de un proveedor (para no escribir el nombre a mano). */
export async function lookupSupplier(docType: string, rawId: string): Promise<{ name: string; email: string } | null> {
  const code = LOOKUP_CODE[docType];
  if (!code) throw new Error("Tipo de documento inválido.");
  const numeric = ["CC", "TI", "NIT"].includes(docType);
  const id = numeric ? String(rawId ?? "").replace(/\D/g, "") : String(rawId ?? "").replace(/[^A-Za-z0-9]/g, "");
  if (id.length < 5) throw new Error("Escribe primero el número de documento.");
  // Solo se consulta con el tipo elegido: probar otro tipo si no aparece daba coincidencias falsas (registros de prueba de la DIAN)
  return getAcquirerInfo(code, id);
}

/** Tarifas de ReteRenta que acepta la API (lista cerrada). La ReteICA no viaja en el documento soporte. */
export const RETE_RENTA_RATES = [0.1, 0.5, 1, 1.5, 2, 2.5, 3, 3.5, 4, 6, 7, 10, 11, 15, 20, 100];

const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
const clean = (v: unknown, max: number) => String(v ?? "").replace(/\s+/g, " ").trim().slice(0, max);

/** Normaliza y valida lo que llega del formulario; lanza un mensaje claro por el primer problema. */
export function normalizeSupportDoc(raw: any): SupportDocInput {
  const s = raw?.supplier || {};
  const origin = s.origin === "11" ? "11" : "10";
  const idType = origin === "10" ? "31" : clean(s.idType, 2);
  if (!SUPPLIER_ID_TYPES[idType]) throw new Error("Tipo de documento del proveedor inválido.");
  const idDigits = idType === "31" ? clean(s.id, 30).replace(/\D/g, "") : clean(s.id, 30).replace(/[^A-Za-z0-9]/g, "");
  const name = clean(s.name, 200);
  const address = clean(s.address, 200);
  const resident = origin === "10";
  const cityCode = resident ? clean(s.cityCode, 5) : "";
  const country = resident ? "CO" : clean(s.country, 2).toUpperCase();
  const cityName = resident ? "" : clean(s.cityName, 80);
  const rawLines: any[] = Array.isArray(raw?.lines) ? raw.lines : [];
  const lines = rawLines.map((l, i) => {
    const line = { description: clean(l?.description, 300), quantity: Number(l?.quantity), unitPrice: Number(l?.unitPrice) };
    if (!line.description) throw new Error(`Falta la descripción del ítem ${i + 1}.`);
    if (!(line.quantity > 0)) throw new Error(`La cantidad del ítem ${i + 1} debe ser mayor que cero.`);
    if (!(line.unitPrice > 0)) throw new Error(`El valor del ítem ${i + 1} debe ser mayor que cero.`);
    return line;
  }).slice(0, 50);
  const concept = lines.length ? lines[0].description : clean(raw?.concept, 300);
  const quantity = lines.length ? lines[0].quantity : Number(raw?.quantity);
  const unitPrice = lines.length ? lines[0].unitPrice : Number(raw?.unitPrice);
  const issueDate = clean(raw?.issueDate, 10) || todayBogota(); // la completa ContaGO al transmitir: por defecto, hoy
  const missing = [
    !name && "nombre del proveedor", !idDigits && "número de identificación del proveedor", !address && "dirección del proveedor",
    resident && !/^\d{5}$/.test(cityCode) && "municipio del proveedor",
    !resident && (!/^[A-Z]{2}$/.test(country) || country === "CO") && "país del proveedor (distinto de Colombia)",
    !resident && !cityName && "ciudad del proveedor", !concept && "descripción de lo adquirido",
  ].filter(Boolean);
  if (missing.length) throw new Error(`Falta: ${missing.join(", ")}.`);
  if (!(quantity > 0)) throw new Error("La cantidad debe ser mayor que cero.");
  if (!(unitPrice > 0)) throw new Error("El valor debe ser mayor que cero.");
  if (!DATE_RE.test(issueDate)) throw new Error("Fecha del documento inválida (AAAA-MM-DD).");
  if (issueDate > todayBogota()) throw new Error("La fecha del documento no puede ser futura.");
  const rete = Number(raw?.reteRentaPct ?? 0);
  if (rete !== 0 && !RETE_RENTA_RATES.includes(rete)) throw new Error(`La tarifa de retención ${rete}% no es válida para la DIAN. Tarifas: ${RETE_RENTA_RATES.join("; ")} (o sin retención).`);
  const notes = (Array.isArray(raw?.notes) ? raw.notes : []).map((n: unknown) => clean(n, 500)).filter(Boolean).slice(0, 5);
  const short = notes.find((n: string) => n.length < 15);
  if (short) throw new Error(`Cada nota debe tener al menos 15 caracteres (la DIAN lo exige): «${short}» tiene ${short.length}.`);
  const billDate = clean(raw?.bill?.date, 10);
  if (billDate && !DATE_RE.test(billDate)) throw new Error("Fecha de la cuenta de cobro inválida.");
  return {
    supplier: {
      origin, organizationType: s.organizationType === 1 ? 1 : 2, idType, docType: clean(s.docType, 4) || undefined, id: idDigits, name, address, ...(resident ? { cityCode } : { country, cityName }),
      email: clean(s.email, 120) || undefined, phone: clean(s.phone, 30) || undefined, postalCode: clean(s.postalCode, 6) || undefined,
    },
    bill: { number: clean(raw?.bill?.number, 40) || undefined, date: billDate || undefined },
    concept, quantity, unitPrice, ...(lines.length > 0 && { lines }), issueDate, reteRentaPct: rete,
    payment: { form: raw?.payment?.form === "2" ? "2" : "1", days: Number(raw?.payment?.days) || undefined, method: clean(raw?.payment?.method, 3) || undefined },
    notes,
  };
}

/** Totales del documento: sin IVA; la retención no cambia el total del documento, solo lo que se paga al proveedor. */
export function supportTotals(input: Pick<SupportDocInput, "quantity" | "unitPrice" | "reteRentaPct"> & { lines?: SupportDocInput["lines"] }) {
  const partes = input.lines?.length ? input.lines.map(l => ({ price: l.unitPrice, quantity: l.quantity })) : [{ price: input.unitPrice, quantity: input.quantity }];
  const subtotal = round2(partes.reduce((t, p) => t + buildItem({ description: "x", price: p.price, quantity: p.quantity, unitCode: "94", taxes: [{ taxCode: "01", percentage: 0 }] }).subtotal, 0));
  const retention = round2((subtotal * (input.reteRentaPct || 0)) / 100);
  return { subtotal, retention, net: round2(subtotal - retention) };
}

/** Arma el payload de `POST /support-documents`. */
export function buildSupportDocPayload(args: { input: SupportDocInput; number: number; cfg: AlanubeIssuerConfig; resolution: SupportDocResolution }): Record<string, unknown> {
  const { input, cfg } = args;
  const s = input.supplier;
  const payment = validatePayment({ form: input.payment?.form, days: input.payment?.days, method: input.payment?.method });
  const partes = input.lines?.length ? input.lines.map(l => ({ concept: l.description, quantity: l.quantity, unitPrice: l.unitPrice })) : [{ concept: input.concept, quantity: input.quantity, unitPrice: input.unitPrice }];
  const items = partes.map(l => buildItem({
    // El documento soporte NO lleva el prefijo «Cloud Computing - » (eso es solo de las facturas y notas crédito de ContaGO)
    description: l.concept.replace(/\s+/g, " ").trim().slice(0, 300), price: l.unitPrice, quantity: l.quantity, unitCode: "94",
    // Sin `taxes` la DIAN rechaza (DSAU04): se declara el IVA al 0 %
    taxes: [{ taxCode: "01", percentage: 0 }], standardCode: { identificationId: "1", id: "999" },
  }));
  const totals = buildTotals(items);
  const rete = input.reteRentaPct || 0;
  const notes = [
    ...(input.bill?.number || input.bill?.date ? [`Cuenta de cobro ${input.bill?.number ? `No. ${input.bill.number}` : ""}${input.bill?.date ? ` del ${input.bill.date}` : ""}`.replace(/\s+/g, " ").trim()] : []),
    ...(input.notes || []),
  ].slice(0, 6);
  return {
    issueDate: input.issueDate,
    number: args.number,
    resolution: args.resolution,
    company: {
      id: cfg.companyId, organizationType: 1, identificationType: "31", identificationNumber: cfg.nit, dv: cfg.dv,
      name: cfg.name, regimeCode: cfg.regimeCode, taxCode: { id: cfg.taxCode }, email: cfg.email, ...(cfg.phone && { phone: cfg.phone }),
      address: { ...cfg.address, country: "CO" },
    },
    supplier: {
      name: s.name, origin: s.origin, organizationType: s.organizationType, identificationType: s.idType, identificationNumber: s.id,
      ...(s.idType === "31" && { dv: nitDv(s.id) }), regimeCode: "R-99-PN", taxCode: "ZZ",
      // En pruebas el correo del proveedor se sustituye por el del emisor para no escribirle a terceros
      email: /sandbox/i.test(process.env.ALANUBE_BASE_URL || "") ? cfg.email : (s.email || cfg.email), ...(s.phone && { phone: s.phone }),
      // Residente: municipio DIAN (y código postal si lo tiene). No residente: solo el país, que no puede ser CO
      address: s.origin === "10"
        ? { address: s.address, city: s.cityCode, ...(s.postalCode && { postalCode: s.postalCode }), country: "CO" }
        : { address: s.address, city: s.cityName, country: s.country },
    },
    items,
    ...(rete > 0 && { withholdingTaxTotal: [{ taxCode: "06", taxPercentage: rete.toFixed(2), taxableAmount: totals.grossTotal, taxAmount: round2((totals.grossTotal * rete) / 100) }] }),
    payments: [{ paymentForm: payment.form, paymentMethod: payment.method, ...(payment.dueDate && { paymentDueDate: payment.dueDate }) }],
    totalAmounts: totals,
    ...(notes.length && { notes }),
  };
}

// ─── Persistencia ────────────────────────────────────────────────────────────

export interface Actor { userId: string; isAdmin: boolean }

const oid = (id: string) => { if (!ObjectId.isValid(id)) throw new Error("Documento inválido."); return new ObjectId(id); };
/** Cada usuario ve solo lo suyo; el administrador ve todo. */
const scope = (a: Actor) => (a.isAdmin ? {} : { ownerId: a.userId });

export async function listSupportDocs(a: Actor) {
  const rows = await docs().find(scope(a), { projection: { attempts: 0 } }).sort({ createdAt: -1 }).limit(500).toArray();
  return rows.map(r => ({ ...r, id: String(r._id), _id: undefined }));
}

export async function getSupportDoc(a: Actor, id: string) {
  const d = await docs().findOne({ _id: oid(id), ...scope(a) });
  if (!d) throw new Error("Documento no encontrado.");
  return d;
}

/**
 * Fecha, retención, forma de pago y notas son datos de TRANSMISIÓN: los completa ContaGO (administrador). Cuando guarda alguien
 * más se conservan los que ya tenía el documento o los valores por defecto, sin importar lo que llegue.
 */
export function mergeTransmission(incoming: SupportDocInput, existing: SupportDocInput | undefined, isAdmin: boolean): SupportDocInput {
  if (isAdmin) return incoming;
  const base = existing ?? normalizeSupportDoc({ ...incoming, issueDate: undefined, reteRentaPct: 0, payment: undefined, notes: [] });
  return { ...incoming, issueDate: base.issueDate, reteRentaPct: base.reteRentaPct, payment: base.payment, notes: base.notes };
}

/** Formato del número de radicado: RAD-000123. */
export const formatRadicado = (n: number) => `RAD-${String(n).padStart(6, "0")}`;

let radicadoIndexReady = false;
/** Siguiente número de radicado (contador atómico) y garantía de unicidad con un índice. */
async function nextRadicado(): Promise<string> {
  if (!radicadoIndexReady) { await docs().createIndex({ radicado: 1 }, { unique: true, sparse: true }); radicadoIndexReady = true; }
  const r: any = await counters().findOneAndUpdate({ _id: "RAD:supportDocs" } as any, { $inc: { seq: 1 } }, { upsert: true, returnDocument: "after" });
  return formatRadicado((r?.value ?? r).seq as number);
}

/** Migración única: asigna número de radicado a los documentos anteriores a este campo, en orden de creación. */
export async function assignMissingRadicados(): Promise<number> {
  const pending = await docs().find({ radicado: { $exists: false } }).sort({ createdAt: 1 }).project({ _id: 1, createdAt: 1 }).toArray();
  for (const d of pending) await docs().updateOne({ _id: d._id }, { $set: { radicado: await nextRadicado(), radicadoAt: d.createdAt } });
  return pending.length;
}

export async function saveSupportDoc(a: Actor, raw: any, id?: string): Promise<{ id: string; radicado: string; created: boolean }> {
  let input = normalizeSupportDoc(raw);
  if (!a.isAdmin) {
    const existing = id ? (await docs().findOne({ _id: oid(id), ...scope(a) }))?.input : undefined;
    input = mergeTransmission(input, existing, false);
  }
  const now = new Date().toISOString();
  if (!id) {
    const radicado = await nextRadicado();
    const r = await docs().insertOne({ ownerId: a.userId, status: "borrador" as SupportDocStatus, input, radicado, radicadoAt: now, createdAt: now, updatedAt: now, hasPdf: false });
    return { id: String(r.insertedId), radicado, created: true };
  }
  const r = await docs().findOneAndUpdate({ _id: oid(id), ...scope(a), status: { $in: ["borrador", "rechazado", "error"] } }, { $set: { input, updatedAt: now } }, { returnDocument: "after" });
  const updated: any = (r as any)?.value ?? r;
  if (!updated) throw new Error("Documento no encontrado o ya emitido: un documento emitido no se modifica.");
  return { id, radicado: updated.radicado || "", created: false };
}

/** Datos del proveedor que se reutilizan en una nueva cuenta de cobro (lo demás —valores, cuenta de cobro— cambia cada vez). */
export function supplierSnapshot(s: SupportDocInput["supplier"]) {
  return {
    docType: s.docType, organizationType: s.organizationType, name: s.name, email: s.email, phone: s.phone, address: s.address,
    cityCode: s.cityCode, country: s.country, cityName: s.cityName, postalCode: s.postalCode,
  };
}

/**
 * Datos de un proveedor que ya radicó antes (los del radicado más reciente con ese documento). Cada usuario solo ve los
 * suyos; el administrador ve los de todos. No distingue tipo de documento: el mismo número es la misma persona.
 */
export async function findPreviousSupplier(a: Actor, docType: string, rawId: string) {
  const numeric = ["CC", "TI", "NIT"].includes(docType);
  const id = numeric ? String(rawId ?? "").replace(/\D/g, "") : String(rawId ?? "").replace(/[^A-Za-z0-9]/g, "");
  if (id.length < 5) return null;
  const d: any = await docs().findOne({ "input.supplier.id": id, ...scope(a) }, { sort: { createdAt: -1 }, projection: { input: 1, radicado: 1 } });
  return d ? { supplier: supplierSnapshot(d.input.supplier), radicado: d.radicado as string | undefined } : null;
}

// ─── Enlace público para que los proveedores radiquen sus cuentas de cobro ──────────────────────

const settings = () => getDb().collection<any>("supportDocsSettings");
const newToken = () => randomBytes(24).toString("base64url");

/** Token del enlace público (se crea la primera vez). Solo lo ve el administrador. */
export async function getPublicLinkToken(): Promise<string> {
  const cur = await settings().findOne({ _id: "publicLink" } as any);
  if (cur?.token) return cur.token as string;
  const token = newToken();
  await settings().updateOne({ _id: "publicLink" } as any, { $setOnInsert: { token, createdAt: new Date().toISOString() } }, { upsert: true });
  return (await settings().findOne({ _id: "publicLink" } as any)).token as string;
}

/** Genera un token nuevo: el enlace anterior deja de funcionar. */
export async function rotatePublicLinkToken(): Promise<string> {
  const token = newToken();
  await settings().updateOne({ _id: "publicLink" } as any, { $set: { token, rotatedAt: new Date().toISOString() } }, { upsert: true });
  return token;
}

/** Comparación en tiempo constante del token del enlace público. */
export async function isValidPublicToken(token: string): Promise<boolean> {
  const cur = await settings().findOne({ _id: "publicLink" } as any);
  if (!cur?.token || typeof token !== "string" || !token) return false;
  const a = Buffer.from(cur.token), b = Buffer.from(token);
  return a.length === b.length && timingSafeEqual(a, b);
}

/** Quien radica por el enlace público no tiene usuario: sus documentos quedan a nombre de "public" (el administrador los ve todos). */
export const PUBLIC_ACTOR: Actor = { userId: "public", isAdmin: false };

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;

/**
 * Radicación desde el enlace público. El correo es obligatorio (es el contacto y, además, lo que permite cargar sus datos la
 * próxima vez). Si el PDF no se puede guardar, la radicación se deshace: nunca queda un radicado sin soporte.
 */
export async function createPublicRadicado(raw: any, file: { buffer: Buffer; originalname: string; mimetype: string } | undefined, extraFiles: { buffer: Buffer; originalname: string }[], meta: { ip?: string }) {
  if (!file) throw new Error("Adjunta el PDF de la cuenta de cobro: es obligatorio.");
  // Se valida todo ANTES de reservar el número de radicado: un archivo inválido no debe dejar saltos en la numeración
  if (file.mimetype !== "application/pdf" || file.buffer.subarray(0, 5).toString() !== "%PDF-") throw new Error("El archivo de la cuenta de cobro debe ser un PDF.");
  if (extraFiles.length > MAX_EXTRAS) throw new Error(`Máximo ${MAX_EXTRAS} documentos adicionales.`);
  for (const x of extraFiles) validateAttachment(x);
  const input = normalizeSupportDoc(raw);
  if (!input.supplier.email || !EMAIL_RE.test(input.supplier.email)) throw new Error("Escribe un correo electrónico válido: ContaGO lo usará para contactarte.");
  const saved = await saveSupportDoc(PUBLIC_ACTOR, raw);
  try {
    await attachBillPdf(PUBLIC_ACTOR, saved.id, file);
    for (const x of extraFiles) await attachExtra(PUBLIC_ACTOR, saved.id, x);
  } catch (e) {
    await deleteSupportDoc(PUBLIC_ACTOR, saved.id).catch(() => {});
    throw e;
  }
  await docs().updateOne({ _id: oid(saved.id) }, { $set: { source: "public", submitterIp: meta.ip || null } });
  return { radicado: saved.radicado, supplier: input.supplier.name, amount: supportTotals(input).subtotal };
}

/**
 * Datos de un proveedor que ya radicó antes, solo con su número de documento (decisión de ContaGO: más cómodo para el
 * proveedor). OJO: en una página pública esto revela los datos de quien tenga ese documento; por eso solo hay coincidencia
 * exacta del número, un límite de consultas por IP y se devuelven únicamente los datos del proveedor (nunca valores ni archivos).
 */
export async function previousSupplierPublic(docType: string, rawId: string) {
  const numeric = ["CC", "TI", "NIT"].includes(docType);
  const id = numeric ? String(rawId ?? "").replace(/\D/g, "") : String(rawId ?? "").replace(/[^A-Za-z0-9]/g, "");
  if (id.length < 5) return null;
  const d: any = await docs().findOne({ "input.supplier.id": id }, { sort: { createdAt: -1 }, projection: { input: 1 } });
  return d ? supplierSnapshot(d.input.supplier) : null;
}

// ─── Documentos adicionales (soportes opcionales de la cuenta de cobro; no viajan a la DIAN) ──────────

const extrasCol = () => getDb().collection<any>("supportDocExtras");
export const MAX_EXTRAS = 5;

/** Solo PDF, PNG o JPG; se comprueba el contenido real del archivo, no la extensión ni el tipo que declare el navegador. */
export function validateAttachment(f: { buffer: Buffer; originalname: string }): { mime: string } {
  const b = f.buffer;
  if (b.length > 10 * 1024 * 1024) throw new Error(`«${f.originalname}» pesa más de 10 MB.`);
  if (b.subarray(0, 5).toString() === "%PDF-") return { mime: "application/pdf" };
  if (b.length > 8 && b[0] === 0x89 && b.subarray(1, 4).toString() === "PNG") return { mime: "image/png" };
  if (b.length > 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return { mime: "image/jpeg" };
  throw new Error(`«${f.originalname}»: solo se admiten archivos PDF, PNG o JPG.`);
}

export async function attachExtra(a: Actor, docId: string, file: { buffer: Buffer; originalname: string }) {
  const d = await getSupportDoc(a, docId);
  if (!a.isAdmin && !["borrador", "rechazado", "error"].includes(d.status)) throw new Error("El documento ya fue emitido: no se pueden agregar archivos.");
  if ((d.extras?.length || 0) >= MAX_EXTRAS) throw new Error(`Máximo ${MAX_EXTRAS} documentos adicionales.`);
  const { mime } = validateAttachment(file);
  const meta = { id: new ObjectId().toHexString(), name: file.originalname.slice(0, 200), size: file.buffer.length, mime };
  await extrasCol().insertOne({ _id: meta.id, docId, ...meta, data: new Binary(file.buffer), uploadedAt: new Date().toISOString() } as any);
  await docs().updateOne({ _id: oid(docId) }, { $push: { extras: meta }, $set: { updatedAt: new Date().toISOString() } } as any);
  return meta;
}

export async function getExtra(a: Actor, docId: string, extraId: string) {
  await getSupportDoc(a, docId);
  const e: any = await extrasCol().findOne({ _id: extraId, docId } as any);
  return e ? { data: Buffer.from(e.data.buffer), filename: e.name as string, contentType: e.mime as string } : null;
}

export async function deleteExtra(a: Actor, docId: string, extraId: string) {
  const d = await getSupportDoc(a, docId);
  if (!a.isAdmin && !["borrador", "rechazado", "error"].includes(d.status)) throw new Error("El documento ya fue emitido: no se pueden quitar archivos.");
  await extrasCol().deleteOne({ _id: extraId, docId } as any);
  await docs().updateOne({ _id: oid(docId) }, { $pull: { extras: { id: extraId } } } as any);
}

export async function deleteSupportDoc(a: Actor, id: string) {
  const r = await docs().deleteOne({ _id: oid(id), ...scope(a), status: { $in: ["borrador", "error"] } });
  if (!r.deletedCount) throw new Error("Solo se pueden eliminar borradores sin emitir.");
  await docFiles().deleteOne({ _id: id } as any);
  await extrasCol().deleteMany({ docId: id } as any);
}

export async function attachBillPdf(a: Actor, id: string, file: { buffer: Buffer; originalname: string; mimetype: string }) {
  await getSupportDoc(a, id);
  if (file.mimetype !== "application/pdf" || file.buffer.subarray(0, 5).toString() !== "%PDF-") throw new Error("El archivo debe ser un PDF.");
  await docFiles().updateOne({ _id: id } as any, { $set: { data: new Binary(file.buffer), originalName: file.originalname.slice(0, 200), size: file.buffer.length, uploadedAt: new Date().toISOString() } }, { upsert: true });
  await docs().updateOne({ _id: oid(id) }, { $set: { hasPdf: true, pdfName: file.originalname.slice(0, 200), updatedAt: new Date().toISOString() } });
}

export async function getBillPdf(a: Actor, id: string) {
  await getSupportDoc(a, id);
  const f = await docFiles().findOne({ _id: id } as any);
  return f ? { data: Buffer.from(f.data.buffer), filename: f.originalName as string } : null;
}

/**
 * Archivos del documento emitido. Alanube solo entrega los XML (el documento y la respuesta de la DIAN), así que el PDF se
 * genera aquí (con CUDS y QR, la primera vez que se pide) y el ZIP lo arma ContaGO con los tres archivos.
 */
export async function getEmittedFile(a: Actor, id: string, kind: "zip" | "pdf" | "xml") {
  const d = await getSupportDoc(a, id);
  const em = d.emission;
  if (!em?.alanubeId || !d.emittedPayload) return null;
  const f = await alanubeFiles().findOne({ _id: em.alanubeId } as any);
  const xml: Buffer | null = f?.xml ? Buffer.from(f.xml.buffer) : null;
  const name = String(em.fullNumber);
  if (kind === "xml") return xml ? { data: xml, filename: `${name}.xml`, contentType: "application/xml" } : null;

  let pdf: Buffer | null = f?.pdf ? Buffer.from(f.pdf.buffer) : null;
  if (!pdf) {
    pdf = await renderSupportDocPdf({ payload: d.emittedPayload, fullNumber: em.fullNumber, cuds: em.cude, legalStatus: em.legalStatus, date: em.date, qrCodeContent: em.qrCodeContent });
    await alanubeFiles().updateOne({ _id: em.alanubeId } as any, { $set: { pdf: new Binary(pdf), pdfAt: new Date().toISOString() } }, { upsert: true });
  }
  if (kind === "pdf") return { data: pdf, filename: `${name}.pdf`, contentType: "application/pdf" };

  if (!xml) return null;
  const JSZip = (await import("jszip")).default;
  const base = String(em.xmlFileName || name).replace(/\.xml$/i, "");
  const zip = new JSZip();
  zip.file(`${base}.xml`, xml);
  if (f?.applicationResponse) zip.file(`${base}-respuesta-dian.xml`, Buffer.from(f.applicationResponse.buffer));
  zip.file(`${base}.pdf`, pdf);
  return { data: await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }), filename: `${name}.zip`, contentType: "application/zip" };
}

// ─── Emisión ─────────────────────────────────────────────────────────────────

/** Reserva el siguiente número: contador atómico + registro único (empresa/prefijo/número) para no repetirlo jamás. */
async function reserveNumber(cfg: AlanubeIssuerConfig, res: SupportDocResolution, docId: string): Promise<number> {
  const id = `DSE:${cfg.companyId}:${res.prefix}`;
  const first = Number(process.env.ALANUBE_DSE_FIRST_NUMBER || res.minNumber + 1);
  await counters().updateOne({ _id: id } as any, { $max: { seq: first - 1 } }, { upsert: true });
  const r: any = await counters().findOneAndUpdate({ _id: id } as any, { $inc: { seq: 1 } }, { returnDocument: "after" });
  const seq: number = (r?.value ?? r).seq;
  if (!(seq >= res.minNumber && seq <= res.maxNumber)) throw new Error("El consecutivo quedó fuera del rango de la resolución.");
  await numbers().insertOne({ _id: `${cfg.companyId}:${res.prefix}:${seq}`, docId, at: new Date().toISOString() } as any); // falla si el número ya existía
  return seq;
}

export interface SupportEmission {
  number: number; prefix: string; fullNumber?: string; cude?: string; legalStatus?: string; accepted?: boolean;
  errors: string[]; alanubeId?: string; at: string; errorMessage?: string;
  /** Para la representación gráfica (PDF). */
  date?: string; qrCodeContent?: string; xmlFileName?: string;
}

/**
 * Emite el documento soporte de PRUEBA. Orden: 1) candado atómico (estado "emitiendo", así un doble clic no emite dos
 * veces), 2) reserva del número, 3) llamada a Alanube, 4) resultado y archivos oficiales. Si la llamada falla el número
 * queda registrado como usado (puede haber un salto, nunca un duplicado).
 */
export async function emitSupportDoc(a: Actor, id: string): Promise<SupportEmission> {
  if (!a.isAdmin) throw new Error("La transmisión a la DIAN la hace ContaGO: tu documento queda guardado y ContaGO lo emite.");
  if (!/sandbox/i.test(process.env.ALANUBE_BASE_URL || "")) throw new Error("El documento soporte de prueba solo se emite contra el sandbox de Alanube.");
  const pre: any = await docs().findOne({ _id: oid(id), ...scope(a) }, { projection: { hasPdf: 1, status: 1 } });
  if (pre && !pre.hasPdf && ["borrador", "rechazado", "error"].includes(pre.status)) throw new Error("A este radicado le falta el PDF de la cuenta de cobro: adjúntalo desde «Editar» antes de emitir.");
  const locked = await docs().findOneAndUpdate(
    { _id: oid(id), ...scope(a), status: { $in: ["borrador", "rechazado", "error"] } },
    { $set: { status: "emitiendo" as SupportDocStatus, updatedAt: new Date().toISOString() } },
    { returnDocument: "after" },
  );
  const doc: any = (locked as any)?.value ?? locked;
  if (!doc || !doc.input) throw new Error("El documento no existe o ya está emitido o en proceso de emisión.");

  const cfg = loadIssuerConfig();
  const res = loadSupportResolution();
  const at = new Date().toISOString();
  let number = 0;
  try {
    number = await reserveNumber(cfg, res, id);
    const payload = buildSupportDocPayload({ input: doc.input, number, cfg, resolution: res });
    const r = await createSupportDocument(payload);
    const accepted = isInvoiceAccepted(r);
    const errors = r.governmentResponse?.errorMessages || [];
    // Los enlaces de Alanube caducan: los XML se guardan ya (si falla no se pierde el documento, solo queda sin archivo)
    const stored: Record<string, unknown> = { fullNumber: r.fullNumber, savedAt: at };
    for (const [key, url] of [["xml", r.xmlUrl], ["applicationResponse", r.applicationResponseUrl]] as const) {
      if (!url) continue;
      try { stored[key] = new Binary(await downloadFile(url)); } catch (e: any) { console.warn(`[SoporteDoc] ${key}:`, e?.message); }
    }
    await alanubeFiles().updateOne({ _id: r.id } as any, { $set: stored }, { upsert: true });
    const emission: SupportEmission = { number, prefix: res.prefix, fullNumber: r.fullNumber, cude: r.cuds || r.cufe, legalStatus: r.legalStatus, accepted, errors, alanubeId: r.id, at, date: r.date, qrCodeContent: r.qrCodeContent, xmlFileName: r.xmlFileName };
    await docs().updateOne({ _id: oid(id) }, { $set: { status: (accepted ? "emitido" : "rechazado") as SupportDocStatus, emission, emittedPayload: payload, updatedAt: new Date().toISOString() }, $push: { attempts: { number, at, outcome: r.legalStatus } } } as any);
    return emission;
  } catch (e: any) {
    const message = e instanceof AlanubeError ? e.message : (e?.message || "Error desconocido");
    await docs().updateOne({ _id: oid(id) }, { $set: { status: "error" as SupportDocStatus, emission: { number, prefix: res.prefix, errors: [], at, errorMessage: message }, updatedAt: new Date().toISOString() }, $push: { attempts: { number, at, outcome: "error", message } } } as any);
    throw new Error(message);
  }
}
