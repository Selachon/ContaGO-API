/**
 * Capa ERP de "Caja por proyecto" — endpoints que viven ENCIMA de los motores
 * de causación (no los tocan). Gated por el tool `causacion-caja`, scoped por
 * empresa vía header X-Siigo-Company (mismo patrón que el resto de tools SIIGO).
 *
 *   GET    /caja-erp/proyectos            lista proyectos
 *   POST   /caja-erp/proyectos            crea proyecto
 *   PATCH  /caja-erp/proyectos/:id        edita proyecto
 *   DELETE /caja-erp/proyectos/:id        elimina proyecto
 *   GET    /caja-erp/tags?kind=mov|invoice   mapa refId→proyectoId
 *   PUT    /caja-erp/tags                 { kind, refId, proyectoId } asigna/limpia
 */
import { Router, type Request, type Response } from "express";
import multer from "multer";
import { requireAuth } from "../middleware/auth.js";
import { requireToolAccess } from "../middleware/requireToolAccess.js";
import { userCanAccessCompany, setCompanyDrive, getCompanyDrive, getCompanyContext } from "../services/siigoCompaniesService.js";
import { runWithSiigoCompany, getPurchaseSupportDocumentById } from "../services/siigoService.js";
import { ingestNewByDateRange, type IngestResult } from "../services/siigoDianIngestService.js";
import { latestDianTokenForNit, getDianMailboxStatus } from "../services/dianMailboxService.js";
import { v4 as uuidv4 } from "uuid";
import { getUserGoogleDrives, getUserGoogleDriveById, updateUserDriveTokens, getDb } from "../services/database.js";
import { ObjectId } from "mongodb";
import { uploadInvoiceFilesToDrive, uploadPaymentSupportToDrive } from "../services/googleDrive.js";
import { encryptToken } from "../utils/encryption.js";
import { addMovements } from "../services/siigoEgresosStoreService.js";
import {
  listProyectos,
  createProyecto,
  updateProyecto,
  deleteProyecto,
  getTags,
  setTag,
  listEntries,
  createEntry,
  updateEntry,
  deleteEntry,
  listInvoices,
  saveInvoice,
  getInvoiceByKey,
  updateInvoiceByKey,
  setInvoicePaid,
  deleteInvoiceByKey,
  listPayments,
  createPayment,
  updatePayment,
  deletePayment,
  getPayment,
  setPaymentItemSupport,
  getProgrammedKeys,
  listBancos,
  createBanco,
  updateBanco,
  deleteBanco,
  listPlanes,
  savePlan,
  deletePlan,
  type TagKind,
} from "../services/cajaErpStore.js";
import ExcelJS from "exceljs";
import { listMovements } from "../services/siigoEgresosStoreService.js";
import { generarCajaPdf } from "../services/cajaPdf.js";

export const CAJA_ERP_TOOL_ID = "causacion-caja";

const router = Router();
router.use(requireAuth);
router.use(requireToolAccess(CAJA_ERP_TOOL_ID));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024, files: 1 } });

/** Resuelve el Drive ligado a la empresa: config + callback de refresh + NIT. */
async function resolveCompanyDrive(companyId: string) {
  const link = await getCompanyDrive(companyId);
  if (!link) return null;
  const driveConfig = await getUserGoogleDriveById(link.ownerUserId, link.connectionId);
  if (!driveConfig) return null;
  const onTokenRefresh = async (token: string, expiry: number) => {
    await updateUserDriveTokens(link.ownerUserId, encryptToken(token), new Date(expiry).toISOString(), link.connectionId);
  };
  return { driveConfig, ownerUserId: link.ownerUserId, nit: link.nit, onTokenRefresh };
}

async function resolveCompany(req: Request, res: Response): Promise<string | null> {
  const companyId = req.header("X-Siigo-Company");
  if (!companyId) {
    res.status(400).json({ ok: false, message: "Falta la empresa (X-Siigo-Company)." });
    return null;
  }
  const userId = req.user?.userId;
  const allowed = userId ? await userCanAccessCompany(companyId, userId) : false;
  if (!allowed) {
    res.status(403).json({ ok: false, message: "No tiene acceso a esta empresa." });
    return null;
  }
  return companyId;
}

const isKind = (k: unknown): k is TagKind => k === "mov" || k === "invoice";

// ── Proyectos ──────────────────────────────────────────────────────────
router.get("/proyectos", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  res.json({ ok: true, data: await listProyectos(companyId) });
});

router.post("/proyectos", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const nombre = String(req.body?.nombre || "").trim();
  if (!nombre) return res.status(400).json({ ok: false, message: "El nombre es obligatorio." });
  const p = await createProyecto(companyId, {
    nombre,
    esTransversal: !!req.body?.esTransversal,
    estado: req.body?.estado === "cerrado" ? "cerrado" : "activo",
    saldoInicial: Number(req.body?.saldoInicial) || 0,
  });
  res.json({ ok: true, data: p });
});

router.patch("/proyectos/:id", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const updated = await updateProyecto(companyId, req.params.id, req.body || {});
  if (!updated) return res.status(404).json({ ok: false, message: "Proyecto no encontrado." });
  res.json({ ok: true, data: updated });
});

router.delete("/proyectos/:id", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const ok = await deleteProyecto(companyId, req.params.id);
  res.json({ ok, data: { deleted: ok } });
});

// ── Etiquetas proyecto↔referencia ──────────────────────────────────────
router.get("/tags", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const kind = req.query.kind;
  if (!isKind(kind)) return res.status(400).json({ ok: false, message: "kind debe ser 'mov' o 'invoice'." });
  res.json({ ok: true, data: await getTags(companyId, kind) });
});

router.put("/tags", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const { kind, refId, proyectoId } = req.body || {};
  if (!isKind(kind) || !refId) return res.status(400).json({ ok: false, message: "Faltan kind/refId válidos." });
  await setTag(companyId, kind, String(refId), String(proyectoId || ""));
  res.json({ ok: true });
});

// ── Gastos/ingresos manuales (solo ERP, NO van a SIIGO) ────────────────
router.get("/entries", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const proyectoId = typeof req.query.proyectoId === "string" ? req.query.proyectoId : undefined;
  res.json({ ok: true, data: await listEntries(companyId, proyectoId) });
});

router.post("/entries", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const b = req.body || {};
  if (!b.proyectoId) return res.status(400).json({ ok: false, message: "Falta el proyecto." });
  if (!(Number(b.valor) > 0)) return res.status(400).json({ ok: false, message: "El valor debe ser mayor que cero." });
  res.json({ ok: true, data: await createEntry(companyId, b) });
});

router.patch("/entries/:id", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const updated = await updateEntry(companyId, req.params.id, req.body || {});
  if (!updated) return res.status(404).json({ ok: false, message: "Gasto no encontrado." });
  res.json({ ok: true, data: updated });
});

router.delete("/entries/:id", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const ok = await deleteEntry(companyId, req.params.id);
  res.json({ ok, data: { deleted: ok } });
});

// ── Registro persistente de facturas (buzón) ───────────────────────────
router.get("/invoices", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  res.json({ ok: true, data: await listInvoices(companyId) });
});

// Guarda una factura CAUSADA en el buzón (con totales, consecutivo, proyecto, PDF).
router.post("/invoices/save", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const rec = req.body?.invoice || {};
  if (!rec.key && !rec.cufe && !(rec.supplierNit && rec.docNumber)) {
    return res.status(400).json({ ok: false, message: "Falta identificar la factura (key/cufe)." });
  }
  await saveInvoice(companyId, { ...rec, estado: "causada" });
  res.json({ ok: true });
});

router.post("/invoices/paid", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const { key, paid, fechaPago } = req.body || {};
  if (!key) return res.status(400).json({ ok: false, message: "Falta key." });
  await setInvoicePaid(companyId, String(key), paid !== false, fechaPago);
  res.json({ ok: true });
});

router.delete("/invoices", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const key = typeof req.query.key === "string" ? req.query.key : "";
  if (!key) return res.status(400).json({ ok: false, message: "Falta key." });
  const ok = await deleteInvoiceByKey(companyId, key);
  res.json({ ok, data: { deleted: ok } });
});

// Reconsulta a SIIGO el estado del timbrado electrónico (DIAN) de un documento
// soporte (DS) ya causado y actualiza el buzón. Útil cuando al causar quedó en
// "procesando" y luego la DIAN lo acepta/rechaza.
function normalizeDianStamp(stamp: any) {
  if (!stamp || typeof stamp !== "object") return { state: "unknown", ok: false, label: "Sin información de la DIAN", errors: [] };
  const raw = String(stamp.status || "").toLowerCase();
  const cuds = stamp.cuds || stamp.cude || "";
  const errors = Array.isArray(stamp.errors) ? stamp.errors.map((e: any) => e?.message || e?.Message || String(e)).filter(Boolean) : [];
  if (raw === "accepted" || raw === "sent" || (cuds && !raw.includes("reject"))) return { state: "accepted", ok: true, cuds, label: "Aceptado por la DIAN", errors: [] };
  if (raw === "draft" || raw === "") return { state: "draft", ok: false, cuds, label: "Sin transmitir (borrador en SIIGO)", errors };
  if (raw.includes("reject")) return { state: "rejected", ok: false, cuds, label: "Rechazado por la DIAN", errors };
  return { state: raw || "pending", ok: false, cuds, label: `Estado DIAN: ${stamp.status}`, errors };
}

router.post("/invoices/dian-status", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const key = (typeof req.query.key === "string" ? req.query.key : req.body?.key) || "";
  if (!key) return res.status(400).json({ ok: false, message: "Falta key." });
  const inv = await getInvoiceByKey(companyId, String(key));
  if (!inv) return res.status(404).json({ ok: false, message: "Factura no encontrada en el buzón." });
  if (!inv.siigoId) return res.status(400).json({ ok: false, message: "Este documento no tiene id de SIIGO; no se puede reconsultar." });
  const ctx = await getCompanyContext(companyId);
  if (!ctx) return res.status(400).json({ ok: false, message: "Empresa sin conexión a SIIGO." });
  try {
    const doc: any = await runWithSiigoCompany(ctx, () => getPurchaseSupportDocumentById(inv.siigoId));
    const dian = normalizeDianStamp(doc?.stamp);
    await updateInvoiceByKey(companyId, String(key), { dianStamp: dian });
    res.json({ ok: true, data: dian });
  } catch (err) {
    res.status(502).json({ ok: false, message: err instanceof Error ? err.message : "Error consultando SIIGO." });
  }
});

// ── Programación de pagos (lotes) ──────────────────────────────────────
router.get("/payments", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  res.json({ ok: true, data: await listPayments(companyId) });
});

// Keys de facturas ya programadas (para excluirlas de la selección).
router.get("/payments/programmed-keys", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  res.json({ ok: true, data: await getProgrammedKeys(companyId) });
});

router.post("/payments", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const items = Array.isArray(req.body?.items) ? req.body.items : [];
  if (!items.length) return res.status(400).json({ ok: false, message: "La programación necesita al menos un documento." });
  res.json({ ok: true, data: await createPayment(companyId, req.body || {}) });
});

router.patch("/payments/:id", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const updated = await updatePayment(companyId, req.params.id, req.body || {});
  if (!updated) return res.status(404).json({ ok: false, message: "Programación no encontrada." });
  res.json({ ok: true, data: updated });
});

router.delete("/payments/:id", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const ok = await deletePayment(companyId, req.params.id);
  res.json({ ok, data: { deleted: ok } });
});

// Envía una programación al módulo de Egresos: crea un movimiento (egreso) por
// cada ítem en la cuenta bancaria del lote, listo para causar el RP. Marca enviado.
router.post("/payments/:id/send-to-egresos", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const prog = await getPayment(companyId, req.params.id);
  if (!prog) return res.status(404).json({ ok: false, message: "Programación no encontrada." });
  if (prog.estado === "enviado") return res.status(409).json({ ok: false, message: "Esta programación ya fue enviada a Egresos." });
  if (!prog.items.length) return res.status(400).json({ ok: false, message: "La programación no tiene documentos." });

  const movs = prog.items.map((it) => ({
    bankAccountId: prog.bankAccountId,
    date: prog.fecha,
    value: it.valor,
    description: it.kind === "manual"
      ? `${it.tipo} · ${it.supplierName}${it.descripcion ? " · " + it.descripcion : ""}`
      : `${it.supplierName}${it.descripcion ? " · " + it.descripcion : ""}`,
    nit: it.supplierNit,
    kind: "egreso" as const,
    direction: "out" as const,
    source: "manual" as const,
  }));

  try {
    const created = await addMovements(companyId, movs);
    await updatePayment(companyId, prog.id, { estado: "enviado" });
    res.json({ ok: true, data: { creados: created.length } });
  } catch (e) {
    res.status(502).json({ ok: false, message: e instanceof Error ? e.message : "Error creando movimientos en Egresos." });
  }
});

// Sube el soporte de PAGO de un ítem de una programación al Drive de la empresa
// y lo asocia (y a la factura si aplica → aparece en el listado Facturas).
router.post("/payments/:id/support", upload.single("file"), async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const file = req.file;
  const itemIndex = Number(req.body?.itemIndex);
  const { date } = req.body || {};
  if (!file || !Number.isInteger(itemIndex)) return res.status(400).json({ ok: false, message: "Falta archivo o itemIndex." });
  const drive = await resolveCompanyDrive(companyId);
  if (!drive) return res.status(409).json({ ok: false, message: "La empresa no tiene Drive vinculado." });
  try {
    const url = await uploadPaymentSupportToDrive(
      file.buffer, file.originalname, file.mimetype,
      String(drive.nit || "").replace(/\D/g, ""), String(date || ""),
      drive.driveConfig, drive.ownerUserId, drive.onTokenRefresh,
    );
    const updated = await setPaymentItemSupport(companyId, req.params.id, itemIndex, url);
    if (!updated) return res.status(404).json({ ok: false, message: "Programación no encontrada." });
    res.json({ ok: true, data: { url, program: updated } });
  } catch (e) {
    res.status(502).json({ ok: false, message: e instanceof Error ? e.message : "Error subiendo el soporte." });
  }
});

// ── Google Drive de la empresa (para archivar las facturas) ────────────
router.get("/drive/status", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const link = await getCompanyDrive(companyId);
  const available = (await getUserGoogleDrives(req.user!.userId)).map((d) => ({ connectionId: d.connection_id, email: d.user_email }));
  let linkedEmail = "";
  if (link) {
    const cfg = await getUserGoogleDriveById(link.ownerUserId, link.connectionId);
    linkedEmail = cfg?.user_email || "";
  }
  res.json({ ok: true, data: { linked: !!link, linkedEmail, connectionId: link?.connectionId || "", available } });
});

router.post("/drive/link", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const connectionId = String(req.body?.connectionId || "");
  if (!connectionId) return res.status(400).json({ ok: false, message: "Falta connectionId." });
  const cfg = await getUserGoogleDriveById(req.user!.userId, connectionId);
  if (!cfg) return res.status(404).json({ ok: false, message: "Conexión de Drive no encontrada." });
  await setCompanyDrive(companyId, req.user!.userId, connectionId);
  res.json({ ok: true, data: { linked: true, linkedEmail: cfg.user_email } });
});

// Archiva el PDF de una factura en el Drive de la empresa (misma estructura que
// el tool de Descarga Masiva) y guarda el enlace en el registro (pdfUrl).
router.post("/invoices/archive", upload.single("file"), async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const file = req.file;
  const { key, docNumber, date } = req.body || {};
  if (!file || !key) return res.status(400).json({ ok: false, message: "Falta archivo o key." });
  const drive = await resolveCompanyDrive(companyId);
  if (!drive) return res.status(409).json({ ok: false, message: "La empresa no tiene Drive vinculado." });
  try {
    const r = await uploadInvoiceFilesToDrive(
      file.buffer, null,                       // solo PDF
      String(docNumber || key),                // nombre = número de factura
      String(drive.nit || "").replace(/\D/g, ""), // NIT de la empresa (receptor)
      String(date || ""),
      drive.driveConfig, drive.ownerUserId, drive.onTokenRefresh,
      "received",
    );
    // Si la factura YA está en el buzón (causada), persiste el enlace ahí mismo.
    // Las pendientes no están en el buzón: el frontend guarda el enlace en la
    // fila de Causación y lo pasa al causar (saveInvoice).
    const existing = await getInvoiceByKey(companyId, String(key));
    if (existing && r.pdfUrl) await updateInvoiceByKey(companyId, String(key), { pdfUrl: r.pdfUrl });
    res.json({ ok: true, data: { pdfUrl: r.pdfUrl || "" } });
  } catch (e) {
    res.status(502).json({ ok: false, message: e instanceof Error ? e.message : "Error archivando en Drive." });
  }
});

// ── Cuentas bancarias (saldo real para comparar con la caja) ───────────
router.get("/bancos", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  res.json({ ok: true, data: await listBancos(companyId) });
});
router.post("/bancos", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  if (!String(req.body?.nombre || "").trim()) return res.status(400).json({ ok: false, message: "El nombre es obligatorio." });
  res.json({ ok: true, data: await createBanco(companyId, req.body || {}) });
});
router.patch("/bancos/:id", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const u = await updateBanco(companyId, req.params.id, req.body || {});
  if (!u) return res.status(404).json({ ok: false, message: "Cuenta no encontrada." });
  res.json({ ok: true, data: u });
});
router.delete("/bancos/:id", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  res.json({ ok: await deleteBanco(companyId, req.params.id) });
});

// ── Reporte de caja en Excel (total o por proyecto) ────────────────────
// Reconstruye los saldos por proyecto = saldoInicial + ingresos − egresos
// (movimientos de banco etiquetados + entradas manuales).
async function buildCajaData(companyId: string) {
  const [proyectos, tagsMap, movs, entries, bancos] = await Promise.all([
    listProyectos(companyId),
    getTags(companyId, "mov"),
    listMovements(companyId),
    listEntries(companyId),
    listBancos(companyId),
  ]);
  const byProj = new Map<string, { items: any[]; in: number; out: number }>();
  const ensure = (pid: string) => { if (!byProj.has(pid)) byProj.set(pid, { items: [], in: 0, out: 0 }); return byProj.get(pid)!; };
  for (const m of movs as any[]) {
    const pid = (tagsMap as any)[m.id]; if (!pid) continue;
    const e = ensure(pid); const v = Number(m.value) || 0;
    if (m.direction === "in") e.in += v; else e.out += v;
    e.items.push({ fecha: m.date, descripcion: m.description, valor: v, direction: m.direction, origen: "Banco" });
  }
  for (const en of entries as any[]) {
    if (!en.proyectoId) continue;
    const e = ensure(en.proyectoId); const v = Number(en.valor) || 0;
    if (en.direction === "in") e.in += v; else e.out += v;
    e.items.push({ fecha: en.fecha, descripcion: en.descripcion, valor: v, direction: en.direction, origen: "Manual", categoria: en.categoria });
  }
  const rows = proyectos.map((p) => {
    const e = byProj.get(p.id) || { items: [], in: 0, out: 0 };
    return { proyecto: p, ingresos: e.in, egresos: e.out, saldo: (p.saldoInicial || 0) + e.in - e.out, items: e.items };
  });
  const totalCaja = rows.reduce((a, r) => a + r.saldo, 0);
  const totalBancos = bancos.reduce((a, b) => a + b.saldo, 0);
  return { rows, bancos, totalCaja, totalBancos };
}

router.get("/report/excel", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const proyectoId = typeof req.query.proyectoId === "string" ? req.query.proyectoId : "";
  const { rows, bancos, totalCaja, totalBancos } = await buildCajaData(companyId);
  const money = (n: number) => Math.round(Number(n) || 0);
  const wb = new ExcelJS.Workbook();

  if (proyectoId) {
    // Detalle de un proyecto: ingresos primero, luego egresos.
    const r = rows.find((x) => x.proyecto.id === proyectoId);
    const ws = wb.addWorksheet("Detalle");
    const nombre = r?.proyecto.nombre || "Proyecto";
    ws.addRow([`Caja del proyecto: ${nombre}`]); ws.getRow(1).font = { bold: true, size: 14 };
    ws.addRow([`Saldo inicial`, money(r?.proyecto.saldoInicial || 0)]);
    ws.addRow([`Ingresos`, money(r?.ingresos || 0)]);
    ws.addRow([`Egresos`, money(r?.egresos || 0)]);
    ws.addRow([`Saldo`, money(r?.saldo || 0)]); ws.getRow(5).font = { bold: true };
    ws.addRow([]);
    const head = ws.addRow(["Fecha", "Descripción", "Origen", "Tipo", "Valor"]); head.font = { bold: true };
    const items = (r?.items || []).slice().sort((a, b) => (a.direction === b.direction ? String(a.fecha).localeCompare(String(b.fecha)) : a.direction === "in" ? -1 : 1));
    for (const it of items) ws.addRow([it.fecha, it.descripcion, it.origen, it.direction === "in" ? "Ingreso" : "Egreso", money(it.valor) * (it.direction === "in" ? 1 : -1)]);
    ws.columns = [{ width: 12 }, { width: 50 }, { width: 10 }, { width: 10 }, { width: 16 }];
  } else {
    // Resumen total + comparación con bancos.
    const ws = wb.addWorksheet("Resumen");
    ws.addRow([`Caja por proyecto — resumen`]); ws.getRow(1).font = { bold: true, size: 14 };
    ws.addRow([`Generado`, new Date().toISOString().slice(0, 10)]);
    ws.addRow([]);
    const head = ws.addRow(["Proyecto", "Ingresos", "Egresos", "Saldo"]); head.font = { bold: true };
    for (const r of rows.slice().sort((a, b) => b.saldo - a.saldo)) ws.addRow([r.proyecto.nombre, money(r.ingresos), money(r.egresos), money(r.saldo)]);
    const totRow = ws.addRow(["TOTAL CAJA", "", "", money(totalCaja)]); totRow.font = { bold: true };
    ws.addRow([]);
    const h2 = ws.addRow(["Cuentas bancarias", "Saldo", "Fecha"]); h2.font = { bold: true };
    for (const b of bancos) ws.addRow([b.nombre, money(b.saldo), b.fechaSaldo]);
    ws.addRow(["TOTAL BANCOS", money(totalBancos)]).font = { bold: true };
    ws.addRow(["DIFERENCIA (banco − caja)", money(totalBancos - totalCaja)]).font = { bold: true };
    ws.columns = [{ width: 32 }, { width: 16 }, { width: 16 }, { width: 16 }];
  }
  const fname = proyectoId ? `caja_proyecto_${new Date().toISOString().slice(0, 10)}.xlsx` : `caja_resumen_${new Date().toISOString().slice(0, 10)}.xlsx`;
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename="${fname}"`);
  await wb.xlsx.write(res);
  res.end();
});

// Saldos actuales por proyecto (para la planeación de pagos).
router.get("/saldos", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const { rows, bancos, totalCaja, totalBancos } = await buildCajaData(companyId);
  res.json({ ok: true, data: {
    proyectos: rows.map((r) => ({ id: r.proyecto.id, nombre: r.proyecto.nombre, esTransversal: r.proyecto.esTransversal, saldo: r.saldo })),
    totalCaja, totalBancos,
  } });
});

// ── Planes de pago (simulación guardable) ──────────────────────────────
router.get("/planes", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  res.json({ ok: true, data: await listPlanes(companyId) });
});
router.post("/planes", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  res.json({ ok: true, data: await savePlan(companyId, req.body || {}) });
});
router.delete("/planes/:id", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  res.json({ ok: await deletePlan(companyId, req.params.id) });
});

// Reporte de caja en PDF profesional (Puppeteer renderiza HTML limpio).
router.get("/report/pdf", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const proyectoId = typeof req.query.proyectoId === "string" ? req.query.proyectoId : "";
  const { rows, bancos, totalCaja, totalBancos } = await buildCajaData(companyId);
  let empresa = "Empresa";
  try { const c = await getDb().collection<any>("siigoCompanies").findOne({ _id: new ObjectId(companyId) }, { projection: { name: 1 } }); if (c?.name) empresa = c.name; } catch { /* noop */ }
  try {
    const pdf = await generarCajaPdf({
      empresa,
      rows: rows.map((r) => ({ proyecto: { id: r.proyecto.id, nombre: r.proyecto.nombre, saldoInicial: r.proyecto.saldoInicial, esTransversal: r.proyecto.esTransversal }, ingresos: r.ingresos, egresos: r.egresos, saldo: r.saldo, items: r.items })),
      bancos: bancos.map((b) => ({ nombre: b.nombre, saldo: b.saldo, fechaSaldo: b.fechaSaldo })),
      totalCaja, totalBancos, proyectoId: proyectoId || undefined,
    });
    const fname = proyectoId ? `caja_proyecto_${new Date().toISOString().slice(0, 10)}.pdf` : `caja_resumen_${new Date().toISOString().slice(0, 10)}.pdf`;
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${fname}"`);
    res.send(pdf);
  } catch (e) {
    res.status(502).json({ ok: false, message: e instanceof Error ? e.message : "Error generando PDF." });
  }
});

// ── Traer facturas nuevas desde DIAN (token del buzón + rango por mes) ──────
// "La magia": toma el token reenviado al buzón de ContaGO para el NIT de la
// empresa activa y descarga las facturas RECIBIDAS del mes elegido, descartando
// las ya importadas (dedup por CUFE). No requiere pegar token ni subir archivos.
const onlyDigits = (s?: string) => String(s || "").replace(/\D/g, "").replace(/^0+/, "");

interface DianFetchJob {
  status: "processing" | "completed" | "cancelled" | "error";
  progress: { step: string; current: number; total: number };
  userId: string;
  createdAt: number;
  result?: IngestResult;
  error?: string;
}
const dianFetchJobs = new Map<string, DianFetchJob>();
const DIAN_FETCH_TTL_MS = 2 * 60 * 60 * 1000;
setInterval(() => {
  const now = Date.now();
  for (const [id, job] of dianFetchJobs) if (now - job.createdAt > DIAN_FETCH_TTL_MS) dianFetchJobs.delete(id);
}, 30 * 60 * 1000).unref?.();

// Rango [primer día, último día] de un mes (YYYY-MM-DD).
function monthRange(year: number, month: number): { fechaInicio: string; fechaFin: string } {
  const mm = String(month).padStart(2, "0");
  const lastDay = new Date(year, month, 0).getDate();
  return { fechaInicio: `${year}-${mm}-01`, fechaFin: `${year}-${mm}-${String(lastDay).padStart(2, "0")}` };
}

// Estado del token disponible para la empresa: ¿hay uno en el buzón? ¿hace cuánto?
router.get("/dian/token-status", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const ctx = await getCompanyContext(companyId);
  const nit = onlyDigits(ctx?.nit || "");
  if (!nit) return res.json({ ok: true, data: { mailbox: await getDianMailboxStatus(), companyNit: "", token: null } });
  let token = null;
  try {
    const t = await latestDianTokenForNit(nit, 7);
    if (t) token = { receivedAt: t.receivedAt, ageMinutes: Math.round((Date.now() - t.receivedAt) / 60000), subject: t.subject };
  } catch { /* buzón no conectado: token queda null */ }
  res.json({ ok: true, data: { mailbox: await getDianMailboxStatus(), companyNit: nit, token } });
});

// Dispara la descarga del mes. Body: { year, month }. Devuelve { jobId } o, si no
// hay token fresco en el buzón, { ok:false, code:"NO_TOKEN" } para que la UI guíe.
router.post("/dian/fetch-new", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const year = Number(req.body?.year);
  const month = Number(req.body?.month);
  if (!year || !month || month < 1 || month > 12) {
    return res.status(400).json({ ok: false, message: "Indica el año y el mes a consultar." });
  }
  const ctx = await getCompanyContext(companyId);
  if (!ctx) return res.status(400).json({ ok: false, message: "Empresa sin conexión a SIIGO." });
  const companyNit = onlyDigits(ctx.nit);
  if (!companyNit) {
    return res.status(400).json({ ok: false, code: "NO_COMPANY_NIT", message: "La empresa no tiene NIT configurado." });
  }

  let token;
  try {
    token = await latestDianTokenForNit(companyNit, 7);
  } catch (e) {
    const msg = e instanceof Error ? e.message : "";
    if (msg === "DIAN_MAILBOX_NOT_CONNECTED") {
      return res.status(409).json({ ok: false, code: "MAILBOX_OFF", message: "El buzón DIAN de ContaGO no está conectado. Avísale al administrador." });
    }
    return res.status(502).json({ ok: false, message: "No se pudo leer el buzón DIAN." });
  }
  if (!token) {
    return res.status(409).json({
      ok: false,
      code: "NO_TOKEN",
      message: `No hay un token reciente para el NIT ${companyNit}. Solicita un token en el portal DIAN y reenvíalo al buzón de ContaGO; luego vuelve a intentar.`,
    });
  }

  const { fechaInicio, fechaFin } = monthRange(year, month);
  const maxDocuments = Math.max(1, Number(process.env.DIAN_MAX_DOCUMENTS || 850));
  const jobId = uuidv4();
  dianFetchJobs.set(jobId, { status: "processing", progress: { step: "En cola...", current: 0, total: 0 }, userId: req.user?.userId || "", createdAt: Date.now() });

  runWithSiigoCompany(ctx, () =>
    ingestNewByDateRange({
      companyId,
      tokenUrl: token!.tokenUrl,
      fechaInicio,
      fechaFin,
      nitReceptor: companyNit,
      maxDocuments,
      onProgress: (p) => { const j = dianFetchJobs.get(jobId); if (j && j.status === "processing") j.progress = p; },
      isCancelled: () => dianFetchJobs.get(jobId)?.status === "cancelled",
    })
  )
    .then((result) => {
      const j = dianFetchJobs.get(jobId);
      if (!j || j.status === "cancelled") return;
      j.status = "completed"; j.result = result;
      j.progress = { step: "Completado", current: result.stats.downloaded, total: result.stats.listed };
    })
    .catch((err) => {
      const j = dianFetchJobs.get(jobId);
      if (!j) return;
      j.status = "error"; j.error = err instanceof Error ? err.message : "Error en la descarga DIAN.";
    });

  res.json({ ok: true, jobId, range: { fechaInicio, fechaFin }, tokenAgeMinutes: Math.round((Date.now() - token.receivedAt) / 60000) });
});

router.get("/dian/fetch-new/status/:jobId", (req, res) => {
  const job = dianFetchJobs.get(req.params.jobId);
  if (!job) return res.status(404).json({ ok: false, message: "Job no encontrado." });
  res.json({ ok: true, status: job.status, progress: job.progress, error: job.error, stats: job.result?.stats });
});

router.get("/dian/fetch-new/result/:jobId", (req, res) => {
  const job = dianFetchJobs.get(req.params.jobId);
  if (!job) return res.status(404).json({ ok: false, message: "Job no encontrado." });
  if (job.status !== "completed" || !job.result) return res.status(409).json({ ok: false, message: "El job aún no ha terminado.", status: job.status });
  res.json({ ok: true, data: job.result.items, stats: job.result.stats, failures: job.result.failures });
});

router.post("/dian/fetch-new/cancel/:jobId", (req, res) => {
  const job = dianFetchJobs.get(req.params.jobId);
  if (!job) return res.status(404).json({ ok: false, message: "Job no encontrado." });
  if (job.status === "processing") job.status = "cancelled";
  res.json({ ok: true });
});

export default router;
