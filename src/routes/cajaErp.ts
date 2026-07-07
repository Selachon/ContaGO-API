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
import { userCanAccessCompany, userOwnsCompany, setCompanyDrive, getCompanyDrive, getCompanyContext } from "../services/siigoCompaniesService.js";
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
  discardInvoice,
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
  listDrafts,
  syncDrafts,
  clearDrafts,
  listGastosFijos,
  createGastoFijo,
  updateGastoFijo,
  deleteGastoFijo,
  initGastosFijosDefaults,
  listPlanProyectos,
  upsertPlanProyecto,
  computeProyeccion,
  getGastosOverrides,
  setGastoOverride,
  listPlanIngresos,
  upsertPlanIngreso,
  deletePlanIngreso,
  computeFlujoEditable,
  getCajaConfig,
  updateCajaConfig,
  listProvDinamica,
  upsertProvDinamica,
  type TagKind,
  getPublicToken,
  generatePublicToken,
  revokePublicToken,
  resolveCompanyByToken,
} from "../services/cajaErpStore.js";
import ExcelJS from "exceljs";
import { listMovements, getMovement, updateMovement as updateBankMov } from "../services/siigoEgresosStoreService.js";
import { generarCajaPdf } from "../services/cajaPdf.js";

export const CAJA_ERP_TOOL_ID = "causacion-caja";

// ── Router público (sin auth) — solo lectura vía token ────────────────────────
export const publicCajaRouter = Router();

publicCajaRouter.get("/public/:token/report", async (req, res) => {
  const companyId = await resolveCompanyByToken(req.params.token);
  if (!companyId) return res.status(404).json({ ok: false, message: "Token inválido." });
  const { rows, bancos, totalCaja, totalBancos } = await buildCajaData(companyId);
  const round = (n: number) => Math.round(Number(n) || 0);
  const totalSaldoInicial = rows.reduce((a, r) => a + (r.proyecto.saldoInicial || 0), 0);
  const totalIngresos = rows.reduce((a, r) => a + r.ingresos, 0);
  const totalEgresos = rows.reduce((a, r) => a + r.egresos, 0);
  res.json({ ok: true, data: {
    generadoEn: new Date().toISOString(),
    kpis: {
      totalSaldoInicial: round(totalSaldoInicial),
      totalIngresos: round(totalIngresos),
      totalEgresos: round(totalEgresos),
      totalCaja: round(totalCaja),
      totalBancos: round(totalBancos),
    },
    proyectos: rows
      .filter((r) => r.saldo !== 0 || r.items.length > 0)
      .map((r) => ({
        id: r.proyecto.id,
        nombre: r.proyecto.nombre,
        esTransversal: r.proyecto.esTransversal,
        saldoInicial: round(r.proyecto.saldoInicial || 0),
        ingresos: round(r.ingresos),
        egresos: round(r.egresos),
        saldo: round(r.saldo),
        items: (r.items || []).map((it: any) => ({
          fecha: it.fecha,
          descripcion: it.descripcion,
          valor: round(it.valor),
          direction: it.direction,
          origen: it.origen || it.source || "",
          categoria: it.categoria || "",
        })),
      })),
    bancos,
  }});
});

publicCajaRouter.get("/public/:token/flujo", async (req, res) => {
  const companyId = await resolveCompanyByToken(req.params.token);
  if (!companyId) return res.status(404).json({ ok: false, message: "Token inválido." });
  const meses = Math.min(12, Math.max(1, Number(req.query.meses) || 6));
  const [flujo, config] = await Promise.all([
    computeFlujoEditable(companyId, meses),
    getCajaConfig(companyId),
  ]);
  res.json({ ok: true, data: { flujo, config } });
});

publicCajaRouter.get("/public/:token/excel", async (req, res) => {
  const companyId = await resolveCompanyByToken(req.params.token);
  if (!companyId) return res.status(404).json({ ok: false, message: "Token inválido." });
  const { rows, bancos: _bancos, totalCaja, totalBancos: _tb } = await buildCajaData(companyId);
  const round = (n: number) => Math.round(Number(n) || 0);
  const wb = new ExcelJS.Workbook();
  wb.creator = "ContaGO"; wb.created = new Date();
  const BLUE = "1E3A5F"; const WHITE = "FFFFFF";
  const HEAD_FILL = { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FF" + BLUE } };
  const ALT_FILL  = { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFF5F5F5" } };

  const { proyectoId } = req.query as { proyectoId?: string };

  function addMovSheet(ws: ExcelJS.Worksheet, items: any[], proyNombre: string) {
    const hRow = ws.addRow(["Fecha", "Mes", "Proyecto", "Descripción", "Categoría", "Valor"]);
    hRow.eachCell((c) => { c.fill = HEAD_FILL; c.font = { bold: true, color: { argb: "FF" + WHITE } }; });
    const sorted = [...items].sort((a, b) => String(a.fecha).localeCompare(String(b.fecha)));
    let ri = 2;
    for (const it of sorted) {
      const val = round(it.valor) * (it.direction === "in" ? 1 : -1);
      const dRow = ws.addRow([it.fecha, String(it.fecha || "").slice(0, 7), proyNombre, it.descripcion, it.categoria || it.origen || "", val]);
      dRow.getCell(6).numFmt = "#,##0";
      dRow.getCell(6).font = { color: { argb: val >= 0 ? "FF1B5E20" : "FFB71C1C" } };
      if (ri % 2 === 0) dRow.eachCell((c) => { c.fill = ALT_FILL; });
      ri++;
    }
    ws.columns = [{ width: 12 }, { width: 10 }, { width: 28 }, { width: 50 }, { width: 26 }, { width: 18 }];
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: 6 } };
    ws.views = [{ state: "frozen", ySplit: 1 }];
  }

  let fname: string;

  if (proyectoId) {
    // Descarga de un solo proyecto
    const row = rows.find((r) => r.proyecto.id === proyectoId);
    if (!row) return res.status(404).json({ ok: false, message: "Proyecto no encontrado." });
    const ws = wb.addWorksheet(row.proyecto.nombre.slice(0, 31));
    addMovSheet(ws, row.items || [], row.proyecto.nombre);
    fname = `caja_${row.proyecto.nombre.replace(/\s+/g, "_").slice(0, 30)}_${new Date().toISOString().slice(0, 10)}.xlsx`;
  } else {
    // Descarga consolidada (todos)
    const wsTodos = wb.addWorksheet("Todos los movimientos");
    const allItems: any[] = [];
    for (const r of rows) for (const it of r.items || []) allItems.push({ ...it, proyNombre: r.proyecto.nombre });
    addMovSheet(wsTodos, allItems.map((it) => ({ ...it, descripcion: it.descripcion })), "");
    // Overwrite column 3 with proyNombre per row already included via the addMovSheet loop
    // Hoja resumen
    const wsRes = wb.addWorksheet("Resumen");
    const rh = wsRes.addRow(["Proyecto", "Saldo inicial", "Ingresos", "Egresos", "Saldo"]);
    rh.eachCell((c) => { c.fill = HEAD_FILL; c.font = { bold: true, color: { argb: "FF" + WHITE } }; });
    for (const r of rows.slice().sort((a, b) => b.saldo - a.saldo))
      wsRes.addRow([r.proyecto.nombre, round(r.proyecto.saldoInicial || 0), round(r.ingresos), round(r.egresos), round(r.saldo)]);
    wsRes.addRow(["TOTAL CAJA", "", "", "", round(totalCaja)]).font = { bold: true };
    wsRes.columns = [{ width: 32 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 16 }];
    fname = `caja_${new Date().toISOString().slice(0, 10)}.xlsx`;
  }

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename="${fname}"`);
  await wb.xlsx.write(res); res.end();
});

// Router principal: requiere tool causacion-caja (para Caja ERP).
const router = Router();
router.use(requireAuth);
router.use(requireToolAccess(CAJA_ERP_TOOL_ID));

// Router de bandeja (drafts): solo requiere auth + acceso a la empresa.
// Lo usan tanto CausacionCaja (tool=caja) como SiigoXmlAccounting (tool=xml)
// que tienen tools distintos, por eso va separado del router principal.
export const draftsRouter = Router();
draftsRouter.use(requireAuth);

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

// Igual que resolveCompany pero permite a admins de ContaGO gestionar la bandeja
// de borradores de cualquier empresa cliente (solo lectura/escritura de drafts,
// no da acceso a las rutas de causación ni a Siigo).
async function resolveCompanyForDrafts(req: Request, res: Response): Promise<string | null> {
  const companyId = req.header("X-Siigo-Company");
  if (!companyId) {
    res.status(400).json({ ok: false, message: "Falta la empresa (X-Siigo-Company)." });
    return null;
  }
  const userId = req.user?.userId;
  const isAdmin = req.user?.isAdmin === true;
  const allowed = isAdmin || (userId ? await userCanAccessCompany(companyId, userId) : false);
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
  const month = typeof req.query.month === "string" ? req.query.month : undefined;
  res.json({ ok: true, data: await listEntries(companyId, proyectoId, month) });
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

// ── Bandeja de PENDIENTES (drafts) compartida por empresa ────────────────
// Facturas importadas aún NO causadas. Persisten en el servidor para sobrevivir
// al cambio de PC y para que las vean todos los usuarios de la empresa compartida.
// Registradas en draftsRouter (sin requireToolAccess) para que funcionen tanto con
// tool=caja (CausacionCaja) como tool=xml (SiigoXmlAccounting).
draftsRouter.get("/drafts", async (req, res) => {
  const companyId = await resolveCompanyForDrafts(req, res);
  if (!companyId) return;
  const tool = String(req.query.tool || "caja");
  res.json({ ok: true, data: await listDrafts(companyId, tool) });
});

draftsRouter.post("/drafts/sync", async (req, res) => {
  const companyId = await resolveCompanyForDrafts(req, res);
  if (!companyId) return;
  const tool = String(req.body?.tool || "caja");
  const upserts = Array.isArray(req.body?.upserts) ? req.body.upserts : [];
  const deletes = Array.isArray(req.body?.deletes) ? req.body.deletes : [];
  try {
    const r = await syncDrafts(companyId, upserts, deletes, req.user?.userId || "", tool);
    res.json({ ok: true, data: r });
  } catch (err) {
    res.status(500).json({ ok: false, message: err instanceof Error ? err.message : "Error al sincronizar pendientes." });
  }
});

draftsRouter.delete("/drafts", async (req, res) => {
  const companyId = await resolveCompanyForDrafts(req, res);
  if (!companyId) return;
  const tool = String(req.query.tool || "caja");
  const deleted = await clearDrafts(companyId, tool);
  res.json({ ok: true, data: { deleted } });
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

// Descarta una factura (no se causará): rechazada / otro motivo. Se registra en el
// buzón con estado "descartada" para mostrarla en su filtro y para que su clave
// bloquee la re-ingesta desde DIAN. (La "causada manualmente" usa /invoices/save.)
router.post("/invoices/discard", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const rec = req.body?.invoice || {};
  const motivo = String(req.body?.motivo || "otro");
  const nota = String(req.body?.nota || "");
  if (!rec.key && !rec.cufe && !(rec.supplierNit && rec.docNumber)) {
    return res.status(400).json({ ok: false, message: "Falta identificar la factura (key/cufe)." });
  }
  await discardInvoice(companyId, rec, motivo, nota);
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

// Sube soporte de pago de una factura individual (pagada por otro medio / descartada)
// Guarda la URL en el campo soportePagoUrl de la factura.
router.post("/invoices/support", upload.single("file"), async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const file = req.file;
  const { key, date } = req.body || {};
  if (!file || !key) return res.status(400).json({ ok: false, message: "Falta archivo o key." });
  const drive = await resolveCompanyDrive(companyId);
  if (!drive) return res.status(409).json({ ok: false, message: "La empresa no tiene Drive vinculado." });
  try {
    const url = await uploadPaymentSupportToDrive(
      file.buffer, file.originalname, file.mimetype,
      String(drive.nit || "").replace(/\D/g, ""), String(date || ""),
      drive.driveConfig, drive.ownerUserId, drive.onTokenRefresh,
    );
    await getDb().collection<any>("cajaErpInvoices").updateOne(
      { companyId, key: String(key) },
      { $set: { paymentPdfUrl: url, updatedAt: new Date().toISOString() } },
    );
    res.json({ ok: true, data: { url } });
  } catch (e) {
    res.status(502).json({ ok: false, message: e instanceof Error ? e.message : "Error subiendo el soporte." });
  }
});

// ── Google Drive de la empresa (para archivar las facturas) ────────────
router.get("/drive/status", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const link = await getCompanyDrive(companyId);
  const canManage = req.user!.isAdmin || (await userOwnsCompany(companyId, req.user!.userId));
  const available = canManage
    ? (await getUserGoogleDrives(req.user!.userId)).map((d) => ({ connectionId: d.connection_id, email: d.user_email }))
    : [];
  let linkedEmail = "";
  if (link) {
    const cfg = await getUserGoogleDriveById(link.ownerUserId, link.connectionId);
    linkedEmail = cfg?.user_email || "";
  }
  res.json({ ok: true, data: { linked: !!link, linkedEmail, connectionId: link?.connectionId || "", available, isOwner: canManage } });
});

router.post("/drive/link", async (req, res) => {
  const companyId = await resolveCompany(req, res);
  if (!companyId) return;
  const canManage = req.user!.isAdmin || (await userOwnsCompany(companyId, req.user!.userId));
  if (!canManage) {
    return res.status(403).json({ ok: false, message: "Solo el administrador de la empresa puede vincular su cuenta de Drive." });
  }
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
    if (m.cajaEntryId) continue; // ya representado como entry manual — evitar doble conteo
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
  const allProjects = req.query.all === "true";
  const { rows, bancos, totalCaja, totalBancos } = await buildCajaData(companyId);
  const money = (n: number) => Math.round(Number(n) || 0);
  const wb = new ExcelJS.Workbook();
  wb.creator = "ContaGO"; wb.created = new Date();

  const BLUE = "1E3A5F"; const WHITE = "FFFFFF";
  const GREEN_FILL = { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFE8F5E9" } };
  const RED_FILL   = { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFFFEBEE" } };
  const HEAD_FILL  = { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FF" + BLUE } };
  const ALT_FILL   = { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFF5F5F5" } };

  const addDetailSheet = (ws: ExcelJS.Worksheet, r: typeof rows[0]) => {
    const nombre = r.proyecto.nombre;
    // Encabezado
    ws.mergeCells("A1:F1");
    const t = ws.getCell("A1"); t.value = `Caja · ${nombre}`;
    t.font = { bold: true, size: 14, color: { argb: "FF" + BLUE } };
    t.alignment = { vertical: "middle" };
    ws.getRow(1).height = 26;

    ws.addRow(["Saldo inicial", "", "", "", "", money(r.proyecto.saldoInicial || 0)]);
    ws.addRow(["Ingresos",      "", "", "", "", money(r.ingresos)]);
    ws.addRow(["Egresos",       "", "", "", "", money(r.egresos)]);
    const sRow = ws.addRow(["Saldo final", "", "", "", "", money(r.saldo)]);
    sRow.font = { bold: true }; ws.addRow([]);

    const hRow = ws.addRow(["Fecha", "Mes", "Proyecto", "Descripción", "Categoría", "Valor"]);
    hRow.eachCell((c) => { c.fill = HEAD_FILL; c.font = { bold: true, color: { argb: "FF" + WHITE } }; c.alignment = { horizontal: "center" }; });
    ws.getRow(hRow.number).height = 20;

    const items = (r.items || []).slice().sort((a: any, b: any) => String(a.fecha).localeCompare(String(b.fecha)));
    let rowIdx = hRow.number + 1;
    for (const it of items) {
      const mes = String(it.fecha || "").slice(0, 7);
      const val = money(it.valor) * (it.direction === "in" ? 1 : -1);
      const dRow = ws.addRow([it.fecha, mes, nombre, it.descripcion, it.categoria || it.origen || "", val]);
      const fill = (rowIdx % 2 === 0) ? ALT_FILL : undefined;
      dRow.getCell(6).numFmt = "#,##0";
      dRow.getCell(6).font = { color: { argb: val >= 0 ? "FF1B5E20" : "FFB71C1C" } };
      if (fill) dRow.eachCell((c) => { c.fill = fill; });
      rowIdx++;
    }
    ws.columns = [{ width: 12 }, { width: 10 }, { width: 28 }, { width: 50 }, { width: 26 }, { width: 18 }];
    ws.autoFilter = { from: { row: hRow.number, column: 1 }, to: { row: hRow.number, column: 6 } };
  };

  if (allProjects) {
    // ── Hoja "Todos" filtrable ─────────────────────────────────────────────
    const wsTodos = wb.addWorksheet("Todos los movimientos");
    const hRow = wsTodos.addRow(["Fecha", "Mes", "Proyecto", "Descripción", "Categoría", "Valor"]);
    hRow.eachCell((c) => { c.fill = HEAD_FILL; c.font = { bold: true, color: { argb: "FF" + WHITE } }; c.alignment = { horizontal: "center" }; });
    wsTodos.getRow(1).height = 22;

    const allItems: any[] = [];
    for (const r of rows) {
      for (const it of r.items || []) allItems.push({ ...it, proyNombre: r.proyecto.nombre });
    }
    allItems.sort((a, b) => String(a.fecha).localeCompare(String(b.fecha)));

    let ri = 2;
    for (const it of allItems) {
      const mes = String(it.fecha || "").slice(0, 7);
      const val = money(it.valor) * (it.direction === "in" ? 1 : -1);
      const dRow = wsTodos.addRow([it.fecha, mes, it.proyNombre, it.descripcion, it.categoria || it.origen || "", val]);
      dRow.getCell(6).numFmt = "#,##0";
      dRow.getCell(6).font = { color: { argb: val >= 0 ? "FF1B5E20" : "FFB71C1C" } };
      if (ri % 2 === 0) dRow.eachCell((c) => { c.fill = ALT_FILL; });
      ri++;
    }
    wsTodos.columns = [{ width: 12 }, { width: 10 }, { width: 28 }, { width: 50 }, { width: 26 }, { width: 18 }];
    wsTodos.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: 6 } };
    wsTodos.views = [{ state: "frozen", ySplit: 1 }];

    // ── Una hoja por proyecto ──────────────────────────────────────────────
    for (const r of rows) {
      if (!r.items?.length && !r.proyecto.saldoInicial) continue;
      const safeName = r.proyecto.nombre.replace(/[\\\/\?\*\[\]]/g, "").slice(0, 31);
      const ws = wb.addWorksheet(safeName);
      addDetailSheet(ws, r);
    }

    // ── Hoja resumen ───────────────────────────────────────────────────────
    const wsRes = wb.addWorksheet("Resumen");
    wsRes.mergeCells("A1:D1");
    const rt = wsRes.getCell("A1"); rt.value = "Caja por proyecto — resumen";
    rt.font = { bold: true, size: 14, color: { argb: "FF" + BLUE } }; wsRes.getRow(1).height = 26;
    wsRes.addRow(["Generado", new Date().toISOString().slice(0, 10)]);
    wsRes.addRow([]);
    const rh = wsRes.addRow(["Proyecto", "Saldo inicial", "Ingresos", "Egresos", "Saldo"]);
    rh.eachCell((c) => { c.fill = HEAD_FILL; c.font = { bold: true, color: { argb: "FF" + WHITE } }; });
    for (const r of rows.slice().sort((a, b) => b.saldo - a.saldo)) {
      const row = wsRes.addRow([r.proyecto.nombre, money(r.proyecto.saldoInicial || 0), money(r.ingresos), money(r.egresos), money(r.saldo)]);
      row.getCell(5).font = { bold: true, color: { argb: money(r.saldo) >= 0 ? "FF1B5E20" : "FFB71C1C" } };
    }
    const totRow = wsRes.addRow(["TOTAL CAJA", "", "", "", money(totalCaja)]);
    totRow.font = { bold: true }; totRow.getCell(5).numFmt = "#,##0";
    wsRes.addRow([]);
    const bh = wsRes.addRow(["Cuenta bancaria", "Saldo banco", "Fecha"]); bh.font = { bold: true };
    for (const b of bancos) wsRes.addRow([b.nombre, money(b.saldo), b.fechaSaldo]);
    wsRes.addRow(["TOTAL BANCOS", money(totalBancos)]).font = { bold: true };
    wsRes.addRow(["Diferencia banco − caja", money(totalBancos - totalCaja)]).font = { bold: true };
    wsRes.columns = [{ width: 32 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 16 }];

  } else if (proyectoId) {
    const r = rows.find((x) => x.proyecto.id === proyectoId);
    if (r) addDetailSheet(wb.addWorksheet("Detalle"), r);
  } else {
    // Resumen simple
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

  const fname = allProjects
    ? `caja_todos_proyectos_${new Date().toISOString().slice(0, 10)}.xlsx`
    : proyectoId ? `caja_proyecto_${new Date().toISOString().slice(0, 10)}.xlsx`
    : `caja_resumen_${new Date().toISOString().slice(0, 10)}.xlsx`;
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

// ─── PLAN FINANCIERO ──────────────────────────────────────────────────────────

router.get("/plan/gastos-fijos", async (req: Request, res: Response) => {
  const companyId = req.headers["x-siigo-company"] as string;
  if (!companyId || !(await userCanAccessCompany(companyId, req.user!.userId)))
    return res.status(403).json({ ok: false, message: "Sin acceso." });
  await initGastosFijosDefaults(companyId);
  const items = await listGastosFijos(companyId);
  res.json({ ok: true, items });
});

router.post("/plan/gastos-fijos", async (req: Request, res: Response) => {
  const companyId = req.headers["x-siigo-company"] as string;
  if (!companyId || !(await userCanAccessCompany(companyId, req.user!.userId)))
    return res.status(403).json({ ok: false, message: "Sin acceso." });
  const { nombre, monto, frecuenciaMeses, activo } = req.body;
  if (!nombre || !monto) return res.status(400).json({ ok: false, message: "nombre y monto requeridos." });
  const item = await createGastoFijo(companyId, {
    nombre,
    monto: Number(monto),
    frecuenciaMeses: Number(frecuenciaMeses) || 1,
    activo: activo !== false,
  });
  res.json({ ok: true, item });
});

router.patch("/plan/gastos-fijos/:id", async (req: Request, res: Response) => {
  const companyId = req.headers["x-siigo-company"] as string;
  if (!companyId || !(await userCanAccessCompany(companyId, req.user!.userId)))
    return res.status(403).json({ ok: false, message: "Sin acceso." });
  const ok = await updateGastoFijo(companyId, req.params.id, req.body);
  res.json({ ok });
});

router.delete("/plan/gastos-fijos/:id", async (req: Request, res: Response) => {
  const companyId = req.headers["x-siigo-company"] as string;
  if (!companyId || !(await userCanAccessCompany(companyId, req.user!.userId)))
    return res.status(403).json({ ok: false, message: "Sin acceso." });
  const ok = await deleteGastoFijo(companyId, req.params.id);
  res.json({ ok });
});

router.get("/plan/proyectos-plan", async (req: Request, res: Response) => {
  const companyId = req.headers["x-siigo-company"] as string;
  if (!companyId || !(await userCanAccessCompany(companyId, req.user!.userId)))
    return res.status(403).json({ ok: false, message: "Sin acceso." });
  const items = await listPlanProyectos(companyId);
  res.json({ ok: true, items });
});

router.patch("/plan/proyectos-plan/:proyectoId", async (req: Request, res: Response) => {
  const companyId = req.headers["x-siigo-company"] as string;
  if (!companyId || !(await userCanAccessCompany(companyId, req.user!.userId)))
    return res.status(403).json({ ok: false, message: "Sin acceso." });
  const { proyectoId } = req.params;
  const { valorContrato, fechaInicio, fechaFin, aporteOficinaMensual, honorariosArquitectasMensual } = req.body;
  const item = await upsertPlanProyecto(companyId, proyectoId, {
    valorContrato: Number(valorContrato) || 0,
    fechaInicio: fechaInicio || "",
    fechaFin: fechaFin || "",
    aporteOficinaMensual: Number(aporteOficinaMensual) || 0,
    honorariosArquitectasMensual: Number(honorariosArquitectasMensual) || 0,
  });
  res.json({ ok: true, item });
});

router.get("/plan/proyeccion", async (req: Request, res: Response) => {
  const companyId = req.headers["x-siigo-company"] as string;
  if (!companyId || !(await userCanAccessCompany(companyId, req.user!.userId)))
    return res.status(403).json({ ok: false, message: "Sin acceso." });
  const meses = Math.min(12, Math.max(1, Number(req.query.meses) || 6));
  const data = await computeProyeccion(companyId, meses);
  res.json({ ok: true, data });
});

// ── Flujo de caja editable ────────────────────────────────────────────────────

router.get("/plan/flujo", async (req: Request, res: Response) => {
  const companyId = req.headers["x-siigo-company"] as string;
  if (!companyId || !(await userCanAccessCompany(companyId, req.user!.userId)))
    return res.status(403).json({ ok: false, message: "Sin acceso." });
  const meses = Math.min(12, Math.max(1, Number(req.query.meses) || 6));
  const data = await computeFlujoEditable(companyId, meses);
  res.json({ ok: true, data });
});

router.patch("/plan/gastos-mes", async (req: Request, res: Response) => {
  const companyId = req.headers["x-siigo-company"] as string;
  if (!companyId || !(await userCanAccessCompany(companyId, req.user!.userId)))
    return res.status(403).json({ ok: false, message: "Sin acceso." });
  const { mes, gastoFijoId, monto } = req.body;
  if (!mes || !gastoFijoId) return res.status(400).json({ ok: false, message: "mes y gastoFijoId requeridos." });
  await setGastoOverride(companyId, mes, gastoFijoId, monto === null || monto === undefined ? null : Number(monto));
  res.json({ ok: true });
});

router.get("/plan/ingresos", async (req: Request, res: Response) => {
  const companyId = req.headers["x-siigo-company"] as string;
  if (!companyId || !(await userCanAccessCompany(companyId, req.user!.userId)))
    return res.status(403).json({ ok: false, message: "Sin acceso." });
  const items = await listPlanIngresos(companyId);
  res.json({ ok: true, items });
});

router.post("/plan/ingresos", async (req: Request, res: Response) => {
  const companyId = req.headers["x-siigo-company"] as string;
  if (!companyId || !(await userCanAccessCompany(companyId, req.user!.userId)))
    return res.status(403).json({ ok: false, message: "Sin acceso." });
  const { concepto, tipo, pagos, notaDescarte, montoEstimado } = req.body;
  if (!concepto) return res.status(400).json({ ok: false, message: "concepto requerido." });
  const tipoVal = tipo === "probable" ? "probable" : tipo === "descartado" ? "descartado" : "confirmado";
  const item = await upsertPlanIngreso(companyId, { concepto, tipo: tipoVal, pagos: pagos || [], notaDescarte, montoEstimado: montoEstimado ? Number(montoEstimado) : undefined });
  res.json({ ok: true, item });
});

router.patch("/plan/ingresos/:id", async (req: Request, res: Response) => {
  const companyId = req.headers["x-siigo-company"] as string;
  if (!companyId || !(await userCanAccessCompany(companyId, req.user!.userId)))
    return res.status(403).json({ ok: false, message: "Sin acceso." });
  const { concepto, tipo, pagos, notaDescarte, montoEstimado } = req.body;
  const tipoVal = tipo === "probable" ? "probable" : tipo === "descartado" ? "descartado" : "confirmado";
  const item = await upsertPlanIngreso(companyId, { id: req.params.id, concepto, tipo: tipoVal, pagos: pagos || [], notaDescarte, montoEstimado: montoEstimado ? Number(montoEstimado) : undefined });
  res.json({ ok: true, item });
});

router.delete("/plan/ingresos/:id", async (req: Request, res: Response) => {
  const companyId = req.headers["x-siigo-company"] as string;
  if (!companyId || !(await userCanAccessCompany(companyId, req.user!.userId)))
    return res.status(403).json({ ok: false, message: "Sin acceso." });
  const ok = await deletePlanIngreso(companyId, req.params.id);
  res.json({ ok });
});

// ── Ingreso bancario → entrada en Caja por proyecto ──────────────────────────
// Crea una cajaErpEntry (direction:"in") a partir de un movimiento del extracto
// bancario y marca ese movimiento como "enviado_caja".

router.post("/from-mov/:movId", async (req: Request, res: Response) => {
  const companyId = req.headers["x-siigo-company"] as string;
  if (!companyId || !(await userCanAccessCompany(companyId, req.user!.userId)))
    return res.status(403).json({ ok: false, message: "Sin acceso." });

  const { proyectoId, rcRef, categoria } = req.body || {};
  if (!proyectoId) return res.status(400).json({ ok: false, message: "proyectoId requerido." });

  const mov = await getMovement(companyId, req.params.movId);
  if (!mov) return res.status(404).json({ ok: false, message: "Movimiento no encontrado." });
  if (mov.kind !== "ingreso") return res.status(400).json({ ok: false, message: "Solo ingresos se pueden enviar a caja." });
  if (mov.cajaEntryId) return res.status(409).json({ ok: false, message: "Este ingreso ya fue enviado a caja.", cajaEntryId: mov.cajaEntryId });

  // Descripción: incluye referencia RC si la hay
  const rcPart = rcRef ? ` [${rcRef}]` : (mov.receipt?.name ? ` [${mov.receipt.name}]` : "");
  const descripcion = `${mov.description || "Ingreso banco"}${rcPart}`.trim();

  const entry = await createEntry(companyId, {
    proyectoId,
    fecha: mov.date || new Date().toISOString().slice(0, 10),
    descripcion,
    valor: mov.value,
    direction: "in",
    categoria: categoria || rcRef || mov.receipt?.name || "Ingreso banco",
  });

  // Marcar el movimiento bancario como procesado
  await updateBankMov(companyId, req.params.movId, {
    cajaEntryId: entry.id,
    status: "enviado_caja",
  });

  res.json({ ok: true, entry });
});

// ── Configuración del flujo (saldo inicial, tarifa ICA) ───────────────────────

router.get("/plan/config", async (req: Request, res: Response) => {
  const companyId = req.headers["x-siigo-company"] as string;
  if (!companyId || !(await userCanAccessCompany(companyId, req.user!.userId)))
    return res.status(403).json({ ok: false, message: "Sin acceso." });
  const config = await getCajaConfig(companyId);
  res.json({ ok: true, config });
});

router.patch("/plan/config", async (req: Request, res: Response) => {
  const companyId = req.headers["x-siigo-company"] as string;
  if (!companyId || !(await userCanAccessCompany(companyId, req.user!.userId)))
    return res.status(403).json({ ok: false, message: "Sin acceso." });
  const { saldoInicial, tarifaIca, saldoBancos, adeudadoProyectos } = req.body;
  const config = await updateCajaConfig(companyId, {
    ...(saldoInicial        !== undefined ? { saldoInicial:        Number(saldoInicial) }        : {}),
    ...(tarifaIca           !== undefined ? { tarifaIca:           Number(tarifaIca) }            : {}),
    ...(saldoBancos         !== undefined ? { saldoBancos:         Number(saldoBancos) }          : {}),
    ...(adeudadoProyectos   !== undefined ? { adeudadoProyectos }                                 : {}),
  });
  res.json({ ok: true, config });
});

// ── Provisiones dinámicas (IVA / ICA por mes) ────────────────────────────────

router.get("/plan/prov-dinamica", async (req: Request, res: Response) => {
  const companyId = req.headers["x-siigo-company"] as string;
  if (!companyId || !(await userCanAccessCompany(companyId, req.user!.userId)))
    return res.status(403).json({ ok: false, message: "Sin acceso." });
  const mesesParam = String(req.query.meses || "");
  const meses = mesesParam ? mesesParam.split(",").map((m) => m.trim()).filter(Boolean) : [];
  if (!meses.length) return res.status(400).json({ ok: false, message: "Param meses requerido (YYYY-MM,…)." });
  const items = await listProvDinamica(companyId, meses);
  res.json({ ok: true, items });
});

router.patch("/plan/prov-dinamica", async (req: Request, res: Response) => {
  const companyId = req.headers["x-siigo-company"] as string;
  if (!companyId || !(await userCanAccessCompany(companyId, req.user!.userId)))
    return res.status(403).json({ ok: false, message: "Sin acceso." });
  const { mes, ivaGenerado, ivaDescontable, ingresosFact } = req.body;
  if (!mes) return res.status(400).json({ ok: false, message: "mes requerido (YYYY-MM)." });
  await upsertProvDinamica(companyId, mes, {
    ...(ivaGenerado    !== undefined ? { ivaGenerado:    Number(ivaGenerado)    } : {}),
    ...(ivaDescontable !== undefined ? { ivaDescontable: Number(ivaDescontable) } : {}),
    ...(ingresosFact   !== undefined ? { ingresosFact:   Number(ingresosFact)   } : {}),
  });
  res.json({ ok: true });
});

// ── Gestión del token público (autenticado) ───────────────────────────────────
router.get("/public-token", async (req, res) => {
  const companyId = await resolveCompany(req, res); if (!companyId) return;
  const token = await getPublicToken(companyId);
  res.json({ ok: true, data: { token } });
});
router.post("/public-token", async (req, res) => {
  const companyId = await resolveCompany(req, res); if (!companyId) return;
  const token = await generatePublicToken(companyId);
  res.json({ ok: true, data: { token } });
});
router.delete("/public-token", async (req, res) => {
  const companyId = await resolveCompany(req, res); if (!companyId) return;
  await revokePublicToken(companyId);
  res.json({ ok: true });
});

export default router;
