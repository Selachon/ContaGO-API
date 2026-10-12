import { Router, Request, Response, NextFunction } from "express";
import multer from "multer";
import { requireAuth } from "../middleware/auth.js";
import { requireToolAccess } from "../middleware/requireToolAccess.js";
import { getDianPlaces } from "../services/alanubeService.js";
import {
  attachBillPdf, attachExtra, deleteExtra, deleteSupportDoc, emitSupportDoc, findPreviousSupplier, getBillPdf, getExtra, getEmittedFile, getSupportDoc, listSupportDocs,
  getPublicLinkToken, loadSupportResolution, lookupSupplier, rotatePublicLinkToken, saveSupportDoc, supportTotals, SUPPLIER_ID_TYPES,
} from "../services/supportDocsService.js";

const router = Router();
const TOOL_ID = "documentos-soporte";
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024, files: 1 } });

router.use(requireAuth);
router.use(requireToolAccess(TOOL_ID));

const actor = (req: Request) => ({ userId: req.user!.userId, isAdmin: !!req.user!.isAdmin });
const withTotals = (d: any) => ({ ...d, id: String(d._id), _id: undefined, totals: d.input ? supportTotals(d.input) : undefined });
const fail = (res: Response, err: any, status = 400) => res.status(status).json({ ok: false, message: err?.message || "Error" });

/** Envuelve un middleware de multer para que sus errores (demasiados archivos, archivo muy pesado…) salgan como JSON y nunca como una página de error. */
const safeUpload = (mw: (req: Request, res: Response, cb: (err?: any) => void) => void) => (req: Request, res: Response, next: NextFunction) => {
  mw(req, res, (err?: any) => {
    if (!err) return next();
    const msg = err.code === "LIMIT_FILE_SIZE" ? "Un archivo pesa más de 10 MB."
      : err.code === "LIMIT_FILE_COUNT" || err.code === "LIMIT_UNEXPECTED_FILE" ? `Máximo ${5} documentos adicionales y un PDF de cuenta de cobro.`
      : "No se pudieron leer los archivos adjuntos.";
    res.status(400).json({ ok: false, message: msg });
  });
};


// Enlace público para proveedores (solo administrador): ver el token y renovarlo (el anterior deja de funcionar)
router.get("/enlace-publico", async (req, res) => {
  if (!req.user!.isAdmin) return res.status(403).json({ ok: false, message: "Solo el administrador gestiona el enlace público." });
  try { res.json({ ok: true, token: await getPublicLinkToken() }); } catch (err) { fail(res, err, 500); }
});
router.post("/enlace-publico/renovar", async (req, res) => {
  if (!req.user!.isAdmin) return res.status(403).json({ ok: false, message: "Solo el administrador gestiona el enlace público." });
  try { res.json({ ok: true, token: await rotatePublicLinkToken() }); } catch (err) { fail(res, err, 500); }
});

// Catálogos para el formulario (departamentos, municipios, tipos de documento) y la resolución en uso
router.get("/catalogs", async (_req, res) => {
  try {
    const places = await getDianPlaces();
    const r = loadSupportResolution();
    res.json({ ok: true, ...places, idTypes: SUPPLIER_ID_TYPES, resolution: { prefix: r.prefix, minNumber: r.minNumber, maxNumber: r.maxNumber } });
  } catch (err) { fail(res, err, 502); }
});

// Nombre y correo registrados en la DIAN para un documento (ayuda para llenar el proveedor)
router.get("/proveedor", async (req, res) => {
  try {
    const found = await lookupSupplier(String(req.query.docType || ""), String(req.query.id || ""));
    res.json({ ok: true, found: !!found, ...(found || {}) });
  } catch (err) { fail(res, err); }
});

// Datos de un proveedor que ya radicó antes (proveedores recurrentes)
router.get("/proveedor-previo", async (req, res) => {
  try {
    const found = await findPreviousSupplier(actor(req), String(req.query.docType || ""), String(req.query.id || ""));
    res.json({ ok: true, found: !!found, ...(found || {}) });
  } catch (err) { fail(res, err); }
});

router.get("/", async (req, res) => {
  try { res.json({ ok: true, rows: (await listSupportDocs(actor(req))).map(r => ({ ...r, totals: r.input ? supportTotals(r.input) : undefined })) }); }
  catch (err) { fail(res, err, 500); }
});

router.get("/:id", async (req, res) => {
  try { res.json({ ok: true, doc: withTotals(await getSupportDoc(actor(req), req.params.id)) }); }
  catch (err) { fail(res, err, 404); }
});

router.post("/", async (req, res) => {
  try { res.json({ ok: true, ...(await saveSupportDoc(actor(req), req.body)) }); }
  catch (err) { fail(res, err); }
});

router.put("/:id", async (req, res) => {
  try { res.json({ ok: true, ...(await saveSupportDoc(actor(req), req.body, req.params.id)) }); }
  catch (err) { fail(res, err); }
});

router.delete("/:id", async (req, res) => {
  try { await deleteSupportDoc(actor(req), req.params.id); res.json({ ok: true }); }
  catch (err) { fail(res, err); }
});

// PDF de la cuenta de cobro del proveedor
router.post("/:id/pdf", safeUpload(upload.single("file") as any), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ ok: false, message: "Adjunte el PDF de la cuenta de cobro." });
    await attachBillPdf(actor(req), req.params.id, req.file);
    res.json({ ok: true });
  } catch (err) { fail(res, err); }
});

router.get("/:id/pdf", async (req, res) => {
  try {
    const f = await getBillPdf(actor(req), req.params.id);
    if (!f) return res.status(404).json({ ok: false, message: "Este documento no tiene PDF de la cuenta de cobro." });
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `inline; filename="${encodeURIComponent(f.filename)}"`);
    res.send(f.data);
  } catch (err) { fail(res, err, 404); }
});

// Documentos adicionales (opcionales): subir, ver y quitar
router.post("/:id/adjuntos", safeUpload(upload.single("file") as any), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ ok: false, message: "Adjunte un archivo." });
    res.json({ ok: true, adjunto: await attachExtra(actor(req), req.params.id, req.file) });
  } catch (err) { fail(res, err); }
});
router.get("/:id/adjuntos/:extraId", async (req, res) => {
  try {
    const f = await getExtra(actor(req), req.params.id, req.params.extraId);
    if (!f) return res.status(404).json({ ok: false, message: "Archivo no encontrado." });
    res.setHeader("Content-Type", f.contentType);
    res.setHeader("Content-Disposition", `inline; filename="${encodeURIComponent(f.filename)}"`);
    res.send(f.data);
  } catch (err) { fail(res, err, 404); }
});
router.delete("/:id/adjuntos/:extraId", async (req, res) => {
  try { await deleteExtra(actor(req), req.params.id, req.params.extraId); res.json({ ok: true }); }
  catch (err) { fail(res, err); }
});

// Emite el documento soporte de PRUEBA con Alanube
router.post("/:id/emitir", async (req, res) => {
  try { res.json({ ok: true, emission: await emitSupportDoc(actor(req), req.params.id) }); }
  catch (err) { console.error("[SoporteDoc] emitir:", (err as Error)?.message); fail(res, err); }
});

// Archivos oficiales del documento emitido: ZIP (AttachedDocument + PDF), PDF o XML
router.get("/:id/archivo/:kind", async (req, res) => {
  try {
    const kind = (["zip", "pdf", "xml"] as const).find(k => k === req.params.kind);
    if (!kind) return res.status(400).json({ ok: false, message: "Solicitud inválida" });
    const f = await getEmittedFile(actor(req), req.params.id, kind);
    if (!f) return res.status(404).json({ ok: false, message: "No hay archivo oficial para este documento." });
    res.setHeader("Content-Type", f.contentType);
    res.setHeader("Content-Disposition", `${kind === "pdf" ? "inline" : "attachment"}; filename="${f.filename}"`);
    res.send(f.data);
  } catch (err) { fail(res, err, 404); }
});

export default router;
