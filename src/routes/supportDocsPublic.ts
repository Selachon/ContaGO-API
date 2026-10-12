import { Router, Request, Response, NextFunction } from "express";
import multer from "multer";
import { rateLimit } from "../middleware/rateLimit.js";
import { getDianPlaces } from "../services/alanubeService.js";
import { createPublicRadicado, isValidPublicToken, previousSupplierPublic, MAX_EXTRAS, SUPPLIER_ID_TYPES } from "../services/supportDocsService.js";

/**
 * Enlace público para que los proveedores radiquen sus cuentas de cobro (sin usuario ni contraseña).
 * Protegido por un token secreto en la URL (renovable por el administrador), límite de peticiones por IP, un campo trampa
 * contra robots y validación del PDF. NO expone ningún dato de otros proveedores sin verificar el correo.
 */
const router = Router();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024, files: 1 + MAX_EXTRAS } });
const fail = (res: Response, err: any, status = 400) => res.status(status).json({ ok: false, message: err?.message || "Error" });

/** Envuelve un middleware de multer para que sus errores (demasiados archivos, archivo muy pesado…) salgan como JSON y nunca como una página de error. */
const safeUpload = (mw: (req: Request, res: Response, cb: (err?: any) => void) => void) => (req: Request, res: Response, next: NextFunction) => {
  mw(req, res, (err?: any) => {
    if (!err) return next();
    const msg = err.code === "LIMIT_FILE_SIZE" ? "Un archivo pesa más de 10 MB."
      : err.code === "LIMIT_FILE_COUNT" || err.code === "LIMIT_UNEXPECTED_FILE" ? `Máximo ${MAX_EXTRAS} documentos adicionales y un PDF de cuenta de cobro.`
      : "No se pudieron leer los archivos adjuntos.";
    res.status(400).json({ ok: false, message: msg });
  });
};


router.use("/:token", async (req: Request, res: Response, next: NextFunction) => {
  try {
    if (!(await isValidPublicToken(req.params.token))) return res.status(404).json({ ok: false, message: "Este enlace no es válido o fue renovado. Pídele a ContaGO el enlace vigente." });
    next();
  } catch (err) { fail(res, err, 500); }
});

// Catálogos para el formulario (sin datos internos como la resolución)
router.get("/:token/catalogs", rateLimit(60, 10 * 60_000), async (_req, res) => {
  try {
    const p = await getDianPlaces();
    res.json({ ok: true, departments: p.departments, municipalities: p.municipalities, countries: p.countries, idTypes: SUPPLIER_ID_TYPES });
  } catch (err) { fail(res, err, 502); }
});

// Cargar datos de un radicado anterior con solo el número de documento (límite de consultas por IP)
router.post("/:token/proveedor-previo", rateLimit(20, 10 * 60_000), async (req, res) => {
  try {
    const found = await previousSupplierPublic(String(req.body?.docType || ""), String(req.body?.id || ""));
    res.json({ ok: true, found: !!found, ...(found ? { supplier: found } : {}) });
  } catch (err) { fail(res, err); }
});

// Radicar: campos como JSON en `data` + el PDF en `file`
router.post("/:token/radicar", rateLimit(15, 60 * 60_000), safeUpload(upload.fields([{ name: "file", maxCount: 1 }, { name: "extras", maxCount: MAX_EXTRAS }]) as any), async (req, res) => {
  try {
    if (String(req.body?.website || "").trim()) return res.json({ ok: true, radicado: "RAD-000000" }); // campo trampa: un robot lo llena, una persona no
    let raw: any;
    try { raw = JSON.parse(String(req.body?.data || "{}")); } catch { return res.status(400).json({ ok: false, message: "Datos inválidos." }); }
    const files = req.files as Record<string, Express.Multer.File[]> | undefined;
    const r = await createPublicRadicado(raw, files?.file?.[0], files?.extras || [], { ip: req.ip });
    res.json({ ok: true, ...r });
  } catch (err: any) {
    if (err?.code === "LIMIT_FILE_SIZE") return res.status(400).json({ ok: false, message: "Un archivo pesa más de 10 MB." });
    if (err?.code === "LIMIT_UNEXPECTED_FILE" || err?.code === "LIMIT_FILE_COUNT") return res.status(400).json({ ok: false, message: `Máximo ${MAX_EXTRAS} documentos adicionales.` });
    fail(res, err);
  }
});

export default router;
