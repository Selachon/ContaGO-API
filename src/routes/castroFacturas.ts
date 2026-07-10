import { Router, type Request, type Response, type NextFunction } from "express";
import multer from "multer";
import fs from "fs";
import path from "path";
import jwt from "jsonwebtoken";
import { requireAuth } from "../middleware/auth.js";
import { getDb } from "../services/database.js";
import { getAuthUrl } from "../services/googleDrive.js";
import { parsePdfInvoice } from "./pdfParse.js";
import {
  upsertFactura,
  listFacturas,
  getFactura,
  claimFactura,
  rechazarFactura,
  setDriveLink,
  setSheetsRow,
  getConfig,
  saveConfig,
  clearGoogleAuth,
  appendToSheet,
  appendToPersonalesSheet,
  uploadPdfToDrive,
  uploadAnexoToDrive,
  ensureIndexes,
} from "../services/castroFacturasService.js";

const router = Router();
const PDF_DIR = path.resolve("data/castro-pdfs");
fs.mkdirSync(PDF_DIR, { recursive: true });
ensureIndexes().catch((e) => console.error("[castro] index error:", e));

// ── Middleware: autenticación de portal ───────────────────────────────────────

function portalAuth(req: Request, res: Response, next: NextFunction): void {
  const authHeader = req.headers.authorization;
  const token = authHeader?.startsWith("Bearer ")
    ? authHeader.slice(7)
    : (req.query.token as string | undefined);

  if (!token) { res.status(401).json({ error: "No autenticado." }); return; }

  try {
    const payload = jwt.verify(token, process.env.JWT_SECRET!) as { email: string; purpose: string };
    if (payload.purpose !== "castro-portal") throw new Error();
    (req as Request & { portalEmail: string }).portalEmail = payload.email;
    next();
  } catch {
    res.status(401).json({ error: "Sesión expirada. Vuelve a iniciar sesión." });
  }
}

// ── Portal: iniciar login con Google ─────────────────────────────────────────

router.get("/auth/portal/login", (req: Request, res: Response) => {
  const returnUrl = String(req.query.returnUrl || "https://contago.com.co/public/facturas-mob");
  const PORTAL_SCOPES = [
    "https://www.googleapis.com/auth/userinfo.email",
    "openid",
  ];
  const state = Buffer.from(JSON.stringify({ purpose: "castro-portal", returnUrl })).toString("base64");
  const authUrl = getAuthUrl(state, PORTAL_SCOPES);
  res.redirect(302, authUrl);
});

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024, files: 500 },
});

const uploadAnexos = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 15 * 1024 * 1024, files: 5 },
});

// ── Admin: initiate Google OAuth for Castro ───────────────────────────────────

router.get("/auth/google", requireAuth, (_req: Request, res: Response) => {
  const CASTRO_SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/userinfo.email",
  ];
  const state = Buffer.from(JSON.stringify({ purpose: "castro" })).toString("base64");
  const authUrl = getAuthUrl(state, CASTRO_SCOPES);
  res.json({ ok: true, authUrl });
});

// ── Admin: Google connection status ──────────────────────────────────────────

router.get("/auth/google/status", requireAuth, async (_req: Request, res: Response) => {
  const config = await getConfig();
  if (!config.googleAuth) {
    res.json({ connected: false });
    return;
  }
  res.json({
    connected: true,
    email: config.googleAuth.user_email,
    connectedAt: config.googleAuth.connected_at,
    driveFolderId: config.googleAuth.drive_folder_id,
  });
});

// ── Admin: disconnect Google ──────────────────────────────────────────────────

router.post("/auth/google/disconnect", requireAuth, async (_req: Request, res: Response) => {
  await clearGoogleAuth();
  res.json({ ok: true });
});

// ── Admin: upload PDF batch ───────────────────────────────────────────────────

router.post(
  "/upload",
  requireAuth,
  upload.array("pdfs", 500),
  async (req: Request, res: Response) => {
    const files = (req.files as Express.Multer.File[]) || [];
    if (!files.length) { res.status(400).json({ error: "No se recibieron PDFs." }); return; }

    const config = await getConfig();
    const driveConnected = Boolean(config.googleAuth?.drive_folder_id);

    const results: { filename: string; cufe?: string; driveLink?: string; error?: string; duplicado?: boolean }[] = [];

    for (const file of files) {
      try {
        const invoice = await parsePdfInvoice(file.buffer, file.originalname);
        const cufe = invoice.cufe || file.originalname.replace(/\.pdf$/i, "");

        // Verificar duplicado por CUFE antes de procesar
        const existente = await getFactura(cufe);
        if (existente) {
          results.push({ filename: file.originalname, cufe, duplicado: true });
          continue;
        }

        // Save PDF to disk
        const pdfFilename = `${cufe}.pdf`;
        const pdfPath = path.join(PDF_DIR, pdfFilename);
        fs.writeFileSync(pdfPath, file.buffer);

        const descriptions = (invoice.lineItems || [])
          .map((li) => li.description)
          .filter(Boolean)
          .join(" / ")
          .slice(0, 400);

        await upsertFactura({
          cufe,
          docNumber: invoice.docNumber || "",
          issuerName: invoice.issuerName || invoice.issuerCommercialName || "Proveedor sin nombre",
          issuerNit: invoice.issuerNit || "",
          issueDateISO: invoice.issueDateISO || "",
          subtotal: invoice.subtotal || 0,
          iva: invoice.iva || 0,
          total: invoice.total || 0,
          lineItemsDescription: descriptions,
          pdfFilename,
          pdfPath,
        });

        // Upload to Drive immediately if connected
        let driveLink: string | undefined;
        if (driveConnected) {
          try {
            driveLink = await uploadPdfToDrive(file.buffer, pdfFilename);
            await setDriveLink(cufe, driveLink);
          } catch (e) {
            console.error(`[castro] Drive upload failed for ${pdfFilename}:`, e);
          }
        }

        results.push({ filename: file.originalname, cufe, driveLink });
      } catch (e) {
        results.push({ filename: file.originalname, error: (e as Error).message });
      }
    }

    const ok = results.filter((r) => !r.error && !r.duplicado).length;
    const duplicados = results.filter((r) => r.duplicado);
    const errors = results.filter((r) => r.error);
    res.json({ ok, total: files.length, duplicados: duplicados.length, driveConnected, duplicadosList: duplicados.length ? duplicados : undefined, errors: errors.length ? errors : undefined });
  }
);

// ── Admin: list all facturas ──────────────────────────────────────────────────

router.get("/admin/facturas", requireAuth, async (_req: Request, res: Response) => {
  const facturas = await listFacturas();
  res.json({ facturas });
});

// ── Admin: config ─────────────────────────────────────────────────────────────

router.get("/config", requireAuth, async (_req: Request, res: Response) => {
  const config = await getConfig();
  res.json({ empleados: config.empleados, socios: config.socios || [], spreadsheetId: config.spreadsheetId });
});

router.put("/config", requireAuth, async (req: Request, res: Response) => {
  const { empleados, spreadsheetId } = req.body;
  const patch: Record<string, unknown> = {};

  if (Array.isArray(empleados)) {
    patch.empleados = empleados.map((e: unknown) => String(e).trim().toLowerCase()).filter(Boolean);
  }
  if (Array.isArray(req.body.socios)) {
    patch.socios = req.body.socios.map((e: unknown) => String(e).trim().toLowerCase()).filter(Boolean);
  }
  if (typeof spreadsheetId === "string") {
    // Accept full URL or just the ID
    const idMatch = spreadsheetId.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    patch.spreadsheetId = idMatch ? idMatch[1] : spreadsheetId.trim();
  }

  await saveConfig(patch as never);
  const updated = await getConfig();
  res.json({ ok: true, empleados: updated.empleados, spreadsheetId: updated.spreadsheetId });
});

// ── Public: list facturas ─────────────────────────────────────────────────────

router.get("/public/facturas", portalAuth, async (_req: Request, res: Response) => {
  const facturas = await listFacturas();
  const safe = facturas.map(({ pdfPath: _p, ...rest }) => rest);
  res.json({ facturas: safe });
});

// ── Public: employee list ─────────────────────────────────────────────────────

router.get("/public/config", portalAuth, async (_req: Request, res: Response) => {
  const config = await getConfig();
  res.json({ empleados: config.empleados, socios: config.socios || [] });
});

// ── Public: serve PDF ─────────────────────────────────────────────────────────

router.get("/public/pdf/:cufe", portalAuth, async (req: Request, res: Response) => {
  const cufe = req.params.cufe.replace(/[^a-f0-9A-F]/g, "").slice(0, 200);
  const factura = await getFactura(cufe);
  if (!factura) { res.status(404).json({ error: "Factura no encontrada." }); return; }

  // Servir desde disco local si existe (dev)
  const pdfPath = path.join(PDF_DIR, `${cufe}.pdf`);
  if (fs.existsSync(pdfPath)) {
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `inline; filename="${factura.pdfFilename}"`);
    fs.createReadStream(pdfPath).pipe(res);
    return;
  }

  // En producción el disco es efímero: redirigir a Drive si está disponible
  if (factura.pdfDriveLink) {
    res.redirect(302, factura.pdfDriveLink);
    return;
  }

  res.status(404).json({ error: "PDF no disponible. Sube el archivo nuevamente." });
});

// ── Public: claim a factura ───────────────────────────────────────────────────

router.post("/public/reclamar/:cufe", portalAuth, uploadAnexos.array("anexos", 5), async (req: Request, res: Response) => {
  const { cufe } = req.params;
  const { correo, concepto, proyecto, formaPago, detalleFormaPago, esSocio } = req.body;

  if (!correo || !concepto || !proyecto || !formaPago) {
    res.status(400).json({ error: "Faltan campos: correo, concepto, proyecto, formaPago." });
    return;
  }

  const factura = await getFactura(cufe);
  if (!factura) { res.status(404).json({ error: "Factura no encontrada." }); return; }
  if (factura.estado === "reclamada") { res.status(409).json({ error: "Esta factura ya fue reclamada." }); return; }

  // Subir anexos a Drive antes de guardar el claim
  const anexosFiles = (req.files as Express.Multer.File[]) || [];
  const anexosDriveLinks: string[] = [];
  if (anexosFiles.length) {
    const config = await getConfig();
    if (config.googleAuth?.drive_folder_id) {
      for (const file of anexosFiles) {
        try {
          const link = await uploadAnexoToDrive(file.buffer, file.originalname, cufe);
          anexosDriveLinks.push(link);
        } catch (e) {
          console.error(`[castro] Error subiendo anexo ${file.originalname}:`, e);
        }
      }
    }
  }

  const claim = {
    correo: String(correo).trim(),
    concepto: String(concepto).trim(),
    proyecto: String(proyecto).trim(),
    formaPago: String(formaPago).trim(),
    detalleFormaPago: String(detalleFormaPago || "").trim(),
    esSocio: esSocio === "true" || esSocio === true,
    anexosDriveLinks,
    reclamadaAt: new Date(),
    sheetsRow: null as number | null,
  };

  const ok = await claimFactura(cufe, claim);
  if (!ok) { res.status(409).json({ error: "Esta factura ya fue reclamada." }); return; }

  // Append to Google Sheets (hoja según si es socio o no)
  let sheetsRow = 0;
  let sheetsWarning: string | undefined;
  try {
    const facturaFresh = await getFactura(cufe);
    if (claim.esSocio) {
      sheetsRow = await appendToPersonalesSheet(facturaFresh!, claim);
    } else {
      sheetsRow = await appendToSheet(facturaFresh!, claim);
    }
    if (sheetsRow > 0) await setSheetsRow(cufe, sheetsRow);
  } catch (e) {
    console.error("[castro] Sheets append failed:", e);
    sheetsWarning = (e as Error).message;
  }

  res.json({ ok: true, sheetsRow, driveLink: factura.pdfDriveLink, sheetsWarning });
});

// ── Public: rechazar una factura ──────────────────────────────────────────────

router.post("/public/rechazar/:cufe", portalAuth, async (req: Request, res: Response) => {
  const { cufe } = req.params;
  const { correo, motivo } = req.body;

  if (!correo || !motivo) {
    res.status(400).json({ error: "Faltan campos: correo, motivo." });
    return;
  }

  const factura = await getFactura(cufe);
  if (!factura) { res.status(404).json({ error: "Factura no encontrada." }); return; }
  if (factura.estado !== "pendiente") { res.status(409).json({ error: "La factura ya fue gestionada." }); return; }

  const rechazo = {
    correo: String(correo).trim(),
    motivo: String(motivo).trim(),
    rechazadaAt: new Date(),
  };

  const ok = await rechazarFactura(cufe, rechazo);
  if (!ok) { res.status(409).json({ error: "La factura ya fue gestionada." }); return; }

  res.json({ ok: true });
});

export default router;
