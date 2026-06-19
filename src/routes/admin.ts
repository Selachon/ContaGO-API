import { Router, Request, Response } from "express";
import { ObjectId } from "mongodb";
import ExcelJS from "exceljs";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";
import { v4 as uuidv4 } from "uuid";
import JSZip from "jszip";
import multer from "multer";
import {
  listUsers,
  getUserById,
  updateUser,
  suspendUser,
  reactivateUser,
  logAdminAction,
  getAuditLogs,
  getAllUsersForExport,
  getPortalStatusConfig,
  updatePortalStatusConfig,
  getAccountingKanbanBoard,
  getAdminKanbanBoard,
  listAdminUsersForKanban,
  updateAccountingKanbanBoard,
  updateAdminKanbanBoard,
  type AdminKanbanBoard,
  type PortalStatusConfig,
} from "../services/adminService.js";
import { requireAuth } from "../middleware/auth.js";
import { createDemoInvite, listDemoInvites, TOOL_SUCCESSOR, updateUserPassword } from "../services/database.js";
import { extractInvoiceDataFromXml } from "../services/xmlParser.js";
import { generateExcelFile, generateExcelFilename } from "../services/excelGenerator.js";
import type { InvoiceData } from "../types/dianExcel.js";

const router = Router();
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DOWNLOADS_DIR = path.join(__dirname, "../../downloads");
const manualDianUpload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 50 * 1024 * 1024,
    files: 2000,
  },
});

const DEMO_ALLOWED_TOOLS = new Set([
  "dian-cufe-downloader",
  "dian-mass-download",
  "dian-third-parties-excel",
]);

function generateTemporaryPassword(): string {
  return String(Math.floor(100000 + Math.random() * 900000));
}

function resolveFrontendBaseUrl(req: Request, rawBaseUrl: unknown): string {
  const requested = typeof rawBaseUrl === "string" ? rawBaseUrl.trim() : "";
  if (/^https?:\/\//i.test(requested)) return requested.replace(/\/+$/, "");

  const envUrl = process.env.FRONTEND_URL || process.env.CORS_ORIGIN?.split(",")[0]?.trim();
  if (envUrl && /^https?:\/\//i.test(envUrl)) return envUrl.replace(/\/+$/, "");

  return (req.protocol + "://" + req.get("host")).replace(/\/+$/, "");
}

// Todas las rutas admin requieren autenticacion y rol admin.
router.use(requireAuth);
router.use((req: Request, res: Response, next) => {
  if (!req.user?.isAdmin) {
    return res.status(403).json({ ok: false, message: "Acceso denegado. Se requiere rol de administrador." });
  }
  next();
});

function normalizeManualValue(value: string | undefined, fallback = "N/A"): string {
  const trimmed = String(value || "").trim();
  return trimmed || fallback;
}

function buildManualInvoiceRow(invoiceData: Partial<InvoiceData>, fallbackDocNumber: string): InvoiceData {
  const cufe = normalizeManualValue(invoiceData.cufe || invoiceData.trackId, "N/A");
  const docNumber = normalizeManualValue(invoiceData.docNumber, fallbackDocNumber);

  return {
    issuerNit: normalizeManualValue(invoiceData.issuerNit),
    issuerName: normalizeManualValue(invoiceData.issuerName),
    issuerEmail: normalizeManualValue(invoiceData.issuerEmail),
    issuerPhone: normalizeManualValue(invoiceData.issuerPhone),
    issuerAddress: normalizeManualValue(invoiceData.issuerAddress),
    issuerCity: normalizeManualValue(invoiceData.issuerCity),
    issuerDepartment: normalizeManualValue(invoiceData.issuerDepartment),
    issuerCountry: normalizeManualValue(invoiceData.issuerCountry),
    issuerCommercialName: normalizeManualValue(invoiceData.issuerCommercialName),
    issuerTaxpayerType: normalizeManualValue(invoiceData.issuerTaxpayerType),
    issuerFiscalRegime: normalizeManualValue(invoiceData.issuerFiscalRegime),
    issuerTaxResponsibility: normalizeManualValue(invoiceData.issuerTaxResponsibility),
    issuerEconomicActivity: normalizeManualValue(invoiceData.issuerEconomicActivity),
    receiverNit: normalizeManualValue(invoiceData.receiverNit),
    receiverName: normalizeManualValue(invoiceData.receiverName),
    receiverEmail: normalizeManualValue(invoiceData.receiverEmail),
    receiverPhone: normalizeManualValue(invoiceData.receiverPhone),
    receiverAddress: normalizeManualValue(invoiceData.receiverAddress),
    receiverCity: normalizeManualValue(invoiceData.receiverCity),
    receiverDepartment: normalizeManualValue(invoiceData.receiverDepartment),
    receiverCountry: normalizeManualValue(invoiceData.receiverCountry),
    receiverCommercialName: normalizeManualValue(invoiceData.receiverCommercialName),
    receiverTaxpayerType: normalizeManualValue(invoiceData.receiverTaxpayerType),
    receiverFiscalRegime: normalizeManualValue(invoiceData.receiverFiscalRegime),
    receiverTaxResponsibility: normalizeManualValue(invoiceData.receiverTaxResponsibility),
    receiverEconomicActivity: normalizeManualValue(invoiceData.receiverEconomicActivity),
    issueDate: normalizeManualValue(invoiceData.issueDate),
    issueDateISO: invoiceData.issueDateISO || "9999-12-31",
    paymentMethod: normalizeManualValue(invoiceData.paymentMethod),
    subtotal: invoiceData.subtotal || 0,
    iva: invoiceData.iva || 0,
    total: invoiceData.total || 0,
    taxes: invoiceData.taxes || [],
    discount: invoiceData.discount || 0,
    surcharge: invoiceData.surcharge || 0,
    concepts: normalizeManualValue(invoiceData.concepts),
    lineItems: invoiceData.lineItems || [],
    documentType: normalizeManualValue(invoiceData.documentType, "Factura Electrónica"),
    isDocumentoSoporte: invoiceData.isDocumentoSoporte || false,
    cufe,
    notes: invoiceData.notes || "",
    trackId: cufe,
    docNumber,
    zipFilename: `${normalizeManualValue(invoiceData.issuerNit)} - ${docNumber}.zip`,
  };
}

async function collectXmlBuffersFromAdminUpload(files: Express.Multer.File[]): Promise<Array<{ buffer: Buffer; filename: string }>> {
  const xmlFiles: Array<{ buffer: Buffer; filename: string }> = [];

  for (const uploaded of files) {
    const lowerName = uploaded.originalname.toLowerCase();

    if (lowerName.endsWith(".xml")) {
      xmlFiles.push({ buffer: uploaded.buffer, filename: uploaded.originalname });
      continue;
    }

    if (!lowerName.endsWith(".zip")) continue;

    const zip = await JSZip.loadAsync(uploaded.buffer);
    for (const [filename, file] of Object.entries(zip.files)) {
      if (file.dir || !filename.toLowerCase().endsWith(".xml")) continue;
      xmlFiles.push({
        buffer: await file.async("nodebuffer"),
        filename: `${uploaded.originalname}/${filename}`,
      });
    }
  }

  return xmlFiles;
}

function inferManualCompany(invoices: InvoiceData[], direction: "received" | "sent", companyName: string, companyNit: string) {
  let inferredName = companyName;
  let inferredNit = companyNit;

  if (inferredName && inferredNit) return { companyName: inferredName, companyNit: inferredNit };

  for (const invoice of invoices) {
    const nit = direction === "sent" ? invoice.issuerNit : invoice.receiverNit;
    const name = direction === "sent" ? invoice.issuerName : invoice.receiverName;
    if (!inferredNit && nit && nit !== "N/A") inferredNit = nit;
    if (!inferredName && name && name !== "N/A") inferredName = name;
    if (inferredName && inferredNit) break;
  }

  return { companyName: inferredName, companyNit: inferredNit };
}

// ============================================
// POST /admin/dian-xml-excel - Excel DIAN desde XML subidos manualmente
// ============================================
router.post("/dian-xml-excel", manualDianUpload.array("files", 2000), async (req: Request, res: Response) => {
  const uploadedFiles = Array.isArray(req.files) ? req.files as Express.Multer.File[] : [];
  const direction = req.body?.document_direction === "sent" ? "sent" : "received";
  const requestedCompanyName = typeof req.body?.company_name === "string" ? req.body.company_name.trim() : "";
  const requestedCompanyNit = typeof req.body?.company_nit === "string" ? req.body.company_nit.trim() : "";

  if (uploadedFiles.length === 0) {
    return res.status(400).json({ ok: false, message: "Sube al menos un archivo .xml o .zip con XML." });
  }

  let excelPath = "";
  try {
    fs.mkdirSync(DOWNLOADS_DIR, { recursive: true });

    const xmlFiles = await collectXmlBuffersFromAdminUpload(uploadedFiles);
    if (xmlFiles.length === 0) {
      return res.status(400).json({ ok: false, message: "No se encontraron XML en los archivos subidos." });
    }

    const invoices: InvoiceData[] = [];
    for (let i = 0; i < xmlFiles.length; i++) {
      const xmlFile = xmlFiles[i];
      const fallbackDocNumber = path.basename(xmlFile.filename).replace(/\.xml$/i, "") || `XML-${i + 1}`;
      const parsed = await extractInvoiceDataFromXml(xmlFile.buffer, {
        id: fallbackDocNumber,
        docnum: fallbackDocNumber,
      });
      invoices.push(buildManualInvoiceRow(parsed, fallbackDocNumber));
    }

    if (invoices.length === 0) {
      return res.status(400).json({ ok: false, message: "No se pudo construir ninguna fila para el Excel." });
    }

    const { companyName, companyNit } = inferManualCompany(invoices, direction, requestedCompanyName, requestedCompanyNit);
    const jobId = `admin-xml-${uuidv4()}`;
    excelPath = path.join(DOWNLOADS_DIR, `${jobId}.xlsx`);
    await generateExcelFile(invoices, excelPath, false, direction === "sent", companyName, companyNit);

    const basePrefix = direction === "sent" ? "Facturas Emitidas DIAN" : "Facturas DIAN";
    const filePrefix = companyName
      ? `${companyNit ? `${companyNit} - ` : ""}${companyName} - ${basePrefix}`
      : (companyNit ? `${companyNit} - ${basePrefix}` : basePrefix);
    const filename = generateExcelFilename(undefined, undefined, filePrefix);

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${encodeURIComponent(filename)}"`);

    const stream = fs.createReadStream(excelPath);
    stream.pipe(res);
    stream.on("close", () => {
      if (excelPath && fs.existsSync(excelPath)) {
        try { fs.unlinkSync(excelPath); } catch {}
      }
    });
  } catch (err) {
    console.error("[Admin] Error generando Excel DIAN manual:", err);
    if (excelPath && fs.existsSync(excelPath)) {
      try { fs.unlinkSync(excelPath); } catch {}
    }
    if (!res.headersSent) {
      res.status(500).json({ ok: false, message: (err as Error).message || "Error al generar Excel desde XML." });
    }
  }
});

// ============================================
// GET /admin/users/export - Exportar todos los usuarios como Excel
// ============================================
router.get("/users/export", async (req: Request, res: Response) => {
  try {
    const users = await getAllUsersForExport();

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Usuarios");

    ws.columns = [
      { header: "Nombre", key: "name", width: 28 },
      { header: "Email", key: "email", width: 32 },
      { header: "Estado", key: "status", width: 12 },
      { header: "Admin", key: "isAdmin", width: 8 },
      { header: "Rol", key: "role", width: 12 },
      { header: "Teléfono", key: "phone", width: 16 },
      { header: "Herramientas", key: "purchasedTools", width: 40 },
      { header: "NITs", key: "nits", width: 40 },
      { header: "Valor pago", key: "paymentAmount", width: 14 },
      { header: "Forma de pago", key: "paymentMethod", width: 22 },
      { header: "Inicio licencia", key: "licenseStartDate", width: 16, style: { numFmt: "yyyy-mm-dd" } },
      { header: "Fin licencia", key: "licenseEndDate", width: 16, style: { numFmt: "yyyy-mm-dd" } },
      { header: "Empresas en plan", key: "companiesInPlan", width: 18 },
      { header: "Factura ref.", key: "invoiceRef", width: 20 },
      { header: "Fecha creación", key: "createdAt", width: 20, style: { numFmt: "yyyy-mm-dd hh:mm" } },
    ];

    // Header style
    ws.getRow(1).eachCell((cell) => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF2563EB" } };
      cell.alignment = { vertical: "middle", horizontal: "center" };
    });
    ws.getRow(1).height = 20;

    const parseDate = (d: string | undefined) => {
      if (!d || !/^\d{4}-\d{2}-\d{2}$/.test(d)) return null;
      return new Date(d + "T12:00:00");
    };

    for (const u of users) {
      ws.addRow({
        name: u.name,
        email: u.email,
        status: u.status === "active" ? "Activo" : "Suspendido",
        isAdmin: u.isAdmin ? "Sí" : "No",
        role: u.role || (u.isAdmin ? "ADMIN" : "USER"),
        phone: u.phone ?? "",
        purchasedTools: u.purchasedTools.join(", "),
        nits: u.nits.join(", "),
        paymentAmount: u.paymentAmount ?? "",
        paymentMethod: u.paymentMethod ?? "",
        licenseStartDate: parseDate(u.licenseStartDate),
        licenseEndDate: parseDate(u.licenseEndDate),
        companiesInPlan: u.companiesInPlan ?? "",
        invoiceRef: u.invoiceRef ?? "",
        createdAt: u.createdAt ? new Date(u.createdAt) : null,
      });
    }

    // Freeze header row
    ws.views = [{ state: "frozen", ySplit: 1 }];

    const filename = `usuarios_contago_${new Date().toISOString().slice(0, 10)}.xlsx`;
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);

    await wb.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error("[Admin] Error exportando usuarios:", err);
    res.status(500).json({ ok: false, message: "Error al exportar usuarios" });
  }
});

// ============================================
// GET /admin/demo-invites - Listar invitaciones DEMO recientes
// ============================================
router.get("/demo-invites", async (_req: Request, res: Response) => {
  try {
    const invites = await listDemoInvites(40);
    res.json({ ok: true, invites });
  } catch (err) {
    console.error("[Admin] Error listando invitaciones demo:", err);
    res.status(500).json({ ok: false, message: "Error al listar invitaciones DEMO" });
  }
});

// ============================================
// POST /admin/demo-invites - Crear enlace DEMO de un uso
// ============================================
router.post("/demo-invites", async (req: Request, res: Response) => {
  try {
    const actorId = req.user!.userId;
    const toolId = typeof req.body?.toolId === "string" ? req.body.toolId.trim() : "";

    if (!DEMO_ALLOWED_TOOLS.has(toolId)) {
      return res.status(400).json({ ok: false, message: "Selecciona una herramienta válida para la DEMO" });
    }

    const created = await createDemoInvite(toolId, actorId);
    if (!created) {
      return res.status(409).json({
        ok: false,
        message: "No se pudo crear el enlace DEMO. Intenta nuevamente.",
      });
    }

    const baseUrl = resolveFrontendBaseUrl(req, req.body?.frontendBaseUrl);
    const inviteUrl = baseUrl + "/demo/" + created.token;

    await logAdminAction({
      actorId,
      action: "create_demo_invite",
      after: {
        inviteId: created.invite.id,
        toolId,
        expiresAt: created.invite.expiresAt,
      },
    });

    res.status(201).json({
      ok: true,
      inviteUrl,
      token: created.token,
      invite: created.invite,
      message: "Enlace DEMO creado. Es de un solo uso; el cliente escribirá su NIT al activarlo.",
    });
  } catch (err) {
    console.error("[Admin] Error creando invitación demo:", err);
    res.status(500).json({ ok: false, message: "Error al crear invitación DEMO" });
  }
});

// ============================================
// GET /admin/users - Listar usuarios con paginacion y filtros
// ============================================
router.get("/users", async (req: Request, res: Response) => {
  try {
    const page = Math.max(1, parseInt(req.query.page as string) || 1);
    const limit = Math.min(100, Math.max(1, parseInt(req.query.limit as string) || 20));
    const search = (req.query.search as string)?.trim() || "";
    const status = req.query.status as string | undefined;
    const tool = req.query.tool as string | undefined;

    const result = await listUsers({ page, limit, search, status, tool });

    res.json({
      ok: true,
      users: result.users,
      pagination: {
        page,
        limit,
        total: result.total,
        totalPages: Math.ceil(result.total / limit),
      },
    });
  } catch (err) {
    console.error("[Admin] Error listando usuarios:", err);
    res.status(500).json({ ok: false, message: "Error interno al listar usuarios" });
  }
});

// ============================================
// GET /admin/users/:id - Detalle de un usuario
// ============================================
router.get("/users/:id", async (req: Request, res: Response) => {
  try {
    const { id } = req.params;

    if (!ObjectId.isValid(id)) {
      return res.status(400).json({ ok: false, message: "ID de usuario invalido" });
    }

    const user = await getUserById(id);
    if (!user) {
      return res.status(404).json({ ok: false, message: "Usuario no encontrado" });
    }

    res.json({ ok: true, user });
  } catch (err) {
    console.error("[Admin] Error obteniendo usuario:", err);
    res.status(500).json({ ok: false, message: "Error interno al obtener usuario" });
  }
});

// ============================================
// PATCH /admin/users/:id - Editar usuario
// ============================================
router.patch("/users/:id", async (req: Request, res: Response) => {
  try {
    const { id } = req.params;
    const actorId = req.user!.userId;

    if (!ObjectId.isValid(id)) {
      return res.status(400).json({ ok: false, message: "ID de usuario invalido" });
    }

    const allowedFields = ["name", "nits", "purchasedTools", "isAdmin", "phone", "paymentAmount", "paymentMethod", "licenseStartDate", "licenseEndDate", "companiesInPlan", "invoiceRef"];
    const updates: Record<string, unknown> = {};
    const before: Record<string, unknown> = {};

    // Obtener estado actual para diff de auditoria
    const currentUser = await getUserById(id);
    if (!currentUser) {
      return res.status(404).json({ ok: false, message: "Usuario no encontrado" });
    }

    // Validar y recoger campos permitidos
    const currentUserObj = currentUser as unknown as Record<string, unknown>;
    for (const field of allowedFields) {
      if (req.body[field] !== undefined) {
        before[field] = currentUserObj[field];

        if (field === "name") {
          const name = req.body.name?.trim();
          if (!name || name.length < 2) {
            return res.status(400).json({ ok: false, message: "El nombre debe tener al menos 2 caracteres" });
          }
          updates.name = name;
        } else if (field === "nits") {
          const nits = Array.isArray(req.body.nits) ? req.body.nits : [];
          updates.nits = [...new Set(nits.filter((n: unknown) => typeof n === "string" && n.trim()).map((n: string) => n.trim()))];
        } else if (field === "purchasedTools") {
          const tools = Array.isArray(req.body.purchasedTools) ? req.body.purchasedTools : [];
          const normalized = tools
            .filter((t: unknown) => typeof t === "string" && (t as string).trim())
            .map((t: string) => TOOL_SUCCESSOR[t.trim()] || t.trim());
          updates.purchasedTools = [...new Set(normalized)];
        } else if (field === "isAdmin") {
          if (id === actorId && req.body.isAdmin === false) {
            return res.status(400).json({ ok: false, message: "No puedes quitarte el rol de administrador a ti mismo" });
          }
          updates.is_admin = Boolean(req.body.isAdmin);
        } else if (field === "phone") {
          updates.phone = req.body.phone ? String(req.body.phone).trim() : undefined;
        } else if (field === "paymentAmount") {
          const v = parseFloat(req.body.paymentAmount);
          updates.paymentAmount = isNaN(v) ? undefined : v;
        } else if (field === "paymentMethod") {
          updates.paymentMethod = req.body.paymentMethod ? String(req.body.paymentMethod).trim() : undefined;
        } else if (field === "licenseStartDate") {
          updates.licenseStartDate = req.body.licenseStartDate ? String(req.body.licenseStartDate).trim() : undefined;
        } else if (field === "licenseEndDate") {
          updates.licenseEndDate = req.body.licenseEndDate ? String(req.body.licenseEndDate).trim() : undefined;
        } else if (field === "companiesInPlan") {
          const v = parseInt(req.body.companiesInPlan, 10);
          updates.companiesInPlan = isNaN(v) ? undefined : v;
        } else if (field === "invoiceRef") {
          updates.invoiceRef = req.body.invoiceRef ? String(req.body.invoiceRef).trim() : undefined;
        }
      }
    }

    if (Object.keys(updates).length === 0) {
      return res.status(400).json({ ok: false, message: "No se proporcionaron campos validos para actualizar" });
    }

    const success = await updateUser(id, updates);
    if (!success) {
      return res.status(500).json({ ok: false, message: "Error al actualizar usuario" });
    }

    // Registrar auditoria
    await logAdminAction({
      actorId,
      action: "update_user",
      targetUserId: id,
      before,
      after: updates,
    });

    const updatedUser = await getUserById(id);
    res.json({ ok: true, user: updatedUser, message: "Usuario actualizado correctamente" });
  } catch (err) {
    console.error("[Admin] Error actualizando usuario:", err);
    res.status(500).json({ ok: false, message: "Error interno al actualizar usuario" });
  }
});

// ============================================
// POST /admin/users/:id/suspend - Suspender usuario
// ============================================
router.post("/users/:id/suspend", async (req: Request, res: Response) => {
  try {
    const { id } = req.params;
    const actorId = req.user!.userId;
    const reason = (req.body.reason as string)?.trim() || "Sin motivo especificado";

    if (!ObjectId.isValid(id)) {
      return res.status(400).json({ ok: false, message: "ID de usuario invalido" });
    }

    // No permitir auto-suspension
    if (id === actorId) {
      return res.status(400).json({ ok: false, message: "No puedes suspenderte a ti mismo" });
    }

    const user = await getUserById(id);
    if (!user) {
      return res.status(404).json({ ok: false, message: "Usuario no encontrado" });
    }

    if (user.status === "suspended") {
      return res.status(400).json({ ok: false, message: "El usuario ya esta suspendido" });
    }

    const success = await suspendUser(id);
    if (!success) {
      return res.status(500).json({ ok: false, message: "Error al suspender usuario" });
    }

    await logAdminAction({
      actorId,
      action: "suspend_user",
      targetUserId: id,
      reason,
    });

    res.json({ ok: true, message: "Usuario suspendido correctamente" });
  } catch (err) {
    console.error("[Admin] Error suspendiendo usuario:", err);
    res.status(500).json({ ok: false, message: "Error interno al suspender usuario" });
  }
});

// ============================================
// POST /admin/users/:id/reactivate - Reactivar usuario
// ============================================
router.post("/users/:id/reactivate", async (req: Request, res: Response) => {
  try {
    const { id } = req.params;
    const actorId = req.user!.userId;
    const reason = (req.body.reason as string)?.trim() || "Sin motivo especificado";

    if (!ObjectId.isValid(id)) {
      return res.status(400).json({ ok: false, message: "ID de usuario invalido" });
    }

    const user = await getUserById(id);
    if (!user) {
      return res.status(404).json({ ok: false, message: "Usuario no encontrado" });
    }

    if (user.status === "active") {
      return res.status(400).json({ ok: false, message: "El usuario ya esta activo" });
    }

    const success = await reactivateUser(id);
    if (!success) {
      return res.status(500).json({ ok: false, message: "Error al reactivar usuario" });
    }

    await logAdminAction({
      actorId,
      action: "reactivate_user",
      targetUserId: id,
      reason,
    });

    res.json({ ok: true, message: "Usuario reactivado correctamente" });
  } catch (err) {
    console.error("[Admin] Error reactivando usuario:", err);
    res.status(500).json({ ok: false, message: "Error interno al reactivar usuario" });
  }
});

// ============================================
// POST /admin/users/:id/reset-password - Restablecer contraseña
// ============================================
router.post("/users/:id/reset-password", async (req: Request, res: Response) => {
  try {
    const { id } = req.params;
    const actorId = req.user!.userId;

    if (!ObjectId.isValid(id)) {
      return res.status(400).json({ ok: false, message: "ID de usuario invalido" });
    }

    const user = await getUserById(id);
    if (!user) {
      return res.status(404).json({ ok: false, message: "Usuario no encontrado" });
    }

    const temporaryPassword = generateTemporaryPassword();
    const updated = await updateUserPassword(id, temporaryPassword, true);
    if (!updated) {
      return res.status(500).json({ ok: false, message: "No se pudo restablecer la contraseña" });
    }

    await logAdminAction({
      actorId,
      action: "reset_password",
      targetUserId: id,
      reason: "Restablecimiento manual por administrador",
    });

    return res.json({
      ok: true,
      message: "Contraseña restablecida correctamente",
      temporaryPassword,
    });
  } catch (err) {
    console.error("[Admin] Error restableciendo contraseña:", err);
    return res.status(500).json({ ok: false, message: "Error interno al restablecer contraseña" });
  }
});

// ============================================
// GET /admin/audit - Logs de auditoria
// ============================================
router.get("/audit", async (req: Request, res: Response) => {
  try {
    const page = Math.max(1, parseInt(req.query.page as string) || 1);
    const limit = Math.min(100, Math.max(1, parseInt(req.query.limit as string) || 20));
    const targetUserId = req.query.userId as string | undefined;
    const action = req.query.action as string | undefined;

    const result = await getAuditLogs({ page, limit, targetUserId, action });

    res.json({
      ok: true,
      logs: result.logs,
      pagination: {
        page,
        limit,
        total: result.total,
        totalPages: Math.ceil(result.total / limit),
      },
    });
  } catch (err) {
    console.error("[Admin] Error obteniendo logs de auditoria:", err);
    res.status(500).json({ ok: false, message: "Error interno al obtener logs" });
  }
});

// ============================================
// GET /admin/portal-status - Configuracion de mantenimiento del portal
// ============================================
router.get("/portal-status", async (_req: Request, res: Response) => {
  try {
    const status = await getPortalStatusConfig();
    res.json({ ok: true, status });
  } catch (err) {
    console.error("[Admin] Error obteniendo portal status:", err);
    res.status(500).json({ ok: false, message: "Error obteniendo configuracion del portal" });
  }
});

// ============================================
// PUT /admin/portal-status - Actualizar configuracion del portal
// ============================================
router.put("/portal-status", async (req: Request, res: Response) => {
  try {
    const actorId = req.user!.userId;
    const current = await getPortalStatusConfig();
    const payload = req.body?.status as PortalStatusConfig | undefined;

    if (!payload || typeof payload !== "object") {
      return res.status(400).json({ ok: false, message: "Payload de configuracion invalido" });
    }

    const updated = await updatePortalStatusConfig(payload, actorId);

    await logAdminAction({
      actorId,
      action: "update_portal_status",
      before: current as unknown as Record<string, unknown>,
      after: updated as unknown as Record<string, unknown>,
    });

    res.json({ ok: true, status: updated, message: "Configuracion del portal actualizada" });
  } catch (err) {
    console.error("[Admin] Error actualizando portal status:", err);
    res.status(500).json({ ok: false, message: "Error actualizando configuracion del portal" });
  }
});

function summarizeKanban(board: AdminKanbanBoard): Record<string, unknown> {
  return {
    statuses: board.statuses?.length || 0,
    tasks: board.tasks?.length || 0,
    done: (board.tasks || []).filter((task) => task.statusId === "done").length,
  };
}

// ============================================
// GET /admin/kanban - Tablero administrativo
// ============================================
router.get("/kanban", async (_req: Request, res: Response) => {
  try {
    const [board, admins] = await Promise.all([
      getAdminKanbanBoard(),
      listAdminUsersForKanban(),
    ]);

    res.json({ ok: true, board, admins });
  } catch (err) {
    console.error("[Admin] Error obteniendo kanban:", err);
    res.status(500).json({ ok: false, message: "Error obteniendo tablero Kanban" });
  }
});

// ============================================
// PUT /admin/kanban - Guardar tablero administrativo
// ============================================
router.put("/kanban", async (req: Request, res: Response) => {
  try {
    const actorId = req.user!.userId;
    const payload = req.body?.board as Partial<AdminKanbanBoard> | undefined;

    if (!payload || typeof payload !== "object") {
      return res.status(400).json({ ok: false, message: "Payload de tablero invalido" });
    }

    const current = await getAdminKanbanBoard();
    const board = await updateAdminKanbanBoard(payload, actorId);

    await logAdminAction({
      actorId,
      action: "update_admin_kanban",
      before: summarizeKanban(current),
      after: summarizeKanban(board),
    });

    res.json({ ok: true, board, message: "Tablero Kanban actualizado" });
  } catch (err) {
    console.error("[Admin] Error actualizando kanban:", err);
    res.status(500).json({ ok: false, message: "Error guardando tablero Kanban" });
  }
});


// ============================================
// GET /admin/accounting-kanban - Tablero contable
// ============================================
router.get("/accounting-kanban", async (_req: Request, res: Response) => {
  try {
    const [board, admins] = await Promise.all([
      getAccountingKanbanBoard(),
      listAdminUsersForKanban(),
    ]);

    res.json({ ok: true, board, admins });
  } catch (err) {
    console.error("[Admin] Error obteniendo kanban contable:", err);
    res.status(500).json({ ok: false, message: "Error obteniendo tablero Kanban contable" });
  }
});

// ============================================
// PUT /admin/accounting-kanban - Guardar tablero contable
// ============================================
router.put("/accounting-kanban", async (req: Request, res: Response) => {
  try {
    const actorId = req.user!.userId;
    const payload = req.body?.board as Partial<AdminKanbanBoard> | undefined;

    if (!payload || typeof payload !== "object") {
      return res.status(400).json({ ok: false, message: "Payload de tablero invalido" });
    }

    const current = await getAccountingKanbanBoard();
    const board = await updateAccountingKanbanBoard(payload, actorId);

    await logAdminAction({
      actorId,
      action: "update_accounting_kanban",
      before: summarizeKanban(current),
      after: summarizeKanban(board),
    });

    res.json({ ok: true, board, message: "Tablero Kanban contable actualizado" });
  } catch (err) {
    console.error("[Admin] Error actualizando kanban contable:", err);
    res.status(500).json({ ok: false, message: "Error guardando tablero Kanban contable" });
  }
});
export default router;
