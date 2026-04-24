import { Router, Request, Response } from "express";
import { requireAuth } from "../middleware/auth.js";
import { getUserNits, hasPurchase } from "../services/database.js";
import { normalizeNit, upsertDianCertificateConfig } from "../services/dianCertificateStore.js";
import { createDianDocumentsForNit } from "../dian/createDianDocuments.js";
import type { DianEnvironment } from "../dian/types/DianResponse.js";
import { DianError } from "../dian/errors/DianError.js";

const router = Router();
const ALLOWED_TOOL_IDS = ["dian-downloader", "dian-excel-exporter"];

function resolveEnvironment(raw?: string): DianEnvironment {
  const source = raw?.trim() ? raw : process.env.DIAN_ENVIRONMENT ?? "hab";
  const value = source.toLowerCase().trim();
  if (value === "hab" || value === "prod") {
    return value;
  }

  console.warn(
    JSON.stringify({
      module: "dian",
      component: "dianSoapRoute",
      level: "warn",
      action: "invalid_environment_fallback",
      raw,
      fallback: "hab",
      timestamp: new Date().toISOString(),
    })
  );

  return "hab";
}

async function requireAnyDianToolAccess(req: Request, res: Response): Promise<boolean> {
  if (!req.user) {
    res.status(401).json({ ok: false, message: "Token no proporcionado" });
    return false;
  }

  if (req.user.isAdmin) {
    return true;
  }

  for (const toolId of ALLOWED_TOOL_IDS) {
    // eslint-disable-next-line no-await-in-loop
    const allowed = await hasPurchase(req.user.userId, toolId);
    if (allowed) {
      return true;
    }
  }

  res.status(403).json({
    ok: false,
    message: "No tienes acceso a las herramientas DIAN requeridas.",
  });
  return false;
}

async function checkNitAccess(req: Request, res: Response, nit: string): Promise<boolean> {
  if (!req.user) {
    res.status(401).json({ ok: false, message: "Token no proporcionado" });
    return false;
  }
  if (req.user.isAdmin) {
    return true;
  }

  const allowedNits = await getUserNits(req.user.userId);
  const normalizedNit = normalizeNit(nit);
  const allowedNormalized = allowedNits.map(normalizeNit);

  if (!allowedNormalized.includes(normalizedNit)) {
    res.status(403).json({
      ok: false,
      message: `No tienes acceso al NIT ${nit}`,
    });
    return false;
  }

  return true;
}

function handleDianError(res: Response, error: unknown): void {
  if (error instanceof DianError) {
    const status = error.code === "DIAN_VALIDATION_ERROR" ? 400 : 502;
    res.status(status).json({
      ok: false,
      error: error.code,
      message: error.message,
      details: error.details,
    });
    return;
  }

  const message = error instanceof Error ? error.message : "Error desconocido";
  res.status(500).json({ ok: false, message });
}

router.use(requireAuth);

router.post("/certificates", async (req: Request, res: Response) => {
  if (!req.user?.isAdmin) {
    return res.status(403).json({
      ok: false,
      message: "Solo administradores pueden registrar certificados DIAN",
    });
  }

  const nit = String(req.body?.nit ?? "");
  const p12Path = String(req.body?.p12_path ?? req.body?.p12Path ?? "");
  const p12Password = String(req.body?.p12_password ?? req.body?.p12Password ?? "");
  const enabled = req.body?.enabled !== false;

  let environment: DianEnvironment;
  try {
    environment = resolveEnvironment(String(req.body?.environment ?? ""));
  } catch {
    environment = resolveEnvironment();
  }

  if (!nit || !p12Path || !p12Password) {
    return res.status(400).json({
      ok: false,
      message: "nit, p12_path y p12_password son requeridos",
    });
  }

  try {
    const saved = await upsertDianCertificateConfig({
      nit,
      p12Path,
      p12Password,
      environment,
      enabled,
      updatedBy: req.user.userId,
    });

    if (!saved) {
      return res.status(500).json({
        ok: false,
        message: "No fue posible guardar la configuración DIAN",
      });
    }

    res.json({
      ok: true,
      message: "Certificado DIAN actualizado",
      nit: normalizeNit(nit),
      environment,
      enabled,
    });
  } catch (error) {
    handleDianError(res, error);
  }
});

router.post("/status", async (req: Request, res: Response) => {
  const allowed = await requireAnyDianToolAccess(req, res);
  if (!allowed) return;

  const nit = String(req.body?.nit ?? "");
  const trackId = String(req.body?.trackId ?? req.body?.track_id ?? "");
  const environment = resolveEnvironment(String(req.body?.environment ?? ""));

  if (!nit || !trackId) {
    return res.status(400).json({ ok: false, message: "nit y trackId son requeridos" });
  }

  if (!(await checkNitAccess(req, res, nit))) return;

  try {
    const documents = await createDianDocumentsForNit({
      nit,
      environment,
      companyId: req.user?.userId,
    });

    if (!documents) {
      return res.status(404).json({
        ok: false,
        message: `No hay certificado DIAN configurado para NIT ${nit} en ${environment}`,
      });
    }

    const result = await documents.getStatus(trackId);
    res.json({ ok: true, data: result });
  } catch (error) {
    handleDianError(res, error);
  }
});

router.post("/status-zip", async (req: Request, res: Response) => {
  const allowed = await requireAnyDianToolAccess(req, res);
  if (!allowed) return;

  const nit = String(req.body?.nit ?? "");
  const trackId = String(req.body?.trackId ?? req.body?.track_id ?? "");
  const environment = resolveEnvironment(String(req.body?.environment ?? ""));

  if (!nit || !trackId) {
    return res.status(400).json({ ok: false, message: "nit y trackId son requeridos" });
  }

  if (!(await checkNitAccess(req, res, nit))) return;

  try {
    const documents = await createDianDocumentsForNit({
      nit,
      environment,
      companyId: req.user?.userId,
    });

    if (!documents) {
      return res.status(404).json({
        ok: false,
        message: `No hay certificado DIAN configurado para NIT ${nit} en ${environment}`,
      });
    }

    const result = await documents.getStatusZip(trackId);
    res.json({ ok: true, data: result });
  } catch (error) {
    handleDianError(res, error);
  }
});

router.post("/xml-by-key", async (req: Request, res: Response) => {
  const allowed = await requireAnyDianToolAccess(req, res);
  if (!allowed) return;

  const nit = String(req.body?.nit ?? "");
  const trackId = String(req.body?.trackId ?? req.body?.track_id ?? "");
  const environment = resolveEnvironment(String(req.body?.environment ?? ""));

  if (!nit || !trackId) {
    return res.status(400).json({ ok: false, message: "nit y trackId son requeridos" });
  }

  if (!(await checkNitAccess(req, res, nit))) return;

  try {
    const documents = await createDianDocumentsForNit({
      nit,
      environment,
      companyId: req.user?.userId,
    });

    if (!documents) {
      return res.status(404).json({
        ok: false,
        message: `No hay certificado DIAN configurado para NIT ${nit} en ${environment}`,
      });
    }

    const result = await documents.getXmlByDocumentKey(trackId);
    res.json({ ok: true, data: result });
  } catch (error) {
    handleDianError(res, error);
  }
});

router.post("/acquirer", async (req: Request, res: Response) => {
  const allowed = await requireAnyDianToolAccess(req, res);
  if (!allowed) return;

  const nit = String(req.body?.nit ?? "");
  const identificationType = String(req.body?.identificationType ?? req.body?.type ?? "");
  const identificationNumber = String(
    req.body?.identificationNumber ?? req.body?.number ?? ""
  );
  const environment = resolveEnvironment(String(req.body?.environment ?? ""));

  if (!nit || !identificationType || !identificationNumber) {
    return res.status(400).json({
      ok: false,
      message: "nit, identificationType e identificationNumber son requeridos",
    });
  }

  if (!(await checkNitAccess(req, res, nit))) return;

  try {
    const documents = await createDianDocumentsForNit({
      nit,
      environment,
      companyId: req.user?.userId,
    });

    if (!documents) {
      return res.status(404).json({
        ok: false,
        message: `No hay certificado DIAN configurado para NIT ${nit} en ${environment}`,
      });
    }

    const result = await documents.getAcquirer(identificationType, identificationNumber);
    res.json({ ok: true, data: result });
  } catch (error) {
    handleDianError(res, error);
  }
});

router.post("/exchange-emails", async (req: Request, res: Response) => {
  const allowed = await requireAnyDianToolAccess(req, res);
  if (!allowed) return;

  const nit = String(req.body?.nit ?? "");
  const environment = resolveEnvironment(String(req.body?.environment ?? ""));

  if (!nit) {
    return res.status(400).json({ ok: false, message: "nit es requerido" });
  }

  if (!(await checkNitAccess(req, res, nit))) return;

  try {
    const documents = await createDianDocumentsForNit({
      nit,
      environment,
      companyId: req.user?.userId,
    });

    if (!documents) {
      return res.status(404).json({
        ok: false,
        message: `No hay certificado DIAN configurado para NIT ${nit} en ${environment}`,
      });
    }

    const result = await documents.getExchangeEmails(nit);
    res.json({ ok: true, data: result });
  } catch (error) {
    handleDianError(res, error);
  }
});

export default router;
