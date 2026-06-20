/**
 * Conciliación bancaria: cruza el extracto del banco (PDF) contra el libro
 * auxiliar de la contabilidad (Excel de SIIGO) y arma el cuadre.
 *
 *  POST /conciliacion/analyze    multipart: extracto (PDF) + auxiliar (Excel)
 *                                → parsea ambos y devuelve la conciliación.
 *  POST /conciliacion/reconcile  JSON: { statement, ledger, options }
 *                                → re-corre el cruce (tolerancia / overrides) sin re-subir.
 *  POST /conciliacion/export     JSON: { statement, ledger, options, meta }
 *                                → genera el reporte Excel.
 */
import { Router, type Request, type Response } from "express";
import multer from "multer";
import { requireAuth } from "../middleware/auth.js";
import { requireToolAccess } from "../middleware/requireToolAccess.js";
import { parseBankPdf, StatementError } from "../services/bankStatementParser.js";
import { parseAuxiliaryLedger, LedgerError } from "../services/auxiliaryLedgerParser.js";
import {
  reconcile,
  type RecStatementInput,
  type RecLedgerInput,
  type RecOptions,
} from "../services/bankReconciliationService.js";
import { generateReconciliationExcel, type GenMeta } from "../services/conciliacionExcelGenerator.js";

export const CONCILIACION_TOOL_ID = "conciliacion-bancaria";

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 25 * 1024 * 1024, files: 2 },
});

const router = Router();

router.use(requireAuth);
router.use(requireToolAccess(CONCILIACION_TOOL_ID));

const toNum = (v: unknown, def: number): number => {
  const n = Number(v);
  return Number.isFinite(n) ? n : def;
};

function isPdf(file: Express.Multer.File): boolean {
  return file.mimetype === "application/pdf" || (file.originalname || "").toLowerCase().endsWith(".pdf");
}

/** Normaliza el extracto (PDF) a la forma que consume el motor. */
async function parseStatement(file: Express.Multer.File, password?: string) {
  if (!isPdf(file)) {
    throw new StatementError(
      "Por ahora el extracto debe subirse en PDF (Bancolombia o Itaú). El soporte de extracto en Excel llegará pronto.",
      422,
      "extracto_excel_pendiente"
    );
  }
  const parsed = await parseBankPdf(file.buffer, password);
  const statement: RecStatementInput = {
    opening: parsed.reconciliation.openingBalance,
    closing: parsed.reconciliation.closingBalance,
    movements: parsed.movements.map((m) => ({
      date: m.date,
      description: m.description,
      value: m.value,
      direction: m.direction,
      kind: m.kind,
    })),
  };
  return { statement, parsed };
}

// ─── POST /analyze: sube extracto + auxiliar y devuelve la conciliación ────
router.post(
  "/analyze",
  upload.fields([
    { name: "extracto", maxCount: 1 },
    { name: "auxiliar", maxCount: 1 },
  ]),
  async (req: Request, res: Response) => {
    try {
      const files = req.files as Record<string, Express.Multer.File[]> | undefined;
      const extractoFile = files?.extracto?.[0];
      const auxiliarFile = files?.auxiliar?.[0];
      if (!extractoFile) return res.status(400).json({ ok: false, message: "Falta el archivo del extracto." });
      if (!auxiliarFile) return res.status(400).json({ ok: false, message: "Falta el archivo del auxiliar contable." });

      const password = typeof req.body?.password === "string" ? req.body.password : undefined;
      const tolerance = toNum(req.body?.tolerance, 100);

      const { statement, parsed } = await parseStatement(extractoFile, password);
      const aux = await parseAuxiliaryLedger(auxiliarFile.buffer);
      const ledger: RecLedgerInput = {
        opening: aux.opening,
        closing: aux.closing,
        entries: aux.entries.map((e) => ({
          date: e.date,
          description: e.description,
          value: e.value,
          direction: e.direction,
          voucher: e.voucher,
          nit: e.nit,
          thirdParty: e.thirdParty,
        })),
      };

      const result = reconcile(statement, ledger, { tolerance });

      return res.json({
        ok: true,
        meta: {
          bank: parsed.bank,
          accountCode: aux.accountCode,
          accountName: aux.accountName,
          period: aux.period,
        },
        extracto: { reconciliation: parsed.reconciliation },
        auxiliar: { check: aux.check, totals: aux.totals },
        statement,
        ledger,
        result,
      });
    } catch (error) {
      return handleErr(res, error);
    }
  }
);

// ─── POST /reconcile: re-corre el cruce con nueva tolerancia / overrides ───
router.post("/reconcile", async (req: Request, res: Response) => {
  try {
    const { statement, ledger, options } = req.body as {
      statement?: RecStatementInput;
      ledger?: RecLedgerInput;
      options?: RecOptions;
    };
    if (!statement?.movements || !ledger?.entries) {
      return res.status(400).json({ ok: false, message: "Faltan los datos de extracto/auxiliar." });
    }
    const result = reconcile(statement, ledger, options ?? {});
    return res.json({ ok: true, result });
  } catch (error) {
    return handleErr(res, error);
  }
});

// ─── POST /export: genera el reporte Excel ────────────────────────────────
router.post("/export", async (req: Request, res: Response) => {
  try {
    const { statement, ledger, options, meta } = req.body as {
      statement?: RecStatementInput;
      ledger?: RecLedgerInput;
      options?: RecOptions;
      meta?: GenMeta;
    };
    if (!statement?.movements || !ledger?.entries) {
      return res.status(400).json({ ok: false, message: "Faltan los datos de extracto/auxiliar." });
    }
    const result = reconcile(statement, ledger, options ?? {});
    const buffer = await generateReconciliationExcel(
      result,
      meta ?? {},
      statement.movements,
      ledger.entries,
    );
    const fname = `Conciliacion_${(meta?.accountName || "banco").replace(/[^\w]+/g, "_")}.xlsx`;
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${fname}"`);
    return res.send(buffer);
  } catch (error) {
    return handleErr(res, error);
  }
});

function handleErr(res: Response, error: unknown) {
  if (error instanceof StatementError || error instanceof LedgerError) {
    return res.status(error.status).json({ ok: false, code: error.code, message: error.message });
  }
  return res.status(400).json({
    ok: false,
    message: error instanceof Error ? error.message : "Error procesando la conciliación.",
  });
}

export default router;
