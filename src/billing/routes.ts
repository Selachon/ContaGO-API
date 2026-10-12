import { Router, Request, Response } from "express";
import { requireAuth } from "../middleware/auth.js";
import { requireToolAccess } from "../middleware/requireToolAccess.js";
import { rateLimit } from "../middleware/rateLimit.js";
import { AlanubeError, getDianPlaces } from "../services/alanubeService.js";
import {
  BILLING_TOOL_ID, REGIME_CODES, TAX_CODES, createCompany, getCompany, listCompanies, lookupNit, normalizeResolution,
  enableCompanySupportDocs, refreshCompanyEnablement, saveCompanyResolution, startCompanyEnablement, sandboxEnableCompany, isSandbox, updateCompany,
} from "./companies.js";
import { assertTokenMatchesNit } from "./dianToken.js";
import { CUSTOMER_DOC_TYPES, TAX_PERCENTS, deleteCustomer, deleteProduct, listCustomers, listProducts, saveCustomer, saveProduct } from "./catalog.js";
import { emitSupportDocument, getSupportDocFile, saveSupportDocDraft } from "./supportDocuments.js";
import { lookupSupplier } from "../services/supportDocsService.js";
import { creditNoteContext, getDoc, deleteDraft, emitCreditNote, emitInvoice, getDocFile, listDocs, saveCreditNoteDraft, saveInvoiceDraft } from "./invoices.js";
import { DianRangesError, fetchDianTestRanges } from "./dianRanges.js";

const router = Router();
router.use(requireAuth);
router.use(requireToolAccess(BILLING_TOOL_ID));

const actor = (req: Request) => ({ userId: req.user!.userId, isAdmin: !!req.user!.isAdmin });

/** Mensajes de la habilitación de la DIAN/Alanube traducidos para el cliente. */
function friendly(err: any): string {
  const m = String(err?.message || "Error");
  if (/COMPANY_DIAN_MISMATCH/.test(m)) return "El enlace de la DIAN no corresponde al NIT de esta empresa. Genera el token con el NIT y la cédula del representante legal de esta empresa.";
  if (/INVALID_DIAN_TOKEN_FORMAT/.test(m)) return "El enlace de la DIAN no tiene el formato esperado. Cópialo completo desde el correo.";
  if (/CONCURRENT_ENABLEMENT|ENABLEMENT_SCOPE_CONFLICT/.test(m)) return "Ya hay una habilitación en curso para esta empresa. Espera a que termine.";
  return m;
}
const fail = (res: Response, err: any, status = 400) => res.status(status).json({ ok: false, message: friendly(err) });

// Catálogos del formulario de empresa
router.get("/catalogs", async (_req, res) => {
  try {
    const p = await getDianPlaces();
    res.json({ ok: true, departments: p.departments, municipalities: p.municipalities, countries: p.countries, regimeCodes: REGIME_CODES, taxCodes: TAX_CODES });
  } catch (err) { fail(res, err, 502); }
});

// Nombre y correo registrados en la DIAN para un NIT
router.get("/nit", rateLimit(30, 10 * 60_000), async (req, res) => {
  try { res.json({ ok: true, ...(await lookupNit(String(req.query.nit || ""))) }); } catch (err) { fail(res, err); }
});

router.get("/empresas", async (req, res) => {
  try { res.json({ ok: true, companies: await listCompanies(actor(req)) }); } catch (err) { fail(res, err, 500); }
});

router.post("/empresas", rateLimit(10, 60 * 60_000), async (req, res) => {
  try { res.json({ ok: true, company: await createCompany(actor(req), req.body) }); } catch (err) { fail(res, err); }
});

router.get("/empresas/:id", async (req, res) => {
  try { res.json({ ok: true, company: await refreshCompanyEnablement(actor(req), req.params.id) }); } catch (err) { fail(res, err, 404); }
});

router.patch("/empresas/:id", async (req, res) => {
  try { res.json({ ok: true, company: await updateCompany(actor(req), req.params.id, req.body) }); } catch (err) { fail(res, err); }
});

// Habilitar facturas en la DIAN con el token del cliente (el token no se guarda)
router.post("/empresas/:id/habilitacion", rateLimit(10, 10 * 60_000), async (req, res) => {
  try { res.json({ ok: true, ...(await startCompanyEnablement(actor(req), req.params.id, String(req.body?.tokenDIAN || ""))) }); } catch (err) { fail(res, err); }
});

// Leer en el portal de la DIAN los rangos de numeración de prueba (abre un navegador: límite estricto)
router.post("/empresas/:id/rangos-prueba", rateLimit(5, 10 * 60_000), async (req, res) => {
  try {
    const c = await getCompany(actor(req), req.params.id);
    const tokenUrl = String(req.body?.tokenDIAN || "");
    assertTokenMatchesNit(tokenUrl, c.nit);
    res.json({ ok: true, ranges: await fetchDianTestRanges(tokenUrl, c.nit) });
  } catch (err: any) {
    console.error("[Facturación] rangos de prueba:", err instanceof DianRangesError ? err.code : err?.message);
    fail(res, err);
  }
});

// Guardar la resolución de facturas (la leída de la DIAN y confirmada, o escrita a mano)
router.post("/empresas/:id/resolucion", async (req, res) => {
  try {
    normalizeResolution(req.body);
    res.json({ ok: true, company: await saveCompanyResolution(actor(req), req.params.id, req.body) });
  } catch (err) { fail(res, err); }
});

// ─── Clientes, productos y facturas de una empresa ────────────────────────────────────────────────────────────

router.get("/opciones", (_req, res) => res.json({ ok: true, customerDocTypes: CUSTOMER_DOC_TYPES, taxPercents: TAX_PERCENTS, sandbox: isSandbox() }));

router.get("/empresas/:id/clientes", async (req, res) => { try { res.json({ ok: true, customers: await listCustomers(actor(req), req.params.id) }); } catch (err) { fail(res, err); } });
router.post("/empresas/:id/clientes", async (req, res) => { try { res.json({ ok: true, customer: await saveCustomer(actor(req), req.params.id, req.body) }); } catch (err) { fail(res, err); } });
router.put("/empresas/:id/clientes/:cid", async (req, res) => { try { res.json({ ok: true, customer: await saveCustomer(actor(req), req.params.id, req.body, req.params.cid) }); } catch (err) { fail(res, err); } });
router.delete("/empresas/:id/clientes/:cid", async (req, res) => { try { await deleteCustomer(actor(req), req.params.id, req.params.cid); res.json({ ok: true }); } catch (err) { fail(res, err); } });

router.get("/empresas/:id/productos", async (req, res) => { try { res.json({ ok: true, products: await listProducts(actor(req), req.params.id) }); } catch (err) { fail(res, err); } });
router.post("/empresas/:id/productos", async (req, res) => { try { res.json({ ok: true, product: await saveProduct(actor(req), req.params.id, req.body) }); } catch (err) { fail(res, err); } });
router.put("/empresas/:id/productos/:pid", async (req, res) => { try { res.json({ ok: true, product: await saveProduct(actor(req), req.params.id, req.body, req.params.pid) }); } catch (err) { fail(res, err); } });
router.delete("/empresas/:id/productos/:pid", async (req, res) => { try { await deleteProduct(actor(req), req.params.id, req.params.pid); res.json({ ok: true }); } catch (err) { fail(res, err); } });

router.get("/empresas/:id/documentos", async (req, res) => { try { res.json({ ok: true, documents: await listDocs(actor(req), req.params.id) }); } catch (err) { fail(res, err); } });
router.post("/empresas/:id/facturas", async (req, res) => { try { res.json({ ok: true, ...(await saveInvoiceDraft(actor(req), req.params.id, req.body)) }); } catch (err) { fail(res, err); } });
router.put("/empresas/:id/facturas/:did", async (req, res) => { try { res.json({ ok: true, ...(await saveInvoiceDraft(actor(req), req.params.id, req.body, req.params.did)) }); } catch (err) { fail(res, err); } });
router.delete("/empresas/:id/documentos/:did", async (req, res) => { try { await deleteDraft(actor(req), req.params.id, req.params.did); res.json({ ok: true }); } catch (err) { fail(res, err); } });
router.post("/empresas/:id/facturas/:did/emitir", rateLimit(30, 10 * 60_000), async (req, res) => {
  try { res.json({ ok: true, emission: await emitInvoice(actor(req), req.params.id, req.params.did) }); }
  catch (err) { console.error("[Facturación] emitir:", (err as Error)?.message); fail(res, err); }
});
router.get("/empresas/:id/facturas/:did/nota-credito/contexto", async (req, res) => { try { res.json({ ok: true, ...(await creditNoteContext(actor(req), req.params.id, req.params.did)) }); } catch (err) { fail(res, err); } });
router.post("/empresas/:id/notas-credito", async (req, res) => { try { res.json({ ok: true, ...(await saveCreditNoteDraft(actor(req), req.params.id, req.body)) }); } catch (err) { fail(res, err); } });
router.put("/empresas/:id/notas-credito/:did", async (req, res) => { try { res.json({ ok: true, ...(await saveCreditNoteDraft(actor(req), req.params.id, req.body, req.params.did)) }); } catch (err) { fail(res, err); } });
router.post("/empresas/:id/notas-credito/:did/emitir", rateLimit(30, 10 * 60_000), async (req, res) => {
  try { res.json({ ok: true, emission: await emitCreditNote(actor(req), req.params.id, req.params.did) }); }
  catch (err) { console.error("[Facturación] emitir nota crédito:", (err as Error)?.message); fail(res, err); }
});
router.post("/empresas/:id/habilitacion-pruebas", rateLimit(20, 10 * 60_000), async (req, res) => { try { res.json({ ok: true, company: await sandboxEnableCompany(actor(req), req.params.id) }); } catch (err) { fail(res, err); } });
router.post("/empresas/:id/resolucion-ds", async (req, res) => { try { res.json({ ok: true, company: await enableCompanySupportDocs(actor(req), req.params.id, req.body) }); } catch (err) { fail(res, err); } });
router.get("/empresas/:id/proveedor", rateLimit(30, 10 * 60_000), async (req, res) => { try { const found = await lookupSupplier(String(req.query.docType || ""), String(req.query.id || "")); res.json({ ok: true, found: !!found, ...(found || {}) }); } catch (err) { fail(res, err); } });
router.post("/empresas/:id/documentos-soporte", async (req, res) => { try { res.json({ ok: true, ...(await saveSupportDocDraft(actor(req), req.params.id, req.body)) }); } catch (err) { fail(res, err); } });
router.put("/empresas/:id/documentos-soporte/:did", async (req, res) => { try { res.json({ ok: true, ...(await saveSupportDocDraft(actor(req), req.params.id, req.body, req.params.did)) }); } catch (err) { fail(res, err); } });
router.post("/empresas/:id/documentos-soporte/:did/emitir", rateLimit(30, 10 * 60_000), async (req, res) => {
  try { res.json({ ok: true, emission: await emitSupportDocument(actor(req), req.params.id, req.params.did) }); }
  catch (err) { console.error("[Facturación] emitir documento soporte:", (err as Error)?.message); fail(res, err); }
});
router.get("/empresas/:id/documentos/:did/archivo/:kind", async (req, res) => {
  try {
    const kind = (["zip", "pdf", "xml"] as const).find(k => k === req.params.kind);
    if (!kind) return res.status(400).json({ ok: false, message: "Solicitud inválida" });
    const isDs = (await getDoc(actor(req), req.params.id, req.params.did)).type === "supportDocument";
    const f = isDs ? await getSupportDocFile(actor(req), req.params.id, req.params.did, kind) : await getDocFile(actor(req), req.params.id, req.params.did, kind);
    if (!f) return res.status(404).json({ ok: false, message: "No hay archivo oficial para este documento." });
    res.setHeader("Content-Type", f.contentType);
    res.setHeader("Content-Disposition", `${kind === "pdf" ? "inline" : "attachment"}; filename="${f.filename}"`);
    res.send(f.data);
  } catch (err) { fail(res, err, 404); }
});

export default router;
