/**
 * Pruebas del dominio del ERP (sin base de datos ni red).
 * Ejecutar: npm run test:erp
 */
import test from "node:test";
import assert from "node:assert/strict";
import { createHmac } from "node:crypto";

import { round2, sum, subtract, percentage, toCents, formatCOP, gte, parseAmount } from "./money.js";
import { assertTransition, availableTransitions, canTransition, TRANSITIONS, WorkflowError } from "./workflow.js";
import { permissionsForRole, ROLE_PERMISSIONS } from "./permissions.js";
import { buildDocument, netPayable, recalcTotals, documentLabel, buildDocumentFilter } from "./documents.js";
import { agingBucket, splitSiigoName, maskAccountNumber } from "./payments.js";
import { buildPurchasePayload, strategyFor, buildPostingKey, translateSiigoError } from "./posting.js";
import { recalcRetentions, retentionsTotal } from "./rules.js";
import { normalizeSlug, validateSlug, DEFAULT_SETTINGS, DEFAULT_FEATURES, type ErpTenant } from "./tenants.js";
import { computeCheckDigit, normalizeNit } from "./suppliers.js";
import { redact } from "./audit.js";
import { upsertDoc } from "./db.js";
import { encodeState, decodeState } from "./oauthState.js";
import { createJob, failJob, finishJob, getJob, updateProgress } from "./jobs.js";
import { slugFromHost } from "./context.js";
import { SiigoError } from "../services/siigoService.js";
import type { ErpDocument } from "./types.js";

// ─── Dinero ──────────────────────────────────────────────────────────────

test("money: suma sin errores de coma flotante", () => {
  assert.equal(sum([0.1, 0.2]), 0.3);
  assert.equal(sum([1_234_567.89, 0.11]), 1_234_568);
  assert.equal(round2(2.675), 2.68);
});

test("money: resta y comparación por centavos", () => {
  assert.equal(subtract(10_000_000, 4_000_000), 6_000_000);
  assert.equal(subtract(0.3, 0.1), 0.2);
  assert.equal(gte(1000.005, 1000.01), true);
  assert.equal(gte(999.99, 1000), false);
});

test("money: retención sobre base gravable", () => {
  assert.equal(percentage(4_200_000, 3.5), 147_000);
  assert.equal(percentage(1_000_000, 2.5), 25_000);
  assert.equal(toCents("1.250.000,50"), 125000050);
  assert.equal(toCents("1,250,000.50"), 125000050);
  assert.equal(toCents("$ 1.250.000"), 125000000);
  assert.equal(parseAmount("1.250.000"), 1_250_000);
  assert.equal(parseAmount("2500,75"), 2500.75);
});

test("money: formato colombiano", () => {
  assert.match(formatCOP(1_250_000), /1\.250\.000/);
});

// ─── Máquina de estados ──────────────────────────────────────────────────

test("workflow: solo permite transiciones declaradas", () => {
  assert.equal(canTransition("APPROVE", "PENDING_APPROVAL"), true);
  assert.equal(canTransition("APPROVE", "PAID"), false);
  assert.equal(canTransition("START_POSTING", "ACCOUNTING_READY"), true);
  assert.equal(canTransition("START_POSTING", "RECEIVED"), false);
});

test("workflow: rechaza transición inválida con mensaje de negocio", () => {
  assert.throws(
    () => assertTransition("SCHEDULE_PAYMENT", "PENDING_APPROVAL"),
    (err: unknown) => err instanceof WorkflowError && /Programar pago/.test((err as Error).message)
  );
});

test("workflow: las acciones del sistema no se ofrecen al usuario", () => {
  const actions = availableTransitions("POSTING_TO_SIIGO", Object.values(TRANSITIONS).map((t) => t.permission));
  assert.equal(actions.length, 0);
});

test("workflow: las acciones dependen de los permisos del usuario", () => {
  const approver = availableTransitions("PENDING_APPROVAL", ROLE_PERMISSIONS.APPROVER);
  assert.ok(approver.some((a) => a.name === "APPROVE"));
  const viewer = availableTransitions("PENDING_APPROVAL", ROLE_PERMISSIONS.VIEWER);
  assert.equal(viewer.length, 0);
});

// ─── Permisos ────────────────────────────────────────────────────────────

test("permisos: tesorería no puede causar en Siigo", () => {
  const treasury = permissionsForRole("TREASURY");
  assert.ok(treasury.includes("payables.pay"));
  assert.ok(!treasury.includes("accounting.post_siigo"));
});

test("permisos: quien aprueba no parametriza contabilidad", () => {
  const approver = permissionsForRole("APPROVER");
  assert.ok(approver.includes("documents.approve"));
  assert.ok(!approver.includes("accounting.edit"));
});

test("permisos: extras se acumulan sin duplicar", () => {
  const perms = permissionsForRole("VIEWER", ["documents.approve", "documents.read"]);
  assert.equal(perms.filter((p) => p === "documents.read").length, 1);
  assert.ok(perms.includes("documents.approve"));
});

// ─── Documentos ──────────────────────────────────────────────────────────

function demoDoc(overrides: Partial<ErpDocument> = {}): ErpDocument {
  const base = buildDocument("t1", {
    source: "DIAN",
    type: "ELECTRONIC_INVOICE",
    supplierNit: "900.123.456",
    supplierName: "Proveedor Demo",
    documentPrefix: "SETP",
    documentNumber: "481",
    cufe: "abc123",
    issueDate: "2026-08-01",
    dueDate: "2026-08-31",
    items: [
      { description: "Servicio", quantity: 1, unitPrice: 1_000_000, subtotal: 1_000_000, taxAmount: 190_000 },
      { description: "Soporte", quantity: 2, unitPrice: 500_000, subtotal: 1_000_000, taxAmount: 190_000 },
    ],
  });
  return { id: "d1", ...base, ...overrides } as ErpDocument;
}

test("documentos: normaliza NIT y calcula totales desde los ítems", () => {
  const doc = demoDoc();
  assert.equal(doc.supplierNit, "900123456");
  assert.equal(doc.subtotal, 2_000_000);
  assert.equal(doc.taxes, 380_000);
  assert.equal(doc.total, 2_380_000);
  assert.equal(doc.balance, 2_380_000);
  assert.equal(doc.version, 1);
});

test("documentos: el neto a pagar descuenta retenciones", () => {
  const doc = demoDoc({ retentionsTotal: 70_000 });
  assert.equal(netPayable(doc), 2_310_000);
});

test("documentos: recalcTotals suma retenciones parametrizadas", () => {
  const doc = demoDoc();
  doc.accounting.retentions = recalcRetentions(
    [{ id: "r1", kind: "RETEFUENTE", name: "Retefuente 3.5%", rate: 3.5, base: 2_000_000, value: 0 }],
    2_000_000
  );
  const totals = recalcTotals(doc);
  assert.equal(totals.retentionsTotal, 70_000);
  assert.equal(retentionsTotal(doc.accounting.retentions), 70_000);
});

test("documentos: etiqueta legible del documento", () => {
  assert.equal(documentLabel(demoDoc()), "SETP-481");
  assert.equal(documentLabel({ documentPrefix: undefined, documentNumber: undefined, type: "OTHER" } as any), "s/n");
});

test("documentos: todo filtro queda acotado al tenant", () => {
  const filter = buildDocumentFilter("t1", { status: ["PAYABLE"], search: "energía" });
  assert.equal((filter as any).tenantId, "t1");
  assert.deepEqual((filter as any).status, { $in: ["PAYABLE"] });
  assert.ok(Array.isArray((filter as any).$or));
});

test("documentos: construir un filtro sin tenant es un error", () => {
  assert.throws(() => buildDocumentFilter("", {}), /aislamiento multi-tenant/);
});

// ─── Cuentas por pagar ───────────────────────────────────────────────────

test("cxp: clasificación por antigüedad", () => {
  assert.equal(agingBucket("2026-09-30", "2026-08-21"), "NOT_DUE");
  assert.equal(agingBucket("2026-08-01", "2026-08-21"), "D_1_30");
  assert.equal(agingBucket("2026-07-01", "2026-08-21"), "D_31_60");
  assert.equal(agingBucket("2026-06-01", "2026-08-21"), "D_61_90");
  assert.equal(agingBucket("2026-01-01", "2026-08-21"), "D_90_PLUS");
});

test("cxp: descompone el comprobante Siigo para el cruce del egreso", () => {
  assert.deepEqual(splitSiigoName("FC-1-482"), { prefix: "FC-1", consecutive: 482 });
  assert.deepEqual(splitSiigoName("RP-1-199"), { prefix: "RP-1", consecutive: 199 });
  assert.deepEqual(splitSiigoName(undefined), {});
});

test("cxp: los números de cuenta se guardan enmascarados", () => {
  assert.equal(maskAccountNumber("12345678904832"), "••••4832");
  assert.equal(maskAccountNumber("483"), "483");
});

// ─── Causación ───────────────────────────────────────────────────────────

const tenant: ErpTenant = {
  id: "t1", slug: "demo", name: "Demo", nit: "901000001",
  timezone: "America/Bogota", currency: "COP", status: "ACTIVE",
  features: DEFAULT_FEATURES, settings: DEFAULT_SETTINGS,
  createdAt: "", updatedAt: "",
};

test("causación: factura electrónica → FC, cuenta de cobro → DS", () => {
  assert.equal(strategyFor(demoDoc(), tenant), "FC");
  assert.equal(strategyFor(demoDoc({ type: "COLLECTION_ACCOUNT" }), tenant), "DS");
  assert.equal(strategyFor(demoDoc({ type: "LOAN" }), tenant), "DS");
});

test("causación: la clave de idempotencia incluye tenant, documento y operación", () => {
  assert.equal(buildPostingKey("t1", "d1"), "t1:d1:ACCOUNTING");
  assert.notEqual(buildPostingKey("t1", "d1"), buildPostingKey("t2", "d1"));
});

test("causación: el payload descuenta las retenciones del valor a pagar", () => {
  const doc = demoDoc();
  doc.accounting = {
    accountCode: "512035",
    siigoDocumentTypeId: 27,
    paymentTypeId: 5401,
    costCenterId: 1,
    retentions: [
      { id: "r1", kind: "RETEFUENTE", name: "Retefuente 3.5%", siigoTaxId: 900, rate: 3.5, base: 2_000_000, value: 70_000 },
    ],
  };

  const payload = buildPurchasePayload(tenant, doc);
  assert.equal(payload.document.id, 27);
  assert.equal(payload.supplier.identification, "900123456");
  assert.equal(payload.items.length, 2);
  assert.equal(payload.items[0].type, "Account");
  assert.equal(payload.items[0].code, "512035");
  assert.equal(payload.items[1].price, 500_000);
  assert.equal(payload.cost_center, 1);
  assert.deepEqual(payload.provider_invoice, { prefix: "SETP", number: "481" });
  assert.deepEqual(payload.retentions, [{ id: 900 }]);
  // 2.380.000 − 70.000 de retención
  assert.equal(payload.payments[0].value, 2_310_000);
  assert.equal(payload.payments[0].due_date, "2026-08-31");
});

test("causación: sin ítems se genera una línea única con el total", () => {
  const doc = demoDoc({ items: [], subtotal: 500_000, taxes: 0, total: 500_000 });
  doc.accounting = { accountCode: "513535", siigoDocumentTypeId: 27, paymentTypeId: 5401, retentions: [] };
  const payload = buildPurchasePayload(tenant, doc);
  assert.equal(payload.items.length, 1);
  assert.equal(payload.items[0].price, 500_000);
  assert.equal(payload.payments[0].value, 500_000);
});

// ─── Traducción de errores ───────────────────────────────────────────────

test("errores: tercero inexistente se traduce con acción concreta", () => {
  const err = new SiigoError("Bad Request", 400, "siigo_error", {
    Errors: [{ Code: "invalid_reference", Message: "supplier with identification 860002964 not found", Params: ["supplier.identification"] }],
  });
  const friendly = translateSiigoError(err);
  assert.equal(friendly.code, "supplier_not_found");
  assert.equal(friendly.action, "CREATE_SUPPLIER");
  assert.match(friendly.message, /860\.002\.964/);
});

test("errores: forma de pago inválida se explica en lenguaje de negocio", () => {
  const err = new SiigoError("Bad Request", 400, "siigo_error", {
    Errors: [{ Code: "invalid_reference", Message: "payment id is not valid", Params: ["payments[0].id"] }],
  });
  const friendly = translateSiigoError(err);
  assert.match(friendly.message, /forma de pago/i);
  assert.equal(friendly.action, "FIX_PARAMETERS");
});

test("errores: un timeout obliga a reconciliar, nunca a reintentar a ciegas", () => {
  const friendly = translateSiigoError(new Error("socket hang up"));
  assert.equal(friendly.code, "timeout");
  assert.equal(friendly.action, "RECONCILE");
});

test("errores: credenciales rechazadas apuntan a Integraciones", () => {
  const friendly = translateSiigoError(new SiigoError("Unauthorized", 401, "siigo_auth"));
  assert.equal(friendly.code, "siigo_auth");
  assert.match(friendly.hint || "", /access key/i);
});

// ─── Tenants ─────────────────────────────────────────────────────────────

test("tenants: normaliza y valida el subdominio", () => {
  assert.equal(normalizeSlug("Empresa ABC S.A.S."), "empresa-abc-s-a-s");
  assert.equal(validateSlug("empresa-abc"), null);
  assert.match(validateSlug("www") || "", /reservado/);
  assert.match(validateSlug("ab") || "", /3 caracteres/);
  assert.match(validateSlug("-mala-") || "", /letras, números y guiones/);
});

test("tenants: el slug se resuelve desde el hostname, no del cliente", () => {
  process.env.ERP_ROOT_DOMAIN = "contago.com.co";
  assert.equal(slugFromHost("empresaabc.contago.com.co"), "empresaabc");
  assert.equal(slugFromHost("contago.com.co"), null);
  assert.equal(slugFromHost("evil.com"), null);
  assert.equal(slugFromHost("a.b.contago.com.co"), null);
});

// ─── Terceros ────────────────────────────────────────────────────────────

test("terceros: dígito de verificación DIAN", () => {
  assert.equal(normalizeNit("890.903.938-8"), "8909039388");
  // NIT reales verificables: Bancolombia 890903938-8, Grupo Éxito 890900608-9.
  assert.equal(computeCheckDigit("890903938"), "8");
  assert.equal(computeCheckDigit("890900608"), "9");
});

// ─── Auditoría ───────────────────────────────────────────────────────────

test("auditoría: nunca registra secretos", () => {
  const clean = redact({
    username: "demo",
    access_key: "super-secreto",
    nested: { refresh_token: "abc", Authorization: "Bearer x", total: 100 },
  }) as any;
  assert.equal(clean.access_key, "[redactado]");
  assert.equal(clean.nested.refresh_token, "[redactado]");
  assert.equal(clean.nested.Authorization, "[redactado]");
  assert.equal(clean.nested.total, 100);
  assert.equal(clean.username, "demo");
});

// ─── Upserts ─────────────────────────────────────────────────────────────

test("upsert: $setOnInsert nunca comparte campos con $set", () => {
  // MongoDB rechaza el update si ambos operadores tocan el mismo campo.
  const update = upsertDoc(
    { name: "Nuevo nombre", active: false },
    { tenantId: "t1", name: "Alta", active: true, createdAt: "2026-08-21" }
  ) as any;

  assert.deepEqual(update.$set, { name: "Nuevo nombre", active: false });
  assert.deepEqual(update.$setOnInsert, { tenantId: "t1", createdAt: "2026-08-21" });
  assert.ok(!("name" in update.$setOnInsert));
  assert.ok(!("active" in update.$setOnInsert));
});

test("upsert: omite $setOnInsert si $set ya cubre todo", () => {
  const update = upsertDoc({ a: 1, b: 2 }, { a: 9, b: 9 }) as any;
  assert.deepEqual(update, { $set: { a: 1, b: 2 } });
});

test("upsert: conserva operadores extra como $inc", () => {
  const update = upsertDoc(
    { name: "X" },
    { tenantId: "t1", documentCount: 0 },
    { $inc: { documentCount: 1 } }
  ) as any;
  assert.deepEqual(update.$inc, { documentCount: 1 });
  assert.equal(update.$setOnInsert.tenantId, "t1");
});

// ─── OAuth state ─────────────────────────────────────────────────────────

test("oauth: el state va firmado y no se puede manipular", () => {
  process.env.JWT_SECRET = "test-secret";
  const state = encodeState({ purpose: "erp-login", slug: "demo" });
  const decoded = decodeState(state);
  assert.equal(decoded?.slug, "demo");
  assert.equal(decoded?.purpose, "erp-login");

  const [payload] = state.split(".");
  assert.equal(decodeState(`${payload}.firmafalsa`), null);
  assert.equal(decodeState("basura"), null);
});

test("oauth: un state caducado se rechaza", () => {
  process.env.JWT_SECRET = "test-secret";
  const stale = Buffer.from(
    JSON.stringify({ purpose: "erp-login", slug: "demo", ts: Date.now() - 20 * 60 * 1000, nonce: "x" })
  ).toString("base64url");
  const sig = createHmac("sha256", "test-secret").update(stale).digest("base64url");
  assert.equal(decodeState(`${stale}.${sig}`), null);
});

// ─── Jobs en segundo plano ────────────────────────────────────────────────

test("jobs: reporta progreso y aísla por tenant", () => {
  const id = createJob("t1");
  const job = getJob("t1", id);
  assert.equal(job?.status, "RUNNING");
  assert.equal(job?.progress.step, "Iniciando…");

  updateProgress(id, "Descargando documentos… (3/10)", 30);
  assert.equal(getJob("t1", id)?.progress.percent, 30);

  // Un tenant nunca puede leer el job de otro, aunque conozca el id.
  assert.equal(getJob("t2", id), null);
});

test("jobs: éxito y error dejan el estado final correcto", () => {
  const okId = createJob("t1");
  finishJob(okId, { created: 3 });
  const ok = getJob("t1", okId);
  assert.equal(ok?.status, "SUCCESS");
  assert.deepEqual(ok?.result, { created: 3 });
  assert.ok(ok?.finishedAt);

  const failId = createJob("t1");
  failJob(failId, "token expirado");
  const failed = getJob("t1", failId);
  assert.equal(failed?.status, "ERROR");
  assert.equal(failed?.error, "token expirado");
});

test("jobs: un job terminado ya no acepta más progreso", () => {
  const id = createJob("t1");
  finishJob(id, { ok: true });
  updateProgress(id, "esto no debería aplicarse", 99);
  assert.equal(getJob("t1", id)?.status, "SUCCESS");
  assert.notEqual(getJob("t1", id)?.progress.step, "esto no debería aplicarse");
});
