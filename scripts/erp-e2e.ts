/**
 * Prueba de humo de extremo a extremo del ERP contra la API en ejecución.
 * Uso: npm run test:erp:e2e   (requiere la API levantada y MongoDB accesible)
 */
import "dotenv/config";
import jwt from "jsonwebtoken";
import { connectMongo, getDb } from "../src/services/database.js";

const BASE = process.env.ERP_E2E_BASE || "http://localhost:8000/erp";

async function main() {
  await connectMongo();
  const admin = await getDb().collection("users").findOne({ is_admin: true });
  if (!admin) throw new Error("No hay usuario admin en la base de datos.");
  const token = jwt.sign(
    { userId: String(admin._id), email: admin.email, isAdmin: true },
    process.env.JWT_SECRET!,
    { expiresIn: "1h" }
  );
  console.log("Admin:", admin.email, "\n");

  const call = async (path: string, init: RequestInit = {}) => {
    const res = await fetch(`${BASE}${path}`, {
      ...init,
      headers: {
        Authorization: `Bearer ${token}`,
        "X-Erp-Tenant": "demo",
        "Content-Type": "application/json",
        ...(init.headers || {}),
      },
    });
    const text = await res.text();
    let body: any;
    try { body = JSON.parse(text); } catch { body = text.slice(0, 200); }
    return { status: res.status, body };
  };

  const steps: Array<[string, () => Promise<{ status: number; body: any }>]> = [
    ["seed demo", () => call("/admin/seed-demo", { method: "POST", body: "{}" })],
    ["admin tenants", () => call("/admin/tenants")],
    ["admin health", () => call("/admin/health")],
    ["my-tenants", () => call("/my-tenants")],
    ["portal", () => call("/portal?slug=demo")],
    ["me", () => call("/me")],
    ["dashboard", () => call("/dashboard")],
    ["documents", () => call("/documents?pageSize=5")],
    ["documents search", () => call("/documents/search?q=demo")],
    ["suppliers", () => call("/documents/suppliers")],
    ["accounting dashboard", () => call("/accounting/dashboard")],
    ["accounting queue", () => call("/accounting/queue?scope=pending")],
    ["accounting errors", () => call("/accounting/queue?scope=errors")],
    ["accounting rules", () => call("/accounting/rules")],
    ["payables", () => call("/treasury/payables")],
    ["treasury dashboard", () => call("/treasury/dashboard")],
    ["schedules", () => call("/treasury/schedules")],
    ["payments", () => call("/treasury/payments")],
    ["bank accounts", () => call("/treasury/bank-accounts")],
    ["cashflow", () => call("/treasury/cashflow")],
    ["cost centers", () => call("/cost-centers")],
    ["members", () => call("/members")],
    ["roles", () => call("/roles")],
    ["integrations", () => call("/integrations")],
    ["notifications", () => call("/notifications")],
    ["audit", () => call("/audit")],
    ["reports list", () => call("/reports")],
    ["report payables", () => call("/reports/payables")],
    ["report cashflow", () => call("/reports/cashflow")],
  ];

  let failures = 0;
  for (const [name, fn] of steps) {
    try {
      const { status, body } = await fn();
      const ok = status >= 200 && status < 300;
      if (!ok) failures += 1;
      const data = body?.data;
      const summary = Array.isArray(data)
        ? `${data.length} items`
        : data?.items ? `${data.items.length}/${data.total} items`
        : data?.rows ? `${data.rows.length} filas`
        : body?.message || (data ? "ok" : "");
      console.log(`${ok ? "✔" : "✖"} ${name.padEnd(24)} ${status}  ${String(summary).slice(0, 90)}`);
    } catch (err) {
      failures += 1;
      console.log(`✖ ${name.padEnd(24)} ERROR ${err instanceof Error ? err.message : err}`);
    }
  }

  // ── Flujo completo sobre un documento ──────────────────────────────────
  const docs = await call("/documents?status=PENDING_APPROVAL&pageSize=1");
  const doc = docs.body?.data?.items?.[0];
  if (doc) {
    console.log(`\n— Flujo sobre ${doc.documentPrefix}-${doc.documentNumber} (${doc.supplierName}) —`);
    const approve = await call(`/documents/${doc.id}/approve`, { method: "POST", body: JSON.stringify({ version: doc.version }) });
    console.log(`${approve.status < 300 ? "✔" : "✖"} aprobar              ${approve.status} → ${approve.body?.data?.status || approve.body?.message}`);

    const suggestion = await call(`/accounting/documents/${doc.id}/suggestion`);
    console.log(`${suggestion.status < 300 ? "✔" : "✖"} sugerencia           ${suggestion.status} → ${suggestion.body?.data?.explanation?.[0] || ""}`);

    const applied = await call(`/accounting/documents/${doc.id}/suggestion/apply`, { method: "POST", body: "{}" });
    console.log(`${applied.status < 300 ? "✔" : "✖"} aplicar sugerencia   ${applied.status} → cuenta ${applied.body?.data?.accounting?.accountCode || "—"}`);

    const validation = await call(`/accounting/documents/${doc.id}/validate`);
    const issues = validation.body?.data?.issues || [];
    console.log(`${validation.status < 300 ? "✔" : "✖"} validar              ${validation.status} → estrategia ${validation.body?.data?.strategy}, ${validation.body?.data?.valid ? "válido" : `${issues.length} hallazgo(s)`}`);
    for (const i of issues) console.log(`     · [${i.severity}] ${i.message}`);
  }

  // ── Cuenta de cobro creada por formulario ──────────────────────────────
  const created = await call("/documents", {
    method: "POST",
    body: JSON.stringify({
      source: "FORM",
      type: "COLLECTION_ACCOUNT",
      supplierNit: "1020304050",
      supplierName: "María Fernanda Ruiz (Cuenta de cobro)",
      issueDate: new Date().toISOString().slice(0, 10),
      description: "Prueba de humo — honorarios",
      items: [{ description: "Honorarios", quantity: 1, unitPrice: 2_000_000, subtotal: 2_000_000, taxAmount: 0, total: 2_000_000 }],
      total: 2_000_000,
    }),
  });
  console.log(`\n${created.status === 201 ? "✔" : "✖"} radicar cuenta de cobro (${created.status}) → ${created.body?.data?.status || created.body?.message}`);

  // Duplicado: el mismo documento no debe crearse dos veces.
  const dupe = await call("/documents", {
    method: "POST",
    body: JSON.stringify({
      source: "DIAN", type: "ELECTRONIC_INVOICE",
      supplierNit: doc?.supplierNit || "900123456",
      supplierName: doc?.supplierName || "X",
      documentPrefix: doc?.documentPrefix, documentNumber: doc?.documentNumber,
      issueDate: doc?.issueDate || new Date().toISOString().slice(0, 10),
      total: 1000,
    }),
  });
  console.log(`${dupe.status === 409 ? "✔" : "✖"} duplicado rechazado (${dupe.status}) → ${dupe.body?.message}`);

  // ── Parametrización manual y causación ─────────────────────────────────
  if (doc) {
    const params = await call(`/accounting/documents/${doc.id}`, {
      method: "PUT",
      body: JSON.stringify({
        accountCode: "512035",
        payableAccountCode: "233595",
        costCenterId: 1,
        siigoDocumentTypeId: 27,
        paymentTypeId: 5401,
        retentions: [{ kind: "RETEFUENTE", name: "Retefuente 3.5%", rate: 3.5, siigoTaxId: 900 }],
      }),
    });
    const retTotal = params.body?.data?.retentionsTotal;
    console.log(`\n${params.status < 300 ? "✔" : "✖"} parametrizar         ${params.status} → retenciones ${retTotal}, saldo ${params.body?.data?.balance}`);

    const ready = await call(`/accounting/documents/${doc.id}/ready`, { method: "POST", body: "{}" });
    console.log(`${ready.status < 300 ? "✔" : "✖"} confirmar            ${ready.status} → ${ready.body?.data?.status}`);

    const post = await call(`/accounting/documents/${doc.id}/post`, { method: "POST", body: "{}" });
    const expected = post.status === 409 && post.body?.code === "siigo_not_connected";
    console.log(`${expected || post.status < 300 ? "✔" : "✖"} causar en Siigo      ${post.status} → ${post.body?.message || post.body?.data?.posting?.siigoName}`);
  }

  // ── Tesorería: programar, pagar, ver el efecto en flujo de caja ─────────
  const payables = await call("/treasury/payables");
  const payable = payables.body?.data?.rows?.[0];
  if (payable) {
    console.log(`\n— Tesorería sobre ${payable.label} (saldo ${payable.balance}) —`);

    const schedule = await call("/treasury/schedules", {
      method: "POST",
      body: JSON.stringify({
        scheduledDate: new Date(Date.now() + 3 * 86400000).toISOString().slice(0, 10),
        allocations: [{ documentId: payable.documentId }],
      }),
    });
    console.log(`${schedule.status === 201 ? "✔" : "✖"} programar pago       ${schedule.status} → total ${schedule.body?.data?.total}`);

    const half = Math.round(payable.balance / 2);
    const payment = await call("/treasury/payments", {
      method: "POST",
      body: JSON.stringify({
        paidAt: new Date().toISOString().slice(0, 10),
        reference: "E2E-001",
        scheduleId: schedule.body?.data?.id,
        allocations: [{ documentId: payable.documentId, value: half }],
      }),
    });
    console.log(`${payment.status === 201 ? "✔" : "✖"} pago parcial         ${payment.status} → ${payment.body?.data?.value}`);

    const after = await call(`/documents/${payable.documentId}`);
    const d = after.body?.data?.document;
    const okPartial = d?.status === "PARTIALLY_PAID" && Math.abs(d.balance - (payable.balance - half)) < 1;
    console.log(`${okPartial ? "✔" : "✖"} saldo tras abono     estado ${d?.status}, saldo ${d?.balance} (esperado ${payable.balance - half})`);

    const rest = await call("/treasury/payments", {
      method: "POST",
      body: JSON.stringify({
        paidAt: new Date().toISOString().slice(0, 10),
        reference: "E2E-002",
        allocations: [{ documentId: payable.documentId, value: payable.balance - half }],
      }),
    });
    console.log(`${rest.status === 201 ? "✔" : "✖"} pago del saldo       ${rest.status}`);

    const final = await call(`/documents/${payable.documentId}`);
    const fd = final.body?.data?.document;
    console.log(`${fd?.status === "PAID" && fd.balance === 0 ? "✔" : "✖"} documento pagado     estado ${fd?.status}, saldo ${fd?.balance}`);

    const excess = await call("/treasury/payments", {
      method: "POST",
      body: JSON.stringify({
        paidAt: new Date().toISOString().slice(0, 10),
        allocations: [{ documentId: payable.documentId, value: 1000 }],
      }),
    });
    console.log(`${excess.status >= 400 ? "✔" : "✖"} sobrepago rechazado  ${excess.status} → ${excess.body?.message}`);

    const flow = await call("/treasury/cashflow");
    const actual = flow.body?.data?.totals?.actual;
    console.log(`${actual > 0 ? "✔" : "✖"} flujo de caja real   ${actual}`);

    const egreso = await call(`/treasury/payments/${payment.body?.data?.id}/post-siigo`, { method: "POST", body: "{}" });
    const egresoOk = egreso.status === 409 || egreso.status < 300 || egreso.status === 422;
    console.log(`${egresoOk ? "✔" : "✖"} egreso Siigo         ${egreso.status} → ${egreso.body?.message || egreso.body?.data?.payment?.siigo?.siigoName}`);
  }

  // ── Aislamiento multi-tenant ───────────────────────────────────────────
  const other = await fetch(`${BASE}/documents`, {
    headers: { Authorization: `Bearer ${token}`, "X-Erp-Tenant": "no-existe" },
  });
  console.log(`${other.status === 404 ? "✔" : "✖"} tenant inexistente rechazado (${other.status})`);

  const anon = await fetch(`${BASE}/documents`, { headers: { "X-Erp-Tenant": "demo" } });
  console.log(`${anon.status === 401 ? "✔" : "✖"} sin sesión rechazado (${anon.status})`);

  console.log(`\n${failures === 0 ? "TODO OK" : `${failures} fallo(s)`}`);
  process.exit(failures === 0 ? 0 : 1);
}

main().catch((err) => { console.error(err); process.exit(1); });
