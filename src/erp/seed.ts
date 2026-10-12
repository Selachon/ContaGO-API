/**
 * Datos de demostración. Permiten recorrer el producto completo sin
 * integraciones reales. Nunca contienen datos de clientes reales.
 */
import { randomUUID } from "node:crypto";
import { COL, col, now, today } from "./db.js";
import { createTenant, getTenantBySlug, inviteMember, type ErpTenant } from "./tenants.js";
import { buildDocument } from "./documents.js";
import { round2 } from "./money.js";
import type { ErpDocument } from "./types.js";

const DEMO_SLUG = "demo";

function shift(days: number): string {
  const d = new Date(Date.parse(`${today()}T12:00:00Z`) + days * 86_400_000);
  return d.toISOString().slice(0, 10);
}

const SUPPLIERS = [
  { nit: "900123456", name: "Servicios Integrales Andina S.A.S." },
  { nit: "901456789", name: "Tecnología y Nube Colombia S.A.S." },
  { nit: "830001234", name: "Arrendamientos del Norte Ltda." },
  { nit: "1020304050", name: "María Fernanda Ruiz (Cuenta de cobro)" },
  { nit: "860002964", name: "Energía y Redes S.A. E.S.P." },
  { nit: "900777888", name: "Logística Express del Caribe S.A.S." },
];

interface SeedSpec {
  supplier: (typeof SUPPLIERS)[number];
  type: ErpDocument["type"];
  source: ErpDocument["source"];
  status: ErpDocument["status"];
  subtotal: number;
  ivaRate: number;
  issueOffset: number;
  dueOffset: number;
  description: string;
  costCenterId?: number;
  posted?: { name: string; id: string };
  error?: string;
  paid?: number;
}

const SPECS: SeedSpec[] = [
  {
    supplier: SUPPLIERS[0], type: "ELECTRONIC_INVOICE", source: "DIAN", status: "PENDING_APPROVAL",
    subtotal: 4_200_000, ivaRate: 19, issueOffset: -3, dueOffset: 27,
    description: "Servicios profesionales de mantenimiento — agosto", costCenterId: 1,
  },
  {
    supplier: SUPPLIERS[1], type: "ELECTRONIC_INVOICE", source: "DIAN", status: "PENDING_ACCOUNTING",
    subtotal: 1_850_000, ivaRate: 19, issueOffset: -6, dueOffset: 24,
    description: "Licenciamiento de plataforma en la nube", costCenterId: 2,
  },
  {
    supplier: SUPPLIERS[2], type: "ELECTRONIC_INVOICE", source: "DIAN", status: "PAYABLE",
    subtotal: 8_000_000, ivaRate: 0, issueOffset: -20, dueOffset: 10,
    description: "Canon de arrendamiento oficina principal", costCenterId: 1,
    posted: { name: "FC-1-482", id: randomUUID() },
  },
  {
    supplier: SUPPLIERS[3], type: "COLLECTION_ACCOUNT", source: "FORM", status: "APPROVED",
    subtotal: 3_500_000, ivaRate: 0, issueOffset: -2, dueOffset: 13,
    description: "Honorarios de asesoría contable — agosto", costCenterId: 3,
  },
  {
    supplier: SUPPLIERS[4], type: "ELECTRONIC_INVOICE", source: "DIAN", status: "POSTING_ERROR",
    subtotal: 2_340_000, ivaRate: 19, issueOffset: -8, dueOffset: -1,
    description: "Consumo de energía — sede norte", costCenterId: 1,
    error: "No se encontró el tercero 860.002.964 en Siigo.",
  },
  {
    supplier: SUPPLIERS[5], type: "ELECTRONIC_INVOICE", source: "DIAN", status: "PAYABLE",
    subtotal: 12_500_000, ivaRate: 19, issueOffset: -35, dueOffset: -5,
    description: "Transporte y distribución nacional", costCenterId: 2,
    posted: { name: "FC-1-476", id: randomUUID() },
  },
  {
    supplier: SUPPLIERS[0], type: "ELECTRONIC_INVOICE", source: "DIAN", status: "PARTIALLY_PAID",
    subtotal: 6_000_000, ivaRate: 19, issueOffset: -25, dueOffset: 5,
    description: "Mantenimiento preventivo trimestral", costCenterId: 1,
    posted: { name: "FC-1-470", id: randomUUID() }, paid: 3_000_000,
  },
  {
    supplier: SUPPLIERS[1], type: "ELECTRONIC_INVOICE", source: "DIAN", status: "PAYMENT_POSTED",
    subtotal: 1_500_000, ivaRate: 19, issueOffset: -45, dueOffset: -15,
    description: "Soporte técnico mensual — julio", costCenterId: 2,
    posted: { name: "FC-1-455", id: randomUUID() }, paid: 1_785_000,
  },
  {
    supplier: SUPPLIERS[2], type: "CREDIT_CARD", source: "MANUAL", status: "PENDING_REVIEW",
    subtotal: 980_000, ivaRate: 0, issueOffset: -1, dueOffset: 29,
    description: "Cuota tarjeta de crédito corporativa", costCenterId: 3,
  },
  {
    supplier: SUPPLIERS[4], type: "ELECTRONIC_INVOICE", source: "DIAN", status: "REJECTED",
    subtotal: 450_000, ivaRate: 19, issueOffset: -4, dueOffset: 26,
    description: "Cargo por reconexión (no corresponde a la sede)", costCenterId: 1,
  },
];

export interface SeedResult {
  tenant: ErpTenant;
  created: boolean;
  documents: number;
  payments: number;
}

/**
 * Crea (o completa) la empresa de demostración con documentos en todos los
 * estados del flujo, más pagos, programaciones y cuentas bancarias.
 */
export async function seedDemoTenant(adminEmail?: string): Promise<SeedResult> {
  let tenant = await getTenantBySlug(DEMO_SLUG);
  let created = false;

  if (!tenant) {
    tenant = await createTenant({
      name: "Demo ContaGO S.A.S.",
      legalName: "Demo ContaGO S.A.S.",
      nit: "901000001",
      dv: "7",
      slug: DEMO_SLUG,
      settings: {
        approvalThreshold: 1_000_000,
        requireApproval: true,
        defaultPaymentTermsDays: 30,
        defaultPayableAccountCode: "233595",
      },
    });
    created = true;
  }

  if (adminEmail) {
    await inviteMember(tenant.id, { email: adminEmail, role: "CLIENT_ADMIN", name: "Administrador demo" });
  }

  const existing = await col(COL.documents).countDocuments({ tenantId: tenant.id });
  if (existing > 0) {
    return {
      tenant,
      created,
      documents: existing,
      payments: await col(COL.payments).countDocuments({ tenantId: tenant.id }),
    };
  }

  // Cuentas bancarias
  const bankIds: string[] = [];
  for (const b of [
    { bank: "Bancolombia", accountName: "Cuenta operativa", maskedAccountNumber: "••••4832", accountType: "CORRIENTE", siigoAccountCode: "111005" },
    { bank: "Davivienda", accountName: "Cuenta nómina", maskedAccountNumber: "••••1197", accountType: "AHORROS", siigoAccountCode: "111010" },
  ]) {
    const res = await col(COL.bankAccounts).insertOne({
      tenantId: tenant.id, ...b, currency: "COP", active: true, createdAt: now(), updatedAt: now(),
    } as any);
    bankIds.push(String(res.insertedId));
  }

  // Centros de costo (con subcentros locales)
  for (const cc of [
    { siigoId: 1, code: "ADM", name: "Administración", subs: ["Gerencia", "Contabilidad"] },
    { siigoId: 2, code: "OPE", name: "Operaciones", subs: ["Logística", "Tecnología"] },
    { siigoId: 3, code: "COM", name: "Comercial", subs: ["Mercadeo", "Ventas"] },
  ]) {
    await col(COL.costCenters).insertOne({
      tenantId: tenant.id,
      siigoId: cc.siigoId,
      code: cc.code,
      name: cc.name,
      active: true,
      subCenters: cc.subs.map((s) => ({ id: randomUUID(), code: s.slice(0, 3).toUpperCase(), name: s, active: true })),
      lastSyncedAt: now(),
    } as any);
  }

  // Reglas contables de ejemplo
  await col(COL.rules).insertMany([
    {
      tenantId: tenant.id, name: "Arrendamientos → 512035", priority: 10, active: true,
      conditions: { descriptionContains: "arrendamiento" },
      actions: {
        accountCode: "512035", payableAccountCode: "233595", costCenterId: 1,
        retentions: [{ kind: "RETEFUENTE", name: "Retefuente arrendamientos 3.5%", rate: 3.5 }],
      },
      createdAt: now(), updatedAt: now(),
    },
    {
      tenantId: tenant.id, name: "Servicios en la nube → 513540", priority: 20, active: true,
      conditions: { supplierNit: "901456789" },
      actions: { accountCode: "513540", payableAccountCode: "233595", costCenterId: 2 },
      createdAt: now(), updatedAt: now(),
    },
  ] as any);

  // Documentos
  let counter = 480;
  const inserted: ErpDocument[] = [];
  for (const spec of SPECS) {
    counter += 1;
    const taxes = round2((spec.subtotal * spec.ivaRate) / 100);
    const total = round2(spec.subtotal + taxes);
    const doc = buildDocument(tenant.id, {
      source: spec.source,
      type: spec.type,
      supplierNit: spec.supplier.nit,
      supplierName: spec.supplier.name,
      documentPrefix: spec.source === "DIAN" ? "SETP" : undefined,
      documentNumber: String(counter),
      cufe: spec.source === "DIAN" ? `demo-${randomUUID().replace(/-/g, "")}` : undefined,
      issueDate: shift(spec.issueOffset),
      dueDate: shift(spec.dueOffset),
      subtotal: spec.subtotal,
      taxes,
      total,
      description: spec.description,
      costCenterId: spec.costCenterId,
      estimatedPaymentDate: shift(spec.dueOffset),
      items: [
        {
          description: spec.description,
          quantity: 1,
          unitPrice: spec.subtotal,
          subtotal: spec.subtotal,
          taxAmount: taxes,
          total,
        },
      ],
      status: spec.status,
    });

    if (spec.posted) {
      doc.siigo = {
        status: "SUCCESS", attempts: 1, strategy: spec.type === "ELECTRONIC_INVOICE" ? "FC" : "DS",
        siigoId: spec.posted.id, siigoName: spec.posted.name,
        siigoNumber: Number(spec.posted.name.split("-").pop()),
        postedAt: now(),
      };
      doc.accounting = {
        accountCode: "512035", accountName: "Arrendamientos", payableAccountCode: "233595",
        costCenterId: spec.costCenterId, retentions: [],
        confirmedBy: "Contador ContaGO", confirmedAt: now(),
        suggestion: { source: "RULE", reason: 'Aplicada la regla "Arrendamientos → 512035".', confidence: 0.95 },
      };
    }
    if (spec.error) {
      doc.siigo = {
        status: "FAILED", attempts: 2, strategy: "FC",
        errorCode: "supplier_not_found", errorMessage: spec.error, friendlyError: spec.error,
        lastAttemptAt: now(),
      };
      doc.accounting = {
        accountCode: "513535", accountName: "Servicios públicos", payableAccountCode: "233595",
        costCenterId: spec.costCenterId, retentions: [],
        confirmedBy: "Contador ContaGO", confirmedAt: now(),
      };
    }
    if (spec.status === "REJECTED") {
      doc.rejectedAt = now();
      doc.rejectedByName = "Laura Gómez";
      doc.rejectionReason = "El cargo no corresponde a esta sede; solicitar nota crédito al proveedor.";
    }
    if (spec.status === "APPROVED" || spec.posted) {
      doc.approvedAt = now();
      doc.approvedByName = "Laura Gómez";
    }
    if (spec.paid) {
      doc.balance = round2(total - spec.paid);
    }

    const res = await col(COL.documents).insertOne(doc as any);
    inserted.push({ id: String(res.insertedId), ...doc });

    await col(COL.suppliers).updateOne(
      { tenantId: tenant.id, identification: spec.supplier.nit },
      {
        $set: { name: spec.supplier.name, updatedAt: now() },
        $inc: { documentCount: 1 },
        $setOnInsert: {
          tenantId: tenant.id, identification: spec.supplier.nit,
          identificationType: spec.supplier.nit.length > 9 ? "31" : "13",
          active: true, createdAt: now(),
        },
      },
      { upsert: true }
    );
  }

  // Pagos: uno parcial y uno total con egreso en Siigo.
  const partial = inserted.find((d) => d.status === "PARTIALLY_PAID");
  const fullyPaid = inserted.find((d) => d.status === "PAYMENT_POSTED");
  let payments = 0;

  if (partial) {
    await col(COL.payments).insertOne({
      tenantId: tenant.id, paidAt: shift(-10), value: 3_000_000,
      bankAccountId: bankIds[0], bankAccountLabel: "Bancolombia ••••4832",
      reference: "TRF-889231",
      allocations: [{
        documentId: partial.id, documentLabel: "SETP-487",
        supplierNit: partial.supplierNit, siigoPrefix: "FC-1", siigoConsecutive: 470, value: 3_000_000,
      }],
      files: [], siigo: { status: "PENDING", attempts: 0 },
      createdAt: now(), updatedAt: now(),
    } as any);
    payments += 1;
  }
  if (fullyPaid) {
    await col(COL.payments).insertOne({
      tenantId: tenant.id, paidAt: shift(-12), value: fullyPaid.total,
      bankAccountId: bankIds[0], bankAccountLabel: "Bancolombia ••••4832",
      reference: "TRF-884410",
      allocations: [{
        documentId: fullyPaid.id, documentLabel: "SETP-488",
        supplierNit: fullyPaid.supplierNit, siigoPrefix: "FC-1", siigoConsecutive: 455, value: fullyPaid.total,
      }],
      files: [],
      siigo: { status: "SUCCESS", attempts: 1, siigoId: randomUUID(), siigoName: "RP-1-199", siigoNumber: 199, postedAt: now() },
      createdAt: now(), updatedAt: now(),
    } as any);
    payments += 1;
  }

  // Una programación futura
  const toSchedule = inserted.filter((d) => d.status === "PAYABLE").slice(0, 2);
  if (toSchedule.length) {
    await col(COL.schedules).insertOne({
      tenantId: tenant.id,
      scheduledDate: shift(4),
      bankAccountId: bankIds[0],
      bankAccountLabel: "Bancolombia ••••4832",
      allocations: toSchedule.map((d) => ({
        documentId: d.id, documentLabel: `SETP-${d.documentNumber}`,
        supplierNit: d.supplierNit, siigoPrefix: "FC-1",
        siigoConsecutive: d.siigo?.siigoNumber, value: d.balance,
      })),
      total: round2(toSchedule.reduce((s, d) => s + d.balance, 0)),
      status: "SCHEDULED",
      observations: "Pago programado del viernes",
      createdAt: now(), updatedAt: now(),
    } as any);
  }

  await col(COL.audit).insertOne({
    tenantId: tenant.id, action: "TENANT_SEEDED", entityType: "tenant", entityId: tenant.id,
    actorName: "Sistema", summary: `Datos de demostración creados (${inserted.length} documentos).`,
    createdAt: now(),
  } as any);

  return { tenant, created, documents: inserted.length, payments };
}
