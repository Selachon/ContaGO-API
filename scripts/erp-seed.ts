/**
 * Crea (o completa) la empresa de demostración del ERP.
 *
 *   npm run erp:seed                      → crea la demo si no existe
 *   npm run erp:seed -- --reset           → borra sus datos y la reconstruye
 *   npm run erp:seed -- correo@empresa.co → además invita a ese correo
 */
import "dotenv/config";
import { connectMongo, getDb } from "../src/services/database.js";
import { COL, ensureErpIndexes } from "../src/erp/db.js";
import { getTenantBySlug } from "../src/erp/tenants.js";
import { seedDemoTenant } from "../src/erp/seed.js";

/** Borra los datos operativos de la demo (nunca toca otras empresas). */
async function resetDemo(): Promise<void> {
  const tenant = await getTenantBySlug("demo");
  if (!tenant) return;
  const collections = [
    COL.documents, COL.suppliers, COL.schedules, COL.payments, COL.bankAccounts,
    COL.costCenters, COL.rules, COL.comments, COL.audit, COL.notifications,
    COL.operations, COL.savedViews, COL.catalogs,
  ];
  for (const name of collections) {
    await getDb().collection(name).deleteMany({ tenantId: tenant.id });
  }
  console.log("Datos de la empresa demo eliminados.");
}

async function main() {
  await connectMongo();
  await ensureErpIndexes();

  const args = process.argv.slice(2);
  if (args.includes("--reset")) await resetDemo();

  const email = args.find((a) => a.includes("@")) || process.env.ADMIN_EMAIL;
  const result = await seedDemoTenant(email);
  console.log(
    `Empresa demo: ${result.tenant.name} (${result.tenant.slug}.contago.com.co)\n` +
    `${result.created ? "creada" : "ya existía"} · ${result.documents} documentos · ${result.payments} pagos`
  );
  process.exit(0);
}

main().catch((err) => { console.error("Fallo el seed:", err); process.exit(1); });
