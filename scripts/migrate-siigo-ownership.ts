/**
 * Migración no destructiva: dueño por empresa en la herramienta Siigo.
 *
 *   - Asigna ownerUserId = admin principal a toda empresa que no tenga dueño.
 *   - Garantiza sharedWith = [] en cada empresa.
 *   - Crea índices (username único, ownerUserId, sharedWith, profiles.companyId).
 *
 * NO borra ni renombra nada. Es idempotente: correrlo varias veces es seguro.
 *
 * Uso:
 *   npx tsx scripts/migrate-siigo-ownership.ts            # DRY-RUN (solo lectura)
 *   npx tsx scripts/migrate-siigo-ownership.ts --apply    # Aplica los cambios
 */
import "dotenv/config";
import { MongoClient, type Collection, type Document } from "mongodb";

const APPLY = process.argv.includes("--apply");

const uri = process.env.MONGODB_URI;
const dbName = process.env.MONGODB_DB;
const adminEmail = (process.env.ADMIN_EMAIL || "").trim().toLowerCase();

function line(label: string, value: unknown) {
  console.log(`  ${label.padEnd(38, ".")} ${value}`);
}

/** createIndex tolerante: ignora falta de espacio (Railway) y reporta duplicados. */
async function safeIndex(
  col: Collection<Document>,
  keys: Record<string, 1 | -1>,
  options: Parameters<Collection["createIndex"]>[1] = {}
): Promise<void> {
  const desc = `${col.collectionName} ${JSON.stringify(keys)}${options.unique ? " (unique)" : ""}`;
  try {
    await col.createIndex(keys, options);
    console.log(`  ✓ índice: ${desc}`);
  } catch (err) {
    const code = (err as { code?: number }).code;
    if (code === 14031) { console.warn(`  ⚠ índice omitido por espacio (Railway): ${desc}`); return; }
    if (code === 11000) { console.warn(`  ⚠ índice ÚNICO no creado por duplicados existentes: ${desc} → ${(err as Error).message}`); return; }
    throw err;
  }
}

async function main() {
  if (!uri) throw new Error("MONGODB_URI no está definido");

  console.log("\n══════════════════════════════════════════════════════════");
  console.log(`  Migración Siigo · ownership   [${APPLY ? "APPLY" : "DRY-RUN (solo lectura)"}]`);
  console.log("══════════════════════════════════════════════════════════\n");

  const client = new MongoClient(uri);
  await client.connect();
  try {
    const db = dbName ? client.db(dbName) : client.db();
    const companies = db.collection("siigoCompanies");
    const users = db.collection("users");
    const profiles = db.collection("siigoSupplierProfiles");
    const config = db.collection("siigoConfig");

    // ── Resolver admin principal ──────────────────────────────────────
    let admin = adminEmail ? await users.findOne({ email: adminEmail }) : null;
    if (!admin) admin = await users.findOne({ is_admin: true }, { sort: { created_at: 1 } });
    if (!admin) throw new Error("No se encontró ningún admin (ADMIN_EMAIL ni is_admin:true).");
    const adminId = admin._id.toString();

    console.log("Admin principal (futuro dueño):");
    line("email", admin.email);
    line("_id", adminId);
    console.log("");

    // ── Diagnóstico ───────────────────────────────────────────────────
    const totalCompanies = await companies.countDocuments({});
    const withoutOwner = await companies.countDocuments({ ownerUserId: { $exists: false } });
    const withOwner = totalCompanies - withoutOwner;
    const withoutShared = await companies.countDocuments({ sharedWith: { $exists: false } });
    const totalProfiles = await profiles.countDocuments({});
    const totalConfig = await config.countDocuments({});

    console.log("Estado actual:");
    line("empresas (total)", totalCompanies);
    line("· con dueño", withOwner);
    line("· SIN dueño (se asignarán)", withoutOwner);
    line("· sin campo sharedWith", withoutShared);
    line("perfiles de proveedor", totalProfiles);
    line("docs de config", totalConfig);
    console.log("");

    // Duplicados de username (bloquean el índice único)
    const dups = await companies.aggregate([
      { $group: { _id: "$username", n: { $sum: 1 } } },
      { $match: { n: { $gt: 1 } } },
    ]).toArray();

    console.log("Listado de empresas:");
    const list = await companies.find({}, { projection: { name: 1, username: 1, ownerUserId: 1 } }).toArray();
    for (const c of list) {
      const owner = c.ownerUserId ? (c.ownerUserId === adminId ? "→ admin" : `→ ${c.ownerUserId}`) : "(SIN dueño)";
      console.log(`  • ${String(c.name || "?").padEnd(32)} ${String(c.username || "?").padEnd(24)} ${owner}`);
    }
    if (dups.length) {
      console.log("\n  ⚠ usernames DUPLICADOS (impedirán índice único hasta resolverlos):");
      for (const d of dups) console.log(`     - "${d._id}" × ${d.n}`);
    }
    console.log("");

    if (!APPLY) {
      console.log("DRY-RUN: no se escribió nada. Repite con --apply para aplicar.\n");
      return;
    }

    // ── Aplicar (idempotente) ─────────────────────────────────────────
    console.log("Aplicando cambios...");
    const r1 = await companies.updateMany(
      { ownerUserId: { $exists: false } },
      { $set: { ownerUserId: adminId } }
    );
    line("empresas con dueño asignado", r1.modifiedCount);

    const r2 = await companies.updateMany(
      { sharedWith: { $exists: false } },
      { $set: { sharedWith: [] } }
    );
    line("empresas con sharedWith inicializado", r2.modifiedCount);
    console.log("");

    console.log("Creando índices...");
    if (dups.length === 0) {
      await safeIndex(companies, { username: 1 }, { unique: true });
    } else {
      console.warn("  ⚠ índice único de username OMITIDO por duplicados (resuélvelos y reintenta).");
      await safeIndex(companies, { username: 1 });
    }
    await safeIndex(companies, { ownerUserId: 1 });
    await safeIndex(companies, { sharedWith: 1 });
    await safeIndex(profiles, { companyId: 1 });
    console.log("");

    // ── Verificación ──────────────────────────────────────────────────
    const stillWithout = await companies.countDocuments({ ownerUserId: { $exists: false } });
    const totalAfter = await companies.countDocuments({});
    console.log("Verificación:");
    line("empresas (total, debe ser igual)", `${totalCompanies} → ${totalAfter}`);
    line("empresas SIN dueño (debe ser 0)", stillWithout);
    console.log(stillWithout === 0 && totalAfter === totalCompanies ? "\n✓ Migración OK.\n" : "\n✗ Revisar: verificación no cuadra.\n");
  } finally {
    await client.close();
  }
}

main().catch((e) => { console.error("\n✗ Error en la migración:", e); process.exit(1); });
