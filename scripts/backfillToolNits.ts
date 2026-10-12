/**
 * Backfill de `user.toolNits` a partir de la lista global `user.nits`.
 *
 * Para cada usuario NO admin, por cada herramienta con alcance por NIT que el
 * usuario tiene comprada y que aún no tiene entrada en `toolNits`, se copia
 * `user.nits` como lista habilitada para esa herramienta.
 *
 * Uso:
 *   tsx scripts/backfillToolNits.ts          # dry-run (no escribe)
 *   tsx scripts/backfillToolNits.ts --apply  # aplica los cambios
 */
import "dotenv/config";
import { connectMongo, db, NIT_SCOPED_TOOLS, TOOL_SUCCESSOR } from "../src/services/database.js";

const APPLY = process.argv.includes("--apply");

function expandTools(tools: string[]): Set<string> {
  const out = new Set(tools);
  for (const t of tools) {
    const succ = TOOL_SUCCESSOR[t];
    if (succ) out.add(succ);
  }
  return out;
}

await connectMongo();
const users = (db as any).collection("users");
const all = await users
  .find({}, { projection: { email: 1, name: 1, nits: 1, is_admin: 1, purchasedTools: 1, toolNits: 1 } })
  .toArray();

let updated = 0;
let sinNits = 0;
const ops: { updateOne: { filter: any; update: any } }[] = [];

for (const u of all) {
  if (u.is_admin) continue;
  const purchased = expandTools(u.purchasedTools || []);
  const scoped = [...purchased].filter((t) => NIT_SCOPED_TOOLS.has(t));
  if (scoped.length === 0) continue;

  const current: Record<string, string[]> = u.toolNits || {};
  const nits: string[] = (u.nits || []).map((n: string) => String(n).trim()).filter(Boolean);
  const patch: Record<string, string[]> = {};

  for (const tool of scoped) {
    if (current[tool]?.length) continue; // ya configurada
    if (nits.length === 0) {
      console.warn(`  ! ${u.email} — herramienta "${tool}" sin NITs para copiar (lista global vacía)`);
      sinNits++;
      continue;
    }
    patch[tool] = nits;
  }

  if (Object.keys(patch).length === 0) continue;

  console.log(`  ~ ${u.email}: ${Object.entries(patch).map(([k, v]) => `${k}=[${v.join(",")}]`).join("  ")}`);
  updated++;
  const set: Record<string, string[]> = {};
  for (const [k, v] of Object.entries(patch)) set[`toolNits.${k}`] = v;
  ops.push({ updateOne: { filter: { _id: u._id }, update: { $set: set } } });
}

console.log(`\n${updated} usuario(s) a actualizar · ${sinNits} herramienta(s) sin NITs para copiar (quedan bloqueadas hasta asignar).`);

if (APPLY && ops.length) {
  const r = await users.bulkWrite(ops);
  console.log(`Aplicado. modified=${r.modifiedCount}`);
} else if (!APPLY) {
  console.log("Dry-run. Ejecuta con --apply para escribir.");
}

process.exit(0);
