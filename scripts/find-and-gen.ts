// Busca un usuario por nombre/email (substring) y genera/regenera su código.
// Uso: npx tsx scripts/find-and-gen.ts "perilla"
import "dotenv/config";
import { connectMongo, resetExtActivation, generateExtActivationCode, db } from "../src/services/database.js";

const q = (process.argv[2] || "perilla").toLowerCase();
await connectMongo();
const users = (db as any).collection("users");
const matches = await users.find({
  $or: [
    { name: { $regex: q, $options: "i" } },
    { email: { $regex: q, $options: "i" } },
  ],
}).limit(10).toArray();

if (!matches.length) { console.error(`✗ Sin coincidencias para "${q}"`); process.exit(1); }
for (const u of matches) {
  console.log(`• ${u.name} <${u.email}> id=${u._id}`);
}
const target = matches[0];
await resetExtActivation(target._id.toString());
const r = await generateExtActivationCode(target._id.toString());
console.log(`\n✓ Código para ${target.name} <${target.email}>: ${r.code}`);
process.exit(0);
