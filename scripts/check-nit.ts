import "dotenv/config";
import { connectMongo, db } from "../src/services/database.js";

const nit = process.argv[2] || "900815934";
const norm = (n: string) => (n || "").replace(/[-\s]/g, "").trim();

await connectMongo();
const users = (db as any).collection("users");
const all = await users.find({}, { projection: { email: 1, name: 1, nits: 1, is_admin: 1 } }).toArray();
const withNit = all.filter((u: any) => (u.nits || []).map(norm).includes(nit));
console.log(`Usuarios con NIT ${nit} autorizado:`, withNit.length ? withNit.map((u: any) => u.email) : "NINGUNO");
console.log("\nCuentas de prueba relevantes:");
for (const u of all) {
  if (/perilla|administracion|karina|pulido|laboratorio|solucion/i.test((u.email || "") + (u.name || ""))) {
    console.log(`  - ${u.email} (${u.name}) | admin:${!!u.is_admin} | nits:[${(u.nits || []).map(norm).join(", ")}]`);
  }
}
process.exit(0);
