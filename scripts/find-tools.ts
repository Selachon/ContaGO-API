import "dotenv/config";
import { connectMongo, db } from "../src/services/database.js";
import { checkToolAccess } from "../src/middleware/requireToolAccess.js";

const q = process.argv[2] || "taborda";
await connectMongo();
const users = (db as any).collection("users");
const found = await users.find(
  { $or: [{ name: { $regex: q, $options: "i" } }, { email: { $regex: q, $options: "i" } }] },
  { projection: { email: 1, name: 1, is_admin: 1, purchasedTools: 1, nits: 1 } }
).toArray();
if (!found.length) { console.log("Sin coincidencias para", q); process.exit(0); }
for (const u of found) {
  const id = u._id.toString();
  const masiva = await checkToolAccess(id, u.is_admin, "dian-mass-download");
  const expo = await checkToolAccess(id, u.is_admin, "dian-cufe-downloader");
  const terc = await checkToolAccess(id, u.is_admin, "dian-third-parties-excel");
  console.log(`${u.name} <${u.email}> | admin:${!!u.is_admin}`);
  console.log(`   tools: [${(u.purchasedTools || []).join(", ")}]  nits:[${(u.nits || []).join(", ")}]`);
  console.log(`   acceso -> masiva:${masiva.ok}  exportador:${expo.ok}  terceros:${terc.ok}`);
}
process.exit(0);
