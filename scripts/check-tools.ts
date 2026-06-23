import "dotenv/config";
import { connectMongo, db, getUserByEmail } from "../src/services/database.js";
import { checkToolAccess } from "../src/middleware/requireToolAccess.js";

await connectMongo();
const users = (db as any).collection("users");
const emails = process.argv.slice(2);
const list = emails.length ? emails : [
  "karina_andreap@hotmail.com", "administracion@contago.com.co", "jose.perilla@proton.me",
];
for (const email of list) {
  const u = await getUserByEmail(email);
  if (!u) { console.log(email, "-> no existe"); continue; }
  const masiva = await checkToolAccess(u.id, u.is_admin, "dian-mass-download");
  const expo = await checkToolAccess(u.id, u.is_admin, "dian-cufe-downloader");
  const terc = await checkToolAccess(u.id, u.is_admin, "dian-third-parties-excel");
  const rec = await users.findOne({ _id: (await import("mongodb")).ObjectId.createFromHexString(u.id) }, { projection: { purchasedTools: 1 } });
  console.log(`${email} | admin:${u.is_admin}`);
  console.log(`   tools comprados: [${(rec?.purchasedTools || []).join(", ")}]`);
  console.log(`   masiva:${masiva.ok}  exportador:${expo.ok}  terceros:${terc.ok}`);
}
process.exit(0);
