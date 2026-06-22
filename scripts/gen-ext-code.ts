// Emite (una sola vez) el código de activación de la extensión para un usuario.
// Uso: npx tsx scripts/gen-ext-code.ts [email]
import "dotenv/config";
import { connectMongo, getUserByEmail, generateExtActivationCode, getExtActivationStatus } from "../src/services/database.js";

const email = process.argv[2] || "administracion@contago.com.co";

await connectMongo();
const user = await getUserByEmail(email);
if (!user) {
  console.error(`✗ No existe el usuario ${email}`);
  process.exit(1);
}
const res = await generateExtActivationCode(user.id);
if (res.already) {
  const st = await getExtActivationStatus(user.id);
  console.log(`ℹ El usuario ya generó un código.${st?.activated ? " Ya fue ACTIVADO (canjeado)." : ` Código vigente: ${st?.code}`}`);
  console.log("  Para regenerar: POST /admin/users/" + user.id + "/reset-ext-activation (admin)");
} else {
  console.log(`✓ Código de activación para ${email}:`);
  console.log("  " + res.code);
}
process.exit(0);
