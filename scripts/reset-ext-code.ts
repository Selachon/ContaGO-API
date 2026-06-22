// Restablece y regenera el código de activación (para pruebas).
// Uso: npx tsx scripts/reset-ext-code.ts [email]
import "dotenv/config";
import { connectMongo, getUserByEmail, resetExtActivation, generateExtActivationCode } from "../src/services/database.js";

const email = process.argv[2] || "administracion@contago.com.co";
await connectMongo();
const u = await getUserByEmail(email);
if (!u) { console.error(`✗ No existe ${email}`); process.exit(1); }
await resetExtActivation(u.id);
const r = await generateExtActivationCode(u.id);
console.log(`✓ NUEVO código para ${email}: ${r.code}`);
process.exit(0);
