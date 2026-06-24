import "dotenv/config";
import { connectMongo, getUserByEmail, getUserGoogleDriveById, updateUserDriveTokens } from "../src/services/database.js";
import { getOrCreateRootFolder, uploadInvoiceFilesToDrive } from "../src/services/googleDrive.js";
import { encryptToken } from "../src/utils/encryption.js";

const CONN = "b0e878a1-8128-44c7-bade-8602282bbb69";
await connectMongo();
const u = await getUserByEmail("jose.perilla@proton.me");
if (!u) { console.error("no user"); process.exit(1); }
const driveConfig = await getUserGoogleDriveById(u.id, CONN);
console.log("driveConfig resuelto:", !!driveConfig);
if (!driveConfig) process.exit(1);
const onTokenRefresh = async (tok: string, exp: number) => {
  await updateUserDriveTokens(u.id, encryptToken(tok), new Date(exp).toISOString(), CONN);
};
try {
  const root = await getOrCreateRootFolder(driveConfig, u.id, onTokenRefresh);
  console.log("rootFolder OK:", root);
  const xml = Buffer.from('<?xml version="1.0"?><Invoice><ID>TEST-DRIVE</ID></Invoice>');
  const res = await uploadInvoiceFilesToDrive(null, xml, "TEST-DRIVE-001", "900815934", "2026-06-24", driveConfig, u.id, onTokenRefresh, "received");
  console.log("✅ UPLOAD OK:", JSON.stringify(res));
} catch (e: any) {
  console.error("❌ UPLOAD ERROR:", e?.message || e);
  console.error((e?.stack || "").split("\n").slice(0, 6).join("\n"));
}
process.exit(0);
