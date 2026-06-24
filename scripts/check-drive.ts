import "dotenv/config";
import { connectMongo, db } from "../src/services/database.js";

await connectMongo();
const users = (db as any).collection("users");
const cur = users.find(
  { $or: [{ "google_drives.0": { $exists: true } }, { google_drive: { $exists: true } }, { exporter_prefs: { $exists: true } }] },
  { projection: { email: 1, exporter_prefs: 1, selected_google_drive_id: 1, google_drives: 1, google_drive: 1 } }
);
const docs = await cur.toArray();
for (const u of docs) {
  const drives = (u.google_drives && u.google_drives.length ? u.google_drives : (u.google_drive ? [u.google_drive] : [])) || [];
  console.log(`\n${u.email}`);
  console.log(`  exporter_prefs:`, JSON.stringify(u.exporter_prefs || null));
  console.log(`  selected_google_drive_id:`, u.selected_google_drive_id || null);
  console.log(`  google_drives conn_ids:`, drives.map((d: any) => ({ conn: d.connection_id, email: d.email })));
}
process.exit(0);
