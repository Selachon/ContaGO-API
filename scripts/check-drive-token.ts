import "dotenv/config";
import { connectMongo, db } from "../src/services/database.js";
await connectMongo();
const users=(db as any).collection("users");
const u=await users.findOne({email:"jose.perilla@proton.me"},{projection:{google_drives:1,google_drive:1,selected_google_drive_id:1}});
const drives=(u?.google_drives?.length?u.google_drives:(u?.google_drive?[u.google_drive]:[]))||[];
for(const d of drives){
  console.log("conn:",d.connection_id);
  console.log("  campos:",Object.keys(d));
  console.log("  email:",d.email||d.account_email||d.user_email);
  console.log("  token_expiry:",d.token_expiry||d.expiry||d.expires_at, "| ahora:", new Date().toISOString());
  console.log("  tiene access_token:",!!(d.access_token||d.token), "| refresh_token:",!!(d.refresh_token), "| root_folder:",d.root_folder_id||d.rootFolderId||d.folder_id);
}
process.exit(0);
