import Database from "better-sqlite3";
import { MongoClient } from "mongodb";
import path from "path";
import { fileURLToPath } from "url";

const MONGODB_URI = process.env.MONGODB_URI;
const MONGODB_DB = process.env.MONGODB_DB;

if (!MONGODB_URI) {
  console.error("MONGODB_URI no está definido.");
  process.exit(1);
}

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SQLITE_PATH = path.join(__dirname, "../data.db");

const sqlite = new Database(SQLITE_PATH, { readonly: true });

type SqlUser = {
  id: number;
  email: string;
  name: string;
  password_hash: string;
  is_admin: number;
  created_at: string;
};

type SqlPurchase = {
  user_id: number;
  tool_id: string;
};

const users = sqlite.prepare("SELECT * FROM users").all() as SqlUser[];
const purchases = sqlite.prepare("SELECT user_id, tool_id FROM user_purchases").all() as SqlPurchase[];

const purchasesByUser = new Map<number, string[]>();
for (const purchase of purchases) {
  const list = purchasesByUser.get(purchase.user_id) || [];
  list.push(purchase.tool_id);
  purchasesByUser.set(purchase.user_id, list);
}

const client = new MongoClient(MONGODB_URI);

async function run(): Promise<void> {
  await client.connect();
  const db = MONGODB_DB ? client.db(MONGODB_DB) : client.db();
  const usersCollection = db.collection("users");

  await usersCollection.createIndex({ email: 1 }, { unique: true });
  await usersCollection.createIndex({ legacyId: 1 });

  let inserted = 0;
  let skipped = 0;

  for (const user of users) {
    const email = user.email.toLowerCase().trim();
    const doc = {
      email,
      name: user.name,
      password_hash: user.password_hash,
      is_admin: !!user.is_admin,
      purchasedTools: purchasesByUser.get(user.id) || [],
      created_at: user.created_at,
      legacyId: user.id,
    };

    const result = await usersCollection.updateOne(
      { email },
      { $setOnInsert: doc },
      { upsert: true }
    );

    if (result.upsertedCount > 0) {
      inserted++;
    } else {
      skipped++;
    }
  }

  console.log(`Migración completa. Insertados: ${inserted}. Omitidos: ${skipped}.`);
}

run()
  .catch((err) => {
    console.error("Error durante la migración:", err);
    process.exitCode = 1;
  })
  .finally(async () => {
    sqlite.close();
    await client.close();
  });
