import bcrypt from "bcrypt";
import { MongoClient, type Db, type Collection, ObjectId } from "mongodb";
import type { User } from "../types/auth.js";

interface UserRecord {
  _id: ObjectId;
  email: string;
  name: string;
  password_hash: string;
  is_admin: boolean;
  purchasedTools: string[];
  nits: string[];
  created_at: string;
  legacyId?: number;
}


let client: MongoClient | null = null;
let db: Db | null = null;

function usersCollection(): Collection<UserRecord> {
  if (!db) {
    throw new Error("MongoDB no está conectado");
  }
  return db.collection<UserRecord>("users");
}

export async function connectMongo(): Promise<void> {
  const uri = process.env.MONGODB_URI;
  const dbName = process.env.MONGODB_DB;

  if (!uri) {
    throw new Error("MONGODB_URI no está definido");
  }

  client = new MongoClient(uri);
  await client.connect();

  db = dbName ? client.db(dbName) : client.db();

  // Índices
  const users = usersCollection();
  await users.createIndex({ email: 1 }, { unique: true });
  await users.createIndex({ legacyId: 1 });
}

function mapUser(record: UserRecord | null): User | null {
  if (!record) return null;
  return {
    id: record._id.toString(),
    email: record.email,
    name: record.name,
    password_hash: record.password_hash,
    is_admin: record.is_admin,
    nits: record.nits || [],
    created_at: record.created_at,
  };
}

// ============================================
// User functions
// ============================================

export async function createUser(
  email: string,
  name: string,
  password: string,
  isAdmin = false,
  nits: string[] = [],
  purchasedTools: string[] = []
): Promise<User | null> {
  try {
    const hash = await bcrypt.hash(password, 10);
    const record: UserRecord = {
      _id: new ObjectId(),
      email: email.toLowerCase().trim(),
      name: name.trim(),
      password_hash: hash,
      is_admin: isAdmin,
      purchasedTools,
      nits,
      created_at: new Date().toISOString(),
    };

    await usersCollection().insertOne(record);
    return mapUser(record);
  } catch (err) {
    console.error("Error creating user:", err);
    return null;
  }
}

export async function getUserById(id: string): Promise<User | null> {
  try {
    const oid = new ObjectId(id);
    const record = await usersCollection().findOne({ _id: oid });
    return mapUser(record);
  } catch {
    return null;
  }
}

export async function getUserByEmail(email: string): Promise<User | null> {
  const record = await usersCollection().findOne({ email: email.toLowerCase().trim() });
  return mapUser(record);
}

export async function verifyPassword(user: User, password: string): Promise<boolean> {
  return bcrypt.compare(password, user.password_hash);
}

// ============================================
// Purchase functions
// ============================================

export async function getUserPurchases(userId: string): Promise<string[]> {
  try {
    const oid = new ObjectId(userId);
    const record = await usersCollection().findOne({ _id: oid }, { projection: { purchasedTools: 1 } });
    return record?.purchasedTools || [];
  } catch {
    return [];
  }
}

export async function addPurchase(userId: string, toolId: string): Promise<boolean> {
  try {
    const oid = new ObjectId(userId);
    const result = await usersCollection().updateOne(
      { _id: oid },
      { $addToSet: { purchasedTools: toolId } }
    );
    return result.modifiedCount > 0;
  } catch {
    return false;
  }
}

export async function hasPurchase(userId: string, toolId: string): Promise<boolean> {
  try {
    const oid = new ObjectId(userId);
    const record = await usersCollection().findOne({ _id: oid, purchasedTools: toolId });
    return !!record;
  } catch {
    return false;
  }
}

// ============================================
// NIT functions
// ============================================

export async function getUserNits(userId: string): Promise<string[]> {
  try {
    const oid = new ObjectId(userId);
    const record = await usersCollection().findOne({ _id: oid }, { projection: { nits: 1 } });
    return record?.nits || [];
  } catch {
    return [];
  }
}

// ============================================
// Seed admin user (env-based, no hardcoded password)
// ============================================

export async function seedAdminUser(): Promise<void> {
  const adminEmail = process.env.ADMIN_EMAIL;
  const adminPassword = process.env.ADMIN_PASSWORD;
  const adminName = process.env.ADMIN_NAME || "Admin";
  const adminNits = process.env.ADMIN_NITS
    ? process.env.ADMIN_NITS.split(",").map((n) => n.trim()).filter(Boolean)
    : [];

  if (!adminEmail || !adminPassword) {
    console.log("ADMIN_EMAIL / ADMIN_PASSWORD not set - skipping admin seed.");
    return;
  }

  const existing = await getUserByEmail(adminEmail);

  if (!existing) {
    console.log(`Creating admin user: ${adminEmail}`);
    await createUser(adminEmail, adminName, adminPassword, true, adminNits);
  } else {
    // Ensure admin flag and sync NITs from env
    const updates: Record<string, unknown> = {};
    if (!existing.is_admin) updates.is_admin = true;
    if (adminNits.length > 0) updates.nits = adminNits;

    if (Object.keys(updates).length > 0) {
      await usersCollection().updateOne(
        { email: adminEmail.toLowerCase().trim() },
        { $set: updates }
      );
      if (updates.is_admin) console.log(`Promoted ${adminEmail} to admin.`);
      if (updates.nits) console.log(`Updated NITs for ${adminEmail}: ${adminNits.join(", ")}`);
    }
  }
}

export { db };
