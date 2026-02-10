import Database from "better-sqlite3";
import bcrypt from "bcrypt";
import path from "path";
import { fileURLToPath } from "url";
import type { User, UserPurchase } from "../types/auth.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DB_PATH = path.join(__dirname, "../../data.db");

const db = new Database(DB_PATH);

// Inicializar tablas
db.exec(`
  CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    email TEXT UNIQUE NOT NULL,
    name TEXT NOT NULL,
    password_hash TEXT NOT NULL,
    is_admin INTEGER DEFAULT 0,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
  );

  CREATE TABLE IF NOT EXISTS user_purchases (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL,
    tool_id TEXT NOT NULL,
    purchased_at TEXT DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (user_id) REFERENCES users(id),
    UNIQUE(user_id, tool_id)
  );
`);

// ============================================
// User functions
// ============================================

export async function createUser(
  email: string,
  name: string,
  password: string,
  isAdmin = false
): Promise<User | null> {
  try {
    const hash = await bcrypt.hash(password, 10);
    const stmt = db.prepare(
      "INSERT INTO users (email, name, password_hash, is_admin) VALUES (?, ?, ?, ?)"
    );
    const result = stmt.run(email, name, hash, isAdmin ? 1 : 0);
    return getUserById(result.lastInsertRowid as number);
  } catch (err) {
    // Email duplicado u otro error
    console.error("Error creating user:", err);
    return null;
  }
}

export function getUserById(id: number): User | null {
  const stmt = db.prepare("SELECT * FROM users WHERE id = ?");
  return stmt.get(id) as User | null;
}

export function getUserByEmail(email: string): User | null {
  const stmt = db.prepare("SELECT * FROM users WHERE email = ?");
  return stmt.get(email) as User | null;
}

export async function verifyPassword(user: User, password: string): Promise<boolean> {
  return bcrypt.compare(password, user.password_hash);
}

// ============================================
// Purchase functions
// ============================================

export function getUserPurchases(userId: number): string[] {
  const stmt = db.prepare("SELECT tool_id FROM user_purchases WHERE user_id = ?");
  const rows = stmt.all(userId) as Array<{ tool_id: string }>;
  return rows.map((r) => r.tool_id);
}

export function addPurchase(userId: number, toolId: string): boolean {
  try {
    const stmt = db.prepare(
      "INSERT OR IGNORE INTO user_purchases (user_id, tool_id) VALUES (?, ?)"
    );
    stmt.run(userId, toolId);
    return true;
  } catch {
    return false;
  }
}

export function hasPurchase(userId: number, toolId: string): boolean {
  const stmt = db.prepare(
    "SELECT 1 FROM user_purchases WHERE user_id = ? AND tool_id = ?"
  );
  return stmt.get(userId, toolId) !== undefined;
}

// ============================================
// Seed admin user
// ============================================

export async function seedAdminUser(): Promise<void> {
  const adminEmail = "info@contago.com.co";
  const existing = getUserByEmail(adminEmail);
  
  if (!existing) {
    console.log("Creating admin user: info@contago.com.co");
    // Password temporal - cambiar en producción
    await createUser(adminEmail, "ContaGO Admin", "ContaGO2024!", true);
  } else if (!existing.is_admin) {
    // Asegurar que sea admin
    const stmt = db.prepare("UPDATE users SET is_admin = 1 WHERE email = ?");
    stmt.run(adminEmail);
  }
}

export { db };
