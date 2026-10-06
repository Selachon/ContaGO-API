import fs from "fs";
import path from "path";
import { Binary } from "mongodb";
import { getDb } from "./database.js";

/**
 * RUT (PDF) de cada usuario. Se guarda en MongoDB (colección userRuts, _id = id de usuario) porque el
 * contenedor del API no tiene volumen: lo que se escribe en disco se pierde en cada despliegue.
 * El disco (data/user-ruts) se mantiene solo como respaldo de lectura para archivos anteriores.
 */
const COLLECTION = "userRuts";
const LEGACY_DIR = path.resolve("data/user-ruts");

interface RutDoc {
  _id: string;
  data: Binary;
  originalName: string;
  size: number;
  uploadedAt: string;
}

const col = () => getDb().collection<RutDoc>(COLLECTION);

export async function saveRut(userId: string, buffer: Buffer, originalName: string): Promise<void> {
  await col().updateOne(
    { _id: userId },
    { $set: { data: new Binary(buffer), originalName, size: buffer.length, uploadedAt: new Date().toISOString() } },
    { upsert: true }
  );
}

/** Devuelve el PDF del RUT. Si solo existe en el disco viejo, lo copia a MongoDB al leerlo. */
export async function getRut(userId: string, legacyFilename?: string): Promise<{ data: Buffer; originalName?: string } | null> {
  const doc = await col().findOne({ _id: userId });
  if (doc) return { data: Buffer.from(doc.data.buffer), originalName: doc.originalName };
  if (!legacyFilename) return null;
  const legacyPath = path.join(LEGACY_DIR, path.basename(legacyFilename));
  if (!fs.existsSync(legacyPath)) return null;
  const data = fs.readFileSync(legacyPath);
  await saveRut(userId, data, legacyFilename).catch(() => {});
  return { data };
}

/** Ids de usuario cuyo PDF de RUT está realmente disponible (MongoDB o disco viejo). */
export async function rutAvailability(users: { id: string; filename?: string }[]): Promise<Set<string>> {
  const ids = users.map((u) => u.id);
  const inDb = await col().find({ _id: { $in: ids } }, { projection: { _id: 1 } }).toArray();
  const available = new Set(inDb.map((d) => d._id));
  for (const u of users) {
    if (!available.has(u.id) && u.filename && fs.existsSync(path.join(LEGACY_DIR, path.basename(u.filename)))) available.add(u.id);
  }
  return available;
}
