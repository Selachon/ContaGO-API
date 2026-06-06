/**
 * Cuentas bancarias por empresa para la herramienta de Causación de Egresos.
 * Cada empresa puede tener varias (Itaú, Bancolombia, …); los movimientos
 * importados se asocian a una cuenta para conciliar por separado.
 */
import { ObjectId } from "mongodb";
import { getDb } from "./database.js";

const COLLECTION = "siigoBankAccounts";

export interface BankAccount {
  id: string;
  companyId: string;
  name: string;
  bank: string;
  number: string;
  accountCode: string; // cuenta contable del banco (para el avanzado)
  createdAt: string;
}

function mapDoc(d: any): BankAccount {
  return {
    id: d._id.toString(),
    companyId: d.companyId,
    name: d.name || "",
    bank: d.bank || "",
    number: d.number || "",
    accountCode: d.accountCode || "",
    createdAt: d.createdAt || "",
  };
}

export async function listBankAccounts(companyId: string): Promise<BankAccount[]> {
  if (!companyId) return [];
  const docs = await getDb()
    .collection<any>(COLLECTION)
    .find({ companyId })
    .sort({ createdAt: 1 })
    .toArray();
  return docs.map(mapDoc);
}

export async function createBankAccount(
  companyId: string,
  data: { name: string; bank?: string; number?: string; accountCode?: string }
): Promise<BankAccount> {
  const name = String(data?.name || "").trim();
  if (!companyId) throw new Error("Falta la empresa.");
  if (!name) throw new Error("El nombre de la cuenta es obligatorio.");
  const doc = {
    companyId,
    name,
    bank: String(data.bank || "").trim(),
    number: String(data.number || "").trim(),
    accountCode: String(data.accountCode || "").replace(/\D/g, ""),
    createdAt: new Date().toISOString(),
  };
  const r = await getDb().collection<any>(COLLECTION).insertOne(doc);
  return mapDoc({ ...doc, _id: r.insertedId });
}

const EDITABLE = ["name", "bank", "number", "accountCode"];

export async function updateBankAccount(
  companyId: string,
  id: string,
  patch: Record<string, unknown>
): Promise<BankAccount | null> {
  if (!companyId || !ObjectId.isValid(id)) return null;
  const set: Record<string, unknown> = {};
  for (const f of EDITABLE) {
    if (f in patch) set[f] = f === "accountCode" ? String(patch[f] ?? "").replace(/\D/g, "") : String(patch[f] ?? "").trim();
  }
  const col = getDb().collection<any>(COLLECTION);
  await col.updateOne({ _id: new ObjectId(id), companyId }, { $set: set });
  const doc = await col.findOne({ _id: new ObjectId(id), companyId });
  return doc ? mapDoc(doc) : null;
}

/** Borra la cuenta y devuelve cuántos movimientos quedaron huérfanos (informativo). */
export async function deleteBankAccount(companyId: string, id: string): Promise<boolean> {
  if (!companyId || !ObjectId.isValid(id)) return false;
  const r = await getDb().collection<any>(COLLECTION).deleteOne({ _id: new ObjectId(id), companyId });
  return r.deletedCount > 0;
}
