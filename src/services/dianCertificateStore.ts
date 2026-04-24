import { ObjectId, type Collection, type Filter } from "mongodb";
import type {
  DianCertificateCredentials,
  DianEnvironment,
} from "../dian/types/DianResponse.js";
import { db } from "./database.js";
import { decryptToken, encryptToken } from "../utils/encryption.js";

interface DianCertificateRecord {
  _id: ObjectId;
  nit: string;
  environment: DianEnvironment;
  p12_path: string;
  encrypted_p12_password: string;
  enabled: boolean;
  created_at: string;
  updated_at: string;
  created_by?: string;
  updated_by?: string;
}

interface UpsertDianCertificateInput {
  nit: string;
  p12Path: string;
  p12Password: string;
  environment: DianEnvironment;
  enabled?: boolean;
  updatedBy?: string;
}

function dianCertificatesCollection(): Collection<DianCertificateRecord> {
  if (!db) {
    throw new Error("MongoDB no está conectado");
  }
  return db.collection<DianCertificateRecord>("dian_certificates");
}

export function normalizeNit(nit: string): string {
  return nit.replace(/[^\d]/g, "").trim();
}

export async function upsertDianCertificateConfig(
  input: UpsertDianCertificateInput
): Promise<boolean> {
  const normalizedNit = normalizeNit(input.nit);
  if (!normalizedNit) {
    throw new Error("NIT inválido para configuración de certificado DIAN");
  }
  if (!input.p12Path.trim()) {
    throw new Error("p12Path es requerido");
  }
  if (!input.p12Password.trim()) {
    throw new Error("p12Password es requerido");
  }

  const now = new Date().toISOString();
  const encryptedPassword = encryptToken(input.p12Password);

  const result = await dianCertificatesCollection().updateOne(
    {
      nit: normalizedNit,
      environment: input.environment,
    },
    {
      $set: {
        p12_path: input.p12Path,
        encrypted_p12_password: encryptedPassword,
        enabled: input.enabled ?? true,
        updated_at: now,
        updated_by: input.updatedBy,
      },
      $setOnInsert: {
        created_at: now,
        created_by: input.updatedBy,
      },
    },
    { upsert: true }
  );

  return result.acknowledged;
}

export async function getDianCertificateCredentialsByNit(
  nit: string,
  environment: DianEnvironment
): Promise<DianCertificateCredentials | null> {
  const normalizedNit = normalizeNit(nit);
  if (!normalizedNit) {
    return null;
  }

  const query = {
    nit: normalizedNit,
    enabled: true,
    $or: [{ environment }, { environment: { $exists: false } }],
  } as Filter<DianCertificateRecord>;

  const record = await dianCertificatesCollection().findOne(query, {
    sort: { updated_at: -1 },
  });
  if (!record) {
    return null;
  }

  return {
    p12Path: record.p12_path,
    p12Password: decryptToken(record.encrypted_p12_password),
    environment: record.environment ?? environment,
  };
}

export async function hasDianCertificateConfigured(
  nit: string,
  environment: DianEnvironment
): Promise<boolean> {
  const credentials = await getDianCertificateCredentialsByNit(nit, environment);
  return credentials !== null;
}
