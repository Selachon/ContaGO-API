export type UserRole = "USER" | "ADMIN" | "DEMO";

// Datos de facturación electrónica del usuario/cliente. No siempre coinciden
// con el nombre/email de la cuenta del portal (p.ej. la razón social de la
// empresa que factura vs. el contacto que usa la herramienta).
export interface BillingInfo {
  idType: "NIT" | "CC" | "CE" | "PA" | "PEP";
  id: string;
  isCompany: boolean;
  firstName?: string;
  lastName?: string;
  companyName?: string;
  email: string;
  city: string;
  address: string;
  rutFile?: {
    filename: string;
    originalName: string;
    uploadedAt: string;
  };
}

export interface DemoAccess {
  nit: string;
  normalizedNit: string;
  toolId: string;
  inviteId: string;
  startedAt: string;
  expiresAt: string;
  trialLimit: number;
}

export interface User {
  id: string;
  email: string;
  name: string;
  password_hash: string;
  is_admin: boolean;
  role: UserRole;
  nits: string[];
  status?: "active" | "suspended";
  force_password_change?: boolean;
  created_at: string;
  demo?: DemoAccess;
  companiesInPlan?: number;
  toolCompanyLimits?: Record<string, number>;
  toolNits?: Record<string, string[]>;
  licenseStartDate?: string;
  billing?: BillingInfo;
}

export interface UserPurchase {
  id: number;
  user_id: string;
  tool_id: string;
  purchased_at: string;
}

export interface JWTPayload {
  userId: string;
  email: string;
  isAdmin: boolean;
  role?: UserRole;
  deviceId?: string;
}

export interface AuthResponse {
  ok: boolean;
  message?: string;
  token?: string;
  user?: {
    id: string;
    email: string;
    name: string;
    isAdmin: boolean;
    role?: UserRole;
    purchasedTools: string[];
    nits: string[];
    toolNits?: Record<string, string[]>;
    companiesInPlan?: number;
    toolCompanyLimits?: Record<string, number>;
    forcePasswordChange?: boolean;
    demo?: DemoAccess & {
      isExpired: boolean;
      remainingMs: number;
    };
  };
  temporaryPassword?: string;
}
