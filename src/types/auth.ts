export type UserRole = "USER" | "ADMIN" | "DEMO";

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
  licenseStartDate?: string;
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
    companiesInPlan?: number;
    forcePasswordChange?: boolean;
    demo?: DemoAccess & {
      isExpired: boolean;
      remainingMs: number;
    };
  };
  temporaryPassword?: string;
}
