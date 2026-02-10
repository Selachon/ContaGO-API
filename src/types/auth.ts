export interface User {
  id: number;
  email: string;
  name: string;
  password_hash: string;
  is_admin: boolean;
  created_at: string;
}

export interface UserPurchase {
  id: number;
  user_id: number;
  tool_id: string;
  purchased_at: string;
}

export interface JWTPayload {
  userId: number;
  email: string;
  isAdmin: boolean;
}

export interface AuthResponse {
  ok: boolean;
  message?: string;
  token?: string;
  user?: {
    id: number;
    email: string;
    name: string;
    isAdmin: boolean;
    purchasedTools: string[];
  };
}
