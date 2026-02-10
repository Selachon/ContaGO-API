export interface User {
  id: string;
  email: string;
  name: string;
  password_hash: string;
  is_admin: boolean;
  created_at: string;
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
    purchasedTools: string[];
  };
}
