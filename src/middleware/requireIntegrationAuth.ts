import type { NextFunction, Request, RequestHandler, Response } from "express";
import { getRequestToken, requireAuth } from "./auth.js";

export type IntegrationAuthMode = "internal_api_key" | "jwt";

declare global {
  namespace Express {
    interface Request {
      integrationAuthMode?: IntegrationAuthMode;
    }
  }
}

function getInternalApiKey(): string {
  const raw = process.env.GPT_INTERNAL_API_KEY;
  return typeof raw === "string" ? raw.trim() : "";
}

function maskToken(token: string | null): string {
  if (!token) return "[none]";
  if (token.length <= 8) return `${token.slice(0, 2)}***`;
  return `${token.slice(0, 4)}...${token.slice(-2)}`;
}

export function createRequireIntegrationAuth(jwtAuth: RequestHandler = requireAuth): RequestHandler {
  return (req: Request, res: Response, next: NextFunction) => {
    const token = getRequestToken(req);
    const hasAuthorizationHeader = typeof req.headers.authorization === "string";
    console.log(
      `[IntegrationAuth] path=${req.originalUrl || req.path} hasAuthorization=${hasAuthorizationHeader} token=${maskToken(token)}`
    );

    const internalApiKey = getInternalApiKey();
    if (token && internalApiKey && token === internalApiKey) {
      req.integrationAuthMode = "internal_api_key";
      console.log(`[IntegrationAuth] path=${req.originalUrl || req.path} authenticated=internal_api_key`);
      return next();
    }

    console.log(`[IntegrationAuth] path=${req.originalUrl || req.path} fallback=jwt`);

    const wrappedNext: NextFunction = (err?: unknown) => {
      if (!err) {
        req.integrationAuthMode = "jwt";
      }
      next(err);
    };

    return jwtAuth(req, res, wrappedNext);
  };
}

export const requireIntegrationAuth = createRequireIntegrationAuth();
