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

function getIntegrationToken(req: Request): string | null {
  const authHeader = req.headers.authorization;
  if (typeof authHeader === "string" && authHeader.trim()) {
    const bearer = authHeader.match(/^Bearer\s+(.+)$/i);
    if (bearer?.[1]?.trim()) return bearer[1].trim();

    // Compatibilidad con clientes que envían Authorization sin prefijo Bearer.
    return authHeader.trim();
  }

  const xApiKey = req.headers["x-api-key"];
  if (typeof xApiKey === "string" && xApiKey.trim()) {
    return xApiKey.trim();
  }

  if (Array.isArray(xApiKey) && typeof xApiKey[0] === "string" && xApiKey[0].trim()) {
    return xApiKey[0].trim();
  }

  return getRequestToken(req);
}

export function createRequireIntegrationAuth(jwtAuth: RequestHandler = requireAuth): RequestHandler {
  return (req: Request, res: Response, next: NextFunction) => {
    const token = getIntegrationToken(req);
    const hasAuthorizationHeader = typeof req.headers.authorization === "string";
    const hasXApiKey = typeof req.headers["x-api-key"] === "string" || Array.isArray(req.headers["x-api-key"]);
    console.log(
      `[IntegrationAuth] path=${req.originalUrl || req.path} hasAuthorization=${hasAuthorizationHeader} hasXApiKey=${hasXApiKey} token=${maskToken(token)}`
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
