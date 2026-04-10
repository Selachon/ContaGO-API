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

export function createRequireIntegrationAuth(jwtAuth: RequestHandler = requireAuth): RequestHandler {
  return (req: Request, res: Response, next: NextFunction) => {
    const token = getRequestToken(req);

    const internalApiKey = getInternalApiKey();
    if (token && internalApiKey && token === internalApiKey) {
      req.integrationAuthMode = "internal_api_key";
      return next();
    }

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
