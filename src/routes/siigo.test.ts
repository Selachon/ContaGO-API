import { beforeEach, afterEach, describe, it } from "node:test";
import assert from "node:assert/strict";
import express, { type RequestHandler } from "express";
import request from "supertest";
import { createSiigoRouter } from "./siigo.js";
import { resetSiigoTokenCache } from "../services/siigoService.js";
import { createRequireIntegrationAuth } from "../middleware/requireIntegrationAuth.js";

interface EnvSnapshot {
  SIIGO_API_BASE_URL?: string;
  SIIGO_PARTNER_ID?: string;
  SIIGO_USERNAME?: string;
  SIIGO_ACCESS_KEY?: string;
  GPT_INTERNAL_API_KEY?: string;
}

const authStub: RequestHandler = (req, _res, next) => {
  req.user = {
    userId: "test-user-id",
    email: "test@example.com",
    isAdmin: false,
  };
  next();
};

function buildApp() {
  const app = express();
  app.use(express.json());
  app.use("/integrations/siigo", createSiigoRouter(authStub));
  return app;
}

function buildAppWithIntegrationAuth(jwtAuth: RequestHandler) {
  const app = express();
  app.use(express.json());
  app.use("/integrations/siigo", createSiigoRouter(createRequireIntegrationAuth(jwtAuth)));
  return app;
}

function setSiigoEnv(overrides: Partial<EnvSnapshot> = {}): void {
  process.env.SIIGO_API_BASE_URL = overrides.SIIGO_API_BASE_URL ?? "https://api.siigo.com";
  process.env.SIIGO_PARTNER_ID = overrides.SIIGO_PARTNER_ID ?? "SentiidoAI";
  process.env.SIIGO_USERNAME = overrides.SIIGO_USERNAME ?? "sandbox@siigoapi.com";
  process.env.SIIGO_ACCESS_KEY = overrides.SIIGO_ACCESS_KEY ?? "fake-access-key";
  process.env.GPT_INTERNAL_API_KEY = overrides.GPT_INTERNAL_API_KEY ?? "gpt-internal-key";
}

function jsonResponse(status: number, payload: unknown, headers: Record<string, string> = {}): Response {
  return new Response(JSON.stringify(payload), {
    status,
    headers: {
      "content-type": "application/json",
      ...headers,
    },
  });
}

const envSnapshot: EnvSnapshot = {
  SIIGO_API_BASE_URL: process.env.SIIGO_API_BASE_URL,
  SIIGO_PARTNER_ID: process.env.SIIGO_PARTNER_ID,
  SIIGO_USERNAME: process.env.SIIGO_USERNAME,
  SIIGO_ACCESS_KEY: process.env.SIIGO_ACCESS_KEY,
  GPT_INTERNAL_API_KEY: process.env.GPT_INTERNAL_API_KEY,
};

describe("Siigo integration routes", () => {
  beforeEach(() => {
    setSiigoEnv();
    resetSiigoTokenCache();
  });

  afterEach(() => {
    resetSiigoTokenCache();
    process.env.SIIGO_API_BASE_URL = envSnapshot.SIIGO_API_BASE_URL;
    process.env.SIIGO_PARTNER_ID = envSnapshot.SIIGO_PARTNER_ID;
    process.env.SIIGO_USERNAME = envSnapshot.SIIGO_USERNAME;
    process.env.SIIGO_ACCESS_KEY = envSnapshot.SIIGO_ACCESS_KEY;
    process.env.GPT_INTERNAL_API_KEY = envSnapshot.GPT_INTERNAL_API_KEY;
    global.fetch = originalFetch;
  });

  it("allows access with JWT via integration auth middleware", async () => {
    const jwtPassStub: RequestHandler = (req, _res, next) => {
      req.user = {
        userId: "jwt-user",
        email: "jwt@example.com",
        isAdmin: false,
      };
      next();
    };

    const app = buildAppWithIntegrationAuth(jwtPassStub);
    const response = await request(app)
      .get("/integrations/siigo/health")
      .set("Authorization", "Bearer jwt-valid-token");

    assert.equal(response.status, 200);
    assert.equal(response.body.ok, true);
    assert.equal(response.body.authMode, "jwt");
  });

  it("allows access with GPT_INTERNAL_API_KEY", async () => {
    const jwtFailStub: RequestHandler = (_req, res) => {
      res.status(401).json({ ok: false, message: "Token inválido o expirado" });
    };

    const app = buildAppWithIntegrationAuth(jwtFailStub);
    const response = await request(app)
      .get("/integrations/siigo/health")
      .set("Authorization", "Bearer gpt-internal-key");

    assert.equal(response.status, 200);
    assert.equal(response.body.ok, true);
    assert.equal(response.body.authMode, "internal_api_key");
  });

  it("rejects invalid bearer token when not JWT or internal key", async () => {
    const jwtFailStub: RequestHandler = (_req, res) => {
      res.status(401).json({ ok: false, message: "Token inválido o expirado" });
    };

    const app = buildAppWithIntegrationAuth(jwtFailStub);
    const response = await request(app)
      .get("/integrations/siigo/health")
      .set("Authorization", "Bearer invalid-token");

    assert.equal(response.status, 401);
    assert.equal(response.body.ok, false);
  });

  it("GET /health reports missing configuration without exposing secrets", async () => {
    process.env.SIIGO_USERNAME = "";
    process.env.SIIGO_ACCESS_KEY = "";

    const app = buildApp();
    const response = await request(app).get("/integrations/siigo/health");

    assert.equal(response.status, 500);
    assert.equal(response.body.ok, false);
    assert.deepEqual(response.body.missing, ["SIIGO_USERNAME", "SIIGO_ACCESS_KEY"]);
    assert.equal(response.body.partnerId, "SentiidoAI");
    assert.equal(response.body.tokenCached, false);
    assert.equal(response.body.access_key, undefined);
    assert.equal(response.body.access_token, undefined);
  });

  it("POST /auth authenticates without returning access token", async () => {
    global.fetch = async () => jsonResponse(200, { access_token: "siigo-token", expires_in: 3600 });

    const app = buildApp();
    const response = await request(app).post("/integrations/siigo/auth");

    assert.equal(response.status, 200);
    assert.equal(response.body.ok, true);
    assert.equal(response.body.authenticated, true);
    assert.ok(response.body.expiresAt);
    assert.equal(response.body.access_token, undefined);
    assert.equal(response.body.token, undefined);
  });

  it("GET /invoices forwards Partner-Id and query params", async () => {
    const fetchCalls: Array<{ url: string; init?: RequestInit }> = [];
    global.fetch = async (url, init) => {
      fetchCalls.push({ url: String(url), init });
      if (String(url).endsWith("/auth")) {
        return jsonResponse(200, { access_token: "siigo-token", expires_in: 3600 });
      }
      return jsonResponse(200, { results: [], pagination: { page: 1 } });
    };

    const app = buildApp();
    const response = await request(app).get(
      "/integrations/siigo/invoices?created_start=2025-01-01&page=1&page_size=25"
    );

    assert.equal(response.status, 200);
    assert.equal(response.body.ok, true);
    assert.equal(fetchCalls.length, 2);
    assert.ok(fetchCalls[1].url.includes("/v1/invoices"));
    assert.ok(fetchCalls[1].url.includes("created_start=2025-01-01"));
    assert.ok(fetchCalls[1].url.includes("page_size=25"));
    const partnerHeader = (fetchCalls[1].init?.headers as Record<string, string>)["Partner-Id"];
    assert.equal(partnerHeader, "SentiidoAI");
  });

  it("retries once after 401 and then succeeds", async () => {
    let call = 0;
    global.fetch = async (url) => {
      call += 1;
      if (String(url).endsWith("/auth") && call === 1) {
        return jsonResponse(200, { access_token: "token-1", expires_in: 3600 });
      }
      if (String(url).includes("/v1/invoices") && call === 2) {
        return jsonResponse(401, { message: "Unauthorized" });
      }
      if (String(url).endsWith("/auth") && call === 3) {
        return jsonResponse(200, { access_token: "token-2", expires_in: 3600 });
      }
      if (String(url).includes("/v1/invoices") && call === 4) {
        return jsonResponse(200, { results: [{ id: "inv-1" }] });
      }
      return jsonResponse(500, { message: "Unexpected call" });
    };

    const app = buildApp();
    const response = await request(app).get("/integrations/siigo/invoices");

    assert.equal(response.status, 200);
    assert.equal(response.body.ok, true);
    assert.equal(call, 4);
  });

  it("returns clean error when Siigo config is missing", async () => {
    process.env.SIIGO_USERNAME = "";

    const app = buildApp();
    const response = await request(app).get("/integrations/siigo/invoices");

    assert.equal(response.status, 500);
    assert.equal(response.body.ok, false);
    assert.equal(response.body.code, "siigo_config_missing");
    assert.ok(String(response.body.message).includes("Configuración incompleta"));
  });
});

const originalFetch = global.fetch;
