import test from "node:test";
import assert from "node:assert/strict";
import { validateDianUrl } from "./validateDianUrl.js";

function run(token_url: unknown) {
  const req: any = { body: { token_url } };
  const result: any = {};
  const res: any = {
    status(c: number) { result.status = c; return this; },
    json(b: unknown) { result.body = b; return this; },
  };
  let nexted = false;
  validateDianUrl(req, res, () => { nexted = true; });
  return { nexted, result, req };
}

test("token de habilitación (hab): se rechaza con mensaje claro, no sigue al scraper", () => {
  const { nexted, result } = run("https://catalogo-vpfe-hab.dian.gov.co/User/AuthToken?pk=10910094%7C93389781&rk=93389781&token=00000000-0000-4000-8000-000000000000");
  assert.equal(nexted, false);
  assert.equal(result.status, 400);
  assert.match(result.body.detalle, /PRUEBAS/);
  assert.match(result.body.detalle, /catalogo-vpfe\.dian\.gov\.co/);
});

test("token real de producción: pasa", () => {
  const { nexted, req } = run("https://catalogo-vpfe.dian.gov.co/User/AuthToken?pk=10910094%7C901965856&rk=901965856&token=abc");
  assert.equal(nexted, true);
  assert.equal(req.body.token_url, "https://catalogo-vpfe.dian.gov.co/User/AuthToken?pk=10910094%7C901965856&rk=901965856&token=abc");
});

test("token de hab envuelto en un SafeLink de Outlook (caso real reportado): también se rechaza con el mismo mensaje", () => {
  const safelink = "https://na01.safelinks.protection.outlook.com/?url=" + encodeURIComponent("https://catalogo-vpfe-hab.dian.gov.co/User/AuthToken?pk=10910094%7C93389781&rk=93389781&token=00000000-0000-4000-8000-000000000000") + "&data=x";
  const { nexted, result } = run(safelink);
  assert.equal(nexted, false);
  assert.equal(result.status, 400);
  assert.match(result.body.detalle, /PRUEBAS/);
});

test("otro host cualquiera: sigue rechazado como antes (SSRF)", () => {
  const { nexted, result } = run("https://evil.example.com/User/AuthToken");
  assert.equal(nexted, false);
  assert.equal(result.status, 400);
  assert.match(result.body.detalle, /dominio de la DIAN/);
});
