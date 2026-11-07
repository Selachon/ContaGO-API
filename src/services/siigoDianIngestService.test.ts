import test from "node:test";
import assert from "node:assert/strict";
import { oldestIfDescending } from "./siigoDianIngestService.js";

const rows = (...dates: string[]) => dates.map((issueDate) => ({ issueDate }));

test("oldestIfDescending: devuelve la fecha más vieja cuando vienen de más nueva a más vieja", () => {
  assert.equal(oldestIfDescending(rows("31/10/2026", "31/10/2026", "15/10/2026", "01/10/2026")), "2026-10-01");
});

test("oldestIfDescending: si el orden no es descendente no se puede confiar", () => {
  assert.equal(oldestIfDescending(rows("10/10/2026", "20/10/2026")), "");
});

test("oldestIfDescending: una fecha ausente o con formato inválido invalida el resultado", () => {
  assert.equal(oldestIfDescending(rows("31/10/2026", "")), "");
  assert.equal(oldestIfDescending(rows("2026-10-31")), "");
});

test("oldestIfDescending: sin filas no hay fecha", () => {
  assert.equal(oldestIfDescending([]), "");
});
