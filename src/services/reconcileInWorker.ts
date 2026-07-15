import { Worker } from "worker_threads";
import { fileURLToPath } from "url";
import path from "path";
import type { RecStatementInput, RecLedgerInput, RecOptions, ReconciliationResult } from "./bankReconciliationService.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
// En dist/services/ → apunta a dist/workers/reconcileWorker.js
const WORKER_PATH = path.join(__dirname, "../workers/reconcileWorker.js");

export function reconcileInWorker(
  statement: RecStatementInput,
  ledger: RecLedgerInput,
  options: RecOptions = {},
): Promise<ReconciliationResult> {
  return new Promise((resolve, reject) => {
    const worker = new Worker(WORKER_PATH);

    const timeout = setTimeout(() => {
      worker.terminate();
      reject(new Error("La conciliación tardó demasiado y fue cancelada. Intenta con un extracto de menor tamaño."));
    }, 120_000);

    worker.once("message", (msg: { ok: boolean; result?: ReconciliationResult; error?: string }) => {
      clearTimeout(timeout);
      worker.terminate();
      if (msg.ok && msg.result) resolve(msg.result);
      else reject(new Error(msg.error ?? "Error en el worker de conciliación."));
    });

    worker.once("error", (err) => {
      clearTimeout(timeout);
      worker.terminate();
      reject(err);
    });

    worker.postMessage({ statement, ledger, options });
  });
}
