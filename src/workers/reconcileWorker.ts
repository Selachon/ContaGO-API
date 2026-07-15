import { parentPort } from "worker_threads";
import { reconcile } from "../services/bankReconciliationService.js";
import type { RecStatementInput, RecLedgerInput, RecOptions } from "../services/bankReconciliationService.js";

parentPort!.once("message", (msg: { statement: RecStatementInput; ledger: RecLedgerInput; options: RecOptions }) => {
  try {
    const result = reconcile(msg.statement, msg.ledger, msg.options);
    parentPort!.postMessage({ ok: true, result });
  } catch (err) {
    parentPort!.postMessage({ ok: false, error: err instanceof Error ? err.message : String(err) });
  }
});
