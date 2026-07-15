/**
 * Motor de conciliación bancaria: cruza el EXTRACTO del banco contra el LIBRO
 * AUXILIAR de la contabilidad (misma cuenta y periodo) y arma el cuadre.
 *
 * Estrategia (acordada con el usuario):
 *  1) Cruce 1:1 por MONTO dentro de tolerancia, separando por dirección
 *     (ingreso↔débito, egreso↔crédito). La fecha se usa como desempate, no
 *     como filtro. Maneja multiplicidad: 3 pagos de $100k en extracto vs 2 en
 *     contabilidad → casa 2, deja 1 como partida (y marca el grupo ambiguo).
 *  2) Subset-sum sobre lo que quedó suelto: varios renglones del extracto que
 *     sumen un renglón de la contabilidad (comisiones desglosadas vs
 *     consolidadas) y viceversa.
 *  3) Lo que siga suelto = partidas conciliatorias en 4 categorías.
 *
 * Cuadre: usa la identidad de saldos
 *   SaldoFinalExt − SaldoFinalCont = (SaldoIniExt − SaldoIniCont)
 *                                    + (netoExt − netoCont)
 * y como los renglones casados se cancelan entre sí, la diferencia queda
 * EXPLICADA por las partidas conciliatorias + el residual por tolerancia.
 */

export type Dir = "in" | "out";

export interface RecMovementInput {
  date: string;
  description: string;
  value: number;
  direction: Dir;
  kind?: string; // ingreso | egreso | bank_fee (del extracto), opcional
}

export interface RecEntryInput {
  date: string;
  description: string;
  value: number;
  direction: Dir;
  voucher?: string;
  nit?: string;
  thirdParty?: string;
}

export interface RecStatementInput {
  opening: number | null;
  closing: number | null;
  movements: RecMovementInput[];
}

export interface RecLedgerInput {
  opening: number | null;
  closing: number | null;
  entries: RecEntryInput[];
}

export interface RecOptions {
  tolerance?: number; // pesos, default 100
  maxSubsetSize?: number; // tope de elementos por grupo en subset-sum, default 12
  ignoredIds?: string[]; // ids a excluir del cruce (el usuario los marca aparte)
  manualMatches?: { statementIds: string[]; ledgerIds: string[] }[]; // cruces forzados por el usuario
}

export interface RecItem {
  id: string;
  side: "statement" | "ledger";
  date: string;
  description: string;
  value: number;
  direction: Dir;
  kind?: string;
  voucher?: string;
  nit?: string;
  thirdParty?: string;
}

export interface RecMatch {
  type: "1:1" | "group" | "manual";
  direction: Dir;
  statement: RecItem[];
  ledger: RecItem[];
  valueStatement: number;
  valueLedger: number;
  residual: number; // valueStatement - valueLedger
  ambiguous: boolean; // hubo varios candidatos del mismo monto
}

export type PartidaCategory =
  | "ingresos_no_contabilizados"
  | "egresos_no_contabilizados"
  | "ingresos_contab_sin_extracto"
  | "egresos_contab_sin_extracto";

export interface Cuadre {
  openingStatement: number | null;
  openingLedger: number | null;
  openingDiff: number;
  closingStatement: number | null;
  closingLedger: number | null;
  closingDiff: number;
  partidas: {
    ingresos_no_contabilizados: number;
    egresos_no_contabilizados: number;
    ingresos_contab_sin_extracto: number;
    egresos_contab_sin_extracto: number;
  };
  toleranceAdjustment: number; // residual neto de los renglones casados
  explained: number; // lo que explican partidas + ajuste + dif. saldo inicial
  unexplained: number; // closingDiff - explained (debería ≈ 0)
  balanced: boolean;
}

export interface ReconciliationResult {
  tolerance: number;
  matches: RecMatch[];
  partidas: Record<PartidaCategory, RecItem[]>;
  cuadre: Cuadre;
  counts: {
    statement: number;
    ledger: number;
    matched1to1: number;
    grouped: number;
    manual: number;
    unmatchedStatement: number;
    unmatchedLedger: number;
  };
  ignored: { statement: RecItem[]; ledger: RecItem[] };
}

const EPS = 0.01;

const round2 = (n: number): number => Math.round(n * 100) / 100;

const BANK_STATEMENT_GROUP_RX =
  /impto gobierno 4x1000|4\s*x\s*1000|gmf|cuota manejo suc virt empresa|servicio pago a otros bancos|iva cuota manejo suc virt emp|cobro iva pagos automaticos|iva boton|abono intereses ahorros|comision boton|servicio por pagos a nequi|servicio pago a proveedores|servicio pago de nomina/i;

const BANK_LEDGER_GROUP_RX =
  /gastos bancarios|intereses y gastos bancarios|comisiones bancarias|cuota manejo|4\s*x\s*1000|gmf|gravamen|servicios bancarios/i;

const isStatementBankGroupItem = (item: RecItem): boolean =>
  item.kind === "bank_fee" || BANK_STATEMENT_GROUP_RX.test(item.description);

const isLedgerBankGroupItem = (item: RecItem): boolean =>
  BANK_LEDGER_GROUP_RX.test([item.description, item.thirdParty || "", item.voucher || ""].join(" ")) ||
  // Comprobantes de Contabilidad (CC-) son asientos de consolidación mensual de
  // gastos bancarios e intereses; no tienen descripción canónica pero sí agrupan.
  /^CC-/i.test(item.voucher ?? "");

// ─── Normalización a items con id estable ────────────────────────────────
function toItems(statement: RecStatementInput, ledger: RecLedgerInput): {
  stmt: RecItem[];
  led: RecItem[];
} {
  const stmt: RecItem[] = statement.movements.map((m, i) => ({
    id: `s${i}`,
    side: "statement",
    date: m.date,
    description: m.description,
    value: round2(Math.abs(m.value)),
    direction: m.direction,
    kind: m.kind,
  }));
  const led: RecItem[] = ledger.entries.map((e, i) => ({
    id: `l${i}`,
    side: "ledger",
    date: e.date,
    description: e.description,
    value: round2(Math.abs(e.value)),
    direction: e.direction,
    voucher: e.voucher,
    nit: e.nit,
    thirdParty: e.thirdParty,
  }));
  return { stmt, led };
}

const dayDiff = (a: string, b: string): number => {
  if (!a || !b) return 999;
  const da = Date.parse(a);
  const db = Date.parse(b);
  if (Number.isNaN(da) || Number.isNaN(db)) return 999;
  return Math.abs(da - db) / 86400000;
};

// ─── Fase 1: cruce 1:1 por monto (tolerancia), fecha como desempate ───────
function matchOneToOne(
  stmt: RecItem[],
  led: RecItem[],
  tol: number
): { matches: RecMatch[]; stmtLeft: RecItem[]; ledLeft: RecItem[] } {
  const matches: RecMatch[] = [];
  const usedLed = new Set<string>();
  const usedStmt = new Set<string>();

  for (const dir of ["in", "out"] as Dir[]) {
    const ss = stmt.filter((s) => s.direction === dir);
    const ll = led.filter((l) => l.direction === dir);

    // Orden estable: por valor luego fecha → empareja primero los montos grandes.
    ss.sort((a, b) => b.value - a.value || a.date.localeCompare(b.date));

    for (const s of ss) {
      if (usedStmt.has(s.id)) continue;
      const candidates = ll.filter((l) => !usedLed.has(l.id) && Math.abs(l.value - s.value) <= tol);
      if (candidates.length === 0) continue;
      // Mejor candidato: menor diferencia de valor, luego fecha más cercana.
      candidates.sort(
        (a, b) =>
          Math.abs(a.value - s.value) - Math.abs(b.value - s.value) ||
          dayDiff(a.date, s.date) - dayDiff(b.date, s.date)
      );
      const l = candidates[0];
      usedStmt.add(s.id);
      usedLed.add(l.id);
      // Ambiguo si había más de un candidato con prácticamente el mismo monto.
      const ambiguous =
        candidates.filter((c) => Math.abs(c.value - l.value) <= tol).length > 1 ||
        ss.filter((o) => Math.abs(o.value - s.value) <= tol).length > 1;
      matches.push({
        type: "1:1",
        direction: dir,
        statement: [s],
        ledger: [l],
        valueStatement: s.value,
        valueLedger: l.value,
        residual: round2(s.value - l.value),
        ambiguous,
      });
    }
  }

  const stmtLeft = stmt.filter((s) => !usedStmt.has(s.id));
  const ledLeft = led.filter((l) => !usedLed.has(l.id));
  return { matches, stmtLeft, ledLeft };
}

// ─── Subset-sum acotado: subconjunto de `pool` que sume ≈ target ──────────
function findSubset(pool: RecItem[], target: number, tol: number, maxSize: number): RecItem[] | null {
  // Orden desc para podar antes y favorecer pocos elementos grandes.
  const items = [...pool].sort((a, b) => b.value - a.value);
  const n = items.length;
  let best: { idx: number[]; residual: number } | null = null;

  const suffix = new Array(n + 1).fill(0);
  for (let i = n - 1; i >= 0; i--) suffix[i] = suffix[i + 1] + items[i].value;

  // Límite de nodos explorados: evita bloquear el event loop indefinidamente
  // cuando el pool tiene muchos ítems con valores repetidos (p.ej. 90 comisiones).
  let nodes = 0;
  const MAX_NODES = 2_000_000;

  const chosen: number[] = [];
  const dfs = (start: number, remaining: number) => {
    if (nodes++ > MAX_NODES) return;
    if (best && Math.abs(best.residual) <= EPS) return; // ya hay exacto
    if (Math.abs(remaining) <= tol) {
      const residual = Math.abs(remaining);
      if (!best || residual < best.residual || (residual === best.residual && chosen.length < best.idx.length)) {
        best = { idx: [...chosen], residual };
      }
      if (residual <= EPS) return;
    }
    if (chosen.length >= maxSize) return;
    if (start >= n) return;
    if (suffix[start] + tol < remaining) return; // ni sumando todo lo restante alcanza
    for (let i = start; i < n; i++) {
      // Skip-duplicate: si este ítem tiene el mismo valor que el anterior
      // al mismo nivel, ya exploramos esa rama — saltamos para evitar la
      // explosión combinatoria con comisiones bancarias repetidas.
      if (i > start && Math.abs(items[i].value - items[i - 1].value) < EPS) continue;
      if (items[i].value - tol > remaining) continue; // este solo ya se pasa
      chosen.push(i);
      dfs(i + 1, remaining - items[i].value);
      chosen.pop();
      if (best && Math.abs(best.residual) <= EPS) return;
      if (nodes > MAX_NODES) return;
    }
  };
  dfs(0, target);

  if (!best) return null;
  // Solo agrupamos si son ≥2 elementos (un 1:1 ya se intentó antes).
  const b = best as { idx: number[]; residual: number };
  if (b.idx.length < 2) return null;
  // Al hacer skip-duplicate, el índice encontrado apunta al primero de cada valor;
  // pero en el resultado queremos ítems reales — mapeamos expandiendo duplicados.
  const result: RecItem[] = [];
  const usedIndexes = new Set(b.idx);
  const valueCounts = new Map<number, number>();
  for (const idx of b.idx) {
    const val = Math.round(items[idx].value * 100);
    valueCounts.set(val, (valueCounts.get(val) ?? 0) + 1);
  }
  const taken = new Map<number, number>();
  for (let i = 0; i < n && result.length < b.idx.length; i++) {
    const val = Math.round(items[i].value * 100);
    const need = valueCounts.get(val) ?? 0;
    const got = taken.get(val) ?? 0;
    if (got < need) { result.push(items[i]); taken.set(val, got + 1); }
  }
  return result;
}

// ─── Fase 2: agrupación por subset-sum (comisiones desglosadas/consolidadas) ─
function matchGroups(
  stmtLeft: RecItem[],
  ledLeft: RecItem[],
  tol: number,
  maxSize: number
): { matches: RecMatch[]; stmtLeft: RecItem[]; ledLeft: RecItem[] } {
  const matches: RecMatch[] = [];
  let sPool = [...stmtLeft];
  let lPool = [...ledLeft];

  // Tolerancia ampliada para grupos bancarios: el IVA sobre comisiones y pequeños
  // redondeos se acumulan cuando hay muchos cargos, por lo que un grupo puede
  // diferir más que una partida individual.
  const bankGroupTol = Math.max(tol, 1000);

  for (const dir of ["in", "out"] as Dir[]) {
    // 2a) Un renglón de contabilidad ↔ varios del extracto (lo más común).
    let ledTargets = lPool.filter((l) => l.direction === dir).sort((a, b) => b.value - a.value);
    for (const target of ledTargets) {
      if (!lPool.includes(target) || !isLedgerBankGroupItem(target)) continue;

      const bankCandidates = sPool.filter(isStatementBankGroupItem);
      const bankIn = round2(bankCandidates.filter((s) => s.direction === "in").reduce((a, s) => a + s.value, 0));
      const bankOut = round2(bankCandidates.filter((s) => s.direction === "out").reduce((a, s) => a + s.value, 0));
      const bankNet = target.direction === "in" ? round2(bankIn - bankOut) : round2(bankOut - bankIn);
      if (bankCandidates.length >= 2 && Math.abs(bankNet - target.value) <= bankGroupTol) {
        matches.push({
          type: "group",
          direction: dir,
          statement: bankCandidates,
          ledger: [target],
          valueStatement: bankNet,
          valueLedger: target.value,
          residual: round2(bankNet - target.value),
          ambiguous: false,
        });
        const bankIds = new Set(bankCandidates.map((s) => s.id));
        sPool = sPool.filter((s) => !bankIds.has(s.id));
        lPool = lPool.filter((l) => l.id !== target.id);
        continue;
      }

      const candidates = sPool.filter((s) => s.direction === dir && isStatementBankGroupItem(s));
      // Intentar con el pool COMPLETO del mismo sentido primero (evita la
      // restricción de maxSize cuando todos los gastos bancarios forman el total).
      const fullSum = round2(candidates.reduce((a, s) => a + s.value, 0));
      const fullMatch = candidates.length >= 1 && Math.abs(fullSum - target.value) <= bankGroupTol;
      const subset = fullMatch ? candidates : findSubset(candidates, target.value, bankGroupTol, maxSize);
      if (!subset) continue;
      const sum = round2(subset.reduce((a, s) => a + s.value, 0));
      matches.push({
        type: "group",
        direction: dir,
        statement: subset,
        ledger: [target],
        valueStatement: sum,
        valueLedger: target.value,
        residual: round2(sum - target.value),
        ambiguous: false,
      });
      const subsetIds = new Set(subset.map((s) => s.id));
      sPool = sPool.filter((s) => !subsetIds.has(s.id));
      lPool = lPool.filter((l) => l.id !== target.id);
    }

    // 2b) Un renglón del extracto ↔ varios de contabilidad.
    let stmtTargets = sPool.filter((s) => s.direction === dir).sort((a, b) => b.value - a.value);
    for (const target of stmtTargets) {
      if (!sPool.includes(target) || !isStatementBankGroupItem(target)) continue;
      const candidates = lPool.filter((l) => l.direction === dir && isLedgerBankGroupItem(l));
      const subset = findSubset(candidates, target.value, bankGroupTol, maxSize);
      if (!subset) continue;
      const sum = round2(subset.reduce((a, l) => a + l.value, 0));
      matches.push({
        type: "group",
        direction: dir,
        statement: [target],
        ledger: subset,
        valueStatement: target.value,
        valueLedger: sum,
        residual: round2(target.value - sum),
        ambiguous: false,
      });
      const subsetIds = new Set(subset.map((l) => l.id));
      lPool = lPool.filter((l) => !subsetIds.has(l.id));
      sPool = sPool.filter((s) => s.id !== target.id);
    }
  }

  return { matches, stmtLeft: sPool, ledLeft: lPool };
}

// ─── Cruces manuales forzados por el usuario ──────────────────────────────
function applyManual(
  stmt: RecItem[],
  led: RecItem[],
  manual: NonNullable<RecOptions["manualMatches"]>
): { matches: RecMatch[]; stmtLeft: RecItem[]; ledLeft: RecItem[] } {
  const matches: RecMatch[] = [];
  const sById = new Map(stmt.map((s) => [s.id, s]));
  const lById = new Map(led.map((l) => [l.id, l]));
  const usedS = new Set<string>();
  const usedL = new Set<string>();
  for (const m of manual) {
    const ss = m.statementIds.map((id) => sById.get(id)).filter((x): x is RecItem => !!x);
    const ll = m.ledgerIds.map((id) => lById.get(id)).filter((x): x is RecItem => !!x);
    if (ss.length === 0 && ll.length === 0) continue;
    ss.forEach((s) => usedS.add(s.id));
    ll.forEach((l) => usedL.add(l.id));
    const vs = round2(ss.reduce((a, s) => a + s.value, 0));
    const vl = round2(ll.reduce((a, l) => a + l.value, 0));
    matches.push({
      type: "manual",
      direction: ss[0]?.direction || ll[0]?.direction || "out",
      statement: ss,
      ledger: ll,
      valueStatement: vs,
      valueLedger: vl,
      residual: round2(vs - vl),
      ambiguous: false,
    });
  }
  return {
    matches,
    stmtLeft: stmt.filter((s) => !usedS.has(s.id)),
    ledLeft: led.filter((l) => !usedL.has(l.id)),
  };
}

/** Concilia extracto vs auxiliar contable y devuelve cruces, partidas y cuadre. */
export function reconcile(
  statement: RecStatementInput,
  ledger: RecLedgerInput,
  options: RecOptions = {}
): ReconciliationResult {
  const tol = options.tolerance ?? 100;
  const maxSize = options.maxSubsetSize ?? 12;
  const ignored = new Set(options.ignoredIds ?? []);

  const { stmt, led } = toItems(statement, ledger);

  const ignoredStmt = stmt.filter((s) => ignored.has(s.id));
  const ignoredLed = led.filter((l) => ignored.has(l.id));
  let sActive = stmt.filter((s) => !ignored.has(s.id));
  let lActive = led.filter((l) => !ignored.has(l.id));

  const allMatches: RecMatch[] = [];

  // 0) Cruces manuales (tienen prioridad).
  if (options.manualMatches?.length) {
    const r = applyManual(sActive, lActive, options.manualMatches);
    allMatches.push(...r.matches);
    sActive = r.stmtLeft;
    lActive = r.ledLeft;
  }

  // 1) 1:1 EXACTO (tolerancia mínima, solo decimales): casa los pares claros
  //    sin que la tolerancia ancha "robe" un renglón que en realidad es un
  //    grupo (p.ej. el único renglón de intereses contable vs. varios del banco).
  const tightTol = Math.min(tol, 1);
  const r1 = matchOneToOne(sActive, lActive, tightTol);
  allMatches.push(...r1.matches);

  // 2) Subset-sum con la tolerancia completa (comisiones desglosadas/consolidadas).
  const r2 = matchGroups(r1.stmtLeft, r1.ledLeft, tol, maxSize);
  allMatches.push(...r2.matches);

  // 3) 1:1 LAXO con la tolerancia completa, sobre lo que aún quedó suelto
  //    (pagos que difieren en unos pesos y no entraron en ningún grupo).
  const r3 = matchOneToOne(r2.stmtLeft, r2.ledLeft, tol);
  allMatches.push(...r3.matches);

  const stmtLeft = r3.stmtLeft;
  const ledLeft = r3.ledLeft;

  // 3) Partidas conciliatorias.
  const partidas: Record<PartidaCategory, RecItem[]> = {
    ingresos_no_contabilizados: stmtLeft.filter((s) => s.direction === "in"),
    egresos_no_contabilizados: stmtLeft.filter((s) => s.direction === "out"),
    ingresos_contab_sin_extracto: ledLeft.filter((l) => l.direction === "in"),
    egresos_contab_sin_extracto: ledLeft.filter((l) => l.direction === "out"),
  };

  const sum = (arr: RecItem[]) => round2(arr.reduce((a, x) => a + x.value, 0));
  const pIngNoContab = sum(partidas.ingresos_no_contabilizados);
  const pEgrNoContab = sum(partidas.egresos_no_contabilizados);
  const pIngContabSinExt = sum(partidas.ingresos_contab_sin_extracto);
  const pEgrContabSinExt = sum(partidas.egresos_contab_sin_extracto);

  // Residual neto de los renglones casados (efecto sobre la diferencia de saldos).
  const toleranceAdjustment = round2(allMatches.reduce((a, m) => a + m.residual, 0));

  const openingStatement = statement.opening;
  const openingLedger = ledger.opening;
  const closingStatement = statement.closing;
  const closingLedger = ledger.closing;
  const openingDiff = round2((openingStatement ?? 0) - (openingLedger ?? 0));
  const closingDiff = round2((closingStatement ?? 0) - (closingLedger ?? 0));

  // closingDiff = openingDiff + (netoExt - netoCont)
  //   netoExt - netoCont (de lo no casado) = +ingNoContab - egrNoContab - ingContabSinExt + egrContabSinExt
  //   + el residual por tolerancia de lo casado
  const explained = round2(
    openingDiff + pIngNoContab - pEgrNoContab - pIngContabSinExt + pEgrContabSinExt + toleranceAdjustment
  );
  const unexplained = round2(closingDiff - explained);

  return {
    tolerance: tol,
    matches: allMatches,
    partidas,
    cuadre: {
      openingStatement,
      openingLedger,
      openingDiff,
      closingStatement,
      closingLedger,
      closingDiff,
      partidas: {
        ingresos_no_contabilizados: pIngNoContab,
        egresos_no_contabilizados: pEgrNoContab,
        ingresos_contab_sin_extracto: pIngContabSinExt,
        egresos_contab_sin_extracto: pEgrContabSinExt,
      },
      toleranceAdjustment,
      explained,
      unexplained,
      balanced: Math.abs(unexplained) <= Math.max(tol, 1),
    },
    counts: {
      statement: stmt.length,
      ledger: led.length,
      matched1to1: allMatches.filter((m) => m.type === "1:1").length,
      grouped: allMatches.filter((m) => m.type === "group").length,
      manual: allMatches.filter((m) => m.type === "manual").length,
      unmatchedStatement: stmtLeft.length,
      unmatchedLedger: ledLeft.length,
    },
    ignored: { statement: ignoredStmt, ledger: ignoredLed },
  };
}
