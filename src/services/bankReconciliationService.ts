/**
 * Motor de conciliación bancaria: cruza el EXTRACTO del banco contra el LIBRO
 * AUXILIAR de la contabilidad (misma cuenta y periodo) y arma el cuadre.
 *
 * Estrategia (acordada con el usuario):
 *  0) Los gastos bancarios del extracto (4x1000, GMF, comisiones, IVA sobre
 *     comisiones, etc.) se excluyen del cruce por completo: casi nunca se
 *     causan en la contabilidad renglón por renglón (ni siquiera
 *     consolidados), así que siempre van 100% a su propia partida
 *     "gastos bancarios sin contabilizar", agrupados por concepto.
 *  1) Cruce 1:1 por MONTO dentro de tolerancia, separando por dirección
 *     (ingreso↔débito, egreso↔crédito). La fecha se usa como desempate, no
 *     como filtro. Maneja multiplicidad: 3 pagos de $100k en extracto vs 2 en
 *     contabilidad → casa 2, deja 1 como partida (y marca el grupo ambiguo).
 *  2) Lo que siga suelto = partidas conciliatorias en 5 categorías.
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
  | "gastos_bancarios_sin_contabilizar"
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
    gastos_bancarios_sin_contabilizar: number;
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

// Gastos bancarios (4x1000, GMF, comisiones, IVA sobre comisiones, cuotas de
// manejo, etc.): casi nunca se causan en la contabilidad renglón por renglón
// (ni siquiera consolidados 1 a 1), así que se excluyen del motor de cruce
// por completo y van siempre a su propia partida "gastos bancarios sin
// contabilizar", agrupados por concepto (ver reconcile()).
const BANK_FEE_RX =
  /impto gobierno 4x1000|4\s*x\s*1\.?000|gmf|cargo por impuesto|cuota manejo suc virt empresa|servicio pago a otros bancos|iva cuota manejo suc virt emp|cobro iva pagos automaticos|iva boton|comision boton|servicio por pagos a nequi|servicio pago a proveedores|servicio pago de nomina/i;

const isBankFeeItem = (item: RecItem): boolean => item.kind === "bank_fee" || BANK_FEE_RX.test(item.description);

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
  const ignored = new Set(options.ignoredIds ?? []);

  const { stmt, led } = toItems(statement, ledger);

  const ignoredStmt = stmt.filter((s) => ignored.has(s.id));
  const ignoredLed = led.filter((l) => ignored.has(l.id));
  let sActive = stmt.filter((s) => !ignored.has(s.id));
  let lActive = led.filter((l) => !ignored.has(l.id));

  // Gastos bancarios (4x1000, GMF, comisiones, IVA sobre comisiones, etc.):
  // se excluyen del cruce por completo y van siempre 100% a su propia
  // partida, sin importar si algún asiento consolidado de la contabilidad
  // "calzaría" contra ellos — casi nunca se causan uno a uno ni consolidado.
  const bankFees = sActive.filter((s) => s.direction === "out" && isBankFeeItem(s));
  const bankFeeIds = new Set(bankFees.map((s) => s.id));
  sActive = sActive.filter((s) => !bankFeeIds.has(s.id));

  const allMatches: RecMatch[] = [];

  // 0) Cruces manuales (tienen prioridad).
  if (options.manualMatches?.length) {
    const r = applyManual(sActive, lActive, options.manualMatches);
    allMatches.push(...r.matches);
    sActive = r.stmtLeft;
    lActive = r.ledLeft;
  }

  // 1) 1:1 dentro de tolerancia (fecha como desempate).
  const r1 = matchOneToOne(sActive, lActive, tol);
  allMatches.push(...r1.matches);

  const stmtLeft = r1.stmtLeft;
  const ledLeft = r1.ledLeft;

  // 2) Partidas conciliatorias.
  const partidas: Record<PartidaCategory, RecItem[]> = {
    ingresos_no_contabilizados: stmtLeft.filter((s) => s.direction === "in"),
    egresos_no_contabilizados: stmtLeft.filter((s) => s.direction === "out"),
    gastos_bancarios_sin_contabilizar: bankFees,
    ingresos_contab_sin_extracto: ledLeft.filter((l) => l.direction === "in"),
    egresos_contab_sin_extracto: ledLeft.filter((l) => l.direction === "out"),
  };

  const sum = (arr: RecItem[]) => round2(arr.reduce((a, x) => a + x.value, 0));
  const pIngNoContab = sum(partidas.ingresos_no_contabilizados);
  const pEgrNoContab = sum(partidas.egresos_no_contabilizados);
  const pGastosBancarios = sum(partidas.gastos_bancarios_sin_contabilizar);
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
    openingDiff +
      pIngNoContab -
      pEgrNoContab -
      pGastosBancarios -
      pIngContabSinExt +
      pEgrContabSinExt +
      toleranceAdjustment
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
        gastos_bancarios_sin_contabilizar: pGastosBancarios,
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
