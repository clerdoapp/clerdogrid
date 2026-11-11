// Background error checking: the rules behind the green corner indicators and their quick fixes.
import { parseFormula } from '../../engine/parser';
import { toR1C1, fromR1C1 } from '../../engine/r1c1';
import { CellError } from '../../engine/types';
import { cellKey, parseRange, type CellRange, type ErrorCheckType } from '../../model/workbook';
import { inRange, type GridStore } from '../../app/store';
import { settings, type ErrorRule } from './settings';
import { N } from '../../ui/i18n';

export interface Issue { rule: ErrorRule; title: string; help: string }

const NUMBER_TEXT = /^\s*[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)?(?:\.\d+)?(?:[eE][-+]?\d+)?\s*%?\s*$/;
const TWO_DIGIT_DATE = /^\s*(?:\d{1,2}[/-]\d{1,2}[/-]\d{2}|[A-Za-z]{3,9}\.? \d{1,2},? \d{2}|\d{1,2}[ -][A-Za-z]{3,9}[ -]\d{2})\s*$/;

export const RULE_TEXT: Record<ErrorRule, { label: string; title: string; help: string }> = {
  numberStoredAsText: { label: N('Numbers stored as text'), title: N('Number stored as text'), help: N('This cell holds a number as text, so formulas may skip it and it sorts as text. Convert it to a number to calculate with it.') },
  formula: { label: N('Formulas that differ from their neighbors'), title: N('Inconsistent formula'), help: N('The formula in this cell doesn\'t follow the pattern of the formulas next to it. Copy the neighboring formula if this one should match.') },
  formulaRange: { label: N('Formulas that leave out cells next to their range'), title: N('Formula leaves out adjacent cells'), help: N('This formula refers to a range that has more numbers right next to it. Extend the range if those cells belong in it.') },
  unlockedFormula: { label: N('Unlocked cells with formulas'), title: N('Unlocked formula cell'), help: N('This cell has a formula but isn\'t locked, so it can be changed even when the sheet is protected. Lock it to keep the formula safe.') },
  twoDigitTextYear: { label: N('Text dates with two-digit years'), title: N('Text date with a two-digit year'), help: N('This text looks like a date with a two-digit year, which can be read as the wrong century. Convert it to a real date.') },
  evalError: { label: N('Formulas that give errors'), title: N('Formula error'), help: N('This formula returns an error. Check the values it refers to, or step through it with Evaluate formula.') },
};

const ignored = (store: GridStore, row: number, col: number, rule: ErrorRule): boolean =>
  !!store.sheet.ignoredErrors?.some((ie) => inRange(ie.range, row, col) && ie.types.includes(rule as ErrorCheckType));

const neighbourFormula = (store: GridStore, row: number, col: number): string | null => {
  const c = store.cell(row, col);
  return c?.f ? toR1C1(c.f, row, col) : null;
};

/** Ranges a formula refers to on its own sheet (single-area references only). */
const ownRanges = (f: string): CellRange[] => {
  const out: CellRange[] = [];
  const re = /(^|[^A-Za-z0-9_.!$'\]])(\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+)(?![A-Za-z0-9_(!])/g;
  const text = f.replace(/"[^"]*"/g, '""');
  let m: RegExpExecArray | null;
  while ((m = re.exec(text))) { const r = parseRange(m[2]!.replace(/\$/g, '')); if (r) out.push(r); }
  return out;
};

const isNum = (store: GridStore, row: number, col: number): boolean => row >= 0 && col >= 0 && typeof store.value(row, col) === 'number' && !store.cell(row, col)?.f;

/** The range the formula probably meant to cover, if numbers sit right next to one of its ranges. */
export const widerRange = (store: GridStore, row: number, col: number): { from: CellRange; to: CellRange } | null => {
  const c = store.cell(row, col);
  if (!c?.f || !/^(SUM|AVERAGE|COUNT|COUNTA|MIN|MAX|PRODUCT|SUBTOTAL|AGGREGATE)\(/i.test(c.f)) return null;
  for (const r of ownRanges(c.f)) {
    if (r.c1 === r.c2 && r.r2 > r.r1 - 1) {
      let r2 = r.r2, r1 = r.r1;
      while (r2 + 1 < row && isNum(store, r2 + 1, r.c1) && !(r2 + 1 === row && r.c1 === col)) r2++;
      if (r.c1 !== col || r.r2 >= row) { while (isNum(store, r2 + 1, r.c1) && !(r2 + 1 === row && r.c1 === col)) r2++; }
      while (r1 > 0 && isNum(store, r1 - 1, r.c1) && !(r1 - 1 === row && r.c1 === col)) r1--;
      if (r2 !== r.r2 || r1 !== r.r1) return { from: r, to: { ...r, r1, r2 } };
    } else if (r.r1 === r.r2) {
      let c2 = r.c2, c1 = r.c1;
      while (isNum(store, r.r1, c2 + 1) && !(c2 + 1 === col && r.r1 === row)) c2++;
      while (c1 > 0 && isNum(store, r.r1, c1 - 1) && !(c1 - 1 === col && r.r1 === row)) c1--;
      if (c2 !== r.c2 || c1 !== r.c1) return { from: r, to: { ...r, c1, c2 } };
    }
  }
  return null;
};

/** The neighbour formula (R1C1) this cell breaks the pattern of, if any. */
export const patternBreak = (store: GridStore, row: number, col: number): { r: number; c: number } | null => {
  const own = neighbourFormula(store, row, col);
  if (!own) return null;
  const pairs: [[number, number], [number, number]][] = [[[row - 1, col], [row + 1, col]], [[row, col - 1], [row, col + 1]]];
  for (const [a, b] of pairs) {
    const fa = neighbourFormula(store, a[0], a[1]), fb = neighbourFormula(store, b[0], b[1]);
    if (fa && fa === fb && fa !== own) return { r: a[0], c: a[1] };
  }
  // At the end of a run, two formulas in a row on one side also set the pattern.
  const runs: [[number, number], [number, number]][] = [[[row - 1, col], [row - 2, col]], [[row + 1, col], [row + 2, col]], [[row, col - 1], [row, col - 2]], [[row, col + 1], [row, col + 2]]];
  for (const [a, b] of runs) {
    if (a[0] < 0 || a[1] < 0 || b[0] < 0 || b[1] < 0) continue;
    const fa = neighbourFormula(store, a[0], a[1]), fb = neighbourFormula(store, b[0], b[1]);
    if (fa && fa === fb && fa !== own) {
      // Only when the cell on the other side doesn't continue this cell's own pattern.
      const other = neighbourFormula(store, row + (row - a[0]), col + (col - a[1]));
      if (other !== own) return { r: a[0], c: a[1] };
    }
  }
  return null;
};

export const issuesAt = (store: GridStore, row: number, col: number): Issue[] => {
  if (!settings.errorChecking) return [];
  const out: Issue[] = [];
  const c = store.sheet.cells.get(cellKey(row, col));
  if (!c) return out;
  const add = (rule: ErrorRule) => { if (settings.rules[rule] && !ignored(store, row, col, rule)) out.push({ rule, ...RULE_TEXT[rule] }); };
  const v = store.value(row, col);
  if (c.f) {
    if (v instanceof CellError) add('evalError');
    if (settings.rules.formula && patternBreak(store, row, col)) add('formula');
    if (settings.rules.formulaRange && widerRange(store, row, col)) add('formulaRange');
    if (store.style(row, col).protection?.locked === false) add('unlockedFormula');
  } else if (typeof v === 'string' && v.trim()) {
    if (NUMBER_TEXT.test(v) && /\d/.test(v)) add('numberStoredAsText');
    else if (TWO_DIGIT_DATE.test(v)) add('twoDigitTextYear');
  }
  return out;
};

// ---------- Fixes ----------

const selectedFlagged = (store: GridStore, rule: ErrorRule): { row: number; col: number }[] => {
  const out: { row: number; col: number }[] = [];
  const b = store.engine.usedBounds(store.sheetIndex);
  for (const r of store.sel.ranges) {
    for (let row = r.r1; row <= Math.min(r.r2, b.rows); row++) for (let col = r.c1; col <= Math.min(r.c2, b.cols); col++) {
      if (out.length > 100000) return out;
      if (issuesAt(store, row, col).some((i) => i.rule === rule)) out.push({ row, col });
    }
  }
  return out;
};

export const convertToNumbers = (store: GridStore): void => {
  const cells = selectedFlagged(store, 'numberStoredAsText');
  if (!cells.length) return;
  store.change(N('Convert to number'), cells.map((p) => ({ area: { r1: p.row, c1: p.col, r2: p.row, c2: p.col } })), () => {
    for (const p of cells) {
      const raw = String(store.value(p.row, p.col)).trim();
      const pct = raw.endsWith('%');
      const n = Number(raw.replace(/[,%\s]/g, ''));
      if (!Number.isFinite(n)) continue;
      store.engine.setValue(store.sheetIndex, p.row, p.col, pct ? n / 100 : n);
      const st = store.style(p.row, p.col);
      if (st.numFmt === '@') store.styleRange({ r1: p.row, c1: p.col, r2: p.row, c2: p.col }, (s) => ({ ...s, numFmt: pct ? '0%' : undefined }));
    }
  });
};

export const copyNeighbourFormula = (store: GridStore): void => {
  const { row, col } = store.sel.active;
  const n = patternBreak(store, row, col);
  if (!n) return;
  const r1 = neighbourFormula(store, n.r, n.c)!;
  store.change(N('Copy formula'), [{ area: { r1: row, c1: col, r2: row, c2: col } }], () => store.enter(row, col, `=${fromR1C1(r1, row, col)}`));
};

export const extendRange = (store: GridStore): void => {
  const { row, col } = store.sel.active;
  const w = widerRange(store, row, col);
  const c = store.cell(row, col);
  if (!w || !c?.f) return;
  const re = new RegExp(`(\\$?)${colLetters(w.from.c1)}(\\$?)${w.from.r1 + 1}:(\\$?)${colLetters(w.from.c2)}(\\$?)${w.from.r2 + 1}(?![0-9])`);
  const next = c.f.replace(re, (_m, d1: string, d2: string, d3: string, d4: string) => `${d1}${colLetters(w.to.c1)}${d2}${w.to.r1 + 1}:${d3}${colLetters(w.to.c2)}${d4}${w.to.r2 + 1}`);
  try { parseFormula(next); } catch { return; }
  store.change(N('Update formula'), [{ area: { r1: row, c1: col, r2: row, c2: col } }], () => store.enter(row, col, `=${next}`));
};

const colLetters = (col: number): string => {
  let s = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
};

export const lockCells = (store: GridStore): void => {
  const cells = selectedFlagged(store, 'unlockedFormula');
  if (!cells.length) return;
  store.change(N('Lock cell'), cells.map((p) => ({ area: { r1: p.row, c1: p.col, r2: p.row, c2: p.col } })), () => {
    for (const p of cells) store.styleRange({ r1: p.row, c1: p.col, r2: p.row, c2: p.col }, (s) => { const { protection, ...rest } = s; return protection?.hidden ? { ...rest, protection: { hidden: true } } : rest; });
  });
};

/** Converts text dates with two-digit years to real dates in the 1900s or 2000s. */
export const convertTwoDigitYears = (store: GridStore, century: 1900 | 2000): void => {
  const cells = selectedFlagged(store, 'twoDigitTextYear');
  if (!cells.length) return;
  store.change(N('Convert dates'), cells.map((p) => ({ area: { r1: p.row, c1: p.col, r2: p.row, c2: p.col } })), () => {
    for (const p of cells) {
      const raw = String(store.value(p.row, p.col)).trim();
      const fixed = raw.replace(/(\d{2})\s*$/, (y) => String(century + Number(y)));
      store.enter(p.row, p.col, fixed);
      if (typeof store.value(p.row, p.col) !== 'number') store.engine.setValue(store.sheetIndex, p.row, p.col, raw);
    }
  });
};

export const ignoreIssue = (store: GridStore, rule: ErrorRule): void => {
  const cells = selectedFlagged(store, rule);
  const target = cells.length ? cells : [store.sel.active];
  store.change(N('Ignore error'), 'sheet', () => {
    const list = store.sheet.ignoredErrors ?? [];
    for (const p of target) {
      const hit = list.find((ie) => ie.range.r1 === p.row && ie.range.r2 === p.row && ie.range.c1 === p.col && ie.range.c2 === p.col);
      if (hit) { if (!hit.types.includes(rule as ErrorCheckType)) hit.types.push(rule as ErrorCheckType); }
      else list.push({ range: { r1: p.row, c1: p.col, r2: p.row, c2: p.col }, types: [rule as ErrorCheckType] });
    }
    store.sheet.ignoredErrors = list;
  });
};

/** Clears every "ignore error" on the sheet so the indicators come back. */
export const resetIgnored = (store: GridStore): void => {
  if (!store.wb.sheets.some((s) => s.ignoredErrors?.length)) return;
  store.change(N('Reset ignored errors'), 'book', () => { for (const s of store.wb.sheets) delete s.ignoredErrors; });
};
