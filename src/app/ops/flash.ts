// Flash Fill on the grid: reads the examples next to a block of data, learns a program
// (see ../flashfill.ts) and fills the empty cells in one undoable step.
import { parseNumber } from '../../engine/coerce';
import { formatValue } from '../../formats/numfmt';
import { activeLocale, activeNames } from '../../formats/locales';
import { synthesize, type FlashOptions, type FlashProgram } from '../flashfill';
import { displayFormat } from '../locale';
import type { GridStore } from '../store';
import { N, T } from '../../ui/i18n';

/** Set when the grey suggestion preview is running (src/features/flashfill); then nothing fills on its own. */
export const flashPreview = { enabled: false };

export interface FlashPlan {
  sheet: number;
  col: number;
  top: number;
  bottom: number;
  program: FlashProgram;
  /** Values for the empty cells, by row. */
  fills: Map<number, string | number>;
  examples: number;
}

const MAX_ROWS = 20_000;
const MAX_INPUT_COLS = 6;

export const flashOptions = (): FlashOptions => {
  const n = activeNames();
  return { dateOrder: activeLocale().dateOrder, months: n.months, monthsShort: n.monthsShort, days: n.days, daysShort: n.daysShort };
};

const shown = (store: GridStore, row: number, col: number): string => {
  const v = store.value(row, col);
  if (v === null) return '';
  if (typeof v === 'string') return v;
  return formatValue(v, displayFormat(store.style(row, col).numFmt) ?? 'General', { date1904: store.wb.date1904 }).text;
};

/** Works out what Flash Fill would put in column `col` around `row`, or a reason it can't. */
export const planFlashFill = (store: GridStore, row: number, col: number, minExamples = 1): FlashPlan | string => {
  if (col === 0) return T('Type an example of what you want in the cell next to your data, then try Flash Fill again.');
  const sheet = store.sheetIndex;
  const has = (r: number, c: number): boolean => store.value(r, c) !== null;
  // Input columns: the filled block directly to the left (and the one to the right, if any).
  const left: number[] = [];
  for (let c = col - 1; c >= 0 && left.length < MAX_INPUT_COLS && has(row, c); c--) left.unshift(c);
  const right: number[] = [];
  for (let c = col + 1; right.length < 2 && has(row, c); c++) right.push(c);
  const inputCols = [...left, ...right];
  const anyInput = (r: number): boolean => inputCols.some((c) => has(r, c));
  let top = row, bottom = row;
  while (top > 0 && anyInput(top - 1) && row - top < MAX_ROWS) top--;
  while (anyInput(bottom + 1) && bottom - row < MAX_ROWS) bottom++;
  if (!inputCols.length) return T('Type an example of what you want in the cell next to your data, then try Flash Fill again.');
  const inputsAt = (r: number): string[] => inputCols.map((c) => shown(store, r, c));
  const examples: { inputs: string[]; output: string }[] = [];
  const empty: number[] = [];
  let numeric = true;
  for (let r = top; r <= bottom; r++) {
    const v = store.value(r, col);
    if (v === null || v === '') { if (anyInput(r)) empty.push(r); continue; }
    if (store.cell(r, col)?.f) continue;
    if (typeof v !== 'number') numeric = false;
    examples.push({ inputs: inputsAt(r), output: shown(store, r, col) });
  }
  if (!examples.length) return T('Type an example of what you want in the cell next to your data, then try Flash Fill again.');
  if (examples.length < minExamples) return T('More examples needed.');
  if (!empty.length) return T('There are no empty cells to fill next to this data.');
  const unlabeled = empty.map(inputsAt);
  const program = synthesize(examples, unlabeled, flashOptions());
  if (!program) return T('Flash Fill couldn\'t find a pattern. Give one or two more examples.');
  const fills = new Map<number, string | number>();
  empty.forEach((r, i) => {
    const out = program.run(unlabeled[i]!);
    if (out === null || out === '') return;
    const n = numeric ? parseNumber(out) : null;
    fills.set(r, n !== null && numeric ? n : out);
  });
  if (!fills.size) return T('Flash Fill couldn\'t find a pattern. Give one or two more examples.');
  return { sheet, col, top, bottom, program, fills, examples: examples.length };
};

export const applyFlashPlan = (store: GridStore, plan: FlashPlan): void => {
  store.change(N('Flash Fill'), [{ area: { r1: plan.top, c1: plan.col, r2: plan.bottom, c2: plan.col }, sheet: plan.sheet }], () => {
    for (const [r, v] of plan.fills) if (store.value(r, plan.col, plan.sheet) === null) store.engine.setValue(plan.sheet, r, plan.col, v);
  });
  store.notify(T('Flash Fill filled {count, plural, one {# cell} other {# cells}}.', { count: plan.fills.size }));
};

/** Ctrl+E: fills the column of the active cell from the examples typed so far. */
export const flashFill = (store: GridStore): boolean => {
  const { row, col } = store.sel.active;
  const plan = planFlashFill(store, row, col);
  if (typeof plan === 'string') { store.notify(plan); return false; }
  applyFlashPlan(store, plan);
  return true;
};
