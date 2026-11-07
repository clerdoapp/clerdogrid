import { parseNumber, toText } from '../engine/coerce';
import { isError, isMatrix, type Scalar } from '../engine/types';
import type { CfOperator, DataValidation } from '../model/workbook';
import type { GridStore } from './store';

const evalOne = (store: GridStore, formula: string | undefined, row: number, col: number): Scalar => {
  if (!formula) return null;
  const f = formula.startsWith('=') ? formula.slice(1) : formula;
  const v = store.engine.evaluate(f, store.sheetIndex, row, col);
  return isMatrix(v) ? v[0]![0]! : v;
};

export const compareOp = (op: CfOperator | undefined, x: number, a: number, b: number): boolean => {
  switch (op ?? 'between') {
    case 'between': return x >= Math.min(a, b) && x <= Math.max(a, b);
    case 'notBetween': return x < Math.min(a, b) || x > Math.max(a, b);
    case 'equal': return x === a;
    case 'notEqual': return x !== a;
    case 'greaterThan': return x > a;
    case 'lessThan': return x < a;
    case 'greaterThanOrEqual': return x >= a;
    case 'lessThanOrEqual': return x <= a;
  }
};

/** The choices of a list rule: "a,b,c" or a range / formula. */
export const listOptions = (store: GridStore, dv: DataValidation, row: number, col: number): string[] => {
  const f = dv.formula1 ?? '';
  if (/^".*"$/s.test(f)) return f.slice(1, -1).split(',').map((s) => s.trim()).filter(Boolean);
  if (!f.startsWith('=') && !/[A-Z]+\d|!/i.test(f) && f.includes(',')) return f.split(',').map((s) => s.trim()).filter(Boolean);
  const v = store.engine.evaluate(f.replace(/^=/, ''), store.sheetIndex, row, col);
  const out: string[] = [];
  const add = (x: Scalar) => { if (x !== null && !isError(x)) { const t = toText(x); if (typeof t === 'string' && t !== '' && !out.includes(t)) out.push(t); } };
  if (isMatrix(v)) for (const r of v) for (const x of r) add(x); else add(v);
  return out;
};

/** Whether `text` is allowed in the cell by its validation rule. */
export const checkValidation = (store: GridStore, dv: DataValidation, row: number, col: number, text: string): boolean => {
  if (dv.type === 'list') {
    const opts = listOptions(store, dv, row, col).map((o) => o.toLowerCase());
    return opts.includes(text.trim().toLowerCase());
  }
  if (dv.type === 'custom') {
    const prev = store.engine.getValue(store.sheetIndex, row, col);
    const n = parseNumber(text);
    // Evaluate as if the value were already in the cell.
    const cell = store.cell(row, col);
    const old = cell?.v;
    if (cell && !cell.f) cell.v = n ?? text;
    const ok = evalOne(store, dv.formula1, row, col);
    if (cell && !cell.f) cell.v = old ?? prev;
    return ok === true || (typeof ok === 'number' && ok !== 0);
  }
  if (dv.type === 'textLength') {
    const a = Number(evalOne(store, dv.formula1, row, col)), b = Number(evalOne(store, dv.formula2, row, col));
    return compareOp(dv.operator, text.length, a, b);
  }
  const n = parseNumber(text, store.wb.date1904);
  if (n === null) return false;
  if (dv.type === 'whole' && !Number.isInteger(n)) return false;
  if (dv.type === 'time' && (n < 0 || n >= 1)) return false;
  const a = Number(evalOne(store, dv.formula1, row, col)), b = Number(evalOne(store, dv.formula2, row, col));
  if (Number.isNaN(a)) return true;
  return compareOp(dv.operator, n, a, b);
};
