import { CellError, Err, isError, isLambda, type FnContext, type FunctionDef, type Scalar, type Value } from '../types';
import { cellsOf, fromRange, toBool } from '../coerce';

/**
 * Logical values of AND/OR/XOR arguments. Inside ranges and arrays only numbers and booleans count
 * (text and blanks are skipped); typed arguments are coerced, so "TRUE" counts and "abc" is #VALUE!.
 */
const logicals = (args: readonly Value[], ctx: FnContext): boolean[] | CellError => {
  const out: boolean[] = [];
  for (const a of args) {
    if (fromRange(a)) {
      for (const c of cellsOf(a, ctx)) {
        if (isError(c)) return c;
        if (typeof c === 'number') out.push(c !== 0);
        else if (typeof c === 'boolean') out.push(c);
      }
      continue;
    }
    if (isLambda(a)) return Err.calc;
    const b = toBool(a as Scalar);
    if (isError(b)) return b;
    out.push(b);
  }
  return out.length ? out : Err.value;
};

const reduce = (name: string, f: (bs: boolean[]) => boolean): FunctionDef => ({
  name, min: 1, max: 255,
  fn: (args, ctx) => {
    const bs = logicals(args, ctx);
    return isError(bs) ? bs : f(bs);
  },
});

export const defs: FunctionDef[] = [
  reduce('AND', (bs) => bs.every(Boolean)),
  reduce('OR', (bs) => bs.some(Boolean)),
  reduce('XOR', (bs) => bs.filter(Boolean).length % 2 === 1),
  { name: 'NOT', min: 1, max: 1, scalar: true, fn: ([v]) => { const b = toBool(v as Scalar); return isError(b) ? b : !b; } },
  { name: 'TRUE', min: 0, max: 0, fn: () => true },
  { name: 'FALSE', min: 0, max: 0, fn: () => false },
];
