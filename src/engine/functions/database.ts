import { CellError, Err, isError, type FnContext, type FunctionDef, type Matrix, type Scalar, type Value } from '../types';
import { first, parseNumber, toMatrix, toText } from '../coerce';

// ---------------------------------------------------------------- criteria matching

type Op = '=' | '<>' | '<' | '<=' | '>' | '>=';
/** One criteria cell compiled to a test. */
type Test = (v: Scalar) => boolean;

const lower = (s: string): string => s.toLowerCase();

/** Wildcard pattern (* ? and ~ escape) to an anchored, case-insensitive regular expression. */
const wildcard = (pattern: string, prefix: boolean): RegExp => {
  let re = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~' && i + 1 < pattern.length && '*?~'.includes(pattern[i + 1])) re += '\\' + pattern[++i];
    else if (ch === '*') re += '[\\s\\S]*';
    else if (ch === '?') re += '[\\s\\S]';
    else re += ch.replace(/[.+^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp(`^${re}${prefix ? '' : '$'}`, 'i');
};

const cmp = (op: Op, d: number): boolean =>
  op === '=' ? d === 0 : op === '<>' ? d !== 0 : op === '<' ? d < 0 : op === '<=' ? d <= 0 : op === '>' ? d > 0 : d >= 0;

/**
 * Compiles a criteria cell following spreadsheet's database rules:
 * - a number or logical matches equal values of the same type;
 * - plain text matches cells that begin with it ("Ap" matches "Apple"), wildcards allowed;
 * - "=text" matches exactly, "=" alone matches blanks, "<>" alone matches non-blanks;
 * - comparison operators compare numbers with numbers and text with text (case-insensitive).
 */
const compile = (c: Scalar, date1904: boolean): Test | null => {
  if (c === null || c === '') return null;
  if (isError(c)) return (v) => isError(v) && v.code === c.code;
  if (typeof c === 'number') return (v) => typeof v === 'number' && v === c;
  if (typeof c === 'boolean') return (v) => v === c;
  const m = /^(<=|>=|<>|=|<|>)?([\s\S]*)$/.exec(c)!;
  const op = (m[1] ?? '') as Op | '';
  const operand = m[2];
  if (!op) {
    const n = parseNumber(operand, date1904);
    if (n !== null) return (v) => typeof v === 'number' && v === n;
    const u = lower(operand);
    if (u === 'true' || u === 'false') return (v) => v === (u === 'true');
    const re = wildcard(operand, true);
    return (v) => typeof v === 'string' && re.test(v);
  }
  if (operand === '') {
    if (op === '=') return (v) => v === null || v === '';
    if (op === '<>') return (v) => v !== null && v !== '';
    return () => false;
  }
  const n = parseNumber(operand, date1904);
  if (n !== null) {
    return (v) => {
      if (typeof v === 'number') return cmp(op, v - n);
      return op === '<>';
    };
  }
  const u = lower(operand);
  if (u === 'true' || u === 'false') {
    const b = u === 'true';
    return (v) => (typeof v === 'boolean' ? cmp(op, Number(v) - Number(b)) : op === '<>');
  }
  if (op === '=' || op === '<>') {
    const re = wildcard(operand, false);
    return (v) => (typeof v === 'string' && re.test(v)) === (op === '=');
  }
  return (v) => typeof v === 'string' && cmp(op, lower(v).localeCompare(u));
};

/** Header lookup is case-insensitive and ignores surrounding spaces. */
const headerKey = (v: Scalar): string => {
  const t = toText(v);
  return isError(t) ? '' : lower(t.trim());
};

/**
 * Rows of the database (without the header) that satisfy the criteria: the conditions in one criteria
 * row are ANDed, and criteria rows are ORed. A condition under a header that is not a database
 * column can never match.
 */
const matchingRows = (db: Matrix, criteria: Matrix, date1904: boolean): Scalar[][] | CellError => {
  if (criteria.length < 2) return Err.value;
  const keys = db[0].map(headerKey);
  const cols = criteria[0].map((h) => {
    const k = headerKey(h);
    return k ? keys.indexOf(k) : -1;
  });
  const rows: { col: number; test: Test }[][] = [];
  for (const crow of criteria.slice(1)) {
    const conds: { col: number; test: Test }[] = [];
    crow.forEach((c, j) => {
      const test = compile(c, date1904);
      if (test) conds.push({ col: cols[j], test });
    });
    rows.push(conds);
  }
  return db.slice(1).filter((rec) => rows.some((conds) => conds.every(({ col, test }) => col >= 0 && test(rec[col] ?? null))));
};

// ---------------------------------------------------------------- shared plumbing

/** Resolves the field argument to a 0-based column: header text or a 1-based index. */
const fieldIndex = (db: Matrix, field: Value, ctx: FnContext): number | CellError => {
  const f = first(field, ctx);
  if (isError(f)) return f;
  if (typeof f === 'number') {
    const i = Math.trunc(f);
    return i >= 1 && i <= db[0].length ? i - 1 : Err.value;
  }
  if (typeof f === 'string') {
    const i = db[0].map(headerKey).indexOf(lower(f.trim()));
    return i >= 0 ? i : Err.value;
  }
  return Err.value;
};

interface Selection { records: Scalar[][]; col: number | null }

const select = (args: Value[], ctx: FnContext, fieldOptional: boolean): Selection | CellError => {
  const [dbArg, fieldArg, critArg] = args;
  if (isError(dbArg)) return dbArg;
  if (isError(critArg)) return critArg;
  const db = toMatrix(dbArg, ctx);
  if (isError(db)) return db;
  const criteria = toMatrix(critArg ?? null, ctx);
  if (isError(criteria)) return criteria;
  if (!db.length || !db[0].length) return Err.value;
  let col: number | null = null;
  if (fieldArg !== null && fieldArg !== undefined) {
    const c = fieldIndex(db, fieldArg, ctx);
    if (isError(c)) return c;
    col = c;
  } else if (!fieldOptional) return Err.value;
  const records = matchingRows(db, criteria, ctx.date1904);
  return isError(records) ? records : { records, col };
};

/** Numbers of the field column in the matching records; error cells propagate. */
const numbers = (s: Selection): number[] | CellError => {
  const out: number[] = [];
  for (const r of s.records) {
    const v = r[s.col!] ?? null;
    if (isError(v)) return v;
    if (typeof v === 'number') out.push(v);
  }
  return out;
};

const fin = (x: number): number | CellError => (Number.isFinite(x) ? x : Err.num);

const numeric = (name: string, f: (xs: number[]) => number | CellError): FunctionDef => ({
  name, min: 3, max: 3,
  fn: (args, ctx) => {
    const s = select(args, ctx, false);
    if (isError(s)) return s;
    const xs = numbers(s);
    if (isError(xs)) return xs;
    const r = f(xs);
    return isError(r) ? r : fin(r);
  },
});

const sum = (xs: number[]): number => xs.reduce((a, b) => a + b, 0);
const variance = (xs: number[], sample: boolean): number | CellError => {
  const n = xs.length;
  if (n < (sample ? 2 : 1)) return Err.div0;
  const mean = sum(xs) / n;
  return xs.reduce((a, x) => a + (x - mean) ** 2, 0) / (sample ? n - 1 : n);
};
const stdev = (xs: number[], sample: boolean): number | CellError => {
  const v = variance(xs, sample);
  return isError(v) ? v : Math.sqrt(v);
};

export const defs: FunctionDef[] = [
  numeric('DSUM', sum),
  numeric('DAVERAGE', (xs) => (xs.length ? sum(xs) / xs.length : Err.div0)),
  numeric('DMAX', (xs) => (xs.length ? Math.max(...xs) : 0)),
  numeric('DMIN', (xs) => (xs.length ? Math.min(...xs) : 0)),
  numeric('DPRODUCT', (xs) => (xs.length ? xs.reduce((a, b) => a * b, 1) : 0)),
  numeric('DVAR', (xs) => variance(xs, true)),
  numeric('DVARP', (xs) => variance(xs, false)),
  numeric('DSTDEV', (xs) => stdev(xs, true)),
  numeric('DSTDEVP', (xs) => stdev(xs, false)),
  {
    name: 'DCOUNT', min: 2, max: 3,
    fn: (args, ctx) => {
      // DCOUNT(db, , criteria) and DCOUNT(db, criteria) both count matching records.
      const a = args.length === 2 ? [args[0], null, args[1]] : args;
      const s = select(a, ctx, true);
      if (isError(s)) return s;
      return s.col === null ? s.records.length : s.records.filter((r) => typeof r[s.col!] === 'number').length;
    },
  },
  {
    name: 'DCOUNTA', min: 2, max: 3,
    fn: (args, ctx) => {
      const a = args.length === 2 ? [args[0], null, args[1]] : args;
      const s = select(a, ctx, true);
      if (isError(s)) return s;
      return s.col === null ? s.records.length : s.records.filter((r) => (r[s.col!] ?? null) !== null && r[s.col!] !== '').length;
    },
  },
  {
    name: 'DGET', min: 3, max: 3,
    fn: (args, ctx) => {
      const s = select(args, ctx, false);
      if (isError(s)) return s;
      if (s.records.length === 0) return Err.value;
      if (s.records.length > 1) return Err.num;
      return s.records[0][s.col!] ?? 0;
    },
  },
];
