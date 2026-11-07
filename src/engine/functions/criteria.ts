import { CellError, isError, isRef, type FnContext, type Matrix, type RangeRef, type Scalar, type Value } from '../types';
import { approxEqual, parseNumber, refArea, toMatrix } from '../coerce';

/**
 * A range or array prepared for scanning. `rows`/`cols` cover the data actually held (references are
 * trimmed to the used area); `fullRows`/`fullCols` are the true shape, for size checks and blank counts.
 */
export interface Grid {
  readonly m: Matrix;
  readonly rows: number;
  readonly cols: number;
  readonly fullRows: number;
  readonly fullCols: number;
  readonly ref?: RangeRef;
  at(i: number, j: number): Scalar;
}

export const toGrid = (v: Value, ctx: FnContext): Grid | CellError => {
  const m = toMatrix(v, ctx);
  if (isError(m)) return m;
  const [fr, fc] = isRef(v) ? refArea(v) : [m.length, m[0]?.length ?? 0];
  return { m, rows: m.length, cols: m[0]?.length ?? 0, fullRows: fr, fullCols: fc, ref: isRef(v) ? v : undefined, at: (i, j) => m[i]?.[j] ?? null };
};

/** spreadsheet wildcard pattern (`*`, `?`, `~` escape) as a whole-string, case-insensitive RegExp. */
export const wildcardRegex = (pattern: string): RegExp => {
  let src = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i]!;
    if (ch === '~' && i + 1 < pattern.length && '*?~'.includes(pattern[i + 1]!)) src += '\\' + pattern[++i];
    else if (ch === '*') src += '[\\s\\S]*';
    else if (ch === '?') src += '[\\s\\S]';
    else src += ch.replace(/[.+^${}()|[\]\\/-]/g, '\\$&');
  }
  return new RegExp(`^${src}$`, 'i');
};

export const hasWildcard = (s: string): boolean => /[*?~]/.test(s);

/** Text equality the way spreadsheet lookups and criteria compare: case-insensitive, optional wildcards. */
export const textMatcher = (pattern: string, wildcards: boolean): ((s: string) => boolean) => {
  if (wildcards && hasWildcard(pattern)) {
    const re = wildcardRegex(pattern);
    return (s) => re.test(s);
  }
  const p = pattern.toLowerCase();
  return (s) => s.toLowerCase() === p;
};

export const textCompare = (a: string, b: string): number => a.localeCompare(b, undefined, { sensitivity: 'accent' });

type Op = '=' | '<>' | '<' | '>' | '<=' | '>=';

const cmpOp = (op: Op, c: number): boolean =>
  op === '=' ? c === 0 : op === '<>' ? c !== 0 : op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0;

const numEq = (v: Scalar, n: number, date1904: boolean): boolean => {
  if (typeof v === 'number') return approxEqual(v, n);
  // COUNTIF-family equality also matches numbers stored as text ("5" matches 5).
  if (typeof v === 'string') { const p = parseNumber(v, date1904); return p !== null && approxEqual(p, n); }
  return false;
};

/**
 * Builds a predicate for the criteria argument of COUNTIF, SUMIFS and friends.
 * A blank criteria cell counts as 0. Wildcards only apply to `=`/`<>` text criteria, and ordering
 * operators only compare values of the same type (numbers with numbers, text with text).
 */
export const makeCriteria = (c: Scalar, date1904 = false): ((v: Scalar) => boolean) => {
  if (c === null) return (v) => numEq(v, 0, date1904);
  if (typeof c === 'number') return (v) => numEq(v, c, date1904);
  if (typeof c === 'boolean') return (v) => v === c;
  if (isError(c)) return (v) => isError(v) && v.code === c.code;

  const m = /^(<=|>=|<>|=|<|>)?([\s\S]*)$/.exec(c)!;
  const op = (m[1] ?? '=') as Op;
  const rest = m[2]!;

  if (rest === '') {
    if (!m[1]) return (v) => v === null || v === '';
    if (op === '=') return (v) => v === null;
    if (op === '<>') return (v) => v !== null;
  }

  const n = parseNumber(rest, date1904);
  if (n !== null) {
    if (op === '=') return (v) => numEq(v, n, date1904);
    if (op === '<>') return (v) => !numEq(v, n, date1904);
    return (v) => typeof v === 'number' && cmpOp(op, approxEqual(v, n) ? 0 : v - n);
  }

  const up = rest.toUpperCase();
  if (up === 'TRUE' || up === 'FALSE') {
    const b = up === 'TRUE';
    if (op === '<>') return (v) => v !== b;
    return (v) => typeof v === 'boolean' && cmpOp(op, Number(v) - Number(b));
  }
  if (/^#/.test(rest)) {
    const code = up;
    const isErr = (v: Scalar): boolean => v instanceof CellError && v.code === code;
    if (op === '=') return isErr;
    if (op === '<>') return (v) => !isErr(v);
  }

  if (op === '=' || op === '<>') {
    const t = textMatcher(rest, true);
    const eq = (v: Scalar): boolean => typeof v === 'string' && t(v);
    return op === '=' ? eq : (v) => !eq(v);
  }
  return (v) => typeof v === 'string' && cmpOp(op, textCompare(v, rest));
};
