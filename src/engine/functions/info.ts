import { Err, isError, isLambda, isMatrix, isRef, isRefs, type FnContext, type FunctionDef, type RangeRef, type Scalar, type Value } from '../types';
import { first, num, text } from '../coerce';

const colName = (c: number): string => {
  let s = '';
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
};

const is = (name: string, test: (v: Scalar) => boolean): FunctionDef => ({ name, min: 1, max: 1, scalar: true, fn: ([v]) => test(v as Scalar) });

const parity = (name: string, odd: boolean): FunctionDef => ({
  name, min: 1, max: 1, scalar: true,
  fn: ([v], ctx) => {
    if (typeof v === 'boolean') return Err.value;
    const n = num(v!, ctx);
    return isError(n) ? n : (Math.abs(Math.trunc(n)) % 2 === 1) === odd;
  },
});

/** The reference a function argument names, if it is one. */
const refOf = (v: Value | undefined): RangeRef | null => (v !== undefined && isRef(v) ? v : v !== undefined && isRefs(v) ? v.areas[0] ?? null : null);

/** CELL("format") codes, from spreadsheet's documented table. */
const formatCode = (fmt: string): string => {
  const f = fmt.split(';')[0]!.replace(/\[[^\]]*\]|"[^"]*"|\\./g, '');
  const paren = /\(/.test(fmt.split(';')[1] ?? '') ? '()' : '';
  const neg = /\[red\]/i.test(fmt.split(';')[1] ?? '') ? '-' : '';
  if (/^general$/i.test(fmt) || fmt === '') return 'G';
  const decimals = (/\.(0+)/.exec(f)?.[1] ?? '').length;
  if (/[dmyhs]/i.test(f)) {
    if (/h/i.test(f)) return /am\/pm|a\/p/i.test(f) ? (/s/i.test(f) ? 'D6' : 'D7') : /s/i.test(f) ? 'D8' : 'D9';
    if (/d/i.test(f) && /y/i.test(f)) return /mmm/i.test(f) ? 'D1' : 'D4';
    if (/d/i.test(f)) return 'D2';
    return /mmm/i.test(f) ? 'D3' : 'D5';
  }
  if (/%/.test(f)) return `P${decimals}`;
  if (/E\+/i.test(f)) return `S${decimals}`;
  if (/[$€£¥]/.test(fmt)) return `C${decimals}${neg}${paren}`;
  if (/,/.test(f)) return `,${decimals}${neg}${paren}`;
  return `F${decimals}${neg}${paren}`;
};

const cellInfo = (args: Value[], ctx: FnContext): Value => {
  const kind = text(args[0]!, ctx);
  if (isError(kind)) return kind;
  const target = args.length > 1 ? args[1] : undefined;
  if (target !== undefined && !isRef(target) && !isRefs(target)) return Err.value;
  const r = refOf(target) ?? { kind: 'ref' as const, sheet: ctx.sheet, r1: ctx.row, c1: ctx.col, r2: ctx.row, c2: ctx.col };
  const v = ctx.cell(r.sheet, r.r1, r.c1);
  const fmt = ctx.numberFormat(r.sheet, r.r1, r.c1);
  switch (kind.toLowerCase()) {
    case 'address': {
      const a = `$${colName(r.c1)}$${r.r1 + 1}`;
      if (r.sheet === ctx.sheet) return a;
      const n = ctx.sheetName(r.sheet);
      return `${/^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`}!${a}`;
    }
    case 'col': return r.c1 + 1;
    case 'row': return r.r1 + 1;
    case 'contents': return v;
    case 'type': return v === null ? 'b' : typeof v === 'string' ? 'l' : 'v';
    case 'width': return [[8, true]];
    case 'format': return formatCode(fmt);
    case 'color': return /;\[red\]/i.test(fmt) ? 1 : 0;
    case 'parentheses': return /;\(|^\(/.test(fmt) ? 1 : 0;
    case 'prefix': return typeof v === 'string' ? "'" : '';
    case 'protect': return 1;
    case 'filename': return '';
    case 'sheet': return ctx.sheetName(r.sheet);
    default: return Err.value;
  }
};

const ERROR_NUMBERS: Partial<Record<string, number>> = {
  '#NULL!': 1, '#DIV/0!': 2, '#VALUE!': 3, '#REF!': 4, '#NAME?': 5, '#NUM!': 6, '#N/A': 7, '#GETTING_DATA': 8,
  '#SPILL!': 9, '#CONNECT!': 10, '#BLOCKED!': 11, '#UNKNOWN!': 12, '#FIELD!': 13, '#CALC!': 14, '#BUSY!': 19,
};

export const defs: FunctionDef[] = [
  { name: 'CELL', min: 1, max: 2, fn: cellInfo },
  {
    name: 'ERROR.TYPE', min: 1, max: 1, scalar: true,
    fn: ([v]) => (isError(v) ? ERROR_NUMBERS[v.code] ?? Err.na : Err.na),
  },
  {
    name: 'INFO', min: 1, max: 1, scalar: true,
    fn: ([t], ctx) => {
      const k = text(t!, ctx);
      if (isError(k)) return k;
      switch (k.toLowerCase()) {
        case 'directory': return '';
        case 'numfile': return ctx.sheetCount();
        case 'origin': return '$A:$A$1';
        case 'osversion': return 'Windows (64-bit) NT 10.00';
        case 'recalc': return 'Automatic';
        case 'release': return '16.0';
        case 'system': return 'pcdos';
        default: return Err.value;
      }
    },
  },
  is('ISBLANK', (v) => v === null),
  is('ISERR', (v) => isError(v) && v.code !== '#N/A'),
  is('ISERROR', (v) => isError(v)),
  is('ISLOGICAL', (v) => typeof v === 'boolean'),
  is('ISNA', (v) => isError(v) && v.code === '#N/A'),
  is('ISNONTEXT', (v) => typeof v !== 'string'),
  is('ISNUMBER', (v) => typeof v === 'number'),
  is('ISTEXT', (v) => typeof v === 'string'),
  parity('ISEVEN', false),
  parity('ISODD', true),
  {
    name: 'ISFORMULA', min: 1, max: 1,
    fn: ([v], ctx) => { const r = refOf(v); return r ? ctx.formula(r.sheet, r.r1, r.c1) !== null : Err.value; },
  },
  {
    name: 'FORMULATEXT', min: 1, max: 1,
    fn: ([v], ctx) => {
      const r = refOf(v);
      if (!r) return isError(v) ? v : Err.na;
      const f = ctx.formula(r.sheet, r.r1, r.c1);
      return f === null ? Err.na : f.startsWith('=') ? f : `=${f}`;
    },
  },
  { name: 'ISREF', min: 1, max: 1, fn: ([v]) => isRef(v!) || isRefs(v!) },
  {
    name: 'N', min: 1, max: 1, scalar: true,
    fn: ([v]) => (typeof v === 'number' ? v : typeof v === 'boolean' ? Number(v) : isError(v) ? v : 0),
  },
  { name: 'NA', min: 0, max: 0, fn: () => Err.na },
  {
    name: 'SHEET', min: 0, max: 1,
    fn: ([v], ctx) => {
      if (v === undefined || v === null) return ctx.sheet + 1;
      const r = refOf(v);
      if (r) return r.sheet + 1;
      const s = first(v, ctx);
      if (typeof s !== 'string') return isError(s) ? s : Err.na;
      const i = ctx.sheetIndex(s);
      return i < 0 ? Err.na : i + 1;
    },
  },
  {
    name: 'SHEETS', min: 0, max: 1,
    fn: ([v], ctx) => {
      if (v === undefined || v === null) return ctx.sheetCount();
      if (isRefs(v)) return new Set(v.areas.map((a) => a.sheet)).size;
      return isRef(v) ? 1 : Err.na;
    },
  },
  {
    name: 'TYPE', min: 1, max: 1,
    fn: ([v], ctx) => {
      if (isLambda(v!)) return 128;
      if (isMatrix(v!) || isRefs(v!)) return 64;
      if (isRef(v!)) {
        if (v.r1 !== v.r2 || v.c1 !== v.c2) return 64;
        return typeOf(ctx.cell(v.sheet, v.r1, v.c1));
      }
      return typeOf(v as Scalar);
    },
  },
];

const typeOf = (v: Scalar): number =>
  v === null || typeof v === 'number' ? 1 : typeof v === 'string' ? 2 : typeof v === 'boolean' ? 4 : 16;
