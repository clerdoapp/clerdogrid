// Parsed-formula tokens (Ptg) of the binary formats: decoding to ClerdoGrid formula text and
// compiling formula text back to tokens. MS-XLS 2.5.198 (Rgce, 16-bit rows) and MS-XLSB 2.5.97
// (32-bit rows, 14-bit columns) share one implementation with a width switch.
import { nodeText, parseFormula, type Node, type RefNode } from '../engine/parser';
import type { ErrorCode } from '../engine/types';
import { stripFuturePrefixes } from './formula-xml';
import { BIFF8_MAX_FTAB, REF_ARGS, REF_RESULT, VOLATILE, ftabByIndex, ftabByName } from './biff8-ftab';

export const ERR_BY_CODE: Record<number, ErrorCode> = {
  0x00: '#NULL!', 0x07: '#DIV/0!', 0x0f: '#VALUE!', 0x17: '#REF!', 0x1d: '#NAME?', 0x24: '#NUM!', 0x2a: '#N/A', 0x2b: '#GETTING_DATA',
};
export const CODE_BY_ERR: Record<string, number> = Object.fromEntries(Object.entries(ERR_BY_CODE).map(([k, v]) => [v, +k]));

export interface SheetSpan { first: string; last?: string }

export interface DecodeCtx {
  xlsb: boolean;
  /** Older (BIFF5) tokens: 8-bit columns, flags in the row word, code-page strings. */
  biff5?: { text: (b: Uint8Array) => string; sheetsByTab: (first: number, last: number) => SheetSpan | null };
  /** Cell the formula belongs to; relative (shared-formula) tokens are offsets from it. */
  row: number;
  col: number;
  /** Sheets for an EXTERNSHEET / XTI index; null for a deleted or external sheet. */
  sheets: (ixti: number) => SheetSpan | null;
  /** Defined name by 1-based index. */
  name: (index: number) => { name: string; sheet?: string } | undefined;
  /** External or add-in name (tNameX). */
  externName: (ixti: number, index: number) => string | undefined;
  /** Table for a structured reference (xlsb PtgList). */
  table?: (id: number) => { name: string; columns: string[] } | undefined;
}

export class Exp { constructor(readonly row: number, readonly col: number) {} }
export class UnsupportedFormula extends Error {}

class Reader {
  o = 0;
  private dv: DataView;
  constructor(readonly b: Uint8Array) { this.dv = new DataView(b.buffer, b.byteOffset, b.byteLength); }
  need(n: number): void { if (this.o + n > this.b.length) throw new UnsupportedFormula('Formula data ends early.'); }
  u8(): number { this.need(1); return this.b[this.o++]!; }
  u16(): number { this.need(2); const v = this.dv.getUint16(this.o, true); this.o += 2; return v; }
  i16(): number { this.need(2); const v = this.dv.getInt16(this.o, true); this.o += 2; return v; }
  u32(): number { this.need(4); const v = this.dv.getUint32(this.o, true); this.o += 4; return v; }
  i32(): number { this.need(4); const v = this.dv.getInt32(this.o, true); this.o += 4; return v; }
  f64(): number { this.need(8); const v = this.dv.getFloat64(this.o, true); this.o += 8; return v; }
  skip(n: number): void { this.need(n); this.o += n; }
  utf16(n: number): string { this.need(n * 2); let s = ''; for (let i = 0; i < n; i++) s += String.fromCharCode(this.dv.getUint16(this.o + i * 2, true)); this.o += n * 2; return s; }
  latin1(n: number): string { this.need(n); let s = ''; for (let i = 0; i < n; i++) s += String.fromCharCode(this.b[this.o + i]!); this.o += n; return s; }
  /** XLUnicodeStringNoCch body. */
  xlChars(n: number): string { const hi = this.u8() & 1; return hi ? this.utf16(n) : this.latin1(n); }
}

const err = (code: ErrorCode): Node => ({ type: 'err', code });
const BIN: Record<number, Extract<Node, { type: 'binary' }>['op']> = {
  0x03: '+', 0x04: '-', 0x05: '*', 0x06: '/', 0x07: '^', 0x08: '&', 0x09: '<', 0x0a: '<=', 0x0b: '=', 0x0c: '>=', 0x0d: '>', 0x0e: '<>', 0x0f: ' ', 0x11: ':',
};

/** Decodes a token array. Returns formula text (no "="), or an Exp marker for shared/array formula cells. */
export const decodeRgce = (rgce: Uint8Array, rgcb: Uint8Array, ctx: DecodeCtx): string | Exp => {
  const r = new Reader(rgce);
  const extra = new Reader(rgcb);
  const wide = ctx.xlsb;
  const MAXR = wide ? 1048576 : 65536;
  const MAXC = wide ? 16384 : 256;
  const stack: Node[] = [];
  const pop = (): Node => { const n = stack.pop(); if (!n) throw new UnsupportedFormula('Formula tokens are out of order.'); return n; };
  const rowOf = (): number => (wide ? r.u32() : r.u16());
  const mkRef = (r1: number, c1w: number, r2: number, c2w: number, area: boolean, rel: boolean): RefNode => {
    const c1Rel = !!(c1w & 0x4000), r1Rel = !!(c1w & 0x8000), c2Rel = !!(c2w & 0x4000), r2Rel = !!(c2w & 0x8000);
    let c1 = c1w & 0x3fff, c2 = c2w & 0x3fff;
    if (rel) {
      // offsets from the formula's cell, wrapped to the sheet size
      const sCol = (c: number) => (wide ? ((c << 18) >> 18) : ((c & 0xff) << 24) >> 24);
      const sRow = (v: number) => (wide ? v | 0 : (v << 16) >> 16);
      if (r1Rel) r1 = (((ctx.row + sRow(r1)) % MAXR) + MAXR) % MAXR;
      if (r2Rel) r2 = (((ctx.row + sRow(r2)) % MAXR) + MAXR) % MAXR;
      if (c1Rel) c1 = (((ctx.col + sCol(c1)) % MAXC) + MAXC) % MAXC;
      if (c2Rel) c2 = (((ctx.col + sCol(c2)) % MAXC) + MAXC) % MAXC;
    } else if (!wide) { c1 &= 0xff; c2 &= 0xff; }
    const ref: RefNode = {
      r1: Math.min(r1, r2), r2: Math.max(r1, r2), c1: Math.min(c1, c2), c2: Math.max(c1, c2),
      r1Abs: !r1Rel, c1Abs: !c1Rel, r2Abs: !r2Rel, c2Abs: !c2Rel, kind: area ? 'area' : 'cell',
    };
    if (area && ref.r1 === 0 && ref.r2 === MAXR - 1) { ref.kind = 'cols'; ref.r2 = 1048575; }
    else if (area && ref.c1 === 0 && ref.c2 === MAXC - 1) { ref.kind = 'rows'; ref.c2 = 16383; }
    return ref;
  };
  const b5 = ctx.biff5;
  // BIFF5 keeps the relative flags in the row word; move them to the column word as BIFF8 has them
  const b5word = (rw: number, col: number) => ({ row: rw & 0x3fff, col: col | (rw & 0xc000) });
  const readRef = (rel: boolean): RefNode => {
    if (b5) { const rw = r.u16(), col = r.u8(); const x = b5word(rw, col); return mkRef(x.row, x.col, x.row, x.col, false, rel); }
    const row = rowOf(); const col = r.u16(); return mkRef(row, col, row, col, false, rel);
  };
  const readArea = (rel: boolean): RefNode => {
    if (b5) { const w1 = r.u16(), w2 = r.u16(), c1 = r.u8(), c2 = r.u8(); const a = b5word(w1, c1), b = b5word(w2, c2); return mkRef(a.row, a.col, b.row, b.col, true, rel); }
    const r1 = rowOf(), r2 = rowOf(), c1 = r.u16(), c2 = r.u16(); return mkRef(r1, c1, r2, c2, true, rel);
  };
  const b5Sheets = (ref: RefNode): Node => {
    const ixals = r.i16(); r.skip(8); const t1 = r.i16(), t2 = r.i16();
    const s = ixals < 0 ? b5!.sheetsByTab(t1, t2) : null;
    if (!s) return err('#REF!');
    ref.sheet = s.first;
    if (s.last && s.last !== s.first) ref.sheet2 = s.last;
    return { type: 'ref', ref };
  };
  const withSheet = (ref: RefNode, ixti: number): Node => {
    const s = ctx.sheets(ixti);
    if (!s) return err('#REF!');
    ref.sheet = s.first;
    if (s.last && s.last !== s.first) ref.sheet2 = s.last;
    return { type: 'ref', ref };
  };
  const callNode = (name: string, args: Node[]): Node => {
    const clean = stripFuturePrefixes(name).replace(/^_xlws\./i, '');
    if (clean === 'TRUE' && !args.length) return { type: 'bool', value: true };
    if (clean === 'FALSE' && !args.length) return { type: 'bool', value: false };
    return { type: 'call', name: clean.toUpperCase(), args };
  };
  const readArray = (): Node => {
    let rows: number, cols: number;
    if (wide) { rows = extra.u32(); cols = extra.u32(); }
    else { cols = extra.u8() + 1; rows = extra.u16() + 1; }
    if (rows * cols > 1_000_000) throw new UnsupportedFormula('Array constant is too large.');
    const out: Node[][] = [];
    for (let i = 0; i < rows; i++) {
      const row: Node[] = [];
      for (let j = 0; j < cols; j++) {
        const t = extra.u8();
        if (wide) {
          if (t === 0x00) row.push({ type: 'num', value: extra.f64() });
          else if (t === 0x01) row.push({ type: 'str', value: extra.utf16(extra.u16()) });
          else if (t === 0x02) row.push({ type: 'bool', value: !!extra.u8() });
          else if (t === 0x04) { row.push(err(ERR_BY_CODE[extra.u8()] ?? '#N/A')); }
          else throw new UnsupportedFormula('Unknown array value.');
        } else if (t === 0x00) { extra.skip(8); row.push({ type: 'str', value: '' }); }
        else if (t === 0x01) row.push({ type: 'num', value: extra.f64() });
        else if (t === 0x02) {
          if (b5) { const n = extra.u8(); const bytes = new Uint8Array(n); for (let k = 0; k < n; k++) bytes[k] = extra.u8(); row.push({ type: 'str', value: b5.text(bytes) }); }
          else { const n = extra.u16(); row.push({ type: 'str', value: extra.xlChars(n) }); }
        }
        else if (t === 0x04) { row.push({ type: 'bool', value: !!extra.u8() }); extra.skip(7); }
        else if (t === 0x10) { row.push(err(ERR_BY_CODE[extra.u8()] ?? '#N/A')); extra.skip(7); }
        else throw new UnsupportedFormula('Unknown array value.');
      }
      out.push(row);
    }
    return { type: 'array', rows: out };
  };

  while (r.o < rgce.length) {
    const ptg = r.u8();
    if (ptg === 0x01) {
      const row = wide ? r.u32() : r.u16();
      const col = wide ? ctx.col : r.u16();
      if (!stack.length && r.o >= rgce.length) return new Exp(row, col);
      throw new UnsupportedFormula('Unexpected shared formula token.');
    }
    if (ptg === 0x02) throw new UnsupportedFormula('Data table formula.');
    if (BIN[ptg]) { const b = pop(), a = pop(); stack.push({ type: 'binary', op: BIN[ptg]!, left: a, right: b }); continue; }
    switch (ptg) {
      case 0x10: { const b = pop(), a = pop(); stack.push({ type: 'union', items: [...(a.type === 'union' ? a.items : [a]), ...(b.type === 'union' ? b.items : [b])] }); continue; }
      case 0x12: stack.push({ type: 'unary', op: '+', arg: pop() }); continue;
      case 0x13: stack.push({ type: 'unary', op: '-', arg: pop() }); continue;
      case 0x14: stack.push({ type: 'unary', op: '%', arg: pop() }); continue;
      case 0x15: continue;
      case 0x16: stack.push({ type: 'missing' }); continue;
      case 0x17: {
        const n = wide ? r.u16() : r.u8();
        if (b5) { r.need(n); stack.push({ type: 'str', value: b5.text(rgce.subarray(r.o, r.o + n)) }); r.skip(n); continue; }
        stack.push({ type: 'str', value: wide ? r.utf16(n) : r.xlChars(n) }); continue;
      }
      case 0x18: {
        const eptg = r.u8();
        if (eptg === 0x19 && wide) {
          // PtgList: structured table reference
          r.u16(); // ixti
          const flags = r.u16();
          const listId = r.u32();
          const colFirst = r.u16(), colLast = r.u16();
          const t = ctx.table?.(listId);
          if (!t) { stack.push(err('#REF!')); continue; }
          const rowType = (flags >> 2) & 0x1f;
          const items: ('#All' | '#Data' | '#Headers' | '#Totals' | '#This Row')[] = [];
          const map: Record<number, (typeof items)[number][]> = { 0: [], 1: ['#All'], 2: ['#Data'], 4: ['#Headers'], 6: ['#Headers', '#Data'], 8: ['#Totals'], 0xc: ['#Data', '#Totals'], 0x10: ['#This Row'] };
          items.push(...(map[rowType] ?? []));
          const colMode = flags & 3;
          const spec: { items: typeof items; col1?: string; col2?: string } = { items };
          if (colMode !== 0) {
            spec.col1 = t.columns[colFirst];
            if (colMode === 2 && colLast !== colFirst) spec.col2 = t.columns[colLast];
          }
          stack.push({ type: 'table', table: t.name, spec });
          continue;
        }
        throw new UnsupportedFormula('Unknown extended token.');
      }
      case 0x19: {
        const kind = r.u8();
        if (kind & 0x04) { const n = r.u16(); r.skip((n + 1) * 2); continue; }
        if (kind & 0x10) { r.u16(); stack.push({ type: 'call', name: 'SUM', args: [pop()] }); continue; }
        r.u16();
        continue;
      }
      case 0x1c: stack.push(err(ERR_BY_CODE[r.u8()] ?? '#N/A')); continue;
      case 0x1d: stack.push({ type: 'bool', value: !!r.u8() }); continue;
      case 0x1e: stack.push({ type: 'num', value: r.u16() }); continue;
      case 0x1f: stack.push({ type: 'num', value: r.f64() }); continue;
      default: break;
    }
    const base = (ptg & 0x1f) | 0x20;
    if (ptg < 0x20 || ptg > 0x7f) throw new UnsupportedFormula(`Unknown token 0x${ptg.toString(16)}.`);
    switch (base) {
      case 0x20: r.skip(wide ? 14 : 7); stack.push(readArray()); break;
      case 0x21: {
        const idx = r.u16();
        const f = ftabByIndex(idx);
        if (!f || f.argc < 0) throw new UnsupportedFormula(`Unknown function ${idx}.`);
        const args = stack.splice(stack.length - f.argc, f.argc);
        if (args.length !== f.argc) throw new UnsupportedFormula('Function is missing arguments.');
        stack.push(callNode(f.name, args));
        break;
      }
      case 0x22: {
        const argc = r.u8() & 0x7f;
        const idx = r.u16() & 0x7fff;
        if (argc > stack.length) throw new UnsupportedFormula('Function is missing arguments.');
        const args = stack.splice(stack.length - argc, argc);
        if (idx === 255) {
          const fn = args.shift();
          const name = fn?.type === 'name' ? fn.name : fn?.type === 'str' ? fn.value : undefined;
          if (!name) { stack.push(err('#NAME?')); break; }
          stack.push(callNode(name, args));
        } else {
          const f = ftabByIndex(idx);
          if (!f) { stack.push(err('#NAME?')); break; }
          stack.push(callNode(f.name, args));
        }
        break;
      }
      case 0x23: {
        const idx = wide ? r.u32() : r.u16();
        if (!wide) r.skip(b5 ? 12 : 2);
        const n = ctx.name(idx);
        stack.push(n ? { type: 'name', name: n.name, sheet: n.sheet } : err('#NAME?'));
        break;
      }
      case 0x24: stack.push({ type: 'ref', ref: readRef(false) }); break;
      case 0x25: stack.push({ type: 'ref', ref: readArea(false) }); break;
      case 0x26: case 0x27: {
        r.skip(4); r.u16();
        if (base === 0x26) { const n = extra.u16(); extra.skip(n * (wide ? 16 : b5 ? 6 : 8)); }
        break;
      }
      case 0x28: case 0x29: r.skip(base === 0x28 ? 6 : 2); break;
      case 0x2a: r.skip(wide ? 6 : b5 ? 3 : 4); stack.push(err('#REF!')); break;
      case 0x2b: r.skip(wide ? 12 : b5 ? 6 : 8); stack.push(err('#REF!')); break;
      case 0x2c: stack.push({ type: 'ref', ref: readRef(true) }); break;
      case 0x2d: stack.push({ type: 'ref', ref: readArea(true) }); break;
      case 0x39: {
        if (b5) {
          r.i16(); r.skip(8); const idx = r.u16(); r.skip(12);
          const n = ctx.name(idx);
          stack.push(n ? { type: 'name', name: n.name, sheet: n.sheet } : err('#NAME?'));
          break;
        }
        const ixti = r.u16();
        const idx = r.u32();
        const nm = ctx.externName(ixti, idx);
        stack.push(nm ? { type: 'name', name: nm } : err('#NAME?'));
        break;
      }
      case 0x3a: case 0x3b: {
        if (b5) {
          const start = r.o;
          r.skip(14);
          const ref = base === 0x3a ? readRef(false) : readArea(false);
          const end = r.o;
          r.o = start;
          stack.push(b5Sheets(ref));
          r.o = end;
          break;
        }
        const ixti = r.u16();
        stack.push(withSheet(base === 0x3a ? readRef(false) : readArea(false), ixti));
        break;
      }
      case 0x3c: r.skip(b5 ? 14 + 3 : 2 + (wide ? 6 : 4)); stack.push(err('#REF!')); break;
      case 0x3d: r.skip(b5 ? 14 + 6 : 2 + (wide ? 12 : 8)); stack.push(err('#REF!')); break;
      default: throw new UnsupportedFormula(`Unknown token 0x${ptg.toString(16)}.`);
    }
  }
  if (stack.length !== 1) throw new UnsupportedFormula('Formula tokens are out of order.');
  return nodeText(stack[0]!);
};

/** Text for a name or add-in function reference left on its own (e.g. "_xlfn.IFERROR"). */
export const cleanName = (s: string): string => stripFuturePrefixes(s);

// ---------- Compiling ----------

export interface EncodeCtx {
  xlsb?: boolean;
  /** 'cell': plain references; 'name': every reference is 3D; 'rel': references relative to (row, col). */
  mode: 'cell' | 'name' | 'rel';
  row?: number;
  col?: number;
  /** Index of a sheet by name, or undefined. */
  sheetIndex: (name: string) => number | undefined;
  /** EXTERNSHEET index for a sheet span (-2 means the workbook itself, for names). */
  ixti: (first: number, last: number) => number;
  /** 1-based defined-name index. */
  nameIndex: (name: string, sheet?: number) => number | undefined;
  /** 1-based index of the hidden function name used for a function the table doesn't hold. */
  addin: (name: string) => number | { ixti: number; index: number };
  /** Expands a structured reference into a plain reference. */
  table?: (n: Extract<Node, { type: 'table' }>) => RefNode | undefined;
}

type Cls = 'R' | 'V' | 'A';
const CLS: Record<Cls, number> = { R: 0x20, V: 0x40, A: 0x60 };
const PREC: Record<string, number> = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5, ' ': 8, ':': 9 };
const OPS: Record<string, number> = Object.fromEntries(Object.entries(BIN).map(([k, v]) => [v, +k]));

class Out {
  bytes: number[] = [];
  extra: number[] = [];
  u8(v: number): void { this.bytes.push(v & 0xff); }
  u16(v: number): void { this.bytes.push(v & 0xff, (v >> 8) & 0xff); }
  u32(v: number): void { this.bytes.push(v & 0xff, (v >> 8) & 0xff, (v >> 16) & 0xff, (v >>> 24) & 0xff); }
  f64(v: number, to = this.bytes): void { const b = new Uint8Array(new Float64Array([v]).buffer); to.push(...b); }
}

/** Compiles formula text (no "=") to tokens. Throws on formulas that can't be stored. */
export const encodeFormula = (formula: string, ctx: EncodeCtx, arrayFormula = false): { rgce: Uint8Array; rgcb: Uint8Array } => {
  const tree = parseFormula(formula);
  const o = new Out();
  const wide = !!ctx.xlsb;
  const MAXR = wide ? 1048576 : 65536, MAXC = wide ? 16384 : 256;
  const top: Cls = arrayFormula ? 'A' : 'V';

  const colWord = (c: number, abs: boolean, rAbs: boolean, rel: boolean): number => {
    let v = c;
    if (rel && !abs) v = ((c - (ctx.col ?? 0)) & (wide ? 0x3fff : 0xff));
    return (v & 0x3fff) | (abs ? 0 : 0x4000) | (rAbs ? 0 : 0x8000);
  };
  const rowWord = (r: number, abs: boolean, rel: boolean): number => {
    if (rel && !abs) return (r - (ctx.row ?? 0)) & (wide ? 0xffffffff : 0xffff);
    return r;
  };
  const putRow = (v: number) => (wide ? o.u32(v) : o.u16(v));

  const ref = (rf: RefNode, cls: Cls): void => {
    if (rf.external) { o.u8(0x1c); o.u8(0x17); return; }
    let r1 = rf.r1, r2 = rf.r2, c1 = rf.c1, c2 = rf.c2;
    if (rf.kind === 'cols') { r1 = 0; r2 = MAXR - 1; }
    if (rf.kind === 'rows') { c1 = 0; c2 = MAXC - 1; }
    const area = rf.kind !== 'cell';
    const bad = r2 >= MAXR || c2 >= MAXC;
    let sheet3d: number | undefined;
    if (rf.sheet !== undefined || ctx.mode === 'name') {
      const a = rf.sheet !== undefined ? ctx.sheetIndex(rf.sheet) : undefined;
      const b = rf.sheet2 !== undefined ? ctx.sheetIndex(rf.sheet2) : a;
      if (rf.sheet !== undefined && (a === undefined || b === undefined)) { o.u8(0x1c); o.u8(0x17); return; }
      sheet3d = ctx.ixti(a ?? 0, b ?? a ?? 0);
    }
    const rel = ctx.mode === 'rel' && sheet3d === undefined;
    if (sheet3d !== undefined) {
      o.u8((area ? (bad ? 0x3d : 0x3b) : (bad ? 0x3c : 0x3a)) - 0x20 + CLS[cls]);
      o.u16(sheet3d);
    } else o.u8((area ? (bad ? 0x2b : rel ? 0x2d : 0x25) : (bad ? 0x2a : rel ? 0x2c : 0x24)) - 0x20 + CLS[cls]);
    if (bad) { for (let i = 0; i < (area ? (wide ? 12 : 8) : (wide ? 6 : 4)); i++) o.u8(0); return; }
    if (area) {
      putRow(rowWord(r1, rf.r1Abs, rel)); putRow(rowWord(r2, rf.r2Abs, rel));
      o.u16(colWord(c1, rf.c1Abs, rf.r1Abs, rel)); o.u16(colWord(c2, rf.c2Abs, rf.r2Abs, rel));
    } else { putRow(rowWord(r1, rf.r1Abs, rel)); o.u16(colWord(c1, rf.c1Abs, rf.r1Abs, rel)); }
  };

  const str = (s: string): void => {
    const t = s.slice(0, 255);
    o.u8(0x17);
    if (wide) { o.u16(t.length); for (const ch of t) o.u16(ch.charCodeAt(0)); return; }
    o.u8(t.length);
    const hi = /[^\x00-\xff]/.test(t);
    o.u8(hi ? 1 : 0);
    for (let i = 0; i < t.length; i++) { if (hi) o.u16(t.charCodeAt(i)); else o.u8(t.charCodeAt(i)); }
  };

  const array = (n: Extract<Node, { type: 'array' }>, cls: Cls): void => {
    o.u8(0x20 - 0x20 + CLS[cls === 'R' ? 'A' : cls]);
    for (let i = 0; i < (wide ? 14 : 7); i++) o.u8(0);
    const rows = n.rows.length, cols = n.rows[0]?.length ?? 0;
    const e = o.extra;
    const p16 = (v: number) => e.push(v & 0xff, (v >> 8) & 0xff);
    const p32 = (v: number) => e.push(v & 0xff, (v >> 8) & 0xff, (v >> 16) & 0xff, (v >>> 24) & 0xff);
    if (wide) { p32(rows); p32(cols); } else { e.push((cols - 1) & 0xff); p16(rows - 1); }
    for (const row of n.rows) for (const x0 of row) {
      let x = x0;
      if (x.type === 'unary' && x.op === '-' && x.arg.type === 'num') x = { type: 'num', value: -x.arg.value };
      if (wide) {
        if (x.type === 'num') { e.push(0); o.f64(x.value, e); }
        else if (x.type === 'str') { e.push(1); p16(x.value.length); for (const ch of x.value) p16(ch.charCodeAt(0)); }
        else if (x.type === 'bool') { e.push(2, x.value ? 1 : 0); }
        else if (x.type === 'err') { e.push(4, CODE_BY_ERR[x.code] ?? 0x2a); }
        else { e.push(0); o.f64(0, e); }
      } else if (x.type === 'num') { e.push(1); o.f64(x.value, e); }
      else if (x.type === 'str') { e.push(2); p16(x.value.length); e.push(1); for (const ch of x.value) p16(ch.charCodeAt(0)); }
      else if (x.type === 'bool') { e.push(4, x.value ? 1 : 0, 0, 0, 0, 0, 0, 0); }
      else if (x.type === 'err') { e.push(0x10, CODE_BY_ERR[x.code] ?? 0x2a, 0, 0, 0, 0, 0, 0); }
      else e.push(0, 0, 0, 0, 0, 0, 0, 0, 0);
    }
  };

  const paren = (child: Node, parent: number, right = false): boolean => {
    const p = child.type === 'binary' ? PREC[child.op]! : child.type === 'unary' && child.op !== '%' ? 6 : 99;
    return right ? p <= parent : p < parent;
  };

  const call = (name: string, args: Node[], cls: Cls): void => {
    const up = name.toUpperCase();
    const f = ftabByName(up);
    const direct = f && (wide || f.index <= BIFF8_MAX_FTAB);
    const refArgs = REF_ARGS[up];
    const argCls = (i: number): Cls => (refArgs === 'all' || refArgs?.includes(i) ? 'R' : cls === 'A' ? 'A' : 'V');
    if (VOLATILE.has(up)) { o.u8(0x19); o.u8(0x01); o.u16(0); }
    if (!direct) {
      const idx = ctx.addin(up);
      if (typeof idx === 'number') { o.u8(0x23); if (wide) o.u32(idx); else { o.u16(idx); o.u16(0); } }
      else { o.u8(0x39); o.u16(idx.ixti); o.u32(idx.index); }
    }
    args.forEach((a, i) => emit(a, argCls(i)));
    const fcls: Cls = cls === 'R' && REF_RESULT.has(up) ? 'R' : cls === 'A' ? 'A' : 'V';
    if (!direct) { o.u8(0x22 - 0x20 + CLS[fcls]); o.u8(args.length + 1); o.u16(255); return; }
    if (f!.argc >= 0 && f!.argc === args.length) { o.u8(0x21 - 0x20 + CLS[fcls]); o.u16(f!.index); return; }
    if (f!.argc >= 0 && f!.argc !== args.length) throw new Error(`${up} takes ${f!.argc} arguments.`);
    o.u8(0x22 - 0x20 + CLS[fcls]); o.u8(args.length); o.u16(f!.index);
  };

  const emit = (n: Node, cls: Cls): void => {
    switch (n.type) {
      case 'num':
        if (Number.isInteger(n.value) && n.value >= 0 && n.value <= 65535) { o.u8(0x1e); o.u16(n.value); }
        else { o.u8(0x1f); o.f64(n.value); }
        return;
      case 'str': str(n.value); return;
      case 'bool': o.u8(0x1d); o.u8(n.value ? 1 : 0); return;
      case 'err': o.u8(0x1c); o.u8(CODE_BY_ERR[n.code] ?? 0x2a); return;
      case 'missing': o.u8(0x16); return;
      case 'array': array(n, cls); return;
      case 'ref': ref(n.ref, cls); return;
      case 'spill': emit(n.arg, cls); return;
      case 'name': {
        const idx = ctx.nameIndex(n.name, n.sheet !== undefined ? ctx.sheetIndex(n.sheet) : undefined);
        if (idx === undefined) {
          if (/^(TRUE|FALSE)$/i.test(n.name)) { o.u8(0x1d); o.u8(/^TRUE$/i.test(n.name) ? 1 : 0); return; }
          o.u8(0x1c); o.u8(0x1d); return;
        }
        o.u8(0x23 - 0x20 + CLS[cls]);
        if (wide) o.u32(idx); else { o.u16(idx); o.u16(0); }
        return;
      }
      case 'table': {
        const rf = ctx.table?.(n);
        if (rf) ref(rf, cls); else { o.u8(0x1c); o.u8(0x17); }
        return;
      }
      case 'call': call(n.name, n.args, cls); return;
      case 'invoke': o.u8(0x1c); o.u8(0x1d); return;
      case 'unary':
        if (n.op === '@') { emit(n.arg, cls); return; }
        emit(n.arg, cls === 'A' ? 'A' : 'V');
        if (n.op !== '%' && paren(n.arg, 6)) o.u8(0x15);
        o.u8(n.op === '-' ? 0x13 : n.op === '+' ? 0x12 : 0x14);
        return;
      case 'union':
        n.items.forEach((x, i) => { emit(x, 'R'); if (i) o.u8(0x10); });
        o.u8(0x15);
        return;
      case 'binary': {
        const p = PREC[n.op]!;
        const opCls: Cls = n.op === ':' || n.op === ' ' ? 'R' : cls === 'A' ? 'A' : 'V';
        emit(n.left, opCls);
        if (paren(n.left, p)) o.u8(0x15);
        emit(n.right, opCls);
        if (paren(n.right, p, true)) o.u8(0x15);
        o.u8(OPS[n.op]!);
        return;
      }
    }
  };
  emit(tree, top);
  return { rgce: new Uint8Array(o.bytes), rgcb: new Uint8Array(o.extra) };
};
