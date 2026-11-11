import { fromSerial, toSerial } from '../../engine/date';
import { shiftFormula } from '../../engine/refactor';
import { isDateFormat } from '../../formats/numfmt';
import { cellKey, internStyle, type CellRange } from '../../model/workbook';
import type { Scalar } from '../../engine/types';
import type { GridStore } from '../store';
import { N } from '../../ui/i18n';

const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const LISTS: string[][] = [DAYS, DAYS.map((d) => d.slice(0, 3)), MONTHS, MONTHS.map((m) => m.slice(0, 3))];

const matchCase = (sample: string, word: string) => (sample === sample.toUpperCase() ? word.toUpperCase() : sample === sample.toLowerCase() ? word.toLowerCase() : word);

interface Src { v: Scalar; f?: string; s?: number; row: number; col: number }

type Gen = (k: number) => { v?: Scalar; f?: string };

/** Builds the generator for one line (row or column) of source cells; k counts from the first new cell (0) outward. */
const seriesFor = (store: GridStore, src: Src[], step: 1 | -1, copyOnly: boolean): Gen => {
  const n = src.length;
  const at = (k: number) => src[((k % n) + n) % n]!;
  const copy: Gen = (k) => ({ v: at(k).v, f: at(k).f });
  // Formulas are shifted by the caller; plain copies just repeat the pattern.
  if (copyOnly || src.some((s) => s.f)) return step > 0 ? copy : (k) => ({ v: at(n - 1 - k).v });
  const nums = src.map((s) => s.v);
  if (nums.every((v) => typeof v === 'number')) {
    const xs = nums as number[];
    const fmt = store.wb.styles[src[0]!.s ?? 0]?.numFmt ?? 'General';
    if (n === 1) {
      if (isDateFormat(fmt)) return (k) => ({ v: xs[0]! + step * (k + 1) });
      return copy;
    }
    if (isDateFormat(fmt) && n >= 2) {
      const d = xs.map((x) => fromSerial(x, store.wb.date1904));
      const months = (d[1]!.year - d[0]!.year) * 12 + d[1]!.month - d[0]!.month;
      if (months && d.every((x) => x.day === d[0]!.day) && d.every((x, i) => i === 0 || (x.year - d[i - 1]!.year) * 12 + x.month - d[i - 1]!.month === months)) {
        const last = step > 0 ? d[n - 1]! : d[0]!;
        return (k) => ({ v: toSerial(last.year, last.month + step * months * (k + 1), last.day, store.wb.date1904) });
      }
    }
    // Least-squares line through the source, continued.
    const mx = (n - 1) / 2;
    const my = xs.reduce((a, b) => a + b, 0) / n;
    let sxy = 0, sxx = 0;
    xs.forEach((y, i) => { sxy += (i - mx) * (y - my); sxx += (i - mx) ** 2; });
    const slope = sxx ? sxy / sxx : 0;
    const icpt = my - slope * mx;
    return (k) => ({ v: +(icpt + slope * (step > 0 ? n + k : -1 - k)).toPrecision(15) });
  }
  if (n >= 1 && nums.every((v) => typeof v === 'string')) {
    const strs = nums as string[];
    for (const list of LISTS) {
      const idx = strs.map((s) => list.findIndex((w) => w.toLowerCase() === s.toLowerCase()));
      if (idx.every((i) => i >= 0)) {
        const d = n > 1 ? idx[1]! - idx[0]! : 1;
        const base = step > 0 ? idx[n - 1]! : idx[0]!;
        return (k) => ({ v: matchCase(strs[0]!, list[(((base + step * d * (k + 1)) % list.length) + list.length) % list.length]!) });
      }
    }
    const q = strs.map((s) => /^(Q|Qtr|Quarter )([1-4])$/i.exec(s));
    if (q.every(Boolean)) {
      const last = +(step > 0 ? q[n - 1]! : q[0]!)[2]!;
      return (k) => ({ v: `${q[0]![1]}${((((last - 1 + step * (k + 1)) % 4) + 4) % 4) + 1}` });
    }
    const parts = strs.map((s) => /^(.*?)(\d+)(\D*)$/.exec(s));
    if (parts.every(Boolean) && parts.every((p) => p![1] === parts[0]![1] && p![3] === parts[0]![3])) {
      const vals = parts.map((p) => +p![2]!);
      const d = n > 1 ? vals[1]! - vals[0]! : 1;
      const width = parts[0]![2]!.length;
      const last = step > 0 ? vals[n - 1]! : vals[0]!;
      return (k) => {
        const x = Math.abs(last + step * d * (k + 1));
        return { v: `${parts[0]![1]}${String(x).padStart(parts[0]![2]!.startsWith('0') ? width : 0, '0')}${parts[0]![3]}` };
      };
    }
  }
  return copy;
};

/**
 * Fill-handle drag: extends `source` to cover `target` (which contains it), continuing series.
 * `copyOnly` is the Ctrl-drag / "Copy cells" choice.
 */
export const autoFill = (store: GridStore, source: CellRange, target: CellRange, copyOnly = false): void => {
  const sheet = store.sheetIndex;
  const vertical = target.c1 === source.c1 && target.c2 === source.c2;
  const forward = vertical ? target.r2 > source.r2 : target.c2 > source.c2;
  const shrinking = vertical ? target.r2 < source.r2 : target.c2 < source.c2;
  store.change(N('AutoFill'), [{ area: target }, { area: source }], () => {
    if (shrinking) {
      // Dragging back into the selection clears what's left behind.
      const clear: CellRange = vertical ? { ...source, r1: target.r2 + 1 } : { ...source, c1: target.c2 + 1 };
      for (let r = clear.r1; r <= clear.r2; r++) for (let c = clear.c1; c <= clear.c2; c++) store.engine.setValue(sheet, r, c, null);
      return;
    }
    const lines = vertical ? source.c2 - source.c1 + 1 : source.r2 - source.r1 + 1;
    const len = vertical ? source.r2 - source.r1 + 1 : source.c2 - source.c1 + 1;
    const extra = vertical ? (forward ? target.r2 - source.r2 : source.r1 - target.r1) : (forward ? target.c2 - source.c2 : source.c1 - target.c1);
    for (let li = 0; li < lines; li++) {
      const src: Src[] = [];
      for (let i = 0; i < len; i++) {
        const row = vertical ? source.r1 + i : source.r1 + li, col = vertical ? source.c1 + li : source.c1 + i;
        const c = store.cell(row, col);
        src.push({ v: c?.v ?? null, f: c?.f, s: c?.s, row, col });
      }
      const gen = seriesFor(store, src, forward ? 1 : -1, copyOnly);
      for (let k = 0; k < extra; k++) {
        const dist = k + 1;
        const row = vertical ? (forward ? source.r2 + dist : source.r1 - dist) : source.r1 + li;
        const col = vertical ? source.c1 + li : forward ? source.c2 + dist : source.c1 - dist;
        const from = src[forward ? k % len : len - 1 - (k % len)]!;
        const out = gen(forward ? k : k);
        if (from.f) store.engine.setFormula(sheet, row, col, shiftFormula(from.f, row - from.row, col - from.col));
        else store.engine.setValue(sheet, row, col, out.v ?? null);
        const cell = store.sheet.cells.get(cellKey(row, col));
        if (from.s) { if (cell) cell.s = from.s; else store.sheet.cells.set(cellKey(row, col), { v: null, s: from.s }); }
        else if (cell) delete cell.s;
      }
    }
  });
  store.selectRange(target, store.sel.active);
};

/** Ctrl+D / Ctrl+R: copy the first row (or column) of the selection over the rest. A single cell takes from above/left. */
export const fillDirection = (store: GridStore, dir: 'down' | 'right' | 'up' | 'left'): void => {
  let r = store.range;
  if (r.r1 === r.r2 && (dir === 'down' || dir === 'up')) r = dir === 'down' ? { ...r, r1: Math.max(0, r.r1 - 1) } : { ...r, r2: r.r2 + 1 };
  if (r.c1 === r.c2 && (dir === 'right' || dir === 'left')) r = dir === 'right' ? { ...r, c1: Math.max(0, r.c1 - 1) } : { ...r, c2: r.c2 + 1 };
  const sheet = store.sheetIndex;
  store.change(`Fill ${dir}`, [{ area: r }], () => {
    for (let row = r.r1; row <= r.r2; row++) for (let col = r.c1; col <= r.c2; col++) {
      const sr = dir === 'down' ? r.r1 : dir === 'up' ? r.r2 : row;
      const sc = dir === 'right' ? r.c1 : dir === 'left' ? r.c2 : col;
      if (sr === row && sc === col) continue;
      const c = store.cell(sr, sc);
      if (c?.f) store.engine.setFormula(sheet, row, col, shiftFormula(c.f, row - sr, col - sc));
      else store.engine.setValue(sheet, row, col, c?.v ?? null);
      const t = store.sheet.cells.get(cellKey(row, col));
      if (c?.s) { if (t) t.s = c.s; else store.sheet.cells.set(cellKey(row, col), { v: null, s: c.s }); }
    }
  });
};

/** Home > Fill > Series: linear, growth, or date steps over the selection. */
export const fillSeries = (store: GridStore, opts: { byRows: boolean; type: 'linear' | 'growth' | 'date'; unit?: 'day' | 'weekday' | 'month' | 'year'; step: number; stop?: number }): void => {
  const r = store.range;
  const sheet = store.sheetIndex;
  store.change(N('Series'), [{ area: r }], () => {
    const lines = opts.byRows ? r.r2 - r.r1 + 1 : r.c2 - r.c1 + 1;
    const len = opts.byRows ? r.c2 - r.c1 + 1 : r.r2 - r.r1 + 1;
    for (let li = 0; li < lines; li++) {
      const start = store.value(opts.byRows ? r.r1 + li : r.r1, opts.byRows ? r.c1 : r.c1 + li);
      if (typeof start !== 'number') continue;
      let v = start;
      for (let i = 1; i < len; i++) {
        if (opts.type === 'linear') v += opts.step;
        else if (opts.type === 'growth') v *= opts.step;
        else {
          const d = fromSerial(v, store.wb.date1904);
          if (opts.unit === 'month') v = toSerial(d.year, d.month + opts.step, d.day, store.wb.date1904);
          else if (opts.unit === 'year') v = toSerial(d.year + opts.step, d.month, d.day, store.wb.date1904);
          else if (opts.unit === 'weekday') { let k = opts.step; while (k > 0) { v++; const wd = new Date(Date.UTC(fromSerial(v).year, fromSerial(v).month - 1, fromSerial(v).day)).getUTCDay(); if (wd !== 0 && wd !== 6) k--; } }
          else v += opts.step;
        }
        if (opts.stop !== undefined && (opts.step >= 0 ? v > opts.stop : v < opts.stop)) break;
        store.engine.setValue(sheet, opts.byRows ? r.r1 + li : r.r1 + i, opts.byRows ? r.c1 + i : r.c1 + li, +v.toPrecision(15));
      }
    }
  });
};

export { flashFill } from './flash';

export { internStyle };
