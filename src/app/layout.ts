import { MAX_COLS, MAX_ROWS } from '../engine/types';
import { keyRow, type CellStyle, type Sheet } from '../model/workbook';

/** Column width in characters → pixels, as spreadsheets do for an 11pt default font. */
export const colPx = (chars: number): number => Math.max(0, Math.trunc(chars * 7 + 5));
export const pxToChars = (px: number): number => Math.max(0, Math.round(((px - 5) / 7) * 100) / 100);
/** Row height in points → pixels. */
export const rowPx = (pt: number): number => Math.round((pt * 4) / 3);
export const pxToPt = (px: number): number => Math.round(px * 0.75 * 4) / 4;

/**
 * Pixel positions of rows and columns. Sizes are stored only up to the last customised index;
 * beyond that every size is the default, so a million rows cost nothing.
 */
class Axis {
  private offsets: Float64Array;
  private readonly last: number;
  constructor(private readonly count: number, private readonly dflt: number, sizes: Map<number, number>) {
    this.last = Math.min(count - 1, Math.max(-1, ...sizes.keys()));
    this.offsets = new Float64Array(this.last + 2);
    for (let i = 0; i <= this.last; i++) this.offsets[i + 1] = this.offsets[i]! + (sizes.get(i) ?? dflt);
  }
  start(i: number): number {
    if (i <= this.last + 1) return this.offsets[Math.max(0, i)]!;
    return this.offsets[this.last + 1]! + (i - this.last - 1) * this.dflt;
  }
  size(i: number): number { return this.start(i + 1) - this.start(i); }
  get total(): number { return this.start(this.count); }
  /** Index at a pixel position. */
  at(px: number): number {
    if (px <= 0) return 0;
    const edge = this.offsets[this.last + 1]!;
    if (px >= edge) return Math.min(this.count - 1, this.last + 1 + Math.floor((px - edge) / (this.dflt || 1)));
    let lo = 0, hi = this.last + 1;
    while (lo < hi) { const m = (lo + hi + 1) >> 1; if (this.offsets[m]! <= px) lo = m; else hi = m - 1; }
    return Math.min(lo, this.count - 1);
  }
}

/** Rows without a set height grow to fit larger fonts and line breaks in wrapped cells. */
const autoHeights = (s: Sheet, styles: CellStyle[]): Map<number, number> => {
  const out = new Map<number, number>();
  for (const [k, c] of s.cells) {
    if (!c.s) continue;
    const row = keyRow(k);
    if (s.rows.get(row)?.height !== undefined) continue;
    const st = styles[c.s];
    const size = st?.font?.size ?? 11;
    const lines = st?.align?.wrap && typeof c.v === 'string' ? c.v.split('\n').length : 1;
    if (size <= 11 && lines === 1) continue;
    const pt = Math.max(s.defaultRowHeight, Math.ceil(size * 1.3 * lines * 4) / 4);
    if (pt > (out.get(row) ?? 0)) out.set(row, pt);
  }
  return out;
};

export class SheetLayout {
  readonly rows: Axis;
  readonly cols: Axis;
  constructor(s: Sheet, styles: CellStyle[] = []) {
    const rowSizes = new Map<number, number>();
    for (const [i, pt] of autoHeights(s, styles)) rowSizes.set(i, rowPx(pt));
    for (const [i, r] of s.rows) if (r.hidden || r.height !== undefined) rowSizes.set(i, r.hidden ? 0 : rowPx(r.height!));
    const colSizes = new Map<number, number>();
    for (const [i, c] of s.cols) if (c.hidden || c.width !== undefined) colSizes.set(i, c.hidden ? 0 : colPx(c.width!));
    this.rows = new Axis(MAX_ROWS, rowPx(s.defaultRowHeight), rowSizes);
    this.cols = new Axis(MAX_COLS, colPx(s.defaultColWidth), colSizes);
  }
}

const cache = new WeakMap<Sheet, { version: number; layout: SheetLayout }>();

/** Cached per sheet and change version; the grid passes the workbook styles so rows can size to their fonts. */
export const layoutOf = (s: Sheet, version: number, styles?: CellStyle[]): SheetLayout => {
  const hit = cache.get(s);
  if (hit && hit.version === version) return hit.layout;
  const layout = new SheetLayout(s, styles);
  cache.set(s, { version, layout });
  return layout;
};
