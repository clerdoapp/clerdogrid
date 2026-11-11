// Display order of rows for a sheet view that sorts without moving data: display row p shows data row
// toData(p). Only rows r1..r2 are reordered; every other row stays where it is.
import type { CellRange } from '../model/workbook';
import type { LayoutAxis, SheetLayout } from './layout';

export class RowOrder {
  readonly r2: number;
  private readonly inv: Int32Array;
  /** `perm[i]` is the data row shown at display row r1 + i. */
  constructor(readonly r1: number, readonly perm: Int32Array) {
    this.r2 = r1 + perm.length - 1;
    this.inv = new Int32Array(perm.length);
    perm.forEach((d, i) => { this.inv[d - r1] = r1 + i; });
  }
  toData(p: number): number { return p >= this.r1 && p <= this.r2 ? this.perm[p - this.r1]! : p; }
  toDisplay(d: number): number { return d >= this.r1 && d <= this.r2 ? this.inv[d - this.r1]! : d; }
  get identity(): boolean { return this.perm.every((d, i) => d === this.r1 + i); }

  /** Display rows r1..r2 as data ranges: one range per run of consecutive data rows. */
  dataRuns(p1: number, p2: number, c1: number, c2: number): CellRange[] {
    const out: CellRange[] = [];
    const push = (a: number, b: number) => {
      const last = out[out.length - 1];
      if (last && last.r2 + 1 === a) last.r2 = b; else out.push({ r1: a, c1, r2: b, c2 });
    };
    const lo = Math.min(p1, p2), hi = Math.max(p1, p2);
    if (lo < this.r1) push(lo, Math.min(hi, this.r1 - 1));
    for (let p = Math.max(lo, this.r1); p <= Math.min(hi, this.r2); p++) { const d = this.toData(p); push(d, d); }
    if (hi > this.r2) push(Math.max(lo, this.r2 + 1), hi);
    return out;
  }

  /** Display rows spanned by a data range (its bounding box on screen). */
  displaySpan(r1: number, r2: number): [number, number] {
    if (r2 < this.r1 || r1 > this.r2 || (r1 <= this.r1 && r2 >= this.r2)) return [r1, r2];
    let lo = Infinity, hi = -Infinity;
    const a = Math.max(r1, this.r1), b = Math.min(r2, this.r2);
    for (let d = a; d <= b; d++) { const p = this.toDisplay(d); if (p < lo) lo = p; if (p > hi) hi = p; }
    if (r1 < this.r1) lo = r1;
    if (r2 > this.r2) hi = r2;
    return [lo, hi];
  }
}

/** Row axis whose sizes follow the display order. */
class PermutedAxis implements LayoutAxis {
  private readonly prefix: Float64Array;
  private readonly base0: number;
  constructor(private readonly base: LayoutAxis, private readonly order: RowOrder) {
    this.base0 = base.start(order.r1);
    this.prefix = new Float64Array(order.perm.length + 1);
    order.perm.forEach((d, i) => { this.prefix[i + 1] = this.prefix[i]! + base.size(d); });
  }
  start(i: number): number {
    const { r1, r2 } = this.order;
    if (i <= r1 || i > r2 + 1) return this.base.start(i);
    return this.base0 + this.prefix[i - r1]!;
  }
  size(i: number): number { return this.start(i + 1) - this.start(i); }
  get total(): number { return this.base.total; }
  at(px: number): number {
    const { r1 } = this.order;
    const end = this.base0 + this.prefix[this.prefix.length - 1]!;
    if (px < this.base0 || px >= end) return this.base.at(px);
    let lo = 0, hi = this.prefix.length - 2;
    const rel = px - this.base0;
    while (lo < hi) { const m = (lo + hi + 1) >> 1; if (this.prefix[m]! <= rel) lo = m; else hi = m - 1; }
    return r1 + lo;
  }
}

const cache = new WeakMap<SheetLayout, { order: RowOrder; layout: SheetLayout }>();

/** The sheet layout with rows in display order. */
export const withRowOrder = (layout: SheetLayout, order: RowOrder | null): SheetLayout => {
  if (!order) return layout;
  const hit = cache.get(layout);
  if (hit && hit.order === order) return hit.layout;
  const out = { rows: new PermutedAxis(layout.rows, order), cols: layout.cols } as SheetLayout;
  cache.set(layout, { order, layout: out });
  return out;
};
