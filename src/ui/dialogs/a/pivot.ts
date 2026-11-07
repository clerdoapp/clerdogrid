import { compareValues } from '../../../engine/coerce';
import { isError, type Scalar } from '../../../engine/types';

export type Summary = 'sum' | 'count' | 'average' | 'max' | 'min';
export const SUMMARIES: { value: Summary; label: string; fn: string }[] = [
  { value: 'sum', label: 'Sum', fn: 'SUM' }, { value: 'count', label: 'Count', fn: 'COUNTA' }, { value: 'average', label: 'Average', fn: 'AVERAGE' },
  { value: 'max', label: 'Max', fn: 'MAX' }, { value: 'min', label: 'Min', fn: 'MIN' },
];

export interface PivotSpec { rows: number[]; col: number | null; value: number; summary: Summary }

/** A static summary table: header row, one row per key, a grand total row and (with a column field) a total column. */
export interface PivotResult { cells: Scalar[][]; rowFieldCount: number }

class Acc {
  sum = 0; n = 0; count = 0; max = -Infinity; min = Infinity;
  add(v: Scalar): void {
    if (v === null || v === '') return;
    this.count++;
    if (typeof v === 'number') { this.sum += v; this.n++; this.max = Math.max(this.max, v); this.min = Math.min(this.min, v); }
  }
  result(s: Summary): Scalar {
    switch (s) {
      case 'sum': return this.sum;
      case 'count': return this.count;
      case 'average': return this.n ? this.sum / this.n : null;
      case 'max': return this.n ? this.max : null;
      case 'min': return this.n ? this.min : null;
    }
  }
}

const keyOf = (vals: Scalar[]): string => JSON.stringify(vals.map((v) => (isError(v) ? `#${v.code}` : typeof v === 'string' ? v.toLowerCase() : v)));
const cmpKeys = (a: Scalar[], b: Scalar[]): number => {
  for (let i = 0; i < a.length; i++) {
    const x = a[i] ?? null, y = b[i] ?? null;
    if (x === null || y === null) { if (x !== y) return x === null ? 1 : -1; continue; }
    if (isError(x) || isError(y)) continue;
    const c = compareValues(x, y);
    if (c) return c;
  }
  return 0;
};

/** `data` excludes the header row; `names` are the column headers. */
export const buildPivot = (names: string[], data: Scalar[][], spec: PivotSpec): PivotResult => {
  const label = `${SUMMARIES.find((s) => s.value === spec.summary)!.label} of ${names[spec.value] ?? 'Value'}`;
  const rowKeys = new Map<string, Scalar[]>(), colKeys = new Map<string, Scalar>();
  const cells = new Map<string, Acc>(), rowTot = new Map<string, Acc>(), colTot = new Map<string, Acc>();
  const grand = new Acc();
  const get = (m: Map<string, Acc>, k: string) => { let a = m.get(k); if (!a) { a = new Acc(); m.set(k, a); } return a; };
  for (const line of data) {
    if (line.every((v) => v === null)) continue;
    const rk = spec.rows.map((c) => line[c] ?? null);
    const rks = keyOf(rk);
    if (!rowKeys.has(rks)) rowKeys.set(rks, rk);
    const v = line[spec.value] ?? null;
    get(rowTot, rks).add(v);
    grand.add(v);
    if (spec.col !== null) {
      const ck = line[spec.col] ?? null;
      const cks = keyOf([ck]);
      if (!colKeys.has(cks)) colKeys.set(cks, ck);
      get(cells, `${rks}\u0001${cks}`).add(v);
      get(colTot, cks).add(v);
    }
  }
  const rows = [...rowKeys].sort((a, b) => cmpKeys(a[1], b[1]));
  const cols = [...colKeys].sort((a, b) => cmpKeys([a[1]], [b[1]]));
  const blankLabel = (v: Scalar): Scalar => (v === null ? '(blank)' : v);
  const out: Scalar[][] = [];
  const rowNames = spec.rows.map((c) => names[c] ?? '');
  if (spec.col === null) {
    out.push([...rowNames, label]);
    for (const [k, rk] of rows) out.push([...rk.map(blankLabel), rowTot.get(k)!.result(spec.summary)]);
    out.push(['Grand total', ...Array(spec.rows.length - 1).fill(null), grand.result(spec.summary)]);
  } else {
    out.push([...rowNames, ...cols.map(([, v]) => blankLabel(v)), 'Grand total']);
    for (const [k, rk] of rows) out.push([...rk.map(blankLabel), ...cols.map(([ck]) => cells.get(`${k}\u0001${ck}`)?.result(spec.summary) ?? null), rowTot.get(k)!.result(spec.summary)]);
    out.push(['Grand total', ...Array(spec.rows.length - 1).fill(null), ...cols.map(([ck]) => colTot.get(ck)!.result(spec.summary)), grand.result(spec.summary)]);
  }
  return { cells: out, rowFieldCount: spec.rows.length };
};
