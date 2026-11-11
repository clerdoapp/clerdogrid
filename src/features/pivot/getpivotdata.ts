// GETPIVOTDATA against real PivotTables, and the formula text made when a table cell is clicked
// while typing a formula.
import { setPivotResolver } from '../../engine/pivotHook';
import { quoteSheet } from '../../engine/parser';
import { fromSerial } from '../../engine/date';
import { Err, isError, type Scalar } from '../../engine/types';
import { colName, type Workbook } from '../../model/workbook';
import { VALUES_FIELD, type PivotCache, type PivotTable } from '../../model/pivot';
import { cacheOf, layoutOf, pivotAt, pivotCellAt } from './apply';
import { fieldItems, itemKey } from './cache';
import { captionOf, dataFieldName, fieldCaption, type Line, type PivotLayout, type PNodeT } from './compute';

const norm = (s: string) => s.trim().toLowerCase();

const findDataField = (pt: PivotTable, cache: PivotCache, name: string): number => {
  const n = norm(name);
  let i = pt.data.findIndex((d) => norm(dataFieldName(cache, d)) === n);
  if (i < 0) i = pt.data.findIndex((d) => norm(cache.fields[d.field]?.name ?? '') === n);
  return i;
};

const findField = (pt: PivotTable, cache: PivotCache, name: string): number => {
  const n = norm(name);
  return cache.fields.findIndex((f, i) => norm(f.name) === n || norm(fieldCaption(pt, cache, i)) === n);
};

const findSlot = (pt: PivotTable, cache: PivotCache, l: PivotLayout, f: number, item: Scalar, date1904: boolean): number => {
  const pf = pt.fields[f];
  if (!pf) return -1;
  const all = fieldItems(cache.fields[f]!);
  const want = typeof item === 'string' ? norm(item) : typeof item === 'number' || typeof item === 'boolean' ? norm(String(item)) : null;
  for (let s = 0; s < pf.items.length; s++) {
    const it = pf.items[s]!;
    if (want !== null && norm(l.caption(f, s)) === want) return s;
    if (it.x >= 0) {
      const v = all[it.x] ?? null;
      if (itemKey(v) === itemKey(item)) return s;
      if (want !== null && norm(captionOf(v, undefined, cache.fields[f]!.dates, date1904)) === want) return s;
    }
  }
  return -1;
};

const childBySlot = (n: PNodeT, f: number, slot: number): PNodeT | null => n.kids.find((k) => k.field === f && k.slot === slot) ?? null;

/** Walks an axis tree along the given items; fields must be given from the outermost in. */
const nodeFor = (root: PNodeT, axis: number[], picks: Map<number, number>): PNodeT | null => {
  let n = root;
  let stopped = false;
  for (const f of axis) {
    if (f === VALUES_FIELD) continue;
    const s = picks.get(f);
    if (s === undefined) { stopped = true; continue; }
    if (stopped) return null;
    const k = childBySlot(n, f, s);
    if (!k) return null;
    n = k;
  }
  return n;
};

const lineFor = (lines: Line[], node: PNodeT, d: number, hasData: boolean): number =>
  lines.findIndex((l) => l.node === node && (!hasData || l.data === d) && ((l.kind === 'item' && l.showValues) || (l.kind === 'sub' && l.fn === 'default') || l.kind === 'grand'));

export const getPivotData = (wb: Workbook, sheet: number, row: number, col: number, dataName: string, pairs: [string, Scalar][]): Scalar => {
  const ref = pivotAt(wb, sheet, row, col);
  if (!ref) return Err.ref;
  const pt = ref.pt;
  const cache = cacheOf(wb, pt.cacheId);
  const l = layoutOf(wb, pt);
  if (!cache || !l) return Err.ref;
  const d = findDataField(pt, cache, dataName);
  if (d < 0) return Err.ref;
  const picks = new Map<number, number>();
  for (const [fname, item] of pairs) {
    const f = findField(pt, cache, fname);
    if (f < 0) return Err.ref;
    const s = findSlot(pt, cache, l, f, item, wb.date1904);
    if (s < 0) return Err.ref;
    const page = pt.pages.find((p) => p.field === f);
    if (page) { if (page.item !== undefined && page.item !== s) return Err.ref; continue; }
    if (!pt.rows.includes(f) && !pt.cols.includes(f)) return Err.ref;
    picks.set(f, s);
  }
  const rn = nodeFor(l.rowRoot, pt.rows, picks), cn = nodeFor(l.colRoot, pt.cols, picks);
  if (!rn || !cn) return Err.ref;
  const multi = pt.data.length > 1;
  const ri = lineFor(l.rowLines, rn, d, multi && pt.rows.includes(VALUES_FIELD));
  const ci = lineFor(l.colLines, cn, d, multi && !pt.rows.includes(VALUES_FIELD));
  if (ri < 0 || ci < 0) return Err.ref;
  const cell = l.cells[l.firstDataRow + ri]?.[l.firstDataCol + ci];
  if (!cell || cell.role !== 'data') return Err.ref;
  return cell.v === null ? Err.ref : cell.v;
};

setPivotResolver(getPivotData);

const itemArg = (v: Scalar, dates: boolean | undefined, date1904: boolean): string => {
  if (typeof v === 'number') {
    if (dates) { const p = fromSerial(Math.floor(v), date1904); return `DATE(${p.year},${p.month},${p.day})`; }
    return String(v);
  }
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (isError(v)) return `"${v.code}"`;
  return `"${String(v ?? '(blank)').replace(/"/g, '""')}"`;
};

/** GETPIVOTDATA(...) text for a value cell, or null when the cell isn't a table value. */
export const getPivotDataText = (wb: Workbook, sheet: number, row: number, col: number, fromSheet: number): string | null => {
  const hit = pivotCellAt(wb, sheet, row, col);
  if (!hit?.cell || hit.cell.role !== 'data' || hit.cell.data === undefined) return null;
  const { pt } = hit.ref;
  const cache = cacheOf(wb, pt.cacheId);
  if (!cache) return null;
  const rl = hit.layout.rowLines[hit.cell.rowLine!], cl = hit.layout.colLines[hit.cell.colLine!];
  if (!rl || !cl) return null;
  if ((rl.kind === 'sub' && rl.fn !== 'default') || (cl.kind === 'sub' && cl.fn !== 'default')) return null;
  const args: string[] = [];
  for (const el of [...rl.path, ...cl.path]) {
    if (el.field === VALUES_FIELD) continue;
    const it = pt.fields[el.field]?.items[el.slot];
    if (!it) continue;
    const v = it.name ?? (it.x >= 0 ? fieldItems(cache.fields[el.field]!)[it.x] ?? null : null);
    args.push(`"${cache.fields[el.field]!.name.replace(/"/g, '""')}"`, itemArg(v, !cache.fields[el.field]!.group && cache.fields[el.field]!.dates, wb.date1904));
  }
  const at = `$${colName(pt.anchor.col)}$${pt.anchor.row + 1}`;
  const target = sheet !== fromSheet ? `${quoteSheet(wb.sheets[sheet]!.name)}!${at}` : at;
  const name = dataFieldName(cache, pt.data[hit.cell.data]!).replace(/"/g, '""');
  return `GETPIVOTDATA("${name}",${target}${args.length ? ',' + args.join(',') : ''})`;
};
