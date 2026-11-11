import { useMemo, useState } from 'react';
import { changeStructure } from '../../../app/ops/structure';
import { currentRegion, looksLikeHeader } from '../../../app/ops/data';
import type { GridStore } from '../../../app/store';
import { shiftFormula } from '../../../engine/refactor';
import { isMatrix, type Scalar } from '../../../engine/types';
import { cellKey, parseAddr, rangeAddr, type Cell, type CellRange } from '../../../model/workbook';
import { refString } from '../../charts/chartData';
import '../../charts/charts.css';
import { Dialog, Field, type DialogProps } from '../Dialog';
import { addSheetAtEnd, BOLD, clipToUsed, headerNames, put, rangeText, resolveRange, setColWidth } from './util';
import { N, T, Tr } from '../../i18n';

const regionOrSelection = (store: GridStore): CellRange => {
  const r = store.range;
  return clipToUsed(store, store.sheetIndex, r.r1 === r.r2 && r.c1 === r.c2 ? currentRegion(store) : r);
};

const sameKey = (a: Scalar, b: Scalar): boolean => (typeof a === 'string' && typeof b === 'string' ? a.toLowerCase() === b.toLowerCase() : a === b);

// ---------- Subtotals ----------

const SUBTOTAL_FNS: { code: number; label: string }[] = [
  { code: 9, label: N('Sum') }, { code: 3, label: N('Count') }, { code: 1, label: N('Average') }, { code: 4, label: N('Max') }, { code: 5, label: N('Min') },
  { code: 6, label: N('Product') }, { code: 2, label: N('Count numbers') }, { code: 7, label: N('Standard deviation') }, { code: 10, label: N('Variance') },
];

/** Rows in the block that hold a SUBTOTAL formula. */
const subtotalRows = (store: GridStore, r: CellRange): number[] => {
  const out: number[] = [];
  for (let row = r.r1; row <= r.r2; row++) for (let c = r.c1; c <= r.c2; c++) if (store.engine.isSubtotalCell(store.sheetIndex, row, c)) { out.push(row); break; }
  return out;
};

const removeSubtotals = (store: GridStore, r: CellRange): number => {
  const rows = subtotalRows(store, r);
  for (const row of [...rows].reverse()) changeStructure(store, 'row', row, -1);
  const remaining = r.r2 - rows.length;
  for (let row = r.r1; row <= remaining; row++) {
    const info = store.sheet.rows.get(row);
    if (info) { delete info.outline; info.hidden = false; }
  }
  return rows.length;
};

export const SubtotalsDialog = ({ store, close }: DialogProps) => {
  const range = useMemo(() => regionOrSelection(store), [store]);
  const header = useMemo(() => looksLikeHeader(store, range), [store, range]);
  const names = useMemo(() => headerNames(store, store.sheetIndex, range, header), [store, range, header]);
  const numericCols = useMemo(() => names.map((_, i) => i).filter((i) => typeof store.value(range.r1 + (header ? 1 : 0), range.c1 + i) === 'number'), [names, store, range, header]);
  const [by, setBy] = useState(0);
  const [fn, setFn] = useState(9);
  const [cols, setCols] = useState<number[]>(() => (numericCols.length ? [numericCols.at(-1)!] : [names.length - 1]));
  const [replace, setReplace] = useState(true);
  const existing = useMemo(() => subtotalRows(store, range).length, [store, range]);

  const ok = (): boolean => {
    if (range.r2 - range.r1 < 1) { store.notify(T('Select a list with a header row and some data.')); return false; }
    if (!cols.length) { store.notify(T('Pick at least one column to add a subtotal to.')); return false; }
    const label = SUBTOTAL_FNS.find((f) => f.code === fn)!.label;
    store.change(N('Subtotal'), 'book', () => {
      let r = { ...range };
      if (replace && existing) r = { ...r, r2: r.r2 - removeSubtotals(store, r) };
      applySubtotals(store, r, header, by, fn, cols, label);
    });
    return true;
  };

  return (
    <Dialog title={T('Subtotal')} close={close} width={420} onOk={ok} footer={existing ? <button onClick={() => { store.change(N('Remove subtotals'), 'book', () => removeSubtotals(store, range)); close(); }}>{T('Remove all')}</button> : undefined}>
      <p className="muted small">{T('List: {range}. Sort it by the grouping column first.', { range: rangeAddr(range) })}</p>
      <Field label={T('At each change in')}>
        <select value={by} onChange={(e) => setBy(+e.target.value)}>{names.map((n, i) => <option key={i} value={i}>{n}</option>)}</select>
      </Field>
      <Field label={T('Use function')}>
        <select value={fn} onChange={(e) => setFn(+e.target.value)}>{SUBTOTAL_FNS.map((f) => <option key={f.code} value={f.code}>{Tr(f.label)}</option>)}</select>
      </Field>
      <div className="seclabel">{T('Add subtotal to')}</div>
      <div className="checks">
        {names.map((n, i) => <label key={i}><input type="checkbox" checked={cols.includes(i)} onChange={(e) => setCols(e.target.checked ? [...cols, i].sort((a, b) => a - b) : cols.filter((c) => c !== i))} /> {n}</label>)}
      </div>
      {existing > 0 && <label className="checkrow"><input type="checkbox" checked={replace} onChange={(e) => setReplace(e.target.checked)} />{' '}{T('Replace current subtotals')}</label>}
    </Dialog>
  );
};

/** Inserts a subtotal row after each group plus a grand total, and outlines the groups. Runs inside a 'book' change. */
const applySubtotals = (store: GridStore, r: CellRange, header: boolean, by: number, fn: number, cols: number[], label: string): void => {
  const sheetIdx = store.sheetIndex;
  const first = r.r1 + (header ? 1 : 0);
  if (first > r.r2) return;
  const groups: { from: number; to: number; key: Scalar }[] = [];
  for (let row = first; row <= r.r2; row++) {
    const k = store.value(row, r.c1 + by);
    const g = groups.at(-1);
    if (g && sameKey(g.key, k)) g.to = row; else groups.push({ from: row, to: row, key: k });
  }
  const extra = groups.length + 1;
  changeStructure(store, 'row', r.r2 + 1, extra);
  const sheet = store.sheet;
  // Snapshot the data rows (after the insert they are unchanged in place).
  const orig: (Cell | undefined)[][] = [];
  const origRows = new Map<number, ReturnType<typeof sheet.rows.get>>();
  for (let row = first; row <= r.r2; row++) {
    const line: (Cell | undefined)[] = [];
    for (let c = r.c1; c <= r.c2; c++) line.push(structuredClone(sheet.cells.get(cellKey(row, c))));
    orig.push(line);
    origRows.set(row, structuredClone(sheet.rows.get(row)));
  }
  for (let row = first; row <= r.r2 + extra; row++) {
    for (let c = r.c1; c <= r.c2; c++) store.engine.clearCell(sheetIdx, row, c);
    sheet.rows.delete(row);
  }
  const colAddr = (c: number, a: number, b: number) => `${colLetters(c)}${a + 1}:${colLetters(c)}${b + 1}`;
  let at = first;
  const placeRow = (src: number) => {
    orig[src - first]!.forEach((cell, j) => {
      if (!cell) return;
      const col = r.c1 + j;
      if (cell.f) { store.engine.setFormula(sheetIdx, at, col, shiftFormula(cell.f, at - src, 0)); const p = sheet.cells.get(cellKey(at, col)); if (p) p.s = cell.s; }
      else { store.engine.setValue(sheetIdx, at, col, cell.v); const p = sheet.cells.get(cellKey(at, col)); if (p) p.s = cell.s; else if (cell.s) sheet.cells.set(cellKey(at, col), { v: null, s: cell.s }); }
    });
    sheet.rows.set(at, { ...(origRows.get(src) ?? {}), outline: 2, hidden: false });
    at++;
  };
  const totalRow = (text: string, a: number, b: number, level: number | undefined) => {
    put(store, sheetIdx, at, r.c1 + by, text, BOLD);
    for (const j of cols) put(store, sheetIdx, at, r.c1 + j, null, BOLD, `SUBTOTAL(${fn},${colAddr(r.c1 + j, a, b)})`);
    sheet.rows.set(at, level ? { outline: level } : {});
    at++;
  };
  const dataStart = first;
  for (const g of groups) {
    const start = at;
    for (let row = g.from; row <= g.to; row++) placeRow(row);
    totalRow(`${g.key === null ? '(blank)' : String(typeof g.key === 'object' ? g.key.code : g.key)} ${label}`, start, at - 1, 1);
  }
  totalRow(`Grand ${label.toLowerCase()}`, dataStart, at - 1, undefined);
};

const colLetters = (c: number): string => rangeAddr({ r1: 0, c1: c, r2: 0, c2: c }).replace(/\d+/g, '');

// ---------- What-if data table ----------

const cellOf = (text: string): { row: number; col: number } | null => (text.trim() ? parseAddr(text.replace(/^=/, '')) : null);

export const DataTableDialog = ({ store, close }: DialogProps) => {
  const [rowIn, setRowIn] = useState('');
  const [colIn, setColIn] = useState('');
  const r = store.range;

  const ok = (): boolean => {
    const ri = cellOf(rowIn), ci = cellOf(colIn);
    if ((rowIn.trim() && !ri) || (colIn.trim() && !ci)) { store.notify(T('Enter input cells as cell addresses, such as B3.')); return false; }
    if (!ri && !ci) { store.notify(T('Enter a row input cell, a column input cell, or both.')); return false; }
    if (r.r2 - r.r1 < 1 || r.c2 - r.c1 < 1) { store.notify(T('Select the whole table first: input values along the edge and formulas in the first row or column.')); return false; }
    const body: CellRange = { r1: r.r1 + 1, c1: r.c1 + 1, r2: r.r2, c2: r.c2 };
    if ((body.r2 - body.r1 + 1) * (body.c2 - body.c1 + 1) > 20000) { store.notify(T('That table is too big. Use 20,000 result cells or fewer.')); return false; }
    for (const p of [ri, ci]) if (p && p.row >= r.r1 && p.row <= r.r2 && p.col >= r.c1 && p.col <= r.c2) { store.notify(T('The input cell must be outside the table.')); return false; }
    const results = computeTable(store, r, ri, ci);
    if (!results) return false;
    store.change(N('Data table'), [{ area: body }], () => {
      results.forEach((line, i) => line.forEach((v, j) => store.engine.setValue(store.sheetIndex, body.r1 + i, body.c1 + j, v)));
    });
    return true;
  };

  return (
    <Dialog title={T('Data table')} close={close} width={380} onOk={ok}>
      <p className="muted small">{T('Table: {range}. Put input values along the top row and/or left column, and the formula in the top-left cell (two inputs), the top row (column input) or the left column (row input).', { range: rangeAddr(r) })}</p>
      <Field label={T('Row input cell')}><input value={rowIn} onChange={(e) => setRowIn(e.target.value)} placeholder={T('Values in the top row replace this cell')} spellCheck={false} /></Field>
      <Field label={T('Column input cell')}><input value={colIn} onChange={(e) => setColIn(e.target.value)} placeholder={T('Values in the left column replace this cell')} spellCheck={false} /></Field>
    </Dialog>
  );
};

/** Fills a what-if table by setting the input cell(s), recalculating and reading the formula cell, then restoring. */
const computeTable = (store: GridStore, r: CellRange, ri: { row: number; col: number } | null, ci: { row: number; col: number } | null): Scalar[][] | null => {
  const sheet = store.sheetIndex;
  const inputs = [ri, ci].filter((p): p is { row: number; col: number } => !!p);
  const saved = inputs.map((p) => ({ p, cell: structuredClone(store.cell(p.row, p.col)) }));
  if (saved.some((s) => s.cell?.f)) { store.notify(T('An input cell must hold a value, not a formula.')); return null; }
  const formulaAt = (row: number, col: number) => !!store.cell(row, col)?.f;
  const set = (p: { row: number; col: number }, v: Scalar) => store.engine.setValue(sheet, p.row, p.col, v);
  const read = (row: number, col: number): Scalar => { const v = store.value(row, col); return isMatrix(v as never) ? null : v; };
  const out: Scalar[][] = [];
  try {
    if (ri && ci) {
      if (!formulaAt(r.r1, r.c1)) { store.notify(T('Put the formula in the top-left cell of the table.')); return null; }
      for (let row = r.r1 + 1; row <= r.r2; row++) {
        const line: Scalar[] = [];
        for (let col = r.c1 + 1; col <= r.c2; col++) { set(ri, store.value(r.r1, col)); set(ci, store.value(row, r.c1)); store.engine.recalc(); line.push(read(r.r1, r.c1)); }
        out.push(line);
      }
    } else if (ci) {
      const fcols = Array.from({ length: r.c2 - r.c1 }, (_, j) => r.c1 + 1 + j);
      if (!fcols.some((c) => formulaAt(r.r1, c))) { store.notify(T('Put the formulas in the top row of the table.')); return null; }
      for (let row = r.r1 + 1; row <= r.r2; row++) { set(ci, store.value(row, r.c1)); store.engine.recalc(); out.push(fcols.map((c) => (formulaAt(r.r1, c) ? read(r.r1, c) : null))); }
    } else if (ri) {
      const frows = Array.from({ length: r.r2 - r.r1 }, (_, i) => r.r1 + 1 + i);
      if (!frows.some((rr) => formulaAt(rr, r.c1))) { store.notify(T('Put the formulas in the left column of the table.')); return null; }
      const cols: Scalar[][] = [];
      for (let col = r.c1 + 1; col <= r.c2; col++) { set(ri, store.value(r.r1, col)); store.engine.recalc(); cols.push(frows.map((rr) => (formulaAt(rr, r.c1) ? read(rr, r.c1) : null))); }
      frows.forEach((_, i) => out.push(cols.map((c) => c[i]!)));
    }
  } finally {
    for (const s of saved) {
      if (s.cell) { set(s.p, s.cell.v); const c = store.cell(s.p.row, s.p.col); if (c) c.s = s.cell.s; }
      else store.engine.clearCell(sheet, s.p.row, s.p.col);
    }
    store.engine.recalc();
  }
  return out;
};

// ---------- Descriptive statistics ----------

const STATS: [string, (r: string) => string][] = [
  [N('Mean'), (r) => `AVERAGE(${r})`],
  [N('Standard error'), (r) => `STDEV.S(${r})/SQRT(COUNT(${r}))`],
  [N('Median'), (r) => `MEDIAN(${r})`],
  [N('Mode'), (r) => `IFERROR(MODE.SNGL(${r}),"None")`],
  [N('Standard deviation'), (r) => `STDEV.S(${r})`],
  [N('Variance'), (r) => `VAR.S(${r})`],
  [N('Kurtosis'), (r) => `KURT(${r})`],
  [N('Skewness'), (r) => `SKEW(${r})`],
  [N('Range'), (r) => `MAX(${r})-MIN(${r})`],
  [N('Minimum'), (r) => `MIN(${r})`],
  [N('Maximum'), (r) => `MAX(${r})`],
  [N('Sum'), (r) => `SUM(${r})`],
  [N('Count'), (r) => `COUNT(${r})`],
];

export const StatsDialog = ({ store, close }: DialogProps) => {
  const [srcText, setSrcText] = useState(() => rangeText(store, regionOrSelection(store)));
  const src = resolveRange(store, srcText);
  const header = useMemo(() => (src ? looksLikeHeader(store, src.range) || typeof store.value(src.range.r1, src.range.c1, src.sheet) === 'string' : false), [src, store]);
  const [useHeader, setUseHeader] = useState<boolean | null>(null);
  const hdr = useHeader ?? header;
  const names = src ? headerNames(store, src.sheet, src.range, hdr) : [];
  const numeric = src ? names.map((_, i) => i).filter((i) => {
    for (let row = src.range.r1 + (hdr ? 1 : 0); row <= Math.min(src.range.r2, src.range.r1 + 50); row++) if (typeof store.value(row, src.range.c1 + i, src.sheet) === 'number') return true;
    return false;
  }) : [];
  const [picked, setPicked] = useState<number[] | null>(null);
  const cols = (picked ?? numeric).filter((i) => i < names.length);

  const ok = (): boolean => {
    if (!src) { store.notify(T('Enter an input range such as B1:D20.')); return false; }
    if (!cols.length) { store.notify(T('Pick at least one column with numbers.')); return false; }
    const first = src.range.r1 + (hdr ? 1 : 0);
    if (first > src.range.r2) { store.notify(T('The input range has no data rows.')); return false; }
    store.change(N('Descriptive statistics'), 'book', () => {
      const sheet = addSheetAtEnd(store, 'Statistics');
      put(store, sheet, 0, 0, 'Statistic', BOLD);
      STATS.forEach(([label], i) => put(store, sheet, i + 1, 0, label));
      cols.forEach((ci, j) => {
        const ref = refString(src.name, { r1: first, c1: src.range.c1 + ci, r2: src.range.r2, c2: src.range.c1 + ci });
        put(store, sheet, 0, j + 1, names[ci]!, { ...BOLD, align: { h: 'right' } });
        STATS.forEach(([, f], i) => put(store, sheet, i + 1, j + 1, null, i === STATS.length - 1 ? undefined : { numFmt: '0.0000' }, f(ref)));
        setColWidth(store, sheet, j + 1, Math.max(12, names[ci]!.length + 2));
      });
      setColWidth(store, sheet, 0, 20);
      store.wb.sheets[sheet]!.view.frozenRows = 1;
    });
    store.selectRange({ r1: 0, c1: 0, r2: 0, c2: 0 });
    return true;
  };

  return (
    <Dialog title={T('Descriptive statistics')} close={close} width={420} onOk={ok}>
      <Field label={T('Input range')}><input value={srcText} onChange={(e) => { setSrcText(e.target.value); setPicked(null); }} aria-invalid={!src} spellCheck={false} /></Field>
      <label className="checkrow"><input type="checkbox" checked={hdr} onChange={(e) => setUseHeader(e.target.checked)} />{' '}{T('First row has labels')}</label>
      <div className="seclabel">{T('Columns')}</div>
      <div className="checks">
        {names.map((n, i) => <label key={i}><input type="checkbox" checked={cols.includes(i)} onChange={(e) => setPicked(e.target.checked ? [...cols, i].sort((a, b) => a - b) : cols.filter((c) => c !== i))} /> {n}</label>)}
        {!names.length && <span className="muted">{T('Enter a valid range.')}</span>}
      </div>
      <p className="muted small">{T('Results go on a new sheet as formulas, so they update when the data changes.')}</p>
    </Dialog>
  );
};

