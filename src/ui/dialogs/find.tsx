import { useMemo, useRef, useState } from 'react';
import { findAll, replaceAll, currentRegion, type FindOptions, type Found } from '../../app/ops/data';
import { MAX_COLS, MAX_ROWS, isError } from '../../engine/types';
import { addr, keyCol, keyRow, type CellRange } from '../../model/workbook';
import { Dialog, Field, type DialogProps } from './Dialog';
import { Check, goTo, Radio, Tabs } from './shared';
import { N, T } from '../i18n';

// ---------- Find & replace ----------

export const FindDialog = ({ store, data, close }: DialogProps<{ tab?: 'find' | 'replace' }>) => {
  const [tab, setTab] = useState<'find' | 'replace'>(data?.tab ?? 'find');
  const [text, setText] = useState('');
  const [repl, setRepl] = useState('');
  const [within, setWithin] = useState<'sheet' | 'book'>('sheet');
  const [order, setOrder] = useState<'rows' | 'cols'>('rows');
  const [lookIn, setLookIn] = useState<'formulas' | 'values'>('formulas');
  const [matchCase, setMatchCase] = useState(false);
  const [whole, setWhole] = useState(false);
  const [results, setResults] = useState<Found[] | null>(null);
  const [more, setMore] = useState(false);
  const cursor = useRef(-1);

  const opts = (): FindOptions => ({ text, matchCase, wholeCell: whole, lookIn: tab === 'replace' ? 'formulas' : lookIn, allSheets: within === 'book', byColumns: order === 'cols' });
  const sortHits = (hits: Found[]): Found[] => {
    const cur = store.sheetIndex;
    return within === 'book' ? [...hits].sort((a, b) => ((a.sheet - cur + store.wb.sheets.length) % store.wb.sheets.length) - ((b.sheet - cur + store.wb.sheets.length) % store.wb.sheets.length) || 0) : hits;
  };
  const search = (): Found[] => {
    const hits = findAll(store, opts()).filter((h) => store.wb.sheets[h.sheet]!.state === 'visible');
    return sortHits(hits);
  };

  const next = (): Found | null => {
    if (!text) { store.notify(T('Type what to look for.')); return null; }
    const hits = search();
    if (!hits.length) { store.notify(T('Couldn\'t find "{text}".', { text })); return null; }
    const { row, col } = store.sel.active;
    const after = (h: Found) => h.sheet !== store.sheetIndex ? false : order === 'cols' ? h.col > col || (h.col === col && h.row > row) : h.row > row || (h.row === row && h.col > col);
    let i = hits.findIndex((h) => h.sheet === store.sheetIndex && after(h));
    if (i < 0) i = hits.findIndex((h) => h.sheet !== store.sheetIndex);
    if (i < 0) i = 0;
    const h = hits[i]!;
    cursor.current = i;
    goTo(store, h.sheet, h.row, h.col);
    return h;
  };

  const replaceOne = () => {
    if (!text) { store.notify(T('Type what to look for.')); return; }
    const { row, col } = store.sel.active;
    const here = findAll(store, { ...opts(), allSheets: false }).find((h) => h.row === row && h.col === col);
    if (here) {
      if (store.isLocked(row, col)) { store.notify(T('This cell is protected.')); return; }
      replaceAll(store, opts(), repl, here);
    }
    next();
  };

  const replaceEvery = () => {
    if (!text) { store.notify(T('Type what to look for.')); return; }
    const n = replaceAll(store, opts(), repl);
    store.notify(n ? T('Replaced {n, plural, one {# cell} other {# cells}}.', { n }) : T('Couldn\'t find "{text}".', { text }));
    setResults(null);
  };

  return (
    <Dialog title={T('Find and replace')} close={close} width={520} okLabel={T('Find next')} onOk={() => { next(); return false; }}
      footer={<>
        <button onClick={() => { if (!text) { store.notify(T('Type what to look for.')); return; } setResults(search()); }}>{T('Find all')}</button>
        {tab === 'replace' && <><button onClick={replaceEvery}>{T('Replace all')}</button><button onClick={replaceOne}>{T('Replace')}</button></>}
      </>}>
      <Tabs tabs={[['find', N('Find')], ['replace', N('Replace')]] as const} value={tab} onChange={setTab} />
      <Field label={T('Find what')}><input autoFocus value={text} onChange={(e) => { setText(e.target.value); setResults(null); }} placeholder={T('Use * and ? as wildcards')} /></Field>
      {tab === 'replace' && <Field label={T('Replace with')}><input value={repl} onChange={(e) => setRepl(e.target.value)} /></Field>}
      <button className="tool b-more" onClick={() => setMore(!more)} aria-expanded={more}>{more ? T('Fewer options') : T('More options')}</button>
      {more && <div className="b-col">
        <div className="grid3">
          <Field label={T('Within')}><select value={within} onChange={(e) => setWithin(e.target.value as 'sheet' | 'book')}><option value="sheet">{T('This sheet')}</option><option value="book">{T('Whole workbook')}</option></select></Field>
          <Field label={T('Search')}><select value={order} onChange={(e) => setOrder(e.target.value as 'rows' | 'cols')}><option value="rows">{T('By rows')}</option><option value="cols">{T('By columns')}</option></select></Field>
          <Field label={T('Look in')}><select value={tab === 'replace' ? 'formulas' : lookIn} disabled={tab === 'replace'} onChange={(e) => setLookIn(e.target.value as 'formulas' | 'values')}><option value="formulas">{T('Formulas')}</option><option value="values">{T('Values')}</option></select></Field>
        </div>
        <div className="b-checks"><Check label={T('Match case')} checked={matchCase} onChange={setMatchCase} /><Check label={T('Match entire cell')} checked={whole} onChange={setWhole} /></div>
      </div>}
      {results && <div className="b-results">
        <div className="muted small">{T('{count, plural, one {# cell found} other {# cells found}}', { count: results.length })}</div>
        <div className="b-table" role="table">
          {results.slice(0, 2000).map((h) => (
            <button key={`${h.sheet}:${h.row}:${h.col}`} className="b-tr b-link" onClick={() => goTo(store, h.sheet, h.row, h.col)}>
              <span>{store.wb.sheets[h.sheet]!.name}</span><span className="mono">{addr(h.row, h.col)}</span><span className="b-ellip">{h.text}</span>
            </button>
          ))}
        </div>
        {results.length > 2000 && <div className="muted small">{T('Showing the first 2,000.')}</div>}
      </div>}
    </Dialog>
  );
};

// ---------- Go to special ----------

type Special = 'blanks' | 'constants' | 'formulas' | 'region' | 'last' | 'visible' | 'cf' | 'dv' | 'rowDiff' | 'colDiff' | 'notes';
const SPECIALS: [Special, string][] = [
  ['notes', N('Notes')], ['constants', N('Constants')], ['formulas', N('Formulas')], ['blanks', N('Blanks')], ['region', N('Current region')], ['last', N('Last cell')],
  ['visible', N('Visible cells only')], ['cf', N('Conditional formats')], ['dv', N('Data validation')], ['rowDiff', N('Row differences')], ['colDiff', N('Column differences')],
];
const MAX_AREAS = 10000;

/** Packs cells into ranges: runs along each row, then merges equal runs on consecutive rows. */
export const packCells = (cells: { row: number; col: number }[]): CellRange[] => {
  const sorted = [...cells].sort((a, b) => a.row - b.row || a.col - b.col);
  const runs: CellRange[] = [];
  for (const c of sorted) {
    const last = runs.at(-1);
    if (last && last.r1 === c.row && last.c2 === c.col - 1) last.c2 = c.col;
    else runs.push({ r1: c.row, r2: c.row, c1: c.col, c2: c.col });
  }
  const open = new Map<string, CellRange>();
  const out: CellRange[] = [];
  for (const r of runs) {
    const k = `${r.c1}:${r.c2}`;
    const prev = open.get(k);
    if (prev && prev.r2 === r.r1 - 1) prev.r2 = r.r1;
    else { open.set(k, r); out.push(r); }
  }
  return out;
};

export const GotoSpecialDialog = ({ store, close }: DialogProps) => {
  const [kind, setKind] = useState<Special>('constants');
  const [types, setTypes] = useState({ numbers: true, text: true, logicals: true, errors: true });
  const [dvAll, setDvAll] = useState<'all' | 'same'>('all');
  const types4 = useMemo(() => Object.entries(types) as [keyof typeof types, boolean][], [types]);

  const pickType = (v: unknown): boolean => {
    if (typeof v === 'number') return types.numbers;
    if (typeof v === 'string') return types.text;
    if (typeof v === 'boolean') return types.logicals;
    if (isError(v)) return types.errors;
    return false;
  };

  const run = (): boolean => {
    const s = store.sheet;
    const si = store.sheetIndex;
    const b = store.engine.usedBounds(si);
    const sel = store.range;
    const single = store.sel.ranges.length === 1 && sel.r1 === sel.r2 && sel.c1 === sel.c2;
    const used: CellRange = { r1: 0, c1: 0, r2: Math.max(0, b.rows), c2: Math.max(0, b.cols) };
    const clip = (r: CellRange): CellRange => ({ r1: r.r1, c1: r.c1, r2: Math.min(r.r2, used.r2), c2: Math.min(r.c2, used.c2) });
    const areas: CellRange[] = single ? [used] : store.sel.ranges.map(clip);
    const inAreas = (row: number, col: number) => areas.some((a) => row >= a.r1 && row <= a.r2 && col >= a.c1 && col <= a.c2);
    const cells: { row: number; col: number }[] = [];
    const cap = MAX_AREAS * 50;

    if (kind === 'region') { const r = currentRegion(store); store.selectRange(r, store.sel.active); return true; }
    if (kind === 'last') { if (b.rows < 0) { store.notify(T('This sheet is empty.')); return false; } goTo(store, si, b.rows, b.cols); return true; }

    if (kind === 'constants' || kind === 'formulas') {
      for (const [k, c] of s.cells) {
        const row = keyRow(k), col = keyCol(k);
        if (!inAreas(row, col)) continue;
        const isF = !!c.f;
        if (kind === 'formulas' ? !isF : isF || c.v === null) continue;
        if (pickType(isF ? store.value(row, col) : c.v)) cells.push({ row, col });
      }
    } else if (kind === 'blanks') {
      for (const a of areas) for (let row = a.r1; row <= a.r2 && cells.length < cap; row++) for (let col = a.c1; col <= a.c2; col++) if (store.value(row, col) === null && !store.cell(row, col)?.f) cells.push({ row, col });
    } else if (kind === 'notes') {
      for (const k of s.notes.keys()) if (inAreas(keyRow(k), keyCol(k)) || single) cells.push({ row: keyRow(k), col: keyCol(k) });
    } else if (kind === 'visible') {
      const src = single ? [used] : store.sel.ranges.map(clip);
      for (const a of src) for (let row = a.r1; row <= a.r2 && cells.length < cap; row++) {
        if (s.rows.get(row)?.hidden) continue;
        for (let col = a.c1; col <= a.c2; col++) if (!s.cols.get(col)?.hidden) cells.push({ row, col });
      }
    } else if (kind === 'cf' || kind === 'dv') {
      const list = kind === 'cf' ? s.conditionalFormats.map((c) => c.ranges) : s.validations.filter((v) => {
        if (dvAll === 'all') return true;
        const { row, col } = store.sel.active;
        const here = s.validations.find((x) => x.ranges.some((r) => row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2));
        return here === v;
      }).map((v) => v.ranges);
      const ranges = list.flat().map((r) => ({ ...r, r2: Math.min(r.r2, MAX_ROWS - 1), c2: Math.min(r.c2, MAX_COLS - 1) }));
      if (!ranges.length) { store.notify(kind === 'cf' ? T('No cells have conditional formatting.') : T('No cells have data validation.')); return false; }
      store.sel = { ranges: ranges.slice(0, MAX_AREAS), active: { row: ranges[0]!.r1, col: ranges[0]!.c1 }, anchor: { row: ranges[0]!.r1, col: ranges[0]!.c1 } };
      store.emit();
      return true;
    } else if (kind === 'rowDiff' || kind === 'colDiff') {
      const a = sel;
      const { row: ar, col: ac } = store.sel.active;
      const key = (row: number, col: number) => { const c = store.cell(row, col); return c?.f ? `f:${c.f}` : `v:${String(store.value(row, col))}`; };
      if (kind === 'rowDiff') for (let row = a.r1; row <= Math.min(a.r2, used.r2); row++) { const ref = key(row, ac); for (let col = a.c1; col <= a.c2; col++) if (col !== ac && key(row, col) !== ref) cells.push({ row, col }); }
      else for (let col = a.c1; col <= Math.min(a.c2, used.c2); col++) { const ref = key(ar, col); for (let row = a.r1; row <= Math.min(a.r2, used.r2); row++) if (row !== ar && key(row, col) !== ref) cells.push({ row, col }); }
    }
    if (!cells.length) { store.notify(T('No cells were found.')); return false; }
    let ranges = packCells(cells.slice(0, cap));
    if (ranges.length > MAX_AREAS) { ranges = ranges.slice(0, MAX_AREAS); store.notify(T('Selected the first {p0} areas.', { p0: MAX_AREAS.toLocaleString() })); }
    const first = ranges[0]!;
    store.sel = { ranges, active: { row: first.r1, col: first.c1 }, anchor: { row: first.r1, col: first.c1 } };
    store.emit();
    return true;
  };

  return (
    <Dialog title={T('Go to special')} close={close} onOk={run} width={440}>
      <div className="b-twocol">
        <Radio name="gs" value={kind} options={SPECIALS} onChange={setKind} />
      </div>
      {(kind === 'constants' || kind === 'formulas') && <div className="b-checks">
        {types4.map(([k, v]) => <Check key={k} label={k[0]!.toUpperCase() + k.slice(1)} checked={v} onChange={(x) => setTypes({ ...types, [k]: x })} />)}
      </div>}
      {kind === 'dv' && <Radio name="dv" row value={dvAll} options={[['all', N('All')], ['same', N('Same as the active cell')]]} onChange={setDvAll} />}
      <p className="muted small">{T('With one cell selected, the whole used area is searched; otherwise only the selection.')}</p>
    </Dialog>
  );
};
