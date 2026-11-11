import { useMemo, useState } from 'react';
import { currentRegion } from '../../../app/ops/data';
import type { GridStore } from '../../../app/store';
import { inRange } from '../../../app/store';
import { newUserSheet as newSheet, parseAddr, type CellRange } from '../../../model/workbook';
import { refString } from '../../charts/chartData';
import { Dialog, Field, type DialogProps } from '../Dialog';
import { Radio } from '../shared';
import { clipToUsed, headerNames, put, rangeText, resolveRange } from './util';
import { createPivot, sourceText } from '../../../features/pivot/apply';
import { newPivotTable } from '../../../features/pivot/model';
import { setPui } from '../../../features/pivot/ui';
import type { PivotSource } from '../../../model/pivot';
import { N, T, Tr } from '../../i18n';

const SUMMARY_FN: [string, string][] = [['SUM', N('Sum')], ['COUNTA', N('Count')], ['AVERAGE', N('Average')], ['MAX', N('Max')], ['MIN', N('Min')]];

const defaultSource = (store: GridStore): string => {
  const { row, col } = store.sel.active;
  const t = store.sheet.tables.find((x) => inRange(x.range, row, col));
  if (t && store.range.r1 === store.range.r2 && store.range.c1 === store.range.c2) return t.name;
  const r = store.range;
  return rangeText(store, clipToUsed(store, store.sheetIndex, r.r1 === r.r2 && r.c1 === r.c2 ? currentRegion(store) : r));
};

/** Parses "H2" or "Sheet2!H2" into a sheet index and cell. */
const parseDest = (store: GridStore, text: string): { sheet: number; row: number; col: number } | null => {
  const m = /^(?:'?(.+?)'?!)?(\$?[A-Za-z]{1,3}\$?\d{1,7})$/.exec(text.trim());
  if (!m) return null;
  const sheet = m[1] ? store.wb.sheets.findIndex((s) => s.name.toLowerCase() === m[1]!.replace(/''/g, "'").toLowerCase()) : store.sheetIndex;
  const a = parseAddr(m[2]!);
  return sheet >= 0 && a ? { sheet, ...a } : null;
};

/** Reads the source box: a range, a table name or a defined name. */
const sourceOf = (store: GridStore, text: string): { src: PivotSource; sheet: number; range: CellRange } | null => {
  const t = text.trim().replace(/^=/, '');
  const r = resolveRange(store, t);
  if (r) return { src: { kind: 'range', sheet: store.wb.sheets[r.sheet]!.name, range: r.range }, sheet: r.sheet, range: r.range };
  for (let i = 0; i < store.wb.sheets.length; i++) {
    const tb = store.wb.sheets[i]!.tables.find((x) => x.name.toLowerCase() === t.toLowerCase());
    if (tb) return { src: { kind: 'name', name: tb.name }, sheet: i, range: { ...tb.range, r2: tb.range.r2 - (tb.totalsRow ? 1 : 0) } };
  }
  const dn = store.wb.names.find((n) => n.name.toLowerCase() === t.toLowerCase());
  if (dn) {
    const rr = resolveRange(store, dn.formula.replace(/^=/, ''));
    if (rr) return { src: { kind: 'name', name: dn.name }, sheet: rr.sheet, range: rr.range };
  }
  return null;
};

export const PivotDialog = ({ store, close }: DialogProps) => {
  const [srcText, setSrcText] = useState(() => defaultSource(store));
  const src = sourceOf(store, srcText);
  const [dest, setDest] = useState<'new' | 'cell'>('new');
  const [destText, setDestText] = useState('');
  const sharing = (store.wb.pivotCaches ?? []).filter((c) => sourceText(c.source).toLowerCase() === (src ? sourceText(src.src) : '').toLowerCase());
  const [share, setShare] = useState(true);
  const [asFormula, setAsFormula] = useState(false);
  const names = useMemo(() => (src ? headerNames(store, src.sheet, src.range, true) : []), [store, src?.sheet, src?.range.r1, src?.range.c1, src?.range.r2, src?.range.c2]); // eslint-disable-line react-hooks/exhaustive-deps
  const [rows, setRows] = useState<number[]>([0]);
  const [col, setCol] = useState<number | null>(null);
  const [value, setValue] = useState<number>(Math.max(0, names.length - 1));
  const [fn, setFn] = useState('SUM');

  const ok = (): boolean => {
    if (!src || src.range.r2 <= src.range.r1) { store.notify(T('Enter a source range, table or name with a header row and at least one row of data.')); return false; }
    let target: { sheet: number; row: number; col: number } | null = null;
    if (dest === 'cell') {
      target = parseDest(store, destText);
      if (!target) { store.notify(T('Enter a destination cell such as H2.')); return false; }
    }
    if (asFormula) {
      const sr = src.range;
      const sheetName = store.wb.sheets[src.sheet]!.name;
      const colRef = (c: number) => refString(sheetName, { r1: sr.r1, c1: sr.c1 + c, r2: sr.r2, c2: sr.c1 + c });
      const valid = rows.filter((r) => r < names.length);
      if (!valid.length) { store.notify(T('Pick at least one field for the rows.')); return false; }
      const rowArg = valid.length === 1 ? colRef(valid[0]!) : `HSTACK(${valid.map(colRef).join(',')})`;
      const formula = col === null ? `GROUPBY(${rowArg},${colRef(value)},${fn},3)` : `PIVOTBY(${rowArg},${colRef(col)},${colRef(value)},${fn},3)`;
      const t = target ?? { sheet: -1, row: 0, col: 0 };
      store.change(N('Summary formula'), 'book', () => {
        let sheet = t.sheet;
        if (sheet < 0) { store.wb.sheets.splice(store.sheetIndex, 0, newSheet(freeSheetName(store))); store.engine.sheetsChanged(); sheet = store.sheetIndex; }
        store.wb.activeSheet = sheet;
        put(store, sheet, t.row, t.col, null, undefined, formula);
        const c = store.wb.sheets[sheet]!.cells.get(t.row * 16384 + t.col);
        if (c) c.dyn = true;
      });
      store.selectRange({ r1: t.row, c1: t.col, r2: t.row, c2: t.col });
      return true;
    }
    const res = createPivot(store, { source: src.src, dest: target, useCache: share && sharing[0] ? sharing[0].id : undefined }, newPivotTable);
    if (typeof res === 'string') { if (res) store.notify(res); return false; }
    setPui(store, { pane: true, paneTarget: res });
    return true;
  };

  const toggleRow = (i: number, on: boolean) => setRows(on ? [...rows.filter((r) => r !== i), i] : rows.filter((r) => r !== i));

  return (
    <Dialog title={T('PivotTable')} close={close} width={480} onOk={ok} okLabel={T('Create')}>
      <Field label={T('Data (a range with a header row, a table or a name)')}><input value={srcText} onChange={(e) => setSrcText(e.target.value)} aria-invalid={!src} spellCheck={false} /></Field>
      <div className="seclabel">{T('Put it in')}</div>
      <Radio name="pvdest" value={dest} onChange={setDest} options={[['new', N('A new sheet')], ['cell', N('This sheet or another one, at a cell')]]} />
      {dest === 'cell' && <Field label={T('Top-left cell')}><input value={destText} onChange={(e) => setDestText(e.target.value)} placeholder={T('H2 or Sheet2!A1')} spellCheck={false} /></Field>}
      {sharing.length > 0 && !asFormula && <label className="checkrow"><input type="checkbox" checked={share} onChange={(e) => setShare(e.target.checked)} />{' '}{T('Share data with the other PivotTables on this source (slicers can filter them together)')}</label>}
      <label className="checkrow"><input type="checkbox" checked={asFormula} onChange={(e) => setAsFormula(e.target.checked)} />{' '}{T('Make a summary formula instead (GROUPBY or PIVOTBY)')}</label>
      {asFormula && (
        <>
          <div className="seclabel">{T('Rows')}</div>
          <div className="checks">
            {names.map((n, i) => <label key={i}><input type="checkbox" checked={rows.includes(i)} onChange={(e) => toggleRow(i, e.target.checked)} /> {n}</label>)}
          </div>
          <div className="grid2">
            <Field label={T('Columns (optional)')}>
              <select value={col ?? ''} onChange={(e) => setCol(e.target.value === '' ? null : +e.target.value)}><option value="">{T('None')}</option>{names.map((n, i) => <option key={i} value={i}>{n}</option>)}</select>
            </Field>
            <Field label={T('Values')}><select value={value} onChange={(e) => setValue(+e.target.value)}>{names.map((n, i) => <option key={i} value={i}>{n}</option>)}</select></Field>
            <Field label={T('Summarize by')}><select value={fn} onChange={(e) => setFn(e.target.value)}>{SUMMARY_FN.map(([v, l]) => <option key={v} value={v}>{Tr(l)}</option>)}</select></Field>
          </div>
        </>
      )}
    </Dialog>
  );
};

const freeSheetName = (store: GridStore): string => {
  const taken = new Set(store.wb.sheets.map((s) => s.name.toLowerCase()));
  let n = store.wb.sheets.length + 1;
  while (taken.has(`sheet${n}`)) n++;
  return `Sheet${n}`;
};
