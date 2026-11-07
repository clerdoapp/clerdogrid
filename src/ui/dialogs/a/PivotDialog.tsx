import { useMemo, useState } from 'react';
import { currentRegion } from '../../../app/ops/data';
import type { GridStore } from '../../../app/store';
import { parseAddr, type CellRange } from '../../../model/workbook';
import { refString } from '../../charts/chartData';
import '../../charts/charts.css';
import { Dialog, Field, type DialogProps } from '../Dialog';
import { buildPivot, SUMMARIES, type Summary } from './pivot';
import { addSheetAtEnd, BOLD, clipToUsed, hasContent, headerNames, put, rangeText, readMatrix, resolveRange, setColWidth } from './util';

const defaultSource = (store: GridStore): CellRange => {
  const r = store.range;
  return clipToUsed(store, store.sheetIndex, r.r1 === r.r2 && r.c1 === r.c2 ? currentRegion(store) : r);
};

/** Parses "H2" or "Sheet2!H2" into a sheet index and cell. */
const parseDest = (store: GridStore, text: string): { sheet: number; row: number; col: number } | null => {
  const m = /^(?:'?(.+?)'?!)?(\$?[A-Za-z]{1,3}\$?\d{1,7})$/.exec(text.trim());
  if (!m) return null;
  const sheet = m[1] ? store.wb.sheets.findIndex((s) => s.name.toLowerCase() === m[1]!.replace(/''/g, "'").toLowerCase()) : store.sheetIndex;
  const a = parseAddr(m[2]!);
  return sheet >= 0 && a ? { sheet, ...a } : null;
};

export const PivotDialog = ({ store, close }: DialogProps) => {
  const [srcText, setSrcText] = useState(() => rangeText(store, defaultSource(store)));
  const src = resolveRange(store, srcText);
  const names = useMemo(() => (src ? headerNames(store, src.sheet, src.range, true) : []), [store, src?.sheet, src?.range.r1, src?.range.c1, src?.range.r2, src?.range.c2]); // eslint-disable-line react-hooks/exhaustive-deps
  const firstNumeric = useMemo(() => {
    if (!src) return 0;
    for (let c = src.range.c1; c <= src.range.c2; c++) if (typeof store.value(src.range.r1 + 1, c, src.sheet) === 'number') return c - src.range.c1;
    return Math.max(0, names.length - 1);
  }, [src, names, store]);
  const [rows, setRows] = useState<number[]>([0]);
  const [col, setCol] = useState<number | null>(null);
  const [value, setValue] = useState<number | null>(null);
  const [summary, setSummary] = useState<Summary>('sum');
  const [dest, setDest] = useState<'new' | 'cell'>('new');
  const [destText, setDestText] = useState('');
  const [asFormula, setAsFormula] = useState(false);
  const val = value ?? firstNumeric;
  const validRows = rows.filter((r) => r < names.length);

  const ok = (): boolean => {
    if (!src || src.range.r2 <= src.range.r1) { store.notify('Enter a source range with a header row and at least one row of data.'); return false; }
    if (!validRows.length) { store.notify('Pick at least one field for the rows.'); return false; }
    let target = { sheet: -1, row: 0, col: 0 };
    if (dest === 'cell') {
      const d = parseDest(store, destText);
      if (!d) { store.notify('Enter a destination cell such as H2.'); return false; }
      target = d;
    }
    const sr = src.range;
    const colRef = (c: number) => refString(src.name, { r1: sr.r1, c1: sr.c1 + c, r2: sr.r2, c2: sr.c1 + c });
    const fnName = SUMMARIES.find((s) => s.value === summary)!.fn;

    let cells: (string | number | boolean | null)[][] = [];
    let formula = '';
    if (asFormula) {
      const rowArg = validRows.length === 1 ? colRef(validRows[0]!) : `HSTACK(${validRows.map(colRef).join(',')})`;
      const agg = `LAMBDA(x,${fnName}(x))`;
      formula = col === null ? `GROUPBY(${rowArg},${colRef(val)},${agg},3)` : `PIVOTBY(${rowArg},${colRef(col)},${colRef(val)},${agg},3)`;
    } else {
      const data = readMatrix(store, src.sheet, sr, 200000).slice(1);
      const res = buildPivot(names, data, { rows: validRows, col, value: val, summary });
      cells = res.cells.map((r) => r.map((v) => (v !== null && typeof v === 'object' ? v.code : v)));
      if (cells.length * (cells[0]?.length ?? 0) > 2_000_000) { store.notify('That summary would be too large. Choose fewer fields.'); return false; }
    }
    const width = asFormula ? 1 : cells[0]!.length, height = asFormula ? 1 : cells.length;
    if (target.sheet >= 0 && hasContent(store, target.sheet, { r1: target.row, c1: target.col, r2: target.row + height - 1, c2: target.col + width - 1 })
      && !confirm('Replace the data in the cells where the summary goes?')) return false;

    const numFmt = summary === 'count' ? undefined : store.style(sr.r1 + 1, sr.c1 + val, src.sheet).numFmt;
    store.change(asFormula ? 'Summary formula' : 'PivotTable', 'book', () => {
      if (target.sheet < 0) { target = { sheet: addSheetAtEnd(store, 'Summary'), row: 0, col: 0 }; }
      else store.wb.activeSheet = target.sheet;
      const { sheet, row: r0, col: c0 } = target;
      if (asFormula) {
        put(store, sheet, r0, c0, null, undefined, formula);
        const c = store.wb.sheets[sheet]!.cells.get(r0 * 16384 + c0);
        if (c) c.dyn = true;
        return;
      }
      const last = cells.length - 1;
      cells.forEach((line, i) => line.forEach((v, j) => {
        const isValue = i > 0 && j >= validRows.length;
        const style = i === 0 || i === last ? { ...BOLD, ...(isValue && numFmt ? { numFmt } : {}) } : isValue && numFmt ? { numFmt } : undefined;
        put(store, sheet, r0 + i, c0 + j, v, style);
      }));
      if (target.row === 0 && target.col === 0) for (let j = 0; j < width; j++) setColWidth(store, sheet, c0 + j, Math.min(40, Math.max(10, ...cells.map((l) => String(l[j] ?? '').length + 2))));
    });
    store.selectRange({ r1: target.row, c1: target.col, r2: target.row, c2: target.col });
    return true;
  };

  const toggleRow = (i: number, on: boolean) => setRows(on ? [...rows.filter((r) => r !== i), i] : rows.filter((r) => r !== i));

  return (
    <Dialog title="PivotTable" close={close} width={480} onOk={ok} okLabel="Create">
      <Field label="Source data (with a header row)"><input value={srcText} onChange={(e) => setSrcText(e.target.value)} aria-invalid={!src} spellCheck={false} /></Field>
      <div className="seclabel">Rows</div>
      <div className="checks">
        {names.map((n, i) => <label key={i}><input type="checkbox" checked={rows.includes(i)} onChange={(e) => toggleRow(i, e.target.checked)} /> {n}{rows.length > 1 && rows.includes(i) ? <span className="muted small"> ({rows.indexOf(i) + 1})</span> : null}</label>)}
        {!names.length && <span className="muted">Enter a valid source range.</span>}
      </div>
      <div className="grid2">
        <Field label="Columns (optional)">
          <select value={col ?? ''} onChange={(e) => setCol(e.target.value === '' ? null : +e.target.value)}>
            <option value="">None</option>{names.map((n, i) => <option key={i} value={i}>{n}</option>)}
          </select>
        </Field>
        <Field label="Values">
          <select value={val} onChange={(e) => setValue(+e.target.value)}>{names.map((n, i) => <option key={i} value={i}>{n}</option>)}</select>
        </Field>
        <Field label="Summarize by">
          <select value={summary} onChange={(e) => setSummary(e.target.value as Summary)}>{SUMMARIES.map((s) => <option key={s.value} value={s.value}>{s.label}</option>)}</select>
        </Field>
        <Field label="Put it in">
          <select value={dest} onChange={(e) => setDest(e.target.value as 'new' | 'cell')}><option value="new">A new sheet</option><option value="cell">A cell I choose</option></select>
        </Field>
      </div>
      {dest === 'cell' && <Field label="Destination cell"><input value={destText} onChange={(e) => setDestText(e.target.value)} placeholder="H2 or Sheet2!A1" spellCheck={false} /></Field>}
      <label className="checkrow"><input type="checkbox" checked={asFormula} onChange={(e) => setAsFormula(e.target.checked)} /> Use a formula instead (updates when the data changes)</label>
    </Dialog>
  );
};
