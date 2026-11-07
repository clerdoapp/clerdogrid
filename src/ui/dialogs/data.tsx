import { useMemo, useState } from 'react';
import { ArrowDown, ArrowUp, Plus, Trash2 } from 'lucide-react';
import { paste, type PasteMode, type PasteOptions } from '../../app/ops/clipboard';
import { fillSeries } from '../../app/ops/fill';
import { dataRange, goalSeek, looksLikeHeader, removeDuplicates, sortRange, textToColumns, type SortKey } from '../../app/ops/data';
import { parseNumber, toText } from '../../engine/coerce';
import { colName, parseAddr } from '../../model/workbook';
import type { GridStore } from '../../app/store';
import { Dialog, Field, type DialogProps } from './Dialog';
import { Check, columnLabel, parseRefs, PALETTE, Radio, rangeLabel, scalarText } from './shared';

// ---------- Paste special ----------

type PasteChoice = PasteMode;
const PASTE_MODES: [PasteChoice, string][] = [
  ['all', 'All'], ['formulas', 'Formulas'], ['values', 'Values'], ['formats', 'Formats'], ['noBorders', 'All except borders'],
  ['columnWidths', 'Column widths'], ['formulasAndFormats', 'Formulas and formats'], ['valuesAndFormats', 'Values and formats'],
];
const OPERATIONS: [NonNullable<PasteOptions['operation']>, string][] = [['none', 'None'], ['add', 'Add'], ['subtract', 'Subtract'], ['multiply', 'Multiply'], ['divide', 'Divide']];

export const PasteSpecialDialog = ({ store, close }: DialogProps) => {
  const [mode, setMode] = useState<PasteChoice>('all');
  const [op, setOp] = useState<NonNullable<PasteOptions['operation']>>('none');
  const [skip, setSkip] = useState(false);
  const [transpose, setTranspose] = useState(false);
  const go = (opts: PasteOptions) => { close(); void paste(store, opts).catch(() => store.notify("Couldn't read the clipboard.")); };
  return (
    <Dialog title="Paste special" close={close} width={460} okLabel="Paste" onOk={() => { go({ mode, operation: op, skipBlanks: skip, transpose }); return false; }}
      footer={<button onClick={() => go({ mode: 'link' })}>Paste link</button>}>
      <div className="grid2">
        <Field label="Paste"><Radio name="pm" value={mode} options={PASTE_MODES} onChange={setMode} /></Field>
        <Field label="Math with the cells below"><Radio name="op" value={op} options={OPERATIONS} onChange={setOp} /></Field>
      </div>
      <div className="b-checks"><Check label="Skip blanks" checked={skip} onChange={setSkip} /><Check label="Transpose" checked={transpose} onChange={setTranspose} /></div>
      <p className="muted small">Options apply to cells copied in ClerdoGrid. Text from other apps is pasted as is.</p>
    </Dialog>
  );
};

// ---------- Series ----------

export const SeriesDialog = ({ store, close }: DialogProps) => {
  const r = store.range;
  const [byRows, setByRows] = useState(r.c2 - r.c1 > r.r2 - r.r1);
  const [type, setType] = useState<'linear' | 'growth' | 'date'>('linear');
  const [unit, setUnit] = useState<'day' | 'weekday' | 'month' | 'year'>('day');
  const [step, setStep] = useState('1');
  const [stop, setStop] = useState('');
  return (
    <Dialog title="Fill series" close={close} width={420} onOk={() => {
      const s = Number(step);
      if (!Number.isFinite(s)) { store.notify('Enter a number for the step.'); return false; }
      const st = stop.trim() ? parseNumber(stop, store.wb.date1904) : undefined;
      if (st === null) { store.notify('Enter a number or date for the stop value, or leave it empty.'); return false; }
      if (r.r1 === r.r2 && r.c1 === r.c2) { store.notify('Select the start cell and the cells to fill.'); return false; }
      fillSeries(store, { byRows, type, unit: type === 'date' ? unit : undefined, step: s, stop: st });
      return true;
    }}>
      <div className="grid2">
        <Field label="Series in"><Radio name="sin" value={byRows ? 'rows' : 'cols'} options={[['rows', 'Rows'], ['cols', 'Columns']]} onChange={(v) => setByRows(v === 'rows')} /></Field>
        <Field label="Type"><Radio name="st" value={type} options={[['linear', 'Linear'], ['growth', 'Growth'], ['date', 'Date']]} onChange={setType} /></Field>
      </div>
      {type === 'date' && <Field label="Date unit"><select value={unit} onChange={(e) => setUnit(e.target.value as typeof unit)}><option value="day">Day</option><option value="weekday">Weekday</option><option value="month">Month</option><option value="year">Year</option></select></Field>}
      <div className="grid2">
        <Field label={type === 'growth' ? 'Multiply by' : 'Step'}><input value={step} inputMode="decimal" onChange={(e) => setStep(e.target.value)} /></Field>
        <Field label="Stop at (optional)"><input value={stop} onChange={(e) => setStop(e.target.value)} /></Field>
      </div>
      <p className="muted small">Starts from the first cell of each {byRows ? 'row' : 'column'} in the selection.</p>
    </Dialog>
  );
};

// ---------- Sort ----------

interface Level { col: number; by: 'value' | 'cellColor' | 'fontColor'; order: 'asc' | 'desc' | 'custom'; color: string; list: string }
const CUSTOM_LISTS = ['Sun, Mon, Tue, Wed, Thu, Fri, Sat', 'Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday', 'Jan, Feb, Mar, Apr, May, Jun, Jul, Aug, Sep, Oct, Nov, Dec', 'January, February, March, April, May, June, July, August, September, October, November, December', 'Low, Medium, High'];

export const SortDialog = ({ store, close }: DialogProps) => {
  const r = useMemo(() => dataRange(store), [store]);
  const [header, setHeader] = useState(() => looksLikeHeader(store, r));
  const [byCols, setByCols] = useState(false);
  const [caseSensitive, setCase] = useState(false);
  const [levels, setLevels] = useState<Level[]>([{ col: byCols ? r.r1 : store.sel.active.col >= r.c1 && store.sel.active.col <= r.c2 ? store.sel.active.col : r.c1, by: 'value', order: 'asc', color: '#ffc7ce', list: CUSTOM_LISTS[0]! }]);
  const lines = byCols ? Array.from({ length: Math.min(r.r2 - r.r1 + 1, 200) }, (_, i) => r.r1 + i) : Array.from({ length: Math.min(r.c2 - r.c1 + 1, 500) }, (_, i) => r.c1 + i);
  const label = (i: number) => byCols ? (header ? scalarText(store.value(i, r.c1)) || `Row ${i + 1}` : `Row ${i + 1}`) : columnLabel(store, r, i, header);
  const colors = useMemo(() => {
    const set = new Set<string>();
    for (let row = r.r1; row <= Math.min(r.r2, r.r1 + 2000); row++) for (let c = r.c1; c <= Math.min(r.c2, r.c1 + 50); c++) { const st = store.style(row, c); if (st.fill?.color) set.add(st.fill.color); if (st.font?.color) set.add(st.font.color); }
    return [...set, ...PALETTE].slice(0, 24);
  }, [store, r]);
  const set = (i: number, p: Partial<Level>) => setLevels(levels.map((l, k) => (k === i ? { ...l, ...p } : l)));
  const ok = () => {
    if (store.sel.ranges.length > 1) { store.notify("Sort can't be used on multiple selections."); return false; }
    const keys: SortKey[] = levels.map((l) => ({
      col: l.col, by: l.by, color: l.by === 'value' ? undefined : l.color, descending: l.order === 'desc',
      customList: l.by === 'value' && l.order === 'custom' ? l.list.split(',').map((x) => x.trim()).filter(Boolean) : undefined,
    }));
    sortRange(store, r, keys, header, byCols, caseSensitive);
    return true;
  };
  return (
    <Dialog title="Sort" close={close} onOk={ok} okLabel="Sort" wide
      footer={<button onClick={() => setLevels([...levels, { ...levels.at(-1)!, col: lines.find((c) => !levels.some((l) => l.col === c)) ?? lines[0]! }])} disabled={levels.length >= 64}><Plus size={15} strokeWidth={1.9} /> Add level</button>}>
      <div className="b-inline">
        <span className="muted small">Range {rangeLabel(r)}</span>
        <Check label="My data has headers" checked={header} onChange={setHeader} />
        <Check label="Case sensitive" checked={caseSensitive} onChange={setCase} />
        <Check label="Sort left to right" checked={byCols} onChange={(v) => { setByCols(v); setLevels([{ ...levels[0]!, col: v ? r.r1 : r.c1 }]); }} />
      </div>
      <div className="b-table" role="table">
        <div className="b-tr b-th b-sortrow" role="row"><span>{byCols ? 'Row' : 'Column'}</span><span>Sort on</span><span>Order</span><span /></div>
        {levels.map((l, i) => (
          <div key={i} className="b-tr b-sortrow" role="row">
            <select aria-label={byCols ? 'Row' : 'Column'} value={l.col} onChange={(e) => set(i, { col: +e.target.value })}>{lines.map((c) => <option key={c} value={c}>{label(c)}</option>)}</select>
            <select aria-label="Sort on" value={l.by} onChange={(e) => set(i, { by: e.target.value as Level['by'] })}><option value="value">Cell value</option><option value="cellColor">Cell color</option><option value="fontColor">Font color</option></select>
            {l.by === 'value'
              ? <span className="b-inline">
                  <select aria-label="Order" value={l.order} onChange={(e) => set(i, { order: e.target.value as Level['order'] })}><option value="asc">A to Z, smallest first</option><option value="desc">Z to A, largest first</option><option value="custom">Custom list</option></select>
                  {l.order === 'custom' && <input aria-label="Custom list" list="b-lists" value={l.list} onChange={(e) => set(i, { list: e.target.value })} />}
                </span>
              : <span className="b-inline">
                  <span className="b-colors">{colors.map((c) => <button key={c} type="button" className={`sw${l.color === c ? ' on' : ''}`} style={{ background: c }} aria-label={c} onClick={() => set(i, { color: c })} />)}</span>
                  <select aria-label="Position" value={l.order === 'desc' ? 'desc' : 'asc'} onChange={(e) => set(i, { order: e.target.value as Level['order'] })}><option value="asc">On top</option><option value="desc">On bottom</option></select>
                </span>}
            <span className="b-rowtools">
              <button className="ico sm" aria-label="Move up" disabled={i === 0} onClick={() => { const n = [...levels]; [n[i - 1], n[i]] = [n[i]!, n[i - 1]!]; setLevels(n); }}><ArrowUp size={15} strokeWidth={1.9} /></button>
              <button className="ico sm" aria-label="Move down" disabled={i === levels.length - 1} onClick={() => { const n = [...levels]; [n[i + 1], n[i]] = [n[i]!, n[i + 1]!]; setLevels(n); }}><ArrowDown size={15} strokeWidth={1.9} /></button>
              <button className="ico sm" aria-label="Delete level" disabled={levels.length === 1} onClick={() => setLevels(levels.filter((_, k) => k !== i))}><Trash2 size={15} strokeWidth={1.9} /></button>
            </span>
          </div>
        ))}
      </div>
      <datalist id="b-lists">{CUSTOM_LISTS.map((l) => <option key={l} value={l} />)}</datalist>
    </Dialog>
  );
};

// ---------- Remove duplicates ----------

export const RemoveDuplicatesDialog = ({ store, close }: DialogProps) => {
  const r = useMemo(() => dataRange(store), [store]);
  const [header, setHeader] = useState(() => looksLikeHeader(store, r));
  const cols = Array.from({ length: Math.min(r.c2 - r.c1 + 1, 500) }, (_, i) => r.c1 + i);
  const [picked, setPicked] = useState<Set<number>>(() => new Set(cols));
  return (
    <Dialog title="Remove duplicates" close={close} okLabel="Remove" onOk={() => {
      if (!picked.size) { store.notify('Pick at least one column.'); return false; }
      const total = r.r2 - r.r1 + 1 - (header ? 1 : 0);
      const n = removeDuplicates(store, r, [...picked].sort((a, b) => a - b), header);
      store.notify(n ? `Removed ${n} duplicate ${n === 1 ? 'row' : 'rows'}. ${total - n} unique left.` : 'No duplicates found.');
      return true;
    }}>
      <p className="muted small">Rows in {rangeLabel(r)} are duplicates when all the ticked columns match.</p>
      <div className="b-inline">
        <button className="tool" onClick={() => setPicked(new Set(cols))}>Select all</button>
        <button className="tool" onClick={() => setPicked(new Set())}>Clear</button>
        <Check label="My data has headers" checked={header} onChange={setHeader} />
      </div>
      <div className="b-list b-checklist">{cols.map((c) => <Check key={c} label={columnLabel(store, r, c, header)} checked={picked.has(c)} onChange={(v) => { const n = new Set(picked); if (v) n.add(c); else n.delete(c); setPicked(n); }} />)}</div>
    </Dialog>
  );
};

// ---------- Text to columns ----------

const splitPreview = (s: string, o: { delims: string[]; other: string; consecutive: boolean; qualifier: string; widths: number[]; fixed: boolean }): string[] => {
  if (o.fixed) {
    const out: string[] = [];
    let at = 0;
    for (const w of [...o.widths, Infinity]) { out.push(s.slice(at, at + w).trim()); at += w; if (at >= s.length) break; }
    return out;
  }
  const set = new Set([...o.delims, ...(o.other ? [o.other] : [])]);
  if (!set.size) return [s];
  const out: string[] = [];
  let cur = '', inQ = false;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i]!;
    if (o.qualifier && ch === o.qualifier) { inQ = !inQ; continue; }
    if (!inQ && set.has(ch)) { out.push(cur); cur = ''; if (o.consecutive) while (i + 1 < s.length && set.has(s[i + 1]!)) i++; continue; }
    cur += ch;
  }
  out.push(cur);
  return out;
};

export const TextToColumnsDialog = ({ store, close }: DialogProps) => {
  const r = store.range;
  const sample = useMemo(() => {
    const out: string[] = [];
    for (let row = r.r1; row <= Math.min(r.r2, r.r1 + 7); row++) { const v = store.value(row, r.c1); out.push(v === null ? '' : String(toText(v))); }
    return out;
  }, [store, r]);
  const [fixed, setFixed] = useState(false);
  const [delims, setDelims] = useState<string[]>(() => (sample.some((s) => s.includes('\t')) ? ['\t'] : sample.some((s) => s.includes(';')) ? [';'] : [',']));
  const [other, setOther] = useState('');
  const [consecutive, setConsecutive] = useState(false);
  const [qualifier, setQualifier] = useState('"');
  const [widthsText, setWidthsText] = useState('10');
  const widths = widthsText.split(',').map((x) => Math.trunc(Number(x.trim()))).filter((n) => n > 0);
  const opts = { delims, other, consecutive, qualifier, widths, fixed };
  const rows = sample.map((s) => splitPreview(s, opts));
  const width = Math.max(1, ...rows.map((x) => x.length));
  const toggle = (d: string, on: boolean) => setDelims(on ? [...delims, d] : delims.filter((x) => x !== d));
  const ok = () => {
    if (r.c1 !== r.c2) { store.notify('Select cells in one column.'); return false; }
    if (store.sel.ranges.length > 1) { store.notify("This can't be used on multiple selections."); return false; }
    if (fixed && !widths.length) { store.notify('Enter column widths like 10, 8, 12.'); return false; }
    textToColumns(store, fixed ? { delimiters: [], fixedWidths: widths } : { delimiters: delims, other: other || undefined, treatConsecutive: consecutive, qualifier });
    return true;
  };
  return (
    <Dialog title="Text to columns" close={close} onOk={ok} okLabel="Split" wide>
      <Radio name="ttc" row value={fixed ? 'fixed' : 'delim'} options={[['delim', 'Split at a character'], ['fixed', 'Fixed width']]} onChange={(v) => setFixed(v === 'fixed')} />
      {!fixed ? <>
        <div className="b-checks">
          <Check label="Tab" checked={delims.includes('\t')} onChange={(v) => toggle('\t', v)} />
          <Check label="Semicolon" checked={delims.includes(';')} onChange={(v) => toggle(';', v)} />
          <Check label="Comma" checked={delims.includes(',')} onChange={(v) => toggle(',', v)} />
          <Check label="Space" checked={delims.includes(' ')} onChange={(v) => toggle(' ', v)} />
          <label className="b-check">Other <input className="b-tiny" maxLength={1} value={other} onChange={(e) => setOther(e.target.value)} aria-label="Other character" /></label>
        </div>
        <div className="b-inline">
          <Check label="Treat repeated characters as one" checked={consecutive} onChange={setConsecutive} />
          <span className="muted small">Text in</span>
          <select value={qualifier} onChange={(e) => setQualifier(e.target.value)} aria-label="Text qualifier"><option value={'"'}>Double quotes</option><option value="'">Single quotes</option><option value="">None</option></select>
        </div>
      </> : <Field label="Column widths (characters, comma-separated)"><input value={widthsText} onChange={(e) => setWidthsText(e.target.value)} placeholder="10, 8, 12" /></Field>}
      <Field label="Preview">
        <div className="b-previewgrid" style={{ gridTemplateColumns: `repeat(${width}, minmax(60px, 1fr))` }}>
          {Array.from({ length: width }, (_, j) => <div key={`h${j}`} className="b-pgh">{colName(r.c1 + j)}</div>)}
          {rows.flatMap((parts, i) => Array.from({ length: width }, (_, j) => <div key={`${i}:${j}`} className="b-pgc">{parts[j] ?? ''}</div>))}
        </div>
      </Field>
      {width > 1 && <p className="muted small">Cells in the {width - 1} {width === 2 ? 'column' : 'columns'} to the right will be replaced.</p>}
    </Dialog>
  );
};

// ---------- Goal Seek ----------

const cellRef = (store: GridStore, text: string): { row: number; col: number } | null => {
  const p = parseRefs(store, text);
  if (p && p.sheet === store.sheetIndex && p.ranges.length === 1) { const g = p.ranges[0]!; if (g.r1 === g.r2 && g.c1 === g.c2) return { row: g.r1, col: g.c1 }; }
  return parseAddr(text);
};

export const GoalSeekDialog = ({ store, close }: DialogProps) => {
  const a = store.sel.active;
  const [setCell, setSetCell] = useState(rangeLabel({ r1: a.row, c1: a.col, r2: a.row, c2: a.col }));
  const [target, setTarget] = useState('');
  const [byCell, setByCell] = useState('');
  const [result, setResult] = useState<string | null>(null);
  const run = () => {
    const sc = cellRef(store, setCell), bc = cellRef(store, byCell);
    if (!sc) { store.notify('Enter the cell with the formula, like B5.'); return; }
    if (!store.cell(sc.row, sc.col)?.f) { store.notify('The "Set cell" needs a formula.'); return; }
    if (!bc) { store.notify('Enter the cell to change, like B2.'); return; }
    if (store.cell(bc.row, bc.col)?.f) { store.notify('The cell to change must hold a number, not a formula.'); return; }
    const t = parseNumber(target, store.wb.date1904);
    if (t === null) { store.notify('Enter the value you want to reach.'); return; }
    if (store.isLocked(bc.row, bc.col)) { store.notify('The cell to change is protected.'); return; }
    const res = goalSeek(store, sc, t, bc);
    setResult(res.ok ? `Found a solution: ${rangeLabel({ r1: bc.row, c1: bc.col, r2: bc.row, c2: bc.col })} = ${+res.value.toPrecision(10)}, and ${setCell.toUpperCase()} = ${scalarText(store.value(sc.row, sc.col))}. Undo to go back.` : `No exact solution found. Closest result: ${scalarText(res.result)}. Nothing was changed.`);
  };
  return (
    <Dialog title="Goal seek" close={close} width={400} footer={<button onClick={run}>Solve</button>}>
      <Field label="Set cell"><input value={setCell} onChange={(e) => setSetCell(e.target.value)} placeholder="B5" /></Field>
      <Field label="To value"><input value={target} inputMode="decimal" onChange={(e) => setTarget(e.target.value)} /></Field>
      <Field label="By changing cell"><input value={byCell} onChange={(e) => setByCell(e.target.value)} placeholder="B2" /></Field>
      {result && <div className="preview b-result" role="status">{result}</div>}
    </Dialog>
  );
};

