import { useEffect, useMemo, useRef, useState } from 'react';
import { addr, newWorkbook, cellKey, internStyle, type DataQuery, type Workbook } from '../../model/workbook';
import type { Scalar } from '../../engine/types';
import type { GridStore } from '../../app/store';
import { Dialog, Field, type DialogProps } from '../../ui/dialogs/Dialog';
import { Check, Radio } from '../../ui/dialogs/shared';
import { formatOpenHandlers, saveBlob, saveFile, workbookFromBytes } from '../../app/files';
import { displayText } from '../../app/ops/clipboard';
import { PasswordRequired } from '../../formats/biff8-read';
import { xlsLimits } from '../../formats/biff8-write';
import { xlsbLimits } from '../../formats/xlsb-write';
import {
  COLUMN_FORMATS, ENCODINGS, LOCALES, convertRows, decodeText, detectBreaks, guessOptions, parseText, splitLines, type ColumnFormat, type TextImportOptions,
} from '../../formats/text-import';
import { exportHtml, exportJson, exportXml, type CellSource } from '../../formats/html-export';
import { computedValues, newQueryId, placeBlock, rememberHandle, workbookBlock, type PickedFile } from './ops';
import { N, T, Tr } from '../../ui/i18n';

// ---------- Text import ----------

export interface TextImportData { file: PickedFile; mode: 'open' | 'insert' }

const DELIMS: [string, string][] = [['\t', N('Tab')], [';', N('Semicolon')], [',', N('Comma')], [' ', N('Space')], ['|', N('Bar')]];

/** Starts fresh for every file (the dialog host keeps one component per dialog id). */
export const TextImportDialog = (p: DialogProps<TextImportData>) => <TextImportBody key={`${p.data?.file.name}:${p.data?.file.data.length}`} {...p} />;

const TextImportBody = ({ store, close, data }: DialogProps<TextImportData>) => {
  const file = data!.file;
  const [o, setO] = useState<TextImportOptions>(() => guessOptions(file.data, file.name));
  const [other, setOther] = useState('');
  const [dest, setDest] = useState<'sheet' | 'cell'>('sheet');
  const [dateOrder, setDateOrder] = useState('');
  const set = (patch: Partial<TextImportOptions>) => setO((x) => ({ ...x, ...patch }));
  const opts: TextImportOptions = { ...o, dateOrder: dateOrder || undefined, delimited: { ...o.delimited, delimiters: [...o.delimited.delimiters.filter((d) => d.length === 1 && DELIMS.some(([k]) => k === d)), ...(other ? [other] : [])] } };
  const preview = useMemo(() => parseText(file.data, opts, 100).fields, [file.data, JSON.stringify(opts)]);
  const cols = Math.min(200, Math.max(0, ...preview.map((r) => r.length)));
  const lines = useMemo(() => splitLines(decodeText(file.data, o.encoding).text.slice(0, 200_000)).slice(Math.max(0, o.startRow - 1), Math.max(0, o.startRow - 1) + 40), [file.data, o.encoding, o.startRow]);
  const ruler = useRef<HTMLDivElement>(null);
  const width = Math.min(400, Math.max(20, ...lines.map((l) => l.length)));

  const toggleBreak = (pos: number) => {
    if (pos <= 0 || pos >= width + 50) return;
    set({ breaks: o.breaks.includes(pos) ? o.breaks.filter((b) => b !== pos) : [...o.breaks, pos].sort((a, b) => a - b) });
  };
  const colFormat = (i: number): ColumnFormat => o.formats[i] ?? 'general';
  const setFormat = (i: number, f: ColumnFormat) => { const next = [...o.formats]; while (next.length <= i) next.push('general'); next[i] = f; set({ formats: next }); };

  const ok = () => {
    const { fields } = parseText(file.data, opts);
    const { values, formats } = convertRows(fields, { formats: opts.formats, locale: opts.locale, dateOrder: opts.dateOrder, trim: opts.kind === 'fixed' });
    if (data!.mode === 'open') {
      const wb = newWorkbook();
      const sheet = wb.sheets[0]!;
      sheet.name = file.name.replace(/\.[^.]+$/, '').slice(0, 31).replace(/[\\/?*[\]:]/g, '_') || 'Sheet1';
      values.forEach((r, i) => r.forEach((v, j) => {
        if (v === null || i >= 1_048_576 || j >= 16_384) return;
        const f = formats[i]?.[j];
        sheet.cells.set(cellKey(i, j), f ? { v, s: internStyle(wb, { numFmt: f }) } : { v });
      }));
      store.load(wb, file.name);
      return true;
    }
    const at = store.sel.active;
    const q: DataQuery = {
      id: newQueryId(), name: file.name.replace(/\.[^.]+$/, '') || 'Text', kind: 'text', file: file.name, options: JSON.stringify(opts),
      target: dest === 'cell' ? { sheet: store.sheet.name, row: at.row, col: at.col, rows: 0, cols: 0 } : { sheet: '', row: 0, col: 0, rows: 0, cols: 0 },
    };
    q.name = uniqueQueryName(store, q.name);
    rememberHandle(q, file);
    placeBlock(store, q, { values, formats }, N('Get data from text'));
    return true;
  };

  return (
    <Dialog title={T('Import {name}', { name: file.name })} close={close} wide okLabel={data!.mode === 'open' ? T('Open') : T('Import')} onOk={ok}>
      <div className="b-col fmt-import">
        <div className="fmt-row">
          <Field label={T('File origin')}>
            <select value={o.encoding} onChange={(e) => set({ encoding: e.target.value })}>
              {ENCODINGS.filter(([k]) => k !== 'auto').map(([k, l]) => <option key={k} value={k}>{Tr(l)}</option>)}
            </select>
          </Field>
          <Field label={T('Start at row')}><input type="number" min={1} value={o.startRow} onChange={(e) => set({ startRow: Math.max(1, Math.floor(+e.target.value || 1)) })} /></Field>
          <Field label={T('Numbers like')}>
            <select value={LOCALES.findIndex(([, , l]) => l.decimal === o.locale.decimal && l.thousands === o.locale.thousands)} onChange={(e) => set({ locale: LOCALES[+e.target.value]![2] })}>
              {LOCALES.map(([k, l], i) => <option key={k} value={i}>{Tr(l)}</option>)}
            </select>
          </Field>
          <Field label={T('Dates in text')}>
            <select value={dateOrder} onChange={(e) => setDateOrder(e.target.value)}>
              <option value="">{T('Year first only')}</option><option value="MDY">{T('Month, day, year')}</option><option value="DMY">{T('Day, month, year')}</option><option value="YMD">{T('Year, month, day')}</option>
            </select>
          </Field>
        </div>
        <Radio name="kind" row value={o.kind} onChange={(k) => set({ kind: k, breaks: k === 'fixed' && !o.breaks.length ? detectBreaks(lines) : o.breaks })} options={[['delimited', N('Separated by characters')], ['fixed', N('Fixed width')]]} />
        {o.kind === 'delimited' ? (
          <div className="fmt-row">
            <div className="b-checks">
              {DELIMS.map(([d, l]) => (
                <Check key={l} label={Tr(l)} checked={o.delimited.delimiters.includes(d)} onChange={(on) => set({ delimited: { ...o.delimited, delimiters: on ? [...o.delimited.delimiters, d] : o.delimited.delimiters.filter((x) => x !== d) } })} />
              ))}
              <label className="b-check">{T('Other')}{' '}<input className="b-tiny" maxLength={3} value={other} onChange={(e) => setOther(e.target.value)} aria-label={T('Other separator')} /></label>
            </div>
            <Check label={T('Treat repeated separators as one')} checked={!!o.delimited.mergeDelimiters} onChange={(v) => set({ delimited: { ...o.delimited, mergeDelimiters: v } })} />
            <Field label={T('Text in quotes')}>
              <select value={o.delimited.qualifier} onChange={(e) => set({ delimited: { ...o.delimited, qualifier: e.target.value } })}>
                <option value={'"'}>{T('" double quotes')}</option><option value="'">{T('\' single quotes')}</option><option value="">{T('None')}</option>
              </select>
            </Field>
          </div>
        ) : (
          <div className="b-col">
            <p className="muted small">{T('Click the ruler or the text to add a column break; click a break to remove it.')}</p>
            <div className="b-inline"><button onClick={() => set({ breaks: detectBreaks(lines) })}>{T('Find breaks')}</button><button onClick={() => set({ breaks: [] })}>{T('Clear breaks')}</button></div>
            <div className="fmt-fixed" ref={ruler} onClick={(e) => {
              const box = ruler.current!.getBoundingClientRect();
              const ch = ruler.current!.querySelector<HTMLElement>('.fmt-ch')?.getBoundingClientRect().width || 8;
              toggleBreak(Math.round((e.clientX - box.left + ruler.current!.scrollLeft - 6) / ch));
            }}>
              <div className="fmt-ruler">{Array.from({ length: Math.ceil(width / 10) + 1 }, (_, i) => <span key={i} style={{ left: `calc(${i * 10}ch + 6px)` }}>{i * 10}</span>)}<span className="fmt-ch">0</span></div>
              <pre>{lines.join('\n')}</pre>
              {o.breaks.map((b) => <div key={b} className="fmt-break" style={{ left: `calc(${b}ch + 6px)` }} title={T('Break at {b}', { b })} />)}
            </div>
          </div>
        )}
        {data!.mode === 'insert' && (
          <Field label={T('Put the data in')}>
            <Radio name="dest" row value={dest} onChange={setDest} options={[['sheet', N('A new sheet')], ['cell', T('This sheet at {cell}', { cell: addr(store.sel.active.row, store.sel.active.col) })]]} />
          </Field>
        )}
        <div className="fmt-preview" role="region" aria-label={T('Preview')}>
          <table>
            <thead><tr>{Array.from({ length: cols }, (_, i) => (
              <th key={i}><select aria-label={T('Column {p0} type', { p0: i + 1 })} value={colFormat(i)} onChange={(e) => setFormat(i, e.target.value as ColumnFormat)}>
                {COLUMN_FORMATS.map(([k, l]) => <option key={k} value={k}>{Tr(l)}</option>)}
              </select></th>
            ))}</tr></thead>
            <tbody>{preview.slice(0, 60).map((r, i) => <tr key={i}>{Array.from({ length: cols }, (_, j) => <td key={j} className={colFormat(j) === 'skip' ? 'fmt-skip' : ''}>{r[j] ?? ''}</td>)}</tr>)}</tbody>
          </table>
        </div>
      </div>
    </Dialog>
  );
};

const uniqueQueryName = (store: GridStore, base: string): string => {
  const names = new Set((store.wb.queries ?? []).map((q) => q.name.toUpperCase()));
  let n = base, k = 2;
  while (names.has(n.toUpperCase())) n = `${base} ${k++}`;
  return n;
};

// ---------- From another workbook ----------

export interface WorkbookImportData { file: PickedFile; wb: Workbook }

export const WorkbookImportDialog = (p: DialogProps<WorkbookImportData>) => <WorkbookImportBody key={`${p.data?.file.name}:${p.data?.file.data.length}`} {...p} />;

const WorkbookImportBody = ({ store, close, data }: DialogProps<WorkbookImportData>) => {
  const { file, wb } = data!;
  const visible = wb.sheets.filter((s) => s.state !== 'veryHidden');
  const tables = wb.sheets.flatMap((s) => s.tables.map((t) => t.name));
  const names = wb.names.filter((n) => !n.hidden && !/^_xlnm\./.test(n.name)).map((n) => n.name);
  const [sheet, setSheet] = useState(visible[0]?.name ?? '');
  const [range, setRange] = useState('');
  const [dest, setDest] = useState<'sheet' | 'cell'>('sheet');
  const [values, setValues] = useState<((s: number, r: number, c: number) => Scalar) | null>(null);
  const [err, setErr] = useState('');
  useEffect(() => { let live = true; void computedValues(wb).then((f) => { if (live) setValues(() => f); }); return () => { live = false; }; }, [wb]);
  const block = values ? workbookBlock(wb, sheet, range.trim() || undefined, values) : null;
  const ok = () => {
    if (!block || typeof block === 'string') { setErr(typeof block === 'string' ? block : T('Still reading the file.')); return false; }
    const at = store.sel.active;
    const q: DataQuery = {
      id: newQueryId(), name: uniqueQueryName(store, range.trim() || sheet || file.name), kind: 'workbook', file: file.name,
      source: { sheet, ...(range.trim() ? { range: range.trim() } : {}) },
      target: dest === 'cell' ? { sheet: store.sheet.name, row: at.row, col: at.col, rows: 0, cols: 0 } : { sheet: '', row: 0, col: 0, rows: 0, cols: 0 },
    };
    rememberHandle(q, file);
    placeBlock(store, q, block, N('Get data from workbook'));
    return true;
  };
  return (
    <Dialog title={T('Get data from {name}', { name: file.name })} close={close} wide okLabel={T('Import')} onOk={ok}>
      <div className="b-split">
        <div className="b-col">
          <span className="b-h3">{T('Sheets')}</span>
          <div className="b-list b-short">{visible.map((s) => <button key={s.name} className={sheet === s.name && !range ? 'on' : ''} onClick={() => { setSheet(s.name); setRange(''); }}>{s.name}</button>)}</div>
          {(tables.length > 0 || names.length > 0) && <>
            <span className="b-h3">{T('Tables and names')}</span>
            <div className="b-list b-short">{[...tables, ...names].map((t) => <button key={t} className={range === t ? 'on' : ''} onClick={() => setRange(t)}>{t}</button>)}</div>
          </>}
        </div>
        <div className="b-col">
          <Field label={T('Range (optional)')}><input value={range} placeholder={T('All used cells, or for example A1:D20')} onChange={(e) => setRange(e.target.value)} /></Field>
          <Field label={T('Put the data in')}>
            <Radio name="wdest" row value={dest} onChange={setDest} options={[['sheet', N('A new sheet')], ['cell', T('This sheet at {cell}', { cell: addr(store.sel.active.row, store.sel.active.col) })]]} />
          </Field>
          {err && <p className="fmt-error" role="alert">{err}</p>}
          <div className="fmt-preview">
            {!block ? <p className="muted small">{T('Reading…')}</p> : typeof block === 'string' ? <p className="muted small">{block}</p> : (
              <table><tbody>{block.values.slice(0, 30).map((r, i) => <tr key={i}>{r.slice(0, 30).map((v, j) => <td key={j}>{v === null ? '' : String(v instanceof Object ? (v as { code: string }).code : v)}</td>)}</tr>)}</tbody></table>
            )}
          </div>
        </div>
      </div>
    </Dialog>
  );
};

// ---------- Password for an older workbook ----------

export interface PasswordData { name: string; data: Uint8Array; path: string | null; handle?: unknown; wrong: boolean }

export const XlsPasswordDialog = ({ store, close, data }: DialogProps<PasswordData>) => {
  const [pw, setPw] = useState('');
  const [wrong, setWrong] = useState(data?.wrong ?? false);
  const ok = () => {
    try {
      const wb = workbookFromBytes(data!.name, data!.data, pw);
      store.load(wb, data!.name, data!.path, data!.handle);
      store.file.password = pw;
      return true;
    } catch (e) {
      if (e instanceof PasswordRequired) { setWrong(true); return false; }
      store.notify(e instanceof Error ? e.message : T('The file could not be opened.'));
      return true;
    }
  };
  return (
    <Dialog title={T('Password')} close={close} okLabel={T('Open')} onOk={ok} width={380}>
      <p className="small">{data?.name}{' '}{T('is protected with a password.')}</p>
      <Field label={T('Password')}><input type="password" autoFocus value={pw} onChange={(e) => { setPw(e.target.value); setWrong(false); }} /></Field>
      {wrong && <p className="fmt-error" role="alert">{T('That password is not right.')}</p>}
    </Dialog>
  );
};

// ---------- Saving as an older workbook ----------

/** Save warnings from the file writers (they stay English there; shown translated here). */
export const SAVE_WARNINGS = [
  N('Rows after 65,536 are left out.'), N('Columns after IV are left out.'), N('Chart types the older format lacks are saved as the nearest simple type.'),
  N('Chart sheets are saved as worksheets with the chart on them.'), N('Tables become plain ranges.'), N('Pictures that are not PNG, JPEG, BMP, EMF or WMF are left out.'),
  N('Some conditional formats are left out.'), N('Cell styles after the first 4,000 use the default style.'), N('PivotTables and slicers are left out.'), N('Sparklines are left out.'),
  N('Tables, color scales, data bars and icon sets are kept for ClerdoGrid only; other apps show the cells without them.'), N('Pictures in headers and footers are left out.'), N('Scripts are left out.'),
];

export const XlsLimitsDialog = ({ store, close }: DialogProps) => {
  const limits = xlsLimits(store.wb);
  return (
    <Dialog title={T('Save as older workbook')} close={close} okLabel={T('Save')} onOk={() => { void saveFile(store, true, 'xls'); return true; }} width={440}>
      <p className="small">{T('The older .xls format can\'t hold everything in this workbook:')}</p>
      <ul className="small">{limits.map((l) => <li key={l}>{Tr(l)}</li>)}</ul>
      <p className="muted small">{T('Keep a copy as .xlsx to keep all of it.')}</p>
    </Dialog>
  );
};

export const XlsbLimitsDialog = ({ store, close }: DialogProps) => {
  const limits = xlsbLimits(store.wb);
  return (
    <Dialog title={T('Save as binary workbook')} close={close} okLabel={T('Save')} onOk={() => { void saveFile(store, true, 'xlsb'); return true; }} width={440}>
      <p className="small">{T('The .xlsb format can\'t hold everything in this workbook:')}</p>
      <ul className="small">{limits.map((l) => <li key={l}>{Tr(l)}</li>)}</ul>
      <p className="muted small">{T('Keep a copy as .xlsx to keep all of it.')}</p>
    </Dialog>
  );
};

// ---------- Export ----------

export interface ExportData { kind: 'html' | 'xml' | 'json' }

export const cellSource = (store: GridStore): CellSource => ({
  value: (s, r, c) => store.engine.getValue(s, r, c),
  text: (s, r, c) => displayText(store, r, c, s),
  style: (s, r, c) => store.style(r, c, s),
  bounds: (s) => store.engine.usedBounds(s),
});

export const ExportDialog = ({ store, close, data }: DialogProps<ExportData>) => {
  const kind = data?.kind ?? 'html';
  const r = store.range;
  const hasSel = r.r1 !== r.r2 || r.c1 !== r.c2;
  const [scope, setScope] = useState<'sheet' | 'book' | 'selection'>(hasSel ? 'selection' : 'sheet');
  const [headers, setHeaders] = useState(true);
  const [shown, setShown] = useState(false);
  const base = `${store.file.name || 'Book1'}${scope === 'book' ? '' : ` - ${store.sheet.name}`}`;
  const ok = () => {
    const src = cellSource(store);
    const area = scope === 'selection' ? r : undefined;
    let text: string, mime: string, ext: string, label: string;
    if (kind === 'html') {
      text = exportHtml(store.wb, src, { sheets: scope === 'book' ? undefined : [store.sheetIndex], area, title: store.file.name || N('Workbook') });
      mime = 'text/html;charset=utf-8'; ext = 'html'; label = T('Web page');
    } else if (kind === 'xml') {
      text = exportXml(store.wb, src, { sheet: store.sheetIndex, area, headers, asShown: shown });
      mime = 'application/xml'; ext = 'xml'; label = T('XML data');
    } else {
      text = exportJson(store.wb, src, { sheet: store.sheetIndex, area, headers, asShown: shown });
      mime = 'application/json'; ext = 'json'; label = T('JSON data');
    }
    void saveBlob(`${base}.${ext}`, new Blob([text], { type: mime }), label, ext);
    return true;
  };
  const title = kind === 'html' ? T('Export as web page') : kind === 'xml' ? T('Export as XML data') : T('Export as JSON data');
  return (
    <Dialog title={Tr(title)} close={close} okLabel={T('Export')} onOk={ok} width={420}>
      <div className="b-col">
        <Radio name="scope" value={scope} onChange={setScope} options={[
          ['sheet', N('This sheet')], ...(kind === 'html' ? [['book', N('All visible sheets')] as const] : []), ...(hasSel ? [['selection', N('Selected cells')] as const] : []),
        ]} />
        {kind !== 'html' && <>
          <Check label={T('First row has field names')} checked={headers} onChange={setHeaders} />
          <Check label={T('Write values as shown')} checked={shown} onChange={setShown} />
        </>}
      </div>
    </Dialog>
  );
};

/** Opens .txt and .prn files through the import dialog, and older workbooks that need a password. */
export const installOpenHandlers = (open: (store: GridStore, id: string, data: unknown) => void): void => {
  formatOpenHandlers.text = (store, file) => open(store, 'formats.textImport', { file: { name: file.name, data: file.data }, mode: 'open' } satisfies TextImportData);
  formatOpenHandlers.password = (store, file) => open(store, 'formats.password', file satisfies PasswordData);
};
