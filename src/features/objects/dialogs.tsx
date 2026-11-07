import { useEffect, useMemo, useState } from 'react';
import { addr, rangeAddr, parseRange } from '../../model/workbook';
import type { GridStore } from '../../app/store';
import { Dialog, Field, type DialogProps } from '../../ui/dialogs/Dialog';
import { accessibilityIssues, cleanExcessFormatting, excessFormatting, sheetStats, spellCells, wordsIn, type SpellCell } from './checks';
import { addEditRange, removeEditRange, unlockEditRange, updateShape } from './ops';

const goTo = (store: GridStore, sheet: number, at?: { row: number; col: number }) => {
  if (sheet !== store.sheetIndex) store.setActiveSheet(sheet);
  if (at) store.select({ row: at.row, col: at.col });
};

// ---------- Workbook statistics ----------

export const StatsDialog = ({ store, close }: DialogProps) => {
  const all = store.wb.sheets.map((s) => sheetStats(store.wb, s));
  const cur = all[store.sheetIndex]!;
  const sum = (k: keyof typeof cur) => all.reduce((n, s) => n + (s[k] as number), 0);
  const rows: [string, keyof typeof cur][] = [['Cells with data', 'cells'], ['Formulas', 'formulas'], ['Tables', 'tables'], ['Charts', 'charts'], ['Pictures', 'pictures'], ['Shapes', 'shapes'], ['Comments', 'comments'], ['Notes', 'notes']];
  return (
    <Dialog title="Workbook statistics" close={close} width={420}>
      <table className="obj-stats">
        <thead><tr><th /><th>This sheet</th><th>Workbook</th></tr></thead>
        <tbody>
          <tr><td>End of sheet</td><td>{cur.endCell}</td><td>{all.length} {all.length === 1 ? 'sheet' : 'sheets'}</td></tr>
          {rows.map(([l, k]) => <tr key={k}><td>{l}</td><td>{(cur[k] as number).toLocaleString()}</td><td>{sum(k).toLocaleString()}</td></tr>)}
        </tbody>
      </table>
    </Dialog>
  );
};

// ---------- Accessibility ----------

export const AccessibilityDialog = ({ store, close }: DialogProps) => {
  const [v, setV] = useState(0);
  const issues = useMemo(() => accessibilityIssues(store.wb), [store, v, store.layoutVersion]);
  const groups = [...new Set(issues.map((i) => i.kind))];
  return (
    <Dialog title="Check accessibility" close={close} width={500}>
      {!issues.length && <p>No problems found.</p>}
      {groups.map((g) => (
        <section key={g} className="obj-issues">
          <h3 className="b-h3">{g} ({issues.filter((i) => i.kind === g).length})</h3>
          {issues.filter((i) => i.kind === g).map((i, n) => (
            <div key={n} className="obj-issue">
              <button className="obj-go" onClick={() => goTo(store, i.sheet, i.go)}><b>{store.wb.sheets[i.sheet]!.name}</b> · {i.where}<span className="muted small">{i.text} {i.tip}</span></button>
              {i.object && <button onClick={() => { goTo(store, i.sheet); store.dialog = { id: 'altText', data: { ...i.object, back: 'accessibility' } }; store.emit(); setV(v + 1); }}>Add alt text</button>}
            </div>
          ))}
        </section>
      ))}
    </Dialog>
  );
};

// ---------- Alt text ----------

export const AltTextDialog = ({ store, close, data }: DialogProps<{ type: 'chart' | 'picture' | 'shape'; id: string; back?: string }>) => {
  const s = store.sheet;
  const obj = data?.type === 'chart' ? s.charts.find((c) => c.id === data.id) : data?.type === 'picture' ? s.pictures.find((p) => p.id === data.id) : s.shapes?.find((x) => x.id === data?.id);
  const [text, setText] = useState(obj?.alt ?? '');
  if (!obj || !data) { queueMicrotask(close); return null; }
  const save = () => {
    const alt = text.trim() || undefined;
    if (data.type === 'shape') updateShape(store, data.id, { alt }, 'Alt text');
    else store.change('Alt text', 'sheet', () => {
      const o = data.type === 'chart' ? store.sheet.charts.find((c) => c.id === data.id) : store.sheet.pictures.find((p) => p.id === data.id);
      if (o) { if (alt) o.alt = alt; else delete o.alt; }
    });
    if (data.back) { store.dialog = { id: data.back }; store.emit(); return false; }
    return true;
  };
  return (
    <Dialog title="Alt text" close={close} width={420} okLabel="Save" onOk={save}>
      <p className="muted small">Describe the {data.type} for people who use a screen reader.</p>
      <textarea autoFocus className="b-note" value={text} maxLength={2000} onChange={(e) => setText(e.target.value)} placeholder="For example: Sales by month, rising from 10 to 40" />
    </Dialog>
  );
};

// ---------- Performance ----------

export const PerformanceDialog = ({ store, close }: DialogProps) => {
  const [done, setDone] = useState<{ removed: number; bytes: number } | null>(null);
  const x = excessFormatting(store.sheet);
  const n = x.cells.length + x.rows.length + x.cols.length;
  const kb = (b: number) => (b < 1024 ? `${b} bytes` : `${(b / 1024).toFixed(1)} KB`);
  return (
    <Dialog title="Check performance" close={close} width={420}
      footer={n > 0 && !done ? <button className="primary" onClick={() => setDone(cleanExcessFormatting(store))}>Clean up</button> : undefined}>
      {done ? <p>Removed formatting from {done.removed.toLocaleString()} {done.removed === 1 ? 'place' : 'places'}. The file is about {kb(done.bytes)} smaller.</p>
        : n ? <>
          <p>This sheet has formatting outside the cells that hold data:</p>
          <ul className="small">
            {x.cells.length > 0 && <li>{x.cells.length.toLocaleString()} empty formatted cells</li>}
            {x.rows.length > 0 && <li>{x.rows.length.toLocaleString()} formatted rows</li>}
            {x.cols.length > 0 && <li>{x.cols.length.toLocaleString()} formatted columns</li>}
          </ul>
          <p className="muted small">Cleaning it up saves about {kb(x.bytes)} and makes the sheet faster. You can undo it.</p>
        </> : <p>No extra formatting found on this sheet.</p>}
    </Dialog>
  );
};

// ---------- Allow edit ranges ----------

export const EditRangesDialog = ({ store, close }: DialogProps) => {
  const s = store.sheet;
  const [title, setTitle] = useState('');
  const [ref, setRef] = useState(rangeAddr(store.range));
  const [pw, setPw] = useState('');
  const [unlock, setUnlock] = useState<{ i: number; pw: string } | null>(null);
  const list = s.editRanges ?? [];
  const add = () => {
    const r = parseRange(ref.replace(/\$/g, '').replace(/^.*!/, ''));
    if (!r) { store.notify('Type a range like A1:B10.'); return; }
    addEditRange(store, title, r, pw || undefined);
    setTitle(''); setPw('');
  };
  return (
    <Dialog title="Allow edit ranges" close={close} width={460}>
      <p className="muted small">People can change these ranges while the sheet is protected. A range with a password asks for it first.</p>
      <div className="b-list obj-ranges">
        {!list.length && <div className="muted small pad">No ranges yet.</div>}
        {list.map((e, i) => (
          <div key={i} className="obj-range">
            <span><b>{e.title}</b> <span className="mono">{rangeAddr(e.range)}</span>{e.password && <span className="muted small"> · password</span>}</span>
            <span>
              {e.password && s.protection?.sheet && <button onClick={() => setUnlock({ i, pw: '' })}>Unlock</button>}
              <button onClick={() => removeEditRange(store, i)} aria-label={`Remove ${e.title}`}>Remove</button>
            </span>
          </div>
        ))}
      </div>
      {unlock && (
        <div className="b-inline">
          <input type="password" autoFocus placeholder="Password" aria-label="Range password" value={unlock.pw} onChange={(e) => setUnlock({ ...unlock, pw: e.target.value })} />
          <button className="primary" onClick={() => { if (unlockEditRange(store, unlock.i, unlock.pw)) { store.notify('Range unlocked.'); setUnlock(null); } else store.notify('That password is not right.'); }}>Unlock</button>
        </div>
      )}
      <h3 className="b-h3">New range</h3>
      <Field label="Title"><input value={title} onChange={(e) => setTitle(e.target.value)} placeholder="Range1" /></Field>
      <Field label="Cells"><input className="mono" value={ref} onChange={(e) => setRef(e.target.value)} /></Field>
      <Field label="Password (optional)"><input type="password" value={pw} onChange={(e) => setPw(e.target.value)} /></Field>
      <div><button className="primary" onClick={add}>Add range</button></div>
    </Dialog>
  );
};

// ---------- Spelling ----------

interface Speller { correct: (w: string) => boolean; suggest: (w: string) => string[]; add: (w: string) => void }
let spellerPromise: Promise<Speller> | null = null;

/** Loads the dictionary only when spelling is first checked, so it stays out of the main bundle. */
const loadSpeller = (): Promise<Speller> => spellerPromise ??= (async () => {
  const [{ default: nspell }, aff, dic] = await Promise.all([
    import('nspell'),
    import('../../../node_modules/dictionary-en/index.aff?raw'),
    import('../../../node_modules/dictionary-en/index.dic?raw'),
  ]);
  const sp = nspell(aff.default, dic.default);
  for (const w of personalWords()) sp.add(w);
  return sp;
})();

const DICT_KEY = 'clerdogrid.dictionary';
const personalWords = (): string[] => { try { return JSON.parse(localStorage.getItem(DICT_KEY) ?? '[]') as string[]; } catch { return []; } };
const savePersonal = (w: string) => { try { localStorage.setItem(DICT_KEY, JSON.stringify([...new Set([...personalWords(), w])].slice(-5000))); } catch { /* private mode */ } };

interface Miss { cell: SpellCell; word: string; start: number }

export const SpellingDialog = ({ store, close }: DialogProps) => {
  const [sp, setSp] = useState<Speller | null>(null);
  const [err, setErr] = useState('');
  const [cells] = useState(() => spellCells(store));
  const [pos, setPos] = useState({ cell: 0, from: 0 });
  const [ignored] = useState(() => new Set<string>());
  const [changeAll] = useState(() => new Map<string, string>());
  const [repl, setRepl] = useState('');
  const [finished, setFinished] = useState(false);

  useEffect(() => { loadSpeller().then(setSp, () => setErr('The dictionary could not be loaded.')); }, []);

  const textOf = (c: SpellCell) => { const v = store.cell(c.row, c.col, c.sheet)?.v; return typeof v === 'string' ? v : ''; };

  /** Next misspelling from a position; applies Change all replacements on the way. */
  const findNext = (from: { cell: number; from: number }): Miss | null => {
    if (!sp) return null;
    for (let i = from.cell; i < cells.length; i++) {
      const c = cells[i]!;
      const text = textOf(c);
      for (const w of wordsIn(text)) {
        if (i === from.cell && w.start < from.from) continue;
        if (ignored.has(w.word.toLowerCase()) || sp.correct(w.word)) continue;
        return { cell: c, word: w.word, start: w.start };
      }
    }
    return null;
  };

  const miss = useMemo(() => (sp && !finished ? findNext(pos) : null), [sp, pos, finished]);
  const suggestions = useMemo(() => (miss && sp ? sp.suggest(miss.word).slice(0, 8) : []), [miss, sp]);
  useEffect(() => {
    if (miss) { goTo(store, miss.cell.sheet, miss.cell); setRepl(changeAll.get(miss.word.toLowerCase()) ?? suggestions[0] ?? miss.word); }
    else if (sp && !finished) setFinished(true);
  }, [miss, suggestions]);

  // Apply any Change-all replacement automatically.
  useEffect(() => {
    if (miss && changeAll.has(miss.word.toLowerCase())) replace(changeAll.get(miss.word.toLowerCase())!);
  }, [miss]);

  const advance = (m: Miss, len: number) => setPos({ cell: cells.indexOf(m.cell), from: m.start + len });
  const replace = (to: string) => {
    if (!miss) return;
    const m = miss;
    const text = textOf(m.cell);
    const next = text.slice(0, m.start) + to + text.slice(m.start + m.word.length);
    store.change('Spelling', [{ area: { r1: m.cell.row, c1: m.cell.col, r2: m.cell.row, c2: m.cell.col }, sheet: m.cell.sheet }], () => store.engine.setValue(m.cell.sheet, m.cell.row, m.cell.col, next));
    advance(m, to.length);
  };

  return (
    <Dialog title="Spelling" close={close} width={460}>
      {err && <p>{err}</p>}
      {!sp && !err && <p className="muted">Loading the dictionary…</p>}
      {sp && !cells.length && <p>There is no text to check on this sheet.</p>}
      {sp && cells.length > 0 && finished && <p>The spelling check is complete.</p>}
      {miss && (
        <div className="b-col">
          <div className="muted small">Not in the dictionary · {addr(miss.cell.row, miss.cell.col)}</div>
          <div className="spell-ctx">{textOf(miss.cell).slice(Math.max(0, miss.start - 40), miss.start)}<mark>{miss.word}</mark>{textOf(miss.cell).slice(miss.start + miss.word.length, miss.start + miss.word.length + 40)}</div>
          <div className="spell-grid">
            <div className="b-col">
              <Field label="Change to"><input value={repl} onChange={(e) => setRepl(e.target.value)} /></Field>
              <div className="b-list b-short" role="listbox" aria-label="Suggestions">
                {suggestions.length ? suggestions.map((s) => <button key={s} className={s === repl ? 'on' : ''} onClick={() => setRepl(s)} onDoubleClick={() => replace(s)}>{s}</button>) : <div className="muted small pad">No suggestions</div>}
              </div>
            </div>
            <div className="spell-btns">
              <button onClick={() => advance(miss, miss.word.length)}>Ignore once</button>
              <button onClick={() => { ignored.add(miss.word.toLowerCase()); advance(miss, miss.word.length); }}>Ignore all</button>
              <button onClick={() => { sp!.add(miss.word); savePersonal(miss.word); advance(miss, miss.word.length); }}>Add to dictionary</button>
              <button className="primary" disabled={!repl.trim()} onClick={() => replace(repl)}>Change</button>
              <button disabled={!repl.trim()} onClick={() => { changeAll.set(miss.word.toLowerCase(), repl); replace(repl); }}>Change all</button>
            </div>
          </div>
        </div>
      )}
    </Dialog>
  );
};
