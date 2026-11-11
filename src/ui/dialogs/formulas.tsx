import { useMemo, useState } from 'react';
import { Pencil, Plus, RotateCcw, Trash2 } from 'lucide-react';
import { CATALOG, catalogByName, searchCatalog, signature, type FnArg, type FnCategory, type FnInfo } from '../../engine/catalog';
import { mapNode, nodeText, parseFormula, quoteSheet, walk, type Node, type RefNode } from '../../engine/parser';
import { CellError, isError, isMatrix, MAX_COLS, MAX_ROWS, type ErrorCode, type Matrix, type Scalar } from '../../engine/types';
import { addr, keyCol, keyRow, parseAddr, type CellRange, type DefinedName } from '../../model/workbook';
import type { GridStore } from '../../app/store';
import { Dialog, Field, type DialogProps } from './Dialog';
import { goTo, rangeLabel, resultText, sheetRef } from './shared';
import { N, T, Tr } from '../i18n';

const evalText = (store: GridStore, text: string, sheet = store.sheetIndex): { ok: boolean; text: string; value?: Scalar | Matrix } => {
  if (!text.trim()) return { ok: true, text: '' };
  try {
    parseFormula(text);
    const { row, col } = store.sel.active;
    const v = store.engine.evaluate(text, sheet, row, col);
    return { ok: true, text: resultText(v), value: v };
  } catch {
    return { ok: false, text: N('Not a valid formula yet') };
  }
};

// ---------- Insert function ----------

const CATEGORIES: (FnCategory | 'All' | 'Recent')[] = ['All', 'Recent', 'Financial', 'Date & Time', 'Math & Trig', 'Statistical', 'Lookup & Reference', 'Database', 'Text', 'Logical', 'Information', 'Engineering', 'Web', 'Compatibility'];
/** Category names for the catalogue (the ids stay English). */
export const CATEGORY_LABELS = [N('All'), N('Recent'), N('Financial'), N('Date & Time'), N('Math & Trig'), N('Statistical'), N('Lookup & Reference'), N('Database'), N('Text'), N('Logical'), N('Information'), N('Engineering'), N('Web'), N('Compatibility')];
const RECENT_KEY = 'clerdogrid.recentFunctions';
const readRecent = (): string[] => { try { const v = JSON.parse(localStorage.getItem(RECENT_KEY) ?? '[]'); return Array.isArray(v) ? v.filter((x): x is string => typeof x === 'string').slice(0, 10) : []; } catch { return []; } };
const pushRecent = (name: string) => { try { localStorage.setItem(RECENT_KEY, JSON.stringify([name, ...readRecent().filter((n) => n !== name)].slice(0, 10))); } catch { /* storage unavailable */ } };

/** Inputs for each argument; repeating groups grow as the last one is filled. */
const argSlots = (f: FnInfo, values: string[]): FnArg[] => {
  const lead = f.args.findIndex((a) => a.repeat);
  if (lead < 0) return [...f.args];
  const rep = f.args.filter((a) => a.repeat);
  const before = f.args.slice(0, lead).filter((a) => !a.repeat);
  const after = f.args.slice(lead).filter((a) => !a.repeat);
  let groups = 1;
  for (;;) {
    const start = before.length + (groups - 1) * rep.length;
    const filled = values.slice(start, start + rep.length).some((v) => v?.trim());
    if (!filled || groups >= Math.floor(255 / rep.length)) break;
    groups++;
  }
  const renumber = (name: string, g: number) => (g === 0 ? name : `${name.replace(/\d+$/, '')}${Number(/\d+$/.exec(name)?.[0] ?? 1) + g}`);
  const slots = [...before];
  for (let g = 0; g < groups; g++) slots.push(...rep.map((a) => ({ ...a, name: renumber(a.name, g) })));
  slots.push(...after);
  return slots;
};

const buildCall = (f: FnInfo, values: string[], slots: FnArg[]): string => {
  let last = -1;
  slots.forEach((_, i) => { if (values[i]?.trim()) last = i; });
  slots.forEach((a, i) => { if (!a.optional && !a.repeat && i > last) last = i; });
  return `${f.name}(${slots.slice(0, last + 1).map((_, i) => values[i]?.trim() ?? '').join(',')})`;
};

export const InsertFunctionDialog = ({ store, data, close }: DialogProps<{ category?: string; name?: string }>) => {
  const [q, setQ] = useState('');
  const [cat, setCat] = useState<string>(data?.category && CATEGORIES.includes(data.category as FnCategory) ? data.category : 'All');
  const [picked, setPicked] = useState<FnInfo | null>(data?.name ? catalogByName(data.name) ?? null : null);
  const [step, setStep] = useState<'pick' | 'args'>(data?.name && catalogByName(data.name) ? 'args' : 'pick');
  const [values, setValues] = useState<string[]>([]);
  const list = useMemo(() => {
    const recent = readRecent();
    let fs: readonly FnInfo[] = q.trim() ? searchCatalog(q, 400) : CATALOG;
    if (cat === 'Recent') fs = recent.map((n) => catalogByName(n)).filter((x): x is FnInfo => !!x);
    else if (cat !== 'All') fs = fs.filter((f) => f.category === cat);
    return fs.slice(0, 600);
  }, [q, cat]);
  const cur = picked ?? list[0] ?? null;

  const insert = (): boolean => {
    if (!cur) return false;
    const slots = argSlots(cur, values);
    const call = buildCall(cur, values, slots);
    pushRecent(cur.name);
    const e = store.editing;
    if (e) {
      const t = e.text;
      store.setEditText(t.startsWith('=') ? (t.length > 1 && !/[(,=+\-*/^&<>:]$/.test(t) ? `${t}+${call}` : `${t}${call}`) : `=${call}`);
      return true;
    }
    store.beginEdit(`=${call}`);
    if (!store.editing) return true;
    if (!store.commitEdit()) return false;
    return true;
  };

  if (step === 'args' && cur) {
    const slots = argSlots(cur, values);
    const call = buildCall(cur, values, slots);
    const res = evalText(store, call);
    return (
      <Dialog title={cur.name} close={close} onOk={insert} okLabel={T('Insert')} width={560} footer={<button onClick={() => setStep('pick')}>{T('Back')}</button>}>
        <p className="muted small">{cur.about}</p>
        <div className="b-args">
          {slots.map((a, i) => {
            const v = values[i] ?? '';
            const pv = v.trim() ? evalText(store, v) : null;
            return (
              <div key={`${a.name}:${i}`} className="b-arg">
                <label className={a.optional || a.repeat ? 'muted' : ''} htmlFor={`b-arg-${i}`}>{a.name}</label>
                <input id={`b-arg-${i}`} className="mono" autoFocus={i === 0} value={v} onChange={(e) => { const n = [...values]; n[i] = e.target.value; setValues(n); }} placeholder={a.optional || a.repeat ? 'optional' : ''} />
                <span className={`b-argval${pv && !pv.ok ? ' bad' : ''}`}>{pv ? `= ${pv.text}` : ''}</span>
                <span className="b-argabout">{a.about}</span>
              </div>
            );
          })}
          {!slots.length && <p className="muted small">{T('This function has no arguments.')}</p>}
        </div>
        <div className="preview b-fxprev"><span className="mono">={call}</span><span className="grow" /><span className={res.ok ? '' : 'b-bad'}>{res.text ? `= ${res.text}` : ''}</span></div>
      </Dialog>
    );
  }

  return (
    <Dialog title={T('Insert function')} close={close} wide okLabel={T('Next')} onOk={() => { if (cur) { setPicked(cur); setValues([]); setStep('args'); } return false; }}>
      <div className="b-fnsearch">
        <input autoFocus value={q} onChange={(e) => { setQ(e.target.value); setPicked(null); }} placeholder={T('Search functions, like sum or date')} aria-label={T('Search functions')} />
        <select value={cat} onChange={(e) => { setCat(e.target.value); setPicked(null); }} aria-label={T('Category')}>{CATEGORIES.map((c) => <option key={c} value={c}>{Tr(c)}</option>)}</select>
      </div>
      <div className="b-split b-fnsplit">
        <div className="b-list b-fnlist" role="listbox" aria-label={T('Functions')}>
          {list.map((f) => <button key={f.name} role="option" aria-selected={cur?.name === f.name} className={cur?.name === f.name ? 'on' : ''} onClick={() => setPicked(f)} onDoubleClick={() => { setPicked(f); setValues([]); setStep('args'); }}>{f.name}</button>)}
          {!list.length && <div className="muted pad">{T('Nothing matches.')}</div>}
        </div>
        <div className="b-col b-fninfo">
          {cur && <>
            <div className="mono b-sig">{signature(cur)}</div>
            <p>{cur.about}</p>
            <p className="muted small">{cur.category}{cur.returns ? T(' · returns {returns}', { returns: cur.returns }) : ''}</p>
            <ul className="b-arglist">{cur.args.map((a, i) => <li key={i}><b>{a.name}</b>{a.optional || a.repeat ? T(' (optional)') : ''}: {a.about}</li>)}</ul>
          </>}
        </div>
      </div>
    </Dialog>
  );
};

// ---------- Names ----------

const RESERVED = /^(TRUE|FALSE|R|C|R\d*C\d*)$/i;
export const nameError = (store: GridStore, name: string, scope: number | undefined, except?: DefinedName): string | null => {
  if (!name) return T('Enter a name.');
  if (name.length > 255) return T('A name can be at most 255 characters.');
  if (!/^[A-Za-z_\\À-￿][A-Za-z0-9_.\\À-￿]*$/.test(name)) return T('Start with a letter or underscore, and use only letters, numbers, periods and underscores (no spaces).');
  if (parseAddr(name) || /^[A-Za-z]{1,3}\d+$/.test(name) || RESERVED.test(name)) return T("A name can't look like a cell address.");
  const clash = store.wb.names.find((n) => n !== except && n.name.toLowerCase() === name.toLowerCase() && n.sheet === scope);
  if (clash) return scope === undefined ? T('"{name}" is already used in this workbook.', { name }) : T('"{name}" is already used on that sheet.', { name });
  if (store.wb.sheets.some((s) => s.tables.some((t) => t.name.toLowerCase() === name.toLowerCase()))) return T('A table already uses that name.');
  return null;
};

const renameInFormulas = (store: GridStore, from: string, to: string): void => {
  const lower = from.toLowerCase();
  store.wb.sheets.forEach((s, si) => {
    for (const [k, c] of s.cells) {
      if (!c.f || !c.f.toLowerCase().includes(lower)) continue;
      try {
        let hit = false;
        const out = mapNode(parseFormula(c.f), (n) => (n.type === 'name' && n.name.toLowerCase() === lower ? (hit = true, { ...n, name: to }) : n));
        if (hit) store.engine.setFormula(si, keyRow(k), keyCol(k), nodeText(out));
      } catch { /* leave formulas we can't parse */ }
    }
  });
};

const NameEditor = ({ store, initial, onDone, onCancel }: { store: GridStore; initial?: DefinedName; onDone: () => void; onCancel?: () => void }) => {
  const [name, setName] = useState(initial?.name ?? '');
  const [scope, setScope] = useState<number | -1>(initial?.sheet ?? -1);
  const [refers, setRefers] = useState(initial ? `=${initial.formula}` : `=${store.sel.ranges.map((r) => sheetRef(store, store.sheetIndex, r)).join(',')}`);
  const [comment, setComment] = useState(initial?.comment ?? '');
  const preview = evalText(store, refers.replace(/^=/, ''), scope >= 0 ? scope : store.sheetIndex);
  const save = (): boolean => {
    const n = name.trim();
    const sc = scope >= 0 ? scope : undefined;
    const err = nameError(store, n, sc, initial);
    if (err) { store.notify(err); return false; }
    const f = refers.trim().replace(/^=/, '');
    if (!f) { store.notify(T('Enter what the name refers to, like =Sheet1!$A$1:$A$9.')); return false; }
    try { parseFormula(f); } catch { store.notify(T('"Refers to" has a problem. Check brackets, quotes and sheet names.')); return false; }
    store.change(initial ? N('Edit name') : N('Define name'), 'book', () => {
      const entry: DefinedName = { name: n, formula: f, sheet: sc, comment: comment.trim() || undefined, hidden: initial?.hidden };
      if (initial) {
        const i = store.wb.names.findIndex((x) => x.name === initial.name && x.sheet === initial.sheet);
        if (i >= 0) store.wb.names[i] = entry; else store.wb.names.push(entry);
        if (initial.name !== n) renameInFormulas(store, initial.name, n);
      } else store.wb.names.push(entry);
      store.engine.sheetsChanged();
      store.engine.recalc(true);
    });
    onDone();
    return true;
  };
  return (
    <Dialog title={initial ? T('Edit name') : T('New name')} close={onCancel ?? onDone} onOk={save} width={460} footer={onCancel ? <button onClick={onCancel}>{T('Back')}</button> : undefined}>
      <Field label={T('Name')}><input autoFocus value={name} onChange={(e) => setName(e.target.value)} placeholder={T('SalesTotal')} /></Field>
      <Field label={T('Scope')}><select value={scope} onChange={(e) => setScope(+e.target.value)}><option value={-1}>{T('Workbook')}</option>{store.wb.sheets.map((s, i) => <option key={i} value={i}>{s.name}</option>)}</select></Field>
      <Field label={T('Refers to')}><input className="mono" value={refers} onChange={(e) => setRefers(e.target.value)} /></Field>
      <div className="muted small">{T('Value:')}{' '}{preview.text || '—'}</div>
      <Field label={T('Comment')}><textarea value={comment} maxLength={255} onChange={(e) => setComment(e.target.value)} /></Field>
    </Dialog>
  );
};

export const DefineNameDialog = ({ store, close }: DialogProps) => <NameEditor store={store} onDone={close} />;

export const NamesDialog = ({ store, close }: DialogProps) => {
  const [filter, setFilter] = useState<'all' | 'book' | 'sheet' | 'errors'>('all');
  const [q, setQ] = useState('');
  const [edit, setEdit] = useState<{ name?: DefinedName } | null>(null);
  const [, bump] = useState(0);
  const rows = store.wb.names
    .map((n) => ({ n, v: evalText(store, n.formula, n.sheet ?? store.sheetIndex) }))
    .filter(({ n, v }) => !n.hidden && (!q || n.name.toLowerCase().includes(q.toLowerCase()))
      && (filter === 'all' || (filter === 'book' && n.sheet === undefined) || (filter === 'sheet' && n.sheet !== undefined) || (filter === 'errors' && (!v.ok || (v.value !== undefined && !isMatrix(v.value) && isError(v.value))))));
  if (edit) return <NameEditor store={store} initial={edit.name} onDone={() => { setEdit(null); bump((x) => x + 1); }} onCancel={() => setEdit(null)} />;
  const remove = (n: DefinedName) => {
    store.change(N('Delete name'), 'book', () => {
      store.wb.names = store.wb.names.filter((x) => x !== n);
      store.engine.sheetsChanged();
      store.engine.recalc(true);
    });
    bump((x) => x + 1);
  };
  return (
    <Dialog title={T('Names')} close={close} wide footer={<button onClick={() => setEdit({})}><Plus size={15} strokeWidth={1.9} />{' '}{T('New name')}</button>}>
      <div className="b-inline">
        <input value={q} onChange={(e) => setQ(e.target.value)} placeholder={T('Filter by name')} aria-label={T('Filter by name')} />
        <select value={filter} onChange={(e) => setFilter(e.target.value as typeof filter)} aria-label={T('Show')}>
          <option value="all">{T('All names')}</option><option value="book">{T('Workbook scope')}</option><option value="sheet">{T('Sheet scope')}</option><option value="errors">{T('With errors')}</option>
        </select>
      </div>
      <div className="b-table" role="table">
        <div className="b-tr b-th b-namerow" role="row"><span>{T('Name')}</span><span>{T('Value')}</span><span>{T('Refers to')}</span><span>{T('Scope')}</span><span /></div>
        {rows.map(({ n, v }) => (
          <div key={`${n.name}:${n.sheet ?? ''}`} className="b-tr b-namerow" role="row" title={n.comment}>
            <span className="b-ellip"><b>{n.name}</b></span>
            <span className="b-ellip">{v.text}</span>
            <span className="b-ellip mono">={n.formula}</span>
            <span className="b-ellip">{n.sheet === undefined ? T('Workbook') : store.wb.sheets[n.sheet]?.name ?? '?'}</span>
            <span className="b-rowtools">
              <button className="ico sm" aria-label={T('Edit {name}', { name: n.name })} onClick={() => setEdit({ name: n })}><Pencil size={15} strokeWidth={1.9} /></button>
              <button className="ico sm" aria-label={T('Delete {name}', { name: n.name })} onClick={() => remove(n)}><Trash2 size={15} strokeWidth={1.9} /></button>
            </span>
          </div>
        ))}
        {!rows.length && <div className="muted pad">{store.wb.names.length ? T('No names match.') : T('No names yet. A name gives a cell, range or formula a readable label.')}</div>}
      </div>
    </Dialog>
  );
};

// ---------- Evaluate formula ----------

const literal = (v: Scalar | Matrix): Node => {
  if (isMatrix(v)) return { type: 'array', rows: v.slice(0, 20).map((r) => r.slice(0, 20).map((x) => literal(x))) };
  if (v === null) return { type: 'num', value: 0 };
  if (v instanceof CellError) return { type: 'err', code: v.code };
  if (typeof v === 'number') return { type: 'num', value: v };
  if (typeof v === 'boolean') return { type: 'bool', value: v };
  return { type: 'str', value: v };
};

const steppable = (n: Node): boolean => n.type === 'ref' || n.type === 'name' || n.type === 'table' || n.type === 'call' || n.type === 'invoke' || n.type === 'unary' || (n.type === 'binary' && n.op !== ':' && n.op !== ' ');

const postOrder = (n: Node, out: Node[]): Node[] => {
  switch (n.type) {
    case 'array': n.rows.forEach((r) => r.forEach((x) => postOrder(x, out))); break;
    case 'call': n.args.forEach((x) => postOrder(x, out)); break;
    case 'invoke': n.args.forEach((x) => postOrder(x, out)); break;
    case 'unary': case 'spill': postOrder(n.arg, out); break;
    case 'binary': if (n.op !== ':' && n.op !== ' ') { postOrder(n.left, out); postOrder(n.right, out); } break;
    case 'union': n.items.forEach((x) => postOrder(x, out)); break;
    default: break;
  }
  if (steppable(n)) out.push(n);
  return out;
};

/** Rebuilds the tree with evaluated nodes replaced by their values, and the next node replaced by a marker. */
const reduce = (n: Node, done: Map<Node, Scalar | Matrix>, mark: Node | null): Node => {
  if (done.has(n)) return literal(done.get(n)!);
  if (n === mark) return { type: 'name', name: '\u0001MARK' };
  const r = (x: Node) => reduce(x, done, mark);
  switch (n.type) {
    case 'array': return { ...n, rows: n.rows.map((row) => row.map(r)) };
    case 'call': return { ...n, args: n.args.map(r) };
    case 'invoke': return { ...n, callee: n.callee, args: n.args.map(r) };
    case 'unary': case 'spill': return { ...n, arg: r(n.arg) };
    case 'binary': return { ...n, left: r(n.left), right: r(n.right) };
    case 'union': return { ...n, items: n.items.map(r) };
    default: return n;
  }
};

export const EvaluateDialog = ({ store, close }: DialogProps) => {
  const { row, col } = store.sel.active;
  const f = store.cell(row, col)?.f;
  const tree = useMemo(() => { try { return f ? parseFormula(f) : null; } catch { return null; } }, [f]);
  const order = useMemo(() => (tree ? postOrder(tree, []) : []), [tree]);
  const [n, setN] = useState(0);
  const done = useMemo(() => {
    const m = new Map<Node, Scalar | Matrix>();
    for (const node of order.slice(0, n)) m.set(node, store.engine.evaluate(nodeText(node), store.sheetIndex, row, col));
    return m;
  }, [order, n, store, row, col]);
  if (!f || !tree) {
    return <Dialog title={T('Evaluate formula')} close={close} width={420}><p>{f ? T('This formula can\'t be read.') : T('{p0} doesn\'t hold a formula. Select a cell with a formula first.', { p0: addr(row, col) })}</p></Dialog>;
  }
  const next = order[n] ?? null;
  const shown = nodeText(reduce(tree, done, next));
  const [pre, post] = shown.includes('\u0001MARK') ? shown.split('\u0001MARK') as [string, string] : [shown, ''];
  const nextText = next ? nodeText(reduce(next, done, null)) : '';
  const finished = !next;
  return (
    <Dialog title={T('Evaluate formula')} close={close} wide footer={<>
      <button onClick={() => setN(0)} disabled={n === 0}><RotateCcw size={15} strokeWidth={1.9} />{' '}{T('Restart')}</button>
      <button onClick={() => setN(order.length)} disabled={finished}>{T('Show result')}</button>
      <button onClick={() => setN(n + 1)} disabled={finished}>{T('Evaluate')}</button>
    </>}>
      <div className="muted small">{store.sheet.name}!{addr(row, col)}</div>
      <div className="preview b-evalbox mono">=
        {pre}{next && <mark className="b-mark">{nextText}</mark>}{post}
      </div>
      <p className="muted small">{finished ? T('Result: {p0}', { p0: resultText(store.value(row, col)) }) : T(next!.type === 'ref' || next!.type === 'name' ? 'Next, look up the highlighted part. Step {step} of {count}.' : 'Next, work out the highlighted part. Step {step} of {count}.', { step: n + 1, count: order.length })}</p>
      <div className="b-table b-steps" role="table">
        {order.slice(0, n).map((node, i) => (
          <div key={i} className="b-tr" role="row"><span className="mono b-ellip">{nodeText(node)}</span><span className="mono b-ellip">{resultText(done.get(node)!)}</span></div>
        ))}
      </div>
    </Dialog>
  );
};

// ---------- Trace precedents / dependents ----------

interface Hit { sheet: number; range: CellRange; label: string; value: string; via?: string }

const refRange = (r: RefNode): CellRange => ({ r1: r.r1, c1: r.c1, r2: r.kind === 'cols' ? MAX_ROWS - 1 : r.r2, c2: r.kind === 'rows' ? MAX_COLS - 1 : r.c2 });

const refsOf = (store: GridStore, formula: string, sheet: number, depth = 0): { sheet: number; range: CellRange; via?: string }[] => {
  let tree: Node;
  try { tree = parseFormula(formula); } catch { return []; }
  const out: { sheet: number; range: CellRange; via?: string }[] = [];
  walk(tree, (n) => {
    if (n.type === 'ref' && !n.ref.external) {
      const si = n.ref.sheet ? store.engine.sheetIndex(n.ref.sheet) : sheet;
      if (si >= 0) out.push({ sheet: si, range: refRange(n.ref) });
    } else if (n.type === 'name' && depth < 4) {
      const def = store.engine.definedName(n.name, sheet);
      if (def) for (const r of refsOf(store, def, sheet, depth + 1)) out.push({ ...r, via: n.name });
    } else if (n.type === 'table') {
      const t = store.engine.table(n.table || store.engine.tableAt(sheet, 0, 0) || '');
      if (t) out.push({ sheet: t.sheet, range: { r1: t.r1, c1: t.c1, r2: t.r2, c2: t.c2 }, via: n.table });
    }
  });
  return out;
};

const describe = (store: GridStore, sheet: number, r: CellRange): Hit => {
  const single = r.r1 === r.r2 && r.c1 === r.c2;
  return { sheet, range: r, label: `${quoteSheet(store.wb.sheets[sheet]!.name)}!${rangeLabel(r)}`, value: single ? resultText(store.value(r.r1, r.c1, sheet)) : `${(r.r2 - r.r1 + 1).toLocaleString()} × ${(r.c2 - r.c1 + 1).toLocaleString()} cells` };
};

export const TraceDialog = ({ store, data, close }: DialogProps<{ dir?: 'precedents' | 'dependents' }>) => {
  const [dir, setDir] = useState<'precedents' | 'dependents'>(data?.dir ?? 'precedents');
  const { row, col } = store.sel.active;
  const si = store.sheetIndex;
  const hits = useMemo<Hit[]>(() => {
    if (dir === 'precedents') {
      const f = store.cell(row, col)?.f;
      if (!f) return [];
      const seen = new Set<string>();
      return refsOf(store, f, si).filter((h) => { const k = `${h.sheet}:${h.range.r1}:${h.range.c1}:${h.range.r2}:${h.range.c2}`; if (seen.has(k)) return false; seen.add(k); return true; })
        .map((h) => ({ ...describe(store, h.sheet, h.range), via: h.via }));
    }
    const out: Hit[] = [];
    let scanned = 0;
    store.wb.sheets.forEach((s, sIdx) => {
      for (const [k, c] of s.cells) {
        if (!c.f || ++scanned > 200000) continue;
        const r = keyRow(k), cc = keyCol(k);
        if (refsOf(store, c.f, sIdx).some((h) => h.sheet === si && row >= h.range.r1 && row <= h.range.r2 && col >= h.range.c1 && col <= h.range.c2)) out.push(describe(store, sIdx, { r1: r, c1: cc, r2: r, c2: cc }));
      }
    });
    return out;
  }, [dir, store, row, col, si]);
  return (
    <Dialog title={dir === 'precedents' ? T('Cells this formula uses') : T('Formulas that use this cell')} close={close} width={520}>
      <div className="b-seg" role="radiogroup">
        <button role="radio" aria-checked={dir === 'precedents'} className={dir === 'precedents' ? 'on' : ''} onClick={() => setDir('precedents')}>{T('Uses')}</button>
        <button role="radio" aria-checked={dir === 'dependents'} className={dir === 'dependents' ? 'on' : ''} onClick={() => setDir('dependents')}>{T('Used by')}</button>
      </div>
      <div className="muted small">{store.sheet.name}!{addr(row, col)}{store.cell(row, col)?.f ? ` = ${store.cell(row, col)!.f}` : ''}</div>
      <div className="b-table" role="table">
        {hits.map((h, i) => (
          <button key={i} className="b-tr b-link" onClick={() => { if (h.sheet !== store.sheetIndex) store.setActiveSheet(h.sheet); store.selectRange(h.range); }}>
            <span className="mono">{h.label}</span><span className="b-ellip">{h.value}</span><span className="muted small">{h.via ? T('via {via}', { via: h.via }) : ''}</span>
          </button>
        ))}
        {!hits.length && <div className="muted pad">{dir === 'precedents' ? (store.cell(row, col)?.f ? T('This formula refers to no cells.') : T('This cell doesn\'t hold a formula.')) : T('No formulas refer to this cell.')}</div>}
      </div>
    </Dialog>
  );
};

// ---------- Error checking ----------

export const ERROR_HELP: Partial<Record<ErrorCode, string>> = {
  '#DIV/0!': N('A number is divided by zero or by an empty cell.'),
  '#N/A': N("A lookup didn't find a match, or a value isn't available."),
  '#NAME?': N("The formula uses a name or function that doesn't exist. Check the spelling."),
  '#NULL!': N("Two ranges that don't overlap are joined with a space."),
  '#NUM!': N('A number is out of range, or a calculation could not find an answer.'),
  '#REF!': N('The formula refers to cells that were deleted or are off the sheet.'),
  '#VALUE!': N('A value has the wrong type, like text where a number is needed.'),
  '#SPILL!': N("The result needs more cells, but some aren't empty. Clear them."),
  '#CALC!': N('The calculation gave an empty result, such as a filter with no matches.'),
  '#FIELD!': N("The field doesn't exist in the data."),
  '#BLOCKED!': N('This feature is not available here.'),
  '#CONNECT!': N("Couldn't reach a data source."),
  '#BUSY!': N('Still working on the result.'),
  '#GETTING_DATA': N('Waiting for data.'),
  '#UNKNOWN!': N("This data type isn't supported."),
};

export const ErrorsDialog = ({ store, close }: DialogProps) => {
  const list = useMemo(() => {
    const out: { sheet: number; row: number; col: number; code: string; f?: string; help: string }[] = [];
    const circ = new Set(store.engine.circular.map((c) => `${c.sheet}:${c.row}:${c.col}`));
    for (const c of store.engine.circular) out.push({ sheet: c.sheet, row: c.row, col: c.col, code: 'Circular', f: store.cell(c.row, c.col, c.sheet)?.f, help: N('The formula refers to its own cell, directly or through other cells.') });
    store.wb.sheets.forEach((s, si) => {
      for (const [k, c] of s.cells) {
        if (out.length > 5000) return;
        const row = keyRow(k), col = keyCol(k);
        if (circ.has(`${si}:${row}:${col}`)) continue;
        const v = store.value(row, col, si);
        if (isError(v)) out.push({ sheet: si, row, col, code: v.code, f: c.f, help: (v.detail ? `${v.detail}. ` : '') + Tr(ERROR_HELP[v.code] ?? N('The formula returned an error.')) });
      }
    });
    return out;
  }, [store]);
  return (
    <Dialog title={T('Error checking')} close={close} wide>
      <div className="muted small">{list.length ? T('{count, plural, one {# problem} other {# problems}} in this workbook', { count: list.length }) : T('No errors found in this workbook.')}</div>
      <div className="b-table" role="table">
        {list.slice(0, 2000).map((e, i) => (
          <button key={i} className="b-tr b-link b-errrow" onClick={() => goTo(store, e.sheet, e.row, e.col)}>
            <span className="mono">{quoteSheet(store.wb.sheets[e.sheet]!.name)}!{addr(e.row, e.col)}</span>
            <span className="b-errcode">{e.code}</span>
            <span className="b-ellip mono">{e.f ? `=${e.f}` : ''}</span>
            <span className="b-errhelp">{Tr(e.help)}</span>
          </button>
        ))}
      </div>
    </Dialog>
  );
};
