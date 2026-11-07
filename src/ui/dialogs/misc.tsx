import { useMemo, useState } from 'react';
import { Lock, LockOpen, Search } from 'lucide-react';
import { allCommands } from '../../app/commands';
import { colPx, layoutOf, rowPx } from '../../app/layout';
import { quoteSheet } from '../../engine/parser';
import { cellKey, mergeStyle, parseRange, type Sheet } from '../../model/workbook';
import { prettyKeys } from '../keys';
import { Dialog, Field, type DialogProps } from './Dialog';
import { protectionHash } from '../../app/protection';
import { Check, cleanStyle, Radio, rangeLabel, scalarText } from './shared';

// ---------- Link ----------

const SAFE_URL = /^(https:\/\/[^\s<>"]+|mailto:[^\s<>"]+)$/i;

export const LinkDialog = ({ store, close }: DialogProps) => {
  const { row, col } = store.sel.active;
  const key = cellKey(row, col);
  const cur = store.sheet.hyperlinks.get(key);
  const internal = cur?.startsWith('#');
  const [kind, setKind] = useState<'web' | 'place' | 'email'>(internal ? 'place' : cur?.startsWith('mailto:') ? 'email' : 'web');
  const [text, setText] = useState(() => scalarText(store.value(row, col)));
  const [url, setUrl] = useState(cur && !internal && !cur.startsWith('mailto:') ? cur : '');
  const [email, setEmail] = useState(cur?.startsWith('mailto:') ? decodeURIComponent(cur.slice(7).split('?')[0]!) : '');
  const [subject, setSubject] = useState(() => { const m = /[?&]subject=([^&]*)/.exec(cur ?? ''); return m ? decodeURIComponent(m[1]!) : ''; });
  const parsedPlace = internal ? cur!.slice(1) : '';
  const placeSheet = parsedPlace.includes('!') ? parsedPlace.slice(0, parsedPlace.lastIndexOf('!')).replace(/^'|'$/g, '').replace(/''/g, "'") : store.sheet.name;
  const [sheet, setSheet] = useState(Math.max(0, store.engine.sheetIndex(placeSheet)));
  const [cellRef, setCellRef] = useState(parsedPlace.includes('!') ? parsedPlace.slice(parsedPlace.lastIndexOf('!') + 1) : 'A1');
  const hasFormula = !!store.cell(row, col)?.f;

  const ok = () => {
    let target: string;
    if (kind === 'web') {
      const u = /^[a-z][a-z0-9+.-]*:/i.test(url.trim()) ? url.trim() : `https://${url.trim()}`;
      if (!url.trim() || !SAFE_URL.test(u)) { store.notify('Enter a web address that starts with https://'); return false; }
      try { new URL(u); } catch { store.notify("That web address doesn't look right."); return false; }
      target = u;
    } else if (kind === 'email') {
      if (!/^[^\s@<>"]+@[^\s@<>"]+\.[^\s@<>"]+$/.test(email.trim())) { store.notify('Enter an email address like name@example.com.'); return false; }
      target = `mailto:${email.trim()}${subject.trim() ? `?subject=${encodeURIComponent(subject.trim())}` : ''}`;
    } else {
      if (!parseRange(cellRef.replace(/\$/g, '').toUpperCase())) { store.notify('Enter a cell like A1 or a range like B2:D9.'); return false; }
      target = `#${quoteSheet(store.wb.sheets[sheet]!.name)}!${cellRef.replace(/\$/g, '').toUpperCase()}`;
    }
    const shown = text.trim() || (kind === 'place' ? target.slice(1) : kind === 'email' ? email.trim() : target);
    store.change('Link', 'sheet', () => {
      store.sheet.hyperlinks.set(key, target);
      if (!hasFormula && scalarText(store.value(row, col)) !== shown) store.engine.setValue(store.sheetIndex, row, col, shown);
      store.styleRange({ r1: row, c1: col, r2: row, c2: col }, (s) => cleanStyle(mergeStyle(s, { font: { color: '#0563c1', underline: 'single' } })));
    });
    return true;
  };

  const removeLink = () => {
    store.change('Remove link', 'sheet', () => {
      store.sheet.hyperlinks.delete(key);
      store.styleRange({ r1: row, c1: col, r2: row, c2: col }, (s) => cleanStyle({ ...s, font: s.font ? { ...s.font, color: undefined, underline: undefined } : undefined }));
    });
    close();
  };

  return (
    <Dialog title={cur ? 'Edit link' : 'Insert link'} close={close} onOk={ok} width={460} footer={cur ? <button onClick={removeLink}>Remove link</button> : undefined}>
      <Radio name="lk" row value={kind} options={[['web', 'Web page'], ['place', 'Place in this workbook'], ['email', 'Email']]} onChange={setKind} />
      <Field label="Text to show"><input value={text} disabled={hasFormula} onChange={(e) => setText(e.target.value)} /></Field>
      {kind === 'web' && <Field label="Address"><input autoFocus value={url} onChange={(e) => setUrl(e.target.value)} placeholder="https://example.com" inputMode="url" /></Field>}
      {kind === 'email' && <div className="b-col">
        <Field label="Email address"><input autoFocus value={email} onChange={(e) => setEmail(e.target.value)} placeholder="name@example.com" inputMode="email" /></Field>
        <Field label="Subject"><input value={subject} onChange={(e) => setSubject(e.target.value)} /></Field>
      </div>}
      {kind === 'place' && <div className="grid2">
        <Field label="Sheet"><select value={sheet} onChange={(e) => setSheet(+e.target.value)}>{store.wb.sheets.map((s, i) => <option key={i} value={i}>{s.name}</option>)}</select></Field>
        <Field label="Cell"><input value={cellRef} onChange={(e) => setCellRef(e.target.value)} placeholder="A1" /></Field>
      </div>}
      <p className="muted small">Only https and email links are allowed. Hold Ctrl (Cmd on Mac) and click the cell to open it.</p>
    </Dialog>
  );
};

// ---------- Note ----------

export const NoteDialog = ({ store, close }: DialogProps) => {
  const { row, col } = store.sel.active;
  const key = cellKey(row, col);
  const cur = store.sheet.notes.get(key);
  const [text, setText] = useState(cur?.text ?? '');
  const write = (t: string | null, label: string) => store.change(label, 'sheet', () => {
    if (t === null) store.sheet.notes.delete(key); else store.sheet.notes.set(key, { author: cur?.author ?? 'You', text: t });
  });
  return (
    <Dialog title={cur ? 'Edit note' : 'New note'} close={close} width={400} okLabel="Save"
      onOk={() => { const t = text.trim(); if (!t) { if (cur) write(null, 'Delete note'); return true; } if (t !== cur?.text) write(t, cur ? 'Edit note' : 'Add note'); return true; }}
      footer={cur ? <button onClick={() => { write(null, 'Delete note'); close(); }}>Delete</button> : undefined}>
      <div className="muted small">{store.sheet.name}!{rangeLabel({ r1: row, c1: col, r2: row, c2: col })}{cur?.author ? ` · ${cur.author}` : ''}</div>
      <textarea autoFocus className="b-note" value={text} maxLength={32767} onChange={(e) => setText(e.target.value)} placeholder="Write a note" />
    </Dialog>
  );
};

// ---------- Symbol ----------

const SYMBOL_GROUPS: [string, string][] = [
  ['Math', '±×÷≠≈≤≥∞√∛∑∏∫∂∆∇∈∉∩∪⊂⊃⊆⊇∧∨¬∀∃∅°′″‰∝∠⊥∥'],
  ['Currency', '$€£¥₹₩₽₺₪₦฿₫₴₱¢₿'],
  ['Arrows', '←↑→↓↔↕↖↗↘↙⇐⇑⇒⇓⇔↩↪⟵⟶⤴⤵'],
  ['Greek', 'αβγδεζηθικλμνξοπρστυφχψωΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ'],
  ['Fractions', '½⅓⅔¼¾⅕⅖⅗⅘⅙⅚⅛⅜⅝⅞¹²³⁴⁵⁶⁷⁸⁹⁰₀₁₂₃₄₅₆₇₈₉'],
  ['Punctuation', '©®™§¶†‡•·…–—‹›«»‚„“”‘’¡¿№'],
  ['Shapes', '■□▪▫▲△▼▽◆◇○●◎★☆✓✔✗✘☐☑☒♠♣♥♦'],
  ['Box drawing', '─│┌┐└┘├┤┬┴┼═║╔╗╚╝╠╣╦╩╬░▒▓█'],
];

export const SymbolDialog = ({ store, close }: DialogProps) => {
  const [group, setGroup] = useState(SYMBOL_GROUPS[0]![0]);
  const [code, setCode] = useState('');
  const [pick, setPick] = useState('');
  const chars = [...(SYMBOL_GROUPS.find(([g]) => g === group)?.[1] ?? '')];
  const fromCode = (): string | null => {
    const n = parseInt(code.replace(/^(U\+|0x)/i, ''), 16);
    if (!Number.isFinite(n) || n < 0x20 || n > 0x10ffff || (n >= 0xd800 && n <= 0xdfff)) return null;
    return String.fromCodePoint(n);
  };
  const insert = (ch: string) => {
    if (store.editing) store.setEditText(store.editing.text + ch);
    else store.beginEdit(ch);
    close();
  };
  const current = code.trim() ? fromCode() : pick;
  return (
    <Dialog title="Symbol" close={close} width={480} okLabel="Insert" onOk={() => { if (!current) { store.notify(code.trim() ? 'Enter a code like 00B0 or U+2713.' : 'Pick a symbol.'); return false; } insert(current); return false; }}>
      <div className="b-chips" role="tablist">{SYMBOL_GROUPS.map(([g]) => <button key={g} role="tab" aria-selected={g === group} className={g === group ? 'on' : ''} onClick={() => setGroup(g)}>{g}</button>)}</div>
      <div className="b-symgrid" role="listbox" aria-label={group}>
        {chars.map((c) => <button key={c} role="option" aria-selected={pick === c} className={pick === c ? 'on' : ''} title={`U+${c.codePointAt(0)!.toString(16).toUpperCase().padStart(4, '0')}`} onClick={() => { setPick(c); setCode(''); }} onDoubleClick={() => insert(c)}>{c}</button>)}
      </div>
      <div className="grid2">
        <Field label="Unicode code (hex)"><input value={code} onChange={(e) => setCode(e.target.value)} placeholder="00B0" /></Field>
        <div className="preview b-sympick" aria-live="polite">{current ?? ''}<span className="muted small">{current ? `U+${current.codePointAt(0)!.toString(16).toUpperCase().padStart(4, '0')}` : ''}</span></div>
      </div>
    </Dialog>
  );
};

// ---------- Protect sheet ----------

const PROTECT_OPTIONS: [string, string, boolean][] = [
  ['selectLockedCells', 'Select locked cells', true], ['selectUnlockedCells', 'Select unlocked cells', true],
  ['formatCells', 'Format cells', false], ['formatColumns', 'Format columns', false], ['formatRows', 'Format rows', false],
  ['insertColumns', 'Insert columns', false], ['insertRows', 'Insert rows', false], ['deleteColumns', 'Delete columns', false], ['deleteRows', 'Delete rows', false],
  ['sort', 'Sort', false], ['autoFilter', 'Use filters', false],
];


export const ProtectDialog = ({ store, close }: DialogProps) => {
  const p = store.sheet.protection;
  const on = !!p?.sheet;
  const [pw, setPw] = useState('');
  const [pw2, setPw2] = useState('');
  const [busy, setBusy] = useState(false);
  const [opts, setOpts] = useState<Record<string, boolean>>(() => Object.fromEntries(PROTECT_OPTIONS.map(([k, , d]) => [k, p?.options?.[k] ?? d])));
  const sheetIdx = store.sheetIndex;
  const apply = (next: Sheet['protection'], label: string) => {
    store.change(label, 'sheet', () => { store.wb.sheets[sheetIdx]!.protection = next; });
  };
  const submit = async () => {
    if (busy) return;
    setBusy(true);
    try {
      if (on) {
        if (p?.password) {
          const h = protectionHash(pw);
          if (h.toUpperCase() !== p.password.toUpperCase()) { store.notify("That password isn't right."); return; }
        }
        apply(undefined, 'Unprotect sheet');
        store.notify('Sheet unprotected.');
        close();
      } else {
        if (pw !== pw2) { store.notify("The passwords don't match."); return; }
        const hash = pw ? protectionHash(pw) : undefined;
        apply({ sheet: true, password: hash, options: opts }, 'Protect sheet');
        store.notify(pw ? 'Sheet protected with a password.' : 'Sheet protected.');
        close();
      }
    } catch {
      store.notify("Couldn't set the password on this device.");
    } finally { setBusy(false); }
  };
  return (
    <Dialog title={on ? 'Unprotect sheet' : 'Protect sheet'} close={close} width={440} okLabel={on ? 'Unprotect' : 'Protect'} onOk={() => { void submit(); return false; }}>
      <div className="b-inline muted small">{on ? <LockOpen size={15} strokeWidth={1.9} aria-hidden="true" /> : <Lock size={15} strokeWidth={1.9} aria-hidden="true" />}{store.sheet.name}</div>
      {on ? <>
        <p>{p?.password ? 'Enter the password to unprotect this sheet.' : 'This sheet is protected without a password.'}</p>
        {p?.password && <Field label="Password"><input type="password" autoFocus value={pw} onChange={(e) => setPw(e.target.value)} /></Field>}
      </> : <>
        <p className="muted small">Locked cells can't be changed while the sheet is protected. Unlock cells first in Format cells, Protection.</p>
        <div className="grid2">
          <Field label="Password (optional)"><input type="password" autoFocus value={pw} onChange={(e) => setPw(e.target.value)} autoComplete="new-password" /></Field>
          <Field label="Repeat password"><input type="password" value={pw2} disabled={!pw} onChange={(e) => setPw2(e.target.value)} autoComplete="new-password" /></Field>
        </div>
        <Field label="Everyone can still">
          <div className="b-list b-checklist">{PROTECT_OPTIONS.map(([k, l]) => <Check key={k} label={l} checked={!!opts[k]} onChange={(v) => setOpts({ ...opts, [k]: v })} />)}</div>
        </Field>
        <p className="muted small">If you forget the password, it can't be recovered.</p>
      </>}
    </Dialog>
  );
};

// ---------- Zoom to selection ----------

export const ZoomToSelectionDialog = ({ store, close }: DialogProps) => {
  const s = store.sheet;
  const r = store.range;
  const layout = layoutOf(s, store.layoutVersion);
  const w = layout.cols.start(Math.min(r.c2 + 1, r.c1 + 2000)) - layout.cols.start(r.c1) || colPx(s.defaultColWidth);
  const h = layout.rows.start(Math.min(r.r2 + 1, r.r1 + 20000)) - layout.rows.start(r.r1) || rowPx(s.defaultRowHeight);
  const wrap = document.querySelector<HTMLElement>('.gridwrap');
  const availW = Math.max(200, (wrap?.clientWidth ?? window.innerWidth - 320) - 50);
  const availH = Math.max(150, (wrap?.clientHeight ?? window.innerHeight - 200) - 30);
  const fit = Math.max(10, Math.min(400, Math.floor(Math.min(availW / w, availH / h) * 100)));
  const [zoom, setZoom] = useState(fit);
  const ok = () => {
    const z = Math.max(10, Math.min(400, Math.round(zoom)));
    if (!Number.isFinite(z)) { store.notify('Enter a zoom from 10 to 400.'); return false; }
    s.view.zoom = z;
    s.view.selection = { ...r };
    store.file.dirty = true;
    store.touch();
    return true;
  };
  return (
    <Dialog title="Zoom" close={close} onOk={ok} width={340}>
      <div className="b-chips">{[50, 75, 100, 150, 200].map((z) => <button key={z} className={zoom === z ? 'on' : ''} onClick={() => setZoom(z)}>{z}%</button>)}<button className={zoom === fit ? 'on' : ''} onClick={() => setZoom(fit)}>Fit selection</button></div>
      <Field label="Zoom (%)"><input type="number" min={10} max={400} value={zoom} onChange={(e) => setZoom(+e.target.value)} /></Field>
      <p className="muted small">Fit selection ({rangeLabel(r)}) is {fit}%.</p>
    </Dialog>
  );
};

// ---------- Properties ----------

const toLocal = (iso: string | undefined): string => { if (!iso) return '—'; const d = new Date(iso); return Number.isNaN(+d) ? iso : d.toLocaleString(); };

export const PropertiesDialog = ({ store, close }: DialogProps) => {
  const p = store.wb.props ?? {};
  const [title, setTitle] = useState(p.title ?? '');
  const [author, setAuthor] = useState(p.creator ?? '');
  const stats = useMemo(() => {
    let cells = 0, formulas = 0;
    for (const s of store.wb.sheets) for (const c of s.cells.values()) { if (c.v !== null || c.f) cells++; if (c.f) formulas++; }
    return { cells, formulas };
  }, [store]);
  return (
    <Dialog title="Workbook properties" close={close} width={420} onOk={() => {
      const next = { ...p, title: title.trim() || undefined, creator: author.trim() || undefined };
      if (next.title === p.title && next.creator === p.creator) return true;
      store.change('Properties', 'book', () => { store.wb.props = next; });
      return true;
    }}>
      <Field label="Title"><input value={title} maxLength={255} onChange={(e) => setTitle(e.target.value)} /></Field>
      <Field label="Author"><input value={author} maxLength={255} onChange={(e) => setAuthor(e.target.value)} /></Field>
      <dl className="b-dl">
        <dt>File</dt><dd>{store.file.name}</dd>
        <dt>Created</dt><dd>{toLocal(p.created)}</dd>
        <dt>Modified</dt><dd>{toLocal(p.modified)}</dd>
        <dt>Sheets</dt><dd>{store.wb.sheets.length}</dd>
        <dt>Cells with data</dt><dd>{stats.cells.toLocaleString()}</dd>
        <dt>Formulas</dt><dd>{stats.formulas.toLocaleString()}</dd>
      </dl>
    </Dialog>
  );
};

// ---------- Shortcuts ----------

export const ShortcutsDialog = ({ close }: DialogProps) => {
  const [q, setQ] = useState('');
  const [onlyKeys, setOnlyKeys] = useState(true);
  const rows = useMemo(() => allCommands()
    .filter((c) => (!onlyKeys || c.keys) && (!q || `${c.label} ${c.group ?? ''} ${c.keys ?? ''}`.toLowerCase().includes(q.toLowerCase())))
    .sort((a, b) => (a.group ?? '').localeCompare(b.group ?? '') || a.label.localeCompare(b.label)), [q, onlyKeys]);
  return (
    <Dialog title="Keyboard shortcuts" close={close} wide>
      <div className="b-inline">
        <label className="b-searchbox"><Search size={15} strokeWidth={1.9} aria-hidden="true" /><input autoFocus value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search commands" aria-label="Search commands" /></label>
        <Check label="Only with shortcuts" checked={onlyKeys} onChange={setOnlyKeys} />
      </div>
      <div className="b-table b-keys" role="table">
        {rows.map((c) => (
          <div key={c.id} className="b-tr" role="row"><span>{c.label.replace(/…$/, '')}</span><span className="muted small">{c.group ?? ''}</span><span className="b-kbds">{c.keys ? c.keys.split(',').map((k) => <kbd key={k}>{prettyKeys(k.trim())}</kbd>) : null}</span></div>
        ))}
        {!rows.length && <div className="muted pad">Nothing matches.</div>}
      </div>
    </Dialog>
  );
};

// ---------- About ----------

export const AboutDialog = ({ close }: DialogProps) => (
  <Dialog title="About" close={close} width={380}>
    <div className="b-about">
      <svg width="48" height="48" viewBox="0 0 48 48" aria-hidden="true"><rect width="48" height="48" rx="12" fill="var(--accent)" /><path d="M14 16h20M14 24h20M14 32h20M22 12v24" stroke="var(--on-accent)" strokeWidth="2.4" strokeLinecap="round" /></svg>
      <h3>ClerdoGrid</h3>
      <div className="muted">Version {__APP_VERSION__}</div>
    </div>
    <p>Free and open source, MIT or Apache-2.0.</p>
    <p>No account, no tracking: your files stay on your device.</p>
    <p className="muted small">Made by Clerdo and contributors. Bundled fonts and icons keep their own open licenses, listed in NOTICE.</p>
  </Dialog>
);
