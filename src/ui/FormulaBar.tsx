import { useEffect, useRef, useState } from 'react';
import { Check, ChevronDown, X } from 'lucide-react';
import { MAX_COLS, MAX_ROWS } from '../engine/types';
import { colName, parseRange } from '../model/workbook';
import { run } from '../app/commands';
import type { GridStore } from '../app/store';
import { FormulaAssist } from './FormulaAssist';
import { cycleAbsolute } from './GridView';
import { gotoReference } from '../app/goto';
import { T } from './i18n';

const S = 1.9;

const selectionLabel = (store: GridStore): string => {
  if (store.r1c1) return r1c1Label(store);
  const r = store.range;
  if (store.editing) return `${colName(store.editing.col)}${store.editing.row + 1}`;
  const name = store.wb.names.find((n) => {
    const m = /^(?:'?(.+?)'?!)?\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(n.formula);
    if (!m) return false;
    if (m[1] && m[1] !== store.sheet.name) return false;
    const rr = parseRange(`${m[2]}${m[3]}${m[4] ? `:${m[4]}${m[5]}` : ''}`);
    return rr && rr.r1 === r.r1 && rr.c1 === r.c1 && rr.r2 === r.r2 && rr.c2 === r.c2;
  });
  if (name) return name.name;
  if (r.r1 === 0 && r.r2 === MAX_ROWS - 1 && r.c1 === 0 && r.c2 === MAX_COLS - 1) return `${colName(store.sel.active.col)}${store.sel.active.row + 1}`;
  if (r.r1 === r.r2 && r.c1 === r.c2) return `${colName(r.c1)}${r.r1 + 1}`;
  if (r.r1 === 0 && r.r2 === MAX_ROWS - 1) return `${colName(r.c1)}:${colName(r.c2)}`;
  if (r.c1 === 0 && r.c2 === MAX_COLS - 1) return `${r.r1 + 1}:${r.r2 + 1}`;
  return `${colName(store.sel.active.col)}${store.sel.active.row + 1}`;
};

const r1c1Label = (store: GridStore): string => {
  const r = store.range;
  if (store.editing) return `R${store.editing.row + 1}C${store.editing.col + 1}`;
  if (r.r1 === 0 && r.r2 === MAX_ROWS - 1 && r.c1 === 0 && r.c2 === MAX_COLS - 1) return `R${store.sel.active.row + 1}C${store.sel.active.col + 1}`;
  if (r.r1 === 0 && r.r2 === MAX_ROWS - 1) return r.c1 === r.c2 ? `C${r.c1 + 1}` : `C${r.c1 + 1}:C${r.c2 + 1}`;
  if (r.c1 === 0 && r.c2 === MAX_COLS - 1) return r.r1 === r.r2 ? `R${r.r1 + 1}` : `R${r.r1 + 1}:R${r.r2 + 1}`;
  return `R${store.sel.active.row + 1}C${store.sel.active.col + 1}`;
};

/** Name box (go to / define names) and the formula bar. */
export const FormulaBar = ({ store }: { store: GridStore }) => {
  const [nameText, setNameText] = useState<string | null>(null);
  const [expanded, setExpanded] = useState(false);
  const input = useRef<HTMLTextAreaElement>(null);
  const assistKey = useRef<((e: KeyboardEvent) => boolean) | null>(null);
  const { row, col } = store.sel.active;
  const text = store.editing ? store.editing.text : store.inputText(row, col);
  const r = store.range;
  const size = r.r1 !== r.r2 || r.c1 !== r.c2 ? `${r.r2 - r.r1 + 1}R × ${r.c2 - r.c1 + 1}C` : '';

  useEffect(() => {
    if (store.editing?.inBar && document.activeElement !== input.current) input.current?.focus();
  }, [store.editing]);

  const onKey = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (assistKey.current?.(e.nativeEvent)) { e.preventDefault(); return; }
    const mod = e.ctrlKey || e.metaKey;
    if (e.key === 'Enter' && !e.altKey) {
      e.preventDefault();
      if (mod && e.shiftKey) { store.commitEdit(false, true); return; }
      if (store.commitEdit(mod)) { if (!mod) store.moveAfterEnter(e.shiftKey); focusGrid(); }
      return;
    }
    if (e.key === 'Enter' && e.altKey) {
      e.preventDefault();
      const t = e.currentTarget, p = t.selectionStart;
      store.setEditText(text.slice(0, p) + '\n' + text.slice(t.selectionEnd));
      requestAnimationFrame(() => t.setSelectionRange(p + 1, p + 1));
      return;
    }
    if (e.key === 'Tab') { e.preventDefault(); if (store.commitEdit()) { store.move(0, e.shiftKey ? -1 : 1, true); focusGrid(); } return; }
    if (e.key === 'Escape') { e.preventDefault(); store.cancelEdit(); focusGrid(); return; }
    if (e.key === 'F4') { e.preventDefault(); cycleAbsolute(e.currentTarget, store); }
  };

  return (
    <div className="formulabar">
      <div className="namebox">
        <input aria-label={T('Name box')} value={nameText ?? selectionLabel(store)} spellCheck={false}
          onFocus={(e) => { setNameText(selectionLabel(store)); requestAnimationFrame(() => e.target.select()); }}
          onChange={(e) => setNameText(e.target.value)}
          onBlur={() => setNameText(null)}
          onKeyDown={(e) => { if (e.key === 'Enter') { gotoReference(store, nameText ?? ''); setNameText(null); focusGrid(); } if (e.key === 'Escape') { setNameText(null); focusGrid(); } }} />
        <button className="ico sm" aria-label={T('Named ranges')} title={T('Name manager')} onClick={() => run(store, 'formula.names')}><ChevronDown size={14} strokeWidth={S} /></button>
      </div>
      <span className="selsize">{size}</span>
      <div className="fbtns">
        <button className="ico sm" aria-label={T('Cancel')} disabled={!store.editing} onClick={() => { store.cancelEdit(); focusGrid(); }}><X size={15} strokeWidth={S} /></button>
        <button className="ico sm" aria-label={T('Enter')} disabled={!store.editing} onClick={() => { store.commitEdit(); focusGrid(); }}><Check size={15} strokeWidth={S} /></button>
        <button className="ico sm fxbtn" aria-label={T('Insert function')} title={T('Insert function')} onClick={() => run(store, 'formula.insertFunction')}><i>fx</i></button>
      </div>
      <div className="fbox">
        <textarea ref={input} className={`finput${expanded ? ' tall' : ''}`} aria-label={T('Formula bar')} spellCheck={false} rows={1}
          value={text}
          readOnly={store.isLocked(row, col)}
          onFocus={() => { if (!store.editing) store.beginEdit(undefined, true); else if (!store.editing.inBar) { store.editing = { ...store.editing, inBar: true }; store.emit(); } }}
          onChange={(e) => { if (!store.editing) store.beginEdit(e.target.value, true); else store.setEditText(e.target.value); }}
          onKeyDown={onKey} />
        {store.editing?.inBar && <FormulaAssist store={store} anchor={{ x: 0, y: 32 }} input={input} keyRef={assistKey} />}
      </div>
      <button className="ico sm" aria-label={expanded ? T('Collapse formula bar') : T('Expand formula bar')} onClick={() => setExpanded(!expanded)}>
        <ChevronDown size={15} strokeWidth={S} style={{ transform: expanded ? 'rotate(180deg)' : undefined }} />
      </button>
    </div>
  );
};

export const focusGrid = (): void => { (document.querySelector('.gridwrap') as HTMLElement | null)?.focus({ preventScroll: true }); };
