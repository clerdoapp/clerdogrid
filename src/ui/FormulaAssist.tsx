import { useEffect, useState, type MutableRefObject, type RefObject } from 'react';
import { catalogByName, searchCatalog, type FnInfo } from '../engine/catalog';
import type { GridStore } from '../app/store';
import { formulaListSep } from '../app/locale';

/** File > Options: function autocomplete and argument tips. */
export const assistPrefs = { autocomplete: true, tips: true };

interface Context { token: string; tokenStart: number; call?: { name: string; arg: number } }

/** What the caret is in: a function name being typed, and the call (and argument) it sits inside. */
export const formulaContext = (text: string, caret: number): Context | null => {
  if (!text.startsWith('=')) return null;
  const before = text.slice(0, caret);
  let inStr = false;
  for (const ch of before) if (ch === '"') inStr = !inStr;
  if (inStr) return null;
  const tok = /([A-Za-z_][A-Za-z0-9_.]*)$/.exec(before);
  const prev = tok ? before.slice(0, before.length - tok[1]!.length) : before;
  const token = tok && /[=(,+\-*/^&<>:;{\s]$/.test(prev) && !/\$$/.test(prev) ? tok[1]! : '';
  // Walk back to find the open call. Arguments are split by the region's separator.
  const sep = formulaListSep();
  const stack: { name: string; arg: number }[] = [];
  let i = 1, q = false;
  while (i < before.length) {
    const ch = before[i]!;
    if (ch === '"') q = !q;
    else if (!q) {
      if (ch === '(') {
        const m = /([A-Za-z_][A-Za-z0-9_.]*)\s*$/.exec(before.slice(0, i));
        stack.push({ name: m ? m[1]!.toUpperCase() : '', arg: 0 });
      } else if (ch === ')') stack.pop();
      else if (ch === sep && stack.length) stack[stack.length - 1]!.arg++;
      else if (ch === '{') stack.push({ name: '{', arg: 0 });
      else if (ch === '}') stack.pop();
    }
    i++;
  }
  const call = [...stack].reverse().find((s) => s.name && s.name !== '{');
  return { token, tokenStart: caret - token.length, call };
};

export const Signature = ({ info, arg }: { info: FnInfo; arg: number }) => {
  const args = info.args;
  const repeatFrom = args.findIndex((a) => a.repeat);
  const current = repeatFrom >= 0 && arg >= repeatFrom ? repeatFrom + ((arg - repeatFrom) % (args.length - repeatFrom)) : arg;
  return (
    <span className="sig">
      <b>{info.name}</b>(
      {args.map((a, i) => (
        <span key={i}>{i ? `${formulaListSep()} ` : ''}<span className={i === current ? 'cur' : ''}>{a.optional ? `[${a.name}]` : a.name}</span></span>
      ))}
      {repeatFrom >= 0 ? ', …' : ''})
    </span>
  );
};

export const FormulaAssist = ({ store, anchor, input, keyRef }: {
  store: GridStore;
  anchor: { x: number; y: number };
  input: RefObject<HTMLTextAreaElement | HTMLInputElement | null>;
  keyRef: MutableRefObject<((e: KeyboardEvent) => boolean) | null>;
}) => {
  const [caret, setCaret] = useState(0);
  const [pick, setPick] = useState(0);
  const [closed, setClosed] = useState('');
  const text = store.editing?.text ?? '';
  useEffect(() => {
    const el = input.current;
    if (!el) return;
    const upd = () => setCaret(el.selectionStart ?? el.value.length);
    upd();
    el.addEventListener('keyup', upd); el.addEventListener('click', upd); el.addEventListener('input', upd);
    return () => { el.removeEventListener('keyup', upd); el.removeEventListener('click', upd); el.removeEventListener('input', upd); };
  }, [input, text]);
  const ctx = formulaContext(text, Math.min(caret, text.length));
  const matches = assistPrefs.autocomplete && ctx?.token && closed !== `${ctx.tokenStart}:${ctx.token}` ? searchCatalog(ctx.token, 12).filter((f) => f.name.startsWith(ctx.token.toUpperCase())) : [];
  const names = assistPrefs.autocomplete && ctx?.token ? store.wb.names.filter((n) => n.name.toLowerCase().startsWith(ctx.token.toLowerCase())).slice(0, 5) : [];
  const call = assistPrefs.tips && ctx?.call ? catalogByName(ctx.call.name) : undefined;
  const sel = Math.min(pick, Math.max(0, matches.length + names.length - 1));
  useEffect(() => setPick(0), [ctx?.token]);

  const accept = (name: string, isFn: boolean) => {
    if (!ctx || !store.editing) return;
    const ins = isFn ? `${name}(` : name;
    const next = text.slice(0, ctx.tokenStart) + ins + text.slice(ctx.tokenStart + ctx.token.length);
    store.setEditText(next);
    const pos = ctx.tokenStart + ins.length;
    requestAnimationFrame(() => { const el = input.current; if (el) { el.focus(); el.setSelectionRange(pos, pos); setCaret(pos); } });
  };

  keyRef.current = (e: KeyboardEvent) => {
    const total = matches.length + names.length;
    if (!total) return false;
    if (e.key === 'ArrowDown') { setPick((sel + 1) % total); return true; }
    if (e.key === 'ArrowUp') { setPick((sel - 1 + total) % total); return true; }
    if (e.key === 'Tab' || (e.key === 'Enter' && !e.shiftKey && !e.ctrlKey && !e.metaKey && ctx?.token && matches[sel]?.name !== ctx.token.toUpperCase())) {
      if (sel < matches.length) accept(matches[sel]!.name, true); else accept(names[sel - matches.length]!.name, false);
      return true;
    }
    if (e.key === 'Escape') { setClosed(`${ctx!.tokenStart}:${ctx!.token}`); return true; }
    return false;
  };

  if (!ctx || (!matches.length && !names.length && !call)) return null;
  return (
    <div className="assist" style={{ left: anchor.x, top: anchor.y }} onMouseDown={(e) => e.preventDefault()}>
      {call && !matches.length && !names.length && (
        <div className="sigtip"><Signature info={call} arg={ctx.call!.arg} /><div className="about">{call.args[Math.min(ctx.call!.arg, call.args.length - 1)]?.about}</div></div>
      )}
      {(matches.length > 0 || names.length > 0) && (
        <div className="fnlist" role="listbox">
          {matches.map((f, i) => (
            <button key={f.name} role="option" aria-selected={i === sel} className={i === sel ? 'on' : ''} onClick={() => accept(f.name, true)}>
              <span className="fx">fx</span><b>{f.name}</b>
            </button>
          ))}
          {names.map((n, i) => (
            <button key={n.name} role="option" aria-selected={matches.length + i === sel} className={matches.length + i === sel ? 'on' : ''} onClick={() => accept(n.name, false)}>
              <span className="fx nm">N</span><b>{n.name}</b>
            </button>
          ))}
          {sel < matches.length && <div className="fnabout">{matches[sel]!.about}</div>}
        </div>
      )}
    </div>
  );
};
