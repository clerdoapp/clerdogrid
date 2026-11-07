import { useLayoutEffect, useMemo, useRef, type KeyboardEvent, type ReactNode } from 'react';

const KEYWORDS = new Set(('break case catch class const continue debugger default delete do else export extends finally for function if import in instanceof let new of return '
  + 'super switch this throw try typeof var void while with yield async await static get set').split(' '));
const LITERALS = new Set(['true', 'false', 'null', 'undefined', 'NaN', 'Infinity']);

const TOKEN = /(\/\/[^\n]*|\/\*[\s\S]*?(?:\*\/|$))|("(?:[^"\\\n]|\\.)*"?|'(?:[^'\\\n]|\\.)*'?|`(?:[^`\\]|\\[\s\S])*`?)|(\b\d[\d_]*(?:\.\d+)?(?:e[+-]?\d+)?\b|\b0x[\da-f]+\b)|([A-Za-z_$][\w$]*)/gi;

/** Splits code into coloured spans: comments, strings, numbers, keywords. */
export const highlight = (code: string): ReactNode[] => {
  const out: ReactNode[] = [];
  let last = 0, i = 0;
  for (const m of code.matchAll(TOKEN)) {
    const at = m.index ?? 0;
    if (at > last) out.push(code.slice(last, at));
    const [text, comment, str, num, word] = m;
    const cls = comment ? 'c' : str ? 's' : num ? 'n' : word && KEYWORDS.has(word) ? 'k' : word && LITERALS.has(word) ? 'l' : '';
    out.push(cls ? <span key={i++} className={`tk-${cls}`}>{text}</span> : text);
    last = at + text.length;
  }
  if (last < code.length) out.push(code.slice(last));
  return out;
};

const INDENT = '  ';

/** A plain textarea with line numbers and a coloured layer underneath; no editor library. */
export const CodeEditor = ({ value, onChange, errorLine, onRun, onSave, label = 'Script code', readOnly }: {
  value: string; onChange: (v: string) => void; errorLine?: number; onRun?: () => void; onSave?: () => void; label?: string; readOnly?: boolean;
}) => {
  const ta = useRef<HTMLTextAreaElement>(null);
  const layer = useRef<HTMLPreElement>(null);
  const gutter = useRef<HTMLDivElement>(null);
  const pendingSel = useRef<[number, number] | null>(null);
  const lines = useMemo(() => value.split('\n').length, [value]);
  const colored = useMemo(() => highlight(value), [value]);

  useLayoutEffect(() => {
    const p = pendingSel.current;
    if (p && ta.current) { ta.current.setSelectionRange(p[0], p[1]); pendingSel.current = null; }
  });

  const sync = () => {
    const t = ta.current;
    if (!t) return;
    if (layer.current) { layer.current.scrollTop = t.scrollTop; layer.current.scrollLeft = t.scrollLeft; }
    if (gutter.current) gutter.current.scrollTop = t.scrollTop;
  };

  const replace = (start: number, end: number, text: string, selStart: number, selEnd = selStart) => {
    onChange(value.slice(0, start) + text + value.slice(end));
    pendingSel.current = [selStart, selEnd];
  };

  const onKey = (e: KeyboardEvent<HTMLTextAreaElement>) => {
    const t = e.currentTarget;
    const { selectionStart: s, selectionEnd: en } = t;
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); onRun?.(); return; }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); e.stopPropagation(); onSave?.(); return; }
    if (readOnly) return;
    if (e.key === 'Tab') {
      e.preventDefault();
      const lineStart = value.lastIndexOf('\n', s - 1) + 1;
      if (s === en && !e.shiftKey) { replace(s, en, INDENT, s + INDENT.length); return; }
      const block = value.slice(lineStart, en);
      const changed = e.shiftKey ? block.replace(/^ {1,2}/gm, '') : block.replace(/^/gm, INDENT);
      const firstDelta = e.shiftKey ? -(block.match(/^ {1,2}/)?.[0].length ?? 0) : INDENT.length;
      replace(lineStart, en, changed, Math.max(lineStart, s + firstDelta), lineStart + changed.length);
      return;
    }
    if (e.key === 'Enter' && !e.shiftKey && !e.altKey) {
      e.preventDefault();
      e.stopPropagation();
      const lineStart = value.lastIndexOf('\n', s - 1) + 1;
      const indent = /^[ \t]*/.exec(value.slice(lineStart, s))![0];
      const before = value.slice(0, s).trimEnd();
      const extra = /[{[(]$/.test(before) ? INDENT : '';
      const closing = /^[ \t]*[}\])]/.test(value.slice(en));
      if (extra && closing) {
        const text = `\n${indent}${extra}\n${indent}`;
        replace(s, en, text, s + indent.length + extra.length + 1);
      } else replace(s, en, `\n${indent}${extra}`, s + 1 + indent.length + extra.length);
      return;
    }
    if (e.key === 'Escape') e.stopPropagation();
  };

  return (
    <div className="au-editor">
      <div className="au-gutter" ref={gutter} aria-hidden="true">
        {Array.from({ length: lines }, (_, i) => <div key={i} className={errorLine === i + 1 ? 'err' : undefined}>{i + 1}</div>)}
      </div>
      <div className="au-code">
        <pre className="au-layer" ref={layer} aria-hidden="true">{colored}{'\n'}</pre>
        <textarea ref={ta} value={value} spellCheck={false} autoCapitalize="off" autoComplete="off" autoCorrect="off" wrap="off" aria-label={label} readOnly={readOnly}
          onChange={(e) => onChange(e.target.value)} onKeyDown={onKey} onScroll={sync} />
      </div>
    </div>
  );
};
