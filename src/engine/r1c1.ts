// Converts formula text between A1 references and R1C1 references relative to the formula's cell.
// Strings, quoted sheet names and structured-reference brackets are left alone; function names are
// never mistaken for references because a reference can't be followed by "(".
import { MAX_COLS, MAX_ROWS } from './types';

const colName = (col: number): string => {
  let s = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
};
const colIndex = (name: string): number => {
  let n = 0;
  for (const ch of name.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
};

/** Splits formula text into literal chunks (strings, quoted names, brackets) and plain code. */
const chunks = (text: string): { code: boolean; text: string }[] => {
  const out: { code: boolean; text: string }[] = [];
  let i = 0, start = 0;
  const flush = (end: number) => { if (end > start) out.push({ code: true, text: text.slice(start, end) }); };
  while (i < text.length) {
    const ch = text[i]!;
    if (ch === '"' || ch === "'") {
      flush(i);
      let j = i + 1;
      while (j < text.length) { if (text[j] === ch) { if (text[j + 1] === ch) { j += 2; continue; } break; } j++; }
      out.push({ code: false, text: text.slice(i, j + 1) });
      i = start = j + 1;
    } else if (ch === '[' && i > 0 && /[A-Za-z0-9_.\]]/.test(text[i - 1]!)) {
      // Structured reference brackets such as Table1[[#Headers],[Col]] (R1C1 brackets follow R or C, handled by the code path).
      const prev = text[i - 1]!;
      const isR1C1 = /[RrCc]/.test(prev) && /^\[-?\d+\]/.test(text.slice(i));
      if (isR1C1) { i++; continue; }
      flush(i);
      let depth = 0, j = i;
      for (; j < text.length; j++) { if (text[j] === '[') depth++; else if (text[j] === ']' && --depth === 0) break; }
      out.push({ code: false, text: text.slice(i, j + 1) });
      i = start = j + 1;
    } else if (ch === '{') {
      flush(i);
      let j = i;
      let inStr = false;
      for (; j < text.length; j++) { if (text[j] === '"') inStr = !inStr; else if (text[j] === '}' && !inStr) break; }
      out.push({ code: false, text: text.slice(i, j + 1) });
      i = start = j + 1;
    } else i++;
  }
  flush(text.length);
  return out;
};

const A1_TOKEN = /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?::(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7}))?|(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})|(\$?)(\d{1,7}):(\$?)(\d{1,7})/y;

const relPart = (letter: 'R' | 'C', abs: boolean, value: number, base: number): string =>
  abs ? `${letter}${value + 1}` : value === base ? letter : `${letter}[${value - base}]`;

/** True when the characters around a match make it a reference rather than part of a name or number. */
const standalone = (code: string, start: number, end: number): boolean => {
  const before = code[start - 1] ?? '';
  const after = code[end] ?? '';
  if (/[A-Za-z0-9_.\\$]/.test(before)) return false;
  if (/[A-Za-z0-9_.(]/.test(after)) return false;
  return true;
};

/** A1 formula text to R1C1, relative to the cell at (row, col), 0-based. */
export const toR1C1 = (formula: string, row: number, col: number): string =>
  chunks(formula).map((c) => {
    if (!c.code) return c.text;
    let out = '';
    let i = 0;
    const s = c.text;
    while (i < s.length) {
      A1_TOKEN.lastIndex = i;
      const m = /[A-Za-z$\d]/.test(s[i]!) && (i === 0 || !/[A-Za-z0-9_.$]/.test(s[i - 1]!)) ? A1_TOKEN.exec(s) : null;
      if (m && standalone(s, i, i + m[0].length)) {
        if (m[2] !== undefined) {
          const c1 = colIndex(m[2]), r1 = +m[4]! - 1;
          if (c1 < MAX_COLS && r1 < MAX_ROWS && r1 >= 0) {
            let t = relPart('R', !!m[3], r1, row) + relPart('C', !!m[1], c1, col);
            if (m[6] !== undefined) t += ':' + relPart('R', !!m[7], +m[8]! - 1, row) + relPart('C', !!m[5], colIndex(m[6]), col);
            out += t; i += m[0].length; continue;
          }
        } else if (m[10] !== undefined) {
          const a = colIndex(m[10]), b = colIndex(m[12]!);
          if (a < MAX_COLS && b < MAX_COLS) {
            const pa = relPart('C', !!m[9], a, col), pb = relPart('C', !!m[11], b, col);
            out += pa === pb ? pa : `${pa}:${pb}`; i += m[0].length; continue;
          }
        } else if (m[14] !== undefined) {
          const a = +m[14] - 1, b = +m[16]! - 1;
          if (a >= 0 && b >= 0 && a < MAX_ROWS && b < MAX_ROWS) {
            const pa = relPart('R', !!m[13], a, row), pb = relPart('R', !!m[15], b, row);
            out += pa === pb ? pa : `${pa}:${pb}`; i += m[0].length; continue;
          }
        }
      }
      // Copy a whole word at once so the middle of a name is never read as a reference.
      const w = /^[A-Za-z0-9_.\\]+/.exec(s.slice(i));
      const n = w ? w[0].length : 1;
      out += s.slice(i, i + n);
      i += n;
    }
    return out;
  }).join('');

const R1C1_PART = /([Rr])(\[-?\d{1,7}\]|\d{1,7})?([Cc])(\[-?\d{1,5}\]|\d{1,5})?|([Rr])(\[-?\d{1,7}\]|\d{1,7})?|([Cc])(\[-?\d{1,5}\]|\d{1,5})?/y;

interface Part { row?: { abs: boolean; v: number }; col?: { abs: boolean; v: number } }

const readPart = (s: string, i: number, row: number, col: number): { part: Part; end: number } | null | 'ref' => {
  R1C1_PART.lastIndex = i;
  const m = R1C1_PART.exec(s);
  if (!m || !m[0]) return null;
  const val = (spec: string | undefined, base: number): { abs: boolean; v: number } => {
    if (spec === undefined) return { abs: false, v: base };
    if (spec.startsWith('[')) return { abs: false, v: base + Number(spec.slice(1, -1)) };
    return { abs: true, v: Number(spec) - 1 };
  };
  const part: Part = {};
  if (m[1]) { part.row = val(m[2], row); part.col = val(m[4], col); }
  else if (m[5]) part.row = val(m[6], row);
  else part.col = val(m[8], col);
  if ((part.row && (part.row.v < 0 || part.row.v >= MAX_ROWS)) || (part.col && (part.col.v < 0 || part.col.v >= MAX_COLS))) return 'ref';
  return { part, end: i + m[0].length };
};

const a1Of = (p: Part): string =>
  (p.col ? `${p.col.abs ? '$' : ''}${colName(p.col.v)}` : '') + (p.row ? `${p.row.abs ? '$' : ''}${p.row.v + 1}` : '');

/** R1C1 formula text to A1, relative to the cell at (row, col), 0-based. References that fall off the sheet become #REF!. */
export const fromR1C1 = (formula: string, row: number, col: number): string =>
  chunks(formula).map((c) => {
    if (!c.code) return c.text;
    const s = c.text;
    let out = '';
    let i = 0;
    while (i < s.length) {
      const starts = /[RrCc]/.test(s[i]!) && (i === 0 || !/[A-Za-z0-9_.$\]]/.test(s[i - 1]!));
      const a = starts ? readPart(s, i, row, col) : null;
      if (a && (a === 'ref' || standalone(s, i, a.end) || s[a.end] === ':')) {
        if (a === 'ref') { R1C1_PART.lastIndex = i; const m = R1C1_PART.exec(s)!; out += '#REF!'; i += m[0].length; continue; }
        let end = a.end;
        let text: string;
        const b = s[end] === ':' ? readPart(s, end + 1, row, col) : null;
        if (b === 'ref') { R1C1_PART.lastIndex = end + 1; const m = R1C1_PART.exec(s)!; out += '#REF!'; i = end + 1 + m[0].length; continue; }
        if (b && standalone(s, end + 1, b.end)) {
          end = b.end;
          const p = a.part, q = b.part;
          if (p.row && p.col && q.row && q.col) text = `${a1Of(p)}:${a1Of(q)}`;
          else if (p.row && !p.col && q.row && !q.col) text = `${a1Of(p)}:${a1Of(q)}`;
          else if (p.col && !p.row && q.col && !q.row) text = `${a1Of(p)}:${a1Of(q)}`;
          else { out += s.slice(i, end); i = end; continue; }
        } else if (standalone(s, i, a.end)) {
          const p = a.part;
          // A lone R or C part is a whole row or column.
          text = p.row && p.col ? a1Of(p) : p.row ? `${a1Of(p)}:${a1Of(p)}` : `${a1Of(p)}:${a1Of(p)}`;
        } else { out += s[i]; i++; continue; }
        out += text;
        i = end;
        continue;
      }
      const w = /^[A-Za-z0-9_.\\]+/.exec(s.slice(i));
      const n = w ? w[0].length : 1;
      out += s.slice(i, i + n);
      i += n;
    }
    return out;
  }).join('');

/** "R1C1" label for a cell (absolute), as the name box shows it in R1C1 mode. */
export const r1c1Address = (row: number, col: number): string => `R${row + 1}C${col + 1}`;
