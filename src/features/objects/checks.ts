// Review checks: workbook statistics, accessibility, performance cleanup and spelling helpers.
import { keyCol, keyRow, addr, type Sheet, type Workbook } from '../../model/workbook';
import type { GridStore } from '../../app/store';

// ---------- Contrast (WCAG 2 relative luminance) ----------

const channel = (v: number): number => { const c = v / 255; return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4; };
export const luminance = (hex: string): number => {
  const m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})/i.exec(hex);
  if (!m) return 1;
  return 0.2126 * channel(parseInt(m[1]!, 16)) + 0.7152 * channel(parseInt(m[2]!, 16)) + 0.0722 * channel(parseInt(m[3]!, 16));
};
export const contrastRatio = (a: string, b: string): number => {
  const la = luminance(a), lb = luminance(b);
  return (Math.max(la, lb) + 0.05) / (Math.min(la, lb) + 0.05);
};

// ---------- Statistics ----------

export interface SheetStats { name: string; endCell: string; cells: number; formulas: number; tables: number; charts: number; pictures: number; shapes: number; comments: number; notes: number }

export const sheetStats = (wb: Workbook, s: Sheet): SheetStats => {
  let rows = -1, cols = -1, cells = 0, formulas = 0;
  for (const [k, c] of s.cells) {
    if (c.v === null && !c.f) continue;
    cells++;
    if (c.f) formulas++;
    rows = Math.max(rows, keyRow(k)); cols = Math.max(cols, keyCol(k));
  }
  void wb;
  return {
    name: s.name, endCell: rows < 0 ? 'A1' : addr(rows, cols), cells, formulas, tables: s.tables.length, charts: s.charts.length,
    pictures: s.pictures.length, shapes: s.shapes?.length ?? 0, comments: s.comments?.size ?? 0, notes: s.notes.size,
  };
};

// ---------- Accessibility ----------

export interface Issue { kind: string; sheet: number; where: string; text: string; tip: string; go?: { row: number; col: number }; object?: { type: 'chart' | 'picture' | 'shape'; id: string } }

export const accessibilityIssues = (wb: Workbook): Issue[] => {
  const out: Issue[] = [];
  wb.sheets.forEach((s, si) => {
    const at = (a: { row: number; col: number }) => addr(a.row, a.col);
    for (const c of s.charts) if (!c.alt?.trim()) out.push({ kind: 'Missing alt text', sheet: si, where: c.title || 'Chart', text: 'Chart has no alternative text.', tip: 'Describe what the chart shows.', go: c.anchor, object: { type: 'chart', id: c.id } });
    for (const p of s.pictures) if (!p.alt?.trim()) out.push({ kind: 'Missing alt text', sheet: si, where: `Picture at ${at(p.anchor)}`, text: 'Picture has no alternative text.', tip: 'Describe the picture in a sentence.', go: p.anchor, object: { type: 'picture', id: p.id } });
    for (const sh of s.shapes ?? []) if (!sh.alt?.trim() && !sh.text?.trim() && sh.kind !== 'line' && sh.kind !== 'arrowLine') out.push({ kind: 'Missing alt text', sheet: si, where: `Shape at ${at(sh.anchor)}`, text: 'Shape has no text or alternative text.', tip: 'Add alt text, or mark it decorative by typing a short label.', go: sh.anchor, object: { type: 'shape', id: sh.id } });
    for (const m of s.merges) out.push({ kind: 'Merged cells', sheet: si, where: `${addr(m.r1, m.c1)}:${addr(m.r2, m.c2)}`, text: 'Merged cells are hard to move through with a screen reader.', tip: 'Unmerge and use Center across selection instead.', go: { row: m.r1, col: m.c1 } });
    if (/^Sheet\d+$/i.test(s.name) && s.cells.size) out.push({ kind: 'Default sheet name', sheet: si, where: s.name, text: 'The sheet name does not say what it holds.', tip: 'Rename the sheet.' });
    if (s.state !== 'visible') out.push({ kind: 'Hidden sheet', sheet: si, where: s.name, text: 'Hidden sheets are easy to miss.', tip: 'Unhide it, or delete it if it is not needed.' });
    for (const t of s.tables) if (!t.headerRow) out.push({ kind: 'Table without headers', sheet: si, where: t.name, text: 'The table has no header row.', tip: 'Turn on the header row.', go: { row: t.range.r1, col: t.range.c1 } });
    let n = 0;
    for (const [k, c] of s.cells) {
      if (c.v === null || c.v === '' || n > 200) continue;
      const st = wb.styles[c.s ?? 0];
      if (!st?.font?.color && !st?.fill?.color) continue;
      const fg = st.font?.color ?? '#000000', bg = st.fill?.color ?? '#FFFFFF';
      const ratio = contrastRatio(fg, bg);
      if (ratio < 4.5) {
        n++;
        out.push({ kind: 'Hard-to-read text', sheet: si, where: addr(keyRow(k), keyCol(k)), text: `Text contrast is ${ratio.toFixed(1)}:1 (needs 4.5:1).`, tip: 'Use a darker text color or a lighter fill.', go: { row: keyRow(k), col: keyCol(k) } });
      }
    }
  });
  return out;
};

// ---------- Performance ----------

/** Cells and row/column formats beyond the last cell with data. Returns how many it would remove and a byte estimate. */
export const excessFormatting = (s: Sheet): { cells: number[]; rows: number[]; cols: number[]; bytes: number } => {
  let lastR = -1, lastC = -1;
  for (const [k, c] of s.cells) if (c.v !== null || c.f) { lastR = Math.max(lastR, keyRow(k)); lastC = Math.max(lastC, keyCol(k)); }
  const cells: number[] = [];
  for (const [k, c] of s.cells) if (c.v === null && !c.f && (keyRow(k) > lastR || keyCol(k) > lastC)) cells.push(k);
  const rows = [...s.rows].filter(([r, info]) => r > lastR && info.s && info.height === undefined && !info.hidden).map(([r]) => r);
  const cols = [...s.cols].filter(([c, info]) => c > lastC && info.s && info.width === undefined && !info.hidden).map(([c]) => c);
  // Roughly what each costs in the saved sheet XML (<c r="AB123" s="5"/> and row/col elements), before compression.
  return { cells, rows, cols, bytes: cells.length * 22 + rows.length * 34 + cols.length * 46 };
};

export const cleanExcessFormatting = (store: GridStore): { removed: number; bytes: number } => {
  const x = excessFormatting(store.sheet);
  const removed = x.cells.length + x.rows.length + x.cols.length;
  if (!removed) return { removed: 0, bytes: 0 };
  store.change('Clean extra formatting', 'sheet', () => {
    const s = store.sheet;
    for (const k of x.cells) s.cells.delete(k);
    for (const r of x.rows) { const info = s.rows.get(r)!; delete info.s; if (!Object.keys(info).length) s.rows.delete(r); }
    for (const c of x.cols) { const info = s.cols.get(c)!; delete info.s; if (!Object.keys(info).length) s.cols.delete(c); }
    store.engine.sheetsChanged();
  });
  return { removed, bytes: x.bytes };
};

// ---------- Spelling ----------

/** Words in a text cell worth checking: skips URLs, emails, ALL-CAPS, words with digits, and single letters. */
export const wordsIn = (text: string): { word: string; start: number }[] => {
  const out: { word: string; start: number }[] = [];
  const skip: [number, number][] = [];
  for (const m of text.matchAll(/(?:https?:\/\/|www\.)\S+|\S+@\S+\.\S+/gi)) skip.push([m.index!, m.index! + m[0].length]);
  for (const m of text.matchAll(/[A-Za-zÀ-ɏ][A-Za-zÀ-ɏ'’]*[A-Za-zÀ-ɏ]|[A-Za-z]/g)) {
    const w = m[0].replace(/’/g, "'"), i = m.index!;
    if (skip.some(([a, b]) => i >= a && i < b)) continue;
    if (w.length < 2 || (w === w.toUpperCase() && /[A-Z]/.test(w))) continue;
    if (/\d/.test(text.slice(Math.max(0, i - 1), i + w.length + 1))) continue;
    out.push({ word: w, start: i });
  }
  return out;
};

export interface SpellCell { sheet: number; row: number; col: number; text: string }

/** Text cells to check, in reading order from the active cell (formulas and numbers are skipped). */
export const spellCells = (store: GridStore): SpellCell[] => {
  const s = store.sheet;
  const list: SpellCell[] = [];
  for (const [k, c] of s.cells) if (!c.f && typeof c.v === 'string' && c.v.trim()) list.push({ sheet: store.sheetIndex, row: keyRow(k), col: keyCol(k), text: c.v });
  list.sort((a, b) => a.row - b.row || a.col - b.col);
  const { row, col } = store.sel.active;
  const i = list.findIndex((c) => c.row > row || (c.row === row && c.col >= col));
  return i <= 0 ? list : [...list.slice(i), ...list.slice(0, i)];
};
