import { adjustForMove, shiftFormula } from '../../engine/refactor';
import { CellError, MAX_COLS, MAX_ROWS, type Scalar } from '../../engine/types';
import { formatValue } from '../../formats/numfmt';
import { cellKey, internStyle, keyCol, keyRow, type Cell, type CellRange, type CellStyle } from '../../model/workbook';
import type { GridStore } from '../store';
import { N, T } from '../../ui/i18n';

interface ClipCell { v: Scalar; f?: string; style: CellStyle; display: string; /** Source cell, when copied rows or columns were skipped. */ at?: [number, number] }
interface Clip {
  sheet: number;
  range: CellRange;
  cells: (ClipCell | null)[][];
  colWidths: (number | undefined)[];
  merges: CellRange[];
  cut: boolean;
  text: string;
}

let clip: Clip | null = null;

export const displayText = (store: GridStore, row: number, col: number, sheet = store.sheetIndex): string => {
  const v = store.engine.getValue(sheet, row, col);
  if (v === null) return '';
  if (v instanceof CellError) return v.code;
  return formatValue(v, store.style(row, col, sheet).numFmt ?? 'General', { date1904: store.wb.date1904 }).text;
};

const escapeHtml = (s: string) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);

const styleCss = (s: CellStyle): string => {
  const out: string[] = [];
  if (s.font?.bold) out.push('font-weight:bold');
  if (s.font?.italic) out.push('font-style:italic');
  if (s.font?.underline) out.push('text-decoration:underline');
  if (s.font?.color) out.push(`color:${s.font.color}`);
  if (s.font?.name) out.push(`font-family:'${s.font.name}'`);
  if (s.font?.size) out.push(`font-size:${s.font.size}pt`);
  if (s.fill?.color) out.push(`background:${s.fill.color}`);
  if (s.align?.h && s.align.h !== 'general') out.push(`text-align:${s.align.h}`);
  if (s.numFmt && s.numFmt !== 'General') out.push(`mso-number-format:"${s.numFmt.replace(/"/g, '\\"')}"`);
  return out.join(';');
};

/** Copies (or cuts) the selection to the internal clipboard and the system clipboard as text and HTML. */
export const copySelection = async (store: GridStore, cut = false): Promise<void> => {
  // Several blocks copy together when they line up (as after selecting visible cells only).
  const ranges = store.sel.ranges.map((x) => clampRange(store, x));
  let rowList: number[] | null = null, colList: number[] | null = null;
  if (ranges.length > 1) {
    const sameCols = ranges.every((x) => x.c1 === ranges[0]!.c1 && x.c2 === ranges[0]!.c2);
    const sameRows = ranges.every((x) => x.r1 === ranges[0]!.r1 && x.r2 === ranges[0]!.r2);
    const lines = (a: [number, number][]) => [...new Set(a.flatMap(([p, q]) => Array.from({ length: q - p + 1 }, (_, i) => p + i)))].sort((x, y) => x - y);
    const rowsSet = lines(ranges.map((x) => [x.r1, x.r2])), colsSet = lines(ranges.map((x) => [x.c1, x.c2]));
    // A grid of blocks (rows x columns), like visible cells around hidden lines.
    const grid = ranges.length === new Set(ranges.map((x) => `${x.r1},${x.r2}`)).size * new Set(ranges.map((x) => `${x.c1},${x.c2}`)).size;
    if (cut || !(sameCols || sameRows || grid)) { store.notify(T('That command can\'t be used on multiple selections.')); return; }
    rowList = rowsSet; colList = colsSet;
  }
  const r = rowList && colList ? { r1: rowList[0]!, c1: colList[0]!, r2: rowList[rowList.length - 1]!, c2: colList[colList.length - 1]! } : ranges[ranges.length - 1]!;
  const rowsToCopy = rowList ?? Array.from({ length: r.r2 - r.r1 + 1 }, (_, i) => r.r1 + i);
  const colsToCopy = colList ?? Array.from({ length: r.c2 - r.c1 + 1 }, (_, i) => r.c1 + i);
  const cells: (ClipCell | null)[][] = [];
  const lines: string[] = [];
  const html: string[] = ['<table>'];
  for (const row of rowsToCopy) {
    if (store.sheet.rows.get(row)?.hidden && !cut) continue;
    const line: (ClipCell | null)[] = [];
    const texts: string[] = [];
    html.push('<tr>');
    for (const col of colsToCopy) {
      const c = store.cell(row, col);
      const display = displayText(store, row, col);
      const style = store.style(row, col);
      line.push(c || style !== store.wb.styles[0] ? { v: c?.v ?? null, f: c?.f, style, display, at: [row, col] } : null);
      texts.push(/[\t\n"]/.test(display) ? `"${display.replace(/"/g, '""')}"` : display);
      html.push(`<td style="${escapeHtml(styleCss(style))}">${escapeHtml(display)}</td>`);
    }
    html.push('</tr>');
    cells.push(line);
    lines.push(texts.join('\t'));
  }
  html.push('</table>');
  const text = lines.join('\r\n') + '\r\n';
  clip = {
    sheet: store.sheetIndex, range: r, cells, cut, text,
    colWidths: colsToCopy.map((c) => store.sheet.cols.get(c)?.width),
    merges: rowList ? [] : store.sheet.merges.filter((m) => m.r1 >= r.r1 && m.r2 <= r.r2 && m.c1 >= r.c1 && m.c2 <= r.c2),
  };
  store.marquee = { sheet: store.sheetIndex, range: r, cut };
  store.emit();
  try {
    if (navigator.clipboard && 'write' in navigator.clipboard && typeof ClipboardItem !== 'undefined') {
      await navigator.clipboard.write([new ClipboardItem({ 'text/plain': new Blob([text], { type: 'text/plain' }), 'text/html': new Blob([html.join('')], { type: 'text/html' }) })]);
    } else await navigator.clipboard?.writeText(text);
  } catch { /* the internal clipboard still works inside the app */ }
};

const clampRange = (store: GridStore, r: CellRange): CellRange => {
  const b = store.engine.usedBounds(store.sheetIndex);
  return { r1: r.r1, c1: r.c1, r2: r.r2 === MAX_ROWS - 1 ? Math.max(r.r1, b.rows) : r.r2, c2: r.c2 === MAX_COLS - 1 ? Math.max(r.c1, b.cols) : r.c2 };
};

export type PasteMode = 'all' | 'values' | 'formulas' | 'formats' | 'valuesAndFormats' | 'formulasAndFormats' | 'noBorders' | 'columnWidths' | 'link';
export interface PasteOptions { mode?: PasteMode; transpose?: boolean; skipBlanks?: boolean; operation?: 'none' | 'add' | 'subtract' | 'multiply' | 'divide' }

/** Pastes: the app's own clipboard when it matches the system clipboard, otherwise text/HTML from other apps. */
export const paste = async (store: GridStore, opts: PasteOptions = {}, external?: { text?: string; html?: string }): Promise<void> => {
  let text = external?.text, html = external?.html;
  if (!external) {
    try {
      if (navigator.clipboard && 'read' in navigator.clipboard) {
        const items = await navigator.clipboard.read();
        for (const it of items) {
          if (it.types.includes('text/html')) html = await (await it.getType('text/html')).text();
          if (it.types.includes('text/plain')) text = await (await it.getType('text/plain')).text();
        }
      } else text = await (navigator as Navigator).clipboard?.readText();
    } catch { /* permission denied: fall back to the internal clipboard */ }
  }
  const norm = (s?: string) => (s ?? '').replace(/\r\n/g, '\n').trimEnd();
  if (clip && (text === undefined || norm(text) === norm(clip.text))) { pasteInternal(store, clip, opts); return; }
  if (html && /<table/i.test(html)) { pasteGrid(store, parseHtmlTable(html)); return; }
  if (text !== undefined) pasteGrid(store, parseTsv(text).map((r) => r.map((t) => ({ text: t }))));
};

const pasteInternal = (store: GridStore, c: Clip, opts: PasteOptions): void => {
  const mode = opts.mode ?? 'all';
  const src = c.cells;
  const h = src.length, w = src[0]?.length ?? 0;
  if (!h || !w) return;
  const sel = store.range;
  // Fill a larger selection with whole copies, as spreadsheets do when it divides evenly.
  const tH = opts.transpose ? w : h, tW = opts.transpose ? h : w;
  const reps = (sel.r2 - sel.r1 + 1) % tH === 0 && (sel.c2 - sel.c1 + 1) % tW === 0 ? [(sel.r2 - sel.r1 + 1) / tH, (sel.c2 - sel.c1 + 1) / tW] : [1, 1];
  const target: CellRange = { r1: sel.r1, c1: sel.c1, r2: sel.r1 + tH * reps[0]! - 1, c2: sel.c1 + tW * reps[1]! - 1 };
  if (target.r2 >= MAX_ROWS || target.c2 >= MAX_COLS) { store.notify(T('The paste area doesn\'t fit on the sheet.')); return; }
  if (c.cut && c.sheet !== store.sheetIndex && store.wb.sheets[c.sheet] === undefined) return;
  const scope = [{ area: target }];
  if (c.cut) scope.push({ area: c.range, sheet: c.sheet } as never);
  store.change(c.cut ? N('Move') : N('Paste'), c.cut ? 'book' : scope, () => {
    const sheet = store.sheetIndex;
    if (c.cut) {
      // Moving: clear the source first, then write, then fix references to the moved block everywhere.
      for (let row = c.range.r1; row <= c.range.r2; row++) for (let col = c.range.c1; col <= c.range.c2; col++) store.engine.clearCell(c.sheet, row, col);
      store.wb.sheets[c.sheet]!.merges = store.wb.sheets[c.sheet]!.merges.filter((m) => !(m.r1 >= c.range.r1 && m.r2 <= c.range.r2 && m.c1 >= c.range.c1 && m.c2 <= c.range.c2));
    }
    for (let rr = 0; rr < reps[0]!; rr++) for (let rc = 0; rc < reps[1]!; rc++) {
      for (let i = 0; i < h; i++) for (let j = 0; j < w; j++) {
        const cc = src[i]![j]!;
        const di = opts.transpose ? j : i, dj = opts.transpose ? i : j;
        const row = target.r1 + rr * tH + di, col = target.c1 + rc * tW + dj;
        if (opts.skipBlanks && (!cc || (cc.v === null && !cc.f))) continue;
        writeCell(store, sheet, row, col, cc, c, i, j, mode, opts.operation ?? 'none');
      }
    }
    if ((mode === 'all' || mode === 'columnWidths') && !opts.transpose) {
      c.colWidths.forEach((wd, j) => { if (wd !== undefined) { const ci = store.sheet.cols.get(target.c1 + j) ?? {}; ci.width = wd; store.sheet.cols.set(target.c1 + j, ci); } });
    }
    if (mode === 'all' && !opts.transpose) for (const m of c.merges) store.sheet.merges.push({ r1: m.r1 - c.range.r1 + target.r1, c1: m.c1 - c.range.c1 + target.c1, r2: m.r2 - c.range.r1 + target.r1, c2: m.c2 - c.range.c1 + target.c1 });
    if (c.cut) {
      const from = { sheet: store.wb.sheets[c.sheet]!.name, ...c.range };
      const to = { sheet: store.sheet.name, row: target.r1, col: target.c1 };
      for (const s of store.wb.sheets) for (const cell of s.cells.values()) if (cell.f) cell.f = adjustForMove(cell.f, s.name, from, to);
      store.engine.sheetsChanged();
      store.engine.recalc(true);
    }
  });
  if (c.cut) { clip = null; store.marquee = null; }
  store.selectRange(target);
};

const writeCell = (store: GridStore, sheet: number, row: number, col: number, cc: ClipCell | null, c: Clip, i: number, j: number, mode: PasteMode, op: string): void => {
  const s = store.wb.sheets[sheet]!;
  const key = cellKey(row, col);
  const withStyle = mode === 'all' || mode === 'formats' || mode === 'valuesAndFormats' || mode === 'formulasAndFormats' || mode === 'noBorders';
  const withContent = mode !== 'formats' && mode !== 'columnWidths';
  if (mode === 'columnWidths') return;
  if (mode === 'link') {
    const ref = `${c.sheet !== sheet ? `'${store.wb.sheets[c.sheet]!.name.replace(/'/g, "''")}'!` : ''}$${colLetters(cc?.at?.[1] ?? c.range.c1 + j)}$${(cc?.at?.[0] ?? c.range.r1 + i) + 1}`;
    store.engine.setFormula(sheet, row, col, ref);
    return;
  }
  if (withContent) {
    if (!cc || (cc.v === null && !cc.f)) { if (op === 'none') store.engine.setValue(sheet, row, col, null); }
    else if (cc.f && (mode === 'all' || mode === 'formulas' || mode === 'formulasAndFormats' || mode === 'noBorders') && op === 'none') {
      const f = c.cut ? cc.f : shiftFormula(cc.f, row - (cc.at?.[0] ?? c.range.r1 + i), col - (cc.at?.[1] ?? c.range.c1 + j));
      store.engine.setFormula(sheet, row, col, f);
    } else if (op !== 'none' && typeof cc.v === 'number') {
      const cur = store.engine.getValue(sheet, row, col);
      const base = typeof cur === 'number' ? cur : cur === null ? 0 : NaN;
      const v = op === 'add' ? base + cc.v : op === 'subtract' ? base - cc.v : op === 'multiply' ? base * cc.v : cc.v === 0 ? NaN : base / cc.v;
      store.engine.setValue(sheet, row, col, Number.isFinite(v) ? v : new CellError(op === 'divide' ? '#DIV/0!' : '#VALUE!'));
    } else store.engine.setValue(sheet, row, col, cc.f ? store.engine.getValue(c.sheet, cc.at?.[0] ?? c.range.r1 + i, cc.at?.[1] ?? c.range.c1 + j) : cc.v);
  }
  if (withStyle) {
    let style = cc?.style ?? {};
    if (mode === 'noBorders') style = { ...style, border: undefined };
    const id = internStyle(store.wb, style);
    const cell = s.cells.get(key);
    if (cell) cell.s = id; else if (id) s.cells.set(key, { v: null, s: id } as Cell);
  }
};

const colLetters = (c: number): string => { let s = ''; let n = c + 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

/** Tab-separated text with quoted fields, as spreadsheets put on the clipboard. */
export const parseTsv = (text: string): string[][] => {
  const rows: string[][] = [];
  let row: string[] = [], field = '', q = false;
  const t = text.replace(/\r\n?/g, '\n').replace(/\n$/, '');
  for (let i = 0; i < t.length; i++) {
    const ch = t[i]!;
    if (q) { if (ch === '"') { if (t[i + 1] === '"') { field += '"'; i++; } else q = false; } else field += ch; continue; }
    if (ch === '"' && field === '') { q = true; continue; }
    if (ch === '\t') { row.push(field); field = ''; continue; }
    if (ch === '\n') { row.push(field); rows.push(row); row = []; field = ''; continue; }
    field += ch;
  }
  row.push(field); rows.push(row);
  return rows;
};

interface ExtCell { text: string; style?: CellStyle }

const parseHtmlTable = (html: string): ExtCell[][] => {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const table = doc.querySelector('table');
  if (!table) return [];
  const out: ExtCell[][] = [];
  for (const tr of Array.from(table.querySelectorAll('tr'))) {
    const row: ExtCell[] = [];
    for (const td of Array.from(tr.querySelectorAll('td,th'))) {
      const el = td as HTMLElement;
      const st: CellStyle = {};
      const css = el.style;
      if (css.fontWeight === 'bold' || +css.fontWeight >= 600 || el.querySelector('b,strong') || td.tagName === 'TH') st.font = { ...st.font, bold: true };
      if (css.fontStyle === 'italic' || el.querySelector('i,em')) st.font = { ...st.font, italic: true };
      if (css.color) st.font = { ...st.font, color: cssHex(css.color) };
      if (css.backgroundColor || css.background) { const c = cssHex(css.backgroundColor || css.background); if (c) st.fill = { color: c }; }
      row.push({ text: (el.innerText ?? el.textContent ?? '').replace(/ /g, ' ').trim(), style: Object.keys(st).length ? st : undefined });
      const span = +(el.getAttribute('colspan') ?? 1);
      for (let k = 1; k < span; k++) row.push({ text: '' });
    }
    out.push(row);
  }
  return out;
};

const cssHex = (c: string): string | undefined => {
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(c);
  if (m) return `#${[m[1], m[2], m[3]].map((x) => (+x!).toString(16).padStart(2, '0')).join('')}`;
  return /^#[0-9a-f]{6}$/i.test(c) ? c : undefined;
};

const pasteGrid = (store: GridStore, rows: ExtCell[][]): void => {
  if (!rows.length) return;
  const sel = store.range;
  const w = Math.max(...rows.map((r) => r.length));
  const target: CellRange = { r1: sel.r1, c1: sel.c1, r2: sel.r1 + rows.length - 1, c2: sel.c1 + w - 1 };
  if (target.r2 >= MAX_ROWS || target.c2 >= MAX_COLS) { store.notify(T('The paste area doesn\'t fit on the sheet.')); return; }
  store.change(N('Paste'), [{ area: target }], () => {
    rows.forEach((r, i) => r.forEach((c, j) => {
      store.enter(target.r1 + i, target.c1 + j, c.text);
      if (c.style) store.styleRange({ r1: target.r1 + i, c1: target.c1 + j, r2: target.r1 + i, c2: target.c1 + j }, (s) => ({ ...s, ...c.style }));
    }));
  });
  store.selectRange(target);
};

export const clearClipboardMarquee = (store: GridStore): void => {
  if (!store.marquee) return;
  store.marquee = null;
  if (clip?.cut) clip = null;
  store.emit();
};

export const hasInternalClip = (): boolean => clip !== null;
export { keyCol, keyRow };
