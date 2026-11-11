// Export as a web page (one sheet or the whole workbook) with styles inline, and as XML or JSON data.
import { cellKey, type CellRange, type CellStyle, type RichRun, type Sheet, type Workbook } from '../model/workbook';
import { isError, type Scalar } from '../engine/types';
import { fileColWidth } from './xlsx-chart';
import { isDateFormat } from './numfmt';
import { fromSerial, timeOf } from '../engine/date';

export interface CellSource {
  value: (sheet: number, row: number, col: number) => Scalar;
  text: (sheet: number, row: number, col: number) => string;
  style: (sheet: number, row: number, col: number) => CellStyle;
  bounds: (sheet: number) => { rows: number; cols: number };
}

const esc = (s: string): string => s.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]!);
const SAFE_LINK = /^(https:|mailto:)/i;
const color = (c: string | undefined): string | undefined => (c && /^#[0-9a-f]{6}([0-9a-f]{2})?$/i.test(c) ? c.slice(0, 7) : undefined);
const fontFamily = (n: string): string => `'${n.replace(/['\\<>;{}]/g, '')}'`;

const BORDER_CSS: Record<string, string> = {
  thin: '1px solid', medium: '2px solid', thick: '3px solid', hair: '1px dotted', dotted: '1px dotted', dashed: '1px dashed', mediumDashed: '2px dashed',
  double: '3px double', dashDot: '1px dashed', mediumDashDot: '2px dashed', dashDotDot: '1px dotted', mediumDashDotDot: '2px dotted', slantDashDot: '2px dashed',
};

const fontCss = (f: CellStyle['font'] | undefined): string[] => {
  if (!f) return [];
  const out: string[] = [];
  if (f.name) out.push(`font-family:${fontFamily(f.name)},sans-serif`);
  if (f.size) out.push(`font-size:${Math.max(1, Math.min(400, f.size))}pt`);
  if (f.bold) out.push('font-weight:700');
  if (f.italic) out.push('font-style:italic');
  const deco = [f.underline && 'underline', f.strike && 'line-through'].filter(Boolean).join(' ');
  if (deco) out.push(`text-decoration:${deco}`);
  if (f.underline?.startsWith('double')) out.push('text-decoration-style:double');
  const c = color(f.color);
  if (c) out.push(`color:${c}`);
  if (f.vertAlign) out.push(`vertical-align:${f.vertAlign === 'superscript' ? 'super' : 'sub'}`);
  return out;
};

export const cellCss = (st: CellStyle, v: Scalar): string => {
  const out = fontCss(st.font);
  const fill = color(st.fill?.color) ?? color(st.fill?.gradient?.stops[0]?.color);
  if (fill) out.push(`background:${fill}`);
  for (const side of ['top', 'right', 'bottom', 'left'] as const) {
    const b = st.border?.[side];
    if (b) out.push(`border-${side}:${BORDER_CSS[b.style] ?? '1px solid'} ${color(b.color) ?? '#000'}`);
  }
  const a = st.align ?? {};
  const h = a.h && a.h !== 'general' ? (a.h === 'centerContinuous' ? 'center' : a.h === 'distributed' ? 'justify' : a.h === 'fill' ? 'left' : a.h) : typeof v === 'number' ? 'right' : typeof v === 'boolean' || isError(v) ? 'center' : undefined;
  if (h) out.push(`text-align:${h}`);
  if (a.v) out.push(`vertical-align:${a.v === 'center' ? 'middle' : a.v === 'top' ? 'top' : a.v === 'bottom' ? 'bottom' : 'middle'}`);
  out.push(a.wrap ? 'white-space:pre-wrap' : 'white-space:pre');
  if (a.indent) out.push(`padding-left:${a.indent * 9 + 3}px`);
  if (a.rtl) out.push('direction:rtl');
  return out.join(';');
};

const richHtml = (runs: RichRun[]): string => runs.map((r) => { const css = fontCss(r.font).join(';'); return css ? `<span style="${css}">${esc(r.text)}</span>` : esc(r.text); }).join('');

export interface HtmlOptions {
  /** Sheets to include (all visible sheets when missing). */
  sheets?: number[];
  /** Only this area of the first listed sheet. */
  area?: CellRange;
  title?: string;
  gridlines?: boolean;
}

const sheetTable = (wb: Workbook, si: number, src: CellSource, area: CellRange | undefined, grid: boolean): string => {
  const s: Sheet = wb.sheets[si]!;
  const b = src.bounds(si);
  const r = area ?? { r1: 0, c1: 0, r2: b.rows, c2: b.cols };
  const covered = new Set<number>();
  const spanAt = new Map<number, { rs: number; cs: number }>();
  for (const m of s.merges) {
    if (m.r2 < r.r1 || m.r1 > r.r2 || m.c2 < r.c1 || m.c1 > r.c2) continue;
    const r1 = Math.max(m.r1, r.r1), c1 = Math.max(m.c1, r.c1), r2 = Math.min(m.r2, r.r2), c2 = Math.min(m.c2, r.c2);
    let rs = 0, cs = 0;
    for (let i = r1; i <= r2; i++) if (!s.rows.get(i)?.hidden) rs++;
    for (let j = c1; j <= c2; j++) if (!s.cols.get(j)?.hidden) cs++;
    spanAt.set(cellKey(r1, c1), { rs: Math.max(1, rs), cs: Math.max(1, cs) });
    for (let i = r1; i <= r2; i++) for (let j = c1; j <= c2; j++) if (i !== r1 || j !== c1) covered.add(cellKey(i, j));
  }
  const cols: string[] = [];
  for (let c = r.c1; c <= r.c2; c++) {
    if (s.cols.get(c)?.hidden) continue;
    cols.push(`<col style="width:${Math.round(fileColWidth(s.cols.get(c)?.width ?? s.defaultColWidth) * 7)}px">`);
  }
  const rows: string[] = [];
  for (let row = r.r1; row <= r.r2; row++) {
    const ri = s.rows.get(row);
    if (ri?.hidden) continue;
    const cells: string[] = [];
    for (let c = r.c1; c <= r.c2; c++) {
      if (s.cols.get(c)?.hidden) continue;
      const k = cellKey(row, c);
      if (covered.has(k)) continue;
      const v = src.value(si, row, c);
      const st = src.style(si, row, c);
      const span = spanAt.get(k);
      const cell = s.cells.get(k);
      const link = s.hyperlinks.get(k);
      const note = s.notes.get(k);
      let body = cell?.rich?.length && typeof v === 'string' ? richHtml(cell.rich) : esc(src.text(si, row, c));
      if (link && SAFE_LINK.test(link)) body = `<a href="${esc(link)}" rel="noopener noreferrer">${body}</a>`;
      else if (link?.startsWith('#')) {
        const target = link.slice(1).split('!')[0]!.replace(/^'|'$/g, '').replace(/''/g, "'");
        const ti = wb.sheets.findIndex((x) => x.name === target);
        if (ti >= 0) body = `<a href="#sheet${ti + 1}">${body}</a>`;
      }
      const at = [span?.rs && span.rs > 1 ? ` rowspan="${span.rs}"` : '', span?.cs && span.cs > 1 ? ` colspan="${span.cs}"` : '', note ? ` title="${esc(note.text)}"` : ''].join('');
      cells.push(`<td${at} style="${cellCss(st, v)}">${body}</td>`);
    }
    rows.push(`<tr${ri?.height ? ` style="height:${Math.round((ri.height * 96) / 72)}px"` : ''}>${cells.join('')}</tr>`);
  }
  const pics = s.pictures.filter((p) => /^image\/(png|jpeg|gif|webp|bmp)$/.test(p.mime)).map((p) => {
    let bin = '';
    for (let i = 0; i < p.data.length; i += 0x8000) bin += String.fromCharCode(...p.data.subarray(i, i + 0x8000));
    return `<img src="data:${p.mime};base64,${btoa(bin)}" width="${Math.round(p.anchor.width)}" height="${Math.round(p.anchor.height)}" alt="${esc(p.alt ?? '')}">`;
  }).join('');
  return `<table class="${grid ? 'grid' : ''}"><colgroup>${cols.join('')}</colgroup><tbody>${rows.join('')}</tbody></table>${pics ? `<div class="pics">${pics}</div>` : ''}`;
};

/** A complete, self-contained HTML document. All text is escaped; only https and mailto links are kept. */
export const exportHtml = (wb: Workbook, src: CellSource, o: HtmlOptions = {}): string => {
  const list = o.sheets ?? wb.sheets.map((s, i) => (s.state === 'visible' ? i : -1)).filter((i) => i >= 0);
  const title = o.title ?? 'Workbook';
  const multi = list.length > 1;
  const nav = multi ? `<nav>${list.map((i) => `<a href="#sheet${i + 1}">${esc(wb.sheets[i]!.name)}</a>`).join('')}</nav>` : '';
  const body = list.map((i, k) => {
    const grid = o.gridlines ?? wb.sheets[i]!.view.showGrid;
    return `<section id="sheet${i + 1}">${multi ? `<h2>${esc(wb.sheets[i]!.name)}</h2>` : ''}${sheetTable(wb, i, src, k === 0 ? o.area : undefined, grid)}</section>`;
  }).join('');
  return `<!doctype html>
<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><meta name="generator" content="ClerdoGrid"><title>${esc(title)}</title>
<style>body{font:11pt Carlito,Calibri,Arial,sans-serif;margin:16px;color:#000;background:#fff}table{border-collapse:collapse;table-layout:fixed;margin-bottom:24px}td{padding:1px 4px;overflow:hidden;vertical-align:bottom}
table.grid td{outline:1px solid #d9d9d9;outline-offset:-1px}nav{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:16px}h2{font-size:14pt;margin:16px 0 8px}.pics img{margin:4px;max-width:100%;height:auto}</style>
</head><body>${nav}${body}</body></html>`;
};

// ---------- XML and JSON data ----------

export interface DataOptions {
  sheet: number;
  area?: CellRange;
  /** First row holds field names. */
  headers: boolean;
  /** Write values as shown (formatted text) rather than raw values. */
  asShown?: boolean;
}

const tableData = (wb: Workbook, src: CellSource, o: DataOptions): { names: string[]; rows: Scalar[][]; shown: string[][] } => {
  const b = src.bounds(o.sheet);
  const r = o.area ?? { r1: 0, c1: 0, r2: b.rows, c2: b.cols };
  const raw: Scalar[][] = [], shown: string[][] = [];
  for (let row = r.r1; row <= r.r2; row++) {
    const vr: Scalar[] = [], sr: string[] = [];
    for (let c = r.c1; c <= r.c2; c++) { vr.push(src.value(o.sheet, row, c)); sr.push(src.text(o.sheet, row, c)); }
    raw.push(vr); shown.push(sr);
  }
  const width = r.c2 - r.c1 + 1;
  let names: string[];
  if (o.headers && raw.length) {
    const seen = new Map<string, number>();
    names = shown[0]!.map((h, i) => {
      const base = h.trim() || `Column${i + 1}`;
      const n = (seen.get(base) ?? 0) + 1;
      seen.set(base, n);
      return n > 1 ? `${base}${n}` : base;
    });
    raw.shift(); shown.shift();
  } else names = Array.from({ length: width }, (_, i) => `Column${i + 1}`);
  // drop trailing empty rows
  while (raw.length && raw.at(-1)!.every((v) => v === null || v === '')) { raw.pop(); shown.pop(); }
  void wb;
  return { names, rows: raw, shown };
};

const xmlName = (s: string): string => {
  let n = s.normalize('NFC').replace(/[^\p{L}\p{N}_.-]+/gu, '_');
  if (!/^[\p{L}_]/u.test(n)) n = `_${n}`;
  if (/^xml/i.test(n)) n = `_${n}`;
  return n || '_';
};

const isoOf = (v: number, code: string | undefined, date1904: boolean): string | undefined => {
  if (!code || !isDateFormat(code)) return undefined;
  const d = fromSerial(Math.floor(v), date1904);
  const t = timeOf(v - Math.floor(v));
  const p = (n: number, w = 2) => String(n).padStart(w, '0');
  const date = `${p(d.year, 4)}-${p(d.month)}-${p(d.day)}`;
  return v % 1 ? `${date}T${p(t.hour)}:${p(t.minute)}:${p(t.second)}` : date;
};

/** XML document: one element per row, one child element per column. */
export const exportXml = (wb: Workbook, src: CellSource, o: DataOptions): string => {
  const { names, rows, shown } = tableData(wb, src, o);
  const tags = names.map(xmlName);
  const r0 = o.area?.r1 ?? 0, c0 = o.area?.c1 ?? 0;
  const body = rows.map((r, i) => `  <row>\n${r.map((v, j) => {
    if (v === null || v === '') return `    <${tags[j]}/>`;
    const st = src.style(o.sheet, r0 + i + (o.headers ? 1 : 0), c0 + j);
    const iso = typeof v === 'number' ? isoOf(v, st.numFmt, wb.date1904) : undefined;
    const t = o.asShown ? shown[i]![j]! : iso ?? (isError(v) ? v.code : String(v));
    const type = typeof v === 'number' ? (iso ? 'date' : 'number') : typeof v === 'boolean' ? 'boolean' : isError(v) ? 'error' : 'string';
    return `    <${tags[j]} type="${type}">${esc(t)}</${tags[j]}>`;
  }).join('\n')}\n  </row>`).join('\n');
  const root = xmlName(wb.sheets[o.sheet]?.name ?? 'data');
  return `<?xml version="1.0" encoding="UTF-8"?>\n<${root}>\n${body}\n</${root}>\n`;
};

/** JSON: an array of objects keyed by the header row (or arrays when there are no headers). */
export const exportJson = (wb: Workbook, src: CellSource, o: DataOptions): string => {
  const { names, rows, shown } = tableData(wb, src, o);
  const r0 = o.area?.r1 ?? 0, c0 = o.area?.c1 ?? 0;
  const cell = (v: Scalar, i: number, j: number): unknown => {
    if (o.asShown) return shown[i]![j];
    if (v === null) return null;
    if (isError(v)) return v.code;
    if (typeof v === 'number') {
      if (!Number.isFinite(v)) return null;
      const iso = isoOf(v, src.style(o.sheet, r0 + i + (o.headers ? 1 : 0), c0 + j).numFmt, wb.date1904);
      return iso ?? v;
    }
    return v;
  };
  const data = rows.map((r, i) => (o.headers ? Object.fromEntries(names.map((n, j) => [n, cell(r[j] ?? null, i, j)])) : r.map((v, j) => cell(v, i, j))));
  return JSON.stringify(data, null, 2) + '\n';
};
