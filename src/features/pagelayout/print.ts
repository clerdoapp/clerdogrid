// Print jobs: which pages to print, loading their fonts, and drawing them to PDF or SVG.
// Used by print preview, printing, "Save as PDF" and the command-line tool.
import { createElement } from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { MAX_COLS } from '../../engine/types';
import { addr, type CellRange, type Chart } from '../../model/workbook';
import { layoutOf } from '../../app/layout';
import type { GridStore } from '../../app/store';
import { PdfDocument } from '../../formats/pdf';
import type { FontBook, FontRequest } from '../../formats/pdf-text';
import { ChartSvg } from '../../ui/charts/ChartSvg';
import { resolveChart } from '../../ui/charts/chartData';
import { paginate, type Page, type Pagination } from './paginate';
import { renderPage, drawHeaderFooter, type RenderContext } from './pageRender';
import { PdfPainter, type TextRasterizer } from './pdfPainter';
import { SvgPainter } from './svgPainter';
import type { Painter, PathCmd, StrokeStyle, TextStyle } from './painter';
import { N } from '../../ui/i18n';

export interface PrintOptions {
  what: 'sheet' | 'workbook' | 'selection';
  /** Sheet to print for "sheet" (default: the active sheet). */
  sheet?: number;
  ignorePrintArea?: boolean;
}

export interface NoteEntry { cell: string; text: string }

export interface PrintPage {
  sheet: number;
  pg: Pagination;
  page: Page;
  /** Pages in this sheet's printout and this page's 1-based position. */
  total: number;
  position: number;
  /** A page of the "comments at the end" list instead of cells. */
  notes?: NoteEntry[];
}

export interface PrintJob { pages: PrintPage[]; title: string }

export interface PrintEnv {
  fonts: FontBook;
  file?: string;
  path?: string;
  date?: Date;
  raster?: TextRasterizer;
  /** Pictures converted to PNG for formats PDF can't hold (GIF, BMP, WebP, SVG). */
  images?: Map<Uint8Array, Uint8Array>;
  imageUrl?: (data: Uint8Array, mime: string) => string;
}

/** Runs fn with `sheet` as the store's active sheet, without notifying the interface. */
export const withSheet = <T>(store: GridStore, sheet: number, fn: () => T): T => {
  const prev = store.wb.activeSheet;
  store.wb.activeSheet = sheet;
  try { return fn(); } finally { store.wb.activeSheet = prev; }
};

export const paginateSheet = (store: GridStore, sheet: number, o: { areas?: CellRange[]; ignorePrintArea?: boolean } = {}): Pagination => {
  const s = store.wb.sheets[sheet]!;
  return paginate(s, { styles: store.wb.styles, layout: layoutOf(s, store.getVersion(), store.wb.styles), areas: o.areas, ignorePrintArea: o.ignorePrintArea });
};

const notesOf = (store: GridStore, sheet: number): NoteEntry[] => {
  const s = store.wb.sheets[sheet]!;
  const out: NoteEntry[] = [];
  const keys = new Set([...s.notes.keys(), ...(s.comments?.keys() ?? [])]);
  for (const k of [...keys].sort((a, b) => (a % MAX_COLS) - (b % MAX_COLS) || Math.floor(a / MAX_COLS) - Math.floor(b / MAX_COLS))) {
    const cell = addr(Math.floor(k / MAX_COLS), k % MAX_COLS);
    const n = s.notes.get(k);
    if (n) out.push({ cell, text: n.author ? `${n.author}: ${n.text}` : n.text });
    const t = s.comments?.get(k);
    if (t) out.push({ cell, text: [`${t.author}: ${t.text}`, ...t.replies.map((r) => `${r.author}: ${r.text}`)].join('\n') });
  }
  return out;
};

const NOTES_PER_PAGE = 18;

export const planPrint = (store: GridStore, o: PrintOptions): PrintJob => {
  const sheets = o.what === 'workbook' ? store.wb.sheets.map((s, i) => (s.state === 'visible' ? i : -1)).filter((i) => i >= 0) : [o.sheet ?? store.sheetIndex];
  const pages: PrintPage[] = [];
  for (const si of sheets) {
    const areas = o.what === 'selection' ? store.sel.ranges.map((r) => ({ ...r })) : undefined;
    const pg = withSheet(store, si, () => paginateSheet(store, si, { areas, ignorePrintArea: o.ignorePrintArea }));
    const sheetPages: PrintPage[] = pg.pages.map((page) => ({ sheet: si, pg, page, total: 0, position: 0 }));
    if (pg.setup.cellComments === 'atEnd') {
      const notes = notesOf(store, si);
      for (let i = 0; i < notes.length; i += NOTES_PER_PAGE) {
        const page: Page = { index: sheetPages.length, number: (pg.setup.firstPageNumber ?? 1) + sheetPages.length, area: -1, rows: [0, 0], cols: [0, 0], x: pg.margins.left, y: pg.margins.top, w: 0, h: 0 };
        sheetPages.push({ sheet: si, pg, page, total: 0, position: 0, notes: notes.slice(i, i + NOTES_PER_PAGE) });
      }
    }
    sheetPages.forEach((p, i) => { p.total = sheetPages.length; p.position = i + 1; });
    pages.push(...sheetPages);
  }
  return { pages, title: store.file.name || N('Book1') };
};

const chartCache = new WeakMap<Chart, { key: string; svg: string }>();
const chartSvgFor = (store: GridStore) => (chart: Chart, width: number, height: number): string | null => {
  try {
    const data = resolveChart(store, chart);
    const key = `${width}x${height}|${JSON.stringify(data.series.map((s) => s.values))}|${JSON.stringify(chart)}`;
    const hit = chartCache.get(chart);
    if (hit && hit.key === key) return hit.svg;
    const svg = renderToStaticMarkup(createElement(ChartSvg, { chart, data, width, height }));
    chartCache.set(chart, { key, svg });
    return svg;
  } catch { return null; }
};

const drawNotesPage = (p: Painter, pp: PrintPage): void => {
  const { pg } = pp;
  const x = pg.margins.left, w = pg.paperW - pg.margins.left - pg.margins.right;
  let y = pg.margins.top + 12;
  const label: TextStyle = { family: 'Calibri', size: 10, bold: true, color: '#000000' };
  const body: TextStyle = { family: 'Calibri', size: 10, color: '#000000' };
  for (const n of pp.notes ?? []) {
    p.text(`Cell: ${n.cell}`, x, y, label);
    y += 14;
    for (const para of n.text.split('\n')) {
      let line = '';
      for (const word of para.split(/(\s+)/)) {
        if (line && p.fonts.measure(line + word, body, 10) > w - 18) { p.text(line.trimEnd(), x + 18, y, body); y += 13; line = word.trimStart(); }
        else line += word;
      }
      p.text(line, x + 18, y, body);
      y += 13;
    }
    y += 8;
    if (y > pg.paperH - pg.margins.bottom) break;
  }
};

const contextFor = (store: GridStore, pp: PrintPage, env: PrintEnv): RenderContext => ({
  store, pg: pp.pg, page: pp.page, total: pp.total, position: pp.position,
  hf: { date: env.date ?? new Date(), file: env.file ?? `${store.file.name || 'Book1'}.xlsx`, path: env.path ?? '' },
  chartSvg: chartSvgFor(store),
});

/** Draws one print page onto any painter. */
export const drawPrintPage = (store: GridStore, pp: PrintPage, p: Painter, env: PrintEnv): void => {
  withSheet(store, pp.sheet, () => {
    const ctx = contextFor(store, pp, env);
    if (pp.notes) { drawNotesPage(p, pp); p.mono = !!pp.pg.setup.blackAndWhite; drawHeaderFooter(p, ctx); }
    else renderPage(p, ctx);
  });
};

/** Records the text a page would draw, so its fonts can be loaded first. */
class TextCollector implements Painter {
  mono = false;
  texts = new Map<string, { req: FontRequest; chars: Set<string>; strings: Set<string> }>();
  constructor(readonly fonts: FontBook) {}
  save(): void { /* nothing */ }
  restore(): void { /* nothing */ }
  clip(): void { /* nothing */ }
  transform(): void { /* nothing */ }
  rect(): void { /* nothing */ }
  gradient(): void { /* nothing */ }
  path(_c: PathCmd[], _f?: string, _s?: StrokeStyle): void { /* nothing */ }
  image(): void { /* nothing */ }
  link(): void { /* nothing */ }
  text(text: string, _x: number, _y: number, st: TextStyle): number {
    const key = `${st.family ?? ''}|${st.bold ? 1 : 0}|${st.italic ? 1 : 0}`;
    let e = this.texts.get(key);
    if (!e) { e = { req: { family: st.family, bold: st.bold, italic: st.italic }, chars: new Set('0123456789#.,-+E '), strings: new Set() }; this.texts.set(key, e); }
    for (const ch of text) e.chars.add(ch);
    // Han characters are drawn in the Chinese, Japanese or Korean font their own text calls for.
    if (/[\u2e80-\u9fff\uac00-\ud7af\uf900-\ufaff]/.test(text) && e.strings.size < 20000) e.strings.add(text);
    return this.fonts.measure(text, st, st.size);
  }
}

/** Loads every font the job needs. Call before drawing. */
export const loadFonts = async (store: GridStore, job: PrintJob, env: PrintEnv): Promise<void> => {
  for (let round = 0; round < 2; round++) {
    const rec = new TextCollector(env.fonts);
    for (const pp of job.pages) drawPrintPage(store, pp, rec, env);
    for (const e of rec.texts.values()) {
      for (const t of e.strings) await env.fonts.ensure(t, e.req);
      await env.fonts.ensure([...e.chars].join(''), e.req);
    }
  }
};

export const jobToPdf = (store: GridStore, job: PrintJob, env: PrintEnv): Uint8Array => {
  const doc = new PdfDocument();
  let lastSheet = -1;
  for (const pp of job.pages) {
    if (pp.sheet !== lastSheet) { doc.bookmark(store.wb.sheets[pp.sheet]!.name, doc.pageCount); lastSheet = pp.sheet; }
    const page = doc.addPage(pp.pg.paperW, pp.pg.paperH);
    drawPrintPage(store, pp, new PdfPainter(page, env.fonts, { raster: env.raster, images: env.images }), env);
    page.end();
  }
  if (!job.pages.length) {
    const pg = paginateSheet(store, store.sheetIndex);
    doc.addPage(pg.paperW, pg.paperH).end();
  }
  return doc.finish({ title: store.wb.props?.title ?? job.title, author: store.wb.props?.creator });
};

export const pageToSvg = (store: GridStore, pp: PrintPage, env: PrintEnv, cssW?: string, cssH?: string): string => {
  const p = new SvgPainter(env.fonts, { imageUrl: env.imageUrl });
  drawPrintPage(store, pp, p, env);
  return p.svg(pp.pg.paperW, pp.pg.paperH, cssW, cssH);
};

/** Builds a PDF of a sheet, the selection or the whole workbook. */
export const buildPdf = async (store: GridStore, o: PrintOptions, env: PrintEnv): Promise<Uint8Array> => {
  const job = planPrint(store, o);
  await loadFonts(store, job, env);
  return jobToPdf(store, job, env);
};
