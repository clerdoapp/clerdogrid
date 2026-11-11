// Splits a sheet into printed pages from the paper, margins, scaling, print area, print titles,
// manual breaks and the row heights and column widths. Shared by the break preview, the page
// layout view, print preview, printing and PDF export. All sizes are in points.
import { MAX_COLS, MAX_ROWS } from '../../engine/types';
import { keyCol, keyRow, parseRange, colIndex, type CellRange, type CellStyle, type PageSetup, type Sheet } from '../../model/workbook';
import { SheetLayout, colPx } from '../../app/layout';
import { lengthPt, paperByCode } from './paper';

export const DEFAULT_MARGINS = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 };

export interface Margins { left: number; right: number; top: number; bottom: number; header: number; footer: number }

export interface Page {
  /** 0-based position in the sheet's page sequence. */
  index: number;
  /** Printed page number (first page number + index). */
  number: number;
  area: number;
  rows: [number, number];
  cols: [number, number];
  titleRows?: [number, number];
  titleCols?: [number, number];
  /** Top-left of the printed block (headings, titles and cells) on the paper. */
  x: number; y: number;
  /** Size of the block before scaling. */
  w: number; h: number;
}

export interface Bands { area: CellRange; rowStarts: number[]; colStarts: number[]; manualRows: Set<number>; manualCols: Set<number> }

export interface Pagination {
  paperW: number; paperH: number;
  /** Margins in points. */
  margins: Margins;
  scale: number;
  /** Percent chosen by fit-to-page, if used. */
  fitPercent?: number;
  headW: number; headH: number;
  areas: CellRange[];
  bands: Bands[];
  pages: Page[];
  layout: SheetLayout;
  setup: PageSetup;
}

const PT = 0.75;

export const paperSize = (s: PageSetup): { w: number; h: number } => {
  const p = paperByCode(s.paper);
  let w = lengthPt(s.paperWidth) ?? p.w, h = lengthPt(s.paperHeight) ?? p.h;
  if ((s.orientation === 'landscape') !== (w > h)) [w, h] = [h, w];
  return { w, h };
};

export const marginsPt = (s: PageSetup): Margins => {
  const m = { ...DEFAULT_MARGINS, ...s.margins };
  return { left: m.left * 72, right: m.right * 72, top: m.top * 72, bottom: m.bottom * 72, header: m.header * 72, footer: m.footer * 72 };
};

/** "A1:D9,F1:G4" → ranges; whole rows/columns ("1:5", "A:C") too. */
export const parseAreas = (text: string | undefined): CellRange[] => {
  if (!text) return [];
  const out: CellRange[] = [];
  for (const raw of text.split(',')) {
    const t = raw.slice(raw.lastIndexOf('!') + 1).replace(/\$/g, '').trim();
    let m = /^(\d+):(\d+)$/.exec(t);
    if (m) { out.push({ r1: Math.min(+m[1]!, +m[2]!) - 1, r2: Math.max(+m[1]!, +m[2]!) - 1, c1: 0, c2: MAX_COLS - 1 }); continue; }
    m = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(t);
    if (m) { const a = colIndex(m[1]!), b = colIndex(m[2]!); out.push({ r1: 0, r2: MAX_ROWS - 1, c1: Math.min(a, b), c2: Math.max(a, b) }); continue; }
    const r = parseRange(t);
    if (r) out.push(r);
  }
  return out;
};

export const parseTitleRows = (t: string | undefined): [number, number] | undefined => {
  const m = t && /^\$?(\d+):\$?(\d+)$/.exec(t.trim());
  return m ? [Math.min(+m[1]!, +m[2]!) - 1, Math.max(+m[1]!, +m[2]!) - 1] : undefined;
};
export const parseTitleCols = (t: string | undefined): [number, number] | undefined => {
  const m = t && /^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$/.exec(t.trim());
  if (!m) return undefined;
  const a = colIndex(m[1]!), b = colIndex(m[2]!);
  return [Math.min(a, b), Math.max(a, b)];
};

const visible = (s: CellStyle | undefined): boolean => !!s && (!!s.fill?.color || !!s.fill?.gradient || (!!s.fill?.pattern && s.fill.pattern !== 'none') || !!s.border && Object.values(s.border).some(Boolean));

/** Cells that print when there is no print area: values, visible formatting and floating objects. */
export const usedArea = (sheet: Sheet, styles: CellStyle[], layout: SheetLayout): CellRange | null => {
  let r1 = Infinity, c1 = Infinity, r2 = -1, c2 = -1;
  const add = (r: number, c: number) => { if (r < r1) r1 = r; if (r > r2) r2 = r; if (c < c1) c1 = c; if (c > c2) c2 = c; };
  for (const [k, cell] of sheet.cells) {
    if ((cell.v === null || cell.v === '') && !cell.f && !visible(styles[cell.s ?? 0])) continue;
    add(keyRow(k), keyCol(k));
  }
  for (const m of sheet.merges) if (sheet.cells.has(m.r1 * MAX_COLS + m.c1)) { add(m.r1, m.c1); add(m.r2, m.c2); }
  const objs = [...sheet.charts.map((c) => c.anchor), ...sheet.pictures.map((p) => p.anchor), ...(sheet.shapes ?? []).map((s) => s.anchor)];
  for (const a of objs) {
    const x = layout.cols.start(a.col) + a.dx + a.width, y = layout.rows.start(a.row) + a.dy + a.height;
    add(a.row, a.col);
    add(layout.rows.at(Math.max(0, y - 1)), layout.cols.at(Math.max(0, x - 1)));
  }
  for (const k of sheet.notes.keys()) if (sheet.cells.has(k)) add(keyRow(k), keyCol(k));
  return r2 < 0 ? null : { r1, c1, r2, c2 };
};

export interface PaginateOptions {
  /** Print these ranges instead of the print area (for "print selection"). */
  areas?: CellRange[];
  ignorePrintArea?: boolean;
  styles?: CellStyle[];
  layout?: SheetLayout;
}

/** Start of each page band along one axis. */
const bandStarts = (start: number, end: number, size: (i: number) => number, avail: (bandStart: number) => number, manual: Set<number> | null): number[] => {
  const starts = [start];
  let cur = start, used = 0;
  for (let i = start; i <= end; i++) {
    const s = size(i);
    if (manual && i > cur && manual.has(i)) { starts.push(i); cur = i; used = 0; }
    if (s <= 0) continue;
    if (used > 0 && used + s > avail(cur) + 0.01) { starts.push(i); cur = i; used = 0; }
    used += s;
  }
  return starts;
};

export const paginate = (sheet: Sheet, opts: PaginateOptions = {}): Pagination => {
  const setup = sheet.pageSetup ?? {};
  const styles = opts.styles ?? [];
  const layout = opts.layout ?? new SheetLayout(sheet, styles);
  const { w: paperW, h: paperH } = paperSize(setup);
  const margins = marginsPt(setup);
  const printW = Math.max(36, paperW - margins.left - margins.right);
  const printH = Math.max(36, paperH - margins.top - margins.bottom);
  const colW = (c: number) => layout.cols.size(c) * PT;
  const rowH = (r: number) => layout.rows.size(r) * PT;

  const used = usedArea(sheet, styles, layout);
  let areas = opts.areas?.length ? opts.areas : !opts.ignorePrintArea ? parseAreas(setup.printArea) : [];
  if (!areas.length && used) areas = [used];
  // Whole rows or columns stop at the used range.
  areas = areas.map((a) => ({
    r1: a.r1, c1: a.c1,
    r2: a.r2 >= MAX_ROWS - 1 ? Math.max(a.r1, used?.r2 ?? a.r1) : a.r2,
    c2: a.c2 >= MAX_COLS - 1 ? Math.max(a.c1, used?.c2 ?? a.c1) : a.c2,
  }));

  const tRows = parseTitleRows(setup.printTitleRows), tCols = parseTitleCols(setup.printTitleCols);
  const sum = (a: number, b: number, f: (i: number) => number) => { let t = 0; for (let i = a; i <= b; i++) t += f(i); return t; };
  const titleH = tRows ? sum(tRows[0], tRows[1], rowH) : 0;
  const titleW = tCols ? sum(tCols[0], tCols[1], colW) : 0;
  const maxRow = Math.max(0, ...areas.map((a) => a.r2));
  const headW = setup.headings ? (colPx(Math.max(3, String(maxRow + 1).length) * 1.0) + 4) * PT : 0;
  const headH = setup.headings ? Math.max(12, sheet.defaultRowHeight) : 0;
  const fit = setup.fitWidth !== undefined || setup.fitHeight !== undefined;
  const manualR = fit ? null : new Set(setup.rowBreaks ?? []);
  const manualC = fit ? null : new Set(setup.colBreaks ?? []);

  const compute = (scale: number): Bands[] => areas.map((a) => {
    const availW = (start: number) => printW / scale - headW - (tCols && start > tCols[1] ? titleW : 0);
    const availH = (start: number) => printH / scale - headH - (tRows && start > tRows[1] ? titleH : 0);
    return {
      area: a,
      colStarts: bandStarts(a.c1, a.c2, colW, availW, manualC),
      rowStarts: bandStarts(a.r1, a.r2, rowH, availH, manualR),
      manualRows: new Set([...(setup.rowBreaks ?? [])].filter((r) => r > a.r1 && r <= a.r2)),
      manualCols: new Set([...(setup.colBreaks ?? [])].filter((c) => c > a.c1 && c <= a.c2)),
    };
  });

  let scale = Math.max(10, Math.min(400, setup.scale ?? 100)) / 100;
  let fitPercent: number | undefined;
  if (fit) {
    const fw = setup.fitWidth ?? 1, fh = setup.fitHeight ?? 1;
    const ok = (p: number) => compute(p / 100).every((b) => (!fw || b.colStarts.length <= fw) && (!fh || b.rowStarts.length <= fh));
    let lo = 10, hi = 100;
    if (ok(100)) lo = 100;
    else while (lo < hi) { const mid = Math.ceil((lo + hi) / 2); if (ok(mid)) lo = mid; else hi = mid - 1; }
    fitPercent = lo;
    scale = lo / 100;
  }
  const bands = compute(scale);

  const pages: Page[] = [];
  const first = setup.firstPageNumber ?? 1;
  const rtl = sheet.view.rtl;
  bands.forEach((b, ai) => {
    const rowRanges = b.rowStarts.map((s, i) => [s, (b.rowStarts[i + 1] ?? b.area.r2 + 1) - 1] as [number, number]);
    const colRanges = b.colStarts.map((s, i) => [s, (b.colStarts[i + 1] ?? b.area.c2 + 1) - 1] as [number, number]);
    const order: [[number, number], [number, number]][] = [];
    if (setup.order === 'overThenDown') { for (const rr of rowRanges) for (const cc of colRanges) order.push([rr, cc]); }
    else for (const cc of colRanges) for (const rr of rowRanges) order.push([rr, cc]);
    for (const [rr, cc] of order) {
      const bodyW = sum(cc[0], cc[1], colW), bodyH = sum(rr[0], rr[1], rowH);
      if (bodyW <= 0 || bodyH <= 0) continue;
      const tr = tRows && rr[0] > tRows[1] ? tRows : undefined;
      const tc = tCols && cc[0] > tCols[1] ? tCols : undefined;
      const w = headW + (tc ? titleW : 0) + bodyW, h = headH + (tr ? titleH : 0) + bodyH;
      const x = setup.centerH ? margins.left + (printW - w * scale) / 2 : rtl ? paperW - margins.right - w * scale : margins.left;
      const y = setup.centerV ? margins.top + (printH - h * scale) / 2 : margins.top;
      pages.push({ index: pages.length, number: first + pages.length, area: ai, rows: rr, cols: cc, titleRows: tr, titleCols: tc, x, y, w, h });
    }
  });
  return { paperW, paperH, margins, scale, fitPercent, headW, headH, areas, bands, pages, layout, setup };
};

/** Row and column positions (unscaled points from the block's top-left) of one page. */
export interface PageGrid { cols: { index: number; x: number; w: number }[]; rows: { index: number; y: number; h: number }[]; headW: number; headH: number }

export const pageGrid = (pg: Pagination, page: Page): PageGrid => {
  const cols: PageGrid['cols'] = [], rows: PageGrid['rows'] = [];
  let x = pg.headW, y = pg.headH;
  const addCols = (a: number, b: number) => { for (let c = a; c <= b; c++) { const w = pg.layout.cols.size(c) * PT; if (w > 0) { cols.push({ index: c, x, w }); x += w; } } };
  const addRows = (a: number, b: number) => { for (let r = a; r <= b; r++) { const h = pg.layout.rows.size(r) * PT; if (h > 0) { rows.push({ index: r, y, h }); y += h; } } };
  if (page.titleCols) addCols(page.titleCols[0], page.titleCols[1]);
  addCols(page.cols[0], page.cols[1]);
  if (page.titleRows) addRows(page.titleRows[0], page.titleRows[1]);
  addRows(page.rows[0], page.rows[1]);
  return { cols, rows, headW: pg.headW, headH: pg.headH };
};
