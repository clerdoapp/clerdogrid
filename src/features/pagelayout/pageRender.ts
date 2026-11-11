// Draws one printed page (cells, objects, headings, header and footer) onto a Painter.
import { CellError, isError, MAX_COLS, type Scalar } from '../../engine/types';
import { formatValue } from '../../formats/numfmt';
import { cellKey, colName, DEFAULT_FONT, type BorderSide, type CellRange, type CellStyle, type Chart, type RichRun, type Shape } from '../../model/workbook';
import { CfEvaluator, type CfResult } from '../../app/cf';
import { inRange, type GridStore } from '../../app/store';
import { shapePath } from '../objects/ShapeLayer';
import { pageGrid, type Page, type Pagination } from './paginate';
import { generalAlign } from '../../app/bidi';
import { parsePathData, drawSvg, roundRectPath, type Painter, type PathCmd, type TextStyle } from './painter';
import { fieldText, hfFor, parseHf, type HfContext, type HfFont, type HfRun } from './headerFooter';

const PT = 0.75;
const BORDER_W: Record<string, number> = { hair: 0.25, thin: 0.5, dotted: 0.5, dashed: 0.5, dashDot: 0.5, dashDotDot: 0.5, medium: 1.25, mediumDashed: 1.25, mediumDashDot: 1.25, mediumDashDotDot: 1.25, slantDashDot: 1.25, thick: 2, double: 0.5 };
const BORDER_DASH: Record<string, number[]> = { dotted: [0.5, 1], hair: [0.5, 0.5], dashed: [2.25, 0.75], mediumDashed: [4.5, 1.5], dashDot: [3.75, 1.5, 0.75, 1.5], mediumDashDot: [4.5, 1.5, 1.5, 1.5], dashDotDot: [3.75, 0.75, 0.75, 0.75, 0.75, 0.75], mediumDashDotDot: [4.5, 1.5, 1.5, 1.5, 1.5, 1.5], slantDashDot: [4.5, 1.5, 1.5, 1.5] };
const NAMED: Record<string, string> = { red: '#ff0000', blue: '#0000ff', green: '#00a000', black: '#000000', white: '#ffffff', yellow: '#c8a800', cyan: '#00a0a0', magenta: '#c000c0' };

export interface RenderContext {
  store: GridStore;
  /** The sheet being printed must be the store's active sheet while drawing. */
  pg: Pagination;
  page: Page;
  /** Pages in this sheet's printout, and this page's 1-based position in it. */
  total: number;
  position: number;
  hf: Omit<HfContext, 'page' | 'pages' | 'sheet'>;
  /** SVG markup of a chart at a size in pixels, or null when charts can't be drawn here. */
  chartSvg?: (chart: Chart, width: number, height: number) => string | null;
  /** Draw the page's cells only (no header/footer): used by the page layout view for click areas too. */
  noHeaderFooter?: boolean;
}

interface Pane { c1: number; c2: number; r1: number; r2: number; x: number; y: number; w: number; h: number; x0: number; y0: number }

const fontOf = (st: CellStyle, scale = 1): TextStyle => {
  const f = st.font ?? {};
  const size = (f.size ?? DEFAULT_FONT.size) * (f.vertAlign ? 0.7 : 1) * scale;
  return { family: f.name ?? DEFAULT_FONT.name, bold: f.bold, italic: f.italic, size, color: f.color ?? '#000000', underline: f.underline ? (f.underline.startsWith('double') ? 'double' : 'single') : undefined, strike: f.strike };
};

const mergeShallow = (a: CellStyle, b: CellStyle): CellStyle => ({ ...a, ...b, font: { ...a.font, ...b.font }, fill: b.fill ?? a.fill, border: { ...a.border, ...b.border } });

interface CellDraw { text: string; color?: string; align: 'left' | 'right' | 'center'; fill?: string; style: CellStyle; cf: CfResult | null; v: Scalar; repeat?: { ch: string; at: number }; rich?: RichRun[] }

export const renderPage = (p: Painter, ctx: RenderContext): void => {
  const { store, pg, page } = ctx;
  const sheet = store.sheet;
  const setup = pg.setup;
  const s = pg.scale;
  const layout = pg.layout;
  const grid = pageGrid(pg, page);
  const rtl = sheet.view.rtl;
  const draft = !!setup.draft;
  const mono = !!setup.blackAndWhite;
  p.mono = mono;
  const cf = new CfEvaluator(store);
  const draws = new Map<number, CellDraw>();
  const errMode = setup.errors ?? 'displayed';

  const cellDraw = (row: number, col: number): CellDraw => {
    const k = cellKey(row, col);
    const hit = draws.get(k);
    if (hit) return hit;
    const style = store.style(row, col);
    const cell = sheet.cells.get(k);
    let v = store.value(row, col);
    const res = sheet.conditionalFormats.length ? cf.evaluate(row, col) : null;
    const merged = res?.style ? mergeShallow(style, res.style) : style;
    let text = '', color: string | undefined, align: CellDraw['align'] = 'left', repeat: CellDraw['repeat'];
    if (v instanceof CellError && errMode !== 'displayed') v = errMode === 'blank' ? null : errMode === 'dash' ? '--' : new CellError('#N/A');
    if (v !== null && !res?.hideValue) {
      if (isError(v)) { text = (v as CellError).code; align = 'center'; }
      else if (typeof v === 'boolean') { text = v ? 'TRUE' : 'FALSE'; align = 'center'; }
      else {
        const f = formatValue(v, merged.numFmt ?? 'General', { date1904: store.wb.date1904 });
        text = f.text; color = f.color;
        if (f.fill !== undefined && f.fillAt !== undefined) repeat = { ch: f.fill, at: f.fillAt };
        // Same rules as the grid: text by its own direction, numbers to the trailing side of the sheet.
        align = generalAlign(v, f.text, rtl, !!merged.align?.rtl);
        if (f.align && typeof v === 'number') align = rtl ? (f.align === 'right' ? 'left' : 'right') : f.align;
      }
    }
    const h = merged.align?.h;
    if (h === 'left' || h === 'right' || h === 'center') align = h;
    else if (h === 'centerContinuous' || h === 'distributed') align = 'center';
    const d: CellDraw = { text, color, align, fill: res?.fill ?? merged.fill?.color, style: merged, cf: res, v, repeat, rich: typeof v === 'string' && cell?.rich?.length ? cell.rich : undefined };
    draws.set(k, d);
    return d;
  };
  const hasContent = (r: number, c: number) => store.value(r, c) !== null;

  // Panes: title columns/rows and the body, laid out in the block's unscaled points.
  const segX: { c1: number; c2: number; x: number }[] = [];
  const segY: { r1: number; r2: number; y: number }[] = [];
  let bx = pg.headW, by = pg.headH;
  const width = (a: number, b: number) => (layout.cols.start(b + 1) - layout.cols.start(a)) * PT;
  const height = (a: number, b: number) => (layout.rows.start(b + 1) - layout.rows.start(a)) * PT;
  if (page.titleCols) { segX.push({ c1: page.titleCols[0], c2: page.titleCols[1], x: bx }); bx += width(page.titleCols[0], page.titleCols[1]); }
  segX.push({ c1: page.cols[0], c2: page.cols[1], x: bx });
  if (page.titleRows) { segY.push({ r1: page.titleRows[0], r2: page.titleRows[1], y: by }); by += height(page.titleRows[0], page.titleRows[1]); }
  segY.push({ r1: page.rows[0], r2: page.rows[1], y: by });
  const panes: Pane[] = [];
  for (const sy of segY) for (const sx of segX) {
    panes.push({ c1: sx.c1, c2: sx.c2, r1: sy.r1, r2: sy.r2, x: sx.x, y: sy.y, w: width(sx.c1, sx.c2), h: height(sy.r1, sy.r2), x0: layout.cols.start(sx.c1) * PT, y0: layout.rows.start(sy.r1) * PT });
  }
  const mirror = (x: number, w: number) => (rtl ? page.w - x - w : x);

  p.save();
  p.transform(s, 0, 0, s, page.x, page.y);

  for (const pane of panes) {
    const colX = (c: number) => pane.x + layout.cols.start(c) * PT - pane.x0;
    const rowY = (r: number) => pane.y + layout.rows.start(r) * PT - pane.y0;
    const cw = (c: number) => layout.cols.size(c) * PT;
    const rh = (r: number) => layout.rows.size(r) * PT;
    const rectOf = (r: CellRange) => { const x = colX(r.c1), w = width(r.c1, r.c2); return { x: mirror(x, w), y: rowY(r.r1), w, h: height(r.r1, r.r2) }; };
    const cellRect = (r: number, c: number) => ({ x: mirror(colX(c), cw(c)), y: rowY(r), w: cw(c), h: rh(r) });
    const merges = sheet.merges.filter((m) => m.r2 >= pane.r1 && m.r1 <= pane.r2 && m.c2 >= pane.c1 && m.c1 <= pane.c2);
    const mergeAt = (r: number, c: number) => merges.find((m) => inRange(m, r, c));
    p.save();
    p.clip(mirror(pane.x, pane.w), pane.y, pane.w, pane.h);

    // Gridlines.
    if (setup.gridlines && !draft) {
      const st = { color: '#9a9a9a', width: 0.35 };
      const cmds: PathCmd[] = [];
      const x1 = mirror(pane.x, pane.w), x2 = x1 + pane.w;
      for (let r = pane.r1; r <= pane.r2 + 1; r++) { const y = rowY(r); cmds.push(['M', x1, y], ['L', x2, y]); }
      for (let c = pane.c1; c <= pane.c2 + 1; c++) { const x = rtl ? mirror(colX(c), 0) : colX(c); cmds.push(['M', x, pane.y], ['L', x, pane.y + pane.h]); }
      p.path(cmds, undefined, st);
      for (const m of merges) {
        const rr = rectOf(m);
        p.rect(rr.x + 0.3, rr.y + 0.3, rr.w - 0.6, rr.h - 0.6, '#ffffff');
      }
    }
    // Fills.
    if (!mono) {
      for (let r = pane.r1; r <= pane.r2; r++) {
        if (!rh(r)) continue;
        for (let c = pane.c1; c <= pane.c2; c++) {
          if (!cw(c)) continue;
          const m = mergeAt(r, c);
          if (m && (r !== Math.max(m.r1, pane.r1) || c !== Math.max(m.c1, pane.c1))) continue;
          const d = m ? cellDraw(m.r1, m.c1) : cellDraw(r, c);
          const rect = m ? rectOf(m) : cellRect(r, c);
          const g = d.style.fill?.gradient;
          if (g && !d.cf?.fill) p.gradient(rect.x, rect.y, rect.w, rect.h, g.angle, g.stops);
          else if (d.fill) p.rect(rect.x, rect.y, rect.w, rect.h, d.fill);
          else if (d.style.fill?.pattern && d.style.fill.pattern !== 'none' && d.style.fill.patternColor) p.rect(rect.x, rect.y, rect.w, rect.h, d.style.fill.patternColor, 0.35);
        }
      }
    }
    // Content.
    for (let r = pane.r1; r <= pane.r2; r++) {
      if (!rh(r)) continue;
      for (let c = Math.max(0, pane.c1 - 12); c <= Math.min(MAX_COLS - 1, pane.c2 + 12); c++) {
        if (!cw(c)) continue;
        const m = mergeAt(r, c) ?? sheet.merges.find((mm) => inRange(mm, r, c));
        if (m && (r !== Math.max(m.r1, pane.r1) || c !== Math.max(m.c1, pane.c1))) continue;
        const ar = m ? m.r1 : r, ac = m ? m.c1 : c;
        if (!hasContent(ar, ac)) continue;
        const d = cellDraw(ar, ac);
        const rect = m ? rectOf(m) : cellRect(r, c);
        if (d.cf?.bar && !mono) drawBar(p, rect, d.cf.bar);
        if (d.cf?.icon) drawIcon(p, rect, d.cf.icon);
        if (d.style.checkbox && typeof d.v === 'boolean') { drawCheckbox(p, rect, d.v); continue; }
        if (!d.text) continue;
        drawText(p, d, rect, r, c, !!m, (rr, cc) => hasContent(rr, cc) || !!sheet.merges.find((mm) => inRange(mm, rr, cc)), (cc) => cellRect(r, cc), pane, mono, d.cf?.icon ? 13.5 : 0, rtl);
      }
    }
    // Borders.
    for (let r = pane.r1; r <= pane.r2; r++) {
      const h = rh(r);
      if (!h) continue;
      for (let c = pane.c1; c <= pane.c2; c++) {
        const w = cw(c);
        if (!w) continue;
        const own = cellDraw(r, c).style.border;
        const right = sheet.cells.has(cellKey(r, c + 1)) ? cellDraw(r, c + 1).style.border?.left : undefined;
        const below = sheet.cells.has(cellKey(r + 1, c)) ? cellDraw(r + 1, c).style.border?.top : undefined;
        if (!own && !right && !below) continue;
        const b = { ...own, right: own?.right ?? right, bottom: own?.bottom ?? below };
        const { x, y } = cellRect(r, c);
        const L = rtl ? b.right : b.left, R = rtl ? b.left : b.right;
        if (b.top) borderLine(p, b.top, x, y, x + w, y, mono);
        if (b.bottom) borderLine(p, b.bottom, x, y + h, x + w, y + h, mono);
        if (L) borderLine(p, L, x, y, x, y + h, mono);
        if (R) borderLine(p, R, x + w, y, x + w, y + h, mono);
        if (b.diagonalDown) borderLine(p, b.diagonalDown, x, y, x + w, y + h, mono);
        if (b.diagonalUp) borderLine(p, b.diagonalUp, x, y + h, x + w, y, mono);
      }
    }
    // Links.
    for (const [k, uri] of sheet.hyperlinks) {
      const r = Math.floor(k / MAX_COLS), c = k % MAX_COLS;
      if (r < pane.r1 || r > pane.r2 || c < pane.c1 || c > pane.c2) continue;
      const rect = cellRect(r, c);
      const tl = { x: rect.x * s + page.x, y: rect.y * s + page.y, w: rect.w * s, h: rect.h * s };
      p.link(tl.x, tl.y, tl.w, tl.h, uri);
    }
    // Floating objects.
    if (!draft) {
      const objRect = (a: Chart['anchor']) => {
        const x = pane.x + (layout.cols.start(a.col) + a.dx) * PT - pane.x0, y = pane.y + (layout.rows.start(a.row) + a.dy) * PT - pane.y0;
        return { x: mirror(x, a.width * PT), y, w: a.width * PT, h: a.height * PT };
      };
      const visibleIn = (b: { x: number; y: number; w: number; h: number }) => b.x < mirror(pane.x, pane.w) + pane.w && b.x + b.w > mirror(pane.x, pane.w) && b.y < pane.y + pane.h && b.y + b.h > pane.y;
      for (const pic of sheet.pictures) { const b = objRect(pic.anchor); if (visibleIn(b)) p.image(pic.data, pic.mime, b.x, b.y, b.w, b.h); }
      for (const ch of sheet.charts) {
        const b = objRect(ch.anchor);
        if (!visibleIn(b)) continue;
        const svg = ctx.chartSvg?.(ch, ch.anchor.width, ch.anchor.height);
        if (svg) { p.rect(b.x, b.y, b.w, b.h, '#ffffff'); drawSvg(p, svg, b.x, b.y, b.w, b.h); p.path(roundRectPath(b.x, b.y, b.w, b.h, 0), undefined, { color: '#d9d9d9', width: 0.5 }); }
      }
      for (const sh of sheet.shapes ?? []) { const b = objRect(sh.anchor); if (visibleIn(b)) drawShape(p, sh, b); }
      if (setup.cellComments === 'asDisplayed') {
        for (const [k, note] of sheet.notes) {
          const r = Math.floor(k / MAX_COLS), c = k % MAX_COLS;
          if (r < pane.r1 || r > pane.r2 || c < pane.c1 || c > pane.c2) continue;
          const rect = cellRect(r, c);
          drawNote(p, rect.x + rect.w + 6, rect.y, (note.author ? `${note.author}:\n` : '') + note.text);
        }
      }
    }
    p.restore();
  }

  // Row and column headings.
  if (setup.headings) {
    const hs: TextStyle = { family: DEFAULT_FONT.name, size: 9, color: '#404040' };
    const line = { color: '#9a9a9a', width: 0.35 };
    for (const pane of panes) {
      if (pane.y === segY.at(-1)!.y && segY.length > 1 && pane.r1 === segY[0]!.r1) continue;
      for (let c = pane.c1; c <= pane.c2; c++) {
        const w = layout.cols.size(c) * PT;
        if (!w || pane !== panes.find((q) => q.c1 === pane.c1)) continue;
        const x = mirror(pane.x + layout.cols.start(c) * PT - pane.x0, w);
        p.rect(x, 0, w, pg.headH, '#efefef');
        p.path(roundRectPath(x, 0, w, pg.headH, 0), undefined, line);
        const t = colName(c);
        const tw = p.fonts.measure(t, hs, hs.size);
        p.text(t, x + (w - tw) / 2, pg.headH * 0.72, hs);
      }
    }
    for (const pane of panes) {
      if (pane !== panes.find((q) => q.r1 === pane.r1)) continue;
      for (let r = pane.r1; r <= pane.r2; r++) {
        const h = layout.rows.size(r) * PT;
        if (!h) continue;
        const y = pane.y + layout.rows.start(r) * PT - pane.y0;
        const x = mirror(0, pg.headW);
        p.rect(x, y, pg.headW, h, '#efefef');
        p.path(roundRectPath(x, y, pg.headW, h, 0), undefined, line);
        const t = String(r + 1);
        const tw = p.fonts.measure(t, hs, hs.size);
        p.text(t, x + (pg.headW - tw) / 2, y + h / 2 + hs.size * 0.35, hs);
      }
    }
    p.rect(mirror(0, pg.headW), 0, pg.headW, pg.headH, '#e2e2e2');
  }
  p.restore();

  if (!ctx.noHeaderFooter) drawHeaderFooter(p, ctx);
  void grid;
};

// ---------- Cell pieces ----------

const borderLine = (p: Painter, b: BorderSide, x1: number, y1: number, x2: number, y2: number, mono: boolean): void => {
  const color = mono ? '#000000' : b.color ?? '#000000';
  const w = BORDER_W[b.style] ?? 0.5;
  if (b.style === 'double') {
    const horiz = y1 === y2;
    const o = 0.9;
    const cmds: PathCmd[] = horiz ? [['M', x1, y1 - o], ['L', x2, y2 - o], ['M', x1, y1 + o], ['L', x2, y2 + o]] : [['M', x1 - o, y1], ['L', x2 - o, y2], ['M', x1 + o, y1], ['L', x2 + o, y2]];
    p.path(cmds, undefined, { color, width: 0.5 });
    return;
  }
  p.path([['M', x1, y1], ['L', x2, y2]], undefined, { color, width: w, dash: BORDER_DASH[b.style], cap: 'butt' });
};

const wrapLines = (p: Painter, st: TextStyle, text: string, width: number): string[] => {
  const out: string[] = [];
  const m = (t: string) => p.fonts.measure(t, st, st.size);
  for (const para of text.split('\n')) {
    const words = para.split(/(\s+)/);
    let line = '';
    for (const w of words) {
      const next = line + w;
      if (line && m(next.trimEnd()) > width) {
        out.push(line.trimEnd());
        line = w.trimStart();
        while (m(line) > width && line.length > 1) {
          let k = line.length;
          while (k > 1 && m(line.slice(0, k)) > width) k--;
          out.push(line.slice(0, k)); line = line.slice(k);
        }
      } else line = next;
    }
    out.push(line);
  }
  return out;
};

const drawText = (p: Painter, d: CellDraw, rect: { x: number; y: number; w: number; h: number }, row: number, col: number, merged: boolean,
  hasContent: (r: number, c: number) => boolean, cellRect: (c: number) => { x: number; w: number }, pane: Pane, mono: boolean, iconPad: number, rtl: boolean): void => {
  const st0 = d.style;
  let font = fontOf(st0);
  font.color = mono ? '#000000' : d.color ? NAMED[d.color.toLowerCase()] ?? d.color : font.color;
  const pad = 2.25;
  const indent = (st0.align?.indent ?? 0) * 6.75;
  const wrap = st0.align?.wrap || st0.align?.h === 'justify' || st0.align?.h === 'distributed';
  const measure = (t: string, f = font) => p.fonts.measure(t, f, f.size);
  let text = d.text;
  const isNum = typeof d.v === 'number';
  let tw = measure(text);
  const avail = rect.w - pad * 2 - indent - iconPad;
  if (st0.align?.shrink && !wrap && tw > avail && avail > 0) { font = { ...font, size: Math.max(1, (font.size * avail) / tw) }; tw = measure(text); }
  if (isNum && !wrap && tw > avail && text && !/[a-z]/i.test(text.replace(/E[+-]/, ''))) {
    const fmt = st0.numFmt ?? 'General';
    if (fmt === 'General') {
      const chars = Math.max(1, Math.floor(avail / (measure('0') || 5)));
      text = formatValue(d.v as number, 'General', { width: chars }).text;
      tw = measure(text);
    }
    if (tw > avail) { const hash = measure('#') || 5; text = '#'.repeat(Math.max(1, Math.floor(avail / hash))); tw = measure(text); }
  }
  if (d.repeat && tw < avail) {
    const cw = measure(d.repeat.ch) || 1;
    const n = Math.floor((avail - tw) / cw);
    if (n > 0) { text = text.slice(0, d.repeat.at) + d.repeat.ch.repeat(n) + text.slice(d.repeat.at); tw = measure(text); }
  }
  if (st0.align?.h === 'fill' && text) { const unit = measure(text); if (unit > 0) { text = text.repeat(Math.max(1, Math.floor(avail / unit))); tw = measure(text); } }
  const align = d.align;
  // Unwrapped text spills into empty neighbours on the same page.
  let clipX = rect.x, clipW = rect.w, x = rect.x, w = rect.w;
  if (!wrap && !merged && !isNum && tw > avail && st0.align?.rotation === undefined) {
    const spill = (dir: 1 | -1, need: number) => {
      let c = col + dir, extra = 0;
      while (c >= pane.c1 && c <= pane.c2 && extra < need && !hasContent(row, c)) { extra += cellRect(c).w; c += dir; }
      return extra;
    };
    const need = tw + pad * 2 + indent - rect.w;
    const goRight = (align === 'left') !== rtl;
    if (align === 'center') { const r = spill(rtl ? -1 : 1, need / 2), l = spill(rtl ? 1 : -1, need / 2); clipX -= l; clipW += l + r; x -= l; w += l + r; }
    else if (goRight) { const e = spill(rtl ? -1 : 1, need); clipW += e; }
    else { const e = spill(rtl ? 1 : -1, need); clipX -= e; clipW += e; }
  }
  const lineH = font.size * 1.22;
  const ascent = font.size * 0.78;
  p.save();
  p.clip(clipX, rect.y, clipW, rect.h);
  const rot = st0.align?.rotation;
  if (rot !== undefined && rot !== 0) {
    p.transform(1, 0, 0, 1, rect.x + rect.w / 2, rect.y + rect.h / 2);
    if (rot === 255) {
      const chars = [...text];
      chars.forEach((ch, i) => { const cw = measure(ch); p.text(ch, -cw / 2, -((chars.length - 1) * lineH) / 2 + i * lineH + ascent / 2, font); });
    } else {
      const angle = rot <= 90 ? -rot : rot - 90;
      const a = (angle * Math.PI) / 180;
      p.transform(Math.cos(a), Math.sin(a), -Math.sin(a), Math.cos(a), 0, 0);
      p.text(text, -tw / 2, ascent / 2, font);
    }
    p.restore();
    return;
  }
  const lines = wrap ? wrapLines(p, font, text, Math.max(1, rect.w - pad * 2 - indent)) : text.split('\n').slice(0, 1);
  const totalH = lines.length * lineH;
  const va = st0.align?.v ?? 'bottom';
  let ty = va === 'top' ? rect.y + pad + ascent : va === 'center' || va === 'justify' || va === 'distributed' ? rect.y + (rect.h - totalH) / 2 + ascent + (lineH - font.size) / 2 : rect.y + rect.h - pad - (totalH - lineH) - (lineH - ascent) + 0.75;
  if (d.rich && !wrap && lines.length === 1) {
    const runs = d.rich.map((r) => ({ text: r.text.split('\n')[0]!, st: { ...font, ...fontOf({ font: { ...st0.font, ...r.font } }), color: mono ? '#000000' : r.font?.color ?? font.color } }));
    const total = runs.reduce((a, r) => a + measure(r.text, r.st), 0);
    let tx = align === 'right' ? x + w - pad - total - indent : align === 'center' ? x + (w - total) / 2 : x + pad + indent + iconPad;
    for (const r of runs) tx += p.text(r.text, tx, ty, r.st);
    p.restore();
    return;
  }
  for (const ln of lines) {
    const lw = measure(ln);
    let tx = align === 'right' ? x + w - pad - lw - indent : align === 'center' ? x + (w - lw) / 2 : x + pad + indent + iconPad;
    if (align === 'right' && iconPad) tx = Math.max(tx, rect.x + iconPad);
    if (st0.align?.h === 'centerContinuous') tx = rect.x + (rect.w - lw) / 2;
    p.text(ln, tx, ty, font, st0.align?.rtl);
    ty += lineH;
  }
  p.restore();
};

const drawCheckbox = (p: Painter, rect: { x: number; y: number; w: number; h: number }, on: boolean): void => {
  const size = Math.max(6, Math.min(10.5, rect.h - 3, rect.w - 3));
  const x = rect.x + (rect.w - size) / 2, y = rect.y + (rect.h - size) / 2;
  p.path(roundRectPath(x, y, size, size, 1.8), on ? '#0c7a4b' : '#ffffff', on ? undefined : { color: '#6b7a71', width: 1 });
  if (on) p.path([['M', x + size * 0.24, y + size * 0.52], ['L', x + size * 0.43, y + size * 0.7], ['L', x + size * 0.77, y + size * 0.32]], undefined, { color: '#ffffff', width: 1.4, cap: 'round', join: 'round' });
};

const drawBar = (p: Painter, rect: { x: number; y: number; w: number; h: number }, bar: NonNullable<CfResult['bar']>): void => {
  const pad = 1.5, fullW = rect.w - pad * 2;
  let x = rect.x + pad;
  const w = fullW * Math.min(1, bar.fraction);
  if (bar.axis !== undefined) {
    const ax = rect.x + pad + fullW * bar.axis;
    x = bar.negative ? ax - w : ax;
    p.path([['M', ax, rect.y], ['L', ax, rect.y + rect.h]], undefined, { color: '#000000', width: 0.5, dash: [1.5, 1.5] });
  }
  const y = rect.y + pad, h = rect.h - pad * 2;
  if (bar.gradient) p.gradient(x, y, w, h, 0, [{ pos: 0, color: bar.color }, { pos: 1, color: '#ffffff' }]);
  else p.rect(x, y, w, h, bar.color);
  p.path(roundRectPath(x, y, w, h, 0), undefined, { color: bar.color, width: 0.5 });
};

const drawIcon = (p: Painter, rect: { x: number; y: number; w: number; h: number }, icon: { set: string; index: number }): void => {
  const s = 9;
  const x = rect.x + 2.25, y = rect.y + (rect.h - s) / 2;
  const n = parseInt(icon.set, 10) || 3;
  const t = n > 1 ? icon.index / (n - 1) : 1;
  const color = /Gray/.test(icon.set) ? '#808080' : t >= 0.99 ? '#1e9e48' : t <= 0.01 ? '#d93025' : t >= 0.5 ? '#f2b400' : '#f28b00';
  const cx = x + s / 2, cy = y + s / 2;
  if (/Arrow/.test(icon.set)) {
    const ang = (-90 + 180 * (1 - t)) * (Math.PI / 180) + Math.PI / 2;
    const pts: [number, number][] = [[0, -s / 2], [s / 2, 0], [s / 5, 0], [s / 5, s / 2], [-s / 5, s / 2], [-s / 5, 0], [-s / 2, 0]];
    const rot = pts.map(([px, py]) => [cx + px * Math.cos(ang) - py * Math.sin(ang), cy + px * Math.sin(ang) + py * Math.cos(ang)] as [number, number]);
    p.path([['M', ...rot[0]!], ...rot.slice(1).map((q) => ['L', ...q] as PathCmd), ['Z']], color);
    return;
  }
  if (/Rating|Quarters|Boxes|Star/.test(icon.set)) {
    p.path(parsePathData(`M${cx + s / 2 - 0.75},${cy}A${s / 2 - 0.75},${s / 2 - 0.75} 0 1 0 ${cx - s / 2 + 0.75},${cy}A${s / 2 - 0.75},${s / 2 - 0.75} 0 1 0 ${cx + s / 2 - 0.75},${cy}Z`), undefined, { color: '#4472c4', width: 0.75 });
    if (t > 0) {
      const a = -Math.PI / 2 + Math.PI * 2 * Math.min(0.9999, t), r = s / 2 - 0.75;
      p.path(parsePathData(`M${cx},${cy}L${cx},${cy - r}A${r},${r} 0 ${t > 0.5 ? 1 : 0} 1 ${cx + r * Math.cos(a)},${cy + r * Math.sin(a)}Z`), '#4472c4');
    }
    return;
  }
  if (/Triangle/.test(icon.set)) {
    if (t > 0.5) p.path([['M', cx, y], ['L', x + s, y + s], ['L', x, y + s], ['Z']], color);
    else if (t < 0.5) p.path([['M', x, y], ['L', x + s, y], ['L', cx, y + s], ['Z']], color);
    else p.rect(x, y + s / 3, s, s / 3, color);
    return;
  }
  if (/Flag/.test(icon.set)) { p.rect(x + 0.75, y, 1.1, s, color); p.path([['M', x + 1.9, y], ['L', x + s, y + s / 4], ['L', x + 1.9, y + s / 2], ['Z']], color); return; }
  p.path(parsePathData(`M${cx + s / 2},${cy}A${s / 2},${s / 2} 0 1 0 ${cx - s / 2},${cy}A${s / 2},${s / 2} 0 1 0 ${cx + s / 2},${cy}Z`), color);
};

const drawShape = (p: Painter, sh: Shape, b: { x: number; y: number; w: number; h: number }): void => {
  p.save();
  if (sh.rotation) {
    const a = (sh.rotation * Math.PI) / 180, cx = b.x + b.w / 2, cy = b.y + b.h / 2;
    p.transform(1, 0, 0, 1, cx, cy); p.transform(Math.cos(a), Math.sin(a), -Math.sin(a), Math.cos(a), 0, 0); p.transform(1, 0, 0, 1, -cx, -cy);
  }
  p.transform(1, 0, 0, 1, b.x, b.y);
  const isLine = sh.kind === 'line' || sh.kind === 'arrowLine';
  const cmds = parsePathData(shapePath(sh.kind, b.w, b.h));
  const line = sh.line ?? (sh.kind === 'textBox' ? undefined : { color: '#0C7A4B', width: 1.5 });
  const fill = isLine ? undefined : sh.fill;
  p.path(cmds, fill, line && line.width > 0 ? { color: line.color, width: line.width * PT, dash: line.dash === 'dash' ? [4, 2] : line.dash === 'dot' ? [1, 1.5] : undefined, join: 'round' } : undefined);
  if (sh.kind === 'arrowLine' && line) {
    const ang = Math.atan2(b.h, b.w), L = 6 + line.width * 1.5;
    const tip: [number, number] = [b.w, b.h];
    const l: [number, number] = [tip[0] - L * Math.cos(ang - 0.4), tip[1] - L * Math.sin(ang - 0.4)], r: [number, number] = [tip[0] - L * Math.cos(ang + 0.4), tip[1] - L * Math.sin(ang + 0.4)];
    p.path([['M', ...tip], ['L', ...l], ['L', ...r], ['Z']], line.color);
  }
  if (sh.text && !isLine) {
    const st = fontOf({ font: { size: 11, ...sh.font } });
    if (!sh.font?.color) st.color = sh.fill && luminance(sh.fill) < 0.4 ? '#ffffff' : '#111111';
    p.clip(0, 0, b.w, b.h);
    const lines = wrapLines(p, st, sh.text, Math.max(1, b.w - 9));
    const lineH = st.size * 1.25;
    let y = (b.h - lines.length * lineH) / 2 + st.size * 0.95;
    for (const ln of lines) {
      const lw = p.fonts.measure(ln, st, st.size);
      const x = sh.align === 'left' ? 4.5 : sh.align === 'right' ? b.w - 4.5 - lw : (b.w - lw) / 2;
      p.text(ln, x, y, st);
      y += lineH;
    }
  }
  p.restore();
};

const luminance = (hex: string): number => {
  const m = /^#?([0-9a-f]{6})/i.exec(hex);
  if (!m) return 1;
  const v = parseInt(m[1]!, 16);
  return (0.299 * (v >> 16) + 0.587 * ((v >> 8) & 255) + 0.114 * (v & 255)) / 255;
};

const drawNote = (p: Painter, x: number, y: number, text: string): void => {
  const st: TextStyle = { family: 'Tahoma', size: 8, color: '#000000' };
  const lines = wrapLines(p, st, text, 120);
  const h = lines.length * 10 + 6;
  p.rect(x, y, 128, h, '#ffffe1');
  p.path(roundRectPath(x, y, 128, h, 0), undefined, { color: '#7f7f7f', width: 0.5 });
  lines.forEach((ln, i) => p.text(ln, x + 4, y + 10 + i * 10, st));
};

// ---------- Header and footer ----------

const hfStyle = (f: HfFont, scale: number): TextStyle => ({
  family: f.name ?? DEFAULT_FONT.name, bold: f.bold, italic: f.italic, size: (f.size ?? DEFAULT_FONT.size) * scale * (f.vert ? 0.7 : 1),
  color: f.color ?? '#000000', underline: f.underline, strike: f.strike,
});

interface HfPiece { text?: string; picture?: string; style: TextStyle; w: number; h: number; rise: number }

export const drawHeaderFooter = (p: Painter, ctx: RenderContext): void => {
  const { pg, page, store } = ctx;
  const s = pg.setup;
  const which = hfFor(s, ctx.position);
  const fscale = s.scaleWithDoc === false ? 1 : pg.scale;
  const align = s.alignWithMargins !== false;
  const left = align ? pg.margins.left : 0.3 * 72, right = pg.paperW - (align ? pg.margins.right : 0.3 * 72);
  const hc: HfContext = { ...ctx.hf, page: page.number, pages: ctx.total + (s.firstPageNumber ?? 1) - 1, sheet: store.sheet.name };
  const pictures = s.headerPictures ?? {};
  for (const [kind, src] of [['H', which.header], ['F', which.footer]] as const) {
    if (!src) continue;
    const secs = parseHf(src);
    for (const sec of ['L', 'C', 'R'] as const) {
      const runs = secs[sec];
      if (!runs.length) continue;
      const slot = pictures[`${sec}${kind}${which.slot}`] ? `${sec}${kind}${which.slot}` : `${sec}${kind}`;
      const lines = layoutLines(p, runs, hc, fscale, slot, pictures[slot]);
      const heights = lines.map((l) => Math.max(...l.map((x) => x.h), DEFAULT_FONT.size * fscale * 1.2));
      const total = heights.reduce((a, b) => a + b, 0);
      let y = kind === 'H' ? pg.margins.header : pg.paperH - pg.margins.footer - total;
      lines.forEach((ln, i) => {
        const lw = ln.reduce((a, x) => a + x.w, 0);
        let x = sec === 'L' ? left : sec === 'R' ? right - lw : (pg.paperW - lw) / 2;
        const base = y + heights[i]! * 0.8;
        for (const piece of ln) {
          if (piece.picture) {
            const pic = pictures[piece.picture];
            if (pic) p.image(pic.data, pic.mime, x, base - piece.h * 0.9, piece.w, piece.h);
          } else if (piece.text) p.text(piece.text, x, base - piece.rise, piece.style);
          x += piece.w;
        }
        y += heights[i]!;
      });
    }
  }
};

const layoutLines = (p: Painter, runs: HfRun[], c: HfContext, scale: number, slot: string, picture: { width: number; height: number } | undefined): HfPiece[][] => {
  const lines: HfPiece[][] = [[]];
  for (const r of runs) {
    const st = hfStyle(r.font, scale);
    const rise = r.font.vert === 'super' ? st.size * 0.45 : r.font.vert === 'sub' ? -st.size * 0.2 : 0;
    if (r.kind === 'picture') {
      if (picture) lines.at(-1)!.push({ picture: slot, style: st, w: picture.width * scale, h: picture.height * scale, rise: 0 });
      continue;
    }
    const text = r.kind === 'text' ? r.text : fieldText(r.field, r.offset, c);
    text.split('\n').forEach((part, i) => {
      if (i) lines.push([]);
      if (!part) return;
      const line = lines.at(-1)!;
      const last = line.at(-1);
      // Neighbouring pieces in the same font are drawn as one string.
      if (last?.text !== undefined && last.rise === rise && JSON.stringify(last.style) === JSON.stringify(st)) { last.text += part; last.w = p.fonts.measure(last.text, st, st.size); }
      else line.push({ text: part, style: st, w: p.fonts.measure(part, st, st.size), h: st.size * 1.2, rise });
    });
  }
  return lines;
};
