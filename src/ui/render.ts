// Draws one sheet onto a canvas. Pure drawing: no event handling here.
import { CellError, MAX_COLS, MAX_ROWS, type Scalar } from '../engine/types';
import { formatValue } from '../formats/numfmt';
import { cellKey, colName, DEFAULT_FONT, type CellRange, type CellStyle, type BorderSide } from '../model/workbook';
import { CfEvaluator, type CfResult } from '../app/cf';
import type { SheetLayout } from '../app/layout';
import { inRange, type GridStore } from '../app/store';

export interface Theme {
  bg: string; grid: string; header: string; headerInk: string; headerSel: string; headerSelInk: string; headerLine: string;
  accent: string; selFill: string; ink: string; frozen: string; dark: boolean;
}

export const readTheme = (el: Element): Theme => {
  const cs = getComputedStyle(el);
  const v = (n: string, d: string) => cs.getPropertyValue(n).trim() || d;
  return {
    bg: v('--sheet', '#ffffff'), grid: v('--gridline', '#e1e3ea'), header: v('--hdr', '#f4f5f9'), headerInk: v('--hdr-ink', '#5c6375'),
    headerSel: v('--hdr-sel', '#d8efe2'), headerSelInk: v('--accent', '#0c7a4b'), headerLine: v('--hdr-line', '#d6d9e3'), accent: v('--accent', '#0c7a4b'),
    selFill: v('--sel', 'rgba(18,165,101,0.13)'), ink: v('--cell-ink', '#111111'), frozen: v('--frozen', '#9aa3b8'), dark: v('--is-dark', '0') === '1',
  };
};

/** Where a pane of cells is drawn on screen. */
export interface Pane { r1: number; r2: number; c1: number; c2: number; x: number; y: number; w: number; h: number; scrollX: number; scrollY: number }

export interface Viewport {
  layout: SheetLayout;
  zoom: number;
  headerW: number;
  headerH: number;
  width: number;
  height: number;
  /** Scroll offset of the main pane in unzoomed sheet pixels, measured from the first unfrozen row/col. */
  scrollX: number;
  scrollY: number;
  frozenRows: number;
  frozenCols: number;
}

export interface Overlay {
  refs: { range: CellRange; color: string }[];
  fillPreview: CellRange | null;
  dragRange: CellRange | null;
  marqueePhase: number;
  editing: { row: number; col: number } | null;
  showFormulas: boolean;
}

/** Font names mapped to bundled fonts with the same metrics, then to common system fonts. */
const FONT_FALLBACK: Record<string, string> = {
  calibri: '"Calibri", "Carlito"', 'calibri light': '"Calibri Light", "Carlito"', cambria: '"Cambria", "Caladea"', arial: '"Arial", "Arimo", "Liberation Sans"',
  'times new roman': '"Times New Roman", "Tinos"', 'courier new': '"Courier New", "Cousine"', aptos: '"Aptos", "Carlito"', 'aptos narrow': '"Aptos Narrow", "Carlito"',
  helvetica: '"Helvetica", "Arimo"', verdana: '"Verdana", "Arimo"', tahoma: '"Tahoma", "Arimo"', segoe: '"Segoe UI", "Inter"', 'segoe ui': '"Segoe UI", "Inter"',
};
export const fontFamily = (name?: string): string => {
  const n = (name ?? DEFAULT_FONT.name).trim();
  return `${FONT_FALLBACK[n.toLowerCase()] ?? `"${n}"`}, "Carlito", "Noto Naskh Arabic", system-ui, sans-serif`;
};
export const cssFont = (s: CellStyle, zoom: number): string => {
  const f = s.font ?? {};
  const px = ((f.size ?? DEFAULT_FONT.size) * 4) / 3 * zoom * (f.vertAlign ? 0.7 : 1);
  return `${f.italic ? 'italic ' : ''}${f.bold ? '700 ' : '400 '}${px.toFixed(2)}px ${fontFamily(f.name)}`;
};

const BORDER_W: Record<string, number> = { thin: 1, hair: 1, dotted: 1, dashed: 1, dashDot: 1, dashDotDot: 1, medium: 2, mediumDashed: 2, mediumDashDot: 2, mediumDashDotDot: 2, slantDashDot: 2, thick: 3, double: 3 };
const BORDER_DASH: Record<string, number[]> = { dotted: [1, 1], hair: [1, 1], dashed: [3, 1], mediumDashed: [6, 2], dashDot: [5, 2, 1, 2], mediumDashDot: [6, 2, 2, 2], dashDotDot: [5, 1, 1, 1, 1, 1], mediumDashDotDot: [6, 2, 2, 2, 2, 2], slantDashDot: [6, 2, 2, 2] };

export const panesOf = (vp: Viewport): Pane[] => {
  const { layout, zoom, headerW, headerH, width, height, frozenRows: fr, frozenCols: fc } = vp;
  const frozenH = layout.rows.start(fr) * zoom, frozenW = layout.cols.start(fc) * zoom;
  const baseY = layout.rows.start(fr), baseX = layout.cols.start(fc);
  const mainY = baseY + vp.scrollY, mainX = baseX + vp.scrollX;
  const visRows = (start: number, px: number) => [layout.rows.at(start), Math.min(MAX_ROWS - 1, layout.rows.at(start + px / zoom))] as const;
  const visCols = (start: number, px: number) => [layout.cols.at(start), Math.min(MAX_COLS - 1, layout.cols.at(start + px / zoom))] as const;
  const mainW = Math.max(0, width - headerW - frozenW), mainH = Math.max(0, height - headerH - frozenH);
  const [mr1, mr2] = visRows(mainY, mainH);
  const [mc1, mc2] = visCols(mainX, mainW);
  const panes: Pane[] = [{ r1: mr1, r2: mr2, c1: mc1, c2: mc2, x: headerW + frozenW, y: headerH + frozenH, w: mainW, h: mainH, scrollX: mainX, scrollY: mainY }];
  if (fr) panes.push({ r1: 0, r2: fr - 1, c1: mc1, c2: mc2, x: headerW + frozenW, y: headerH, w: mainW, h: frozenH, scrollX: mainX, scrollY: 0 });
  if (fc) panes.push({ r1: mr1, r2: mr2, c1: 0, c2: fc - 1, x: headerW, y: headerH + frozenH, w: frozenW, h: mainH, scrollX: 0, scrollY: mainY });
  if (fr && fc) panes.push({ r1: 0, r2: fr - 1, c1: 0, c2: fc - 1, x: headerW, y: headerH, w: frozenW, h: frozenH, scrollX: 0, scrollY: 0 });
  return panes;
};

/** Screen rectangle of a cell range inside a pane (may extend past the pane). */
export const rectIn = (p: Pane, vp: Viewport, r: CellRange): { x: number; y: number; w: number; h: number } => {
  const { layout, zoom } = vp;
  const x = p.x + (layout.cols.start(r.c1) - p.scrollX) * zoom;
  const y = p.y + (layout.rows.start(r.r1) - p.scrollY) * zoom;
  return { x, y, w: (layout.cols.start(r.c2 + 1) - layout.cols.start(r.c1)) * zoom, h: (layout.rows.start(r.r2 + 1) - layout.rows.start(r.r1)) * zoom };
};

export const paneFor = (panes: Pane[], vp: Viewport, row: number, col: number): Pane => {
  const inFr = row < vp.frozenRows, inFc = col < vp.frozenCols;
  return panes.find((p) => (inFr ? p.scrollY === 0 && p.r2 < vp.frozenRows : p.r1 >= vp.frozenRows || p.scrollY !== 0) && (inFc ? p.scrollX === 0 && p.c2 < vp.frozenCols : p.c1 >= vp.frozenCols || p.scrollX !== 0)) ?? panes[0]!;
};

interface CellDraw { text: string; color?: string; align: 'left' | 'right' | 'center'; fill?: string; style: CellStyle; cf: CfResult | null; v: Scalar; repeat?: { ch: string; at: number } }

export class SheetRenderer {
  private measureCache = new Map<string, number>();
  constructor(private readonly ctx: CanvasRenderingContext2D) {}

  private measure(font: string, text: string): number {
    const k = font + '\u0001' + text;
    let w = this.measureCache.get(k);
    if (w === undefined) {
      this.ctx.font = font;
      w = this.ctx.measureText(text).width;
      if (this.measureCache.size > 20000) this.measureCache.clear();
      this.measureCache.set(k, w);
    }
    return w;
  }

  draw(store: GridStore, vp: Viewport, theme: Theme, ov: Overlay): void {
    const ctx = this.ctx;
    const sheet = store.sheet;
    const panes = panesOf(vp);
    const cf = new CfEvaluator(store);
    ctx.save();
    ctx.fillStyle = theme.bg;
    ctx.fillRect(0, 0, vp.width, vp.height);
    for (const p of panes) this.drawPane(store, vp, theme, p, cf, ov);
    this.drawHeaders(store, vp, theme, panes);
    // Frozen pane lines.
    ctx.strokeStyle = theme.frozen;
    ctx.lineWidth = 1;
    if (vp.frozenRows) { const y = vp.headerH + vp.layout.rows.start(vp.frozenRows) * vp.zoom + 0.5; ctx.beginPath(); ctx.moveTo(0, y); ctx.lineTo(vp.width, y); ctx.stroke(); }
    if (vp.frozenCols) { const x = vp.headerW + vp.layout.cols.start(vp.frozenCols) * vp.zoom + 0.5; ctx.beginPath(); ctx.moveTo(x, 0); ctx.lineTo(x, vp.height); ctx.stroke(); }
    // Selection, references and fill preview drawn over every pane, clipped to the cell area.
    ctx.save();
    ctx.beginPath(); ctx.rect(vp.headerW, vp.headerH, vp.width - vp.headerW, vp.height - vp.headerH); ctx.clip();
    for (const p of panes) {
      ctx.save();
      ctx.beginPath(); ctx.rect(p.x, p.y, p.w, p.h); ctx.clip();
      this.drawSelection(store, vp, theme, p, ov);
      ctx.restore();
    }
    ctx.restore();
    // Corner box.
    ctx.fillStyle = theme.header;
    ctx.fillRect(0, 0, vp.headerW, vp.headerH);
    ctx.fillStyle = theme.headerLine;
    ctx.beginPath(); ctx.moveTo(vp.headerW - 4, vp.headerH - 4); ctx.lineTo(vp.headerW - 4, 6); ctx.lineTo(vp.headerW - 4 - (vp.headerH - 10), vp.headerH - 4); ctx.closePath(); ctx.fill();
    void sheet;
    ctx.restore();
  }

  private cellDraw(store: GridStore, row: number, col: number, cf: CfEvaluator, showFormulas: boolean): CellDraw {
    const style = store.style(row, col);
    const c = store.sheet.cells.get(cellKey(row, col));
    const v = store.value(row, col);
    const res = store.sheet.conditionalFormats.length ? cf.evaluate(row, col) : null;
    const merged = res?.style ? mergeShallow(style, res.style) : style;
    let text = '', color: string | undefined, align: CellDraw['align'] = 'left';
    let repeat: CellDraw['repeat'];
    if (showFormulas && c?.f) { text = `=${c.f}`; }
    else if (v !== null && !res?.hideValue) {
      if (v instanceof CellError) { text = v.code; align = 'center'; }
      else if (typeof v === 'boolean') { text = v ? 'TRUE' : 'FALSE'; align = 'center'; }
      else {
        const f = formatValue(v, merged.numFmt ?? 'General', { date1904: store.wb.date1904 });
        text = f.text; color = f.color;
        if (f.fill !== undefined && f.fillAt !== undefined) repeat = { ch: f.fill, at: f.fillAt };
        align = typeof v === 'number' ? 'right' : 'left';
        if (f.align) align = f.align;
      }
    }
    const h = merged.align?.h;
    if (h === 'left' || h === 'right' || h === 'center') align = h;
    else if (h === 'centerContinuous' || h === 'distributed') align = 'center';
    return { text, color, align, fill: res?.fill ?? merged.fill?.color, style: merged, cf: res, v, repeat };
  }

  private drawPane(store: GridStore, vp: Viewport, theme: Theme, p: Pane, cf: CfEvaluator, ov: Overlay): void {
    const ctx = this.ctx;
    const { layout, zoom } = vp;
    const sheet = store.sheet;
    ctx.save();
    ctx.beginPath(); ctx.rect(p.x, p.y, p.w, p.h); ctx.clip();
    const colX = (c: number) => p.x + (layout.cols.start(c) - p.scrollX) * zoom;
    const rowY = (r: number) => p.y + (layout.rows.start(r) - p.scrollY) * zoom;
    const merges = sheet.merges.filter((m) => m.r2 >= p.r1 && m.r1 <= p.r2 && m.c2 >= p.c1 && m.c1 <= p.c2);
    const mergeAt = (r: number, c: number) => merges.find((m) => inRange(m, r, c));
    // Text may spill sideways, so look a little beyond the visible columns.
    const xc1 = Math.max(0, p.c1 - 12), xc2 = Math.min(MAX_COLS - 1, p.c2 + 12);
    const draws = new Map<number, CellDraw>();
    const get = (r: number, c: number): CellDraw => {
      const k = cellKey(r, c);
      let d = draws.get(k);
      if (!d) { d = this.cellDraw(store, r, c, cf, ov.showFormulas); draws.set(k, d); }
      return d;
    };
    const hasContent = (r: number, c: number) => store.value(r, c) !== null || (ov.editing?.row === r && ov.editing.col === c);

    // Fills.
    for (let r = p.r1; r <= p.r2; r++) {
      const h = layout.rows.size(r) * zoom;
      if (!h) continue;
      const y = rowY(r);
      for (let c = p.c1; c <= p.c2; c++) {
        const w = layout.cols.size(c) * zoom;
        if (!w) continue;
        const m = mergeAt(r, c);
        if (m && (m.r1 !== r || m.c1 !== c) && !(r === p.r1 && m.r1 < r) && !(c === p.c1 && m.c1 < c)) continue;
        const d = m ? get(m.r1, m.c1) : get(r, c);
        if (!d.fill && !d.style.fill?.pattern && !d.style.fill?.gradient) continue;
        const rect = m ? rectIn(p, vp, m) : { x: colX(c), y, w, h };
        if (d.style.fill?.gradient && !d.cf?.fill) {
          const g = d.style.fill.gradient;
          const rad = (g.angle * Math.PI) / 180;
          const gr = ctx.createLinearGradient(rect.x, rect.y, rect.x + Math.cos(rad) * rect.w, rect.y + Math.sin(rad) * rect.h);
          for (const s of g.stops) gr.addColorStop(s.pos, s.color);
          ctx.fillStyle = gr;
        } else ctx.fillStyle = d.fill ?? d.style.fill?.patternColor ?? theme.bg;
        ctx.fillRect(rect.x, rect.y, rect.w, rect.h);
      }
    }
    // Gridlines.
    if (sheet.view.showGrid) {
      ctx.strokeStyle = theme.grid;
      ctx.lineWidth = 1;
      ctx.beginPath();
      for (let c = p.c1; c <= p.c2 + 1; c++) { const x = Math.round(colX(c)) - 0.5; ctx.moveTo(x, p.y); ctx.lineTo(x, p.y + p.h); }
      for (let r = p.r1; r <= p.r2 + 1; r++) { const y = Math.round(rowY(r)) - 0.5; ctx.moveTo(p.x, y); ctx.lineTo(p.x + p.w, y); }
      ctx.stroke();
      // Cover gridlines inside merges and under filled cells.
      for (const m of merges) {
        const d = get(m.r1, m.c1);
        const rect = rectIn(p, vp, m);
        ctx.fillStyle = d.fill ?? theme.bg;
        ctx.fillRect(rect.x, rect.y, rect.w - 1, rect.h - 1);
      }
      for (let r = p.r1; r <= p.r2; r++) for (let c = p.c1; c <= p.c2; c++) {
        if (mergeAt(r, c)) continue;
        const d = get(r, c);
        if (!d.fill) continue;
        ctx.fillStyle = d.fill;
        const x = colX(c), y = rowY(r), w = layout.cols.size(c) * zoom, h = layout.rows.size(r) * zoom;
        ctx.fillRect(x - 1, y - 1, w + 1, h + 1);
      }
    }
    // Content.
    for (let r = p.r1; r <= p.r2; r++) {
      const h = layout.rows.size(r) * zoom;
      if (!h) continue;
      const y = rowY(r);
      for (let c = xc1; c <= xc2; c++) {
        const w = layout.cols.size(c) * zoom;
        if (!w) continue;
        const m = mergeAt(r, c) ?? sheet.merges.find((mm) => inRange(mm, r, c));
        if (m && (m.r1 !== r || m.c1 !== c)) { if (!(r === p.r1 && c === Math.max(p.c1, m.c1))) continue; }
        const anchorR = m ? m.r1 : r, anchorC = m ? m.c1 : c;
        const ck = cellKey(anchorR, anchorC);
        if (!hasContent(anchorR, anchorC) && !store.sheet.notes.has(ck) && !store.sheet.comments?.has(ck)) continue;
        if (ov.editing && ov.editing.row === anchorR && ov.editing.col === anchorC) continue;
        const d = get(anchorR, anchorC);
        const rect = m ? rectIn(p, vp, m) : { x: colX(c), y, w, h };
        if (d.cf?.bar) this.drawBar(rect, d.cf.bar, zoom);
        if (d.cf?.icon) this.drawIcon(rect, d.cf.icon, zoom);
        if (store.sheet.notes.has(cellKey(anchorR, anchorC))) {
          ctx.fillStyle = '#d93025';
          ctx.beginPath(); ctx.moveTo(rect.x + rect.w - 7 * zoom, rect.y + 1); ctx.lineTo(rect.x + rect.w - 1, rect.y + 1); ctx.lineTo(rect.x + rect.w - 1, rect.y + 7 * zoom); ctx.fill();
        }
        if (store.sheet.comments?.has(ck)) {
          // Comment threads: a violet corner, top-left so it never hides a note's red corner.
          ctx.fillStyle = store.sheet.comments.get(ck)!.resolved ? theme.headerInk : theme.dark ? '#b197fc' : '#7048e8';
          ctx.beginPath(); ctx.moveTo(rect.x + 1, rect.y + 1); ctx.lineTo(rect.x + 8 * zoom, rect.y + 1); ctx.lineTo(rect.x + 1, rect.y + 8 * zoom); ctx.fill();
        }
        if (d.style.checkbox && typeof d.v === 'boolean' && !ov.showFormulas) { this.drawCheckbox(rect, d.v, zoom, theme); continue; }
        if (!d.text) continue;
        this.drawText(store, vp, p, d, rect, anchorR, anchorC, !!m, hasContent, theme, d.cf?.icon ? 18 * zoom : 0);
      }
    }
    // Borders.
    for (let r = p.r1; r <= p.r2; r++) {
      const y = rowY(r), h = layout.rows.size(r) * zoom;
      if (!h) continue;
      for (let c = p.c1; c <= p.c2; c++) {
        const w = layout.cols.size(c) * zoom;
        if (!w) continue;
        const b = this.borderOf(store, r, c, get);
        if (!b) continue;
        const x = colX(c);
        if (b.top) this.line(b.top, x, y, x + w, y, theme);
        if (b.bottom) this.line(b.bottom, x, y + h, x + w, y + h, theme);
        if (b.left) this.line(b.left, x, y, x, y + h, theme);
        if (b.right) this.line(b.right, x + w, y, x + w, y + h, theme);
        if (b.diagonalDown) this.line(b.diagonalDown, x, y, x + w, y + h, theme);
        if (b.diagonalUp) this.line(b.diagonalUp, x, y + h, x + w, y, theme);
      }
    }
    // Filter buttons and the validation list arrow.
    const af = sheet.autoFilter;
    if (af && af.range.r1 >= p.r1 && af.range.r1 <= p.r2) {
      for (let c = Math.max(af.range.c1, p.c1); c <= Math.min(af.range.c2, p.c2); c++) {
        const active = af.columns.some((f) => f.col === c);
        this.dropButton(colX(c + 1) - 17 * zoom, rowY(af.range.r1) + layout.rows.size(af.range.r1) * zoom - 17 * zoom, 16 * zoom, theme, active ? 'filter' : 'arrow');
      }
    }
    for (const t of sheet.tables) {
      if (!t.headerRow || sheet.autoFilter) continue;
      if (t.range.r1 < p.r1 || t.range.r1 > p.r2) continue;
      for (let c = Math.max(t.range.c1, p.c1); c <= Math.min(t.range.c2, p.c2); c++) this.dropButton(colX(c + 1) - 17 * zoom, rowY(t.range.r1) + layout.rows.size(t.range.r1) * zoom - 17 * zoom, 16 * zoom, theme, 'arrow');
    }
    ctx.restore();
  }

  private borderOf(store: GridStore, r: number, c: number, get: (r: number, c: number) => CellDraw): CellStyle['border'] | null {
    const own = get(r, c).style.border;
    // A neighbour's matching edge also draws (files store borders on either cell).
    const right = c + 1 < MAX_COLS && store.sheet.cells.has(cellKey(r, c + 1)) ? get(r, c + 1).style.border?.left : undefined;
    const below = r + 1 < MAX_ROWS && store.sheet.cells.has(cellKey(r + 1, c)) ? get(r + 1, c).style.border?.top : undefined;
    if (!own && !right && !below) return null;
    return { ...own, right: own?.right ?? right, bottom: own?.bottom ?? below };
  }

  private line(b: BorderSide, x1: number, y1: number, x2: number, y2: number, theme: Theme): void {
    const ctx = this.ctx;
    const w = BORDER_W[b.style] ?? 1;
    ctx.strokeStyle = b.color ?? (theme.dark ? '#d8dbe6' : '#000000');
    ctx.setLineDash(BORDER_DASH[b.style] ?? []);
    const off = w % 2 ? 0.5 : 0;
    const snap = (v: number) => Math.round(v) - 1 + off;
    if (b.style === 'double') {
      ctx.lineWidth = 1;
      ctx.beginPath();
      if (y1 === y2) { ctx.moveTo(x1, snap(y1) - 1); ctx.lineTo(x2, snap(y2) - 1); ctx.moveTo(x1, snap(y1) + 1); ctx.lineTo(x2, snap(y2) + 1); }
      else { ctx.moveTo(snap(x1) - 1, y1); ctx.lineTo(snap(x2) - 1, y2); ctx.moveTo(snap(x1) + 1, y1); ctx.lineTo(snap(x2) + 1, y2); }
      ctx.stroke();
    } else {
      ctx.lineWidth = w;
      ctx.beginPath();
      if (y1 === y2) { ctx.moveTo(x1 - 1, snap(y1)); ctx.lineTo(x2, snap(y2)); }
      else if (x1 === x2) { ctx.moveTo(snap(x1), y1 - 1); ctx.lineTo(snap(x2), y2); }
      else { ctx.moveTo(x1, y1); ctx.lineTo(x2, y2); }
      ctx.stroke();
    }
    ctx.setLineDash([]);
  }

  private drawText(_store: GridStore, vp: Viewport, p: Pane, d: CellDraw, rect: { x: number; y: number; w: number; h: number }, row: number, col: number, merged: boolean,
    hasContent: (r: number, c: number) => boolean, theme: Theme, iconPad: number): void {
    const ctx = this.ctx;
    const { zoom, layout } = vp;
    const st = d.style;
    const font = cssFont(st, zoom);
    const pad = 3 * zoom;
    const indent = (st.align?.indent ?? 0) * 9 * zoom;
    const wrap = st.align?.wrap || st.align?.h === 'justify' || st.align?.h === 'distributed';
    const fontPx = ((st.font?.size ?? DEFAULT_FONT.size) * 4) / 3 * zoom;
    const lineH = fontPx * 1.22;
    let text = d.text;
    let { x, w } = rect;
    const isNum = typeof d.v === 'number';
    let tw = this.measure(font, text);
    const avail = rect.w - pad * 2 - indent - iconPad;
    if (isNum && !wrap && tw > avail && d.text && !/[a-z]/i.test(d.text.replace(/E[+-]/, ''))) {
      // Numbers never spill: General shortens, other formats show ####.
      const fmt = st.numFmt ?? 'General';
      if (fmt === 'General') {
        const chars = Math.max(1, Math.floor(avail / (this.measure(font, '0') || 7)));
        text = formatValue(d.v as number, 'General', { width: chars }).text;
        tw = this.measure(font, text);
      }
      if (tw > avail) { const hash = this.measure(font, '#'); text = '#'.repeat(Math.max(1, Math.floor(avail / hash))); tw = this.measure(font, text); }
    }
    if (d.repeat && tw < avail) {
      // "*x" in the format: repeat x so the text fills the cell (accounting formats pin $ to the left).
      const cw = this.measure(font, d.repeat.ch) || 1;
      const n = Math.floor((avail - tw) / cw);
      if (n > 0) { text = text.slice(0, d.repeat.at) + d.repeat.ch.repeat(n) + text.slice(d.repeat.at); tw = this.measure(font, text); }
    }
    if (st.align?.h === 'fill' && text) {
      const unit = this.measure(font, text);
      if (unit > 0) { text = text.repeat(Math.max(1, Math.floor(avail / unit))); tw = this.measure(font, text); }
    }
    // Overflow into empty neighbours for unwrapped text.
    let clipX = rect.x, clipW = rect.w;
    if (!wrap && !merged && !isNum && tw > avail && st.align?.rotation === undefined) {
      if (d.align === 'left' || d.align === 'center') {
        let c = col + 1;
        let extra = 0;
        while (c < MAX_COLS && extra + rect.w < tw + pad * 2 + indent && !hasContent(row, c) && c <= p.c2 + 12) { extra += layout.cols.size(c) * zoom; c++; }
        clipW += extra;
        if (d.align === 'center') { let cl = col - 1, left = 0; while (cl >= 0 && left < extra && !hasContent(row, cl)) { left += layout.cols.size(cl) * zoom; cl--; } clipX -= left; clipW += left; x -= left; w += left + extra; }
      }
      if (d.align === 'right') {
        let c = col - 1, extra = 0;
        while (c >= 0 && extra + rect.w < tw + pad * 2 && !hasContent(row, c)) { extra += layout.cols.size(c) * zoom; c--; }
        clipX -= extra; clipW += extra;
      }
    }
    ctx.save();
    ctx.beginPath(); ctx.rect(clipX, rect.y, clipW - 1, rect.h - 1); ctx.clip();
    ctx.font = font;
    ctx.fillStyle = d.color ? colorOf(d.color) : st.font?.color ?? (theme.dark && !d.fill ? theme.ink : '#000000');
    ctx.textBaseline = 'alphabetic';
    const lines = wrap ? this.wrapLines(font, text, Math.max(1, rect.w - pad * 2 - indent)) : text.split('\n').slice(0, 1);
    const ascent = fontPx * 0.78;
    const totalH = lines.length * lineH;
    const va = st.align?.v ?? 'bottom';
    let ty = va === 'top' ? rect.y + pad + ascent : va === 'center' || va === 'justify' || va === 'distributed' ? rect.y + (rect.h - totalH) / 2 + ascent + (lineH - fontPx) / 2 : rect.y + rect.h - pad - (totalH - lineH) - (lineH - ascent) + 1;
    const rot = st.align?.rotation;
    if (rot !== undefined && rot !== 0) {
      const angle = rot === 255 ? 0 : rot <= 90 ? -rot : rot - 90;
      ctx.translate(rect.x + rect.w / 2, rect.y + rect.h / 2);
      if (rot === 255) {
        const chars = [...text];
        chars.forEach((ch, i) => { const cw = this.measure(font, ch); ctx.fillText(ch, -cw / 2, -((chars.length - 1) * lineH) / 2 + i * lineH + ascent / 2); });
      } else {
        ctx.rotate((angle * Math.PI) / 180);
        ctx.fillText(text, -tw / 2, ascent / 2);
      }
      ctx.restore();
      return;
    }
    for (const ln of lines) {
      const lw = this.measure(font, ln);
      let tx = d.align === 'right' ? x + w - pad - lw - indent : d.align === 'center' ? x + (w - lw) / 2 : x + pad + indent + iconPad;
      if (d.align === 'right' && iconPad) tx = Math.max(tx, rect.x + iconPad);
      if (st.align?.h === 'centerContinuous') tx = rect.x + (rect.w - lw) / 2;
      ctx.fillText(ln, tx, ty);
      if (st.font?.underline) {
        ctx.fillRect(tx, ty + 2 * zoom, lw, Math.max(1, zoom));
        if (st.font.underline.startsWith('double')) ctx.fillRect(tx, ty + 4 * zoom, lw, Math.max(1, zoom));
      }
      if (st.font?.strike) ctx.fillRect(tx, ty - fontPx * 0.28, lw, Math.max(1, zoom));
      ty += lineH;
    }
    ctx.restore();
  }

  private wrapLines(font: string, text: string, width: number): string[] {
    const out: string[] = [];
    for (const para of text.split('\n')) {
      const words = para.split(/(\s+)/);
      let line = '';
      for (const w of words) {
        const next = line + w;
        if (line && this.measure(font, next.trimEnd()) > width) {
          out.push(line.trimEnd());
          line = w.trimStart();
          while (this.measure(font, line) > width && line.length > 1) {
            let k = line.length;
            while (k > 1 && this.measure(font, line.slice(0, k)) > width) k--;
            out.push(line.slice(0, k)); line = line.slice(k);
          }
        } else line = next;
      }
      out.push(line);
    }
    return out;
  }

  private drawCheckbox(rect: { x: number; y: number; w: number; h: number }, on: boolean, zoom: number, theme: Theme): void {
    const ctx = this.ctx;
    const size = Math.max(8, Math.min(14 * zoom, rect.h - 4, rect.w - 4));
    const x = Math.round(rect.x + (rect.w - size) / 2) + 0.5, y = Math.round(rect.y + (rect.h - size) / 2) + 0.5;
    ctx.save();
    ctx.lineWidth = Math.max(1, 1.4 * zoom);
    ctx.beginPath(); ctx.roundRect(x, y, size, size, 2.5 * zoom);
    if (on) { ctx.fillStyle = theme.accent; ctx.fill(); }
    else { ctx.fillStyle = theme.bg; ctx.fill(); ctx.strokeStyle = theme.dark ? '#7c8c83' : '#6b7a71'; ctx.stroke(); }
    if (on) {
      ctx.strokeStyle = theme.dark ? '#06200f' : '#ffffff';
      ctx.lineWidth = Math.max(1.4, 1.9 * zoom); ctx.lineCap = 'round'; ctx.lineJoin = 'round';
      ctx.beginPath(); ctx.moveTo(x + size * 0.24, y + size * 0.52); ctx.lineTo(x + size * 0.43, y + size * 0.7); ctx.lineTo(x + size * 0.77, y + size * 0.32); ctx.stroke();
    }
    ctx.restore();
  }

  private drawBar(rect: { x: number; y: number; w: number; h: number }, bar: NonNullable<CfResult['bar']>, zoom: number): void {
    const ctx = this.ctx;
    const pad = 2 * zoom;
    const fullW = rect.w - pad * 2;
    let x = rect.x + pad, w = fullW * Math.min(1, bar.fraction);
    if (bar.axis !== undefined) {
      const ax = rect.x + pad + fullW * bar.axis;
      x = bar.negative ? ax - w : ax;
      ctx.strokeStyle = '#000'; ctx.setLineDash([2, 2]); ctx.beginPath(); ctx.moveTo(ax, rect.y); ctx.lineTo(ax, rect.y + rect.h); ctx.stroke(); ctx.setLineDash([]);
    }
    const y = rect.y + pad, h = rect.h - pad * 2;
    if (bar.gradient) {
      const g = ctx.createLinearGradient(x, 0, x + w, 0);
      g.addColorStop(0, bar.color); g.addColorStop(1, '#ffffff');
      ctx.fillStyle = g;
    } else ctx.fillStyle = bar.color;
    ctx.fillRect(x, y, w, h);
    ctx.strokeStyle = bar.color; ctx.lineWidth = 1; ctx.strokeRect(x + 0.5, y + 0.5, w - 1, h - 1);
  }

  private drawIcon(rect: { x: number; y: number; w: number; h: number }, icon: { set: string; index: number }, zoom: number): void {
    const ctx = this.ctx;
    const s = 12 * zoom;
    const x = rect.x + 3 * zoom, y = rect.y + (rect.h - s) / 2;
    const n = parseInt(icon.set, 10) || 3;
    const t = n > 1 ? icon.index / (n - 1) : 1;
    const color = t >= 0.99 ? '#1e9e48' : t <= 0.01 ? '#d93025' : t >= 0.5 ? '#f2b400' : '#f28b00';
    ctx.fillStyle = /Gray/.test(icon.set) ? '#808080' : color;
    ctx.beginPath();
    if (/Arrow/.test(icon.set)) {
      const cx = x + s / 2, cy = y + s / 2;
      const ang = (-90 + 180 * (1 - t)) * (Math.PI / 180);
      ctx.save(); ctx.translate(cx, cy); ctx.rotate(ang + Math.PI / 2);
      ctx.moveTo(0, -s / 2); ctx.lineTo(s / 2, 0); ctx.lineTo(s / 5, 0); ctx.lineTo(s / 5, s / 2); ctx.lineTo(-s / 5, s / 2); ctx.lineTo(-s / 5, 0); ctx.lineTo(-s / 2, 0); ctx.closePath();
      ctx.fill(); ctx.restore(); return;
    }
    if (/Flag/.test(icon.set)) { ctx.fillRect(x + 1, y, 1.5 * zoom, s); ctx.moveTo(x + 2.5 * zoom, y); ctx.lineTo(x + s, y + s / 4); ctx.lineTo(x + 2.5 * zoom, y + s / 2); ctx.fill(); return; }
    if (/Rating|Quarters|Boxes|Star/.test(icon.set)) {
      ctx.strokeStyle = '#4472c4'; ctx.fillStyle = '#4472c4';
      ctx.arc(x + s / 2, y + s / 2, s / 2 - 1, 0, Math.PI * 2); ctx.stroke();
      ctx.beginPath(); ctx.moveTo(x + s / 2, y + s / 2); ctx.arc(x + s / 2, y + s / 2, s / 2 - 1, -Math.PI / 2, -Math.PI / 2 + Math.PI * 2 * t); ctx.closePath(); ctx.fill(); return;
    }
    if (/Triangle/.test(icon.set)) { if (t > 0.5) { ctx.moveTo(x + s / 2, y); ctx.lineTo(x + s, y + s); ctx.lineTo(x, y + s); } else if (t < 0.5) { ctx.moveTo(x, y); ctx.lineTo(x + s, y); ctx.lineTo(x + s / 2, y + s); } else ctx.rect(x, y + s / 3, s, s / 3); ctx.fill(); return; }
    ctx.arc(x + s / 2, y + s / 2, s / 2, 0, Math.PI * 2);
    ctx.fill();
  }

  private dropButton(x: number, y: number, s: number, theme: Theme, kind: 'arrow' | 'filter'): void {
    const ctx = this.ctx;
    ctx.fillStyle = theme.dark ? '#2a2f3d' : '#ffffff';
    ctx.strokeStyle = theme.headerLine;
    ctx.fillRect(x, y, s, s); ctx.strokeRect(x + 0.5, y + 0.5, s - 1, s - 1);
    ctx.fillStyle = theme.headerInk;
    ctx.beginPath();
    if (kind === 'filter') { ctx.moveTo(x + s * 0.2, y + s * 0.25); ctx.lineTo(x + s * 0.8, y + s * 0.25); ctx.lineTo(x + s * 0.56, y + s * 0.55); ctx.lineTo(x + s * 0.56, y + s * 0.8); ctx.lineTo(x + s * 0.44, y + s * 0.72); ctx.lineTo(x + s * 0.44, y + s * 0.55); }
    else { ctx.moveTo(x + s * 0.3, y + s * 0.4); ctx.lineTo(x + s * 0.7, y + s * 0.4); ctx.lineTo(x + s * 0.5, y + s * 0.65); }
    ctx.closePath(); ctx.fill();
  }

  private drawHeaders(store: GridStore, vp: Viewport, theme: Theme, panes: Pane[]): void {
    if (!store.sheet.view.showHeaders) return;
    const ctx = this.ctx;
    const { layout, zoom } = vp;
    const sel = store.sel.ranges;
    const fullCol = (c: number) => sel.some((r) => r.r1 === 0 && r.r2 === MAX_ROWS - 1 && c >= r.c1 && c <= r.c2);
    const fullRow = (r: number) => sel.some((x) => x.c1 === 0 && x.c2 === MAX_COLS - 1 && r >= x.r1 && r <= x.r2);
    const colSel = (c: number) => sel.some((r) => c >= r.c1 && c <= r.c2);
    const rowSel = (r: number) => sel.some((x) => r >= x.r1 && r <= x.r2);
    ctx.font = `400 ${(11 * zoom).toFixed(1)}px "Hanken Grotesk", system-ui, sans-serif`;
    ctx.textBaseline = 'middle';
    ctx.textAlign = 'center';
    // Column header strip.
    ctx.fillStyle = theme.header; ctx.fillRect(vp.headerW, 0, vp.width, vp.headerH);
    for (const p of panes.filter((x) => x.y === vp.headerH + layout.rows.start(vp.frozenRows) * zoom || x.r2 < vp.frozenRows ? true : true)) {
      if (p.r2 < vp.frozenRows && panes.some((q) => q !== p && q.c1 === p.c1 && q.r2 >= vp.frozenRows)) continue;
      ctx.save(); ctx.beginPath(); ctx.rect(p.x, 0, p.w, vp.headerH); ctx.clip();
      for (let c = p.c1; c <= p.c2; c++) {
        const w = layout.cols.size(c) * zoom;
        if (!w) continue;
        const x = p.x + (layout.cols.start(c) - p.scrollX) * zoom;
        const on = colSel(c), full = fullCol(c);
        if (on) { ctx.fillStyle = full ? theme.accent : theme.headerSel; ctx.fillRect(x, 0, w, vp.headerH); }
        ctx.fillStyle = full ? '#ffffff' : on ? theme.headerSelInk : theme.headerInk;
        ctx.fillText(colName(c), x + w / 2, vp.headerH / 2 + 0.5);
        ctx.fillStyle = theme.headerLine; ctx.fillRect(Math.round(x + w) - 1, 0, 1, vp.headerH);
        if (on && !full) { ctx.fillStyle = theme.accent; ctx.fillRect(x, vp.headerH - 2, w, 2); }
      }
      ctx.restore();
    }
    // Row header strip.
    ctx.fillStyle = theme.header; ctx.fillRect(0, vp.headerH, vp.headerW, vp.height);
    for (const p of panes) {
      if (p.c2 < vp.frozenCols && panes.some((q) => q !== p && q.r1 === p.r1 && q.c2 >= vp.frozenCols)) continue;
      ctx.save(); ctx.beginPath(); ctx.rect(0, p.y, vp.headerW, p.h); ctx.clip();
      for (let r = p.r1; r <= p.r2; r++) {
        const h = layout.rows.size(r) * zoom;
        if (!h) continue;
        const y = p.y + (layout.rows.start(r) - p.scrollY) * zoom;
        const on = rowSel(r), full = fullRow(r);
        if (on) { ctx.fillStyle = full ? theme.accent : theme.headerSel; ctx.fillRect(0, y, vp.headerW, h); }
        ctx.fillStyle = full ? '#ffffff' : on ? theme.headerSelInk : theme.headerInk;
        ctx.fillText(String(r + 1), vp.headerW / 2, y + h / 2 + 0.5);
        ctx.fillStyle = theme.headerLine; ctx.fillRect(0, Math.round(y + h) - 1, vp.headerW, 1);
        if (on && !full) { ctx.fillStyle = theme.accent; ctx.fillRect(vp.headerW - 2, y, 2, h); }
      }
      ctx.restore();
    }
    ctx.fillStyle = theme.headerLine;
    ctx.fillRect(0, vp.headerH - 1, vp.width, 1);
    ctx.fillRect(vp.headerW - 1, 0, 1, vp.height);
    ctx.textAlign = 'start';
  }

  private drawSelection(store: GridStore, vp: Viewport, theme: Theme, p: Pane, ov: Overlay): void {
    const ctx = this.ctx;
    const sel = store.sel;
    const editing = !!ov.editing;
    // Formula references while editing (each in its own color).
    for (const r of ov.refs) {
      const rect = rectIn(p, vp, r.range);
      ctx.strokeStyle = r.color; ctx.lineWidth = 2;
      ctx.strokeRect(rect.x + 0.5, rect.y + 0.5, rect.w - 2, rect.h - 2);
      ctx.fillStyle = r.color + '22'; ctx.fillRect(rect.x, rect.y, rect.w, rect.h);
    }
    // Spill outline around the active cell's dynamic array.
    const sp = store.engine.spillAnchorOf(store.sheetIndex, sel.active.row, sel.active.col);
    if (sp) {
      const range = store.engine.spillRange(store.sheetIndex, sp.row, sp.col);
      if (range && (range.r2 > range.r1 || range.c2 > range.c1)) {
        const rect = rectIn(p, vp, { r1: range.r1, c1: range.c1, r2: range.r2, c2: range.c2 });
        ctx.strokeStyle = theme.accent; ctx.setLineDash([3, 2]); ctx.lineWidth = 1;
        ctx.strokeRect(rect.x + 0.5, rect.y + 0.5, rect.w - 1, rect.h - 1); ctx.setLineDash([]);
      }
    }
    if (editing) return;
    // Shade the selection except the active cell, which stays clear.
    const am = store.mergeAt(sel.active.row, sel.active.col) ?? { r1: sel.active.row, c1: sel.active.col, r2: sel.active.row, c2: sel.active.col };
    const ar = rectIn(p, vp, am);
    ctx.fillStyle = theme.selFill;
    for (const r of sel.ranges) {
      const rect = rectIn(p, vp, clampVisible(r, p));
      const hasActive = inRange(r, am.r1, am.c1);
      if (!hasActive) { ctx.fillRect(rect.x, rect.y, rect.w, rect.h); continue; }
      ctx.fillRect(rect.x, rect.y, rect.w, Math.max(0, ar.y - rect.y));
      ctx.fillRect(rect.x, ar.y + ar.h, rect.w, Math.max(0, rect.y + rect.h - ar.y - ar.h));
      ctx.fillRect(rect.x, ar.y, Math.max(0, ar.x - rect.x), ar.h);
      ctx.fillRect(ar.x + ar.w, ar.y, Math.max(0, rect.x + rect.w - ar.x - ar.w), ar.h);
    }
    const last = sel.ranges.at(-1)!;
    const lr = rectIn(p, vp, last);
    ctx.strokeStyle = theme.accent; ctx.lineWidth = 2;
    ctx.strokeRect(lr.x, lr.y, lr.w - 1, lr.h - 1);
    if (sel.ranges.length === 1 && (last.r1 !== last.r2 || last.c1 !== last.c2)) {
      ctx.lineWidth = 1; ctx.strokeStyle = theme.accent;
      ctx.strokeRect(ar.x + 1.5, ar.y + 1.5, ar.w - 4, ar.h - 4);
    }
    // Fill handle.
    if (sel.ranges.length === 1) {
      ctx.fillStyle = theme.accent;
      ctx.fillRect(lr.x + lr.w - 4, lr.y + lr.h - 4, 7, 7);
      ctx.fillStyle = theme.bg;
      ctx.fillRect(lr.x + lr.w - 5, lr.y + lr.h - 5, 1, 9); ctx.fillRect(lr.x + lr.w - 5, lr.y + lr.h - 5, 9, 1);
    }
    if (ov.fillPreview) {
      const fr = rectIn(p, vp, ov.fillPreview);
      ctx.strokeStyle = theme.headerInk; ctx.setLineDash([4, 3]); ctx.lineWidth = 1;
      ctx.strokeRect(fr.x + 0.5, fr.y + 0.5, fr.w - 2, fr.h - 2); ctx.setLineDash([]);
    }
    if (ov.dragRange) {
      const dr = rectIn(p, vp, ov.dragRange);
      ctx.strokeStyle = theme.headerInk; ctx.lineWidth = 2; ctx.setLineDash([4, 3]);
      ctx.strokeRect(dr.x + 1, dr.y + 1, dr.w - 3, dr.h - 3); ctx.setLineDash([]);
    }
    const mq = store.marquee;
    if (mq && mq.sheet === store.sheetIndex) {
      const mr = rectIn(p, vp, mq.range);
      ctx.strokeStyle = theme.accent; ctx.lineWidth = 2; ctx.setLineDash([5, 4]); ctx.lineDashOffset = -ov.marqueePhase;
      ctx.strokeRect(mr.x + 1, mr.y + 1, mr.w - 3, mr.h - 3); ctx.setLineDash([]); ctx.lineDashOffset = 0;
    }
  }
}

const clampVisible = (r: CellRange, p: Pane): CellRange => ({ r1: Math.max(r.r1, p.r1 - 1), c1: Math.max(r.c1, p.c1 - 1), r2: Math.min(r.r2, p.r2 + 1), c2: Math.min(r.c2, p.c2 + 1) });

const mergeShallow = (a: CellStyle, b: CellStyle): CellStyle => ({ ...a, ...b, font: { ...a.font, ...b.font }, fill: b.fill ?? a.fill, border: { ...a.border, ...b.border } });

const NAMED: Record<string, string> = { red: '#ff0000', blue: '#0000ff', green: '#00a000', black: '#000000', white: '#ffffff', yellow: '#c8a800', cyan: '#00a0a0', magenta: '#c000c0' };
const colorOf = (c: string): string => NAMED[c.toLowerCase()] ?? c;
