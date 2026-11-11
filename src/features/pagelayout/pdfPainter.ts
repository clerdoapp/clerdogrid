import type { PdfPage, PathCmd, StrokeStyle } from '../../formats/pdf';
import type { FontBook, ShapedText } from '../../formats/pdf-text';
import { grayOf, type Painter, type TextStyle } from './painter';

/** Draws text that no bundled font covers as a picture (browser only). */
export type TextRasterizer = (text: string, style: TextStyle) => { png: Uint8Array; width: number; ascent: number; height: number } | null;

const missing = (t: ShapedText) => t.runs.some((r) => r.missing && r.text.trim());

export class PdfPainter implements Painter {
  mono = false;
  constructor(readonly page: PdfPage, readonly fonts: FontBook, private readonly opts: { raster?: TextRasterizer; images?: Map<Uint8Array, Uint8Array> } = {}) {}

  private c(color: string): string { return this.mono ? grayOf(color) : color; }

  save(): void { this.page.save(); }
  restore(): void { this.page.restore(); }
  clip(x: number, y: number, w: number, h: number): void { this.page.clipRect(x, y, w, h); }
  transform(a: number, b: number, c: number, d: number, e: number, f: number): void { this.page.transform(a, b, c, d, e, f); }

  rect(x: number, y: number, w: number, h: number, fill: string, alpha = 1): void {
    if (w <= 0 || h <= 0) return;
    if (alpha < 1) this.page.path([['M', x, y], ['L', x + w, y], ['L', x + w, y + h], ['L', x, y + h], ['Z']], this.c(fill).slice(0, 7), undefined, false, alpha);
    else this.page.rect(x, y, w, h, this.c(fill));
  }

  gradient(x: number, y: number, w: number, h: number, angle: number, stops: { pos: number; color: string }[]): void {
    this.page.gradientRect(x, y, w, h, angle, stops.map((s) => ({ pos: s.pos, color: this.c(s.color) })));
  }

  path(cmds: PathCmd[], fill?: string, stroke?: StrokeStyle, o: { evenOdd?: boolean; fillAlpha?: number; strokeAlpha?: number } = {}): void {
    this.page.path(cmds, fill && this.c(fill), stroke && { ...stroke, color: this.c(stroke.color) }, o.evenOdd, o.fillAlpha ?? 1, o.strokeAlpha ?? 1);
  }

  text(text: string, x: number, y: number, style: TextStyle, rtl?: boolean): number {
    if (!text) return 0;
    const req = { family: style.family, bold: style.bold, italic: style.italic };
    const shaped = this.fonts.shape(text, req, rtl);
    const width = (shaped.width * style.size) / 1000;
    const color = this.c(style.color);
    const f = shaped.runs.find((x) => x.font && !x.missing)?.font ?? shaped.runs.find((x) => x.font)?.font;
    const synth = { bold: !!style.bold && !!f && !f.bold, skew: !!style.italic && !!f && !f.italic };
    if (!this.opts.raster || !missing(shaped)) {
      this.page.text(shaped, x, y, style.size, color, synth);
    } else {
      // Runs no bundled font covers are drawn by the system as pictures; the rest stays real text.
      let cx = x;
      for (const run of shaped.runs) {
        const w = (run.width * style.size) / 1000;
        const r = run.missing && run.text.trim() ? this.opts.raster(run.text, { ...style, color, underline: undefined, strike: undefined }) : null;
        if (r) {
          const scale = style.size / 100;
          this.page.image(r.png, cx, y - r.ascent * scale, w, r.height * scale, run.text);
        } else this.page.text({ runs: [run], width: run.width, rtl: run.rtl }, cx, y, style.size, color, synth);
        cx += w;
      }
    }
    this.decorate(x, y, width, style, color);
    return width;
  }

  private decorate(x: number, y: number, w: number, style: TextStyle, color: string): void {
    const t = Math.max(0.4, style.size * 0.06);
    if (style.underline) {
      this.page.rect(x, y + style.size * 0.12, w, t, color);
      if (style.underline === 'double') this.page.rect(x, y + style.size * 0.12 + t * 2, w, t, color);
    }
    if (style.strike) this.page.rect(x, y - style.size * 0.3, w, t, color);
  }

  image(data: Uint8Array, _mime: string, x: number, y: number, w: number, h: number): void {
    const bytes = this.opts.images?.get(data) ?? data;
    if (!this.page.image(bytes, x, y, w, h, undefined, this.mono)) {
      this.page.path([['M', x, y], ['L', x + w, y], ['L', x + w, y + h], ['L', x, y + h], ['Z']], '#f2f2f2', { color: '#bfbfbf', width: 0.5 });
    }
  }

  link(x: number, y: number, w: number, h: number, uri: string): void { this.page.link(x, y, w, h, uri); }
}
