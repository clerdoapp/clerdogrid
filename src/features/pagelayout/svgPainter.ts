// Builds an SVG document for one page: used by print preview, the page layout view and printing.
import type { FontBook } from '../../formats/pdf-text';
import { fontFamily } from '../../ui/render';
import { grayOf, type Painter, type PathCmd, type StrokeStyle, type TextStyle } from './painter';

const esc = (s: string) => s.replace(/[&<>"]/g, (c) => (c === '&' ? '&amp;' : c === '<' ? '&lt;' : c === '>' ? '&gt;' : '&quot;'));
const n = (v: number) => (Number.isFinite(v) ? String(Math.round(v * 100) / 100) : '0');
const RTL = /[֐-ࣿיִ-﷿ﹰ-﻿]/;

const colorAttr = (name: string, c: string): string => {
  if (c.length === 9) return ` ${name}="${c.slice(0, 7)}" ${name}-opacity="${n(parseInt(c.slice(7), 16) / 255)}"`;
  return ` ${name}="${c}"`;
};

let seq = 0;
const toBase64 = (b: Uint8Array): string => {
  let s = '';
  for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode(...b.subarray(i, i + 0x8000));
  return btoa(s);
};

export const svgPath = (cmds: PathCmd[]): string => cmds.map((c) => (c[0] === 'Z' ? 'Z' : c[0] + c.slice(1).map((v) => n(v as number)).join(' '))).join('');

export class SvgPainter implements Painter {
  mono = false;
  private out: string[] = [];
  private stack: number[] = [];
  private open = 0;
  private defs: string[] = [];
  private id = `p${(++seq).toString(36)}`;
  private clips = 0;

  constructor(readonly fonts: FontBook, private readonly opts: { imageUrl?: (data: Uint8Array, mime: string) => string } = {}) {}

  private c(color: string): string { return this.mono ? grayOf(color) : color; }

  save(): void { this.stack.push(this.open); this.open = 0; }
  restore(): void { this.out.push('</g>'.repeat(this.open)); this.open = this.stack.pop() ?? 0; }
  clip(x: number, y: number, w: number, h: number): void {
    const id = `${this.id}c${++this.clips}`;
    this.defs.push(`<clipPath id="${id}"><rect x="${n(x)}" y="${n(y)}" width="${n(Math.max(0, w))}" height="${n(Math.max(0, h))}"/></clipPath>`);
    this.out.push(`<g clip-path="url(#${id})">`);
    this.open++;
  }
  transform(a: number, b: number, c: number, d: number, e: number, f: number): void {
    this.out.push(`<g transform="matrix(${[a, b, c, d, e, f].map(n).join(' ')})">`);
    this.open++;
  }

  rect(x: number, y: number, w: number, h: number, fill: string, alpha = 1): void {
    if (w <= 0 || h <= 0) return;
    this.out.push(`<rect x="${n(x)}" y="${n(y)}" width="${n(w)}" height="${n(h)}"${colorAttr('fill', this.c(fill))}${alpha < 1 ? ` fill-opacity="${n(alpha)}"` : ''}/>`);
  }

  gradient(x: number, y: number, w: number, h: number, angle: number, stops: { pos: number; color: string }[]): void {
    const id = `${this.id}g${this.defs.length}`;
    const rad = (angle * Math.PI) / 180;
    this.defs.push(`<linearGradient id="${id}" gradientUnits="userSpaceOnUse" x1="${n(x)}" y1="${n(y)}" x2="${n(x + Math.cos(rad) * w)}" y2="${n(y + Math.sin(rad) * h)}">${stops.map((s) => `<stop offset="${n(s.pos)}" stop-color="${this.c(s.color).slice(0, 7)}"/>`).join('')}</linearGradient>`);
    this.out.push(`<rect x="${n(x)}" y="${n(y)}" width="${n(w)}" height="${n(h)}" fill="url(#${id})"/>`);
  }

  path(cmds: PathCmd[], fill?: string, stroke?: StrokeStyle, o: { evenOdd?: boolean; fillAlpha?: number; strokeAlpha?: number } = {}): void {
    if (!cmds.length) return;
    let a = ` d="${svgPath(cmds)}"`;
    a += fill ? colorAttr('fill', this.c(fill)) : ' fill="none"';
    if (fill && (o.fillAlpha ?? 1) < 1) a += ` fill-opacity="${n(o.fillAlpha!)}"`;
    if (o.evenOdd) a += ' fill-rule="evenodd"';
    if (stroke) {
      a += `${colorAttr('stroke', this.c(stroke.color))} stroke-width="${n(stroke.width)}"`;
      if (stroke.dash?.length) a += ` stroke-dasharray="${stroke.dash.map(n).join(' ')}"`;
      if (stroke.cap) a += ` stroke-linecap="${stroke.cap}"`;
      if (stroke.join) a += ` stroke-linejoin="${stroke.join}"`;
      if ((o.strokeAlpha ?? 1) < 1) a += ` stroke-opacity="${n(o.strokeAlpha!)}"`;
    }
    this.out.push(`<path${a}/>`);
  }

  text(text: string, x: number, y: number, style: TextStyle, rtl?: boolean): number {
    if (!text) return 0;
    const req = { family: style.family, bold: style.bold, italic: style.italic };
    const shaped = this.fonts.shape(text, req, rtl ?? (RTL.test(text) ? undefined : false));
    const k = style.size / 1000;
    const fam = esc(fontFamily(style.family));
    const deco = [style.underline ? 'underline' : '', style.strike ? 'line-through' : ''].filter(Boolean).join(' ');
    const attrs = ` font-family="${fam}" font-size="${n(style.size)}"${style.bold ? ' font-weight="700"' : ''}${style.italic ? ' font-style="italic"' : ''}${colorAttr('fill', this.c(style.color))}`
      + `${deco ? ` text-decoration="${deco}"` : ''}${style.underline === 'double' ? ' style="text-decoration-style:double"' : ''} xml:space="preserve"`;
    // One element per direction run, placed where the PDF layout puts it, so both outputs match.
    let cx = x;
    for (const run of shaped.runs) {
      const w = run.width * k;
      if (run.text.trim() || deco) {
        this.out.push(run.rtl
          ? `<text x="${n(cx + w)}" y="${n(y)}" direction="rtl" unicode-bidi="embed"${attrs}>${esc(run.text)}</text>`
          : `<text x="${n(cx)}" y="${n(y)}"${attrs}>${esc(run.text)}</text>`);
      }
      cx += w;
    }
    return shaped.width * k;
  }

  image(data: Uint8Array, mime: string, x: number, y: number, w: number, h: number): void {
    const url = this.opts.imageUrl ? this.opts.imageUrl(data, mime) : `data:${mime};base64,${toBase64(data)}`;
    this.out.push(`<image x="${n(x)}" y="${n(y)}" width="${n(w)}" height="${n(h)}" preserveAspectRatio="none" href="${esc(url)}"${this.mono ? ` filter="url(#${this.grayFilter()})"` : ''}/>`);
  }

  private gray = '';
  private grayFilter(): string {
    if (!this.gray) { this.gray = `${this.id}gray`; this.defs.push(`<filter id="${this.gray}"><feColorMatrix type="saturate" values="0"/></filter>`); }
    return this.gray;
  }

  link(x: number, y: number, w: number, h: number, uri: string): void {
    if (!/^(https:|mailto:)/i.test(uri)) return;
    this.out.push(`<a href="${esc(uri)}" target="_blank"><rect x="${n(x)}" y="${n(y)}" width="${n(w)}" height="${n(h)}" fill="transparent"/></a>`);
  }

  /** The finished page, sized in points (or to `cssW` × `cssH` when given). */
  svg(width: number, height: number, cssW?: string, cssH?: string): string {
    while (this.stack.length) this.restore();
    this.out.push('</g>'.repeat(this.open));
    this.open = 0;
    return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${n(width)} ${n(height)}" width="${cssW ?? `${n(width)}pt`}" height="${cssH ?? `${n(height)}pt`}"><defs>${this.defs.join('')}</defs><rect width="100%" height="100%" fill="#ffffff"/>${this.out.join('')}</svg>`;
  }
}
