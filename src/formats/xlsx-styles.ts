// styles.xml and theme colours (ECMA-376 Part 1, 18.8 Styles; 20.1.6 Theme).
import type { BorderSide, BorderStyle, CellStyle, HAlign, Underline, VAlign } from '../model/workbook';
import { DEFAULT_FONT } from '../model/workbook';
import { INDEXED_COLORS, builtinFormat, builtinId } from './numfmt';
import { attr, boolAttr, el, esc, find, kid, kids, numAttr, parseRoot, type XNode } from './xml';

// ---------------------------------------------------------------- colours

/** Office 2013+ default theme, in SpreadsheetML theme-index order (lt1, dk1, lt2, dk2, accent1-6, hlink, folHlink). */
export const DEFAULT_THEME = ['#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#0563C1', '#954F72'];

export const readTheme = (xml: string | undefined): string[] => {
  if (!xml) return DEFAULT_THEME.slice();
  const scheme = find(parseRoot(xml), 'clrScheme');
  if (!scheme) return DEFAULT_THEME.slice();
  const col = (name: string, i: number): string => {
    const n = kid(scheme, name);
    const c = n?.children[0];
    const v = c?.name === 'sysClr' ? attr(c, 'lastClr') : attr(c, 'val');
    return v && /^[0-9A-Fa-f]{6}$/.test(v) ? '#' + v.toUpperCase() : DEFAULT_THEME[i]!;
  };
  return [col('lt1', 0), col('dk1', 1), col('lt2', 2), col('dk2', 3), col('accent1', 4), col('accent2', 5), col('accent3', 6),
    col('accent4', 7), col('accent5', 8), col('accent6', 9), col('hlink', 10), col('folHlink', 11)];
};

const rgbToHls = (r: number, g: number, b: number): [number, number, number] => {
  const mx = Math.max(r, g, b), mn = Math.min(r, g, b);
  const l = (mx + mn) / 2;
  if (mx === mn) return [0, l, 0];
  const d = mx - mn;
  const s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
  const h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, l, s];
};
const hlsToRgb = (h: number, l: number, s: number): [number, number, number] => {
  if (s === 0) return [l, l, l];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const f = (t0: number): number => {
    let t = t0;
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    return t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p;
  };
  return [f(h + 1 / 3), f(h), f(h - 1 / 3)];
};
const hex2 = (x: number): string => Math.round(Math.max(0, Math.min(1, x)) * 255).toString(16).padStart(2, '0').toUpperCase();

/** spreadsheet's tint: darken (tint < 0) or lighten (tint > 0) the HLS luminance. */
export const applyTint = (hex: string, tint: number): string => {
  if (!tint) return hex;
  const n = parseInt(hex.slice(1, 7), 16);
  const [h, l0, s] = rgbToHls(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255);
  const l = tint < 0 ? l0 * (1 + tint) : l0 * (1 - tint) + tint;
  const [r, g, b] = hlsToRgb(h, l, s);
  return '#' + hex2(r) + hex2(g) + hex2(b);
};

/** CT_Color → "#RRGGBB" (undefined for auto / system colours). */
export const readColor = (n: XNode | undefined, theme: readonly string[], indexed: readonly string[] = INDEXED_COLORS): string | undefined => {
  if (!n || boolAttr(n, 'auto')) return undefined;
  let c: string | undefined;
  const rgb = attr(n, 'rgb');
  const th = numAttr(n, 'theme');
  const ix = numAttr(n, 'indexed');
  if (rgb && /^[0-9A-Fa-f]{6,8}$/.test(rgb)) c = '#' + (rgb.length === 8 ? rgb.slice(2) : rgb).toUpperCase();
  else if (th !== undefined) c = theme[th];
  else if (ix !== undefined) c = ix >= 64 ? undefined : indexed[ix];
  if (!c) return undefined;
  const tint = numAttr(n, 'tint');
  return tint ? applyTint(c, tint) : c;
};

/** "#RRGGBB" / "#RRGGBBAA" → ARGB as SpreadsheetML wants it. */
export const argb = (c: string): string => {
  const h = c.replace('#', '').toUpperCase();
  if (h.length === 8) return h.slice(6) + h.slice(0, 6);
  if (h.length === 3) return 'FF' + h.split('').map((x) => x + x).join('');
  return 'FF' + h.padStart(6, '0').slice(0, 6);
};

const colorEl = (name: string, c: string | undefined): string => (c ? el(name, { rgb: argb(c) }) : '');

// ---------------------------------------------------------------- reading

type Font = NonNullable<CellStyle['font']>;
type Fill = NonNullable<CellStyle['fill']>;
type Border = NonNullable<CellStyle['border']>;

export interface StylesRead {
  /** Converted cellXfs. */
  xfs: CellStyle[];
  dxfs: CellStyle[];
}

const valAttr = (n: XNode, name: string): string | undefined => {
  const k = kid(n, name);
  return k ? attr(k, 'val') ?? '' : undefined;
};
const flag = (n: XNode, name: string): boolean | undefined => {
  const k = kid(n, name);
  if (!k) return undefined;
  const v = attr(k, 'val');
  return v === undefined || v === '1' || v === 'true';
};

/** A <font> or run <rPr>. `dropDefaults` omits the workbook default name/size/text colour. */
export const readFont = (n: XNode, theme: readonly string[], indexed: readonly string[], dropDefaults: boolean): Font | undefined => {
  const f: Font = {};
  const name = valAttr(n, 'name') ?? valAttr(n, 'rFont');
  if (name && !(dropDefaults && name === DEFAULT_FONT.name)) f.name = name;
  const sz = valAttr(n, 'sz');
  if (sz && !(dropDefaults && +sz === DEFAULT_FONT.size)) f.size = +sz;
  if (flag(n, 'b')) f.bold = true;
  if (flag(n, 'i')) f.italic = true;
  if (flag(n, 'strike')) f.strike = true;
  const u = valAttr(n, 'u');
  if (u !== undefined && u !== 'none') f.underline = (u || 'single') as Underline;
  const va = valAttr(n, 'vertAlign');
  if (va === 'superscript' || va === 'subscript') f.vertAlign = va;
  const cn = kid(n, 'color');
  const isDefaultColor = cn && (attr(cn, 'theme') === '1' && !numAttr(cn, 'tint')) || (cn && (attr(cn, 'indexed') === '8' || attr(cn, 'indexed') === '64'));
  if (cn && !(dropDefaults && isDefaultColor)) {
    const c = readColor(cn, theme, indexed);
    if (c) f.color = c;
  }
  return Object.keys(f).length ? f : undefined;
};

const readFill = (n: XNode, theme: readonly string[], indexed: readonly string[], dxf: boolean): Fill | undefined => {
  const pf = kid(n, 'patternFill');
  if (pf) {
    const type = attr(pf, 'patternType') ?? (dxf ? 'solid' : 'none');
    if (type === 'none') return undefined;
    const fg = readColor(kid(pf, 'fgColor'), theme, indexed);
    const bg = readColor(kid(pf, 'bgColor'), theme, indexed);
    // In differential formats the solid colour is stored as bgColor.
    if (type === 'solid') { const c = dxf ? bg ?? fg : fg ?? bg; return c ? { color: c } : undefined; }
    const out: Fill = { pattern: type };
    if (bg) out.color = bg;
    if (fg) out.patternColor = fg;
    return out;
  }
  const gf = kid(n, 'gradientFill');
  if (gf) {
    const stops = kids(gf, 'stop').map((s) => ({ pos: numAttr(s, 'position') ?? 0, color: readColor(kid(s, 'color'), theme, indexed) ?? '#000000' }));
    return { gradient: { angle: numAttr(gf, 'degree') ?? 0, stops } };
  }
  return undefined;
};

const readSide = (n: XNode | undefined, theme: readonly string[], indexed: readonly string[]): BorderSide | undefined => {
  const style = attr(n, 'style');
  if (!n || !style || style === 'none') return undefined;
  const side: BorderSide = { style: style as BorderStyle };
  const c = readColor(kid(n, 'color'), theme, indexed);
  if (c) side.color = c;
  return side;
};

const readBorder = (n: XNode, theme: readonly string[], indexed: readonly string[]): Border | undefined => {
  const b: Border = {};
  for (const s of ['left', 'right', 'top', 'bottom'] as const) {
    const side = readSide(kid(n, s) ?? kid(n, s === 'left' ? 'start' : s === 'right' ? 'end' : s), theme, indexed);
    if (side) b[s] = side;
  }
  const diag = readSide(kid(n, 'diagonal'), theme, indexed);
  if (diag) {
    if (boolAttr(n, 'diagonalUp')) b.diagonalUp = diag;
    if (boolAttr(n, 'diagonalDown')) b.diagonalDown = { ...diag };
  }
  return Object.keys(b).length ? b : undefined;
};

const readAlign = (n: XNode | undefined): CellStyle['align'] => {
  if (!n) return undefined;
  const a: NonNullable<CellStyle['align']> = {};
  const h = attr(n, 'horizontal');
  if (h && h !== 'general') a.h = h as HAlign;
  const v = attr(n, 'vertical');
  if (v) a.v = v as VAlign;
  if (boolAttr(n, 'wrapText')) a.wrap = true;
  const ind = numAttr(n, 'indent');
  if (ind) a.indent = ind;
  const rot = numAttr(n, 'textRotation');
  if (rot) a.rotation = rot > 90 && rot <= 180 ? 90 - rot : rot;
  if (boolAttr(n, 'shrinkToFit')) a.shrink = true;
  if (attr(n, 'readingOrder') === '2') a.rtl = true;
  return Object.keys(a).length ? a : undefined;
};

const readProtection = (n: XNode | undefined): CellStyle['protection'] => {
  if (!n) return undefined;
  const p: NonNullable<CellStyle['protection']> = {};
  if (attr(n, 'locked') !== undefined) p.locked = boolAttr(n, 'locked');
  if (attr(n, 'hidden') !== undefined) p.hidden = boolAttr(n, 'hidden');
  return Object.keys(p).length ? p : undefined;
};

const clean = (s: CellStyle): CellStyle => {
  const o: CellStyle = {};
  for (const k of ['font', 'fill', 'border', 'align', 'numFmt', 'protection'] as const) if (s[k] !== undefined) (o as Record<string, unknown>)[k] = s[k];
  return o;
};

export const readStyles = (xml: string | undefined, theme: readonly string[]): StylesRead => {
  if (!xml) return { xfs: [{}], dxfs: [] };
  const root = parseRoot(xml);
  let indexed: readonly string[] = INDEXED_COLORS;
  const custom = kid(kid(root, 'colors'), 'indexedColors');
  if (custom) {
    const list = INDEXED_COLORS.slice();
    kids(custom, 'rgbColor').forEach((c, i) => { const v = attr(c, 'rgb'); if (v) list[i] = '#' + v.slice(-6).toUpperCase(); });
    indexed = list;
  }
  const fmts = new Map<number, string>();
  for (const n of kids(kid(root, 'numFmts'), 'numFmt')) fmts.set(numAttr(n, 'numFmtId') ?? -1, attr(n, 'formatCode') ?? 'General');
  const fmtCode = (id: number | undefined): string | undefined => {
    if (!id) return undefined;
    const c = fmts.get(id) ?? builtinFormat(id);
    return c && c !== 'General' ? c : undefined;
  };
  const fonts = kids(kid(root, 'fonts'), 'font').map((f) => readFont(f, theme, indexed, true));
  const fills = kids(kid(root, 'fills'), 'fill').map((f) => readFill(f, theme, indexed, false));
  const borders = kids(kid(root, 'borders'), 'border').map((b) => readBorder(b, theme, indexed));
  const xfs = kids(kid(root, 'cellXfs'), 'xf').map((x) =>
    clean({
      font: fonts[numAttr(x, 'fontId') ?? 0],
      fill: fills[numAttr(x, 'fillId') ?? 0],
      border: borders[numAttr(x, 'borderId') ?? 0],
      numFmt: fmtCode(numAttr(x, 'numFmtId')),
      align: readAlign(kid(x, 'alignment')),
      protection: readProtection(kid(x, 'protection')),
    }),
  );
  const dxfs = kids(kid(root, 'dxfs'), 'dxf').map((d) => {
    const nf = kid(d, 'numFmt');
    return clean({
      font: kid(d, 'font') ? readFont(kid(d, 'font')!, theme, indexed, false) : undefined,
      fill: kid(d, 'fill') ? readFill(kid(d, 'fill')!, theme, indexed, true) : undefined,
      border: kid(d, 'border') ? readBorder(kid(d, 'border')!, theme, indexed) : undefined,
      numFmt: nf ? attr(nf, 'formatCode') ?? fmtCode(numAttr(nf, 'numFmtId')) : undefined,
      align: readAlign(kid(d, 'alignment')),
      protection: readProtection(kid(d, 'protection')),
    });
  });
  return { xfs: xfs.length ? xfs : [{}], dxfs };
};

// ---------------------------------------------------------------- writing

export const fontXml = (f: Font | undefined, tag: 'font' | 'rPr', withDefaults: boolean): string => {
  const ff: Font = withDefaults ? { name: DEFAULT_FONT.name, size: DEFAULT_FONT.size, ...f } : { ...f };
  let s = '';
  if (ff.bold) s += '<b/>';
  if (ff.italic) s += '<i/>';
  if (ff.strike) s += '<strike/>';
  if (ff.underline) s += ff.underline === 'single' ? '<u/>' : el('u', { val: ff.underline });
  if (ff.vertAlign) s += el('vertAlign', { val: ff.vertAlign });
  if (ff.size !== undefined) s += el('sz', { val: ff.size });
  s += colorEl('color', ff.color);
  if (ff.name) s += el(tag === 'rPr' ? 'rFont' : 'name', { val: ff.name });
  if (withDefaults && tag === 'font') s += el('family', { val: 2 });
  return tag === 'rPr' && !s ? '' : `<${tag}>${s}</${tag}>`;
};

const fillXml = (f: Fill | undefined, dxf: boolean): string => {
  if (!f || (!f.color && !f.pattern && !f.gradient)) return '<fill><patternFill patternType="none"/></fill>';
  if (f.gradient) {
    const stops = f.gradient.stops.map((s) => `<stop position="${s.pos}">${colorEl('color', s.color)}</stop>`).join('');
    return `<fill><gradientFill degree="${f.gradient.angle}">${stops}</gradientFill></fill>`;
  }
  const type = f.pattern && f.pattern !== 'solid' ? f.pattern : 'solid';
  if (type === 'solid') {
    return dxf
      ? `<fill><patternFill patternType="solid">${colorEl('bgColor', f.color)}</patternFill></fill>`
      : `<fill><patternFill patternType="solid">${colorEl('fgColor', f.color)}<bgColor indexed="64"/></patternFill></fill>`;
  }
  return `<fill><patternFill patternType="${esc(type)}">${colorEl('fgColor', f.patternColor)}${colorEl('bgColor', f.color)}</patternFill></fill>`;
};

const sideXml = (name: string, s: BorderSide | undefined): string =>
  s ? `<${name} style="${s.style}">${colorEl('color', s.color)}</${name}>` : `<${name}/>`;

const borderXml = (b: Border | undefined): string => {
  const d = b?.diagonalUp ?? b?.diagonalDown;
  return `<border${b?.diagonalUp ? ' diagonalUp="1"' : ''}${b?.diagonalDown ? ' diagonalDown="1"' : ''}>${sideXml('left', b?.left)}${sideXml('right', b?.right)}${sideXml('top', b?.top)}${sideXml('bottom', b?.bottom)}${sideXml('diagonal', d)}</border>`;
};

const alignXml = (a: CellStyle['align']): string => {
  if (!a) return '';
  const rot = a.rotation === undefined ? undefined : a.rotation < 0 ? 90 - a.rotation : a.rotation;
  return el('alignment', {
    horizontal: a.h, vertical: a.v, wrapText: a.wrap || undefined, indent: a.indent || undefined,
    textRotation: rot || undefined, shrinkToFit: a.shrink || undefined, readingOrder: a.rtl ? 2 : undefined,
  });
};

const protXml = (p: CellStyle['protection']): string =>
  p ? el('protection', { locked: p.locked, hidden: p.hidden }) : '';

/** Builds styles.xml: fonts/fills/borders/numFmts are shared between equal entries; one cellXf per style. */
export class StyleWriter {
  private fonts: string[] = [];
  private fills: string[] = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
  private borders: string[] = [];
  private numFmts = new Map<string, number>();
  private xfs: string[] = [];
  private dxfs: string[] = [];
  private dxfIndex = new Map<string, number>();

  constructor(styles: readonly CellStyle[]) {
    const base = styles[0] ?? {};
    this.fonts.push(fontXml(base.font, 'font', true));
    this.borders.push(borderXml(undefined));
    for (const st of styles.length ? styles : [{}]) this.xfs.push(this.xf(st));
  }

  private intern(list: string[], xml: string): number {
    const i = list.indexOf(xml);
    if (i >= 0) return i;
    list.push(xml);
    return list.length - 1;
  }

  numFmtId(code: string | undefined): number {
    if (!code || code === 'General') return 0;
    const b = builtinId(code);
    if (b !== undefined) return b;
    let id = this.numFmts.get(code);
    if (id === undefined) { id = 164 + this.numFmts.size; this.numFmts.set(code, id); }
    return id;
  }

  private xf(st: CellStyle): string {
    const fontId = st.font ? this.intern(this.fonts, fontXml(st.font, 'font', true)) : 0;
    const fillId = st.fill && (st.fill.color || st.fill.pattern || st.fill.gradient) ? this.intern(this.fills, fillXml(st.fill, false)) : 0;
    const borderId = st.border ? this.intern(this.borders, borderXml(st.border)) : 0;
    const numFmtId = this.numFmtId(st.numFmt);
    const inner = alignXml(st.align) + protXml(st.protection);
    return el('xf', {
      numFmtId, fontId, fillId, borderId, xfId: 0,
      applyNumberFormat: numFmtId ? true : undefined, applyFont: fontId ? true : undefined, applyFill: fillId ? true : undefined,
      applyBorder: borderId ? true : undefined, applyAlignment: st.align ? true : undefined, applyProtection: st.protection ? true : undefined,
    }, inner);
  }

  dxf(st: CellStyle): number {
    let s = '';
    if (st.font) s += fontXml(st.font, 'font', false);
    if (st.numFmt) s += el('numFmt', { numFmtId: this.numFmtId(st.numFmt) || 164, formatCode: st.numFmt });
    if (st.fill && (st.fill.color || st.fill.pattern || st.fill.gradient)) s += fillXml(st.fill, true);
    s += alignXml(st.align);
    if (st.border) s += borderXml(st.border);
    s += protXml(st.protection);
    const xml = `<dxf>${s}</dxf>`;
    let i = this.dxfIndex.get(xml);
    if (i === undefined) { i = this.dxfs.length; this.dxfs.push(xml); this.dxfIndex.set(xml, i); }
    return i;
  }

  get counts(): { fonts: number; fills: number; borders: number; numFmts: number } {
    return { fonts: this.fonts.length, fills: this.fills.length, borders: this.borders.length, numFmts: this.numFmts.size };
  }

  xml(): string {
    const list = (tag: string, items: string[]): string => `<${tag} count="${items.length}">${items.join('')}</${tag}>`;
    const nf = [...this.numFmts].map(([code, id]) => el('numFmt', { numFmtId: id, formatCode: code }));
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
      + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      + (nf.length ? list('numFmts', nf) : '')
      + list('fonts', this.fonts) + list('fills', this.fills) + list('borders', this.borders)
      + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
      + list('cellXfs', this.xfs)
      + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
      + list('dxfs', this.dxfs)
      + '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>'
      + '</styleSheet>';
  }
}
