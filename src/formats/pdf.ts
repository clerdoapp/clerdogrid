// A PDF 1.7 writer (ISO 32000-1), no dependencies besides fflate for Flate compression.
// Pages are drawn with top-left origin coordinates in points; text uses embedded TrueType
// subsets (CIDFontType2, Identity-H) with ToUnicode maps so it can be selected and searched.
import { zlibSync } from 'fflate';
import type { LoadedFont, ShapedText } from './pdf-text';
import { decodePng, isJpeg, isPng, jpegInfo } from './pdf-image';

const enc = new TextEncoder();
const num = (n: number): string => {
  if (!Number.isFinite(n)) return '0';
  const r = Math.round(n * 1000) / 1000;
  return Object.is(r, -0) ? '0' : String(r);
};

/** PDF text string: literal for ASCII, UTF-16BE hex otherwise. */
export const pdfString = (s: string): string => {
  if (/^[\x20-\x7e]*$/.test(s)) return `(${s.replace(/[\\()]/g, (c) => '\\' + c)})`;
  let hex = 'FEFF';
  for (let i = 0; i < s.length; i++) hex += s.charCodeAt(i).toString(16).padStart(4, '0');
  return `<${hex.toUpperCase()}>`;
};
const pdfName = (s: string): string => '/' + s.replace(/[^A-Za-z0-9+_.-]/g, (c) => '#' + c.charCodeAt(0).toString(16).padStart(2, '0'));

export type Rgb = [number, number, number];
/** "#rrggbb" or "#rgb" → 0..1 components; unknown text gives black. */
export const rgbOf = (c: string | undefined): Rgb => {
  const m = c && /^#?([0-9a-f]{6}|[0-9a-f]{3})/i.exec(c.trim());
  if (!m) return [0, 0, 0];
  const h = m[1]!.length === 3 ? m[1]!.split('').map((x) => x + x).join('') : m[1]!;
  return [parseInt(h.slice(0, 2), 16) / 255, parseInt(h.slice(2, 4), 16) / 255, parseInt(h.slice(4, 6), 16) / 255];
};
const alphaOf = (c: string | undefined): number => {
  const m = c && /^#[0-9a-f]{6}([0-9a-f]{2})$/i.exec(c.trim());
  return m ? parseInt(m[1]!, 16) / 255 : 1;
};

class ObjectStore {
  private bodies: (Uint8Array | null)[] = [];
  alloc(): number { this.bodies.push(null); return this.bodies.length; }
  set(n: number, body: string | Uint8Array): void { this.bodies[n - 1] = typeof body === 'string' ? enc.encode(body) : body; }
  add(body: string | Uint8Array): number { const n = this.alloc(); this.set(n, body); return n; }
  stream(dict: string, data: Uint8Array, compress = true, n = this.alloc()): number {
    const d = compress ? zlibSync(data, { level: 6 }) : data;
    const head = enc.encode(`<<${dict}${compress ? '/Filter/FlateDecode' : ''}/Length ${d.length}>>\nstream\n`);
    const tail = enc.encode('\nendstream');
    const out = new Uint8Array(head.length + d.length + tail.length);
    out.set(head); out.set(d, head.length); out.set(tail, head.length + d.length);
    this.set(n, out);
    return n;
  }
  build(root: number, info: number, id: string): Uint8Array {
    const chunks: Uint8Array[] = [];
    let size = 0;
    const push = (b: Uint8Array) => { chunks.push(b); size += b.length; };
    push(enc.encode('%PDF-1.7\n'));
    push(new Uint8Array([0x25, 0xe2, 0xe3, 0xcf, 0xd3, 0x0a]));
    const offsets: number[] = [];
    this.bodies.forEach((b, i) => {
      offsets.push(size);
      push(enc.encode(`${i + 1} 0 obj\n`));
      push(b ?? enc.encode('null'));
      push(enc.encode('\nendobj\n'));
    });
    const xref = size;
    let x = `xref\n0 ${this.bodies.length + 1}\n0000000000 65535 f \n`;
    for (const o of offsets) x += `${String(o).padStart(10, '0')} 00000 n \n`;
    x += `trailer\n<</Size ${this.bodies.length + 1}/Root ${root} 0 R/Info ${info} 0 R/ID[<${id}><${id}>]>>\nstartxref\n${xref}\n%%EOF\n`;
    push(enc.encode(x));
    const out = new Uint8Array(size);
    let p = 0;
    for (const c of chunks) { out.set(c, p); p += c.length; }
    return out;
  }
}

interface FontUse { lf: LoadedFont; ref: number; glyphs: Map<number, string> }
interface ImageUse { ref: number; name: string }

export interface PdfInfo { title?: string; author?: string; subject?: string; creator?: string }

export class PdfDocument {
  private store = new ObjectStore();
  private pagesRef = this.store.alloc();
  private pageRefs: number[] = [];
  private fonts = new Map<number, FontUse>();
  private images = new Map<Uint8Array, ImageUse | null>();
  private grayImages = new Map<Uint8Array, ImageUse | null>();
  private gstates = new Map<string, string>();
  private shadings = 0;
  private outlines: { title: string; page: number }[] = [];

  addPage(width: number, height: number): PdfPage {
    const page = new PdfPage(this, width, height);
    return page;
  }

  /** Called by PdfPage.end(). */
  commitPage(page: PdfPage, content: string, resources: string, annots: string[]): void {
    const contents = this.store.stream('', enc.encode(content));
    const ref = this.store.alloc();
    const annotRefs = annots.map((a) => this.store.add(a));
    this.store.set(ref, `<</Type/Page/Parent ${this.pagesRef} 0 R/MediaBox[0 0 ${num(page.width)} ${num(page.height)}]/Resources ${resources}/Contents ${contents} 0 R${annotRefs.length ? `/Annots[${annotRefs.map((r) => `${r} 0 R`).join(' ')}]` : ''}>>`);
    this.pageRefs.push(ref);
  }

  get pageCount(): number { return this.pageRefs.length; }

  /** Adds a bookmark pointing at a page (0-based) that is already added or will be next. */
  bookmark(title: string, page: number): void { this.outlines.push({ title, page }); }

  fontUse(lf: LoadedFont): FontUse {
    let u = this.fonts.get(lf.id);
    if (!u) { u = { lf, ref: this.store.alloc(), glyphs: new Map() }; this.fonts.set(lf.id, u); }
    return u;
  }

  /** XObject for picture bytes (PNG or JPEG); null if the format can't be embedded. `gray` turns a PNG into shades of gray. */
  imageUse(data: Uint8Array, gray = false): ImageUse | null {
    const cache = gray && isPng(data) ? this.grayImages : this.images;
    if (cache.has(data)) return cache.get(data)!;
    let use: ImageUse | null = null;
    try {
      if (isJpeg(data)) {
        const j = jpegInfo(data);
        if (j) {
          const cs = j.components === 1 ? '/DeviceGray' : j.components === 4 ? '/DeviceCMYK' : '/DeviceRGB';
          const ref = this.store.stream(`/Type/XObject/Subtype/Image/Width ${j.width}/Height ${j.height}/ColorSpace${cs}/BitsPerComponent 8/Filter/DCTDecode${j.adobeInverted ? '/Decode[1 0 1 0 1 0 1 0]' : ''}`, data, false);
          use = { ref, name: `Im${ref}` };
        }
      } else if (isPng(data)) {
        const img = decodePng(data);
        if (gray && img.channels === 3) {
          const g = new Uint8Array(img.width * img.height);
          for (let i = 0; i < g.length; i++) g[i] = Math.round(0.299 * img.data[i * 3]! + 0.587 * img.data[i * 3 + 1]! + 0.114 * img.data[i * 3 + 2]!);
          img.data = g;
          img.channels = 1;
        }
        const smask = img.alpha ? this.store.stream(`/Type/XObject/Subtype/Image/Width ${img.width}/Height ${img.height}/ColorSpace/DeviceGray/BitsPerComponent 8`, img.alpha) : 0;
        const ref = this.store.stream(`/Type/XObject/Subtype/Image/Width ${img.width}/Height ${img.height}/ColorSpace/${img.channels === 1 ? 'DeviceGray' : 'DeviceRGB'}/BitsPerComponent 8${smask ? `/SMask ${smask} 0 R` : ''}`, img.data);
        use = { ref, name: `Im${ref}` };
      }
    } catch { use = null; }
    cache.set(data, use);
    return use;
  }

  gstate(fillAlpha: number, strokeAlpha: number): string {
    const k = `${num(fillAlpha)}/${num(strokeAlpha)}`;
    let name = this.gstates.get(k);
    if (!name) {
      const ref = this.store.add(`<</Type/ExtGState/ca ${num(fillAlpha)}/CA ${num(strokeAlpha)}>>`);
      name = `GS${ref}`;
      this.gstates.set(k, name);
      this.gsRefs.set(name, ref);
    }
    return name;
  }
  gsRefs = new Map<string, number>();

  /** Graphics state that paints with the Saturation blend mode (used to grey JPEG pictures). */
  desaturate(): string {
    const k = 'desat';
    let name = this.gstates.get(k);
    if (!name) { const ref = this.store.add('<</Type/ExtGState/BM/Saturation>>'); name = `GS${ref}`; this.gstates.set(k, name); this.gsRefs.set(name, ref); }
    return name;
  }

  /** Axial shading from (x0,y0) to (x1,y1) in page space with color stops. */
  shading(x0: number, y0: number, x1: number, y1: number, stops: { pos: number; color: string }[]): { name: string; ref: number } {
    const s = [...stops].sort((a, b) => a.pos - b.pos);
    if (s.length === 1) s.push({ ...s[0]!, pos: 1 });
    const fn = (a: string, b: string) => `<</FunctionType 2/Domain[0 1]/C0[${rgbOf(a).map(num).join(' ')}]/C1[${rgbOf(b).map(num).join(' ')}]/N 1>>`;
    let func: string;
    if (s.length === 2) func = fn(s[0]!.color, s[1]!.color);
    else {
      const fns = s.slice(1).map((st, i) => fn(s[i]!.color, st.color));
      const bounds = s.slice(1, -1).map((st) => num(Math.min(1, Math.max(0, st.pos))));
      func = `<</FunctionType 3/Domain[0 1]/Functions[${fns.join('')}]/Bounds[${bounds.join(' ')}]/Encode[${s.slice(1).map(() => '0 1').join(' ')}]>>`;
    }
    const lo = Math.min(1, Math.max(0, s[0]!.pos)), hi = Math.min(1, Math.max(0, s.at(-1)!.pos));
    const ax = x0 + (x1 - x0) * lo, ay = y0 + (y1 - y0) * lo, bx = x0 + (x1 - x0) * hi, by = y0 + (y1 - y0) * hi;
    const ref = this.store.add(`<</ShadingType 2/ColorSpace/DeviceRGB/Coords[${[ax, ay, bx, by].map(num).join(' ')}]/Function ${func}/Extend[true true]>>`);
    this.shadings++;
    return { name: `Sh${ref}`, ref };
  }

  private writeFonts(): void {
    for (const u of this.fonts.values()) {
      const f = u.lf.font;
      const scale = 1000 / f.unitsPerEm;
      const tag = subsetTag(u.lf.id, u.glyphs);
      const base = `${tag}+${f.postscriptName}`;
      const file = f.subset(u.glyphs.keys());
      const ff = this.store.stream(`/Length1 ${file.length}`, file);
      const flags = (f.fixedPitch ? 1 : 0) | 4 | (u.lf.italic || f.italicAngle ? 64 : 0);
      const fd = this.store.add(`<</Type/FontDescriptor/FontName${pdfName(base)}/Flags ${flags}/FontBBox[${f.bbox.map((v) => num(v * scale)).join(' ')}]/ItalicAngle ${num(f.italicAngle)}/Ascent ${num(f.ascent * scale)}/Descent ${num(f.descent * scale)}/CapHeight ${num(f.capHeight * scale)}/StemV ${f.weight >= 600 ? 140 : 80}/FontFile2 ${ff} 0 R>>`);
      const gids = [...u.glyphs.keys()].sort((a, b) => a - b);
      let w = '';
      for (let i = 0; i < gids.length;) {
        let j = i;
        while (j + 1 < gids.length && gids[j + 1] === gids[j]! + 1) j++;
        w += `${gids[i]}[${gids.slice(i, j + 1).map((g) => num(f.advance(g) * scale)).join(' ')}]`;
        i = j + 1;
      }
      const cid = this.store.add(`<</Type/Font/Subtype/CIDFontType2/BaseFont${pdfName(base)}/CIDSystemInfo<</Registry(Adobe)/Ordering(Identity)/Supplement 0>>/FontDescriptor ${fd} 0 R/DW ${num(f.advance(0) * scale)}/W[${w}]/CIDToGIDMap/Identity>>`);
      const tu = this.store.stream('', enc.encode(toUnicodeCMap(u.glyphs)));
      this.store.set(u.ref, `<</Type/Font/Subtype/Type0/BaseFont${pdfName(base)}/Encoding/Identity-H/DescendantFonts[${cid} 0 R]/ToUnicode ${tu} 0 R>>`);
    }
  }

  finish(info: PdfInfo = {}): Uint8Array {
    this.writeFonts();
    this.store.set(this.pagesRef, `<</Type/Pages/Kids[${this.pageRefs.map((r) => `${r} 0 R`).join(' ')}]/Count ${this.pageRefs.length}>>`);
    let outlines = '';
    const marks = this.outlines.filter((o) => this.pageRefs[o.page]);
    if (marks.length) {
      const root = this.store.alloc();
      const refs = marks.map(() => this.store.alloc());
      marks.forEach((m, i) => this.store.set(refs[i]!, `<</Title ${pdfString(m.title)}/Parent ${root} 0 R${i ? `/Prev ${refs[i - 1]} 0 R` : ''}${i < refs.length - 1 ? `/Next ${refs[i + 1]} 0 R` : ''}/Dest[${this.pageRefs[m.page]} 0 R/XYZ null null null]>>`));
      this.store.set(root, `<</Type/Outlines/First ${refs[0]} 0 R/Last ${refs.at(-1)} 0 R/Count ${refs.length}>>`);
      outlines = `/Outlines ${root} 0 R/PageMode/UseOutlines`;
    }
    const catalog = this.store.add(`<</Type/Catalog/Pages ${this.pagesRef} 0 R${outlines}>>`);
    const d = new Date();
    const pad = (n: number) => String(n).padStart(2, '0');
    const date = `D:${d.getUTCFullYear()}${pad(d.getUTCMonth() + 1)}${pad(d.getUTCDate())}${pad(d.getUTCHours())}${pad(d.getUTCMinutes())}${pad(d.getUTCSeconds())}Z`;
    const infoRef = this.store.add(`<<${info.title ? `/Title ${pdfString(info.title)}` : ''}${info.author ? `/Author ${pdfString(info.author)}` : ''}${info.subject ? `/Subject ${pdfString(info.subject)}` : ''}/Creator ${pdfString(info.creator ?? 'ClerdoGrid')}/Producer (ClerdoGrid)/CreationDate (${date})/ModDate (${date})>>`);
    let id = '';
    const seed = `${date}${this.pageRefs.length}${Math.random()}`;
    let h = 0x811c9dc5;
    for (let i = 0; i < 16; i++) { for (const ch of seed + i) h = Math.imul(h ^ ch.charCodeAt(0), 0x01000193) >>> 0; id += (h & 255).toString(16).padStart(2, '0'); }
    return this.store.build(catalog, infoRef, id.toUpperCase());
  }
}

const subsetTag = (id: number, glyphs: Map<number, string>): string => {
  let h = id * 2654435761;
  for (const g of glyphs.keys()) h = (Math.imul(h ^ g, 16777619) >>> 0);
  let t = '';
  for (let i = 0; i < 6; i++) { t += String.fromCharCode(65 + (h % 26)); h = Math.floor(h / 26) + (i + 1) * 7919; }
  return t;
};

const utf16hex = (s: string): string => {
  let h = '';
  for (let i = 0; i < s.length; i++) h += s.charCodeAt(i).toString(16).padStart(4, '0');
  return h.toUpperCase();
};

const toUnicodeCMap = (glyphs: Map<number, string>): string => {
  const entries = [...glyphs].filter(([, s]) => s.length > 0).sort((a, b) => a[0] - b[0]);
  let body = '';
  for (let i = 0; i < entries.length; i += 100) {
    const chunk = entries.slice(i, i + 100);
    body += `${chunk.length} beginbfchar\n${chunk.map(([g, s]) => `<${g.toString(16).padStart(4, '0').toUpperCase()}> <${utf16hex(s.slice(0, 32))}>`).join('\n')}\nendbfchar\n`;
  }
  return '/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo <</Registry (Adobe) /Ordering (UCS) /Supplement 0>> def\n/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n'
    + '1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n' + body + 'endcmap\nCMapName currentdict /CMapResource defineresource pop\nend\nend\n';
};

export interface StrokeStyle { color: string; width: number; dash?: number[]; cap?: 'butt' | 'round' | 'square'; join?: 'miter' | 'round' | 'bevel' }
export type PathCmd = ['M', number, number] | ['L', number, number] | ['C', number, number, number, number, number, number] | ['Z'];

/** One page. Coordinates are points from the top-left corner. */
export class PdfPage {
  private ops: string[] = [];
  private fontsUsed = new Map<string, number>();
  private xobjects = new Map<string, number>();
  private gs = new Set<string>();
  private sh = new Map<string, number>();
  private annots: string[] = [];
  private ended = false;
  constructor(private readonly doc: PdfDocument, readonly width: number, readonly height: number) {
    this.ops.push(`1 0 0 -1 0 ${num(height)} cm`);
  }

  save(): void { this.ops.push('q'); }
  restore(): void { this.ops.push('Q'); }
  transform(a: number, b: number, c: number, d: number, e: number, f: number): void { this.ops.push(`${[a, b, c, d, e, f].map(num).join(' ')} cm`); }

  private fillColor(c: string): void {
    const [r, g, b] = rgbOf(c);
    this.ops.push(`${num(r)} ${num(g)} ${num(b)} rg`);
    const a = alphaOf(c);
    if (a < 1) this.alpha(a, 1);
  }
  private strokeColor(c: string): void {
    const [r, g, b] = rgbOf(c);
    this.ops.push(`${num(r)} ${num(g)} ${num(b)} RG`);
  }
  alpha(fill: number, stroke: number): void {
    const name = this.doc.gstate(fill, stroke);
    this.gs.add(name);
    this.ops.push(`/${name} gs`);
  }

  private pathOps(cmds: PathCmd[]): void {
    for (const c of cmds) {
      if (c[0] === 'M') this.ops.push(`${num(c[1])} ${num(c[2])} m`);
      else if (c[0] === 'L') this.ops.push(`${num(c[1])} ${num(c[2])} l`);
      else if (c[0] === 'C') this.ops.push(`${c.slice(1).map((v) => num(v as number)).join(' ')} c`);
      else this.ops.push('h');
    }
  }

  rect(x: number, y: number, w: number, h: number, fill: string): void {
    this.ops.push('q');
    this.fillColor(fill);
    this.ops.push(`${num(x)} ${num(y)} ${num(w)} ${num(h)} re f`, 'Q');
  }

  clipRect(x: number, y: number, w: number, h: number): void { this.ops.push(`${num(x)} ${num(y)} ${num(w)} ${num(h)} re W n`); }
  clipPath(cmds: PathCmd[], evenOdd = false): void { this.pathOps(cmds); this.ops.push(evenOdd ? 'W* n' : 'W n'); }

  path(cmds: PathCmd[], fill?: string, stroke?: StrokeStyle, evenOdd = false, fillAlpha = 1, strokeAlpha = 1): void {
    if (!cmds.length || (!fill && !stroke)) return;
    this.ops.push('q');
    if (fillAlpha < 1 || strokeAlpha < 1) this.alpha(fillAlpha, strokeAlpha);
    if (fill) this.fillColor(fill);
    if (stroke) {
      this.strokeColor(stroke.color);
      this.ops.push(`${num(Math.max(0, stroke.width))} w`);
      if (stroke.dash?.length) this.ops.push(`[${stroke.dash.map(num).join(' ')}] 0 d`);
      if (stroke.cap) this.ops.push(`${stroke.cap === 'butt' ? 0 : stroke.cap === 'round' ? 1 : 2} J`);
      if (stroke.join) this.ops.push(`${stroke.join === 'miter' ? 0 : stroke.join === 'round' ? 1 : 2} j`);
    }
    this.pathOps(cmds);
    this.ops.push(fill && stroke ? (evenOdd ? 'B*' : 'B') : fill ? (evenOdd ? 'f*' : 'f') : 'S', 'Q');
  }

  line(x1: number, y1: number, x2: number, y2: number, s: StrokeStyle): void { this.path([['M', x1, y1], ['L', x2, y2]], undefined, s); }

  gradientRect(x: number, y: number, w: number, h: number, angle: number, stops: { pos: number; color: string }[]): void {
    const rad = (angle * Math.PI) / 180;
    const sh = this.doc.shading(x, y, x + Math.cos(rad) * w, y + Math.sin(rad) * h, stops);
    this.sh.set(sh.name, sh.ref);
    this.ops.push('q', `${num(x)} ${num(y)} ${num(w)} ${num(h)} re W n`, `/${sh.name} sh`, 'Q');
  }

  /** Draws picture bytes into a box. Returns false if the format isn't supported. */
  image(data: Uint8Array, x: number, y: number, w: number, h: number, actualText?: string, gray = false): boolean {
    const use = this.doc.imageUse(data, gray);
    if (!use) return false;
    this.xobjects.set(use.name, use.ref);
    // Text drawn as a picture keeps its words for copy and search.
    if (actualText) this.ops.push(`/Span<</ActualText ${pdfString(actualText)}>>BDC`);
    this.ops.push('q', `${num(w)} 0 0 ${num(-h)} ${num(x)} ${num(y + h)} cm`, `/${use.name} Do`, 'Q');
    if (gray && !isPng(data)) {
      // A gray box painted in Saturation mode takes the colour out of what's under it.
      const gs = this.doc.desaturate();
      this.gs.add(gs);
      this.ops.push('q', `/${gs} gs`, '0.5 g', `${num(x)} ${num(y)} ${num(w)} ${num(h)} re f`, 'Q');
    }
    if (actualText) this.ops.push('EMC');
    return true;
  }

  /**
   * Draws laid-out text with its left edge at x and baseline at y.
   * `skew` slants for a missing italic; `bold` strokes the outline for a missing bold weight.
   */
  text(t: ShapedText, x: number, y: number, size: number, color: string, opts: { skew?: boolean; bold?: boolean } = {}): void {
    const [r, g, b] = rgbOf(color);
    const a = alphaOf(color);
    this.ops.push('q');
    if (a < 1) this.alpha(a, a);
    this.ops.push(`${num(r)} ${num(g)} ${num(b)} rg`);
    if (opts.bold) this.ops.push(`${num(r)} ${num(g)} ${num(b)} RG ${num(size * 0.035)} w 2 Tr`);
    this.ops.push('BT');
    const skew = opts.skew ? Math.tan((12 * Math.PI) / 180) : 0;
    let cx = x;
    for (const run of t.runs) {
      if (!run.font || !run.glyphs.length) { cx += (run.width * size) / 1000; continue; }
      const u = this.doc.fontUse(run.font);
      const name = `F${run.font.id}`;
      this.fontsUsed.set(name, u.ref);
      const f = run.font.font;
      const em = 1000 / f.unitsPerEm;
      this.ops.push(`/${name} ${num(size)} Tf`);
      // Glyphs are placed where layout put them: kerning goes into TJ adjustments, marks get their own position.
      let parts: string[] = [];
      let lineY: number | null = null;
      let expect = 0;
      const flush = () => { if (parts.length) this.ops.push(`[${parts.join('')}] TJ`); parts = []; };
      for (const gl of run.glyphs) {
        if (!u.glyphs.has(gl.gid) || (!u.glyphs.get(gl.gid) && gl.chars)) u.glyphs.set(gl.gid, gl.chars);
        if (lineY === null || Math.abs(gl.y - lineY) > 0.01) {
          flush();
          lineY = gl.y;
          this.ops.push(`1 0 ${num(skew)} -1 ${num(cx + (gl.x * size) / 1000)} ${num(y - (gl.y * size) / 1000)} Tm`);
          expect = gl.x;
        }
        const gap = gl.x - expect;
        if (Math.abs(gap) > 0.01) parts.push(num(-gap));
        parts.push(`<${gl.gid.toString(16).padStart(4, '0')}>`);
        expect = gl.x + f.advance(gl.gid) * em;
      }
      flush();
      cx += (run.width * size) / 1000;
    }
    this.ops.push('ET', 'Q');
  }

  /** A clickable link to an https or mailto address. */
  link(x: number, y: number, w: number, h: number, uri: string): void {
    if (!/^(https:|mailto:)/i.test(uri)) return;
    const y1 = this.height - y - h, y2 = this.height - y;
    this.annots.push(`<</Type/Annot/Subtype/Link/Rect[${[x, y1, x + w, y2].map(num).join(' ')}]/Border[0 0 0]/A<</S/URI/URI ${pdfString(uri)}>>>>`);
  }

  raw(op: string): void { this.ops.push(op); }

  end(): void {
    if (this.ended) return;
    this.ended = true;
    const res: string[] = [];
    if (this.fontsUsed.size) res.push(`/Font<<${[...this.fontsUsed].map(([n, r]) => `/${n} ${r} 0 R`).join('')}>>`);
    if (this.xobjects.size) res.push(`/XObject<<${[...this.xobjects].map(([n, r]) => `/${n} ${r} 0 R`).join('')}>>`);
    if (this.gs.size) res.push(`/ExtGState<<${[...this.gs].map((n) => `/${n} ${this.doc.gsRefs.get(n)} 0 R`).join('')}>>`);
    if (this.sh.size) res.push(`/Shading<<${[...this.sh].map(([n, r]) => `/${n} ${r} 0 R`).join('')}>>`);
    this.doc.commitPage(this, this.ops.join('\n'), `<<${res.join('')}/ProcSet[/PDF/Text/ImageB/ImageC/ImageI]>>`, this.annots);
  }
}
