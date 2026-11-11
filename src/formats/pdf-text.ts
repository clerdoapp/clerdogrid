// Text layout for PDF export: font choice and fallback across the bundled fonts, bidirectional
// ordering (a compact form of the Unicode Bidirectional Algorithm, UAX #9, without explicit
// embeddings), Arabic joining (UAX #9 / Unicode ArabicShaping data) and OpenType substitution.
import { TrueTypeFont, type GlyphInfo } from './pdf-font';

export interface FontRequest { family?: string; bold?: boolean; italic?: boolean }

/**
 * Where font files come from: names like "carlito-latin-400-normal.woff". `ranges` gives each package's
 * subset coverage as CSS unicode-range text (the packages' unicode.json), so only needed chunks load.
 */
export interface FontSource { files(): string[]; load(file: string): Promise<Uint8Array>; ranges?(): Record<string, Record<string, string>> }

export interface LoadedFont { id: number; file: string; font: TrueTypeFont; bold: boolean; italic: boolean }

/** Glyph advances are in thousandths of the font size. */
/** `adv` is the glyph's advance after positioning; `x`/`y` place it from the run's left edge and baseline (y up). */
export interface ShapedGlyph { gid: number; chars: string; adv: number; x: number; y: number }
export interface ShapedRun { font: LoadedFont | null; glyphs: ShapedGlyph[]; width: number; rtl: boolean; text: string; missing?: boolean }
export interface ShapedText { runs: ShapedRun[]; width: number; rtl: boolean }

/** Bundled font packages by the names people pick in cells. */
const FAMILY: Record<string, string> = {
  calibri: 'carlito', 'calibri light': 'carlito', aptos: 'carlito', 'aptos narrow': 'carlito', 'aptos display': 'carlito', carlito: 'carlito',
  arial: 'arimo', helvetica: 'arimo', 'helvetica neue': 'arimo', verdana: 'arimo', tahoma: 'arimo', 'arial narrow': 'arimo', 'liberation sans': 'arimo', arimo: 'arimo',
  'times new roman': 'tinos', times: 'tinos', georgia: 'tinos', 'liberation serif': 'tinos', tinos: 'tinos', garamond: 'tinos', 'book antiqua': 'tinos',
  'courier new': 'cousine', courier: 'cousine', consolas: 'cousine', 'lucida console': 'cousine', 'liberation mono': 'cousine', cousine: 'cousine',
  cambria: 'caladea', caladea: 'caladea',
  'segoe ui': 'inter', segoe: 'inter', inter: 'inter', 'system-ui': 'inter',
  'hanken grotesk': 'hanken-grotesk', lato: 'lato', montserrat: 'montserrat', 'open sans': 'open-sans', roboto: 'roboto', 'roboto mono': 'roboto-mono',
  'noto sans sc': 'noto-sans-sc', simsun: 'noto-sans-sc', nsimsun: 'noto-sans-sc', simhei: 'noto-sans-sc', 'microsoft yahei': 'noto-sans-sc', dengxian: 'noto-sans-sc', 'pingfang sc': 'noto-sans-sc',
  pmingliu: 'noto-sans-sc', mingliu: 'noto-sans-sc', 'microsoft jhenghei': 'noto-sans-sc',
  'noto sans jp': 'noto-sans-jp', 'ms gothic': 'noto-sans-jp', 'ms pgothic': 'noto-sans-jp', 'ms mincho': 'noto-sans-jp', meiryo: 'noto-sans-jp', 'yu gothic': 'noto-sans-jp', 'yu mincho': 'noto-sans-jp', 'hiragino sans': 'noto-sans-jp',
  'noto sans kr': 'noto-sans-kr', 'malgun gothic': 'noto-sans-kr', batang: 'noto-sans-kr', gulim: 'noto-sans-kr', dotum: 'noto-sans-kr', 'apple sd gothic neo': 'noto-sans-kr',
  'noto naskh arabic': 'noto-naskh-arabic', 'arabic typesetting': 'noto-naskh-arabic', 'traditional arabic': 'noto-naskh-arabic', 'simplified arabic': 'noto-naskh-arabic', 'sakkal majalla': 'noto-naskh-arabic',
};
const PACKAGES = ['noto-naskh-arabic', 'noto-sans-sc', 'noto-sans-jp', 'noto-sans-kr', 'noto-emoji', 'hanken-grotesk', 'roboto-mono', 'open-sans', 'carlito', 'caladea', 'arimo', 'tinos', 'cousine', 'inter', 'lato', 'montserrat', 'roboto'];
const FALLBACK = ['noto-naskh-arabic', 'arimo', 'carlito', 'open-sans', 'roboto', 'inter', 'tinos', 'cousine'];
/** Chinese, Japanese and Korean fonts, ordered by what the text looks like. */
const cjkOrder = (text: string): string[] =>
  /[\u3040-\u30ff\u31f0-\u31ff]/.test(text) ? ['noto-sans-jp', 'noto-sans-sc', 'noto-sans-kr']
    : /[\uac00-\ud7af\u1100-\u11ff\u3130-\u318f]/.test(text) ? ['noto-sans-kr', 'noto-sans-sc', 'noto-sans-jp']
      : ['noto-sans-sc', 'noto-sans-jp', 'noto-sans-kr'];

/** CSS unicode-range text ("U+0460-052F,U+20B4,U+4??") → ranges. */
export const parseUnicodeRange = (text: string): [number, number][] => text.split(',').map((t) => t.trim().replace(/^U\+/i, '')).flatMap((t): [number, number][] => {
  if (t.includes('?')) return [[parseInt(t.replace(/\?/g, '0'), 16), parseInt(t.replace(/\?/g, 'f'), 16)]];
  const [a, b] = t.split('-');
  const lo = parseInt(a!, 16), hi = b ? parseInt(b, 16) : lo;
  return Number.isFinite(lo) && Number.isFinite(hi) ? [[lo, hi]] : [];
});

export const fontPackage = (family?: string): string => {
  const n = (family ?? '').trim().toLowerCase().replace(/^["']|["']$/g, '');
  return FAMILY[n] ?? (PACKAGES.includes(n.replace(/\s+/g, '-')) ? n.replace(/\s+/g, '-') : 'carlito');
};

/** Rough Unicode coverage of each font subset, used to decide which files to load first. */
const SUBSET: Record<string, [number, number][]> = {
  latin: [[0, 0xff], [0x131, 0x131], [0x152, 0x153], [0x2bb, 0x2bc], [0x2c6, 0x2c6], [0x2da, 0x2da], [0x2dc, 0x2dc], [0x304, 0x304], [0x308, 0x308], [0x329, 0x329], [0x2000, 0x206f], [0x20ac, 0x20ac], [0x2122, 0x2122], [0x2191, 0x2193], [0x2212, 0x2215], [0xfeff, 0xfeff], [0xfffd, 0xfffd]],
  'latin-ext': [[0x100, 0x2ff], [0x1d00, 0x1dbf], [0x1e00, 0x1eff], [0x2020, 0x2020], [0x20a0, 0x20c0], [0x2113, 0x2113], [0x2c60, 0x2c7f], [0xa720, 0xa7ff]],
  vietnamese: [[0x102, 0x103], [0x110, 0x111], [0x128, 0x129], [0x168, 0x169], [0x1a0, 0x1b0], [0x300, 0x323], [0x1ea0, 0x1ef9], [0x20ab, 0x20ab]],
  cyrillic: [[0x301, 0x301], [0x400, 0x45f], [0x490, 0x491], [0x4b0, 0x4b1], [0x2116, 0x2116]],
  'cyrillic-ext': [[0x460, 0x52f], [0x1c80, 0x1c8a], [0x20b4, 0x20b4], [0x2de0, 0x2dff], [0xa640, 0xa69f], [0xfe2e, 0xfe2f]],
  greek: [[0x370, 0x3ff]],
  'greek-ext': [[0x1f00, 0x1fff]],
  hebrew: [[0x590, 0x5ff], [0x20aa, 0x20aa], [0x25cc, 0x25cc], [0xfb1d, 0xfb4f]],
  arabic: [[0x600, 0x6ff], [0x750, 0x77f], [0x870, 0x8ff], [0x200c, 0x200e], [0x2010, 0x2011], [0x204f, 0x204f], [0x2e41, 0x2e41], [0xfb50, 0xfdff], [0xfe70, 0xfefc]],
  math: [[0x2200, 0x22ff], [0x2300, 0x23ff], [0x27c0, 0x27ef], [0x2980, 0x2aff], [0x1d400, 0x1d7ff]],
  symbols: [[0x2100, 0x21ff], [0x2190, 0x21ff], [0x2300, 0x23ff], [0x2400, 0x27ff], [0x2b00, 0x2bff], [0x1f000, 0x1faff]],
};
const subsetOf = (file: string, pkg: string): string => file.slice(pkg.length + 1).replace(/-\d{3}-(normal|italic)\.(woff2?|ttf|otf)$/, '');

const isMarkCp = (cp: number): boolean => /\p{M}/u.test(String.fromCodePoint(cp));
const isIgnorable = (cp: number): boolean => cp === 0x200c || cp === 0x200d || (cp >= 0xfe00 && cp <= 0xfe0f) || cp === 0xfeff || cp < 0x20 || (cp >= 0x7f && cp < 0xa0);

export class FontBook {
  private loaded = new Map<string, LoadedFont | null>();
  private pending = new Map<string, Promise<LoadedFont | null>>();
  private nextId = 1;
  private shapeCache = new Map<string, ShapedText>();
  private names: string[];
  private ranges = new Map<string, Map<string, [number, number][]>>();
  private fileCache = new Map<string, { files: string[]; bold: boolean; italic: boolean }>();

  /** `measureMissing` sizes text no bundled font covers (in thousandths of the font size), e.g. with the system's fonts. */
  constructor(private readonly source: FontSource, private readonly opts: { measureMissing?: (text: string, req: FontRequest) => number | null } = {}) {
    this.names = source.files().filter((f) => /\.woff$|\.ttf$|\.otf$/.test(f));
    for (const [pkg, subs] of Object.entries(source.ranges?.() ?? {})) {
      this.ranges.set(pkg, new Map(Object.entries(subs).map(([k, v]) => [k.replace(/^\[|\]$/g, ''), parseUnicodeRange(v)])));
    }
  }

  /** Whether a subset file may hold cp; `known` is false when there's no coverage data for it. */
  private covers(pkg: string, subset: string, cp: number): { yes: boolean; known: boolean } {
    const r = this.ranges.get(pkg)?.get(subset) ?? SUBSET[subset];
    return r ? { yes: r.some(([a, b]) => cp >= a && cp <= b), known: true } : { yes: true, known: false };
  }

  /** Files of a package in a weight and style, in the order to try them (latin first). */
  private filesFor(pkg: string, bold: boolean, italic: boolean): { files: string[]; bold: boolean; italic: boolean } {
    const key = `${pkg}|${bold ? 1 : 0}${italic ? 1 : 0}`;
    let hit = this.fileCache.get(key);
    if (!hit) { hit = this.findFiles(pkg, bold, italic); this.fileCache.set(key, hit); }
    return hit;
  }

  private findFiles(pkg: string, bold: boolean, italic: boolean): { files: string[]; bold: boolean; italic: boolean } {
    const of = (w: number, it: boolean) => this.names.filter((f) => f.startsWith(pkg + '-') && new RegExp(`-${w}-${it ? 'italic' : 'normal'}\\.`).test(f) && !PACKAGES.some((p) => p !== pkg && p.startsWith(pkg + '-') && f.startsWith(p + '-')));
    const tries: [number, boolean][] = [[bold ? 700 : 400, italic], [bold ? 700 : 400, false], [400, italic], [400, false]];
    for (const [w, it] of tries) {
      const files = of(w, it);
      if (files.length) {
        const order = ['latin', 'latin-ext', 'arabic', 'cyrillic', 'greek', 'hebrew', 'vietnamese', 'cyrillic-ext', 'greek-ext', 'symbols', 'math'];
        files.sort((a, b) => { const x = order.indexOf(subsetOf(a, pkg)), y = order.indexOf(subsetOf(b, pkg)); return (x < 0 ? 99 : x) - (y < 0 ? 99 : y); });
        return { files, bold: bold && w === 700, italic: italic && it };
      }
    }
    return { files: [], bold: false, italic: false };
  }

  private load(file: string, bold: boolean, italic: boolean): Promise<LoadedFont | null> {
    if (this.loaded.has(file)) return Promise.resolve(this.loaded.get(file)!);
    let p = this.pending.get(file);
    if (!p) {
      p = this.source.load(file).then((data) => {
        const lf: LoadedFont = { id: this.nextId++, file, font: new TrueTypeFont(data), bold, italic };
        this.loaded.set(file, lf);
        return lf;
      }).catch(() => { this.loaded.set(file, null); return null; });
      this.pending.set(file, p);
    }
    return p;
  }

  private chain(req: FontRequest, text = ''): string[] {
    const first = fontPackage(req.family);
    return [...new Set([first, ...FALLBACK, ...cjkOrder(text), 'noto-emoji'])];
  }

  /** Loads what is needed to draw `text` in this font. */
  async ensure(text: string, req: FontRequest = {}): Promise<void> {
    const cps = new Set<number>();
    for (const ch of text) { const cp = ch.codePointAt(0)!; if (!isIgnorable(cp) && cp !== 0x20) cps.add(cp); }
    const chain = this.chain(req, text);
    const first = this.filesFor(chain[0]!, !!req.bold, !!req.italic);
    if (first.files[0]) await this.load(first.files[0], first.bold, first.italic);
    for (const cp of cps) {
      if (this.resolve(cp, req, chain)) continue;
      let found = false;
      for (const pkg of chain) {
        const v = this.filesFor(pkg, !!req.bold, !!req.italic);
        const known = this.ranges.has(pkg);
        const likely = v.files.filter((f) => { const c = this.covers(pkg, subsetOf(f, pkg), cp); return c.yes && (c.known || !known); });
        // With coverage data only the chunks that claim cp are fetched; without it, one more file is tried.
        const rest = known ? [] : v.files.filter((x) => !likely.includes(x));
        for (const f of [...likely, ...rest]) {
          if (this.loaded.has(f) && !this.loaded.get(f)) continue;
          const lf = await this.load(f, v.bold, v.italic);
          if (lf?.font.has(cp)) { found = true; break; }
          if (!likely.length) break;
        }
        if (found) break;
      }
    }
    this.shapeCache.clear();
  }

  /** The first loaded font in the fallback chain with a glyph for cp. */
  private resolve(cp: number, req: FontRequest, chain = this.chain(req)): LoadedFont | null {
    for (const pkg of chain) {
      const v = this.filesFor(pkg, !!req.bold, !!req.italic);
      for (const f of v.files) { const lf = this.loaded.get(f); if (lf?.font.has(cp)) return lf; }
    }
    return null;
  }

  /** The main loaded font for a request (for metrics and missing glyphs). */
  primary(req: FontRequest): LoadedFont | null {
    for (const pkg of this.chain(req)) {
      const v = this.filesFor(pkg, !!req.bold, !!req.italic);
      for (const f of v.files) { const lf = this.loaded.get(f); if (lf) return lf; }
    }
    return null;
  }

  allLoaded(): LoadedFont[] { return [...this.loaded.values()].filter((x): x is LoadedFont => !!x); }

  /** Lays out one line of text. Widths are in thousandths of the font size. */
  shape(text: string, req: FontRequest = {}, rtl?: boolean): ShapedText {
    const key = `${req.family ?? ''}|${req.bold ? 1 : 0}${req.italic ? 1 : 0}|${rtl === undefined ? '' : rtl ? 'r' : 'l'}|${text}`;
    const hit = this.shapeCache.get(key);
    if (hit) return hit;
    const cps = [...text].map((c) => c.codePointAt(0)!);
    const { levels, base } = bidiLevels(cps, rtl);
    const prim = this.primary(req);
    const chain = this.chain(req, text);
    // Itemise into runs of one level and one font.
    const fonts: (LoadedFont | null)[] = [];
    cps.forEach((cp, i) => {
      const prev = i > 0 ? fonts[i - 1]! : null;
      if ((isMarkCp(cp) || isIgnorable(cp)) && prev && (prev.font.has(cp) || isIgnorable(cp))) fonts.push(prev);
      else if (cp === 0x20 || cp === 0xa0) fonts.push(prev && prev.font.has(cp) ? prev : prim);
      else fonts.push(this.resolve(cp, req, chain));
    });
    const runs: ShapedRun[] = [];
    const runLevels: number[] = [];
    let start = 0;
    for (let i = 1; i <= cps.length; i++) {
      if (i < cps.length && levels[i] === levels[start] && fonts[i] === fonts[start]) continue;
      const seg = cps.slice(start, i);
      const lvl = levels[start]!;
      runs.push(this.shapeRun(seg, fonts[start] ?? null, prim, lvl % 2 === 1, req));
      runLevels.push(lvl);
      start = i;
    }
    // Visual order (rule L2).
    const order = runs.map((_, i) => i);
    const max = Math.max(0, ...runLevels);
    const minOdd = Math.min(...runLevels.filter((l) => l % 2 === 1), 99);
    for (let lv = max; lv >= minOdd && lv >= 1; lv--) {
      for (let i = 0; i < order.length;) {
        if (runLevels[order[i]!]! >= lv) {
          let j = i;
          while (j < order.length && runLevels[order[j]!]! >= lv) j++;
          const rev = order.slice(i, j).reverse();
          order.splice(i, j - i, ...rev);
          i = j;
        } else i++;
      }
    }
    const out: ShapedText = { runs: order.map((i) => runs[i]!), width: runs.reduce((a, r) => a + r.width, 0), rtl: base === 1 };
    if (this.shapeCache.size > 50000) this.shapeCache.clear();
    this.shapeCache.set(key, out);
    return out;
  }

  private shapeRun(cps: number[], lf: LoadedFont | null, prim: LoadedFont | null, rtl: boolean, req: FontRequest): ShapedRun {
    const text = String.fromCodePoint(...cps);
    if (!lf && text.trim() && this.opts.measureMissing) {
      const w = this.opts.measureMissing(text, req);
      if (w !== null) return { font: null, glyphs: [], width: w, rtl, text, missing: true };
    }
    const f = lf ?? prim;
    if (!f) return { font: null, glyphs: [], width: cps.length * 550, rtl, text, missing: true };
    const font = f.font;
    const scale = 1000 / font.unitsPerEm;
    const arabic = cps.some((cp) => cp >= 0x600 && cp <= 0x8ff);
    const forms = arabic ? joiningForms(cps) : null;
    let buf: GlyphInfo[] = cps.map((cp, i) => {
      const m = rtl ? MIRROR[cp] : undefined;
      const gid = lf ? font.glyph(m ?? cp) || font.glyph(cp) : 0;
      return { gid, chars: String.fromCodePoint(cp), mask: COMMON | (forms ? forms[i]! : 0), cluster: i };
    }).filter((g, i) => !(isIgnorable(cps[i]!) && g.gid === 0));
    let script = 'latn';
    if (lf) {
      script = arabic ? 'arab' : cps.some((cp) => cp >= 0x590 && cp <= 0x5ff) ? 'hebr' : cps.some((cp) => cp >= 0x400 && cp <= 0x52f) ? 'cyrl' : cps.some((cp) => cp >= 0x370 && cp <= 0x3ff) ? 'grek'
        : cps.some((cp) => cp >= 0x3040 && cp <= 0x30ff) ? 'kana' : cps.some((cp) => cp >= 0xac00 && cp <= 0xd7af || cp >= 0x1100 && cp <= 0x11ff) ? 'hang' : cps.some((cp) => cp >= 0x4e00 && cp <= 0x9fff || cp >= 0x3400 && cp <= 0x4dbf) ? 'hani' : 'latn';
    }
    if (lf && font.gsub) {
      try { buf = font.gsub.apply(buf, lookupsOf(f, script, 'GSUB')); } catch { /* keep unshaped */ }
    }
    if (lf && font.gpos) {
      for (const g of buf) g.mask = 0xffff;
      try { buf = font.gpos.apply(buf, lookupsOf(f, script, 'GPOS')); } catch { /* keep unpositioned */ }
    }
    // Positions in font units, then thousandths of the size. Marks sit on the glyph they attach to.
    const adv = buf.map((g) => font.advance(g.gid) + (g.xAdv ?? 0));
    const total = adv.reduce((a, b) => a + b, 0);
    const xs: number[] = [], ys: number[] = [];
    let pen = rtl ? total : 0;
    buf.forEach((g, i) => {
      if (rtl) pen -= adv[i]!;
      xs.push(pen + (g.attach === undefined ? g.xOff ?? 0 : 0));
      ys.push(g.attach === undefined ? g.yOff ?? 0 : 0);
      if (!rtl) pen += adv[i]!;
    });
    buf.forEach((g, i) => {
      if (g.attach === undefined || g.attach < 0 || g.attach >= buf.length) return;
      const b = g.attach;
      // A mark's offset is measured from the base's origin.
      xs[i] = xs[b]! + (g.xOff ?? 0);
      ys[i] = ys[b]! + (g.yOff ?? 0);
    });
    let glyphs: ShapedGlyph[] = buf.map((g, i) => ({ gid: g.gid, chars: g.chars, adv: adv[i]! * scale, x: xs[i]! * scale, y: ys[i]! * scale }));
    if (rtl) glyphs = glyphs.reverse();
    return { font: f, glyphs, width: total * scale, rtl, text, missing: !lf };
  }

  /** Width in points. */
  measure(text: string, req: FontRequest, size: number): number { return (this.shape(text, req).width * size) / 1000; }
}

const FEATURE = { isol: 1, fina: 2, medi: 4, init: 8, ccmp: 16, locl: 32, rlig: 64, liga: 128, calt: 256, clig: 512, mset: 1024 } as const;
const POS_FEATURES = { kern: 1, mark: 1, mkmk: 1, curs: 1, dist: 1, abvm: 1, blwm: 1, palt: 0 } as const;
const COMMON = FEATURE.ccmp | FEATURE.locl | FEATURE.rlig | FEATURE.liga | FEATURE.calt | FEATURE.clig | FEATURE.mset;
const lookupCache = new WeakMap<TrueTypeFont, Map<string, [number, number][]>>();
const lookupsOf = (lf: LoadedFont, script: string, table: 'GSUB' | 'GPOS'): [number, number][] => {
  let m = lookupCache.get(lf.font);
  if (!m) { m = new Map(); lookupCache.set(lf.font, m); }
  const key = table + script;
  let l = m.get(key);
  if (!l) { l = table === 'GSUB' ? lf.font.gsub!.lookupsFor(script, FEATURE) : lf.font.gpos!.lookupsFor(script, POS_FEATURES); m.set(key, l); }
  return l;
};

const MIRROR: Record<number, number> = { 0x28: 0x29, 0x29: 0x28, 0x3c: 0x3e, 0x3e: 0x3c, 0x5b: 0x5d, 0x5d: 0x5b, 0x7b: 0x7d, 0x7d: 0x7b, 0xab: 0xbb, 0xbb: 0xab, 0x2039: 0x203a, 0x203a: 0x2039, 0x2264: 0x2265, 0x2265: 0x2264 };

// ---------- Arabic joining ----------

const RIGHT_JOINING: [number, number][] = [
  [0x622, 0x625], [0x627, 0x627], [0x629, 0x629], [0x62f, 0x632], [0x648, 0x648], [0x671, 0x673], [0x675, 0x677], [0x688, 0x699], [0x6c0, 0x6c0], [0x6c3, 0x6cb],
  [0x6cd, 0x6cd], [0x6cf, 0x6cf], [0x6d2, 0x6d3], [0x6d5, 0x6d5], [0x6ee, 0x6ef], [0x759, 0x75b], [0x76b, 0x76c], [0x771, 0x771], [0x773, 0x774], [0x778, 0x779],
  [0x8aa, 0x8ac], [0x8ae, 0x8ae], [0x8b1, 0x8b2], [0x8b9, 0x8b9],
];
const NON_JOINING: [number, number][] = [[0x600, 0x605], [0x608, 0x608], [0x60b, 0x60b], [0x60c, 0x60f], [0x61b, 0x61f], [0x621, 0x621], [0x660, 0x66d], [0x674, 0x674], [0x6d4, 0x6d4], [0x6dd, 0x6dd], [0x6e5, 0x6e6], [0x6e9, 0x6e9], [0x6f0, 0x6f9], [0x6fd, 0x6fe]];
const inRanges = (cp: number, r: [number, number][]) => r.some(([a, b]) => cp >= a && cp <= b);

type Join = 'U' | 'R' | 'D' | 'T' | 'C';
const joinType = (cp: number): Join => {
  if (cp === 0x640 || cp === 0x200d) return 'C';
  if (isMarkCp(cp) && cp >= 0x600 && cp <= 0x8ff) return 'T';
  if (cp < 0x620 || cp > 0x8ff) return isMarkCp(cp) ? 'T' : 'U';
  if (inRanges(cp, NON_JOINING)) return 'U';
  if (inRanges(cp, RIGHT_JOINING)) return 'R';
  if ((cp >= 0x620 && cp <= 0x64a) || (cp >= 0x66e && cp <= 0x6d3) || (cp >= 0x6fa && cp <= 0x6ff) || (cp >= 0x750 && cp <= 0x77f) || (cp >= 0x8a0 && cp <= 0x8c9)) return 'D';
  return 'U';
};

/** isol/fina/medi/init feature bit for each character, in logical order. */
const joiningForms = (cps: number[]): number[] => {
  const types = cps.map(joinType);
  const out = new Array<number>(cps.length).fill(0);
  const prevOf = (i: number) => { for (let j = i - 1; j >= 0; j--) if (types[j] !== 'T') return types[j]; return 'U'; };
  const nextOf = (i: number) => { for (let j = i + 1; j < types.length; j++) if (types[j] !== 'T') return types[j]; return 'U'; };
  types.forEach((t, i) => {
    if (t !== 'D' && t !== 'R') return;
    const p = prevOf(i), n = nextOf(i);
    const joinPrev = p === 'D' || p === 'C';
    const joinNext = t === 'D' && (n === 'D' || n === 'R' || n === 'C');
    out[i] = joinPrev && joinNext ? FEATURE.medi : joinPrev ? FEATURE.fina : joinNext ? FEATURE.init : FEATURE.isol;
  });
  return out;
};

// ---------- Bidirectional ordering ----------

type Bc = 'L' | 'R' | 'AL' | 'EN' | 'ES' | 'ET' | 'AN' | 'CS' | 'NSM' | 'WS' | 'ON' | 'BN';
const bidiClass = (cp: number): Bc => {
  if (cp >= 0x30 && cp <= 0x39) return 'EN';
  if (cp === 0x2b || cp === 0x2d || cp === 0x207a || cp === 0x207b || cp === 0xff0b || cp === 0xff0d || cp === 0x2212) return 'ES';
  if (cp === 0x23 || cp === 0x24 || cp === 0x25 || (cp >= 0xa2 && cp <= 0xa5) || cp === 0xb0 || cp === 0xb1 || cp === 0x2030 || cp === 0x2031 || cp === 0x2032 || cp === 0x2033 || (cp >= 0x20a0 && cp <= 0x20cf) || cp === 0x66a) return 'ET';
  if (cp === 0x2c || cp === 0x2e || cp === 0x2f || cp === 0x3a || cp === 0xa0 || cp === 0x60c || cp === 0xff0c || cp === 0xff0e || cp === 0xff0f || cp === 0xff1a) return 'CS';
  if ((cp >= 0x660 && cp <= 0x669) || cp === 0x66b || cp === 0x66c || (cp >= 0x600 && cp <= 0x605) || cp === 0x6dd) return 'AN';
  if ((cp >= 0x6f0 && cp <= 0x6f9) || cp === 0xb2 || cp === 0xb3 || cp === 0xb9 || (cp >= 0x2070 && cp <= 0x2079) || (cp >= 0xff10 && cp <= 0xff19)) return 'EN';
  if (cp === 0x200c || cp === 0x200d || cp === 0xfeff || (cp >= 0xfe00 && cp <= 0xfe0f)) return 'BN';
  if (cp === 0x200e) return 'L';
  if (cp === 0x200f) return 'R';
  if (cp === 0x61c) return 'AL';
  const ch = String.fromCodePoint(cp);
  if (/[\p{Mn}\p{Me}]/u.test(ch)) return 'NSM';
  if (/\s/.test(ch)) return 'WS';
  if ((cp >= 0x590 && cp <= 0x5ff) || (cp >= 0x7c0 && cp <= 0x85f) || (cp >= 0xfb1d && cp <= 0xfb4f) || (cp >= 0x10800 && cp <= 0x10fff) || (cp >= 0x1e800 && cp <= 0x1efff)) return 'R';
  if ((cp >= 0x600 && cp <= 0x7bf) || (cp >= 0x860 && cp <= 0x8ff) || (cp >= 0xfb50 && cp <= 0xfdff) || (cp >= 0xfe70 && cp <= 0xfeff)) return 'AL';
  if (/[\p{L}\p{Mc}]/u.test(ch) || cp === 0x200e) return 'L';
  return 'ON';
};

/** Embedding level of each character; `rtl` forces the paragraph direction, otherwise the first strong character decides. */
export const bidiLevels = (cps: number[], rtl?: boolean): { levels: Uint8Array; base: 0 | 1 } => {
  const n = cps.length;
  const t: Bc[] = cps.map(bidiClass);
  let base: 0 | 1 = 0;
  if (rtl !== undefined) base = rtl ? 1 : 0;
  else for (const c of t) { if (c === 'L') break; if (c === 'R' || c === 'AL') { base = 1; break; } }
  const sos: Bc = base ? 'R' : 'L';
  if (!t.some((c) => c === 'R' || c === 'AL' || c === 'AN') && !base) return { levels: new Uint8Array(n), base };
  // W1
  for (let i = 0; i < n; i++) if (t[i] === 'NSM' || t[i] === 'BN') t[i] = i ? t[i - 1]! : sos;
  // W2, W3
  let strong: Bc = sos;
  for (let i = 0; i < n; i++) {
    const c = t[i]!;
    if (c === 'L' || c === 'R' || c === 'AL') strong = c;
    else if (c === 'EN' && strong === 'AL') t[i] = 'AN';
  }
  for (let i = 0; i < n; i++) if (t[i] === 'AL') t[i] = 'R';
  // W4
  for (let i = 1; i < n - 1; i++) {
    if (t[i] === 'ES' && t[i - 1] === 'EN' && t[i + 1] === 'EN') t[i] = 'EN';
    else if (t[i] === 'CS' && (t[i - 1] === 'EN' || t[i - 1] === 'AN') && t[i + 1] === t[i - 1]) t[i] = t[i - 1]!;
  }
  // W5
  for (let i = 0; i < n; i++) {
    if (t[i] !== 'ET') continue;
    let j = i;
    while (j < n && t[j] === 'ET') j++;
    if ((i > 0 && t[i - 1] === 'EN') || (j < n && t[j] === 'EN')) for (let k = i; k < j; k++) t[k] = 'EN';
    i = j - 1;
  }
  // W6
  for (let i = 0; i < n; i++) if (t[i] === 'ES' || t[i] === 'ET' || t[i] === 'CS') t[i] = 'ON';
  // W7
  strong = sos;
  for (let i = 0; i < n; i++) {
    const c = t[i]!;
    if (c === 'L' || c === 'R') strong = c;
    else if (c === 'EN' && strong === 'L') t[i] = 'L';
  }
  // N1, N2
  const dirOf = (c: Bc): 'L' | 'R' | null => (c === 'L' ? 'L' : c === 'R' || c === 'EN' || c === 'AN' ? 'R' : null);
  for (let i = 0; i < n; i++) {
    if (dirOf(t[i]!)) continue;
    let j = i;
    while (j < n && !dirOf(t[j]!)) j++;
    const before = i > 0 ? dirOf(t[i - 1]!)! : (sos as 'L' | 'R');
    const after = j < n ? dirOf(t[j]!)! : (sos as 'L' | 'R');
    const d: Bc = before === after ? before : base ? 'R' : 'L';
    for (let k = i; k < j; k++) t[k] = d;
    i = j - 1;
  }
  // I1, I2
  const levels = new Uint8Array(n);
  for (let i = 0; i < n; i++) {
    const c = t[i]!;
    if (base === 0) levels[i] = c === 'R' ? 1 : c === 'AN' || c === 'EN' ? 2 : 0;
    else levels[i] = c === 'L' || c === 'EN' || c === 'AN' ? 2 : 1;
  }
  // L1: trailing whitespace takes the paragraph level.
  for (let i = n - 1; i >= 0 && /\s/.test(String.fromCodePoint(cps[i]!)); i--) levels[i] = base;
  return { levels, base };
};
