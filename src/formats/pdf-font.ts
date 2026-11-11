// TrueType font reading for PDF export, written from the OpenType specification:
// WOFF 1.0 unwrapping, cmap, metrics, GDEF/GSUB glyph substitution and glyf subsetting.
import { unzlibSync } from 'fflate';

class Reader {
  readonly v: DataView;
  constructor(readonly b: Uint8Array) { this.v = new DataView(b.buffer, b.byteOffset, b.byteLength); }
  u8(o: number): number { return o < this.b.length ? this.b[o]! : 0; }
  u16(o: number): number { return o + 2 <= this.b.length ? this.v.getUint16(o) : 0; }
  i16(o: number): number { return o + 2 <= this.b.length ? this.v.getInt16(o) : 0; }
  u32(o: number): number { return o + 4 <= this.b.length ? this.v.getUint32(o) : 0; }
  tag(o: number): string { return String.fromCharCode(this.u8(o), this.u8(o + 1), this.u8(o + 2), this.u8(o + 3)); }
}

const MAX_TABLE = 64 * 1024 * 1024;

/** Table data of a TrueType/OpenType font, or of a WOFF 1.0 file. */
export const fontTables = (data: Uint8Array): Map<string, Uint8Array> => {
  const r = new Reader(data);
  const sig = r.tag(0);
  const out = new Map<string, Uint8Array>();
  if (sig === 'wOFF') {
    const n = r.u16(12);
    for (let i = 0; i < n; i++) {
      const e = 44 + i * 20;
      const tag = r.tag(e), off = r.u32(e + 4), comp = r.u32(e + 8), orig = r.u32(e + 12);
      if (off + comp > data.length || orig > MAX_TABLE) throw new Error('Damaged font file.');
      const raw = data.subarray(off, off + comp);
      out.set(tag, comp < orig ? unzlibSync(raw) : raw.slice());
    }
    return out;
  }
  if (sig === 'wOF2') throw new Error('WOFF2 fonts are not supported here.');
  const v = r.u32(0);
  if (v !== 0x00010000 && sig !== 'true' && sig !== 'OTTO') throw new Error('Not a font file.');
  const n = r.u16(4);
  for (let i = 0; i < n; i++) {
    const e = 12 + i * 16;
    const tag = r.tag(e), off = r.u32(e + 8), len = r.u32(e + 12);
    if (off + len > data.length) throw new Error('Damaged font file.');
    out.set(tag, data.subarray(off, off + len));
  }
  return out;
};

// ---------- Coverage and class definitions ----------

const coverage = (r: Reader, o: number): Map<number, number> => {
  const m = new Map<number, number>();
  const fmt = r.u16(o);
  if (fmt === 1) {
    const n = r.u16(o + 2);
    for (let i = 0; i < n; i++) m.set(r.u16(o + 4 + i * 2), i);
  } else if (fmt === 2) {
    const n = r.u16(o + 2);
    for (let i = 0; i < n; i++) {
      const p = o + 4 + i * 6;
      const s = r.u16(p), e = r.u16(p + 2), idx = r.u16(p + 4);
      for (let g = s; g <= e && g - s < 65536; g++) m.set(g, idx + g - s);
    }
  }
  return m;
};

const classDef = (r: Reader, o: number): Map<number, number> => {
  const m = new Map<number, number>();
  if (!o) return m;
  const fmt = r.u16(o);
  if (fmt === 1) {
    const start = r.u16(o + 2), n = r.u16(o + 4);
    for (let i = 0; i < n; i++) m.set(start + i, r.u16(o + 6 + i * 2));
  } else if (fmt === 2) {
    const n = r.u16(o + 2);
    for (let i = 0; i < n; i++) {
      const p = o + 4 + i * 6;
      const s = r.u16(p), e = r.u16(p + 2), c = r.u16(p + 4);
      for (let g = s; g <= e && g - s < 65536; g++) m.set(g, c);
    }
  }
  return m;
};

// ---------- GSUB and GPOS (shared lookup machinery) ----------

/**
 * One glyph in a shaping buffer. `chars` is the text it stands for (for copy and search).
 * Positioning (font units) is filled in by GPOS: extra advance, placement offsets, and the
 * index of the glyph a mark is attached to.
 */
export interface GlyphInfo { gid: number; chars: string; mask: number; cluster: number; xAdv?: number; xOff?: number; yOff?: number; attach?: number; lig?: number }

interface Lookup { type: number; flag: number; subs: number[]; markSet?: number }

abstract class OtLayout {
  protected r: Reader;
  protected lookups: Lookup[] = [];
  private scriptList: number;
  private featureList: number;
  private covCache = new Map<number, Map<number, number>>();
  private classCache = new Map<number, Map<number, number>>();
  protected abstract readonly extensionType: number;
  protected abstract readonly contextType: number;

  constructor(data: Uint8Array, protected readonly gdef: Gdef | null) {
    this.r = new Reader(data);
    this.scriptList = this.r.u16(4);
    this.featureList = this.r.u16(6);
    const ll = this.r.u16(8);
    const n = this.r.u16(ll);
    for (let i = 0; i < n; i++) {
      const lo = ll + this.r.u16(ll + 2 + i * 2);
      const type = this.r.u16(lo), flag = this.r.u16(lo + 2), cnt = this.r.u16(lo + 4);
      const subs: number[] = [];
      for (let k = 0; k < cnt; k++) subs.push(lo + this.r.u16(lo + 6 + k * 2));
      this.lookups.push({ type, flag, subs, markSet: flag & 0x10 ? this.r.u16(lo + 6 + cnt * 2) : undefined });
    }
  }

  protected cov(o: number) { let c = this.covCache.get(o); if (!c) { c = coverage(this.r, o); this.covCache.set(o, c); } return c; }
  protected cls(o: number) { let c = this.classCache.get(o); if (!c) { c = classDef(this.r, o); this.classCache.set(o, c); } return c; }

  /** Lookup indexes for the given features (tag → bit) in a script, with the feature bits that switch each on. */
  lookupsFor(script: string, features: Record<string, number>): [number, number][] {
    const r = this.r;
    const sl = this.scriptList;
    const n = r.u16(sl);
    let so = 0;
    for (const want of [script, 'DFLT', 'dflt', 'latn']) {
      for (let i = 0; i < n; i++) if (r.tag(sl + 2 + i * 6) === want) { so = sl + r.u16(sl + 2 + i * 6 + 4); break; }
      if (so) break;
    }
    if (!so) return [];
    const ls = r.u16(so) ? so + r.u16(so) : 0;
    if (!ls) return [];
    const fl = this.featureList;
    const idx: number[] = [];
    const req = r.u16(ls + 2);
    if (req !== 0xffff) idx.push(req);
    const fc = r.u16(ls + 4);
    for (let i = 0; i < fc; i++) idx.push(r.u16(ls + 6 + i * 2));
    const masks = new Map<number, number>();
    for (const fi of idx) {
      if (fi >= r.u16(fl)) continue;
      const rec = fl + 2 + fi * 6;
      const bit = features[r.tag(rec)];
      if (!bit) continue;
      const fo = fl + r.u16(rec + 4);
      const lc = r.u16(fo + 2);
      for (let k = 0; k < lc; k++) { const li = r.u16(fo + 4 + k * 2); masks.set(li, (masks.get(li) ?? 0) | bit); }
    }
    return [...masks].filter(([li]) => li < this.lookups.length).sort((a, b) => a[0] - b[0]);
  }

  protected glyphClass(g: number): number { return this.gdef?.glyphClass(g) ?? 0; }

  protected skip(g: GlyphInfo, flag: number): boolean {
    if (!this.gdef || !(flag & 0xff0e)) return false;
    const c = this.gdef.glyphClass(g.gid);
    if (flag & 2 && c === 1) return true;
    if (flag & 4 && c === 2) return true;
    if (flag & 8 && c === 3) return true;
    if (c === 3 && flag & 0xff00 && this.gdef.markClass(g.gid) !== flag >> 8) return true;
    return false;
  }

  apply(buf: GlyphInfo[], lookups: [number, number][]): GlyphInfo[] {
    for (const [li, mask] of lookups) {
      const lk = this.lookups[li]!;
      for (let i = 0; i < buf.length; i++) {
        if (!(buf[i]!.mask & mask) || this.skip(buf[i]!, lk.flag)) continue;
        const before = buf.length;
        if (this.applyAt(buf, i, li, 0)) i += Math.max(0, buf.length - before);
      }
    }
    return buf;
  }

  /** Applies lookup `li` at position i. Returns true if something changed. */
  protected applyAt(buf: GlyphInfo[], i: number, li: number, depth: number): boolean {
    const lk = this.lookups[li];
    if (!lk || depth > 8) return false;
    for (const st of lk.subs) {
      let type = lk.type, o = st;
      if (type === this.extensionType) { type = this.r.u16(o + 2); o = o + this.r.u32(o + 4); }
      if (type === this.contextType ? this.context(buf, i, o, this.r.u16(o), lk.flag, depth)
        : type === this.contextType + 1 ? this.chain(buf, i, o, this.r.u16(o), lk.flag, depth)
        : this.applySub(buf, i, type, o, lk.flag)) return true;
    }
    return false;
  }

  protected abstract applySub(buf: GlyphInfo[], i: number, type: number, o: number, flag: number): boolean;

  protected next(buf: GlyphInfo[], i: number, flag: number, dir = 1): number {
    let j = i + dir;
    while (j >= 0 && j < buf.length && this.skip(buf[j]!, flag)) j += dir;
    return j;
  }

  /** Positions of `count` input glyphs starting at i (skipping ignored glyphs), checked by `test`. */
  private matchInput(buf: GlyphInfo[], i: number, count: number, flag: number, test: (k: number, gid: number) => boolean): number[] | null {
    const pos = [i];
    let j = i;
    for (let k = 1; k < count; k++) {
      j = this.next(buf, j, flag);
      if (j >= buf.length || !test(k, buf[j]!.gid)) return null;
      pos.push(j);
    }
    return pos;
  }

  private matchSide(buf: GlyphInfo[], i: number, count: number, flag: number, dir: number, test: (k: number, gid: number) => boolean): boolean {
    let j = i;
    for (let k = 0; k < count; k++) {
      j = this.next(buf, j, flag, dir);
      if (j < 0 || j >= buf.length || !test(k, buf[j]!.gid)) return false;
    }
    return true;
  }

  private runRecords(buf: GlyphInfo[], pos: number[], recs: number, count: number, depth: number): boolean {
    const r = this.r;
    const p = [...pos];
    for (let k = 0; k < count; k++) {
      const si = r.u16(recs + k * 4), li = r.u16(recs + k * 4 + 2);
      if (si >= p.length) continue;
      const before = buf.length;
      this.applyAt(buf, p[si]!, li, depth + 1);
      const delta = buf.length - before;
      if (delta) for (let q = si + 1; q < p.length; q++) p[q]! += delta;
    }
    return true;
  }

  private context(buf: GlyphInfo[], i: number, o: number, fmt: number, flag: number, depth: number): boolean {
    const r = this.r;
    const gid = buf[i]!.gid;
    if (fmt === 1 || fmt === 2) {
      const ci = this.cov(o + r.u16(o + 2)).get(gid);
      if (ci === undefined) return false;
      const cd = fmt === 2 ? this.cls(o + r.u16(o + 4)) : null;
      const setsAt = fmt === 2 ? o + 6 : o + 4;
      const idx = cd ? cd.get(gid) ?? 0 : ci;
      if (idx >= r.u16(setsAt)) return false;
      const off = r.u16(setsAt + 2 + idx * 2);
      if (!off) return false;
      const rs = o + off;
      const n = r.u16(rs);
      for (let k = 0; k < n; k++) {
        const ro = rs + r.u16(rs + 2 + k * 2);
        const gc = r.u16(ro), sc = r.u16(ro + 2);
        const pos = this.matchInput(buf, i, gc, flag, (q, g) => (cd ? (cd.get(g) ?? 0) : g) === r.u16(ro + 4 + (q - 1) * 2));
        if (pos) return this.runRecords(buf, pos, ro + 4 + (gc - 1) * 2, sc, depth);
      }
      return false;
    }
    if (fmt === 3) {
      const gc = r.u16(o + 2), sc = r.u16(o + 4);
      if (!this.cov(o + r.u16(o + 6)).has(gid)) return false;
      const pos = this.matchInput(buf, i, gc, flag, (q, g) => this.cov(o + r.u16(o + 6 + q * 2)).has(g));
      return pos ? this.runRecords(buf, pos, o + 6 + gc * 2, sc, depth) : false;
    }
    return false;
  }

  private chain(buf: GlyphInfo[], i: number, o: number, fmt: number, flag: number, depth: number): boolean {
    const r = this.r;
    const gid = buf[i]!.gid;
    if (fmt === 1 || fmt === 2) {
      const ci = this.cov(o + r.u16(o + 2)).get(gid);
      if (ci === undefined) return false;
      const bc = fmt === 2 ? this.cls(o + r.u16(o + 4)) : null, ic = fmt === 2 ? this.cls(o + r.u16(o + 6)) : null, lc = fmt === 2 ? this.cls(o + r.u16(o + 8)) : null;
      const setsAt = fmt === 2 ? o + 10 : o + 4;
      const idx = ic ? ic.get(gid) ?? 0 : ci;
      if (idx >= r.u16(setsAt)) return false;
      const off = r.u16(setsAt + 2 + idx * 2);
      if (!off) return false;
      const rs = o + off;
      const n = r.u16(rs);
      const val = (m: Map<number, number> | null, g: number) => (m ? m.get(g) ?? 0 : g);
      for (let k = 0; k < n; k++) {
        let p = rs + r.u16(rs + 2 + k * 2);
        const bn = r.u16(p); const bAt = p + 2; p = bAt + bn * 2;
        const inN = r.u16(p); const inAt = p + 2; p = inAt + (inN - 1) * 2;
        const ln = r.u16(p); const lAt = p + 2; p = lAt + ln * 2;
        const sc = r.u16(p); const recs = p + 2;
        const pos = this.matchInput(buf, i, inN, flag, (q, g) => val(ic, g) === r.u16(inAt + (q - 1) * 2));
        if (!pos) continue;
        if (!this.matchSide(buf, i, bn, flag, -1, (q, g) => val(bc, g) === r.u16(bAt + q * 2))) continue;
        if (!this.matchSide(buf, pos[pos.length - 1]!, ln, flag, 1, (q, g) => val(lc, g) === r.u16(lAt + q * 2))) continue;
        return this.runRecords(buf, pos, recs, sc, depth);
      }
      return false;
    }
    if (fmt === 3) {
      let p = o + 2;
      const bn = r.u16(p); const bAt = p + 2; p = bAt + bn * 2;
      const inN = r.u16(p); const inAt = p + 2; p = inAt + inN * 2;
      const ln = r.u16(p); const lAt = p + 2; p = lAt + ln * 2;
      const sc = r.u16(p); const recs = p + 2;
      if (!inN || !this.cov(o + r.u16(inAt)).has(gid)) return false;
      const pos = this.matchInput(buf, i, inN, flag, (q, g) => this.cov(o + r.u16(inAt + q * 2)).has(g));
      if (!pos) return false;
      if (!this.matchSide(buf, i, bn, flag, -1, (q, g) => this.cov(o + r.u16(bAt + q * 2)).has(g))) return false;
      if (!this.matchSide(buf, pos[pos.length - 1]!, ln, flag, 1, (q, g) => this.cov(o + r.u16(lAt + q * 2)).has(g))) return false;
      return this.runRecords(buf, pos, recs, sc, depth);
    }
    return false;
  }
}

export class Gsub extends OtLayout {
  protected readonly extensionType = 7;
  protected readonly contextType = 5;

  protected applySub(buf: GlyphInfo[], i: number, type: number, o: number, flag: number): boolean {
    const r = this.r;
    const g = buf[i]!;
    const fmt = r.u16(o);
    switch (type) {
      case 1: {
        const ci = this.cov(o + r.u16(o + 2)).get(g.gid);
        if (ci === undefined) return false;
        g.gid = fmt === 1 ? (g.gid + r.i16(o + 4)) & 0xffff : r.u16(o + 6 + ci * 2);
        return true;
      }
      case 2: case 3: {
        const ci = this.cov(o + r.u16(o + 2)).get(g.gid);
        if (ci === undefined || ci >= r.u16(o + 4)) return false;
        const so = o + r.u16(o + 6 + ci * 2);
        const n = r.u16(so);
        if (type === 3) { if (n) g.gid = r.u16(so + 2); return n > 0; }
        if (!n) return false;
        const seq: GlyphInfo[] = [];
        for (let k = 0; k < n; k++) seq.push({ gid: r.u16(so + 2 + k * 2), chars: k === 0 ? g.chars : '', mask: g.mask, cluster: g.cluster });
        buf.splice(i, 1, ...seq);
        return true;
      }
      case 4: {
        const ci = this.cov(o + r.u16(o + 2)).get(g.gid);
        if (ci === undefined || ci >= r.u16(o + 4)) return false;
        const lso = o + r.u16(o + 6 + ci * 2);
        const n = r.u16(lso);
        for (let k = 0; k < n; k++) {
          const lo = lso + r.u16(lso + 2 + k * 2);
          const lig = r.u16(lo), cc = r.u16(lo + 2);
          const pos: number[] = [];
          let j = i, ok = true;
          for (let c = 1; c < cc; c++) {
            j = this.next(buf, j, flag);
            if (j >= buf.length || buf[j]!.gid !== r.u16(lo + 4 + (c - 1) * 2)) { ok = false; break; }
            pos.push(j);
          }
          if (!ok) continue;
          g.gid = lig;
          g.lig = cc;
          for (const p of pos) g.chars += buf[p]!.chars;
          for (let p = pos.length - 1; p >= 0; p--) buf.splice(pos[p]!, 1);
          return true;
        }
        return false;
      }
      default: return false;
    }
  }
}

export class Gpos extends OtLayout {
  protected readonly extensionType = 9;
  protected readonly contextType = 7;

  /** Reads a ValueRecord; returns its size in bytes. */
  private value(o: number, fmt: number, g: GlyphInfo | null): number {
    const r = this.r;
    let p = o;
    const take = () => { const v = r.i16(p); p += 2; return v; };
    if (fmt & 1) { const v = take(); if (g) g.xOff = (g.xOff ?? 0) + v; }
    if (fmt & 2) { const v = take(); if (g) g.yOff = (g.yOff ?? 0) + v; }
    if (fmt & 4) { const v = take(); if (g) g.xAdv = (g.xAdv ?? 0) + v; }
    for (let b = 8; b <= 0x80; b <<= 1) if (fmt & b) p += 2;
    return p - o;
  }
  private static size(fmt: number): number { let n = 0; for (let b = 1; b <= 0x80; b <<= 1) if (fmt & b) n += 2; return n; }

  private anchor(o: number): [number, number] { return [this.r.i16(o + 2), this.r.i16(o + 4)]; }

  /** Attaches the mark at i to the glyph at `base` using the two anchors. */
  private attach(buf: GlyphInfo[], i: number, base: number, ba: [number, number], ma: [number, number]): void {
    const m = buf[i]!;
    m.attach = base;
    m.xOff = ba[0] - ma[0];
    m.yOff = ba[1] - ma[1];
  }

  protected applySub(buf: GlyphInfo[], i: number, type: number, o: number, flag: number): boolean {
    const r = this.r;
    const g = buf[i]!;
    const fmt = r.u16(o);
    switch (type) {
      case 1: {
        const ci = this.cov(o + r.u16(o + 2)).get(g.gid);
        if (ci === undefined) return false;
        const vf = r.u16(o + 4);
        this.value(fmt === 1 ? o + 6 : o + 8 + ci * Gpos.size(vf), vf, g);
        return true;
      }
      case 2: {
        const ci = this.cov(o + r.u16(o + 2)).get(g.gid);
        if (ci === undefined) return false;
        const j = this.next(buf, i, flag);
        if (j >= buf.length) return false;
        const h = buf[j]!;
        const vf1 = r.u16(o + 4), vf2 = r.u16(o + 6);
        const s1 = Gpos.size(vf1), s2 = Gpos.size(vf2);
        if (fmt === 1) {
          if (ci >= r.u16(o + 8)) return false;
          const ps = o + r.u16(o + 10 + ci * 2);
          const n = r.u16(ps);
          // Pair records are sorted by second glyph.
          let lo = 0, hi = n - 1;
          while (lo <= hi) {
            const mid = (lo + hi) >> 1;
            const rec = ps + 2 + mid * (2 + s1 + s2);
            const second = r.u16(rec);
            if (second === h.gid) { this.value(rec + 2, vf1, g); this.value(rec + 2 + s1, vf2, h); return true; }
            if (second < h.gid) lo = mid + 1; else hi = mid - 1;
          }
          return false;
        }
        if (fmt === 2) {
          const c1 = this.cls(o + r.u16(o + 8)).get(g.gid) ?? 0, c2 = this.cls(o + r.u16(o + 10)).get(h.gid) ?? 0;
          const n1 = r.u16(o + 12), n2 = r.u16(o + 14);
          if (c1 >= n1 || c2 >= n2) return false;
          const rec = o + 16 + (c1 * n2 + c2) * (s1 + s2);
          this.value(rec, vf1, g); this.value(rec + s1, vf2, h);
          return true;
        }
        return false;
      }
      case 3: {
        const ci = this.cov(o + r.u16(o + 2)).get(g.gid);
        if (ci === undefined) return false;
        const j = this.next(buf, i, flag);
        if (j >= buf.length) return false;
        const cj = this.cov(o + r.u16(o + 2)).get(buf[j]!.gid);
        if (cj === undefined) return false;
        const exitOff = r.u16(o + 6 + ci * 4 + 2), entryOff = r.u16(o + 6 + cj * 4);
        if (!exitOff || !entryOff) return false;
        const ex = this.anchor(o + exitOff), en = this.anchor(o + entryOff);
        buf[j]!.yOff = (buf[j]!.yOff ?? 0) + ex[1] - en[1];
        return true;
      }
      case 4: case 5: case 6: {
        if (this.glyphClass(g.gid) !== 3 && this.gdef) return false;
        const mci = this.cov(o + r.u16(o + 2)).get(g.gid);
        if (mci === undefined) return false;
        // Find what the mark sits on: the previous base (4), ligature (5) or mark (6).
        let j = i - 1;
        if (type === 6) { while (j >= 0 && this.skip(buf[j]!, flag)) j--; if (j < 0 || (this.gdef && this.glyphClass(buf[j]!.gid) !== 3)) return false; }
        else while (j >= 0 && this.glyphClass(buf[j]!.gid) === 3) j--;
        if (j < 0) return false;
        const bci = this.cov(o + r.u16(o + 4)).get(buf[j]!.gid);
        if (bci === undefined) return false;
        const classCount = r.u16(o + 6);
        const markArray = o + r.u16(o + 8), baseArray = o + r.u16(o + 10);
        if (mci >= r.u16(markArray)) return false;
        const mrec = markArray + 2 + mci * 4;
        const mclass = r.u16(mrec), manchor = markArray + r.u16(mrec + 2);
        if (mclass >= classCount) return false;
        let aOff: number;
        if (type === 5) {
          if (bci >= r.u16(baseArray)) return false;
          const la = baseArray + r.u16(baseArray + 2 + bci * 2);
          const comps = r.u16(la);
          if (!comps) return false;
          // Without per-component tracking the mark goes on the last component it follows.
          const comp = comps - 1;
          const off = r.u16(la + 2 + (comp * classCount + mclass) * 2);
          if (!off) return false;
          aOff = la + off;
        } else {
          if (bci >= r.u16(baseArray)) return false;
          const off = r.u16(baseArray + 2 + (bci * classCount + mclass) * 2);
          if (!off) return false;
          aOff = baseArray + off;
        }
        this.attach(buf, i, j, this.anchor(aOff), this.anchor(manchor));
        return true;
      }
      default: return false;
    }
  }
}

export class Gdef {
  private glyph: Map<number, number>;
  private mark: Map<number, number>;
  constructor(data: Uint8Array) {
    const r = new Reader(data);
    const gc = r.u16(4), mc = r.u16(10);
    this.glyph = classDef(r, gc);
    this.mark = classDef(r, mc);
  }
  glyphClass(g: number): number { return this.glyph.get(g) ?? 0; }
  markClass(g: number): number { return this.mark.get(g) ?? 0; }
}

// ---------- The font ----------

export class TrueTypeFont {
  readonly tables: Map<string, Uint8Array>;
  readonly unitsPerEm: number;
  readonly numGlyphs: number;
  readonly ascent: number;
  readonly descent: number;
  readonly lineGap: number;
  readonly capHeight: number;
  readonly bbox: [number, number, number, number];
  readonly italicAngle: number;
  readonly fixedPitch: boolean;
  readonly weight: number;
  readonly postscriptName: string;
  readonly gsub: Gsub | null;
  readonly gpos: Gpos | null;
  readonly gdef: Gdef | null;
  private advances: Uint16Array;
  private cmap = new Map<number, number>();
  private locaFmt: number;

  constructor(data: Uint8Array) {
    this.tables = fontTables(data);
    const t = (n: string) => this.tables.get(n);
    const head = t('head'), hhea = t('hhea'), maxp = t('maxp'), hmtx = t('hmtx');
    if (!head || !hhea || !maxp || !hmtx || !t('glyf') || !t('loca')) throw new Error('This font has no TrueType outlines.');
    const h = new Reader(head);
    this.unitsPerEm = h.u16(18) || 1000;
    this.bbox = [h.i16(36), h.i16(38), h.i16(40), h.i16(42)];
    this.locaFmt = h.i16(50);
    this.numGlyphs = new Reader(maxp).u16(4);
    const hh = new Reader(hhea);
    this.ascent = hh.i16(4); this.descent = hh.i16(6); this.lineGap = hh.i16(8);
    const nh = Math.max(1, hh.u16(34));
    const hm = new Reader(hmtx);
    this.advances = new Uint16Array(this.numGlyphs);
    let last = 0;
    for (let g = 0; g < this.numGlyphs; g++) { if (g < nh) last = hm.u16(g * 4); this.advances[g] = last; }
    const os2 = t('OS/2');
    const o = os2 ? new Reader(os2) : null;
    this.weight = o?.u16(4) ?? 400;
    this.capHeight = o && o.u16(0) >= 2 && os2!.length >= 90 ? o.i16(88) : Math.round(this.ascent * 0.7);
    if (o && os2!.length >= 72) { const ta = o.i16(68), td = o.i16(70); if (ta > 0) { this.ascent = Math.max(this.ascent, ta); this.descent = Math.min(this.descent, td); } }
    const post = t('post');
    const p = post ? new Reader(post) : null;
    this.italicAngle = p ? p.i16(4) + p.u16(6) / 65536 : 0;
    this.fixedPitch = p ? p.u32(12) !== 0 : false;
    this.postscriptName = this.readName() ?? 'Font';
    this.readCmap();
    const gdef = t('GDEF');
    this.gdef = gdef ? new Gdef(gdef) : null;
    const gsub = t('GSUB');
    this.gsub = gsub ? new Gsub(gsub, this.gdef) : null;
    const gpos = t('GPOS');
    this.gpos = gpos ? new Gpos(gpos, this.gdef) : null;
  }

  private readName(): string | undefined {
    const n = this.tables.get('name');
    if (!n) return undefined;
    const r = new Reader(n);
    const count = r.u16(2), strOff = r.u16(4);
    let best: string | undefined;
    for (let i = 0; i < count; i++) {
      const e = 6 + i * 12;
      const pid = r.u16(e), nid = r.u16(e + 6), len = r.u16(e + 8), off = r.u16(e + 10);
      if (nid !== 6) continue;
      const bytes = n.subarray(strOff + off, strOff + off + len);
      let s = '';
      if (pid === 3 || pid === 0) for (let k = 0; k + 1 < bytes.length; k += 2) s += String.fromCharCode((bytes[k]! << 8) | bytes[k + 1]!);
      else for (const b of bytes) s += String.fromCharCode(b);
      s = s.replace(/[^A-Za-z0-9-]/g, '');
      if (s) { best = s; if (pid === 3) break; }
    }
    return best;
  }

  private readCmap(): void {
    const c = this.tables.get('cmap');
    if (!c) return;
    const r = new Reader(c);
    const n = r.u16(2);
    let pick = -1, score = -1;
    for (let i = 0; i < n; i++) {
      const pid = r.u16(4 + i * 8), eid = r.u16(6 + i * 8), off = r.u32(8 + i * 8);
      const fmt = r.u16(off);
      const s = fmt === 12 && (pid === 3 && eid === 10 || pid === 0) ? 3 : fmt === 4 && (pid === 3 && eid === 1 || pid === 0) ? 2 : fmt === 4 || fmt === 12 ? 1 : 0;
      if (s > score) { score = s; pick = off; }
    }
    if (pick < 0 || score <= 0) return;
    const fmt = r.u16(pick);
    if (fmt === 4) {
      const segX2 = r.u16(pick + 6);
      const ends = pick + 14, starts = ends + segX2 + 2, deltas = starts + segX2, ranges = deltas + segX2;
      for (let s = 0; s < segX2; s += 2) {
        const end = r.u16(ends + s), start = r.u16(starts + s), delta = r.i16(deltas + s), ro = r.u16(ranges + s);
        for (let ch = start; ch <= end && ch !== 0xffff; ch++) {
          let g: number;
          if (!ro) g = (ch + delta) & 0xffff;
          else { g = r.u16(ranges + s + ro + (ch - start) * 2); if (g) g = (g + delta) & 0xffff; }
          if (g && g < this.numGlyphs) this.cmap.set(ch, g);
        }
      }
    } else if (fmt === 12) {
      const groups = r.u32(pick + 12);
      for (let i = 0; i < groups && i < 100000; i++) {
        const e = pick + 16 + i * 12;
        const s = r.u32(e), en = r.u32(e + 4), g0 = r.u32(e + 8);
        for (let ch = s; ch <= en && ch - s < 70000; ch++) if (g0 + ch - s < this.numGlyphs) this.cmap.set(ch, g0 + ch - s);
      }
    }
  }

  glyph(cp: number): number { return this.cmap.get(cp) ?? 0; }
  has(cp: number): boolean { return this.cmap.has(cp); }
  advance(gid: number): number { return this.advances[gid] ?? 0; }
  isMark(gid: number): boolean { return this.gdef?.glyphClass(gid) === 3; }

  private glyphRange(gid: number): [number, number] {
    const loca = new Reader(this.tables.get('loca')!);
    return this.locaFmt ? [loca.u32(gid * 4), loca.u32(gid * 4 + 4)] : [loca.u16(gid * 2) * 2, loca.u16(gid * 2 + 2) * 2];
  }

  /** A font file with only the given glyphs' outlines; glyph ids stay the same. */
  subset(used: Iterable<number>): Uint8Array {
    const glyf = this.tables.get('glyf')!;
    const gr = new Reader(glyf);
    const keep = new Set<number>([0]);
    const stack = [...used].filter((g) => g >= 0 && g < this.numGlyphs);
    while (stack.length) {
      const g = stack.pop()!;
      if (keep.has(g) && g !== 0) continue;
      keep.add(g);
      const [s, e] = this.glyphRange(g);
      if (e <= s || gr.i16(s) >= 0) continue;
      let p = s + 10;
      for (let guard = 0; guard < 256; guard++) {
        const flags = gr.u16(p), comp = gr.u16(p + 2);
        if (!keep.has(comp)) stack.push(comp);
        p += 4 + (flags & 1 ? 4 : 2) + (flags & 8 ? 2 : flags & 0x40 ? 4 : flags & 0x80 ? 8 : 0);
        if (!(flags & 0x20)) break;
      }
    }
    const parts: Uint8Array[] = [];
    const loca = new Uint8Array((this.numGlyphs + 1) * 4);
    const lv = new DataView(loca.buffer);
    let off = 0;
    for (let g = 0; g < this.numGlyphs; g++) {
      lv.setUint32(g * 4, off);
      if (!keep.has(g)) continue;
      const [s, e] = this.glyphRange(g);
      if (e <= s || e > glyf.length) continue;
      const len = e - s, padded = (len + 3) & ~3;
      const piece = new Uint8Array(padded);
      piece.set(glyf.subarray(s, e));
      parts.push(piece);
      off += padded;
    }
    lv.setUint32(this.numGlyphs * 4, off);
    const newGlyf = new Uint8Array(off);
    let q = 0;
    for (const p of parts) { newGlyf.set(p, q); q += p.length; }
    const head = this.tables.get('head')!.slice();
    const hv = new DataView(head.buffer);
    hv.setUint32(8, 0);
    hv.setInt16(50, 1);
    const out = new Map<string, Uint8Array>([['head', head], ['hhea', this.tables.get('hhea')!], ['maxp', this.tables.get('maxp')!], ['hmtx', this.tables.get('hmtx')!], ['loca', loca], ['glyf', newGlyf]]);
    for (const k of ['cvt ', 'fpgm', 'prep', 'OS/2']) { const t = this.tables.get(k); if (t) out.set(k, t); }
    const file = buildSfnt(out);
    const sum = checksum(file);
    new DataView(file.buffer).setUint32(sfntTableOffset(file, 'head') + 8, (0xb1b0afba - sum) >>> 0);
    return file;
  }
}

const checksum = (b: Uint8Array): number => {
  let s = 0;
  const v = new DataView(b.buffer, b.byteOffset, b.byteLength);
  const n = b.length & ~3;
  for (let i = 0; i < n; i += 4) s = (s + v.getUint32(i)) >>> 0;
  if (b.length > n) { let last = 0; for (let i = n; i < b.length; i++) last |= b[i]! << (24 - (i - n) * 8); s = (s + (last >>> 0)) >>> 0; }
  return s;
};

const sfntTableOffset = (file: Uint8Array, tag: string): number => {
  const r = new Reader(file);
  const n = r.u16(4);
  for (let i = 0; i < n; i++) if (r.tag(12 + i * 16) === tag) return r.u32(12 + i * 16 + 8);
  return 0;
};

const buildSfnt = (tables: Map<string, Uint8Array>): Uint8Array => {
  const tags = [...tables.keys()].sort();
  const n = tags.length;
  let size = 12 + n * 16;
  for (const t of tags) size += (tables.get(t)!.length + 3) & ~3;
  const out = new Uint8Array(size);
  const v = new DataView(out.buffer);
  v.setUint32(0, 0x00010000);
  v.setUint16(4, n);
  const es = Math.floor(Math.log2(n));
  v.setUint16(6, 16 * 2 ** es); v.setUint16(8, es); v.setUint16(10, n * 16 - 16 * 2 ** es);
  let off = 12 + n * 16;
  tags.forEach((t, i) => {
    const d = tables.get(t)!;
    const e = 12 + i * 16;
    for (let k = 0; k < 4; k++) out[e + k] = t.charCodeAt(k);
    v.setUint32(e + 4, checksum(d)); v.setUint32(e + 8, off); v.setUint32(e + 12, d.length);
    out.set(d, off);
    off += (d.length + 3) & ~3;
  });
  return out;
};
