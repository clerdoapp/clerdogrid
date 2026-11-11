import { beforeAll, describe, expect, it } from 'vitest';
import { execFileSync } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { strToU8, zlibSync } from 'fflate';
import { GridStore } from '../src/app/store';
import { registerAllFunctions } from '../src/engine/functions';
import { cellKey, internStyle } from '../src/model/workbook';
import { PdfDocument } from '../src/formats/pdf';
import { FontBook, bidiLevels } from '../src/formats/pdf-text';
import { TrueTypeFont } from '../src/formats/pdf-font';
import { decodePng, jpegInfo } from '../src/formats/pdf-image';
import { buildPdf, loadFonts, pageToSvg, planPrint } from '../src/features/pagelayout/print';
import { parsePathData } from '../src/features/pagelayout/painter';
import { nodeFontSource } from './pdfFonts';

beforeAll(() => registerAllFunctions());

const latin1 = (b: Uint8Array) => { let s = ''; for (let i = 0; i < b.length; i++) s += String.fromCharCode(b[i]!); return s; };

/** Checks the cross-reference table: every offset points at its object. */
const checkStructure = (pdf: Uint8Array) => {
  const s = latin1(pdf);
  expect(s.startsWith('%PDF-1.7\n')).toBe(true);
  expect(s.trimEnd().endsWith('%%EOF')).toBe(true);
  const sx = /startxref\n(\d+)\n%%EOF\s*$/.exec(s);
  expect(sx).toBeTruthy();
  const xref = +sx![1]!;
  expect(s.slice(xref, xref + 4)).toBe('xref');
  const head = /^xref\n0 (\d+)\n/.exec(s.slice(xref))!;
  const n = +head[1]!;
  const table = s.slice(xref + head[0].length, xref + head[0].length + n * 20);
  for (let i = 1; i < n; i++) {
    const e = table.slice(i * 20, i * 20 + 20);
    expect(e).toMatch(/^\d{10} 00000 n \n$/);
    const off = +e.slice(0, 10);
    expect(s.slice(off, off + `${i} 0 obj`.length)).toBe(`${i} 0 obj`);
  }
  expect(s).toMatch(new RegExp(`trailer\\n<</Size ${n}/Root \\d+ 0 R/Info \\d+ 0 R/ID\\[<[0-9A-F]{32}><[0-9A-F]{32}>\\]>>`));
  return s;
};

const tool = (name: string) => ['/usr/bin/', '/usr/local/bin/'].map((d) => d + name).find((p) => existsSync(p));

describe('text layout', () => {
  it('orders mixed-direction text', () => {
    const cps = [...'abc אבג 123'].map((c) => c.codePointAt(0)!);
    // Digits after Hebrew stay right-to-left context numbers (level 2), as UAX #9 says.
    expect([...bidiLevels(cps).levels]).toEqual([0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2]);
    expect([...bidiLevels([...'abc 123'].map((c) => c.codePointAt(0)!)).levels]).toEqual([0, 0, 0, 0, 0, 0, 0]);
    const ar = [...'مرحبا 12'].map((c) => c.codePointAt(0)!);
    const r = bidiLevels(ar);
    expect(r.base).toBe(1);
    expect([...r.levels]).toEqual([1, 1, 1, 1, 1, 1, 2, 2]);
  });

  it('joins Arabic letters and forms lam-alef', async () => {
    const fb = new FontBook(nodeFontSource());
    await fb.ensure('بب ب لا', {});
    const gids = (t: string) => fb.shape(t, {}).runs[0]!.glyphs.map((g) => g.gid);
    const alone = gids('ب');
    expect(gids('بب')).not.toEqual([...alone, ...alone]);
    const la = fb.shape('لا', {}).runs[0]!.glyphs;
    expect(la.map((g) => g.gid)).not.toEqual([...gids('ا'), ...gids('ل')]);
    expect(la.map((g) => g.chars).reverse().join('')).toBe('لا');
  });

  it('falls back across bundled fonts', async () => {
    const fb = new FontBook(nodeFontSource());
    const t = 'Abc Привет Ελλάδα שלום';
    await fb.ensure(t, { family: 'Calibri' });
    const shaped = fb.shape(t, { family: 'Calibri' });
    expect(shaped.runs.some((r) => r.missing)).toBe(false);
    expect(new Set(shaped.runs.map((r) => r.font!.file.split('-')[0])).size).toBeGreaterThan(1);
  });

  it('subsets a font and keeps glyph ids', async () => {
    const fb = new FontBook(nodeFontSource());
    await fb.ensure('Hi', { family: 'Arial' });
    const lf = fb.primary({ family: 'Arial' })!;
    const g = lf.font.glyph(72);
    const sub = new TrueTypeFont(lf.font.subset([g]));
    expect(sub.numGlyphs).toBe(lf.font.numGlyphs);
    expect(sub.advance(g)).toBe(lf.font.advance(g));
  });
});

describe('images', () => {
  const chunk = (type: string, data: Uint8Array) => {
    const out = new Uint8Array(12 + data.length);
    new DataView(out.buffer).setUint32(0, data.length);
    out.set(strToU8(type), 4); out.set(data, 8);
    return out;
  };
  const png = (w: number, h: number, depth: number, ctype: number, rows: number[][], extra: Uint8Array[] = []) => {
    const ihdr = new Uint8Array(13);
    const v = new DataView(ihdr.buffer);
    v.setUint32(0, w); v.setUint32(4, h); ihdr[8] = depth; ihdr[9] = ctype;
    const raw = Uint8Array.from(rows.flatMap((r) => [0, ...r]));
    const parts = [Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr), ...extra, chunk('IDAT', zlibSync(raw)), chunk('IEND', new Uint8Array())];
    const out = new Uint8Array(parts.reduce((a, p) => a + p.length, 0));
    let o = 0;
    for (const p of parts) { out.set(p, o); o += p.length; }
    return out;
  };
  it('decodes RGBA, palette with transparency and 1-bit gray PNGs', () => {
    const rgba = decodePng(png(2, 1, 8, 6, [[255, 0, 0, 255, 0, 0, 255, 128]]));
    expect([...rgba.data]).toEqual([255, 0, 0, 0, 0, 255]);
    expect([...rgba.alpha!]).toEqual([255, 128]);
    const pal = decodePng(png(2, 1, 8, 3, [[0, 1]], [chunk('PLTE', Uint8Array.from([10, 20, 30, 40, 50, 60])), chunk('tRNS', Uint8Array.from([0]))]));
    expect([...pal.data]).toEqual([10, 20, 30, 40, 50, 60]);
    expect([...pal.alpha!]).toEqual([0, 255]);
    const bit = decodePng(png(3, 1, 1, 0, [[0b10100000]]));
    expect([...bit.data]).toEqual([255, 0, 255]);
    expect(bit.alpha).toBeUndefined();
  });
  it('reads JPEG sizes', () => {
    const j = Uint8Array.from([0xff, 0xd8, 0xff, 0xe0, 0, 4, 0, 0, 0xff, 0xc0, 0, 11, 8, 0, 20, 0, 30, 3, 1, 0x22, 0]);
    expect(jpegInfo(j)).toEqual({ width: 30, height: 20, components: 3, adobeInverted: false });
  });
});

describe('path data', () => {
  it('reads relative commands, curves and arcs', () => {
    const p = parsePathData('M10 10h5v5l-5 0zM0,0 Q5,5 10,0 A5 5 0 0 1 20 0');
    expect(p[0]).toEqual(['M', 10, 10]);
    expect(p[1]).toEqual(['L', 15, 10]);
    expect(p[2]).toEqual(['L', 15, 15]);
    expect(p[3]).toEqual(['L', 10, 15]);
    expect(p[4]).toEqual(['Z']);
    expect(p[6]![0]).toBe('C');
    const end = p.at(-1)!.slice(-2) as number[];
    expect(end[0]).toBeCloseTo(20, 6);
    expect(end[1]).toBeCloseTo(0, 6);
  });
});

describe('PDF documents', () => {
  it('writes a valid file with fonts, ToUnicode, images, gradients and links', async () => {
    const fb = new FontBook(nodeFontSource());
    await fb.ensure('Hello مرحبا', {});
    const doc = new PdfDocument();
    const p = doc.addPage(612, 792);
    p.rect(10, 10, 100, 20, '#ff000080');
    p.gradientRect(10, 40, 100, 20, 0, [{ pos: 0, color: '#000000' }, { pos: 0.5, color: '#ffffff' }, { pos: 1, color: '#00ff00' }]);
    p.text(fb.shape('Hello مرحبا', {}), 20, 100, 12, '#000000');
    p.link(20, 90, 100, 12, 'https://clerdo.app');
    p.link(20, 90, 100, 12, 'javascript:alert(1)');
    p.end();
    doc.bookmark('Sheet1', 0);
    const s = checkStructure(doc.finish({ title: 'Tést' }));
    expect(s).toContain('/Subtype/CIDFontType2');
    expect(s).toContain('/Encoding/Identity-H');
    expect(s).toContain('/ToUnicode');
    expect(s).toContain('/FontFile2');
    expect(s).toContain('/ShadingType 2');
    expect(s).toContain('/URI (https://clerdo.app)');
    expect(s).not.toContain('javascript');
    expect(s).toContain('/Title <FEFF005400E900730074>');
  });

  it('exports a formatted sheet that other tools read back', async () => {
    const s = new GridStore();
    s.enter(0, 0, 'Region'); s.enter(0, 1, 'Sales'); s.enter(0, 2, 'Notes');
    const rows: [string, number, string][] = [['North', 1234.5, 'Grew in spring'], ['South', -250, 'الجنوب'], ['East', 0.125, 'Long text that wraps inside the cell']];
    rows.forEach(([a, b, c], i) => { s.enter(i + 1, 0, a); s.enter(i + 1, 1, String(b)); s.enter(i + 1, 2, c); });
    s.enter(5, 1, '=SUM(B2:B4)'); s.enter(6, 1, '=1/0');
    s.engine.recalc();
    const sh = s.sheet;
    sh.cells.get(cellKey(0, 0))!.s = internStyle(s.wb, { font: { bold: true, color: '#ffffff', size: 14 }, fill: { color: '#0c7a4b' }, border: { bottom: { style: 'double', color: '#000000' } } });
    sh.cells.get(cellKey(1, 1))!.s = internStyle(s.wb, { numFmt: '#,##0.00', border: { top: { style: 'dashed' }, left: { style: 'thick', color: '#c00000' } } });
    sh.cells.get(cellKey(3, 2))!.s = internStyle(s.wb, { align: { wrap: true, v: 'top' } });
    sh.cells.get(cellKey(1, 2))!.s = internStyle(s.wb, { align: { rotation: 45 } });
    sh.merges.push({ r1: 8, c1: 0, r2: 9, c2: 2 });
    s.enter(8, 0, 'Merged title');
    sh.cells.get(cellKey(8, 0))!.s = internStyle(s.wb, { align: { h: 'center', v: 'center' }, fill: { gradient: { angle: 0, stops: [{ pos: 0, color: '#ffffff' }, { pos: 1, color: '#2f6fdf' }] } } });
    sh.conditionalFormats.push({ ranges: [{ r1: 1, c1: 1, r2: 3, c2: 1 }], rules: [{ type: 'dataBar', min: { type: 'min' }, max: { type: 'max' }, color: '#638ec6' }] });
    sh.charts.push({ id: 'c1', type: 'column', title: 'Sales', series: [{ name: 'Sales', values: 'Sheet1!$B$2:$B$4', categories: 'Sheet1!$A$2:$A$4' }], anchor: { row: 11, col: 0, dx: 0, dy: 0, width: 320, height: 200 } });
    sh.shapes = [{ id: 's1', kind: 'roundRect', anchor: { row: 11, col: 6, dx: 0, dy: 0, width: 140, height: 60 }, text: 'Shape text', fill: '#ffe066', line: { color: '#0c7a4b', width: 2 } }];
    sh.hyperlinks.set(cellKey(1, 0), 'https://clerdo.app/north');
    const logo = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='), (c) => c.charCodeAt(0));
    sh.pageSetup = { header: '&L&A&C&"Arial,Bold"&16Quarterly report&RPage &P of &N', footer: '&C&F&R&G', gridlines: true, headings: true, errors: 'NA', printTitleRows: '1:1', headerPictures: { RF: { data: logo, mime: 'image/png', width: 30, height: 15 } } };
    const fonts = new FontBook(nodeFontSource());
    const pdf = await buildPdf(s, { what: 'sheet' }, { fonts, date: new Date(2026, 10, 8) });
    const text = checkStructure(pdf);
    expect(text).toContain('/Count 1');
    expect(text).toContain('/Subtype/Image');
    expect(text).toContain('/SMask');
    expect(text).toContain('/Subtype/Link');
    const dir = mkdtempSync(join(tmpdir(), 'pl-pdf-'));
    const file = join(dir, 'sheet.pdf');
    writeFileSync(file, pdf);
    const qpdf = tool('qpdf');
    if (qpdf) expect(execFileSync(qpdf, ['--check', file]).toString()).toContain('No syntax or stream encoding errors');
    const pdftotext = tool('pdftotext');
    if (pdftotext) {
      const out = execFileSync(pdftotext, ['-layout', file, '-']).toString();
      for (const want of ['984.625', 'Quarterly report', 'Page 1 of 1', 'Region', '1,234.50', '#N/A', 'Merged title', 'Sales', 'Shape text', 'Sheet1', 'Book1.xlsx', 'الجنوب']) expect(out).toContain(want);
      expect(out).not.toContain('#DIV/0!');
    }
    const pdftoppm = tool('pdftoppm');
    if (pdftoppm) {
      execFileSync(pdftoppm, ['-r', '60', '-png', '-f', '1', '-l', '1', file, join(dir, 'pg')]);
      const img = readFileSync(join(dir, 'pg-1.png'));
      expect(img.length).toBeGreaterThan(5000);
      if (process.env.PDF_SNAPSHOT) writeFileSync(process.env.PDF_SNAPSHOT, img);
    }
  }, 60_000);

  it('makes one page per page in an SVG for preview and printing', async () => {
    const s = new GridStore();
    s.enter(0, 0, 'a < b & c');
    s.sheet.pageSetup = { footer: '&CPage &P' };
    const fonts = new FontBook(nodeFontSource());
    const job = planPrint(s, { what: 'sheet' });
    await loadFonts(s, job, { fonts });
    const svg = pageToSvg(s, job.pages[0]!, { fonts });
    expect(svg).toMatch(/^<svg xmlns="http:\/\/www.w3.org\/2000\/svg" viewBox="0 0 612 792"/);
    expect(svg).toContain('a &lt; b &amp; c');
    expect(svg).toContain('>Page 1</text>');
  });

  it('writes the whole workbook with a bookmark per sheet', async () => {
    const s = new GridStore();
    s.enter(0, 0, 'one');
    s.wb.sheets.push({ ...structuredClone(s.sheet), name: 'Second' });
    const pdf = await buildPdf(s, { what: 'workbook' }, { fonts: new FontBook(nodeFontSource()) });
    const t = checkStructure(pdf);
    expect(t).toContain('/Count 2');
    expect(t).toContain('/Title (Second)');
  });
});

describe('PDF export without the app', () => {
  it('exports a workbook read from .xlsx with the bundled fonts', async () => {
    const { exportPdf } = await import('../src/formats/pdf-export');
    const { newWorkbook } = await import('../src/model/workbook');
    const wb = newWorkbook();
    wb.sheets[0]!.cells.set(0, { v: 'From the command line' });
    const pdf = await exportPdf(wb, { sheet: 0, fileName: 'x.xlsx' });
    const t = checkStructure(pdf);
    expect(t).toContain('/FontFile2');
  });
});

describe('glyph positioning and more scripts', () => {
  it('kerns pairs and places Arabic marks on their letters', async () => {
    const fb = new FontBook(nodeFontSource());
    await fb.ensure('AVA بِ', {});
    const av = fb.shape('AV', {}).runs[0]!;
    const a = fb.shape('A', {}).width, v = fb.shape('V', {}).width;
    expect(av.width).toBeLessThan(a + v - 5);
    expect(av.glyphs[1]!.x).toBeCloseTo(av.glyphs[0]!.adv, 5);
    const beh = fb.shape('بِ', {}).runs[0]!.glyphs;
    const kasra = beh.find((g) => g.chars === 'ِ')!;
    const base = beh.find((g) => g.chars === 'ب')!;
    expect(kasra.y).toBeLessThan(-50);
    expect(kasra.x).toBeGreaterThanOrEqual(base.x - 50);
    expect(kasra.x).toBeLessThan(base.x + base.adv + 50);
  });

  it('embeds Chinese, Japanese and Korean as real text, loading only the chunks it needs', async () => {
    const fb = new FontBook(nodeFontSource());
    const text = '中文 日本語のかな 한국어';
    await fb.ensure(text, {});
    const shaped = fb.shape(text, {});
    expect(shaped.runs.some((r) => r.missing)).toBe(false);
    const files = fb.allLoaded().map((f) => f.file);
    expect(files.some((f) => f.startsWith('noto-sans-sc-'))).toBe(true);
    expect(files.some((f) => f.startsWith('noto-sans-kr-'))).toBe(true);
    expect(files.length).toBeLessThan(20);
    expect(files.some((f) => /chinese-simplified|japanese|korean/.test(f))).toBe(false);
    // Kana in the text makes Japanese the first choice for Han characters.
    expect(fb.shape('日本語のかな', {}).runs[0]!.font!.file).toMatch(/^noto-sans-jp-/);
    const doc = new PdfDocument();
    const p = doc.addPage(300, 100);
    p.text(shaped, 10, 50, 14, '#000000');
    p.end();
    const pdf = doc.finish();
    checkStructure(pdf);
    const pdftotext = tool('pdftotext');
    if (pdftotext) {
      const dir = mkdtempSync(join(tmpdir(), 'pl-cjk-'));
      writeFileSync(join(dir, 'c.pdf'), pdf);
      const out = execFileSync(pdftotext, [join(dir, 'c.pdf'), '-']).toString();
      for (const w of ['中文', '日本語のかな', '한국어']) expect(out).toContain(w);
    }
  });

  it('uses the monochrome emoji font under Node', async () => {
    const fb = new FontBook(nodeFontSource());
    await fb.ensure('ok 😀', {});
    const run = fb.shape('ok 😀', {}).runs.find((r) => r.text.includes('😀'))!;
    expect(run.missing).toBeFalsy();
    expect(run.font!.file).toMatch(/^noto-emoji-/);
  });

  it('greys pictures in black and white mode', () => {
    const png = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='), (c) => c.charCodeAt(0));
    const jpg = Uint8Array.from([0xff, 0xd8, 0xff, 0xc0, 0, 11, 8, 0, 1, 0, 1, 3, 1, 0x22, 0, 0xff, 0xd9]);
    const doc = new PdfDocument();
    const p = doc.addPage(100, 100);
    p.image(png, 0, 0, 10, 10, undefined, true);
    p.image(jpg, 20, 0, 10, 10, undefined, true);
    p.end();
    const s = checkStructure(doc.finish());
    expect(s).toMatch(/\/Width 1\/Height 1\/ColorSpace\/DeviceGray\/BitsPerComponent 8\/SMask/);
    expect(s).not.toMatch(/\/ColorSpace\/DeviceRGB\/BitsPerComponent 8\/SMask/);
    expect(s).toContain('/BM/Saturation');
  });

  it('takes chart colours from the shared palette', async () => {
    const { CHART_LIGHT, chartPaletteCss } = await import('../src/ui/charts/palette');
    const { chartColorVars, cssColor } = await import('../src/features/pagelayout/painter');
    expect(chartColorVars()['--chart-1']).toBe(CHART_LIGHT['chart-1']);
    expect(cssColor('var(--chart-grid)')).toBe(CHART_LIGHT['chart-grid']);
    expect(chartPaletteCss()).toContain(`--chart-1: ${CHART_LIGHT['chart-1']};`);
  });
});
