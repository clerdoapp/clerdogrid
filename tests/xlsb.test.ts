import { describe, expect, it } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import { cellKey } from '../src/model/workbook';
import { CellError } from '../src/engine/types';
import { readXlsb, xlsbRecords } from '../src/formats/xlsb-read';
import { encodeFormula, type EncodeCtx } from '../src/formats/biff8-formula';

// Builds .xlsb parts record by record (MS-XLSB) so the reader is checked against the format itself.
class B {
  b: number[] = [];
  u8(v: number) { this.b.push(v & 0xff); return this; }
  u16(v: number) { this.b.push(v & 0xff, (v >> 8) & 0xff); return this; }
  u32(v: number) { this.b.push(v & 0xff, (v >> 8) & 0xff, (v >> 16) & 0xff, (v >>> 24) & 0xff); return this; }
  f64(v: number) { this.b.push(...new Uint8Array(new Float64Array([v]).buffer)); return this; }
  wide(s: string) { this.u32(s.length); for (const c of s) this.u16(c.charCodeAt(0)); return this; }
  bytes(a: ArrayLike<number>) { for (let i = 0; i < a.length; i++) this.b.push(a[i]!); return this; }
  color(rgb?: string) { if (!rgb) return this.u8(0).u8(0).u16(0).u32(0); const n = parseInt(rgb.slice(1), 16); return this.u8(5).u8(0).u16(0).u8(n >> 16).u8(n >> 8).u8(n).u8(255); }
}
const part = (recs: [number, B | number[]][]): Uint8Array => {
  const out: number[] = [];
  for (const [type, body] of recs) {
    const data = body instanceof B ? body.b : body;
    let t = type;
    do { out.push((t & 0x7f) | (t > 0x7f ? 0x80 : 0)); t >>= 7; } while (t);
    let n = data.length;
    do { out.push((n & 0x7f) | (n > 0x7f ? 0x80 : 0)); n >>= 7; } while (n);
    out.push(...data);
  }
  return new Uint8Array(out);
};
const rel = (id: string, type: string, target: string, external = false) =>
  `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`;
const rels = (items: string[]) => strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items.join('')}</Relationships>`);

const sheets = ['Data', 'Other'];
const ctx: EncodeCtx = { xlsb: true, mode: 'cell', sheetIndex: (n) => { const i = sheets.indexOf(n); return i < 0 ? undefined : i; }, ixti: (a) => a, nameIndex: (n) => (n === 'Rate' ? 1 : undefined), addin: () => 1 };
const fmla = (b: B, f: string, row = 0, col = 0) => { const t = encodeFormula(f, { ...ctx, row, col }); return b.u16(0).u32(t.rgce.length).bytes(t.rgce).u32(t.rgcb.length).bytes(t.rgcb); };
const cell = (col: number, xf: number) => new B().u32(col).u32(xf);

const build = (): Uint8Array => {
  const wbBin = part([
    [131, []], [153, new B().u32(0)],
    [156, new B().u32(0).u32(1).wide('rId1').wide('Data')],
    [156, new B().u32(1).u32(2).wide('rId2').wide('Other')],
    [357, []], [362, new B().u32(2).u32(0).u32(0).u32(0).u32(0).u32(1).u32(1)],
    [39, (() => { const b = new B().u32(0).u8(0).u32(0xffffffff).wide('Rate'); const t = encodeFormula('Data!$A$1', { ...ctx, mode: 'name' }); return b.u32(t.rgce.length).bytes(t.rgce).u32(0); })()],
    [158, new B().u32(0).u32(0).u32(1000).u32(800).u32(600).u32(0).u32(1).u8(0)],
    [132, []],
  ]);
  const styles = part([
    [615, []], [44, new B().u16(164).wide('0.000')], [616, []],
    [611, []], [43, new B().u16(220).u16(0).u16(400).u16(0).u8(0).u8(2).u8(0).u8(0).color().u8(2).wide('Calibri')],
    [43, new B().u16(280).u16(2).u16(700).u16(0).u8(1).u8(2).u8(0).u8(0).color('#C00000').u8(0).wide('Arial')], [612, []],
    [603, []], [45, new B().u32(0).color().color().u32(0).f64(0).f64(0).f64(0).f64(0).f64(0).u32(0)],
    [45, new B().u32(1).color('#FFFF00').color().u32(0).f64(0).f64(0).f64(0).f64(0).f64(0).u32(0)], [604, []],
    [613, []], [46, new B().u8(0).u8(0).u8(0).color().u8(0).u8(0).color().u8(0).u8(0).color().u8(0).u8(0).color().u8(0).u8(0).color()],
    [46, new B().u8(0).u8(1).u8(0).color().u8(6).u8(0).color('#0000FF').u8(0).u8(0).color().u8(0).u8(0).color().u8(0).u8(0).color()], [614, []],
    [617, []],
    [47, new B().u16(0xffff).u16(0).u16(0).u16(0).u16(0).u8(0).u8(0).u16(0x1000 | (2 << 3)).u16(0)],
    [47, new B().u16(0).u16(164).u16(1).u16(1).u16(1).u8(45).u8(0).u16(0x2 | (0 << 3) | 0x40).u16(0)],
    [618, []],
  ]);
  const sst = part([[159, new B().u32(3).u32(2)], [19, new B().u8(0).wide('hello')], [19, new B().u8(1).wide('bold it').u32(2).u16(0).u16(1).u16(4).u16(0)], [160, []]]);
  const sheet1 = part([
    [129, []],
    [147, new B().u16(0).u8(0).color('#00B050').u32(0).u32(0).wide('')],
    [133, []], [137, new B().u16(0x4 | 0x8).u32(0).u32(0).u32(0).u8(64).u8(0).u16(0).u16(150).u16(0).u16(0).u16(0).u32(0)],
    [151, new B().f64(1).f64(2).u32(2).u32(1).u32(0).u8(1)], [138, []], [134, []],
    [485, new B().u32(0xffffffff).u16(8).u16(300).u16(0).u8(0).u8(0)],
    [390, []], [60, new B().u32(0).u32(0).u32(Math.round(20.71 * 256)).u32(0).u16(2)], [60, new B().u32(3).u32(3).u32(2000).u32(0).u16(3)], [391, []],
    [145, []],
    [0, new B().u32(0).u32(0).u16(600).u16(0x2000).u8(0).u32(0)],
    [5, cell(0, 0).f64(1.5)], [2, cell(1, 1).u32((10 << 2) | 2)], [7, cell(2, 0).u32(0)], [4, cell(3, 0).u8(1)], [3, cell(4, 0).u8(0x07)],
    [6, cell(5, 0).wide('inline')], [7, cell(6, 0).u32(1)],
    [0, new B().u32(1).u32(0).u16(300).u16(0x1000).u8(0).u32(0)],
    [9, fmla(cell(0, 0).f64(11.5), 'A1+B1')], [8, fmla(cell(1, 0).wide('helloX'), 'C1&"X"')], [10, fmla(cell(2, 0).u8(1), 'AND(D1,TRUE)')],
    [9, fmla(cell(3, 0).f64(3), 'Other!A1+Rate*2')], [9, fmla(cell(4, 0).f64(2), 'SUM({1,1})')], [11, fmla(cell(5, 0).u8(0x2a), 'IFERROR(NA(),NA())')],
    [0, new B().u32(2).u32(0).u16(300).u16(0).u8(0).u32(0)],
    [9, cell(0, 0).f64(2).u16(0).u32(5).bytes([0x01, 2, 0, 0, 0]).u32(0)],
    [425, (() => { const b = new B().u32(2).u32(3).u32(0).u32(0); const t = encodeFormula('B3*2', { ...ctx, mode: 'rel', row: 2, col: 0 }); return b.u32(t.rgce.length).bytes(t.rgce).u32(0); })()],
    [0, new B().u32(3).u32(0).u16(300).u16(0).u8(0).u32(0)],
    [9, cell(0, 0).f64(4).u16(0).u32(5).bytes([0x01, 2, 0, 0, 0]).u32(0)],
    [146, []],
    [177, new B().u32(1)], [176, new B().u32(5).u32(6).u32(0).u32(2)], [178, []],
    [494, new B().u32(0).u32(0).u32(6).u32(6).wide('rId9').wide('').wide('').wide('')],
    [494, new B().u32(1).u32(1).u32(6).u32(6).wide('').wide("Other!A1").wide('').wide('')],
    [494, new B().u32(2).u32(2).u32(6).u32(6).wide('rId8').wide('').wide('').wide('')],
    [535, new B().u16(0xcc3d).u32(1).u32(1).u32(1).u32(0)],
    [476, new B().f64(0.5).f64(0.6).f64(0.7).f64(0.8).f64(0.2).f64(0.3)],
    [478, new B().u32(9).u32(80).u32(600).u32(600).u32(1).u32(1).u32(1).u32(1).u16(2).wide('')],
    [161, new B().u32(0).u32(3).u32(0).u32(2)], [162, []],
    [130, []],
  ]);
  const sheet2 = part([[129, []], [145, []], [0, new B().u32(0).u32(0).u16(300).u16(0).u8(0).u32(0)], [2, cell(0, 0).u32((5 << 2) | 2)], [146, []], [130, []]]);
  const comments = part([[628, []], [630, []], [632, new B().wide('Ann')], [631, []], [633, []], [635, new B().u32(0).u32(0).u32(0).u32(0).u32(0).bytes(new Array(16).fill(0))], [637, new B().u8(0).wide('Note text')], [636, []], [634, []], [629, []]]);
  return zipSync({
    '[Content_Types].xml': strToU8('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'),
    '_rels/.rels': rels([rel('rId1', 'officeDocument', 'xl/workbook.bin')]),
    'xl/workbook.bin': wbBin,
    'xl/_rels/workbook.bin.rels': rels([rel('rId1', 'worksheet', 'worksheets/sheet1.bin'), rel('rId2', 'worksheet', 'worksheets/sheet2.bin'), rel('rId3', 'styles', 'styles.bin'), rel('rId4', 'sharedStrings', 'sharedStrings.bin')]),
    'xl/styles.bin': styles,
    'xl/sharedStrings.bin': sst,
    'xl/worksheets/sheet1.bin': sheet1,
    'xl/worksheets/_rels/sheet1.bin.rels': rels([rel('rId9', 'hyperlink', 'https://example.com/', true), rel('rId8', 'hyperlink', 'javascript:alert(1)', true), rel('rId7', 'comments', '../comments1.bin')]),
    'xl/worksheets/sheet2.bin': sheet2,
    'xl/comments1.bin': comments,
  });
};

describe('.xlsb reading', () => {
  const wb = readXlsb(build());
  const s = wb.sheets[0]!;
  const c = (r: number, col: number) => s.cells.get(cellKey(r, col))!;

  it('splits records with variable-length headers', () => {
    expect(xlsbRecords(new Uint8Array([0x80, 0x01, 0x02, 0xaa, 0xbb, 0x00, 0x00]))).toEqual([{ type: 128, data: new Uint8Array([0xaa, 0xbb]) }, { type: 0, data: new Uint8Array(0) }]);
  });
  it('reads sheets, states and the active tab', () => {
    expect(wb.sheets.map((x) => [x.name, x.state])).toEqual([['Data', 'visible'], ['Other', 'hidden']]);
    expect(wb.activeSheet).toBe(1);
    expect(wb.sheets[1]!.cells.get(0)!.v).toBe(5);
  });
  it('reads every cell type and shared strings with runs', () => {
    expect(c(0, 0).v).toBe(1.5);
    expect(c(0, 1).v).toBe(10);
    expect(c(0, 2).v).toBe('hello');
    expect(c(0, 3).v).toBe(true);
    expect((c(0, 4).v as CellError).code).toBe('#DIV/0!');
    expect(c(0, 5).v).toBe('inline');
    expect(c(0, 6).v).toBe('bold it');
    expect(c(0, 6).rich).toEqual([{ text: 'bold', font: { name: 'Arial', size: 14, bold: true, italic: true, underline: 'single', color: '#C00000' } }, { text: ' it' }]);
  });
  it('decodes formulas, names and shared formulas', () => {
    expect(c(1, 0)).toMatchObject({ f: 'A1+B1', v: 11.5 });
    expect(c(1, 1)).toMatchObject({ f: 'C1&"X"', v: 'helloX' });
    expect(c(1, 2)).toMatchObject({ f: 'AND(D1,TRUE)', v: true });
    expect(c(1, 3).f).toBe('Other!A1+Rate*2');
    expect(c(1, 4).f).toBe('SUM({1,1})');
    expect(c(1, 5).f).toBe('IFERROR(NA(),NA())');
    expect(c(2, 0).f).toBe('B3*2');
    expect(c(3, 0).f).toBe('B4*2');
    expect(wb.names).toEqual([{ name: 'Rate', formula: 'Data!$A$1' }]);
  });
  it('reads styles', () => {
    const st = wb.styles[c(0, 1).s!]!;
    expect(st.font).toMatchObject({ name: 'Arial', size: 14, bold: true, italic: true, color: '#C00000' });
    expect(st.fill).toEqual({ color: '#FFFF00' });
    expect(st.border?.bottom).toEqual({ style: 'double', color: '#0000FF' });
    expect(st.numFmt).toBe('0.000');
    expect(st.align).toMatchObject({ h: 'center', wrap: true, rotation: 45 });
    expect(st.protection).toEqual({ locked: false });
  });
  it('reads layout, view, links, notes, protection and page setup', () => {
    expect(s.rows.get(0)!.height).toBe(30);
    expect(s.rows.get(1)!.hidden).toBe(true);
    expect(s.cols.get(0)!.width).toBeCloseTo(20, 1);
    expect(s.cols.get(3)!.hidden).toBe(true);
    expect(s.defaultRowHeight).toBe(15);
    expect(s.view).toMatchObject({ frozenRows: 2, frozenCols: 1, zoom: 150 });
    expect(s.tabColor).toBe('#00B050');
    expect(s.merges).toEqual([{ r1: 5, r2: 6, c1: 0, c2: 2 }]);
    expect(s.hyperlinks.get(cellKey(0, 6))).toBe('https://example.com/');
    expect(s.hyperlinks.get(cellKey(1, 6))).toBe('#Other!A1');
    expect(s.hyperlinks.has(cellKey(2, 6))).toBe(false);
    expect(s.notes.get(0)).toEqual({ author: 'Ann', text: 'Note text' });
    expect(s.protection).toMatchObject({ sheet: true, password: 'CC3D' });
    expect(s.pageSetup).toMatchObject({ paper: 9, scale: 80, orientation: 'landscape', margins: { left: 0.5, right: 0.6, top: 0.7, bottom: 0.8, header: 0.2, footer: 0.3 } });
    expect(s.autoFilter?.range).toEqual({ r1: 0, r2: 3, c1: 0, c2: 2 });
  });
});
