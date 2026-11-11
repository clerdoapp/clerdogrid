import { describe, expect, it } from 'vitest';
import { cellKey, internStyle, newSheet, newWorkbook, type Workbook } from '../src/model/workbook';
import { exportHtml, exportJson, exportXml, type CellSource } from '../src/formats/html-export';
import { formatValue } from '../src/formats/numfmt';

const book = (): Workbook => {
  const wb = newWorkbook();
  const s = wb.sheets[0]!;
  s.name = 'Sales <2024>';
  const bold = internStyle(wb, { font: { bold: true, color: '#FF0000', name: "x'; } body {" }, fill: { color: '#FFFF00' }, border: { bottom: { style: 'double', color: '#0000FF' } }, align: { h: 'center' } });
  s.cells.set(cellKey(0, 0), { v: 'Item', s: bold });
  s.cells.set(cellKey(0, 1), { v: 'Price & tax', s: bold });
  s.cells.set(cellKey(0, 2), { v: 'Day' });
  s.cells.set(cellKey(1, 0), { v: '<script>alert(1)</script>' });
  s.cells.set(cellKey(1, 1), { v: 1.5, s: internStyle(wb, { numFmt: '0.00' }) });
  s.cells.set(cellKey(1, 2), { v: 45000, s: internStyle(wb, { numFmt: 'yyyy-mm-dd' }) });
  s.cells.set(cellKey(2, 0), { v: 'hidden row' });
  s.rows.set(2, { hidden: true });
  s.cells.set(cellKey(3, 0), { v: 'merged' });
  s.merges.push({ r1: 3, c1: 0, r2: 3, c2: 1 });
  s.cells.set(cellKey(4, 0), { v: 'safe' });
  s.hyperlinks.set(cellKey(4, 0), 'https://example.com/?a=1&b="2"');
  s.cells.set(cellKey(4, 1), { v: 'unsafe' });
  s.hyperlinks.set(cellKey(4, 1), 'javascript:alert(1)');
  s.cells.set(cellKey(5, 0), { v: true });
  s.notes.set(cellKey(5, 0), { text: 'a "note"' });
  const t = newSheet('Two');
  t.cells.set(0, { v: 'second' });
  wb.sheets.push(t);
  return wb;
};

const source = (wb: Workbook): CellSource => ({
  value: (si, r, c) => wb.sheets[si]!.cells.get(cellKey(r, c))?.v ?? null,
  text: (si, r, c) => { const cell = wb.sheets[si]!.cells.get(cellKey(r, c)); return cell ? formatValue(cell.v, wb.styles[cell.s ?? 0]?.numFmt ?? 'General').text : ''; },
  style: (si, r, c) => wb.styles[wb.sheets[si]!.cells.get(cellKey(r, c))?.s ?? 0] ?? {},
  bounds: (si) => { let rows = 0, cols = 0; for (const k of wb.sheets[si]!.cells.keys()) { rows = Math.max(rows, Math.floor(k / 16384)); cols = Math.max(cols, k % 16384); } return { rows, cols }; },
});

describe('web page export', () => {
  const wb = book();
  const html = exportHtml(wb, source(wb), { title: 'Report <x>' });
  it('escapes text and keeps only safe links', () => {
    expect(html).toContain('&lt;script&gt;alert(1)&lt;/script&gt;');
    expect(html).not.toContain('<script>');
    expect(html).toContain('<title>Report &lt;x&gt;</title>');
    expect(html).toContain('href="https://example.com/?a=1&amp;b=&quot;2&quot;"');
    expect(html).not.toContain('javascript:');
    expect(html).toContain('title="a &quot;note&quot;"');
  });
  it('writes styles inline, merges and skips hidden rows', () => {
    expect(html).toContain('font-weight:700');
    expect(html).toContain('background:#FFFF00');
    expect(html).toContain('border-bottom:3px double #0000FF');
    expect(html).toContain('colspan="2"');
    expect(html).not.toContain('hidden row');
    expect(html).toContain('>1.50<');
    expect(html).toContain('>2023-03-15<');
    expect(html).not.toMatch(/font-family:'x'; }/);
  });
  it('lists every sheet with links between them', () => {
    expect(html).toContain('<a href="#sheet2">Two</a>');
    expect(html).toContain('<h2>Sales &lt;2024&gt;</h2>');
    expect(html).toContain('second');
    const one = exportHtml(wb, source(wb), { sheets: [1] });
    expect(one).not.toContain('<nav>');
  });
});

describe('XML and JSON export', () => {
  const wb = book();
  it('writes XML with safe element names', () => {
    const xml = exportXml(wb, source(wb), { sheet: 0, area: { r1: 0, c1: 0, r2: 1, c2: 2 }, headers: true });
    expect(xml).toMatch(/^<\?xml[^\n]*\n<Sales_2024_>/);
    expect(xml).toContain('<Price_tax type="number">1.5</Price_tax>');
    expect(xml).toContain('<Item type="string">&lt;script&gt;alert(1)&lt;/script&gt;</Item>');
    expect(xml).toContain('<Day type="date">2023-03-15</Day>');
  });
  it('writes JSON objects or arrays', () => {
    const data = JSON.parse(exportJson(wb, source(wb), { sheet: 0, area: { r1: 0, c1: 0, r2: 1, c2: 2 }, headers: true }));
    expect(data).toEqual([{ Item: '<script>alert(1)</script>', 'Price & tax': 1.5, Day: '2023-03-15' }]);
    const raw = JSON.parse(exportJson(wb, source(wb), { sheet: 0, area: { r1: 1, c1: 1, r2: 1, c2: 2 }, headers: false, asShown: true }));
    expect(raw).toEqual([['1.50', '2023-03-15']]);
  });
});
