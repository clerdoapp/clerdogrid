import { beforeAll, describe, expect, it } from 'vitest';
import { strFromU8, unzipSync } from 'fflate';
import { GridStore } from '../src/app/store';
import { registerAllFunctions } from '../src/engine/functions';
import { cellKey, internStyle, newWorkbook, parseAddr, type ShapeKind } from '../src/model/workbook';
import { readXlsx, writeXlsx } from '../src/formats/xlsx';
import {
  addComment, addEditRange, addShape, addStroke, clearInk, defaultShape, deleteComment, editComment, eraseStrokes, insertCheckboxes, orderShape,
  removeCheckboxes, replyComment, resolveComment, strokesAt, threadAt, toggleCheckboxes, unlockEditRange, updateShape,
} from '../src/features/objects/ops';
import { accessibilityIssues, cleanExcessFormatting, contrastRatio, excessFormatting, sheetStats, wordsIn } from '../src/features/objects/checks';

beforeAll(() => registerAllFunctions());

const setup = () => {
  const s = new GridStore();
  const type = (a: string, text: string) => { const p = parseAddr(a)!; s.select(p); s.beginEdit(text); expect(s.commitEdit()).toBe(true); };
  const val = (a: string) => { const p = parseAddr(a)!; return s.value(p.row, p.col); };
  const sel = (a: string) => { const [p, q] = a.split(':'); const x = parseAddr(p!)!, y = parseAddr(q ?? p!)!; s.selectRange({ r1: x.row, c1: x.col, r2: y.row, c2: y.col }); };
  return { s, type, val, sel };
};

const roundTrip = (s: GridStore) => readXlsx(writeXlsx(s.wb));

describe('comment threads', () => {
  it('adds, replies, edits, resolves and deletes, each undoable', () => {
    const { s } = setup();
    addComment(s, 1, 1, 'Check this total', 'Ana');
    expect(threadAt(s, 1, 1)).toMatchObject({ author: 'Ana', text: 'Check this total', replies: [] });
    replyComment(s, 1, 1, 'Fixed it', 'Ben');
    expect(threadAt(s, 1, 1)!.replies).toHaveLength(1);
    editComment(s, 1, 1, 0, 'Fixed it now');
    expect(threadAt(s, 1, 1)!.replies[0]!.text).toBe('Fixed it now');
    resolveComment(s, 1, 1, true);
    expect(threadAt(s, 1, 1)!.resolved).toBe(true);
    s.undo();
    expect(threadAt(s, 1, 1)!.resolved).toBeUndefined();
    deleteComment(s, 1, 1, 0);
    expect(threadAt(s, 1, 1)!.replies).toHaveLength(0);
    deleteComment(s, 1, 1);
    expect(threadAt(s, 1, 1)).toBeUndefined();
    expect(s.sheet.comments).toBeUndefined();
    s.undo();
    expect(threadAt(s, 1, 1)!.text).toBe('Check this total');
    s.undo(); s.undo();
    expect(threadAt(s, 1, 1)!.replies[0]!.text).toBe('Fixed it');
    s.redo();
    expect(threadAt(s, 1, 1)!.replies[0]!.text).toBe('Fixed it now');
  });

  it('ignores empty text', () => {
    const { s } = setup();
    addComment(s, 0, 0, '   ');
    expect(s.sheet.comments).toBeUndefined();
    expect(s.canUndo).toBe(false);
  });
});

describe('checkboxes', () => {
  it('formats cells, toggles them in one undo step, and formulas read TRUE/FALSE', () => {
    const { s, type, val, sel } = setup();
    type('B1', '=COUNTIF(A1:A3,TRUE)');
    type('C1', '=IF(A1,"done","open")');
    sel('A1:A3');
    insertCheckboxes(s);
    expect(val('A1')).toBe(false);
    expect(s.style(0, 0).checkbox).toBe(true);
    expect(val('C1')).toBe('open');
    toggleCheckboxes(s, [{ row: 0, col: 0 }]);
    expect(val('A1')).toBe(true);
    expect(val('B1')).toBe(1);
    expect(val('C1')).toBe('done');
    sel('A1:A3');
    toggleCheckboxes(s);
    expect([val('A1'), val('A2'), val('A3')]).toEqual([true, true, true]);
    expect(val('B1')).toBe(3);
    s.undo();
    expect([val('A1'), val('A2'), val('A3')]).toEqual([true, false, false]);
    expect(val('B1')).toBe(1);
    removeCheckboxes(s);
    expect(s.style(0, 0).checkbox).toBeUndefined();
    expect(val('A1')).toBe(true);
  });

  it('keeps existing values and skips formulas', () => {
    const { s, type, val, sel } = setup();
    type('A1', 'TRUE');
    type('A2', '=1=1');
    sel('A1:A2');
    insertCheckboxes(s);
    expect(val('A1')).toBe(true);
    expect(toggleCheckboxes(s, [{ row: 1, col: 0 }])).toBe(false);
    expect(val('A2')).toBe(true);
  });

  it('protected cells do not toggle', () => {
    const { s, sel } = setup();
    sel('A1');
    insertCheckboxes(s);
    s.sheet.protection = { sheet: true };
    expect(toggleCheckboxes(s, [{ row: 0, col: 0 }])).toBe(false);
  });
});

describe('shapes', () => {
  it('adds, formats, orders and undoes', () => {
    const { s } = setup();
    const a = defaultShape('rect', { row: 1, col: 1, dx: 0, dy: 0, width: 100, height: 50 });
    const b = defaultShape('ellipse', { row: 2, col: 2, dx: 0, dy: 0, width: 80, height: 80 });
    addShape(s, a); addShape(s, b);
    updateShape(s, a.id, { fill: '#FF0000', text: 'Hi' });
    expect(s.sheet.shapes![0]!.fill).toBe('#FF0000');
    orderShape(s, a.id, 'forward');
    expect(s.sheet.shapes!.map((x) => x.id)).toEqual([b.id, a.id]);
    s.undo();
    expect(s.sheet.shapes!.map((x) => x.id)).toEqual([a.id, b.id]);
    s.undo();
    expect(s.sheet.shapes![0]!.fill).toBe('#D6EFE2');
  });
});

describe('ink', () => {
  it('adds, erases by touching and clears', () => {
    const { s } = setup();
    addStroke(s, { id: 'k1', tool: 'pen', color: '#000000', width: 2, points: [[10, 10], [100, 10]] });
    addStroke(s, { id: 'k2', tool: 'highlighter', color: '#ffe066', width: 12, points: [[10, 50], [100, 50]] });
    expect(strokesAt(s.sheet.ink!, [50, 12], 3).map((k) => k.id)).toEqual(['k1']);
    eraseStrokes(s, new Set(['k1']));
    expect(s.sheet.ink!.map((k) => k.id)).toEqual(['k2']);
    clearInk(s);
    expect(s.sheet.ink).toBeUndefined();
    s.undo(); s.undo();
    expect(s.sheet.ink).toHaveLength(2);
  });
});

describe('save and reopen', () => {
  it('keeps comment threads exactly and writes them as legacy comments too', () => {
    const { s } = setup();
    addComment(s, 0, 0, 'First', 'Ana');
    replyComment(s, 0, 0, 'Second', 'Ben');
    resolveComment(s, 0, 0, true);
    s.sheet.notes.set(cellKey(3, 3), { author: 'Cy', text: 'A plain note' });
    const bytes = writeXlsx(s.wb);
    const files = unzipSync(bytes);
    const legacy = strFromU8(files['xl/comments1.xml']!);
    expect(legacy).toContain('Ana: First');
    expect(legacy).toContain('Ben: Second');
    expect(strFromU8(files['[Content_Types].xml']!)).toContain('/xl/clerdogrid/objects.json');
    const wb = readXlsx(bytes);
    const t = wb.sheets[0]!.comments!.get(cellKey(0, 0))!;
    expect(t).toMatchObject({ author: 'Ana', text: 'First', resolved: true, replies: [{ author: 'Ben', text: 'Second' }] });
    expect(t.time).toBe(threadAt(s, 0, 0)!.time);
    expect(wb.sheets[0]!.notes.has(cellKey(0, 0))).toBe(false);
    expect(wb.sheets[0]!.notes.get(cellKey(3, 3))!.text).toBe('A plain note');
  });

  it('keeps shapes as DrawingML in the same drawing part as pictures', () => {
    const { s } = setup();
    const kinds: ShapeKind[] = ['rect', 'roundRect', 'ellipse', 'triangle', 'rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'line', 'arrowLine', 'star5', 'diamond', 'pentagon', 'hexagon', 'callout', 'textBox'];
    kinds.forEach((k, i) => addShape(s, defaultShape(k, { row: i * 3, col: 1, dx: 4, dy: 2, width: 120, height: 60 })));
    updateShape(s, s.sheet.shapes![0]!.id, { text: 'Hello\nWorld', font: { size: 14, bold: true, color: '#C00000' }, align: 'right', rotation: 30, line: { color: '#1971C2', width: 2, dash: 'dash' }, alt: 'A box' });
    updateShape(s, s.sheet.shapes![15]!.id, { text: 'Note box' });
    s.sheet.pictures.push({ id: 'p1', data: new Uint8Array([0x89, 0x50, 0x4e, 0x47]), mime: 'image/png', anchor: { row: 0, col: 5, dx: 0, dy: 0, width: 20, height: 20 }, alt: 'Logo' });
    const bytes = writeXlsx(s.wb);
    const files = unzipSync(bytes);
    expect(Object.keys(files).filter((f) => /^xl\/drawings\/drawing\d+\.xml$/.test(f))).toEqual(['xl/drawings/drawing1.xml']);
    const dxml = strFromU8(files['xl/drawings/drawing1.xml']!);
    for (const p of ['rect', 'roundRect', 'ellipse', 'triangle', 'rightArrow', 'line', 'star5', 'wedgeRectCallout', 'hexagon']) expect(dxml).toContain(`prst="${p}"`);
    expect(dxml).toContain('txBox="1"');
    const wb = readXlsx(bytes);
    const sh = wb.sheets[0]!.shapes!;
    expect(sh.map((x) => x.kind)).toEqual(kinds);
    const first = sh[0]!;
    expect(first).toMatchObject({ text: 'Hello\nWorld', align: 'right', rotation: 30, alt: 'A box', fill: '#D6EFE2' });
    expect(first.font).toMatchObject({ size: 14, bold: true, color: '#C00000' });
    expect(first.line).toMatchObject({ color: '#1971C2', width: 2, dash: 'dash' });
    expect(first.anchor).toMatchObject({ row: 0, col: 1, dx: 4, dy: 2, width: 120, height: 60 });
    expect(sh[15]!.text).toBe('Note box');
    expect(sh[15]!.fill).toBeUndefined();
    expect(wb.sheets[0]!.pictures[0]!.alt).toBe('Logo');
  });

  it('keeps checkboxes and ink, and other apps see booleans', () => {
    const { s, sel } = setup();
    sel('A1:A2');
    insertCheckboxes(s);
    toggleCheckboxes(s, [{ row: 1, col: 0 }]);
    addStroke(s, { id: 'k1', tool: 'marker', color: '#e03131', width: 4.4, points: [[1, 2], [3.5, 4]] });
    const bytes = writeXlsx(s.wb);
    expect(strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']!)).toMatch(/<c r="A2"[^>]*t="b"[^>]*><v>1<\/v>/);
    const wb = readXlsx(bytes);
    const sh = wb.sheets[0]!;
    expect(sh.cells.get(cellKey(1, 0))!.v).toBe(true);
    expect(wb.styles[sh.cells.get(cellKey(0, 0))!.s!]!.checkbox).toBe(true);
    expect(sh.ink).toEqual([{ id: 'k1', tool: 'marker', color: '#e03131', width: 4.4, points: [[1, 2], [3.5, 4]] }]);
  });

  it('ignores a malformed objects part', async () => {
    const { applyObjectsPart } = await import('../src/formats/xlsx-objects');
    const wb = newWorkbook();
    applyObjectsPart(wb, '{"sheets":{"0":{"comments":[{"ref":"ZZZZ9","thread":{}}],"ink":[{"points":"x"}],"checkboxes":[5]}}}');
    applyObjectsPart(wb, 'not json');
    expect(wb.sheets[0]!.comments).toBeUndefined();
    expect(wb.sheets[0]!.ink).toBeUndefined();
  });

  it('keeps allow-edit ranges', () => {
    const { s } = setup();
    addEditRange(s, 'Inputs here', { r1: 0, c1: 0, r2: 4, c2: 1 }, 'pw');
    const wb = roundTrip(s);
    expect(wb.sheets[0]!.editRanges).toEqual([{ title: 'Inputs here', range: { r1: 0, c1: 0, r2: 4, c2: 1 }, password: s.sheet.editRanges![0]!.password }]);
  });
});

describe('review checks', () => {
  it('computes contrast ratios', () => {
    expect(contrastRatio('#000000', '#FFFFFF')).toBeCloseTo(21, 1);
    expect(contrastRatio('#FFFFFF', '#FFFFFF')).toBeCloseTo(1, 5);
    expect(contrastRatio('#777777', '#FFFFFF')).toBeCloseTo(4.48, 1);
    expect(contrastRatio('#767676', '#FFFFFF')).toBeGreaterThanOrEqual(4.5);
  });

  it('finds accessibility problems', () => {
    const wb = newWorkbook();
    const s = wb.sheets[0]!;
    s.cells.set(cellKey(0, 0), { v: 'faint', s: internStyle(wb, { font: { color: '#CCCCCC' } }) });
    s.cells.set(cellKey(1, 0), { v: 'fine', s: internStyle(wb, { font: { color: '#000000' }, fill: { color: '#FFFFFF' } }) });
    s.merges.push({ r1: 3, c1: 0, r2: 3, c2: 2 });
    s.pictures.push({ id: 'p', data: new Uint8Array(), mime: 'image/png', anchor: { row: 0, col: 0, dx: 0, dy: 0, width: 1, height: 1 } });
    const kinds = accessibilityIssues(wb).map((i) => i.kind);
    expect(kinds).toContain('Hard-to-read text');
    expect(kinds).toContain('Merged cells');
    expect(kinds).toContain('Missing alt text');
    expect(kinds).toContain('Default sheet name');
    expect(kinds.filter((k) => k === 'Hard-to-read text')).toHaveLength(1);
  });

  it('cleans formatting outside the used range', () => {
    const { s, type } = setup();
    type('B2', '5');
    const st = internStyle(s.wb, { fill: { color: '#FFFF00' } });
    s.sheet.cells.set(cellKey(1, 0), { v: null, s: st });
    s.sheet.cells.set(cellKey(50, 0), { v: null, s: st });
    s.sheet.cells.set(cellKey(0, 20), { v: null, s: st });
    s.sheet.rows.set(100, { s: st });
    const x = excessFormatting(s.sheet);
    expect(x.cells).toHaveLength(2);
    expect(x.rows).toEqual([100]);
    const r = cleanExcessFormatting(s);
    expect(r.removed).toBe(3);
    expect(r.bytes).toBeGreaterThan(0);
    expect(s.sheet.cells.has(cellKey(1, 0))).toBe(true);
    expect(s.sheet.cells.has(cellKey(50, 0))).toBe(false);
    s.undo();
    expect(s.sheet.cells.has(cellKey(50, 0))).toBe(true);
  });

  it('counts statistics', () => {
    const { s, type } = setup();
    type('A1', '1'); type('C4', '=A1*2');
    addComment(s, 0, 0, 'x');
    const st = sheetStats(s.wb, s.sheet);
    expect(st).toMatchObject({ endCell: 'C4', cells: 2, formulas: 1, comments: 1 });
  });

  it('picks words to spell-check', () => {
    expect(wordsIn('Teh NASA report at https://x.io/abc and a@b.com v2 ok').map((w) => w.word)).toEqual(['Teh', 'report', 'at', 'and', 'ok']);
  });
});

describe('allow edit ranges', () => {
  it('unlock cells on a protected sheet', () => {
    const { s } = setup();
    s.sheet.protection = { sheet: true };
    expect(s.isLocked(0, 0)).toBe(true);
    addEditRange(s, 'Open', { r1: 0, c1: 0, r2: 1, c2: 1 });
    expect(s.isLocked(0, 0)).toBe(false);
    expect(s.isLocked(2, 2)).toBe(true);
    addEditRange(s, 'Secret', { r1: 5, c1: 0, r2: 5, c2: 0 }, 'letmein');
    expect(s.isLocked(5, 0)).toBe(true);
    expect(unlockEditRange(s, 1, 'wrong')).toBe(false);
    expect(s.isLocked(5, 0)).toBe(true);
    expect(unlockEditRange(s, 1, 'letmein')).toBe(true);
    expect(s.isLocked(5, 0)).toBe(false);
  });
});

