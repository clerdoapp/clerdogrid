import { describe, expect, it } from 'vitest';
import { cellKey, newWorkbook, type SparklineGroup } from '../src/model/workbook';
import { readXlsx, writeXlsx } from '../src/formats/xlsx';
import { GridStore } from '../src/app/store';
import { buildSparklines, defaultGroup, groupScales, sparkGeometry, sparkValues } from '../src/features/sparklines/model';
import { sparklinesExtXml } from '../src/formats/xlsx-sparklines';
import { unzipSync, strFromU8 } from 'fflate';

describe('placing sparklines', () => {
  it('one per data row along a column, one per data column along a row', () => {
    const rows = buildSparklines('S', { range: { r1: 1, c1: 1, r2: 3, c2: 6 } }, { r1: 1, c1: 7, r2: 3, c2: 7 });
    expect(rows).toEqual([{ ref: 'S!$B$2:$G$2', row: 1, col: 7 }, { ref: 'S!$B$3:$G$3', row: 2, col: 7 }, { ref: 'S!$B$4:$G$4', row: 3, col: 7 }]);
    const cols = buildSparklines('S', { sheet: 'Data', range: { r1: 0, c1: 0, r2: 9, c2: 2 } }, { r1: 10, c1: 0, r2: 10, c2: 2 });
    expect(cols).toEqual([{ ref: 'Data!$A$1:$A$10', row: 10, col: 0 }, { ref: 'Data!$B$1:$B$10', row: 10, col: 1 }, { ref: 'Data!$C$1:$C$10', row: 10, col: 2 }]);
    expect(buildSparklines('S', { range: { r1: 0, c1: 0, r2: 0, c2: 5 } }, { r1: 0, c1: 6, r2: 0, c2: 6 })).toEqual([{ ref: 'S!$A$1:$F$1', row: 0, col: 6 }]);
    expect(typeof buildSparklines('S', { range: { r1: 0, c1: 0, r2: 3, c2: 3 } }, { r1: 0, c1: 5, r2: 2, c2: 5 })).toBe('string');
    expect(typeof buildSparklines('S', { range: { r1: 0, c1: 0, r2: 3, c2: 3 } }, { r1: 0, c1: 5, r2: 1, c2: 6 })).toBe('string');
  });
});

describe('sparkline values and shapes', () => {
  const wb = newWorkbook();
  const s = wb.sheets[0]!;
  [[1, 5, null, -2, 4], [10, 20, 30, 40, 50]].forEach((r, ri) => r.forEach((v, ci) => { if (v !== null) s.cells.set(cellKey(ri, ci), { v }); }));
  s.cols.set(4, { hidden: true });
  const store = new GridStore(wb);
  const g = defaultGroup('line', [{ ref: 'Sheet1!$A$1:$E$1', row: 0, col: 6 }, { ref: 'Sheet1!$A$2:$E$2', row: 1, col: 6 }]);
  it('reads values, skipping hidden cells unless asked', () => {
    expect(sparkValues(store, 0, g, g.sparklines[0]!)).toEqual([1, 5, null, -2]);
    expect(sparkValues(store, 0, { ...g, displayHidden: true }, g.sparklines[0]!)).toEqual([1, 5, null, -2, 4]);
    expect(sparkValues(store, 0, { ...g, displayEmptyCellsAs: 'zero' }, g.sparklines[0]!)).toEqual([1, 5, 0, -2]);
  });
  it('scales each sparkline, the group, or to fixed bounds', () => {
    const vals = [[1, 5, -2], [10, 50]];
    expect(groupScales(g, vals)).toEqual([{ min: -2, max: 5 }, { min: 10, max: 50 }]);
    expect(groupScales({ ...g, minAxis: 'group', maxAxis: 'group' }, vals)).toEqual([{ min: -2, max: 50 }, { min: -2, max: 50 }]);
    expect(groupScales({ ...g, minAxis: 'custom', manualMin: 0, maxAxis: 'custom', manualMax: 100 }, vals)[1]).toEqual({ min: 0, max: 100 });
  });
  const box = { x: 0, y: 0, w: 100, h: 20 };
  it('lines break at gaps unless joined, with marked points', () => {
    const geo = sparkGeometry({ ...g, high: true, low: true }, [1, 5, null, -2], box, { min: -2, max: 5 });
    expect(geo.lines.map((l) => l.length)).toEqual([2, 1]);
    expect(geo.marks.map((m) => m.color)).toEqual([g.highColor, g.lowColor]);
    expect(sparkGeometry({ ...g, displayEmptyCellsAs: 'span' }, [1, 5, null, -2], box, { min: -2, max: 5 }).lines).toHaveLength(1);
    const rtl = sparkGeometry({ ...g, rightToLeft: true }, [1, 2], box, { min: 1, max: 2 });
    expect(rtl.lines[0]![0]![0]).toBeGreaterThan(rtl.lines[0]![1]![0]);
  });
  it('columns and win/loss bars', () => {
    const col = sparkGeometry({ ...defaultGroup('column', []), negative: true }, [3, -1, 2], box, { min: -1, max: 3 });
    expect(col.bars).toHaveLength(3);
    expect(col.bars[1]!.color).toBe(defaultGroup('column', []).negativeColor);
    const wl = sparkGeometry(defaultGroup('winLoss', []), [3, -1, 0, 2], box, { min: -1, max: 3 });
    expect(wl.bars).toHaveLength(3);
    expect(wl.bars[0]!.y1 - wl.bars[0]!.y0).toBeCloseTo(wl.bars[1]!.y1 - wl.bars[1]!.y0, 6);
  });
  it('spaces points by date when there is a date range', () => {
    const geo = sparkGeometry(g, [1, 2, 3], box, { min: 1, max: 3 }, [0, 1, 10]);
    const xs = geo.lines[0]!.map((p) => p[0]);
    expect(xs[1]! - xs[0]!).toBeLessThan((xs[2]! - xs[1]!) / 5);
  });
});

describe('sparklines in .xlsx', () => {
  it('save and reopen with every option', () => {
    const wb = newWorkbook();
    const s = wb.sheets[0]!;
    for (let c = 0; c < 6; c++) s.cells.set(cellKey(0, c), { v: c * 2 - 3 });
    const groups: SparklineGroup[] = [
      { ...defaultGroup('line', [{ ref: 'Sheet1!$A$1:$F$1', row: 2, col: 0 }, { ref: 'Sheet1!$A$1:$F$1', row: 3, col: 0 }]), markers: true, high: true, low: true, first: true, last: true, lineWeight: 1.5, rightToLeft: true, displayEmptyCellsAs: 'span', displayHidden: true, dateRange: 'Sheet1!$A$5:$F$5' },
      { ...defaultGroup('column', [{ ref: 'Sheet1!$A$1:$F$1', row: 2, col: 1 }]), negative: true, showAxis: true, minAxis: 'custom', manualMin: -5, maxAxis: 'group', displayEmptyCellsAs: 'zero' },
      { ...defaultGroup('winLoss', [{ ref: 'Sheet1!$A$1:$F$1', row: 2, col: 2 }]), maxAxis: 'custom', manualMax: 10, minAxis: 'group' },
    ];
    s.sparklines = groups;
    const bytes = writeXlsx(wb);
    const xml = strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']!);
    expect(xml).toContain('<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"');
    expect(xml).toContain('<x14:sparklineGroup');
    expect(xml).toContain('type="stacked"');
    expect(xml).toContain('<xm:sqref>A3</xm:sqref>');
    expect(xml.lastIndexOf('<extLst>')).toBeGreaterThan(xml.lastIndexOf('</sheetData>'));
    expect(readXlsx(bytes).sheets[0]!.sparklines).toEqual(groups);
  });
  it('writes nothing for a sheet without sparklines', () => {
    expect(sparklinesExtXml(newWorkbook().sheets[0]!)).toBe('');
  });
});
