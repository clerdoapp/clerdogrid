import { describe, expect, it } from 'vitest';
import { cellKey } from '../src/model/workbook';
import { writeCfb } from '../src/formats/cfb';
import { readXls, PasswordRequired } from '../src/formats/biff8-read';
import { xorArray, xorEncode, xorKey, xorVerifier } from '../src/formats/biff8-crypto';
import { isOldBiff, readOldBiff } from '../src/formats/biff-old-read';
import { workbookFromBytes } from '../src/app/files';

// Record streams built byte by byte from the format descriptions.
const rec = (type: number, ...parts: (number[] | Uint8Array)[]): number[] => {
  const body = parts.flatMap((p) => [...p]);
  return [type & 255, type >> 8, body.length & 255, body.length >> 8, ...body];
};
const u16 = (n: number) => [n & 255, (n >> 8) & 255];
const u32 = (n: number) => [n & 255, (n >> 8) & 255, (n >> 16) & 255, (n >>> 24) & 255];
const f64 = (n: number) => [...new Uint8Array(new Float64Array([n]).buffer)];
const bytes = (s: string) => [...s].map((c) => c.charCodeAt(0) & 255);

/** A two-sheet Excel 95 workbook: code-page text, number formats, formulas with 3D references, a note. */
const biff5 = (encrypt?: string): Uint8Array => {
  const sheetRecs = (cells: number[][]) => [rec(0x0809, u16(0x0500), u16(0x10), u16(0), u16(0)), ...cells, rec(0x000a)];
  const globals = (pos1: number, pos2: number) => [
    rec(0x0809, u16(0x0500), u16(5), u16(0), u16(0)),
    ...(encrypt ? [rec(0x002f, u16(xorKey(encrypt)), u16(xorVerifier(encrypt)))] : []),
    rec(0x0042, u16(1252)),
    rec(0x0031, u16(200), u16(0), u16(0x7fff), u16(400), u16(0), [0, 0, 0, 0], [5], bytes('Arial')),
    rec(0x041e, u16(164), [5], bytes('0.000')),
    rec(0x00e0, u16(0), u16(0), u16(0xfff5), [0x20, 0], u32(0x20c0), u32(0)),
    rec(0x00e0, u16(0), u16(164), u16(0x0001), [0x22, 0], u32(0x20c0), u32(0)),
    rec(0x0085, u32(pos1), [0, 0], [5], bytes('Daten')),
    rec(0x0085, u32(pos2), [0, 0], [4], bytes('Zwei')),
    rec(0x000a),
  ];
  const s1 = sheetRecs([
    rec(0x0204, u16(0), u16(0), u16(0), u16(5), bytes('Caf\xe9!')),
    rec(0x0203, u16(1), u16(0), u16(1), f64(1.5)),
    // =A2*2 : tRefV row 1 (relative row and column), col 0 ; tInt 2 ; tMul
    rec(0x0006, u16(1), u16(1), u16(0), f64(3), u16(0), u32(0), u16(8), [0x44], u16(0xc001), [0], [0x1e], u16(2), [0x05]),
    // =Zwei!A1&"x" : tRef3dV ixals -1, 8 reserved, itab 1..1, row 0 col 0 ; tStr "x" ; tConcat
    rec(0x0006, u16(2), u16(1), u16(0), [0, 0, 0, 0, 0, 0, 0xff, 0xff], u16(0), u32(0), u16(23), [0x5a], u16(0xffff), [0, 0, 0, 0, 0, 0, 0, 0], u16(1), u16(1), u16(0xc000), [0], [0x17, 1], bytes('x'), [0x08]),
    rec(0x0207, u16(3), bytes('hix')),
    rec(0x001c, u16(0), u16(0), u16(4), bytes('Note')),
  ]);
  const s2 = sheetRecs([rec(0x0204, u16(0), u16(0), u16(0), u16(2), bytes('hi'))]);
  const lenOf = (a: number[][]) => a.reduce((n, r) => n + r.length, 0);
  const g0 = globals(0, 0);
  const pos1 = lenOf(g0), pos2 = pos1 + lenOf(s1);
  const stream = new Uint8Array([...globals(pos1, pos2).flat(), ...s1.flat(), ...s2.flat()]);
  if (encrypt) {
    const arr = xorArray(encrypt, xorKey(encrypt), xorVerifier(encrypt))!;
    const dv = new DataView(stream.buffer);
    let o = 0;
    let after = false;
    while (o + 4 <= stream.length) {
      const t = dv.getUint16(o, true), n = dv.getUint16(o + 2, true);
      if (after && t !== 0x0809) xorEncode(stream, o + 4 + (t === 0x0085 ? 4 : 0), o + 4 + n, n, arr);
      if (t === 0x002f) after = true;
      o += 4 + n;
    }
  }
  return writeCfb(new Map([['Book', stream]]));
};

describe('Excel 95 (BIFF5) workbooks', () => {
  it('reads text in the code page, number formats, formulas and notes', () => {
    const wb = readXls(biff5());
    expect(wb.sheets.map((s) => s.name)).toEqual(['Daten', 'Zwei']);
    const s = wb.sheets[0]!;
    expect(s.cells.get(0)!.v).toBe('Café!');
    expect(s.cells.get(cellKey(1, 0))!.v).toBe(1.5);
    expect(wb.styles[s.cells.get(cellKey(1, 0))!.s!]!).toMatchObject({ numFmt: '0.000', align: { h: 'center' } });
    expect(s.cells.get(cellKey(1, 1))).toMatchObject({ f: 'A2*2', v: 3 });
    expect(s.cells.get(cellKey(2, 1))).toMatchObject({ f: 'Zwei!A1&"x"', v: 'hix' });
    expect(s.notes.get(0)).toEqual({ text: 'Note' });
    expect(wb.sheets[1]!.cells.get(0)!.v).toBe('hi');
  });
  it('opens XOR-obfuscated files with their password', () => {
    const bytesEnc = biff5('secret');
    expect(() => readXls(bytesEnc)).toThrow(PasswordRequired);
    expect(() => readXls(bytesEnc, { password: 'nope' })).toThrow(/not right/);
    const wb = readXls(bytesEnc, { password: 'secret' });
    expect(wb.sheets[0]!.cells.get(0)!.v).toBe('Café!');
    expect(wb.sheets[0]!.cells.get(cellKey(1, 1))!.f).toBe('A2*2');
  });
  it('derives the XOR key and verifier per the specification', () => {
    expect(xorKey('abc').toString(16)).toBe('514a');
    expect(xorVerifier('abc').toString(16)).toBe('cc1a');
  });
});

describe('BIFF2-4 worksheets', () => {
  it('reads values, cached results, formats and widths of a BIFF2 sheet', () => {
    const stream = new Uint8Array([
      ...rec(0x0009, u16(2), u16(0x10)),
      ...rec(0x001e, [7], bytes('General')), ...rec(0x001e, [4], bytes('0.00')),
      ...rec(0x0024, [0], [0], u16(20 * 256)),
      ...rec(0x0004, u16(0), u16(0), [0, 0, 0], [5], bytes('Hello')),
      ...rec(0x0003, u16(1), u16(0), [0, 1, 0], f64(2.5)),
      ...rec(0x0002, u16(2), u16(0), [0, 0, 0], u16(7)),
      ...rec(0x0005, u16(3), u16(0), [0, 0, 0], [1], [0]),
      ...rec(0x0006, u16(4), u16(0), [0, 0, 0], f64(10), [0], [3], [0x1e, 10, 0]),
      ...rec(0x000a),
    ]);
    expect(isOldBiff(stream)).toBe(true);
    const wb = workbookFromBytes('old.xls', stream);
    const s = wb.sheets[0]!;
    expect(s.cells.get(0)!.v).toBe('Hello');
    expect(s.cells.get(cellKey(1, 0))!.v).toBe(2.5);
    expect(wb.styles[s.cells.get(cellKey(1, 0))!.s!]!.numFmt).toBe('0.00');
    expect(s.cells.get(cellKey(2, 0))!.v).toBe(7);
    expect(s.cells.get(cellKey(3, 0))!.v).toBe(true);
    expect(s.cells.get(cellKey(4, 0))!.v).toBe(10);
    expect(s.cols.get(0)!.width).toBeCloseTo(19.29, 1);
  });
  it('reads BIFF4 workbooks with several sheets', () => {
    const sheet = (v: number) => [...rec(0x0409, u16(0), u16(0x10), u16(0)), ...rec(0x0203, u16(0), u16(0), u16(0), f64(v)), ...rec(0x000a)];
    const stream = new Uint8Array([
      ...rec(0x0409, u16(0), u16(0x100), u16(0)),
      ...rec(0x008f, u32(0), u16(0), [2], bytes('S1')), ...rec(0x008f, u32(0), u16(0), [2], bytes('S2')),
      ...rec(0x000a), ...sheet(1), ...sheet(2),
    ]);
    const wb = readOldBiff(stream);
    expect(wb.sheets.map((s) => [s.name, s.cells.get(0)!.v])).toEqual([['S1', 1], ['S2', 2]]);
  });
});
