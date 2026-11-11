import { describe, expect, it } from 'vitest';
import { cfbGet, isCfb, readCfb, writeCfb } from '../src/formats/cfb';

const filled = (n: number, seed: number) => {
  const a = new Uint8Array(n);
  for (let i = 0; i < n; i++) a[i] = (i * 31 + seed) & 0xff;
  return a;
};

describe('compound file', () => {
  it('writes and reads back small, large and nested streams', () => {
    const streams = new Map<string, Uint8Array>([
      ['Workbook', filled(10_000, 1)],
      ['\u0005SummaryInformation', filled(200, 2)],
      ['EncryptionInfo', filled(4095, 3)],
      ['\u0006DataSpaces/Version', filled(76, 4)],
      ['\u0006DataSpaces/DataSpaceMap', filled(4096, 5)],
      ['Empty', new Uint8Array(0)],
    ]);
    const bytes = writeCfb(streams);
    expect(isCfb(bytes)).toBe(true);
    const back = readCfb(bytes);
    for (const [k, v] of streams) expect(back.get(k), k).toEqual(v);
    expect(cfbGet(back, 'WORKBOOK')).toEqual(streams.get('Workbook'));
  });

  it('handles many streams and a big file needing extra FAT sectors', () => {
    const streams = new Map<string, Uint8Array>();
    for (let i = 0; i < 300; i++) streams.set(`S${i}`, filled(i * 37 + 1, i));
    streams.set('Big', filled(8 * 1024 * 1024, 9));
    const back = readCfb(writeCfb(streams));
    expect(back.size).toBe(301);
    for (const [k, v] of streams) expect(back.get(k)!.length, k).toBe(v.length);
    const a = back.get('Big')!, b = streams.get('Big')!;
    let same = a.length === b.length;
    for (let i = 0; same && i < a.length; i++) same = a[i] === b[i];
    expect(same).toBe(true);
    expect(back.get('S77')).toEqual(streams.get('S77'));
  });

  it('rejects files that are not compound files', () => {
    expect(() => readCfb(new Uint8Array(1024))).toThrow();
  });
});
