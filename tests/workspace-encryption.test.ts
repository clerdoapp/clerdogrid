import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { decryptWorkbook, encryptWorkbook, EncryptionError, isEncryptedPackage } from '../src/formats/encryption';
import { readCfb, writeCfb } from '../src/formats/cfb';
import { readXlsx, writeXlsx } from '../src/formats/xlsx';
import { newWorkbook, cellKey } from '../src/model/workbook';

const fx = (n: string) => new Uint8Array(readFileSync(new URL(`./workspace-fixtures/${n}`, import.meta.url)));

describe('encrypted workbooks', () => {
  it('opens a file encrypted by another tool', () => {
    const enc = fx('enc_py.xlsx');
    expect(isEncryptedPackage(enc)).toBe(true);
    const zip = decryptWorkbook(enc, 'pa$$wörd');
    expect(zip).toEqual(fx('plain.xlsx'));
    const wb = readXlsx(zip);
    expect(wb.sheets[0]!.cells.get(cellKey(0, 0))?.v).toBe('secret');
  });

  it('rejects a wrong password', () => {
    expect(() => decryptWorkbook(fx('enc_py.xlsx'), 'nope')).toThrow(EncryptionError);
    try { decryptWorkbook(fx('enc_py.xlsx'), 'nope'); } catch (e) { expect((e as EncryptionError).code).toBe('password'); }
  });

  it('saves with a password and reopens', () => {
    const wb = newWorkbook();
    wb.sheets[0]!.cells.set(cellKey(3, 2), { v: 'hidden value' });
    const zip = writeXlsx(wb);
    const enc = encryptWorkbook(zip, 'Correct horse', 1000);
    expect(isEncryptedPackage(enc)).toBe(true);
    const streams = readCfb(enc);
    expect([...streams.keys()]).toContain('\u0006DataSpaces/DataSpaceMap');
    const back = decryptWorkbook(enc, 'Correct horse');
    expect(back).toEqual(zip);
    expect(readXlsx(back).sheets[0]!.cells.get(cellKey(3, 2))?.v).toBe('hidden value');
    expect(() => decryptWorkbook(enc, 'correct horse')).toThrow(/password/);
  });

  it('detects tampering through the integrity check', () => {
    const enc = encryptWorkbook(new Uint8Array(5000).fill(7), 'x', 10);
    const streams = readCfb(enc);
    const pkg = streams.get('EncryptedPackage')!;
    pkg[100]! ^= 1;
    expect(() => decryptWorkbook(writeCfb(streams), 'x')).toThrow(/damaged/);
  });
});
