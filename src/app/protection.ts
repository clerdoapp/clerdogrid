/**
 * The sheet-protection password hash defined by the file format (ECMA-376), so protected sheets
 * unlock the same way in other apps. It is a short hash: protection prevents accidental edits,
 * it is not security.
 */
export const protectionHash = (password: string): string => {
  let h = 0;
  for (let i = password.length - 1; i >= 0; i--) {
    h = ((h >> 14) & 0x01) | ((h << 1) & 0x7fff);
    h ^= password.charCodeAt(i);
  }
  h = ((h >> 14) & 0x01) | ((h << 1) & 0x7fff);
  h ^= password.length;
  h ^= 0xce4b;
  return h.toString(16).toUpperCase().padStart(4, '0');
};

import { concat, fromBase64, hashFn, randomBytes, toBase64, utf16le } from '../formats/crypto';
import type { PasswordHash } from '../model/workbook';
import { N, Tr } from '../ui/i18n';

/**
 * The modern protection hash (ECMA-376 Part 1, algorithmName/hashValue/saltValue/spinCount):
 * H(salt + password as UTF-16LE), then spinCount rounds of H(previous + 4-byte little-endian round number).
 */
export const modernHash = (password: string, saltValue: string, spinCount: number, algorithmName = 'SHA-512'): string | null => {
  const H = hashFn(algorithmName);
  if (!H || spinCount < 0 || spinCount > 10_000_000) return null;
  let h = H(concat(fromBase64(saltValue), utf16le(password)));
  const buf = new Uint8Array(h.length + 4);
  const dv = new DataView(buf.buffer);
  for (let i = 0; i < spinCount; i++) { buf.set(h, 0); dv.setUint32(h.length, i, true); h = H(buf); }
  return toBase64(h);
};

export const makePasswordHash = (password: string, spinCount = 100000): PasswordHash => {
  const saltValue = toBase64(randomBytes(16));
  return { algorithmName: 'SHA-512', saltValue, spinCount, hashValue: modernHash(password, saltValue, spinCount)! };
};

/** Checks a password against a legacy hash, a modern hash, or both. No hash at all means no password. */
export const checkPassword = (password: string, p: { password?: string; hash?: PasswordHash }): boolean => {
  if (p.hash) {
    const h = modernHash(password, p.hash.saltValue, p.hash.spinCount, p.hash.algorithmName);
    if (h !== null) return h === p.hash.hashValue;
  }
  if (p.password) return protectionHash(password).toUpperCase() === p.password.toUpperCase();
  return true;
};

export const hasPassword = (p: { password?: string; hash?: PasswordHash } | undefined): boolean => !!(p?.password || p?.hash);

const STRUCTURE_MESSAGE = N('The workbook structure is protected. Unprotect the workbook to change sheets.');
/** True (and tells the person) when workbook structure protection blocks sheet changes. */
export const structureLocked = (store: { wb: { protection?: { structure: boolean } }; notify: (t: string) => void }): boolean => {
  if (!store.wb.protection?.structure) return false;
  store.notify(Tr(STRUCTURE_MESSAGE));
  return true;
};
