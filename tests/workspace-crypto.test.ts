import { describe, expect, it } from 'vitest';
import { createHash, createHmac } from 'node:crypto';
import { Aes, hmac, sha1, sha256, sha384, sha512 } from '../src/formats/crypto';

const hex = (b: Uint8Array) => Buffer.from(b).toString('hex');
const unhex = (s: string) => new Uint8Array(Buffer.from(s, 'hex'));

describe('hashes', () => {
  const inputs = ['', 'abc', 'a'.repeat(111), 'b'.repeat(112), 'c'.repeat(1000), 'x'.repeat(55), 'y'.repeat(64)];
  for (const s of inputs) {
    it(`hash ${s.length}`, () => {
      const b = new TextEncoder().encode(s);
      expect(hex(sha512(b))).toBe(createHash('sha512').update(b).digest('hex'));
      expect(hex(sha384(b))).toBe(createHash('sha384').update(b).digest('hex'));
      expect(hex(sha256(b))).toBe(createHash('sha256').update(b).digest('hex'));
      expect(hex(sha1(b))).toBe(createHash('sha1').update(b).digest('hex'));
    });
  }
  it('hmac', () => {
    const k = unhex('0b'.repeat(80)), d = new TextEncoder().encode('Hi There');
    expect(hex(hmac('SHA512', k, d))).toBe(createHmac('sha512', k).update(d).digest('hex'));
    expect(hex(hmac('SHA1', k.subarray(0, 20), d))).toBe(createHmac('sha1', k.subarray(0, 20)).update(d).digest('hex'));
  });
});

describe('AES', () => {
  // FIPS 197 appendix C examples.
  const pt = unhex('00112233445566778899aabbccddeeff');
  const cases: [string, string][] = [
    ['000102030405060708090a0b0c0d0e0f', '69c4e0d86a7b0430d8cdb78070b4c55a'],
    ['000102030405060708090a0b0c0d0e0f1011121314151617', 'dda97ca4864cdfe06eaf70a0ec0d7191'],
    ['000102030405060708090a0b0c0d0e0f101112131415161718191a1b1c1d1e1f', '8ea2b7ca516745bfeafc49904b496089'],
  ];
  for (const [k, ct] of cases) it(`key ${k.length * 4} bits`, () => {
    const a = new Aes(unhex(k));
    expect(hex(a.encryptEcb(pt))).toBe(ct);
    expect(hex(a.decryptEcb(unhex(ct)))).toBe(hex(pt));
  });
  it('CBC round trip', () => {
    const a = new Aes(unhex('2b7e151628aed2a6abf7158809cf4f3c'));
    const iv = unhex('000102030405060708090a0b0c0d0e0f');
    const data = unhex('6bc1bee22e409f96e93d7e117393172aae2d8a571e03ac9c9eb76fac45af8e51');
    const c = a.encryptCbc(data, iv);
    expect(hex(c)).toBe('7649abac8119b246cee98e9b12e9197d5086cb9b507219ee95db113a917678b2');
    expect(hex(a.decryptCbc(c, iv))).toBe(hex(data));
  });
});
