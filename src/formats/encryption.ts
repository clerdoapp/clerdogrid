// Password-encrypted workbooks: a compound file holding EncryptionInfo and EncryptedPackage streams
// around the normal zip package (MS-OFFCRYPTO open specification, sections 2.3.4.5 to 2.3.4.15).
// Reads agile encryption (any AES key size, SHA-1/256/384/512) and standard encryption (AES-ECB, SHA-1);
// writes agile encryption with AES-256 and SHA-512.
import { cfbGet, isCfb, readCfb, writeCfb } from './cfb';
import { Aes, concat, fromBase64, hashFn, hmac, randomBytes, toBase64, u32le, utf16le } from './crypto';
import { attr, find, kid, numAttr, parseXml } from './xml';

export class EncryptionError extends Error {
  constructor(message: string, readonly code: 'password' | 'unsupported' | 'corrupt') { super(message); }
}

const PASSWORD_ENCRYPTOR = 'http://schemas.microsoft.com/office/2006/keyEncryptor/password';
const BLOCK_VERIFIER_INPUT = new Uint8Array([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]);
const BLOCK_VERIFIER_VALUE = new Uint8Array([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]);
const BLOCK_KEY_VALUE = new Uint8Array([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]);
const BLOCK_HMAC_KEY = new Uint8Array([0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6]);
const BLOCK_HMAC_VALUE = new Uint8Array([0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33]);
const SEGMENT = 4096;
const MAX_SPIN = 10_000_000;

/** True when the bytes are a compound file that holds an encrypted workbook package. */
export const isEncryptedPackage = (bytes: Uint8Array): boolean => {
  if (!isCfb(bytes)) return false;
  try {
    const s = readCfb(bytes);
    return !!cfbGet(s, 'EncryptionInfo') && !!cfbGet(s, 'EncryptedPackage');
  } catch { return false; }
};

/** Pads or truncates to `n` bytes, padding with `fill`. */
const fit = (b: Uint8Array, n: number, fill = 0x36): Uint8Array => {
  if (b.length >= n) return b.slice(0, n);
  const out = new Uint8Array(n).fill(fill);
  out.set(b);
  return out;
};
const padTo = (b: Uint8Array, block: number): Uint8Array => (b.length % block ? fit(b, b.length + block - (b.length % block), 0) : b);

interface KeyParams { salt: Uint8Array; blockSize: number; keyBits: number; hashSize: number; hash: string; cipher: string; chaining: string }

const keyParams = (n: ReturnType<typeof find>): KeyParams => ({
  salt: fromBase64(attr(n, 'saltValue') ?? ''),
  blockSize: numAttr(n, 'blockSize') ?? 16,
  keyBits: numAttr(n, 'keyBits') ?? 256,
  hashSize: numAttr(n, 'hashSize') ?? 64,
  hash: attr(n, 'hashAlgorithm') ?? 'SHA512',
  cipher: attr(n, 'cipherAlgorithm') ?? 'AES',
  chaining: attr(n, 'cipherChaining') ?? 'ChainingModeCBC',
});

/** Agile password key derivation: H(salt+password), spun with a leading iterator, then a block key. */
const agileHash = (password: string, p: KeyParams, spin: number): Uint8Array => {
  const H = hashFn(p.hash)!;
  let h = H(concat(p.salt, utf16le(password)));
  const buf = new Uint8Array(4 + h.length);
  const dv = new DataView(buf.buffer);
  for (let i = 0; i < spin; i++) { dv.setUint32(0, i, true); buf.set(h, 4); h = H(buf); }
  return h;
};
const agileKey = (base: Uint8Array, block: Uint8Array, p: KeyParams): Uint8Array => fit(hashFn(p.hash)!(concat(base, block)), p.keyBits / 8);
const agileIv = (p: KeyParams, block?: Uint8Array): Uint8Array => fit(block ? hashFn(p.hash)!(concat(p.salt, block)) : p.salt, p.blockSize);

const checkParams = (p: KeyParams): void => {
  if (p.cipher !== 'AES') throw new EncryptionError(`Files encrypted with ${p.cipher} can't be opened.`, 'unsupported');
  if (p.chaining !== 'ChainingModeCBC') throw new EncryptionError('This encryption mode isn\'t supported.', 'unsupported');
  if (![128, 192, 256].includes(p.keyBits) || p.blockSize !== 16) throw new EncryptionError('This encryption key size isn\'t supported.', 'unsupported');
  if (!hashFn(p.hash)) throw new EncryptionError(`The ${p.hash} hash isn't supported.`, 'unsupported');
};

const decryptAgile = (xmlText: string, pkg: Uint8Array, password: string): Uint8Array => {
  const doc = parseXml(xmlText);
  const keyData = find(doc, 'keyData');
  const enc = [...(find(doc, 'keyEncryptors')?.children ?? [])].find((k) => attr(k, 'uri') === PASSWORD_ENCRYPTOR);
  const ek = enc ? kid(enc, 'encryptedKey') : undefined;
  if (!keyData || !ek) throw new EncryptionError('This file is protected in a way ClerdoGrid can\'t open.', 'unsupported');
  const kp = keyParams(keyData), pp = keyParams(ek);
  checkParams(kp); checkParams(pp);
  const spin = numAttr(ek, 'spinCount') ?? 100000;
  if (spin < 0 || spin > MAX_SPIN) throw new EncryptionError('This file has invalid encryption settings.', 'corrupt');
  const base = agileHash(password, pp, spin);
  const iv = agileIv(pp);
  const H = hashFn(pp.hash)!;
  const verifierInput = new Aes(agileKey(base, BLOCK_VERIFIER_INPUT, pp)).decryptCbc(padTo(fromBase64(attr(ek, 'encryptedVerifierHashInput') ?? ''), 16), iv).slice(0, pp.salt.length);
  const verifierHash = new Aes(agileKey(base, BLOCK_VERIFIER_VALUE, pp)).decryptCbc(padTo(fromBase64(attr(ek, 'encryptedVerifierHashValue') ?? ''), 16), iv).slice(0, pp.hashSize);
  const expected = H(verifierInput);
  if (expected.length !== verifierHash.length || expected.some((b, i) => b !== verifierHash[i])) throw new EncryptionError("That password isn't right.", 'password');
  const secret = new Aes(agileKey(base, BLOCK_KEY_VALUE, pp)).decryptCbc(padTo(fromBase64(attr(ek, 'encryptedKeyValue') ?? ''), 16), iv).slice(0, kp.keyBits / 8);

  const di = find(doc, 'dataIntegrity');
  if (di) {
    const aes = new Aes(secret);
    const hk = aes.decryptCbc(padTo(fromBase64(attr(di, 'encryptedHmacKey') ?? ''), 16), agileIv(kp, BLOCK_HMAC_KEY)).slice(0, kp.hashSize);
    const hv = aes.decryptCbc(padTo(fromBase64(attr(di, 'encryptedHmacValue') ?? ''), 16), agileIv(kp, BLOCK_HMAC_VALUE)).slice(0, kp.hashSize);
    const actual = hmac(kp.hash, hk, pkg);
    if (actual.some((b, i) => b !== hv[i])) throw new EncryptionError('This file is damaged: its contents don\'t match its integrity check.', 'corrupt');
  }
  return decryptSegments(pkg, secret, kp);
};

const decryptSegments = (pkg: Uint8Array, secret: Uint8Array, kp: KeyParams): Uint8Array => {
  if (pkg.length < 8) throw new EncryptionError('This encrypted file is damaged.', 'corrupt');
  const dv = new DataView(pkg.buffer, pkg.byteOffset, 8);
  const size = dv.getUint32(0, true) + dv.getUint32(4, true) * 0x100000000;
  if (size > pkg.length) throw new EncryptionError('This encrypted file is damaged.', 'corrupt');
  const aes = new Aes(secret);
  const out = new Uint8Array(size);
  const body = pkg.subarray(8);
  for (let i = 0, off = 0; off < size; i++, off += SEGMENT) {
    const chunk = body.subarray(off, Math.min(body.length, off + SEGMENT));
    const plain = aes.decryptCbc(chunk.length % 16 ? chunk.subarray(0, chunk.length - (chunk.length % 16)) : chunk, agileIv(kp, u32le(i)));
    out.set(plain.subarray(0, Math.min(plain.length, size - off)), off);
  }
  return out;
};

// ---------- Standard encryption (read only) ----------

const decryptStandard = (info: Uint8Array, pkg: Uint8Array, password: string): Uint8Array => {
  const dv = new DataView(info.buffer, info.byteOffset, info.byteLength);
  const headerSize = dv.getUint32(8, true);
  const h = 12;
  const algId = dv.getUint32(h + 8, true);
  const keyBits = dv.getUint32(h + 16, true) || 128;
  if (![0x660e, 0x660f, 0x6610].includes(algId)) throw new EncryptionError('This file uses an older encryption ClerdoGrid can\'t open.', 'unsupported');
  let v = h + headerSize;
  const saltSize = dv.getUint32(v, true); v += 4;
  const salt = info.slice(v, v + saltSize); v += saltSize;
  const encVerifier = info.slice(v, v + 16); v += 16;
  const verifierHashSize = dv.getUint32(v, true); v += 4;
  const encVerifierHash = info.slice(v, v + 32);
  // Key derivation from MS-OFFCRYPTO 2.3.4.7.
  const sha1 = hashFn('SHA1')!;
  let hh = sha1(concat(salt, utf16le(password)));
  for (let i = 0; i < 50000; i++) hh = sha1(concat(u32le(i), hh));
  const hf = sha1(concat(hh, u32le(0)));
  const x1 = sha1(new Uint8Array(64).fill(0x36).map((b, i) => b ^ (hf[i] ?? 0)));
  const x2 = sha1(new Uint8Array(64).fill(0x5c).map((b, i) => b ^ (hf[i] ?? 0)));
  const key = concat(x1, x2).slice(0, keyBits / 8);
  const aes = new Aes(key);
  const verifier = aes.decryptEcb(encVerifier);
  const vh = aes.decryptEcb(encVerifierHash).slice(0, verifierHashSize);
  const expected = sha1(verifier);
  if (expected.some((b, i) => b !== vh[i])) throw new EncryptionError("That password isn't right.", 'password');
  const size = new DataView(pkg.buffer, pkg.byteOffset, 8).getUint32(0, true);
  return aes.decryptEcb(pkg.subarray(8)).slice(0, size);
};

/** Decrypts a password-protected workbook and returns the zip package inside. */
export const decryptWorkbook = (bytes: Uint8Array, password: string): Uint8Array => {
  let streams: Map<string, Uint8Array>;
  try { streams = readCfb(bytes); } catch { throw new EncryptionError('This file is damaged.', 'corrupt'); }
  const info = cfbGet(streams, 'EncryptionInfo');
  const pkg = cfbGet(streams, 'EncryptedPackage');
  if (!info || !pkg || info.length < 8) throw new EncryptionError('This file isn\'t an encrypted workbook.', 'corrupt');
  const major = info[0]! | (info[1]! << 8), minor = info[2]! | (info[3]! << 8);
  if (major === 4 && minor === 4) return decryptAgile(new TextDecoder().decode(info.subarray(8)), pkg, password);
  if ((major === 3 || major === 4) && minor === 2) return decryptStandard(info, pkg, password);
  throw new EncryptionError('This file uses an encryption ClerdoGrid can\'t open.', 'unsupported');
};

// ---------- Writing ----------

const lpUnicode = (s: string): Uint8Array => {
  const body = utf16le(s);
  const pad = (4 - (body.length % 4)) % 4;
  return concat(u32le(body.length), body, new Uint8Array(pad));
};
const u16pair = (a: number, b: number): Uint8Array => new Uint8Array([a & 255, a >> 8, b & 255, b >> 8]);

/** The data-spaces streams that mark the package as encrypted (MS-OFFCRYPTO 2.1). */
const dataSpaces = (): [string, Uint8Array][] => {
  const ver = concat(lpUnicode('Microsoft.Container.DataSpaces'), u16pair(1, 0), u16pair(1, 0), u16pair(1, 0));
  const entryBody = concat(u32le(1), u32le(0), lpUnicode('EncryptedPackage'), lpUnicode('StrongEncryptionDataSpace'));
  const map = concat(u32le(8), u32le(1), u32le(entryBody.length + 4), entryBody);
  const def = concat(u32le(8), u32le(1), lpUnicode('StrongEncryptionTransform'));
  const id = lpUnicode('{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}');
  const primary = concat(u32le(8 + id.length), u32le(1), id, lpUnicode('Microsoft.Container.EncryptionTransform'),
    u16pair(1, 0), u16pair(1, 0), u16pair(1, 0), u32le(0), u32le(0), u32le(0), u32le(4));
  return [
    ['\u0006DataSpaces/Version', ver],
    ['\u0006DataSpaces/DataSpaceMap', map],
    ['\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace', def],
    ['\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary', primary],
  ];
};

/** Encrypts a zip package with a password (agile encryption, AES-256, SHA-512). */
export const encryptWorkbook = (zip: Uint8Array, password: string, spinCount = 100000): Uint8Array => {
  const kp: KeyParams = { salt: randomBytes(16), blockSize: 16, keyBits: 256, hashSize: 64, hash: 'SHA512', cipher: 'AES', chaining: 'ChainingModeCBC' };
  const pp: KeyParams = { ...kp, salt: randomBytes(16) };
  const secret = randomBytes(32);
  const aes = new Aes(secret);

  const parts: Uint8Array[] = [];
  const size = new Uint8Array(8);
  const sdv = new DataView(size.buffer);
  sdv.setUint32(0, zip.length >>> 0, true);
  sdv.setUint32(4, Math.floor(zip.length / 0x100000000), true);
  parts.push(size);
  for (let i = 0, off = 0; off < zip.length; i++, off += SEGMENT) {
    parts.push(aes.encryptCbc(padTo(zip.subarray(off, Math.min(zip.length, off + SEGMENT)), 16), agileIv(kp, u32le(i))));
  }
  const pkg = concat(...parts);

  const hmacKey = randomBytes(64);
  const encHmacKey = aes.encryptCbc(hmacKey, agileIv(kp, BLOCK_HMAC_KEY));
  const encHmacValue = aes.encryptCbc(hmac('SHA512', hmacKey, pkg), agileIv(kp, BLOCK_HMAC_VALUE));

  const base = agileHash(password, pp, spinCount);
  const iv = agileIv(pp);
  const verifierInput = randomBytes(16);
  const encVerifierInput = new Aes(agileKey(base, BLOCK_VERIFIER_INPUT, pp)).encryptCbc(verifierInput, iv);
  const encVerifierHash = new Aes(agileKey(base, BLOCK_VERIFIER_VALUE, pp)).encryptCbc(padTo(hashFn('SHA512')!(verifierInput), 16), iv);
  const encKey = new Aes(agileKey(base, BLOCK_KEY_VALUE, pp)).encryptCbc(secret, iv);

  const common = 'blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512"';
  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    + '<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">'
    + `<keyData saltSize="16" ${common} saltValue="${toBase64(kp.salt)}"/>`
    + `<dataIntegrity encryptedHmacKey="${toBase64(encHmacKey)}" encryptedHmacValue="${toBase64(encHmacValue)}"/>`
    + `<keyEncryptors><keyEncryptor uri="${PASSWORD_ENCRYPTOR}"><p:encryptedKey spinCount="${spinCount}" saltSize="16" ${common} saltValue="${toBase64(pp.salt)}"`
    + ` encryptedVerifierHashInput="${toBase64(encVerifierInput)}" encryptedVerifierHashValue="${toBase64(encVerifierHash)}" encryptedKeyValue="${toBase64(encKey)}"/></keyEncryptor></keyEncryptors></encryption>`;
  const info = concat(u16pair(4, 4), u32le(0x40), new TextEncoder().encode(xml));

  return writeCfb(new Map<string, Uint8Array>([...dataSpaces(), ['EncryptionInfo', info], ['EncryptedPackage', pkg]]));
};
