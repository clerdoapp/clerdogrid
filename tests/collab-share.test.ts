import { beforeAll, describe, expect, it } from 'vitest';
import { GridStore } from '../src/app/store';
import { registerAllFunctions } from '../src/engine/functions';
import { workbookFromBytes } from '../src/app/files';
import { cellKey, parseAddr } from '../src/model/workbook';
import { makeShareLink, openShareHash, parseShareHash } from '../src/features/collab/share';
import { capacity, encodeQr, formatBits, qrPath } from '../src/features/collab/qr';
import { b64urlDecode, b64urlEncode, packText, unpackText } from '../src/features/collab/b64';
import { checkTurn, decodeCode, encodeCode, iceServers, readCode } from '../src/features/collab/rtc';

beforeAll(() => registerAllFunctions());

const type = (s: GridStore, a: string, text: string) => { const p = parseAddr(a)!; s.select(p); s.beginEdit(text); s.commitEdit(); };

describe('share links', () => {
  it('saves a workbook into a link and reopens it through the file reader', () => {
    const s = new GridStore();
    type(s, 'A1', 'Price'); type(s, 'A2', '12.5'); type(s, 'B2', '=A2*2');
    s.selectRange({ r1: 0, c1: 0, r2: 0, c2: 0 }); s.applyStyle({ font: { bold: true } });
    s.file.name = 'Prices Q4';
    const link = makeShareLink(s, true);
    if (!('url' in link)) throw new Error('expected a link');
    const hash = link.url.slice(link.url.indexOf('#'));
    expect(hash).toMatch(/^#wb=[A-Za-z0-9_-]+&ro=1&n=Prices%20Q4$/);
    const shared = parseShareHash(hash)!;
    expect(shared.readOnly).toBe(true);
    expect(shared.name).toBe('Prices Q4');
    const wb = workbookFromBytes('x.xlsx', shared.bytes);
    expect(wb.sheets[0]!.cells.get(cellKey(1, 1))?.f).toBe('A2*2');
    expect(wb.styles[wb.sheets[0]!.cells.get(cellKey(0, 0))!.s!]!.font?.bold).toBe(true);

    const t = new GridStore();
    expect(openShareHash(t, hash)).toBe(true);
    expect(t.readOnly).toBe(true);
    expect(t.value(1, 1)).toBe(25);
    expect(t.file.name).toBe('Prices Q4');
    t.select({ row: 5, col: 5 }); t.beginEdit('x');
    expect(t.editing).toBe(null);
  });

  it('refuses broken or hostile links without throwing', () => {
    expect(parseShareHash('#something=1')).toBe(null);
    expect(() => parseShareHash('#wb=@@@')).toThrow();
    expect(() => parseShareHash('#wb=AAAA')).toThrow();
    const t = new GridStore();
    // Valid deflate of bytes that aren't a workbook: the reader rejects it, the app keeps going.
    expect(openShareHash(t, `#wb=${packText('not a zip')}`)).toBe(false);
    expect(t.toast).toBeTruthy();
  });

  it('caps inflation so a tiny link cannot expand without limit', () => {
    const big = packText('0'.repeat(5_000_000));
    expect(big.length).toBeLessThan(20_000);
    expect(() => unpackText(big, 1_000_000)).toThrow();
  });
});

describe('codes', () => {
  it('round-trips offer and answer codes and tells them apart', () => {
    const sdp = 'v=0\r\no=- 1 2 IN IP4 127.0.0.1\r\ns=-\r\nt=0 0\r\nm=application 9 UDP/DTLS/SCTP webrtc-datachannel\r\na=fingerprint:sha-256 AA:BB\r\n';
    const offer = encodeCode('offer', sdp);
    expect(offer).toMatch(/^CGO1\.[A-Za-z0-9_-]+$/);
    expect(decodeCode(offer, 'offer')).toBe(sdp);
    expect(decodeCode(`https://x.example/app/#join=${offer}`, 'offer')).toBe(sdp);
    expect(() => decodeCode(offer, 'answer')).toThrow(/invite code/);
    expect(() => decodeCode(offer.slice(0, 20), 'offer')).toThrow();
    expect(() => decodeCode(encodeCode('offer', 'hello'), 'offer')).toThrow();
  });
  it('carries the host\'s TURN relay in an invite when asked, and checks it', () => {
    const sdp = 'v=0\r\ns=-\r\nm=application 9 UDP/DTLS/SCTP webrtc-datachannel\r\n';
    const turn = { urls: 'turn:relay.example.com:3478', username: 'team', credential: 's3cret' };
    const code = encodeCode('offer', sdp, turn);
    expect(code).toMatch(/^CGO2\./);
    expect(readCode(code, 'offer')).toEqual({ sdp, turn });
    expect(readCode(encodeCode('offer', sdp), 'offer').turn).toBe(null);
    expect(encodeCode('answer', sdp, turn)).toMatch(/^CGA1\./); // replies never carry it
    expect(() => readCode(code, 'answer')).toThrow(/invite code/);
    expect(checkTurn({ urls: 'http://evil', username: '', credential: '' })).toBe(null);
    expect(checkTurn({ urls: 'turns:relay.example.com:5349?transport=tcp', username: 'a', credential: 'b' })).not.toBe(null);
    // A guest uses the relay from the code alongside its own servers.
    expect(iceServers(turn).some((s) => s.urls === turn.urls && s.username === 'team')).toBe(true);
  });
  it('base64url round-trips every byte value', () => {
    const b = new Uint8Array(256).map((_, i) => i);
    for (let n = 0; n < 6; n++) expect(b64urlDecode(b64urlEncode(b.subarray(n)))).toEqual(b.subarray(n));
  });
});

describe('QR codes (ISO/IEC 18004, byte mode, level M)', () => {
  it('has the standard capacities and format words', () => {
    expect(capacity(1)).toBe(14);
    expect(capacity(10)).toBe(213);
    expect(capacity(40)).toBe(2331);
    expect(formatBits(0)).toBe(0b101010000010010);
    expect(formatBits(5)).toBe(0b100000011001110);
  });
  it('picks the smallest version and draws finder patterns, timing and the dark module', () => {
    const q = encodeQr('HELLO WORLD');
    expect(q.version).toBe(1);
    expect(q.size).toBe(21);
    const finder = [[1, 1, 1, 1, 1, 1, 1], [1, 0, 0, 0, 0, 0, 1], [1, 0, 1, 1, 1, 0, 1], [1, 0, 1, 1, 1, 0, 1], [1, 0, 1, 1, 1, 0, 1], [1, 0, 0, 0, 0, 0, 1], [1, 1, 1, 1, 1, 1, 1]];
    for (const [ox, oy] of [[0, 0], [14, 0], [0, 14]]) for (let y = 0; y < 7; y++) for (let x = 0; x < 7; x++) expect(q.dark(ox! + x, oy! + y)).toBe(finder[y]![x] === 1);
    for (let i = 8; i < 13; i++) { expect(q.dark(i, 6)).toBe(i % 2 === 0); expect(q.dark(6, i)).toBe(i % 2 === 0); }
    expect(q.dark(8, q.size - 8)).toBe(true);
    expect(encodeQr('x'.repeat(2331)).version).toBe(40);
    expect(() => encodeQr('x'.repeat(2332))).toThrow();
    expect(qrPath(q)).toMatch(/^M\d/);
  });
  it('writes the version information block from version 7', () => {
    const q = encodeQr('y'.repeat(120));
    expect(q.version).toBe(7);
    // Version 7's 18-bit word is 000111110010010100, least significant bit first, bottom-left block by columns.
    const word = 0b000111110010010100;
    for (let i = 0; i < 18; i++) expect(q.dark(Math.floor(i / 3), q.size - 11 + (i % 3))).toBe(((word >>> i) & 1) === 1);
  });
});
