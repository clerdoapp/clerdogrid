import { describe, expect, it } from 'vitest';
import { checkUrl, isPublicAddress, pinnedFetch, relayFetch, safeLookup, setRelayAllowed } from '../electron/netRelay';

type Resolver = Parameters<typeof safeLookup>[3];
const fake = (addrs: { address: string; family: number }[]): Resolver =>
  ((_h: string, _o: unknown, cb: (e: null, a: unknown) => void) => cb(null, addrs)) as unknown as Resolver;
const run = (addrs: { address: string; family: number }[], all = false) => new Promise<{ err: string | null; out: unknown }>((res) =>
  safeLookup('example.com', { all }, (err, address) => res({ err: err ? (err as NodeJS.ErrnoException).code ?? err.message : null, out: address }), fake(addrs)));

describe('relay address checks', () => {
  it('knows public and private addresses', () => {
    for (const a of ['93.184.216.34', '8.8.8.8', '2606:2800:220:1::1', '::ffff:8.8.8.8', '64:ff9b::8.8.8.8']) expect(isPublicAddress(a), a).toBe(true);
    for (const a of ['127.0.0.1', '10.1.2.3', '172.20.0.1', '192.168.1.1', '169.254.169.254', '100.64.0.1', '0.0.0.0', '224.0.0.1', '255.255.255.255',
      '::1', '::', 'fe80::1', 'fd00::1', 'fc00::5', 'ff02::1', '2001:db8::1', '::ffff:127.0.0.1', '::ffff:10.0.0.1', '64:ff9b::7f00:1', '64:ff9b::10.0.0.1', 'not-an-ip'])
      expect(isPublicAddress(a), a).toBe(false);
  });

  it('refuses private literals in URLs', () => {
    expect(checkUrl('https://[::ffff:127.0.0.1]/')).toBeNull();
    expect(checkUrl('https://169.254.169.254/latest')).toBeNull();
    expect(checkUrl('https://8.8.8.8/')?.hostname).toBe('8.8.8.8');
  });

  it('blocks names that resolve to a private address', async () => {
    expect((await run([{ address: '10.0.0.5', family: 4 }])).err).toBe('EADDRBLOCKED');
    expect((await run([{ address: '93.184.216.34', family: 4 }, { address: '127.0.0.1', family: 4 }])).err).toBe('EADDRBLOCKED');
    expect((await run([])).err).toBe('EADDRBLOCKED');
    expect(await run([{ address: '93.184.216.34', family: 4 }])).toEqual({ err: null, out: '93.184.216.34' });
    expect((await run([{ address: '93.184.216.34', family: 4 }], true)).out).toEqual([{ address: '93.184.216.34', family: 4 }]);
  });

  it('reports a blocked connection as not allowed', async () => {
    setRelayAllowed(true);
    // localhost.localtest-style names aren't needed: the default fetch resolves "localhost" through safeLookup, which refuses it.
    const blocked = (async () => { throw Object.assign(new Error('x'), { code: 'EADDRBLOCKED' }); }) as unknown as typeof fetch;
    expect(await relayFetch({ url: 'https://example.com/', kind: 'text', allowed: true }, blocked)).toEqual({ ok: false, error: 'address not allowed' });
    await expect(pinnedFetch('https://localhost.:1/')).rejects.toBeTruthy();
    setRelayAllowed(false);
  });
});
