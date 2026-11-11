import { afterEach, describe, expect, it } from 'vitest';
import { checkUrl, relayFetch, setRelayAllowed } from '../electron/netRelay';
import { defs } from '../src/engine/functions/web';
import { clearOnlineCache, setOnlineAllowed, settled } from '../src/engine/online';
import { makeCtx } from './fnctx';
import type { Value } from '../src/engine/types';

afterEach(() => { setRelayAllowed(false); setOnlineAllowed(false); clearOnlineCache(); delete (globalThis as { clerdo?: unknown }).clerdo; });

const respond = (body: string | Uint8Array, init: ResponseInit = {}) => new Response(body as BodyInit, init);
const seq = (...rs: Response[]): { f: typeof fetch; urls: string[]; inits: RequestInit[] } => {
  const urls: string[] = [], inits: RequestInit[] = [];
  let i = 0;
  return { urls, inits, f: (async (u: string, init: RequestInit) => { urls.push(u); inits.push(init); return rs[i++]!; }) as unknown as typeof fetch };
};

describe('desktop network relay', () => {
  it('accepts only public https addresses', () => {
    expect(checkUrl('https://example.com/x')?.href).toBe('https://example.com/x');
    for (const bad of ['http://example.com', 'https://user:pw@example.com', 'https://localhost/', 'https://127.0.0.1/', 'https://10.1.2.3/', 'https://192.168.0.5/',
      'https://169.254.1.1/', 'https://[::1]/', 'https://[fd00::1]/', 'https://printer.local/', 'https://example.com:8443/', 'https://2130706433/', 'file:///etc/passwd', 42, 'x'.repeat(3000)]) {
      expect(checkUrl(bad), String(bad).slice(0, 40)).toBeNull();
    }
  });
  it('needs both its own setting and the page flag', async () => {
    const { f } = seq(respond('hi'));
    expect(await relayFetch({ url: 'https://example.com', kind: 'text', allowed: true }, f)).toEqual({ ok: false, error: 'internet access is off' });
    setRelayAllowed(true);
    expect(await relayFetch({ url: 'https://example.com', kind: 'text', allowed: false }, f)).toEqual({ ok: false, error: 'internet access is off' });
    expect(setRelayAllowed('yes')).toBe(false);
    setRelayAllowed(true);
    expect(await relayFetch({ url: 'https://example.com', kind: 'evil', allowed: true }, f)).toEqual({ ok: false, error: 'bad request' });
    expect(await relayFetch(null, f)).toEqual({ ok: false, error: 'bad request' });
  });
  it('fetches without credentials and follows only https redirects', async () => {
    setRelayAllowed(true);
    const s = seq(respond('', { status: 302, headers: { location: '/next' } }), respond('done', { headers: { 'content-type': 'text/plain' } }));
    const r = await relayFetch({ url: 'https://example.com/start', kind: 'text', allowed: true }, s.f);
    expect(r).toMatchObject({ ok: true, text: 'done', url: 'https://example.com/next' });
    expect(s.inits[0]).toMatchObject({ credentials: 'omit', redirect: 'manual', referrerPolicy: 'no-referrer' });
    const bad = seq(respond('', { status: 301, headers: { location: 'http://example.com/' } }));
    expect(await relayFetch({ url: 'https://example.com/', kind: 'text', allowed: true }, bad.f)).toEqual({ ok: false, error: 'redirect not allowed' });
    const toLocal = seq(respond('', { status: 307, headers: { location: 'https://127.0.0.1/admin' } }));
    expect(await relayFetch({ url: 'https://example.com/', kind: 'text', allowed: true }, toLocal.f)).toEqual({ ok: false, error: 'redirect not allowed' });
  });
  it('caps size and checks image types', async () => {
    setRelayAllowed(true);
    expect(await relayFetch({ url: 'https://example.com/', kind: 'text', allowed: true, maxBytes: 4 }, seq(respond('12345')).f)).toEqual({ ok: false, error: 'too large' });
    expect(await relayFetch({ url: 'https://example.com/a.png', kind: 'binary', allowed: true }, seq(respond('<html>', { headers: { 'content-type': 'text/html' } })).f)).toEqual({ ok: false, error: 'not an image' });
    const img = await relayFetch({ url: 'https://example.com/a.png', kind: 'binary', allowed: true }, seq(respond(new Uint8Array([137, 80, 78, 71]), { headers: { 'content-type': 'image/png' } })).f);
    expect(img).toMatchObject({ ok: true, mime: 'image/png' });
    expect(img.ok && [...img.data!]).toEqual([137, 80, 78, 71]);
  });
  it('the page uses the relay when the desktop app provides one', async () => {
    const calls: unknown[] = [];
    let mainAllowed = false;
    (globalThis as { clerdo?: unknown }).clerdo = {
      netSetAllowed: async (on: boolean) => { mainAllowed = on; return on; },
      netFetch: async (req: { url: string; allowed: boolean }) => { calls.push(req); return mainAllowed && req.allowed ? { ok: true, status: 200, mime: 'text/plain', text: 'via relay' } : { ok: false, error: 'off' }; },
    };
    setOnlineAllowed(true);
    await Promise.resolve();
    expect(mainAllowed).toBe(true);
    const call = (...a: Value[]) => defs.find((d) => d.name === 'WEBSERVICE')!.fn(a, makeCtx());
    call('https://example.com/r');
    await settled();
    expect(call('https://example.com/r')).toBe('via relay');
    expect(calls[0]).toMatchObject({ url: 'https://example.com/r', kind: 'text', allowed: true });
    setOnlineAllowed(false);
    await Promise.resolve();
    expect(mainAllowed).toBe(false);
  });
});
