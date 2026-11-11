// Device-to-device links over WebRTC data channels. There is no server: the two sides swap an offer code and an
// answer code by hand (copy and paste, or a QR code). The channel is encrypted by DTLS.
import { packText, unpackText } from './b64';
import type { Link } from './session';
import { LIMITS, type Msg } from './validate';
import { T } from '../../ui/i18n';

export const STUN_KEY = 'clerdogrid.stun';
export const DEFAULT_STUN = 'stun:stun.l.google.com:19302';

export const TURN_KEY = 'clerdogrid.turn';

/** A relay for devices that can't reach each other directly (strict NATs, UDP blocked). */
export interface TurnServer { urls: string; username: string; credential: string }
export interface TurnSettings extends TurnServer { /** Put the relay (with its login) in invite codes, so guests use it too. */ share: boolean }

const turnUrlOk = (u: string): boolean => /^turns?:[^\s/]+(\?transport=(udp|tcp))?$/i.test(u);

export const checkTurn = (x: unknown): TurnServer | null => {
  if (!x || typeof x !== 'object') return null;
  const t = x as Record<string, unknown>;
  const ok = (v: unknown) => typeof v === 'string' && v.length <= 512;
  if (!ok(t.urls) || !ok(t.username) || !ok(t.credential) || !turnUrlOk((t.urls as string).trim())) return null;
  return { urls: (t.urls as string).trim(), username: t.username as string, credential: t.credential as string };
};

export const loadTurn = (): TurnSettings => {
  try {
    const v = JSON.parse(localStorage.getItem(TURN_KEY) ?? 'null') as Record<string, unknown> | null;
    const t = checkTurn(v);
    if (t) return { ...t, share: v?.share === true };
    if (v && typeof v === 'object') return { urls: typeof v.urls === 'string' ? v.urls : '', username: typeof v.username === 'string' ? v.username : '', credential: typeof v.credential === 'string' ? v.credential : '', share: v.share === true };
  } catch { /* nothing saved */ }
  return { urls: '', username: '', credential: '', share: false };
};
export const saveTurn = (t: TurnSettings): void => { try { localStorage.setItem(TURN_KEY, JSON.stringify(t)); } catch { /* storage blocked */ } };

/** STUN servers (a space- or comma-separated list; empty means none, for local networks only), plus a TURN relay if set. */
export const iceServers = (extra?: TurnServer | null): RTCIceServer[] => {
  let v: string | null = null;
  try { v = localStorage.getItem(STUN_KEY); } catch { /* storage blocked */ }
  const urls = v === null ? [DEFAULT_STUN] : v.split(/[\s,]+/).filter((u) => /^stuns?:[^\s/]+$/i.test(u));
  const out: RTCIceServer[] = urls.length ? [{ urls }] : [];
  const own = checkTurn(loadTurn());
  for (const t of [own, extra ? checkTurn(extra) : null]) if (t && !out.some((s) => s.urls === t.urls)) out.push({ urls: t.urls, username: t.username, credential: t.credential });
  return out;
};

const PREFIX = { offer: 'CGO1.', answer: 'CGA1.' } as const;
/** An invite that also carries the host's TURN relay. */
const OFFER_TURN = 'CGO2.';
const MAX_SDP = 64 * 1024;

/** Drops SDP lines a data channel doesn't need, so codes and QR codes stay small. TCP candidates rarely help without a relay. */
export const minifySdp = (sdp: string): string => sdp.split(/\r?\n/).filter((l) => l && !/^a=(extmap-allow-mixed|msid-semantic|ice-options:trickle)/.test(l) && !(/^a=candidate:/.test(l) && / tcp /i.test(l))).join('\r\n') + '\r\n';

export const encodeCode = (kind: 'offer' | 'answer', sdp: string, turn?: TurnServer | null): string => {
  const t = kind === 'offer' && turn ? checkTurn(turn) : null;
  if (t) return OFFER_TURN + packText(JSON.stringify({ s: minifySdp(sdp), t }));
  return PREFIX[kind] + packText(minifySdp(sdp));
};

export interface CodeContent { sdp: string; turn: TurnServer | null }

/** Reads a pasted code (or a join link holding one). Throws a short, readable error when it isn't the expected kind. */
export const readCode = (text: string, kind: 'offer' | 'answer'): CodeContent => {
  let t = text.trim();
  const m = /[#&]join=([A-Za-z0-9._-]+)/.exec(t);
  if (m) t = decodeURIComponent(m[1]!);
  t = t.replace(/\s+/g, '');
  const isOffer = t.startsWith(PREFIX.offer) || t.startsWith(OFFER_TURN);
  if (kind === 'offer' ? t.startsWith(PREFIX.answer) : isOffer) throw new Error(kind === 'offer' ? T('That\'s a reply code. Paste the invite code from the host.') : T('That\'s an invite code. Paste the reply code from the person joining.'));
  const prefix = kind === 'offer' && t.startsWith(OFFER_TURN) ? OFFER_TURN : PREFIX[kind];
  const broken = T('That code isn\'t complete. Copy the whole code and try again.');
  if (!t.startsWith(prefix) || t.length > MAX_SDP) throw new Error(broken);
  let body: string;
  try { body = unpackText(t.slice(prefix.length), MAX_SDP); } catch { throw new Error(broken); }
  let sdp = body, turn: TurnServer | null = null;
  if (prefix === OFFER_TURN) {
    let v: unknown;
    try { v = JSON.parse(body); } catch { throw new Error(broken); }
    const o = v as { s?: unknown; t?: unknown } | null;
    if (!o || typeof o.s !== 'string') throw new Error(broken);
    sdp = o.s;
    turn = checkTurn(o.t);
    if (!turn) throw new Error(T('That code\'s relay server settings aren\'t valid.'));
  }
  if (!sdp.startsWith('v=0') || !/\r?\nm=application /.test(sdp)) throw new Error(T('That code isn\'t a ClerdoGrid session code.'));
  return { sdp, turn };
};

export const decodeCode = (text: string, kind: 'offer' | 'answer'): string => readCode(text, kind).sdp;

export const joinUrl = (offerCode: string): string => {
  const base = typeof location !== 'undefined' && /^https?:$/.test(location.protocol) ? `${location.origin}${location.pathname}` : 'https://clerdoapp.github.io/clerdogrid/';
  return `${base}#join=${offerCode}`;
};

/** Waits until ICE candidates are gathered (no trickle: the code carries them), or a few seconds at most. */
const gathered = (pc: RTCPeerConnection, ms = 5000): Promise<void> => new Promise((resolve) => {
  if (pc.iceGatheringState === 'complete') { resolve(); return; }
  const done = () => { pc.removeEventListener('icegatheringstatechange', check); clearTimeout(t); resolve(); };
  const check = () => { if (pc.iceGatheringState === 'complete') done(); };
  pc.addEventListener('icegatheringstatechange', check);
  const t = setTimeout(done, ms);
});

export interface Pairing {
  /** The code to give to the other side. */
  code: string;
  /** Resolves when the data channel opens. */
  link: Promise<Link>;
  cancel(): void;
}

export interface Invite extends Pairing { accept(answer: string): Promise<void> }

const OPEN_TIMEOUT = 3 * 60 * 1000;

const channelLink = (pc: RTCPeerConnection, dcPromise: Promise<RTCDataChannel>): Promise<Link> => new Promise((resolve, reject) => {
  const t = setTimeout(() => { reject(new Error(T('The connection timed out.'))); pc.close(); }, OPEN_TIMEOUT);
  pc.addEventListener('connectionstatechange', () => { if (pc.connectionState === 'failed') { clearTimeout(t); reject(new Error(T('Couldn\'t connect. Both devices need a network connection; on different networks a STUN server must be reachable.'))); } });
  dcPromise.then((dc) => {
    const ready = () => { clearTimeout(t); resolve(new RtcLink(pc, dc)); };
    if (dc.readyState === 'open') ready(); else dc.addEventListener('open', ready, { once: true });
  }, reject);
});

/** Host side: makes an invite code for one person. */
export const createInvite = async (): Promise<Invite> => {
  const pc = new RTCPeerConnection({ iceServers: iceServers() });
  const turn = loadTurn();
  const shared = turn.share ? checkTurn(turn) : null;
  const dc = pc.createDataChannel('clerdogrid', { ordered: true });
  await pc.setLocalDescription(await pc.createOffer());
  await gathered(pc);
  const code = encodeCode('offer', pc.localDescription!.sdp, shared);
  const link = channelLink(pc, Promise.resolve(dc));
  link.catch(() => undefined);
  return {
    code, link,
    cancel: () => pc.close(),
    accept: async (answer: string) => {
      const sdp = decodeCode(answer, 'answer');
      if (pc.signalingState !== 'have-local-offer') throw new Error(T('This invite was already used. Make a new one.'));
      await pc.setRemoteDescription({ type: 'answer', sdp });
    },
  };
};

/** Guest side: answers an invite code; the returned code goes back to the host. */
export const answerInvite = async (offer: string): Promise<Pairing> => {
  const { sdp, turn } = readCode(offer, 'offer');
  const pc = new RTCPeerConnection({ iceServers: iceServers(turn) });
  const dc = new Promise<RTCDataChannel>((resolve) => { pc.ondatachannel = (e) => resolve(e.channel); });
  await pc.setRemoteDescription({ type: 'offer', sdp });
  await pc.setLocalDescription(await pc.createAnswer());
  await gathered(pc);
  const link = channelLink(pc, dc);
  link.catch(() => undefined);
  return { code: encodeCode('answer', pc.localDescription!.sdp), link, cancel: () => pc.close() };
};

// ---------- Framing ----------
// Messages are JSON text. Long ones are split into frames "\u0001id:index:count:" + piece and joined on arrival.

const FRAME = 16 * 1024;
const HIGH_WATER = 4 * 1024 * 1024;

class RtcLink implements Link {
  onMessage: ((m: unknown) => void) | null = null;
  onClose: (() => void) | null = null;
  onState: ((s: 'connected' | 'reconnecting') => void) | null = null;
  private queue: string[] = [];
  private nextId = 1;
  private parts = new Map<number, { n: number; got: number; size: number; pieces: string[] }>();
  private closed = false;

  constructor(private pc: RTCPeerConnection, private dc: RTCDataChannel) {
    dc.bufferedAmountLowThreshold = 1024 * 1024;
    dc.onbufferedamountlow = () => this.pump();
    dc.onmessage = (e) => this.receive(e.data);
    dc.onclose = () => this.lost();
    pc.addEventListener('connectionstatechange', () => {
      const s = pc.connectionState;
      if (s === 'disconnected') this.onState?.('reconnecting');
      else if (s === 'connected') this.onState?.('connected');
      else if (s === 'failed' || s === 'closed') this.lost();
    });
  }

  send(m: Msg): void {
    if (this.closed) return;
    const text = JSON.stringify(m);
    if (text.length <= FRAME) this.queue.push(text);
    else {
      const id = this.nextId++;
      const n = Math.ceil(text.length / FRAME);
      for (let i = 0; i < n; i++) this.queue.push(`\u0001${id}:${i}:${n}:${text.slice(i * FRAME, (i + 1) * FRAME)}`);
    }
    this.pump();
  }

  private pump(): void {
    while (this.queue.length && this.dc.readyState === 'open' && this.dc.bufferedAmount < HIGH_WATER) {
      try { this.dc.send(this.queue.shift()!); } catch { this.lost(); return; }
    }
  }

  private receive(data: unknown): void {
    if (typeof data !== 'string' || data.length > LIMITS.frame) return;
    if (data.charCodeAt(0) === 1) {
      const m = /^\u0001(\d{1,9}):(\d{1,9}):(\d{1,9}):/.exec(data);
      if (!m) return;
      const id = +m[1]!, i = +m[2]!, n = +m[3]!;
      if (n < 2 || i >= n || n * FRAME > LIMITS.message * 1.1) return;
      let a = this.parts.get(id);
      if (!a) {
        if (this.parts.size >= 4) return;
        a = { n, got: 0, size: 0, pieces: new Array<string>(n) };
        this.parts.set(id, a);
      }
      if (a.n !== n || a.pieces[i] !== undefined) return;
      const piece = data.slice(m[0].length);
      a.pieces[i] = piece; a.got++; a.size += piece.length;
      if (a.size > LIMITS.message) { this.parts.delete(id); return; }
      if (a.got === n) { this.parts.delete(id); this.deliver(a.pieces.join('')); }
      return;
    }
    this.deliver(data);
  }

  private deliver(text: string): void {
    let v: unknown;
    try { v = JSON.parse(text); } catch { return; }
    this.onMessage?.(v);
  }

  private lost(): void {
    if (this.closed) return;
    this.closed = true;
    try { this.dc.close(); this.pc.close(); } catch { /* already closed */ }
    this.onClose?.();
  }

  close(): void {
    if (this.closed) return;
    // Let queued frames (a goodbye) go out first.
    const finish = () => { this.closed = true; try { this.dc.close(); this.pc.close(); } catch { /* closed */ } };
    // A data channel can drop what is still buffered when it closes, so wait for the buffer to drain.
    const started = Date.now();
    const wait = () => { this.pump(); if ((this.queue.length || this.dc.bufferedAmount) && Date.now() - started < 3000) setTimeout(wait, 50); else setTimeout(finish, 200); };
    wait();
    this.onClose = null;
  }
}
