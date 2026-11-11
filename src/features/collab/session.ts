// A live session: one host (the authority that orders every change) and guests linked to it (a star).
// Guests apply their own edits at once, send them with the last host sequence they had seen, and rebase them on
// whatever the host ordered first. The host transforms incoming edits past everything they hadn't seen yet.
import type { ChangeEvent, GridStore } from '../../app/store';
import { captureLight, opsFromEvent, type Op } from './ops';
import { applyOp, BLANK_SHEET, deletedSheets, rebaseHistory, settleSheets } from './apply';
import { cloneSheet } from '../../app/clone';
import { shiftRange, shiftPos, xformLists } from './ot';
import { bookHash, decodeBook, encodeBook } from './wire';
import { placeSheets } from './sheets';

/** The host gives every sheet an id and sort key when a session starts, including the copies in its undo history. */
const identifySheets = (store: GridStore): void => {
  placeSheets(store.wb.sheets);
  const live = store.wb.sheets;
  const copy = (to: { cid?: string; cpos?: string; name: string }, i: number) => {
    if (to.cid) return;
    const src = live.find((x) => x.name === to.name) ?? live[i];
    if (src) { to.cid = src.cid; to.cpos = src.cpos; }
  };
  store.rebaseHistory((step) => {
    for (const snap of step.snaps) {
      if (snap.kind === 'book') { snap.data.sheets.forEach(copy); placeSheets(snap.data.sheets); }
      else if (snap.kind === 'sheet') copy(snap.data, snap.sheet);
    }
    return true;
  });
};
import * as V from './validate';
import type { Msg, PeerInfo, Presence } from './validate';
import { T } from '../../ui/i18n';

export const PROTOCOL = 1;

/** A message pipe to one other device. Messages are plain JSON-safe objects. */
export interface Link {
  send(m: Msg): void;
  close(): void;
  onMessage: ((m: unknown) => void) | null;
  onClose: (() => void) | null;
  /** Unstable network: the link may come back by itself. */
  onState?: ((s: 'connected' | 'reconnecting') => void) | null;
}

export const COLORS = ['#e5484d', '#3e63dd', '#d6409f', '#f76b15', '#8e4ec6', '#0090ff', '#a18072', '#c2a300', '#e54666', '#5b5bd6'];

export type Status = 'connecting' | 'live' | 'reconnecting' | 'lost' | 'ended';

interface Guest { info: PeerInfo; link: Link | null; lastOpId: number; lastSeen: number }
interface LogEntry { seq: number; origin: string; ops: Op[] }
interface Pending { id: number; base: number; ops: Op[] }

const LOG_MAX = 5000;
const PING_MS = 4000;
// Generous: background tabs run timers rarely, and real drops are noticed by the connection itself.
const DEAD_MS = 75000;
const HASH_MS = 2500;

export const newPeerId = (): string => {
  const b = new Uint8Array(9);
  crypto.getRandomValues(b);
  return Array.from(b, (x) => x.toString(36).padStart(2, '0')).join('').slice(0, 16);
};

const rangesOf = (store: GridStore): [number, number, number, number][] => store.sel.ranges.slice(-V.LIMITS.ranges).map((r) => [r.r1, r.c1, r.r2, r.c2]);

export class Session {
  readonly role: 'host' | 'guest';
  readonly store: GridStore;
  me: PeerInfo;
  roster: PeerInfo[] = [];
  presence = new Map<string, Presence>();
  status: Status = 'connecting';
  following: string | null = null;
  /** Called after anything visible changes (roster, presence, status). */
  onUpdate: () => void = () => undefined;
  onNotice: (text: string) => void = () => undefined;

  // host
  private guests = new Map<string, Guest>();
  private log: LogEntry[] = [];
  seq = 0;
  private hostOpId = 0;
  private hashSeq = -1;
  private nextHashAt = 0;
  // guest
  private host: Link | null = null;
  acked = 0;
  pending: Pending[] = [];
  private nextOpId = 1;
  private lastResync = 0;
  private hostSeen = 0;

  private wbRef: unknown;
  private stopObserve: () => void;
  private stopSub: () => void;
  private timers: ReturnType<typeof setInterval>[] = [];
  private lastPresence = '';
  private presenceTimer: ReturnType<typeof setTimeout> | null = null;
  private followSel = '';
  private loading = false;

  constructor(store: GridStore, role: 'host' | 'guest', me: PeerInfo) {
    this.store = store;
    this.role = role;
    this.me = { ...me, host: role === 'host' || undefined };
    this.wbRef = store.wb;
    if (role === 'host') { this.status = 'live'; this.roster = [this.me]; identifySheets(store); }
    this.stopObserve = store.observe({ begin: (e, scope) => captureLight(store, e, scope), end: (e) => this.local(e) });
    this.stopSub = store.subscribe(() => this.watch());
    this.timers.push(setInterval(() => this.tick(), PING_MS));
    if (role === 'host') this.timers.push(setInterval(() => this.sendHash(), HASH_MS));
  }

  get peers(): PeerInfo[] { return this.roster.filter((p) => p.id !== this.me.id); }
  get liveGuests(): number { return [...this.guests.values()].filter((g) => g.link).length; }

  // ---------- Local changes ----------

  private local(e: ChangeEvent): void {
    if (this.status === 'ended') return;
    let ops: Op[];
    try { ops = opsFromEvent(this.store, e); } catch (err) { console.error(err); this.requestFullSync(); return; }
    if (!ops.length) return;
    // Sheets deleted here are kept like deleted sheets from others (apply.ts), so they take later edits too.
    for (const o of ops) if (o.t === 'sheetDel') {
      const book = e.snaps.find((x) => x.kind === 'book');
      const sh = book?.kind === 'book' ? book.data.sheets.find((x) => x.cid === o.s) : undefined;
      if (sh) deletedSheets(this.store.wb).set(o.s, cloneSheet(sh));
    }
    // A sheet brought back here (undo) uses the kept copy too, so it matches what everyone else brings back.
    const back = ops.filter((o): o is Extract<Op, { t: 'sheetAdd' }> => o.t === 'sheetAdd' && deletedSheets(this.store.wb).has(o.s));
    if (back.length) this.store.applyRemote(() => {
      const wb = this.store.wb;
      for (const o of back) {
        const kept = deletedSheets(wb).get(o.s)!;
        deletedSheets(wb).delete(o.s);
        const i = wb.sheets.findIndex((x) => x.cid === o.s);
        if (i < 0) continue;
        kept.name = wb.sheets[i]!.name; kept.cpos = wb.sheets[i]!.cpos;
        wb.sheets[i] = kept;
      }
      this.store.engine.sheetsChanged();
    });
    // A stand-in sheet (blank, or a deleted one shown because nothing else was left) goes away once there is a
    // real sheet, as it does on every other copy.
    if (ops.some((o) => o.t === 'sheetAdd')) {
      const before = this.store.wb.sheets.map((x) => x.cid);
      const wb = this.store.wb;
      const dead = deletedSheets(wb);
      if (wb.sheets.some((x) => x.cid === BLANK_SHEET || (x.cid && dead.has(x.cid)))) {
        this.store.applyRemote(() => { settleSheets(wb); this.store.engine.sheetsChanged(); });
        const gone = before.filter((id) => !wb.sheets.some((x) => x.cid === id));
        for (const id of gone) rebaseHistory(this.store, { t: 'sheetDel', s: id!, name: '' }, before);
      }
    }
    if (this.role === 'host') { this.commit(ops, this.me.id, ++this.hostOpId); return; }
    const p: Pending = { id: this.nextOpId++, base: this.acked, ops };
    this.pending.push(p);
    if (this.pending.length > V.LIMITS.pending) { this.onNotice(T('Too many changes are waiting to be sent. Reloading the shared workbook.')); this.pending = []; this.requestFullSync(); return; }
    this.flush();
  }

  /**
   * One batch is in flight at a time: the host must know about a guest's previous batch before it transforms the next,
   * so everything waiting is merged into one batch when the previous one is acknowledged.
   */
  private inflight = false;
  private flush(): void {
    if (this.inflight || !this.pending.length || !this.host || this.status !== 'live') return;
    if (this.pending.length > 1) {
      const all = this.pending.splice(0);
      this.pending = [{ id: all[all.length - 1]!.id, base: this.acked, ops: all.flatMap((x) => x.ops) }];
    }
    const p = this.pending[0]!;
    p.base = this.acked;
    this.inflight = true;
    this.host.send({ t: 'op', id: p.id, base: p.base, ops: p.ops });
  }

  private requestFullSync(): void {
    if (this.role === 'host') { this.broadcastSnapshot(); return; }
    if (this.host && Date.now() - this.lastResync > 3000) { this.lastResync = Date.now(); this.host.send({ t: 'resync' }); }
  }

  /** Applies ops from someone else to this copy (no undo step), patching undo history and cursors. */
  private applyIncoming(ops: Op[]): void {
    const store = this.store;
    store.applyRemote(() => {
      for (const op of ops) {
        const before = store.wb.sheets.map((s) => s.cid);
        let ok = false;
        try { ok = applyOp(store.wb, op, { engine: store.engine, store }); } catch (err) { console.error(err); }
        if (!ok) continue;
        rebaseHistory(store, op, before);
        if (op.t === 'struct') this.shiftPresence(op.s, op.a, op.at, op.n);
      }
    });
    this.wbRef = store.wb;
  }

  private shiftPresence(sheet: string, axis: 'row' | 'col', at: number, n: number): void {
    for (const p of this.presence.values()) {
      if (p.sheet !== sheet) continue;
      const ranges = p.ranges.map(([r1, c1, r2, c2]) => shiftRange({ r1, c1, r2, c2 }, axis, at, n)).filter((r) => !!r).map((r) => [r!.r1, r!.c1, r!.r2, r!.c2] as [number, number, number, number]);
      const a = shiftPos({ row: p.active[0], col: p.active[1] }, axis, at, n);
      p.active = [a.row, a.col];
      p.ranges = ranges.length ? ranges : [[a.row, a.col, a.row, a.col]];
      if (p.editing) { const e = shiftPos({ row: p.editing[0], col: p.editing[1] }, axis, at, n); p.editing = [e.row, e.col]; }
    }
  }

  // ---------- Host ----------

  private commit(ops: Op[], origin: string, id: number): void {
    this.seq++;
    this.log.push({ seq: this.seq, origin, ops });
    if (this.log.length > LOG_MAX) this.log.splice(0, this.log.length - LOG_MAX);
    if (origin === this.me.id) for (const op of ops) if (op.t === 'struct') this.shiftPresence(op.s, op.a, op.at, op.n);
    this.broadcast({ t: 'op', seq: this.seq, origin, id, ops });
  }

  private broadcast(m: Msg, except?: string): void {
    for (const [id, g] of this.guests) if (g.link && id !== except) g.link.send(m);
  }

  /** The host takes a new guest connection; the guest says who it is in its first message. */
  addGuestLink(link: Link, ro = false): void {
    let gid: string | null = null;
    link.onMessage = (raw) => {
      let m: Msg;
      try { m = V.message(raw); } catch (err) { console.warn(err); return; }
      if (gid) { const g = this.guests.get(gid); if (g) g.lastSeen = Date.now(); }
      if (!gid) {
        if (m.t !== 'hello') return;
        if (m.proto !== PROTOCOL) { this.onNotice(T('{p0} uses a different version of ClerdoGrid and can\'t join.', { p0: m.name || T('Someone') })); link.send({ t: 'end' }); link.close(); return; }
        gid = m.id;
        this.welcome(link, m, ro);
        return;
      }
      this.hostReceive(gid, m);
    };
    link.onClose = () => {
      if (!gid) return;
      const g = this.guests.get(gid);
      if (g && g.link === link) { g.link = null; g.info.away = true; this.presence.delete(gid); this.sendRoster(); this.onNotice(T('{name} lost the connection.', { name: g.info.name })); }
    };
  }

  private welcome(link: Link, m: Extract<Msg, { t: 'hello' }>, ro: boolean): void {
    const old = this.guests.get(m.id);
    if (old?.link && old.link !== link) { old.link.onClose = null; old.link.close(); }
    const used = new Set(this.roster.filter((p) => p.id !== m.id).map((p) => p.color));
    const color = old?.info.color ?? COLORS.find((c) => !used.has(c)) ?? COLORS[this.guests.size % COLORS.length]!;
    const name = m.name.trim().slice(0, 60) || T('Guest');
    const g: Guest = old ?? { info: { id: m.id, name, color }, link, lastOpId: 0, lastSeen: Date.now() };
    g.link = link; g.lastSeen = Date.now();
    g.info = { id: m.id, name, color, ro: (old ? old.info.ro : ro) || undefined };
    this.guests.set(m.id, g);
    // Changes the guest made while disconnected.
    if (m.resume && !g.info.ro) for (const p of m.resume.pending) this.acceptOp(g, p.id, p.base, p.ops);
    link.send({ t: 'welcome', you: m.id, color, seq: this.seq, book: encodeBook(this.store.wb), name: this.store.file.name, ro: !!g.info.ro, roster: this.currentRoster() });
    this.sendRoster();
    this.sendPresenceNow(true);
    for (const p of this.presence.values()) if (p.id !== m.id) link.send({ t: 'presence', p });
    this.onNotice(old ? T('{name} is back.', { name }) : T('{name} joined.', { name }));
  }

  private acceptOp(g: Guest, id: number, base: number, incoming: Op[]): void {
    if (id <= g.lastOpId) return;
    const start = this.log.length ? this.log[0]!.seq : this.seq + 1;
    if (base > this.seq || (base < start - 1 && this.log.some((e) => e.origin !== g.info.id))) {
      // Too old to transform safely: the guest gets the whole workbook again.
      g.lastOpId = id;
      g.link?.send({ t: 'snap', seq: this.seq, book: encodeBook(this.store.wb), name: this.store.file.name });
      return;
    }
    let ops = incoming;
    for (const e of this.log) if (e.seq > base && e.origin !== g.info.id) ops = xformLists(ops, e.ops, true)[0];
    g.lastOpId = id;
    this.applyIncoming(ops);
    this.commit(ops, g.info.id, id);
  }

  private hostReceive(gid: string, m: Msg): void {
    const g = this.guests.get(gid);
    if (!g) return;
    switch (m.t) {
      case 'op':
        if (g.info.ro || m.base === undefined) return;
        this.acceptOp(g, m.id, m.base, m.ops);
        break;
      case 'presence':
        if (m.p.id !== gid) return;
        this.presence.set(gid, m.p);
        this.broadcast(m, gid);
        this.followUpdate(gid);
        this.onUpdate();
        break;
      case 'resync': g.link?.send({ t: 'snap', seq: this.seq, book: encodeBook(this.store.wb), name: this.store.file.name }); break;
      case 'bye':
        g.link?.close();
        this.guests.delete(gid);
        this.presence.delete(gid);
        if (this.following === gid) this.following = null;
        this.sendRoster();
        this.onNotice(T('{name} left.', { name: g.info.name }));
        break;
      case 'ping': g.link?.send({ t: 'pong' }); break;
      default: break;
    }
  }

  private currentRoster(): PeerInfo[] {
    return [this.me, ...[...this.guests.values()].map((g) => ({ ...g.info, away: !g.link || undefined }))];
  }

  private sendRoster(): void {
    this.roster = this.currentRoster();
    this.broadcast({ t: 'roster', peers: this.roster });
    this.onUpdate();
  }

  /** Lets a guest who lost the connection be removed from the list. */
  forget(id: string): void {
    if (this.role !== 'host') return;
    const g = this.guests.get(id);
    if (!g) return;
    g.link?.close();
    this.guests.delete(id);
    this.presence.delete(id);
    this.sendRoster();
  }

  private broadcastSnapshot(): void {
    this.seq++;
    this.log = [];
    const book = encodeBook(this.store.wb);
    this.broadcast({ t: 'snap', seq: this.seq, book, name: this.store.file.name });
  }

  private sendHash(): void {
    if (this.status !== 'live' || this.hashSeq === this.seq || !this.liveGuests || this.store.editing || Date.now() < this.nextHashAt) return;
    this.hashSeq = this.seq;
    const t0 = Date.now();
    const h = bookHash(this.store.wb);
    // Big workbooks are checked less often, so the check never takes more than a few percent of the time.
    this.nextHashAt = Date.now() + (Date.now() - t0) * 30;
    this.broadcast({ t: 'hash', seq: this.seq, h });
  }

  // ---------- Guest ----------

  /** Connects (or reconnects) to the host. */
  attachHost(link: Link): void {
    if (this.host && this.host !== link) { this.host.onClose = null; this.host.close(); }
    this.host = link;
    this.hostSeen = Date.now();
    this.status = 'connecting';
    // Nothing can be typed until the host's copy arrives, so no edit is made against a state about to be replaced.
    this.store.readOnly = true;
    link.onMessage = (raw) => {
      let m: Msg;
      try { m = V.message(raw); } catch (err) { console.warn(err); return; }
      this.hostSeen = Date.now();
      this.guestReceive(m);
    };
    link.onClose = () => {
      if (this.host !== link || this.status === 'ended') return;
      this.host = null;
      this.status = 'lost';
      this.presence.clear();
      this.onNotice(T('Lost the connection to the host. Ask them for a new invite code to reconnect; your changes are kept.'));
      this.onUpdate();
    };
    link.onState = (s) => { if (this.host === link && this.status !== 'ended' && this.status !== 'lost') { this.status = s === 'reconnecting' ? 'reconnecting' : 'live'; this.onUpdate(); } };
    this.inflight = false;
    const resume = this.acked || this.pending.length ? { acked: this.acked, pending: this.pending.length ? [{ id: this.pending[this.pending.length - 1]!.id, base: this.acked, ops: this.pending.flatMap((p) => p.ops) }] : [] } : undefined;
    if (this.pending.length > 1) this.pending = [{ ...resume!.pending[0]! }];
    link.send({ t: 'hello', proto: PROTOCOL, id: this.me.id, name: this.me.name, resume });
    this.onUpdate();
  }

  private loadBook(book: Parameters<typeof decodeBook>[0], name: string): void {
    const wb = decodeBook(book);
    const keepSheet = this.store.sheet?.cid;
    this.loading = true;
    try { this.store.load(wb, name); this.wbRef = this.store.wb; } finally { this.loading = false; }
    const i = keepSheet ? wb.sheets.findIndex((s) => s.cid === keepSheet) : -1;
    if (i >= 0 && i !== this.store.sheetIndex) this.store.setActiveSheet(i);
    this.wbRef = this.store.wb;
  }

  private guestReceive(m: Msg): void {
    switch (m.t) {
      case 'welcome':
        this.me = { ...this.me, color: m.color, ro: m.ro || undefined };
        this.loadBook(m.book, m.name);
        this.store.readOnly = m.ro;
        this.acked = m.seq;
        this.pending = [];
        this.inflight = false;
        this.roster = m.roster;
        this.status = 'live';
        this.lastPresence = '';
        this.sendPresenceNow();
        this.onUpdate();
        break;
      case 'snap':
        this.loadBook(m.book, m.name);
        this.acked = m.seq;
        this.pending = [];
        this.inflight = false;
        this.onUpdate();
        break;
      case 'op': {
        if (m.seq === undefined || m.origin === undefined) return;
        if (m.seq <= this.acked) return;
        if (m.seq !== this.acked + 1) { this.requestFullSync(); return; }
        if (m.origin === this.me.id) {
          if (this.pending[0]?.id === m.id) { this.pending.shift(); this.inflight = false; }
          else if (this.pending.some((p) => p.id === m.id)) { this.acked = m.seq; this.requestFullSync(); return; }
          this.acked = m.seq;
          this.flush();
          return;
        }
        let remote = m.ops;
        for (const p of this.pending) { const [p2, r2] = xformLists(p.ops, remote, true); p.ops = p2; remote = r2; }
        this.acked = m.seq;
        this.applyIncoming(remote);
        break;
      }
      case 'presence':
        if (m.p.id === this.me.id) return;
        this.presence.set(m.p.id, m.p);
        this.followUpdate(m.p.id);
        this.onUpdate();
        break;
      case 'roster':
        this.roster = m.peers;
        for (const id of [...this.presence.keys()]) if (!m.peers.some((p) => p.id === id && !p.away)) this.presence.delete(id);
        if (this.following && !m.peers.some((p) => p.id === this.following)) this.following = null;
        { const mine = m.peers.find((p) => p.id === this.me.id); if (mine) this.me = { ...this.me, color: mine.color }; }
        this.onUpdate();
        break;
      case 'hash':
        if (m.seq === this.acked && !this.pending.length && !this.store.editing && bookHash(this.store.wb) !== m.h) this.requestFullSync();
        break;
      case 'end':
        this.onNotice(T('The host ended the session. You can keep working on your copy.'));
        this.shutdown();
        break;
      case 'ping': this.host?.send({ t: 'pong' }); break;
      default: break;
    }
  }

  // ---------- Presence and follow ----------

  private presenceNow(): Presence {
    const s = this.store;
    const e = s.editing;
    return { id: this.me.id, sheet: s.sheet.cid ?? '', ranges: rangesOf(s), active: [s.sel.active.row, s.sel.active.col], editing: e ? [e.row, e.col] : null };
  }

  private sendPresenceNow(force = false): void {
    if (this.status === 'ended') return;
    const p = this.presenceNow();
    if (!p.sheet) return;
    const key = JSON.stringify(p);
    if (!force && key === this.lastPresence) return;
    this.lastPresence = key;
    const m: Msg = { t: 'presence', p };
    if (this.role === 'host') this.broadcast(m);
    else if (this.host && this.status === 'live') this.host.send(m);
  }

  /** Called on every store update: presence, follow, and noticing another workbook being opened. */
  private watch(): void {
    if (this.status === 'ended' || this.loading) return;
    if (this.store.wb !== this.wbRef) {
      this.wbRef = this.store.wb;
      if (this.role === 'host') { this.store.readOnly = false; identifySheets(this.store); this.broadcastSnapshot(); }
      else if (!this.store.remote) { this.onNotice(T('You opened another workbook, so you left the live session.')); this.leave(); return; }
    }
    if (this.following) {
      const cur = `${this.store.sheetIndex}|${JSON.stringify(this.store.sel.ranges)}`;
      if (this.followSel && cur !== this.followSel) { this.following = null; this.onUpdate(); }
    }
    if (this.presenceTimer) return;
    this.presenceTimer = setTimeout(() => { this.presenceTimer = null; this.sendPresenceNow(); }, 50);
  }

  follow(id: string | null): void {
    this.following = id;
    this.followSel = '';
    if (id) this.followUpdate(id, true);
    this.onUpdate();
  }

  /** Jumps to where a person is, once. */
  goTo(id: string): void {
    const p = this.presence.get(id);
    if (!p) return;
    const i = this.store.wb.sheets.findIndex((s) => s.cid === p.sheet);
    if (i < 0) return;
    this.store.commitEdit();
    if (i !== this.store.sheetIndex) this.store.setActiveSheet(i);
    const last = p.ranges[p.ranges.length - 1]!;
    this.store.selectRange({ r1: last[0], c1: last[1], r2: last[2], c2: last[3] }, { row: p.active[0], col: p.active[1] });
  }

  private followUpdate(id: string, force = false): void {
    if (this.following !== id && !force) return;
    this.goTo(id);
    this.followSel = `${this.store.sheetIndex}|${JSON.stringify(this.store.sel.ranges)}`;
  }

  // ---------- Life cycle ----------

  private tick(): void {
    const now = Date.now();
    if (this.role === 'host') {
      for (const g of this.guests.values()) {
        if (!g.link) continue;
        if (now - g.lastSeen > DEAD_MS) { const l = g.link; l.onClose?.(); l.onClose = null; l.close(); continue; }
        g.link.send({ t: 'ping' });
      }
    } else if (this.host) {
      if (now - this.hostSeen > DEAD_MS) { const l = this.host; l.onClose?.(); l.onClose = null; l.close(); return; }
      this.host.send({ t: 'ping' });
    }
  }

  /** Guest: leave the session. Host: end it for everyone. */
  leave(): void {
    if (this.status === 'ended') return;
    if (this.role === 'host') { this.broadcast({ t: 'end' }); for (const g of this.guests.values()) { if (g.link) { g.link.onClose = null; g.link.close(); } } }
    else if (this.host) { this.host.send({ t: 'bye' }); this.host.onClose = null; this.host.close(); }
    this.shutdown();
  }

  private shutdown(): void {
    this.status = 'ended';
    for (const t of this.timers) clearInterval(t);
    if (this.presenceTimer) clearTimeout(this.presenceTimer);
    this.stopObserve();
    this.stopSub();
    if (this.host) { this.host.onClose = null; this.host.close(); this.host = null; }
    this.store.readOnly = false;
    this.presence.clear();
    this.following = null;
    this.onUpdate();
  }
}
