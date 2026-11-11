import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { GridStore } from '../src/app/store';
import { registerAllFunctions } from '../src/engine/functions';
import { addSheet, changeStructure, deleteSheet, moveSheet, renameSheet } from '../src/app/ops/structure';
import { CellError } from '../src/engine/types';
import { cellKey, parseAddr, parseRange } from '../src/model/workbook';
import { Session, type Link } from '../src/features/collab/session';
import { bookHash, decodeBook, encodeBook } from '../src/features/collab/wire';
import * as V from '../src/features/collab/validate';
import { xform, xformLists, shiftRange } from '../src/features/collab/ot';
import type { Op } from '../src/features/collab/ops';
import type { Msg } from '../src/features/collab/validate';

beforeAll(() => registerAllFunctions());

// ---------- An in-memory network with manual delivery, so tests control what is concurrent ----------

class Net {
  q: (() => void)[] = [];
  pump(max = 100000) { let n = 0; while (this.q.length && n++ < max) this.q.shift()!(); }
  /** Delivers queued messages in random order per link direction (order within one direction is kept). */
  pair(): [Link, Link] {
    const mk = (): Link & { peer?: Link; open: boolean } => ({ open: true, onMessage: null, onClose: null, send(m: Msg) { if (!this.open) return; const s = JSON.stringify(m); const to = this.peer!; net().q.push(() => { if ((to as unknown as { open: boolean }).open) to.onMessage?.(JSON.parse(s)); }); }, close() { if (!this.open) return; this.open = false; const to = this.peer! as Link & { open: boolean }; net().q.push(() => { if (to.open) { to.open = false; to.onClose?.(); } }); } });
    const net = () => this;
    const a = mk(), b = mk();
    a.peer = b; b.peer = a;
    return [a, b];
  }
}

const sessions: Session[] = [];
afterEach(() => { for (const s of sessions.splice(0)) s.leave(); });

const helpers = (s: GridStore) => ({
  type: (a: string, text: string) => { const p = parseAddr(a)!; s.select({ row: p.row, col: p.col }); s.beginEdit(text); expect(s.commitEdit()).toBe(true); },
  val: (a: string) => { const p = parseAddr(a)!; const v = s.value(p.row, p.col); return v instanceof CellError ? v.code : v; },
  f: (a: string) => { const p = parseAddr(a)!; return s.cell(p.row, p.col)?.f; },
});

const setup = (guests = 1, ro = false) => {
  const net = new Net();
  const hostStore = new GridStore();
  const host = new Session(hostStore, 'host', { id: 'host1', name: 'Hana Host', color: '#e5484d' });
  sessions.push(host);
  const gs: { store: GridStore; session: Session }[] = [];
  for (let i = 0; i < guests; i++) {
    const store = new GridStore();
    const session = new Session(store, 'guest', { id: `guest${i + 1}`, name: `Gus ${i + 1}`, color: '#000000' });
    sessions.push(session);
    const [h, g] = net.pair();
    host.addGuestLink(h, ro);
    session.attachHost(g);
    net.pump();
    gs.push({ store, session });
  }
  return { net, hostStore, host, guests: gs };
};

const same = (...stores: GridStore[]) => {
  const h = stores.map((s) => bookHash(s.wb));
  for (const x of h) expect(x).toBe(h[0]);
};

describe('wire format', () => {
  it('round-trips a workbook with styles, merges, notes, comments, charts and pictures', () => {
    const s = new GridStore();
    const { type } = helpers(s);
    type('A1', '12.5%'); type('B2', '=A1*2'); type('C3', 'hello');
    s.selectRange(parseRange('A1:B1')!); s.applyStyle({ font: { bold: true }, fill: { color: '#ffee00' } });
    s.sheet.merges.push(parseRange('D1:E2')!);
    s.sheet.notes.set(cellKey(0, 0), { author: 'x', text: 'note' });
    s.sheet.comments = new Map([[cellKey(1, 1), { id: 'c1', author: 'a', text: 'hi', time: '2026-01-01T00:00:00Z', replies: [] }]]);
    s.sheet.pictures.push({ id: 'p1', data: new Uint8Array([1, 2, 3, 250]), mime: 'image/png', anchor: { row: 1, col: 1, dx: 0, dy: 0, width: 10, height: 10 } });
    s.sheet.rows.set(3, { height: 30, s: s.wb.styles.length - 1 });
    const enc = JSON.parse(JSON.stringify(encodeBook(s.wb)));
    V.wbook(enc);
    const wb = decodeBook(enc);
    expect(bookHash(wb)).toBe(bookHash(s.wb));
    expect(wb.sheets[0]!.pictures[0]!.data).toEqual(new Uint8Array([1, 2, 3, 250]));
  });
});

describe('sessions', () => {
  it('a guest gets the workbook on join and edits sync both ways', () => {
    const net = new Net();
    const hs = new GridStore();
    helpers(hs).type('A1', 'before join');
    const host = new Session(hs, 'host', { id: 'h', name: 'H', color: '#e5484d' });
    const gs = new GridStore();
    const guest = new Session(gs, 'guest', { id: 'g', name: 'G', color: '#000000' });
    sessions.push(host, guest);
    const [a, b] = net.pair();
    host.addGuestLink(a); guest.attachHost(b); net.pump();
    expect(guest.status).toBe('live');
    expect(helpers(gs).val('A1')).toBe('before join');
    helpers(hs).type('B1', '5'); net.pump();
    helpers(gs).type('B2', '=B1*3'); net.pump();
    expect(helpers(hs).val('B2')).toBe(15);
    expect(helpers(gs).val('B2')).toBe(15);
    expect(host.roster.map((p) => p.name)).toEqual(['H', 'G']);
    expect(guest.roster.length).toBe(2);
    same(hs, gs);
  });

  it('remote edits are not on the local undo stack and recalc happens', () => {
    const { net, hostStore, guests: [g] } = setup();
    helpers(hostStore).type('A1', '2');
    helpers(hostStore).type('A2', '=A1+1');
    net.pump();
    expect(g!.store.canUndo).toBe(false);
    helpers(g!.store).type('A1', '10');
    net.pump();
    expect(helpers(hostStore).val('A2')).toBe(11);
    expect(hostStore.canUndo).toBe(true);
    hostStore.undo(); // undoes the host's own "=A1+1", not the guest's 10
    net.pump();
    expect(helpers(hostStore).val('A1')).toBe(10);
    expect(helpers(g!.store).val('A2')).toBe(null);
    same(hostStore, g!.store);
  });

  it("undo doesn't revert other people's later edits", () => {
    const { net, hostStore, guests: [g] } = setup();
    const gst = g!.store;
    gst.selectRange(parseRange('A1:A3')!);
    gst.beginEdit('mine'); gst.commitEdit(true);
    net.pump();
    helpers(hostStore).type('A2', 'theirs');
    net.pump();
    gst.undo();
    net.pump();
    expect(helpers(gst).val('A1')).toBe(null);
    expect(helpers(gst).val('A2')).toBe('theirs');
    expect(helpers(hostStore).val('A3')).toBe(null);
    same(hostStore, gst);
  });

  it('undo past your own row insert keeps others\' later edits in the right place', () => {
    const { net, hostStore, guests: [g] } = setup();
    const gs = g!.store;
    helpers(gs).type('A1', 'top');
    gs.select({ row: 0, col: 0 });
    changeStructure(gs, 'row', 0, 1); // "top" moves to A2
    helpers(gs).type('B5', 'mine');
    net.pump();
    helpers(hostStore).type('C3', 'theirs');
    net.pump();
    gs.undo(); // removes "mine"
    expect(helpers(gs).val('B5')).toBe(null);
    gs.undo(); // removes the inserted row: "theirs" moves up with it
    expect(helpers(gs).val('A1')).toBe('top');
    expect(helpers(gs).val('C2')).toBe('theirs');
    gs.undo(); // removes "top"
    expect(helpers(gs).val('A1')).toBe(null);
    expect(helpers(gs).val('C2')).toBe('theirs');
    expect(gs.canUndo).toBe(false);
    net.pump();
    same(hostStore, gs);
    // Redo goes forward through the insert again, keeping "theirs".
    gs.redo(); gs.redo();
    expect(helpers(gs).val('A2')).toBe('top');
    expect(helpers(gs).val('C3')).toBe('theirs');
    net.pump();
    same(hostStore, gs);
  });

  it('undo after someone else inserts or deletes rows lands on the moved cells', () => {
    const { net, hostStore, guests: [g] } = setup();
    const gs = g!.store;
    for (let r = 1; r <= 8; r++) helpers(hostStore).type(`A${r}`, `r${r}`);
    net.pump();
    helpers(gs).type('B5', 'mine'); // next to r5
    gs.selectRange({ r1: 0, c1: 3, r2: 1048575, c2: 3 }); gs.applyStyle({ font: { bold: true } }); // whole column D: a sheet snapshot
    helpers(gs).type('C1', '=A5'); // a formula pointing at r5
    net.pump();
    hostStore.select({ row: 0, col: 0 });
    changeStructure(hostStore, 'row', 1, 2); // r5 is now on row 7
    net.pump();
    hostStore.select({ row: 0, col: 0 });
    changeStructure(hostStore, 'row', 7, -1); // delete r6 (row 8)
    net.pump();
    expect(helpers(gs).val('B7')).toBe('mine');
    expect(helpers(gs).f('C1')).toBe('A7');
    gs.undo(); // the formula
    expect(helpers(gs).val('C1')).toBe(null);
    gs.undo(); // the bold column
    expect(gs.style(6, 3).font?.bold).toBeFalsy();
    expect(helpers(gs).val('A7')).toBe('r5');
    expect(helpers(gs).val('A8')).toBe('r7');
    gs.undo(); // "mine", now on row 7
    expect(helpers(gs).val('B7')).toBe(null);
    expect(helpers(gs).val('B5')).toBe(null);
    net.pump();
    same(hostStore, gs);
  });

  it('undo of your row delete brings the rows back with others\' edits around them', () => {
    const { net, hostStore, guests: [g] } = setup();
    const gs = g!.store;
    for (let r = 1; r <= 6; r++) helpers(hostStore).type(`A${r}`, `r${r}`);
    net.pump();
    gs.select({ row: 0, col: 0 });
    changeStructure(gs, 'row', 1, -2); // delete r2 and r3
    net.pump();
    helpers(hostStore).type('B3', 'theirs'); // next to r5 (row 3 now)
    hostStore.select({ row: 0, col: 0 });
    changeStructure(hostStore, 'row', 0, 1); // everything moves down one
    net.pump();
    gs.undo();
    const col = (s: GridStore, c: string) => Array.from({ length: 8 }, (_, i) => helpers(s).val(`${c}${i + 1}`));
    expect(col(gs, 'A')).toEqual([null, 'r1', 'r2', 'r3', 'r4', 'r5', 'r6', null]);
    expect(helpers(gs).val('B6')).toBe('theirs');
    net.pump();
    same(hostStore, gs);
  });

  it('concurrent insert row and cell edit land on the same data', () => {
    const { net, hostStore, guests: [g] } = setup();
    for (let r = 1; r <= 6; r++) helpers(hostStore).type(`A${r}`, `r${r}`);
    net.pump();
    // Concurrent: host inserts two rows above row 3 while the guest edits A5 and types a formula pointing at A5.
    hostStore.select({ row: 2, col: 0 });
    changeStructure(hostStore, 'row', 2, 2);
    helpers(g!.store).type('B5', 'next to r5');
    helpers(g!.store).type('C1', '=A5');
    net.pump();
    for (const s of [hostStore, g!.store]) {
      expect(helpers(s).val('A7')).toBe('r5');
      expect(helpers(s).val('B7')).toBe('next to r5');
      expect(helpers(s).f('C1')).toBe('A7');
      expect(helpers(s).val('C1')).toBe('r5');
    }
    same(hostStore, g!.store);
  });

  it('concurrent deletes and inserts on both sides converge', () => {
    const { net, hostStore, guests: [g] } = setup();
    for (let r = 1; r <= 10; r++) helpers(hostStore).type(`A${r}`, `r${r}`);
    net.pump();
    hostStore.select({ row: 0, col: 0 });
    changeStructure(hostStore, 'row', 2, -4); // delete rows 3-6
    changeStructure(g!.store, 'row', 4, 2); // insert 2 rows inside that block
    helpers(g!.store).type('A5', 'new');
    net.pump();
    const col = (s: GridStore) => Array.from({ length: 10 }, (_, i) => helpers(s).val(`A${i + 1}`));
    expect(col(g!.store)).toEqual(['r1', 'r2', 'new', null, 'r7', 'r8', 'r9', 'r10', null, null]);
    expect(col(hostStore)).toEqual(['r1', 'r2', 'new', null, 'r7', 'r8', 'r9', 'r10', null, null]);
    same(hostStore, g!.store);
  });

  it('a later edit of the same cell wins on every copy', () => {
    const { net, hostStore, guests: [g] } = setup();
    helpers(g!.store).type('A1', 'guest');
    helpers(hostStore).type('A1', 'host');
    net.pump();
    same(hostStore, g!.store);
    expect(helpers(hostStore).val('A1')).toBe('guest'); // the host ordered the guest's edit after its own
  });

  it('syncs sheets, styles, merges, comments, charts and names', () => {
    const { net, hostStore, guests: [g] } = setup();
    const gs = g!.store;
    addSheet(hostStore);
    net.pump();
    expect(gs.wb.sheets.map((s) => s.name)).toEqual(['Sheet1', 'Sheet2']);
    renameSheet(gs, 1, 'Data');
    helpers(hostStore).type('A1', '=Sheet1!B1'); // host is on Sheet2 (now Data after sync)
    net.pump();
    same(hostStore, gs);
    expect(hostStore.wb.sheets[1]!.name).toBe('Data');
    moveSheet(hostStore, 1, 0);
    net.pump();
    expect(gs.wb.sheets.map((s) => s.name)).toEqual(['Data', 'Sheet1']);
    gs.setActiveSheet(1);
    gs.selectRange(parseRange('B2:C3')!);
    gs.applyStyle({ font: { italic: true } });
    gs.change('Merge', 'sheet', () => { gs.sheet.merges.push(parseRange('B2:C3')!); });
    gs.change('Comment', 'sheet', () => { gs.sheet.comments = new Map([[cellKey(1, 1), { id: 'c9', author: 'Gus', text: 'look', time: '2026-11-08T00:00:00Z', replies: [] }]]); });
    gs.change('Chart', 'sheet', () => { gs.sheet.charts.push({ id: 'ch1', type: 'column', series: [{ values: 'Sheet1!B2:B3' }], anchor: { row: 4, col: 4, dx: 0, dy: 0, width: 300, height: 200 } }); });
    gs.change('Name', 'book', () => { gs.wb.names.push({ name: 'Total', formula: 'Sheet1!$B$2' }); });
    net.pump();
    const hs1 = hostStore.wb.sheets.find((s) => s.name === 'Sheet1')!;
    expect(hs1.merges).toEqual([parseRange('B2:C3')]);
    expect(hs1.comments?.get(cellKey(1, 1))?.text).toBe('look');
    expect(hs1.charts[0]!.id).toBe('ch1');
    expect(hostStore.wb.styles[hs1.cells.get(cellKey(1, 1))!.s!]!.font?.italic).toBe(true);
    expect(hostStore.wb.names[0]!.name).toBe('Total');
    same(hostStore, gs);
    deleteSheet(hostStore, 0);
    net.pump();
    expect(gs.wb.sheets.map((s) => s.name)).toEqual(['Sheet1']);
    same(hostStore, gs);
  });

  it('relays between three people and shows presence', () => {
    const { net, hostStore, host, guests } = setup(2);
    const [g1, g2] = guests;
    helpers(g1!.store).type('A1', 'from g1');
    net.pump();
    expect(helpers(g2!.store).val('A1')).toBe('from g1');
    g2!.store.selectRange(parseRange('C3:D4')!);
    // presence goes out on a short timer
    return new Promise<void>((done) => setTimeout(() => {
      net.pump();
      const p = g1!.session.presence.get('guest2');
      expect(p?.ranges[0]).toEqual([2, 2, 3, 3]);
      expect(host.presence.get('guest2')?.active).toEqual([2, 2]);
      expect(g1!.session.roster.map((x) => x.id)).toEqual(['host1', 'guest1', 'guest2']);
      same(hostStore, g1!.store, g2!.store);
      done();
    }, 80));
  });

  it('remote cursors move with inserted rows', async () => {
    const { net, hostStore, host, guests: [g] } = setup();
    g!.store.select({ row: 5, col: 1 });
    await new Promise((r) => setTimeout(r, 80));
    net.pump();
    expect(host.presence.get('guest1')?.active).toEqual([5, 1]);
    hostStore.select({ row: 0, col: 0 });
    changeStructure(hostStore, 'row', 1, 3);
    expect(host.presence.get('guest1')?.active).toEqual([8, 1]);
    net.pump();
    expect(g!.store.sel.active).toEqual({ row: 8, col: 1 });
  });

  it('view-only guests can not change the workbook', () => {
    const { net, hostStore, guests: [g] } = setup(1, true);
    expect(g!.store.readOnly).toBe(true);
    g!.store.select({ row: 0, col: 0 });
    g!.store.beginEdit('x');
    expect(g!.store.editing).toBe(null);
    net.pump();
    expect(helpers(hostStore).val('A1')).toBe(null);
  });

  it('reconnects and sends changes made while offline', () => {
    const { net, hostStore, host, guests: [g] } = setup();
    const [h1, g1] = net.pair();
    void h1; void g1;
    // Drop the link.
    (host as unknown as { guests: Map<string, { link: Link }> }).guests.get('guest1')!.link.close();
    net.pump();
    expect(g!.session.status).toBe('lost');
    helpers(g!.store).type('A1', 'offline edit');
    helpers(hostStore).type('B1', 'host edit');
    const [h2, g2] = net.pair();
    host.addGuestLink(h2);
    g!.session.attachHost(g2);
    net.pump();
    expect(g!.session.status).toBe('live');
    expect(helpers(hostStore).val('A1')).toBe('offline edit');
    expect(helpers(g!.store).val('B1')).toBe('host edit');
    same(hostStore, g!.store);
  });

  it('heals a diverged copy from the host hash check', async () => {
    const { net, hostStore, guests: [g] } = setup();
    helpers(hostStore).type('A1', '1');
    net.pump();
    g!.store.sheet.cells.set(cellKey(9, 9), { v: 'drift' }); // a change that bypassed change()
    (g!.session as unknown as { lastResync: number }).lastResync = 0;
    (hostStore as unknown as { emit: () => void }).emit();
    await new Promise((r) => setTimeout(r, 2700));
    net.pump();
    net.pump();
    expect(g!.store.sheet.cells.has(cellKey(9, 9))).toBe(false);
    same(hostStore, g!.store);
  }, 6000);

  /** Seeded ids for new sheets, so a failing run can be replayed. */
  const seedIds = (seed: number) => {
    let x = seed * 7919 + 1;
    return vi.spyOn(globalThis.crypto, 'getRandomValues').mockImplementation(<T extends ArrayBufferView | null>(b: T): T => {
      const a = b as unknown as Uint8Array;
      for (let i = 0; i < a.length; i++) { x = (x * 1103515245 + 12345) & 0x7fffffff; a[i] = x >> 16; }
      return b;
    });
  };

  const randomRun = (sheetOps: boolean, rounds: number, steps: number, seed0: number) => {
    const failed: number[] = [];
    for (let round = 0; round < rounds; round++) {
      let seed = seed0 + round;
      const rnd = (n: number) => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed % n; };
      const ids = seedIds(seed0 + round);
      const { net, hostStore, guests } = setup(2);
      const stores = [hostStore, ...guests.map((x) => x.store)];
      for (let step = 0; step < steps; step++) {
        const s = stores[rnd(stores.length)]!;
        const k = rnd(sheetOps ? 13 : 10);
        if (k >= 10) {
          const j = rnd(5);
          if (j === 0) addSheet(s);
          else if (j === 1) renameSheet(s, s.sheetIndex, `N${rnd(1000)}`);
          else if (j === 2 && s.wb.sheets.length > 1) deleteSheet(s);
          else if (j === 3 && s.wb.sheets.length > 1) moveSheet(s, rnd(s.wb.sheets.length), rnd(s.wb.sheets.length));
          else if (s.wb.sheets.length > 1) s.setActiveSheet(rnd(s.wb.sheets.length));
        } else if (k < 5) { s.select({ row: rnd(12), col: rnd(4) }); s.beginEdit(rnd(3) ? String(rnd(100)) : `=A${rnd(12) + 1}+1`); s.commitEdit(); }
        else if (k < 7) { s.select({ row: 0, col: 0 }); changeStructure(s, rnd(4) ? 'row' : 'col', rnd(10), rnd(2) ? 1 + rnd(2) : -(1 + rnd(2))); }
        else if (k < 8) { s.selectRange({ r1: rnd(6), c1: rnd(3), r2: 6 + rnd(4), c2: 3 }); s.applyStyle({ font: { bold: !!rnd(2) } }); }
        else if (k < 9) { s.selectRange({ r1: rnd(8), c1: 0, r2: 9, c2: 2 }); s.change('Clear', [{ area: s.range }], () => { for (let r = s.range.r1; r <= s.range.r2; r++) for (let c = s.range.c1; c <= s.range.c2; c++) s.engine.setValue(s.sheetIndex, r, c, null); }); }
        else if (s.canUndo) s.undo();
        else if (s.canRedo) s.redo();
        if (rnd(3) === 0) net.pump(rnd(6));
      }
      net.pump();
      // No checksum heal here: every copy must already agree.
      const h = stores.map((x) => bookHash(x.wb));
      if (h.some((x) => x !== h[0])) failed.push(seed0 + round);
      for (const x of sessions.splice(0)) x.leave();
      ids.mockRestore();
    }
    expect(failed).toEqual([]);
  };

  it('random concurrent cell, row/column, style, clear and undo edits converge exactly', () => randomRun(false, 120, 60, 7));
  it('random edits with sheet inserts, renames, deletes, moves, undo and redo converge without the checksum (1,500 runs)', () => randomRun(true, 1500, 60, 1), 180_000);
});

describe('transforms', () => {
  const ins = (at: number, n: number): Op => ({ t: 'struct', s: 'S', nm: 'S', a: 'row', at, n });
  it('moves cell edits past inserts and drops edits in deleted rows', () => {
    const cells: Op = { t: 'cells', s: 'S', c: [[4, 0, { v: 'x' }], [1, 0, { v: 'y' }]] };
    expect(xform(cells, ins(2, 3), true)).toEqual([{ t: 'cells', s: 'S', c: [[7, 0, { v: 'x' }], [1, 0, { v: 'y' }]] }]);
    expect(xform(cells, ins(3, -2), true)).toEqual([{ t: 'cells', s: 'S', c: [[1, 0, { v: 'y' }]] }]);
    expect(xform(cells, { t: 'struct', s: 'Other', nm: 'Other', a: 'row', at: 0, n: 5 }, true)).toEqual([cells]);
  });
  it('orders inserts at the same place by host order', () => {
    expect(xform(ins(2, 1), ins(2, 3), true)).toEqual([ins(5, 1)]);
    expect(xform(ins(2, 1), ins(2, 3), false)).toEqual([ins(2, 1)]);
  });
  it('splits a delete around an insert inside it', () => {
    expect(xform(ins(2, -4), ins(4, 2), true)).toEqual([ins(6, -2), ins(2, -2)]);
    expect(xform(ins(2, -4), ins(3, -5), true)).toEqual([ins(2, -1)]);
    expect(xform(ins(2, -2), ins(1, -5), true)).toEqual([]);
  });
  it('drops the earlier of two edits to the same cell', () => {
    const a: Op = { t: 'cells', s: 'S', c: [[0, 0, { v: 1 }], [0, 1, { v: 2 }]] };
    const b: Op = { t: 'cells', s: 'S', c: [[0, 0, { v: 3 }]] };
    expect(xform(a, b, false)).toEqual([{ t: 'cells', s: 'S', c: [[0, 1, { v: 2 }]] }]);
    expect(xform(a, b, true)).toEqual([a]);
  });
  it('gives clashing sheet names a suffix on every copy', () => {
    const a: Op = { t: 'sheetRen', s: 'sAAAA1111', from: 'X', to: 'Data' };
    const b: Op = { t: 'sheetRen', s: 'sBBBB2222', from: 'Y', to: 'data' };
    expect(xform(a, b, true)).toEqual([{ ...a, to: 'Data (AAAA)' }]);
    expect(xform(b, a, false)).toEqual([{ t: 'sheetRen', s: 'sAAAA1111', from: 'Data', to: 'Data (AAAA)' }, b]);
    // A delete and a re-add (undo) of the same sheet: the later one wins.
    const del: Op = { t: 'sheetDel', s: 'sAAAA1111', name: 'X' };
    const add: Op = { t: 'sheetAdd', s: 'sAAAA1111', name: 'X', pos: 'i', d: { name: 'X', cells: [], rows: [], cols: [], parts: {} } };
    expect(xform(add, del, true)).toEqual([add]);
    expect(xform(del, add, false)).toEqual([]);
  });
  it('follows sheet renames and deletes', () => {
    const a: Op = { t: 'cells', s: 'S', c: [[0, 0, { f: 'S!A1+T!B2' }]] };
    expect(xform(a, { t: 'sheetRen', s: 'idT', from: 'T', to: 'New' }, true)).toEqual([{ t: 'cells', s: 'S', c: [[0, 0, { f: 'S!A1+New!B2' }]] }]);
    // Edits to a deleted sheet still go to its kept copy (so an undo brings them back on every copy).
    expect(xform(a, { t: 'sheetDel', s: 'S', name: 'S' }, true)).toEqual([{ t: 'cells', s: 'S', c: [[0, 0, { f: '#REF!+T!B2' }]] }]);
    expect(xform(a, { t: 'sheetDel', s: 'idT', name: 'T' }, true)).toEqual([{ t: 'cells', s: 'S', c: [[0, 0, { f: 'S!A1+#REF!' }]] }]);
  });
  it('shifts merges and rules with rows', () => {
    const p: Op = { t: 'prop', s: 'S', k: 'merges', v: [{ r1: 3, c1: 0, r2: 4, c2: 1 }, { r1: 8, c1: 0, r2: 8, c2: 2 }] };
    expect(xform(p, ins(0, 2), true)).toEqual([{ ...p, v: [{ r1: 5, c1: 0, r2: 6, c2: 1 }, { r1: 10, c1: 0, r2: 10, c2: 2 }] }]);
    expect(shiftRange({ r1: 2, c1: 0, r2: 6, c2: 0 }, 'row', 3, -2)).toEqual({ r1: 2, c1: 0, r2: 4, c2: 0 });
  });
  it('transforms lists both ways', () => {
    const [a, b] = xformLists([ins(1, 1)], [ins(0, 1), ins(5, -1)], true);
    expect(a).toEqual([ins(2, 1)]);
    expect(b).toEqual([ins(0, 1), ins(6, -1)]);
  });
});

describe('message checks', () => {
  it('rejects malformed messages and ignores unknown op types', () => {
    expect(() => V.message({ t: 'op', id: 1, base: 0, ops: [{ t: 'cells', s: 'S', c: [[-1, 0, null]] }] })).toThrow();
    expect(() => V.message({ t: 'op', id: 1, base: 0, ops: [{ t: 'cells', s: 'S', c: [[0, 99999, null]] }] })).toThrow();
    expect(() => V.message({ t: 'op', id: 1, base: 0, ops: [{ t: 'cells', s: 'S', c: [[0, 0, { v: 1, evil: 2 }]] }] })).toThrow();
    expect(() => V.message({ t: 'op', id: 1, base: 0, ops: [{ t: 'keyed', s: 'S', k: 'hyperlinks', c: [[0, 0, 'javascript:alert(1)']] }] })).toThrow();
    expect(() => V.message({ t: 'presence', p: { id: 'x', sheet: 'S', ranges: [[5, 0, 1, 0]], active: [0, 0], editing: null } })).toThrow();
    expect(() => V.message({ t: 'hello', proto: 1, id: 'bad id!', name: 'x' })).toThrow();
    expect(() => V.message({ t: 'nonsense' })).toThrow();
    const m = V.message({ t: 'op', id: 1, base: 0, ops: [{ t: 'future', x: 1 }, { t: 'cells', s: 'S', c: [[0, 0, { v: 'ok' }]] }] });
    expect((m as { ops: Op[] }).ops).toHaveLength(1);
    expect(() => V.message({ t: 'op', id: 1, base: 0, ops: [{ t: 'book', k: '__proto__', v: 1 }] })).toThrow();
  });
});
