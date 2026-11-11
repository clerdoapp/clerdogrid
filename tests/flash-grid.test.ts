import { beforeAll, describe, expect, it } from 'vitest';
import { GridStore } from '../src/app/store';
import { registerAllFunctions } from '../src/engine/functions';
import { flashFill } from '../src/app/ops/fill';
import { planFlashFill } from '../src/app/ops/flash';
import { accept, ghost, watch } from '../src/features/flashfill/ghost';
import { parseAddr } from '../src/model/workbook';

beforeAll(() => registerAllFunctions());

const setup = (rows: string[][]) => {
  const s = new GridStore();
  rows.forEach((r, i) => r.forEach((v, j) => { if (v !== '') s.enter(i, j, v); }));
  s.engine.recalc();
  const type = (a: string, text: string) => { const p = parseAddr(a)!; s.select({ row: p.row, col: p.col }); s.beginEdit(text); expect(s.commitEdit()).toBe(true); };
  const val = (a: string) => { const p = parseAddr(a)!; return s.value(p.row, p.col); };
  return { s, type, val };
};

describe('Flash Fill on the grid', () => {
  it('Ctrl+E fills the column in one undoable step', () => {
    const { s, type, val } = setup([['Ann Lee'], ['Bob Stone'], ['Cy Twombly'], ['Di Wu']]);
    type('B1', 'Lee');
    s.select({ row: 1, col: 1 });
    expect(flashFill(s)).toBe(true);
    expect([val('B2'), val('B3'), val('B4')]).toEqual(['Stone', 'Twombly', 'Wu']);
    s.undo();
    expect(val('B2')).toBeNull();
    expect(val('B1')).toBe('Lee');
  });
  it('uses dates as they are shown, and keeps numbers as numbers', () => {
    const { s, type, val } = setup([['3/5/2024', 'Order 17'], ['12/25/2023', 'Order 204'], ['7/14/2022', 'Order 3']]);
    type('C1', '2024');
    type('C2', '2023');
    s.select({ row: 2, col: 2 });
    expect(flashFill(s)).toBe(true);
    expect(val('C3')).toBe(2022);
  });
  it('explains when there is nothing to learn from', () => {
    const { s } = setup([['a'], ['b']]);
    expect(planFlashFill(s, 0, 1)).toMatch(/example/);
  });
  it('suggests after two examples and Enter accepts', () => {
    const { s, type, val } = setup([['john.smith@x.com'], ['mary.jones@y.org'], ['bob.lee@z.net'], ['ann.ray@q.io']]);
    const stop = watch(s);
    try {
      type('B1', 'john');
      expect(ghost.plan).toBeNull();
      type('B2', 'mary');
      expect(ghost.plan).not.toBeNull();
      expect([...ghost.plan!.fills.values()]).toEqual(['bob', 'ann']);
      expect(val('B3')).toBeNull();
      expect(accept(s)).toBe(true);
      expect([val('B3'), val('B4')]).toEqual(['bob', 'ann']);
      s.undo();
      expect(val('B3')).toBeNull();
    } finally { stop(); }
  });
  it('respects the automatic suggestion setting', () => {
    const g = globalThis as unknown as { localStorage?: unknown };
    const old = g.localStorage;
    g.localStorage = { getItem: (k: string) => (k === 'clerdogrid.flashAuto' ? '0' : null), setItem: () => undefined, removeItem: () => undefined };
    const { s, type } = setup([['a b'], ['c d'], ['e f']]);
    const stop = watch(s);
    try {
      type('B1', 'b');
      type('B2', 'd');
      expect(ghost.plan).toBeNull();
    } finally { stop(); g.localStorage = old; ghost.plan = null; }
  });
});
