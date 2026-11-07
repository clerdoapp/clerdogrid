import { MAX_COLS, MAX_ROWS } from '../engine/types';
import { parseRange, rangeAddr } from '../model/workbook';
import type { GridStore } from './store';

/** Name box / Go To: jumps to an address, range, row/column span or defined name; an unknown valid name defines it. */
export const gotoReference = (store: GridStore, t: string): boolean => {
  const v = t.trim();
  if (!v) return false;
  const named = store.wb.names.find((n) => n.name.toLowerCase() === v.toLowerCase());
  if (named) {
    const m = /^(?:'?(.+?)'?!)?(.+)$/.exec(named.formula.replace(/\$/g, ''));
    const si = m?.[1] ? store.engine.sheetIndex(m[1]) : store.sheetIndex;
    const rr = parseRange(m?.[2] ?? '');
    if (rr) { if (si >= 0) store.setActiveSheet(si); store.selectRange(rr); return true; }
  }
  const m = /^(?:'?(.+?)'?!)?(.+)$/.exec(v);
  const target = m?.[2]?.replace(/\$/g, '').toUpperCase() ?? '';
  let rr = parseRange(target);
  if (!rr && /^[A-Z]{1,3}:[A-Z]{1,3}$/.test(target)) { const [a, b] = target.split(':'); rr = { r1: 0, r2: MAX_ROWS - 1, c1: parseRange(`${a}1`)!.c1, c2: parseRange(`${b}1`)!.c1 }; }
  if (!rr && /^\d+:\d+$/.test(target)) { const [a, b] = target.split(':').map(Number); rr = { r1: a! - 1, r2: b! - 1, c1: 0, c2: MAX_COLS - 1 }; }
  if (rr) {
    if (m?.[1]) { const si = store.engine.sheetIndex(m[1]); if (si < 0) { store.notify(`There's no sheet called ${m[1]}.`); return false; } store.setActiveSheet(si); }
    store.selectRange(rr);
    return true;
  }
  // Typing a new name defines it for the selection, as in other spreadsheets.
  if (/^[A-Za-z_\\][A-Za-z0-9_.]{0,254}$/.test(v) && !parseRange(v)) {
    store.change('Define name', 'book', () => {
      store.wb.names.push({ name: v, formula: `${/^[A-Za-z_][A-Za-z0-9_]*$/.test(store.sheet.name) ? store.sheet.name : `'${store.sheet.name.replace(/'/g, "''")}'`}!${rangeAddr(store.range).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')}` });
      store.engine.sheetsChanged();
      store.engine.recalc(true);
    });
    store.notify(`Named ${rangeAddr(store.range)} "${v}".`);
    return true;
  }
  store.notify('Enter a cell address like B7, a range like A1:C9, or a valid name.');
  return false;
};

