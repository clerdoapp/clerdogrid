// Sheet identity in a live session. Each sheet has a stable id (so edits follow it through renames and moves) and a
// sort key; the sheet list is always ordered by (key, id). Moving or adding a sheet only sets its own key, so
// concurrent sheet changes can't disagree about the order.
import type { Sheet, Workbook } from '../../model/workbook';

const DIGITS = '0123456789abcdefghijklmnopqrstuvwxyz';

export const newSheetId = (): string => {
  const b = new Uint8Array(8);
  crypto.getRandomValues(b);
  return 's' + Array.from(b, (x) => DIGITS[x % 36]).join('');
};

/** A key strictly between `a` and `b` (null: no bound), as a base-36 fraction. Keys never end in "0". */
export const keyBetween = (a: string | null, b: string | null): string => {
  const lo = a ?? '';
  let hi = b;
  if (hi !== null && hi <= lo) hi = null;
  let out = '';
  for (let i = 0; ; i++) {
    const da = i < lo.length ? DIGITS.indexOf(lo[i]!) : 0;
    const db = hi !== null ? (i < hi.length ? DIGITS.indexOf(hi[i]!) : 0) : 36;
    if (da === db) { out += DIGITS[da]; continue; }
    const mid = Math.floor((da + db) / 2);
    if (mid > da) return out + DIGITS[mid];
    // Adjacent digits: keep the lower one and look for room past it with no upper bound.
    out += DIGITS[da];
    hi = null;
  }
};

export const sheetOrder = (a: Sheet, b: Sheet): number => {
  const pa = a.cpos ?? '', pb = b.cpos ?? '';
  if (pa !== pb) return pa < pb ? -1 : 1;
  const ia = a.cid ?? '', ib = b.cid ?? '';
  return ia < ib ? -1 : ia > ib ? 1 : 0;
};

/** Puts the sheets in key order, keeping defined names pointing at the same sheets. */
export const sortSheets = (wb: Workbook): void => {
  const old = [...wb.sheets];
  wb.sheets.sort(sheetOrder);
  if (old.every((s, i) => wb.sheets[i] === s)) return;
  wb.names = wb.names.map((n) => (n.sheet === undefined || !old[n.sheet] ? n : { ...n, sheet: wb.sheets.indexOf(old[n.sheet]!) }));
};

/**
 * Gives every sheet an id (new or duplicated ones get a fresh id; `keep` wins a duplicate) and sets keys so that
 * the current list order is the key order, changing as few keys as possible. Returns the sheets whose key changed.
 */
export const placeSheets = (list: Sheet[], keep?: Set<Sheet>): Sheet[] => {
  const seen = new Map<string, Sheet>();
  const owners = [...list].sort((a, b) => Number(keep?.has(b) ?? false) - Number(keep?.has(a) ?? false));
  for (const s of owners) {
    if (!s.cid || seen.has(s.cid)) { s.cid = newSheetId(); delete s.cpos; }
    seen.set(s.cid, s);
  }
  // Longest run of sheets already in key order stays; the rest get new keys between their neighbours.
  const n = list.length;
  const len = new Array<number>(n).fill(0), prev = new Array<number>(n).fill(-1);
  let best = -1;
  for (let i = 0; i < n; i++) {
    if (list[i]!.cpos === undefined) continue;
    len[i] = 1;
    for (let j = 0; j < i; j++) if (list[j]!.cpos !== undefined && len[j]! + 1 > len[i]! && list[j]!.cpos! < list[i]!.cpos!) { len[i] = len[j]! + 1; prev[i] = j; }
    if (best < 0 || len[i]! > len[best]!) best = i;
  }
  const fixed = new Set<number>();
  for (let i = best; i >= 0; i = prev[i]!) fixed.add(i);
  const changed: Sheet[] = [];
  for (let i = 0; i < n; i++) {
    if (fixed.has(i)) continue;
    const lo = i > 0 ? list[i - 1]!.cpos ?? null : null;
    let hi: string | null = null;
    for (let j = i + 1; j < n; j++) if (fixed.has(j)) { hi = list[j]!.cpos!; break; }
    list[i]!.cpos = keyBetween(lo, hi);
    changed.push(list[i]!);
  }
  return changed;
};

export const sheetIndexById = (wb: Workbook, id: string): number => wb.sheets.findIndex((s) => s.cid === id);
