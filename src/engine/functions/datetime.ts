import { CellError, Err, isError, type FnContext, type FunctionDef, type Scalar, type Value } from '../types';
import { cellsOf, first, fromRange, int, num, optNum, text, toNumber } from '../coerce';
import { daysInMonth, fromSerial, isLeap, nowSerial, parseDateTime, todaySerial, weekdayOf, type DateParts } from '../date';

const DAY_MS = 86_400_000;
/** Last valid serial: 9999-12-31. */
export const maxSerial = (d1904: boolean): number => (d1904 ? 2_957_003 : 2_958_465);

/** spreadsheet serial for y-m-1 plus (d - 1) days; month and day may overflow either way. Linear in spreadsheet's calendar. */
export const serialOf = (y: number, m: number, d: number, d1904 = false): number => {
  const dt = new Date(0);
  dt.setUTCFullYear(y, m - 1, 1);
  const ms = dt.getTime();
  if (d1904) return Math.round((ms - Date.UTC(1904, 0, 1)) / DAY_MS) + d - 1;
  const days = Math.round((ms - Date.UTC(1899, 11, 31)) / DAY_MS);
  return days + (days >= 60 ? 1 : 0) + d - 1; // the phantom 1900-02-29 is serial 60
};

export const partsOf = (serial: number, d1904 = false): DateParts => fromSerial(serial, d1904);

/** Days in a month in spreadsheet's calendar, where February 1900 has 29 days. */
export const monthDays = (y: number, m: number, d1904 = false): number => (!d1904 && y === 1900 && m === 2 ? 29 : daysInMonth(y, m));

export const isMonthEnd = (serial: number, d1904 = false): boolean => {
  const p = partsOf(serial, d1904);
  return p.day === monthDays(p.year, p.month, d1904);
};

/** Shifts a date by whole months, clamping the day (or pinning to month end when `eom`). */
export const addMonths = (serial: number, months: number, d1904 = false, eom = false): number => {
  const p = partsOf(serial, d1904);
  const t = p.year * 12 + (p.month - 1) + months;
  const y = Math.floor(t / 12), m = (t % 12) + 1;
  const dim = monthDays(y, m, d1904);
  return serialOf(y, m, eom ? dim : Math.min(p.day, dim), d1904);
};

/** Weekday, 0 = Sunday. spreadsheet serials are linear, so this is plain arithmetic (1900-01-01 is a Sunday). */
const wd = (s: number, d1904: boolean): number => (((s + (d1904 ? 5 : 6)) % 7) + 7) % 7;

/** A date argument: coerced (text dates are read), truncated, and within 0..9999-12-31. */
export const toDate = (v: Scalar, d1904: boolean): number | CellError => {
  const n = toNumber(v, d1904);
  if (isError(n)) return n;
  if (n < 0 || n >= maxSerial(d1904) + 1) return Err.num;
  return Math.floor(n);
};
export const dateArg = (v: Value, ctx: FnContext): number | CellError => toDate(first(v, ctx), ctx.date1904);

/** 30/360 as YEARFRAC basis 0 and the coupon functions count it (NASD, with the February rules). */
export const days360US = (s: number, e: number, d1904 = false): number => {
  const a = partsOf(s, d1904), b = partsOf(e, d1904);
  let sd = a.day, ed = b.day;
  const sFebEnd = a.month === 2 && isMonthEnd(s, d1904), eFebEnd = b.month === 2 && isMonthEnd(e, d1904);
  if (sd === 31 && ed === 31) { sd = 30; ed = 30; }
  else if (sd === 31) sd = 30;
  else if (sd === 30 && ed === 31) ed = 30;
  else if (sFebEnd && eFebEnd) { sd = 30; ed = 30; }
  else if (sFebEnd) sd = 30;
  return (b.year - a.year) * 360 + (b.month - a.month) * 30 + ed - sd;
};

export const days360EU = (s: number, e: number, d1904 = false): number => {
  const a = partsOf(s, d1904), b = partsOf(e, d1904);
  return (b.year - a.year) * 360 + (b.month - a.month) * 30 + Math.min(b.day, 30) - Math.min(a.day, 30);
};

/** YEARFRAC, including spreadsheet's basis 1 (actual/actual) rules for spans within and beyond one year. */
export const yearFrac = (start: number, end: number, basis: number, d1904 = false): number => {
  if (start > end) [start, end] = [end, start];
  const days = end - start;
  switch (basis) {
    case 0: return days360US(start, end, d1904) / 360;
    case 2: return days / 360;
    case 3: return days / 365;
    case 4: return days360EU(start, end, d1904) / 360;
  }
  const a = partsOf(start, d1904), b = partsOf(end, d1904);
  const withinYear = a.year === b.year || (a.year + 1 === b.year && (a.month > b.month || (a.month === b.month && a.day >= b.day)));
  if (withinYear) {
    const mar1 = (y: number): number => serialOf(y, 3, 1, d1904);
    const feb29Between = (isLeap(a.year) && start < mar1(a.year) && end >= mar1(a.year)) || (isLeap(b.year) && end >= mar1(b.year) && start < mar1(b.year));
    const leap = (a.year === b.year && isLeap(a.year)) || feb29Between || (b.month === 2 && b.day === 29);
    return days / (leap ? 366 : 365);
  }
  const years = b.year - a.year + 1;
  const avg = (serialOf(b.year + 1, 1, 1, d1904) - serialOf(a.year, 1, 1, d1904)) / years;
  return days / avg;
};

const dateFn = (name: string, f: (d: DateParts, s: number, ctx: FnContext) => number): FunctionDef => ({
  name, min: 1, max: 1, scalar: true,
  fn: ([v], ctx) => { const s = dateArg(v, ctx); return isError(s) ? s : f(partsOf(s, ctx.date1904), s, ctx); },
});

/** Time-of-day in whole seconds; spreadsheets round to the nearest second. */
const secondsOf = (v: Value, ctx: FnContext): number | CellError => {
  const n = num(v, ctx);
  if (isError(n)) return n;
  if (n < 0 || n >= maxSerial(ctx.date1904) + 1) return Err.num;
  return Math.round((n - Math.floor(n)) * 86_400 + 1e-9) % 86_400;
};
const timeFn = (name: string, f: (secs: number) => number): FunctionDef => ({
  name, min: 1, max: 1, scalar: true,
  fn: ([v], ctx) => { const s = secondsOf(v, ctx); return isError(s) ? s : f(s); },
});

const textArg = (v: Value, ctx: FnContext): string | CellError => {
  const s = first(v, ctx);
  return isError(s) ? s : typeof s === 'string' ? s : Err.value;
};

/** Weekend days for the .INTL functions, indexed by weekday (0 = Sunday). */
const weekendMask = (v: Value | undefined, ctx: FnContext, forWorkday: boolean): boolean[] | CellError => {
  const w = v === undefined ? null : first(v, ctx);
  if (isError(w)) return w;
  const mask = Array<boolean>(7).fill(false);
  if (w === null) { mask[0] = mask[6] = true; return mask; }
  if (typeof w === 'string') {
    if (!/^[01]{7}$/.test(w)) return Err.value;
    for (let i = 0; i < 7; i++) mask[(i + 1) % 7] = w[i] === '1'; // string starts on Monday
    return forWorkday && mask.every(Boolean) ? Err.value : mask;
  }
  if (typeof w === 'boolean') return Err.value;
  const k = Math.trunc(w);
  if (k >= 1 && k <= 7) { mask[(k + 5) % 7] = mask[(k + 6) % 7] = true; return mask; }
  if (k >= 11 && k <= 17) { mask[k - 11] = true; return mask; }
  return Err.num;
};

const holidaySet = (v: Value | undefined, ctx: FnContext): Set<number> | CellError => {
  const out = new Set<number>();
  if (v === undefined || v === null) return out;
  const range = fromRange(v);
  for (const c of cellsOf(v, ctx)) {
    if (c === null) continue;
    if (range && typeof c === 'string') return Err.value;
    const d = toDate(c, ctx.date1904);
    if (isError(d)) return d;
    out.add(d);
  }
  return out;
};

const countWorkdays = (s: number, e: number, mask: boolean[], hol: Set<number>, d1904: boolean): number => {
  const sign = s > e ? -1 : 1;
  if (s > e) [s, e] = [e, s];
  const perWeek = mask.filter((x) => !x).length;
  const total = e - s + 1;
  const weeks = Math.floor(total / 7);
  let n = weeks * perWeek;
  for (let d = s + weeks * 7; d <= e; d++) if (!mask[wd(d, d1904)]) n++;
  for (const h of hol) if (h >= s && h <= e && !mask[wd(h, d1904)]) n--;
  return n * sign;
};

const networkdays = (args: Value[], ctx: FnContext, intl: boolean): Value => {
  const s = dateArg(args[0]!, ctx);
  if (isError(s)) return s;
  const e = dateArg(args[1]!, ctx);
  if (isError(e)) return e;
  const mask = weekendMask(intl ? args[2] : undefined, ctx, false);
  if (isError(mask)) return mask;
  const hol = holidaySet(intl ? args[3] : args[2], ctx);
  if (isError(hol)) return hol;
  return countWorkdays(s, e, mask, hol, ctx.date1904);
};

const workday = (args: Value[], ctx: FnContext, intl: boolean): Value => {
  const s = dateArg(args[0]!, ctx);
  if (isError(s)) return s;
  const n = int(args[1]!, ctx);
  if (isError(n)) return n;
  const mask = weekendMask(intl ? args[2] : undefined, ctx, true);
  if (isError(mask)) return mask;
  const hol = holidaySet(intl ? args[3] : args[2], ctx);
  if (isError(hol)) return hol;
  const step = n < 0 ? -1 : 1, max = maxSerial(ctx.date1904);
  let d = s, left = Math.abs(n);
  while (left > 0) {
    d += step;
    if (d < 0 || d > max) return Err.num;
    if (!mask[wd(d, ctx.date1904)] && !hol.has(d)) left--;
  }
  return d;
};

const WEEK_START: Record<number, number> = { 1: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6, 17: 0 };

const isoWeek = (s: number, d1904: boolean): number => {
  const thursday = s - ((wd(s, d1904) + 6) % 7) + 3;
  const y = partsOf(thursday, d1904).year;
  return Math.floor((thursday - serialOf(y, 1, 1, d1904)) / 7) + 1;
};

const basisArg = (v: Value | undefined, ctx: FnContext): number | CellError => {
  const b = optNum(v, ctx, 0);
  if (isError(b)) return b;
  const t = Math.trunc(b);
  return t < 0 || t > 4 ? Err.num : t;
};

export const defs: FunctionDef[] = [
  {
    name: 'DATE', min: 3, max: 3, scalar: true,
    fn: ([y, m, d], ctx) => {
      let yy = int(y!, ctx); if (isError(yy)) return yy;
      const mm = int(m!, ctx); if (isError(mm)) return mm;
      const dd = int(d!, ctx); if (isError(dd)) return dd;
      if (yy < 0 || yy >= 10000) return Err.num;
      if (yy < 1900) yy += 1900;
      const s = serialOf(yy, mm, dd, ctx.date1904);
      return s < 0 || s > maxSerial(ctx.date1904) ? Err.num : s;
    },
  },
  {
    name: 'DATEDIF', min: 3, max: 3, scalar: true,
    fn: ([a, b, u], ctx) => {
      const s = dateArg(a!, ctx); if (isError(s)) return s;
      const e = dateArg(b!, ctx); if (isError(e)) return e;
      const unit = text(u!, ctx); if (isError(unit)) return unit;
      if (s > e) return Err.num;
      const p = partsOf(s, ctx.date1904), q = partsOf(e, ctx.date1904);
      const months = (q.year - p.year) * 12 + q.month - p.month - (q.day < p.day ? 1 : 0);
      switch (unit.toUpperCase()) {
        case 'Y': return Math.floor(months / 12);
        case 'M': return months;
        case 'D': return e - s;
        case 'YM': return months % 12;
        // spreadsheet's MD counts from the start day in the month before the end date, overflowing as DATE does.
        case 'MD': return q.day >= p.day ? q.day - p.day : e - serialOf(q.year, q.month - 1, p.day, ctx.date1904);
        case 'YD': {
          const anniv = serialOf(q.year, p.month, p.day, ctx.date1904);
          return anniv <= e ? e - anniv : e - serialOf(q.year - 1, p.month, p.day, ctx.date1904);
        }
        default: return Err.num;
      }
    },
  },
  {
    name: 'DATEVALUE', min: 1, max: 1, scalar: true,
    fn: ([v], ctx) => {
      const s = textArg(v!, ctx); if (isError(s)) return s;
      const n = parseDateTime(s, ctx.date1904);
      return n === null ? Err.value : Math.floor(n);
    },
  },
  dateFn('DAY', (d) => d.day),
  dateFn('MONTH', (d) => d.month),
  dateFn('YEAR', (d) => d.year),
  {
    name: 'DAYS', min: 2, max: 2, scalar: true,
    fn: ([e, s], ctx) => {
      const a = dateArg(e!, ctx); if (isError(a)) return a;
      const b = dateArg(s!, ctx); if (isError(b)) return b;
      return a - b;
    },
  },
  {
    name: 'DAYS360', min: 2, max: 3, scalar: true,
    fn: ([a, b, m], ctx) => {
      const s = dateArg(a!, ctx); if (isError(s)) return s;
      const e = dateArg(b!, ctx); if (isError(e)) return e;
      const eu = m === undefined || m === null ? false : toNumber(first(m, ctx), ctx.date1904);
      if (isError(eu)) return eu;
      if (eu) return days360EU(s, e, ctx.date1904);
      const p = partsOf(s, ctx.date1904), q = partsOf(e, ctx.date1904);
      let sd = p.day, ed = q.day, em = q.month;
      if (isMonthEnd(s, ctx.date1904)) sd = 30;
      if (ed === 31) { if (sd < 30) { ed = 1; em++; } else ed = 30; }
      return (q.year - p.year) * 360 + (em - p.month) * 30 + ed - sd;
    },
  },
  {
    name: 'EDATE', min: 2, max: 2, scalar: true,
    fn: ([a, m], ctx) => {
      const s = dateArg(a!, ctx); if (isError(s)) return s;
      const n = int(m!, ctx); if (isError(n)) return n;
      const r = addMonths(s, n, ctx.date1904);
      return r < 0 || r > maxSerial(ctx.date1904) ? Err.num : r;
    },
  },
  {
    name: 'EOMONTH', min: 2, max: 2, scalar: true,
    fn: ([a, m], ctx) => {
      const s = dateArg(a!, ctx); if (isError(s)) return s;
      const n = int(m!, ctx); if (isError(n)) return n;
      const r = addMonths(s, n, ctx.date1904, true);
      return r < 0 || r > maxSerial(ctx.date1904) ? Err.num : r;
    },
  },
  timeFn('HOUR', (s) => Math.floor(s / 3600)),
  timeFn('MINUTE', (s) => Math.floor((s % 3600) / 60)),
  timeFn('SECOND', (s) => s % 60),
  dateFn('ISOWEEKNUM', (_d, s, ctx) => isoWeek(s, ctx.date1904)),
  { name: 'NETWORKDAYS', min: 2, max: 3, scalar: [0, 1], fn: (a, ctx) => networkdays(a, ctx, false) },
  { name: 'NETWORKDAYS.INTL', min: 2, max: 4, scalar: [0, 1, 2], fn: (a, ctx) => networkdays(a, ctx, true) },
  { name: 'WORKDAY', min: 2, max: 3, scalar: [0, 1], fn: (a, ctx) => workday(a, ctx, false) },
  { name: 'WORKDAY.INTL', min: 2, max: 4, scalar: [0, 1, 2], fn: (a, ctx) => workday(a, ctx, true) },
  { name: 'NOW', min: 0, max: 0, volatile: true, fn: (_a, ctx) => nowSerial(ctx.date1904) },
  { name: 'TODAY', min: 0, max: 0, volatile: true, fn: (_a, ctx) => todaySerial(ctx.date1904) },
  {
    name: 'TIME', min: 3, max: 3, scalar: true,
    fn: ([h, m, s], ctx) => {
      const hh = int(h!, ctx); if (isError(hh)) return hh;
      const mm = int(m!, ctx); if (isError(mm)) return mm;
      const ss = int(s!, ctx); if (isError(ss)) return ss;
      if (hh > 32767 || mm > 32767 || ss > 32767) return Err.num;
      const total = hh * 3600 + mm * 60 + ss;
      return total < 0 ? Err.num : (total % 86_400) / 86_400;
    },
  },
  {
    name: 'TIMEVALUE', min: 1, max: 1, scalar: true,
    fn: ([v], ctx) => {
      const s = textArg(v!, ctx); if (isError(s)) return s;
      const n = parseDateTime(s, ctx.date1904);
      return n === null ? Err.value : n - Math.floor(n);
    },
  },
  {
    name: 'WEEKDAY', min: 1, max: 2, scalar: true,
    fn: ([v, t], ctx) => {
      const s = dateArg(v!, ctx); if (isError(s)) return s;
      const type = optNum(t, ctx, 1); if (isError(type)) return type;
      const w = weekdayOf(s, ctx.date1904), k = Math.trunc(type);
      if (k === 1) return w + 1;
      if (k === 2) return ((w + 6) % 7) + 1;
      if (k === 3) return (w + 6) % 7;
      if (k >= 11 && k <= 17) return ((w - WEEK_START[k]! + 7) % 7) + 1;
      return Err.num;
    },
  },
  {
    name: 'WEEKNUM', min: 1, max: 2, scalar: true,
    fn: ([v, t], ctx) => {
      const s = dateArg(v!, ctx); if (isError(s)) return s;
      const type = optNum(t, ctx, 1); if (isError(type)) return type;
      const k = Math.trunc(type);
      if (k === 21) return isoWeek(s, ctx.date1904);
      const start = WEEK_START[k];
      if (start === undefined) return Err.num;
      const jan1 = serialOf(partsOf(s, ctx.date1904).year, 1, 1, ctx.date1904);
      const off = (wd(jan1, ctx.date1904) - start + 7) % 7;
      return Math.floor((s - jan1 + off) / 7) + 1;
    },
  },
  {
    name: 'YEARFRAC', min: 2, max: 3, scalar: true,
    fn: ([a, b, c], ctx) => {
      const s = dateArg(a!, ctx); if (isError(s)) return s;
      const e = dateArg(b!, ctx); if (isError(e)) return e;
      const basis = basisArg(c, ctx); if (isError(basis)) return basis;
      return yearFrac(s, e, basis, ctx.date1904);
    },
  },
];
