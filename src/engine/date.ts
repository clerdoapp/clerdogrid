// spreadsheet date serials. In the 1900 system serial 1 is 1900-01-01 and serial 60 is the
// non-existent 1900-02-29 that spreadsheets keep for compatibility; in the 1904 system serial 0 is 1904-01-01.

import { activeLocale, activeNames } from '../formats/locales';

const DAY_MS = 86_400_000;
const EPOCH_1900 = Date.UTC(1899, 11, 30);
const EPOCH_1904 = Date.UTC(1904, 0, 1);

export interface DateParts { year: number; month: number; day: number }

export const isLeap = (y: number): boolean => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;

export const daysInMonth = (y: number, m: number): number => new Date(Date.UTC(y, m, 0)).getUTCDate();

/** Serial for a date; month and day may overflow, as with DATE(2024, 14, 40). */
export const toSerial = (year: number, month: number, day: number, date1904 = false): number => {
  if (!date1904 && year >= 0 && year < 1900) year += 1900;
  const ms = Date.UTC(year, month - 1, 1) + (day - 1) * DAY_MS;
  if (date1904) return Math.round((ms - EPOCH_1904) / DAY_MS);
  const serial = Math.round((ms - EPOCH_1900) / DAY_MS);
  return serial <= 60 ? serial - 1 : serial;
};

export const fromSerial = (serial: number, date1904 = false): DateParts => {
  const whole = Math.floor(serial);
  if (!date1904 && whole === 60) return { year: 1900, month: 2, day: 29 };
  if (!date1904 && whole === 0) return { year: 1900, month: 1, day: 0 };
  const ms = date1904 ? EPOCH_1904 + whole * DAY_MS : EPOCH_1900 + (whole < 60 ? whole + 1 : whole) * DAY_MS;
  const d = new Date(ms);
  return { year: d.getUTCFullYear(), month: d.getUTCMonth() + 1, day: d.getUTCDate() };
};

/** 0 = Sunday, matching JavaScript. */
export const weekdayOf = (serial: number, date1904 = false): number => {
  const whole = Math.floor(serial);
  if (!date1904 && whole < 61) return ((whole % 7) + 6) % 7; // spreadsheets count 1900-01-01 as a Sunday
  const { year, month, day } = fromSerial(whole, date1904);
  return new Date(Date.UTC(year, month - 1, day)).getUTCDay();
};

export interface TimeParts { hour: number; minute: number; second: number; ms: number }

export const timeOf = (serial: number): TimeParts => {
  let frac = serial - Math.floor(serial);
  let total = Math.round(frac * DAY_MS);
  if (total >= DAY_MS) total = 0;
  frac = total;
  const hour = Math.floor(frac / 3_600_000);
  const minute = Math.floor((frac % 3_600_000) / 60_000);
  const second = Math.floor((frac % 60_000) / 1000);
  return { hour, minute, second, ms: frac % 1000 };
};

export const timeSerial = (h: number, m: number, s: number): number => (h * 3600 + m * 60 + s) / 86_400;

/** Serial for "now" in local time. */
export const nowSerial = (date1904 = false, at = new Date()): number => {
  const day = toSerial(at.getFullYear(), at.getMonth() + 1, at.getDate(), date1904);
  return day + timeSerial(at.getHours(), at.getMinutes(), at.getSeconds() + at.getMilliseconds() / 1000);
};

export const todaySerial = (date1904 = false, at = new Date()): number => toSerial(at.getFullYear(), at.getMonth() + 1, at.getDate(), date1904);

const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
const FULL = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
/** English month names (full, or any prefix of three letters or more), then the active locale's names. */
const monthIndex = (s: string): number => {
  const low = s.toLowerCase().replace(/\.$/, '');
  const i = MONTHS.indexOf(low.slice(0, 3));
  if (i >= 0 && FULL[i]!.startsWith(low)) return i + 1;
  const n = activeNames();
  for (const list of [n.months, n.monthsShort]) {
    const k = list.findIndex((x) => x.toLowerCase().replace(/\.$/, '') === low);
    if (k >= 0) return k + 1;
  }
  return -1;
};

const parseTime = (s: string): number | null => {
  const m = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}(?:\.\d+)?))?\s*(am|pm|a|p)?$/i.exec(s.trim());
  if (!m) return null;
  let h = +m[1]!;
  const mi = +m[2]!;
  const sec = m[3] ? +m[3] : 0;
  if (mi > 59 || sec >= 60) return null;
  if (m[4]) {
    if (h < 1 || h > 12) return null;
    const pm = m[4].toLowerCase().startsWith('p');
    h = (h % 12) + (pm ? 12 : 0);
  }
  return timeSerial(h, mi, sec);
};

const valid = (y: number, m: number, d: number): boolean => m >= 1 && m <= 12 && d >= 1 && d <= daysInMonth(y, m) && y >= 1900 && y <= 9999;

/**
 * Reads dates and times typed as text, the way spreadsheets do for entry and DATEVALUE/TIMEVALUE:
 * 2024-03-15, 3/15/2024, 15-Mar-2024, March 15, 2024, 15 March 2024, Mar-24, 14:30, 2:30 PM,
 * and a date followed by a time. Month/day order follows the US convention.
 */
export const parseDateTime = (input: string, date1904 = false): number | null => {
  const s = input.trim();
  if (!s) return null;
  const t = parseTime(s);
  if (t !== null) return t;
  const split = /^(.*?)[\sT]+(\d{1,2}:\d{1,2}(?::\d{1,2}(?:\.\d+)?)?\s*(?:am|pm|a|p)?)$/i.exec(s);
  if (split && /\d/.test(split[1]!)) {
    const d = parseDateTime(split[1]!, date1904);
    const tt = parseTime(split[2]!);
    if (d !== null && tt !== null && d === Math.floor(d)) return d + tt;
  }
  const thisYear = new Date().getFullYear();
  const year = (y: string): number => {
    const n = +y;
    return y.length <= 2 ? (n < 30 ? 2000 + n : 1900 + n) : n;
  };
  let m: RegExpExecArray | null;
  let y = 0, mo = 0, d = 0;
  const order = activeLocale().dateOrder;
  if ((m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})\.?$/.exec(s))) { y = +m[1]!; mo = +m[2]!; d = +m[3]!; }
  else if ((m = /^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})$/.exec(s))) {
    // Numbers only: the region decides which is the day and which the month.
    if (order === 'DMY') { d = +m[1]!; mo = +m[2]!; } else { mo = +m[1]!; d = +m[2]!; }
    y = year(m[3]!);
  }
  else if ((m = /^(\d{1,2})[/\-.](\d{1,2})\.?$/.exec(s)) && (order !== 'MDY' || !s.includes('.'))) {
    if (order === 'MDY') { mo = +m[1]!; d = +m[2]!; }
    else if (order === 'DMY') { d = +m[1]!; mo = +m[2]!; }
    else { mo = +m[1]!; d = +m[2]!; }
    y = thisYear;
  }
  else if ((m = /^(\d{1,2})\.?[\s\-/]+(\p{L}{3,}\.?)(?:[\s\-/,]+(\d{2,4}))?$/iu.exec(s))) { d = +m[1]!; mo = monthIndex(m[2]!); y = m[3] ? year(m[3]) : thisYear; }
  else if ((m = /^(\p{L}{3,}\.?)[\s\-/]+(\d{1,2})(?:(?:st|nd|rd|th)?,?[\s\-/]+(\d{4}))?$/iu.exec(s))) { mo = monthIndex(m[1]!); d = +m[2]!; y = m[3] ? +m[3] : thisYear; }
  else if ((m = /^(\p{L}{3,}\.?)[\s\-/]+(\d{4})$/iu.exec(s))) { mo = monthIndex(m[1]!); d = 1; y = +m[2]!; }
  else return null;
  if (mo < 1 || !valid(y, mo, d)) return null;
  return toSerial(y, mo, d, date1904);
};
