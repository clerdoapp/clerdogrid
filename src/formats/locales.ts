// Regional conventions for showing and typing numbers and dates. The table holds the separators,
// date order, currency and week start for each supported region; month and day names come from the
// platform's Intl data (Gregorian calendar), with English as the fallback. The active locale is set by
// app/locale.ts; lower layers (number formats, date entry) read it from here.

export type DateOrder = 'MDY' | 'DMY' | 'YMD';

export interface LocaleInfo {
  id: string;
  /** Name in its own language, for the picker. */
  name: string;
  /** Windows locale id, as used in number format codes like [$-407]. */
  lcid: number;
  dateOrder: DateOrder;
  dateSep: string;
  decimal: string;
  thousands: string;
  /** Formula argument separator shown in the formula bar. */
  listSep: ',' | ';';
  /** Separators between columns and between rows inside array constants such as {1,2;3,4}. */
  arrayCol: string;
  arrayRow: string;
  currency: string;
  currencyAfter: boolean;
  currencySpace: boolean;
  /** 0 = Sunday, 1 = Monday, 6 = Saturday. */
  firstDay: 0 | 1 | 6;
  /** Format code used to show the system short date (built-in format 14). */
  shortDate: string;
  longDate: string;
  time: string;
  rtl?: boolean;
  /** Native digits offered as a display option. */
  nativeDigits?: 'arab' | 'deva';
}

interface Row { ac?: string; id: string; name: string; lcid: number; o: DateOrder; ds: string; dec: string; th: string; cur: string; after: boolean; sp: boolean; fd: 0 | 1 | 6; sd: string; ld: string; t: string; rtl?: boolean; nd?: 'arab' | 'deva' }

const NB = ' ';
const ROWS: Row[] = [
  { id: 'en-US', name: 'English (United States)', lcid: 0x0409, o: 'MDY', ds: '/', dec: '.', th: ',', cur: '$', after: false, sp: false, fd: 0, sd: 'm/d/yyyy', ld: 'dddd, mmmm d, yyyy', t: 'h:mm AM/PM' },
  { id: 'en-GB', name: 'English (United Kingdom)', lcid: 0x0809, o: 'DMY', ds: '/', dec: '.', th: ',', cur: '£', after: false, sp: false, fd: 1, sd: 'dd/mm/yyyy', ld: 'dddd, d mmmm yyyy', t: 'hh:mm' },
  { id: 'ar-SA', name: 'العربية', lcid: 0x0401, o: 'DMY', ds: '/', dec: '.', th: ',', cur: 'ر.س.', after: true, sp: true, fd: 0, sd: 'dd/mm/yyyy', ld: 'dddd، d mmmm yyyy', t: 'hh:mm AM/PM', rtl: true, nd: 'arab' },
  { ac: '.', id: 'de-DE', name: 'Deutsch', lcid: 0x0407, o: 'DMY', ds: '.', dec: ',', th: '.', cur: '€', after: true, sp: true, fd: 1, sd: 'dd.mm.yyyy', ld: 'dddd, d. mmmm yyyy', t: 'hh:mm' },
  { id: 'de-CH', name: 'Deutsch (Schweiz)', lcid: 0x0807, o: 'DMY', ds: '.', dec: '.', th: '’', cur: 'CHF', after: false, sp: true, fd: 1, sd: 'dd.mm.yyyy', ld: 'dddd, d. mmmm yyyy', t: 'hh:mm' },
  { id: 'es-ES', name: 'Español (España)', lcid: 0x0c0a, o: 'DMY', ds: '/', dec: ',', th: '.', cur: '€', after: true, sp: true, fd: 1, sd: 'dd/mm/yyyy', ld: 'dddd, d "de" mmmm "de" yyyy', t: 'h:mm' },
  { id: 'es-MX', name: 'Español (México)', lcid: 0x080a, o: 'DMY', ds: '/', dec: '.', th: ',', cur: '$', after: false, sp: false, fd: 0, sd: 'dd/mm/yyyy', ld: 'dddd, d "de" mmmm "de" yyyy', t: 'hh:mm AM/PM' },
  { ac: '.', id: 'fr-FR', name: 'Français', lcid: 0x040c, o: 'DMY', ds: '/', dec: ',', th: NB, cur: '€', after: true, sp: true, fd: 1, sd: 'dd/mm/yyyy', ld: 'dddd d mmmm yyyy', t: 'hh:mm' },
  { ac: '.', id: 'fr-CA', name: 'Français (Canada)', lcid: 0x0c0c, o: 'YMD', ds: '-', dec: ',', th: NB, cur: '$', after: true, sp: true, fd: 0, sd: 'yyyy-mm-dd', ld: 'dddd d mmmm yyyy', t: 'hh:mm' },
  { id: 'it-IT', name: 'Italiano', lcid: 0x0410, o: 'DMY', ds: '/', dec: ',', th: '.', cur: '€', after: true, sp: true, fd: 1, sd: 'dd/mm/yyyy', ld: 'dddd d mmmm yyyy', t: 'hh:mm' },
  { id: 'pt-BR', name: 'Português (Brasil)', lcid: 0x0416, o: 'DMY', ds: '/', dec: ',', th: '.', cur: 'R$', after: false, sp: true, fd: 0, sd: 'dd/mm/yyyy', ld: 'dddd, d "de" mmmm "de" yyyy', t: 'hh:mm' },
  { id: 'pt-PT', name: 'Português (Portugal)', lcid: 0x0816, o: 'DMY', ds: '/', dec: ',', th: NB, cur: '€', after: true, sp: true, fd: 1, sd: 'dd/mm/yyyy', ld: 'dddd, d "de" mmmm "de" yyyy', t: 'hh:mm' },
  { id: 'nl-NL', name: 'Nederlands', lcid: 0x0413, o: 'DMY', ds: '-', dec: ',', th: '.', cur: '€', after: false, sp: true, fd: 1, sd: 'd-m-yyyy', ld: 'dddd d mmmm yyyy', t: 'hh:mm' },
  { id: 'sv-SE', name: 'Svenska', lcid: 0x041d, o: 'YMD', ds: '-', dec: ',', th: NB, cur: 'kr', after: true, sp: true, fd: 1, sd: 'yyyy-mm-dd', ld: '"den" d mmmm yyyy', t: 'hh:mm' },
  { id: 'pl-PL', name: 'Polski', lcid: 0x0415, o: 'DMY', ds: '.', dec: ',', th: NB, cur: 'zł', after: true, sp: true, fd: 1, sd: 'dd.mm.yyyy', ld: 'dddd, d mmmm yyyy', t: 'hh:mm' },
  { id: 'tr-TR', name: 'Türkçe', lcid: 0x041f, o: 'DMY', ds: '.', dec: ',', th: '.', cur: '₺', after: false, sp: false, fd: 1, sd: 'dd.mm.yyyy', ld: 'd mmmm yyyy dddd', t: 'hh:mm' },
  { id: 'ru-RU', name: 'Русский', lcid: 0x0419, o: 'DMY', ds: '.', dec: ',', th: NB, cur: '₽', after: true, sp: true, fd: 1, sd: 'dd.mm.yyyy', ld: 'd mmmm yyyy "г."', t: 'h:mm' },
  { id: 'uk-UA', name: 'Українська', lcid: 0x0422, o: 'DMY', ds: '.', dec: ',', th: NB, cur: '₴', after: true, sp: true, fd: 1, sd: 'dd.mm.yyyy', ld: 'd mmmm yyyy "р."', t: 'h:mm' },
  { id: 'zh-CN', name: '中文（简体）', lcid: 0x0804, o: 'YMD', ds: '/', dec: '.', th: ',', cur: '¥', after: false, sp: false, fd: 1, sd: 'yyyy/m/d', ld: 'yyyy"年"m"月"d"日"', t: 'h:mm' },
  { id: 'zh-TW', name: '中文（繁體）', lcid: 0x0404, o: 'YMD', ds: '/', dec: '.', th: ',', cur: 'NT$', after: false, sp: false, fd: 0, sd: 'yyyy/m/d', ld: 'yyyy"年"m"月"d"日"', t: 'AM/PM hh:mm' },
  { id: 'ja-JP', name: '日本語', lcid: 0x0411, o: 'YMD', ds: '/', dec: '.', th: ',', cur: '¥', after: false, sp: false, fd: 0, sd: 'yyyy/mm/dd', ld: 'yyyy"年"m"月"d"日"', t: 'h:mm' },
  { id: 'ko-KR', name: '한국어', lcid: 0x0412, o: 'YMD', ds: '-', dec: '.', th: ',', cur: '₩', after: false, sp: false, fd: 0, sd: 'yyyy-mm-dd', ld: 'yyyy"년" m"월" d"일" dddd', t: 'AM/PM h:mm' },
  { id: 'hi-IN', name: 'हिन्दी', lcid: 0x0439, o: 'DMY', ds: '-', dec: '.', th: ',', cur: '₹', after: false, sp: false, fd: 0, sd: 'dd-mm-yyyy', ld: 'dd mmmm yyyy', t: 'hh:mm AM/PM', nd: 'deva' },
  { id: 'id-ID', name: 'Bahasa Indonesia', lcid: 0x0421, o: 'DMY', ds: '/', dec: ',', th: '.', cur: 'Rp', after: false, sp: false, fd: 1, sd: 'dd/mm/yyyy', ld: 'dddd, dd mmmm yyyy', t: 'h:mm' },
  { id: 'he-IL', name: 'עברית', lcid: 0x040d, o: 'DMY', ds: '/', dec: '.', th: ',', cur: '₪', after: true, sp: true, fd: 0, sd: 'dd/mm/yyyy', ld: 'dddd dd mmmm yyyy', t: 'hh:mm', rtl: true },
];

export const LOCALES: readonly LocaleInfo[] = ROWS.map((r) => ({
  id: r.id, name: r.name, lcid: r.lcid, dateOrder: r.o, dateSep: r.ds, decimal: r.dec, thousands: r.th,
  listSep: r.dec === ',' ? ';' : ',', arrayCol: r.ac ?? (r.dec === ',' ? '\\' : ','), arrayRow: ';', currency: r.cur, currencyAfter: r.after, currencySpace: r.sp, firstDay: r.fd,
  shortDate: r.sd, longDate: r.ld, time: r.t, rtl: r.rtl, nativeDigits: r.nd,
}));

/** More Windows locale ids, used only to pick month and day names for [$-xxxx] format codes. */
const EXTRA_LCID: Record<number, string> = {
  0x0c09: 'en-AU', 0x1009: 'en-CA', 0x4009: 'en-IN', 0x1409: 'en-NZ', 0x0c07: 'de-AT', 0x080c: 'fr-BE', 0x100c: 'fr-CH',
  0x040a: 'es-ES', 0x2c0a: 'es-AR', 0x0813: 'nl-BE', 0x0c04: 'zh-HK', 0x1004: 'zh-SG', 0x0c01: 'ar-EG', 0x0801: 'ar-IQ',
  0x3801: 'ar-AE', 0x0406: 'da-DK', 0x0414: 'nb-NO', 0x040b: 'fi-FI', 0x0405: 'cs-CZ', 0x040e: 'hu-HU', 0x0408: 'el-GR',
  0x0418: 'ro-RO', 0x041e: 'th-TH', 0x042a: 'vi-VN', 0x0429: 'fa-IR', 0x041b: 'sk-SK', 0x0424: 'sl-SI', 0x041a: 'hr-HR',
  0x0402: 'bg-BG', 0x0425: 'et-EE', 0x0426: 'lv-LV', 0x0427: 'lt-LT', 0x0403: 'ca-ES', 0x0420: 'ur-PK', 0x0445: 'bn-IN',
  0x043e: 'ms-MY', 0x0464: 'fil-PH',
};

export const localeById = (id: string): LocaleInfo | undefined => {
  const low = id.toLowerCase();
  return LOCALES.find((l) => l.id.toLowerCase() === low) ?? LOCALES.find((l) => l.id.slice(0, 2).toLowerCase() === low.slice(0, 2));
};

/** BCP 47 tag for a Windows locale id, or undefined. Language-only ids match the first region with that language. */
export const tagForLcid = (lcid: number): string | undefined => {
  const lang = lcid & 0xffff;
  const hit = LOCALES.find((l) => l.lcid === lang)?.id ?? EXTRA_LCID[lang];
  if (hit) return hit;
  const primary = lang & 0x3ff;
  if (!primary) return undefined;
  return LOCALES.find((l) => (l.lcid & 0x3ff) === primary)?.id ?? Object.entries(EXTRA_LCID).find(([k]) => (+k & 0x3ff) === primary)?.[1];
};

// ---------- names ----------

export interface DateNames { months: string[]; monthsShort: string[]; days: string[]; daysShort: string[]; am: string; pm: string }

const EN: DateNames = {
  months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  monthsShort: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  daysShort: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
  am: 'AM', pm: 'PM',
};

const namesCache = new Map<string, DateNames>();
/** Month, day and AM/PM names for a language tag. English for "en" tags and wherever Intl can't help. */
export const dateNames = (tag: string): DateNames => {
  if (/^en\b/i.test(tag)) return EN;
  const hit = namesCache.get(tag);
  if (hit) return hit;
  let res = EN;
  try {
    const opt = { timeZone: 'UTC', calendar: 'gregory', numberingSystem: 'latn' } as const;
    const fm = (o: Intl.DateTimeFormatOptions, d: number): string => new Intl.DateTimeFormat(tag, { ...opt, ...o }).format(d);
    const months = Array.from({ length: 12 }, (_, i) => fm({ month: 'long' }, Date.UTC(2021, i, 15)));
    const monthsShort = Array.from({ length: 12 }, (_, i) => fm({ month: 'short' }, Date.UTC(2021, i, 15)));
    const days = Array.from({ length: 7 }, (_, i) => fm({ weekday: 'long' }, Date.UTC(2021, 0, 3 + i)));
    const daysShort = Array.from({ length: 7 }, (_, i) => fm({ weekday: 'short' }, Date.UTC(2021, 0, 3 + i)));
    const period = (h: number): string => new Intl.DateTimeFormat(tag, { ...opt, hour: 'numeric', hour12: true }).formatToParts(Date.UTC(2021, 0, 1, h)).find((p) => p.type === 'dayPeriod')?.value ?? (h < 12 ? 'AM' : 'PM');
    // An engine without data for the language answers in English; months that look like numbers mean the same.
    if (months.every((m) => !/^\d+$/.test(m))) res = { months, monthsShort, days, daysShort, am: period(9), pm: period(21) };
  } catch { res = EN; }
  namesCache.set(tag, res);
  return res;
};

// ---------- active locale ----------

let current: LocaleInfo = LOCALES[0]!;
let currentNames: DateNames = EN;

export const activeLocale = (): LocaleInfo => current;
export const activeNames = (): DateNames => currentNames;

/** Sets the active locale (used by app/locale.ts; other code should go through that module). */
export const setActiveLocaleInfo = (loc: LocaleInfo): void => {
  current = loc;
  currentNames = dateNames(loc.id);
};

/** Names for a format code's [$-xxxx] locale, or the active locale's when there is none. */
export const namesForLcid = (lcid: number | undefined): DateNames => {
  if (lcid === undefined || (lcid & 0xffff) === 0) return currentNames;
  const tag = tagForLcid(lcid);
  return tag ? dateNames(tag) : currentNames;
};

const DIGITS: Record<string, string> = { arab: '٠١٢٣٤٥٦٧٨٩', pers: '۰۱۲۳۴۵۶۷۸۹', deva: '०१२३४५६७८९' };
/** Replaces ASCII digits with a native digit set (and the Arabic decimal/thousands marks for Arabic digits). */
export const toNativeDigits = (s: string, set: 'arab' | 'pers' | 'deva'): string => {
  const d = DIGITS[set]!;
  let out = s.replace(/[0-9]/g, (c) => d[+c]!);
  if (set === 'arab') out = out.replace(/([٠-٩])\.(?=[٠-٩])/g, '$1٫').replace(/([٠-٩]),(?=[٠-٩])/g, '$1٬');
  return out;
};
/** Turns any native digits (Arabic-Indic, Persian, Devanagari, full-width) back into ASCII. */
export const fromNativeDigits = (s: string): string =>
  s.replace(/[٠-٩]/g, (c) => String(c.charCodeAt(0) - 0x660))
    .replace(/[۰-۹]/g, (c) => String(c.charCodeAt(0) - 0x6f0))
    .replace(/[०-९]/g, (c) => String(c.charCodeAt(0) - 0x966))
    .replace(/[０-９]/g, (c) => String(c.charCodeAt(0) - 0xff10))
    .replace(/٫/g, '.').replace(/٬/g, ',');

let symCache: { id: string; list: string[] } | null = null;
/** Currency symbols recognised when numbers are typed (longest first so "R$" wins over "$"). */
export const currencySymbols = (): string[] => {
  if (symCache?.id === current.id) return symCache.list;
  const list = [...new Set([current.currency, '$', '€', '£', '¥', '￥', '₹', '₩', '₺', '₽', '₴', '₪', '₫', '฿', 'zł', 'kr', 'R$', 'Rp', 'CHF', 'NT$', 'ر.س.'])]
    .sort((a, b) => b.length - a.length);
  symCache = { id: current.id, list };
  return list;
};

/** Number format for an amount typed with `symbol` before or after it. */
export const currencyFormatFor = (symbol: string, after: boolean, decimals: boolean): string => {
  const loc = activeLocale();
  const num = decimals ? '#,##0.00' : '#,##0';
  if (symbol === '$' && !after && loc.currency === '$' && loc.id === 'en-US') return decimals ? '$#,##0.00' : '$#,##0';
  const own = symbol === loc.currency;
  const tag = own ? `[$${symbol}-${loc.lcid.toString(16).toUpperCase()}]` : `"${symbol}"`;
  const space = own ? loc.currencySpace : false;
  const useAfter = own ? loc.currencyAfter : after;
  return useAfter ? `${num}${space ? '\\ ' : ''}${tag}` : `${tag}${space ? '\\ ' : ''}${num}`;
};
