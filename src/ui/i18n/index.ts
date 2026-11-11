// Interface languages. English text is the key: en.ts lists every message, and each other language file
// maps the same keys to its own text (typed as Catalog, so a missing key is a type error).
// Messages use {name} for values and ICU-style plurals: "{count, plural, one {# row} other {# rows}}".
// Function names, error codes, cell contents and file-format strings are never translated.
import { en } from './en';
import { userSheetDefaults } from '../../model/workbook';

export type MsgKey = keyof typeof en;
export type Catalog = Record<MsgKey, string>;
export type Params = Record<string, string | number>;

export interface LangInfo { code: string; name: string; english: string; rtl?: boolean; region: string }

/** The interface languages, in the order the Options list shows them. */
export const LANGS: readonly LangInfo[] = [
  { code: 'en', name: 'English', english: 'English', region: 'en-US' },
  { code: 'ar', name: 'العربية', english: 'Arabic', rtl: true, region: 'ar-SA' },
  { code: 'de', name: 'Deutsch', english: 'German', region: 'de-DE' },
  { code: 'es', name: 'Español', english: 'Spanish', region: 'es-ES' },
  { code: 'fr', name: 'Français', english: 'French', region: 'fr-FR' },
  { code: 'it', name: 'Italiano', english: 'Italian', region: 'it-IT' },
  { code: 'pt', name: 'Português', english: 'Portuguese', region: 'pt-BR' },
  { code: 'nl', name: 'Nederlands', english: 'Dutch', region: 'nl-NL' },
  { code: 'sv', name: 'Svenska', english: 'Swedish', region: 'sv-SE' },
  { code: 'pl', name: 'Polski', english: 'Polish', region: 'pl-PL' },
  { code: 'tr', name: 'Türkçe', english: 'Turkish', region: 'tr-TR' },
  { code: 'ru', name: 'Русский', english: 'Russian', region: 'ru-RU' },
  { code: 'uk', name: 'Українська', english: 'Ukrainian', region: 'uk-UA' },
  { code: 'zh', name: '简体中文', english: 'Chinese (Simplified)', region: 'zh-CN' },
  { code: 'ja', name: '日本語', english: 'Japanese', region: 'ja-JP' },
  { code: 'ko', name: '한국어', english: 'Korean', region: 'ko-KR' },
  { code: 'hi', name: 'हिन्दी', english: 'Hindi', region: 'hi-IN' },
  { code: 'id', name: 'Bahasa Indonesia', english: 'Indonesian', region: 'id-ID' },
  { code: 'he', name: 'עברית', english: 'Hebrew', rtl: true, region: 'he-IL' },
];

export const LANG_KEY = 'clerdogrid.lang';
/** Options key: new sheets start right to left when the interface language is. */
export const RTL_SHEETS_KEY = 'clerdogrid.rtlSheets';

// One lazy chunk per language, so only the chosen one is loaded.
const loaders = import.meta.glob<{ default: Catalog }>('./lang/*.ts');

let lang = 'en';
let current: Catalog = en;
const plurals = new Map<string, Intl.PluralRules>();
const listeners = new Set<() => void>();

const read = (k: string): string | null => { try { return localStorage.getItem(k); } catch { return null; } };
const write = (k: string, v: string): void => { try { localStorage.setItem(k, v); } catch { /* private mode */ } };

/** Best interface language for a BCP 47 tag ("pt-BR" → pt, "iw" → he); English when there is none. */
export const matchLang = (tag: string | null | undefined): string => {
  if (!tag) return 'en';
  const base = tag.toLowerCase().split(/[-_]/)[0]!;
  const alias: Record<string, string> = { iw: 'he', in: 'id' };
  const code = alias[base] ?? base;
  return LANGS.some((l) => l.code === code) ? code : 'en';
};

export const getLang = (): string => lang;
export const langInfo = (code = lang): LangInfo => LANGS.find((l) => l.code === code) ?? LANGS[0]!;
export const isRtlLang = (code = lang): boolean => !!langInfo(code).rtl;
export const onLangChange = (f: () => void): (() => void) => { listeners.add(f); return () => listeners.delete(f); };

const pluralRules = (code: string): Intl.PluralRules => {
  let p = plurals.get(code);
  if (!p) { p = new Intl.PluralRules(code); plurals.set(code, p); }
  return p;
};

/** Splits "{a} and {b, plural, one {…} other {…}}" into top-level text and {…} groups. */
const groups = (msg: string): { text: string; inner: string | null }[] => {
  const out: { text: string; inner: string | null }[] = [];
  let depth = 0, start = 0, buf = '';
  for (let i = 0; i < msg.length; i++) {
    const ch = msg[i]!;
    if (ch === '{') { if (depth === 0) { buf += msg.slice(start, i); start = i + 1; } depth++; }
    else if (ch === '}' && depth > 0) {
      depth--;
      if (depth === 0) { out.push({ text: buf, inner: msg.slice(start, i) }); buf = ''; start = i + 1; }
    }
  }
  out.push({ text: buf + msg.slice(start), inner: null });
  return out;
};

/** "one {# row} other {# rows}" → { one: '# row', other: '# rows' }. */
export const pluralCases = (body: string): Record<string, string> | null => {
  const cases: Record<string, string> = {};
  let i = 0;
  while (i < body.length) {
    while (i < body.length && /\s/.test(body[i]!)) i++;
    if (i >= body.length) break;
    const m = /^(=\d+|zero|one|two|few|many|other)\s*\{/.exec(body.slice(i));
    if (!m) return null;
    i += m[0].length;
    let depth = 1;
    const from = i;
    while (i < body.length && depth) { if (body[i] === '{') depth++; else if (body[i] === '}') depth--; i++; }
    if (depth) return null;
    cases[m[1]!] = body.slice(from, i - 1);
  }
  return cases.other !== undefined ? cases : null;
};

/** Fills {name} values and picks plural forms for `code`. Unknown names stay as written. */
export const formatMessage = (msg: string, params: Params | undefined, code = lang): string => {
  if (!msg.includes('{')) return msg;
  let out = '';
  for (const g of groups(msg)) {
    out += g.text;
    if (g.inner === null) continue;
    const pm = /^\s*(\w+)\s*,\s*plural\s*,([\s\S]*)$/.exec(g.inner);
    if (pm) {
      const n = Number(params?.[pm[1]!] ?? 0);
      const cases = pluralCases(pm[2]!) ?? { other: pm[2]! };
      const chosen = cases[`=${n}`] ?? cases[pluralRules(code).select(n)] ?? cases.other!;
      out += formatMessage(chosen.replace(/#/g, String(n)), params, code);
      continue;
    }
    const name = g.inner.trim();
    out += params && name in params ? String(params[name]) : `{${g.inner}}`;
  }
  return out;
};

/** The message for `key` in the interface language. */
export const t = (key: MsgKey, params?: Params): string => formatMessage(current[key] || en[key] || key, params);

/** Translates text that may be a catalogue key (command labels, data tables); other text comes back as given. */
export function tr(text: string | undefined, params?: Params): string;
/** Anything that isn't text (an element, a number) comes back unchanged. */
export function tr<V>(text: V): V;
export function tr(text: unknown, params?: Params): unknown {
  if (typeof text !== 'string') return text ?? '';
  if (!text) return '';
  return Object.prototype.hasOwnProperty.call(en, text) ? t(text as MsgKey, params) : params ? formatMessage(text, params) : text;
}

/** Marks text in data tables as a catalogue key; it is translated where it is shown, with tr(). */
export const N = (k: MsgKey): string => k;

// Short names for use in components, where `t` and `tr` are often local variables.
export { t as T, tr as Tr };

/** True when `text` is a catalogue key. */
export const isKey = (text: string): text is MsgKey => Object.prototype.hasOwnProperty.call(en, text);

const loadCatalog = async (code: string): Promise<Catalog> => {
  if (code === 'en') return en;
  const load = loaders[`./lang/${code}.ts`];
  if (!load) return en;
  try { return { ...en, ...(await load()).default }; } catch { return en; }
};

const applyDocument = (): void => {
  if (typeof document === 'undefined') return;
  const root = document.documentElement;
  root.lang = lang === 'zh' ? 'zh-CN' : lang;
  root.dir = isRtlLang() ? 'rtl' : 'ltr';
};

/** Switches the interface language. The UI re-renders through onLangChange listeners; no reload. */
export const setLang = async (code: string, opts: { persist?: boolean } = {}): Promise<void> => {
  const next = LANGS.some((l) => l.code === code) ? code : 'en';
  const first = read(LANG_KEY) === null;
  const cat = await loadCatalog(next);
  lang = next;
  current = cat;
  if (opts.persist) {
    write(LANG_KEY, next);
    if (first) await matchRegion(next);
  }
  applyDocument();
  userSheetDefaults.rtl = rtlSheetsDefault();
  for (const f of listeners) f();
};

/** The first time a language is chosen, the region follows it unless the person already picked one. */
const matchRegion = async (code: string): Promise<void> => {
  if (read('clerdogrid.locale') !== null || code === 'en') return;
  const { setLocale } = await import('../../app/locale');
  setLocale(langInfo(code).region, { persist: true });
};

/** Start-up: the saved language, else the browser's, else English. */
export const initLang = async (): Promise<void> => {
  const saved = read(LANG_KEY);
  const nav = typeof navigator !== 'undefined' ? navigator.language : undefined;
  await setLang(saved && LANGS.some((l) => l.code === saved) ? saved : matchLang(nav));
};

/** Whether new sheets start right to left: the Options choice, else on when the interface is right to left. */
export const rtlSheetsDefault = (): boolean => {
  const v = read(RTL_SHEETS_KEY);
  return v === null ? isRtlLang() : v === '1';
};
export const setRtlSheetsDefault = (on: boolean | null): void => {
  try { if (on === null) localStorage.removeItem(RTL_SHEETS_KEY); else localStorage.setItem(RTL_SHEETS_KEY, on ? '1' : '0'); } catch { /* private mode */ }
  userSheetDefaults.rtl = rtlSheetsDefault();
};

/** For tests: use a catalogue directly. */
export const useCatalogForTest = (code: string, cat: Catalog | null): void => { lang = code; current = cat ?? en; };
