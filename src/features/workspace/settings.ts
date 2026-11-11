// App settings from File > Options, kept on this device. Workbook settings (calculation, 1904 dates,
// R1C1 references) live in the workbook itself and are applied from there.
import type { GridStore } from '../../app/store';

export type Aggregate = 'average' | 'count' | 'numCount' | 'min' | 'max' | 'sum';
export type ErrorRule = 'numberStoredAsText' | 'formula' | 'formulaRange' | 'unlockedFormula' | 'twoDigitTextYear' | 'evalError';

export interface AppSettings {
  defaultFont: string;
  defaultSize: number;
  sheetCount: number;
  enterMove: boolean;
  enterDirection: 'down' | 'right' | 'up' | 'left';
  autocomplete: boolean;
  formulaAutocomplete: boolean;
  functionTips: boolean;
  flashFillAuto: boolean;
  /** Empty means the number format's own separators. */
  decimalSep: string;
  thousandsSep: string;
  errorChecking: boolean;
  errorColor: string;
  rules: Record<ErrorRule, boolean>;
  statusItems: Aggregate[];
  focusColor: string;
  versionsKeep: number;
  versionsEvery: number;
  versionsOn: boolean;
  showFormulaBar: boolean;
  showStatusBar: boolean;
}

export const DEFAULTS: AppSettings = {
  defaultFont: 'Calibri',
  defaultSize: 11,
  sheetCount: 1,
  enterMove: true,
  enterDirection: 'down',
  autocomplete: true,
  formulaAutocomplete: true,
  functionTips: true,
  flashFillAuto: true,
  decimalSep: '',
  thousandsSep: '',
  errorChecking: true,
  errorColor: '#1e9e4a',
  rules: { numberStoredAsText: true, formula: true, formulaRange: true, unlockedFormula: true, twoDigitTextYear: true, evalError: true },
  statusItems: ['average', 'count', 'sum'],
  focusColor: '#f5c542',
  versionsKeep: 50,
  versionsEvery: 5,
  versionsOn: true,
  showFormulaBar: true,
  showStatusBar: true,
};

const KEY = 'clerdogrid.options';
export const USER_KEY = 'clerdogrid.user';
export const THEME_KEY = 'clerdogrid.theme';
export const LANG_KEY = 'clerdogrid.lang';

const read = (k: string): string | null => { try { return localStorage.getItem(k); } catch { return null; } };
const write = (k: string, v: string | null): void => { try { if (v === null) localStorage.removeItem(k); else localStorage.setItem(k, v); } catch { /* private mode */ } };

const load = (): AppSettings => {
  const raw = read(KEY);
  if (!raw) return structuredClone(DEFAULTS);
  try {
    const o = JSON.parse(raw) as Partial<AppSettings>;
    return { ...structuredClone(DEFAULTS), ...o, rules: { ...DEFAULTS.rules, ...o.rules }, statusItems: Array.isArray(o.statusItems) ? o.statusItems : DEFAULTS.statusItems };
  } catch { return structuredClone(DEFAULTS); }
};

export let settings: AppSettings = load();

export const saveSettings = (store: GridStore | null, patch: Partial<AppSettings>): void => {
  settings = { ...settings, ...patch };
  write(KEY, JSON.stringify(settings));
  if (store) { applySettings(store); store.touch(); }
};

export const userName = (): string => read(USER_KEY)?.trim() ?? '';
export const setUserName = (n: string): void => write(USER_KEY, n.trim() || null);
export const language = (): string => read(LANG_KEY) ?? 'en';
export const setLanguage = (l: string): void => write(LANG_KEY, l);
export const setThemeSetting = (t: 'system' | 'light' | 'dark'): void => write(THEME_KEY, t === 'system' ? null : t);

/** Pushes the device settings into the store and the page. */
export const applySettings = (store: GridStore): void => {
  store.enterDirection = settings.enterMove ? settings.enterDirection : 'none';
  const root = typeof document !== 'undefined' ? document.documentElement : null;
  if (root) {
    root.toggleAttribute('data-no-fbar', !settings.showFormulaBar);
    root.toggleAttribute('data-no-status', !settings.showStatusBar);
  }
};
