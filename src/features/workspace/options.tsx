// File > Options: device settings (kept in this browser or app) and workbook settings (saved in the file).
import { useState } from 'react';
import { formatValue } from '../../formats/numfmt';
import { keyCol, keyRow, newUserSheet, userSheetDefaults, type Workbook } from '../../model/workbook';
import type { GridStore } from '../../app/store';
import { setSeparators } from '../../app/locale';
import { assistPrefs } from '../../ui/FormulaAssist';
import { statusPrefs, STATUS_ITEMS, type StatusItem } from '../../ui/StatusBar';
import { FONT_NAMES } from '../../ui/fonts';
import { Dialog, Field, type DialogProps } from '../../ui/dialogs/Dialog';
import { Check, ColorPick, Radio, Tabs } from '../../ui/dialogs/shared';
import { applySettings, saveSettings, setThemeSetting, setUserName, settings, userName, type AppSettings, type ErrorRule } from './settings';
import { RULE_TEXT, resetIgnored } from './errors';
import { LANGS, N, RTL_SHEETS_KEY, T, Tr, getLang, setLang as setInterfaceLang, setRtlSheetsDefault } from '../../ui/i18n';

type Tab = 'general' | 'formulas' | 'proofing' | 'advanced' | 'save';
const TABS: [Tab, string][] = [['general', N('General')], ['formulas', N('Formulas')], ['proofing', N('Error checking')], ['advanced', N('Advanced')], ['save', N('Versions')]];
const LANGUAGES: [string, string][] = LANGS.map((l) => [l.code, l.code === 'en' ? l.name : `${l.name} (${l.english})`]);

type CalcMode = 'auto' | 'autoNoTable' | 'manual';

const DATE_CODE = /[dmyhs]/i;
/** A number rounded to what its format shows (precision as displayed). General and dates keep full precision. */
export const roundToDisplay = (v: number, fmt: string | undefined): number => {
  if (!fmt || fmt === 'General' || fmt === '@' || DATE_CODE.test(fmt.replace(/"[^"]*"|\[[^\]]*\]|\\./g, ''))) return v;
  const text = formatValue(v, fmt).text;
  if (/E[+-]?\d+/i.test(text)) { const n = Number(text.replace(/[^\d.E+-]/gi, '')); return Number.isFinite(n) ? n : v; }
  const digits = text.replace(/[^\d.]/g, '');
  if (!/\d/.test(digits)) return v;
  let n = Number(digits);
  if (!Number.isFinite(n)) return v;
  if (text.includes('%')) n /= 100;
  // Thousands scaling ("#,##0," shows thousands) can't be undone from text alone: keep full precision then.
  if (/[0#?],+(?:[;"\s]|$)/.test(fmt.split(';')[0]!)) return v;
  return v < 0 ? -n : n;
};

/** Pushes the workbook's own settings into the engine and the store. */
export const applyWorkbookSettings = (store: GridStore): void => {
  const c = store.wb.calc ?? {};
  store.engine.manual = !!c.manual;
  store.engine.iterate = !!c.iterate;
  store.engine.maxIterations = c.maxIterations ?? 100;
  store.engine.maxChange = c.maxChange ?? 0.001;
  store.r1c1 = !!c.r1c1;
  store.engine.roundResult = c.precisionAsDisplayed ? (sheet, row, col, v) => roundToDisplay(v, store.engine.numberFormat(sheet, row, col)) : null;
};

const roundConstants = (wb: Workbook, store: GridStore): void => {
  wb.sheets.forEach((s, si) => {
    for (const [k, c] of s.cells) {
      if (c.f || typeof c.v !== 'number') continue;
      const r = roundToDisplay(c.v, store.style(keyRow(k), keyCol(k), si).numFmt);
      if (r !== c.v) { c.v = r; store.engine.markChanged(si, keyRow(k), keyCol(k)); }
    }
  });
};

export const OptionsDialog = ({ store, close, data }: DialogProps<{ tab?: Tab }>) => {
  const [tab, setTab] = useState<Tab>(data?.tab ?? 'general');
  const [s, setS] = useState<AppSettings>(() => structuredClone(settings));
  const set = (p: Partial<AppSettings>) => setS((x) => ({ ...x, ...p }));
  const calc0 = store.wb.calc ?? {};
  const [calcMode, setCalcMode] = useState<CalcMode>(calc0.manual ? 'manual' : calc0.autoNoTable ? 'autoNoTable' : 'auto');
  const [iterate, setIterate] = useState(!!calc0.iterate);
  const [maxIt, setMaxIt] = useState(String(calc0.maxIterations ?? 100));
  const [maxCh, setMaxCh] = useState(String(calc0.maxChange ?? 0.001));
  const [r1c1, setR1c1] = useState(!!calc0.r1c1);
  const [pad, setPad] = useState(!!calc0.precisionAsDisplayed);
  const [d1904, setD1904] = useState(store.wb.date1904);
  const [user, setUser] = useState(userName());
  const [theme, setTheme] = useState(store.view.theme);
  const [lang, setLang] = useState(getLang());
  const [rtlSheets, setRtlSheets] = useState<'auto' | 'on' | 'off'>(() => { try { const v = localStorage.getItem(RTL_SHEETS_KEY); return v === '1' ? 'on' : v === '0' ? 'off' : 'auto'; } catch { return 'auto'; } });
  const [systemSep, setSystemSep] = useState(!settings.decimalSep && !settings.thousandsSep);

  const ok = (): boolean => {
    const it = Number(maxIt), ch = Number(maxCh);
    if (iterate && (!Number.isInteger(it) || it < 1 || it > 32767)) { store.notify(T('Maximum iterations must be a whole number from 1 to 32767.')); setTab('formulas'); return false; }
    if (iterate && (!Number.isFinite(ch) || ch < 0)) { store.notify(T('Maximum change must be zero or more.')); setTab('formulas'); return false; }
    if (!systemSep && (s.decimalSep === s.thousandsSep || !s.decimalSep)) { store.notify(T('Choose different decimal and thousands separators.')); setTab('advanced'); return false; }
    if (s.defaultSize < 1 || s.defaultSize > 409) { store.notify(T('Font size must be from 1 to 409.')); return false; }
    const turnOnPad = pad && !calc0.precisionAsDisplayed;
    if (turnOnPad && !confirm(T('Numbers in this workbook will be permanently rounded to what their formats show.'))) return false;

    // Device settings.
    const next: AppSettings = { ...s, sheetCount: Math.max(1, Math.min(255, Math.round(s.sheetCount))), decimalSep: systemSep ? '' : s.decimalSep, thousandsSep: systemSep ? '' : s.thousandsSep };
    saveSettings(null, next);
    setUserName(user);
    setRtlSheetsDefault(rtlSheets === 'auto' ? null : rtlSheets === 'on');
    if (lang !== getLang()) void setInterfaceLang(lang, { persist: true });
    setThemeSetting(theme);
    store.view.theme = theme;
    applyDeviceSettings(store);

    // Workbook settings, saved in the file.
    const calc: NonNullable<Workbook['calc']> = { ...calc0, manual: calcMode === 'manual' || undefined, autoNoTable: calcMode === 'autoNoTable' || undefined, iterate: iterate || undefined,
      maxIterations: iterate ? it : calc0.maxIterations, maxChange: iterate ? ch : calc0.maxChange, r1c1: r1c1 || undefined, precisionAsDisplayed: pad || undefined };
    for (const k of Object.keys(calc) as (keyof typeof calc)[]) if (calc[k] === undefined) delete calc[k];
    const changedBook = JSON.stringify(calc) !== JSON.stringify(calc0) || d1904 !== store.wb.date1904;
    if (changedBook) {
      // One undo step, so undo brings the calculation settings and date system back too.
      store.change(N('Options'), 'book', () => {
        store.wb.calc = Object.keys(calc).length ? calc : undefined;
        store.wb.date1904 = d1904;
        applyWorkbookSettings(store);
        if (turnOnPad) roundConstants(store.wb, store);
        store.engine.recalc(true);
      });
    }
    store.touch();
    return true;
  };

  return (
    <Dialog title={T('Options')} close={close} onOk={ok} wide>
      <Tabs tabs={TABS} value={tab} onChange={setTab} />
      <div className="ws-options">
        {tab === 'general' && <>
          <Field label={T('Your name (shown on comments and notes)')}><input value={user} onChange={(e) => setUser(e.target.value)} maxLength={100} placeholder={T('You')} /></Field>
          <Field label={T('Theme')}><Radio name="theme" row value={theme} options={[['system', N('Match the system')], ['light', N('Light')], ['dark', N('Dark')]]} onChange={setTheme} /></Field>
          <Field label={T('Interface language')}>
            <select value={lang} onChange={(e) => setLang(e.target.value)}>{LANGUAGES.map(([v, l]) => <option key={v} value={v}>{Tr(l)}</option>)}</select>
          </Field>
          <Field label={T('New sheets run right to left')}>
            <select value={rtlSheets} onChange={(e) => setRtlSheets(e.target.value as 'auto' | 'on' | 'off')}>
              <option value="auto">{T('When the interface language does')}</option><option value="on">{T('Always')}</option><option value="off">{T('Never')}</option>
            </select>
          </Field>
          <h4>{T('New workbooks')}</h4>
          <div className="grid2">
            <Field label={T('Font')}><select value={s.defaultFont} onChange={(e) => set({ defaultFont: e.target.value })}>{(FONT_NAMES.includes(s.defaultFont) ? FONT_NAMES : [s.defaultFont, ...FONT_NAMES]).map((f) => <option key={f}>{f}</option>)}</select></Field>
            <Field label={T('Font size')}><input type="number" min={1} max={409} value={s.defaultSize} onChange={(e) => set({ defaultSize: Number(e.target.value) })} /></Field>
          </div>
          <Field label={T('Sheets in a new workbook')}><input type="number" min={1} max={255} value={s.sheetCount} onChange={(e) => set({ sheetCount: Number(e.target.value) })} /></Field>
        </>}
        {tab === 'formulas' && <>
          <h4>{T('Calculation (saved in this workbook)')}</h4>
          <Radio name="calc" value={calcMode} options={[['auto', N('Automatic')], ['autoNoTable', N('Automatic except for data tables')], ['manual', N('Manual (press F9 to calculate)')]]} onChange={setCalcMode} />
          <Check label={T('Iterative calculation')} checked={iterate} onChange={setIterate} />
          <div className="grid2">
            <Field label={T('Maximum iterations')}><input type="number" min={1} max={32767} disabled={!iterate} value={maxIt} onChange={(e) => setMaxIt(e.target.value)} /></Field>
            <Field label={T('Maximum change')}><input type="number" min={0} step="any" disabled={!iterate} value={maxCh} onChange={(e) => setMaxCh(e.target.value)} /></Field>
          </div>
          <Check label={T('Set precision as displayed')} checked={pad} onChange={setPad} />
          <Check label={T('Use the 1904 date system')} checked={d1904} onChange={setD1904} />
          <h4>{T('Working with formulas')}</h4>
          <Check label={T('R1C1 reference style')} checked={r1c1} onChange={setR1c1} />
          <Check label={T('Formula autocomplete')} checked={s.formulaAutocomplete} onChange={(v) => set({ formulaAutocomplete: v })} />
          <Check label={T('Show function argument tips')} checked={s.functionTips} onChange={(v) => set({ functionTips: v })} />
        </>}
        {tab === 'proofing' && <>
          <Check label={T('Check for errors in the background')} checked={s.errorChecking} onChange={(v) => set({ errorChecking: v })} />
          <Field label={T('Indicator color')}><ColorPick value={s.errorColor} onChange={(c) => set({ errorColor: c ?? '#1e9e4a' })} /></Field>
          <h4>{T('Rules')}</h4>
          <div className="b-list">
            {(Object.keys(RULE_TEXT) as ErrorRule[]).map((r) => <Check key={r} label={Tr(RULE_TEXT[r].label)} disabled={!s.errorChecking} checked={s.rules[r]} onChange={(v) => set({ rules: { ...s.rules, [r]: v } })} />)}
          </div>
          <button className="sm" onClick={() => { resetIgnored(store); store.notify(T('Ignored errors will show again.')); }}>{T('Reset ignored errors')}</button>
        </>}
        {tab === 'advanced' && <>
          <h4>{T('Editing')}</h4>
          <Check label={T('After pressing Enter, move the selection')} checked={s.enterMove} onChange={(v) => set({ enterMove: v })} />
          <Field label={T('Direction')}><select disabled={!s.enterMove} value={s.enterDirection} onChange={(e) => set({ enterDirection: e.target.value as AppSettings['enterDirection'] })}>
            <option value="down">{T('Down')}</option><option value="right">{T('Right')}</option><option value="up">{T('Up')}</option><option value="left">{T('Left')}</option></select></Field>
          <Check label={T('AutoComplete cell values from the column')} checked={s.autocomplete} onChange={(v) => set({ autocomplete: v })} />
          <Check label={T('Flash Fill automatically')} checked={s.flashFillAuto} onChange={(v) => set({ flashFillAuto: v })} />
          <h4>{T('Numbers')}</h4>
          <Check label={T('Use the region\'s separators')} checked={systemSep} onChange={setSystemSep} />
          <div className="grid2">
            <Field label={T('Decimal separator')}><input maxLength={1} disabled={systemSep} value={systemSep ? '' : s.decimalSep} placeholder="." onChange={(e) => set({ decimalSep: e.target.value })} /></Field>
            <Field label={T('Thousands separator')}><input maxLength={1} disabled={systemSep} value={systemSep ? '' : s.thousandsSep} placeholder="," onChange={(e) => set({ thousandsSep: e.target.value })} /></Field>
          </div>
          <h4>{T('Display')}</h4>
          <Check label={T('Show the formula bar')} checked={s.showFormulaBar} onChange={(v) => set({ showFormulaBar: v })} />
          <Check label={T('Show the status bar')} checked={s.showStatusBar} onChange={(v) => set({ showStatusBar: v })} />
          <Field label={T('Status bar shows')}>
            <div className="b-list b-checklist">{STATUS_ITEMS.map(([k, l]) => <Check key={k} label={Tr(l)} checked={s.statusItems.includes(k)} onChange={(v) => set({ statusItems: v ? [...s.statusItems, k] : s.statusItems.filter((x) => x !== k) })} />)}</div>
          </Field>
          <Field label={T('Focus cell color')}><ColorPick value={s.focusColor} onChange={(c) => set({ focusColor: c ?? '#f5c542' })} /></Field>
        </>}
        {tab === 'save' && <>
          <Check label={T('Keep version history on this device')} checked={s.versionsOn} onChange={(v) => set({ versionsOn: v })} />
          <div className="grid2">
            <Field label={T('Save a version every (minutes)')}><input type="number" min={1} max={240} disabled={!s.versionsOn} value={s.versionsEvery} onChange={(e) => set({ versionsEvery: Math.max(1, Number(e.target.value) || 5) })} /></Field>
            <Field label={T('Versions to keep per file')}><input type="number" min={1} max={500} disabled={!s.versionsOn} value={s.versionsKeep} onChange={(e) => set({ versionsKeep: Math.max(1, Math.min(500, Number(e.target.value) || 50)) })} /></Field>
          </div>
          <p className="muted small">{T('Versions are also saved each time you save. Named versions are always kept. Encrypted workbooks don\'t keep versions.')}</p>
        </>}
      </div>
    </Dialog>
  );
};

/** Applies device settings to the shared modules that read them. */
export const applyDeviceSettings = (store: GridStore): void => {
  applySettings(store);
  assistPrefs.autocomplete = settings.formulaAutocomplete;
  assistPrefs.tips = settings.functionTips;
  statusPrefs.items = settings.statusItems.filter((x): x is StatusItem => STATUS_ITEMS.some(([k]) => k === x));
  setSeparators(settings.decimalSep, settings.thousandsSep);
};

/** Gives the empty start-up or new workbook the default font, size and sheet count from Options. */
export const applyNewWorkbookDefaults = (wb: Workbook): void => {
  for (const sh of wb.sheets) sh.view.rtl = userSheetDefaults.rtl;
  if (settings.defaultFont !== 'Calibri' || settings.defaultSize !== 11) wb.styles[0] = { ...wb.styles[0], font: { ...wb.styles[0]?.font, name: settings.defaultFont, size: settings.defaultSize } };
  for (let i = wb.sheets.length; i < settings.sheetCount; i++) wb.sheets.push(newUserSheet(`Sheet${i + 1}`));
};
