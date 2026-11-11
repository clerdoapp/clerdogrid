import { useState } from 'react';
import type { GridStore } from '../../app/store';
import { Check } from '../../ui/dialogs/shared';
import { activeLocale, displayFormat, localizeFormula, localizeNumberText, LOCALES, setLocale, setNativeDigits, useNativeDigits } from '../../app/locale';
import { formatValue } from '../../formats/numfmt';
import { todaySerial } from '../../engine/date';
import { T, getLang } from '../../ui/i18n';

/** Day name in the interface language (0 = Sunday). */
const dayName = (d: number): string => new Intl.DateTimeFormat(getLang(), { weekday: 'long', timeZone: 'UTC' }).format(Date.UTC(2023, 0, 1 + d));

/** File > Region: how numbers, dates and formulas are typed and shown. */
export const RegionPanel = ({ store }: { store: GridStore }) => {
  const [, force] = useState(0);
  const loc = activeLocale();
  const apply = (f: () => void) => { f(); store.engine.recalc(true); store.touch(); force((n) => n + 1); };
  const today = todaySerial(store.wb.date1904);
  const sample = localizeNumberText(formatValue(1234567.89, '#,##0.00').text, '#,##0.00');
  const date = localizeNumberText(formatValue(today, displayFormat('m/d/yyyy')!).text, 'm/d/yyyy');
  const long = localizeNumberText(formatValue(today, '[$-F800]dddd, mmmm d, yyyy').text, 'dddd');
  return (
    <div className="region-panel">
      <label className="field">
        <select aria-label={T('Region')} value={loc.id} onChange={(e) => apply(() => setLocale(e.target.value, { persist: true }))}>
          {LOCALES.map((l) => <option key={l.id} value={l.id}>{l.name}</option>)}
        </select>
      </label>
      {loc.nativeDigits && <Check label={loc.nativeDigits === 'arab' ? T('Show Arabic-Indic digits') : T('Show native digits')} checked={useNativeDigits()} onChange={(v) => apply(() => setNativeDigits(v))} />}
      <dl className="region-sample">
        <dt>{T('Number')}</dt><dd>{sample}</dd>
        <dt>{T('Date')}</dt><dd>{date}</dd>
        <dt>{T('Long date')}</dt><dd>{long}</dd>
        <dt>{T('Formula')}</dt><dd><code>={localizeFormula('ROUND(A1*1.5,2)')}</code></dd>
        <dt>{T('Currency')}</dt><dd>{loc.currency}</dd>
        <dt>{T('Week starts')}</dt><dd>{dayName(loc.firstDay)}</dd>
      </dl>
    </div>
  );
};
