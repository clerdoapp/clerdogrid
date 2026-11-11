import { useMemo, useState } from 'react';
import { ArrowDown, ArrowUp, Pencil, Plus, Trash2 } from 'lucide-react';
import { ICON_SETS } from '../../app/cf';
import { parseFormula } from '../../engine/parser';
import type { CellStyle, CfOperator, CfRule, CfValue, ConditionalFormat } from '../../model/workbook';
import { Dialog, Field, type DialogProps } from './Dialog';
import { Check, ColorPick, cssOf, overlaps, parseRefs, rangeLabel } from './shared';
import { N, T, Tr } from '../i18n';

// ---------- Rule model used by the editor ----------

type Kind = 'cell' | 'text' | 'date' | 'dup' | 'blank' | 'error' | 'top' | 'avg' | 'bar' | 'scale2' | 'scale3' | 'icons' | 'formula';
const KINDS: [Kind, string][] = [
  ['cell', N('Cell value')], ['text', N('Text')], ['date', N('Dates')], ['dup', N('Duplicate or unique')], ['blank', N('Blanks')], ['error', N('Errors')],
  ['top', N('Top or bottom')], ['avg', N('Above or below average')], ['bar', N('Data bar')], ['scale2', '2-color scale'], ['scale3', '3-color scale'], ['icons', N('Icon set')], ['formula', N('Use a formula')],
];
const OPS: [CfOperator, string][] = [
  ['greaterThan', 'greater than'], ['lessThan', 'less than'], ['between', 'between'], ['notBetween', 'not between'], ['equal', 'equal to'], ['notEqual', 'not equal to'],
  ['greaterThanOrEqual', 'greater than or equal to'], ['lessThanOrEqual', 'less than or equal to'],
];
const TEXT_OPS: ['containsText' | 'notContainsText' | 'beginsWith' | 'endsWith', string][] = [['containsText', 'contains'], ['notContainsText', 'does not contain'], ['beginsWith', 'begins with'], ['endsWith', 'ends with']];
type Period = Extract<CfRule, { type: 'timePeriod' }>['period'];
const PERIODS: [Period, string][] = [['yesterday', N('Yesterday')], ['today', N('Today')], ['tomorrow', N('Tomorrow')], ['last7Days', N('In the last 7 days')], ['lastWeek', N('Last week')], ['thisWeek', N('This week')], ['nextWeek', N('Next week')], ['lastMonth', N('Last month')], ['thisMonth', N('This month')], ['nextMonth', N('Next month')]];
const VALUE_TYPES: [CfValue['type'], string][] = [['min', N('Lowest value')], ['max', N('Highest value')], ['num', N('Number')], ['percent', N('Percent')], ['percentile', N('Percentile')], ['formula', N('Formula')]];

export const FORMAT_PRESETS: { label: string; style: CellStyle }[] = [
  { label: N('Light red fill, dark red text'), style: { font: { color: '#9c0006' }, fill: { color: '#ffc7ce' } } },
  { label: N('Yellow fill, dark yellow text'), style: { font: { color: '#9c5700' }, fill: { color: '#ffeb9c' } } },
  { label: N('Green fill, dark green text'), style: { font: { color: '#006100' }, fill: { color: '#c6efce' } } },
  { label: N('Light red fill'), style: { fill: { color: '#ffc7ce' } } },
  { label: N('Red text'), style: { font: { color: '#9c0006' } } },
  { label: N('Red border'), style: { border: { top: { style: 'thin', color: '#c00000' }, bottom: { style: 'thin', color: '#c00000' }, left: { style: 'thin', color: '#c00000' }, right: { style: 'thin', color: '#c00000' } } } },
  { label: N('Bold blue text'), style: { font: { bold: true, color: '#0C7A4B' } } },
];

interface Draft {
  kind: Kind;
  op: CfOperator; v1: string; v2: string;
  textOp: (typeof TEXT_OPS)[number][0]; text: string;
  period: Period;
  unique: boolean;
  notBlank: boolean; notError: boolean;
  bottom: boolean; percent: boolean; rank: number;
  below: boolean; equal: boolean; stdDev: number;
  barColor: string; gradient: boolean; barMin: CfValue; barMax: CfValue; showValue: boolean;
  scale: { value: CfValue; color: string }[];
  iconSet: string; iconReverse: boolean; iconValues: CfValue[];
  formula: string;
  style: CellStyle;
  stop: boolean;
}

const iconDefaults = (set: string): CfValue[] => {
  const n = ICON_SETS[set] ?? 3;
  return Array.from({ length: n }, (_, i) => ({ type: 'percent' as const, value: String(Math.round((i * 100) / n)) }));
};

const emptyDraft = (): Draft => ({
  kind: 'cell', op: 'greaterThan', v1: '', v2: '', textOp: 'containsText', text: '', period: 'today', unique: false, notBlank: false, notError: false,
  bottom: false, percent: false, rank: 10, below: false, equal: false, stdDev: 0,
  barColor: '#0C7A4B', gradient: true, barMin: { type: 'min' }, barMax: { type: 'max' }, showValue: true,
  scale: [{ value: { type: 'min' }, color: '#f8696b' }, { value: { type: 'percentile', value: '50' }, color: '#ffeb84' }, { value: { type: 'max' }, color: '#63be7b' }],
  iconSet: '3Arrows', iconReverse: false, iconValues: iconDefaults('3Arrows'),
  formula: '', style: FORMAT_PRESETS[0]!.style, stop: false,
});

const toDraft = (rule: CfRule): Draft => {
  const d = emptyDraft();
  if ('style' in rule) d.style = rule.style;
  if ('stop' in rule) d.stop = !!rule.stop;
  switch (rule.type) {
    case 'cellIs': return { ...d, kind: 'cell', op: rule.operator, v1: rule.formulas[0] ?? '', v2: rule.formulas[1] ?? '' };
    case 'expression': return { ...d, kind: 'formula', formula: rule.formula };
    case 'containsText': case 'notContainsText': case 'beginsWith': case 'endsWith': return { ...d, kind: 'text', textOp: rule.type, text: rule.text };
    case 'timePeriod': return { ...d, kind: 'date', period: rule.period };
    case 'duplicateValues': case 'uniqueValues': return { ...d, kind: 'dup', unique: rule.type === 'uniqueValues' };
    case 'containsBlanks': case 'notContainsBlanks': return { ...d, kind: 'blank', notBlank: rule.type === 'notContainsBlanks' };
    case 'containsErrors': case 'notContainsErrors': return { ...d, kind: 'error', notError: rule.type === 'notContainsErrors' };
    case 'top10': return { ...d, kind: 'top', rank: rule.rank, bottom: !!rule.bottom, percent: !!rule.percent };
    case 'aboveAverage': return { ...d, kind: 'avg', below: !!rule.below, equal: !!rule.equal, stdDev: rule.stdDev ?? 0 };
    case 'dataBar': return { ...d, kind: 'bar', barColor: rule.color, gradient: rule.gradient !== false, barMin: rule.min, barMax: rule.max, showValue: rule.showValue !== false };
    case 'colorScale': return { ...d, kind: rule.stops.length === 2 ? 'scale2' : 'scale3', scale: rule.stops.length === 2 ? [rule.stops[0]!, d.scale[1]!, rule.stops[1]!] : rule.stops.slice(0, 3) };
    case 'iconSet': return { ...d, kind: 'icons', iconSet: rule.set, iconReverse: !!rule.reverse, iconValues: rule.values, showValue: rule.showValue !== false };
  }
};

const fx = (s: string): string => s.trim().replace(/^=/, '');
/** A plain number or "=formula" becomes formula text; anything else is a quoted text constant. */
const operand = (s: string): string => {
  const t = s.trim();
  if (t.startsWith('=')) return t.slice(1);
  if (t !== '' && !Number.isNaN(Number(t))) return t;
  return `"${t.replace(/"/g, '""')}"`;
};
const fromOperand = (s: string): string => (/^".*"$/s.test(s) ? s.slice(1, -1).replace(/""/g, '"') : Number.isNaN(Number(s)) ? `=${s}` : s);

const toRule = (d: Draft): CfRule | string => {
  const stop = d.stop || undefined;
  const style = d.style;
  const checkFormula = (f: string) => { try { parseFormula(f); return null; } catch { return T('That formula has a problem. Check the brackets and quotes.'); } };
  switch (d.kind) {
    case 'cell': {
      if (!d.v1.trim()) return T('Enter a value to compare with.');
      const two = d.op === 'between' || d.op === 'notBetween';
      if (two && !d.v2.trim()) return T('Enter the second value.');
      const formulas = two ? [operand(d.v1), operand(d.v2)] : [operand(d.v1)];
      for (const f of formulas) { const e = checkFormula(f); if (e) return e; }
      return { type: 'cellIs', operator: d.op, formulas, style, stop };
    }
    case 'text': return d.text ? { type: d.textOp, text: d.text, style, stop } : T('Enter the text to look for.');
    case 'date': return { type: 'timePeriod', period: d.period, style, stop };
    case 'dup': return { type: d.unique ? 'uniqueValues' : 'duplicateValues', style, stop };
    case 'blank': return { type: d.notBlank ? 'notContainsBlanks' : 'containsBlanks', style, stop };
    case 'error': return { type: d.notError ? 'notContainsErrors' : 'containsErrors', style, stop };
    case 'top': return d.rank >= 1 && d.rank <= (d.percent ? 100 : 1000) ? { type: 'top10', rank: Math.trunc(d.rank), bottom: d.bottom || undefined, percent: d.percent || undefined, style, stop } : T('Enter a number from 1 to 1000 (or 1 to 100 for percent).');
    case 'avg': return { type: 'aboveAverage', below: d.below || undefined, equal: d.equal || undefined, stdDev: d.stdDev || undefined, style, stop };
    case 'bar': return { type: 'dataBar', min: d.barMin, max: d.barMax, color: d.barColor, gradient: d.gradient, showValue: d.showValue ? undefined : false };
    case 'scale2': return { type: 'colorScale', stops: [d.scale[0]!, d.scale[2]!] };
    case 'scale3': return { type: 'colorScale', stops: d.scale.slice(0, 3) };
    case 'icons': return { type: 'iconSet', set: d.iconSet, values: d.iconValues, reverse: d.iconReverse || undefined, showValue: d.showValue ? undefined : false };
    case 'formula': {
      const f = fx(d.formula);
      if (!f) return 'Enter a formula that returns TRUE or FALSE, like =A1>100.';
      return checkFormula(f) ?? { type: 'expression', formula: f, style, stop };
    }
  }
};

export const describeRule = (rule: CfRule): string => {
  const op = (o: CfOperator) => Tr(OPS.find(([v]) => v === o)?.[1] ?? o);
  switch (rule.type) {
    case 'cellIs': return rule.formulas[1] !== undefined ? T('Cell value {op} {a} and {b}', { op: op(rule.operator), a: fromOperand(rule.formulas[0] ?? ''), b: fromOperand(rule.formulas[1]) }) : T('Cell value {op} {a}', { op: op(rule.operator), a: fromOperand(rule.formulas[0] ?? '') });
    case 'expression': return T('Formula: {formula}', { formula: `=${rule.formula}` });
    case 'containsText': case 'notContainsText': case 'beginsWith': case 'endsWith': return T('Text {test} "{text}"', { test: Tr(TEXT_OPS.find(([v]) => v === rule.type)?.[1]), text: rule.text });
    case 'timePeriod': return T('Date: {period}', { period: Tr(PERIODS.find(([v]) => v === rule.period)?.[1]) });
    case 'duplicateValues': return T('Duplicate values');
    case 'uniqueValues': return T('Unique values');
    case 'containsBlanks': return T('Blank cells');
    case 'notContainsBlanks': return T('Cells with content');
    case 'containsErrors': return T('Cells with errors');
    case 'notContainsErrors': return T('Cells without errors');
    case 'top10': return T(rule.bottom ? 'Bottom {rank}' : 'Top {rank}', { rank: `${rule.rank}${rule.percent ? '%' : ''}` });
    case 'aboveAverage': return T(rule.below ? (rule.equal ? 'Below or equal to average' : 'Below average') : rule.equal ? 'Above or equal to average' : 'Above average') + (rule.stdDev ? ` (${T('{n} std dev', { n: rule.stdDev })})` : '');
    case 'dataBar': return T('Data bar');
    case 'colorScale': return T('{n}-color scale', { n: rule.stops.length });
    case 'iconSet': return T('Icon set: {name}', { name: iconLabel(rule.set) });
  }
};

const iconLabel = (set: string): string => set.replace(/^(\d)/, '$1 ').replace(/([a-z])([A-Z0-9])/g, '$1 $2').toLowerCase();

const RuleSwatch = ({ rule }: { rule: CfRule }) => {
  if (rule.type === 'dataBar') return <span className="b-cfswatch" style={{ background: rule.gradient === false ? rule.color : `linear-gradient(90deg, ${rule.color}, transparent)` }} />;
  if (rule.type === 'colorScale') return <span className="b-cfswatch" style={{ background: `linear-gradient(90deg, ${rule.stops.map((s) => s.color).join(', ')})` }} />;
  if (rule.type === 'iconSet') return <span className="b-cfswatch b-cficon">{ICON_SETS[rule.set] ?? 3}</span>;
  return <span className="b-cfswatch" style={cssOf(rule.style, false)}>{T('AaBb')}</span>;
};

// ---------- Editor pieces ----------

const ValuePicker = ({ label, v, onChange, allowMinMax = true }: { label: string; v: CfValue; onChange: (v: CfValue) => void; allowMinMax?: boolean }) => (
  <div className="b-valuepick">
    <Field label={Tr(label)}>
      <select value={v.type} onChange={(e) => onChange({ type: e.target.value as CfValue['type'], value: e.target.value === 'min' || e.target.value === 'max' ? undefined : v.value ?? '0' })}>
        {VALUE_TYPES.filter(([t]) => allowMinMax || (t !== 'min' && t !== 'max')).map(([t, l]) => <option key={t} value={t}>{Tr(l)}</option>)}
      </select>
    </Field>
    {v.type !== 'min' && v.type !== 'max' && <Field label={T('Value')}><input value={v.value ?? ''} onChange={(e) => onChange({ ...v, value: e.target.value })} /></Field>}
  </div>
);

const StylePicker = ({ style, onChange }: { style: CellStyle; onChange: (s: CellStyle) => void }) => {
  const [custom, setCustom] = useState(!FORMAT_PRESETS.some((p) => JSON.stringify(p.style) === JSON.stringify(style)));
  const f = style.font ?? {};
  return (
    <div className="b-col">
      <Field label={T('Format with')}>
        <select value={custom ? 'custom' : String(FORMAT_PRESETS.findIndex((p) => JSON.stringify(p.style) === JSON.stringify(style)))} onChange={(e) => {
          if (e.target.value === 'custom') { setCustom(true); return; }
          setCustom(false); onChange(FORMAT_PRESETS[+e.target.value]!.style);
        }}>
          {FORMAT_PRESETS.map((p, i) => <option key={p.label} value={i}>{Tr(p.label)}</option>)}
          <option value="custom">{T('Custom format')}</option>
        </select>
      </Field>
      {custom && <>
        <div className="b-checks">
          <Check label={T('Bold')} checked={!!f.bold} onChange={(v) => onChange({ ...style, font: { ...f, bold: v || undefined } })} />
          <Check label={T('Italic')} checked={!!f.italic} onChange={(v) => onChange({ ...style, font: { ...f, italic: v || undefined } })} />
          <Check label={T('Strikethrough')} checked={!!f.strike} onChange={(v) => onChange({ ...style, font: { ...f, strike: v || undefined } })} />
        </div>
        <Field label={T('Text color')}><ColorPick value={f.color} none={T('Automatic')} onChange={(c) => onChange({ ...style, font: { ...f, color: c } })} /></Field>
        <Field label={T('Fill')}><ColorPick value={style.fill?.color} none={T('No fill')} onChange={(c) => onChange({ ...style, fill: c ? { color: c } : undefined })} /></Field>
      </>}
      <div className="preview"><span style={{ ...cssOf(style, false), padding: '2px 8px' }}>{T('AaBbCc 123')}</span></div>
    </div>
  );
};

const RuleEditor = ({ d, set }: { d: Draft; set: (p: Partial<Draft>) => void }) => {
  const usesStyle = !['bar', 'scale2', 'scale3', 'icons'].includes(d.kind);
  return (
    <div className="b-col">
      <Field label={T('Rule type')}><select value={d.kind} onChange={(e) => set({ kind: e.target.value as Kind })}>{KINDS.map(([k, l]) => <option key={k} value={k}>{Tr(l)}</option>)}</select></Field>
      {d.kind === 'cell' && <div className="b-inline">
        <select value={d.op} onChange={(e) => set({ op: e.target.value as CfOperator })} aria-label={T('Comparison')}>{OPS.map(([v, l]) => <option key={v} value={v}>{Tr(l)}</option>)}</select>
        <input value={d.v1} onChange={(e) => set({ v1: e.target.value })} placeholder={T('Value or =formula')} aria-label={T('Value')} />
        {(d.op === 'between' || d.op === 'notBetween') && <><span className="muted">{T('and')}</span><input value={d.v2} onChange={(e) => set({ v2: e.target.value })} placeholder={T('Value')} aria-label={T('Second value')} /></>}
      </div>}
      {d.kind === 'text' && <div className="b-inline">
        <select value={d.textOp} onChange={(e) => set({ textOp: e.target.value as Draft['textOp'] })} aria-label={T('Text test')}>{TEXT_OPS.map(([v, l]) => <option key={v} value={v}>{Tr(l)}</option>)}</select>
        <input value={d.text} onChange={(e) => set({ text: e.target.value })} placeholder={T('Text')} aria-label={T('Text')} />
      </div>}
      {d.kind === 'date' && <Field label={T('Date is')}><select value={d.period} onChange={(e) => set({ period: e.target.value as Period })}>{PERIODS.map(([v, l]) => <option key={v} value={v}>{Tr(l)}</option>)}</select></Field>}
      {d.kind === 'dup' && <Field label={T('Highlight')}><select value={d.unique ? 'u' : 'd'} onChange={(e) => set({ unique: e.target.value === 'u' })}><option value="d">{T('Duplicate values')}</option><option value="u">{T('Unique values')}</option></select></Field>}
      {d.kind === 'blank' && <Field label={T('Highlight')}><select value={d.notBlank ? 'n' : 'b'} onChange={(e) => set({ notBlank: e.target.value === 'n' })}><option value="b">{T('Blank cells')}</option><option value="n">{T('Cells that aren\'t blank')}</option></select></Field>}
      {d.kind === 'error' && <Field label={T('Highlight')}><select value={d.notError ? 'n' : 'e'} onChange={(e) => set({ notError: e.target.value === 'n' })}><option value="e">{T('Cells with errors')}</option><option value="n">{T('Cells without errors')}</option></select></Field>}
      {d.kind === 'top' && <div className="b-inline">
        <select value={d.bottom ? 'b' : 't'} onChange={(e) => set({ bottom: e.target.value === 'b' })} aria-label={T('Top or bottom')}><option value="t">{T('Top')}</option><option value="b">{T('Bottom')}</option></select>
        <input type="number" min={1} max={1000} value={d.rank} onChange={(e) => set({ rank: Math.trunc(+e.target.value) })} aria-label={T('How many')} />
        <Check label={T('Percent')} checked={d.percent} onChange={(v) => set({ percent: v })} />
      </div>}
      {d.kind === 'avg' && <div className="grid2">
        <Field label={T('Values that are')}><select value={`${d.below ? 'b' : 'a'}${d.equal ? 'e' : ''}`} onChange={(e) => set({ below: e.target.value.startsWith('b'), equal: e.target.value.endsWith('e') })}>
          <option value="a">{T('Above average')}</option><option value="b">{T('Below average')}</option><option value="ae">{T('Equal or above average')}</option><option value="be">{T('Equal or below average')}</option>
        </select></Field>
        <Field label={T('Standard deviations')}><select value={d.stdDev} onChange={(e) => set({ stdDev: +e.target.value })}><option value={0}>{T('None')}</option><option value={1}>1</option><option value={2}>2</option><option value={3}>3</option></select></Field>
      </div>}
      {d.kind === 'bar' && <>
        <div className="grid2"><ValuePicker label={T('Shortest bar')} v={d.barMin} onChange={(v) => set({ barMin: v })} /><ValuePicker label={T('Longest bar')} v={d.barMax} onChange={(v) => set({ barMax: v })} /></div>
        <Field label={T('Bar color')}><ColorPick value={d.barColor} onChange={(c) => set({ barColor: c ?? '#0C7A4B' })} /></Field>
        <div className="b-checks"><Check label={T('Gradient fill')} checked={d.gradient} onChange={(v) => set({ gradient: v })} /><Check label={T('Show the value')} checked={d.showValue} onChange={(v) => set({ showValue: v })} /></div>
      </>}
      {(d.kind === 'scale2' || d.kind === 'scale3') && <div className={d.kind === 'scale3' ? 'grid3' : 'grid2'}>
        {(d.kind === 'scale2' ? [0, 2] : [0, 1, 2]).map((i) => (
          <div key={i} className="b-col">
            <ValuePicker label={i === 0 ? T('Lowest') : i === 2 ? T('Highest') : T('Middle')} v={d.scale[i]!.value} allowMinMax={i !== 1} onChange={(v) => set({ scale: d.scale.map((s, k) => (k === i ? { ...s, value: v } : s)) })} />
            <input type="color" aria-label={T('Color')} value={d.scale[i]!.color} onChange={(e) => set({ scale: d.scale.map((s, k) => (k === i ? { ...s, color: e.target.value } : s)) })} />
          </div>
        ))}
        <div className="b-scaleprev" style={{ gridColumn: '1 / -1', background: `linear-gradient(90deg, ${(d.kind === 'scale2' ? [d.scale[0]!, d.scale[2]!] : d.scale).map((s) => s.color).join(', ')})` }} />
      </div>}
      {d.kind === 'icons' && <>
        <Field label={T('Icon style')}><select value={d.iconSet} onChange={(e) => set({ iconSet: e.target.value, iconValues: iconDefaults(e.target.value) })}>{Object.keys(ICON_SETS).map((k) => <option key={k} value={k}>{iconLabel(k)}</option>)}</select></Field>
        <div className="b-col">
          {d.iconValues.map((v, i) => i === 0 ? null : (
            <div key={i} className="b-inline">
              <span className="muted small">{T('Icon {n} when value ≥', { n: d.iconValues.length - i })}</span>
              <input value={v.value ?? ''} onChange={(e) => set({ iconValues: d.iconValues.map((x, k) => (k === i ? { ...x, value: e.target.value } : x)) })} aria-label={T('Threshold')} />
              <select value={v.type} aria-label={T('Type')} onChange={(e) => set({ iconValues: d.iconValues.map((x, k) => (k === i ? { ...x, type: e.target.value as CfValue['type'] } : x)) })}>
                <option value="percent">{T('Percent')}</option><option value="num">{T('Number')}</option><option value="percentile">{T('Percentile')}</option><option value="formula">{T('Formula')}</option>
              </select>
            </div>
          ))}
        </div>
        <div className="b-checks"><Check label={T('Reverse icon order')} checked={d.iconReverse} onChange={(v) => set({ iconReverse: v })} /><Check label={T('Show the value')} checked={d.showValue} onChange={(v) => set({ showValue: v })} /></div>
      </>}
      {d.kind === 'formula' && <Field label={T('Format cells where this formula is true')}><input className="mono" value={d.formula} onChange={(e) => set({ formula: e.target.value })} placeholder="=$B2>100" /></Field>}
      {d.kind === 'formula' && <p className="muted small">{T('Write it for the top-left cell of the range; relative references move with each cell.')}</p>}
      {usesStyle && <StylePicker style={d.style} onChange={(s) => set({ style: s })} />}
      {usesStyle && <Check label={T('Stop checking other rules when this one matches')} checked={d.stop} onChange={(v) => set({ stop: v })} />}
    </div>
  );
};

// ---------- Manager ----------

interface Row { cf: ConditionalFormat; ruleIndex: number; rule: CfRule }

/** One ConditionalFormat per rule keeps priority and ranges independent. */
const flatten = (cfs: ConditionalFormat[]): ConditionalFormat[] => {
  const out: ConditionalFormat[] = [];
  const sorted = [...cfs].map((c, i) => ({ c, i })).sort((a, b) => (a.c.priority ?? a.i) - (b.c.priority ?? b.i));
  for (const { c } of sorted) for (const r of c.rules) out.push({ ranges: c.ranges.map((x) => ({ ...x })), rules: [r] });
  return out.map((c, i) => ({ ...c, priority: i + 1 }));
};

export const ConditionalDialog = ({ store, close, data }: DialogProps<{ kind?: Kind }>) => {
  const [list, setList] = useState<ConditionalFormat[]>(() => flatten(structuredClone(store.sheet.conditionalFormats)));
  const [scope, setScope] = useState<'selection' | 'sheet'>('selection');
  const [editing, setEditing] = useState<{ index: number | null; draft: Draft; ref: string } | null>(() => data?.kind ? { index: null, draft: { ...emptyDraft(), kind: data.kind }, ref: store.sel.ranges.map(rangeLabel).join(',') } : null);
  const [picked, setPicked] = useState<number | null>(null);
  const sel = store.sel.ranges;
  const rows = useMemo(() => list.map((cf, i) => ({ cf, i })).filter(({ cf }) => scope === 'sheet' || cf.ranges.some((r) => sel.some((s) => overlaps(r, s)))), [list, scope, sel]);
  const dirty = useMemo(() => JSON.stringify(list) !== JSON.stringify(flatten(store.sheet.conditionalFormats)), [list, store.sheet.conditionalFormats]);

  const save = (next: ConditionalFormat[]) => {
    store.change(N('Conditional formatting'), 'sheet', () => { store.sheet.conditionalFormats = next.map((c, i) => ({ ...c, priority: i + 1 })); });
  };

  const commitEditor = (): boolean => {
    if (!editing) return true;
    const rule = toRule(editing.draft);
    if (typeof rule === 'string') { store.notify(rule); return false; }
    const refs = parseRefs(store, editing.ref);
    if (!refs || refs.sheet !== store.sheetIndex) { store.notify(T('Enter the cells the rule applies to, like A1:A20 (on this sheet).')); return false; }
    const cf: ConditionalFormat = { ranges: refs.ranges, rules: [rule] };
    const next = editing.index === null ? [cf, ...list] : list.map((c, i) => (i === editing.index ? cf : c));
    setList(next);
    setEditing(null);
    return true;
  };

  const move = (i: number, d: -1 | 1) => {
    const j = i + d;
    if (j < 0 || j >= list.length) return;
    const next = [...list];
    [next[i], next[j]] = [next[j]!, next[i]!];
    setList(next);
    setPicked(j);
  };

  if (editing) {
    return (
      <Dialog title={editing.index === null ? T('New rule') : T('Edit rule')} close={close} wide okLabel={T('Done')} onOk={() => {
        if (!commitEditor()) return false;
        // Opened straight into a new rule: save right away.
        if (data?.kind && editing.index === null) {
          const rule = toRule(editing.draft); const refs = parseRefs(store, editing.ref);
          if (typeof rule !== 'string' && refs) save([{ ranges: refs.ranges, rules: [rule] }, ...list]);
          return true;
        }
        return false;
      }} footer={data?.kind ? undefined : <button onClick={() => setEditing(null)}>{T('Back')}</button>}>
        <Field label={T('Applies to')}><input value={editing.ref} onChange={(e) => setEditing({ ...editing, ref: e.target.value })} placeholder="A1:A20" /></Field>
        <RuleEditor d={editing.draft} set={(p) => setEditing({ ...editing, draft: { ...editing.draft, ...p } })} />
      </Dialog>
    );
  }

  const ruleRows: Row[] = rows.map(({ cf }) => ({ cf, ruleIndex: 0, rule: cf.rules[0]! }));
  return (
    <Dialog title={T('Conditional formatting')} close={close} wide okLabel={T('Apply')} onOk={() => { if (dirty) save(list); }}
      footer={<button onClick={() => setEditing({ index: null, draft: emptyDraft(), ref: sel.map(rangeLabel).join(',') })}><Plus size={15} strokeWidth={1.9} />{' '}{T('New rule')}</button>}>
      <div className="b-inline">
        <span className="muted small">{T('Show rules for')}</span>
        <select value={scope} onChange={(e) => setScope(e.target.value as 'selection' | 'sheet')}><option value="selection">{T('The selection')}</option><option value="sheet">{T('This sheet')}</option></select>
      </div>
      <div className="b-table b-cftable" role="table">
        <div className="b-tr b-th" role="row"><span>{T('Rule (first applies first)')}</span><span>{T('Format')}</span><span>{T('Applies to')}</span><span>{T('Stop')}</span><span /></div>
        {ruleRows.map(({ cf, rule }, k) => {
          const i = rows[k]!.i;
          return (
            <div key={i} role="row" className={`b-tr${picked === i ? ' on' : ''}`} onClick={() => setPicked(i)}>
              <span>{describeRule(rule)}</span>
              <span><RuleSwatch rule={rule} /></span>
              <span><input key={cf.ranges.map(rangeLabel).join(',')} aria-label={T('Applies to')} defaultValue={cf.ranges.map(rangeLabel).join(',')} onBlur={(e) => {
                const refs = parseRefs(store, e.target.value);
                if (!refs || refs.sheet !== store.sheetIndex) { store.notify(T('Enter a range on this sheet, like A1:A20.')); e.target.value = cf.ranges.map(rangeLabel).join(','); return; }
                setList(list.map((c, x) => (x === i ? { ...c, ranges: refs.ranges } : c)));
              }} /></span>
              <span>{'style' in rule ? <input type="checkbox" aria-label={T('Stop if true')} checked={!!rule.stop} onChange={(e) => setList(list.map((c, x) => (x === i ? { ...c, rules: [{ ...rule, stop: e.target.checked || undefined }] } : c)))} /> : null}</span>
              <span className="b-rowtools">
                <button className="ico sm" aria-label={T('Move up')} title={T('Move up')} disabled={i === 0} onClick={(e) => { e.stopPropagation(); move(i, -1); }}><ArrowUp size={15} strokeWidth={1.9} /></button>
                <button className="ico sm" aria-label={T('Move down')} title={T('Move down')} disabled={i === list.length - 1} onClick={(e) => { e.stopPropagation(); move(i, 1); }}><ArrowDown size={15} strokeWidth={1.9} /></button>
                <button className="ico sm" aria-label={T('Edit rule')} title={T('Edit')} onClick={(e) => { e.stopPropagation(); setEditing({ index: i, draft: toDraft(rule), ref: cf.ranges.map(rangeLabel).join(',') }); }}><Pencil size={15} strokeWidth={1.9} /></button>
                <button className="ico sm" aria-label={T('Delete rule')} title={T('Delete')} onClick={(e) => { e.stopPropagation(); setList(list.filter((_, x) => x !== i)); setPicked(null); }}><Trash2 size={15} strokeWidth={1.9} /></button>
              </span>
            </div>
          );
        })}
        {!ruleRows.length && <div className="muted pad">{scope === 'selection' ? T('No rules for the selection yet. Add one with New rule.') : T('No rules on this sheet yet. Add one with New rule.')}</div>}
      </div>
      {dirty && <p className="muted small">{T('Changes are saved when you tap Apply.')}</p>}
    </Dialog>
  );
};

