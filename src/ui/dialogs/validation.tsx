import { useState } from 'react';
import { parseFormula } from '../../engine/parser';
import type { CfOperator, DataValidation } from '../../model/workbook';
import { Dialog, Field, type DialogProps } from './Dialog';
import { Check, overlaps, subtractRange, Tabs } from './shared';

type Allow = DataValidation['type'];
const ALLOW: [Allow, string][] = [['any', 'Any value'], ['whole', 'Whole number'], ['decimal', 'Decimal'], ['list', 'List'], ['date', 'Date'], ['time', 'Time'], ['textLength', 'Text length'], ['custom', 'Custom formula']];
const OPS: [CfOperator, string][] = [['between', 'between'], ['notBetween', 'not between'], ['equal', 'equal to'], ['notEqual', 'not equal to'], ['greaterThan', 'greater than'], ['lessThan', 'less than'], ['greaterThanOrEqual', 'greater than or equal to'], ['lessThanOrEqual', 'less than or equal to']];

/** The stored formula without "=": a list typed as "a, b" becomes a quoted constant. */
const listFormula = (src: string): string => {
  const t = src.trim();
  if (t.startsWith('=')) return t.slice(1);
  return `"${t.split(',').map((x) => x.trim()).filter(Boolean).join(',').replace(/"/g, '')}"`;
};
const listSource = (f: string | undefined): string => (!f ? '' : /^".*"$/s.test(f) ? f.slice(1, -1).split(',').join(', ') : `=${f.replace(/^=/, '')}`);
const showValue = (f: string | undefined): string => (!f ? '' : /^-?\d+(\.\d+)?$/.test(f) ? f : `=${f.replace(/^=/, '')}`);
/** Values typed in the dialog: numbers and plain dates/times stay as typed; "=..." is a formula. */
const valueFormula = (v: string, allow: Allow): string | null => {
  const t = v.trim();
  if (!t) return null;
  if (t.startsWith('=')) return t.slice(1);
  if (!Number.isNaN(Number(t))) return t;
  if (allow === 'date') return `DATEVALUE("${t.replace(/"/g, '')}")`;
  if (allow === 'time') return `TIMEVALUE("${t.replace(/"/g, '')}")`;
  return null;
};

export const ValidationDialog = ({ store, close }: DialogProps) => {
  const { row, col } = store.sel.active;
  const cur = store.sheet.validations.find((v) => v.ranges.some((r) => overlaps(r, { r1: row, c1: col, r2: row, c2: col })));
  const [tab, setTab] = useState<'settings' | 'input' | 'error'>('settings');
  const [allow, setAllow] = useState<Allow>(cur?.type ?? 'any');
  const [op, setOp] = useState<CfOperator>(cur?.operator ?? 'between');
  const [v1, setV1] = useState(cur?.type === 'list' ? listSource(cur.formula1) : cur?.type === 'custom' ? `=${cur.formula1 ?? ''}` : showValue(cur?.formula1));
  const [v2, setV2] = useState(showValue(cur?.formula2));
  const [blank, setBlank] = useState(cur?.allowBlank !== false);
  const [dropdown, setDropdown] = useState(cur?.showDropdown !== false);
  const [showPrompt, setShowPrompt] = useState(!!cur?.prompt);
  const [pTitle, setPTitle] = useState(cur?.prompt?.title ?? '');
  const [pText, setPText] = useState(cur?.prompt?.text ?? '');
  const [showError, setShowError] = useState(cur ? !!cur.error : true);
  const [eStyle, setEStyle] = useState<'stop' | 'warning' | 'information'>(cur?.error?.style ?? 'stop');
  const [eTitle, setETitle] = useState(cur?.error?.title ?? '');
  const [eText, setEText] = useState(cur?.error?.text ?? '');
  const two = op === 'between' || op === 'notBetween';
  const ranged = allow !== 'any' && allow !== 'list' && allow !== 'custom';

  const writeRules = (rule: DataValidation | null, label: string) => {
    const sel = store.sel.ranges;
    store.change(label, 'sheet', () => {
      const kept: DataValidation[] = [];
      for (const dv of store.sheet.validations) {
        let parts = dv.ranges;
        for (const s of sel) parts = parts.flatMap((r) => subtractRange(r, s));
        if (parts.length) kept.push({ ...dv, ranges: parts });
      }
      if (rule) kept.push(rule);
      store.sheet.validations = kept;
    });
  };

  const ok = () => {
    if (allow === 'any' && !showPrompt) { writeRules(null, 'Data validation'); return true; }
    let formula1: string | undefined, formula2: string | undefined;
    if (allow === 'list') {
      if (!v1.trim()) { store.notify('Enter the list items separated by commas, or a range like =$A$1:$A$9.'); setTab('settings'); return false; }
      formula1 = listFormula(v1);
      if (formula1.length > 255 && formula1.startsWith('"')) { store.notify('A typed list can be at most 255 characters. Put the items in cells and use a range.'); return false; }
    } else if (allow === 'custom') {
      formula1 = v1.trim().replace(/^=/, '');
      if (!formula1) { store.notify('Enter a formula that returns TRUE for allowed values, like =A1>0.'); return false; }
    } else if (ranged) {
      const a = valueFormula(v1, allow), b = two ? valueFormula(v2, allow) : undefined;
      if (a === null || (two && b === null)) { store.notify(two ? 'Enter a minimum and a maximum (a number, date, or =formula).' : 'Enter a value (a number, date, or =formula).'); setTab('settings'); return false; }
      formula1 = a; formula2 = b ?? undefined;
    }
    for (const f of [formula1, formula2]) if (f) { try { parseFormula(f); } catch { store.notify('One of the formulas has a problem. Check brackets and quotes.'); return false; } }
    const rule: DataValidation = {
      ranges: store.sel.ranges.map((r) => ({ ...r })),
      type: allow,
      operator: ranged ? op : undefined,
      formula1, formula2,
      allowBlank: blank,
      showDropdown: allow === 'list' ? dropdown : undefined,
      prompt: showPrompt && pText.trim() ? { title: pTitle.trim() || undefined, text: pText.trim() } : undefined,
      error: showError ? { style: eStyle, title: eTitle.trim() || undefined, text: eText.trim() } : undefined,
    };
    writeRules(rule, 'Data validation');
    return true;
  };

  return (
    <Dialog title="Data validation" close={close} onOk={ok} width={480}
      footer={<button onClick={() => { writeRules(null, 'Clear validation'); store.notify('Validation removed from the selection.'); close(); }}>Clear all</button>}>
      <Tabs tabs={[['settings', 'Settings'], ['input', 'Input message'], ['error', 'Error alert']] as const} value={tab} onChange={setTab} />
      {tab === 'settings' && <div className="b-col">
        <div className="grid2">
          <Field label="Allow"><select value={allow} onChange={(e) => setAllow(e.target.value as Allow)}>{ALLOW.map(([v, l]) => <option key={v} value={v}>{l}</option>)}</select></Field>
          {ranged && <Field label="Data"><select value={op} onChange={(e) => setOp(e.target.value as CfOperator)}>{OPS.map(([v, l]) => <option key={v} value={v}>{l}</option>)}</select></Field>}
        </div>
        {ranged && <div className="grid2">
          <Field label={two ? 'Minimum' : 'Value'}><input value={v1} onChange={(e) => setV1(e.target.value)} placeholder={allow === 'date' ? '1/1/2026 or =TODAY()' : allow === 'time' ? '9:00 or =B1' : '0 or =B1'} /></Field>
          {two && <Field label="Maximum"><input value={v2} onChange={(e) => setV2(e.target.value)} placeholder={allow === 'date' ? '12/31/2026' : '100'} /></Field>}
        </div>}
        {allow === 'list' && <>
          <Field label="Source"><input value={v1} onChange={(e) => setV1(e.target.value)} placeholder="Yes, No, Maybe   or   =$A$1:$A$9" /></Field>
          <Check label="Show a dropdown in the cell" checked={dropdown} onChange={setDropdown} />
        </>}
        {allow === 'custom' && <Field label="Formula"><input className="mono" value={v1} onChange={(e) => setV1(e.target.value)} placeholder="=ISNUMBER(A1)" /></Field>}
        {allow !== 'any' && <Check label="Ignore blank cells" checked={blank} onChange={setBlank} />}
        <p className="muted small">Applies to the selected cells. Rules already on these cells are replaced.</p>
      </div>}
      {tab === 'input' && <div className="b-col">
        <Check label="Show a message when the cell is selected" checked={showPrompt} onChange={setShowPrompt} />
        <Field label="Title"><input value={pTitle} maxLength={32} disabled={!showPrompt} onChange={(e) => setPTitle(e.target.value)} /></Field>
        <Field label="Message"><textarea value={pText} maxLength={255} disabled={!showPrompt} onChange={(e) => setPText(e.target.value)} /></Field>
      </div>}
      {tab === 'error' && <div className="b-col">
        <Check label="Show an alert when the entry isn't allowed" checked={showError} onChange={setShowError} />
        <Field label="Kind"><select value={eStyle} disabled={!showError} onChange={(e) => setEStyle(e.target.value as typeof eStyle)}>
          <option value="stop">Stop: block the entry</option><option value="warning">Warning: ask before keeping it</option><option value="information">Information: just tell</option>
        </select></Field>
        <Field label="Title"><input value={eTitle} maxLength={32} disabled={!showError} onChange={(e) => setETitle(e.target.value)} /></Field>
        <Field label="Message"><textarea value={eText} maxLength={255} disabled={!showError} onChange={(e) => setEText(e.target.value)} placeholder="This value isn't allowed here." /></Field>
      </div>}
    </Dialog>
  );
};
