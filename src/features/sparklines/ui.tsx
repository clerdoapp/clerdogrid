// Sparkline dialog and tool-panel section.
import { useEffect, useState, type ReactNode } from 'react';
import { BarChart, Eraser, LineChart, Pencil, Group as GroupIcon, Ungroup, Trash2, Activity } from 'lucide-react';
import { openDialog, run } from '../../app/commands';
import type { GridStore } from '../../app/store';
import type { SparklineGroup, SparklineKind } from '../../model/workbook';
import { rangeAddr } from '../../model/workbook';
import { Dialog, Field, type DialogProps } from '../../ui/dialogs/Dialog';
import { rangeText, resolveRange } from '../../ui/dialogs/a/util';
import { buildSparklines, defaultGroup, groupAt, SPARK_STYLES } from './model';
import { refString } from '../../ui/charts/chartData';
import { N, T, Tr } from '../../ui/i18n';

const S = 1.9;
const KINDS: [SparklineKind, string][] = [['line', N('Line')], ['column', N('Column')], ['winLoss', N('Win/loss')]];

/** Edits the group that holds the active cell, as one undo step. */
export const editGroup = (store: GridStore, label: string, fn: (g: SparklineGroup) => void): void => {
  const { row, col } = store.sel.active;
  if (!groupAt(store.sheet, row, col)) return;
  store.change(label, 'sheet', () => {
    const hit = groupAt(store.sheet, row, col);
    if (hit) fn(hit.group);
  });
};

export const SparklineDialog = ({ store, data, close }: DialogProps<{ kind?: SparklineKind; edit?: boolean }>) => {
  const hit = data?.edit ? groupAt(store.sheet, store.sel.active.row, store.sel.active.col) : null;
  const r = store.range;
  const [kind, setKind] = useState<SparklineKind>(hit?.group.kind ?? data?.kind ?? 'line');
  const [dataText, setDataText] = useState(() => {
    if (hit) {
      const refs = hit.group.sparklines.map((s) => resolveRange(store, s.ref)).filter((x): x is NonNullable<typeof x> => !!x);
      if (!refs.length) return hit.group.sparklines[0]?.ref ?? '';
      const box = refs.reduce((b, x) => ({ r1: Math.min(b.r1, x.range.r1), c1: Math.min(b.c1, x.range.c1), r2: Math.max(b.r2, x.range.r2), c2: Math.max(b.c2, x.range.c2) }), refs[0]!.range);
      return rangeText(store, box, refs[0]!.sheet);
    }
    return r.r1 === r.r2 && r.c1 === r.c2 ? '' : rangeAddr(r);
  });
  const [locText, setLocText] = useState(() => {
    if (hit) {
      const sps = hit.group.sparklines;
      return rangeAddr({ r1: Math.min(...sps.map((s) => s.row)), c1: Math.min(...sps.map((s) => s.col)), r2: Math.max(...sps.map((s) => s.row)), c2: Math.max(...sps.map((s) => s.col)) });
    }
    // Next to a block of rows: the column to its right.
    return r.r1 === r.r2 && r.c1 === r.c2 ? '' : rangeAddr({ r1: r.r1, c1: r.c2 + 1, r2: r.r2, c2: r.c2 + 1 });
  });
  const ok = (): boolean => {
    const d = resolveRange(store, dataText);
    if (!d) { store.notify(T('Enter the data range, such as B2:G9.')); return false; }
    const loc = resolveRange(store, locText);
    if (!loc || loc.sheet !== store.sheetIndex) { store.notify(T('Enter where to put the sparklines on this sheet, such as H2:H9.')); return false; }
    const sps = buildSparklines(store.sheet.name, { sheet: store.wb.sheets[d.sheet]!.name, range: d.range }, loc.range);
    if (typeof sps === 'string') { store.notify(sps); return false; }
    store.change(hit ? N('Edit sparklines') : N('Insert sparklines'), 'sheet', () => {
      const sh = store.sheet;
      const taken = new Set(sps.map((s) => `${s.row},${s.col}`));
      const old = hit ? sh.sparklines?.find((g) => g.sparklines.some((s) => s.row === hit.group.sparklines[0]?.row && s.col === hit.group.sparklines[0]?.col)) : undefined;
      // New sparklines replace any already in those cells.
      sh.sparklines = (sh.sparklines ?? []).filter((g) => g !== old).map((g) => ({ ...g, sparklines: g.sparklines.filter((s) => !taken.has(`${s.row},${s.col}`)) })).filter((g) => g.sparklines.length);
      const base = old ? { ...old, kind, sparklines: sps } : defaultGroup(kind, sps);
      if (kind === 'winLoss' && base.negative === undefined) base.negative = true;
      sh.sparklines.push(base);
    });
    return true;
  };
  return (
    <Dialog title={hit ? T('Edit sparklines') : T('Insert sparklines')} close={close} onOk={ok} okLabel={hit ? T('Save') : T('Insert')}>
      <Field label={T('Kind')}>
        <div className="segmented" role="radiogroup" aria-label={T('Kind')}>
          {KINDS.map(([k, l]) => <button key={k} role="radio" aria-checked={kind === k} className={kind === k ? 'on' : ''} onClick={() => setKind(k)}>{Tr(l)}</button>)}
        </div>
      </Field>
      <Field label={T('Data range')}><input value={dataText} onChange={(e) => setDataText(e.target.value)} placeholder="B2:G9" spellCheck={false} autoFocus /></Field>
      <Field label={T('Location (one cell per row or column of data)')}><input value={locText} onChange={(e) => setLocText(e.target.value)} placeholder="H2:H9" spellCheck={false} /></Field>
    </Dialog>
  );
};

const Chk = ({ label, on, set }: { label: string; on: boolean; set: (v: boolean) => void }) => (
  <label className="cfcheck"><input type="checkbox" checked={on} onChange={(e) => set(e.target.checked)} /> {Tr(label)}</label>
);
const Row = ({ label, children }: { label: string; children: ReactNode }) => <label className="cfrow"><span>{Tr(label)}</span>{children}</label>;
const ColorIn = ({ value, set, label }: { value: string | undefined; set: (v: string) => void; label: string }) => (
  <input type="color" className="sparkcolor" aria-label={Tr(label)} value={value && /^#[0-9a-f]{6}$/i.test(value) ? value : '#12a565'} onChange={(e) => set(e.target.value.toUpperCase())} />
);
const NumIn = ({ value, set, label }: { value: number | undefined; set: (v: number | undefined) => void; label: string }) => {
  const [t, setT] = useState(value?.toString() ?? '');
  useEffect(() => setT(value?.toString() ?? ''), [value]);
  const commit = () => { const n = t.trim() ? Number(t) : undefined; if (n === undefined || Number.isFinite(n)) { if (n !== value) set(n); } else setT(value?.toString() ?? ''); };
  return <input aria-label={Tr(label)} inputMode="decimal" value={t} placeholder="0" onChange={(e) => setT(e.target.value)} onBlur={commit} onKeyDown={(e) => { e.stopPropagation(); if (e.key === 'Enter') commit(); }} />;
};

const SparkFormat = ({ store, g }: { store: GridStore; g: SparklineGroup }) => {
  const up = (p: Partial<SparklineGroup>, label = N('Format sparklines')) => editGroup(store, label, (x) => {
    Object.assign(x, p);
    for (const k of Object.keys(p) as (keyof SparklineGroup)[]) if (p[k] === undefined) delete x[k];
  });
  const [date, setDate] = useState(g.dateRange ?? '');
  useEffect(() => setDate(g.dateRange ?? ''), [g.dateRange]);
  return (
    <div className="chartformat sparkfmt">
      <div className="mini">{T('Selected sparkline group ({count})', { count: g.sparklines.length })}</div>
      <div className="segmented" role="radiogroup" aria-label={T('Sparkline kind')}>
        {KINDS.map(([k, l]) => <button key={k} role="radio" aria-checked={g.kind === k} className={g.kind === k ? 'on' : ''} onClick={() => up({ kind: k }, N('Sparkline kind'))}>{Tr(l)}</button>)}
      </div>
      <div className="mini">{T('Style')}</div>
      <div className="sparkstyles" role="group" aria-label={T('Sparkline styles')}>
        {SPARK_STYLES.map((st) => (
          <button key={st.name} className="sparkstyle" title={Tr(st.name)} aria-label={Tr(st.name)} onClick={() => up({ color: st.color, negativeColor: st.negativeColor, markersColor: st.markersColor, highColor: st.highColor, lowColor: st.lowColor, firstColor: st.firstColor, lastColor: st.lastColor }, N('Sparkline style'))}>
            <i style={{ background: st.color }} /><i style={{ background: st.highColor }} /><i style={{ background: st.negativeColor }} />
          </button>
        ))}
      </div>
      <div className="mini">{T('Show')}</div>
      <div className="sparkshow">
        {([['high', 'highColor', N('High point')], ['low', 'lowColor', N('Low point')], ['first', 'firstColor', N('First point')], ['last', 'lastColor', N('Last point')], ['negative', 'negativeColor', N('Negative points')], ...(g.kind === 'line' ? [['markers', 'markersColor', N('Markers')]] : [])] as [keyof SparklineGroup, keyof SparklineGroup, string][])
          .map(([k, ck, l]) => (
            <div key={k} className="sparkopt">
              <Chk label={Tr(l)} on={!!g[k]} set={(v) => up({ [k]: v || undefined })} />
              <ColorIn label={T('{l} colour', { l: Tr(l) })} value={g[ck] as string | undefined} set={(c) => up({ [ck]: c })} />
            </div>
          ))}
      </div>
      <Row label={T('Colour')}><ColorIn label={T('Sparkline colour')} value={g.color} set={(c) => up({ color: c })} /></Row>
      {g.kind === 'line' && <Row label={T('Weight (pt)')}><NumIn label={T('Line weight')} value={g.lineWeight} set={(v) => up({ lineWeight: v === undefined ? undefined : Math.max(0.25, Math.min(10, v)) })} /></Row>}
      <div className="mini">{T('Axis')}</div>
      <Row label={T('Minimum')}>
        <select aria-label={T('Axis minimum')} value={g.minAxis ?? 'individual'} onChange={(e) => up({ minAxis: e.target.value === 'individual' ? undefined : (e.target.value as 'group' | 'custom') })}>
          <option value="individual">{T('Each sparkline')}</option><option value="group">{T('Same for all')}</option><option value="custom">{T('Fixed')}</option>
        </select>
      </Row>
      {g.minAxis === 'custom' && <Row label={T('Fixed minimum')}><NumIn label={T('Fixed minimum')} value={g.manualMin} set={(v) => up({ manualMin: v })} /></Row>}
      <Row label={T('Maximum')}>
        <select aria-label={T('Axis maximum')} value={g.maxAxis ?? 'individual'} onChange={(e) => up({ maxAxis: e.target.value === 'individual' ? undefined : (e.target.value as 'group' | 'custom') })}>
          <option value="individual">{T('Each sparkline')}</option><option value="group">{T('Same for all')}</option><option value="custom">{T('Fixed')}</option>
        </select>
      </Row>
      {g.maxAxis === 'custom' && <Row label={T('Fixed maximum')}><NumIn label={T('Fixed maximum')} value={g.manualMax} set={(v) => up({ manualMax: v })} /></Row>}
      <div className="sparkopt"><Chk label={T('Show axis')} on={!!g.showAxis} set={(v) => up({ showAxis: v || undefined })} /><ColorIn label={T('Axis colour')} value={g.axisColor} set={(c) => up({ axisColor: c })} /></div>
      <Chk label={T('Right to left')} on={!!g.rightToLeft} set={(v) => up({ rightToLeft: v || undefined })} />
      <Row label={T('Dates')}>
        <input aria-label={T('Date range')} value={date} placeholder={T('Even spacing')} spellCheck={false} onChange={(e) => setDate(e.target.value)}
          onBlur={() => {
            const v = date.trim();
            if (v === (g.dateRange ?? '')) return;
            const r = v ? resolveRange(store, v) : null;
            if (v && !r) { store.notify(T('Enter the dates as a range, such as B1:G1.')); return; }
            up({ dateRange: r ? refString(store.wb.sheets[r.sheet]!.name, r.range) : undefined }, N('Sparkline dates'));
          }}
          onKeyDown={(e) => { e.stopPropagation(); if (e.key === 'Enter') (e.target as HTMLInputElement).blur(); }} />
      </Row>
      <Row label={T('Empty cells')}>
        <select aria-label={T('Show empty cells as')} value={g.displayEmptyCellsAs ?? 'gap'} onChange={(e) => up({ displayEmptyCellsAs: e.target.value as 'gap' | 'zero' | 'span' })}>
          <option value="gap">{T('Gaps')}</option><option value="zero">{T('Zero')}</option><option value="span">{T('Join with a line')}</option>
        </select>
      </Row>
      <Chk label={T('Show data in hidden rows and columns')} on={!!g.displayHidden} set={(v) => up({ displayHidden: v || undefined })} />
      <div className="toollist">
        <button className="tool" onClick={() => openDialog(store, 'sparklines', { edit: true })}><Pencil size={16} strokeWidth={S} /><span>{T('Edit data…')}</span></button>
        <button className="tool" onClick={() => run(store, 'sparkline.group')}><GroupIcon size={16} strokeWidth={S} /><span>{T('Group selected')}</span></button>
        <button className="tool" onClick={() => run(store, 'sparkline.ungroup')}><Ungroup size={16} strokeWidth={S} /><span>{T('Ungroup')}</span></button>
        <button className="tool" onClick={() => run(store, 'sparkline.clear')}><Eraser size={16} strokeWidth={S} /><span>{T('Clear selected sparklines')}</span></button>
        <button className="tool" onClick={() => run(store, 'sparkline.clearGroup')}><Trash2 size={16} strokeWidth={S} /><span>{T('Clear group')}</span></button>
      </div>
    </div>
  );
};

export const SparklineSection = ({ store }: { store: GridStore }) => {
  const hit = groupAt(store.sheet, store.sel.active.row, store.sel.active.col);
  return (
    <>
      <div className="toollist">
        <button className="tool" onMouseDown={(e) => e.preventDefault()} onClick={() => run(store, 'insert.sparkline', 'line')}><LineChart size={16} strokeWidth={S} /><span>{T('Line')}</span></button>
        <button className="tool" onMouseDown={(e) => e.preventDefault()} onClick={() => run(store, 'insert.sparkline', 'column')}><BarChart size={16} strokeWidth={S} /><span>{T('Column')}</span></button>
        <button className="tool" onMouseDown={(e) => e.preventDefault()} onClick={() => run(store, 'insert.sparkline', 'winLoss')}><Activity size={16} strokeWidth={S} /><span>{T('Win/loss')}</span></button>
      </div>
      {hit && <SparkFormat key={`${hit.group.sparklines[0]?.row},${hit.group.sparklines[0]?.col}`} store={store} g={hit.group} />}
    </>
  );
};
