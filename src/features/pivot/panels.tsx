// Tool-panel groups for PivotTables, slicers and outlines.
import type { ReactNode } from 'react';
import {
  BarChart3, Calculator, Columns3, Database, Filter, FoldVertical, Group, LayoutList, ListTree, MoveRight, PanelRight, RefreshCw, RefreshCcw, Rows3,
  Settings2, SlidersHorizontal, Sigma, SquareMousePointer, Table2, Timer, Trash2, Ungroup, UnfoldVertical, type LucideIcon,
} from 'lucide-react';
import { command, run } from '../../app/commands';
import type { GridStore } from '../../app/store';
import { prettyKeys } from '../../ui/keys';
import { VALUES_FIELD } from '../../model/pivot';
import { PRESETS, NO_STYLE, swatch } from './styles';
import { currentPivot, pui } from './ui';
import { setAllFields, setGrandTotals, setStyle, setSubtotalPlacement, setTableLayout, moveValues, update } from './ops';
import { N, T, Tr } from '../../ui/i18n';

const S = 1.9;

export const Tool = ({ store, id, icon: Icon, label, arg }: { store: GridStore; id: string; icon?: LucideIcon; label?: string; arg?: unknown }) => {
  const c = command(id);
  if (!c) return null;
  const on = c.active?.(store) ?? false;
  const text = label ?? c.label;
  return (
    <button className={`tool${on ? ' on' : ''}`} disabled={c.enabled ? !c.enabled(store) : false} aria-pressed={c.active ? on : undefined} title={c.keys ? `${text} (${prettyKeys(c.keys)})` : text}
      onMouseDown={(e) => e.preventDefault()} onClick={() => run(store, id, arg)}>
      {Icon && <Icon size={16} strokeWidth={S} />}<span>{Tr(text)}</span>
    </button>
  );
};
const List = ({ children }: { children: ReactNode }) => <div className="toollist">{children}</div>;

export const pivotSection = (s: GridStore) => {
  const ref = currentPivot(s);
  return (
    <List>
      <Tool store={s} id="data.pivot" icon={Table2} label={T('New PivotTable…')} />
      {ref && <Tool store={s} id="pivot.fieldList" icon={PanelRight} />}
      {ref && <Tool store={s} id="pivot.refresh" icon={RefreshCw} />}
      <Tool store={s} id="pivot.refreshAll" icon={RefreshCcw} />
      {ref && <Tool store={s} id="pivot.changeSource" icon={Database} />}
      {ref && <Tool store={s} id="pivot.options" icon={Settings2} />}
      {ref && <Tool store={s} id="pivot.fieldSettings" icon={SlidersHorizontal} />}
      {ref && <Tool store={s} id="pivot.select" icon={SquareMousePointer} />}
      {ref && <Tool store={s} id="pivot.move" icon={MoveRight} />}
      {ref && <Tool store={s} id="pivot.delete" icon={Trash2} />}
      <Tool store={s} id="pivot.gpd" />
    </List>
  );
};

export const pivotLayoutSection = (s: GridStore) => {
  const ref = currentPivot(s);
  if (!ref) return <p className="muted small">{T('Select a cell in a PivotTable to change its layout.')}</p>;
  const pt = ref.pt;
  const axisFields = [...pt.rows, ...pt.cols].filter((f) => f >= 0);
  const subWhere = axisFields.every((f) => pt.fields[f]?.subtotals?.length === 0) ? 'none' : axisFields.some((f) => pt.fields[f]?.subtotalTop === false) ? 'bottom' : 'top';
  const grand = pt.rowGrand && pt.colGrand ? 'both' : pt.rowGrand ? 'rows' : pt.colGrand ? 'cols' : 'none';
  const repeat = axisFields.length > 0 && axisFields.every((f) => pt.fields[f]?.repeatLabels);
  const blank = axisFields.length > 0 && axisFields.every((f) => pt.fields[f]?.blankRow);
  return (
    <div className="pvtools">
      <label className="field"><span>{T('Report layout')}</span>
        <select value={pt.layout} onChange={(e) => setTableLayout(s, ref, e.target.value as 'compact')}><option value="compact">{T('Compact')}</option><option value="outline">{T('Outline')}</option><option value="tabular">{T('Tabular')}</option></select>
      </label>
      <label className="field"><span>{T('Subtotals')}</span>
        <select value={subWhere} onChange={(e) => setSubtotalPlacement(s, ref, e.target.value as 'top')}><option value="top">{T('At the top of each group')}</option><option value="bottom">{T('At the bottom of each group')}</option><option value="none">{T('Don\'t show')}</option></select>
      </label>
      <label className="field"><span>{T('Grand totals')}</span>
        <select value={grand} onChange={(e) => setGrandTotals(s, ref, e.target.value as 'both')}><option value="both">{T('For rows and columns')}</option><option value="cols">{T('Total row only')}</option><option value="rows">{T('Total column only')}</option><option value="none">{T('Off')}</option></select>
      </label>
      {pt.data.length > 1 && (
        <label className="field"><span>{T('Values go in')}</span>
          <select value={pt.rows.includes(VALUES_FIELD) ? 'rows' : 'cols'} onChange={(e) => moveValues(s, ref, e.target.value as 'rows')}><option value="cols">{T('Columns')}</option><option value="rows">{T('Rows')}</option></select>
        </label>
      )}
      <label className="b-check"><input type="checkbox" checked={repeat} onChange={(e) => setAllFields(s, ref, (f) => ({ ...f, repeatLabels: e.target.checked || undefined }), N('Repeat item labels'))} />{T('Repeat item labels')}</label>
      <label className="b-check"><input type="checkbox" checked={blank} onChange={(e) => setAllFields(s, ref, (f) => ({ ...f, blankRow: e.target.checked || undefined }), N('Blank rows'))} />{T('Blank row after each item')}</label>
      <label className="b-check"><input type="checkbox" checked={pt.showDrill} onChange={(e) => update(s, ref, (p) => ({ ...p, showDrill: e.target.checked }), N('+/- buttons'))} />{T('+/- buttons')}</label>
      <label className="b-check"><input type="checkbox" checked={pt.showHeaders} onChange={(e) => update(s, ref, (p) => ({ ...p, showHeaders: e.target.checked }), N('Field headers'))} />{T('Field headers')}</label>
      <List>
        <Tool store={s} id="pivot.expandAll" icon={UnfoldVertical} />
        <Tool store={s} id="pivot.collapseAll" icon={FoldVertical} />
      </List>
    </div>
  );
};

export const pivotCalcSection = (s: GridStore) => (
  <List>
    <Tool store={s} id="pivot.group" icon={Group} />
    <Tool store={s} id="pivot.ungroup" icon={Ungroup} />
    <Tool store={s} id="pivot.calcField" icon={Calculator} />
    <Tool store={s} id="pivot.calcItem" icon={Sigma} />
    <Tool store={s} id="pivot.showPages" icon={LayoutList} />
    <Tool store={s} id="pivot.drillDown" icon={ListTree} />
    <Tool store={s} id="pivot.clearFilters" icon={Filter} />
  </List>
);

export const pivotStyleSection = (s: GridStore) => {
  const ref = currentPivot(s);
  if (!ref) return <p className="muted small">{T('Select a cell in a PivotTable to style it.')}</p>;
  const st = ref.pt.style;
  const all = [null, ...PRESETS];
  return (
    <div className="pvtools">
      <div className="pvstyles" role="radiogroup" aria-label={T('PivotTable styles')}>
        {all.map((p) => {
          const name = p?.name ?? NO_STYLE;
          const sw = swatch(p);
          const on = (st.name || NO_STYLE) === name || (!PRESETS.some((x) => x.name === st.name) && name === 'PivotStyleLight16' && st.name !== NO_STYLE && !!st.name);
          return (
            <button key={name} role="radio" aria-checked={on} className={`pvstyle${on ? ' on' : ''}`} title={p?.label ?? T('No style')} aria-label={p?.label ?? T('No style')} onClick={() => setStyle(s, ref, { name })}>
              <span style={{ background: sw.head, color: sw.headInk }} />
              <span style={{ background: sw.band }} /><span /><span style={{ background: sw.band }} />
              <span style={{ background: sw.total }} />
            </button>
          );
        })}
      </div>
      <label className="b-check"><input type="checkbox" checked={st.rowHeaders} onChange={(e) => setStyle(s, ref, { rowHeaders: e.target.checked })} />{T('Row headers')}</label>
      <label className="b-check"><input type="checkbox" checked={st.colHeaders} onChange={(e) => setStyle(s, ref, { colHeaders: e.target.checked })} />{T('Column headers')}</label>
      <label className="b-check"><input type="checkbox" checked={st.bandedRows} onChange={(e) => setStyle(s, ref, { bandedRows: e.target.checked })} />{T('Banded rows')}</label>
      <label className="b-check"><input type="checkbox" checked={st.bandedCols} onChange={(e) => setStyle(s, ref, { bandedCols: e.target.checked })} />{T('Banded columns')}</label>
    </div>
  );
};

export const slicerSection = (s: GridStore) => (
  <List>
    <Tool store={s} id="insert.slicer" icon={SlidersHorizontal} />
    <Tool store={s} id="insert.timeline" icon={Timer} />
    {pui.slicer && <Tool store={s} id="slicer.settings" icon={Settings2} />}
    {pui.slicer && <Tool store={s} id="slicer.delete" icon={Trash2} />}
  </List>
);

export const insertPivotSection = (s: GridStore) => (
  <List>
    <Tool store={s} id="data.pivot" icon={Table2} label={T('PivotTable…')} />
    <Tool store={s} id="insert.pivotChart" icon={BarChart3} />
    <Tool store={s} id="insert.slicer" icon={SlidersHorizontal} />
    <Tool store={s} id="insert.timeline" icon={Timer} />
  </List>
);

export const outlineSection = (s: GridStore) => (
  <List>
    <Tool store={s} id="cells.group" icon={Group} />
    <Tool store={s} id="cells.ungroup" icon={Ungroup} />
    <Tool store={s} id="outline.auto" icon={Rows3} />
    <Tool store={s} id="outline.clear" icon={Trash2} />
    <Tool store={s} id="outline.showDetail" icon={UnfoldVertical} />
    <Tool store={s} id="outline.hideDetail" icon={FoldVertical} />
    <Tool store={s} id="outline.settings" icon={Settings2} />
    <Tool store={s} id="data.subtotal" icon={Sigma} />
    <Tool store={s} id="data.consolidate" icon={Columns3} />
    <Tool store={s} id="edit.selectVisible" icon={SquareMousePointer} />
  </List>
);
