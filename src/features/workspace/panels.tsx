// Tool panel groups for the View, Review, Formulas and File rail tabs.
import type { ReactNode } from 'react';
import {
  Calculator, Columns2, Crosshair, Eye, EyeOff, FileLock2, History, LayoutGrid, Lock, Maximize2, Navigation, Plus, Settings2, Shield, SplitSquareVertical,
  SquareStack, ZoomIn, type LucideIcon,
} from 'lucide-react';
import { command, run } from '../../app/commands';
import type { GridStore } from '../../app/store';
import { prettyKeys } from '../../ui/keys';
import { activeViewOf, switchView } from './views';
import { settings, saveSettings } from './settings';
import { desktopWindow } from './desktop';
import { T, Tr } from '../../ui/i18n';

const S = 1.9;

const Btn = ({ store, id, icon: Icon, label }: { store: GridStore; id: string; icon?: LucideIcon; label?: string }) => {
  const c = command(id);
  if (!c) return null;
  const on = c.active?.(store) ?? false;
  const enabled = c.enabled?.(store) ?? true;
  const text = label ?? c.label;
  return (
    <button className={`tool${on ? ' on' : ''}`} disabled={!enabled} aria-pressed={c.active ? on : undefined} title={c.keys ? `${text} (${prettyKeys(c.keys)})` : text}
      onMouseDown={(e) => e.preventDefault()} onClick={() => run(store, id)}>
      {Icon && <Icon size={16} strokeWidth={S} />}<span>{Tr(text)}</span>
    </button>
  );
};
const List = ({ children }: { children: ReactNode }) => <div className="toollist">{children}</div>;

const FOCUS_COLORS = ['#f5c542', '#7bd389', '#6cb4ff', '#ff8fa3', '#c39bff', '#ffa94d'];

export const focusSection = (s: GridStore) => (
  <>
    <List><Btn store={s} id="view.focusCell" icon={Crosshair} /></List>
    <div className="swatches ws-swatches" aria-label={T('Focus cell color')}>
      {FOCUS_COLORS.map((c) => (
        <button key={c} className={`sw${settings.focusColor === c ? ' on' : ''}`} style={{ background: c }} title={c} aria-label={T('Focus color {c}', { c })} onClick={() => saveSettings(s, { focusColor: c })} />
      ))}
      <label className="sw custom" title={T('More colors')} aria-label={T('More colors')}><input type="color" value={settings.focusColor} onChange={(e) => saveSettings(s, { focusColor: e.target.value })} /></label>
    </div>
  </>
);

export const sheetViewSection = (s: GridStore) => {
  const views = s.sheet.sheetViews ?? [];
  const cur = activeViewOf(s.sheet);
  return (
    <>
      <select aria-label={T('Sheet view')} className="wide" value={cur === null ? '__default' : cur === '' ? '__temp' : cur} onChange={(e) => {
        const v = e.target.value;
        switchView(s, v === '__default' ? null : v === '__temp' ? '' : v);
      }}>
        <option value="__default">{T('Default')}</option>
        {cur === '' && <option value="__temp">{T('Temporary view')}</option>}
        {views.map((v) => <option key={v.id} value={v.id}>{v.name}</option>)}
      </select>
      <List>
        <Btn store={s} id="view.newSheetView" icon={Plus} />
        <Btn store={s} id="view.tempView" />
        <Btn store={s} id="view.sheetViews" icon={SquareStack} />
        <Btn store={s} id="view.customViews" icon={LayoutGrid} />
      </List>
    </>
  );
};

export const windowSection = (s: GridStore) => (
  <List>
    <Btn store={s} id="view.navPane" icon={Navigation} />
    <Btn store={s} id="view.fullScreen" icon={Maximize2} />
    <Btn store={s} id="view.split" icon={SplitSquareVertical} />
    <Btn store={s} id="view.newWindow" icon={Columns2} />
    {desktopWindow?.arrangeWindows && <Btn store={s} id="view.arrange" icon={LayoutGrid} />}
    <Btn store={s} id="view.zoomFit" icon={ZoomIn} />
  </List>
);

export const showSection = (s: GridStore) => (
  <List>
    <Btn store={s} id="view.formulaBar" icon={settings.showFormulaBar ? Eye : EyeOff} />
    <Btn store={s} id="view.statusBar" />
  </List>
);

export const protectSection = (s: GridStore) => (
  <List>
    <Btn store={s} id="sheet.protect" icon={Lock} />
    <Btn store={s} id="review.protectWorkbook" icon={Shield} />
    <Btn store={s} id="file.encrypt" icon={FileLock2} />
  </List>
);

export const calcSection = (s: GridStore) => (
  <List>
    <Btn store={s} id="formula.calcManual" icon={Calculator} />
    <Btn store={s} id="formula.calcSheet" />
    <Btn store={s} id="formula.calcOptions" icon={Settings2} />
  </List>
);

export const fileSection = (s: GridStore) => (
  <List>
    <Btn store={s} id="file.versions" icon={History} />
    <Btn store={s} id="file.encrypt" icon={FileLock2} />
    <Btn store={s} id="file.options" icon={Settings2} />
  </List>
);
