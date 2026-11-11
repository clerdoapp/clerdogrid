import type { ReactNode } from 'react';
import {
  Columns2, Eye, FileDown, FileText, Grid3x3, LayoutTemplate, PanelTop, Printer, RectangleHorizontal, RectangleVertical, Rows2, ScissorsLineDashed,
  Settings2, SquareDashed, type LucideIcon,
} from 'lucide-react';
import { command, run } from '../../app/commands';
import type { GridStore } from '../../app/store';
import { prettyKeys } from '../../ui/keys';
import { COMMON_PAPERS, MARGIN_PRESETS, paperByCode } from './paper';
import { DEFAULT_MARGINS } from './paginate';
import { setupOf, updateSetup } from './setup';
import { N, T, Tr } from '../../ui/i18n';

const S = 1.9;

const Tool = ({ store, id, icon: Icon, label }: { store: GridStore; id: string; icon?: LucideIcon; label?: string }) => {
  const c = command(id);
  if (!c) return null;
  const on = c.active?.(store) ?? false;
  const text = label ?? c.label;
  return (
    <button className={`tool${on ? ' on' : ''}`} disabled={c.enabled ? !c.enabled(store) : false} aria-pressed={c.active ? on : undefined} title={c.keys ? `${text} (${prettyKeys(c.keys)})` : text}
      onMouseDown={(e) => e.preventDefault()} onClick={() => run(store, id)}>
      {Icon && <Icon size={16} strokeWidth={S} />}<span>{Tr(text)}</span>
    </button>
  );
};
const List = ({ children }: { children: ReactNode }) => <div className="toollist">{children}</div>;

const Pick = ({ label, value, onChange, children }: { label: string; value: string; onChange: (v: string) => void; children: ReactNode }) => (
  <label className="pl-pick"><span>{Tr(label)}</span><select value={value} onChange={(e) => onChange(e.target.value)}>{children}</select></label>
);

export const setupSection = (s: GridStore) => {
  const ps = setupOf(s);
  const m = { ...DEFAULT_MARGINS, ...ps.margins };
  const preset = MARGIN_PRESETS.find((p) => (Object.keys(p.m) as (keyof typeof m)[]).every((k) => Math.abs(p.m[k] - m[k]) < 0.001))?.id ?? 'custom';
  const fit = ps.fitWidth !== undefined || ps.fitHeight !== undefined;
  const fitOpt = (n: number | undefined) => (n === undefined ? '1' : String(n));
  return (
    <>
      <div className="iconrow">
        <Tool store={s} id="page.portrait" icon={RectangleVertical} /><Tool store={s} id="page.landscape" icon={RectangleHorizontal} />
      </div>
      <Pick label={T('Size')} value={String(ps.paper ?? 1)} onChange={(v) => updateSetup(s, N('Paper size'), { paper: +v, paperWidth: undefined, paperHeight: undefined })}>
        {[...new Set([...COMMON_PAPERS, ps.paper ?? 1])].map((c) => <option key={c} value={c}>{Tr(paperByCode(c).name)}</option>)}
      </Pick>
      <Pick label={T('Margins')} value={preset} onChange={(v) => { const p = MARGIN_PRESETS.find((x) => x.id === v); if (p) updateSetup(s, N('Margins'), { margins: { ...p.m } }); else run(s, 'page.setup', { tab: 'margins' }); }}>
        {MARGIN_PRESETS.map((p) => <option key={p.id} value={p.id}>{Tr(p.label)}</option>)}
        <option value="custom">{preset === 'custom' ? T('Custom') : T('Custom…')}</option>
      </Pick>
      <Pick label={T('Width')} value={fit ? (ps.fitWidth === 0 ? 'auto' : fitOpt(ps.fitWidth)) : 'auto'} onChange={(v) => updateSetup(s, N('Scale to fit'), v === 'auto' && (!fit || !ps.fitHeight) ? { fitWidth: undefined, fitHeight: undefined } : { fitWidth: v === 'auto' ? 0 : +v, fitHeight: fit ? ps.fitHeight ?? 1 : 0, scale: undefined })}>
        <option value="auto">{T('Automatic')}</option>{[1, 2, 3, 4, 5, 6, 7, 8, 9].map((n) => <option key={n} value={n}>{T('{count, plural, one {# page} other {# pages}}', { count: n })}</option>)}
      </Pick>
      <Pick label={T('Height')} value={fit ? (ps.fitHeight === 0 ? 'auto' : fitOpt(ps.fitHeight)) : 'auto'} onChange={(v) => updateSetup(s, N('Scale to fit'), v === 'auto' && (!fit || !ps.fitWidth) ? { fitWidth: undefined, fitHeight: undefined } : { fitHeight: v === 'auto' ? 0 : +v, fitWidth: fit ? ps.fitWidth ?? 1 : 0, scale: undefined })}>
        <option value="auto">{T('Automatic')}</option>{[1, 2, 3, 4, 5, 6, 7, 8, 9].map((n) => <option key={n} value={n}>{T('{count, plural, one {# page} other {# pages}}', { count: n })}</option>)}
      </Pick>
      <label className="pl-pick"><span>{T('Scale')}</span>
        <input type="number" min={10} max={400} disabled={fit} value={fit ? '' : ps.scale ?? 100} placeholder={fit ? T('Fit') : undefined}
          onChange={(e) => { const n = Math.round(+e.target.value); if (n >= 10 && n <= 400) updateSetup(s, N('Scale'), { scale: n }); }} />
      </label>
      <List><Tool store={s} id="page.setup" icon={Settings2} /></List>
    </>
  );
};

export const areaSection = (s: GridStore) => (
  <List>
    <Tool store={s} id="page.printAreaSet" icon={SquareDashed} /><Tool store={s} id="page.printAreaAdd" /><Tool store={s} id="page.printAreaClear" />
    <Tool store={s} id="page.titles" icon={Rows2} />
  </List>
);

export const breaksSection = (s: GridStore) => (
  <List><Tool store={s} id="page.breakInsert" icon={ScissorsLineDashed} /><Tool store={s} id="page.breakRemove" /><Tool store={s} id="page.breakReset" /></List>
);

export const viewsSection = (s: GridStore) => (
  <List><Tool store={s} id="page.viewNormal" icon={Grid3x3} /><Tool store={s} id="page.viewBreaks" icon={Columns2} /><Tool store={s} id="page.viewLayout" icon={LayoutTemplate} /></List>
);

export const sheetOptionsSection = (s: GridStore) => (
  <List><Tool store={s} id="page.printGridlines" /><Tool store={s} id="page.printHeadings" /><Tool store={s} id="page.headerFooter" icon={PanelTop} /></List>
);

export const printSection = (s: GridStore) => (
  <List><Tool store={s} id="file.exportPdf" icon={Printer} /><Tool store={s} id="file.savePdf" icon={FileDown} /><Tool store={s} id="file.openPdf" icon={FileText} /><Tool store={s} id="page.setup" icon={Settings2} /></List>
);

export const previewSection = (s: GridStore) => (
  <List><Tool store={s} id="file.exportPdf" icon={Eye} label={T('Print preview…')} /><Tool store={s} id="file.savePdf" icon={FileDown} /></List>
);
