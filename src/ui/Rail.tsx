import type React from 'react';
import type { ReactNode } from 'react';
import {
  AlignCenter, AlignLeft, AlignRight, AlignVerticalJustifyCenter, AlignVerticalJustifyEnd, AlignVerticalJustifyStart, ArrowDownAZ, ArrowUpAZ, BarChart3, Bold, BookOpen,
  Calculator, ClipboardPaste, Columns3, Copy, Database, ChevronsLeft, ChevronsRight, Eraser, Eye, File, FileDown, FilePlus2, Filter, FilterX, FolderOpen, FunctionSquare,
  Grid3x3, Group, Home, IndentDecrease, IndentIncrease, Italic, Link2, ListChecks, ListX, Lock, Merge, MessageSquare, Moon, Paintbrush, PanelLeftClose, Percent, Plus,
  Printer, Redo2, RefreshCw, Rows3, Save, Scissors, Search, Sigma, Snowflake, SplitSquareHorizontal, Strikethrough, Table2, Target, TextCursorInput, Underline, Undo2, Ungroup, WrapText,
  Image, Omega, Wand2, Brush, Code2, ZoomIn, ZoomOut, DollarSign, Hash, CalendarDays, Clock, Type, Sparkles, Palette, PaintBucket, Baseline, type LucideIcon,
} from 'lucide-react';
import { command, run } from '../app/commands';
import type { GridStore } from '../app/store';
import { prettyKeys } from './keys';
import { FONT_NAMES } from './fonts';
import { panelSections } from './panelRegistry';
import { QUICK_CHARTS } from './charts/chartTypes';
import { N, T, Tr } from './i18n';

const S = 1.9;

const TABS: { id: string; label: string; icon: LucideIcon }[] = [
  { id: 'home', label: N('Home'), icon: Home },
  { id: 'insert', label: N('Insert'), icon: Plus },
  { id: 'formulas', label: N('Formulas'), icon: FunctionSquare },
  { id: 'page', label: N('Page'), icon: Printer },
  { id: 'data', label: N('Data'), icon: Database },
  { id: 'review', label: N('Review'), icon: MessageSquare },
  { id: 'view', label: N('View'), icon: Eye },
  { id: 'draw', label: N('Draw'), icon: Brush },
  { id: 'automate', label: N('Automate'), icon: Code2 },
  { id: 'file', label: N('File'), icon: File },
];

export const Rail = ({ store }: { store: GridStore }) => (
  <nav className="rail" aria-label={T('Tools')}>
    {TABS.map((t) => (
      <button key={t.id} className={store.view.panel === t.id ? 'on' : ''} aria-pressed={store.view.panel === t.id}
        onClick={() => { store.view.panel = store.view.panel === t.id ? null : t.id; store.emit(); }}>
        <t.icon size={20} strokeWidth={S} /><span>{Tr(t.label)}</span>
      </button>
    ))}
  </nav>
);

const Btn = ({ store, id, icon: Icon, label, iconOnly, arg }: { store: GridStore; id: string; icon?: LucideIcon; label?: string; iconOnly?: boolean; arg?: unknown }) => {
  const c = command(id);
  if (!c) return null;
  const on = c.active?.(store) ?? false;
  const enabled = c.enabled?.(store) ?? true;
  const text = Tr(label ?? c.label);
  const tip = c.keys ? `${text} (${prettyKeys(c.keys)})` : text;
  return (
    <button className={`tool${on ? ' on' : ''}${iconOnly ? ' icon' : ''}`} disabled={!enabled} aria-pressed={c.active ? on : undefined} aria-label={text} title={tip}
      onMouseDown={(e) => e.preventDefault()} onClick={() => run(store, id, arg)}>
      {Icon && <Icon size={16} strokeWidth={S} />}{!iconOnly && <span>{text}</span>}
    </button>
  );
};

const Group_ = ({ title, children }: { title: string; children: ReactNode }) => <section className="group"><h3>{Tr(title)}</h3>{children}</section>;
const Row = ({ children }: { children: ReactNode }) => <div className="iconrow">{children}</div>;
const List = ({ children }: { children: ReactNode }) => <div className="toollist">{children}</div>;

const COLORS = ['#000000', '#404040', '#7f7f7f', '#bfbfbf', '#ffffff', '#c00000', '#ff0000', '#ffc000', '#ffff00', '#92d050', '#00b050', '#00b0f0', '#0070c0', '#2347ff', '#7030a0', '#f2dcdb', '#fde9d9', '#fff2cc', '#e2efda', '#ddebf7', '#e4dfec'];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];

const Swatches = ({ store, id, none }: { store: GridStore; id: string; none?: string }) => (
  <div className="swatches">
    {none && <button className="sw none" title={none} aria-label={none} onMouseDown={(e) => e.preventDefault()} onClick={() => run(store, id, '')} />}
    {COLORS.map((c) => <button key={c} className="sw" style={{ background: c }} title={c} aria-label={c} onMouseDown={(e) => e.preventDefault()} onClick={() => run(store, id, c)} />)}
    <label className="sw custom" title={T('More colors')} aria-label={T('More colors')}><input type="color" onChange={(e) => run(store, id, e.target.value)} /></label>
  </div>
);

const FontControls = ({ store }: { store: GridStore }) => {
  const st = store.style(store.sel.active.row, store.sel.active.col);
  const name = st.font?.name ?? 'Calibri';
  const size = st.font?.size ?? 11;
  return (
    <div className="fontctl">
      <select aria-label={T('Font')} value={name} onChange={(e) => run(store, 'format.fontName', e.target.value)}>
        {!FONT_NAMES.includes(name) && <option value={name}>{name}</option>}
        {FONT_NAMES.map((f) => <option key={f} value={f} style={{ fontFamily: f }}>{f}</option>)}
      </select>
      <input aria-label={T('Font size')} list="fontsizes" value={size} onChange={(e) => { const n = Number(e.target.value); if (n > 0 && n <= 409) run(store, 'format.fontSize', n); }} />
      <datalist id="fontsizes">{SIZES.map((s) => <option key={s} value={s} />)}</datalist>
    </div>
  );
};

const NumberFormatPicker = ({ store }: { store: GridStore }) => {
  const fmt = store.style(store.sel.active.row, store.sel.active.col).numFmt ?? 'General';
  const opts: [string, string][] = [[N('General'), N('General')], ['0.00', N('Number')], ['"$"#,##0.00', N('Currency')], ['_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)', N('Accounting')], ['m/d/yyyy', N('Short date')], ['dddd, mmmm d, yyyy', N('Long date')], ['h:mm:ss AM/PM', N('Time')], ['0.00%', N('Percentage')], ['# ?/?', N('Fraction')], ['0.00E+00', N('Scientific')], ['@', N('Text')]];
  return (
    <select aria-label={T('Number format')} className="wide" value={opts.some(([v]) => v === fmt) ? fmt : 'custom'} onChange={(e) => e.target.value === 'custom' ? run(store, 'cells.formatDialog') : run(store, 'number.format', e.target.value)}>
      {opts.map(([v, l]) => <option key={l} value={v}>{Tr(l)}</option>)}
      <option value="custom">{opts.some(([v]) => v === fmt) ? T('More number formats…') : T('Custom: {fmt}', { fmt })}</option>
    </select>
  );
};

const CURRENCIES: [string, string][] = [['$', N('Dollar ($)')], ['€', N('Euro (€)')], ['£', N('Pound (£)')], ['¥', N('Yen (¥)')], ['SAR', N('Saudi riyal (SAR)')], ['AED', N('UAE dirham (AED)')], ['₹', N('Rupee (₹)')], ['CHF', N('Swiss franc (CHF)')]];
const accountingFor = (sym: string): string => {
  const s = /^[A-Z]{3}$/.test(sym) ? `"${sym} "` : `"${sym}"`;
  return `_(${s}* #,##0.00_);_(${s}* \\(#,##0.00\\);_(${s}* "-"??_);_(@_)`;
};
const CurrencyPicker = ({ store }: { store: GridStore }) => (
  <select aria-label={T('Accounting currency')} className="wide" value="" onChange={(e) => { if (e.target.value) run(store, 'number.format', accountingFor(e.target.value)); }}>
    <option value="">{T('Accounting currency…')}</option>
    {CURRENCIES.map(([v, l]) => <option key={v} value={v}>{Tr(l)}</option>)}
  </select>
);

const panels: Record<string, (s: GridStore) => ReactNode> = {
  home: (s) => (<>
    <Group_ title={T('Clipboard')}><Row>
      <Btn store={s} id="edit.paste" icon={ClipboardPaste} iconOnly /><Btn store={s} id="edit.cut" icon={Scissors} iconOnly /><Btn store={s} id="edit.copy" icon={Copy} iconOnly />
      <Btn store={s} id="edit.formatPainter" icon={Paintbrush} iconOnly /><Btn store={s} id="edit.undo" icon={Undo2} iconOnly /><Btn store={s} id="edit.redo" icon={Redo2} iconOnly />
    </Row><Btn store={s} id="edit.pasteSpecial" /></Group_>
    <Group_ title={T('Font')}>
      <FontControls store={s} />
      <Row>
        <Btn store={s} id="format.bold" icon={Bold} iconOnly /><Btn store={s} id="format.italic" icon={Italic} iconOnly /><Btn store={s} id="format.underline" icon={Underline} iconOnly />
        <Btn store={s} id="format.strike" icon={Strikethrough} iconOnly /><Btn store={s} id="format.growFont" icon={Type} iconOnly label={T('Bigger')} /><Btn store={s} id="format.shrinkFont" icon={Baseline} iconOnly label={T('Smaller')} />
      </Row>
      <div className="mini"><Palette size={14} strokeWidth={S} />{' '}{T('Text color')}</div><Swatches store={s} id="format.fontColor" none={T('Automatic')} />
      <div className="mini"><PaintBucket size={14} strokeWidth={S} />{' '}{T('Fill color')}</div><Swatches store={s} id="format.fillColor" none={T('No fill')} />
    </Group_>
    <Group_ title={T('Borders')}><BorderButtons store={s} /></Group_>
    <Group_ title={T('Alignment')}>
      <Row>
        <Btn store={s} id="format.align.left" icon={AlignLeft} iconOnly /><Btn store={s} id="format.align.center" icon={AlignCenter} iconOnly /><Btn store={s} id="format.align.right" icon={AlignRight} iconOnly />
        <Btn store={s} id="format.valign.top" icon={AlignVerticalJustifyStart} iconOnly /><Btn store={s} id="format.valign.center" icon={AlignVerticalJustifyCenter} iconOnly /><Btn store={s} id="format.valign.bottom" icon={AlignVerticalJustifyEnd} iconOnly />
      </Row>
      <Row><Btn store={s} id="format.outdent" icon={IndentDecrease} iconOnly /><Btn store={s} id="format.indent" icon={IndentIncrease} iconOnly /><Btn store={s} id="format.wrap" icon={WrapText} iconOnly /><Btn store={s} id="format.mergeCenter" icon={Merge} iconOnly /></Row>
      <List><Btn store={s} id="format.mergeAcross" /><Btn store={s} id="format.unmerge" /></List>
    </Group_>
    <Group_ title={T('Number')}>
      <NumberFormatPicker store={s} />
      <Row>
        <Btn store={s} id="number.currency" icon={DollarSign} iconOnly /><Btn store={s} id="number.percent" icon={Percent} iconOnly /><Btn store={s} id="number.comma" icon={Hash} iconOnly />
        <Btn store={s} id="number.moreDecimals" icon={ChevronsRight} iconOnly /><Btn store={s} id="number.lessDecimals" icon={ChevronsLeft} iconOnly />
        <Btn store={s} id="number.shortDate" icon={CalendarDays} iconOnly /><Btn store={s} id="number.time" icon={Clock} iconOnly />
      </Row>
      <CurrencyPicker store={s} />
    </Group_>
    <Group_ title={T('Styles')}><List>
      <Btn store={s} id="insert.conditional" icon={Sparkles} label={T('Conditional formatting…')} /><Btn store={s} id="cells.formatTable" icon={Table2} label={T('Format as table…')} /><Btn store={s} id="cells.styles" icon={Palette} label={T('Cell styles…')} />
    </List></Group_>
    <Group_ title={T('Cells')}><List>
      <Btn store={s} id="cells.insertRows" icon={Rows3} label={T('Insert rows')} /><Btn store={s} id="cells.insertCols" icon={Columns3} label={T('Insert columns')} /><Btn store={s} id="cells.insertDialog" label={T('Insert cells…')} />
      <Btn store={s} id="cells.deleteRows" label={T('Delete rows')} /><Btn store={s} id="cells.deleteCols" label={T('Delete columns')} /><Btn store={s} id="cells.deleteDialog" label={T('Delete cells…')} />
      <Btn store={s} id="cells.formatDialog" icon={Grid3x3} label={T('Format cells…')} /><Btn store={s} id="cells.autofitCols" /><Btn store={s} id="cells.autofitRows" /><Btn store={s} id="cells.colWidth" /><Btn store={s} id="cells.rowHeight" />
      <Btn store={s} id="cells.hideRows" /><Btn store={s} id="cells.unhideRows" /><Btn store={s} id="cells.hideCols" /><Btn store={s} id="cells.unhideCols" />
    </List></Group_>
    <Group_ title={T('Editing')}><List>
      <Btn store={s} id="formula.autoSUM" icon={Sigma} /><Btn store={s} id="edit.fillDown" /><Btn store={s} id="edit.fillRight" /><Btn store={s} id="edit.series" /><Btn store={s} id="edit.flashFill" icon={Wand2} />
      <Btn store={s} id="edit.clearAll" icon={Eraser} /><Btn store={s} id="edit.clearFormats" /><Btn store={s} id="edit.clearContents" />
      <Btn store={s} id="data.sortAsc" icon={ArrowDownAZ} /><Btn store={s} id="data.sortDesc" icon={ArrowUpAZ} /><Btn store={s} id="data.filter" icon={Filter} />
      <Btn store={s} id="edit.find" icon={Search} /><Btn store={s} id="edit.replace" /><Btn store={s} id="edit.goto" /><Btn store={s} id="edit.gotoSpecial" />
    </List></Group_>
  </>),
  insert: (s) => (<>
    <Group_ title={T('Tables')}><List><Btn store={s} id="cells.formatTable" icon={Table2} label={T('Table…')} /><Btn store={s} id="data.pivot" icon={Grid3x3} label={T('PivotTable…')} /></List></Group_>
    <Group_ title={T('Charts')}><List><Btn store={s} id="insert.chart" icon={BarChart3} label={T('Recommended charts…')} /></List><ChartQuick store={s} /></Group_>
    <Group_ title={T('Illustrations')}><List><Btn store={s} id="insert.picture" icon={Image} /></List></Group_>
    <Group_ title={T('Links and notes')}><List><Btn store={s} id="insert.link" icon={Link2} /><Btn store={s} id="insert.note" icon={MessageSquare} /></List></Group_>
    <Group_ title={T('Symbols')}><List><Btn store={s} id="insert.symbol" icon={Omega} /><Btn store={s} id="formula.insertFunction" icon={FunctionSquare} /></List></Group_>
  </>),
  formulas: (s) => (<>
    <Group_ title={T('Function library')}><List>
      <Btn store={s} id="formula.insertFunction" icon={FunctionSquare} /><Btn store={s} id="formula.autoSUM" icon={Sigma} /><Btn store={s} id="formula.autoAVERAGE" /><Btn store={s} id="formula.autoCOUNT" /><Btn store={s} id="formula.autoMAX" /><Btn store={s} id="formula.autoMIN" />
    </List><CategoryButtons store={s} /></Group_>
    <Group_ title={T('Defined names')}><List><Btn store={s} id="formula.names" icon={BookOpen} /><Btn store={s} id="formula.defineName" /><Btn store={s} id="formula.createFromSelection" /></List></Group_>
    <Group_ title={T('Formula auditing')}><List>
      <Btn store={s} id="formula.tracePrecedents" /><Btn store={s} id="formula.traceDependents" /><Btn store={s} id="formula.showFormulas" icon={TextCursorInput} /><Btn store={s} id="formula.errorCheck" /><Btn store={s} id="formula.evaluate" />
    </List></Group_>
    <Group_ title={T('Calculation')}><List><Btn store={s} id="formula.calcNow" icon={Calculator} /><Btn store={s} id="formula.iterative" icon={RefreshCw} /></List></Group_>
  </>),
  data: (s) => (<>
    <Group_ title={T('Get data')}><List><Btn store={s} id="data.importCsv" icon={FileDown} /></List></Group_>
    <Group_ title={T('Sort and filter')}><List>
      <Btn store={s} id="data.sortAsc" icon={ArrowDownAZ} /><Btn store={s} id="data.sortDesc" icon={ArrowUpAZ} /><Btn store={s} id="data.sort" /><Btn store={s} id="data.filter" icon={Filter} /><Btn store={s} id="data.clearFilter" icon={FilterX} /><Btn store={s} id="data.reapply" />
    </List></Group_>
    <Group_ title={T('Data tools')}><List>
      <Btn store={s} id="data.textToColumns" icon={SplitSquareHorizontal} /><Btn store={s} id="edit.flashFill" icon={Wand2} /><Btn store={s} id="data.removeDuplicates" icon={ListX} /><Btn store={s} id="data.validation" icon={ListChecks} />
    </List></Group_>
    <Group_ title={T('Forecast')}><List><Btn store={s} id="data.goalSeek" icon={Target} /><Btn store={s} id="data.whatIfTable" /><Btn store={s} id="data.stats" /></List></Group_>
    <Group_ title={T('Outline')}><List><Btn store={s} id="cells.group" icon={Group} /><Btn store={s} id="cells.ungroup" icon={Ungroup} /><Btn store={s} id="data.subtotal" /></List></Group_>
  </>),
  review: (s) => (<>
    <Group_ title={T('Notes')}><List><Btn store={s} id="insert.note" icon={MessageSquare} /><Btn store={s} id="edit.clearNotes" /></List></Group_>
    <Group_ title={T('Protect')}><List><Btn store={s} id="sheet.protect" icon={Lock} /></List></Group_>
    <Group_ title={T('Check')}><List><Btn store={s} id="formula.errorCheck" /><Btn store={s} id="file.properties" /></List></Group_>
  </>),
  view: (s) => (<>
    <Group_ title={T('Show')}><List><Btn store={s} id="view.gridlines" icon={Grid3x3} /><Btn store={s} id="view.headings" /><Btn store={s} id="formula.showFormulas" /><Btn store={s} id="view.theme" icon={Moon} /><Btn store={s} id="view.rtl" /></List></Group_>
    <Group_ title={T('Freeze')}><List><Btn store={s} id="view.freezePanes" icon={Snowflake} /><Btn store={s} id="view.freezeTopRow" /><Btn store={s} id="view.freezeFirstCol" /><Btn store={s} id="view.unfreeze" /></List></Group_>
    <Group_ title={T('Zoom')}><Row><Btn store={s} id="view.zoomOut" icon={ZoomOut} iconOnly /><Btn store={s} id="view.zoomIn" icon={ZoomIn} iconOnly /></Row><List><Btn store={s} id="view.zoom100" /><Btn store={s} id="view.zoomSelection" /></List></Group_>
    <Group_ title={T('Sheets')}><List><Btn store={s} id="sheet.insert" /><Btn store={s} id="sheet.duplicate" /><Btn store={s} id="sheet.rename" /><Btn store={s} id="sheet.delete" /></List></Group_>
  </>),
  file: (s) => (<>
    <Group_ title={T('Workbook')}><List>
      <Btn store={s} id="file.new" icon={FilePlus2} /><Btn store={s} id="file.open" icon={FolderOpen} /><Btn store={s} id="file.save" icon={Save} /><Btn store={s} id="file.saveAs" /><Btn store={s} id="file.templates" />
    </List></Group_>
    <Group_ title={T('Export')}><List><Btn store={s} id="file.exportCsv" icon={FileDown} /><Btn store={s} id="file.exportPdf" icon={Printer} /></List></Group_>
    <Group_ title={T('Info')}><List><Btn store={s} id="file.properties" /><Btn store={s} id="help.shortcuts" /><Btn store={s} id="help.about" /></List></Group_>
  </>),
};

const BorderButtons = ({ store }: { store: GridStore }) => (
  <div className="borderbtns">
    {([['all', N('All')], ['outside', N('Outside')], ['thickBox', N('Thick box')], ['inside', N('Inside')], ['top', N('Top')], ['bottom', N('Bottom')], ['left', N('Left')], ['right', N('Right')], ['doubleBottom', N('Double bottom')], ['topBottom', N('Top and bottom')], ['none', N('None')]] as const).map(([id, label]) => (
      <button key={id} className="borderbtn" title={Tr(command(`border.${id}`)?.label)} aria-label={Tr(command(`border.${id}`)?.label)} onMouseDown={(e) => e.preventDefault()} onClick={() => run(store, `border.${id}`)}>
        <BorderIcon kind={id} /><span>{Tr(label)}</span>
      </button>
    ))}
  </div>
);

const BorderIcon = ({ kind }: { kind: string }) => {
  const thin = 'currentColor';
  const dash = <path d="M3 3h12v12H3z" stroke={thin} strokeOpacity=".3" strokeDasharray="1.5 1.5" fill="none" />;
  const ln = (d: string, w = 1.6) => <path d={d} stroke={thin} strokeWidth={w} fill="none" strokeLinecap="square" />;
  const shapes: Record<string, React.ReactElement> = {
    all: <>{ln('M3 3h12v12H3zM9 3v12M3 9h12')}</>, outside: <>{dash}{ln('M3 3h12v12H3z')}</>, thickBox: <>{dash}{ln('M3 3h12v12H3z', 3)}</>, inside: <>{dash}{ln('M9 3v12M3 9h12')}</>,
    top: <>{dash}{ln('M3 3h12')}</>, bottom: <>{dash}{ln('M3 15h12')}</>, left: <>{dash}{ln('M3 3v12')}</>, right: <>{dash}{ln('M15 3v12')}</>,
    doubleBottom: <>{dash}{ln('M3 13.5h12M3 16h12', 1.1)}</>, topBottom: <>{dash}{ln('M3 3h12M3 15h12')}</>, none: dash,
  };
  return <svg viewBox="0 0 18 18" width="16" height="16" aria-hidden="true">{shapes[kind]}</svg>;
};

const ChartQuick = ({ store }: { store: GridStore }) => (
  <div className="chartquick">
    {QUICK_CHARTS.map(([t, l]) => (
      <button key={t} className="tool" onClick={() => run(store, 'insert.chart', { type: t, quick: true })}>{Tr(l)}</button>
    ))}
  </div>
);

const CategoryButtons = ({ store }: { store: GridStore }) => (
  <div className="styles">
    {[N('Financial'), N('Logical'), N('Text'), N('Date & Time'), N('Lookup & Reference'), N('Math & Trig'), N('Statistical'), N('Engineering'), N('Information'), N('Database')].map((c) => (
      <button key={c} className="stylechip" onClick={() => run(store, 'formula.insertFunction', { category: c })}>{Tr(c)}</button>
    ))}
  </div>
);

export const ToolPanel = ({ store }: { store: GridStore }) => {
  const id = store.view.panel;
  if (!id || (!panels[id] && !panelSections(id).length)) return null;
  return (
    <aside className="toolpanel" aria-label={T('{name} tools', { name: Tr(TABS.find((t) => t.id === id)?.label ?? panelSections(id)[0]?.title) })}>
      <div className="phead"><h2>{TABS.find((t) => t.id === id)?.label ?? panelSections(id)[0]?.title}</h2>
        <button className="ico" aria-label={T('Close panel')} onClick={() => { store.view.panel = null; store.emit(); }}><PanelLeftClose size={16} strokeWidth={S} /></button></div>
      <div className="pbody">{panels[id]?.(store)}{panelSections(id).map((sec) => <Group_ key={sec.title} title={Tr(sec.title)}>{sec.render(store)}</Group_>)}</div>
    </aside>
  );
};
