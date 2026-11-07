import { useMemo, useState } from 'react';
import { Ban, Grid3x3, Square } from 'lucide-react';
import { MAX_COLS, MAX_ROWS } from '../../engine/types';
import { mergeStyle, type BorderSide, type BorderStyle, type CellRange, type CellStyle, type HAlign, type Underline, type VAlign } from '../../model/workbook';
import type { GridStore } from '../../app/store';
import { FONT_NAMES } from '../fonts';
import { Dialog, Field, type DialogProps } from './Dialog';
import { buildCode, NumberTab, parseCode, type NumOpts } from './numberFormat';
import { Check, cleanStyle, ColorPick, cssOf, overlaps, Tabs } from './shared';

type Tab = 'number' | 'align' | 'font' | 'border' | 'fill' | 'protect';
const TABS: [Tab, string][] = [['number', 'Number'], ['align', 'Alignment'], ['font', 'Font'], ['border', 'Border'], ['fill', 'Fill'], ['protect', 'Protection']];

export const BORDER_STYLES: BorderStyle[] = ['hair', 'dotted', 'dashDotDot', 'dashDot', 'dashed', 'thin', 'mediumDashDotDot', 'slantDashDot', 'mediumDashDot', 'mediumDashed', 'medium', 'thick', 'double'];
const PATTERNS: [string, string][] = [
  ['solid', 'Solid'], ['gray125', 'Light dots'], ['gray0625', 'Fine dots'], ['lightGray', 'Light gray'], ['mediumGray', 'Medium gray'], ['darkGray', 'Dark gray'],
  ['darkHorizontal', 'Horizontal stripes'], ['darkVertical', 'Vertical stripes'], ['darkDown', 'Diagonal down'], ['darkUp', 'Diagonal up'], ['darkGrid', 'Grid'], ['darkTrellis', 'Trellis'],
  ['lightHorizontal', 'Thin horizontal'], ['lightVertical', 'Thin vertical'], ['lightDown', 'Thin diagonal down'], ['lightUp', 'Thin diagonal up'], ['lightGrid', 'Thin grid'], ['lightTrellis', 'Thin trellis'],
];
const H_ALIGNS: [HAlign, string][] = [['general', 'General'], ['left', 'Left'], ['center', 'Center'], ['right', 'Right'], ['fill', 'Fill'], ['justify', 'Justify'], ['centerContinuous', 'Center across selection'], ['distributed', 'Distributed']];
const V_ALIGNS: [VAlign, string][] = [['top', 'Top'], ['center', 'Center'], ['bottom', 'Bottom'], ['justify', 'Justify'], ['distributed', 'Distributed']];
const UNDERLINES: [Underline | '', string][] = [['', 'None'], ['single', 'Single'], ['double', 'Double'], ['singleAccounting', 'Single accounting'], ['doubleAccounting', 'Double accounting']];

type Edge = 'top' | 'bottom' | 'left' | 'right' | 'insideH' | 'insideV' | 'diagonalUp' | 'diagonalDown';
/** undefined = leave as is, null = remove. */
type Edges = Partial<Record<Edge, BorderSide | null>>;

/** SVG dash pattern that approximates each border style in the pickers. */
export const dashOf = (s: BorderStyle): { w: number; dash?: string; double?: boolean } => {
  const w = s === 'thick' ? 3 : s.startsWith('medium') || s === 'slantDashDot' ? 2 : s === 'hair' ? 0.6 : 1;
  const dash: Partial<Record<BorderStyle, string>> = { dotted: '1 2', hair: '1 1', dashed: '4 2', mediumDashed: '6 3', dashDot: '5 2 1 2', mediumDashDot: '6 2 2 2', slantDashDot: '6 2 2 2', dashDotDot: '5 2 1 2 1 2', mediumDashDotDot: '6 2 2 2 2 2' };
  return { w, dash: dash[s], double: s === 'double' };
};

export const LineSample = ({ style, color = 'currentColor' }: { style: BorderStyle; color?: string }) => {
  const d = dashOf(style);
  return (
    <svg width="64" height="10" aria-hidden="true">
      {d.double ? <><line x1="2" y1="3" x2="62" y2="3" stroke={color} strokeWidth="1" /><line x1="2" y1="7" x2="62" y2="7" stroke={color} strokeWidth="1" /></>
        : <line x1="2" y1="5" x2="62" y2="5" stroke={color} strokeWidth={d.w} strokeDasharray={d.dash} />}
    </svg>
  );
};

const degToRotation = (deg: number): number => (deg >= 0 ? deg : 90 - deg);
const rotationToDeg = (rot: number | undefined): number => (rot === undefined || rot === 255 ? 0 : rot <= 90 ? rot : 90 - rot);

/** Splits a range into up to 9 parts (corners, edges, middle) where the border logic is uniform. */
const bands = (a: number, b: number): [number, number][] => (a === b ? [[a, a]] : b - a === 1 ? [[a, a], [b, b]] : [[a, a], [a + 1, b - 1], [b, b]]);

const applyEdges = (store: GridStore, r: CellRange, e: Edges): void => {
  const pick = (outer: boolean, out: BorderSide | null | undefined, inner: BorderSide | null | undefined) => (outer ? out : inner);
  for (const [r1, r2] of bands(r.r1, r.r2)) for (const [c1, c2] of bands(r.c1, r.c2)) {
    const sides = {
      top: pick(r1 === r.r1, e.top, e.insideH),
      bottom: pick(r2 === r.r2, e.bottom, e.insideH),
      left: pick(c1 === r.c1, e.left, e.insideV),
      right: pick(c2 === r.c2, e.right, e.insideV),
      diagonalUp: e.diagonalUp,
      diagonalDown: e.diagonalDown,
    };
    if (Object.values(sides).every((x) => x === undefined)) continue;
    store.styleRange({ r1, r2, c1, c2 }, (s) => {
      const b = { ...s.border };
      for (const [k, v] of Object.entries(sides) as [keyof typeof sides, BorderSide | null | undefined][]) {
        if (v === null) delete b[k]; else if (v) b[k] = { ...v };
      }
      return cleanStyle({ ...s, border: Object.keys(b).length ? b : undefined });
    });
  }
  // Neighbours also draw a shared edge, so clear theirs where the outline changed.
  const strip = (area: CellRange, key: 'top' | 'bottom' | 'left' | 'right') => {
    if (area.r1 < 0 || area.c1 < 0 || area.r2 >= MAX_ROWS || area.c2 >= MAX_COLS) return;
    store.styleRange(area, (s) => { if (!s.border?.[key]) return s; const b = { ...s.border }; delete b[key]; return cleanStyle({ ...s, border: Object.keys(b).length ? b : undefined }); });
  };
  if (e.top !== undefined) strip({ r1: r.r1 - 1, r2: r.r1 - 1, c1: r.c1, c2: r.c2 }, 'bottom');
  if (e.bottom !== undefined) strip({ r1: r.r2 + 1, r2: r.r2 + 1, c1: r.c1, c2: r.c2 }, 'top');
  if (e.left !== undefined) strip({ r1: r.r1, r2: r.r2, c1: r.c1 - 1, c2: r.c1 - 1 }, 'right');
  if (e.right !== undefined) strip({ r1: r.r1, r2: r.r2, c1: r.c2 + 1, c2: r.c2 + 1 }, 'left');
};

/** Patch merge where `undefined` fields in the patch mean "keep". Fields set to null are not used. */
const deepPatch = (base: CellStyle, p: CellStyle): CellStyle => cleanStyle(mergeStyle(base, p));

export const FormatCellsDialog = ({ store, data, close }: DialogProps<{ tab?: Tab }>) => {
  const { row, col } = store.sel.active;
  const base = store.style(row, col);
  const r = store.range;
  const [tab, setTab] = useState<Tab>(data?.tab ?? 'number');
  const [num, setNum] = useState<NumOpts>(() => parseCode(base.numFmt ?? 'General'));
  const [numTouched, setNumTouched] = useState(false);
  const [patch, setPatch] = useState<CellStyle>({});
  const [edges, setEdges] = useState<Edges>({});
  const [line, setLine] = useState<BorderStyle>('thin');
  const [lineColor, setLineColor] = useState<string | undefined>(undefined);
  const wasMerged = !!store.mergeAt(row, col) && store.sheet.merges.some((m) => m.r1 === r.r1 && m.c1 === r.c1 && m.r2 === r.r2 && m.c2 === r.c2);
  const [merge, setMerge] = useState(wasMerged);
  const [fillMode, setFillMode] = useState<'none' | 'color' | 'gradient'>(base.fill?.gradient ? 'gradient' : base.fill?.color || base.fill?.pattern ? 'color' : 'none');
  const [fillTouched, setFillTouched] = useState(false);

  const eff = useMemo(() => deepPatch(base, patch), [base, patch]);
  const sub = <K extends 'font' | 'align' | 'protection'>(k: K, v: NonNullable<CellStyle[K]>) => setPatch((p) => ({ ...p, [k]: { ...p[k], ...v } }));
  const setFill = (f: NonNullable<CellStyle['fill']>) => { setFillTouched(true); setPatch((p) => ({ ...p, fill: { ...p.fill, ...f } })); };
  const formats = useMemo(() => [...new Set(store.wb.styles.map((s) => s.numFmt).filter((x): x is string => !!x))].slice(0, 40), [store.wb.styles]);

  const shownEdge = (k: Edge): BorderSide | null | undefined => {
    if (k in edges) return edges[k];
    if (k === 'insideH' || k === 'insideV') return undefined;
    return base.border?.[k];
  };
  const toggleEdge = (k: Edge) => {
    const cur = shownEdge(k);
    const same = cur && cur.style === line && (cur.color ?? '') === (lineColor ?? '');
    setEdges((e) => ({ ...e, [k]: same ? null : { style: line, ...(lineColor ? { color: lineColor } : {}) } }));
  };
  const preset = (p: 'none' | 'outline' | 'inside') => {
    const side: BorderSide = { style: line, ...(lineColor ? { color: lineColor } : {}) };
    if (p === 'none') setEdges({ top: null, bottom: null, left: null, right: null, insideH: null, insideV: null, diagonalUp: null, diagonalDown: null });
    else if (p === 'outline') setEdges((e) => ({ ...e, top: side, bottom: side, left: side, right: side }));
    else setEdges((e) => ({ ...e, insideH: r.r1 !== r.r2 ? side : e.insideH, insideV: r.c1 !== r.c2 ? side : e.insideV }));
  };

  const ok = () => {
    let p: CellStyle = { ...patch };
    if (numTouched) p.numFmt = buildCode(num);
    if (fillTouched) {
      if (fillMode === 'none') p.fill = undefined;
      else if (fillMode === 'color') p.fill = { color: p.fill?.color ?? base.fill?.color, pattern: p.fill?.pattern ?? base.fill?.pattern, patternColor: p.fill?.patternColor ?? base.fill?.patternColor };
      else p.fill = { gradient: p.fill?.gradient ?? base.fill?.gradient ?? { angle: 0, stops: [{ pos: 0, color: '#ffffff' }, { pos: 1, color: '#0C7A4B' }] } };
      if (p.fill?.pattern === 'solid') p.fill = { ...p.fill, pattern: undefined };
    }
    const fillChange = fillTouched;
    const hasEdges = Object.keys(edges).length > 0;
    const mergeChange = merge !== wasMerged;
    const ranges = store.sel.ranges;
    if (mergeChange && merge && ranges.some((x) => store.sheet.tables.some((t) => overlaps(t.range, x)))) { store.notify("Cells in a table can't be merged."); return false; }
    const whole = ranges.some((x) => (x.r1 === 0 && x.r2 === MAX_ROWS - 1) || (x.c1 === 0 && x.c2 === MAX_COLS - 1));
    const grow = (x: CellRange) => ({ r1: Math.max(0, x.r1 - 1), c1: Math.max(0, x.c1 - 1), r2: Math.min(MAX_ROWS - 1, x.r2 + 1), c2: Math.min(MAX_COLS - 1, x.c2 + 1) });
    const scope = whole || mergeChange ? 'sheet' as const : ranges.map((x) => ({ area: hasEdges ? grow(x) : x }));
    store.change('Format cells', scope, () => {
      for (const x of ranges) {
        store.styleRange(x, (s) => {
          let next = deepPatch(s, p);
          if (fillChange) next = cleanStyle({ ...next, fill: p.fill ? cleanStyle({ fill: p.fill }).fill : undefined });
          if ('numFmt' in p && p.numFmt === 'General') { const { numFmt: _drop, ...rest } = next; void _drop; next = rest; }
          return next;
        });
        if (hasEdges) {
          const b = store.engine.usedBounds(store.sheetIndex);
          const clipped = { ...x, r2: x.r1 === 0 && x.r2 === MAX_ROWS - 1 ? Math.max(0, b.rows) : x.r2, c2: x.c1 === 0 && x.c2 === MAX_COLS - 1 ? Math.max(0, b.cols) : x.c2 };
          applyEdges(store, clipped, edges);
        }
        if (mergeChange) {
          store.sheet.merges = store.sheet.merges.filter((m) => !overlaps(m, x));
          if (merge && (x.r1 !== x.r2 || x.c1 !== x.c2)) store.sheet.merges.push({ ...x });
        }
      }
    });
    return true;
  };

  const f = eff.font ?? {};
  const a = eff.align ?? {};
  const fontStyle = f.bold && f.italic ? 'bi' : f.bold ? 'b' : f.italic ? 'i' : 'r';
  const deg = rotationToDeg(a.rotation);
  const edgeBtn = (k: Edge, label: string, icon: React.ReactNode) => {
    const on = !!shownEdge(k);
    return <button type="button" className={`tool icon${on ? ' on' : ''}`} aria-pressed={on} title={label} aria-label={label} onClick={() => toggleEdge(k)}>{icon}</button>;
  };
  const edgeIcon = (d: string) => <svg width="18" height="18" viewBox="0 0 18 18" aria-hidden="true"><rect x="2.5" y="2.5" width="13" height="13" fill="none" stroke="currentColor" strokeOpacity=".3" strokeDasharray="1 2" /><path d={d} stroke="currentColor" strokeWidth="1.9" fill="none" /></svg>;
  const sideLine = (k: Edge, x1: number, y1: number, x2: number, y2: number) => {
    const s = shownEdge(k);
    const hit = <line x1={x1} y1={y1} x2={x2} y2={y2} stroke="transparent" strokeWidth="14" style={{ cursor: 'pointer' }} onClick={() => toggleEdge(k)} />;
    if (!s) return <g key={k}><line x1={x1} y1={y1} x2={x2} y2={y2} stroke="var(--line)" strokeDasharray="2 3" />{hit}</g>;
    const d = dashOf(s.style);
    return <g key={k}><line x1={x1} y1={y1} x2={x2} y2={y2} stroke={s.color ?? 'var(--ink)'} strokeWidth={d.double ? 3 : d.w + 0.4} strokeDasharray={d.dash} />{hit}</g>;
  };
  const multiR = r.r1 !== r.r2, multiC = r.c1 !== r.c2;

  return (
    <Dialog title="Format cells" close={close} onOk={ok} wide>
      <Tabs tabs={TABS} value={tab} onChange={setTab} />
      {tab === 'number' && <NumberTab o={num} set={(p) => { setNumTouched(true); setNum((o) => ({ ...o, ...p })); }} sample={store.value(row, col)} date1904={store.wb.date1904} formats={formats} />}

      {tab === 'align' && (
        <div className="b-col">
          <div className="grid2">
            <Field label="Horizontal"><select value={a.h ?? 'general'} onChange={(e) => sub('align', { h: e.target.value as HAlign })}>{H_ALIGNS.map(([v, l]) => <option key={v} value={v}>{l}</option>)}</select></Field>
            <Field label="Vertical"><select value={a.v ?? 'bottom'} onChange={(e) => sub('align', { v: e.target.value as VAlign })}>{V_ALIGNS.map(([v, l]) => <option key={v} value={v}>{l}</option>)}</select></Field>
          </div>
          <Field label="Indent"><input type="number" min={0} max={250} value={a.indent ?? 0} onChange={(e) => sub('align', { indent: Math.max(0, Math.min(250, Math.trunc(+e.target.value || 0))) || undefined })} /></Field>
          <div className="b-checks">
            <Check label="Wrap text" checked={!!a.wrap} onChange={(v) => sub('align', { wrap: v || undefined, shrink: v ? undefined : a.shrink })} />
            <Check label="Shrink to fit" checked={!!a.shrink} disabled={!!a.wrap} onChange={(v) => sub('align', { shrink: v || undefined })} />
            <Check label="Merge cells" checked={merge} disabled={!multiR && !multiC && !wasMerged} onChange={setMerge} />
            <Check label="Right-to-left" checked={!!a.rtl} onChange={(v) => sub('align', { rtl: v || undefined })} />
          </div>
          <Field label="Text angle">
            <div className="b-angle">
              <input type="range" min={-90} max={90} value={deg} disabled={a.rotation === 255} onChange={(e) => sub('align', { rotation: +e.target.value ? degToRotation(+e.target.value) : undefined })} aria-label="Text angle in degrees" />
              <input type="number" min={-90} max={90} value={deg} disabled={a.rotation === 255} onChange={(e) => { const v = Math.max(-90, Math.min(90, Math.trunc(+e.target.value || 0))); sub('align', { rotation: v ? degToRotation(v) : undefined }); }} />
              <span className="muted small">degrees</span>
            </div>
          </Field>
          <Check label="Stack letters vertically" checked={a.rotation === 255} onChange={(v) => sub('align', { rotation: v ? 255 : undefined })} />
          <div className="preview b-alignprev" style={{ justifyContent: a.h === 'center' || a.h === 'centerContinuous' ? 'center' : a.h === 'right' ? 'flex-end' : 'flex-start', alignItems: a.v === 'top' ? 'flex-start' : a.v === 'center' ? 'center' : 'flex-end', direction: a.rtl ? 'rtl' : undefined }}>
            <span style={{ display: 'inline-block', transform: a.rotation === 255 ? undefined : `rotate(${-deg}deg)`, writingMode: a.rotation === 255 ? 'vertical-rl' : undefined, textOrientation: a.rotation === 255 ? 'upright' : undefined, paddingInlineStart: (a.indent ?? 0) * 9 }}>Text</span>
          </div>
        </div>
      )}

      {tab === 'font' && (
        <div className="b-col">
          <div className="b-fontrow">
            <Field label="Font">
              <input list="b-fontnames" value={f.name ?? 'Calibri'} onChange={(e) => sub('font', { name: e.target.value || undefined })} />
              <datalist id="b-fontnames">{FONT_NAMES.map((n) => <option key={n} value={n} />)}</datalist>
            </Field>
            <Field label="Style">
              <select value={fontStyle} onChange={(e) => sub('font', { bold: e.target.value.includes('b') || undefined, italic: e.target.value.includes('i') || undefined })}>
                <option value="r">Regular</option><option value="i">Italic</option><option value="b">Bold</option><option value="bi">Bold italic</option>
              </select>
            </Field>
            <Field label="Size"><input type="number" min={1} max={409} step={0.5} value={f.size ?? 11} onChange={(e) => { const v = +e.target.value; if (v >= 1 && v <= 409) sub('font', { size: v }); }} /></Field>
          </div>
          <Field label="Underline"><select value={f.underline ?? ''} onChange={(e) => sub('font', { underline: (e.target.value || undefined) as Underline | undefined })}>{UNDERLINES.map(([v, l]) => <option key={v} value={v}>{l}</option>)}</select></Field>
          <Field label="Color"><ColorPick value={f.color} none="Automatic" onChange={(c) => sub('font', { color: c })} /></Field>
          <div className="b-checks">
            <Check label="Strikethrough" checked={!!f.strike} onChange={(v) => sub('font', { strike: v || undefined })} />
            <Check label="Superscript" checked={f.vertAlign === 'superscript'} onChange={(v) => sub('font', { vertAlign: v ? 'superscript' : undefined })} />
            <Check label="Subscript" checked={f.vertAlign === 'subscript'} onChange={(v) => sub('font', { vertAlign: v ? 'subscript' : undefined })} />
          </div>
          <div className="preview b-fontprev"><span style={cssOf({ font: f })}>The quick fox 123</span></div>
        </div>
      )}

      {tab === 'border' && (
        <div className="b-split">
          <div className="b-col">
            <Field label="Line">
              <div className="b-list b-lines" role="listbox">
                {BORDER_STYLES.map((s) => <button key={s} role="option" aria-selected={line === s} aria-label={s} className={line === s ? 'on' : ''} onClick={() => setLine(s)}><LineSample style={s} color={lineColor} /></button>)}
              </div>
            </Field>
            <Field label="Color"><ColorPick value={lineColor} none="Automatic" onChange={setLineColor} /></Field>
          </div>
          <div className="b-col">
            <div className="iconrow b-presets">
              <button type="button" className="tool" onClick={() => preset('none')}><Ban size={16} strokeWidth={1.9} /> None</button>
              <button type="button" className="tool" onClick={() => preset('outline')}><Square size={16} strokeWidth={1.9} /> Outline</button>
              <button type="button" className="tool" disabled={!multiR && !multiC} onClick={() => preset('inside')}><Grid3x3 size={16} strokeWidth={1.9} /> Inside</button>
            </div>
            <div className="b-borderbox">
              <div className="b-edgecol">
                {edgeBtn('top', 'Top', edgeIcon('M2.5 2.5h13'))}
                {multiR && edgeBtn('insideH', 'Inside horizontal', edgeIcon('M2.5 9h13'))}
                {edgeBtn('bottom', 'Bottom', edgeIcon('M2.5 15.5h13'))}
                {edgeBtn('diagonalUp', 'Diagonal up', edgeIcon('M2.5 15.5l13-13'))}
              </div>
              <svg className="b-bprev" viewBox="0 0 160 120" role="img" aria-label="Border preview">
                <rect x="20" y="15" width="120" height="90" fill="var(--sheet)" />
                {multiR && sideLine('insideH', 20, 60, 140, 60)}
                {multiC && sideLine('insideV', 80, 15, 80, 105)}
                {sideLine('diagonalDown', 20, 15, 140, 105)}
                {sideLine('diagonalUp', 20, 105, 140, 15)}
                {sideLine('top', 20, 15, 140, 15)}
                {sideLine('bottom', 20, 105, 140, 105)}
                {sideLine('left', 20, 15, 20, 105)}
                {sideLine('right', 140, 15, 140, 105)}
                <text x="50" y="45" fontSize="11" fill="var(--muted)" textAnchor="middle">Text</text>
                {multiC && <text x="110" y="45" fontSize="11" fill="var(--muted)" textAnchor="middle">Text</text>}
              </svg>
              <div className="b-edgecol">
                {edgeBtn('left', 'Left', edgeIcon('M2.5 2.5v13'))}
                {multiC && edgeBtn('insideV', 'Inside vertical', edgeIcon('M9 2.5v13'))}
                {edgeBtn('right', 'Right', edgeIcon('M15.5 2.5v13'))}
                {edgeBtn('diagonalDown', 'Diagonal down', edgeIcon('M2.5 2.5l13 13'))}
              </div>
            </div>
            <p className="muted small">Pick a line and color, then tap an edge or a preset.</p>
          </div>
        </div>
      )}

      {tab === 'fill' && (
        <div className="b-col">
          <div className="b-seg" role="radiogroup">
            {(['none', 'color', 'gradient'] as const).map((m) => <button key={m} role="radio" aria-checked={fillMode === m} className={fillMode === m ? 'on' : ''} onClick={() => { setFillTouched(true); setFillMode(m); }}>{m === 'none' ? 'No fill' : m === 'color' ? 'Color' : 'Gradient'}</button>)}
          </div>
          {fillMode === 'color' && <>
            <Field label="Background color"><ColorPick value={eff.fill?.color} onChange={(c) => setFill({ color: c })} /></Field>
            <div className="grid2">
              <Field label="Pattern"><select value={eff.fill?.pattern ?? 'solid'} onChange={(e) => setFill({ pattern: e.target.value })}>{PATTERNS.map(([v, l]) => <option key={v} value={v}>{l}</option>)}</select></Field>
              <Field label="Pattern color"><input type="color" value={eff.fill?.patternColor ?? '#000000'} disabled={(eff.fill?.pattern ?? 'solid') === 'solid'} onChange={(e) => setFill({ patternColor: e.target.value })} /></Field>
            </div>
          </>}
          {fillMode === 'gradient' && (() => {
            const g = eff.fill?.gradient ?? { angle: 0, stops: [{ pos: 0, color: '#ffffff' }, { pos: 1, color: '#0C7A4B' }] };
            const c1 = g.stops[0]?.color ?? '#ffffff', c2 = g.stops.at(-1)?.color ?? '#0C7A4B';
            const setG = (p: { angle?: number; c1?: string; c2?: string }) => setFill({ gradient: { angle: p.angle ?? g.angle, stops: [{ pos: 0, color: p.c1 ?? c1 }, { pos: 1, color: p.c2 ?? c2 }] } });
            return <>
              <div className="grid3">
                <Field label="Color 1"><input type="color" value={c1} onChange={(e) => setG({ c1: e.target.value })} /></Field>
                <Field label="Color 2"><input type="color" value={c2} onChange={(e) => setG({ c2: e.target.value })} /></Field>
                <Field label="Direction"><select value={g.angle} onChange={(e) => setG({ angle: +e.target.value })}><option value={0}>Left to right</option><option value={90}>Top to bottom</option><option value={45}>Diagonal down</option><option value={135}>Diagonal up</option></select></Field>
              </div>
            </>;
          })()}
          <div className="preview b-fillprev" style={fillMode === 'none' ? undefined : fillMode === 'gradient' ? cssOf({ fill: { gradient: eff.fill?.gradient ?? { angle: 0, stops: [{ pos: 0, color: '#ffffff' }, { pos: 1, color: '#0C7A4B' }] } } }) : { background: eff.fill?.color ?? eff.fill?.patternColor }} />
        </div>
      )}

      {tab === 'protect' && (
        <div className="b-col">
          <Check label="Locked" checked={eff.protection?.locked !== false} onChange={(v) => sub('protection', { locked: v ? undefined : false })} />
          <Check label="Hide formulas" checked={!!eff.protection?.hidden} onChange={(v) => sub('protection', { hidden: v || undefined })} />
          <p className="muted small">These take effect after you protect the sheet (Review, Protect sheet).</p>
        </div>
      )}
    </Dialog>
  );
};
