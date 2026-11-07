import type { ReactNode } from 'react';
import {
  Accessibility, AlignCenter, AlignLeft, AlignRight, Bold, ChevronDown, ChevronUp, Eraser, Eye, EyeOff, Gauge, Highlighter, LockOpen, MessageSquarePlus,
  MessagesSquare, MousePointer2, Pen, PenTool, ScanText, Sigma, SpellCheck, SquareCheck, SquareX, Trash2, Italic, type LucideIcon,
} from 'lucide-react';
import { command, run } from '../../app/commands';
import type { GridStore } from '../../app/store';
import type { Shape } from '../../model/workbook';
import { prettyKeys } from '../../ui/keys';
import { setUi, ui, type DrawTool } from './state';
import { SHAPE_KINDS, updateShape } from './ops';
import { shapeIcon } from './ShapeLayer';

const S = 1.9;

const Tool = ({ store, id, icon: Icon, label }: { store: GridStore; id: string; icon?: LucideIcon; label?: string }) => {
  const c = command(id);
  if (!c) return null;
  const on = c.active?.(store) ?? false;
  const text = label ?? c.label;
  return (
    <button className={`tool${on ? ' on' : ''}`} disabled={c.enabled ? !c.enabled(store) : false} aria-pressed={c.active ? on : undefined} title={c.keys ? `${text} (${prettyKeys(c.keys)})` : text}
      onMouseDown={(e) => e.preventDefault()} onClick={() => run(store, id)}>
      {Icon && <Icon size={16} strokeWidth={S} />}<span>{text}</span>
    </button>
  );
};
const List = ({ children }: { children: ReactNode }) => <div className="toollist">{children}</div>;

export const commentsSection = (s: GridStore) => (
  <List>
    <Tool store={s} id="comment.new" icon={MessageSquarePlus} />
    <Tool store={s} id="comment.pane" icon={MessagesSquare} />
    <Tool store={s} id="comment.prev" icon={ChevronUp} /><Tool store={s} id="comment.next" icon={ChevronDown} />
    <Tool store={s} id="comment.resolve" /><Tool store={s} id="comment.delete" icon={Trash2} />
  </List>
);

export const notesSection = (s: GridStore) => (
  <List>
    <Tool store={s} id="note.delete" icon={Trash2} />
    <Tool store={s} id="note.showAll" icon={ui.showAllNotes ? EyeOff : Eye} label={ui.showAllNotes ? 'Hide all notes' : 'Show all notes'} />
    <Tool store={s} id="note.prev" icon={ChevronUp} /><Tool store={s} id="note.next" icon={ChevronDown} />
  </List>
);

export const checkSection = (s: GridStore) => (
  <List>
    <Tool store={s} id="review.spelling" icon={SpellCheck} /><Tool store={s} id="review.stats" icon={Sigma} />
    <Tool store={s} id="review.accessibility" icon={Accessibility} /><Tool store={s} id="review.performance" icon={Gauge} />
    <Tool store={s} id="review.altText" icon={ScanText} /><Tool store={s} id="review.editRanges" icon={LockOpen} />
  </List>
);

export const checkboxSection = (s: GridStore) => (
  <List><Tool store={s} id="insert.checkbox" icon={SquareCheck} /><Tool store={s} id="insert.removeCheckbox" icon={SquareX} /></List>
);

const FILLS = ['', '#ffffff', '#d6efe2', '#a9dfc2', '#0c7a4b', '#fff2cc', '#ffd8a8', '#ffc9c9', '#d0ebff', '#e5dbff', '#dee2e6', '#212529'];
const LINES = ['', '#0c7a4b', '#212529', '#868e96', '#e03131', '#f08c00', '#1971c2', '#7048e8'];
const TEXTS = ['#111111', '#ffffff', '#0c7a4b', '#e03131', '#1971c2', '#7048e8'];

const Swatch = ({ color, on, none, onPick }: { color: string; on: boolean; none: string; onPick: (c: string) => void }) => (
  <button className={`sw${color ? '' : ' none'}${on ? ' on' : ''}`} style={color ? { background: color } : undefined} title={color || none} aria-label={color || none} aria-pressed={on}
    onMouseDown={(e) => e.preventDefault()} onClick={() => onPick(color)} />
);

export const shapesSection = (s: GridStore) => {
  const sh = s.sheet.shapes?.find((x) => x.id === ui.shape);
  return (
    <>
      <div className="shapegallery" role="group" aria-label="Insert a shape">
        {SHAPE_KINDS.map(({ kind, label }) => (
          <button key={kind} className="tool icon" title={label} aria-label={label} onMouseDown={(e) => e.preventDefault()} onClick={() => run(s, 'insert.shape', kind)}>{shapeIcon(kind)}</button>
        ))}
      </div>
      {sh && <ShapeFormat store={s} sh={sh} />}
    </>
  );
};

const ShapeFormat = ({ store, sh }: { store: GridStore; sh: Shape }) => {
  const line = sh.kind === 'line' || sh.kind === 'arrowLine';
  const up = (patch: Partial<Shape>, label?: string) => updateShape(store, sh.id, patch, label);
  const f = sh.font ?? {};
  const ln = sh.line ?? { color: '#0c7a4b', width: 0, dash: 'solid' as const };
  return (
    <div className="shapefmt">
      <div className="mini">Selected shape</div>
      {!line && <>
        <div className="mini">Fill</div>
        <div className="swatches">{FILLS.map((c) => <Swatch key={c || 'none'} color={c} none="No fill" on={(sh.fill ?? '') === c} onPick={(c2) => up({ fill: c2 || undefined }, 'Shape fill')} />)}
          <label className="sw custom" title="More colors" aria-label="More fill colors"><input type="color" onChange={(e) => up({ fill: e.target.value }, 'Shape fill')} /></label></div>
      </>}
      <div className="mini">Outline</div>
      <div className="swatches">{LINES.map((c) => <Swatch key={c || 'none'} color={c} none="No outline" on={c ? ln.width > 0 && ln.color === c : ln.width === 0} onPick={(c2) => up({ line: c2 ? { ...ln, color: c2, width: ln.width || 1.5 } : undefined }, 'Shape outline')} />)}</div>
      <div className="shapefmt-row">
        <select aria-label="Outline width" value={ln.width} onChange={(e) => up({ line: { ...ln, width: +e.target.value } }, 'Shape outline')}>
          {[0, 0.75, 1, 1.5, 2, 3, 4.5, 6].map((w) => <option key={w} value={w}>{w ? `${w} px` : 'None'}</option>)}
        </select>
        <select aria-label="Outline dash" value={ln.dash ?? 'solid'} onChange={(e) => up({ line: { ...ln, width: ln.width || 1.5, dash: e.target.value as 'solid' } }, 'Shape outline')}>
          <option value="solid">Solid</option><option value="dash">Dashed</option><option value="dot">Dotted</option>
        </select>
      </div>
      {!line && <>
        <div className="mini">Text</div>
        <div className="shapefmt-row">
          <select aria-label="Text size" value={f.size ?? 11} onChange={(e) => up({ font: { ...f, size: +e.target.value } }, 'Shape text')}>
            {[8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72].map((n) => <option key={n} value={n}>{n}</option>)}
          </select>
          <button className={`tool icon${f.bold ? ' on' : ''}`} aria-label="Bold" aria-pressed={!!f.bold} onClick={() => up({ font: { ...f, bold: !f.bold || undefined } }, 'Shape text')}><Bold size={16} strokeWidth={S} /></button>
          <button className={`tool icon${f.italic ? ' on' : ''}`} aria-label="Italic" aria-pressed={!!f.italic} onClick={() => up({ font: { ...f, italic: !f.italic || undefined } }, 'Shape text')}><Italic size={16} strokeWidth={S} /></button>
          {(['left', 'center', 'right'] as const).map((a) => {
            const I = a === 'left' ? AlignLeft : a === 'center' ? AlignCenter : AlignRight;
            const cur = sh.align ?? (sh.kind === 'textBox' ? 'left' : 'center');
            return <button key={a} className={`tool icon${cur === a ? ' on' : ''}`} aria-label={`Align ${a}`} aria-pressed={cur === a} onClick={() => up({ align: a }, 'Shape text')}><I size={16} strokeWidth={S} /></button>;
          })}
        </div>
        <div className="swatches">{TEXTS.map((c) => <Swatch key={c} color={c} none="" on={(f.color ?? '#111111').toLowerCase() === c} onPick={(c2) => up({ font: { ...f, color: c2 === '#111111' ? undefined : c2 } }, 'Shape text')} />)}</div>
      </>}
    </div>
  );
};

const PEN_COLORS = ['#1f2a24', '#e03131', '#1971c2', '#0c7a4b', '#f08c00', '#7048e8'];
const HL_COLORS = ['#ffe066', '#8ce99a', '#74c0fc', '#ffa8a8', '#e599f7'];
const TOOLS: { id: DrawTool; label: string; icon: LucideIcon }[] = [
  { id: 'select', label: 'Select', icon: MousePointer2 }, { id: 'pen', label: 'Pen', icon: Pen }, { id: 'marker', label: 'Marker', icon: PenTool },
  { id: 'highlighter', label: 'Highlighter', icon: Highlighter }, { id: 'eraser', label: 'Eraser', icon: Eraser },
];

export const drawToolsSection = (s: GridStore) => (
  <div className="toollist">
    {TOOLS.map((t) => (
      <button key={t.id} className={`tool${ui.tool === t.id ? ' on' : ''}`} aria-pressed={ui.tool === t.id} onClick={() => setUi(s, { tool: t.id, shape: null, pinned: null })}>
        <t.icon size={16} strokeWidth={S} /><span>{t.label}</span>
      </button>
    ))}
  </div>
);

export const drawStyleSection = (s: GridStore) => {
  const hl = ui.tool === 'highlighter';
  const colors = hl ? HL_COLORS : PEN_COLORS;
  const cur = hl ? ui.hlColor : ui.penColor;
  return (
    <>
      <div className="mini">{hl ? 'Highlighter color' : 'Pen color'}</div>
      <div className="swatches">
        {colors.map((c) => <Swatch key={c} color={c} none="" on={cur === c} onPick={(c2) => setUi(s, hl ? { hlColor: c2 } : { penColor: c2 })} />)}
        <label className="sw custom" title="More colors" aria-label="More colors"><input type="color" value={cur} onChange={(e) => setUi(s, hl ? { hlColor: e.target.value } : { penColor: e.target.value })} /></label>
      </div>
      <div className="mini">Thickness</div>
      <div className="inkwidths" role="group" aria-label="Thickness">
        {[1, 2, 3, 5, 8].map((w) => (
          <button key={w} className={`tool icon${ui.penWidth === w ? ' on' : ''}`} aria-label={`${w} px`} aria-pressed={ui.penWidth === w} onClick={() => setUi(s, { penWidth: w })}>
            <span className="inkdot" style={{ width: Math.min(18, w * 2 + 2), height: Math.min(18, w * 2 + 2) }} />
          </button>
        ))}
      </div>
    </>
  );
};

export const drawInkSection = (s: GridStore) => (
  <div className="toollist">
    <button className="tool" disabled={!s.sheet.ink?.length} onClick={() => run(s, 'draw.clear')}><Trash2 size={16} strokeWidth={S} /><span>Clear ink on this sheet</span></button>
    <p className="muted small">Draw with a mouse, finger or pencil. After you use a pencil, fingers scroll the sheet instead of drawing.</p>
  </div>
);
