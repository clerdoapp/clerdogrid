// Model operations for comments, notes, checkboxes, shapes and ink. Every change is one undo step.
import { MAX_COLS, MAX_ROWS } from '../../engine/types';
import { cellKey, keyCol, keyRow, type CellRange, type CommentThread, type InkStroke, type Shape, type ShapeKind } from '../../model/workbook';
import { editRangeKey, inRange, type GridStore } from '../../app/store';
import { protectionHash } from '../../app/protection';

export const currentUser = (): string => {
  try { return localStorage.getItem('clerdogrid.user')?.trim() || 'You'; } catch { return 'You'; }
};
const now = (): string => new Date().toISOString();
let seq = 0;
export const newId = (prefix: string): string => `${prefix}${Date.now().toString(36)}${(seq++).toString(36)}`;

// ---------- Comments ----------

export const threadAt = (s: GridStore, row: number, col: number): CommentThread | undefined => s.sheet.comments?.get(cellKey(row, col));

const editThread = (s: GridStore, label: string, row: number, col: number, fn: (t: CommentThread | undefined, map: Map<number, CommentThread>) => void): void => {
  s.change(label, 'sheet', () => {
    const map = (s.sheet.comments ??= new Map());
    fn(map.get(cellKey(row, col)), map);
    if (!map.size) delete s.sheet.comments;
  });
};

export const addComment = (s: GridStore, row: number, col: number, text: string, author = currentUser()): void => {
  const t = text.trim();
  if (!t) return;
  editThread(s, 'New comment', row, col, (cur, map) => {
    if (cur) cur.replies.push({ author, text: t, time: now() });
    else map.set(cellKey(row, col), { id: newId('c'), author, text: t, time: now(), replies: [] });
  });
};

export const replyComment = (s: GridStore, row: number, col: number, text: string, author = currentUser()): void => {
  const t = text.trim();
  if (!t || !threadAt(s, row, col)) return;
  editThread(s, 'Reply', row, col, (cur) => { cur!.replies.push({ author, text: t, time: now() }); cur!.resolved = undefined; });
};

/** Edits the first message (index -1) or a reply. */
export const editComment = (s: GridStore, row: number, col: number, index: number, text: string): void => {
  const t = text.trim();
  if (!t || !threadAt(s, row, col)) return;
  editThread(s, 'Edit comment', row, col, (cur) => {
    if (index < 0) cur!.text = t; else if (cur!.replies[index]) cur!.replies[index]!.text = t;
  });
};

/** Deletes a reply, or the whole thread when index is -1. */
export const deleteComment = (s: GridStore, row: number, col: number, index = -1): void => {
  if (!threadAt(s, row, col)) return;
  editThread(s, index < 0 ? 'Delete comment' : 'Delete reply', row, col, (cur, map) => {
    if (index < 0) map.delete(cellKey(row, col)); else cur!.replies.splice(index, 1);
  });
};

export const resolveComment = (s: GridStore, row: number, col: number, resolved: boolean): void => {
  if (!threadAt(s, row, col)) return;
  editThread(s, resolved ? 'Resolve comment' : 'Reopen comment', row, col, (cur) => { cur!.resolved = resolved || undefined; });
};

/** Cells with a thread (or note), in reading order. */
export const sortedKeys = (keys: Iterable<number>): number[] => [...keys].sort((a, b) => keyRow(a) - keyRow(b) || keyCol(a) - keyCol(b));

/** Next or previous cell (by reading order) from the active cell among these keys. */
export const stepTo = (s: GridStore, keys: Iterable<number>, dir: 1 | -1): { row: number; col: number } | null => {
  const list = sortedKeys(keys);
  if (!list.length) return null;
  const cur = cellKey(s.sel.active.row, s.sel.active.col);
  const at = (k: number) => keyRow(k) * MAX_COLS + keyCol(k);
  const here = at(cur);
  const k = dir > 0 ? list.find((x) => at(x) > here) ?? list[0]! : [...list].reverse().find((x) => at(x) < here) ?? list[list.length - 1]!;
  return { row: keyRow(k), col: keyCol(k) };
};

// ---------- Notes ----------

export const deleteNote = (s: GridStore): void => {
  const keys = [...s.sheet.notes.keys()].filter((k) => s.sel.ranges.some((r) => inRange(r, keyRow(k), keyCol(k))));
  if (!keys.length) return;
  s.change('Delete note', 'sheet', () => { for (const k of keys) s.sheet.notes.delete(k); });
};

// ---------- Checkboxes ----------

/** Selected areas clipped to the used range, so whole columns don't create a million cells. */
const clipped = (s: GridStore): CellRange[] => {
  const b = s.engine.usedBounds(s.sheetIndex);
  return s.sel.ranges.map((r) => {
    const whole = (r.r1 === 0 && r.r2 === MAX_ROWS - 1) || (r.c1 === 0 && r.c2 === MAX_COLS - 1);
    return whole ? { r1: r.r1, c1: r.c1, r2: Math.min(r.r2, Math.max(r.r1, b.rows)), c2: Math.min(r.c2, Math.max(r.c1, b.cols)) } : r;
  }).filter((r) => (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1) <= 200_000);
};

export const isCheckbox = (s: GridStore, row: number, col: number): boolean => !!s.style(row, col).checkbox;

export const insertCheckboxes = (s: GridStore): void => {
  const areas = clipped(s);
  if (!areas.length) { s.notify('Select fewer cells.'); return; }
  if (areas.some((r) => { for (let row = r.r1; row <= r.r2; row++) for (let col = r.c1; col <= r.c2; col++) if (s.isLocked(row, col)) return true; return false; })) { s.notify('These cells are protected.'); return; }
  s.change('Checkbox', areas.map((area) => ({ area })), () => {
    for (const r of areas) {
      s.styleRange(r, (st) => ({ ...st, checkbox: true }));
      for (let row = r.r1; row <= r.r2; row++) for (let col = r.c1; col <= r.c2; col++) {
        const c = s.cell(row, col);
        if (!c?.f && (c?.v ?? null) === null) s.engine.setValue(s.sheetIndex, row, col, false);
      }
    }
  });
};

export const removeCheckboxes = (s: GridStore): void => {
  const areas = clipped(s);
  s.change('Remove checkboxes', areas.map((area) => ({ area })), () => {
    for (const r of areas) s.styleRange(r, (st) => { const { checkbox: _drop, ...rest } = st; void _drop; return rest; });
  });
};

/** Flips checkbox cells. With no cells given, flips the checkbox cells in the selection: all on unless all are on already. */
export const toggleCheckboxes = (s: GridStore, cells?: { row: number; col: number }[]): boolean => {
  const list = cells ?? (() => {
    const out: { row: number; col: number }[] = [];
    for (const r of clipped(s)) for (let row = r.r1; row <= r.r2; row++) for (let col = r.c1; col <= r.c2; col++) if (isCheckbox(s, row, col)) out.push({ row, col });
    return out;
  })();
  const targets = list.filter((p) => isCheckbox(s, p.row, p.col) && !s.cell(p.row, p.col)?.f && !s.isLocked(p.row, p.col));
  if (!targets.length) return false;
  const allOn = targets.every((p) => s.value(p.row, p.col) === true);
  const next = targets.length === 1 ? s.value(targets[0]!.row, targets[0]!.col) !== true : !allOn;
  s.change('Checkbox', targets.map((p) => ({ area: { r1: p.row, c1: p.col, r2: p.row, c2: p.col } })), () => {
    for (const p of targets) s.engine.setValue(s.sheetIndex, p.row, p.col, next);
  });
  return true;
};

// ---------- Shapes ----------

export const SHAPE_KINDS: { kind: ShapeKind; label: string }[] = [
  { kind: 'textBox', label: 'Text box' }, { kind: 'rect', label: 'Rectangle' }, { kind: 'roundRect', label: 'Rounded rectangle' }, { kind: 'ellipse', label: 'Oval' },
  { kind: 'triangle', label: 'Triangle' }, { kind: 'diamond', label: 'Diamond' }, { kind: 'pentagon', label: 'Pentagon' }, { kind: 'hexagon', label: 'Hexagon' },
  { kind: 'star5', label: 'Star' }, { kind: 'callout', label: 'Callout' }, { kind: 'rightArrow', label: 'Right arrow' }, { kind: 'leftArrow', label: 'Left arrow' },
  { kind: 'upArrow', label: 'Up arrow' }, { kind: 'downArrow', label: 'Down arrow' }, { kind: 'line', label: 'Line' }, { kind: 'arrowLine', label: 'Arrow line' },
];

export const defaultShape = (kind: ShapeKind, anchor: Shape['anchor']): Shape => {
  const line = kind === 'line' || kind === 'arrowLine';
  if (kind === 'textBox') return { id: newId('shape'), kind, anchor, text: '', line: { color: '#9AA79F', width: 1, dash: 'solid' }, font: { size: 11 }, align: 'left' };
  return { id: newId('shape'), kind, anchor: line ? { ...anchor, height: Math.round(anchor.height / 2) } : anchor, fill: line ? undefined : '#D6EFE2', line: { color: '#0C7A4B', width: line ? 2 : 1.5, dash: 'solid' }, font: { size: 11 }, align: 'center' };
};

export const addShape = (s: GridStore, shape: Shape): void => {
  s.change('Insert shape', 'sheet', () => { (s.sheet.shapes ??= []).push(shape); });
};

export const updateShape = (s: GridStore, id: string, patch: Partial<Shape>, label = 'Format shape'): void => {
  if (!s.sheet.shapes?.some((x) => x.id === id)) return;
  s.change(label, 'sheet', () => {
    const sh = s.sheet.shapes!.find((x) => x.id === id)!;
    Object.assign(sh, patch);
    for (const k of Object.keys(patch) as (keyof Shape)[]) if (patch[k] === undefined) delete sh[k];
  });
};

export const deleteShape = (s: GridStore, id: string): void => {
  if (!s.sheet.shapes?.some((x) => x.id === id)) return;
  s.change('Delete shape', 'sheet', () => {
    s.sheet.shapes = s.sheet.shapes!.filter((x) => x.id !== id);
    if (!s.sheet.shapes.length) delete s.sheet.shapes;
  });
};

/** Moves a shape one step up or down the stack, or to the top or bottom. */
export const orderShape = (s: GridStore, id: string, to: 'forward' | 'backward' | 'front' | 'back'): void => {
  const list = s.sheet.shapes;
  const i = list?.findIndex((x) => x.id === id) ?? -1;
  if (!list || i < 0) return;
  const j = to === 'forward' ? Math.min(list.length - 1, i + 1) : to === 'backward' ? Math.max(0, i - 1) : to === 'front' ? list.length - 1 : 0;
  if (i === j) return;
  s.change(to === 'forward' || to === 'front' ? 'Bring forward' : 'Send backward', 'sheet', () => {
    const arr = s.sheet.shapes!;
    const [sh] = arr.splice(i, 1);
    arr.splice(j, 0, sh!);
  });
};

// ---------- Ink ----------

export const addStroke = (s: GridStore, stroke: InkStroke): void => {
  if (stroke.points.length < 1) return;
  s.change('Draw', 'sheet', () => { (s.sheet.ink ??= []).push(stroke); });
};

export const eraseStrokes = (s: GridStore, ids: Set<string>): void => {
  if (!ids.size || !s.sheet.ink?.some((k) => ids.has(k.id))) return;
  s.change('Erase', 'sheet', () => {
    s.sheet.ink = s.sheet.ink!.filter((k) => !ids.has(k.id));
    if (!s.sheet.ink.length) delete s.sheet.ink;
  });
};

export const clearInk = (s: GridStore): void => {
  if (!s.sheet.ink?.length) return;
  s.change('Clear ink', 'sheet', () => { delete s.sheet.ink; });
};

/** Distance from point p to segment ab. */
const segDist = (p: [number, number], a: [number, number], b: [number, number]): number => {
  const dx = b[0] - a[0], dy = b[1] - a[1];
  const len = dx * dx + dy * dy;
  const t = len ? Math.max(0, Math.min(1, ((p[0] - a[0]) * dx + (p[1] - a[1]) * dy) / len)) : 0;
  return Math.hypot(p[0] - (a[0] + t * dx), p[1] - (a[1] + t * dy));
};

/** Strokes touched by an eraser at p (sheet pixels). */
export const strokesAt = (ink: InkStroke[], p: [number, number], radius: number): InkStroke[] =>
  ink.filter((k) => {
    const r = radius + k.width / 2;
    if (k.points.length === 1) return Math.hypot(p[0] - k.points[0]![0], p[1] - k.points[0]![1]) <= r;
    for (let i = 1; i < k.points.length; i++) if (segDist(p, k.points[i - 1]!, k.points[i]!) <= r) return true;
    return false;
  });

// ---------- Allow edit ranges ----------

export const addEditRange = (s: GridStore, title: string, range: CellRange, password?: string): void => {
  s.change('Allow edit range', 'sheet', () => {
    (s.sheet.editRanges ??= []).push({ title: title.trim() || `Range${(s.sheet.editRanges?.length ?? 0) + 1}`, range, ...(password ? { password: protectionHash(password) } : {}) });
  });
};

export const removeEditRange = (s: GridStore, index: number): void => {
  if (!s.sheet.editRanges?.[index]) return;
  s.change('Remove edit range', 'sheet', () => {
    s.sheet.editRanges!.splice(index, 1);
    if (!s.sheet.editRanges!.length) delete s.sheet.editRanges;
  });
};

/** Unlocks a password-protected edit range for this session. Returns false for a wrong password. */
export const unlockEditRange = (s: GridStore, index: number, password: string): boolean => {
  const e = s.sheet.editRanges?.[index];
  if (!e) return false;
  if (e.password && protectionHash(password) !== e.password) return false;
  s.unlockedRanges.add(editRangeKey(s.sheet.name, e.range));
  s.emit();
  return true;
};
