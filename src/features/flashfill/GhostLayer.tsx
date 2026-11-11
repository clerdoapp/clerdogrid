import type { GridStore } from '../../app/store';
import { cssFont, paneFor, panesOf, rectIn, type Viewport } from '../../ui/render';
import { ghost } from './ghost';
import { T } from '../../ui/i18n';

/** Grey preview of the suggested Flash Fill values. */
export const GhostLayer = ({ store, vp }: { store: GridStore; vp: Viewport }) => {
  const plan = ghost.plan;
  if (!plan || plan.sheet !== store.sheetIndex) return null;
  const panes = panesOf(vp);
  const items: React.ReactNode[] = [];
  let first: { left: number; top: number; height: number } | null = null;
  for (const [row, v] of plan.fills) {
    const r = rectIn(paneFor(panes, vp, row, plan.col), vp, { r1: row, c1: plan.col, r2: row, c2: plan.col });
    if (r.y + r.h < vp.headerH || r.y > vp.height) continue;
    first ??= { left: r.x + r.w + 6, top: r.y, height: r.h };
    const style = store.style(row, plan.col);
    items.push(
      <div key={row} className={`ff-ghost${typeof v === 'number' ? ' num' : ''}`} style={{ left: r.x + 2, top: r.y + 2, width: r.w - 4, height: r.h - 3, font: cssFont(style, vp.zoom) }}>{String(v)}</div>,
    );
    if (items.length > 400) break;
  }
  return (<>
    {items}
    {first && <div className="ff-tip" style={{ left: first.left, top: first.top + (first.height - 24) / 2 }}>{T('Flash Fill: Enter to accept, Esc to dismiss')}</div>}
  </>);
};
