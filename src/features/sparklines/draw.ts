// Draws sparklines into their cells on the grid canvas.
import type { GridStore } from '../../app/store';
import type { SparklineGroup } from '../../model/workbook';
import type { PanePainter } from '../../ui/render';
import { readNumbers } from '../../ui/charts/chartData';
import { groupScales, sparkGeometry, sparkValues } from './model';

// Panes of one repaint share the values they read (values can change between repaints).
const cache = new WeakMap<SparklineGroup, { at: number; sheet: number; vals: (number | null)[][]; dates?: (number | null)[] }>();

const groupData = (store: GridStore, g: SparklineGroup) => {
  const hit = cache.get(g);
  const now = performance.now();
  if (hit && now - hit.at < 8 && hit.sheet === store.sheetIndex) return hit;
  const vals = g.sparklines.map((sp) => sparkValues(store, store.sheetIndex, g, sp));
  const dates = g.dateRange ? readNumbers(store, g.dateRange) : undefined;
  const next = { at: now, sheet: store.sheetIndex, vals, dates };
  cache.set(g, next);
  return next;
};

export const paintSparklines: PanePainter = (ctx, store, a) => {
  const groups = store.sheet.sparklines;
  if (!groups?.length) return;
  const { pane: p, vp } = a;
  const { layout, zoom } = vp;
  const toDisplay = (r: number) => (vp.rowOrder ? vp.rowOrder.toDisplay(r) : r);
  for (const g of groups) {
    const shown = g.sparklines.map((sp, i) => ({ sp, i, dr: toDisplay(sp.row) })).filter(({ sp, dr }) => dr >= p.r1 && dr <= p.r2 && sp.col >= p.c1 && sp.col <= p.c2);
    if (!shown.length) continue;
    const data = groupData(store, g);
    const scales = groupScales(g, data.vals);
    for (const { sp, i, dr } of shown) {
      const m = a.merges.find((mm) => mm.r1 === sp.row && mm.c1 === sp.col);
      const x = a.colX(sp.col), y = a.rowY(dr);
      const w = (m ? layout.cols.start(m.c2 + 1) - layout.cols.start(m.c1) : layout.cols.size(sp.col)) * zoom;
      const h = (m ? layout.rows.start(m.r2 + 1) - layout.rows.start(m.r1) : layout.rows.size(dr)) * zoom;
      if (w < 3 || h < 3) continue;
      const geo = sparkGeometry(g, data.vals[i] ?? [], { x, y, w, h }, scales[i]!, data.dates);
      ctx.save();
      ctx.beginPath(); ctx.rect(x, y, w, h); ctx.clip();
      if (geo.axis !== null) {
        ctx.strokeStyle = g.axisColor ?? (a.dark ? '#a9b8af' : '#000000');
        ctx.lineWidth = Math.max(1, zoom * 0.75);
        ctx.beginPath(); ctx.moveTo(x + 1, Math.round(geo.axis) + 0.5); ctx.lineTo(x + w - 1, Math.round(geo.axis) + 0.5); ctx.stroke();
      }
      for (const b of geo.bars) {
        ctx.fillStyle = b.color;
        ctx.fillRect(b.x, b.y0, b.w, Math.max(1, b.y1 - b.y0));
      }
      if (geo.lines.length) {
        ctx.strokeStyle = g.color ?? '#12A565';
        ctx.lineWidth = Math.max(0.75, ((g.lineWeight ?? 0.75) * 96) / 72) * zoom;
        ctx.lineJoin = 'round'; ctx.lineCap = 'round';
        for (const run of geo.lines) {
          ctx.beginPath();
          run.forEach(([px, py], k) => (k ? ctx.lineTo(px, py) : ctx.moveTo(px, py)));
          if (run.length === 1) ctx.lineTo(run[0]![0] + 0.01, run[0]![1]);
          ctx.stroke();
        }
      }
      const r = Math.max(1.5, Math.min(3, h / 8)) * Math.max(1, zoom * 0.8);
      for (const mk of geo.marks) {
        ctx.fillStyle = mk.color;
        ctx.beginPath(); ctx.arc(mk.x, mk.y, r, 0, Math.PI * 2); ctx.fill();
      }
      ctx.restore();
    }
  }
};
