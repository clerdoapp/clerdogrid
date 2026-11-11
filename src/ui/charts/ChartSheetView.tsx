// A chart sheet: its one chart fills the grid area.
import { Paintbrush, Pencil } from 'lucide-react';
import { openDialog } from '../../app/commands';
import type { GridStore } from '../../app/store';
import type { Viewport } from '../render';
import { resolveChart } from './chartData';
import { ChartSvg } from './ChartSvg';
import { openFormat } from './ChartView';
import { T } from '../i18n';

const PAD = 24;

export const ChartSheetView = ({ store, vp }: { store: GridStore; vp: Viewport }) => {
  const sheet = store.sheet;
  if (sheet.kind !== 'chart') return null;
  const chart = sheet.charts[0];
  // The chart keeps its sheet's proportions, scaled to the window.
  const W = vp.width - vp.headerW, H = vp.height - vp.headerH;
  const aw = Math.max(80, W - PAD * 2), ah = Math.max(60, H - PAD * 2);
  const ratio = chart ? chart.anchor.width / Math.max(1, chart.anchor.height) : 1.4;
  const w = Math.min(aw, ah * ratio), h = w / ratio;
  return (
    <div className="chartsheet" style={{ position: 'absolute', left: vp.headerW, top: vp.headerH, width: W, height: H }}
      onPointerDown={(e) => e.stopPropagation()} onDoubleClick={() => chart && openDialog(store, 'chart', { id: chart.id })}>
      {chart ? (
        <div className="chartsheetbox" data-chart-id={chart.id} style={{ left: (W - w) / 2, top: (H - h) / 2, width: w, height: h }} role="figure" aria-label={chart.title || T('Chart')}>
          <ChartSvg chart={chart} data={resolveChart(store, chart)} width={w} height={h} />
          <div className="chartedit">
            <button className="ico sm" aria-label={T('Format chart')} title={T('Format chart')} onClick={() => openFormat(store)}><Paintbrush size={15} strokeWidth={1.9} /></button>
            <button className="ico sm" aria-label={T('Edit chart')} title={T('Edit chart')} onClick={() => openDialog(store, 'chart', { id: chart.id })}><Pencil size={15} strokeWidth={1.9} /></button>
          </div>
        </div>
      ) : <p className="muted pad">{T('This chart sheet has no chart.')}</p>}
    </div>
  );
};
