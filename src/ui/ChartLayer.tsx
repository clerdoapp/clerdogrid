import type { GridStore } from '../app/store';
import type { Viewport } from './render';
import { ChartView } from './charts/ChartView';
import { anchorBox } from './render';
import { overlays } from './panelRegistry';

/** Charts and pictures floating over the grid, positioned from their cell anchors. */
export const ChartLayer = ({ store, vp }: { store: GridStore; vp: Viewport }) => {
  const sheet = store.sheet;
  const extra = overlays();
  if (!sheet.charts.length && !sheet.pictures.length && !extra.length) return null;
  const place = (a: { row: number; col: number; dx: number; dy: number; width: number; height: number }) => anchorBox(vp, a);
  return (
    <div className="objlayer" style={{ left: vp.headerW, top: vp.headerH, right: 0, bottom: 0 }}>
      <div style={{ position: 'absolute', left: -vp.headerW, top: -vp.headerH }}>
        {sheet.kind !== 'chart' && sheet.charts.map((c) => <ChartView key={c.id} store={store} chart={c} box={place(c.anchor)} zoom={vp.zoom} />)}
        {sheet.pictures.map((p) => <PictureView key={p.id} store={store} id={p.id} box={place(p.anchor)} />)}
      </div>
      {extra.map((l) => <div key={l.id} className="overlay" style={{ position: 'absolute', left: -vp.headerW, top: -vp.headerH }}>{l.render(store, vp)}</div>)}
    </div>
  );
};

const urls = new WeakMap<Uint8Array, string>();
const PictureView = ({ store, id, box }: { store: GridStore; id: string; box: { left: number; top: number; width: number; height: number } }) => {
  const p = store.sheet.pictures.find((x) => x.id === id)!;
  let url = urls.get(p.data);
  if (!url) { url = URL.createObjectURL(new Blob([p.data as BlobPart], { type: p.mime })); urls.set(p.data, url); }
  return <img className="picture" src={url} alt="" draggable={false} style={box} />;
};
