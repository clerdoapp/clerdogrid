import { useMemo, useState } from 'react';
import { BarChart3, LineChart, PieChart, ScatterChart, Send, Table2, TrendingUp, TriangleAlert, Link2, Crown } from 'lucide-react';
import { Dialog, type DialogProps } from '../../ui/dialogs/Dialog';
import { rangeAddr } from '../../model/workbook';
import { formatValue } from '../../formats/numfmt';
import { analyzeBlock, insertSourceChart, insertSummary, readBlock, runAsk, type DataBlock } from './insightActions';
import { parseAsk, type ColumnSummary, type Insights, type Visual } from './insights';

const S = 1.9;
const num = (v: number | undefined, date = false): string => {
  if (v === undefined || !Number.isFinite(v)) return '–';
  if (date) return formatValue(v, 'yyyy-mm-dd').text;
  return Math.abs(v) >= 1e6 || (Math.abs(v) < 0.01 && v !== 0) ? v.toExponential(2) : (+v.toFixed(2)).toLocaleString();
};
const pct = (n: number, d: number) => (d ? `${Math.round((n / d) * 100)}%` : '0%');

const VIS_ICON = { line: LineChart, column: BarChart3, bar: BarChart3, pie: PieChart, scatter: ScatterChart } as const;

const ColumnCard = ({ c, rows }: { c: ColumnSummary; rows: number }) => {
  const date = c.type === 'date';
  return (
    <div className="in-col">
      <div className="in-colhead"><b title={c.name}>{c.name}</b><span className={`in-type t-${c.type}`}>{c.type}</span></div>
      <dl>
        <div><dt>Filled</dt><dd>{c.count}</dd></div>
        <div><dt>Blank</dt><dd>{c.blanks ? `${c.blanks} (${pct(c.blanks, rows)})` : 0}</dd></div>
        <div><dt>Distinct</dt><dd>{c.distinct}</dd></div>
        {(c.type === 'number' || date) && <>
          <div><dt>Min</dt><dd>{num(c.min, date)}</dd></div>
          <div><dt>Max</dt><dd>{num(c.max, date)}</dd></div>
          {!date && <div><dt>Mean</dt><dd>{num(c.mean)}</dd></div>}
          <div><dt>Median</dt><dd>{num(c.median, date)}</dd></div>
        </>}
      </dl>
      {c.top && <ul className="in-top">{c.top.slice(0, 3).map((t) => <li key={t.value}><span title={t.value}>{t.value}</span><i>{t.count}</i></li>)}</ul>}
    </div>
  );
};

const Card = ({ icon: Icon, title, children, onInsert, insertLabel = 'Insert' }: { icon: typeof Crown; title: string; children: React.ReactNode; onInsert?: () => void; insertLabel?: string }) => (
  <div className="in-card">
    <div className="in-cardhead"><Icon size={16} strokeWidth={S} /><b>{title}</b><div className="grow" />{onInsert && <button className="tool in-insert" onClick={onInsert}>{insertLabel}</button>}</div>
    <div className="in-cardbody">{children}</div>
  </div>
);

const Bars = ({ items }: { items: { key: string; total: number }[] }) => {
  const max = Math.max(...items.map((i) => Math.abs(i.total)), 1);
  return <div className="in-bars">{items.map((i) => (
    <div key={i.key} className="in-bar"><span title={i.key}>{i.key}</span><div><i style={{ width: `${(Math.abs(i.total) / max) * 100}%` }} /></div><em>{num(i.total)}</em></div>
  ))}</div>;
};

export const InsightsDialog = ({ store, close }: DialogProps) => {
  const block = useMemo<DataBlock | null>(() => readBlock(store), [store]);
  const ins = useMemo<Insights | null>(() => (block ? analyzeBlock(store, block) : null), [store, block]);
  const [q, setQ] = useState('');
  const [answer, setAnswer] = useState<string | null>(null);
  const done = (msg: string) => { store.notify(msg); close(); };

  if (!block || !ins) {
    return <Dialog title="Analyze data" close={close}><p>Select a block of data with a header row, then try again.</p></Dialog>;
  }
  const h = block.headers;

  const insertVisual = (v: Visual) => {
    if (v.kind === 'scatter') { insertSourceChart(store, block, 'scatter', v.val, [v.val2!], v.title); done(`Added a chart: ${v.title}.`); return; }
    if (v.kind === 'line') { insertSourceChart(store, block, 'line', v.cat!, [v.val], v.title); done(`Added a chart: ${v.title}.`); return; }
    insertSummary(store, block, { by: v.cat!, val: v.val, agg: v.agg ?? 'sum', chart: v.kind, title: v.title });
    done(`Added a summary sheet: ${v.title}.`);
  };

  const ask = () => {
    const a = parseAsk(q, h, ins.columns);
    if (!a) { setAnswer(`Try "total ${h.find((_, i) => ins.columns[i]?.type === 'number') ?? 'Sales'} by ${h.find((_, i) => ins.columns[i]?.type === 'text') ?? 'Region'}", "top 5 … by …", "count by …" or "… over time".`); return; }
    const msg = runAsk(store, block, a);
    if (a.kind === 'total') setAnswer(msg); else done(msg);
  };

  const outCols = [...new Set(ins.outliers.map((o) => o.col))];
  const firstRow = block.range.r1 + (block.hasHeader ? 1 : 0);

  return (
    <Dialog title="Analyze data" close={close} wide>
      <p className="muted in-sub">{block.sheetName}!{rangeAddr(block.range)} · {ins.rows} rows · {h.length} columns{block.hasHeader ? '' : ' · no header row found'}</p>

      <form className="in-ask" onSubmit={(e) => { e.preventDefault(); ask(); }}>
        <input value={q} onChange={(e) => { setQ(e.target.value); setAnswer(null); }} placeholder={'Ask about your data, for example "total sales by region"'} aria-label="Ask about your data" />
        <button className="primary" type="submit" disabled={!q.trim()}><Send size={15} strokeWidth={S} />Ask</button>
      </form>
      {answer && <p className="in-answer" role="status">{answer}</p>}

      <div className="in-grid">
        {ins.trend && (
          <Card icon={TrendingUp} title={`${ins.trend.yName} over ${ins.trend.xName}`} onInsert={() => { insertSourceChart(store, block, 'line', ins.trend!.x, [ins.trend!.y], `${ins.trend!.yName} over ${ins.trend!.xName}`); done('Added a line chart.'); }}>
            <p>{ins.trend.slope >= 0 ? 'Rises' : 'Falls'} by about <b>{num(Math.abs(ins.trend.slope))}</b> per {ins.columns[ins.trend.x]?.type === 'date' ? 'day' : 'step'}. The straight line fits {ins.trend.r2 >= 0.7 ? 'well' : ins.trend.r2 >= 0.3 ? 'loosely' : 'poorly'} (R² {ins.trend.r2.toFixed(2)}).</p>
          </Card>
        )}
        {ins.topCategories.map((t) => (
          <Card key={`${t.cat}-${t.val}`} icon={Crown} title={`Top ${t.catName} by ${t.valName}`} insertLabel="Insert table"
            onInsert={() => { insertSummary(store, block, { by: t.cat, val: t.val, agg: 'sum', chart: 'column', title: `${t.valName} by ${t.catName}` }); done('Added a summary sheet.'); }}>
            <Bars items={t.items} />
          </Card>
        ))}
        {ins.correlations.slice(0, 3).map((c) => (
          <Card key={`${c.a}-${c.b}`} icon={Link2} title={`${c.aName} and ${c.bName}`} onInsert={() => { insertSourceChart(store, block, 'scatter', c.a, [c.b], `${c.bName} against ${c.aName}`); done('Added a scatter chart.'); }}>
            <p>They move {c.r > 0 ? 'together' : 'in opposite directions'}: r = <b>{c.r.toFixed(2)}</b> over {c.n} rows.</p>
          </Card>
        ))}
        {outCols.map((col) => {
          const list = ins.outliers.filter((o) => o.col === col);
          return (
            <Card key={`o${col}`} icon={TriangleAlert} title={`Unusual ${h[col]}`} insertLabel="Select"
              onInsert={() => { const o = list[0]!; store.selectRange({ r1: firstRow + o.row, c1: block.range.c1 + col, r2: firstRow + o.row, c2: block.range.c1 + col }); close(); }}>
              <p>{list.length} {list.length === 1 ? 'value is' : 'values are'} outside {num(list[0]!.low)} to {num(list[0]!.high)}: {list.slice(0, 4).map((o) => num(o.value)).join(', ')}{list.length > 4 ? '…' : ''}</p>
            </Card>
          );
        })}
        {ins.visuals.length > 0 && (
          <Card icon={Table2} title="Suggested charts">
            <div className="in-vis">{ins.visuals.map((v) => { const I = VIS_ICON[v.kind]; return <button key={v.title + v.kind} className="tool" onClick={() => insertVisual(v)}><I size={16} strokeWidth={S} /><span>{v.title}</span></button>; })}</div>
          </Card>
        )}
        {!ins.trend && !ins.topCategories.length && !ins.correlations.length && !ins.outliers.length && <p className="muted">No patterns stood out. The column summaries are below.</p>}
      </div>

      <h3 className="b-h3">Columns</h3>
      <div className="in-cols">{ins.columns.map((c) => <ColumnCard key={c.index} c={c} rows={ins.rows} />)}</div>
    </Dialog>
  );
};
