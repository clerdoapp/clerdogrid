import { useMemo, useState } from 'react';
import { BarChart3, LineChart, PieChart, ScatterChart, Send, Table2, TrendingUp, TriangleAlert, Link2, Crown } from 'lucide-react';
import { Dialog, type DialogProps } from '../../ui/dialogs/Dialog';
import { rangeAddr } from '../../model/workbook';
import { formatValue } from '../../formats/numfmt';
import { analyzeBlock, insertSourceChart, insertSummary, readBlock, runAsk, type DataBlock } from './insightActions';
import { parseAsk, type ColumnSummary, type Insights, type Visual } from './insights';
import { N, T, Tr } from '../../ui/i18n';

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
      <div className="in-colhead"><b title={c.name}>{c.name}</b><span className={`in-type t-${c.type}`}>{Tr(TYPE_LABEL[c.type])}</span></div>
      <dl>
        <div><dt>{T('Filled')}</dt><dd>{c.count}</dd></div>
        <div><dt>{T('Blank')}</dt><dd>{c.blanks ? `${c.blanks} (${pct(c.blanks, rows)})` : 0}</dd></div>
        <div><dt>{T('Distinct')}</dt><dd>{c.distinct}</dd></div>
        {(c.type === 'number' || date) && <>
          <div><dt>{T('Min')}</dt><dd>{num(c.min, date)}</dd></div>
          <div><dt>{T('Max')}</dt><dd>{num(c.max, date)}</dd></div>
          {!date && <div><dt>{T('Mean')}</dt><dd>{num(c.mean)}</dd></div>}
          <div><dt>{T('Median')}</dt><dd>{num(c.median, date)}</dd></div>
        </>}
      </dl>
      {c.top && <ul className="in-top">{c.top.slice(0, 3).map((t) => <li key={t.value}><span title={t.value}>{t.value}</span><i>{t.count}</i></li>)}</ul>}
    </div>
  );
};

const TYPE_LABEL: Record<ColumnSummary['type'], string> = { number: N('Numbers'), date: N('Dates'), text: N('Text'), boolean: N('TRUE/FALSE'), empty: N('Empty') };

const Card = ({ icon: Icon, title, children, onInsert, insertLabel = N('Insert') }: { icon: typeof Crown; title: string; children: React.ReactNode; onInsert?: () => void; insertLabel?: string }) => (
  <div className="in-card">
    <div className="in-cardhead"><Icon size={16} strokeWidth={S} /><b>{Tr(title)}</b><div className="grow" />{onInsert && <button className="tool in-insert" onClick={onInsert}>{Tr(insertLabel)}</button>}</div>
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
    return <Dialog title={T('Analyze data')} close={close}><p>{T('Select a block of data with a header row, then try again.')}</p></Dialog>;
  }
  const h = block.headers;

  const insertVisual = (v: Visual) => {
    if (v.kind === 'scatter') { insertSourceChart(store, block, 'scatter', v.val, [v.val2!], v.title); done(T('Added a chart: {title}.', { title: v.title })); return; }
    if (v.kind === 'line') { insertSourceChart(store, block, 'line', v.cat!, [v.val], v.title); done(T('Added a chart: {title}.', { title: v.title })); return; }
    insertSummary(store, block, { by: v.cat!, val: v.val, agg: v.agg ?? 'sum', chart: v.kind, title: v.title });
    done(T('Added a summary sheet: {title}.', { title: v.title }));
  };

  const ask = () => {
    const a = parseAsk(q, h, ins.columns);
    if (!a) { setAnswer(T('Try "total {value} by {category}", "top 5 … by …", "count by …" or "… over time".', { value: h.find((_, i) => ins.columns[i]?.type === 'number') ?? T('Sales'), category: h.find((_, i) => ins.columns[i]?.type === 'text') ?? T('Region') })); return; }
    const msg = runAsk(store, block, a);
    if (a.kind === 'total') setAnswer(msg); else done(msg);
  };

  const outCols = [...new Set(ins.outliers.map((o) => o.col))];
  const firstRow = block.range.r1 + (block.hasHeader ? 1 : 0);

  return (
    <Dialog title={T('Analyze data')} close={close} wide>
      <p className="muted in-sub">{block.sheetName}!{rangeAddr(block.range)} · {T('{rows, plural, one {# row} other {# rows}}', { rows: ins.rows })} · {T('{count, plural, one {# column} other {# columns}}', { count: h.length })}{block.hasHeader ? '' : ` · ${T('no header row found')}`}</p>

      <form className="in-ask" onSubmit={(e) => { e.preventDefault(); ask(); }}>
        <input value={q} onChange={(e) => { setQ(e.target.value); setAnswer(null); }} placeholder={T('Ask about your data, for example "total sales by region"')} aria-label={T('Ask about your data')} />
        <button className="primary" type="submit" disabled={!q.trim()}><Send size={15} strokeWidth={S} />{T('Ask')}</button>
      </form>
      {answer && <p className="in-answer" role="status">{answer}</p>}

      <div className="in-grid">
        {ins.trend && (
          <Card icon={TrendingUp} title={T('{yName} over {xName}', { yName: ins.trend.yName, xName: ins.trend.xName })} onInsert={() => { insertSourceChart(store, block, 'line', ins.trend!.x, [ins.trend!.y], T('{yName} over {xName}', { yName: ins.trend!.yName, xName: ins.trend!.xName })); done(T('Added a line chart.')); }}>
            <p>{T(ins.trend.slope >= 0
              ? ins.columns[ins.trend.x]?.type === 'date' ? 'Rises by about {amount} per day.' : 'Rises by about {amount} per step.'
              : ins.columns[ins.trend.x]?.type === 'date' ? 'Falls by about {amount} per day.' : 'Falls by about {amount} per step.', { amount: num(Math.abs(ins.trend.slope)) })}{' '}
              {T(ins.trend.r2 >= 0.7 ? 'The straight line fits well (R² {r2}).' : ins.trend.r2 >= 0.3 ? 'The straight line fits loosely (R² {r2}).' : 'The straight line fits poorly (R² {r2}).', { r2: ins.trend.r2.toFixed(2) })}</p>
          </Card>
        )}
        {ins.topCategories.map((t) => (
          <Card key={`${t.cat}-${t.val}`} icon={Crown} title={T('Top {catName} by {valName}', { catName: t.catName, valName: t.valName })} insertLabel={T('Insert table')}
            onInsert={() => { insertSummary(store, block, { by: t.cat, val: t.val, agg: 'sum', chart: 'column', title: T('{valName} by {catName}', { valName: t.valName, catName: t.catName }) }); done(T('Added a summary sheet.')); }}>
            <Bars items={t.items} />
          </Card>
        ))}
        {ins.correlations.slice(0, 3).map((c) => (
          <Card key={`${c.a}-${c.b}`} icon={Link2} title={T('{aName} and {bName}', { aName: c.aName, bName: c.bName })} onInsert={() => { insertSourceChart(store, block, 'scatter', c.a, [c.b], T('{bName} against {aName}', { bName: c.bName, aName: c.aName })); done(T('Added a scatter chart.')); }}>
            <p>{T(c.r > 0 ? 'They move together: r = {r} over {count, plural, one {# row} other {# rows}}.' : 'They move in opposite directions: r = {r} over {count, plural, one {# row} other {# rows}}.', { r: c.r.toFixed(2), count: c.n })}</p>
          </Card>
        ))}
        {outCols.map((col) => {
          const list = ins.outliers.filter((o) => o.col === col);
          return (
            <Card key={`o${col}`} icon={TriangleAlert} title={T('Unusual {p0}', { p0: h[col] })} insertLabel={T('Select')}
              onInsert={() => { const o = list[0]!; store.selectRange({ r1: firstRow + o.row, c1: block.range.c1 + col, r2: firstRow + o.row, c2: block.range.c1 + col }); close(); }}>
              <p>{T('{count, plural, one {# value is} other {# values are}} outside {low} to {high}: {values}', { count: list.length, low: num(list[0]!.low), high: num(list[0]!.high), values: list.slice(0, 4).map((o) => num(o.value)).join(', ') + (list.length > 4 ? '…' : '') })}</p>
            </Card>
          );
        })}
        {ins.visuals.length > 0 && (
          <Card icon={Table2} title={T('Suggested charts')}>
            <div className="in-vis">{ins.visuals.map((v) => { const I = VIS_ICON[v.kind]; return <button key={v.title + v.kind} className="tool" onClick={() => insertVisual(v)}><I size={16} strokeWidth={S} /><span>{Tr(v.title)}</span></button>; })}</div>
          </Card>
        )}
        {!ins.trend && !ins.topCategories.length && !ins.correlations.length && !ins.outliers.length && <p className="muted">{T('No patterns stood out. The column summaries are below.')}</p>}
      </div>

      <h3 className="b-h3">{T('Columns')}</h3>
      <div className="in-cols">{ins.columns.map((c) => <ColumnCard key={c.index} c={c} rows={ins.rows} />)}</div>
    </Dialog>
  );
};
