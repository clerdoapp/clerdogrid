import { useState } from 'react';
import { Code2, FileText, LayoutGrid, Play, Sparkles, Send } from 'lucide-react';
import { openDialog } from '../../app/commands';
import type { GridStore } from '../../app/store';
import { bookScripts } from './library';
import { runSavedScript } from './ScriptEditor';
import { analyzeBlock, readBlock, runAsk } from './insightActions';
import { parseAsk } from './insights';
import { T } from '../../ui/i18n';

const S = 1.9;

export const AutomatePanel = ({ store }: { store: GridStore }) => {
  const list = bookScripts(store);
  return (<>
    <div className="toollist">
      <button className="tool" onClick={() => openDialog(store, 'scripts')}><Code2 size={16} strokeWidth={S} /><span>{T('Script editor')}</span></button>
      <button className="tool" onClick={() => openDialog(store, 'scripts', { tab: 'sample' })}><LayoutGrid size={16} strokeWidth={S} /><span>{T('Sample scripts')}</span></button>
    </div>
    <div className="au-plist" aria-label={T('Scripts in this workbook')}>
      {list.map((s) => (
        <div key={s.id} className="au-prow">
          <button className="tool" title={T('Edit {name}', { name: s.name })} onClick={() => openDialog(store, 'scripts', { open: { where: 'book', id: s.id } })}><FileText size={16} strokeWidth={S} /><span>{s.name}</span></button>
          <button className="tool icon" aria-label={T('Run {name}', { name: s.name })} title={T('Run')} onClick={() => void runSavedScript(store, s.id)}><Play size={15} strokeWidth={S} /></button>
        </div>
      ))}
      {!list.length && <p className="mini">{T('No scripts in this workbook.')}</p>}
    </div>
  </>);
};

export const InsightsPanel = ({ store }: { store: GridStore }) => {
  const [q, setQ] = useState('');
  const ask = () => {
    const block = readBlock(store);
    if (!block) { store.notify(T('Select a block of data with a header row first.')); return; }
    const a = parseAsk(q, block.headers, analyzeBlock(store, block).columns);
    if (!a) { openDialog(store, 'insights'); return; }
    store.notify(runAsk(store, block, a));
    setQ('');
  };
  return (<>
    <div className="toollist"><button className="tool" onClick={() => openDialog(store, 'insights')}><Sparkles size={16} strokeWidth={S} /><span>{T('Analyze data')}</span></button></div>
    <form className="au-pask" onSubmit={(e) => { e.preventDefault(); ask(); }}>
      <input value={q} onChange={(e) => setQ(e.target.value)} placeholder={T('total sales by region')} aria-label={T('Ask about your data')} />
      <button className="tool icon" type="submit" aria-label={T('Ask')} disabled={!q.trim()}><Send size={15} strokeWidth={S} /></button>
    </form>
  </>);
};
