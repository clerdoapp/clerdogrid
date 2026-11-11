import type { ReactNode } from 'react';
import { FileCode2, FileJson, FileSpreadsheet, Globe, RefreshCw, Trash2, type LucideIcon } from 'lucide-react';
import { command, run } from '../../app/commands';
import type { GridStore } from '../../app/store';
import { prettyKeys } from '../../ui/keys';
import { deleteQuery, refreshQuery, targetText } from './ops';
import { T, Tr } from '../../ui/i18n';

const S = 1.9;

const Tool = ({ store, id, icon: Icon, label }: { store: GridStore; id: string; icon?: LucideIcon; label?: string }) => {
  const c = command(id);
  if (!c) return null;
  const text = Tr(label ?? c.label);
  return (
    <button className="tool" disabled={c.enabled ? !c.enabled(store) : false} title={c.keys ? `${text} (${prettyKeys(c.keys)})` : text}
      onMouseDown={(e) => e.preventDefault()} onClick={() => run(store, id)}>
      {Icon && <Icon size={16} strokeWidth={S} />}<span>{Tr(text)}</span>
    </button>
  );
};
const List = ({ children }: { children: ReactNode }) => <div className="toollist">{children}</div>;

export const saveTypesSection = (s: GridStore) => (
  <List>
    <Tool store={s} id="file.saveAsXlsx" icon={FileSpreadsheet} />
    <Tool store={s} id="file.saveAsOds" icon={FileSpreadsheet} />
    <Tool store={s} id="file.saveAsOts" />
    <Tool store={s} id="file.saveAsXls" />
    <Tool store={s} id="file.saveAsXlsb" />
  </List>
);

export const exportDataSection = (s: GridStore) => (
  <List>
    <Tool store={s} id="file.exportHtml" icon={Globe} />
    <Tool store={s} id="file.exportXml" icon={FileCode2} />
    <Tool store={s} id="file.exportJson" icon={FileJson} />
  </List>
);

export const getDataSection = (s: GridStore) => (
  <>
    <List>
      <Tool store={s} id="data.fromWorkbook" icon={FileSpreadsheet} />
      <Tool store={s} id="data.refreshAll" icon={RefreshCw} />
    </List>
    {!!s.wb.queries?.length && (
      <div className="fmt-queries" role="list" aria-label={T('Queries')}>
        {s.wb.queries.map((q) => (
          <div key={q.id} className="fmt-query" role="listitem">
            <div className="fmt-qtext"><b>{q.name}</b><span className="muted small">{q.file} → {targetText(q)}{q.refreshed ? ` · ${new Date(q.refreshed).toLocaleString()}` : ''}</span></div>
            <button className="ico" title={T('Refresh {name}', { name: q.name })} aria-label={T('Refresh {name}', { name: q.name })} onClick={() => void refreshQuery(s, q)}><RefreshCw size={15} strokeWidth={S} /></button>
            <button className="ico" title={T('Remove the query {name} (the data stays)', { name: q.name })} aria-label={T('Remove {name}', { name: q.name })} onClick={() => deleteQuery(s, q)}><Trash2 size={15} strokeWidth={S} /></button>
          </div>
        ))}
      </div>
    )}
  </>
);
