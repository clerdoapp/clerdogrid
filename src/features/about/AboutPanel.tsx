import { Info } from 'lucide-react';
import { openDialog } from '../../app/commands';
import type { GridStore } from '../../app/store';
import { Links } from './AboutDialog';
import { aboutInfo } from './info';
import { T } from '../../ui/i18n';

export const AboutPanel = ({ store }: { store: GridStore }) => {
  const info = aboutInfo();
  return (
    <div className="ab-panel">
      <div className="ab-line">{T('Version {version}', { version: info.version })} <span>· {info.license.replace(' OR ', ` ${T('or')} `)}</span></div>
      <div className="ab-line"><span>{T('Made by')}</span> {info.contributors.map((c) => c.name).join(', ')}</div>
      <Links info={info} />
      <div className="toollist">
        <button className="tool" onClick={() => openDialog(store, 'about.full')}><Info size={16} strokeWidth={1.9} /><span>{T('Contributors and notices')}</span></button>
      </div>
    </div>
  );
};
