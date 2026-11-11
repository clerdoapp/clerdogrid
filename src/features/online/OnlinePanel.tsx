import { useState } from 'react';
import type { GridStore } from '../../app/store';
import { Check } from '../../ui/dialogs/shared';
import { onlineAllowed } from '../../engine/online';
import { needsReload, setAllowed } from './state';
import { T } from '../../ui/i18n';

/** File > Online functions: the opt-in for WEBSERVICE, IMAGE and STOCKHISTORY. */
export const OnlinePanel = ({ store }: { store: GridStore }) => {
  const [, force] = useState(0);
  return (
    <div className="region-panel">
      <Check label={T('Allow functions to use the internet')} checked={onlineAllowed()} onChange={(v) => { setAllowed(store, v); force((n) => n + 1); }} />
      <p className="mini">{T('WEBSERVICE, IMAGE, STOCKHISTORY and IMPORTCSV fetch data over https. While this is off they show #CONNECT! and nothing is fetched.')}</p>
      {needsReload() && onlineAllowed() && <button className="tool" onClick={() => location.reload()}><span>{T('Reload to finish turning this on')}</span></button>}
    </div>
  );
};
