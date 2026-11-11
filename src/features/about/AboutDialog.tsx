import { ExternalLink } from 'lucide-react';
import { Dialog, type DialogProps } from '../../ui/dialogs/Dialog';
import { aboutInfo, type AboutInfo } from './info';
import { Mark } from './Mark';
import { T, Tr } from '../../ui/i18n';

export const Links = ({ info }: { info: AboutInfo }) => (
  <div className="ab-links">
    {info.links.map((l) => <a key={l.url} href={l.url} target="_blank" rel="noopener noreferrer">{Tr(l.label)}<ExternalLink size={12} strokeWidth={1.9} aria-hidden="true" /></a>)}
  </div>
);

export const AboutDialog = ({ close }: DialogProps) => {
  const info = aboutInfo();
  return (
    <Dialog title={T('About {name}', { name: info.name })} close={close} width={460}>
      <div className="ab-head">
        <Mark />
        <div>
          <h3>{info.name}</h3>
          <div className="muted">{T('Version')}{' '}{info.version}</div>
        </div>
      </div>
      <div className="ab-sec">
        <p>{T('Free and open source under {license}, your choice.', { license: info.license.replace(' OR ', ` ${T('or')} `) })}</p>
        <p>{T('No account, no tracking: your files stay on your device.')}</p>
      </div>
      {info.links.length > 0 && <div className="ab-sec"><h4>{T('Links')}</h4><Links info={info} /></div>}
      <div className="ab-sec">
        <h4>{T('Contributors')}</h4>
        <ul className="ab-people">
          {info.contributors.map((c) => (
            <li key={c.name}>
              {c.url ? <a href={c.url} target="_blank" rel="noopener noreferrer">{c.name}</a> : <span>{c.name}</span>}
              {c.role && <span className="muted">{c.role}</span>}
            </li>
          ))}
        </ul>
      </div>
      {info.notices.length > 0 && (
        <div className="ab-sec">
          <h4>{T('Third-party notices')}</h4>
          <ul className="ab-notices">{info.notices.map((n) => <li key={n}>{n}</li>)}</ul>
        </div>
      )}
    </Dialog>
  );
};
