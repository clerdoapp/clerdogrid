import { useEffect, useRef, type ReactNode } from 'react';
import { X } from 'lucide-react';

export interface DialogProps<T = unknown> {
  store: import('../../app/store').GridStore;
  data?: T;
  close: () => void;
}

/** Shared dialog frame: title bar, scrollable body, action row; Escape closes, Enter submits. */
export const Dialog = ({ title, children, onOk, okLabel = 'OK', close, width = 420, footer, wide }: {
  title: string; children: ReactNode; onOk?: () => void | boolean; okLabel?: string; close: () => void; width?: number; footer?: ReactNode; wide?: boolean;
}) => {
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const first = ref.current?.querySelector<HTMLElement>('[autofocus], input, select, textarea, button.primary');
    first?.focus();
  }, []);
  const ok = () => { if (onOk) { if (onOk() !== false) close(); } else close(); };
  return (
    <div className="backdrop" onPointerDown={(e) => { if (e.target === e.currentTarget) close(); }}>
      <div className="dialog" ref={ref} role="dialog" aria-modal="true" aria-label={title} style={{ width: wide ? Math.min(760, window.innerWidth - 24) : Math.min(width, window.innerWidth - 24) }}
        onKeyDown={(e) => {
          if (e.key === 'Escape') { e.stopPropagation(); close(); }
          if (e.key === 'Enter' && !(e.target instanceof HTMLTextAreaElement) && !(e.target instanceof HTMLButtonElement) && onOk) { e.preventDefault(); ok(); }
        }}>
        <div className="dhead"><h2>{title}</h2><button className="ico" aria-label="Close" onClick={close}><X size={16} strokeWidth={1.9} /></button></div>
        <div className="dbody">{children}</div>
        <div className="dfoot">{footer}<div className="grow" />{onOk && <button onClick={close}>Cancel</button>}<button className="primary" onClick={ok}>{onOk ? okLabel : 'Close'}</button></div>
      </div>
    </div>
  );
};

export const Field = ({ label, children }: { label: string; children: ReactNode }) => <label className="field"><span>{label}</span>{children}</label>;
