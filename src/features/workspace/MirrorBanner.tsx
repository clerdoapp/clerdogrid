import { Columns2, Unlink } from 'lucide-react';
import type { GridStore } from '../../app/store';
import type { Viewport } from '../../ui/render';
import { detachMirror, mirrorState } from './mirror';
import { T } from '../../ui/i18n';

/** Shown in a second window: both windows edit the same workbook. */
export const MirrorBanner = ({ store, vp }: { store: GridStore; vp: Viewport }) => {
  if (!mirrorState.following) return null;
  return (
    <div className="ws-banner" role="region" aria-label={T('Second window')} style={{ left: vp.headerW + 8, width: Math.max(200, vp.width - vp.headerW - 16), top: Math.max(vp.headerH + 8, vp.height - 54) }} onPointerDown={(e) => e.stopPropagation()}>
      <Columns2 size={16} strokeWidth={1.9} />
      <span>{T('Second window on this workbook · changes in either window show in both')}</span>
      <button className="sm" onClick={() => detachMirror(store)}><Unlink size={14} strokeWidth={1.9} />{' '}{T('Make separate copy')}</button>
    </div>
  );
};
