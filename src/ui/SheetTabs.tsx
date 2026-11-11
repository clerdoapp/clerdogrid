import { useState } from 'react';
import { ChevronLeft, ChevronRight, Menu, Plus } from 'lucide-react';
import { addSheet, moveSheet, renameSheet } from '../app/ops/structure';
import type { GridStore } from '../app/store';
import { T } from './i18n';

const S = 1.9;

export const SheetTabs = ({ store }: { store: GridStore }) => {
  const [renaming, setRenaming] = useState<{ i: number; text: string } | null>(null);
  const [drag, setDrag] = useState<number | null>(null);
  const finish = () => { if (renaming) { renameSheet(store, renaming.i, renaming.text); setRenaming(null); } };
  return (
    <div className="sheettabs" role="tablist" aria-label={T('Sheets')}>
      <button className="ico sm" aria-label={T('Previous sheet')} onClick={() => document.querySelector('.tabs')?.scrollBy({ left: -160, behavior: 'smooth' })}><ChevronLeft size={15} strokeWidth={S} /></button>
      <button className="ico sm" aria-label={T('Next sheet')} onClick={() => document.querySelector('.tabs')?.scrollBy({ left: 160, behavior: 'smooth' })}><ChevronRight size={15} strokeWidth={S} /></button>
      <button className="ico sm" aria-label={T('All sheets')} title={T('All sheets')} onClick={(e) => { store.menu = { x: e.clientX, y: e.clientY - 8, kind: 'sheet', data: { list: true } }; store.emit(); }}><Menu size={15} strokeWidth={S} /></button>
      <div className="tabs">
        {store.wb.sheets.map((s, i) => s.state !== 'visible' ? null : (
          <div key={i} role="tab" aria-selected={i === store.sheetIndex} tabIndex={0} draggable={!renaming}
            className={`tab${i === store.sheetIndex ? ' on' : ''}${drag !== null && drag !== i ? ' droptarget' : ''}`}
            style={s.tabColor ? { boxShadow: `inset 0 -3px 0 ${s.tabColor}` } : undefined}
            onClick={() => store.setActiveSheet(i)}
            onDoubleClick={() => setRenaming({ i, text: s.name })}
            onContextMenu={(e) => { e.preventDefault(); store.setActiveSheet(i); store.menu = { x: e.clientX, y: e.clientY, kind: 'sheet', data: { index: i } }; store.emit(); }}
            onDragStart={() => setDrag(i)} onDragEnd={() => setDrag(null)}
            onDragOver={(e) => e.preventDefault()} onDrop={() => { if (drag !== null) moveSheet(store, drag, i); setDrag(null); }}
            onKeyDown={(e) => { if (e.key === 'Enter') setRenaming({ i, text: s.name }); }}>
            {renaming?.i === i ? (
              <input autoFocus value={renaming.text} aria-label={T('Sheet name')} maxLength={31}
                onChange={(e) => setRenaming({ i, text: e.target.value })} onBlur={finish}
                onKeyDown={(e) => { if (e.key === 'Enter') finish(); if (e.key === 'Escape') setRenaming(null); e.stopPropagation(); }} />
            ) : s.name}
          </div>
        ))}
      </div>
      <button className="ico sm" aria-label={T('New sheet')} title={T('New sheet')} onClick={() => addSheet(store, store.wb.sheets.length)}><Plus size={16} strokeWidth={S} /></button>
    </div>
  );
};
