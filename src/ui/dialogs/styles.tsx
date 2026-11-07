import { useMemo, useState } from 'react';
import { mergeStyle, type BorderSide, type CellRange, type CellStyle, type TableDef } from '../../model/workbook';
import { MAX_COLS, MAX_ROWS } from '../../engine/types';
import { currentRegion, looksLikeHeader } from '../../app/ops/data';
import type { GridStore } from '../../app/store';
import { Dialog, Field, type DialogProps } from './Dialog';
import { Check, cleanStyle, cssOf, overlaps, parseRefs, rangeLabel, scalarText } from './shared';

// ---------- Cell styles ----------

const thin = (color = '#7f7f7f'): BorderSide => ({ style: 'thin', color });
const ACCENTS = ['#0C7A4B', '#e8590c', '#1c9a5b', '#9c36b5', '#0c8599', '#c2255c'];
const tint = (hex: string, t: number): string => {
  const p = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16));
  return `#${p.map((x) => Math.round(x + (255 - x) * t).toString(16).padStart(2, '0')).join('')}`;
};

interface NamedStyle { name: string; group: string; style: CellStyle }

export const NAMED_STYLES: NamedStyle[] = [
  { name: 'Normal', group: 'Basic', style: {} },
  { name: 'Good', group: 'Good, bad and neutral', style: { font: { color: '#006100' }, fill: { color: '#c6efce' } } },
  { name: 'Bad', group: 'Good, bad and neutral', style: { font: { color: '#9c0006' }, fill: { color: '#ffc7ce' } } },
  { name: 'Neutral', group: 'Good, bad and neutral', style: { font: { color: '#9c5700' }, fill: { color: '#ffeb9c' } } },
  { name: 'Calculation', group: 'Data and model', style: { font: { bold: true, color: '#fa7d00' }, fill: { color: '#f2f2f2' }, border: { top: thin(), bottom: thin(), left: thin(), right: thin() } } },
  { name: 'Check cell', group: 'Data and model', style: { font: { bold: true, color: '#ffffff' }, fill: { color: '#a5a5a5' }, border: { top: { style: 'double', color: '#3f3f3f' }, bottom: { style: 'double', color: '#3f3f3f' }, left: { style: 'double', color: '#3f3f3f' }, right: { style: 'double', color: '#3f3f3f' } } } },
  { name: 'Explanatory', group: 'Data and model', style: { font: { italic: true, color: '#7f7f7f' } } },
  { name: 'Input', group: 'Data and model', style: { font: { color: '#3f3f76' }, fill: { color: '#ffcc99' }, border: { top: thin(), bottom: thin(), left: thin(), right: thin() } } },
  { name: 'Linked cell', group: 'Data and model', style: { font: { color: '#fa7d00' }, border: { bottom: { style: 'double', color: '#ff8001' } } } },
  { name: 'Note', group: 'Data and model', style: { fill: { color: '#ffffcc' }, border: { top: thin('#b2b2b2'), bottom: thin('#b2b2b2'), left: thin('#b2b2b2'), right: thin('#b2b2b2') } } },
  { name: 'Output', group: 'Data and model', style: { font: { bold: true, color: '#3f3f3f' }, fill: { color: '#f2f2f2' }, border: { top: thin('#3f3f3f'), bottom: thin('#3f3f3f'), left: thin('#3f3f3f'), right: thin('#3f3f3f') } } },
  { name: 'Warning text', group: 'Data and model', style: { font: { color: '#ff0000' } } },
  { name: 'Title', group: 'Titles and headings', style: { font: { size: 18, color: '#1f2a44', name: 'Calibri Light' } } },
  { name: 'Heading 1', group: 'Titles and headings', style: { font: { bold: true, size: 15, color: '#1f2a44' }, border: { bottom: { style: 'thick', color: '#0C7A4B' } } } },
  { name: 'Heading 2', group: 'Titles and headings', style: { font: { bold: true, size: 13, color: '#1f2a44' }, border: { bottom: { style: 'thick', color: '#a3b1ff' } } } },
  { name: 'Heading 3', group: 'Titles and headings', style: { font: { bold: true, color: '#1f2a44' }, border: { bottom: { style: 'medium', color: '#a3b1ff' } } } },
  { name: 'Heading 4', group: 'Titles and headings', style: { font: { bold: true, color: '#1f2a44' } } },
  { name: 'Total', group: 'Titles and headings', style: { font: { bold: true }, border: { top: thin('#0C7A4B'), bottom: { style: 'double', color: '#0C7A4B' } } } },
  ...ACCENTS.flatMap((c, i) => [
    { name: `20% Accent ${i + 1}`, group: 'Accents', style: { fill: { color: tint(c, 0.8) } } },
    { name: `40% Accent ${i + 1}`, group: 'Accents', style: { fill: { color: tint(c, 0.6) } } },
    { name: `60% Accent ${i + 1}`, group: 'Accents', style: { font: { color: '#ffffff' }, fill: { color: tint(c, 0.4) } } },
    { name: `Accent ${i + 1}`, group: 'Accents', style: { font: { color: '#ffffff' }, fill: { color: c } } },
  ]),
  { name: 'Comma', group: 'Number format', style: { numFmt: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)' } },
  { name: 'Comma [0]', group: 'Number format', style: { numFmt: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)' } },
  { name: 'Currency', group: 'Number format', style: { numFmt: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)' } },
  { name: 'Currency [0]', group: 'Number format', style: { numFmt: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)' } },
  { name: 'Percent', group: 'Number format', style: { numFmt: '0%' } },
];

export const CellStylesDialog = ({ store, close }: DialogProps) => {
  const groups = useMemo(() => [...new Set(NAMED_STYLES.map((s) => s.group))], []);
  const apply = (ns: NamedStyle) => {
    // A named style replaces the parts it defines and resets the rest of those parts.
    store.applyStyle((s) => {
      if (ns.name === 'Normal') return s.protection ? { protection: s.protection } : {};
      const next: CellStyle = { ...s };
      if (ns.style.font || ns.group !== 'Number format') next.font = ns.style.font ? { ...ns.style.font } : undefined;
      if (ns.group !== 'Number format') { next.fill = ns.style.fill; next.border = ns.style.border; }
      if (ns.style.numFmt) next.numFmt = ns.style.numFmt;
      return cleanStyle(next);
    }, `Style: ${ns.name}`);
    close();
  };
  return (
    <Dialog title="Cell styles" close={close} wide>
      {groups.map((g) => (
        <section key={g} className="b-section">
          <h3 className="b-h3">{g}</h3>
          <div className="b-gallery">
            {NAMED_STYLES.filter((s) => s.group === g).map((s) => (
              <button key={s.name} className="b-stylechip" style={{ ...cssOf(s.style), fontSize: s.style.font?.size && s.style.font.size > 13 ? 16 : 13 }} onClick={() => apply(s)} title={s.name}>
                {s.group === 'Number format' ? <><span>{s.name}</span><span className="muted small">{s.name.startsWith('Percent') ? '12%' : s.name.startsWith('Currency') ? '$ 1,234' : '1,234'}</span></> : s.name}
              </button>
            ))}
          </div>
        </section>
      ))}
    </Dialog>
  );
};

// ---------- Tables ----------

interface TableLook { id: string; label: string; header: string; headerInk: string; band: string; line: string; dark?: boolean }

const look = (kind: 'Light' | 'Medium' | 'Dark', i: number, c: string): TableLook => kind === 'Light'
  ? { id: `TableStyleLight${i}`, label: `Light ${i}`, header: '#ffffff', headerInk: c, band: tint(c, 0.88), line: c }
  : kind === 'Medium'
    ? { id: `TableStyleMedium${i}`, label: `Medium ${i}`, header: c, headerInk: '#ffffff', band: tint(c, 0.8), line: tint(c, 0.5) }
    : { id: `TableStyleDark${i}`, label: `Dark ${i}`, header: '#1f2430', headerInk: '#ffffff', band: tint(c, 0.25), line: c, dark: true };

export const TABLE_LOOKS: TableLook[] = [
  look('Light', 1, '#3f3f46'), look('Light', 2, '#0C7A4B'), look('Light', 3, '#e8590c'), look('Light', 4, '#1c9a5b'),
  look('Medium', 2, '#0C7A4B'), look('Medium', 3, '#e8590c'), look('Medium', 4, '#1c9a5b'), look('Medium', 5, '#9c36b5'), look('Medium', 6, '#0c8599'), look('Medium', 7, '#52525b'),
  look('Dark', 2, '#0C7A4B'), look('Dark', 3, '#e8590c'),
];

const TableSwatch = ({ t }: { t: TableLook }) => (
  <svg viewBox="0 0 60 44" width="60" height="44" aria-hidden="true">
    <rect x="0" y="0" width="60" height="44" fill={t.dark ? t.band : '#ffffff'} />
    <rect x="0" y="0" width="60" height="9" fill={t.header} />
    {[1, 3].map((k) => <rect key={k} x="0" y={9 + k * 8.75} width="60" height="8.75" fill={t.dark ? tint(t.band, 0.2) : t.band} />)}
    <line x1="0" y1="9" x2="60" y2="9" stroke={t.line} strokeWidth="1.5" />
    {[0, 1, 2].map((k) => <rect key={k} x={4 + k * 19} y="3" width="14" height="3" rx="1" fill={t.headerInk} opacity=".85" />)}
    <rect x="0.5" y="0.5" width="59" height="43" fill="none" stroke={t.line} strokeOpacity=".5" />
  </svg>
);

const uniqueTableName = (store: GridStore): string => {
  const taken = new Set([...store.wb.sheets.flatMap((s) => s.tables.map((t) => t.name.toLowerCase())), ...store.wb.names.map((n) => n.name.toLowerCase())]);
  for (let i = 1; ; i++) if (!taken.has(`table${i}`)) return `Table${i}`;
};

/** Formats a table's cells from its look. Exported so table options can re-apply it. */
export const styleTable = (store: GridStore, t: TableDef, lk: TableLook): void => {
  const r = t.range;
  const first = r.r1 + (t.headerRow ? 1 : 0), last = r.r2 - (t.totalsRow ? 1 : 0);
  const reset = (s: CellStyle): CellStyle => cleanStyle({ ...s, fill: undefined, border: undefined, font: s.font ? { ...s.font, bold: undefined, color: undefined } : undefined });
  store.styleRange(r, reset);
  if (t.headerRow) store.styleRange({ ...r, r2: r.r1 }, (s) => cleanStyle(mergeStyle(s, { font: { bold: true, color: lk.headerInk }, fill: lk.header === '#ffffff' ? undefined : { color: lk.header }, border: { bottom: { style: 'thin', color: lk.line } } })));
  for (let row = first; row <= last; row++) {
    const k = row - first;
    const banded = t.bandedRows && k % 2 === 0;
    for (let c = r.c1; c <= r.c2; c++) {
      const bandCol = t.bandedCols && (c - r.c1) % 2 === 0;
      const color = banded || bandCol ? (lk.dark ? tint(lk.band, 0.2) : lk.band) : lk.dark ? lk.band : undefined;
      if (!color && !lk.dark) continue;
      store.styleRange({ r1: row, r2: row, c1: c, c2: c }, (s) => cleanStyle(mergeStyle(s, { fill: color ? { color } : undefined, font: lk.dark ? { color: '#ffffff' } : undefined })));
    }
  }
  if (t.totalsRow) store.styleRange({ ...r, r1: r.r2 }, (s) => cleanStyle(mergeStyle(s, { font: { bold: true }, border: { top: { style: 'double', color: lk.line } } })));
};

const MAX_TABLE_CELLS = 2_000_000;

export const TableDialog = ({ store, close }: DialogProps) => {
  const existing = store.sheet.tables.find((t) => overlaps(t.range, { r1: store.sel.active.row, c1: store.sel.active.col, r2: store.sel.active.row, c2: store.sel.active.col }));
  const initial = useMemo<CellRange>(() => {
    if (existing) return existing.range;
    const r = store.range;
    return r.r1 === r.r2 && r.c1 === r.c2 ? currentRegion(store) : r;
  }, [store, existing]);
  const [ref, setRef] = useState(rangeLabel(initial));
  const [headers, setHeaders] = useState(existing ? existing.headerRow : looksLikeHeader(store, initial) || typeof store.value(initial.r1, initial.c1) === 'string');
  const [lookId, setLookId] = useState(existing?.style && TABLE_LOOKS.some((l) => l.id === existing.style) ? existing.style : 'TableStyleMedium2');
  const [totals, setTotals] = useState(existing?.totalsRow ?? false);
  const [bandRows, setBandRows] = useState(existing?.bandedRows ?? true);
  const [bandCols, setBandCols] = useState(existing?.bandedCols ?? false);
  const [filter, setFilter] = useState(existing ? !!store.sheet.autoFilter && overlaps(store.sheet.autoFilter.range, existing.range) : true);

  const ok = () => {
    const parsed = parseRefs(store, ref);
    if (!parsed || parsed.ranges.length !== 1 || parsed.sheet !== store.sheetIndex) { store.notify('Enter a range on this sheet, like A1:D20.'); return false; }
    let r = parsed.ranges[0]!;
    const b = store.engine.usedBounds(store.sheetIndex);
    if (r.r2 === MAX_ROWS - 1) r = { ...r, r2: Math.max(r.r1, b.rows) };
    if (r.c2 === MAX_COLS - 1) r = { ...r, c2: Math.max(r.c1, b.cols) };
    if ((r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1) > MAX_TABLE_CELLS) { store.notify('That range is too big for a table.'); return false; }
    const others = store.sheet.tables.filter((t) => t !== existing);
    if (others.some((t) => overlaps(t.range, r))) { store.notify("A table can't overlap another table."); return false; }
    if (store.sheet.merges.some((m) => overlaps(m, r))) { store.notify("A table can't contain merged cells. Unmerge them first."); return false; }
    const lk = TABLE_LOOKS.find((l) => l.id === lookId) ?? TABLE_LOOKS[4]!;
    const rowFree = (row: number) => row < MAX_ROWS && Array.from({ length: r.c2 - r.c1 + 1 }, (_, j) => store.value(row, r.c1 + j)).every((v) => v === null);
    const needsRowBelow = (!existing && !headers && (r.r1 === 0 || !rowFree(r.r1 - 1))) || (totals && !(existing?.totalsRow ?? false));
    if (needsRowBelow && !rowFree(r.r2 + 1)) { store.notify('The row below the data needs to be empty for the new row. Insert a row there first.'); return false; }
    const sheetIdx = store.sheetIndex;
    store.change(existing ? 'Table options' : 'Format as table', 'sheet', () => {
      // Without headers, insert a header row above the data (as when the first row is data).
      if (!headers && !existing) {
        if (r.r1 === 0 || [...Array(r.c2 - r.c1 + 1)].some((_, j) => store.value(r.r1 - 1, r.c1 + j) !== null)) {
          // No room above: the first row becomes the header row and data shifts down one row.
          for (let row = r.r2; row >= r.r1; row--) for (let c = r.c1; c <= r.c2; c++) {
            const cell = store.cell(row, c);
            if (cell?.f) store.engine.setFormula(sheetIdx, row + 1, c, cell.f); else store.engine.setValue(sheetIdx, row + 1, c, cell?.v ?? null);
            const placed = store.cell(row + 1, c); if (placed) placed.s = cell?.s;
          }
          r = { ...r, r2: r.r2 + 1 };
        } else r = { ...r, r1: r.r1 - 1 };
        for (let c = r.c1; c <= r.c2; c++) store.engine.setValue(sheetIdx, r.r1, c, `Column${c - r.c1 + 1}`);
      }
      const names: string[] = [];
      for (let c = r.c1; c <= r.c2; c++) {
        let n = scalarText(store.value(r.r1, c)).trim() || `Column${c - r.c1 + 1}`;
        const stem = n;
        for (let k = 2; names.some((x) => x.toLowerCase() === n.toLowerCase()); k++) n = `${stem}${k}`;
        names.push(n);
        if (scalarText(store.value(r.r1, c)) !== n) store.engine.setValue(sheetIdx, r.r1, c, n);
      }
      // Totals row: add one row below the data, or remove it.
      const hadTotals = existing?.totalsRow ?? false;
      if (totals && !hadTotals) r = { ...r, r2: r.r2 + 1 };
      if (!totals && hadTotals) r = { ...r, r2: r.r2 - 1 };
      const name = existing?.name ?? uniqueTableName(store);
      const t: TableDef = {
        name, range: r, headerRow: true, totalsRow: totals, style: lk.id, bandedRows: bandRows, bandedCols: bandCols,
        columns: names.map((n, i) => ({ name: n, totalsFunction: existing?.columns[i]?.totalsFunction ?? (totals && i === 0 ? 'label' : undefined) })),
      };
      if (existing) Object.assign(existing, t); else store.sheet.tables.push(t);
      store.engine.sheetsChanged();
      if (totals && !hadTotals) {
        const tr = r.r2;
        const esc = (n: string) => n.replace(/([[\]#'@])/g, "'$1");
        store.engine.setValue(sheetIdx, tr, r.c1, 'Total');
        names.forEach((n, i) => {
          if (i === 0) return;
          const c = r.c1 + i;
          let numeric = false;
          for (let row = r.r1 + 1; row < tr && !numeric; row++) numeric = typeof store.value(row, c) === 'number';
          if (numeric || i === names.length - 1) {
            store.engine.setFormula(sheetIdx, tr, c, `SUBTOTAL(109,${name}[${esc(n)}])`);
            t.columns[i]!.totalsFunction = 'sum';
          }
        });
      } else if (!totals && hadTotals) {
        for (let c = r.c1; c <= r.c2; c++) store.engine.clearCell(sheetIdx, r.r2 + 1, c);
      }
      styleTable(store, existing ?? t, lk);
      if (!totals && hadTotals) store.styleRange({ ...r, r1: r.r2 + 1, r2: r.r2 + 1 }, (s) => cleanStyle({ ...s, border: undefined, font: s.font ? { ...s.font, bold: undefined } : undefined }));
      const filterRange = { ...r, r2: totals ? r.r2 - 1 : r.r2 };
      if (filter) store.sheet.autoFilter = { range: filterRange, columns: [] };
      else if (store.sheet.autoFilter && overlaps(store.sheet.autoFilter.range, r)) store.sheet.autoFilter = undefined;
      store.engine.recalc(true);
    });
    store.notify(existing ? `Updated ${existing.name}.` : `Made a table from ${rangeLabel(r)}.`);
    return true;
  };

  const lk = TABLE_LOOKS.find((l) => l.id === lookId);
  return (
    <Dialog title={existing ? `Table: ${existing.name}` : 'Format as table'} close={close} onOk={ok} okLabel={existing ? 'Apply' : 'Create'} wide>
      <div className="grid2">
        <Field label="Data range"><input value={ref} disabled={!!existing} onChange={(e) => setRef(e.target.value)} placeholder="A1:D20" /></Field>
        <div className="b-checks b-pad">
          <Check label="My table has headers" checked={headers} disabled={!!existing} onChange={setHeaders} />
        </div>
      </div>
      <Field label="Style">
        <div className="b-gallery b-tables" role="listbox">
          {TABLE_LOOKS.map((t) => <button key={t.id} role="option" aria-selected={t.id === lookId} aria-label={t.label} title={t.label} className={`b-tablechip${t.id === lookId ? ' on' : ''}`} onClick={() => setLookId(t.id)}><TableSwatch t={t} /></button>)}
        </div>
      </Field>
      <div className="b-checks">
        <Check label="Header row" checked disabled onChange={() => undefined} />
        <Check label="Totals row" checked={totals} onChange={setTotals} />
        <Check label="Banded rows" checked={bandRows} onChange={setBandRows} />
        <Check label="Banded columns" checked={bandCols} onChange={setBandCols} />
        <Check label="Filter buttons" checked={filter} onChange={setFilter} />
      </div>
      {lk && <p className="muted small">{lk.label}. Columns are named from the first row{!headers && !existing ? ` (a header row Column1, Column2… is added)` : ''}. Totals use SUBTOTAL so filtered rows are left out.</p>}
    </Dialog>
  );
};
