// Saved Get Data queries as a small XML part (a custom XML item in .xlsx, "clerdo/queries.xml" in .ods).
import type { DataQuery } from '../model/workbook';
import { attr, kids, parseRoot } from './xml';
import { esc } from './xml';

export const QUERIES_NS = 'https://clerdo.app/ns/grid/queries/1';

export const queriesToXml = (list: DataQuery[]): string =>
  `<?xml version="1.0" encoding="UTF-8"?>\n<queries xmlns="${QUERIES_NS}">` + list.map((q) => {
    const a = (k: string, v: string | number | undefined) => (v === undefined ? '' : ` ${k}="${esc(String(v))}"`);
    return `<query${a('id', q.id)}${a('name', q.name)}${a('kind', q.kind)}${a('file', q.file)}${a('path', q.path)}${a('refreshed', q.refreshed)}`
      + `${a('sheet', q.target.sheet)}${a('row', q.target.row)}${a('col', q.target.col)}${a('rows', q.target.rows)}${a('cols', q.target.cols)}`
      + `${a('sourceSheet', q.source?.sheet)}${a('sourceRange', q.source?.range)}>${q.options ? esc(q.options) : ''}</query>`;
  }).join('') + '</queries>';

const int = (s: string | undefined, max: number): number => { const n = Math.floor(Number(s)); return Number.isFinite(n) && n >= 0 ? Math.min(n, max) : 0; };

/** Reads the part; anything malformed is skipped. */
export const queriesFromXml = (text: string): DataQuery[] => {
  let root;
  try { root = parseRoot(text); } catch { return []; }
  if (root.name !== 'queries' || attr(root, 'xmlns') !== QUERIES_NS) return [];
  return kids(root, 'query').slice(0, 500).flatMap((n): DataQuery[] => {
    const kind = attr(n, 'kind');
    if (kind !== 'text' && kind !== 'workbook') return [];
    const options = n.text.trim();
    if (options) { try { JSON.parse(options); } catch { return []; } }
    const q: DataQuery = {
      id: (attr(n, 'id') ?? '').replace(/[^\w-]/g, '').slice(0, 64) || `q${Math.random().toString(36).slice(2, 8)}`,
      name: (attr(n, 'name') ?? 'Query').slice(0, 200),
      kind,
      file: (attr(n, 'file') ?? '').slice(0, 260),
      target: { sheet: (attr(n, 'sheet') ?? '').slice(0, 31), row: int(attr(n, 'row'), 1_048_575), col: int(attr(n, 'col'), 16_383), rows: int(attr(n, 'rows'), 1_048_576), cols: int(attr(n, 'cols'), 16_384) },
    };
    if (options) q.options = options.slice(0, 100_000);
    const path = attr(n, 'path');
    if (path) q.path = path.slice(0, 4096);
    const refreshed = attr(n, 'refreshed');
    if (refreshed && !Number.isNaN(Date.parse(refreshed))) q.refreshed = refreshed;
    const ss = attr(n, 'sourceSheet');
    if (ss !== undefined) q.source = { sheet: ss.slice(0, 31), ...(attr(n, 'sourceRange') ? { range: attr(n, 'sourceRange')!.slice(0, 255) } : {}) };
    return [q];
  });
};
