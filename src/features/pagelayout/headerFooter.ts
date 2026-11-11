// Header and footer codes (ECMA-376 Part 1, 18.3.1.39): sections, fields and font switches.
import { T } from '../../ui/i18n';
import type { PageSetup } from '../../model/workbook';

export type HfField = 'P' | 'N' | 'D' | 'T' | 'F' | 'Z' | 'A';
export interface HfFont { name?: string; bold?: boolean; italic?: boolean; underline?: 'single' | 'double'; strike?: boolean; vert?: 'super' | 'sub'; size?: number; color?: string; outline?: boolean; shadow?: boolean }
export type HfRun =
  | { kind: 'text'; text: string; font: HfFont }
  | { kind: 'field'; field: HfField; offset?: number; font: HfFont }
  | { kind: 'picture'; font: HfFont };
export interface HfSections { L: HfRun[]; C: HfRun[]; R: HfRun[] }

const THEME = ['#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#0563C1', '#954F72'];

const tint = (hex: string, t: number): string => {
  const c = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16));
  const out = c.map((v) => Math.round(t < 0 ? v * (1 + t) : v + (255 - v) * t));
  return '#' + out.map((v) => Math.max(0, Math.min(255, v)).toString(16).padStart(2, '0')).join('').toUpperCase();
};

/** Parses a header or footer string into its left, center and right sections. */
export const parseHf = (src: string | undefined): HfSections => {
  const out: HfSections = { L: [], C: [], R: [] };
  if (!src) return out;
  let sec: keyof HfSections = 'C';
  let font: HfFont = {};
  let text = '';
  const flush = () => { if (text) { out[sec].push({ kind: 'text', text, font: { ...font } }); text = ''; } };
  for (let i = 0; i < src.length; i++) {
    const ch = src[i]!;
    if (ch !== '&') { text += ch; continue; }
    const n = src[i + 1];
    if (n === undefined) { text += '&'; break; }
    if (n === '&') { text += '&'; i++; continue; }
    const up = n.toUpperCase();
    if (up === 'L' || up === 'C' || up === 'R') { flush(); sec = up; font = {}; i++; continue; }
    if ('PNDTFZA'.includes(up)) {
      flush();
      let j = i + 2;
      let offset: number | undefined;
      if (up === 'P') { const m = /^([+-]\d+)/.exec(src.slice(j)); if (m) { offset = +m[1]!; j += m[1]!.length; } }
      out[sec].push(offset !== undefined ? { kind: 'field', field: up as HfField, offset, font: { ...font } } : { kind: 'field', field: up as HfField, font: { ...font } });
      i = j - 1;
      continue;
    }
    if (up === 'G') { flush(); out[sec].push({ kind: 'picture', font: { ...font } }); i++; continue; }
    flush();
    switch (up) {
      case 'B': font.bold = !font.bold; i++; continue;
      case 'I': font.italic = !font.italic; i++; continue;
      case 'U': font.underline = font.underline === 'single' ? undefined : 'single'; i++; continue;
      case 'E': font.underline = font.underline === 'double' ? undefined : 'double'; i++; continue;
      case 'S': font.strike = !font.strike; i++; continue;
      case 'X': font.vert = font.vert === 'super' ? undefined : 'super'; i++; continue;
      case 'Y': font.vert = font.vert === 'sub' ? undefined : 'sub'; i++; continue;
      case 'O': font.outline = !font.outline; i++; continue;
      case 'H': font.shadow = !font.shadow; i++; continue;
      case 'K': {
        const six = src.slice(i + 2, i + 8);
        const th = /^(\d{2})([+-])(\d{3})/.exec(src.slice(i + 2, i + 9));
        if (th) { font.color = tint(THEME[+th[1]! % THEME.length]!, (th[2] === '-' ? -1 : 1) * Math.min(1, +th[3]! / 100)); i += 7; }
        else if (/^[0-9A-Fa-f]{6}$/.test(six)) { font.color = '#' + six.toUpperCase(); i += 7; }
        else i++;
        continue;
      }
      case '"': {
        const end = src.indexOf('"', i + 2);
        if (end < 0) { i = src.length; continue; }
        const [name, style = ''] = src.slice(i + 2, end).split(',');
        if (name && name !== '-') font.name = name.trim();
        const st = style.toLowerCase();
        if (st) { font.bold = /bold|heavy|black/.test(st); font.italic = /italic|oblique/.test(st); }
        i = end;
        continue;
      }
      default: {
        const m = /^\d{1,3}/.exec(src.slice(i + 1));
        if (m) { const sz = +m[0]; if (sz > 0 && sz <= 409) font.size = sz; i += m[0].length; continue; }
        text += '&';
      }
    }
  }
  flush();
  return out;
};

/** Raw text of each section (codes kept), for the editor. */
export const splitHf = (src: string | undefined): { L: string; C: string; R: string } => {
  const out = { L: '', C: '', R: '' };
  if (!src) return out;
  let sec: 'L' | 'C' | 'R' = 'C';
  for (let i = 0; i < src.length; i++) {
    const ch = src[i]!;
    if (ch === '&' && i + 1 < src.length) {
      const n = src[i + 1]!;
      if (n === '&') { out[sec] += '&&'; i++; continue; }
      const up = n.toUpperCase();
      if (up === 'L' || up === 'C' || up === 'R') { sec = up; i++; continue; }
      if (n === '"') { const end = src.indexOf('"', i + 2); const stop = end < 0 ? src.length - 1 : end; out[sec] += src.slice(i, stop + 1); i = stop; continue; }
    }
    out[sec] += ch;
  }
  return out;
};

export const joinHf = (s: { L: string; C: string; R: string }): string =>
  (s.L ? `&L${s.L}` : '') + (s.C ? `&C${s.C}` : '') + (s.R ? `&R${s.R}` : '');

export interface HfContext { page: number; pages: number; date: Date; file: string; path: string; sheet: string }

export const fieldText = (f: HfField, offset: number | undefined, c: HfContext): string => {
  switch (f) {
    case 'P': return String(c.page + (offset ?? 0));
    case 'N': return String(c.pages);
    case 'D': return c.date.toLocaleDateString();
    case 'T': return c.date.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
    case 'F': return c.file;
    case 'Z': return c.path;
    case 'A': return c.sheet;
  }
};

/** Plain text of one section (for previews and lists). */
export const sectionText = (runs: HfRun[], c: HfContext): string =>
  runs.map((r) => (r.kind === 'text' ? r.text : r.kind === 'field' ? fieldText(r.field, r.offset, c) : '[Picture]')).join('');

/** Which header and footer strings apply to a page (1-based position within the printout). */
export const hfFor = (s: PageSetup, position: number): { header?: string; footer?: string; slot: '' | 'FIRST' | 'EVEN' } => {
  if (s.differentFirst && position === 1) return { header: s.firstHeader, footer: s.firstFooter, slot: 'FIRST' };
  if (s.differentOddEven && position % 2 === 0) return { header: s.evenHeader, footer: s.evenFooter, slot: 'EVEN' };
  return { header: s.header, footer: s.footer, slot: '' };
};

/** Header text in the interface language, with &P / &N codes for the page fields and & escaped. */
const hfText = (msg: string): string => msg.replace(/&/g, '&&').replace(/\u0001/g, '&P').replace(/\u0002/g, '&N');
const pageOf = (): string => hfText(T('Page {page} of {pages}', { page: '\u0001', pages: '\u0002' }));
const pageN = (): string => hfText(T('Page {page}', { page: '\u0001' }));

/** Ready-made headers and footers: [label, code string]. */
export const hfPresets = (author?: string): [string, string][] => [
  [T('(none)'), ''],
  [T('Page 1'), `&C${pageN()}`],
  [T('Page 1 of ?'), `&C${pageOf()}`],
  [T('Sheet name'), '&C&A'],
  [T('File name'), '&C&F'],
  [T('File path'), '&C&Z&F'],
  [T('Sheet name, Page 1'), `&L&A&R${pageN()}`],
  [T('File name, Page 1'), `&L&F&R${pageN()}`],
  [T('Date and time'), '&L&D&R&T'],
  [T('Confidential, date, page'), `&L${hfText(T('Confidential'))}&C&D&R${pageN()}`],
  ...(author ? [[T('Prepared by {author}, date, page', { author }), `&L${hfText(T('Prepared by {author}', { author }))} &D&R${pageN()}`] as [string, string]] : []),
  [T('Sheet name, Page 1 of ?'), `&L&A&C${pageOf()}&R&D`],
];
