import { N } from '../../ui/i18n';
// Paper sizes by their file-format code (ECMA-376 Part 1, 18.3.1.63 pageSetup@paperSize), in points.
const IN = 72, MM = 72 / 25.4;

export interface Paper { code: number; name: string; w: number; h: number }

const P = (code: number, name: string, w: number, h: number, unit: number): Paper => ({ code, name, w: w * unit, h: h * unit });

export const PAPERS: Paper[] = [
  P(1, N('Letter'), 8.5, 11, IN), P(5, N('Legal'), 8.5, 14, IN), P(7, N('Executive'), 7.25, 10.5, IN), P(3, N('Tabloid'), 11, 17, IN), P(4, N('Ledger'), 17, 11, IN),
  P(6, N('Statement'), 5.5, 8.5, IN), P(14, N('Folio'), 8.5, 13, IN), P(16, N('10 × 14 in'), 10, 14, IN), P(17, N('11 × 17 in'), 11, 17, IN), P(18, N('Note'), 8.5, 11, IN),
  P(8, 'A3', 297, 420, MM), P(9, 'A4', 210, 297, MM), P(10, N('A4 small'), 210, 297, MM), P(11, 'A5', 148, 210, MM),
  P(12, 'B4 (JIS)', 257, 364, MM), P(13, 'B5 (JIS)', 182, 257, MM), P(15, N('Quarto'), 215, 275, MM),
  P(24, N('C size sheet'), 17, 22, IN), P(25, N('D size sheet'), 22, 34, IN), P(26, N('E size sheet'), 34, 44, IN),
  P(19, N('Envelope #9'), 3.875, 8.875, IN), P(20, N('Envelope #10'), 4.125, 9.5, IN), P(21, N('Envelope #11'), 4.5, 10.375, IN), P(22, N('Envelope #12'), 4.75, 11, IN), P(23, N('Envelope #14'), 5, 11.5, IN),
  P(27, N('Envelope DL'), 110, 220, MM), P(28, N('Envelope C5'), 162, 229, MM), P(29, N('Envelope C3'), 324, 458, MM), P(30, N('Envelope C4'), 229, 324, MM), P(31, N('Envelope C6'), 114, 162, MM), P(32, N('Envelope C65'), 114, 229, MM),
  P(33, N('Envelope B4'), 250, 353, MM), P(34, N('Envelope B5'), 176, 250, MM), P(35, N('Envelope B6'), 176, 125, MM), P(36, N('Envelope Italy'), 110, 230, MM),
  P(37, N('Envelope Monarch'), 3.875, 7.5, IN), P(38, N('Envelope 6¾'), 3.625, 6.5, IN),
  P(39, N('US standard fanfold'), 14.875, 11, IN), P(40, N('German standard fanfold'), 8.5, 12, IN), P(41, N('German legal fanfold'), 8.5, 13, IN),
  P(42, 'B4 (ISO)', 250, 353, MM), P(43, N('Japanese postcard'), 100, 148, MM), P(44, N('9 × 11 in'), 9, 11, IN), P(45, N('10 × 11 in'), 10, 11, IN), P(46, N('15 × 11 in'), 15, 11, IN),
  P(47, N('Envelope invite'), 220, 220, MM), P(50, N('Letter extra'), 9.275, 12, IN), P(51, N('Legal extra'), 9.275, 15, IN), P(52, N('Tabloid extra'), 11.69, 18, IN),
  P(53, N('A4 extra'), 236, 322, MM), P(54, N('Letter transverse'), 8.275, 11, IN), P(55, N('A4 transverse'), 210, 297, MM), P(56, N('Letter extra transverse'), 9.275, 12, IN),
  P(57, N('Super A/A4'), 227, 356, MM), P(58, N('Super B/A3'), 305, 487, MM), P(59, N('Letter plus'), 8.5, 12.69, IN), P(60, N('A4 plus'), 210, 330, MM),
  P(61, N('A5 transverse'), 148, 210, MM), P(62, N('B5 (JIS) transverse'), 182, 257, MM), P(63, N('A3 extra'), 322, 445, MM), P(64, N('A5 extra'), 174, 235, MM),
  P(65, N('B5 (ISO) extra'), 201, 276, MM), P(66, 'A2', 420, 594, MM), P(67, N('A3 transverse'), 297, 420, MM), P(68, N('A3 extra transverse'), 322, 445, MM),
];

/** Sizes offered first in the page setup list. */
export const COMMON_PAPERS = [1, 5, 7, 3, 8, 9, 11, 12, 13, 66, 20, 27, 28, 30, 37];

const BY_CODE = new Map(PAPERS.map((p) => [p.code, p]));
export const paperByCode = (code: number | undefined): Paper => BY_CODE.get(code ?? 1) ?? BY_CODE.get(1)!;

/** "210mm" / "8.5in" / "12cm" / "600pt" → points. */
export const lengthPt = (s: string | undefined): number | undefined => {
  const m = s && /^(\d+(?:\.\d+)?)(mm|cm|in|pt|pc|pi)$/.exec(s.trim());
  if (!m) return undefined;
  const v = +m[1]!;
  const pt = m[2] === 'mm' ? v * MM : m[2] === 'cm' ? v * MM * 10 : m[2] === 'in' ? v * IN : m[2] === 'pt' ? v : v * 12;
  return pt > 36 && pt < 14400 ? pt : undefined;
};

export interface MarginPreset { id: string; label: string; m: { left: number; right: number; top: number; bottom: number; header: number; footer: number } }
export const MARGIN_PRESETS: MarginPreset[] = [
  { id: 'normal', label: N('Normal'), m: { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7, header: 0.3, footer: 0.3 } },
  { id: 'wide', label: N('Wide'), m: { top: 1, bottom: 1, left: 1, right: 1, header: 0.5, footer: 0.5 } },
  { id: 'narrow', label: N('Narrow'), m: { top: 0.75, bottom: 0.75, left: 0.25, right: 0.25, header: 0.3, footer: 0.3 } },
  { id: 'none', label: N('None'), m: { top: 0, bottom: 0, left: 0, right: 0, header: 0, footer: 0 } },
];
