// The earliest binary worksheet files (BIFF2, BIFF3 and BIFF4, and BIFF4 workbooks): plain record
// streams without a compound-file wrapper. Values, cached formula results, number formats, column
// widths and row heights are read; the formulas themselves use an older token set and are kept as values.
import { cellKey, internStyle, newSheet, type Cell, type Sheet, type Workbook } from '../model/workbook';
import { CellError, type Scalar } from '../engine/types';
import { ERR_BY_CODE } from './biff8-formula';
import { builtinFormat } from './numfmt';
import { shownColWidth } from './xlsx-chart';

/** True for a BIFF2/3/4 stream (a BOF record of one of those versions at the start). */
export const isOldBiff = (b: Uint8Array): boolean =>
  b.length >= 8 && ((b[0] === 0x09 && (b[1] === 0x00 || b[1] === 0x02 || b[1] === 0x04)) && b[2]! >= 4 && b[2]! <= 16 && b[3] === 0);

// The format list of the first versions: built-in codes in this order when the file has none.
const OLD_BUILTIN = ['General', '0', '0.00', '#,##0', '#,##0.00', '"$"#,##0_);\\("$"#,##0\\)', '"$"#,##0_);[Red]\\("$"#,##0\\)',
  '"$"#,##0.00_);\\("$"#,##0.00\\)', '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)', '0%', '0.00%', '0.00E+00', 'm/d/yy', 'd-mmm-yy', 'd-mmm', 'mmm-yy',
  'h:mm AM/PM', 'h:mm:ss AM/PM', 'h:mm', 'h:mm:ss', 'm/d/yy h:mm'];

const rk = (v: number): number => {
  let n: number;
  if (v & 2) n = v >> 2;
  else { const b = new DataView(new ArrayBuffer(8)); b.setUint32(4, v & 0xfffffffc, true); n = b.getFloat64(0, true); }
  return v & 1 ? n / 100 : n;
};

export const readOldBiff = (data: Uint8Array): Workbook => {
  const dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const wb: Workbook = { sheets: [], names: [], styles: [{}], date1904: false, activeSheet: 0 };
  let decoder = new TextDecoder('windows-1252');
  const text = (o: number, n: number) => decoder.decode(data.subarray(o, Math.min(data.length, o + n)));
  const formats: string[] = [];
  const xfFormat: number[] = [];
  const names: string[] = [];
  let sheet: Sheet | null = null;
  let version = 2;
  let pendingString: Cell | null = null;
  const styleCache = new Map<number, number>();
  const styleOfFormat = (ifmt: number): number => {
    const hit = styleCache.get(ifmt);
    if (hit !== undefined) return hit;
    const code = formats[ifmt] ?? OLD_BUILTIN[ifmt] ?? builtinFormat(ifmt);
    const id = code && code !== 'General' ? internStyle(wb, { numFmt: code }) : 0;
    styleCache.set(ifmt, id);
    return id;
  };
  const put = (row: number, col: number, v: Scalar, ifmt: number): Cell | null => {
    if (!sheet || row >= 65536 || col >= 256) return null;
    const cell: Cell = { v };
    const s = styleOfFormat(ifmt);
    if (s) cell.s = s;
    sheet.cells.set(cellKey(row, col), cell);
    return cell;
  };
  const fmtOfXf = (ixfe: number): number => xfFormat[ixfe] ?? 0;
  const formulaValue = (o: number): Scalar | 'string' => {
    if (dv.getUint16(o + 6, true) === 0xffff) {
      const t = data[o]!;
      if (t === 0) return 'string';
      if (t === 1) return !!data[o + 2];
      if (t === 2) return new CellError(ERR_BY_CODE[data[o + 2]!] ?? '#N/A');
      return '';
    }
    return dv.getFloat64(o, true);
  };

  let o = 0;
  let records = 0;
  while (o + 4 <= data.length && records++ < 5_000_000) {
    const type = dv.getUint16(o, true), size = dv.getUint16(o + 2, true);
    const d = o + 4;
    const end = Math.min(data.length, d + size);
    o = end;
    try {
      switch (type) {
        case 0x0009: case 0x0209: case 0x0409: {
          version = type === 0x0009 ? 2 : type === 0x0209 ? 3 : 4;
          const dt = dv.getUint16(d + 2, true);
          if (dt === 0x0100) break; // BIFF4 workbook globals: sheets follow
          sheet = newSheet(names[wb.sheets.length] ?? `Sheet${wb.sheets.length + 1}`);
          wb.sheets.push(sheet);
          break;
        }
        case 0x008f: { // BIFF4 workbook: BUNDLESHEET
          const cch = data[d + 6]!;
          names.push(text(d + 7, cch).slice(0, 31));
          break;
        }
        case 0x0042: { const cp = dv.getUint16(d, true); try { decoder = new TextDecoder(cp === 0x8000 ? 'macintosh' : `windows-${cp}`); } catch { /* keep */ } break; }
        case 0x0022: wb.date1904 = dv.getUint16(d, true) === 1; break;
        case 0x001e: formats.push(text(d + 1, data[d]!)); break;
        case 0x041e: { const id = dv.getUint16(d, true); formats[id] = text(d + 3, data[d + 2]!); break; }
        case 0x0043: xfFormat.push(data[d + 2]! & 0x3f); break;
        case 0x0243: case 0x0443: xfFormat.push(data[d + 1]!); break;
        // BIFF2 cells: row, column, three attribute bytes (the number format is in the second)
        case 0x0002: put(dv.getUint16(d, true), dv.getUint16(d + 2, true), dv.getUint16(d + 7, true), data[d + 5]! & 0x3f); break;
        case 0x0003: put(dv.getUint16(d, true), dv.getUint16(d + 2, true), dv.getFloat64(d + 7, true), data[d + 5]! & 0x3f); break;
        case 0x0004: put(dv.getUint16(d, true), dv.getUint16(d + 2, true), text(d + 8, data[d + 7]!), 0); break;
        case 0x0005: {
          const v = data[d + 7]!, err = data[d + 8]!;
          put(dv.getUint16(d, true), dv.getUint16(d + 2, true), err ? new CellError(ERR_BY_CODE[v] ?? '#N/A') : !!v, 0);
          break;
        }
        case 0x0006: {
          const v = formulaValue(d + 7);
          const c = put(dv.getUint16(d, true), dv.getUint16(d + 2, true), v === 'string' ? '' : v, data[d + 5]! & 0x3f);
          pendingString = v === 'string' ? c : null;
          break;
        }
        case 0x0007: if (pendingString) { pendingString.v = text(d + 1, data[d]!); pendingString = null; } break;
        // BIFF3/4 cells: row, column, XF index
        case 0x0203: put(dv.getUint16(d, true), dv.getUint16(d + 2, true), dv.getFloat64(d + 6, true), fmtOfXf(dv.getUint16(d + 4, true))); break;
        case 0x027e: put(dv.getUint16(d, true), dv.getUint16(d + 2, true), rk(dv.getUint32(d + 6, true)), fmtOfXf(dv.getUint16(d + 4, true))); break;
        case 0x0204: put(dv.getUint16(d, true), dv.getUint16(d + 2, true), text(d + 8, dv.getUint16(d + 6, true)), 0); break;
        case 0x0205: {
          const v = data[d + 6]!, err = data[d + 7]!;
          put(dv.getUint16(d, true), dv.getUint16(d + 2, true), err ? new CellError(ERR_BY_CODE[v] ?? '#N/A') : !!v, 0);
          break;
        }
        case 0x0206: case 0x0406: {
          const v = formulaValue(d + 6);
          const c = put(dv.getUint16(d, true), dv.getUint16(d + 2, true), v === 'string' ? '' : v, fmtOfXf(dv.getUint16(d + 4, true)));
          pendingString = v === 'string' ? c : null;
          break;
        }
        case 0x0207: if (pendingString) { pendingString.v = text(d + 2, dv.getUint16(d, true)); pendingString = null; } break;
        case 0x0024: if (sheet) { const w = dv.getUint16(d + 2, true); for (let c = data[d]!; c <= data[d + 1]!; c++) sheet.cols.set(c, { width: shownColWidth(w / 256) }); } break;
        case 0x007d: if (sheet) {
          const c1 = dv.getUint16(d, true), c2 = dv.getUint16(d + 2, true), w = dv.getUint16(d + 4, true), fl = dv.getUint16(d + 8, true);
          for (let c = c1; c <= Math.min(c2, 255); c++) sheet.cols.set(c, { width: shownColWidth(w / 256), ...(fl & 1 ? { hidden: true } : {}) });
        } break;
        case 0x0008: case 0x0208: if (sheet) {
          const row = dv.getUint16(d, true), h = dv.getUint16(d + 6, true);
          if (!(h & 0x8000) && h) sheet.rows.set(row, { height: (h & 0x7fff) / 20 });
        } break;
        case 0x000a: if (version < 4 || names.length === 0) { /* end of a sheet */ } break;
        default: break;
      }
    } catch { /* a damaged record is skipped */ }
  }
  if (!wb.sheets.length) throw new Error("This file isn't a workbook ClerdoGrid can read.");
  return wb;
};
