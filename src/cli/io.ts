import { readCsv, writeCsv } from '../formats/csv';
import { readXlsx, writeXlsx } from '../formats/xlsx';
import { formatValue } from '../formats/numfmt';
import { CellError, ERROR_CODES, type ErrorCode, type Scalar } from '../engine/types';
import { Engine } from '../engine/engine';
import { cellKey, internStyle, newSheet, newWorkbook, rangeAddr, type CellRange, type Workbook } from '../model/workbook';
import { EXIT, UserError } from './errors';

export const MAX_FILE_BYTES = 512 * 1024 * 1024;

export type FileKind = 'xlsx' | 'csv' | 'tsv' | 'json' | 'ods' | 'xls' | 'pdf' | 'html';

const EXT_KIND: Record<string, FileKind> = {
  xlsx: 'xlsx', xlsm: 'xlsx', xltx: 'xlsx', xltm: 'xlsx',
  csv: 'csv', txt: 'csv', tsv: 'tsv', tab: 'tsv', json: 'json',
  ods: 'ods', ots: 'ods', xls: 'xls', xlt: 'xls', pdf: 'pdf', html: 'html', htm: 'html',
};

export const kindOf = (name: string): FileKind | undefined => EXT_KIND[/\.([^./\\]+)$/.exec(name)?.[1]?.toLowerCase() ?? ''];

// ---------- Optional format modules ----------
// Other format modules may or may not be present in src/formats; they are found by name at build time
// and loaded only when a file of that kind is used.

type Loader = () => Promise<Record<string, unknown>>;
const formatModules = import.meta.glob('../formats/{encryption,ods,ods-*,opendocument*,biff8-read,biff8-write,xls-read,xls-write,pdf-export,pdf-workbook,html-export,html-read,html}.ts') as Record<string, Loader>;

interface OptionalFormat { files: RegExp; read: RegExp; write: RegExp }
const OPTIONAL: Partial<Record<FileKind, OptionalFormat>> = {
  ods: { files: /\/(ods|opendocument)[^/]*\.ts$/i, read: /^(read|parse|import)Ods$/i, write: /^(write|export|build)Ods$/i },
  xls: { files: /\/(biff8?|xls)[^/]*\.ts$/i, read: /^(read|parse|import)(Xls|Biff8?)$/i, write: /^(write|export|build)(Xls|Biff8?)$/i },
  pdf: { files: /\/pdf[^/]*\.ts$/i, read: /^$/, write: /^(write|export|build|make)(Pdf|WorkbookPdf|SheetPdf)$/i },
  html: { files: /\/html[^/]*\.ts$/i, read: /^(read|parse|import)Html$/i, write: /^(write|export|build|make)(Html|WorkbookHtml|SheetHtml)$/i },
};

const findExport = async (kind: FileKind, which: 'read' | 'write'): Promise<((...a: unknown[]) => unknown) | null> => {
  const spec = OPTIONAL[kind];
  if (!spec) return null;
  for (const [path, load] of Object.entries(formatModules)) {
    if (!spec.files.test(path)) continue;
    const mod = await load();
    for (const [name, value] of Object.entries(mod)) if (typeof value === 'function' && spec[which].test(name)) return value as (...a: unknown[]) => unknown;
  }
  return null;
};

const encryption = async (): Promise<{ isEncryptedPackage: (b: Uint8Array) => boolean; decryptWorkbook: (b: Uint8Array, p: string) => Uint8Array; encryptWorkbook: (b: Uint8Array, p: string) => Uint8Array } | null> => {
  const entry = Object.entries(formatModules).find(([p]) => /\/encryption\.ts$/.test(p));
  const mod = entry ? await entry[1]() : null;
  return mod && typeof mod.decryptWorkbook === 'function' && typeof mod.encryptWorkbook === 'function' && typeof mod.isEncryptedPackage === 'function' ? mod as never : null;
};

export const canEncrypt = async (): Promise<boolean> => !!(await encryption());

/** Which file kinds this build can read and write. */
export const supportedKinds = async (): Promise<{ read: FileKind[]; write: FileKind[] }> => {
  const read: FileKind[] = ['xlsx', 'csv', 'tsv', 'json'];
  const write: FileKind[] = ['xlsx', 'csv', 'tsv', 'json'];
  for (const k of Object.keys(OPTIONAL) as FileKind[]) {
    if (await findExport(k, 'read')) read.push(k);
    if (await findExport(k, 'write')) write.push(k);
  }
  return { read, write };
};

const toBytes = (out: unknown): Uint8Array => {
  if (out instanceof Uint8Array) return out;
  if (out instanceof ArrayBuffer) return new Uint8Array(out);
  if (typeof out === 'string') return new TextEncoder().encode(out);
  throw new UserError('The writer for this format returned nothing usable.', EXIT.internal);
};

// ---------- Reading ----------

const sheetNameFrom = (name: string): string =>
  name.replace(/^.*[\\/]/, '').replace(/\.[^.]+$/, '').slice(0, 31).replace(/[\\/?*[\]:]/g, '_').replace(/^'+|'+$/g, '') || 'Sheet1';

const fromRows = (name: string, rows: Scalar[][], formats?: (string | undefined)[][]): Workbook => {
  const wb = newWorkbook();
  const sheet = wb.sheets[0]!;
  sheet.name = sheetNameFrom(name);
  rows.forEach((r, i) => r.forEach((v, j) => {
    if (v === null) return;
    const fmt = formats?.[i]?.[j];
    sheet.cells.set(cellKey(i, j), fmt ? { v, s: internStyle(wb, { numFmt: fmt }) } : { v });
  }));
  return wb;
};

const jsonScalar = (v: unknown): { v: Scalar; f?: string } => {
  if (v === null || v === undefined) return { v: null };
  if (typeof v === 'number') return { v: Number.isFinite(v) ? v : new CellError('#NUM!') };
  if (typeof v === 'boolean') return { v };
  if (typeof v === 'string') {
    if (v.startsWith('=') && v.length > 1) return { v: null, f: v.slice(1) };
    if (ERROR_CODES.includes(v as ErrorCode)) return { v: new CellError(v as ErrorCode) };
    return { v: v.length > 32767 ? v.slice(0, 32767) : v };
  }
  return { v: JSON.stringify(v) };
};

/**
 * Accepts the shape written by toJson ({ sheets: [{ name, rows }] }), a plain array of rows,
 * or an array of records (keys become the header row). Strings starting with "=" are formulas.
 */
export const workbookFromJson = (name: string, text: string): Workbook => {
  let data: unknown;
  try { data = JSON.parse(text); } catch { throw new UserError(`${name} isn't valid JSON.`, EXIT.file); }
  const asRows = (x: unknown): unknown[][] => {
    if (!Array.isArray(x)) throw new UserError(`${name}: expected an array of rows.`, EXIT.file);
    if (x.every((r) => Array.isArray(r))) return x as unknown[][];
    if (x.every((r) => r && typeof r === 'object' && !Array.isArray(r))) {
      const keys: string[] = [];
      for (const r of x as Record<string, unknown>[]) for (const k of Object.keys(r)) if (!keys.includes(k)) keys.push(k);
      return [keys, ...(x as Record<string, unknown>[]).map((r) => keys.map((k) => r[k] ?? null))];
    }
    throw new UserError(`${name}: rows must all be arrays or all be objects.`, EXIT.file);
  };
  const specs: { name: string; rows: unknown[][] }[] = Array.isArray(data)
    ? [{ name: sheetNameFrom(name), rows: asRows(data) }]
    : data && typeof data === 'object' && Array.isArray((data as { sheets?: unknown }).sheets)
      ? ((data as { sheets: { name?: unknown; rows?: unknown; data?: unknown }[] }).sheets).map((s, i) => ({ name: typeof s?.name === 'string' && s.name ? s.name.slice(0, 31) : `Sheet${i + 1}`, rows: asRows(s?.rows ?? s?.data ?? []) }))
      : (() => { throw new UserError(`${name}: expected { "sheets": [...] } or an array of rows.`, EXIT.file); })();
  const wb = newWorkbook();
  wb.sheets = [];
  const used = new Set<string>();
  for (const spec of specs) {
    let n = spec.name.replace(/[\\/?*[\]:]/g, '_') || 'Sheet';
    for (let i = 2; used.has(n.toLowerCase()); i++) n = `${spec.name.slice(0, 27)} (${i})`;
    used.add(n.toLowerCase());
    const sheet = newSheet(n);
    spec.rows.forEach((r, i) => (Array.isArray(r) ? r : [r]).forEach((v, j) => {
      const c = jsonScalar(v);
      if (c.v !== null || c.f) sheet.cells.set(cellKey(i, j), c.f ? { v: null, f: c.f } : { v: c.v });
    }));
    wb.sheets.push(sheet);
  }
  if (!wb.sheets.length) wb.sheets.push(newSheet('Sheet1'));
  new Engine(wb).recalc(true);
  return wb;
};

export interface ReadOptions { password?: string }

export const readWorkbook = async (name: string, data: Uint8Array, ro: ReadOptions = {}): Promise<Workbook> => {
  if (data.length > MAX_FILE_BYTES) throw new UserError(`${name} is too large to open (limit ${MAX_FILE_BYTES / 1024 / 1024} MB).`, EXIT.file);
  const kind = kindOf(name);
  const isZip = data[0] === 0x50 && data[1] === 0x4b;
  const isCfb = data[0] === 0xd0 && data[1] === 0xcf && data[2] === 0x11 && data[3] === 0xe0;
  try {
    if (kind === 'csv' || kind === 'tsv') {
      const { rows, formats } = readCsv(new TextDecoder().decode(data), kind === 'tsv' ? { delimiter: '\t' } : {});
      return fromRows(name, rows, formats);
    }
    if (kind === 'json') return workbookFromJson(name, new TextDecoder().decode(data));
    if (isCfb) {
      const enc = await encryption();
      if (enc?.isEncryptedPackage(data)) {
        if (!ro.password) throw new UserError(`${name} is protected with a password. Give it with --password or CLERDOGRID_PASSWORD.`, EXIT.file);
        let plain: Uint8Array;
        try { plain = enc.decryptWorkbook(data, ro.password); } catch (e) {
          throw new UserError(`${name}: ${(e as { code?: string }).code === 'password' ? 'that password is not right.' : e instanceof Error ? e.message : String(e)}`, EXIT.file);
        }
        return readXlsx(plain);
      }
    }
    if (kind === 'ods' || kind === 'xls' || kind === 'html' || (isCfb && kind !== 'xlsx')) {
      const k: FileKind = isCfb ? 'xls' : kind!;
      const reader = await findExport(k, 'read');
      if (!reader) throw new UserError(`This build can't read .${k} files.`, EXIT.file);
      try {
        return (await reader(k === 'html' ? new TextDecoder().decode(data) : data, { password: ro.password })) as Workbook;
      } catch (e) {
        if (e instanceof Error && 'wrong' in e) throw new UserError(`${name}: ${e.message}${ro.password ? '' : ' Give it with --password or CLERDOGRID_PASSWORD.'}`, EXIT.file);
        throw e;
      }
    }
    if (kind === 'pdf') throw new UserError("PDF files can't be opened as workbooks.", EXIT.file);
    if (isCfb) throw new UserError(`${name} is an older binary workbook or is password-protected.`, EXIT.file);
    if (!isZip) throw new UserError(`${name} isn't a workbook ClerdoGrid can read.`, EXIT.file);
    return readXlsx(data);
  } catch (e) {
    if (e instanceof UserError) throw e;
    throw new UserError(`${name} could not be read: ${e instanceof Error ? e.message : String(e)}`, EXIT.file);
  }
};

// ---------- Writing ----------

export interface ExportOptions {
  sheet?: number;
  range?: CellRange;
  /** Write raw values instead of the text each cell shows. */
  raw?: boolean;
  /** Write formulas ("=...") instead of their results. */
  formulas?: boolean;
  /** Encrypt an .xlsx file with this password. */
  encryptWith?: string;
}

export const scalarText = (v: Scalar): string => (v === null ? '' : v instanceof CellError ? v.code : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v));

/** What a cell shows, using its number format. */
export const displayText = (wb: Workbook, engine: Engine, sheet: number, row: number, col: number): string => {
  const v = engine.getValue(sheet, row, col);
  if (v === null) return '';
  if (v instanceof CellError) return v.code;
  const s = wb.sheets[sheet]!.cells.get(cellKey(row, col))?.s;
  const fmt = wb.styles[s ?? 0]?.numFmt ?? 'General';
  return formatValue(v, fmt, { date1904: wb.date1904 }).text;
};

export const usedRange = (engine: Engine, sheet: number): CellRange | null => {
  const b = engine.usedBounds(sheet);
  return b.rows < 0 || b.cols < 0 ? null : { r1: 0, c1: 0, r2: b.rows, c2: b.cols };
};

const grid = <T>(range: CellRange | null, cell: (r: number, c: number) => T): T[][] => {
  if (!range) return [];
  const out: T[][] = [];
  for (let r = range.r1; r <= range.r2; r++) {
    const line: T[] = [];
    for (let c = range.c1; c <= range.c2; c++) line.push(cell(r, c));
    out.push(line);
  }
  return out;
};

export const sheetText = (wb: Workbook, engine: Engine, o: ExportOptions = {}): string[][] => {
  const sheet = o.sheet ?? wb.activeSheet;
  return grid(o.range ?? usedRange(engine, sheet), (r, c) => {
    const f = o.formulas ? wb.sheets[sheet]!.cells.get(cellKey(r, c))?.f : undefined;
    if (f) return `=${f}`;
    return o.raw ? scalarText(engine.getValue(sheet, r, c)) : displayText(wb, engine, sheet, r, c);
  });
};

export const jsonValue = (v: Scalar): string | number | boolean | null => (v instanceof CellError ? v.code : v);

export const toJson = (wb: Workbook, engine: Engine, o: ExportOptions = {}): string => {
  const sheets = o.sheet !== undefined ? [o.sheet] : wb.sheets.map((_, i) => i);
  return JSON.stringify({
    sheets: sheets.map((i) => {
      const range = (o.sheet !== undefined ? o.range : undefined) ?? usedRange(engine, i);
      return {
        name: wb.sheets[i]!.name,
        range: range ? rangeAddr(range) : null,
        rows: grid(range, (r, c) => {
          const f = o.formulas ? wb.sheets[i]!.cells.get(cellKey(r, c))?.f : undefined;
          return f ? `=${f}` : jsonValue(engine.getValue(i, r, c));
        }),
      };
    }),
  }, null, 2) + '\n';
};

/** Writes a workbook in the format given by the file name. CSV and TSV hold one sheet. */
export const writeWorkbook = async (wb: Workbook, name: string, o: ExportOptions = {}): Promise<Uint8Array> => {
  const kind = kindOf(name);
  if (!kind) throw new UserError(`Can't tell the format from "${name}". Use .xlsx, .csv, .tsv or .json.`, EXIT.usage);
  if (o.encryptWith && kind !== 'xlsx') throw new UserError('Only .xlsx files can be saved with a password.', EXIT.usage);
  const engine = new Engine(wb);
  engine.recalc(true);
  switch (kind) {
    case 'xlsx': {
      const zip = writeXlsx(wb);
      if (!o.encryptWith) return zip;
      const enc = await encryption();
      if (!enc) throw new UserError("This build can't encrypt workbooks.", EXIT.usage);
      return enc.encryptWorkbook(zip, o.encryptWith);
    }
    case 'csv': return new TextEncoder().encode(writeCsv(sheetText(wb, engine, o)));
    case 'tsv': return new TextEncoder().encode(sheetText(wb, engine, o).map((r) => r.map((f) => f.replace(/[\t\r\n]+/g, ' ')).join('\t')).join('\n') + '\n');
    case 'json': return new TextEncoder().encode(toJson(wb, engine, o));
    default: {
      const writer = await findExport(kind, 'write');
      if (!writer) throw new UserError(`This build can't write .${kind} files.`, EXIT.usage);
      return toBytes(await writer(wb, { sheet: o.sheet }));
    }
  }
};
