// Formula text transforms needed when reading/writing SpreadsheetML.
import { colIndex, colName } from '../model/workbook';
import { MAX_COLS, MAX_ROWS } from '../engine/types';

/** Applies `f` to each stretch of the formula outside string literals, quoted sheet names and [structured refs]. */
const mapOutside = (formula: string, f: (code: string) => string): string => {
  let out = '';
  let buf = '';
  for (let i = 0; i < formula.length; i++) {
    const c = formula[i]!;
    if (c === '"' || c === "'") {
      let j = i + 1;
      for (; j < formula.length; j++) {
        if (formula[j] === c) {
          if (formula[j + 1] === c) j++;
          else break;
        }
      }
      // a quoted sheet name stays attached to the following reference: let `f` see it as a placeholder
      out += f(buf);
      buf = '';
      out += formula.slice(i, j + 1);
      i = j;
    } else if (c === '[') {
      let depth = 0, j = i;
      for (; j < formula.length; j++) {
        if (formula[j] === '[') depth++;
        else if (formula[j] === ']' && --depth === 0) break;
      }
      out += f(buf);
      buf = '';
      out += formula.slice(i, j + 1);
      i = j;
    } else buf += c;
  }
  return out + f(buf);
};

export const stripFuturePrefixes = (formula: string): string =>
  formula.indexOf('_xl') < 0 ? formula : mapOutside(formula, (s) => s.replace(/_xl(?:fn|ws|pm)\./gi, ''));

// Functions added after the first edition of the format, which files store with the "_xlfn." prefix.
const FUTURE = new Set(`ACOT ACOTH AGGREGATE ANCHORARRAY ARABIC ARRAYTOTEXT BASE BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE
BINOM.INV BITAND BITLSHIFT BITOR BITRSHIFT BITXOR BYCOL BYROW CEILING.MATH CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT
CHISQ.INV CHISQ.INV.RT CHISQ.TEST CHOOSECOLS CHOOSEROWS COMBINA CONCAT CONFIDENCE.NORM CONFIDENCE.T COT COTH
COVARIANCE.P COVARIANCE.S CSC CSCH DAYS DECIMAL DROP ECMA.CEILING ENCODEURL ERF.PRECISE ERFC.PRECISE EXPAND EXPON.DIST
F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FIELDVALUE FILTERXML FLOOR.MATH FLOOR.PRECISE FORECAST.ETS FORECAST.ETS.CONFINT
FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FORECAST.LINEAR FORMULATEXT GAMMA GAMMA.DIST GAMMA.INV GAMMALN.PRECISE GAUSS
GROUPBY HSTACK HYPGEOM.DIST IFNA IFS IMAGE IMCOSH IMCOT IMCSC IMCSCH IMSEC IMSECH IMSINH IMTAN ISFORMULA ISO.CEILING
ISOMITTED ISOWEEKNUM LAMBDA LET LOGNORM.DIST LOGNORM.INV MAKEARRAY MAP MAXIFS MINIFS MODE.MULT MODE.SNGL MUNIT
NEGBINOM.DIST NETWORKDAYS.INTL NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV NUMBERVALUE PDURATION PERCENTILE.EXC
PERCENTILE.INC PERCENTOF PERCENTRANK.EXC PERCENTRANK.INC PERMUTATIONA PHI PIVOTBY POISSON.DIST QUARTILE.EXC QUARTILE.INC
QUERYSTRING RANDARRAY RANK.AVG RANK.EQ REDUCE REGEXEXTRACT REGEXREPLACE REGEXTEST RRI SCAN SEC SECH SEQUENCE SHEET SHEETS
SINGLE SKEW.P SORTBY STDEV.P STDEV.S STOCKHISTORY SWITCH T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST TAKE TEXTAFTER
TEXTBEFORE TEXTJOIN TEXTSPLIT TOCOL TOROW TRIMRANGE UNICHAR UNICODE UNIQUE VALUETOTEXT VAR.P VAR.S VSTACK WEBSERVICE
WEIBULL.DIST WORKDAY.INTL WRAPCOLS WRAPROWS XLOOKUP XMATCH XOR Z.TEST`.split(/\s+/));
const WORKSHEET_FUTURE = new Set(['FILTER', 'SORT']);

const FN_RE = /(^|[^A-Za-z0-9_.!$])([A-Za-z][A-Za-z0-9_.]*)(?=\s*\()/g;

/** Positions of the arguments of the call whose "(" is at `open`: [argStart, argEnd) pairs and the closing index. */
const callArgs = (s: string, open: number): { args: [number, number][]; close: number } => {
  const args: [number, number][] = [];
  let depth = 0;
  let start = open + 1;
  let i = open;
  for (; i < s.length; i++) {
    const c = s[i];
    if (c === '"') { i = s.indexOf('"', i + 1); if (i < 0) break; continue; }
    if (c === "'") { i = s.indexOf("'", i + 1); if (i < 0) break; continue; }
    if (c === '(' || c === '{' || c === '[') depth++;
    else if (c === ')' || c === '}' || c === ']') {
      depth--;
      if (depth === 0) { args.push([start, i]); break; }
    } else if (c === ',' && depth === 1) { args.push([start, i]); start = i + 1; }
  }
  return { args, close: i < 0 ? s.length : i };
};

/** Adds "_xlfn." / "_xlfn._xlws." prefixes, and "_xlpm." on LET/LAMBDA parameter names, as the file format stores them. */
export const addFuturePrefixes = (formula: string): string => {
  const params = new Set<string>();
  const code = formula;
  const scan = new RegExp(FN_RE.source, 'g');
  let m: RegExpExecArray | null;
  while ((m = scan.exec(code))) {
    const name = m[2]!.toUpperCase().replace(/^_XLFN\./, '');
    if (name !== 'LET' && name !== 'LAMBDA') continue;
    const { args } = callArgs(code, code.indexOf('(', m.index + m[0].length));
    const names = name === 'LET' ? args.filter((_, i) => i % 2 === 0 && i < args.length - 1) : args.slice(0, -1);
    for (const [a, b] of names) {
      const p = code.slice(a, b).trim().replace(/^_xlpm\./i, '');
      if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(p)) params.add(p.toUpperCase());
    }
  }
  return mapOutside(formula, (s) => {
    let r = s.replace(FN_RE, (all, pre: string, name: string) => {
      const u = name.toUpperCase();
      const rest = all.slice(pre.length + name.length);
      if (WORKSHEET_FUTURE.has(u)) return `${pre}_xlfn._xlws.${name}${rest}`;
      if (FUTURE.has(u)) return `${pre}_xlfn.${name}${rest}`;
      return all;
    });
    if (params.size) {
      r = r.replace(/(^|[^A-Za-z0-9_.!$:\\])([A-Za-z_][A-Za-z0-9_.]*)(?![A-Za-z0-9_.(!:])/g, (all, pre: string, id: string) =>
        params.has(id.toUpperCase()) ? `${pre}_xlpm.${id}` : all);
    }
    return r;
  });
};

// ---------------------------------------------------------------- reference shifting

const CELL = /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/y;

/**
 * Moves the relative parts of every A1 reference by (dr, dc), as when a shared formula is
 * expanded or a formula is copied. Absolute parts ($) stay. References pushed off the sheet become #REF!.
 */
export const shiftFormula = (formula: string, dr: number, dc: number): string => {
  if (!dr && !dc) return formula;
  return mapOutside(formula, (s) => shiftCode(s, dr, dc));
};

const isIdent = (c: string | undefined): boolean => !!c && /[A-Za-z0-9_.\\]/.test(c);

const shiftCol = (abs: string, col: number, dc: number): string | null => {
  const c = abs ? col : col + dc;
  return c < 0 || c >= MAX_COLS ? null : abs + colName(c);
};
const shiftRow = (abs: string, row: number, dr: number): string | null => {
  const r = abs ? row : row + dr;
  return r < 0 || r >= MAX_ROWS ? null : abs + (r + 1);
};

const shiftCode = (s: string, dr: number, dc: number): string => {
  let out = '';
  let i = 0;
  while (i < s.length) {
    const prev = s[i - 1];
    if (!isIdent(prev) || prev === undefined) {
      // whole columns A:C / whole rows 1:3
      const colRange = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_(!])/.exec(s.slice(i));
      if (colRange && colIndex(colRange[2]!) < MAX_COLS && colIndex(colRange[4]!) < MAX_COLS) {
        const a = shiftCol(colRange[1]!, colIndex(colRange[2]!), dc);
        const b = shiftCol(colRange[3]!, colIndex(colRange[4]!), dc);
        out += a && b ? `${a}:${b}` : '#REF!';
        i += colRange[0].length;
        continue;
      }
      const rowRange = /^(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![A-Za-z0-9_.(])/.exec(s.slice(i));
      if (rowRange) {
        const a = shiftRow(rowRange[1]!, +rowRange[2]! - 1, dr);
        const b = shiftRow(rowRange[3]!, +rowRange[4]! - 1, dr);
        out += a && b ? `${a}:${b}` : '#REF!';
        i += rowRange[0].length;
        continue;
      }
      CELL.lastIndex = i;
      const m = CELL.exec(s);
      if (m && !isIdent(s[i + m[0].length]) && s[i + m[0].length] !== '(' && s[i + m[0].length] !== '!') {
        const col = colIndex(m[2]!);
        if (col < MAX_COLS) {
          const c = shiftCol(m[1]!, col, dc);
          const r = shiftRow(m[3]!, +m[4]! - 1, dr);
          out += c && r ? c + r : '#REF!';
          i += m[0].length;
          continue;
        }
      }
      // skip a whole identifier (function or defined name) so its tail isn't read as a reference
      const id = /^[A-Za-z_\\][A-Za-z0-9_.\\]*/.exec(s.slice(i));
      if (id) { out += id[0]; i += id[0].length; continue; }
      const numTok = /^\d+(\.\d*)?([eE][+-]?\d+)?/.exec(s.slice(i));
      if (numTok) { out += numTok[0]; i += numTok[0].length; continue; }
    }
    out += s[i];
    i++;
  }
  return out;
};

/** Sheet name as it must appear in a reference. */
export const quoteSheet = (name: string): string =>
  /^[A-Za-z_À-￿][A-Za-z0-9_.À-￿]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) && !/^(R\d*C\d*|R|C|TRUE|FALSE)$/i.test(name)
    ? name
    : `'${name.replace(/'/g, "''")}'`;
