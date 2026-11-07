import { describe, expect, it } from 'vitest';
import { CATALOG, catalogByName, searchCatalog, signature } from '../src/engine/catalog';

const sig = (name: string): string => {
  const f = catalogByName(name);
  if (!f) throw new Error(`missing ${name}`);
  return signature(f);
};

describe('function catalog', () => {
  it('has roughly the full spreadsheet function set', () => {
    expect(CATALOG.length).toBeGreaterThan(450);
  });

  it('has no duplicate names', () => {
    const names = CATALOG.map((f) => f.name);
    expect(new Set(names).size).toBe(names.length);
  });

  it('gives every entry a category and short about text', () => {
    for (const f of CATALOG) {
      expect(f.category, f.name).toBeTruthy();
      expect(f.about.length, f.name).toBeGreaterThan(5);
      expect(f.about.length, f.name).toBeLessThanOrEqual(110);
      for (const a of f.args) expect(a.about.length, `${f.name}.${a.name}`).toBeGreaterThan(2);
    }
  });

  it('formats signatures', () => {
    expect(sig('VLOOKUP')).toBe('VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])');
    expect(sig('SUM')).toBe('SUM(number1, [number2], ...)');
    expect(sig('IF')).toBe('IF(logical_test, value_if_true, [value_if_false])');
    expect(sig('PI')).toBe('PI()');
  });

  it('looks up names case-insensitively', () => {
    expect(catalogByName('xlookup')?.name).toBe('XLOOKUP');
    expect(catalogByName('_xlfn.Filter')?.name).toBe('FILTER');
    expect(catalogByName('nope')).toBeUndefined();
  });

  it('ranks search results: prefix, then name contains, then description', () => {
    expect(searchCatalog('xlo')[0]?.name).toBe('XLOOKUP');
    const r = searchCatalog('lookup', 100).map((f) => f.name);
    expect(r[0]).toBe('LOOKUP');
    expect(r.indexOf('VLOOKUP')).toBeGreaterThan(0);
    expect(searchCatalog('su', 3)).toHaveLength(3);
    expect(searchCatalog('su')[0]?.name).toBe('SUM');
    expect(searchCatalog('depreciation').length).toBeGreaterThan(3);
  });

  it('includes well-known functions', () => {
    const names = `SUM AVERAGE COUNT COUNTA COUNTIF COUNTIFS SUMIF SUMIFS AVERAGEIFS MAX MIN IF IFS IFERROR AND OR NOT XOR
      VLOOKUP HLOOKUP XLOOKUP XMATCH INDEX MATCH OFFSET INDIRECT FILTER SORT SORTBY UNIQUE SEQUENCE RANDARRAY LET LAMBDA
      MAP REDUCE SCAN BYROW BYCOL MAKEARRAY ISOMITTED TAKE DROP CHOOSEROWS CHOOSECOLS EXPAND HSTACK VSTACK TOCOL TOROW
      WRAPROWS WRAPCOLS TEXTSPLIT TEXTBEFORE TEXTAFTER VALUETOTEXT ARRAYTOTEXT REGEXTEST REGEXEXTRACT REGEXREPLACE GROUPBY
      PIVOTBY PERCENTOF TRIMRANGE IMAGE WEBSERVICE FILTERXML ENCODEURL CUBEVALUE CUBEMEMBER CONCATENATE STDEV PMT NPV XIRR
      DATE TODAY NETWORKDAYS.INTL TEXT TEXTJOIN LEFT MID SUBSTITUTE ROUND SUMPRODUCT AGGREGATE DSUM BIN2DEC CONVERT ISNUMBER`
      .split(/\s+/).filter(Boolean);
    const missing = names.filter((n) => !catalogByName(n));
    expect(missing).toEqual([]);
  });

  it('places functions in the expected categories', () => {
    expect(catalogByName('WEBSERVICE')?.category).toBe('Web');
    expect(catalogByName('CUBESET')?.category).toBe('Cube');
    expect(catalogByName('STDEV')?.category).toBe('Compatibility');
    expect(catalogByName('DGET')?.category).toBe('Database');
  });
});
