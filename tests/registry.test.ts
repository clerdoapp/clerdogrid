import { describe, expect, it } from 'vitest';
import { registerAllFunctions } from '../src/engine/functions';
import { functionNames, SPECIAL_FORMS } from '../src/engine/registry';
import { CATALOG } from '../src/engine/catalog';

describe('function registry', () => {
  it('implements every function in the catalog except connected-service ones', () => {
    registerAllFunctions();
    const have = new Set([...functionNames(), ...SPECIAL_FORMS]);
    const online = new Set(['WEBSERVICE', 'FILTERXML', 'STOCKHISTORY', 'FIELDVALUE', 'IMAGE', 'COPILOT', 'IMPORTCSV', 'IMPORTTEXT', 'CALL', 'REGISTER.ID', 'EUROCONVERT', 'RTD']);
    const missing = CATALOG.map((f) => f.name).filter((n) => !have.has(n) && !online.has(n) && !n.startsWith('CUBE'));
    console.log('registered', have.size, 'missing', missing.join(' '));
    expect(missing).toEqual([]);
  });
});
