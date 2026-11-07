import { toText } from '../coerce';
import { CellError, isError, type FunctionDef, type Scalar } from '../types';

const offline = new CellError('#CONNECT!', 'ClerdoGrid works offline and does not fetch data from the internet.');

export const defs: FunctionDef[] = [
  {
    name: 'ENCODEURL', min: 1, max: 1, scalar: true,
    fn: ([v]) => { const t = toText(v as Scalar); return isError(t) ? t : encodeURIComponent(t).replace(/[!'()*]/g, (c) => `%${c.charCodeAt(0).toString(16).toUpperCase()}`); },
  },
  ...['WEBSERVICE', 'STOCKHISTORY', 'IMAGE', 'FIELDVALUE'].map((name): FunctionDef => ({ name, min: 1, max: 255, fn: () => offline })),
  { name: 'FILTERXML', min: 2, max: 2, fn: () => new CellError('#VALUE!') },
];
