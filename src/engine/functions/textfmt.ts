import { Err, isError, type FunctionDef, type Scalar } from '../types';
import { parseNumber, toText } from '../coerce';
import { formatValue, isDateFormat } from '../../formats/numfmt';

export const defs: FunctionDef[] = [
  {
    name: 'TEXT', min: 2, max: 2, scalar: true,
    fn: ([value, fmt], ctx) => {
      const code = toText(fmt as Scalar);
      if (isError(code)) return code;
      let v = value as Scalar;
      if (isError(v)) return v;
      if (v === null) v = 0;
      if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
      if (typeof v === 'string') {
        // Numeric-looking text is formatted as a number, as usual; other text uses the text section.
        const n = parseNumber(v, ctx.date1904);
        if (n !== null) v = n;
      }
      // TEXT ignores fill characters and colours; "*x" contributes nothing.
      if (typeof v === 'number' && isDateFormat(code) && (v < 0 || v >= 2958466)) return Err.value;
      return formatValue(v, code, { date1904: ctx.date1904 }).text;
    },
  },
];
