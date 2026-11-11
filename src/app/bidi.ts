// Text direction helpers shared by the grid and the page renderer.
import type { Scalar } from '../engine/types';

/** First strong direction of a string (Unicode bidi rules P2/P3): 'rtl', 'ltr', or null when it has none. */
export const firstStrong = (s: string): 'rtl' | 'ltr' | null => {
  const m = /[A-Za-z\u00C0-\u024F\u0370-\u052F\u0900-\u0DFF\u0E00-\u0FFF\u1100-\u11FF\u3040-\u9FFF\uAC00-\uD7AF]|[\u0590-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFF]/.exec(s);
  if (!m) return null;
  const c = m[0].charCodeAt(0);
  return (c >= 0x590 && c <= 0x8ff) || (c >= 0xfb1d && c <= 0xfdff) || c >= 0xfe70 ? 'rtl' : 'ltr';
};

/** Where General alignment puts a value on screen: numbers to the trailing edge of the sheet, text by its own first strong direction. */
export const generalAlign = (v: Scalar, text: string, sheetRtl: boolean, cellRtl: boolean): 'left' | 'right' => {
  if (typeof v === 'number') return sheetRtl ? 'left' : 'right';
  if (cellRtl) return 'right';
  const d = firstStrong(text);
  return d === 'rtl' ? 'right' : d === 'ltr' ? 'left' : sheetRtl ? 'right' : 'left';
};
