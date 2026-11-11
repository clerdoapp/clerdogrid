import { describe, expect, it } from 'vitest';
import { fromR1C1, toR1C1 } from '../src/engine/r1c1';

describe('A1 to R1C1', () => {
  const at = (f: string, r = 4, c = 3) => toR1C1(f, r, c); // D5
  it('converts relative, absolute and mixed cells', () => {
    expect(at('D5')).toBe('RC');
    expect(at('A1')).toBe('R[-4]C[-3]');
    expect(at('$A$1')).toBe('R1C1');
    expect(at('$A1+A$1')).toBe('R[-4]C1+R1C[-3]');
    expect(at('SUM(E6:F10)')).toBe('SUM(R[1]C[1]:R[5]C[2])');
  });
  it('converts whole rows and columns', () => {
    expect(at('SUM(A:C)')).toBe('SUM(C[-3]:C[-1])');
    expect(at('SUM($A:$A)')).toBe('SUM(C1)');
    expect(at('SUM(5:5)')).toBe('SUM(R)');
    expect(at('SUM($1:$3)')).toBe('SUM(R1:R3)');
  });
  it('leaves function names, strings, names and sheets alone', () => {
    expect(at('LOG10(A1)&"B2"')).toBe('LOG10(R[-4]C[-3])&"B2"');
    expect(at("'Q1 B2'!A1+Sheet2!$B$2")).toBe("'Q1 B2'!R[-4]C[-3]+Sheet2!R2C2");
    expect(at('Tax_2023*Table1[Amount]+ATAN2(1,2)')).toBe('Tax_2023*Table1[Amount]+ATAN2(1,2)');
    expect(at('{1,2;3,4}+A1')).toBe('{1,2;3,4}+R[-4]C[-3]');
    expect(at('A1#')).toBe('R[-4]C[-3]#');
    expect(at('1.5E+3+A2')).toBe('1.5E+3+R[-3]C[-3]');
  });
});

describe('R1C1 to A1', () => {
  const at = (f: string, r = 4, c = 3) => fromR1C1(f, r, c);
  it('converts back', () => {
    expect(at('RC')).toBe('D5');
    expect(at('R[-4]C[-3]')).toBe('A1');
    expect(at('R1C1')).toBe('$A$1');
    expect(at('R[-4]C1+R1C[-3]')).toBe('$A1+A$1');
    expect(at('SUM(R[1]C[1]:R[5]C[2])')).toBe('SUM(E6:F10)');
    expect(at('SUM(C[-3]:C[-1])')).toBe('SUM(A:C)');
    expect(at('SUM(C1)')).toBe('SUM($A:$A)');
    expect(at('SUM(R)')).toBe('SUM(5:5)');
    expect(at('rc[1]')).toBe('E5');
  });
  it('gives #REF! for references off the sheet', () => {
    expect(at('R[-10]C')).toBe('#REF!');
    expect(at('R[-1]C[-9]+1')).toBe('#REF!+1');
  });
  it('keeps names and functions', () => {
    expect(at('ROUND(RC[-1],2)+COUNT(R1C1:R3C3)')).toBe('ROUND(C5,2)+COUNT($A$1:$C$3)');
    expect(at('Rate*RC')).toBe('Rate*D5');
    expect(at('"RC"&RC')).toBe('"RC"&D5');
  });
  it('round-trips', () => {
    for (const f of ['SUM($A$1:B7)*C$3-INDEX(Sheet2!A:A,MATCH(D5,Sheet2!$B:$B,0))', 'IF(A1>0,"R1C1",F6)', "'My Sheet'!$Z$100+AB12"]) {
      expect(fromR1C1(toR1C1(f, 9, 9), 9, 9)).toBe(f);
    }
  });
});
