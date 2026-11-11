export type ErrorCode =
  | '#NULL!' | '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#NUM!' | '#N/A'
  | '#GETTING_DATA' | '#SPILL!' | '#CALC!' | '#FIELD!' | '#BLOCKED!' | '#CONNECT!' | '#BUSY!' | '#UNKNOWN!';

export const ERROR_CODES: readonly ErrorCode[] = [
  '#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA',
  '#SPILL!', '#CALC!', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#BUSY!', '#UNKNOWN!',
];

export class CellError {
  constructor(readonly code: ErrorCode, readonly detail?: string) {}
  toString(): string { return this.code; }
}

export const Err = {
  null: new CellError('#NULL!'),
  div0: new CellError('#DIV/0!'),
  value: new CellError('#VALUE!'),
  ref: new CellError('#REF!'),
  name: new CellError('#NAME?'),
  num: new CellError('#NUM!'),
  na: new CellError('#N/A'),
  spill: new CellError('#SPILL!'),
  calc: new CellError('#CALC!'),
  field: new CellError('#FIELD!'),
  blocked: new CellError('#BLOCKED!'),
} as const;

/** One cell value. `null` is a blank cell (not the same as an empty string). */
export type Scalar = number | string | boolean | CellError | null;

/** Rectangular, row-major, at least 1×1. */
export type Matrix = Scalar[][];

/** A rectangular block of cells. Rows/cols are 0-based and inclusive. */
export interface RangeRef {
  readonly kind: 'ref';
  readonly sheet: number;
  readonly r1: number;
  readonly c1: number;
  readonly r2: number;
  readonly c2: number;
}

/** Several areas, from the union operator: =SUM((A1:A3,C1:C3)). */
export interface RefUnion {
  readonly kind: 'refs';
  readonly areas: readonly RangeRef[];
}

export interface LambdaValue {
  readonly kind: 'lambda';
  readonly params: readonly string[];
  readonly invoke: (args: Value[]) => Value;
  /** Name shown in GROUPBY/PIVOTBY headers (the function name for SUM passed as a value). */
  readonly label?: string;
  /** Several functions stacked with HSTACK/VSTACK, for GROUPBY/PIVOTBY. */
  readonly list?: readonly LambdaValue[];
  readonly vertical?: boolean;
}

export type Value = Scalar | Matrix | RangeRef | RefUnion | LambdaValue;

export const MAX_ROWS = 1_048_576;
export const MAX_COLS = 16_384;

export const isError = (v: unknown): v is CellError => v instanceof CellError;
export const isMatrix = (v: Value): v is Matrix => Array.isArray(v);
export const isRef = (v: Value): v is RangeRef => typeof v === 'object' && v !== null && !Array.isArray(v) && (v as RangeRef).kind === 'ref';
export const isRefs = (v: Value): v is RefUnion => typeof v === 'object' && v !== null && !Array.isArray(v) && (v as RefUnion).kind === 'refs';
export const isLambda = (v: Value): v is LambdaValue => typeof v === 'object' && v !== null && !Array.isArray(v) && (v as LambdaValue).kind === 'lambda';
export const isScalar = (v: Value): v is Scalar => v === null || typeof v !== 'object' || v instanceof CellError;

export const ref = (sheet: number, r1: number, c1: number, r2 = r1, c2 = c1): RangeRef => ({
  kind: 'ref', sheet, r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2),
});

/** What a worksheet function can ask of the workbook while it runs. */
export interface FnContext {
  /** The cell being calculated. */
  readonly sheet: number;
  readonly row: number;
  readonly col: number;
  /** Values of a range, trimmed to the sheet's used area (whole-column references stay small). */
  values(r: RangeRef): Matrix;
  cell(sheet: number, row: number, col: number): Scalar;
  /** Last used row/col of a sheet (0-based, -1 when empty). */
  usedBounds(sheet: number): { rows: number; cols: number };
  formula(sheet: number, row: number, col: number): string | null;
  sheetIndex(name: string): number;
  sheetName(index: number): string;
  sheetCount(): number;
  /** Calls a LAMBDA value. */
  call(fn: LambdaValue, args: Value[]): Value;
  /** True when the workbook uses the 1904 date system. */
  readonly date1904: boolean;
  /** Hidden rows (manual or filtered), for SUBTOTAL/AGGREGATE. */
  rowHidden(sheet: number, row: number): boolean;
  /** Rows hidden by a filter (SUBTOTAL 1-11 skips these but not manually hidden rows). */
  rowFiltered?(sheet: number, row: number): boolean;
  /** Whether a cell holds a SUBTOTAL or AGGREGATE formula, which SUBTOTAL skips. */
  isSubtotalCell(sheet: number, row: number, col: number): boolean;
  /** Number format code of a cell, for CELL("format"). */
  numberFormat(sheet: number, row: number, col: number): string;
  /** GETPIVOTDATA: a value from the PivotTable that covers a cell. */
  pivotData?(sheet: number, row: number, col: number, dataField: string, pairs: [string, Scalar][]): Scalar;
}

export interface FunctionDef {
  /** Upper-case, with any prefix (e.g. "_XLFN." is stripped before lookup). */
  readonly name: string;
  readonly min: number;
  readonly max: number;
  readonly volatile?: boolean;
  /**
   * Arguments that take one value each. When an array or multi-cell range is passed in one of
   * these positions, the engine calls `fn` once per element and returns an array (dynamic arrays).
   * `true` lifts every argument. Single-cell references are always passed in as their value here.
   */
  readonly scalar?: boolean | readonly number[];
  /** Arguments passed unevaluated as thunks (IF, IFERROR, CHOOSE, LET...). Handled by the evaluator. */
  readonly lazy?: boolean;
  readonly fn: (args: Value[], ctx: FnContext) => Value;
}
