import type { Scalar } from './types';
import type { Workbook } from '../model/workbook';

/** Looks up a value in a PivotTable for GETPIVOTDATA. Set by the pivot feature; null means no tables. */
export type PivotResolver = (wb: Workbook, sheet: number, row: number, col: number, dataField: string, pairs: [string, Scalar][]) => Scalar;

let resolver: PivotResolver | null = null;
export const setPivotResolver = (r: PivotResolver | null): void => { resolver = r; };
export const pivotResolver = (): PivotResolver | null => resolver;
