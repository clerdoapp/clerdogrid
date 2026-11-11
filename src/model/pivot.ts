// PivotTable, slicer and timeline model. Shapes follow ECMA-376 Part 1 §18.10 (pivot tables) so
// reading and writing files stays close to one-to-one.
import type { Scalar } from '../engine/types';
import type { CellRange } from './workbook';

export type PivotAgg =
  | 'sum' | 'count' | 'average' | 'max' | 'min' | 'product' | 'countNums'
  | 'stdDev' | 'stdDevp' | 'var' | 'varp' | 'distinctCount';

export type ShowAs =
  | 'normal' | 'percentOfTotal' | 'percentOfCol' | 'percentOfRow' | 'percentOfParentRow' | 'percentOfParentCol'
  | 'percentOfParent' | 'percent' | 'difference' | 'percentDiff' | 'runTotal' | 'percentOfRunningTotal'
  | 'rankAscending' | 'rankDescending' | 'index';

export type DateGroupBy = 'years' | 'quarters' | 'months' | 'days' | 'hours' | 'minutes' | 'seconds';

/** Grouping of a cache field's values (§18.10.1.36 fieldGroup). */
export interface FieldGroup {
  /** Field whose source values are grouped (the field itself for a grouped database field). */
  base: number;
  /** Next coarser grouping field, if any (par). */
  parent?: number;
  /** Range grouping: numbers in steps, or dates by a calendar unit. */
  range?: { by: 'range' | DateGroupBy; start: number; end: number; autoStart: boolean; autoEnd: boolean; interval: number; dates: boolean };
  /** Manual groups: for each item of the base field, the index of its group in `items`. */
  discrete?: number[];
  /** Group item captions in display order. */
  items: string[];
}

export interface CacheField {
  name: string;
  numFmt?: string;
  /** Unique source values (shared items). Dates are serial numbers with `dates` set. */
  items: Scalar[];
  dates?: boolean;
  /** Calculated field formula (without "="), e.g. "Sales*0.1". */
  formula?: string;
  group?: FieldGroup;
  /** False for fields that exist only through grouping or a formula. */
  database?: boolean;
}

export type PivotSource =
  | { kind: 'range'; sheet: string; range: CellRange }
  | { kind: 'name'; name: string };

export interface PivotCache {
  id: number;
  source: PivotSource;
  fields: CacheField[];
  /** One row per source record, one value per database field (in field order). */
  records: Scalar[][];
  refreshOnLoad?: boolean;
  refreshedBy?: string;
  refreshedDate?: number;
  /** Items no longer in the data are dropped on refresh (missingItemsLimit 0). */
  dropMissing?: boolean;
}

export type Subtotal = 'default' | 'sum' | 'count' | 'average' | 'max' | 'min' | 'product' | 'countNums' | 'stdDev' | 'stdDevp' | 'var' | 'varp';

export interface PivotItem {
  /** Index into the field's items (shared items, or group items when grouped). -1 for a calculated item. */
  x: number;
  hidden?: boolean;
  /** Details hidden (collapsed). */
  collapsed?: boolean;
  /** Custom caption. */
  name?: string;
  /** Calculated item: a formula over other items of the same field, e.g. "East+West". */
  formula?: string;
}

export type LabelOp =
  | 'equal' | 'notEqual' | 'beginsWith' | 'notBeginsWith' | 'endsWith' | 'notEndsWith' | 'contains' | 'notContains'
  | 'greaterThan' | 'greaterThanOrEqual' | 'lessThan' | 'lessThanOrEqual' | 'between' | 'notBetween';
export type ValueOp = 'equal' | 'notEqual' | 'greaterThan' | 'greaterThanOrEqual' | 'lessThan' | 'lessThanOrEqual' | 'between' | 'notBetween';

export type PivotFilter =
  | { kind: 'label'; op: LabelOp; value1: string; value2?: string }
  | { kind: 'value'; op: ValueOp; dataField: number; value1: number; value2?: number }
  | { kind: 'top'; top: boolean; by: 'count' | 'percent' | 'sum'; value: number; dataField: number }
  /** Date range from a timeline: source dates from start up to the end of the end day. */
  | { kind: 'dates'; start: number; end: number; timeline?: string };

export interface PivotField {
  /** Custom caption for the field. */
  name?: string;
  /** Items in display order (manual sort order). */
  items: PivotItem[];
  sort?: 'ascending' | 'descending';
  /** Sort by this data field's values instead of the labels. */
  sortByData?: number;
  /** [] means no subtotals. Missing means automatic. */
  subtotals?: Subtotal[];
  /** Subtotals above the group (outline and compact forms). Default true. */
  subtotalTop?: boolean;
  /** Field shares the column of the previous field (compact form). Default true. */
  compact?: boolean;
  /** Outline form (false: tabular). Default true. */
  outline?: boolean;
  repeatLabels?: boolean;
  blankRow?: boolean;
  pageBreak?: boolean;
  /** Show items with no data. */
  showAll?: boolean;
  filters?: PivotFilter[];
  numFmt?: string;
  /** Several items can be picked in a report filter. */
  multiPage?: boolean;
  /** Includes new items in a manual filter (default true). */
  includeNew?: boolean;
}

export interface DataField {
  name: string;
  field: number;
  agg: PivotAgg;
  showAs?: ShowAs;
  baseField?: number;
  /** Index into the base field's items, or the previous/next item. */
  baseItem?: number | 'prev' | 'next';
  numFmt?: string;
}

/** Field index used for the Σ Values pseudo field in rows/cols (as in the file format). */
export const VALUES_FIELD = -2;

export type ItemType = 'data' | 'default' | 'sum' | 'countA' | 'avg' | 'max' | 'min' | 'product' | 'count' | 'stdDev' | 'stdDevP' | 'var' | 'varP' | 'grand' | 'blank';
/** One rowItems/colItems entry (§18.10.1.44 i). */
export interface AxisItem { t?: ItemType; r?: number; i?: number; x: number[] }

export interface PivotStyleInfo { name: string; rowHeaders: boolean; colHeaders: boolean; bandedRows: boolean; bandedCols: boolean; lastColumn?: boolean }

export interface PivotTable {
  name: string;
  cacheId: number;
  /** Top-left cell of the table body (below any report filters). */
  anchor: { row: number; col: number };
  /** Area last drawn, with the offsets of the header and data starts (location). */
  location?: { ref: CellRange; firstHeaderRow: number; firstDataRow: number; firstDataCol: number; pageRows?: number };
  /** One per cache field. */
  fields: PivotField[];
  rows: number[];
  cols: number[];
  /** Report filters; `item` is an index into the field's items when one item is picked. */
  pages: { field: number; item?: number }[];
  data: DataField[];
  layout: 'compact' | 'outline' | 'tabular';
  /** Totals column on the right (totals of each row). */
  rowGrand: boolean;
  /** Totals row at the bottom (totals of each column). */
  colGrand: boolean;
  style: PivotStyleInfo;
  emptyText?: string;
  showEmpty?: boolean;
  errorText?: string;
  showError?: boolean;
  preserveFormatting: boolean;
  autoFit: boolean;
  showDrill: boolean;
  showHeaders: boolean;
  rowHeaderCaption?: string;
  colHeaderCaption?: string;
  grandTotalCaption?: string;
  dataCaption?: string;
  /** Report filters laid out over then down, wrapped after this many fields (0: one column). */
  pageWrap?: number;
  pageOverThenDown?: boolean;
  indent?: number;
  /** Totals include items hidden by filters (visualTotals false). */
  includeFilteredInTotals?: boolean;
  /** Last layout written, kept for the file format. */
  rowItems?: AxisItem[];
  colItems?: AxisItem[];
  /** Defer layout updates from the field list. */
  deferLayout?: boolean;
  /** Formatting the user applied over the table, kept when it's refreshed (pivot area key to style). */
  formats?: { key: string; style: number }[];
}

// ---------- Slicers and timelines ----------

export type SlicerSource =
  | { kind: 'pivot'; cacheId: number; field: number; pivots: { sheet: string; name: string }[] }
  | { kind: 'table'; table: string; column: number };

export interface SlicerCache {
  name: string;
  source: SlicerSource;
  /** Hidden item keys (pivot: cache item indexes as strings; table: display texts). Empty means no filter. */
  hidden: string[];
  sort: 'ascending' | 'descending' | 'none';
  hideNoData?: boolean;
  showNoDataLast?: boolean;
  /** Timeline: a date range filter instead of item buttons. */
  timeline?: { start?: number; end?: number; level: 'years' | 'quarters' | 'months' | 'days' };
}

export interface Slicer {
  name: string;
  cache: string;
  caption: string;
  kind: 'slicer' | 'timeline';
  anchor: { row: number; col: number; dx: number; dy: number; width: number; height: number };
  columns: number;
  /** Button height in pixels. */
  buttonHeight: number;
  style: string;
  showCaption: boolean;
  /** Timeline options. */
  level?: 'years' | 'quarters' | 'months' | 'days';
  showSelectionLabel?: boolean;
  showTimeLevel?: boolean;
  showScrollbar?: boolean;
}

export const DEFAULT_PIVOT_STYLE: PivotStyleInfo = { name: 'PivotStyleLight16', rowHeaders: true, colHeaders: true, bandedRows: false, bandedCols: false };
