import type { Scalar } from '../engine/types';
import { MAX_COLS } from '../engine/types';
import type { ChartOptionsExtra, ChartSeriesExtra, ChartTypeExtra, SparklineGroup } from './chart';
export type * from './chart';

export const cellKey = (row: number, col: number): number => row * MAX_COLS + col;
export const keyRow = (key: number): number => Math.floor(key / MAX_COLS);
export const keyCol = (key: number): number => key % MAX_COLS;

/** Inclusive, 0-based. */
export interface CellRange { r1: number; c1: number; r2: number; c2: number }

export type BorderStyle = 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'hair' | 'mediumDashed' | 'dashDot' | 'mediumDashDot' | 'dashDotDot' | 'mediumDashDotDot' | 'slantDashDot';
export interface BorderSide { style: BorderStyle; color?: string }

export type HAlign = 'general' | 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
export type VAlign = 'top' | 'center' | 'bottom' | 'justify' | 'distributed';
export type Underline = 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';

/** Colors are "#RRGGBB" (or "#RRGGBBAA"). */
export interface CellStyle {
  font?: { name?: string; size?: number; bold?: boolean; italic?: boolean; underline?: Underline; strike?: boolean; color?: string; vertAlign?: 'superscript' | 'subscript' };
  fill?: { color?: string; pattern?: string; patternColor?: string; gradient?: { angle: number; stops: { pos: number; color: string }[] } };
  border?: { left?: BorderSide; right?: BorderSide; top?: BorderSide; bottom?: BorderSide; diagonalUp?: BorderSide; diagonalDown?: BorderSide };
  align?: { h?: HAlign; v?: VAlign; wrap?: boolean; indent?: number; rotation?: number; shrink?: boolean; rtl?: boolean };
  numFmt?: string;
  protection?: { locked?: boolean; hidden?: boolean };
  /** The cell shows a checkbox for its TRUE/FALSE value. */
  checkbox?: boolean;
}

export interface RichRun { text: string; font?: CellStyle['font'] }

export interface Cell {
  /** The stored or last calculated value. */
  v: Scalar;
  /** Formula without the leading "=", in English with comma separators, e.g. "SUM(A1:A3)". */
  f?: string;
  /** For a dynamic-array formula: the formula spills its result (written as an array formula in .xlsx). */
  dyn?: boolean;
  /** Legacy Ctrl+Shift+Enter array formula covering this range (set on the top-left cell). */
  arrayRange?: CellRange;
  /** Index into Workbook.styles; 0 or missing is the default style. */
  s?: number;
  rich?: RichRun[];
}

export interface ColInfo { width?: number; hidden?: boolean; s?: number; outline?: number; /** The outline group ending next to this column is collapsed. */ collapsed?: boolean }
export interface RowInfo { height?: number; hidden?: boolean; s?: number; outline?: number; /** Hidden by the sheet's filter rather than by hand. */ filtered?: boolean; /** The outline group ending next to this row is collapsed. */ collapsed?: boolean }

export type CfOperator = 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual';
export interface CfValue { type: 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula' | 'autoMin' | 'autoMax'; value?: string }

export type CfRule =
  | { type: 'cellIs'; operator: CfOperator; formulas: string[]; style: CellStyle; stop?: boolean }
  | { type: 'expression'; formula: string; style: CellStyle; stop?: boolean }
  | { type: 'colorScale'; stops: { value: CfValue; color: string }[] }
  | { type: 'dataBar'; min: CfValue; max: CfValue; color: string; gradient?: boolean; showValue?: boolean }
  | { type: 'iconSet'; set: string; values: CfValue[]; reverse?: boolean; showValue?: boolean }
  | { type: 'top10'; rank: number; bottom?: boolean; percent?: boolean; style: CellStyle; stop?: boolean }
  | { type: 'aboveAverage'; below?: boolean; equal?: boolean; stdDev?: number; style: CellStyle; stop?: boolean }
  | { type: 'duplicateValues' | 'uniqueValues'; style: CellStyle; stop?: boolean }
  | { type: 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith'; text: string; style: CellStyle; stop?: boolean }
  | { type: 'containsBlanks' | 'notContainsBlanks' | 'containsErrors' | 'notContainsErrors'; style: CellStyle; stop?: boolean }
  | { type: 'timePeriod'; period: 'today' | 'yesterday' | 'tomorrow' | 'last7Days' | 'thisWeek' | 'lastWeek' | 'nextWeek' | 'thisMonth' | 'lastMonth' | 'nextMonth'; style: CellStyle; stop?: boolean };

export interface ConditionalFormat { ranges: CellRange[]; rules: CfRule[]; priority?: number }

export interface DataValidation {
  ranges: CellRange[];
  type: 'any' | 'whole' | 'decimal' | 'list' | 'date' | 'time' | 'textLength' | 'custom';
  operator?: CfOperator;
  formula1?: string;
  formula2?: string;
  allowBlank?: boolean;
  showDropdown?: boolean;
  prompt?: { title?: string; text: string };
  error?: { style: 'stop' | 'warning' | 'information'; title?: string; text: string };
}

export interface FilterColumn {
  col: number;
  values?: string[];
  blanks?: boolean;
  custom?: { and?: boolean; rules: { operator: CfOperator; value: string }[] };
  top10?: { top: boolean; percent: boolean; value: number };
  color?: string;
}
export interface AutoFilter { range: CellRange; columns: FilterColumn[]; sort?: { col: number; descending?: boolean }[] }

export interface TableDef {
  name: string;
  range: CellRange;
  headerRow: boolean;
  totalsRow: boolean;
  columns: { name: string; totalsFunction?: string; calculatedFormula?: string }[];
  style?: string;
  bandedRows?: boolean;
  bandedCols?: boolean;
}

export type ChartType =
  | 'column' | 'stackedColumn' | 'percentColumn'
  | 'bar' | 'stackedBar' | 'percentBar'
  | 'line' | 'stackedLine' | 'percentLine' | 'lineMarkers' | 'stackedLineMarkers' | 'percentLineMarkers'
  | 'area' | 'stackedArea' | 'percentArea'
  | 'pie' | 'doughnut'
  | 'scatter' | 'scatterSmoothMarkers' | 'scatterSmooth' | 'scatterLinesMarkers' | 'scatterLines' | 'bubble'
  | 'histogram' | 'pareto' | 'boxWhisker'
  | 'combo' | 'comboSecondary' | 'comboAreaColumn'
  | 'treemap' | 'sunburst'
  | 'waterfall' | 'radar' | 'radarMarkers' | 'radarFilled' | 'funnel'
  | ChartTypeExtra;
export interface ChartSeries extends ChartSeriesExtra {
  name?: string; values: string; categories?: string; color?: string;
  /** How a series is drawn in a combo chart. */
  type?: 'column' | 'line' | 'area';
  /** Bubble sizes. */
  sizes?: string;
  /** Category labels on several levels (PivotCharts): the label range, and per level (innermost first) one label per point, '' where a label spans. */
  multiLevel?: { ref: string; levels: string[][] };
}
/** Per-chart settings. Every field is optional; a missing field means the automatic choice. */
export interface ChartOptions extends ChartOptionsExtra {
  /** Series indexes measured on a second value axis at the right. */
  secondary?: number[];
  /** Histogram bins: a fixed number of bins, or a fixed bin width. */
  binCount?: number;
  binWidth?: number;
  /** Waterfall: category indexes drawn as totals. */
  totals?: number[];
  /** What data labels show on pie, doughnut, funnel and treemap charts. */
  labelMode?: 'value' | 'percent' | 'both';
  /** False hides the value gridlines. */
  gridlines?: boolean;
  /** Fixed value-axis bounds. */
  min?: number;
  max?: number;
  /** Number format for the value axis and labels. */
  numFmt?: string;
  /** Colour set (0 is the default set). */
  style?: number;
  /** Box and whisker: hide the mean marker or the outlier points. */
  hideMean?: boolean;
  hideOutliers?: boolean;
}
export interface Chart {
  id: string;
  type: ChartType;
  title?: string;
  series: ChartSeries[];
  /** Position over the grid: top-left cell plus pixel offsets, and size in pixels. */
  anchor: { row: number; col: number; dx: number; dy: number; width: number; height: number };
  legend?: 'right' | 'bottom' | 'top' | 'left' | 'none';
  axisTitles?: { x?: string; y?: string; y2?: string };
  dataLabels?: boolean;
  options?: ChartOptions;
  /** PivotChart: the PivotTable ("Sheet!Name") whose layout feeds the series. */
  pivot?: string;
  /** Alternative text for screen readers. */
  alt?: string;
}

export interface Picture { id: string; data: Uint8Array; mime: string; anchor: Chart['anchor']; alt?: string }

export interface Note { author?: string; text: string }

/** A comment thread with replies (separate from plain notes). */
export interface CommentThread { id: string; author: string; text: string; time: string; resolved?: boolean; replies: { author: string; text: string; time: string }[] }

export type ShapeKind = 'rect' | 'roundRect' | 'ellipse' | 'triangle' | 'rightArrow' | 'leftArrow' | 'upArrow' | 'downArrow' | 'line' | 'arrowLine' | 'star5' | 'diamond' | 'pentagon' | 'hexagon' | 'callout' | 'textBox';
/** A floating shape or text box over the grid. */
export interface Shape {
  id: string;
  kind: ShapeKind;
  anchor: Chart['anchor'];
  text?: string;
  fill?: string;
  line?: { color: string; width: number; dash?: 'solid' | 'dash' | 'dot' };
  font?: CellStyle['font'];
  align?: 'left' | 'center' | 'right';
  rotation?: number;
  alt?: string;
}

/** A pen, highlighter or marker stroke drawn over the sheet, in sheet pixels at 100% zoom. */
export interface InkStroke { id: string; tool: 'pen' | 'highlighter' | 'marker'; color: string; width: number; points: [number, number][] }

export interface NamedSheetView { name: string; filter?: AutoFilter; hiddenRows: number[]; hiddenCols: number[]; /** Stable id ({GUID}) used in the file. */ id?: string }

/** Error-checking rules that can be ignored per range. */
export type ErrorCheckType = 'numberStoredAsText' | 'formula' | 'formulaRange' | 'unlockedFormula' | 'twoDigitTextYear' | 'evalError' | 'emptyCellReference' | 'calculatedColumn' | 'listDataValidation';

/** A modern password hash (ECMA-376 algorithmName / hashValue / saltValue / spinCount). */
export interface PasswordHash { algorithmName: string; hashValue: string; saltValue: string; spinCount: number }

/** A saved custom view of the workbook: per-sheet display, hidden rows and columns, filters and print settings. */
export interface CustomView {
  name: string;
  guid: string;
  includePrint: boolean;
  includeHidden: boolean;
  activeSheet: number;
  sheets: {
    sheet: number;
    state?: Sheet['state'];
    zoom?: number;
    frozenRows?: number;
    frozenCols?: number;
    showGrid?: boolean;
    showHeaders?: boolean;
    selection?: CellRange;
    hiddenRows?: number[];
    hiddenCols?: number[];
    filter?: AutoFilter;
    pageSetup?: Sheet['pageSetup'];
  }[];
}

export interface SheetView {
  frozenRows: number;
  frozenCols: number;
  showGrid: boolean;
  showHeaders: boolean;
  zoom: number;
  rtl: boolean;
  selection?: CellRange;
  /** Page break preview or page layout view instead of the normal grid. */
  pageView?: 'pageBreakPreview' | 'pageLayout';
  /** Split window (not frozen): sizes of the top and left panes and their own scroll offsets, in sheet pixels at 100%. */
  split?: { x: number; y: number; scrollX: number; scrollY: number };
}

export interface Sheet {
  name: string;
  /** Live collaboration: a stable id and a sort key for the sheet's place in the list (not saved to files). */
  cid?: string;
  cpos?: string;
  cells: Map<number, Cell>;
  cols: Map<number, ColInfo>;
  rows: Map<number, RowInfo>;
  merges: CellRange[];
  view: SheetView;
  defaultColWidth: number;
  defaultRowHeight: number;
  conditionalFormats: ConditionalFormat[];
  validations: DataValidation[];
  autoFilter?: AutoFilter;
  tables: TableDef[];
  charts: Chart[];
  pictures: Picture[];
  shapes?: Shape[];
  ink?: InkStroke[];
  /** PivotTables placed on this sheet. */
  pivots?: import('./pivot').PivotTable[];
  /** Slicers and timelines floating over this sheet. */
  slicers?: import('./pivot').Slicer[];
  /** Outline settings: summary rows below detail and summary columns right of detail (both default true). */
  outlinePr?: { summaryBelow?: boolean; summaryRight?: boolean };
  comments?: Map<number, CommentThread>;
  sheetViews?: NamedSheetView[];
  /** Error-check indicators turned off for these ranges (worksheet ignoredErrors). */
  ignoredErrors?: { range: CellRange; types: ErrorCheckType[] }[];
  /** Ranges people can edit while the sheet is protected. */
  editRanges?: { title: string; range: CellRange; password?: string }[];
  /** 'chart' is a chart sheet: no cells, one chart filling the view (charts[0]). */
  kind?: 'chart';
  sparklines?: SparklineGroup[];
  hyperlinks: Map<number, string>;
  notes: Map<number, Note>;
  tabColor?: string;
  state: 'visible' | 'hidden' | 'veryHidden';
  protection?: { password?: string; hash?: PasswordHash; sheet: boolean; options?: Record<string, boolean> };
  pageSetup?: PageSetup;
}

/** A picture in a header or footer (the &G code). Slot names: L/C/R + H/F, plus FIRST or EVEN for those variants, e.g. "CH", "LFFIRST". */
export interface HeaderPicture { data: Uint8Array; mime: string; /** Size in points. */ width: number; height: number }

/** Print settings of a sheet. Margins are in inches. A missing field means the default. */
export interface PageSetup {
  orientation?: 'portrait' | 'landscape';
  /** Paper size code from the file format (1 Letter, 9 A4, ...). */
  paper?: number;
  /** Custom paper size, e.g. "210mm" or "8.5in". */
  paperWidth?: string; paperHeight?: string;
  /** Percent, 10-400. */
  scale?: number;
  /** Fit to pages wide / tall; set either to switch to fit mode. 0 means no limit in that direction. */
  fitWidth?: number; fitHeight?: number;
  /** Ranges like "A1:D20" or "A1:D20,F1:H9" (no sheet name). Stored in the file as the Print_Area defined name. */
  printArea?: string;
  margins?: { left: number; right: number; top: number; bottom: number; header: number; footer: number };
  /** Odd-page (or every page) header and footer, with &-codes. */
  header?: string; footer?: string;
  evenHeader?: string; evenFooter?: string; firstHeader?: string; firstFooter?: string;
  differentOddEven?: boolean; differentFirst?: boolean;
  /** Both default to true. */
  scaleWithDoc?: boolean; alignWithMargins?: boolean;
  headerPictures?: Record<string, HeaderPicture>;
  /** Rows to repeat at the top, "1:2"; columns to repeat at the left, "A:B". */
  printTitleRows?: string; printTitleCols?: string;
  gridlines?: boolean; headings?: boolean;
  centerH?: boolean; centerV?: boolean;
  firstPageNumber?: number;
  order?: 'downThenOver' | 'overThenDown';
  /** Manual page breaks: the index (0-based) of the first row or column of the new page. */
  rowBreaks?: number[]; colBreaks?: number[];
  blackAndWhite?: boolean; draft?: boolean;
  cellComments?: 'none' | 'atEnd' | 'asDisplayed';
  errors?: 'displayed' | 'blank' | 'dash' | 'NA';
  horizontalDpi?: number; verticalDpi?: number; copies?: number;
}

/** A refreshable import from a text file or another workbook. */
export interface DataQuery {
  id: string;
  name: string;
  kind: 'text' | 'workbook';
  /** Name of the source file. */
  file: string;
  /** Full path of the source file on the desktop app, so Refresh can read it again. */
  path?: string;
  /** Text import settings as JSON (text sources). */
  options?: string;
  /** Sheet and range or table in the source workbook. */
  source?: { sheet: string; range?: string };
  /** Where the data goes: top-left cell and the size last written. */
  target: { sheet: string; row: number; col: number; rows: number; cols: number };
  /** ISO time of the last refresh. */
  refreshed?: string;
}

export interface DefinedName { name: string; formula: string; sheet?: number; hidden?: boolean; comment?: string }

export interface Workbook {
  sheets: Sheet[];
  names: DefinedName[];
  /** styles[0] is the default style. */
  styles: CellStyle[];
  date1904: boolean;
  activeSheet: number;
  props?: { title?: string; creator?: string; created?: string; modified?: string };
  calc?: { iterate?: boolean; maxIterations?: number; maxChange?: number; manual?: boolean; /** Automatic except data tables. */ autoNoTable?: boolean; /** Round stored values to their displayed precision. */ precisionAsDisplayed?: boolean; /** Show references in R1C1 style. */ r1c1?: boolean };
  /** Workbook structure protection: sheets can't be added, removed, renamed, moved or unhidden. */
  protection?: { structure: boolean; windows?: boolean; password?: string; hash?: PasswordHash };
  /** Saved custom views (workbook customWorkbookViews). */
  customViews?: CustomView[];
  /** Saved automation scripts (Automate). */
  scripts?: { id: string; name: string; code: string }[];
  /** Package parts the reader didn't understand, written back untouched where possible. */
  extraParts?: Map<string, Uint8Array>;
  /** Data imported from other files (Get Data), kept so it can be refreshed. */
  queries?: DataQuery[];
  /** Pivot caches shared by PivotTables. Treated as immutable: replace a cache instead of changing it. */
  pivotCaches?: import('./pivot').PivotCache[];
  /** Slicer and timeline caches (their filter state). */
  slicerCaches?: import('./pivot').SlicerCache[];
}

export const DEFAULT_COL_WIDTH = 8.43;
export const DEFAULT_ROW_HEIGHT = 15;
export const DEFAULT_FONT = { name: 'Calibri', size: 11 } as const;

export const newSheet = (name: string): Sheet => ({
  name,
  cells: new Map(),
  cols: new Map(),
  rows: new Map(),
  merges: [],
  view: { frozenRows: 0, frozenCols: 0, showGrid: true, showHeaders: true, zoom: 100, rtl: false },
  defaultColWidth: DEFAULT_COL_WIDTH,
  defaultRowHeight: DEFAULT_ROW_HEIGHT,
  conditionalFormats: [],
  validations: [],
  tables: [],
  charts: [],
  pictures: [],
  hyperlinks: new Map(),
  notes: new Map(),
  state: 'visible',
});

/** Defaults for sheets people add (not sheets read from files): right to left follows the interface language. */
export const userSheetDefaults = { rtl: false };
export const newUserSheet = (name: string): Sheet => { const s = newSheet(name); s.view.rtl = userSheetDefaults.rtl; return s; };

export const newWorkbook = (): Workbook => ({ sheets: [newSheet('Sheet1')], names: [], styles: [{}], date1904: false, activeSheet: 0 });

/** Finds or adds a style and returns its index, so equal styles share one entry. */
export const internStyle = (wb: Workbook, style: CellStyle): number => {
  const key = JSON.stringify(style);
  if (key === '{}') return 0;
  const cache = styleIndex.get(wb) ?? new Map<string, number>();
  if (!styleIndex.has(wb)) {
    wb.styles.forEach((s, i) => cache.set(JSON.stringify(s), i));
    styleIndex.set(wb, cache);
  }
  const hit = cache.get(key);
  if (hit !== undefined && JSON.stringify(wb.styles[hit]) === key) return hit;
  wb.styles.push(style);
  cache.set(key, wb.styles.length - 1);
  return wb.styles.length - 1;
};
const styleIndex = new WeakMap<Workbook, Map<string, number>>();

export const colName = (col: number): string => {
  let s = '';
  let n = col + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
};

export const colIndex = (name: string): number => {
  let n = 0;
  for (const ch of name.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
};

export const addr = (row: number, col: number): string => `${colName(col)}${row + 1}`;
export const rangeAddr = (r: CellRange): string => (r.r1 === r.r2 && r.c1 === r.c2 ? addr(r.r1, r.c1) : `${addr(r.r1, r.c1)}:${addr(r.r2, r.c2)}`);

export const parseAddr = (a: string): { row: number; col: number } | null => {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(a.trim());
  if (!m) return null;
  const col = colIndex(m[1]!);
  const row = +m[2]! - 1;
  return row >= 0 && col >= 0 && col < MAX_COLS ? { row, col } : null;
};

export const parseRange = (a: string): CellRange | null => {
  const [p, q] = a.split(':');
  const s = parseAddr(p ?? '');
  const e = q ? parseAddr(q) : s;
  if (!s || !e) return null;
  return { r1: Math.min(s.row, e.row), c1: Math.min(s.col, e.col), r2: Math.max(s.row, e.row), c2: Math.max(s.col, e.col) };
};

/** Merge styles: `over` wins field by field. */
export const mergeStyle = (base: CellStyle, over: CellStyle): CellStyle => ({
  ...base,
  ...over,
  font: base.font || over.font ? { ...base.font, ...over.font } : undefined,
  fill: base.fill || over.fill ? { ...base.fill, ...over.fill } : undefined,
  border: base.border || over.border ? { ...base.border, ...over.border } : undefined,
  align: base.align || over.align ? { ...base.align, ...over.align } : undefined,
  protection: base.protection || over.protection ? { ...base.protection, ...over.protection } : undefined,
});
