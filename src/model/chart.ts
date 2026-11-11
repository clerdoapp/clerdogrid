// Chart and sparkline settings beyond the basics in workbook.ts. Every field is optional:
// a missing field means the automatic choice.

export type ChartTypeExtra =
  | 'stockHLC' | 'stockOHLC' | 'stockVHLC' | 'stockVOHLC'
  | 'surface3D' | 'surfaceWireframe3D' | 'contour' | 'contourWireframe'
  | 'column3DClustered' | 'column3DStacked' | 'column3DPercent' | 'column3D'
  | 'bar3DClustered' | 'bar3DStacked' | 'bar3DPercent'
  | 'cylinderColumn' | 'coneColumn' | 'pyramidColumn'
  | 'line3D' | 'area3D' | 'stackedArea3D' | 'percentArea3D'
  | 'pie3D' | 'pieExploded' | 'pie3DExploded' | 'doughnutExploded' | 'barOfPie' | 'pieOfPie'
  | 'regionMap' | 'bubble3D';

export type LineDash = 'solid' | 'dot' | 'dash' | 'dashDot' | 'longDash' | 'longDashDot';
export type MarkerSymbol = 'auto' | 'none' | 'circle' | 'square' | 'diamond' | 'triangle' | 'x' | 'star' | 'plus' | 'dash' | 'dot';

export interface Gradient { from: string; to: string; /** Degrees, 0 = left to right, 90 = top to bottom. */ angle?: number }
export interface FillFormat { color?: string; gradient?: Gradient; none?: boolean; /** 0 (opaque) to 100 (clear). */ transparency?: number }
/** Width in points. */
export interface LineFormat { color?: string; width?: number; dash?: LineDash; none?: boolean }
export interface ShapeFormat { fill?: FillFormat; line?: LineFormat }
export interface MarkerFormat { symbol?: MarkerSymbol; size?: number; fill?: string; line?: string }

export interface PointFormat { idx: number; fill?: FillFormat; line?: LineFormat; marker?: MarkerFormat; /** Pie slice pulled out, percent of radius. */ explosion?: number }

export type TrendKind = 'linear' | 'exp' | 'log' | 'poly' | 'power' | 'movingAvg';
export interface Trendline {
  kind: TrendKind;
  /** Polynomial order, 2 to 6. */
  order?: number;
  /** Moving average period. */
  period?: number;
  /** Forecast in x units. */
  forward?: number;
  backward?: number;
  /** Fixed value where the line crosses the y axis (linear, exponential, polynomial). */
  intercept?: number;
  showEquation?: boolean;
  showR2?: boolean;
  name?: string;
  line?: LineFormat;
}

export type ErrorKind = 'fixed' | 'percent' | 'stdDev' | 'stdErr' | 'custom';
export interface ErrorBars {
  /** x bars only for scatter and bubble charts. */
  dir?: 'x' | 'y';
  kind: ErrorKind;
  /** Amount for fixed, percent and standard deviation (multiplier). */
  value?: number;
  /** Custom amounts: a cell reference or a list like {1,2,3}. */
  plus?: string;
  minus?: string;
  type?: 'both' | 'plus' | 'minus';
  noCap?: boolean;
  line?: LineFormat;
}

export type LabelPos = 'auto' | 'center' | 'insideEnd' | 'insideBase' | 'outsideEnd' | 'left' | 'right' | 'above' | 'below' | 'bestFit';
export interface DataLabelOptions {
  value?: boolean;
  seriesName?: boolean;
  categoryName?: boolean;
  percent?: boolean;
  bubbleSize?: boolean;
  legendKey?: boolean;
  /** Label text from a range of cells, one per point. */
  fromCells?: string;
  leaderLines?: boolean;
  position?: LabelPos;
  separator?: string;
  numFmt?: string;
}

export type DisplayUnits = 'hundreds' | 'thousands' | 'tenThousands' | 'hundredThousands' | 'millions' | 'tenMillions' | 'hundredMillions' | 'billions' | 'trillions';
export type TickMark = 'none' | 'out' | 'in' | 'cross';
export interface AxisOptions {
  min?: number;
  max?: number;
  majorUnit?: number;
  minorUnit?: number;
  logBase?: number;
  reverse?: boolean;
  /** Where the other axis crosses this one. */
  crosses?: 'auto' | 'min' | 'max';
  crossesAt?: number;
  dispUnits?: DisplayUnits;
  dispUnitsLabel?: boolean;
  numFmt?: string;
  majorTick?: TickMark;
  minorTick?: TickMark;
  labelPos?: 'nextTo' | 'low' | 'high' | 'none';
  majorGrid?: boolean;
  minorGrid?: boolean;
  /** Category axis: true treats numbers as dates, false never does; missing means automatic. */
  dateAxis?: boolean;
  baseUnit?: 'days' | 'months' | 'years';
  hidden?: boolean;
  line?: LineFormat;
}

export interface View3D {
  /** Tilt, -90 to 90 degrees. */
  rotX?: number;
  /** Turn, 0 to 360 degrees. */
  rotY?: number;
  /** Field of view, 0 to 240 (0 is flat projection). */
  perspective?: number;
  rightAngle?: boolean;
  /** Depth of the plot as a percent of its width, 20 to 2000. */
  depthPercent?: number;
  /** Gap between rows of a 3D column chart, 0 to 500 percent. */
  gapDepth?: number;
  heightPercent?: number;
}

export interface ChartOptionsExtra {
  axes?: { x?: AxisOptions; y?: AxisOptions; x2?: AxisOptions; y2?: AxisOptions };
  view3D?: View3D;
  /** 3D column and bar shape. */
  shape?: 'box' | 'cylinder' | 'cone' | 'pyramid';
  /** Percent of a bar's width, 0 to 500. */
  gapWidth?: number;
  /** -100 to 100. */
  overlap?: number;
  varyColors?: boolean;
  firstSliceAngle?: number;
  holeSize?: number;
  ofPie?: { split?: 'pos' | 'val' | 'percent'; splitValue?: number; secondSize?: number; gapWidth?: number };
  upDownBars?: boolean;
  upFill?: string;
  downFill?: string;
  hiLowLines?: boolean;
  dropLines?: boolean;
  seriesLines?: boolean;
  dataTable?: { keys?: boolean; horizontal?: boolean; vertical?: boolean; outline?: boolean };
  /** What data labels show, for every series without its own label settings. */
  labels?: DataLabelOptions;
  legendOverlay?: boolean;
  chartArea?: ShapeFormat;
  plotArea?: ShapeFormat;
  walls?: ShapeFormat;
  floor?: ShapeFormat;
  /** Look preset (backgrounds, outlines, gridlines). 0 is the plain look. */
  look?: number;
  /** How empty cells show: gaps, zero, or lines joined across them. */
  blanks?: 'gap' | 'zero' | 'span';
  /** Plot values in hidden rows and columns. */
  showHidden?: boolean;
  map?: { area?: 'auto' | 'world' | 'us'; labels?: 'none' | 'best' | 'all'; low?: string; mid?: string; high?: string; diverging?: boolean };
  bubbleScale?: number;
}

export interface ChartSeriesExtra {
  trendlines?: Trendline[];
  errorBars?: ErrorBars[];
  labels?: DataLabelOptions;
  fill?: FillFormat;
  line?: LineFormat;
  marker?: MarkerFormat;
  smooth?: boolean;
  invertIfNegative?: boolean;
  explosion?: number;
  points?: PointFormat[];
}

// ---------- Sparklines ----------

export type SparklineKind = 'line' | 'column' | 'winLoss';
export interface Sparkline {
  /** Source range, like Sheet1!A2:F2. */
  ref: string;
  /** Cell the sparkline is drawn in. */
  row: number;
  col: number;
}
export interface SparklineGroup {
  kind: SparklineKind;
  sparklines: Sparkline[];
  color?: string;
  negativeColor?: string;
  axisColor?: string;
  markersColor?: string;
  firstColor?: string;
  lastColor?: string;
  highColor?: string;
  lowColor?: string;
  /** Line width in points. */
  lineWeight?: number;
  markers?: boolean;
  high?: boolean;
  low?: boolean;
  first?: boolean;
  last?: boolean;
  negative?: boolean;
  showAxis?: boolean;
  /** Axis bounds: each sparkline on its own, the same for the group, or a fixed value. */
  minAxis?: 'individual' | 'group' | 'custom';
  maxAxis?: 'individual' | 'group' | 'custom';
  manualMin?: number;
  manualMax?: number;
  /** Date range for a date axis (spacing by date). */
  dateRange?: string;
  rightToLeft?: boolean;
  displayEmptyCellsAs?: 'gap' | 'zero' | 'span';
  displayHidden?: boolean;
}
