// The model of a chart as it is shown: what a chart part (xl/charts/chartN.xml)
// states about its series, axes, title and legend, as far as the engine draws
// with it. A view for reading, like the other objects of a drawing: the part
// itself stays in the package and is written back from there.
//
// A series holds its data twice: the reference the file gives (the formula
// text, as the file has it) and the values the file cached beside it. Where
// the reference can be followed to cells of this workbook the cells are what
// is drawn; the cache is what is drawn otherwise (see chartRefAreas()).

import { foldSheetName, MAX_COLS, MAX_ROWS, sheetNameMap, type CellRange, type Sheet, type Workbook } from './types';

/** Series read from one chart; a chart with more shows the first ones. */
export const MAX_CHART_SERIES = 255;
/** Points of one series that are read from the cache or from cells; further ones are left out. */
export const MAX_CHART_POINTS = 100000;
/** Charts read from one workbook, all sheets together; further ones are shown as placeholders. */
export const MAX_WORKBOOK_CHARTS = 2000;
/** Cached values (<c:pt>) read from the charts of one workbook, all together; a chart that would go beyond is shown as a placeholder. */
export const MAX_WORKBOOK_CHART_POINTS = 4000000;
/** Areas of one reference ("(A1:A3,C1:C3)") that are followed. */
export const MAX_CHART_AREAS = 64;
/** Axes read from one chart. */
export const MAX_CHART_AXES = 16;
/** Trend lines read from one series. */
export const MAX_CHART_TRENDLINES = 8;

/** "stock" is a line chart whose series are the open, high, low and close of each category, in that order. */
export type ChartKind = 'bar' | 'line' | 'pie' | 'doughnut' | 'area' | 'scatter' | 'radar' | 'bubble' | 'stock';

/** How the text of a part of a chart is set; what is missing is the default of the place it stands in. */
export interface ChartTextStyle {
  /** Points. */
  size?: number;
  bold?: boolean;
  italic?: boolean;
  /** "#RRGGBB". */
  color?: string;
  /** Font name as the file states it; never trusted. */
  font?: string;
}

/** A line: `color` is "#RRGGBB", null for a line that is not drawn, undefined for the one the chart gives by itself. */
export interface ChartLine {
  color?: string | null;
  /** Pixels. */
  width?: number;
  dash?: boolean;
}

/** Fill and outline of an area of the chart (c:spPr). */
export interface ChartShape {
  /** "#RRGGBB", null for no fill, undefined when the file states none (the chart then gives its own). */
  fill?: string | null;
  line?: ChartLine;
}

/**
 * The values a chart part cached for a data source, as far as the file has them: one entry
 * per <c:pt>, whatever the indexes are. A part that names point 99,999 and nothing else
 * costs one entry, not a hundred thousand.
 */
export interface ChartCache {
  /** The indexes of the points the file has a value for, ascending, each once. */
  index: number[];
  /** The value of each of them: numbers for "num" (null for a point whose text is no number), texts for "str". */
  values: (number | string | null)[];
}

/** Where a series gets names, categories or values from. */
export interface ChartData {
  /** What the cache holds: numbers (c:numRef, c:numLit) or texts (c:strRef, c:strLit, c:multiLvlStrRef). */
  kind: 'num' | 'str';
  /** The reference as the file has it (c:f), without a leading "="; undefined for literal data. */
  ref?: string;
  /** The reference names several levels of categories (c:multiLvlStrRef); `cache` is the level next to the axis. */
  multiLevel?: boolean;
  /** The further levels of the cache of such a reference, from the axis outwards; the first of them is shown as a second row. */
  levels?: ChartCache[];
  /** The values of the file, only for the points the file has any for. See chartCached(). */
  cache: ChartCache;
  /** How many points the file says there are (c:ptCount); at least one more than the last index of `cache`. */
  count: number;
  /** Number format of the cached numbers (c:formatCode). */
  formatCode?: string;
}

/** The place of a point index in a cache, or -1 when the file has no value for the point. */
function cachePlace(cache: ChartCache, index: number): number {
  const list = cache.index;
  // Most caches have every point: the place is the index.
  if (index >= 0 && index < list.length && list[index] === index) return index;
  let low = 0;
  let high = list.length - 1;
  while (low <= high) {
    const mid = (low + high) >> 1;
    const at = list[mid];
    if (at === index) return mid;
    if (at < index) low = mid + 1;
    else high = mid - 1;
  }
  return -1;
}

/** The value a chart part cached for the point `index`; null for a point the file has no value for. */
export function chartCached(data: ChartData | undefined, index: number): number | string | null {
  if (!data || !data.cache || !Array.isArray(data.cache.index) || !Array.isArray(data.cache.values)) return null;
  const place = cachePlace(data.cache, index);
  if (place < 0) return null;
  const value = data.cache.values[place];
  return value === undefined ? null : value;
}

/**
 * The cache as a list with one place per point, empty places included: for tests and for
 * charts of a few points. At most `limit` places are made (default 10,000).
 */
export function chartCacheList(data: ChartData | undefined, limit = 10000): (number | string | null)[] {
  const out: (number | string | null)[] = [];
  if (!data || !data.cache || !Array.isArray(data.cache.index) || !Array.isArray(data.cache.values)) return out;
  const { index, values } = data.cache;
  for (let i = 0; i < index.length; i++) {
    const at = index[i];
    if (!(at >= 0) || at >= limit) break;
    while (out.length < at) out.push(null);
    const value = values[i];
    out[at] = value === undefined ? null : value;
  }
  return out;
}

/** Where a data label stands beside its point (c:dLblPos). */
export type ChartLabelPosition = 'outEnd' | 'inEnd' | 'ctr' | 'inBase' | 'bestFit' | 't' | 'b' | 'l' | 'r';

/**
 * What data labels show and how (c:dLbls of a chart type or of a series, c:dLbl of a point).
 * What a level leaves open is what the level above it says: point, series, chart type.
 */
export interface ChartDataLabels {
  /** The labels are taken away (c:delete). */
  deleted?: boolean;
  showValue?: boolean;
  showCategory?: boolean;
  showSeries?: boolean;
  /** The share of the point in its series, pies and doughnuts only. */
  showPercent?: boolean;
  showBubbleSize?: boolean;
  /** Stated, never drawn. */
  showLegendKey?: boolean;
  /** What stands between the parts of a label. */
  separator?: string;
  formatCode?: string;
  /** The numbers are formatted as the cells of the series are. */
  sourceLinked?: boolean;
  position?: ChartLabelPosition;
  style?: ChartTextStyle;
  /** A label of a point with a text of its own (c:tx): the plain text of it. */
  text?: string;
}

/** The label of one point (c:dLbl). */
export interface ChartPointLabel extends ChartDataLabels {
  index: number;
}

/** A trend line of a series (c:trendline). */
export interface ChartTrendline {
  type: 'linear' | 'exp' | 'log' | 'poly' | 'power' | 'movingAvg';
  /** Polynomial: the degree, 2..6. */
  order?: number;
  /** Moving average: how many points are averaged, 2 and up. */
  period?: number;
  /** How far the line goes on beyond the last point and before the first, in units of the horizontal axis. */
  forward?: number;
  backward?: number;
  name?: string;
  line?: ChartLine;
}

/** Error bars of a series (c:errBars). */
export interface ChartErrorBars {
  /** "y": along the values; "x": along the horizontal values of a scatter or bubble series. */
  direction: 'x' | 'y';
  type: 'both' | 'plus' | 'minus';
  valueType: 'fixedVal' | 'percentage' | 'stdDev' | 'stdErr' | 'cust';
  /** The amount (fixedVal), the per cent (percentage), or the number of deviations (stdDev). */
  value?: number;
  /** Amounts of each point for "cust". */
  plus?: ChartData;
  minus?: ChartData;
  noEndCap?: boolean;
  line?: ChartLine;
}

/** A place set by hand (c:manualLayout): shares of the frame of the chart. Taken as stated; what is no place is not applied. */
export interface ChartManualLayout {
  /** Plot area only: "inner" is the area of the series, "outer" includes the labels of the axes. */
  target?: 'inner' | 'outer';
  /** "edge": from the corner of the frame; "factor": moved from where the part would stand by itself. */
  xMode?: 'edge' | 'factor';
  yMode?: 'edge' | 'factor';
  /** "edge": `w` (`h`) is where the right (bottom) edge stands, not a width (height). Missing: a size. */
  wMode?: 'edge';
  hMode?: 'edge';
  x?: number;
  y?: number;
  w?: number;
  h?: number;
}

export interface ChartMarker {
  /** "none", "auto", "circle", "square", "diamond", "triangle", "x", "plus", "star", "dash", "dot" ("picture" reads as "square"). */
  symbol: string;
  /** Points across (2..72). */
  size?: number;
  fill?: string | null;
  line?: ChartLine;
}

/** One point of a series that is drawn its own way (c:dPt): a slice of a pie, a single column. */
export interface ChartPoint {
  index: number;
  shape?: ChartShape;
  marker?: ChartMarker;
  /** A slice of a pie: how far it is pulled out, as a percentage of the radius. */
  explosion?: number;
}

export interface ChartSeries {
  /** c:idx: which colour of the chart's own row of colours the series gets. */
  index: number;
  /** c:order: where the series stands among the others. */
  order: number;
  /** The name: a reference with its cached text, or a literal text. */
  name?: { ref?: string; text?: string };
  /** c:cat, or c:xVal of a scatter chart. */
  categories?: ChartData;
  /** c:val, or c:yVal of a scatter chart. */
  values?: ChartData;
  shape?: ChartShape;
  /** Undefined when the series states none: the chart type decides. */
  marker?: ChartMarker;
  points?: ChartPoint[];
  invertIfNegative?: boolean;
  /** The file asks for a curve through the points. */
  smooth?: boolean;
  /** The data labels of the series (c:dLbls) and of single points of it (c:dLbl). */
  dataLabels?: ChartDataLabels;
  pointLabels?: ChartPointLabel[];
  trendlines?: ChartTrendline[];
  /** At most one along the values and one along the horizontal values. */
  errorBars?: ChartErrorBars[];
  /** Pie: how far every slice is pulled out, as a percentage of the radius. */
  explosion?: number;
  /** Bubble chart: the sizes of the bubbles (c:bubbleSize). */
  bubbleSizes?: ChartData;
}

/** The boxes between the first and the last series of a line or stock chart (c:upDownBars). */
export interface ChartUpDownBars {
  /** Space between the boxes as a percentage of their width (0..500). */
  gapWidth?: number;
  /** The look of a box whose last series lies above its first, and of one where it lies below. */
  up?: ChartShape;
  down?: ChartShape;
}

/** One chart type of a plot area with its series. A combination chart has several. */
export interface ChartGroup {
  kind: ChartKind;
  /** The file has the 3-D form of the type (bar3DChart, ...); it is drawn flat. */
  threeD?: boolean;
  /** Bar charts: "col" for upright columns, "bar" for bars that lie. */
  barDir?: 'col' | 'bar';
  /** "standard": side by side (bars: clustered) or one over another without adding up (lines, areas). */
  grouping: 'standard' | 'stacked' | 'percentStacked';
  /** Space between the clusters as a percentage of the width of a bar (0..500). */
  gapWidth?: number;
  /** How far the bars of a cluster overlap, as a percentage of their width (-100..100). */
  overlap?: number;
  /** Each point gets a colour of its own. */
  varyColors?: boolean;
  /** Doughnut: the hole as a percentage of the diameter. */
  holeSize?: number;
  /** Pie and doughnut: where the first slice starts, in degrees clockwise from the top. */
  firstSliceAngle?: number;
  /** Line chart: series without a marker of their own get one. */
  markers?: boolean;
  /** Scatter chart: "lineMarker", "line", "marker", "smooth", "smoothMarker", "none". */
  scatterStyle?: string;
  /** The data labels of every series of the type that states none of its own (c:dLbls). */
  dataLabels?: ChartDataLabels;
  /** Radar chart: "standard" (lines), "marker" (lines with markers), "filled". */
  radarStyle?: 'standard' | 'marker' | 'filled';
  /** Bubble chart: the size of the largest bubble as a percentage of the default (0..300). */
  bubbleScale?: number;
  /** Bubble chart: the size of a bubble is its area ("area", the default) or its width ("w"). */
  sizeRepresents?: 'area' | 'w';
  /** Bubble chart: bubbles of a negative size are drawn (by the size without its sign). */
  showNegBubbles?: boolean;
  /** Lines from each point down to the category axis (line and area charts). */
  dropLines?: ChartLine;
  /** Lines from the highest to the lowest series at each category (line and stock charts). */
  hiLowLines?: ChartLine;
  upDownBars?: ChartUpDownBars;
  /**
   * The file has a pie with a second plot (c:ofPieChart): only the pie is drawn, the points of the
   * second plot as one slice of it. `split` says which points those are: the last `position` ones
   * ("pos", and "auto" without a position: the last third), those below a value ("val") or a per
   * cent ("percent"), or the ones listed ("cust").
   */
  ofPie?: { split: 'auto' | 'pos' | 'val' | 'percent' | 'cust'; position?: number; points?: number[] };
  /** Ids of the axes the group is drawn against. */
  axisIds: number[];
  series: ChartSeries[];
}

export interface ChartTitle {
  /** The text; undefined for a title that shows what the chart gives by itself. */
  text?: string;
  /** The cell the title is read from (c:strRef), as the file has it; `text` is then what the file cached. */
  ref?: string;
  style?: ChartTextStyle;
  layout?: ChartManualLayout;
}

export interface ChartAxis {
  id: number;
  /** "cat" also stands for a date axis and a series axis. */
  kind: 'cat' | 'val';
  /** The file has a date axis (c:dateAx): categories that are dates stand where their dates lie, not side by side. */
  date?: boolean;
  /** Date axis: the unit a category is as wide as, and the unit `majorUnit` counts in. */
  baseTimeUnit?: 'days' | 'months' | 'years';
  majorTimeUnit?: 'days' | 'months' | 'years';
  deleted?: boolean;
  /** Side of the plot area the axis stands on. */
  position: 'b' | 'l' | 't' | 'r';
  min?: number;
  max?: number;
  /** c:orientation maxMin: the axis runs the other way. */
  reversed?: boolean;
  majorUnit?: number;
  /** c:logBase (2..1000): the axis counts in powers of it. */
  logBase?: number;
  formatCode?: string;
  /** The numbers are formatted as the cells of the first series are. */
  sourceLinked?: boolean;
  majorGridlines?: ChartLine;
  minorGridlines?: ChartLine;
  /** "nextTo", "low", "high", "none". */
  tickLabels: 'nextTo' | 'low' | 'high' | 'none';
  title?: ChartTitle;
  /**
   * Where this axis crosses the one it names (`crossAxis`): at its zero or first category, at its minimum, or at
   * its maximum. A value axis that crosses at the maximum stands at the last category (the far side of the plot
   * area); a category axis that crosses at the maximum stands at the largest value.
   */
  crosses?: 'autoZero' | 'min' | 'max';
  /**
   * The place on the other axis this one crosses at, in place of `crosses`. Of a category axis: the value the
   * series stand on. Of a value axis: a category number, which only moves the axis when it lies beyond the last
   * category (to the far side); between the categories the axis is drawn at the start.
   */
  crossesAt?: number;
  /** Id of the axis this one crosses. */
  crossAxis?: number;
  /** Value axis: "between" puts points in the middle of their categories, "midCat" on the tick marks. */
  crossBetween?: 'between' | 'midCat';
  line?: ChartLine;
  labelStyle?: ChartTextStyle;
  /** How the labels are turned, in degrees clockwise (-90..90), when the file says so (a:bodyPr rot of c:txPr). */
  labelRotation?: number;
}

export interface ChartLegend {
  position: 'r' | 'l' | 't' | 'b' | 'tr';
  /** Indexes of entries that are left out (c:legendEntry with c:delete). */
  deleted?: number[];
  style?: ChartTextStyle;
  layout?: ChartManualLayout;
}

export interface Chart {
  /** The chart types of the plot area, in the order of the file. Never empty. */
  groups: ChartGroup[];
  axes: ChartAxis[];
  title?: ChartTitle;
  /** The file says that no title is made up for a chart without one. */
  autoTitleDeleted?: boolean;
  legend?: ChartLegend;
  /** Cells of hidden rows and columns are left out. Default true. */
  plotVisOnly: boolean;
  /** What stands for a cell without a number. */
  blanks: 'gap' | 'zero' | 'span';
  roundedCorners?: boolean;
  /** Dates of the cache count from 1904. */
  date1904?: boolean;
  /** Fill and outline of the whole chart and of the area the series are drawn in. */
  area?: ChartShape;
  plotArea?: ChartShape;
  /** Where the plot area stands when the file sets it by hand. */
  plotLayout?: ChartManualLayout;
  /** Text of the chart where a part states none of its own. */
  textStyle?: ChartTextStyle;
  /** The six accent colours of the theme ("#RRGGBB"): the colours series get that state none. */
  palette: string[];
  /** What the chart has that is not drawn, or is drawn differently, one line each. */
  notes: string[];
}

// ------------------------------------------------------------------ references

/** One area of a reference of a chart: cells of a sheet of this workbook. */
export interface ChartRefArea {
  sheet: Sheet;
  range: CellRange;
}

const CELL = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/;
const COLUMN = /^\$?([A-Za-z]{1,3})$/;
const ROW = /^\$?([0-9]{1,7})$/;

function columnIndex(letters: string): number {
  let value = 0;
  for (let i = 0; i < letters.length; i++) value = value * 26 + ((letters.charCodeAt(i) & 0xdf) - 64);
  return value - 1;
}

/** The pieces of a reference between the commas that stand outside quotes; undefined for a quote that is never closed. */
function splitAreas(text: string): string[] | undefined {
  const pieces: string[] = [];
  let start = 0;
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    if (code === 39) quoted = !quoted;
    else if (code === 44 && !quoted) {
      pieces.push(text.slice(start, i));
      start = i + 1;
      if (pieces.length > MAX_CHART_AREAS) return undefined;
    }
  }
  if (quoted) return undefined;
  pieces.push(text.slice(start));
  return pieces;
}

function areaOf(workbook: Workbook, piece: string, names: ReturnType<typeof sheetNameMap>): ChartRefArea | undefined {
  const text = piece.trim();
  let name: string;
  let rest: string;
  if (text.charCodeAt(0) === 39) {
    // A name in quotes; a quote inside it is written twice.
    let end = -1;
    for (let i = 1; i < text.length; i++) {
      if (text.charCodeAt(i) !== 39) continue;
      if (text.charCodeAt(i + 1) === 39) i++;
      else {
        end = i;
        break;
      }
    }
    if (end < 0 || text.charCodeAt(end + 1) !== 33) return undefined;
    name = text.slice(1, end).replace(/''/g, "'");
    rest = text.slice(end + 2);
  } else {
    const bang = text.indexOf('!');
    if (bang <= 0) return undefined;
    name = text.slice(0, bang);
    rest = text.slice(bang + 1);
  }
  // Another workbook ("[1]Sheet1"), or a run of sheets ("Sheet1:Sheet3"): not followed.
  if (name === '' || name.indexOf('[') >= 0 || name.indexOf(':') >= 0) return undefined;
  // The reference names sheets as the file does.
  const now = names ? names(name) : undefined;
  if (now === null) return undefined;
  const wanted = foldSheetName(now === undefined ? name : now);
  let sheet: Sheet | undefined;
  for (const candidate of workbook.sheets) {
    if (foldSheetName(candidate.name) === wanted) {
      sheet = candidate;
      break;
    }
  }
  if (!sheet || sheet.kind !== 'worksheet') return undefined;
  const colon = rest.indexOf(':');
  const first = colon < 0 ? rest : rest.slice(0, colon);
  const second = colon < 0 ? rest : rest.slice(colon + 1);
  let range: CellRange | undefined;
  const a = CELL.exec(first);
  const b = CELL.exec(second);
  if (a && b) {
    const r1 = Number(a[2]) - 1;
    const r2 = Number(b[2]) - 1;
    const c1 = columnIndex(a[1]);
    const c2 = columnIndex(b[1]);
    range = { top: Math.min(r1, r2), bottom: Math.max(r1, r2), left: Math.min(c1, c2), right: Math.max(c1, c2) };
  } else if (colon >= 0 && COLUMN.test(first) && COLUMN.test(second)) {
    // Whole columns end where the sheet does.
    const c1 = columnIndex((COLUMN.exec(first) as RegExpExecArray)[1]);
    const c2 = columnIndex((COLUMN.exec(second) as RegExpExecArray)[1]);
    range = { top: 0, bottom: Math.max(sheet.rowCount - 1, 0), left: Math.min(c1, c2), right: Math.max(c1, c2) };
  } else if (colon >= 0 && ROW.test(first) && ROW.test(second)) {
    const r1 = Number((ROW.exec(first) as RegExpExecArray)[1]) - 1;
    const r2 = Number((ROW.exec(second) as RegExpExecArray)[1]) - 1;
    range = { top: Math.min(r1, r2), bottom: Math.max(r1, r2), left: 0, right: Math.max(sheet.colCount - 1, 0) };
  }
  if (!range || range.top < 0 || range.left < 0 || range.bottom >= MAX_ROWS || range.right >= MAX_COLS) return undefined;
  return { sheet, range };
}

/**
 * The cells a reference of a chart names, area by area in the order they are
 * written: "Sheet1!$A$1:$A$5", "'My sheet'!$B$2", "(Sheet1!$A$1:$A$3,Sheet1!$C$1:$C$3)".
 * Undefined for whatever cannot be followed to worksheets of this workbook: a
 * reference without a sheet, into another workbook, to a deleted sheet, a
 * defined name, "#REF!", a formula. The values the file cached are drawn then.
 * Never throws.
 */
export function chartRefAreas(workbook: Workbook, reference: string | undefined): ChartRefArea[] | undefined {
  if (typeof reference !== 'string' || reference.length === 0 || reference.length > 8192) return undefined;
  let text = reference.trim();
  if (text.charCodeAt(0) === 61) text = text.slice(1).trim();
  if (text.charCodeAt(0) === 40 && text.charCodeAt(text.length - 1) === 41) text = text.slice(1, text.length - 1);
  const pieces = splitAreas(text);
  if (!pieces) return undefined;
  let names: ReturnType<typeof sheetNameMap>;
  try {
    names = sheetNameMap(workbook);
  } catch (error) {
    return undefined;
  }
  const areas: ChartRefArea[] = [];
  for (const piece of pieces) {
    const area = areaOf(workbook, piece, names);
    if (!area) return undefined;
    areas.push(area);
  }
  return areas.length > 0 ? areas : undefined;
}
