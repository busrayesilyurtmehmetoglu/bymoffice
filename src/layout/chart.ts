// Works out what a chart shows: the numbers of its series (from the cells the
// series refer to, or from the values the file cached), the scale of its axes,
// and where everything is drawn inside the frame of the chart. Pure
// calculation: text is measured through a function that is handed in, and the
// result is a flat list of items for a painter (see render/chart.ts).
//
// Excel is not at hand to compare with: rules that follow what Excel is known
// to do but could not be checked against it are marked UNSURE.

import { formatValue, isDateFormat, serialToDate } from '../format/number-format';
import { builtinFormatId, cellDisplayText, isoToSerial } from '../format/cell-format';
import type { FormatLocale } from '../format/locale';
import {
  chartCached,
  chartRefAreas,
  MAX_CHART_POINTS,
  type Chart,
  type ChartAxis,
  type ChartData,
  type ChartDataLabels,
  type ChartErrorBars,
  type ChartGroup,
  type ChartKind,
  type ChartLine,
  type ChartManualLayout,
  type ChartMarker,
  type ChartPoint,
  type ChartRefArea,
  type ChartSeries,
  type ChartTextStyle,
  type ChartTrendline,
} from '../model/sheet/chart';
import { resolveCellStyle } from '../model/sheet/color';
import { clipLine, fitTrend, seriesSpread, smoothPath } from './chart-fit';
import type { Cell, Sheet, Workbook } from '../model/sheet/types';

/** Positions along the category axis that are drawn for one series; a series with more points is sampled down to it. */
export const MAX_DRAWN_POINTS = 4000;
/** Slices drawn of one ring of a pie; what is left is drawn as one grey slice. */
export const MAX_PIE_SLICES = 1000;
/** Entries a legend lists. */
export const MAX_LEGEND_ENTRIES = 100;
/** Ticks of one axis. */
export const MAX_AXIS_TICKS = 200;
/**
 * Columns, points of lines and markers drawn for one chart, all series together: a chart of very
 * many series shows fewer positions of each, so that painting a frame costs no more than this.
 */
export const MAX_DRAWN_ITEMS = 40000;
/** Spokes of a radar chart; one of more categories shows the largest value of each run of them. */
export const MAX_RADAR_SPOKES = 360;
/** Trend lines drawn for one chart, all series together. */
export const MAX_DRAWN_TRENDS = 64;
/** Data labels worked out for one chart, all series together: further points get none. */
export const MAX_DATA_LABELS = 500;
/**
 * What the layout of one chart weighs at most, whatever the chart holds: see chartLayoutWeight(). A chart whose
 * layout would weigh more is laid out again with fewer positions of each series, until it does not.
 */
export const MAX_LAYOUT_WEIGHT = 1000000;
/** Cells read for one chart, all series, names and categories together; series beyond it are drawn without values. */
export const MAX_CHART_CELLS = 4000000;

// ------------------------------------------------------------------------ data

export interface ChartDataOptions {
  /** The workbook the references of the series are followed in; without it the cached values are drawn. */
  workbook?: Workbook;
  locale?: FormatLocale;
  /** The name of a series that has none: `n` counts from 1. Default "Series1", "Series2", ... */
  seriesName?(n: number): string;
}

/**
 * The numbers of the points of a series. Only the points that hold a number are kept, so
 * that a series of a hundred thousand points of which one has a number costs one number.
 */
export class ChartNumbers {
  /** How many points the series has, those without a number included. */
  readonly length: number;
  /** What a point without a number is drawn as: nothing (null), or zero when the chart says so. */
  readonly blank: number | null;
  /** The indexes of the points that hold a number, ascending; undefined when they are 0, 1, 2, ...: the place of a number is then its index. */
  private readonly index: number[] | undefined;
  private readonly numbers: number[];
  /** Every point holds its place counted from one (1, 2, 3, ...); nothing is stored. */
  private readonly counted: boolean;

  /**
   * `index`: the indexes of the points that hold a number, ascending without repeats, or undefined when the
   * first numbers.length points are the ones (a series without gaps then costs its numbers alone). `numbers`:
   * the value at each of them. Neither is copied.
   */
  constructor(length: number, index: number[] | undefined, numbers: number[], blank: number | null = null, counted = false) {
    this.index = index;
    this.numbers = numbers;
    this.blank = blank;
    this.counted = counted;
    const last = index ? (index.length > 0 ? index[index.length - 1] + 1 : 0) : numbers.length;
    this.length = Math.max(Number.isFinite(length) && length > 0 ? Math.floor(length) : 0, last);
  }

  /** From a list with one place per point; null (or anything that is no finite number) for a point without a number. */
  static of(list: ReadonlyArray<number | null | undefined>, blank: number | null = null): ChartNumbers {
    const gather = new NumberGatherer();
    for (let i = 0; i < list.length; i++) {
      const value = list[i];
      if (typeof value === 'number' && Number.isFinite(value)) gather.add(i, value);
    }
    return gather.numbers(list.length, blank);
  }

  /** 1, 2, 3, ... for `length` points: where the points of a scatter series without numbers for x are placed. */
  static counting(length: number): ChartNumbers {
    return new ChartNumbers(length, undefined, [], null, true);
  }

  /** How many points hold a number of their own. */
  get known(): number {
    return this.counted ? this.length : this.numbers.length;
  }

  /** The index of the k-th point that holds a number. */
  knownIndex(k: number): number {
    return this.index ? this.index[k] : k;
  }

  /** The number of the k-th point that holds one. */
  knownValue(k: number): number {
    return this.counted ? k + 1 : this.numbers[k];
  }

  /** The number of point `i`: what stands for a point without one (null or zero), and null beyond the last point. */
  at(i: number): number | null {
    if (!(i >= 0 && i < this.length) || Math.floor(i) !== i) return null;
    if (this.counted) return i + 1;
    const index = this.index;
    if (!index) return i < this.numbers.length ? this.numbers[i] : this.blank;
    let low = 0;
    let high = index.length - 1;
    while (low <= high) {
      const mid = (low + high) >> 1;
      const found = index[mid];
      if (found === i) return this.numbers[mid];
      if (found < i) low = mid + 1;
      else high = mid - 1;
    }
    return this.blank;
  }

  /** Calls `visit` for every point that holds a number of its own, in the order of the points. */
  each(visit: (value: number, index: number) => void): void {
    const count = this.known;
    for (let k = 0; k < count; k++) visit(this.knownValue(k), this.knownIndex(k));
  }

  /** One place per point, as at() gives them: for tests and for series of a few points. At most `limit` places (default 100,000). */
  toArray(limit = 100000): (number | null)[] {
    const out: (number | null)[] = [];
    const count = Math.min(this.length, limit);
    for (let i = 0; i < count; i++) out.push(this.at(i));
    return out;
  }
}

/** Gathers the numbers of a series point by point, in the order of the points; keeps no indexes while there is no gap. */
class NumberGatherer {
  private index: number[] | undefined;
  private readonly values: number[] = [];

  /** How many numbers were added. */
  get size(): number {
    return this.values.length;
  }

  /** The index of the number added last; -1 before the first. */
  get last(): number {
    return this.values.length === 0 ? -1 : this.index ? this.index[this.index.length - 1] : this.values.length - 1;
  }

  /** `at` has to lie beyond every index added before. */
  add(at: number, value: number): void {
    if (!this.index && at !== this.values.length) {
      // The first gap: from here on the places are no longer the indexes.
      this.index = [];
      for (let i = 0; i < this.values.length; i++) this.index.push(i);
    }
    if (this.index) this.index.push(at);
    this.values.push(value);
  }

  numbers(length: number, blank: number | null): ChartNumbers {
    return new ChartNumbers(length, this.index, this.values, blank);
  }
}

/** One series with the numbers it is drawn from. */
export interface ChartSeriesValues {
  /** Index of the chart type of the plot area the series belongs to. */
  group: number;
  kind: ChartKind;
  series: ChartSeries;
  name: string;
  /** The number of each point; see ChartNumbers. */
  values: ChartNumbers;
  /** Scatter and bubble charts: the horizontal place of each point. */
  xs?: ChartNumbers;
  /** Bubble charts: the size of each bubble; undefined when the series states none. */
  sizes?: ChartNumbers;
  /** The amounts of the error bars of the series that state them point by point ("cust"), in the order of `series.errorBars`. */
  errors?: ({ plus?: ChartNumbers; minus?: ChartNumbers } | undefined)[];
  /** Number format of the values: that of the first cell, or the one of the cache. */
  format?: { code: string; builtinId?: number };
  /** The values come from cells of the workbook, not from the cache of the file. */
  live: boolean;
}

export interface ChartValues {
  series: ChartSeriesValues[];
  /** Number of categories: the most points any series has. */
  count: number;
  /** The label of a category; formatted when asked for, since only a few of many are ever shown. */
  label(index: number): string;
  /** With two levels of categories: the label of the outer level that starts at a category, '' where none does. */
  outerLabel?(index: number): string;
  /** The number of a category, where the categories are numbers or dates: what a date axis places its points by. */
  categoryNumber?(index: number): number | null;
  /** The number format of the categories that are numbers. */
  categoryFormat?: { code: string; builtinId?: number };
  /** The title, when the chart has one that is read from a cell. */
  title?: string;
  /** What is held for the chart: numbers of points and cells of labels, counted. What a keeper of many charts weighs them by. */
  size: number;
}

/** The columns of a range that are read: all of them, or those that are not hidden. */
function columnsOf(sheet: Sheet, left: number, right: number, visibleOnly: boolean): number[] {
  const width = right - left + 1;
  const hidden = new Uint8Array(width);
  const columns = visibleOnly ? sheet.columns || [] : [];
  for (let i = 0; i < columns.length; i++) {
    const column = columns[i];
    if (!column || !column.hidden) continue;
    const from = Math.max(column.min, left);
    const to = Math.min(column.max, right);
    for (let c = from; c <= to; c++) hidden[c - left] = 1;
  }
  const out: number[] = [];
  for (let c = 0; c < width; c++) if (hidden[c] === 0) out.push(left + c);
  return out;
}

/**
 * Visits the cells of the areas in the order they are written, row by row, at most `limit`
 * of them, and says how many there were. A row that holds nothing costs no more than being counted.
 */
function eachCell(areas: ChartRefArea[], visibleOnly: boolean, limit: number, visit: (cell: Cell, position: number) => void): number {
  let position = 0;
  for (let a = 0; a < areas.length && position < limit; a++) {
    const { sheet, range } = areas[a];
    const columns = columnsOf(sheet, range.left, range.right, visibleOnly);
    if (columns.length === 0) continue;
    for (let r = range.top; r <= range.bottom && position < limit; r++) {
      const row = sheet.rows.get(r);
      if (visibleOnly && row && row.hidden) continue;
      if (!row) {
        position = Math.min(position + columns.length, limit);
        continue;
      }
      for (let i = 0; i < columns.length && position < limit; i++) {
        const cell = row.cells[columns[i]];
        if (cell) visit(cell, position);
        position++;
      }
    }
  }
  return position;
}

/** The number a cell gives a chart: null for a cell without one (empty, text, an error, TRUE). */
function cellNumber(cell: Cell | undefined, date1904: boolean): number | null {
  if (!cell) return null;
  if (cell.kind === 'number' && typeof cell.value === 'number' && Number.isFinite(cell.value)) return cell.value;
  if (cell.kind === 'date' && typeof cell.value === 'string') {
    const serial = isoToSerial(cell.value, date1904);
    return serial === undefined ? null : serial;
  }
  // UNSURE: Excel draws a text cell as zero; here it is a cell without a number, like an empty one.
  return null;
}

interface Numbers {
  values: ChartNumbers;
  format?: { code: string; builtinId?: number };
  live: boolean;
}

/** How many more cells may be read for the chart that is being worked out, and what is held for it so far. */
interface Budget {
  left: number;
  held: number;
}

const NO_NUMBERS = new ChartNumbers(0, undefined, []);

/** True for a reference of a chart that follows the edits of its workbook (Chart.live) and names cells that are gone. */
const gone = (chart: Chart, data: ChartData): boolean => chart.live === true && data.ref !== undefined && data.ref.indexOf('#REF!') >= 0;

function numbersOf(chart: Chart, data: ChartData | undefined, options: ChartDataOptions, budget: Budget, zero = chart.blanks === 'zero'): Numbers {
  if (!data) return { values: NO_NUMBERS, live: false };
  const workbook = options.workbook;
  // The cells a series showed were deleted since the file was read: it shows nothing, not what the file cached of them.
  if (workbook && gone(chart, data)) return { values: NO_NUMBERS, live: true };
  const areas = workbook && data.ref !== undefined ? chartRefAreas(workbook, data.ref, chart.live === true) : undefined;
  const gather = new NumberGatherer();
  if (areas && workbook) {
    const date1904 = workbook.date1904 === true;
    let format: Numbers['format'];
    const count = eachCell(areas, chart.plotVisOnly, Math.max(Math.min(MAX_CHART_POINTS, budget.left), 0), (cell, position) => {
      const value = cellNumber(cell, date1904);
      if (value === null) return;
      if (format === undefined) {
        const style = resolveCellStyle(workbook.styles, cell.style);
        format = { code: style.numFmt, builtinId: builtinFormatId(workbook.styles, style.numFmtId) };
      }
      gather.add(position, value);
    });
    budget.left -= count;
    budget.held += gather.size;
    return { values: gather.numbers(count, zero ? 0 : null), format, live: true };
  }
  const cache = data.cache;
  // A model that no reader made may hold anything here: what is no cache of the model holds no numbers.
  if (data.kind === 'num' && cache && Array.isArray(cache.index) && Array.isArray(cache.values)) {
    for (let i = 0; i < cache.index.length; i++) {
      const at = cache.index[i];
      const value = cache.values[i];
      if (!(at >= 0) || at >= MAX_CHART_POINTS) break;
      if (typeof value !== 'number' || !Number.isFinite(value)) continue;
      // A cache made by hand may be out of order: what does not go on upwards is left out.
      if (at <= gather.last || Math.floor(at) !== at) continue;
      gather.add(at, value);
    }
  }
  budget.held += gather.size;
  const count = Math.min(Math.max(data.count || 0, cachedLength(data)), MAX_CHART_POINTS);
  return { values: gather.numbers(count, zero ? 0 : null), format: data.formatCode !== undefined ? { code: data.formatCode } : undefined, live: false };
}

/** One more than the last index a cache has a value for. */
function cachedLength(data: ChartData): number {
  const index = data.cache ? data.cache.index : undefined;
  const last = Array.isArray(index) && index.length > 0 ? index[index.length - 1] : -1;
  return last >= 0 && Number.isFinite(last) ? Math.floor(last) + 1 : 0;
}

interface Labels {
  count: number;
  label(index: number): string;
  outer?(index: number): string;
  number(index: number): number | null;
  format?: { code: string; builtinId?: number };
}

/**
 * The number format a category axis gives its labels: only one the file states as not taken
 * from the cells (sourceLinked="0"). UNSURE: "General" stated that way is left to the cells
 * too - files not written by Excel state it for every axis, and a date would show as its
 * serial number.
 */
function categoryFormat(axis: ChartAxis | undefined): string | undefined {
  if (!axis || axis.sourceLinked !== false || !axis.formatCode) return undefined;
  return /^general$/i.test(axis.formatCode) ? undefined : axis.formatCode;
}

function formatted(value: number, code: string, locale: FormatLocale | undefined, date1904: boolean): string | undefined {
  try {
    return formatValue(value, code, { locale, date1904 }).text;
  } catch (error) {
    return undefined;
  }
}

function labelsOf(chart: Chart, data: ChartData | undefined, options: ChartDataOptions, budget: Budget, axis: ChartAxis | undefined): Labels | undefined {
  if (!data) return undefined;
  const workbook = options.workbook;
  if (workbook && gone(chart, data)) return undefined;
  const code = categoryFormat(axis);
  let areas = workbook && data.ref !== undefined ? chartRefAreas(workbook, data.ref, chart.live === true) : undefined;
  let outerAreas: ChartRefArea[] | undefined;
  if (areas && data.multiLevel) {
    // UNSURE: of a block of several levels the last column (or, for a block that lies, the last row) is the level next to the axis,
    // and the one before it the level outside that.
    if (areas.length !== 1) areas = undefined;
    else {
      const { sheet, range } = areas[0];
      const tall = range.bottom - range.top >= range.right - range.left;
      areas = [{ sheet, range: tall ? { ...range, left: range.right } : { ...range, top: range.bottom } }];
      if (tall && range.right > range.left) outerAreas = [{ sheet, range: { ...range, left: range.right - 1, right: range.right - 1 } }];
      else if (!tall && range.bottom > range.top) outerAreas = [{ sheet, range: { ...range, top: range.bottom - 1, bottom: range.bottom - 1 } }];
    }
  }
  if (areas && workbook) {
    // Only the cells that hold something are kept, by their place among the categories.
    const cells = new Map<number, Cell>();
    const count = eachCell(areas, chart.plotVisOnly, Math.max(Math.min(MAX_CHART_POINTS, budget.left), 0), (cell, position) => void cells.set(position, cell));
    budget.left -= count;
    budget.held += cells.size;
    const date1904 = workbook.date1904 === true;
    let outerCells: Map<number, Cell> | undefined;
    if (outerAreas) {
      const kept = new Map<number, Cell>();
      budget.left -= eachCell(outerAreas, chart.plotVisOnly, Math.max(Math.min(MAX_CHART_POINTS, budget.left), 0), (cell, position) => void kept.set(position, cell));
      budget.held += kept.size;
      outerCells = kept;
    }
    // The format of the first category that is a number: what a date axis writes the dates between its points in.
    let format: Labels['format'];
    cells.forEach((cell) => {
      if (format !== undefined || cellNumber(cell, date1904) === null) return;
      const style = resolveCellStyle(workbook.styles, cell.style);
      format = { code: style.numFmt, builtinId: builtinFormatId(workbook.styles, style.numFmtId) };
    });
    return {
      count,
      format,
      number: (index) => cellNumber(cells.get(index), date1904),
      outer: outerCells ? (index) => cellDisplayText(workbook, (outerCells as Map<number, Cell>).get(index), options.locale) : undefined,
      label: (index) => {
        const cell = cells.get(index);
        if (code !== undefined) {
          // A number or a date of the sheet in the format the axis states for its labels.
          const value = cellNumber(cell, date1904);
          const text = value === null ? undefined : formatted(value, code, options.locale, date1904);
          if (text !== undefined) return text;
        }
        return cellDisplayText(workbook, cell, options.locale);
      },
    };
  }
  const count = Math.min(Math.max(data.count || 0, cachedLength(data)), MAX_CHART_POINTS);
  const date1904 = chart.date1904 === true;
  const level = Array.isArray(data.levels) && data.levels[0] ? ({ kind: 'str', cache: data.levels[0], count } as ChartData) : undefined;
  return {
    count,
    format: data.kind === 'num' && data.formatCode !== undefined ? { code: data.formatCode } : undefined,
    number: (index) => {
      const value = data.kind === 'num' ? chartCached(data, index) : null;
      return typeof value === 'number' && Number.isFinite(value) ? value : null;
    },
    outer: level
      ? (index) => {
          const value = chartCached(level, index);
          return typeof value === 'string' ? value : typeof value === 'number' ? String(value) : '';
        }
      : undefined,
    label: (index) => {
      const value = chartCached(data, index);
      if (typeof value === 'string') return value;
      if (typeof value === 'number') return formatted(value, code !== undefined ? code : data.formatCode || 'General', options.locale, date1904) || '';
      return '';
    },
  };
}

/** The text of the first cell a reference names; undefined when it cannot be followed. */
function cellText(reference: string | undefined, options: ChartDataOptions, live: boolean): string | undefined {
  const workbook = options.workbook;
  if (!workbook || reference === undefined) return undefined;
  const areas = chartRefAreas(workbook, reference, live);
  if (!areas) return undefined;
  const { sheet, range } = areas[0];
  const row = sheet.rows.get(range.top);
  return cellDisplayText(workbook, row ? row.cells[range.left] : undefined, options.locale);
}

/** Where the points of a scatter series stand: at the numbers of its x values, or at 1, 2, 3, ... when a point that has a y has no number for x. */
function placesOf(xs: ChartNumbers | undefined, ys: ChartNumbers): ChartNumbers {
  if (!xs || xs.length === 0) return ChartNumbers.counting(ys.length);
  const within = Math.min(xs.length, ys.length);
  if (ys.blank !== null) {
    // Every point has a y: every point within the x values needs a number for x.
    if (xs.known < within) return ChartNumbers.counting(ys.length);
    let before = 0;
    while (before < xs.known && xs.knownIndex(before) < within) before++;
    if (before < within) return ChartNumbers.counting(ys.length);
  } else {
    for (let k = 0; k < ys.known; k++) {
      const at = ys.knownIndex(k);
      if (at >= within) break;
      if (xs.at(at) === null) return ChartNumbers.counting(ys.length);
    }
  }
  return xs;
}

/**
 * The numbers a chart is drawn from. Where a reference of a series can be
 * followed to cells of the workbook the cells as they are now are taken:
 * hidden rows and columns are left out when the chart says so, and a cell
 * without a number is what the chart makes of empty cells. Otherwise the
 * values the file cached are taken. What is kept is in proportion to the
 * numbers that are found, not to the points a series says it has. Never
 * throws for what a chart states.
 */
export function chartValues(chart: Chart, options: ChartDataOptions = {}): ChartValues {
  const series: ChartSeriesValues[] = [];
  let labels: Labels | undefined;
  let count = 0;
  let position = 0;
  const budget: Budget = { left: MAX_CHART_CELLS, held: 0 };
  for (let g = 0; g < chart.groups.length; g++) {
    const group = chart.groups[g];
    for (const item of group.series) {
      position++;
      const numbers = numbersOf(chart, item.values, options, budget);
      let name = item.name ? cellText(item.name.ref, options, chart.live === true) : undefined;
      if (name === undefined) name = item.name && item.name.text !== undefined ? item.name.text : options.seriesName ? options.seriesName(position) : `Series${position}`;
      const entry: ChartSeriesValues = { group: g, kind: group.kind, series: item, name, values: numbers.values, format: numbers.format, live: numbers.live };
      if (group.kind === 'scatter' || group.kind === 'bubble') {
        // A point without a number for x is not made zero, whatever the chart makes of empty cells.
        const xs = item.categories && item.categories.kind === 'num' ? numbersOf(chart, item.categories, options, budget, false).values : undefined;
        entry.xs = placesOf(xs, numbers.values);
        if (group.kind === 'bubble' && item.bubbleSizes) entry.sizes = numbersOf(chart, item.bubbleSizes, options, budget, false).values;
      } else if (!labels && item.categories) labels = labelsOf(chart, item.categories, options, budget, categoryAxis(chart, group));
      if (Array.isArray(item.errorBars) && item.errorBars.some((bars) => bars && bars.valueType === 'cust')) {
        // The amounts a series states point by point; a point without one has no bar.
        entry.errors = item.errorBars.map((bars) =>
          bars && bars.valueType === 'cust'
            ? { plus: bars.plus ? numbersOf(chart, bars.plus, options, budget, false).values : undefined, minus: bars.minus ? numbersOf(chart, bars.minus, options, budget, false).values : undefined }
            : undefined,
        );
      }
      if (numbers.values.length > count) count = numbers.values.length;
      series.push(entry);
    }
  }
  const known = labels;
  const out: ChartValues = {
    series,
    count,
    label: (index) => {
      if (!(index >= 0 && index < count)) return '';
      if (!known) return String(index + 1);
      try {
        return index < known.count ? known.label(index) : '';
      } catch (error) {
        return '';
      }
    },
    size: budget.held,
  };
  if (known) {
    const { outer, number } = known;
    if (outer) {
      out.outerLabel = (index) => {
        if (!(index >= 0 && index < count && index < known.count)) return '';
        try {
          return outer(index);
        } catch (error) {
          return '';
        }
      };
    }
    out.categoryNumber = (index) => (index >= 0 && index < count && index < known.count ? number(index) : null);
    if (known.format) out.categoryFormat = known.format;
  }
  const title = chart.title ? cellText(chart.title.ref, options, chart.live === true) : undefined;
  if (title !== undefined) out.title = title;
  return out;
}

// ----------------------------------------------------------------------- scale

export interface ChartScale {
  min: number;
  max: number;
  /** The distance between two ticks; on a logarithmic scale the number of powers between them. */
  step: number;
  /** The values that get a tick and a label, from `min` upwards. */
  ticks: number[];
  /** A logarithmic scale: the base its ticks are powers of. Equal factors are then equal distances. */
  log?: number;
}

export interface ChartScaleOptions {
  /** What the axis states; anything else is worked out. */
  min?: number;
  max?: number;
  step?: number;
  /** The most steps there is room for (2..10 is usual). Default 8. */
  maxSteps?: number;
}

/** Takes the noise out of a sum of floating point steps: 0.30000000000000004 is 0.3. */
function tidy(value: number): number {
  return value === 0 || !Number.isFinite(value) ? 0 : Number(value.toPrecision(12));
}

/** The smallest of 1, 2, 5 times a power of ten that is at least `raw`. */
function niceStep(raw: number): number {
  if (!(raw > 0) || !Number.isFinite(raw)) return 1;
  const power = Math.pow(10, Math.floor(Math.log(raw) / Math.LN10));
  const share = raw / power;
  return tidy((share <= 1 ? 1 : share <= 2 ? 2 : share <= 5 ? 5 : 10) * power);
}

/** Numbers beyond this are taken as this: sums and differences of two of them are still numbers. */
const SCALE_LIMIT = 1e300;
/** A stated step that would give more ticks than this is no step; up to it the step is doubled until the ticks fit. */
const MAX_STATED_STEPS = 10000;

interface ScaleWish {
  min?: number;
  max?: number;
  step?: number;
}

/** A scale for data from `lo` to `hi` with what the axis states; undefined when that gives no axis (no length, no step, numbers that are none). */
function scaleFor(lo: number, hi: number, wish: ScaleWish, maxSteps: number): ChartScale | undefined {
  const statedMin = wish.min;
  const statedMax = wish.max;
  let autoMin: number;
  let autoMax: number;
  if (lo === hi) {
    // One value (or none): from zero to it, and a unit of room when it is zero.
    autoMin = lo < 0 ? lo * 1.05 : 0;
    autoMax = hi > 0 ? hi * 1.05 : lo < 0 ? 0 : 1;
  } else if (lo >= 0) {
    autoMin = lo > 0 && (hi - lo) / hi < 1 / 6 ? lo - (hi - lo) / 2 : 0;
    autoMax = hi + 0.05 * (hi - autoMin);
  } else if (hi <= 0) {
    autoMax = hi < 0 && (hi - lo) / -lo < 1 / 6 ? hi + (hi - lo) / 2 : 0;
    autoMin = lo - 0.05 * (autoMax - lo);
  } else {
    autoMin = lo - 0.05 * (hi - lo);
    autoMax = hi + 0.05 * (hi - lo);
  }
  let from = statedMin !== undefined ? statedMin : autoMin;
  let to = statedMax !== undefined ? statedMax : autoMax;
  if (!(to > from)) {
    // What the axis states leaves the data outside: a range of its own beside the stated end.
    if (statedMin !== undefined) to = from + Math.max(Math.abs(from), 1);
    else from = to - Math.max(Math.abs(to), 1);
  }
  // A range too small beside the numbers themselves has no steps that can be told apart.
  const size = Math.max(Math.abs(from), Math.abs(to));
  if (to - from < size * 1e-9) {
    if (statedMin !== undefined && statedMax !== undefined) return undefined;
    if (statedMax !== undefined) from = to - size * 1e-9;
    else to = from + size * 1e-9;
  }
  if (!Number.isFinite(from) || !Number.isFinite(to) || !(to > from)) return undefined;
  let step = wish.step !== undefined ? wish.step : niceStep((to - from) / maxSteps);
  if (!(step > 0) || !Number.isFinite(step)) return undefined;
  // A stated step somewhat too small for the range is doubled until its ticks fit; one far too small is none.
  if (!((to - from) / step <= MAX_STATED_STEPS)) return undefined;
  const min = statedMin !== undefined ? statedMin : tidy(Math.floor(tidy(from / step)) * step);
  let max = statedMax !== undefined ? statedMax : tidy(Math.ceil(tidy(to / step)) * step);
  if (!(max > min)) max = tidy(min + step);
  if (!Number.isFinite(min) || !Number.isFinite(max) || !(max > min)) return undefined;
  if (!((max - min) / step <= MAX_STATED_STEPS)) return undefined;
  for (let i = 0; i < 64 && (max - min) / step > MAX_AXIS_TICKS; i++) step *= 2;
  // A step that does not move the numbers it is added to.
  if (!Number.isFinite(step) || !(min + step > min) || !(max - step < max)) return undefined;
  const ticks: number[] = [];
  for (let i = 0; i <= MAX_AXIS_TICKS; i++) {
    const value = tidy(min + i * step);
    if (!Number.isFinite(value) || value > max + step * 1e-9) break;
    if (ticks.length > 0 && !(value > ticks[ticks.length - 1])) break;
    ticks.push(value);
  }
  if (ticks.length === 0) return undefined;
  return { min, max, step, ticks };
}

/**
 * The scale of a value axis for data between `low` and `high`. Whatever goes
 * in, what comes out are finite numbers with `max` above `min`, a step above
 * zero and at most MAX_AXIS_TICKS + 1 ticks.
 *
 * What the axis states is taken as it is where it makes an axis. A stated end
 * that is no finite number, a minimum that is not below the maximum, and a
 * step that is not above zero or far too small for the range are not: the
 * axis then works those out by itself, as if nothing were stated.
 *
 * UNSURE (close to what Excel does, not checked against it): data of one sign
 * starts at zero, unless it lies within the top sixth of its largest value,
 * in which case the axis starts half a range below the smallest; 5 % of the
 * range is left free at the open end; both ends are multiples of the step,
 * which is 1, 2 or 5 times a power of ten.
 */
export function chartScale(low: number, high: number, options: ChartScaleOptions = {}): ChartScale {
  // Numbers too large to add up are drawn at the largest ones that can be.
  const bounded = (value: number): number => (Number.isFinite(value) ? clamp(value, -SCALE_LIMIT, SCALE_LIMIT) : 0);
  let lo = bounded(low);
  let hi = bounded(high);
  if (lo > hi) {
    const swap = lo;
    lo = hi;
    hi = swap;
  }
  const end = (value: number | undefined): number | undefined => (value !== undefined && Number.isFinite(value) && Math.abs(value) <= SCALE_LIMIT ? value : undefined);
  let min = end(options.min);
  let max = end(options.max);
  if (min !== undefined && max !== undefined && !(max > min)) {
    min = undefined;
    max = undefined;
  }
  const step = options.step !== undefined && options.step > 0 && Number.isFinite(options.step) ? options.step : undefined;
  const maxSteps = options.maxSteps !== undefined && options.maxSteps >= 1 ? Math.min(Math.floor(options.maxSteps), 50) : 8;
  return (
    scaleFor(lo, hi, { min, max, step }, maxSteps) ||
    // The stated step gives no axis: without it.
    (step !== undefined ? scaleFor(lo, hi, { min, max }, maxSteps) : undefined) ||
    // The stated ends give none either.
    (min !== undefined || max !== undefined ? scaleFor(lo, hi, {}, maxSteps) : undefined) || { min: 0, max: 1, step: 0.2, ticks: [0, 0.2, 0.4, 0.6, 0.8, 1] }
  );
}

/**
 * The scale of a logarithmic value axis for data between `low` and `high`: from the power of
 * `base` at or below the smallest value to the one at or above the largest, with a tick at
 * every power, or at every n-th where there is no room for all. Only numbers above zero are
 * taken, of the data and of what the axis states (`min`, `max`; a stated step is not taken).
 * Whatever goes in, what comes out are finite numbers above zero with `max` above `min`.
 * A base that is none (below 2, above 1000) is ten.
 */
export function chartLogScale(low: number, high: number, base: number, options: ChartScaleOptions = {}): ChartScale {
  const b = Number.isFinite(base) && base >= 2 && base <= 1000 ? base : 10;
  const ln = Math.log(b);
  // Powers beyond these are no numbers that can be told apart, or no numbers at all.
  const lowest = Math.ceil(Math.log(1e-300) / ln);
  const highest = Math.floor(Math.log(1e300) / ln);
  const exponent = (value: number, up: boolean): number => {
    const raw = Math.log(value) / ln;
    const whole = Math.round(raw);
    // A value that is a power, but for the noise of its logarithm, is that power.
    return clamp(Math.abs(raw - whole) < 1e-9 ? whole : up ? Math.ceil(raw) : Math.floor(raw), lowest, highest);
  };
  const positive = (value: number | undefined): number | undefined => (typeof value === 'number' && Number.isFinite(value) && value > 0 ? value : undefined);
  let lo = positive(low);
  let hi = positive(high);
  if (lo !== undefined && hi !== undefined && lo > hi) {
    const swap = lo;
    lo = hi;
    hi = swap;
  }
  if (lo === undefined) lo = hi;
  if (hi === undefined) hi = lo;
  // Ends beyond the powers that can be told apart are no ends.
  const end = (value: number | undefined): number | undefined => (value !== undefined && value >= 1e-300 && value <= 1e300 ? value : undefined);
  let statedMin = end(positive(options.min));
  let statedMax = end(positive(options.max));
  if (statedMin !== undefined && statedMax !== undefined && !(statedMax > statedMin)) {
    statedMin = undefined;
    statedMax = undefined;
  }
  let from = statedMin !== undefined ? exponent(statedMin, true) : lo === undefined ? 0 : exponent(lo, false);
  let to = statedMax !== undefined ? exponent(statedMax, false) : hi === undefined ? 1 : exponent(hi, true);
  let min = statedMin !== undefined ? statedMin : Math.pow(b, from);
  let max = statedMax !== undefined ? statedMax : Math.pow(b, to);
  if (!(max > min) || !Number.isFinite(max) || !(min > 0)) {
    // One value, or ends that leave the data outside: one power beside the end that is there.
    if (statedMax !== undefined && statedMin === undefined) {
      to = exponent(statedMax, false);
      from = to - 1;
      max = statedMax;
      min = Math.pow(b, from) < max ? Math.pow(b, from) : max / b;
    } else {
      from = statedMin !== undefined ? exponent(statedMin, true) : clamp(from, lowest, highest - 1);
      to = from + 1;
      min = statedMin !== undefined ? statedMin : Math.pow(b, from);
      max = Math.pow(b, to) > min ? Math.pow(b, to) : min * b;
    }
  }
  const maxSteps = options.maxSteps !== undefined && options.maxSteps >= 1 ? Math.min(Math.floor(options.maxSteps), 50) : 8;
  const every = Math.max(Math.ceil((to - from) / maxSteps), 1);
  const ticks: number[] = [];
  for (let e = from; e <= to && ticks.length <= MAX_AXIS_TICKS; e += every) {
    const value = Math.pow(b, e);
    if (Number.isFinite(value) && value >= min * (1 - 1e-12) && value <= max * (1 + 1e-12) && (ticks.length === 0 || value > ticks[ticks.length - 1])) ticks.push(e < 0 ? tidy(value) : value);
  }
  if (!Number.isFinite(min) || !Number.isFinite(max) || !(min > 0) || !(max > min)) return { min: 1, max: b, step: 1, ticks: [1, b], log: b };
  // No power between the ends that are stated: the ends themselves.
  if (ticks.length < 2) return { min, max, step: every, ticks: [min, max], log: b };
  return { min, max, step: every, ticks, log: b };
}

/** How far along its axis a value lies: 0 at the minimum of the scale, 1 at the maximum. */
function shareOf(scale: ChartScale, value: number): number {
  if (scale.log) {
    const low = Math.log(scale.min);
    const span = Math.log(scale.max) - low;
    return value > 0 && span > 0 ? (Math.log(value) - low) / span : 0;
  }
  return (value - scale.min) / (scale.max - scale.min);
}

// ---------------------------------------------------------------------- layout

/** A font of a chart: the size in pixels. */
export interface ChartFont {
  size: number;
  bold: boolean;
  italic: boolean;
  /** Font name as the file states it; the painter makes it safe. */
  family?: string;
}

export type ChartMeasure = (text: string, font: ChartFont) => number;

export type ChartItem =
  | { kind: 'rect'; role: 'frame' | 'plot' | 'bar' | 'swatch' | 'updown'; x: number; y: number; width: number; height: number; fill?: string; stroke?: string; lineWidth?: number; radius?: number; series?: number; point?: number }
  | {
      kind: 'path';
      role: 'grid' | 'axis' | 'line' | 'area' | 'swatch' | 'trend' | 'error' | 'drop' | 'hilow';
      /** x and y in turn. With `curve`: the first point, then two control points and an end point of a cubic Bézier piece for every further one. */
      points: number[];
      closed?: boolean;
      curve?: boolean;
      /** Every two points are a line of their own. */
      segments?: boolean;
      fill?: string;
      stroke?: string;
      lineWidth?: number;
      dash?: boolean;
      series?: number;
    }
  | { kind: 'markers'; role: 'marker' | 'swatch'; symbol: string; size: number; points: number[]; fill?: string; stroke?: string; series?: number }
  /** Filled circles of sizes of their own: x, y and the radius in turn. */
  | { kind: 'circles'; role: 'bubble'; points: number[]; fill?: string; stroke?: string; lineWidth?: number; series?: number }
  | { kind: 'slice'; cx: number; cy: number; radius: number; inner: number; start: number; end: number; fill: string; stroke?: string; lineWidth?: number; series: number; point: number }
  | {
      kind: 'text';
      role: 'title' | 'tick' | 'category' | 'axisTitle' | 'legend' | 'label';
      text: string;
      x: number;
      y: number;
      font: ChartFont;
      color: string;
      align: 'left' | 'center' | 'right';
      /** Read from the bottom up, around (x, y). */
      rotated?: boolean;
      /** Turned around (x, y) by so many degrees clockwise; the text starts, ends or is centred there as `align` says. */
      angle?: number;
    };

export interface ChartLayout {
  width: number;
  height: number;
  /** "full": everything; "compact": no words; "bare": the series alone. */
  detail: 'full' | 'compact' | 'bare';
  /** Corner radius of the frame, pixels. */
  radius: number;
  /** The area the series are drawn in. */
  plot: { x: number; y: number; width: number; height: number };
  /** The scale of the value axis (the vertical one of a scatter chart); undefined for a pie. */
  scale?: ChartScale;
  /** The scale of the horizontal axis of a scatter chart. */
  xScale?: ChartScale;
  /** The scale of the second value axis, when series are drawn against one. */
  scale2?: ChartScale;
  /** What is drawn, in the order it is drawn. */
  items: ChartItem[];
}

export interface ChartLayoutOptions {
  locale?: FormatLocale;
  /** Dates of the workbook count from 1904. */
  date1904?: boolean;
}

const HEX_COLOR = /^#[0-9A-Fa-f]{6}$/;
const TEXT_COLOR = '#595959';
const GRID_COLOR = '#D9D9D9';
const MINOR_GRID_COLOR = '#F2F2F2';
const REST_COLOR = '#BFBFBF';
/** Lighter and darker forms the accent colours are given once all six are used: [share kept, share of white added]. */
const TONES: [number, number][] = [
  [1, 0],
  [0.6, 0],
  [0.8, 0.2],
  [0.8, 0],
  [0.6, 0.4],
  [0.5, 0],
  [0.7, 0.3],
  [0.7, 0],
  [0.5, 0.5],
];

const safe = (color: string | null | undefined): string | undefined => (typeof color === 'string' && HEX_COLOR.test(color) ? color : undefined);
const clamp = (value: number, low: number, high: number): number => (value < low ? low : value > high ? high : value);

/**
 * The colour the chart gives the n-th series (or point) that states none: the
 * accent colours of the theme in turn, then lighter and darker forms of them.
 * UNSURE: the order of the forms is that of the default colours of a new
 * chart; the style number of the chart, which may ask for others, is not applied.
 */
export function chartAutoColor(chart: Pick<Chart, 'palette'>, index: number): string {
  const palette = chart.palette || [];
  const at = index >= 0 && Number.isFinite(index) ? Math.floor(index) : 0;
  const base = safe(palette[at % 6]) || '#808080';
  const tone = TONES[Math.floor(at / 6) % TONES.length];
  let out = '#';
  for (let i = 0; i < 3; i++) {
    const channel = Math.round(clamp(parseInt(base.slice(1 + i * 2, 3 + i * 2), 16) * tone[0] + 255 * tone[1], 0, 255));
    out += (channel < 16 ? '0' : '') + channel.toString(16).toUpperCase();
  }
  return out;
}

function fontOf(style: ChartTextStyle | undefined, base: ChartTextStyle | undefined, points: number, bold = false): ChartFont {
  const size = (style && style.size) || (base && base.size) || points;
  const family = (style && style.font) || (base && base.font);
  return {
    size: Math.round(((clamp(size, 4, 96) * 96) / 72) * 100) / 100,
    bold: style && style.bold !== undefined ? style.bold : base && base.bold !== undefined ? base.bold : bold,
    italic: style && style.italic !== undefined ? style.italic : base && base.italic !== undefined ? base.italic : false,
    family: typeof family === 'string' ? family : undefined,
  };
}

const colorOf = (style: ChartTextStyle | undefined, base: ChartTextStyle | undefined): string => safe(style && style.color) || safe(base && base.color) || TEXT_COLOR;

/** One line without control characters, of a length that can be measured. */
function clean(text: string): string {
  return (text.length > 300 ? text.slice(0, 300) : text).replace(/[\u0000-\u001f\u007f]+/g, ' ');
}

/** `text` cut so that it is no wider than `room`, with an ellipsis; '' when not even that fits. */
function fit(text: string, font: ChartFont, room: number, measure: ChartMeasure): string {
  if (text === '' || !(room > 0)) return '';
  if (measure(text, font) <= room) return text;
  let low = 0;
  let high = text.length - 1;
  while (low < high) {
    const mid = (low + high + 1) >> 1;
    if (measure(text.slice(0, mid) + '…', font) <= room) low = mid;
    else high = mid - 1;
  }
  return low > 0 ? text.slice(0, low) + '…' : '';
}

interface Box {
  x: number;
  y: number;
  width: number;
  height: number;
}

interface LegendEntry {
  label: string;
  color: string;
  kind: 'box' | 'line' | 'marker';
  symbol?: string;
}

/** What a series is drawn with, once the chart's own colours stand in for what the file leaves open. */
interface SeriesLook {
  color: string;
  fill: string | undefined;
  stroke: string | undefined;
  lineWidth: number;
  dash: boolean;
  marker: { symbol: string; size: number; fill: string | undefined; stroke: string | undefined } | undefined;
  points: Map<number, ChartPoint> | undefined;
}

function pointsOf(series: ChartSeries): Map<number, ChartPoint> | undefined {
  if (!series.points || series.points.length === 0) return undefined;
  const map = new Map<number, ChartPoint>();
  for (const point of series.points) if (point && !map.has(point.index)) map.set(point.index, point);
  return map;
}

function lineColor(line: ChartLine | undefined, automatic: string | undefined): string | undefined {
  if (!line || line.color === undefined) return automatic;
  return line.color === null ? undefined : safe(line.color) || automatic;
}

function markerOf(marker: ChartMarker | undefined, shown: boolean, color: string): SeriesLook['marker'] {
  if (marker ? marker.symbol === 'none' : !shown) return undefined;
  // UNSURE: a marker the file leaves to the chart is a circle for every series (older versions of Excel went through a row of shapes).
  const symbol = marker && marker.symbol !== 'auto' ? marker.symbol : 'circle';
  const size = Math.round(((clamp((marker && marker.size) || 5, 2, 72) * 96) / 72) * 100) / 100;
  const fill = marker && marker.fill !== undefined ? (marker.fill === null ? undefined : safe(marker.fill) || color) : color;
  return { symbol, size, fill, stroke: lineColor(marker ? marker.line : undefined, fill === undefined ? color : undefined) };
}

function lookOf(chart: Chart, group: ChartGroup, series: ChartSeries): SeriesLook {
  const color = chartAutoColor(chart, series.index);
  const shape = series.shape;
  const stated = shape ? shape.fill : undefined;
  const filled = group.kind === 'bar' || group.kind === 'area' || group.kind === 'pie' || group.kind === 'doughnut' || group.kind === 'bubble' || (group.kind === 'radar' && group.radarStyle === 'filled');
  const line = shape ? shape.line : undefined;
  let stroke: string | undefined;
  let marker: SeriesLook['marker'];
  if (group.kind === 'line') {
    stroke = lineColor(line, color);
    marker = markerOf(series.marker, group.markers === true, stroke || color);
  } else if (group.kind === 'radar') {
    stroke = lineColor(line, filled ? undefined : color);
    marker = markerOf(series.marker, group.radarStyle === 'marker', stroke || color);
  } else if (group.kind === 'stock') {
    // UNSURE: the series of a stock chart have neither a line nor markers unless the file gives them one (Excel writes both out).
    stroke = lineColor(line, undefined);
    marker = markerOf(series.marker, false, color);
  } else if (group.kind === 'scatter') {
    const style = group.scatterStyle || 'marker';
    const lined = style === 'lineMarker' || style === 'line' || style === 'smooth' || style === 'smoothMarker';
    stroke = lined ? lineColor(line, color) : undefined;
    marker = markerOf(series.marker, style === 'lineMarker' || style === 'marker' || style === 'smoothMarker', stroke || color);
  } else stroke = lineColor(line, undefined);
  const width = line && line.width !== undefined && line.width > 0 ? line.width : filled ? 1 : 3;
  return {
    color: filled ? (stated === undefined || stated === null ? color : safe(stated) || color) : stroke || (marker && marker.fill) || color,
    fill: filled ? (stated === null ? undefined : safe(stated) || color) : undefined,
    stroke,
    lineWidth: clamp(width, 0.5, 32),
    dash: line !== undefined && line.dash === true,
    marker,
    points: pointsOf(series),
  };
}

/** Fill and outline of one point: what the point states, else what its series has (or, with a colour per point, the chart's n-th colour). */
function pointLook(chart: Chart, look: SeriesLook, index: number, vary: boolean): { fill: string | undefined; stroke: string | undefined; lineWidth: number } {
  const own = look.points ? look.points.get(index) : undefined;
  const shape = own ? own.shape : undefined;
  let fill = vary ? chartAutoColor(chart, index) : look.fill;
  if (shape && shape.fill !== undefined) fill = shape.fill === null ? undefined : safe(shape.fill) || fill;
  let stroke = look.stroke;
  let lineWidth = look.lineWidth;
  if (shape && shape.line) {
    stroke = lineColor(shape.line, stroke);
    if (shape.line.width !== undefined && shape.line.width > 0) lineWidth = clamp(shape.line.width, 0.5, 32);
  }
  return { fill, stroke, lineWidth };
}

/** The largest and smallest value of each of `buckets` runs of the `count` points of a chart, for one series. Costs what the series holds, not what it says it has. */
function sample(values: ChartNumbers, count: number, buckets: number): { low: (number | null)[]; high: (number | null)[]; peak: (number | null)[] } {
  const low: (number | null)[] = [];
  const high: (number | null)[] = [];
  const found: number[] = [];
  for (let b = 0; b < buckets; b++) {
    low.push(null);
    high.push(null);
    found.push(0);
  }
  const from = (b: number): number => Math.floor((b * count) / buckets);
  const to = (b: number): number => Math.max(from(b + 1), from(b) + 1);
  let bucket = 0;
  values.each((value, index) => {
    if (index >= count) return;
    while (bucket < buckets - 1 && index >= to(bucket)) bucket++;
    if (index < from(bucket) || index >= to(bucket)) return;
    const lo = low[bucket];
    const hi = high[bucket];
    if (lo === null || value < lo) low[bucket] = value;
    if (hi === null || value > hi) high[bucket] = value;
    found[bucket]++;
  });
  const peak: (number | null)[] = [];
  const blank = values.blank;
  for (let b = 0; b < buckets; b++) {
    // Points without a number that are drawn as zero are part of their run.
    if (blank !== null && found[b] < Math.min(to(b), values.length) - from(b)) {
      const lo = low[b];
      const hi = high[b];
      if (lo === null || blank < lo) low[b] = blank;
      if (hi === null || blank > hi) high[b] = blank;
    }
    const lo = low[b];
    const hi = high[b];
    peak.push(lo === null || hi === null ? null : Math.abs(lo) > Math.abs(hi) ? lo : hi);
  }
  return { low, high, peak };
}

/** The numbers of the first `places` points of a series, one place each. */
function spread(values: ChartNumbers, places: number): (number | null)[] {
  const out: (number | null)[] = [];
  if (values.blank === null && values.known * 8 < places) {
    // Few numbers among many points: the places are filled from the numbers.
    for (let i = 0; i < places; i++) out.push(null);
    values.each((value, index) => {
      if (index < places) out[index] = value;
    });
    return out;
  }
  for (let i = 0; i < places; i++) out.push(values.at(i));
  return out;
}

function tickText(value: number, format: { code: string; builtinId?: number } | undefined, options: ChartLayoutOptions): string {
  const code = format && format.code !== '' ? format.code : 'General';
  try {
    return clean(formatValue(tidy(value), code, { locale: options.locale, date1904: options.date1904, builtinId: format ? format.builtinId : undefined }).text);
  } catch (error) {
    return String(tidy(value));
  }
}

interface Context {
  chart: Chart;
  values: ChartValues;
  measure: ChartMeasure;
  options: ChartLayoutOptions;
  items: ChartItem[];
  detail: ChartLayout['detail'];
  /** The size of the frame. */
  width: number;
  height: number;
  /** Columns, points of lines and markers that may be drawn, all series together: MAX_DRAWN_ITEMS, or less for a chart that weighs too much with that. */
  budget: number;
}

/** What a layout holds, counted in numbers: eight for each item and one for each coordinate of its points. */
export function chartLayoutWeight(layout: ChartLayout): number {
  let weight = 0;
  for (const item of layout.items) weight += 8 + (item.kind === 'path' || item.kind === 'markers' || item.kind === 'circles' ? item.points.length : 0);
  return weight;
}

/**
 * The place a file sets by hand as a box of the frame; undefined for what is no place: one that is
 * not counted from the corner of the frame ("edge", for both directions), that lies outside the
 * frame, or whose size is none. A place without a size has a width and a height of -1, unless a
 * size is asked for (`sized`).
 */
function manualBox(manual: ChartManualLayout | undefined, width: number, height: number, sized: boolean): Box | undefined {
  if (!manual || manual.xMode !== 'edge' || manual.yMode !== 'edge') return undefined;
  const share = (value: number | undefined): value is number => typeof value === 'number' && Number.isFinite(value) && value >= 0 && value <= 1;
  const { x, y } = manual;
  if (!share(x) || !share(y)) return undefined;
  // UNSURE: a width or a height that is stated as an edge is the place of the far edge.
  const w = manual.wMode === 'edge' && typeof manual.w === 'number' ? manual.w - x : manual.w;
  const h = manual.hMode === 'edge' && typeof manual.h === 'number' ? manual.h - y : manual.h;
  if (sized && (w === undefined || h === undefined)) return undefined;
  if (w !== undefined && (!share(w) || !(w > 0) || x + w > 1 + 1e-9)) return undefined;
  if (h !== undefined && (!share(h) || !(h > 0) || y + h > 1 + 1e-9)) return undefined;
  return { x: x * width, y: y * height, width: w === undefined ? -1 : w * width, height: h === undefined ? -1 : h * height };
}

/**
 * Where everything of a chart is drawn in a frame of `width` by `height`
 * pixels, from (0, 0). A frame too small for words shows the series with
 * their axes, and a still smaller one the series alone. Never throws for
 * what a chart states: sizes that are no sizes give an empty layout.
 */
export function layoutChart(chart: Chart, values: ChartValues, width: number, height: number, measure: ChartMeasure, options: ChartLayoutOptions = {}): ChartLayout {
  let budget = MAX_DRAWN_ITEMS;
  let layout = layoutWithin(chart, values, width, height, measure, options, budget);
  // What a point carries besides itself (error bars, a curve, a label) may make a layout of very many points weigh
  // more than a frame is worth: it is then made again with fewer positions of each series.
  for (let i = 0; i < 10 && budget > 256 && chartLayoutWeight(layout) > MAX_LAYOUT_WEIGHT; i++) {
    budget = Math.floor(budget / 2);
    layout = layoutWithin(chart, values, width, height, measure, options, budget);
  }
  return layout;
}

function layoutWithin(chart: Chart, values: ChartValues, width: number, height: number, measure: ChartMeasure, options: ChartLayoutOptions, budget: number): ChartLayout {
  const w = Number.isFinite(width) && width > 0 ? Math.min(width, 1e5) : 0;
  const h = Number.isFinite(height) && height > 0 ? Math.min(height, 1e5) : 0;
  const detail: ChartLayout['detail'] = w < 80 || h < 60 ? 'bare' : w < 180 || h < 120 ? 'compact' : 'full';
  const items: ChartItem[] = [];
  const layout: ChartLayout = { width: w, height: h, detail, radius: chart.roundedCorners ? Math.min(8, w / 4, h / 4) : 0, plot: { x: 0, y: 0, width: w, height: h }, items };
  if (w === 0 || h === 0) return layout;
  const context: Context = { chart, values, measure, options, items, detail, width: w, height: h, budget };

  // The frame: white with a thin grey outline unless the file says otherwise.
  const area = chart.area;
  const frameFill = area && area.fill !== undefined ? (area.fill === null ? undefined : safe(area.fill)) : '#FFFFFF';
  const frameStroke = lineColor(area ? area.line : undefined, GRID_COLOR);
  items.push({ kind: 'rect', role: 'frame', x: 0, y: 0, width: w, height: h, fill: frameFill, stroke: frameStroke, lineWidth: clamp((area && area.line && area.line.width) || 1, 0.5, 8), radius: layout.radius });

  const pad = detail === 'full' ? 8 : detail === 'compact' ? 4 : Math.min(2, w / 8, h / 8);
  const box: Box = { x: pad, y: pad, width: w - pad * 2, height: h - pad * 2 };

  const pies = chart.groups.some((group) => group.kind === 'pie' || group.kind === 'doughnut');
  // A pie or a radar beside other chart types is drawn alone: they share no axes with them.
  const radar = !pies && chart.groups.some((group) => group.kind === 'radar');
  const shown = values.series.filter((entry) => (pies ? entry.kind === 'pie' || entry.kind === 'doughnut' : radar ? entry.kind === 'radar' : true));

  if (detail === 'full') {
    placeTitle(context, box, shown);
    placeLegend(context, box, shown, pies);
  }
  // A plot area the file sets by hand: with the labels of its axes ("outer", and where the file does not say), or the
  // area of the series alone ("inner"). A place too small to draw in is none.
  let place = manualBox(chart.plotLayout, w, h, true);
  if (place && (place.width < 10 || place.height < 10)) place = undefined;
  const inner = place !== undefined && chart.plotLayout !== undefined && chart.plotLayout.target === 'inner';
  if (pies) layoutPie(context, layout, place || box, shown);
  else if (radar) layoutRadar(context, layout, place || box, shown);
  else layoutAxes(context, layout, place && !inner ? place : box, shown, inner ? place : undefined);
  return layout;
}

function placeTitle(context: Context, box: Box, shown: ChartSeriesValues[]): void {
  const { chart, values, measure, items } = context;
  const title = chart.title;
  if (!title) return;
  let text = values.title !== undefined ? values.title : title.text;
  // UNSURE: a title without a text of its own is the name of the series when the chart has one series; a chart with
  // several shows "Chart Title" in Excel, which is nothing worth showing. A chart without a title element gets none.
  if (text === undefined && shown.length === 1) text = shown[0].name;
  if (text === undefined || text === '') return;
  const font = fontOf(title.style, chart.textStyle, 14);
  const color = colorOf(title.style, chart.textStyle);
  const lines = text.split('\n').slice(0, 3);
  const lineHeight = font.size * 1.2;
  const taken = lines.length * lineHeight + 6;
  if (taken > box.height * 0.4) return;
  // A place set by hand: counted from the corner of the frame ("edge"), or moved from where the title stands by
  // itself ("factor", which is what a direction without a mode means). UNSURE: the room the title takes at the top
  // stays where it is - the plot area does not follow the title.
  const place = title.layout;
  const stated = (value: number | undefined, mode: string | undefined): boolean =>
    mode === 'edge' ? typeof value === 'number' && Number.isFinite(value) && value >= 0 && value <= 1 : value === undefined || (Number.isFinite(value) && Math.abs(value) <= 1);
  const moved = place !== undefined && stated(place.x, place.xMode) && stated(place.y, place.yMode) ? place : undefined;
  const left = moved && moved.xMode === 'edge' ? (moved.x as number) * context.width : undefined;
  const top = moved && moved.yMode === 'edge' ? Math.min((moved.y as number) * context.height, Math.max(context.height - lines.length * lineHeight, 0)) : box.y + (moved ? (moved.y || 0) * context.height : 0);
  const middle = box.x + box.width / 2 + (moved && left === undefined ? (moved.x || 0) * context.width : 0);
  for (let i = 0; i < lines.length; i++) {
    const line = fit(clean(lines[i]), font, left !== undefined ? context.width : box.width, measure);
    if (line === '') continue;
    const y = clamp(top + i * lineHeight + font.size, font.size, Math.max(context.height, font.size));
    // A title placed by its left edge is kept whole inside the frame.
    if (left !== undefined) items.push({ kind: 'text', role: 'title', text: line, x: Math.max(Math.min(left, context.width - measure(line, font)), 0), y, font, color, align: 'left' });
    else {
      // A title that is moved stays whole inside the frame too.
      const half = moved ? Math.min(measure(line, font), context.width) / 2 : 0;
      items.push({ kind: 'text', role: 'title', text: line, x: clamp(middle, half, context.width - half), y: moved ? Math.min(y, Math.max(context.height - (lines.length - 1 - i) * lineHeight, font.size)) : y, font, color, align: 'center' });
    }
  }
  box.y += taken;
  box.height -= taken;
}

function legendEntries(context: Context, shown: ChartSeriesValues[], pies: boolean): LegendEntry[] {
  const { chart, values } = context;
  const entries: LegendEntry[] = [];
  const first = shown[0];
  const perPoint = first !== undefined && (pies || (shown.length === 1 && first.kind === 'bar' && chart.groups[first.group].varyColors === true));
  if (perPoint) {
    // One entry per point: the slices of a pie, the columns of a series whose points differ in colour.
    const group = chart.groups[first.group];
    const look = lookOf(chart, group, first.series);
    const count = Math.min(first.values.length, MAX_LEGEND_ENTRIES);
    for (let i = 0; i < count; i++) entries.push({ label: clean(values.label(i)), color: pointLook(chart, look, i, group.varyColors !== false).fill || look.color, kind: 'box' });
  } else {
    for (let i = 0; i < shown.length && entries.length < MAX_LEGEND_ENTRIES; i++) {
      const entry = shown[i];
      const look = lookOf(chart, chart.groups[entry.group], entry.series);
      const lined = entry.kind === 'line' || entry.kind === 'scatter' || entry.kind === 'stock' || (entry.kind === 'radar' && look.fill === undefined);
      entries.push({ label: clean(entry.name), color: look.color, kind: lined ? (look.stroke ? 'line' : 'marker') : 'box', symbol: look.marker ? look.marker.symbol : undefined });
    }
  }
  const deleted = chart.legend && chart.legend.deleted;
  return deleted && deleted.length > 0 ? entries.filter((_entry, index) => deleted.indexOf(index) < 0) : entries;
}

function placeLegend(context: Context, box: Box, shown: ChartSeriesValues[], pies: boolean): void {
  const { chart, measure, items } = context;
  const legend = chart.legend;
  if (!legend) return;
  const entries = legendEntries(context, shown, pies);
  if (entries.length === 0) return;
  const font = fontOf(legend.style, chart.textStyle, 9);
  const color = colorOf(legend.style, chart.textStyle);
  const rowHeight = Math.max(font.size * 1.35, 10);
  const key = (entry: LegendEntry): number => (entry.kind === 'box' ? Math.min(font.size * 0.7, 10) : 22);
  const gap = 4;
  // A place set by hand: the entries are listed in it. UNSURE: the plot area keeps the room the legend would
  // take where it stands by itself.
  const place = manualBox(legend.layout, context.width, context.height, false);
  const draw = (entry: LegendEntry, label: string, x: number, top: number): void => {
    if (!place) paint(entry, label, x, top);
  };
  const paint = (entry: LegendEntry, label: string, x: number, top: number): void => {
    const size = key(entry);
    const middle = top + rowHeight / 2;
    if (entry.kind === 'box') items.push({ kind: 'rect', role: 'swatch', x, y: middle - size / 2, width: size, height: size, fill: entry.color });
    else {
      if (entry.kind === 'line') items.push({ kind: 'path', role: 'swatch', points: [x, middle, x + size, middle], stroke: entry.color, lineWidth: 2 });
      if (entry.symbol !== undefined || entry.kind === 'marker') items.push({ kind: 'markers', role: 'swatch', symbol: entry.symbol || 'circle', size: 6, points: [x + size / 2, middle], fill: entry.color });
    }
    items.push({ kind: 'text', role: 'legend', text: label, x: x + size + gap, y: middle + font.size * 0.35, font, color, align: 'left' });
  };
  const side = legend.position === 'r' || legend.position === 'l' || legend.position === 'tr';
  if (place) {
    // The corner of the place, moved in where the frame would cut the first entry.
    const left = Math.max(Math.min(place.x, context.width - 40), 0);
    const top = Math.max(Math.min(place.y, context.height - rowHeight), 0);
    const width = Math.min(place.width >= 0 ? place.width : context.width, context.width - left);
    const height = Math.min(place.height >= 0 ? place.height : context.height, context.height - top);
    if (side) {
      // One below the other, as many as the height holds.
      const rows = Math.min(entries.length, Math.max(Math.floor(height / rowHeight), 0));
      for (let i = 0; i < rows; i++) {
        const label = fit(entries[i].label, font, width - key(entries[i]) - gap, measure);
        if (label !== '') paint(entries[i], label, left, top + i * rowHeight);
      }
    } else {
      // Side by side, going on in a new row where the width ends.
      let x = left;
      let row = 0;
      for (const entry of entries) {
        const label = fit(entry.label, font, Math.max(width - key(entry) - gap, 0), measure);
        const itemWidth = key(entry) + gap + measure(label, font);
        if (x > left && x + itemWidth > left + width) {
          x = left;
          row++;
        }
        if ((row + 1) * rowHeight > height) break;
        if (label !== '') paint(entry, label, x, top + row * rowHeight);
        x += itemWidth + 14;
      }
    }
  }
  if (side) {
    const room = box.width * 0.35;
    let widest = 0;
    for (const entry of entries) widest = Math.max(widest, key(entry) + gap + measure(entry.label, font));
    const taken = Math.min(widest, room);
    // Not worth the room it takes from the series.
    if (room < 30 || box.width - taken - 8 < 60) return;
    const rows = Math.min(entries.length, Math.max(Math.floor(box.height / rowHeight), 0));
    if (rows === 0) return;
    const x = legend.position === 'l' ? box.x : box.x + box.width - taken;
    const top = legend.position === 'tr' ? box.y : box.y + (box.height - rows * rowHeight) / 2;
    for (let i = 0; i < rows; i++) {
      const label = fit(entries[i].label, font, taken - key(entries[i]) - gap, measure);
      draw(entries[i], label, x, top + i * rowHeight);
    }
    if (legend.position === 'l') box.x += taken + 8;
    box.width -= taken + 8;
    return;
  }
  // Above or below: rows of entries, each row centred.
  const maxRows = Math.max(Math.min(3, Math.floor((box.height * 0.3) / rowHeight)), 0);
  if (maxRows === 0) return;
  const between = 14;
  const rows: { entries: { entry: LegendEntry; label: string; width: number }[]; width: number }[] = [];
  let row: (typeof rows)[number] = { entries: [], width: 0 };
  for (const entry of entries) {
    const label = fit(entry.label, font, Math.max(box.width * 0.45 - key(entry) - gap, 0), measure);
    const itemWidth = key(entry) + gap + measure(label, font);
    if (row.entries.length > 0 && row.width + between + itemWidth > box.width) {
      rows.push(row);
      row = { entries: [], width: 0 };
      if (rows.length >= maxRows) break;
    }
    row.width += (row.entries.length > 0 ? between : 0) + itemWidth;
    row.entries.push({ entry, label, width: itemWidth });
  }
  if (rows.length < maxRows && row.entries.length > 0) rows.push(row);
  if (rows.length === 0) return;
  const taken = rows.length * rowHeight + 6;
  if (box.height - taken < 40) return;
  const top = legend.position === 't' ? box.y : box.y + box.height - rows.length * rowHeight;
  for (let r = 0; r < rows.length; r++) {
    let x = box.x + Math.max((box.width - rows[r].width) / 2, 0);
    for (const item of rows[r].entries) {
      draw(item.entry, item.label, x, top + r * rowHeight);
      x += item.width + between;
    }
  }
  if (legend.position === 't') box.y += taken;
  box.height -= taken;
}

// ----------------------------------------------------------------- data labels

/** Where a data label belongs: the column, the point or the slice it is the label of. */
type LabelAnchor =
  /** `low`: the end of the bar that shows its value is the one at the smaller coordinate (the top of a column that stands). */
  | { kind: 'bar'; x: number; y: number; width: number; height: number; lying: boolean; low: boolean; stacked: boolean }
  /** `centred`: the label lies on the point unless the file says otherwise (areas). */
  | { kind: 'point'; x: number; y: number; radius: number; centred?: boolean }
  | { kind: 'slice'; cx: number; cy: number; inner: number; outer: number; start: number; end: number; ring: boolean };

/** What a label may say besides the value of its point. */
interface LabelFacts {
  /** The share of the point in its series (pies). */
  share?: number;
  /** The horizontal value of the point (scatter and bubble charts): it stands for the category. */
  x?: number;
  /** The size of the bubble. */
  size?: number;
}

/** The labels of one series: what all of its points show, and the points that show something of their own. */
interface SeriesLabels {
  base: ChartDataLabels | undefined;
  own: Map<number, ChartDataLabels> | undefined;
}

/** What a level says over the level above it: the point over the series, the series over the chart type. */
function mergeLabels(above: ChartDataLabels | undefined, own: ChartDataLabels | undefined): ChartDataLabels | undefined {
  if (!own) return above;
  if (own.deleted) return { deleted: true };
  const out: ChartDataLabels = {};
  if (above && !above.deleted) Object.assign(out, above);
  Object.assign(out, own);
  if (above && !above.deleted && above.style && own.style) out.style = { ...above.style, ...own.style };
  return out;
}

const labelShows = (labels: ChartDataLabels | undefined): labels is ChartDataLabels =>
  labels !== undefined && !labels.deleted && (labels.showValue === true || labels.showCategory === true || labels.showSeries === true || labels.showPercent === true || labels.showBubbleSize === true || (typeof labels.text === 'string' && labels.text !== ''));

/**
 * Places the data labels of a chart. A label is a line of text centred on a place beside its
 * point; one that would lie over a label placed before it is left out, one that would leave
 * the frame is moved into it. No lines lead from a label to its point.
 */
class LabelPlacer {
  /** The labels that were placed, to be drawn above everything else. */
  readonly items: ChartItem[] = [];
  /** Left, top, right and bottom of every label placed so far. */
  private readonly boxes: number[] = [];
  private considered = 0;
  private readonly looks = new Map<ChartDataLabels, { font: ChartFont; color: string }>();

  constructor(
    private readonly context: Context,
    private readonly width: number,
    private readonly height: number,
  ) {}

  /** The labels of a series; undefined for a series none of whose points shows one. */
  of(entry: ChartSeriesValues): SeriesLabels | undefined {
    const group = this.context.chart.groups[entry.group];
    const base = mergeLabels(group ? group.dataLabels : undefined, entry.series.dataLabels);
    let own: Map<number, ChartDataLabels> | undefined;
    const points = entry.series.pointLabels;
    if (Array.isArray(points)) {
      for (const point of points) {
        if (!point || !(point.index >= 0)) continue;
        if (!own) own = new Map();
        if (!own.has(point.index)) own.set(point.index, mergeLabels(base, point) as ChartDataLabels);
      }
    }
    return labelShows(base) || own ? { base, own } : undefined;
  }

  /** The place the labels of a series stand at when no point says otherwise. */
  position(labels: SeriesLabels | undefined): string | undefined {
    return labels && labelShows(labels.base) ? labels.base.position : undefined;
  }

  /** Places the label of point `index` of a series, whose value is `value`. */
  add(entry: ChartSeriesValues, labels: SeriesLabels, index: number, value: number, anchor: LabelAnchor, facts: LabelFacts = {}): void {
    const spec = labels.own && labels.own.has(index) ? labels.own.get(index) : labels.base;
    if (!labelShows(spec) || this.considered >= MAX_DATA_LABELS) return;
    const { measure, options, values, chart } = this.context;
    let text: string;
    if (typeof spec.text === 'string' && spec.text !== '') text = clean(spec.text);
    else {
      const parts: string[] = [];
      if (spec.showSeries) parts.push(clean(entry.name));
      if (spec.showCategory) parts.push(facts.x !== undefined ? tickText(facts.x, undefined, options) : clean(values.label(index)));
      const stated = spec.formatCode !== undefined && spec.sourceLinked !== true ? spec.formatCode : undefined;
      if (spec.showValue && Number.isFinite(value)) parts.push(tickText(value, stated !== undefined ? { code: stated } : entry.format, options));
      // The share in the format of the labels when that is one of per cent.
      if (spec.showPercent && facts.share !== undefined && Number.isFinite(facts.share)) parts.push(tickText(facts.share, { code: stated !== undefined && stated.indexOf('%') >= 0 ? stated : '0%' }, options));
      if (spec.showBubbleSize && facts.size !== undefined && Number.isFinite(facts.size)) parts.push(tickText(facts.size, undefined, options));
      // UNSURE: a comma and a space stand between the parts where the file names nothing (Excel takes the list separator of the system).
      text = parts.filter((part) => part !== '').join(typeof spec.separator === 'string' ? clean(spec.separator) : ', ');
    }
    if (text === '') return;
    this.considered++;
    let look = this.looks.get(spec);
    if (!look) {
      look = { font: fontOf(spec.style, chart.textStyle, 9), color: colorOf(spec.style, chart.textStyle) };
      this.looks.set(spec, look);
    }
    const font = look.font;
    let w = measure(text, font);
    const widest = this.width * 0.45;
    if (w > widest) {
      text = fit(text, font, widest, measure);
      if (text === '') return;
      w = measure(text, font);
    }
    const h = font.size * 1.2;
    const gap = 3;
    let cx: number;
    let cy: number;
    const wanted = spec.position;
    if (anchor.kind === 'bar') {
      const place = wanted === 'outEnd' || wanted === 'inEnd' || wanted === 'ctr' || wanted === 'inBase' ? wanted : anchor.stacked ? 'ctr' : 'outEnd';
      // Away from the bar at the end that shows its value.
      const out = anchor.low ? -1 : 1;
      const half = (anchor.lying ? w : h) / 2 + gap;
      const start = anchor.lying ? anchor.x : anchor.y;
      const length = anchor.lying ? anchor.width : anchor.height;
      const end = anchor.low ? start : start + length;
      const foot = anchor.low ? start + length : start;
      const along = place === 'outEnd' ? end + out * half : place === 'inEnd' ? end - out * half : place === 'inBase' ? foot + out * half : start + length / 2;
      cx = anchor.lying ? along : anchor.x + anchor.width / 2;
      cy = anchor.lying ? anchor.y + anchor.height / 2 : along;
    } else if (anchor.kind === 'point') {
      const place = wanted === 'r' || wanted === 'l' || wanted === 't' || wanted === 'b' || wanted === 'ctr' ? wanted : anchor.centred ? 'ctr' : 'r';
      cx = anchor.x + (place === 'r' ? anchor.radius + 4 + w / 2 : place === 'l' ? -(anchor.radius + 4 + w / 2) : 0);
      cy = anchor.y + (place === 'b' ? anchor.radius + gap + h / 2 : place === 't' ? -(anchor.radius + gap + h / 2) : 0);
    } else {
      const middle = (anchor.start + anchor.end) / 2;
      const cos = Math.cos(middle);
      const sin = Math.sin(middle);
      // How far the box of the label reaches towards the rim, and how much of the arc it takes.
      const reach = (Math.abs(cos) * w + Math.abs(sin) * h) / 2;
      const breadth = Math.abs(sin) * w + Math.abs(cos) * h;
      const inside = Math.max(anchor.outer - 4 - reach, anchor.outer * 0.3);
      let radius: number;
      if (anchor.ring) radius = (anchor.inner + anchor.outer) / 2;
      else if (wanted === 'ctr') radius = anchor.outer * 0.55;
      else if (wanted === 'inEnd') radius = inside;
      else if (wanted === 'outEnd') radius = anchor.outer + 4 + reach;
      // "bestFit" and whatever else: inside where the slice has room for the label, outside where it has none.
      else radius = (anchor.end - anchor.start) * inside >= breadth + 2 && anchor.outer - 4 - reach >= anchor.outer * 0.3 ? inside : anchor.outer + 4 + reach;
      cx = anchor.cx + cos * radius;
      cy = anchor.cy + sin * radius;
    }
    if (!Number.isFinite(cx) || !Number.isFinite(cy)) return;
    // Moved into the frame rather than cut by it.
    const left = clamp(cx - w / 2, 0, Math.max(this.width - w, 0));
    const top = clamp(cy - h / 2, 0, Math.max(this.height - h, 0));
    const boxes = this.boxes;
    for (let i = 0; i < boxes.length; i += 4) {
      if (left < boxes[i + 2] && boxes[i] < left + w && top < boxes[i + 3] && boxes[i + 1] < top + h) return;
    }
    boxes.push(left, top, left + w, top + h);
    this.items.push({ kind: 'text', role: 'label', text, x: left + w / 2, y: top + font.size * 0.85, font, color: look.color, align: 'center' });
  }
}

// ------------------------------------------------------------------------- pie

function layoutPie(context: Context, layout: ChartLayout, box: Box, shown: ChartSeriesValues[]): void {
  const { chart, items } = context;
  const placer = context.detail === 'full' ? new LabelPlacer(context, layout.width, layout.height) : undefined;
  const labelled = placer ? shown.map((entry) => placer.of(entry)) : [];
  // Labels that stand outside the pie take a ring around it.
  let around = 0;
  if (placer && shown.length > 0 && shown[0].kind === 'pie' && placer.position(labelled[0]) === 'outEnd') {
    const font = fontOf(labelled[0] && labelled[0].base ? labelled[0].base.style : undefined, chart.textStyle, 9);
    around = Math.min(font.size * 2.4 + 8, Math.min(box.width, box.height) * 0.3);
  }
  // A pie shows its first series; a doughnut one ring per series.
  const first: ChartSeriesValues | undefined = shown[0];
  const rings = !first ? [] : first.kind === 'doughnut' ? shown.filter((entry) => entry.group === first.group).slice(0, 32) : [first];
  // Slices that are pulled out move away from the centre by a share of the radius: the pie is made smaller by the
  // largest of those shares, so that they stay in its place.
  const share = (value: number | undefined): number => (typeof value === 'number' && Number.isFinite(value) ? clamp(value, 0, 400) / 100 : 0);
  let pulled = 0;
  for (const entry of rings) {
    pulled = Math.max(pulled, share(entry.series.explosion));
    if (Array.isArray(entry.series.points)) for (const point of entry.series.points) if (point) pulled = Math.max(pulled, share(point.explosion));
  }
  const diameter = Math.max((Math.min(box.width, box.height) - (context.detail === 'full' ? 8 : 0) - around * 2) / (1 + pulled), 1);
  const radius = diameter / 2;
  const cx = box.x + box.width / 2;
  const cy = box.y + box.height / 2;
  layout.plot = { x: cx - radius, y: cy - radius, width: diameter, height: diameter };
  if (!first) return;
  const group = chart.groups[first.group];
  const hole = first.kind === 'doughnut' ? (radius * clamp(group.holeSize === undefined ? 50 : group.holeSize, 1, 90)) / 100 : 0;
  const thickness = (radius - hole) / rings.length;
  const turn = Math.PI * 2;
  const begin = -Math.PI / 2 + ((group.firstSliceAngle || 0) * Math.PI) / 180;
  for (let r = 0; r < rings.length; r++) {
    const entry = rings[r];
    const look = lookOf(chart, group, entry.series);
    // UNSURE: a negative value takes the share of its size, as Excel draws it; the first series is the innermost ring.
    let total = 0;
    entry.values.each((value) => {
      total += Math.abs(value);
    });
    if (!(total > 0) || !Number.isFinite(total)) continue;
    const inner = hole + r * thickness;
    const outer = inner + thickness;
    let angle = begin;
    const index = context.values.series.indexOf(entry);
    const known = entry.values.known;
    // A pie with a second plot: the points of that plot are one slice, after the others.
    const split = group.ofPie;
    const points = entry.values.length;
    let apart: ((i: number, size: number) => boolean) | undefined;
    if (split) {
      const at = split.position;
      if (split.split === 'val') apart = (_i, size) => at !== undefined && size < at;
      else if (split.split === 'percent') apart = (_i, size) => at !== undefined && (size / total) * 100 < at;
      else if (split.split === 'cust') {
        const listed = new Set<number>(Array.isArray(split.points) ? split.points : []);
        apart = (i) => listed.has(i);
      } else {
        // UNSURE: without a number of its own the last third of the points, rounded up, is the second plot.
        const last = at !== undefined && at >= 0 ? Math.floor(at) : Math.ceil(points / 3);
        apart = (i) => i >= points - last;
      }
    }
    let rest = 0;
    for (let k = 0; k < known; k++) {
      const i = entry.values.knownIndex(k);
      const size = Math.abs(entry.values.knownValue(k));
      if (!(size > 0)) continue;
      if (apart && apart(i, size)) {
        rest += size;
        continue;
      }
      const end = angle + (size / total) * turn;
      if (i >= MAX_PIE_SLICES) {
        // What is left of a pie of very many slices is one slice.
        items.push({ kind: 'slice', cx, cy, radius: outer, inner, start: angle, end: begin + turn, fill: REST_COLOR, series: index, point: i });
        break;
      }
      const point = pointLook(chart, look, i, group.varyColors !== false);
      // Pulled out along its middle: by what the point says, else by what its series says.
      const own = look.points ? look.points.get(i) : undefined;
      const out = pulled > 0 ? share(own && own.explosion !== undefined ? own.explosion : entry.series.explosion) * radius : 0;
      // In a frame too small for the pie itself the centre of a slice still lies in the frame.
      const sx = out > 0 ? clamp(cx + Math.cos((angle + end) / 2) * out, 0, context.width) : cx;
      const sy = out > 0 ? clamp(cy + Math.sin((angle + end) / 2) * out, 0, context.height) : cy;
      items.push({ kind: 'slice', cx: sx, cy: sy, radius: outer, inner, start: angle, end, fill: point.fill || '#FFFFFF', stroke: point.stroke, lineWidth: point.lineWidth, series: index, point: i });
      const labels = labelled[shown.indexOf(entry)];
      if (placer && labels) placer.add(entry, labels, i, entry.values.knownValue(k), { kind: 'slice', cx: sx, cy: sy, inner, outer, start: angle, end, ring: first.kind === 'doughnut' }, { share: size / total });
      angle = end;
    }
    // UNSURE: the slice of the second plot has the colour that comes after that of the last point.
    if (rest > 0) items.push({ kind: 'slice', cx, cy, radius: outer, inner, start: angle, end: Math.min(angle + (rest / total) * turn, begin + turn), fill: chartAutoColor(chart, points), stroke: look.stroke, lineWidth: look.lineWidth, series: index, point: points });
  }
  if (placer) for (const item of placer.items) items.push(item);
}

// ----------------------------------------------------------------------- radar

/**
 * A radar chart: the categories are spokes around a centre, the first one pointing up and the others
 * following clockwise; a value lies on its spoke as far from the centre as the scale says. The value
 * axis is linear (a logarithmic base is not applied) and its labels run along the first spoke.
 */
function layoutRadar(context: Context, layout: ChartLayout, box: Box, shown: ChartSeriesValues[]): void {
  const { chart, values, measure, options, items, detail } = context;
  const words = detail === 'full';
  const first = shown[0];
  if (!first) return;
  const group = chart.groups[first.group];
  const catAxis = axisOf(chart, group, 'cat');
  const valAxis = axisOf(chart, group, 'val');
  let count = 0;
  for (const entry of shown) count = Math.max(count, entry.values.length);
  // No more spokes than a circle can show, and no more of everything than a frame is worth painting.
  const spokes = Math.min(count, MAX_RADAR_SPOKES, Math.max(Math.floor(context.budget / Math.max(shown.length, 1)), 3));
  const catFont = fontOf(catAxis ? catAxis.labelStyle : undefined, chart.textStyle, 9);
  const catColor = colorOf(catAxis ? catAxis.labelStyle : undefined, chart.textStyle);
  const labelFont = fontOf(valAxis ? valAxis.labelStyle : undefined, chart.textStyle, 9);
  const labelColor = colorOf(valAxis ? valAxis.labelStyle : undefined, chart.textStyle);
  const categoryLabels = words && spokes > 0 && !(catAxis && (catAxis.deleted || catAxis.tickLabels === 'none'));
  const valueLabels = words && !(valAxis && (valAxis.deleted || valAxis.tickLabels === 'none'));
  const source = (i: number): number => (spokes < count ? Math.floor((i * count) / spokes) : i);

  // Room around the circle for the labels of the categories.
  let widest = 0;
  if (categoryLabels) {
    const stride = Math.max(Math.floor(spokes / 50), 1);
    for (let i = 0; i < spokes; i += stride) widest = Math.max(widest, measure(clean(values.label(source(i))), catFont));
    widest = Math.min(widest, box.width * 0.25);
  }
  const marginX = categoryLabels ? widest + 6 : 0;
  const marginY = categoryLabels ? catFont.size * 1.3 + 4 : words ? 4 : 0;
  const radius = Math.max(Math.min(box.width / 2 - marginX, box.height / 2 - marginY), 1);
  const cx = box.x + box.width / 2;
  const cy = box.y + box.height / 2;
  layout.plot = { x: cx - radius, y: cy - radius, width: radius * 2, height: radius * 2 };

  // The series as they are drawn: one value for each spoke.
  const drawn: { entry: ChartSeriesValues; list: (number | null)[] }[] = [];
  let low = Infinity;
  let high = -Infinity;
  for (const entry of shown.slice(0, Math.max(Math.floor(context.budget / Math.max(spokes, 1)), 1))) {
    const list = spokes < count ? sample(entry.values, count, spokes).peak : spread(entry.values, spokes);
    for (const value of list) {
      if (value === null || !Number.isFinite(value)) continue;
      if (value < low) low = value;
      if (value > high) high = value;
    }
    drawn.push({ entry, list });
  }
  if (low === Infinity) {
    low = 0;
    high = 0;
  }
  const scale = chartScale(low, high, {
    min: valAxis ? valAxis.min : undefined,
    max: valAxis ? valAxis.max : undefined,
    step: valAxis ? valAxis.majorUnit : undefined,
    maxSteps: clamp(Math.floor(radius / Math.max(labelFont.size * 2.2, 8)), 1, 10),
  });
  layout.scale = scale;
  if (spokes < 1) return;
  const span = scale.max - scale.min;
  const reach = (value: number): number => clamp((value - scale.min) / span, 0, 1) * radius;
  const turn = (Math.PI * 2) / spokes;
  const cos = (i: number): number => Math.cos(-Math.PI / 2 + i * turn);
  const sin = (i: number): number => Math.sin(-Math.PI / 2 + i * turn);

  if (detail !== 'bare') {
    // A ring for every tick above the centre: a polygon through the spokes.
    const grid = valAxis && !valAxis.deleted ? valAxis.majorGridlines : undefined;
    const stroke = grid ? lineColor(grid, GRID_COLOR) : undefined;
    if (grid && stroke) {
      for (const tick of scale.ticks) {
        const r = reach(tick);
        if (!(r > 0)) continue;
        const points: number[] = [];
        for (let i = 0; i < spokes; i++) points.push(cx + cos(i) * r, cy + sin(i) * r);
        if (points.length >= 6) items.push({ kind: 'path', role: 'grid', points, closed: true, stroke, lineWidth: clamp(grid.width || 1, 0.5, 8), dash: grid.dash === true });
      }
    }
    // The spokes.
    if (!(valAxis && valAxis.deleted)) {
      const line = valAxis ? valAxis.line : undefined;
      const colour = lineColor(line, GRID_COLOR);
      if (colour) for (let i = 0; i < spokes; i++) items.push({ kind: 'path', role: 'axis', points: [cx, cy, cx + cos(i) * radius, cy + sin(i) * radius], stroke: colour, lineWidth: clamp((line && line.width) || 1, 0.5, 8), dash: line !== undefined && line.dash === true });
    }
  }

  const placer = words ? new LabelPlacer(context, layout.width, layout.height) : undefined;
  const spanning = chart.blanks === 'span';
  for (const { entry, list } of drawn) {
    const look = lookOf(chart, group, entry.series);
    const index = values.series.indexOf(entry);
    const labels = placer && spokes === count ? placer.of(entry) : undefined;
    // The outline: runs of points, broken where a point has no number; the last point is joined to the first.
    const runs: number[][] = [];
    const marks: number[] = [];
    let run: number[] = [];
    let whole = true;
    for (let i = 0; i < spokes; i++) {
      const value = list[i];
      if (value === null || value === undefined || !Number.isFinite(value)) {
        if (spanning) continue;
        whole = false;
        if (run.length > 0) runs.push(run);
        run = [];
        continue;
      }
      const r = reach(value);
      const x = cx + cos(i) * r;
      const y = cy + sin(i) * r;
      run.push(x, y);
      marks.push(x, y);
      if (placer && labels) placer.add(entry, labels, i, value, { kind: 'point', x, y, radius: look.marker ? look.marker.size / 2 : 2 });
    }
    if (run.length > 0) runs.push(run);
    if (!whole && runs.length > 1 && list[0] !== null && list[spokes - 1] !== null) {
      // The run that ends at the last spoke goes on with the one that starts at the first.
      const last = runs.pop() as number[];
      runs[0] = last.concat(runs[0]);
    }
    for (const points of runs) {
      const closed = whole && points.length >= 6;
      if (look.fill !== undefined && closed) items.push({ kind: 'path', role: 'area', points, closed: true, fill: look.fill, stroke: look.stroke, lineWidth: look.lineWidth, series: index });
      else if (points.length >= 4 && (look.stroke || look.fill)) items.push({ kind: 'path', role: 'line', points, closed: closed ? true : undefined, stroke: look.stroke || look.fill, lineWidth: look.lineWidth, dash: look.dash, series: index });
    }
    if (look.marker && marks.length > 0) items.push({ kind: 'markers', role: 'marker', symbol: look.marker.symbol, size: look.marker.size, points: marks, fill: look.marker.fill, stroke: look.marker.stroke, series: index });
  }
  if (!words) return;

  // Labels of the values, beside the spoke that points up.
  if (valueLabels) {
    const format = valAxis && valAxis.formatCode && !valAxis.sourceLinked ? { code: valAxis.formatCode } : first.format || (valAxis && valAxis.formatCode ? { code: valAxis.formatCode } : undefined);
    const every = Math.max(Math.ceil((labelFont.size * 1.1 * scale.ticks.length) / Math.max(radius, 1)), 1);
    for (let i = 0; i < scale.ticks.length; i += every) {
      const text = fit(tickText(scale.ticks[i], format, options), labelFont, Math.max(cx - 4, 0), measure);
      if (text !== '') items.push({ kind: 'text', role: 'tick', text, x: cx - 4, y: cy - reach(scale.ticks[i]) + labelFont.size * 0.35, font: labelFont, color: labelColor, align: 'right' });
    }
  }
  // Labels of the categories, beyond the ends of their spokes: as many as there is room for around the circle.
  if (categoryLabels) {
    const room = Math.max(Math.floor((Math.PI * 2 * radius) / (catFont.size * 1.6)), 4);
    const every = Math.max(Math.ceil(spokes / room), 1);
    for (let i = 0; i < spokes; i += every) {
      const c = cos(i);
      const s = sin(i);
      const x = cx + c * (radius + 6);
      const align = Math.abs(c) < 0.2 ? 'center' : c > 0 ? 'left' : 'right';
      const y = cy + s * (radius + 6) + (s < -0.5 ? -2 : s > 0.5 ? catFont.size * 0.85 : catFont.size * 0.35);
      const text = fit(clean(values.label(source(i))), catFont, align === 'left' ? layout.width - x : align === 'right' ? x : Math.min(x, layout.width - x) * 2, measure);
      if (text !== '') items.push({ kind: 'text', role: 'category', text, x, y, font: catFont, color: catColor, align });
    }
  }
  if (placer) for (const item of placer.items) items.push(item);
}

// ------------------------------------------------------------------------ axes

function axisById(chart: Chart, id: number): ChartAxis | undefined {
  for (const axis of chart.axes) if (axis.id === id) return axis;
  return undefined;
}

/** The axis of a kind among those a chart type is drawn against; one of the chart that the type does not name when it names none of the kind. */
function axisOf(chart: Chart, group: ChartGroup, kind: 'cat' | 'val'): ChartAxis | undefined {
  let fallback: ChartAxis | undefined;
  for (const axis of chart.axes) {
    if (axis.kind !== kind) continue;
    if (group.axisIds.indexOf(axis.id) >= 0) return axis;
    if (!fallback) fallback = axis;
  }
  return fallback;
}

/** The axis the categories of a chart type run along: a category or date axis (c:catAx, c:dateAx), never a value axis. */
function categoryAxis(chart: Chart, group: ChartGroup): ChartAxis | undefined {
  return group.kind === 'scatter' || group.kind === 'pie' || group.kind === 'doughnut' ? undefined : axisOf(chart, group, 'cat');
}

/**
 * The two axes of a scatter chart, both value axes. Which is which is the order the chart
 * type names them in: the first is the horizontal one, the second the vertical one. The
 * side an axis says it stands on (c:axPos) is only believed where the two agree with each
 * other - one lies, one stands: files not written by Excel put both at the left, and the
 * settings of x must not end up on y then.
 */
export function scatterAxes(chart: Chart, group: ChartGroup): { x: ChartAxis | undefined; y: ChartAxis | undefined } {
  const pair: ChartAxis[] = [];
  for (const id of group.axisIds) {
    const axis = axisById(chart, id);
    if (axis && pair.indexOf(axis) < 0 && pair.length < 2) pair.push(axis);
  }
  // A chart type that names fewer than two axes that are there: the value axes of the chart in the order of the file.
  for (const axis of chart.axes) if (pair.length < 2 && axis.kind === 'val' && pair.indexOf(axis) < 0) pair.push(axis);
  const lies = (axis: ChartAxis): boolean => axis.position === 'b' || axis.position === 't';
  if (pair.length === 0) return { x: undefined, y: undefined };
  if (pair.length === 1) return lies(pair[0]) ? { x: pair[0], y: undefined } : { x: undefined, y: pair[0] };
  if (lies(pair[1]) && !lies(pair[0])) return { x: pair[1], y: pair[0] };
  return { x: pair[0], y: pair[1] };
}

/** The points of a line as they are gathered: runs that are joined, the places of markers, and points that stand alone. */
class Trace {
  private readonly runs: number[][] = [];
  private readonly marks: number[] = [];
  /** Points with a gap on either side: no line shows them. */
  private readonly lonely: number[] = [];
  private run: number[] = [];

  /**
   * `spanning`: a point without a number does not break the line. `smooth`: the line is a curve
   * through its points, inside the frame of that size.
   */
  constructor(
    private readonly spanning: boolean,
    private readonly smooth?: { width: number; height: number },
  ) {}

  push(x: number, y: number, marked: boolean): void {
    this.run.push(x, y);
    if (marked) this.marks.push(x, y);
  }

  /** A point without a number. */
  gap(): void {
    if (!this.spanning) this.cut();
  }

  private cut(): void {
    if (this.run.length >= 4) this.runs.push(this.run);
    else if (this.run.length === 2) this.lonely.push(this.run[0], this.run[1]);
    this.run = [];
  }

  /** Adds what was gathered to the items of a layout. `dots`: a point between two gaps gets a dot of the width of the line. */
  emit(items: ChartItem[], look: SeriesLook, series: number, markers: boolean, dots: boolean): void {
    this.cut();
    if (look.stroke) {
      for (const points of this.runs) {
        const frame = this.smooth;
        if (frame && points.length >= 6) {
          // UNSURE: a curve through the points that may swing beyond them (a Catmull-Rom spline); how Excel shapes its curve is not known.
          const curve = smoothPath(points);
          // What shapes the curve stays in the frame, and so does the curve.
          for (let i = 0; i < curve.length; i += 2) {
            curve[i] = clamp(curve[i], 0, frame.width);
            curve[i + 1] = clamp(curve[i + 1], 0, frame.height);
          }
          items.push({ kind: 'path', role: 'line', points: curve, curve: true, stroke: look.stroke, lineWidth: look.lineWidth, dash: look.dash, series });
        } else items.push({ kind: 'path', role: 'line', points, stroke: look.stroke, lineWidth: look.lineWidth, dash: look.dash, series });
      }
    }
    const marker = look.marker;
    if (marker && this.marks.length > 0 && markers) items.push({ kind: 'markers', role: 'marker', symbol: marker.symbol, size: marker.size, points: this.marks, fill: marker.fill, stroke: marker.stroke, series });
    else if (!marker && dots && this.lonely.length > 0 && look.stroke) {
      items.push({ kind: 'markers', role: 'marker', symbol: 'circle', size: Math.max(look.lineWidth, 2), points: this.lonely, fill: look.stroke, series });
    }
  }
}

type Side = 'l' | 'r' | 't' | 'b';
const OPPOSITE: { [side in Side]: Side } = { l: 'r', r: 'l', t: 'b', b: 't' };

/** Drawn at places of two numbers: a scatter series, a bubble series. */
const isXY = (kind: ChartKind): boolean => kind === 'scatter' || kind === 'bubble';
/** Drawn as lines through points along the categories. */
const isLine = (kind: ChartKind): boolean => kind === 'line' || kind === 'stock';

type NumberFormat = { code: string; builtinId?: number };
const isDateCode = (code: string, builtinId: number | undefined): boolean => builtinId === 14 || builtinId === 22 || isDateFormat(code);

/**
 * One pair of axes of a plot area with the scale of its value axis. A chart has one, or two
 * when some of its chart types name a value axis of their own: the series of those are drawn
 * against the second scale. The categories are shared.
 */
interface AxisSet {
  cat: ChartAxis | undefined;
  val: ChartAxis | undefined;
  /** What the value axis has to hold. */
  low: number;
  high: number;
  /** Scatter charts: what the horizontal axis has to hold. */
  xLow: number;
  xHigh: number;
  /** Its series are shares of 100 %. */
  percent: boolean;
  format: NumberFormat | undefined;
  scale: ChartScale;
  xScale: ChartScale | undefined;
  tickLabels: string[];
  font: ChartFont;
  color: string;
  /** The value axis shows labels. */
  labels: boolean;
  /** The side of the plot area the value axis stands on, and the side its labels stand on. */
  side: Side;
  labelSide: Side;
  /** The room the labels of the value axis take across their side. */
  band: number;
  /** The value the series stand on: where the category axis crosses. */
  base: number;
}

const UNIT_SCALE: ChartScale = { min: 0, max: 1, step: 0.2, ticks: [0, 0.2, 0.4, 0.6, 0.8, 1] };

interface Plotted {
  entry: ChartSeriesValues;
  index: number;
  group: ChartGroup;
  set: AxisSet;
  look: SeriesLook;
  /** One number per position that is drawn. */
  values: (number | null)[];
  /** Lines of very many points: the smallest and the largest value at each position. */
  low?: (number | null)[];
  high?: (number | null)[];
  /** Stacked series: the value each position is drawn at (null for one that is not drawn) and the value it stands on. */
  to?: (number | null)[];
  from?: number[];
}

/** A scatter series: drawn from its own numbers, at the places its x values say. */
interface Scattered {
  entry: ChartSeriesValues;
  index: number;
  group: ChartGroup;
  set: AxisSet;
  look: SeriesLook;
}

/** How the labels of the categories of an upright chart are set: flat, or turned where they do not fit. */
interface CategoryPlan {
  /** Degrees clockwise; 0 for labels that lie flat. */
  angle: number;
  /** Every n-th label is shown. */
  every: number;
  /** The height the labels take, the gap to the axis included. */
  band: number;
  /** The longest a label may be, pixels. */
  room: number;
}

type TimeUnit = 'days' | 'months' | 'years';
const UNIT_DAYS: { [unit in TimeUnit]: number } = { days: 1, months: 30.4375, years: 365.25 };
const UNIT_RANK: { [unit in TimeUnit]: number } = { days: 0, months: 1, years: 2 };
/** The last day a serial number can name: 31 December 9999. */
const MAX_SERIAL = 2958465;
/** The steps a date axis picks from when the file states none, from the finest on. */
const DATE_STEPS: [TimeUnit, number][] = [
  ['days', 1],
  ['days', 2],
  ['days', 7],
  ['days', 14],
  ['months', 1],
  ['months', 2],
  ['months', 3],
  ['months', 6],
  ['years', 1],
  ['years', 2],
  ['years', 5],
  ['years', 10],
  ['years', 20],
  ['years', 50],
  ['years', 100],
  ['years', 200],
  ['years', 500],
  ['years', 1000],
  ['years', 2000],
  ['years', 5000],
];

/** Where the categories of a date axis stand: each at the unit of time its date lies in. */
interface DatePlan {
  unit: TimeUnit;
  /** The unit of the earliest date, counted from the start of the calendar. */
  first: number;
  /** How many units the axis spans. */
  units: number;
  /** The unit of each position that is drawn, counted from `first`. */
  position: number[];
}

/** The unit of time a serial number lies in, counted from the start of the calendar; undefined for a number that is no date. */
function timeUnitOf(serial: number, unit: TimeUnit, date1904: boolean): number | undefined {
  if (!(serial >= 0) || serial > MAX_SERIAL) return undefined;
  if (unit === 'days') return Math.floor(serial);
  const date = serialToDate(serial, date1904);
  if (!Number.isFinite(date.year) || !Number.isFinite(date.month)) return undefined;
  return unit === 'months' ? date.year * 12 + date.month - 1 : date.year;
}

/** The serial number of a day of the calendar; undefined for a day that is none or lies outside what a serial number can name. */
function serialOfDay(year: number, month: number, day: number, date1904: boolean): number | undefined {
  if (!(year >= 1900 && year <= 9999)) return undefined;
  const two = (value: number): string => (value < 10 ? '0' : '') + value;
  const serial = isoToSerial(`${year}-${two(month)}-${two(day)}`, date1904);
  return serial !== undefined && serial >= 0 && serial <= MAX_SERIAL ? serial : undefined;
}

/** `inner`: the place of the area of the series when the file sets it by hand; taken where the labels of the axes have room around it. */
function layoutAxes(context: Context, layout: ChartLayout, box: Box, shown: ChartSeriesValues[], inner?: Box): void {
  const { chart, values, measure, options, items, detail } = context;
  const words = detail === 'full';
  const scatter = shown.length > 0 && shown.every((entry) => isXY(entry.kind));
  // Series of a scatter type beside bars, lines or areas are drawn on the axes of those (see below).
  const categorical = scatter ? [] : shown.filter((entry) => !isXY(entry.kind));
  const loose = shown.filter((entry) => isXY(entry.kind));
  const firstGroup = categorical.length > 0 ? chart.groups[categorical[0].group] : shown.length > 0 ? chart.groups[shown[0].group] : chart.groups[0];
  const barGroup = chart.groups.filter((group) => group.kind === 'bar')[0];
  // Bars that lie turn the whole chart: categories run up the side, values along the bottom.
  const lying = !scatter && barGroup !== undefined && barGroup.barDir === 'bar';

  // The pairs of axes: that of the first chart type, and a second one for the chart types that name another value axis.
  const sets: AxisSet[] = [];
  const setOfGroup = new Map<ChartGroup, AxisSet>();
  const setFor = (group: ChartGroup): AxisSet => {
    let set = setOfGroup.get(group);
    if (set) return set;
    let cat: ChartAxis | undefined;
    let val: ChartAxis | undefined;
    if (scatter) {
      const both = scatterAxes(chart, group);
      cat = both.x;
      val = both.y;
    } else if (isXY(group.kind) && sets.length > 0) set = sets[0];
    else {
      cat = axisOf(chart, group, 'cat');
      val = axisOf(chart, group, 'val');
    }
    if (!set) for (const other of sets) if (other.val === val) set = other;
    if (!set) {
      // A third value axis has no side left to stand on: its series go with the second.
      if (sets.length >= 2) set = sets[1];
      else {
        set = {
          cat,
          val,
          low: Infinity,
          high: -Infinity,
          xLow: Infinity,
          xHigh: -Infinity,
          percent: false,
          format: undefined,
          scale: UNIT_SCALE,
          xScale: undefined,
          tickLabels: [],
          font: fontOf(val ? val.labelStyle : undefined, chart.textStyle, 9),
          color: colorOf(val ? val.labelStyle : undefined, chart.textStyle),
          labels: words && !(val && (val.deleted || val.tickLabels === 'none')),
          side: 'l',
          labelSide: 'l',
          band: 0,
          base: 0,
        };
        sets.push(set);
      }
    }
    setOfGroup.set(group, set);
    return set;
  };
  setFor(firstGroup);
  for (const entry of shown) setFor(chart.groups[entry.group]);
  if (sets.length === 2) {
    // Which pair is the first is not a matter of which chart type the file writes first: the first pair is the one
    // whose category axis is shown and whose value axis crosses at the start, as Excel writes its primary axes (the
    // second pair has its category axis deleted and its value axis crossing at the maximum). UNSURE: where both
    // pairs or neither look like that, the pair whose value axis the file lists first is taken - Excel lists the
    // primary axes first.
    const first = (set: AxisSet): boolean => set.cat !== undefined && !set.cat.deleted && !(set.val !== undefined && set.val.crosses === 'max');
    const listed = (set: AxisSet): number => (set.val ? chart.axes.indexOf(set.val) : -1);
    const a = first(sets[0]);
    const b = first(sets[1]);
    if (a !== b ? b : listed(sets[1]) >= 0 && (listed(sets[0]) < 0 || listed(sets[1]) < listed(sets[0]))) sets.reverse();
  }
  const primary = sets[0];
  const second: AxisSet | undefined = sets[1];
  const catAxis = primary.cat;
  const valAxis = primary.val;
  const catFont = fontOf(catAxis ? catAxis.labelStyle : undefined, chart.textStyle, 9);
  const catColor = colorOf(catAxis ? catAxis.labelStyle : undefined, chart.textStyle);
  const categoryLabels = words && !(catAxis && (catAxis.deleted || catAxis.tickLabels === 'none'));
  // The category axis of the second pair is drawn where the file keeps it: on the side the first leaves free.
  const cat2 = second && second.cat && second.cat !== catAxis && !second.cat.deleted ? second.cat : undefined;
  const cat2Font = fontOf(cat2 ? cat2.labelStyle : undefined, chart.textStyle, 9);
  const cat2Color = colorOf(cat2 ? cat2.labelStyle : undefined, chart.textStyle);
  const category2Labels = words && cat2 !== undefined && cat2.tickLabels !== 'none';

  // The series as they are drawn: sampled down where there are more points than places to draw them. The
  // categories are those of the series that have any: a scatter series beside them brings none of its own.
  let count = 0;
  for (const entry of categorical) count = Math.max(count, entry.values.length);
  count = scatter ? 0 : Math.max(count, 1);

  // Which side every axis stands on. The value axis stands where it crosses the categories: at the first of them,
  // or at the last when it says so; categories in reverse order take it along. A side the file states (c:axPos)
  // moves it to the far side, never back: files not written by Excel say "l" for every axis.
  // UNSURE: Excel is taken to place an axis by where it crosses alone and to write c:axPos to match.
  const catReversed = catAxis !== undefined && catAxis.reversed === true;
  const valNear: Side = lying ? 'b' : 'l';
  const catNear: Side = lying ? 'l' : 'b';
  sets.forEach((set, index) => {
    const val = set.val;
    // A value axis that crosses at a category number stands inside the plot area, which is not drawn: it stays at
    // the start. Only a number beyond the last category takes it to the far side, as "max" does. UNSURE.
    const atEnd = val !== undefined && (val.crosses === 'max' || (val.crossesAt !== undefined && !scatter && val.crossesAt > count));
    let side = atEnd !== catReversed ? OPPOSITE[valNear] : valNear;
    if (side === valNear && val && val.position === OPPOSITE[valNear]) side = OPPOSITE[valNear];
    // UNSURE: a second value axis that does not say where it stands is put on the side the first leaves free
    // (Excel would draw the two over each other).
    if (index > 0 && side === primary.side) side = OPPOSITE[side];
    set.side = side;
    const where = val ? val.tickLabels : 'nextTo';
    let labelSide = where === 'high' ? (catReversed ? valNear : OPPOSITE[valNear]) : where === 'low' ? (catReversed ? OPPOSITE[valNear] : valNear) : side;
    if (index > 0 && set.labels && primary.labels && labelSide === primary.labelSide) labelSide = OPPOSITE[labelSide];
    set.labelSide = labelSide;
  });
  // The category axis stands where it crosses the values: at zero (the near end), or at the maximum when it says so;
  // values in reverse order take it along.
  const valueReversed = valAxis !== undefined && valAxis.reversed === true;
  const catAtEnd = catAxis !== undefined && catAxis.crosses === 'max' && catAxis.crossesAt === undefined;
  let catLine: Side = catAtEnd !== valueReversed ? OPPOSITE[catNear] : catNear;
  if (catLine === catNear && catAxis && catAxis.position === OPPOSITE[catNear]) catLine = OPPOSITE[catNear];
  const catWhere = catAxis ? catAxis.tickLabels : 'nextTo';
  const catSide: Side = catWhere === 'high' ? (valueReversed ? catNear : OPPOSITE[catNear]) : catWhere === 'low' ? (valueReversed ? OPPOSITE[catNear] : catNear) : catLine;
  const cat2Side = OPPOSITE[catSide];

  // Titles of the axes take a band on the side their labels stand on.
  interface AxisTitle {
    text: string;
    font: ChartFont;
    color: string;
    side: Side;
  }
  const titles: AxisTitle[] = [];
  const titleOn = (axis: ChartAxis | undefined, side: Side): void => {
    if (!words || !axis || axis.deleted || !axis.title || axis.title.text === undefined || titles.some((other) => other.side === side)) return;
    const text = clean(axis.title.text.replace(/\n/g, ' '));
    const font = fontOf(axis.title.style, chart.textStyle, 10, true);
    if (text === '') return;
    const taken = font.size * 1.4;
    if (side === 'b' || side === 't') {
      if (!(box.height > 120)) return;
      box.height -= taken;
      if (side === 't') box.y += taken;
    } else {
      if (!(box.width > 160)) return;
      box.width -= taken;
      if (side === 'l') box.x += taken;
    }
    titles.push({ text, font, color: colorOf(axis.title.style, chart.textStyle), side });
  };
  titleOn(catAxis, catSide);
  titleOn(valAxis, primary.labelSide);
  if (second) titleOn(second.val, second.labelSide);
  if (cat2) titleOn(cat2, cat2Side);

  const alongGuess = Math.max(lying ? box.height : box.width, 1);
  // No more positions than pixels, and no more of everything together than a frame is worth painting.
  const drawn = Math.max(Math.min(Math.max(Math.floor(alongGuess), 16), MAX_DRAWN_POINTS, Math.floor(context.budget / Math.max(shown.length, 1))), 16);
  const places = scatter ? 0 : count > drawn * 2 ? drawn : count;
  /** Points drawn of one scatter series. */
  const drawnScatter = Math.max(Math.min(Math.max(Math.floor(alongGuess), 16) * 2, MAX_DRAWN_POINTS * 2, Math.floor(context.budget / Math.max(shown.length, 1))), 16);
  const plotted: Plotted[] = [];
  for (const entry of categorical) {
    const group = chart.groups[entry.group];
    const set = setFor(group);
    const item: Plotted = { entry, index: values.series.indexOf(entry), group, set, look: lookOf(chart, group, entry.series), values: [] };
    if (places < count) {
      const sampled = sample(entry.values, count, places);
      item.values = sampled.peak;
      if (isLine(group.kind) && group.grouping === 'standard') {
        item.low = sampled.low;
        item.high = sampled.high;
      }
    } else item.values = spread(entry.values, places);
    if (set.val && set.val.logBase) {
      // A logarithmic axis has no place for what is not above zero: such a point is one without a number.
      for (const list of [item.values, item.low, item.high]) {
        if (!list) continue;
        for (let i = 0; i < list.length; i++) {
          const value = list[i];
          if (value !== null && !(value > 0)) list[i] = null;
        }
      }
    }
    plotted.push(item);
  }
  const scattered: Scattered[] = loose.map((entry) => {
    const group = chart.groups[entry.group];
    return { entry, index: values.series.indexOf(entry), group, set: setFor(group), look: lookOf(chart, group, entry.series) };
  });

  // Points stand in the middle of their categories, or on the tick marks when the value axis says so and no bars need the room.
  const onTicks = !scatter && valAxis !== undefined && valAxis.crossBetween === 'midCat' && !plotted.some((item) => item.group.kind === 'bar') && places > 1;

  // A date axis: the categories stand where their dates lie. Categories that are no dates stand side by side, and so
  // do those of bars that lie (UNSURE: Excel places those by date too; the dates are labels here).
  const date1904 = options.date1904 === true;
  const dates = ((): DatePlan | undefined => {
    const number = values.categoryNumber;
    if (scatter || lying || !catAxis || !catAxis.date || !number || places < 1) return undefined;
    let unit = catAxis.baseTimeUnit;
    if (!unit) {
      // UNSURE: without a unit of its own the axis counts in the smallest step between two of its dates:
      // years for steps of a year, months for steps of a month, days otherwise.
      let gap = Infinity;
      let before: number | null = null;
      for (let k = 0; k < Math.min(count, 200); k++) {
        const value = number(k);
        if (value === null) return undefined;
        if (before !== null && value !== before) gap = Math.min(gap, Math.abs(value - before));
        before = value;
      }
      unit = gap === Infinity ? 'days' : gap >= 360 ? 'years' : gap >= 28 ? 'months' : 'days';
    }
    const position: number[] = [];
    let first = Infinity;
    let last = -Infinity;
    for (let i = 0; i <= places; i++) {
      // The last point too, where it is not among the positions that are drawn.
      const source = i === places ? count - 1 : places < count ? Math.floor((i * count) / places) : i;
      const value = number(source);
      const at = value === null ? undefined : timeUnitOf(value, unit, date1904);
      if (at === undefined) return undefined;
      if (i < places) position.push(at);
      if (at < first) first = at;
      if (at > last) last = at;
    }
    for (let i = 0; i < position.length; i++) position[i] -= first;
    return { unit, first, units: last - first + 1, position };
  })();
  // The order lines and areas go through the categories of a date axis: by date.
  let order: number[] | undefined;
  if (dates && places === count && dates.position.some((at, i) => i > 0 && at < dates.position[i - 1])) {
    order = dates.position.map((_at, i) => i).sort((a, b) => dates.position[a] - dates.position[b] || a - b);
  }

  // What each value axis has to hold: the values themselves, or the sums of what is stacked.
  const see = (set: AxisSet, value: number | null | undefined): void => {
    if (value === null || value === undefined || !Number.isFinite(value)) return;
    if (set.val && set.val.logBase && !(value > 0)) return;
    if (value < set.low) set.low = value;
    if (value > set.high) set.high = value;
  };
  for (const item of scattered) {
    const set = item.set;
    item.entry.values.each((value) => see(set, value));
    if (item.entry.values.blank !== null && item.entry.values.known < item.entry.values.length) see(set, item.entry.values.blank);
  }
  const piles = new Map<number, { up: number[]; down: number[]; total: number[] }>();
  for (const item of plotted) {
    if (item.group.grouping === 'standard') {
      for (let i = 0; i < item.values.length; i++) see(item.set, item.values[i]);
      if (item.low && item.high) for (let i = 0; i < item.low.length; i++) (see(item.set, item.low[i]), see(item.set, item.high[i]));
      continue;
    }
    let pile = piles.get(item.entry.group);
    if (!pile) {
      pile = { up: [], down: [], total: [] };
      piles.set(item.entry.group, pile);
    }
    if (item.group.grouping === 'percentStacked') item.set.percent = true;
    // A share of 100 % is a share of the sizes: of the absolute values of what is stacked.
    for (let i = 0; i < places; i++) pile.total[i] = (pile.total[i] || 0) + Math.abs(item.values[i] || 0);
  }
  for (const item of plotted) {
    const pile = piles.get(item.entry.group);
    if (item.group.grouping === 'standard' || !pile) continue;
    const share = item.group.grouping === 'percentStacked';
    const bar = item.group.kind === 'bar';
    const to: (number | null)[] = [];
    const from: number[] = [];
    for (let i = 0; i < places; i++) {
      const raw = item.values[i];
      if ((raw === null || raw === undefined) && bar) {
        to.push(null);
        from.push(0);
        continue;
      }
      let value = raw || 0;
      if (share) {
        const total = pile.total[i] || 0;
        value = total > 0 ? value / total : 0;
      }
      // Bars pile up on two sides: what is above zero on what is above zero, what is below on what is below.
      // UNSURE: lines and areas have no sides. Each series is drawn at the sum of the series up to it,
      // whatever the signs: with 10 and then -3 the second line runs at 7.
      const heap = bar && value < 0 ? pile.down : pile.up;
      const base = heap[i] || 0;
      heap[i] = base + value;
      to.push(base + value);
      from.push(base);
      see(item.set, base + value);
    }
    item.to = to;
    item.from = from;
    // What is stacked stands on zero.
    see(item.set, 0);
  }
  for (const set of sets) {
    if (set.low === Infinity) {
      set.low = 0;
      set.high = 0;
    }
    // Bars and areas stand on zero: it is part of what the axis shows.
    if (!scatter && !(set.val && set.val.logBase) && plotted.some((item) => item.set === set && (item.group.kind === 'bar' || item.group.kind === 'area'))) {
      if (set.low > 0) set.low = 0;
      if (set.high < 0) set.high = 0;
    }
    const val = set.val;
    let firstFormat: NumberFormat | undefined;
    for (const item of plotted) if (item.set === set && firstFormat === undefined) firstFormat = item.entry.format;
    if (plotted.every((item) => item.set !== set)) for (const item of scattered) if (item.set === set && firstFormat === undefined) firstFormat = item.entry.format;
    set.format =
      set.percent && !(val && val.formatCode && !val.sourceLinked && val.formatCode !== 'General')
        ? { code: '0%' }
        : val && val.formatCode && !val.sourceLinked
          ? { code: val.formatCode }
          : firstFormat || (val && val.formatCode ? { code: val.formatCode } : undefined);
  }

  // Room for the labels: those of the values are measured once the scale is known, which needs the length of its axis.
  const tickGap = 6;
  const lineBand = (font: ChartFont): number => font.size * 1.3 + tickGap;
  const scaleOptions = (axis: ChartAxis | undefined, length: number, each: number): ChartScaleOptions => ({
    min: axis ? axis.min : undefined,
    max: axis ? axis.max : undefined,
    step: axis ? axis.majorUnit : undefined,
    maxSteps: clamp(Math.floor(length / each), 1, 10),
  });
  const valueScale = (set: AxisSet, length: number, each: number): ChartScale => {
    const val = set.val;
    if (val && val.logBase) return chartLogScale(set.low, set.high, val.logBase, scaleOptions(val, length, each));
    return set.percent && !(val && (val.min !== undefined || val.max !== undefined))
      ? chartScale(set.low, set.high, { ...scaleOptions(val, length, each), min: set.low < 0 ? -1 : 0, max: 1 })
      : chartScale(set.low, set.high, scaleOptions(val, length, each));
  };
  const outerLabel = !scatter && !lying && !dates && categoryLabels ? values.outerLabel : undefined;
  const outerBand = outerLabel ? catFont.size * 1.3 : 0;
  const need = { l: 0, r: 0, t: 0, b: 0 };
  let plot: Box = { x: box.x, y: box.y, width: box.width, height: box.height };
  let catWidth = 0;
  /** Works out the scales and the plot area for category labels that take `catBand` across their side. */
  const arrange = (catBand: number, leastLeft: number): void => {
    need.l = need.r = need.t = need.b = 0;
    const across = { l: 0, r: 0, t: 0, b: 0 };
    let edge: number;
    if (lying) {
      if (categoryLabels) across[catSide] += catWidth;
      for (const set of sets) if (set.labels) need[set.labelSide] += lineBand(set.font);
      // Room at the end of the axis for half of its last label.
      edge = sets.some((set) => set.labels) ? 16 : words ? 8 : 0;
    } else {
      if (categoryLabels) need[catSide] += catBand;
      if (category2Labels) need[cat2Side] += lineBand(cat2Font);
      // Half a line above the plot area for the label of the topmost tick.
      let headroom = 0;
      for (const set of sets) if (set.labels) headroom = Math.max(headroom, set.font.size * 0.6);
      need.t = Math.max(need.t, headroom);
      const plotHeight = Math.max(box.height - need.t - need.b, Math.min(box.height, 1));
      for (const set of sets) {
        set.scale = valueScale(set, plotHeight, Math.max(set.font.size * 2.2, 8));
        set.tickLabels = set.scale.ticks.map((tick) => tickText(tick, set.format, options));
        set.band = 0;
        if (!set.labels) continue;
        let widest = 0;
        for (const label of set.tickLabels) widest = Math.max(widest, measure(label, set.font));
        set.band = Math.min(widest, box.width * 0.4) + tickGap;
        across[set.labelSide] += set.band;
      }
      // Room at the far side for half of the last label, where the last point stands on the edge of the plot area.
      edge = words ? 6 : 0;
      if (categoryLabels && scatter) edge = 16;
      else if (categoryLabels && onTicks) edge = Math.max(edge, Math.min(measure(clean(values.label(count - 1)), catFont) / 2, 40) + 2);
    }
    need.l = across.l > 0 ? across.l : across.r > 0 ? edge : 0;
    need.r = Math.max(across.r, edge);
    need.l = Math.max(need.l, Math.min(leastLeft, box.width * 0.3));
    plot = { x: box.x + need.l, y: box.y + need.t, width: Math.max(box.width - need.l - need.r, Math.min(box.width, 1)), height: Math.max(box.height - need.t - need.b, Math.min(box.height, 1)) };
    if (lying) {
      for (const set of sets) {
        set.scale = valueScale(set, plot.width, Math.max(set.font.size * 3.2, 36));
        set.tickLabels = set.scale.ticks.map((tick) => tickText(tick, set.format, options));
        set.band = set.labels ? lineBand(set.font) : 0;
      }
    }
  };
  let plan: CategoryPlan | undefined;
  /** The width of the widest of a sample of the category labels. */
  let widestCategory = 0;
  /** Works out the plot area inside `box` with everything that stands around it. */
  const settle = (): void => {
    plan = undefined;
    widestCategory = 0;
    catWidth = 0;
    if (lying) {
      // Category labels beside the plot area: as wide as the widest of them, up to a third of the frame.
      if (categoryLabels) {
        const stride = Math.max(Math.floor(count / 50), 1);
        for (let i = 0; i < count; i += stride) catWidth = Math.max(catWidth, measure(clean(values.label(i)), catFont));
        catWidth = Math.min(catWidth, box.width * 0.33) + tickGap;
      }
      arrange(0, 0);
    } else {
      const flat = categoryLabels ? lineBand(catFont) : 0;
      arrange(flat + outerBand, 0);
      if (categoryLabels && !scatter && !dates) {
        // The widest of a sample decides whether the labels fit, and how many are left out where they do not.
        const stride = Math.max(Math.floor(places / 50), 1);
        for (let i = 0; i < places; i += stride) widestCategory = Math.max(widestCategory, measure(clean(values.label(Math.floor((i * count) / places))), catFont));
        const slotNow = onTicks ? plot.width / (places - 1) : plot.width / Math.max(places, 1);
        const stated = catAxis ? catAxis.labelRotation : undefined;
        const line = catFont.size * 1.2;
        // Labels that would run into each other are turned: by 45 degrees while a slanted line fits its
        // category, then upright. Where not even that fits they lie flat and only every n-th is shown.
        let angle = 0;
        if (stated !== undefined && Number.isFinite(stated)) angle = clamp(Math.round(stated), -90, 90);
        else if (widestCategory + 8 > slotNow) angle = line * Math.SQRT2 <= slotNow ? -45 : line <= slotNow ? -90 : 0;
        if (angle === 0) {
          const wanted = Math.min(widestCategory, 120) + 8;
          const every = Math.max(Math.ceil(wanted / Math.max(slotNow, 1e-6)), 1);
          plan = { angle: 0, every, band: flat, room: 0 };
        } else {
          const turn = (Math.abs(angle) * Math.PI) / 180;
          const sin = Math.sin(turn);
          const cos = Math.cos(turn);
          // Turned labels take at most a third of the height; longer ones are cut.
          const room = Math.max(Math.min(widestCategory, (box.height * 0.3) / Math.max(sin, 0.2)), catFont.size);
          const pitch = Math.min(room * cos + 8, line / Math.max(sin, 1e-3));
          plan = { angle, every: Math.max(Math.ceil(pitch / Math.max(slotNow, 1e-6)), 1), band: room * sin + catFont.size * cos + tickGap, room };
          // A label that ends at its tick runs back to the left: the first one needs that much room before the plot area.
          const back = catSide === 'b' ? angle < 0 : angle > 0;
          let least = 0;
          if (back) {
            const reach = room * cos - box.x;
            const half = onTicks || catReversed ? 0 : 1 / (2 * Math.max(places, 1));
            least = (reach - (box.width - need.r) * half) / (1 - half);
          }
          arrange(plan.band + outerBand, Math.max(least, 0));
        }
      }
    }
  };
  settle();
  if (inner) {
    // The area of the series where the file puts it, when the labels and titles around it stay in the frame.
    const band = { l: 0, r: 0, t: 0, b: 0 };
    for (const title of titles) band[title.side] = title.font.size * 1.4;
    if (inner.x - need.l - band.l >= 0 && inner.y - need.t - band.t >= 0 && inner.x + inner.width + need.r + band.r <= layout.width && inner.y + inner.height + need.b + band.b <= layout.height) {
      box.x = inner.x - need.l;
      box.y = inner.y - need.t;
      box.width = inner.width + need.l + need.r;
      box.height = inner.height + need.t + need.b;
      settle();
      plot = { x: inner.x, y: inner.y, width: inner.width, height: inner.height };
    }
  }
  if (scatter) {
    for (const set of sets) {
      const logX = set.cat !== undefined && set.cat.logBase !== undefined;
      const seeX = (x: number | null): void => {
        if (x === null || !Number.isFinite(x) || (logX && !(x > 0))) return;
        if (x < set.xLow) set.xLow = x;
        if (x > set.xHigh) set.xHigh = x;
      };
      for (const item of scattered) {
        const xs = item.entry.xs;
        const ys = item.entry.values;
        if (!xs || item.set !== set) continue;
        if (ys.blank !== null) {
          // Every point has a y: every x counts.
          if (xs.known > 0) (seeX(xs.knownValue(0)), seeX(xs.knownValue(xs.known - 1)));
          xs.each((x, i) => {
            if (i < ys.length) seeX(x);
          });
        } else {
          // The points that are drawn: those with a number for both.
          for (let k = 0; k < ys.known; k++) seeX(xs.at(ys.knownIndex(k)));
        }
      }
      if (set.xLow === Infinity) {
        set.xLow = 0;
        set.xHigh = 0;
      }
      const wish = scaleOptions(set.cat, plot.width, Math.max(catFont.size * 3.2, 36));
      set.xScale = set.cat && set.cat.logBase ? chartLogScale(set.xLow, set.xHigh, set.cat.logBase, wish) : chartScale(set.xLow, set.xHigh, wish);
    }
  }
  const scale = primary.scale;
  const xScale = primary.xScale;
  layout.plot = plot;
  layout.scale = scale;
  layout.xScale = xScale;
  if (second) layout.scale2 = second.scale;
  const plotRight = plot.x + plot.width;
  const plotBottom = plot.y + plot.height;

  // The area of the series, when the file gives it a fill or an outline.
  const plotShape = chart.plotArea;
  if (plotShape) {
    const fill = plotShape.fill === null ? undefined : safe(plotShape.fill);
    const stroke = lineColor(plotShape.line, undefined);
    if (fill || stroke) items.push({ kind: 'rect', role: 'plot', x: plot.x, y: plot.y, width: plot.width, height: plot.height, fill, stroke, lineWidth: clamp((plotShape.line && plotShape.line.width) || 1, 0.5, 8) });
  }

  // Values to pixels along the value axis of their pair; categories to pixels along theirs.
  const valueAt = (set: AxisSet, value: number): number => {
    const part = shareOf(set.scale, value);
    const share = set.val !== undefined && set.val.reversed === true ? 1 - part : part;
    return lying ? plot.x + share * plot.width : plot.y + plot.height - share * plot.height;
  };
  const along = lying ? plot.height : plot.width;
  const slot = scatter ? 0 : dates ? along / dates.units : onTicks ? along / (places - 1) : along / Math.max(places, 1);
  /** A distance from the start of the category axis as a pixel. */
  const alongAt = (offset: number, reversed = catReversed): number => {
    // Categories of bars that lie run from the bottom up, as the values of an upright chart do.
    const flipped = lying ? !reversed : reversed;
    return lying ? plot.y + (flipped ? along - offset : offset) : plot.x + (flipped ? along - offset : offset);
  };
  const catAt = (position: number): number => {
    if (dates) return alongAt(((dates.position[position] || 0) + 0.5) * slot);
    return alongAt(onTicks ? position * slot : (position + 0.5) * slot);
  };
  const at = (set: AxisSet, cat: number, value: number): [number, number] => (lying ? [valueAt(set, value), catAt(cat)] : [catAt(cat), valueAt(set, value)]);
  /** Where the horizontal value `x` of a scatter chart stands; undefined for one outside the axis. */
  const xAt = (set: AxisSet, x: number): number | undefined => {
    const own = set.xScale;
    if (!own || x < own.min || x > own.max) return undefined;
    return alongAt(shareOf(own, x) * along, set.cat !== undefined && set.cat.reversed === true);
  };
  /**
   * Where a scatter point with the horizontal value `x` stands; undefined for one outside the axis.
   * UNSURE: beside bars, lines or areas there is no axis of numbers to place it on. The n-th category
   * stands at n on the category axis, so a point is put where its x lies among 1, 2, 3, ...: 2.5 is
   * halfway between the second and the third category. (Excel is known to draw such a combination on
   * the axes of the other chart type; how it counts along a category axis was not checked against it.)
   */
  const scatterAt = (set: AxisSet, x: number): number | undefined => {
    if (scatter) return xAt(set, x);
    // Where many categories share a position, so do the numbers between them.
    const position = ((x - 1) * places) / Math.max(count, 1);
    const offset = onTicks ? position * slot : (position + 0.5) * slot;
    return offset >= 0 && offset <= along ? alongAt(offset) : undefined;
  };

  // Grid lines under everything else.
  const gridLine = (line: ChartLine | undefined, automatic: string): { stroke: string | undefined; width: number; dash: boolean } => ({
    stroke: lineColor(line, automatic),
    width: clamp((line && line.width) || 1, 0.5, 8),
    dash: line !== undefined && line.dash === true,
  });
  const across = (position: number, look: { stroke: string | undefined; width: number; dash: boolean }, vertical: boolean): void => {
    if (!look.stroke) return;
    // A hairline is put on the centre of its pixel row.
    const snapped = look.width <= 1 ? Math.round(position - 0.5) + 0.5 : position;
    const points = vertical ? [snapped, plot.y, snapped, plot.y + plot.height] : [plot.x, snapped, plot.x + plot.width, snapped];
    items.push({ kind: 'path', role: 'grid', points, stroke: look.stroke, lineWidth: look.width, dash: look.dash });
  };
  if (detail !== 'bare') {
    for (const set of sets) {
      const val = set.val;
      if (!val || val.deleted) continue;
      if (val.minorGridlines) {
        const minor = gridLine(val.minorGridlines, MINOR_GRID_COLOR);
        if (set.scale.log) {
          // Between two powers: at their multiples (2, 3, ... times the lower one), where the powers follow each other.
          const base = set.scale.log;
          const parts = Math.min(Math.floor(base) - 1, 9);
          for (let i = 0; set.scale.step === 1 && i + 1 < set.scale.ticks.length; i++) {
            for (let k = 2; k <= parts; k++) {
              const value = set.scale.ticks[i] * (base <= 10 ? k : (k * base) / (parts + 1));
              if (value < set.scale.max) across(valueAt(set, value), minor, lying);
            }
          }
        } else {
          const fine = set.scale.step / 5;
          for (let i = 0; i <= MAX_AXIS_TICKS * 5; i++) {
            const value = set.scale.min + i * fine;
            if (value > set.scale.max + fine * 1e-6) break;
            if (i % 5 !== 0) across(valueAt(set, value), minor, lying);
          }
        }
      }
      if (val.majorGridlines) {
        const major = gridLine(val.majorGridlines, GRID_COLOR);
        for (const tick of set.scale.ticks) across(valueAt(set, tick), major, lying);
      }
    }
  }
  if (detail !== 'bare' && catAxis && !catAxis.deleted && catAxis.majorGridlines) {
    const major = gridLine(catAxis.majorGridlines, GRID_COLOR);
    if (xScale) {
      for (const tick of xScale.ticks) {
        const pixel = xAt(primary, tick);
        if (pixel !== undefined) across(pixel, major, !lying);
      }
    } else {
      const lines = Math.min(dates ? dates.units : places, 500);
      for (let i = 0; i <= lines; i++) across((lying ? plot.y : plot.x) + (i * along) / lines, major, !lying);
    }
  }

  // Where the series stand: on zero, or on the end of the axis when zero is not on it.
  // The category axis says where it crosses the value axis.
  for (const set of sets) {
    const cat = scatter ? undefined : set.cat;
    const crosses = cat ? cat.crosses : undefined;
    set.base =
      cat && cat.crossesAt !== undefined
        ? clamp(cat.crossesAt, set.scale.min, set.scale.max)
        : crosses === 'max'
          ? set.scale.max
          : crosses === 'min'
            ? set.scale.min
            : clamp(0, set.scale.min, set.scale.max);
  }
  const base = primary.base;

  /** The value a point of a series is drawn at and the value it stands on; undefined for a point that is not drawn. */
  const ends = (item: Plotted, i: number): [number, number] | undefined => {
    if (item.group.grouping !== 'standard' && item.to && item.from) {
      const top = item.to[i];
      return top === null || top === undefined ? undefined : [top, item.from[i] || 0];
    }
    const raw = item.values[i];
    return raw === null || raw === undefined ? undefined : [raw, item.set.base];
  };
  const within = (set: AxisSet, value: number): number => clamp(value, set.scale.min, set.scale.max);

  // Data labels: of the series whose every point is drawn, placed as the series are drawn and painted above everything.
  const placer = words ? new LabelPlacer(context, layout.width, layout.height) : undefined;

  // Error bars: gathered point by point as the series are drawn, painted above them.
  interface ErrorSet {
    bars: ChartErrorBars;
    own: { plus?: ChartNumbers; minus?: ChartNumbers } | undefined;
    /** The mean of the series and the amount every point has (standard deviation, standard error). */
    mean: number;
    amount: number;
    points: number[];
    series: number;
  }
  const errorSets: ErrorSet[] = [];
  const errorsOf = (entry: ChartSeriesValues, series: number, numbers: ChartNumbers | undefined, direction: 'x' | 'y'): ErrorSet | undefined => {
    const all = entry.series.errorBars;
    if (!Array.isArray(all) || !numbers) return undefined;
    for (let n = 0; n < all.length; n++) {
      const bars = all[n];
      if (!bars || bars.direction !== direction) continue;
      let mean = 0;
      let amount = 0;
      if (bars.valueType === 'stdDev' || bars.valueType === 'stdErr') {
        const known: number[] = [];
        numbers.each((value) => void known.push(value));
        const spread = seriesSpread(known);
        mean = spread.mean;
        // UNSURE: the standard error is the deviation of the sample over the root of the number of points.
        amount = bars.valueType === 'stdDev' ? spread.deviation * (bars.value !== undefined ? Math.abs(bars.value) : 1) : spread.error;
      }
      const set: ErrorSet = { bars, own: entry.errors ? entry.errors[n] : undefined, mean, amount, points: [], series };
      errorSets.push(set);
      return set;
    }
    return undefined;
  };
  /**
   * The values an error bar of the point `i` with the value `value` runs between; undefined for a point without one.
   * `unit`: what one of the series' own numbers is on the axis - 1, or for a series drawn as a share of 100 % one
   * over the sum of its category. UNSURE: Excel's rule for error bars of shares is not known; an amount is taken
   * as a share of the same sum, so that a bar is as long beside its column as the amount is beside the value.
   */
  const errorSpan = (set: ErrorSet, i: number, value: number, centre: number, unit = 1): [number, number] | undefined => {
    const bars = set.bars;
    let plus = 0;
    let minus = 0;
    let middle = centre;
    if (bars.valueType === 'cust') {
      const up = set.own && set.own.plus ? set.own.plus.at(i) : null;
      const down = set.own && set.own.minus ? set.own.minus.at(i) : null;
      plus = up === null ? 0 : Math.abs(up);
      minus = down === null ? 0 : Math.abs(down);
    } else if (bars.valueType === 'percentage') plus = minus = (Math.abs(value) * Math.abs(bars.value !== undefined ? bars.value : 5)) / 100;
    else if (bars.valueType === 'fixedVal') plus = minus = Math.abs(bars.value !== undefined ? bars.value : 0);
    else {
      plus = minus = set.amount;
      // UNSURE: bars of the standard deviation stand around the mean of the series, the same at every point (as Excel is known to draw them).
      if (bars.valueType === 'stdDev') middle = centre + (set.mean - value) * unit;
    }
    plus *= unit;
    minus *= unit;
    if (bars.type === 'plus') minus = 0;
    if (bars.type === 'minus') plus = 0;
    const from = middle - minus;
    const to = middle + plus;
    return Number.isFinite(from) && Number.isFinite(to) && to > from ? [from, to] : undefined;
  };
  const CAP = 4;
  /** An error bar along the value axis of a point that stands at `across` on the category axis. */
  const valueError = (set: ErrorSet, axes: AxisSet, i: number, value: number, centre: number, across: number, unit = 1): void => {
    const span = errorSpan(set, i, value, centre, unit);
    if (!span) return;
    const a = valueAt(axes, within(axes, span[0]));
    const b = valueAt(axes, within(axes, span[1]));
    // Caps end where the frame does.
    const far = lying ? layout.height : layout.width;
    const line = (x1: number, y1: number, x2: number, y2: number): void => void (lying ? set.points.push(y1, clamp(x1, 0, far), y2, clamp(x2, 0, far)) : set.points.push(clamp(x1, 0, far), y1, clamp(x2, 0, far), y2));
    line(across, a, across, b);
    if (set.bars.noEndCap) return;
    // A cap at each end the bar reaches out to.
    if (set.bars.type !== 'plus') line(across - CAP, a, across + CAP, a);
    if (set.bars.type !== 'minus') line(across - CAP, b, across + CAP, b);
  };

  // Trend lines: worked out from the numbers of a series, drawn above the series and cut to the plot area.
  let trendsDrawn = 0;
  const trendItems: ChartItem[] = [];
  /**
   * `xs` and `ys`: the points; `pixel`: where a value at a place along the horizontal axis is drawn, undefined
   * for what has no place. `curved`: the axes are such that a straight trend is no straight line.
   */
  // Points of one moving average that are drawn: of more, each run shows its lowest and its highest.
  const trendPoints = Math.max(Math.min(Math.max(Math.floor(alongGuess), 16), MAX_DRAWN_POINTS, Math.floor(context.budget / MAX_DRAWN_TRENDS)), 16);
  /** `visible`: the places along the horizontal axis that lie in the plot area, from the first to the last. */
  const drawTrends = (series: ChartSeries, look: SeriesLook, index: number, xs: number[], ys: number[], pixel: (x: number, y: number) => [number, number] | undefined, curved: boolean, visible: [number, number]): void => {
    const list = series.trendlines;
    if (!Array.isArray(list) || xs.length < 2) return;
    let first = Infinity;
    let last = -Infinity;
    for (const x of xs) {
      if (x < first) first = x;
      if (x > last) last = x;
    }
    for (const trend of list as ChartTrendline[]) {
      if (!trend || trendsDrawn >= MAX_DRAWN_TRENDS) return;
      const fit = fitTrend(trend.type, xs, ys, { order: trend.order, period: trend.period });
      if (!fit) continue;
      const points: number[] = [];
      const put = (x: number, y: number | undefined): void => {
        const at2 = y === undefined || !Number.isFinite(y) ? undefined : pixel(x, y);
        // What has no place breaks the line. The line is cut to the plot area as it is: its places along the horizontal
        // axis lie within the axis already, and a value far beyond the other one is only kept from being no number.
        if (!at2 || !Number.isFinite(at2[0]) || !Number.isFinite(at2[1])) points.push(NaN, NaN);
        else points.push(clamp(at2[0], -1e12, 1e12), clamp(at2[1], -1e12, 1e12));
      };
      if (fit.points) {
        const all = fit.points;
        const count2 = Math.floor(all.length / 2);
        if (count2 <= trendPoints * 2) for (let i = 0; i + 1 < all.length; i += 2) put(all[i], all[i + 1]);
        else {
          // A moving average of very many points: like its series, each run of points shows its lowest and its highest.
          const runs = Math.max(Math.floor(trendPoints / 2), 1);
          for (let run = 0; run < runs; run++) {
            const start = Math.floor((run * count2) / runs);
            const end = Math.max(Math.floor(((run + 1) * count2) / runs), start + 1);
            let low = start;
            let high = start;
            for (let k = start + 1; k < end && k < count2; k++) {
              if (all[k * 2 + 1] < all[low * 2 + 1]) low = k;
              if (all[k * 2 + 1] > all[high * 2 + 1]) high = k;
            }
            const before = Math.min(low, high);
            const after = Math.max(low, high);
            put(all[before * 2], all[before * 2 + 1]);
            if (after !== before) put(all[after * 2], all[after * 2 + 1]);
          }
        }
      } else {
        // Only as far as the axis reaches: what lies beyond it is not drawn, and need not be worked out.
        const from = Math.max(first - (trend.backward !== undefined && trend.backward > 0 && Number.isFinite(trend.backward) ? trend.backward : 0), visible[0]);
        const to = Math.min(last + (trend.forward !== undefined && trend.forward > 0 && Number.isFinite(trend.forward) ? trend.forward : 0), visible[1]);
        if (!(to > from)) continue;
        const steps = trend.type === 'linear' && !curved ? 1 : 48;
        for (let n = 0; n <= steps; n++) {
          const x = from + ((to - from) * n) / steps;
          put(x, fit.at(x));
        }
      }
      const runs = clipLine(points, plot);
      if (runs.length === 0) continue;
      trendsDrawn++;
      const stated = trend.line;
      // UNSURE: a trend line the file gives no look is dashed, two pixels wide, in the colour of its series.
      const stroke = lineColor(stated, look.color);
      if (!stroke) continue;
      for (const run of runs) {
        trendItems.push({ kind: 'path', role: 'trend', points: run, stroke, lineWidth: stated && stated.width !== undefined && stated.width > 0 ? clamp(stated.width, 0.5, 32) : 2, dash: stated ? (stated.dash === true ? true : undefined) : true, series: index });
      }
    }
  };
  /** Where the n-th category (counted from one, fractions between two of them) stands along the category axis. */
  const categoryPixel = (x: number): number => {
    if (dates) return alongAt((x - 0.5) * slot);
    const position = ((x - 1) * places) / Math.max(count, 1);
    return alongAt(onTicks ? position * slot : (position + 0.5) * slot);
  };
  /** The category numbers at the two ends of the category axis. */
  const categoriesVisible = ((): [number, number] => {
    if (dates) return [0.5, dates.units + 0.5];
    const each = Math.max(count, 1) / Math.max(places, 1);
    return onTicks ? [1, 1 + (places - 1) * each] : [1 - each / 2, 1 + (places - 0.5) * each];
  })();

  // Lines that belong to a chart type, not to a series: from the highest point of each category down to the axis,
  // from the highest to the lowest, and boxes from the first series to the last.
  const groupLines = (): void => {
    for (const group of chart.groups) {
      if (!group.dropLines && !group.hiLowLines && !group.upDownBars) continue;
      const members = plotted.filter((item) => item.group === group);
      if (members.length === 0) continue;
      const set = members[0].set;
      const drop: number[] = [];
      const hiLow: number[] = [];
      // The boxes are drawn over the lines.
      const boxItems: ChartItem[] = [];
      const boxes = group.upDownBars;
      const boxWidth = slot / (1 + clamp(boxes && boxes.gapWidth !== undefined ? boxes.gapWidth : 150, 0, 500) / 100);
      for (let i = 0; i < places; i++) {
        let high = -Infinity;
        let low = Infinity;
        let found = 0;
        for (const member of members) {
          const span2 = ends(member, i);
          if (!span2) continue;
          found++;
          if (span2[0] > high) high = span2[0];
          if (span2[0] < low) low = span2[0];
        }
        if (found === 0) continue;
        const top = at(set, i, within(set, high));
        if (group.dropLines) {
          const foot = at(set, i, set.base);
          drop.push(top[0], top[1], foot[0], foot[1]);
        }
        if (group.hiLowLines && found >= 2) {
          const foot = at(set, i, within(set, low));
          hiLow.push(top[0], top[1], foot[0], foot[1]);
        }
        if (boxes && members.length >= 2) {
          const open = ends(members[0], i);
          const close = ends(members[members.length - 1], i);
          if (!open || !close) continue;
          const up = close[0] >= open[0];
          const shape = up ? boxes.up : boxes.down;
          // UNSURE: a box that rose is white and one that fell is black, both outlined in black, where the file states no look.
          const automatic = up ? '#FFFFFF' : '#000000';
          const fill = shape && shape.fill !== undefined ? (shape.fill === null ? undefined : safe(shape.fill) || automatic) : automatic;
          const a = valueAt(set, within(set, open[0]));
          const b = valueAt(set, within(set, close[0]));
          const centre = catAt(i);
          // A box of a category that stands on the edge of the plot area is cut there.
          const from = Math.max(centre - boxWidth / 2, lying ? plot.y : plot.x);
          const to = Math.min(centre + boxWidth / 2, lying ? plot.y + plot.height : plot.x + plot.width);
          if (!(to >= from)) continue;
          const rect = lying ? { x: Math.min(a, b), y: from, width: Math.abs(a - b), height: to - from } : { x: from, y: Math.min(a, b), width: to - from, height: Math.abs(a - b) };
          boxItems.push({ kind: 'rect', role: 'updown', ...rect, fill, stroke: lineColor(shape ? shape.line : undefined, '#000000'), lineWidth: clamp((shape && shape.line && shape.line.width) || 1, 0.5, 8), point: i });
        }
      }
      const thin = (line: ChartLine | undefined): { stroke: string | undefined; lineWidth: number; dash: boolean } => ({ stroke: lineColor(line, TEXT_COLOR), lineWidth: clamp((line && line.width) || 1, 0.5, 8), dash: line !== undefined && line.dash === true });
      if (drop.length > 0 && thin(group.dropLines).stroke) items.push({ kind: 'path', role: 'drop', points: drop, segments: true, ...thin(group.dropLines) });
      if (hiLow.length > 0 && thin(group.hiLowLines).stroke) items.push({ kind: 'path', role: 'hilow', points: hiLow, segments: true, ...thin(group.hiLowLines) });
      for (const item of boxItems) items.push(item);
    }
  };

  const bars = plotted.filter((item) => item.group.kind === 'bar');
  for (const kind of ['area', 'bar', 'line'] as ChartKind[]) {
    // Under the lines and their markers, above the areas and the bars.
    if (kind === 'line') groupLines();
    for (const item of plotted) {
      if ((isLine(item.group.kind) ? 'line' : item.group.kind) !== kind) continue;
      const group = item.group;
      const set = item.set;
      const stacked = group.grouping !== 'standard';
      const labels = placer && places === count ? placer.of(item.entry) : undefined;
      const errors = places === count ? errorsOf(item.entry, item.index, item.entry.values, 'y') : undefined;
      if (!stacked && group.kind !== 'radar' && Array.isArray(item.entry.series.trendlines) && !(dates && places < count)) {
        // The n-th point stands at n: the first category is 1. On a date axis a point stands at its unit of time.
        const xs: number[] = [];
        const ys: number[] = [];
        if (places === count) {
          for (let i = 0; i < places; i++) {
            const value = item.values[i];
            if (value === null || value === undefined) continue;
            xs.push(dates ? dates.position[i] + 1 : i + 1);
            ys.push(value);
          }
        } else {
          const positive = set.val !== undefined && set.val.logBase !== undefined;
          item.entry.values.each((value, i) => {
            if (positive && !(value > 0)) return;
            xs.push(i + 1);
            ys.push(value);
          });
        }
        drawTrends(
          item.entry.series,
          item.look,
          item.index,
          xs,
          ys,
          (x, y) => {
            if (set.scale.log && !(y > 0)) return undefined;
            const across2 = categoryPixel(x);
            const value = valueAt(set, y);
            return lying ? [value, across2] : [across2, value];
          },
          set.scale.log !== undefined,
          categoriesVisible,
        );
      }
      // A series drawn as a share of 100 %: what one of its numbers is on the axis, category by category.
      const pile = group.grouping === 'percentStacked' ? piles.get(item.entry.group) : undefined;
      const unitAt = (i: number): number => (!pile ? 1 : pile.total[i] > 0 ? 1 / pile.total[i] : 0);

      if (kind === 'bar') {
        const cluster = stacked ? [item] : bars.filter((other) => other.group === group);
        const members = cluster.length;
        const place = cluster.indexOf(item);
        const overlap = stacked ? 0 : clamp(group.overlap === undefined ? 0 : group.overlap, -100, 100) / 100;
        const gapWidth = clamp(group.gapWidth === undefined ? 150 : group.gapWidth, 0, 500) / 100;
        const barWidth = slot / (1 + (members - 1) * (1 - overlap) + gapWidth);
        const offset = (barWidth * gapWidth) / 2 + place * barWidth * (1 - overlap);
        const vary = group.varyColors === true && plotted.length === 1;
        for (let i = 0; i < places; i++) {
          const span2 = ends(item, i);
          if (!span2) continue;
          const a = valueAt(set, within(set, span2[0]));
          const b = valueAt(set, within(set, span2[1]));
          const centre = catAt(i);
          // Bars that lie are counted from the bottom of their slot, like their categories.
          const flipped = lying ? !catReversed : catReversed;
          const start = flipped ? centre + slot / 2 - offset - barWidth : centre - slot / 2 + offset;
          const point = pointLook(chart, item.look, places < count ? -1 : i, vary);
          let fill = point.fill;
          let stroke = point.stroke;
          // UNSURE: "invert if negative" without a second colour shows the bar white with the outline of its series.
          if (item.entry.series.invertIfNegative && (item.values[i] || 0) < 0) {
            stroke = stroke || fill || item.look.color;
            fill = '#FFFFFF';
          }
          const rect = lying
            ? { x: Math.min(a, b), y: start, width: Math.abs(a - b), height: barWidth }
            : { x: start, y: Math.min(a, b), width: barWidth, height: Math.abs(a - b) };
          items.push({ kind: 'rect', role: 'bar', ...rect, fill, stroke, lineWidth: point.lineWidth, series: item.index, point: i });
          if (errors) valueError(errors, set, i, item.values[i] || 0, span2[0], start + barWidth / 2, unitAt(i));
          if (placer && labels) {
            const value = item.values[i] || 0;
            placer.add(item.entry, labels, i, value, { kind: 'bar', ...rect, lying, low: lying ? a < b || (a === b && value < 0) : a < b || (a === b && !(value < 0)), stacked });
          }
        }
        continue;
      }

      if (kind === 'area') {
        // The outline along the values and back along what the series stands on.
        const top: number[] = [];
        const bottom: number[] = [];
        for (let n = 0; n < places; n++) {
          const i = order ? order[n] : n;
          // UNSURE: a point without a number is drawn as zero in an area.
          const span2 = ends(item, i) || [set.base, set.base];
          const p = at(set, i, within(set, span2[0]));
          const q = at(set, i, within(set, span2[1]));
          top.push(p[0], p[1]);
          bottom.push(q[0], q[1]);
          const value = item.values[i];
          if (errors && value !== null && value !== undefined) valueError(errors, set, i, value, span2[0], catAt(i), unitAt(i));
          if (placer && labels && value !== null && value !== undefined) placer.add(item.entry, labels, i, value, { kind: 'point', x: (p[0] + q[0]) / 2, y: (p[1] + q[1]) / 2, radius: 0, centred: true });
        }
        const points = top.slice();
        for (let i = bottom.length - 2; i >= 0; i -= 2) points.push(bottom[i], bottom[i + 1]);
        if (points.length >= 6) items.push({ kind: 'path', role: 'area', points, closed: true, fill: item.look.fill, stroke: item.look.stroke, lineWidth: item.look.lineWidth, series: item.index });
        continue;
      }

      // Lines: runs of points, broken where a point has no number unless the chart spans such points.
      const trace = new Trace(chart.blanks === 'span', item.entry.series.smooth === true && !(item.low && item.high) ? layout : undefined);
      for (let n = 0; n < places; n++) {
        const i = order ? order[n] : n;
        if (item.low && item.high) {
          const lo = item.low[i];
          const hi = item.high[i];
          if (lo === null || hi === null) {
            trace.gap();
            continue;
          }
          const p = at(set, i, within(set, lo));
          const q = at(set, i, within(set, hi));
          trace.push(p[0], p[1], false);
          if (hi !== lo) trace.push(q[0], q[1], false);
          continue;
        }
        const span2 = ends(item, i);
        if (!span2) {
          trace.gap();
          continue;
        }
        const p = at(set, i, within(set, span2[0]));
        trace.push(p[0], p[1], true);
        if (errors) valueError(errors, set, i, item.values[i] || 0, span2[0], catAt(i), unitAt(i));
        if (placer && labels) placer.add(item.entry, labels, i, item.values[i] as number, { kind: 'point', x: p[0], y: p[1], radius: item.look.marker ? item.look.marker.size / 2 : 2 });
      }
      trace.emit(items, item.look, item.index, !(item.low && item.high), true);
    }
  }

  // Scatter series, above the others: every point at its own x. Of very many points every n-th one is drawn,
  // so that no series asks for more markers than its share of what a frame is worth painting.
  // The size of the largest bubble that is drawn decides how large the others are.
  let largest = 0;
  for (const item of scattered) {
    if (item.group.kind !== 'bubble') continue;
    const sizes = item.entry.sizes;
    const negative = item.group.showNegBubbles === true;
    if (!sizes) largest = Math.max(largest, 1);
    else {
      sizes.each((size) => {
        const taken = negative ? Math.abs(size) : size;
        if (taken > largest && Number.isFinite(taken)) largest = taken;
      });
    }
  }
  for (const item of scattered) {
    const xs = item.entry.xs;
    const ys = item.entry.values;
    const set = item.set;
    if (!xs) continue;
    const logY = set.val !== undefined && set.val.logBase !== undefined;
    const logX = scatter && set.cat !== undefined && set.cat.logBase !== undefined;
    const bubble = item.group.kind === 'bubble';
    const smooth = item.entry.series.smooth === true || item.group.scatterStyle === 'smooth' || item.group.scatterStyle === 'smoothMarker';
    const trace = new Trace(chart.blanks === 'span', smooth ? layout : undefined);
    const plain = ys.blank === null;
    // Points without a number are skipped where they are nothing; where they are zero every point is one.
    const total = plain ? ys.known : ys.length;
    const stride = Math.max(Math.ceil(total / drawnScatter), 1);
    const labels = placer && stride === 1 ? placer.of(item.entry) : undefined;
    const yErrors = stride === 1 ? errorsOf(item.entry, item.index, ys, 'y') : undefined;
    const xErrors = stride === 1 && scatter ? errorsOf(item.entry, item.index, xs, 'x') : undefined;
    // Bubbles: x, y and the radius of each. UNSURE: the largest bubble of the chart is a quarter of the smaller
    // side of the plot area across (times what the chart type says), and a bubble of a negative size, where the
    // chart shows such, is drawn by the size without its sign: white with the outline of its series.
    const full: number[] = [];
    const hollow: number[] = [];
    const sizes = item.entry.sizes;
    const widest = (Math.min(plot.width, plot.height) * 0.125 * clamp(item.group.bubbleScale === undefined ? 100 : item.group.bubbleScale, 0, 300)) / 100;
    if (Array.isArray(item.entry.series.trendlines)) {
      const tx: number[] = [];
      const ty: number[] = [];
      for (let k = 0; k < ys.known; k++) {
        const i = ys.knownIndex(k);
        const x = xs.at(i);
        const y = ys.knownValue(k);
        if (x === null || (logY && !(y > 0)) || (logX && !(x > 0))) continue;
        tx.push(x);
        ty.push(y);
      }
      drawTrends(
        item.entry.series,
        item.look,
        item.index,
        tx,
        ty,
        (x, y) => {
          if ((logY && !(y > 0)) || (logX && !(x > 0))) return undefined;
          const own = set.xScale;
          const across2 = scatter && own ? alongAt(shareOf(own, x) * along, set.cat !== undefined && set.cat.reversed === true) : categoryPixel(x);
          const value = valueAt(set, y);
          return lying ? [value, across2] : [across2, value];
        },
        logX || logY,
        scatter && set.xScale ? [set.xScale.min, set.xScale.max] : categoriesVisible,
      );
    }
    let before = -1;
    for (let k = 0; k < total; k += stride) {
      const i = plain ? ys.knownIndex(k) : k;
      const y = plain ? ys.knownValue(k) : ys.at(i);
      const x = xs.at(i);
      // A point that was left out between two that are drawn, because it has no number, breaks the line.
      if (stride === 1 && before >= 0 && i !== before + 1) trace.gap();
      before = i;
      const px = x === null || y === null || (logY && !(y > 0)) ? undefined : scatterAt(set, x);
      if (px === undefined || y === null) {
        trace.gap();
        continue;
      }
      const py = valueAt(set, within(set, y));
      let radius = item.look.marker ? item.look.marker.size / 2 : 2;
      let size: number | undefined;
      if (bubble) {
        const stated = sizes ? sizes.at(i) : 1;
        // A bubble without a size, of the size nothing, or of a negative size the chart does not show, is not drawn.
        if (stated === null || stated === 0 || !Number.isFinite(stated) || (stated < 0 && item.group.showNegBubbles !== true) || !(largest > 0)) continue;
        const share = Math.min(Math.abs(stated) / largest, 1);
        radius = widest * (item.group.sizeRepresents === 'w' ? share : Math.sqrt(share));
        size = stated;
        (stated < 0 ? hollow : full).push(lying ? py : px, lying ? px : py, radius);
      } else if (lying) trace.push(py, px, true);
      else trace.push(px, py, true);
      if (yErrors) valueError(yErrors, set, i, y, y, px);
      if (xErrors && x !== null && set.xScale) {
        const span = errorSpan(xErrors, i, x, x);
        const own = set.xScale;
        if (span) {
          const a = xAt(set, clamp(span[0], own.min, own.max));
          const b = xAt(set, clamp(span[1], own.min, own.max));
          if (a !== undefined && b !== undefined) {
            xErrors.points.push(a, py, b, py);
            const up = Math.max(py - CAP, 0);
            const down = Math.min(py + CAP, layout.height);
            if (!xErrors.bars.noEndCap && xErrors.bars.type !== 'plus') xErrors.points.push(a, up, a, down);
            if (!xErrors.bars.noEndCap && xErrors.bars.type !== 'minus') xErrors.points.push(b, up, b, down);
          }
        }
      }
      if (placer && labels && x !== null) placer.add(item.entry, labels, i, y, { kind: 'point', x: lying ? py : px, y: lying ? px : py, radius }, { x, size });
    }
    if (bubble) {
      const fill = item.look.fill || item.look.color;
      const vary = item.group.varyColors === true && scattered.length === 1 && full.length <= 3 * 256;
      if (vary) {
        // A colour for each bubble.
        for (let i = 0; i + 2 < full.length; i += 3) items.push({ kind: 'circles', role: 'bubble', points: full.slice(i, i + 3), fill: pointLook(chart, item.look, i / 3, true).fill || fill, stroke: item.look.stroke, lineWidth: item.look.lineWidth, series: item.index });
      } else if (full.length > 0) items.push({ kind: 'circles', role: 'bubble', points: full, fill, stroke: item.look.stroke, lineWidth: item.look.lineWidth, series: item.index });
      if (hollow.length > 0) items.push({ kind: 'circles', role: 'bubble', points: hollow, fill: '#FFFFFF', stroke: item.look.color, lineWidth: 1, series: item.index });
    } else trace.emit(items, item.look, item.index, true, false);
  }

  // Error bars and trend lines, above the series.
  for (const set of errorSets) {
    if (set.points.length === 0) continue;
    const stated = set.bars.line;
    const stroke = lineColor(stated, TEXT_COLOR);
    if (stroke) items.push({ kind: 'path', role: 'error', points: set.points, segments: true, stroke, lineWidth: stated && stated.width !== undefined && stated.width > 0 ? clamp(stated.width, 0.5, 8) : 1, dash: stated !== undefined && stated.dash === true ? true : undefined, series: set.series });
  }
  for (const item of trendItems) items.push(item);

  if (detail === 'bare') return;

  // The lines of the axes, above the series.
  const axisLine = (axis: ChartAxis | undefined, points: number[]): void => {
    if (axis && axis.deleted) return;
    const look = gridLine(axis ? axis.line : undefined, GRID_COLOR);
    if (!look.stroke) return;
    const snapped = points.map((value, index) => (look.width <= 1 && points[index % 2] === points[(index % 2) + 2] ? Math.round(value - 0.5) + 0.5 : value));
    items.push({ kind: 'path', role: 'axis', points: snapped, stroke: look.stroke, lineWidth: look.width, dash: look.dash });
  };
  /** A line along one side of the plot area. */
  const edgeLine = (side: Side): number[] =>
    side === 'l' ? [plot.x, plot.y, plot.x, plotBottom] : side === 'r' ? [plotRight, plot.y, plotRight, plotBottom] : side === 't' ? [plot.x, plot.y, plotRight, plot.y] : [plot.x, plotBottom, plotRight, plotBottom];
  const basePixel = valueAt(primary, base);
  axisLine(catAxis, lying ? [basePixel, plot.y, basePixel, plotBottom] : [plot.x, basePixel, plotRight, basePixel]);
  for (const set of sets) axisLine(set.val, edgeLine(set.side));
  if (cat2) axisLine(cat2, edgeLine(cat2Side));
  if (!words) return;

  /** The baseline of a line of text that lies above or below the plot area, `gap` away from it. */
  const lineY = (side: Side, font: ChartFont, gap: number): number => (side === 't' ? plot.y - gap - font.size * 0.25 : plotBottom + gap + font.size * 0.85);

  // Labels of the value axes: all of them when there is room, else every n-th.
  for (const set of sets) {
    if (!set.labels) continue;
    const { tickLabels, font, color } = set;
    if (lying) {
      let widest = 0;
      for (const label of tickLabels) widest = Math.max(widest, measure(label, font));
      const every = Math.max(Math.ceil(((widest + 8) * tickLabels.length) / Math.max(plot.width, 1)), 1);
      for (let i = 0; i < tickLabels.length; i += every) {
        if (tickLabels[i] !== '') items.push({ kind: 'text', role: 'tick', text: tickLabels[i], x: valueAt(set, set.scale.ticks[i]), y: lineY(set.labelSide, font, tickGap), font, color, align: 'center' });
      }
    } else {
      const every = Math.max(Math.ceil((font.size * 1.1 * tickLabels.length) / Math.max(plot.height, 1)), 1);
      const room = set.band - tickGap;
      const left = set.labelSide === 'l';
      for (let i = 0; i < tickLabels.length; i += every) {
        const text = fit(tickLabels[i], font, room, measure);
        if (text !== '') items.push({ kind: 'text', role: 'tick', text, x: left ? plot.x - tickGap : plotRight + tickGap, y: valueAt(set, set.scale.ticks[i]) + font.size * 0.35, font, color, align: left ? 'right' : 'left' });
      }
    }
  }

  // Labels of the categories.
  /** The labels of a horizontal axis of numbers: all of them when there is room, else every n-th. */
  const numberLabels = (set: AxisSet, axis: ChartAxis | undefined, side: Side, font: ChartFont, color: string): void => {
    const own = set.xScale;
    if (!own) return;
    const labels = own.ticks.map((tick) => tickText(tick, axis && axis.formatCode && !axis.sourceLinked ? { code: axis.formatCode } : undefined, options));
    let widest = 0;
    for (const label of labels) widest = Math.max(widest, measure(label, font));
    const every = Math.max(Math.ceil(((widest + 8) * labels.length) / Math.max(plot.width, 1)), 1);
    for (let i = 0; i < labels.length; i += every) {
      const x = xAt(set, own.ticks[i]);
      if (labels[i] !== '' && x !== undefined) items.push({ kind: 'text', role: 'category', text: labels[i], x, y: lineY(side, font, tickGap), font, color, align: 'center' });
    }
  };
  /** Category labels that lie flat above or below the plot area: where they would run into each other only every n-th is shown, cut to the room it has. */
  const flatLabels = (side: Side, font: ChartFont, color: string, every: number): void => {
    const room = Math.max(slot * every - 4, 0);
    for (let i = 0; i < places && i / every < 2000; i += every) {
      const text = fit(clean(values.label(Math.floor((i * count) / places))), font, room, measure);
      if (text !== '') items.push({ kind: 'text', role: 'category', text, x: catAt(i), y: lineY(side, font, tickGap), font, color, align: 'center' });
    }
  };
  if (categoryLabels) {
    if (scatter) numberLabels(primary, catAxis, catSide, catFont, catColor);
    else if (lying) {
      const every = Math.max(Math.ceil((catFont.size * 1.2) / Math.max(slot, 1e-6)), 1);
      const left = catSide === 'l';
      for (let i = 0; i < places && i / every < 2000; i += every) {
        const text = fit(clean(values.label(Math.floor((i * count) / places))), catFont, catWidth - tickGap, measure);
        if (text !== '') items.push({ kind: 'text', role: 'category', text, x: left ? plot.x - tickGap : plotRight + tickGap, y: catAt(i) + catFont.size * 0.35, font: catFont, color: catColor, align: left ? 'right' : 'left' });
      }
    } else if (dates) {
      // A label where a step of the axis starts, with or without a point there.
      const stated = categoryFormat(catAxis);
      const own = values.categoryFormat;
      const format: NumberFormat = stated !== undefined ? { code: stated } : own && isDateCode(own.code, own.builtinId) ? own : { code: 'm/d/yyyy', builtinId: 14 };
      const first = dates.unit === 'days' ? dates.first : dates.unit === 'months' ? serialOfDay(Math.floor(dates.first / 12), (dates.first % 12) + 1, 1, date1904) : serialOfDay(dates.first, 1, 1, date1904);
      const start = first === undefined ? undefined : serialToDate(first, date1904);
      if (first !== undefined && start && Number.isFinite(start.year)) {
        const sample = tickText(first, format, options);
        const room = Math.max(Math.floor(along / (measure(sample, catFont) + 8)), 1);
        const span = dates.units * UNIT_DAYS[dates.unit];
        let unit: TimeUnit = catAxis && catAxis.majorTimeUnit && UNIT_RANK[catAxis.majorTimeUnit] >= UNIT_RANK[dates.unit] ? catAxis.majorTimeUnit : dates.unit;
        let step = catAxis && catAxis.majorUnit !== undefined ? Math.max(Math.floor(catAxis.majorUnit), 1) : 0;
        if (step === 0) {
          // The finest step whose labels have room.
          let pick = DATE_STEPS[DATE_STEPS.length - 1];
          for (const choice of DATE_STEPS) {
            if (UNIT_RANK[choice[0]] < UNIT_RANK[dates.unit] || span / (UNIT_DAYS[choice[0]] * choice[1]) > room) continue;
            pick = choice;
            break;
          }
          unit = pick[0];
          step = pick[1];
        } else {
          // A stated step that asks for more ticks than an axis has is doubled until it does not.
          for (let i = 0; i < 64 && span / (UNIT_DAYS[unit] * step) > MAX_AXIS_TICKS; i++) step *= 2;
        }
        const ticks: { at: number; serial: number }[] = [];
        let serial: number | undefined = first;
        let year = start.year;
        let month = start.month;
        // Every tick aims at the day of the first; a month that has no such day gives its last one (31 January,
        // 29 February, 31 March).
        const day = Math.max(start.day, 1);
        const dayIn = (y: number, m: number): number | undefined => {
          for (let d = day; d >= 28; d--) {
            const found = serialOfDay(y, m, d, date1904);
            if (found !== undefined) return found;
          }
          return day < 28 ? serialOfDay(y, m, day, date1904) : undefined;
        };
        for (let n = 0; n <= MAX_AXIS_TICKS && serial !== undefined; n++) {
          const unitAt = timeUnitOf(serial, dates.unit, date1904);
          if (unitAt === undefined || unitAt - dates.first > dates.units - 1) break;
          ticks.push({ at: unitAt - dates.first, serial });
          if (unit === 'days') serial += step;
          else {
            if (unit === 'months') {
              month += step;
              year += Math.floor((month - 1) / 12);
              month = ((month - 1) % 12) + 1;
            } else year += step;
            serial = unit === 'years' && dates.unit === 'years' ? serialOfDay(year, month, 1, date1904) : dayIn(year, month);
          }
        }
        // Labels of a stated step that would run into each other: every n-th.
        const every = Math.max(Math.ceil(ticks.length / room), 1);
        for (let i = 0; i < ticks.length; i += every) {
          const text = i === 0 ? sample : tickText(ticks[i].serial, format, options);
          if (text !== '') items.push({ kind: 'text', role: 'category', text, x: alongAt((ticks[i].at + 0.5) * slot), y: lineY(catSide, catFont, tickGap), font: catFont, color: catColor, align: 'center' });
        }
      }
    } else if (plan && plan.angle !== 0) {
      // Turned labels: each ends (or starts) at its tick and runs away from the plot area.
      const turn = (plan.angle * Math.PI) / 180;
      const sin = Math.abs(Math.sin(turn));
      const cos = Math.cos(turn);
      const below = catSide === 'b';
      // Below the axis a label turned upwards ends at its tick; above the axis it starts there.
      const back = below ? plan.angle < 0 : plan.angle > 0;
      const y = below ? plotBottom + tickGap + catFont.size * 0.7 * cos : plot.y - tickGap - catFont.size * 0.1 * cos;
      for (let i = 0; i < places && i / plan.every < 2000; i += plan.every) {
        // The middle of the height of the letters lies on the tick.
        const x = catAt(i) + (plan.angle < 0 ? 1 : -1) * catFont.size * 0.35 * sin;
        // No further than the edge of the frame.
        const reach = cos > 1e-6 ? (back ? x : layout.width - x) / cos : Infinity;
        const text = fit(clean(values.label(Math.floor((i * count) / places))), catFont, Math.min(plan.room, Math.max(reach, 0)), measure);
        if (text !== '') items.push({ kind: 'text', role: 'category', text, x, y, font: catFont, color: catColor, align: back ? 'right' : 'left', angle: plan.angle });
      }
    } else flatLabels(catSide, catFont, catColor, plan ? plan.every : 1);
    if (outerLabel && plan) {
      // The outer level: each label across the categories up to the next one, beyond the labels of the inner level.
      const inner = plan.band;
      const y = catSide === 't' ? plot.y - inner - catFont.size * 0.35 : plotBottom + inner + catFont.size * 0.85;
      let from = -1;
      let text = '';
      const flush = (to: number): void => {
        if (from < 0 || text === '') return;
        const a = catAt(from) - (catReversed ? -1 : 1) * (onTicks ? 0 : slot / 2);
        const b = catAt(to - 1) + (catReversed ? -1 : 1) * (onTicks ? 0 : slot / 2);
        const shown2 = fit(text, catFont, Math.max(Math.abs(b - a) - 4, 0), measure);
        if (shown2 !== '') items.push({ kind: 'text', role: 'category', text: shown2, x: (a + b) / 2, y, font: catFont, color: catColor, align: 'center' });
      };
      for (let i = 0; i < places && i < 4000; i++) {
        const label = clean(outerLabel(Math.floor((i * count) / places)));
        if (label === '') continue;
        flush(i);
        from = i;
        text = label;
      }
      flush(places);
    }
  }
  if (category2Labels && second) {
    if (scatter) numberLabels(second, cat2, cat2Side, cat2Font, cat2Color);
    else if (!lying) {
      const wanted = Math.min(widestCategory, 120) + 8;
      flatLabels(cat2Side, cat2Font, cat2Color, Math.max(Math.ceil(wanted / Math.max(slot, 1e-6)), 1));
    }
  }

  // Titles of the axes, centred on the plot area.
  for (const title of titles) {
    const size = title.font.size;
    if (title.side === 'b' || title.side === 't') {
      const text = fit(title.text, title.font, plot.width, measure);
      const y = title.side === 'b' ? plotBottom + need.b + size * 1.1 : plot.y - need.t - size * 0.4;
      if (text !== '') items.push({ kind: 'text', role: 'axisTitle', text, x: plot.x + plot.width / 2, y, font: title.font, color: title.color, align: 'center' });
    } else {
      const text = fit(title.text, title.font, plot.height, measure);
      // Turned text lies to the left of the line it is written on.
      const x = title.side === 'l' ? plot.x - need.l - size * 0.4 : plotRight + need.r + size * 1.4;
      if (text !== '') items.push({ kind: 'text', role: 'axisTitle', text, x, y: plot.y + plot.height / 2, font: title.font, color: title.color, align: 'center', rotated: true });
    }
  }
  if (placer) for (const item of placer.items) items.push(item);
}
