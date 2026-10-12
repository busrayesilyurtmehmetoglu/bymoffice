// Works out what a chart shows: the numbers of its series (from the cells the
// series refer to, or from the values the file cached), the scale of its axes,
// and where everything is drawn inside the frame of the chart. Pure
// calculation: text is measured through a function that is handed in, and the
// result is a flat list of items for a painter (see render/chart.ts).
//
// Excel is not at hand to compare with: rules that follow what Excel is known
// to do but could not be checked against it are marked UNSURE.

import { formatValue } from '../format/number-format';
import { builtinFormatId, cellDisplayText, isoToSerial } from '../format/cell-format';
import type { FormatLocale } from '../format/locale';
import {
  chartCached,
  chartRefAreas,
  MAX_CHART_POINTS,
  type Chart,
  type ChartAxis,
  type ChartData,
  type ChartGroup,
  type ChartKind,
  type ChartLine,
  type ChartMarker,
  type ChartPoint,
  type ChartRefArea,
  type ChartSeries,
  type ChartTextStyle,
} from '../model/sheet/chart';
import { resolveCellStyle } from '../model/sheet/color';
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
  /** Scatter charts: the horizontal place of each point. */
  xs?: ChartNumbers;
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

function numbersOf(chart: Chart, data: ChartData | undefined, options: ChartDataOptions, budget: Budget, zero = chart.blanks === 'zero'): Numbers {
  if (!data) return { values: NO_NUMBERS, live: false };
  const workbook = options.workbook;
  const areas = workbook && data.ref !== undefined ? chartRefAreas(workbook, data.ref) : undefined;
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
  const code = categoryFormat(axis);
  let areas = workbook && data.ref !== undefined ? chartRefAreas(workbook, data.ref) : undefined;
  if (areas && data.multiLevel) {
    // UNSURE: of a block of several levels the last column (or, for a block that lies, the last row) is the level next to the axis.
    if (areas.length !== 1) areas = undefined;
    else {
      const range = areas[0].range;
      const tall = range.bottom - range.top >= range.right - range.left;
      areas = [{ sheet: areas[0].sheet, range: tall ? { ...range, left: range.right } : { ...range, top: range.bottom } }];
    }
  }
  if (areas && workbook) {
    // Only the cells that hold something are kept, by their place among the categories.
    const cells = new Map<number, Cell>();
    const count = eachCell(areas, chart.plotVisOnly, Math.max(Math.min(MAX_CHART_POINTS, budget.left), 0), (cell, position) => void cells.set(position, cell));
    budget.left -= count;
    budget.held += cells.size;
    const date1904 = workbook.date1904 === true;
    return {
      count,
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
  return {
    count,
    label: (index) => {
      const value = chartCached(data, index);
      if (typeof value === 'string') return value;
      if (typeof value === 'number') return formatted(value, code !== undefined ? code : data.formatCode || 'General', options.locale, date1904) || '';
      return '';
    },
  };
}

/** The text of the first cell a reference names; undefined when it cannot be followed. */
function cellText(reference: string | undefined, options: ChartDataOptions): string | undefined {
  const workbook = options.workbook;
  if (!workbook || reference === undefined) return undefined;
  const areas = chartRefAreas(workbook, reference);
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
      let name = item.name ? cellText(item.name.ref, options) : undefined;
      if (name === undefined) name = item.name && item.name.text !== undefined ? item.name.text : options.seriesName ? options.seriesName(position) : `Series${position}`;
      const entry: ChartSeriesValues = { group: g, kind: group.kind, series: item, name, values: numbers.values, format: numbers.format, live: numbers.live };
      if (group.kind === 'scatter') {
        // A point without a number for x is not made zero, whatever the chart makes of empty cells.
        const xs = item.categories && item.categories.kind === 'num' ? numbersOf(chart, item.categories, options, budget, false).values : undefined;
        entry.xs = placesOf(xs, numbers.values);
      } else if (!labels && item.categories) labels = labelsOf(chart, item.categories, options, budget, categoryAxis(chart, group));
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
  const title = chart.title ? cellText(chart.title.ref, options) : undefined;
  if (title !== undefined) out.title = title;
  return out;
}

// ----------------------------------------------------------------------- scale

export interface ChartScale {
  min: number;
  max: number;
  step: number;
  /** The values that get a tick and a label, from `min` upwards. */
  ticks: number[];
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
  | { kind: 'rect'; role: 'frame' | 'plot' | 'bar' | 'swatch'; x: number; y: number; width: number; height: number; fill?: string; stroke?: string; lineWidth?: number; radius?: number; series?: number; point?: number }
  | { kind: 'path'; role: 'grid' | 'axis' | 'line' | 'area' | 'swatch'; points: number[]; closed?: boolean; fill?: string; stroke?: string; lineWidth?: number; dash?: boolean; series?: number }
  | { kind: 'markers'; role: 'marker' | 'swatch'; symbol: string; size: number; points: number[]; fill?: string; stroke?: string; series?: number }
  | { kind: 'slice'; cx: number; cy: number; radius: number; inner: number; start: number; end: number; fill: string; stroke?: string; lineWidth?: number; series: number; point: number }
  | { kind: 'text'; role: 'title' | 'tick' | 'category' | 'axisTitle' | 'legend'; text: string; x: number; y: number; font: ChartFont; color: string; align: 'left' | 'center' | 'right'; rotated?: boolean };

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
  const filled = group.kind === 'bar' || group.kind === 'area' || group.kind === 'pie' || group.kind === 'doughnut';
  const line = shape ? shape.line : undefined;
  let stroke: string | undefined;
  let marker: SeriesLook['marker'];
  if (group.kind === 'line') {
    stroke = lineColor(line, color);
    marker = markerOf(series.marker, group.markers === true, stroke || color);
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
}

/**
 * Where everything of a chart is drawn in a frame of `width` by `height`
 * pixels, from (0, 0). A frame too small for words shows the series with
 * their axes, and a still smaller one the series alone. Never throws for
 * what a chart states: sizes that are no sizes give an empty layout.
 */
export function layoutChart(chart: Chart, values: ChartValues, width: number, height: number, measure: ChartMeasure, options: ChartLayoutOptions = {}): ChartLayout {
  const w = Number.isFinite(width) && width > 0 ? Math.min(width, 1e5) : 0;
  const h = Number.isFinite(height) && height > 0 ? Math.min(height, 1e5) : 0;
  const detail: ChartLayout['detail'] = w < 80 || h < 60 ? 'bare' : w < 180 || h < 120 ? 'compact' : 'full';
  const items: ChartItem[] = [];
  const layout: ChartLayout = { width: w, height: h, detail, radius: chart.roundedCorners ? Math.min(8, w / 4, h / 4) : 0, plot: { x: 0, y: 0, width: w, height: h }, items };
  if (w === 0 || h === 0) return layout;
  const context: Context = { chart, values, measure, options, items, detail };

  // The frame: white with a thin grey outline unless the file says otherwise.
  const area = chart.area;
  const frameFill = area && area.fill !== undefined ? (area.fill === null ? undefined : safe(area.fill)) : '#FFFFFF';
  const frameStroke = lineColor(area ? area.line : undefined, GRID_COLOR);
  items.push({ kind: 'rect', role: 'frame', x: 0, y: 0, width: w, height: h, fill: frameFill, stroke: frameStroke, lineWidth: clamp((area && area.line && area.line.width) || 1, 0.5, 8), radius: layout.radius });

  const pad = detail === 'full' ? 8 : detail === 'compact' ? 4 : Math.min(2, w / 8, h / 8);
  const box: Box = { x: pad, y: pad, width: w - pad * 2, height: h - pad * 2 };

  const pies = chart.groups.some((group) => group.kind === 'pie' || group.kind === 'doughnut');
  const shown = values.series.filter((entry) => (pies ? entry.kind === 'pie' || entry.kind === 'doughnut' : true));

  if (detail === 'full') {
    placeTitle(context, box, shown);
    placeLegend(context, box, shown, pies);
  }
  if (pies) layoutPie(context, layout, box, shown);
  else layoutAxes(context, layout, box, shown);
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
  for (let i = 0; i < lines.length; i++) {
    const line = fit(clean(lines[i]), font, box.width, measure);
    if (line !== '') items.push({ kind: 'text', role: 'title', text: line, x: box.x + box.width / 2, y: box.y + i * lineHeight + font.size, font, color, align: 'center' });
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
      const lined = entry.kind === 'line' || entry.kind === 'scatter';
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
  const draw = (entry: LegendEntry, label: string, x: number, top: number): void => {
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

// ------------------------------------------------------------------------- pie

function layoutPie(context: Context, layout: ChartLayout, box: Box, shown: ChartSeriesValues[]): void {
  const { chart, items } = context;
  const diameter = Math.max(Math.min(box.width, box.height) - (context.detail === 'full' ? 8 : 0), 1);
  const radius = diameter / 2;
  const cx = box.x + box.width / 2;
  const cy = box.y + box.height / 2;
  layout.plot = { x: cx - radius, y: cy - radius, width: diameter, height: diameter };
  // A pie shows its first series; a doughnut one ring per series.
  const first = shown[0];
  if (!first) return;
  const group = chart.groups[first.group];
  const rings = first.kind === 'doughnut' ? shown.filter((entry) => entry.group === first.group).slice(0, 32) : [first];
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
    for (let k = 0; k < known; k++) {
      const i = entry.values.knownIndex(k);
      const size = Math.abs(entry.values.knownValue(k));
      if (!(size > 0)) continue;
      const end = angle + (size / total) * turn;
      if (i >= MAX_PIE_SLICES) {
        // What is left of a pie of very many slices is one slice.
        items.push({ kind: 'slice', cx, cy, radius: outer, inner, start: angle, end: begin + turn, fill: REST_COLOR, series: index, point: i });
        break;
      }
      const point = pointLook(chart, look, i, group.varyColors !== false);
      items.push({ kind: 'slice', cx, cy, radius: outer, inner, start: angle, end, fill: point.fill || '#FFFFFF', stroke: point.stroke, lineWidth: point.lineWidth, series: index, point: i });
      angle = end;
    }
  }
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

  /** `spanning`: a point without a number does not break the line. */
  constructor(private readonly spanning: boolean) {}

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
    if (look.stroke) for (const points of this.runs) items.push({ kind: 'path', role: 'line', points, stroke: look.stroke, lineWidth: look.lineWidth, dash: look.dash, series });
    const marker = look.marker;
    if (marker && this.marks.length > 0 && markers) items.push({ kind: 'markers', role: 'marker', symbol: marker.symbol, size: marker.size, points: this.marks, fill: marker.fill, stroke: marker.stroke, series });
    else if (!marker && dots && this.lonely.length > 0 && look.stroke) {
      items.push({ kind: 'markers', role: 'marker', symbol: 'circle', size: Math.max(look.lineWidth, 2), points: this.lonely, fill: look.stroke, series });
    }
  }
}

interface Plotted {
  entry: ChartSeriesValues;
  index: number;
  group: ChartGroup;
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
  look: SeriesLook;
}

function layoutAxes(context: Context, layout: ChartLayout, box: Box, shown: ChartSeriesValues[]): void {
  const { chart, values, measure, options, items, detail } = context;
  const words = detail === 'full';
  const scatter = shown.length > 0 && shown.every((entry) => entry.kind === 'scatter');
  // Series of a scatter type beside bars, lines or areas are drawn on the axes of those (see below).
  const categorical = scatter ? [] : shown.filter((entry) => entry.kind !== 'scatter');
  const loose = shown.filter((entry) => entry.kind === 'scatter');
  const firstGroup = categorical.length > 0 ? chart.groups[categorical[0].group] : shown.length > 0 ? chart.groups[shown[0].group] : chart.groups[0];
  const barGroup = chart.groups.filter((group) => group.kind === 'bar')[0];
  // Bars that lie turn the whole chart: categories run up the side, values along the bottom.
  const lying = !scatter && barGroup !== undefined && barGroup.barDir === 'bar';
  const both = scatter ? scatterAxes(chart, firstGroup) : undefined;
  const catAxis = both ? both.x : axisOf(chart, firstGroup, 'cat');
  const valAxis = both ? both.y : axisOf(chart, firstGroup, 'val');
  const labelFont = fontOf(valAxis ? valAxis.labelStyle : undefined, chart.textStyle, 9);
  const labelColor = colorOf(valAxis ? valAxis.labelStyle : undefined, chart.textStyle);
  const catFont = fontOf(catAxis ? catAxis.labelStyle : undefined, chart.textStyle, 9);
  const catColor = colorOf(catAxis ? catAxis.labelStyle : undefined, chart.textStyle);
  const valueLabels = words && !(valAxis && (valAxis.deleted || valAxis.tickLabels === 'none'));
  const categoryLabels = words && !(catAxis && (catAxis.deleted || catAxis.tickLabels === 'none'));

  // Titles of the axes take a band below and a band at the left.
  const bottomAxis = lying ? valAxis : catAxis;
  const leftAxis = lying ? catAxis : valAxis;
  const axisTitle = (axis: ChartAxis | undefined): { text: string; font: ChartFont; color: string } | undefined => {
    if (!words || !axis || axis.deleted || !axis.title || axis.title.text === undefined) return undefined;
    const text = clean(axis.title.text.replace(/\n/g, ' '));
    return text === '' ? undefined : { text, font: fontOf(axis.title.style, chart.textStyle, 10, true), color: colorOf(axis.title.style, chart.textStyle) };
  };
  const bottomTitle = axisTitle(bottomAxis);
  const leftTitle = axisTitle(leftAxis);
  let bottomTitleY: number | undefined;
  if (bottomTitle && box.height > 120) {
    box.height -= bottomTitle.font.size * 1.4;
    bottomTitleY = box.y + box.height + bottomTitle.font.size * 1.1;
  }
  let leftTitleX: number | undefined;
  if (leftTitle && box.width > 160) {
    const taken = leftTitle.font.size * 1.4;
    leftTitleX = box.x + leftTitle.font.size;
    box.x += taken;
    box.width -= taken;
  }

  // The series as they are drawn: sampled down where there are more points than places to draw them. The
  // categories are those of the series that have any: a scatter series beside them brings none of its own.
  let count = 0;
  for (const entry of categorical) count = Math.max(count, entry.values.length);
  count = scatter ? 0 : Math.max(count, 1);
  const alongGuess = Math.max(lying ? box.height : box.width, 1);
  // No more positions than pixels, and no more of everything together than a frame is worth painting.
  const drawn = Math.max(Math.min(Math.max(Math.floor(alongGuess), 16), MAX_DRAWN_POINTS, Math.floor(MAX_DRAWN_ITEMS / Math.max(shown.length, 1))), 16);
  const places = scatter ? 0 : count > drawn * 2 ? drawn : count;
  /** Points drawn of one scatter series. */
  const drawnScatter = Math.max(Math.min(Math.max(Math.floor(alongGuess), 16) * 2, MAX_DRAWN_POINTS * 2, Math.floor(MAX_DRAWN_ITEMS / Math.max(shown.length, 1))), 16);
  const plotted: Plotted[] = [];
  for (const entry of categorical) {
    const group = chart.groups[entry.group];
    const item: Plotted = { entry, index: values.series.indexOf(entry), group, look: lookOf(chart, group, entry.series), values: [] };
    if (places < count) {
      const sampled = sample(entry.values, count, places);
      item.values = sampled.peak;
      if (group.kind === 'line' && group.grouping === 'standard') {
        item.low = sampled.low;
        item.high = sampled.high;
      }
    } else item.values = spread(entry.values, places);
    plotted.push(item);
  }
  const scattered: Scattered[] = loose.map((entry) => ({ entry, index: values.series.indexOf(entry), group: chart.groups[entry.group], look: lookOf(chart, chart.groups[entry.group], entry.series) }));

  // Points stand in the middle of their categories, or on the tick marks when the value axis says so and no bars need the room.
  const onTicks = !scatter && valAxis !== undefined && valAxis.crossBetween === 'midCat' && !plotted.some((item) => item.group.kind === 'bar') && places > 1;

  // What the value axis has to hold: the values themselves, or the sums of what is stacked.
  let low = Infinity;
  let high = -Infinity;
  const see = (value: number | null | undefined): void => {
    if (value === null || value === undefined || !Number.isFinite(value)) return;
    if (value < low) low = value;
    if (value > high) high = value;
  };
  for (const item of scattered) {
    item.entry.values.each(see);
    if (item.entry.values.blank !== null && item.entry.values.known < item.entry.values.length) see(item.entry.values.blank);
  }
  const piles = new Map<number, { up: number[]; down: number[]; total: number[] }>();
  let percent = false;
  for (const item of plotted) {
    if (item.group.grouping === 'standard') {
      for (let i = 0; i < item.values.length; i++) see(item.values[i]);
      if (item.low && item.high) for (let i = 0; i < item.low.length; i++) (see(item.low[i]), see(item.high[i]));
      continue;
    }
    let pile = piles.get(item.entry.group);
    if (!pile) {
      pile = { up: [], down: [], total: [] };
      piles.set(item.entry.group, pile);
    }
    if (item.group.grouping === 'percentStacked') percent = true;
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
      see(base + value);
    }
    item.to = to;
    item.from = from;
    // What is stacked stands on zero.
    see(0);
  }
  if (low === Infinity) {
    low = 0;
    high = 0;
  }
  // Bars and areas stand on zero: it is part of what the axis shows.
  if (!scatter && plotted.some((item) => item.group.kind === 'bar' || item.group.kind === 'area')) {
    if (low > 0) low = 0;
    if (high < 0) high = 0;
  }

  const firstFormat = plotted.length > 0 ? plotted[0].entry.format : scattered.length > 0 ? scattered[0].entry.format : undefined;
  const format = percent && !(valAxis && valAxis.formatCode && !valAxis.sourceLinked && valAxis.formatCode !== 'General')
    ? { code: '0%' }
    : valAxis && valAxis.formatCode && !valAxis.sourceLinked
      ? { code: valAxis.formatCode }
      : firstFormat || (valAxis && valAxis.formatCode ? { code: valAxis.formatCode } : undefined);

  // Room for the labels: those of the values are measured once the scale is known, which needs the length of its axis.
  const tickGap = 6;
  const catBand = categoryLabels ? catFont.size * 1.3 + tickGap : 0;
  const valBand = valueLabels ? labelFont.size * 1.3 + tickGap : 0;
  let plot: Box;
  let scale: ChartScale;
  let xScale: ChartScale | undefined;
  const scaleOptions = (axis: ChartAxis | undefined, length: number, each: number): ChartScaleOptions => ({
    min: axis ? axis.min : undefined,
    max: axis ? axis.max : undefined,
    step: axis ? axis.majorUnit : undefined,
    maxSteps: clamp(Math.floor(length / each), 1, 10),
  });
  const valueScale = (length: number, each: number): ChartScale =>
    percent && !(valAxis && (valAxis.min !== undefined || valAxis.max !== undefined))
      ? chartScale(low, high, { ...scaleOptions(valAxis, length, each), min: low < 0 ? -1 : 0, max: 1 })
      : chartScale(low, high, scaleOptions(valAxis, length, each));
  let tickLabels: string[] = [];
  let catWidth = 0;
  if (lying) {
    // Category labels at the left: as wide as the widest of them, up to a third of the frame.
    if (categoryLabels) {
      const stride = Math.max(Math.floor(count / 50), 1);
      for (let i = 0; i < count; i += stride) catWidth = Math.max(catWidth, measure(clean(values.label(i)), catFont));
      catWidth = Math.min(catWidth, box.width * 0.33) + tickGap;
    }
    // Room at the end of the axis for half of its last label.
    const edge = valueLabels ? 16 : words ? 8 : 0;
    plot = { x: box.x + catWidth, y: box.y, width: Math.max(box.width - catWidth - edge, Math.min(box.width, 1)), height: Math.max(box.height - valBand, Math.min(box.height, 1)) };
    scale = valueScale(plot.width, Math.max(labelFont.size * 3.2, 36));
    tickLabels = scale.ticks.map((tick) => tickText(tick, format, options));
  } else {
    // Half a line above the plot area for the label of the topmost tick.
    const headroom = valueLabels ? labelFont.size * 0.6 : 0;
    const plotHeight = Math.max(box.height - catBand - headroom, Math.min(box.height, 1));
    scale = valueScale(plotHeight, Math.max(labelFont.size * 2.2, 8));
    tickLabels = scale.ticks.map((tick) => tickText(tick, format, options));
    let widest = 0;
    if (valueLabels) for (const label of tickLabels) widest = Math.max(widest, measure(label, labelFont));
    const band = valueLabels ? Math.min(widest, box.width * 0.4) + tickGap : 0;
    // Room at the right for half of the last label, where the last point stands on the edge of the plot area.
    let edge = words ? 6 : 0;
    if (categoryLabels && scatter) edge = 16;
    else if (categoryLabels && onTicks) edge = Math.max(edge, Math.min(measure(clean(values.label(count - 1)), catFont) / 2, 40) + 2);
    plot = { x: box.x + band, y: box.y + headroom, width: Math.max(box.width - band - edge, Math.min(box.width, 1)), height: plotHeight };
  }
  if (scatter) {
    let xLow = Infinity;
    let xHigh = -Infinity;
    const seeX = (x: number | null): void => {
      if (x === null || !Number.isFinite(x)) return;
      if (x < xLow) xLow = x;
      if (x > xHigh) xHigh = x;
    };
    for (const item of scattered) {
      const xs = item.entry.xs;
      const ys = item.entry.values;
      if (!xs) continue;
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
    if (xLow === Infinity) {
      xLow = 0;
      xHigh = 0;
    }
    xScale = chartScale(xLow, xHigh, scaleOptions(catAxis, plot.width, Math.max(catFont.size * 3.2, 36)));
  }
  layout.plot = plot;
  layout.scale = scale;
  layout.xScale = xScale;

  // The area of the series, when the file gives it a fill or an outline.
  const plotShape = chart.plotArea;
  if (plotShape) {
    const fill = plotShape.fill === null ? undefined : safe(plotShape.fill);
    const stroke = lineColor(plotShape.line, undefined);
    if (fill || stroke) items.push({ kind: 'rect', role: 'plot', x: plot.x, y: plot.y, width: plot.width, height: plot.height, fill, stroke, lineWidth: clamp((plotShape.line && plotShape.line.width) || 1, 0.5, 8) });
  }

  // Values to pixels along the value axis; categories to pixels along theirs.
  const valueReversed = valAxis !== undefined && valAxis.reversed === true;
  const span = scale.max - scale.min;
  const share = (value: number): number => {
    const part = (value - scale.min) / span;
    return valueReversed ? 1 - part : part;
  };
  const valueAt = (value: number): number => (lying ? plot.x + share(value) * plot.width : plot.y + plot.height - share(value) * plot.height);
  const along = lying ? plot.height : plot.width;
  const catReversed = catAxis !== undefined && catAxis.reversed === true;
  const slot = scatter ? 0 : onTicks ? along / (places - 1) : along / Math.max(places, 1);
  const xSpan = xScale ? xScale.max - xScale.min : 1;
  /** A distance from the start of the category axis as a pixel. */
  const alongAt = (offset: number): number => {
    // Categories of bars that lie run from the bottom up, as the values of an upright chart do.
    const flipped = lying ? !catReversed : catReversed;
    return lying ? plot.y + (flipped ? along - offset : offset) : plot.x + (flipped ? along - offset : offset);
  };
  const catAt = (position: number): number => {
    if (xScale) return alongAt(((position - xScale.min) / xSpan) * along);
    return alongAt(onTicks ? position * slot : (position + 0.5) * slot);
  };
  const at = (cat: number, value: number): [number, number] => (lying ? [valueAt(value), catAt(cat)] : [catAt(cat), valueAt(value)]);
  /**
   * Where a scatter point with the horizontal value `x` stands; undefined for one outside the axis.
   * UNSURE: beside bars, lines or areas there is no axis of numbers to place it on. The n-th category
   * stands at n on the category axis, so a point is put where its x lies among 1, 2, 3, ...: 2.5 is
   * halfway between the second and the third category. (Excel is known to draw such a combination on
   * the axes of the other chart type; how it counts along a category axis was not checked against it.)
   */
  const scatterAt = (x: number): number | undefined => {
    if (xScale) return x < xScale.min || x > xScale.max ? undefined : catAt(x);
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
  if (detail !== 'bare' && valAxis && !valAxis.deleted) {
    if (valAxis.minorGridlines) {
      const minor = gridLine(valAxis.minorGridlines, MINOR_GRID_COLOR);
      const fine = scale.step / 5;
      for (let i = 0; i <= MAX_AXIS_TICKS * 5; i++) {
        const value = scale.min + i * fine;
        if (value > scale.max + fine * 1e-6) break;
        if (i % 5 !== 0) across(valueAt(value), minor, lying);
      }
    }
    if (valAxis.majorGridlines) {
      const major = gridLine(valAxis.majorGridlines, GRID_COLOR);
      for (const tick of scale.ticks) across(valueAt(tick), major, lying);
    }
  }
  if (detail !== 'bare' && catAxis && !catAxis.deleted && catAxis.majorGridlines) {
    const major = gridLine(catAxis.majorGridlines, GRID_COLOR);
    if (xScale) for (const tick of xScale.ticks) across(catAt(tick), major, !lying);
    else {
      const lines = Math.min(places, 500);
      for (let i = 0; i <= lines; i++) across((lying ? plot.y : plot.x) + (i * along) / lines, major, !lying);
    }
  }

  // Where the series stand: on zero, or on the end of the axis when zero is not on it.
  // The category axis says where it crosses the value axis.
  const crosses = catAxis && !scatter ? catAxis.crosses : undefined;
  const base =
    catAxis && !scatter && catAxis.crossesAt !== undefined
      ? clamp(catAxis.crossesAt, scale.min, scale.max)
      : crosses === 'max'
        ? scale.max
        : crosses === 'min'
          ? scale.min
          : clamp(0, scale.min, scale.max);

  const bars = plotted.filter((item) => item.group.kind === 'bar');
  for (const kind of ['area', 'bar', 'line'] as ChartKind[]) {
    for (const item of plotted) {
      if (item.group.kind !== kind) continue;
      const group = item.group;
      const stacked = group.grouping !== 'standard';
      /** The value a point is drawn at and the value it stands on; undefined for a point that is not drawn. */
      const ends = (i: number): [number, number] | undefined => {
        if (stacked && item.to && item.from) {
          const top = item.to[i];
          return top === null || top === undefined ? undefined : [top, item.from[i] || 0];
        }
        const raw = item.values[i];
        return raw === null || raw === undefined ? undefined : [raw, base];
      };

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
          const span2 = ends(i);
          if (!span2) continue;
          const a = valueAt(clamp(span2[0], scale.min, scale.max));
          const b = valueAt(clamp(span2[1], scale.min, scale.max));
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
        }
        continue;
      }

      if (kind === 'area') {
        // The outline along the values and back along what the series stands on.
        const top: number[] = [];
        const bottom: number[] = [];
        for (let i = 0; i < places; i++) {
          // UNSURE: a point without a number is drawn as zero in an area.
          const span2 = ends(i) || [base, base];
          const p = at(i, clamp(span2[0], scale.min, scale.max));
          const q = at(i, clamp(span2[1], scale.min, scale.max));
          top.push(p[0], p[1]);
          bottom.push(q[0], q[1]);
        }
        const points = top.slice();
        for (let i = bottom.length - 2; i >= 0; i -= 2) points.push(bottom[i], bottom[i + 1]);
        if (points.length >= 6) items.push({ kind: 'path', role: 'area', points, closed: true, fill: item.look.fill, stroke: item.look.stroke, lineWidth: item.look.lineWidth, series: item.index });
        continue;
      }

      // Lines: runs of points, broken where a point has no number unless the chart spans such points.
      const trace = new Trace(chart.blanks === 'span');
      for (let i = 0; i < places; i++) {
        if (item.low && item.high) {
          const lo = item.low[i];
          const hi = item.high[i];
          if (lo === null || hi === null) {
            trace.gap();
            continue;
          }
          const p = at(i, clamp(lo, scale.min, scale.max));
          const q = at(i, clamp(hi, scale.min, scale.max));
          trace.push(p[0], p[1], false);
          if (hi !== lo) trace.push(q[0], q[1], false);
          continue;
        }
        const span2 = ends(i);
        if (!span2) {
          trace.gap();
          continue;
        }
        const p = at(i, clamp(span2[0], scale.min, scale.max));
        trace.push(p[0], p[1], true);
      }
      trace.emit(items, item.look, item.index, !(item.low && item.high), true);
    }
  }

  // Scatter series, above the others: every point at its own x. Of very many points every n-th one is drawn,
  // so that no series asks for more markers than its share of what a frame is worth painting.
  for (const item of scattered) {
    const xs = item.entry.xs;
    const ys = item.entry.values;
    if (!xs) continue;
    const trace = new Trace(chart.blanks === 'span');
    const plain = ys.blank === null;
    // Points without a number are skipped where they are nothing; where they are zero every point is one.
    const total = plain ? ys.known : ys.length;
    const stride = Math.max(Math.ceil(total / drawnScatter), 1);
    let before = -1;
    for (let k = 0; k < total; k += stride) {
      const i = plain ? ys.knownIndex(k) : k;
      const y = plain ? ys.knownValue(k) : ys.at(i);
      const x = xs.at(i);
      // A point that was left out between two that are drawn, because it has no number, breaks the line.
      if (stride === 1 && before >= 0 && i !== before + 1) trace.gap();
      before = i;
      const px = x === null || y === null ? undefined : scatterAt(x);
      if (px === undefined || y === null) {
        trace.gap();
        continue;
      }
      const py = valueAt(clamp(y, scale.min, scale.max));
      if (lying) trace.push(py, px, true);
      else trace.push(px, py, true);
    }
    trace.emit(items, item.look, item.index, true, false);
  }

  if (detail === 'bare') return;

  // The lines of the axes, above the series.
  const axisLine = (axis: ChartAxis | undefined, points: number[]): void => {
    if (axis && axis.deleted) return;
    const look = gridLine(axis ? axis.line : undefined, GRID_COLOR);
    if (!look.stroke) return;
    const snapped = points.map((value, index) => (look.width <= 1 && points[index % 2] === points[(index % 2) + 2] ? Math.round(value - 0.5) + 0.5 : value));
    items.push({ kind: 'path', role: 'axis', points: snapped, stroke: look.stroke, lineWidth: look.width, dash: look.dash });
  };
  const basePixel = valueAt(base);
  if (lying) {
    axisLine(catAxis, [basePixel, plot.y, basePixel, plot.y + plot.height]);
    axisLine(valAxis, [plot.x, plot.y + plot.height, plot.x + plot.width, plot.y + plot.height]);
  } else {
    axisLine(catAxis, [plot.x, basePixel, plot.x + plot.width, basePixel]);
    axisLine(valAxis, [plot.x, plot.y, plot.x, plot.y + plot.height]);
  }
  if (!words) return;

  // Labels of the value axis: all of them when there is room, else every n-th.
  if (valueLabels) {
    if (lying) {
      let widest = 0;
      for (const label of tickLabels) widest = Math.max(widest, measure(label, labelFont));
      const every = Math.max(Math.ceil(((widest + 8) * tickLabels.length) / Math.max(plot.width, 1)), 1);
      for (let i = 0; i < tickLabels.length; i += every) {
        if (tickLabels[i] !== '') items.push({ kind: 'text', role: 'tick', text: tickLabels[i], x: valueAt(scale.ticks[i]), y: plot.y + plot.height + tickGap + labelFont.size * 0.85, font: labelFont, color: labelColor, align: 'center' });
      }
    } else {
      const every = Math.max(Math.ceil((labelFont.size * 1.1 * tickLabels.length) / Math.max(plot.height, 1)), 1);
      const room = plot.x - box.x - tickGap;
      for (let i = 0; i < tickLabels.length; i += every) {
        const text = fit(tickLabels[i], labelFont, room, measure);
        if (text !== '') items.push({ kind: 'text', role: 'tick', text, x: plot.x - tickGap, y: valueAt(scale.ticks[i]) + labelFont.size * 0.35, font: labelFont, color: labelColor, align: 'right' });
      }
    }
  }

  // Labels of the categories. They are never turned: where they would run into each other only every n-th is shown,
  // cut to the room it has.
  if (categoryLabels) {
    if (xScale) {
      const labels = xScale.ticks.map((tick) => tickText(tick, catAxis && catAxis.formatCode && !catAxis.sourceLinked ? { code: catAxis.formatCode } : undefined, options));
      let widest = 0;
      for (const label of labels) widest = Math.max(widest, measure(label, catFont));
      const every = Math.max(Math.ceil(((widest + 8) * labels.length) / Math.max(plot.width, 1)), 1);
      for (let i = 0; i < labels.length; i += every) {
        if (labels[i] !== '') items.push({ kind: 'text', role: 'category', text: labels[i], x: catAt(xScale.ticks[i]), y: plot.y + plot.height + tickGap + catFont.size * 0.85, font: catFont, color: catColor, align: 'center' });
      }
    } else if (lying) {
      const every = Math.max(Math.ceil((catFont.size * 1.2) / Math.max(slot, 1e-6)), 1);
      for (let i = 0; i < places && i / every < 2000; i += every) {
        const text = fit(clean(values.label(Math.floor((i * count) / places))), catFont, catWidth - tickGap, measure);
        if (text !== '') items.push({ kind: 'text', role: 'category', text, x: plot.x - tickGap, y: catAt(i) + catFont.size * 0.35, font: catFont, color: catColor, align: 'right' });
      }
    } else {
      // The widest of a sample decides how many are left out.
      let widest = 0;
      const stride = Math.max(Math.floor(places / 50), 1);
      for (let i = 0; i < places; i += stride) widest = Math.max(widest, measure(clean(values.label(Math.floor((i * count) / places))), catFont));
      const wanted = Math.min(widest, 120) + 8;
      const every = Math.max(Math.ceil(wanted / Math.max(slot, 1e-6)), 1);
      const room = Math.max(slot * every - 4, 0);
      for (let i = 0; i < places && i / every < 2000; i += every) {
        const text = fit(clean(values.label(Math.floor((i * count) / places))), catFont, room, measure);
        if (text !== '') items.push({ kind: 'text', role: 'category', text, x: catAt(i), y: plot.y + plot.height + tickGap + catFont.size * 0.85, font: catFont, color: catColor, align: 'center' });
      }
    }
  }

  // Titles of the axes, centred on the plot area.
  if (bottomTitle && bottomTitleY !== undefined) {
    const text = fit(bottomTitle.text, bottomTitle.font, plot.width, measure);
    if (text !== '') items.push({ kind: 'text', role: 'axisTitle', text, x: plot.x + plot.width / 2, y: bottomTitleY, font: bottomTitle.font, color: bottomTitle.color, align: 'center' });
  }
  if (leftTitle && leftTitleX !== undefined) {
    const text = fit(leftTitle.text, leftTitle.font, plot.height, measure);
    if (text !== '') items.push({ kind: 'text', role: 'axisTitle', text, x: leftTitleX, y: plot.y + plot.height / 2, font: leftTitle.font, color: leftTitle.color, align: 'center', rotated: true });
  }
}
