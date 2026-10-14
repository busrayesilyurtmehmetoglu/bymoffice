// Charts as the editor changes them: a chart added from a range of cells, moved
// or resized, changed (its type, title, legend, axes and series) and deleted.
// Like the other commands these functions only build operations; the caller
// applies them through a History so that each is one undo step.
//
// Sheet.drawings holds the drawings of the file as the file has them. The list
// is replaced, never altered ("setDrawings"), so undoing puts the very list
// back and a save after it leaves the parts of the file alone.
//
// - A chart that is added is in no file: its model is all there is of it, and
//   the writer makes a part from it. Its references are `live` (Chart.live):
//   they name sheets as they are called now and cells where they are now, and
//   every edit of rows, columns and sheets keeps them so (chartsMoved(),
//   chartsRenamed() below).
// - A chart of the file whose part is exactly what the engine writes for its
//   model (DrawingChart.known "all") is changed like one that was added.
// - Any other chart of the file whose references the model holds (known
//   "refs") can be changed in what CHART_PART_CHANGES lists: the writer
//   changes those elements of the part and leaves everything else of it as it
//   is. A change the part cannot take without losing what the model does not
//   hold (another chart type) is refused.
// - A chart of which neither is known is moved and deleted, not changed.
// - A chart that is copied (copyChart(), pasteChart(), duplicateChart(), cutChart()) is a new
//   chart of the model. A copy of a chart whose part holds more than its model is made from that
//   part when the workbook is written (DrawingChart.source), and is changed the way the chart it
//   was copied from is.
//
// A drawing that was added, moved or changed stands where it was put, which
// is a place of the sheet as it was then (SheetDrawing.since), like a note.

import { adjustFormula, formulaNamesSheetsUnread, formulaReferences, quoteSheetName, renameSheetsInFormula, type StructuralEdit } from '../calc/refs';
import { scatterAxes } from '../layout/chart';
import { createSheetMetrics, type SheetMetrics } from '../layout/sheet-metrics';
import { chartKindGrouped, chartKindWritten, chartPartCost, chartPartProblem, chartRefAreas, chartRefs, chartWritable, mapChartRefs, MAX_CHART_PART_BYTES, MAX_CHART_POINTS, MAX_CHART_SERIES, MAX_SHEET_CHARTS, MAX_WORKBOOK_CHART_BYTES, MAX_WORKBOOK_CHART_POINTS, MAX_WORKBOOK_CHARTS, remapLegendEntries, type Chart, type ChartAxis, type ChartData, type ChartDataLabels, type ChartGroup, type ChartLegend, type ChartLine, type ChartSeries, type ChartShape } from '../model/sheet/chart';
import {
  MAX_COLS,
  MAX_ROWS,
  sheetNameMap,
  type Cell,
  type CellRange,
  type ChartOrigin,
  type DrawingAnchor,
  type DrawingChart,
  type DrawingMarker,
  type Sheet,
  foldSheetName,
  type SheetDrawing,
  type SheetNameMap,
  type SheetStructuralEdit,
  type Workbook,
} from '../model/sheet/types';
import { anchorRect, drawingAnchorNow } from '../render/sheet-drawings';
import { withCommand, type Op } from './ops';
import type { RuleLists } from './rules';

/** The chart types that are made here. "column": upright bars; "bar": bars that lie. */
export type ChartType = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter';
export const CHART_TYPES: ChartType[] = ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter'];
/** "standard": side by side (bars: clustered), or lines and areas that do not add up. */
export type ChartGrouping = 'standard' | 'stacked' | 'percentStacked';
export type ChartLegendPlace = ChartLegend['position'] | 'none';

/** Drawings one sheet holds: as many as are read from a file. (Charts: MAX_SHEET_CHARTS of them, and MAX_WORKBOOK_CHARTS in the workbook.) */
export const MAX_SHEET_DRAWINGS = 20000;
export { MAX_SHEET_CHARTS, MAX_WORKBOOK_CHARTS };
/** Longest title of a chart or an axis, and longest name of a series. */
export const MAX_CHART_TEXT = 255;
/** The size of a new chart in pixels: near the one Excel gives a new chart. */
export const DEFAULT_CHART_WIDTH = 480;
export const DEFAULT_CHART_HEIGHT = 288;
/** The smallest size a chart is given. */
export const MIN_CHART_SIZE = 24;
const EMU_PER_PIXEL = 9525;
const OFFICE_ACCENTS = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];
const HEX6 = /^[0-9A-Fa-f]{6}$/;
const COLOR = /^#[0-9A-Fa-f]{6}$/;

/**
 * What can be changed in a chart of the file that the engine did not write itself (known
 * "refs"), by the names of the fields of a ChartPatch. Everything but "type"; `grouping` only
 * within the type the chart has.
 */
export const CHART_PART_CHANGES = ['grouping', 'title', 'legend', 'axisTitles', 'axisBounds', 'dataLabels', 'series', 'range', 'seriesIn'];

/** One series of a chart as it is asked for and listed. References are texts such as "Sheet1!$B$2:$B$9". */
export interface ChartSeriesSpec {
  /** The name as a text. */
  name?: string;
  /** The cell the name is read from; takes the place of `name`. */
  nameRef?: string;
  /** The categories (the horizontal values of a scatter chart). */
  categories?: string;
  values?: string;
  /** "#RRGGBB": the fill of bars, areas and markers, the colour of a line. */
  color?: string;
}

/** A series of a chart that is changed: the series it was (`from`: its place among the series of the chart before), with what is said here in place of what it had. Without `from` a new series. */
export interface ChartSeriesEdit {
  from?: number;
  /** null takes the name away. */
  name?: string | null;
  nameRef?: string | null;
  /** null takes the categories away. */
  categories?: string | null;
  values?: string;
  /** null gives the series the colour the chart gives by itself. */
  color?: string | null;
}

export interface ChartSpec {
  type: ChartType;
  /** Bars, lines and areas. Default "standard". */
  grouping?: ChartGrouping;
  /**
   * The cells the chart shows. A first row without numbers names the series and a first column
   * without numbers names the categories (see seriesOfRange()).
   */
  range?: CellRange;
  /** The series are the columns of the range, or its rows. Default: the columns, unless the range has more columns of values than rows. */
  seriesIn?: 'columns' | 'rows';
  /** The series one by one, in place of `range`. */
  series?: ChartSeriesSpec[];
  title?: string;
  /** Default "r"; "none" for a chart without a legend. */
  legend?: ChartLegendPlace;
  /** Where the chart stands: an anchor, or the cells it covers. Default: beside the range, 480 by 288 pixels. */
  anchor?: DrawingAnchor | CellRange;
}

export interface ChartPatch {
  type?: ChartType;
  grouping?: ChartGrouping;
  /** null or '' takes the title away. */
  title?: string | null;
  legend?: ChartLegendPlace;
  /** The titles of the axis of the categories (the horizontal values of a scatter chart) and of the values; null takes one away. */
  axisTitles?: { category?: string | null; value?: string | null };
  /** The smallest and the largest value of the value axis; null lets the chart work it out. */
  axisBounds?: { min?: number | null; max?: number | null };
  /** The values are written beside the points, or not. */
  dataLabels?: boolean;
  /** The series, all of them, in their new order: see ChartSeriesEdit. */
  series?: ChartSeriesEdit[];
  /** The series are made anew from a range of cells, as for a new chart. */
  range?: CellRange;
  seriesIn?: 'columns' | 'rows';
}

/** A chart of a sheet as the commands see it. */
export interface ChartInfo {
  /** The place of the chart in Sheet.drawings: what the commands name it by. It changes when a drawing before it is deleted. */
  id: number;
  /** Undefined for a chart of a kind that is not made here (radar, bubble, stock, a combination, 3-D). */
  type?: ChartType;
  grouping: ChartGrouping;
  title?: string;
  legend: ChartLegendPlace;
  /** Where the chart is now. */
  anchor: DrawingAnchor;
  series: ChartSeriesSpec[];
  axisTitles: { category?: string; value?: string };
  axisBounds: { min?: number; max?: number };
  dataLabels: boolean;
  /** The block of cells of this sheet the series are made of, when they are what seriesOfRange() makes of one. */
  range?: CellRange;
  seriesIn?: 'columns' | 'rows';
  /** "all": every change; "parts": what CHART_PART_CHANGES lists; "none": the chart is moved and deleted only. */
  editable: 'all' | 'parts' | 'none';
  /** The chart is in the file the workbook was read from. */
  ofFile: boolean;
  /** Whether a series can be added, and whether series can be taken away or put in another order. */
  canAddSeries: boolean;
  canReorderSeries: boolean;
}

export { drawingAnchorNow };

// ------------------------------------------------------------------ helpers

function sheetOf(workbook: Workbook, sheetIndex: number): Sheet {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet) throw new RangeError(`No sheet at index ${sheetIndex}`);
  if (sheet.kind !== 'worksheet') throw new RangeError(`Cannot change the charts of a ${sheet.kind}`);
  return sheet;
}

/** Whether a chart has a part that holds more than its model may: a chart of the file, and a copy that is made from the part of one (DrawingChart.source). */
const partBound = (object: DrawingChart): boolean => object.part !== undefined || object.source !== undefined;

/** What every command says of a chart a reader marked as only shown (Chart.shownOnly): it is neither changed, moved nor deleted. */
const SHOWN_ONLY = 'Cannot change this chart: it was read from an xls file and is only shown';
/** The same, of a chart of a newer type (Chart.extended). */
const SHOWN_ONLY_EXTENDED = 'Cannot change this chart: a chart of this type is only shown';
const shownOnly = (chart: Chart | undefined): string => (chart && chart.extended === true ? SHOWN_ONLY_EXTENDED : SHOWN_ONLY);

/** The chart of a sheet a command is about to change. Throws RangeError when `id` names none, and for a chart that is only shown. */
function chartAt(sheet: Sheet, id: number): { drawing: SheetDrawing; object: DrawingChart } {
  const drawing = sheet.drawings && Math.floor(id) === id ? sheet.drawings[id] : undefined;
  if (!drawing || !drawing.object || drawing.object.type !== 'chart') throw new RangeError(`No chart at index ${id}`);
  if (drawing.object.chart && drawing.object.chart.shownOnly === true) throw new RangeError(shownOnly(drawing.object.chart));
  return { drawing, object: drawing.object };
}

const isRange = (value: unknown): value is CellRange => {
  const range = value as CellRange;
  return (
    !!range &&
    [range.top, range.left, range.bottom, range.right].every((n) => typeof n === 'number' && n >= 0 && Math.floor(n) === n) &&
    range.top <= range.bottom &&
    range.left <= range.right &&
    range.bottom < MAX_ROWS &&
    range.right < MAX_COLS
  );
};

const LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
function columnName(col: number): string {
  let name = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) name = LETTERS.charAt((n - 1) % 26) + name;
  return name;
}

/** "Sheet1!$A$1:$B$5": a range of a sheet as a chart names it. */
export function chartRangeRef(sheetName: string, range: CellRange): string {
  const first = `$${columnName(range.left)}$${range.top + 1}`;
  const last = `$${columnName(range.right)}$${range.bottom + 1}`;
  return `${quoteSheetName(sheetName)}!${first === last ? first : `${first}:${last}`}`;
}

// --------------------------------------------------------------- the series of a range

function cellAt(sheet: Sheet, row: number, col: number): Cell | undefined {
  const held = sheet.rows.get(row);
  return held ? held.cells[col] : undefined;
}

const isEmpty = (cell: Cell | undefined): boolean => !cell || cell.kind === 'blank' || cell.value === undefined || cell.value === null || cell.value === '';
const isNumber = (cell: Cell | undefined): boolean => !!cell && cell.kind === 'number' && typeof cell.value === 'number';

/** Whether the cells of a line of the range (a row, or a column) name things: at least one of them holds something, and none holds a number. */
function namesThings(sheet: Sheet, cells: [row: number, col: number][]): boolean {
  let some = false;
  for (const [row, col] of cells) {
    const cell = cellAt(sheet, row, col);
    if (isEmpty(cell)) continue;
    if (isNumber(cell)) return false;
    some = true;
  }
  return some;
}

/** What a range of cells holds for a chart: see seriesOfRange(). */
export interface RangeSeries {
  seriesIn: 'columns' | 'rows';
  /** The first row of the range names things (the series when they are columns, the categories when they are rows). */
  headerRow: boolean;
  /** The first column of the range names things. */
  headerColumn: boolean;
  series: ChartSeriesSpec[];
}

/**
 * The series a range of cells gives a chart.
 *
 * UNSURE (worked out from what Excel is seen to do, not checked against it): the first row is
 * a row of names when it has more than one row, at least one of its cells holds something and
 * none of them holds a number; the same for the first column (the corner cell left aside for
 * both). A cell with a date counts as a name, not as a number. The series are the columns
 * unless the range has more columns of values than rows of values. A scatter chart always
 * takes the first column (row) of two or more as its horizontal values.
 *
 * Throws RangeError for a range without a cell of values, and for one that gives more than
 * MAX_CHART_SERIES series.
 */
export function seriesOfRange(workbook: Workbook, sheetIndex: number, range: CellRange, seriesIn?: 'columns' | 'rows', type?: ChartType): RangeSeries {
  const sheet = sheetOf(workbook, sheetIndex);
  if (!isRange(range)) throw new RangeError('The range of a chart is a range of cells of the sheet');
  const height = range.bottom - range.top + 1;
  const width = range.right - range.left + 1;
  // Cells are looked at as far as the sheet holds any: a whole column is not walked to its end.
  const bottom = Math.min(range.bottom, Math.max(sheet.rowCount - 1, range.top + 1));
  const right = Math.min(range.right, Math.max(sheet.colCount - 1, range.left + 1));
  const rowCells: [number, number][] = [];
  for (let col = range.left + 1; col <= right && rowCells.length < 16384; col++) rowCells.push([range.top, col]);
  const colCells: [number, number][] = [];
  for (let row = range.top + 1; row <= bottom && colCells.length < 100000; row++) colCells.push([row, range.left]);
  let headerRow = height > 1 && (width === 1 ? namesThings(sheet, [[range.top, range.left]]) : namesThings(sheet, rowCells));
  let headerColumn = width > 1 && (height === 1 ? namesThings(sheet, [[range.top, range.left]]) : namesThings(sheet, colCells));
  const dataRows = height - (headerRow ? 1 : 0);
  const dataCols = width - (headerColumn ? 1 : 0);
  const along = seriesIn === 'rows' || seriesIn === 'columns' ? seriesIn : dataCols > dataRows ? 'rows' : 'columns';
  if (type === 'scatter') {
    // The horizontal values stand in the first column (row) whatever they are.
    if (along === 'columns' && width > 1) headerColumn = true;
    if (along === 'rows' && height > 1) headerRow = true;
  }
  const top = range.top + (headerRow ? 1 : 0);
  const left = range.left + (headerColumn ? 1 : 0);
  if (top > range.bottom || left > range.right) throw new RangeError('The range holds no values for a chart');
  const name = sheet.name;
  const series: ChartSeriesSpec[] = [];
  if (along === 'columns') {
    if (range.right - left + 1 > MAX_CHART_SERIES) throw new RangeError(`A chart has at most ${MAX_CHART_SERIES} series`);
    for (let col = left; col <= range.right; col++) {
      const item: ChartSeriesSpec = { values: chartRangeRef(name, { top, bottom: range.bottom, left: col, right: col }) };
      if (headerRow) item.nameRef = chartRangeRef(name, { top: range.top, bottom: range.top, left: col, right: col });
      if (headerColumn) item.categories = chartRangeRef(name, { top, bottom: range.bottom, left: range.left, right: range.left });
      series.push(item);
    }
  } else {
    if (range.bottom - top + 1 > MAX_CHART_SERIES) throw new RangeError(`A chart has at most ${MAX_CHART_SERIES} series`);
    for (let row = top; row <= range.bottom; row++) {
      const item: ChartSeriesSpec = { values: chartRangeRef(name, { top: row, bottom: row, left, right: range.right }) };
      if (headerColumn) item.nameRef = chartRangeRef(name, { top: row, bottom: row, left: range.left, right: range.left });
      if (headerRow) item.categories = chartRangeRef(name, { top: range.top, bottom: range.top, left, right: range.right });
      series.push(item);
    }
  }
  return { seriesIn: along, headerRow, headerColumn, series };
}

// ------------------------------------------------------------------ the model of a chart

/** The accent colours of the workbook's theme, as the reader gives them to a chart. */
function paletteOf(workbook: Workbook): string[] {
  const theme = (workbook.styles && (workbook.styles as { themeColors?: string[] }).themeColors) || [];
  const palette: string[] = [];
  for (let i = 0; i < 6; i++) {
    const color = theme[4 + i];
    palette.push('#' + (typeof color === 'string' && HEX6.test(color) ? color.toUpperCase() : OFFICE_ACCENTS[i]));
  }
  return palette;
}

/** A reference of a chart as it is kept: one that names cells of this workbook; one without a sheet is of `sheet`. Throws RangeError otherwise. */
function checkedRef(workbook: Workbook, sheet: Sheet, text: string, what: string): string {
  let ref = typeof text === 'string' ? text.trim() : '';
  if (ref.charAt(0) === '=') ref = ref.slice(1).trim();
  if (ref === '' || ref.length > 8192) throw new RangeError(`The ${what} of a series are a range of cells`);
  if (chartRefAreas(workbook, ref, true)) return ref;
  const own = `${quoteSheetName(sheet.name)}!${ref}`;
  if (ref.indexOf('!') < 0 && ref.indexOf(',') < 0 && chartRefAreas(workbook, own, true)) return own;
  throw new RangeError(`"${ref.slice(0, 80)}" is not a range of cells of this workbook`);
}

function checkedText(text: unknown, what: string): string | undefined {
  if (text === undefined || text === null) return undefined;
  const value = String(text).replace(/\r\n?/g, '\n').replace(/^\n+|\n+$/g, '');
  if (value.trim() === '') return undefined;
  if (value.length > MAX_CHART_TEXT) throw new RangeError(`The ${what} cannot be longer than ${MAX_CHART_TEXT} characters`);
  if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/.test(value)) throw new RangeError(`The ${what} holds a character a file cannot hold`);
  return value;
}

function checkedColor(color: unknown): string | undefined {
  if (color === undefined || color === null) return undefined;
  if (typeof color !== 'string' || !COLOR.test(color)) throw new RangeError('A colour is written "#RRGGBB"');
  return color.toUpperCase();
}

/** Whether the cells a reference names hold numbers and nothing else (empty cells aside): what makes categories numbers. */
function holdsNumbers(workbook: Workbook, ref: string): boolean {
  const areas = chartRefAreas(workbook, ref, true);
  if (!areas) return false;
  let some = false;
  let seen = 0;
  for (const { sheet, range } of areas) {
    const bottom = Math.min(range.bottom, sheet.rowCount - 1);
    for (let row = range.top; row <= bottom; row++) {
      const held = sheet.rows.get(row);
      if (!held) continue;
      for (let col = range.left; col <= range.right && col < held.cells.length; col++) {
        if (++seen > 100000) return some;
        const cell = held.cells[col];
        if (isEmpty(cell)) continue;
        if (!isNumber(cell)) return false;
        some = true;
      }
    }
  }
  return some;
}

const emptyData = (kind: 'num' | 'str', ref: string): ChartData => (kind === 'num' ? { kind, ref, cache: { index: [], values: [] }, count: 0, formatCode: 'General' } : { kind, ref, cache: { index: [], values: [] }, count: 0 });

/** How the kind of a chart type is written in the model. */
function kindOf(type: ChartType): { kind: ChartGroup['kind']; barDir?: 'col' | 'bar' } {
  if (type === 'column') return { kind: 'bar', barDir: 'col' };
  if (type === 'bar') return { kind: 'bar', barDir: 'bar' };
  return { kind: type };
}

/** The look a series has by its colour, for the kind of chart it is in; undefined for the look the chart gives by itself. */
function lookOf(kind: ChartGroup['kind'], color: string | undefined): Pick<ChartSeries, 'shape' | 'marker'> {
  if (kind === 'scatter') {
    // Points without a line between them.
    const marker: ChartSeries['marker'] = { symbol: 'circle', size: 5 };
    if (color !== undefined) marker.fill = color;
    return { shape: { line: { color: null, width: 2 } }, marker };
  }
  if (kind === 'line') return color !== undefined ? { shape: { line: { color } }, marker: { symbol: 'none' } } : { marker: { symbol: 'none' } };
  return color !== undefined ? { shape: { fill: color } } : {};
}

/** The colour a series was given, whatever kind of chart it is in. */
function colorOf(kind: ChartGroup['kind'], series: ChartSeries): string | undefined {
  let color: string | null | undefined;
  if (kind === 'scatter') color = (series.marker && series.marker.fill) || (series.shape && series.shape.line && series.shape.line.color) || undefined;
  else if (kind === 'line') color = series.shape && series.shape.line ? series.shape.line.color : undefined;
  else color = series.shape ? series.shape.fill : undefined;
  return typeof color === 'string' && COLOR.test(color) ? color.toUpperCase() : undefined;
}

function specOf(kind: ChartGroup['kind'], series: ChartSeries): ChartSeriesSpec {
  const out: ChartSeriesSpec = {};
  if (series.name && series.name.ref !== undefined) out.nameRef = series.name.ref;
  else if (series.name && series.name.text !== undefined) out.name = series.name.text;
  if (series.categories && series.categories.ref !== undefined) out.categories = series.categories.ref;
  if (series.values && series.values.ref !== undefined) out.values = series.values.ref;
  const color = colorOf(kind, series);
  if (color !== undefined) out.color = color;
  return out;
}

function seriesOf(workbook: Workbook, sheet: Sheet, kind: ChartGroup['kind'], spec: ChartSeriesSpec, place: number): ChartSeries {
  const series: ChartSeries = { index: place, order: place };
  if (spec.nameRef !== undefined && spec.nameRef !== null) series.name = { ref: checkedRef(workbook, sheet, spec.nameRef, 'name') };
  else {
    const text = checkedText(spec.name, 'name of a series');
    if (text !== undefined) series.name = { text };
  }
  if (spec.categories !== undefined && spec.categories !== null) {
    const ref = checkedRef(workbook, sheet, spec.categories, 'categories');
    series.categories = emptyData(kind === 'scatter' || holdsNumbers(workbook, ref) ? 'num' : 'str', ref);
  }
  if (spec.values === undefined || spec.values === null) throw new RangeError('A series needs a range of values');
  series.values = emptyData('num', checkedRef(workbook, sheet, spec.values, 'values'));
  Object.assign(series, lookOf(kind, checkedColor(spec.color)));
  return series;
}

function titleOf(text: string | undefined): { text: string } | undefined {
  return text === undefined ? undefined : { text };
}

const CATEGORY_AXIS = 1;
const VALUE_AXIS = 2;

function categoryAxis(position: 'b' | 'l'): ChartAxis {
  return { id: CATEGORY_AXIS, kind: 'cat', position, tickLabels: 'nextTo', formatCode: 'General', sourceLinked: true, crosses: 'autoZero', crossAxis: VALUE_AXIS };
}

function valueAxis(id: number, position: 'b' | 'l', crossAxis: number, between: 'between' | 'midCat', gridlines: boolean): ChartAxis {
  const axis: ChartAxis = { id, kind: 'val', position, tickLabels: 'nextTo', formatCode: 'General', sourceLinked: true, crosses: 'autoZero', crossAxis, crossBetween: between };
  if (gridlines) axis.majorGridlines = {};
  return axis;
}

/** The axes a chart of one type has when it is made here. */
function axesOf(type: ChartType): ChartAxis[] {
  if (type === 'pie' || type === 'doughnut') return [];
  if (type === 'scatter') return [valueAxis(CATEGORY_AXIS, 'b', VALUE_AXIS, 'midCat', false), valueAxis(VALUE_AXIS, 'l', CATEGORY_AXIS, 'midCat', true)];
  if (type === 'bar') return [categoryAxis('l'), valueAxis(VALUE_AXIS, 'b', CATEGORY_AXIS, 'between', true)];
  return [categoryAxis('b'), valueAxis(VALUE_AXIS, 'l', CATEGORY_AXIS, type === 'area' ? 'midCat' : 'between', true)];
}

/** The one chart type (group) of a chart that is made here, without its series. */
function groupOf(type: ChartType, grouping: ChartGrouping): ChartGroup {
  const { kind, barDir } = kindOf(type);
  const round = type === 'pie' || type === 'doughnut';
  const group: ChartGroup = { kind, grouping: round || type === 'scatter' ? 'standard' : grouping, axisIds: round ? [] : [CATEGORY_AXIS, VALUE_AXIS], series: [], varyColors: round };
  if (kind === 'bar') {
    group.barDir = barDir;
    group.gapWidth = 150;
    if (group.grouping !== 'standard') group.overlap = 100;
  }
  if (kind === 'line') group.markers = true;
  if (round) group.firstSliceAngle = 0;
  if (type === 'doughnut') group.holeSize = 50;
  if (type === 'scatter') group.scatterStyle = 'lineMarker';
  return group;
}

function checkType(type: unknown): ChartType {
  if (typeof type !== 'string' || CHART_TYPES.indexOf(type as ChartType) < 0) throw new RangeError(`A chart is one of: ${CHART_TYPES.join(', ')}`);
  return type as ChartType;
}

function checkGrouping(grouping: unknown): ChartGrouping {
  if (grouping === undefined) return 'standard';
  if (grouping !== 'standard' && grouping !== 'stacked' && grouping !== 'percentStacked') throw new RangeError('The series of a chart stand side by side ("standard"), are stacked ("stacked") or stacked to a hundred per cent ("percentStacked")');
  return grouping;
}

function checkLegend(legend: unknown): ChartLegendPlace {
  if (legend === undefined) return 'r';
  if (legend !== 'r' && legend !== 'l' && legend !== 't' && legend !== 'b' && legend !== 'tr' && legend !== 'none') throw new RangeError('The legend of a chart stands at "r", "l", "t", "b" or "tr", or there is "none"');
  return legend;
}

/** The type of a chart of the model as the commands name it; undefined for one that is not made here. */
function typeOf(chart: Chart): ChartType | undefined {
  if (!chart.groups || chart.groups.length !== 1) return undefined;
  const group = chart.groups[0];
  if (group.threeD || group.ofPie) return undefined;
  if (group.kind === 'bar') return group.barDir === 'bar' ? 'bar' : 'column';
  if (group.kind === 'line' || group.kind === 'area' || group.kind === 'pie' || group.kind === 'doughnut' || group.kind === 'scatter') return group.kind;
  return undefined;
}

const LABELS_ON: ChartDataLabels = { showLegendKey: false, showValue: true, showCategory: false, showSeries: false, showPercent: false, showBubbleSize: false };

function labelsShown(chart: Chart): boolean {
  const shows = (labels: ChartDataLabels | undefined) => !!labels && !labels.deleted && (labels.showValue === true || labels.showPercent === true || labels.showCategory === true || labels.showSeries === true);
  return chart.groups.some((group) => shows(group.dataLabels) || group.series.some((series) => shows(series.dataLabels)));
}

/** The axis the categories stand on (the horizontal values of a scatter chart), and the one of the values, of the first chart type. */
function axesNamed(chart: Chart): { category?: ChartAxis; value?: ChartAxis } {
  const group = chart.groups[0];
  const own = (chart.axes || []).filter((axis) => !group || group.axisIds.indexOf(axis.id) >= 0);
  if (group && group.kind === 'scatter') {
    // The two axes are told apart the way they are drawn: by where they say they lie, and where
    // both say the same (a file that states "l" for each) the first the chart type names is the horizontal one.
    const pair = scatterAxes(chart, group);
    return { category: pair.x, value: pair.y };
  }
  return { category: own.find((axis) => axis.kind === 'cat'), value: own.find((axis) => axis.kind === 'val') };
}

// ------------------------------------------------------------------ anchors

/** An anchor that covers a rectangle of the sheet (pixels from the corner of cell A1) and moves and sizes with the cells under it. */
export function anchorOfRect(metrics: SheetMetrics, x: number, y: number, width: number, height: number): DrawingAnchor {
  const limitX = metrics.cols.total;
  const limitY = metrics.rows.total;
  const w = Math.max(Math.min(width, limitX), MIN_CHART_SIZE);
  const h = Math.max(Math.min(height, limitY), MIN_CHART_SIZE);
  const left = Math.max(0, Math.min(x, limitX - w));
  const top = Math.max(0, Math.min(y, limitY - h));
  const marker = (px: number, py: number): DrawingMarker => {
    const col = metrics.cols.indexAt(px);
    const row = metrics.rows.indexAt(py);
    return { col, colOff: Math.max(0, Math.round((px - metrics.cols.offset(col)) * EMU_PER_PIXEL)), row, rowOff: Math.max(0, Math.round((py - metrics.rows.offset(row)) * EMU_PER_PIXEL)) };
  };
  return { kind: 'twoCell', from: marker(left, top), to: marker(left + w, top + h) };
}

function checkedAnchor(anchor: DrawingAnchor | CellRange): DrawingAnchor {
  if (isRange(anchor)) {
    const to = { col: Math.min(anchor.right + 1, MAX_COLS - 1), row: Math.min(anchor.bottom + 1, MAX_ROWS - 1) };
    return { kind: 'twoCell', from: { col: anchor.left, colOff: 0, row: anchor.top, rowOff: 0 }, to: { col: to.col, colOff: 0, row: to.row, rowOff: 0 } };
  }
  const given = anchor as DrawingAnchor;
  const whole = (n: unknown, limit: number) => typeof n === 'number' && n >= 0 && n < limit && Math.floor(n) === n;
  const offset = (n: unknown) => typeof n === 'number' && n >= 0 && n <= 2e12 && Number.isFinite(n);
  const marker = (m: DrawingMarker | undefined): DrawingMarker | undefined =>
    m && whole(m.col, MAX_COLS) && whole(m.row, MAX_ROWS) && offset(m.colOff) && offset(m.rowOff) ? { col: m.col, colOff: Math.round(m.colOff), row: m.row, rowOff: Math.round(m.rowOff) } : undefined;
  if (given && given.kind === 'twoCell') {
    const from = marker(given.from);
    const to = marker(given.to);
    if (from && to && (to.row > from.row || (to.row === from.row && to.rowOff >= from.rowOff)) && (to.col > from.col || (to.col === from.col && to.colOff >= from.colOff))) {
      const out: DrawingAnchor = { kind: 'twoCell', from, to };
      if (given.editAs === 'oneCell' || given.editAs === 'absolute') out.editAs = given.editAs;
      return out;
    }
  } else if (given && given.kind === 'oneCell') {
    const from = marker(given.from);
    if (from && offset(given.width) && offset(given.height)) return { kind: 'oneCell', from, width: Math.round(given.width as number), height: Math.round(given.height as number) };
  } else if (given && given.kind === 'absolute') {
    if (offset(given.x) && offset(given.y) && offset(given.width) && offset(given.height)) {
      return { kind: 'absolute', x: Math.round(given.x as number), y: Math.round(given.y as number), width: Math.round(given.width as number), height: Math.round(given.height as number) };
    }
  }
  throw new RangeError('The place of a chart is an anchor, or a range of cells');
}

/** Where a new chart is put when nobody says: beside the cells it shows (below them when there is no room beside), in the default size. */
function defaultAnchor(sheet: Sheet, range: CellRange | undefined): DrawingAnchor {
  const metrics = createSheetMetrics(sheet);
  let col = range ? range.right + 2 : 1;
  let row = range ? range.top : 1;
  if (col >= MAX_COLS - 32) {
    col = range ? range.left : 1;
    row = range ? Math.min(range.bottom + 2, MAX_ROWS - 64) : 1;
  }
  return anchorOfRect(metrics, metrics.cols.offset(col), metrics.rows.offset(Math.min(row, MAX_ROWS - 64)), DEFAULT_CHART_WIDTH, DEFAULT_CHART_HEIGHT);
}

// ------------------------------------------------------------------ commands

/** Why no chart can be added to a sheet, moved or deleted on it; undefined when that can be done. */
export function chartsProblem(workbook: Workbook, sheetIndex: number): string | undefined {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet) return `No sheet at index ${sheetIndex}`;
  if (sheet.kind !== 'worksheet') return `Cannot change the charts of a ${sheet.kind}`;
  return placesProblem(sheet);
}

/**
 * Why no chart of a sheet is added, moved or deleted: the drawing part of the file takes no
 * change (Sheet.drawingsFixed). A chart of such a sheet is still changed in its own part.
 */
function placesProblem(sheet: Sheet): string | undefined {
  return sheet.drawingsFixed !== undefined && sheet.part !== '' ? `Cannot change the charts of this sheet: ${sheet.drawingsFixed}` : undefined;
}

/** The sheet of a command that changes where charts are, or which there are. Throws when its drawing takes no change. */
function placedSheetOf(workbook: Workbook, sheetIndex: number): Sheet {
  const sheet = sheetOf(workbook, sheetIndex);
  const problem = placesProblem(sheet);
  if (problem) throw new RangeError(problem);
  return sheet;
}

function setDrawings(sheetIndex: number, list: SheetDrawing[]): Op[] {
  return [{ type: 'setDrawings', sheet: sheetIndex, drawings: list.length > 0 ? list : undefined }];
}

/** Throws RangeError when a sheet has no room for one more chart. */
function checkRoom(workbook: Workbook, sheet: Sheet): void {
  if ((sheet.drawings || []).length >= MAX_SHEET_DRAWINGS) throw new RangeError('The sheet has too many drawings');
  // No more charts than are read back from a file: what is written can be opened again.
  const chartsOn = (other: Sheet): number => (other.drawings || []).reduce((count, drawing) => (drawing && drawing.object && drawing.object.type === 'chart' ? count + 1 : count), 0);
  if (chartsOn(sheet) >= MAX_SHEET_CHARTS) throw new RangeError(`A sheet has at most ${MAX_SHEET_CHARTS} charts`);
  if (workbook.sheets.reduce((count, other) => count + chartsOn(other), 0) >= MAX_WORKBOOK_CHARTS) throw new RangeError(`A workbook has at most ${MAX_WORKBOOK_CHARTS} charts`);
}

/** Cells that are looked at when the charts of a workbook are weighed; beyond them a chart weighs what its model caches. */
const MAX_WEIGHED_CELLS = 2000000;

/**
 * What a chart will weigh in the file that is written, as the budgets of a reader count it: the
 * bytes of its part and the values its part caches. A chart of the file that was not changed
 * weighs what its part was read as. Any other is written with the cells as they are: its values
 * are counted from the cells, and its bytes are what its part had (the part a copy is made
 * from, the part that is changed) and what the values and the series take - an estimate that
 * errs on the side of too much.
 */
function chartWeight(workbook: Workbook, object: DrawingChart, budget: { cells: number }): { bytes: number; points: number } {
  const read = chartPartCost(object.origin ? object.origin.chart : object.chart);
  if (object.part !== undefined && !object.origin && read) return read;
  const chart = object.chart;
  let points = 0;
  let series = 0;
  const count = (ref: string | undefined, cached: number): void => {
    const areas = ref !== undefined && budget.cells > 0 ? chartRefAreas(workbook, ref, chart.live === true) : undefined;
    if (!areas) {
      points += cached;
      return;
    }
    let seen = 0;
    let held = 0;
    for (const { sheet, range } of areas) {
      const bottom = Math.min(range.bottom, sheet.rowCount - 1);
      for (let row = range.top; row <= bottom && seen < MAX_CHART_POINTS; row++) {
        const cells = sheet.rows.get(row);
        seen += range.right - range.left + 1;
        if (!cells) continue;
        for (let col = range.left; col <= range.right && col < cells.cells.length; col++) if (!isEmpty(cells.cells[col])) held++;
      }
    }
    budget.cells -= seen;
    points += Math.min(held, MAX_CHART_POINTS);
  };
  for (const group of chart.groups) {
    for (const item of group.series) {
      series++;
      if (item.name && item.name.ref !== undefined) points++;
      for (const data of [item.categories, item.values, item.bubbleSizes]) if (data) count(data.ref, data.cache ? data.cache.index.length : 0);
    }
  }
  return { bytes: (read ? read.bytes : 2048) + 64 * points + 512 * series, points: Math.max(points, read ? read.points : 0) };
}

/**
 * Throws RangeError when the charts of a workbook, with `adding`, would hold more than a
 * reader takes of one workbook (MAX_WORKBOOK_CHART_BYTES, MAX_WORKBOOK_CHART_POINTS,
 * MAX_CHART_PART_BYTES for one part): a chart beyond that would be a placeholder when the file
 * is opened again.
 */
function checkBudget(workbook: Workbook, adding: DrawingChart): void {
  const budget = { cells: MAX_WEIGHED_CELLS };
  const added = chartWeight(workbook, adding, budget);
  let bytes = added.bytes;
  let points = added.points;
  for (const sheet of workbook.sheets) {
    for (const drawing of sheet.drawings || []) {
      if (!drawing || !drawing.object || drawing.object.type !== 'chart') continue;
      const weight = chartWeight(workbook, drawing.object, budget);
      bytes += weight.bytes;
      points += weight.points;
    }
  }
  if (added.bytes > MAX_CHART_PART_BYTES || bytes > MAX_WORKBOOK_CHART_BYTES || points > MAX_WORKBOOK_CHART_POINTS) {
    throw new RangeError('Cannot add the chart: the charts of the workbook would hold more than is read back from a file');
  }
}

/** The name a chart that is added to a sheet gets: "Chart" and one more than the charts the sheet has. */
function nextName(sheet: Sheet): string {
  let number = 1;
  for (const drawing of sheet.drawings || []) if (drawing && drawing.object && drawing.object.type === 'chart') number++;
  return `Chart ${number}`;
}

/**
 * Adds a chart to a sheet, above its other drawings. Returns the operations; the chart is the
 * last entry of Sheet.drawings once they are applied.
 * Throws RangeError for a chart that cannot be made (no values, a reference that names no
 * cells of the workbook, an unknown type, too many drawings).
 */
export function addChart(workbook: Workbook, sheetIndex: number, spec: ChartSpec): Op[] {
  return withCommand(addChartPlain(workbook, sheetIndex, spec), 'addChart', [spec]);
}

function addChartPlain(workbook: Workbook, sheetIndex: number, spec: ChartSpec): Op[] {
  const sheet = placedSheetOf(workbook, sheetIndex);
  if (!spec || typeof spec !== 'object') throw new RangeError('A chart needs a type and a range of cells');
  const type = checkType(spec.type);
  const grouping = checkGrouping(spec.grouping);
  const legend = checkLegend(spec.legend);
  const drawings = sheet.drawings || [];
  checkRoom(workbook, sheet);
  const group = groupOf(type, grouping);
  const specs = Array.isArray(spec.series) ? spec.series : spec.range !== undefined ? seriesOfRange(workbook, sheetIndex, spec.range, spec.seriesIn, type).series : undefined;
  if (!specs || specs.length === 0) throw new RangeError('A chart needs a range of cells, or series');
  if (specs.length > MAX_CHART_SERIES) throw new RangeError(`A chart has at most ${MAX_CHART_SERIES} series`);
  group.series = specs.map((item, i) => seriesOf(workbook, sheet, group.kind, item, i));
  const chart: Chart = { live: true, groups: [group], axes: axesOf(type), plotVisOnly: true, blanks: 'gap', palette: paletteOf(workbook), notes: [] };
  const title = checkedText(spec.title, 'title');
  if (title !== undefined) chart.title = { text: title };
  else chart.autoTitleDeleted = true;
  if (legend !== 'none') chart.legend = { position: legend };
  if (workbook.date1904) chart.date1904 = true;
  const anchor = spec.anchor !== undefined ? checkedAnchor(spec.anchor) : defaultAnchor(sheet, isRange(spec.range) ? spec.range : undefined);
  const made: DrawingChart = { type: 'chart', chart, name: nextName(sheet) };
  checkBudget(workbook, made);
  const added: SheetDrawing = { anchor, object: made, since: sheet.structuralEdits ? sheet.structuralEdits.length : 0 };
  return setDrawings(sheetIndex, drawings.concat([added]));
}

/**
 * Deletes a chart. Throws RangeError when `id` names none, and for a chart of the file that
 * shares a wrapper of alternatives with another drawing (SheetDrawing.wrapped).
 */
export function removeChart(workbook: Workbook, sheetIndex: number, id: number): Op[] {
  return withCommand(removeChartPlain(workbook, sheetIndex, id), 'removeChart', [id]);
}

function removeChartPlain(workbook: Workbook, sheetIndex: number, id: number): Op[] {
  const sheet = placedSheetOf(workbook, sheetIndex);
  const { drawing } = chartAt(sheet, id);
  const inFile = fileDrawing(drawing);
  if (inFile && inFile.wrapped) throw new RangeError('Cannot delete this chart: it shares its place in the drawing with another object');
  const next = (sheet.drawings as SheetDrawing[]).slice();
  next.splice(id, 1);
  return setDrawings(sheetIndex, next);
}

/** A drawing of the file as it was read, for one that was moved or changed since; the drawing itself for one of the file that was not; undefined for one that is in no file. */
function fileDrawing(drawing: SheetDrawing): SheetDrawing | undefined {
  if (drawing.from) return drawing.from;
  return drawing.since === undefined ? drawing : undefined;
}

/** A copy of a drawing that stands where the sheet is now, with what `change` puts in place of what it had. */
function changed(sheet: Sheet, drawing: SheetDrawing, change: Partial<Pick<SheetDrawing, 'anchor' | 'object'>>): SheetDrawing {
  const count = sheet.structuralEdits ? sheet.structuralEdits.length : 0;
  const next: SheetDrawing = { anchor: change.anchor || drawingAnchorNow(drawing, sheet.structuralEdits), object: change.object || drawing.object, since: count };
  const from = fileDrawing(drawing);
  if (from) next.from = from;
  return next;
}

/**
 * Moves or resizes a chart: `anchor` is where it stands from now on (an anchor, or the cells it
 * covers). Returns no operation when the chart is there already.
 */
export function moveChart(workbook: Workbook, sheetIndex: number, id: number, anchor: DrawingAnchor | CellRange): Op[] {
  return withCommand(moveChartPlain(workbook, sheetIndex, id, anchor), 'moveChart', [id, anchor]);
}

function moveChartPlain(workbook: Workbook, sheetIndex: number, id: number, anchor: DrawingAnchor | CellRange): Op[] {
  const sheet = placedSheetOf(workbook, sheetIndex);
  const { drawing } = chartAt(sheet, id);
  const wanted = checkedAnchor(anchor);
  if (JSON.stringify(wanted) === JSON.stringify(drawingAnchorNow(drawing, sheet.structuralEdits))) return [];
  const next = (sheet.drawings as SheetDrawing[]).slice();
  next[id] = changed(sheet, drawing, { anchor: wanted });
  return setDrawings(sheetIndex, next);
}

/**
 * Why a chart cannot be changed the way `patch` says, as a message; undefined when it can.
 * Without a patch: why it cannot be changed at all. With one, the change is worked out and
 * thrown away: what this lets pass, updateChart() takes and the writer writes.
 */
export function chartProblem(workbook: Workbook, sheetIndex: number, id: number, patch?: ChartPatch): string | undefined {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet) return `No sheet at index ${sheetIndex}`;
  if (sheet.kind !== 'worksheet') return `Cannot change the charts of a ${sheet.kind}`;
  const drawing = sheet.drawings ? sheet.drawings[id] : undefined;
  if (!drawing || !drawing.object || drawing.object.type !== 'chart') return `No chart at index ${id}`;
  const object = drawing.object;
  if (object.chart && object.chart.shownOnly === true) return shownOnly(object.chart);
  if (object.part !== undefined && object.known !== 'all' && object.known !== 'refs') return 'Cannot change this chart: its part holds references the model does not, so it is only moved and deleted';
  if (!patch) return undefined;
  try {
    updateChart(workbook, sheetIndex, id, patch);
  } catch (error) {
    if (error instanceof RangeError) return error.message;
    throw error;
  }
  return undefined;
}

/** The origin of a chart of the file that is about to be changed: the one it has, or a fresh one that says every series is where it was. */
function originOf(object: DrawingChart): ChartOrigin {
  return object.origin || { chart: object.chart, series: object.chart.groups.map((group) => group.series.map((_series, i) => i)) };
}

function withTitle<T extends { title?: Chart['title'] }>(holder: T, text: string | undefined): T {
  const next = { ...holder };
  if (text === undefined) delete next.title;
  // What the title had beside its text (its look, its place) stays; a reference gives way to the text.
  else {
    const title = { ...(holder.title || {}), text };
    delete title.ref;
    next.title = title;
  }
  return next;
}

/** A chart of one type made of the series of another, for charts the engine writes itself. */
function retyped(chart: Chart, type: ChartType, grouping: ChartGrouping, workbook: Workbook): Chart {
  const before = axesNamed(chart);
  const fromKind = chart.groups[0].kind;
  const group = groupOf(type, grouping);
  const all: ChartSeries[] = [];
  for (const old of chart.groups) for (const series of old.series) all.push(series);
  group.series = all.map((series, i) => {
    const next: ChartSeries = { index: series.index, order: i };
    if (series.name) next.name = series.name;
    if (series.categories) {
      const ref = series.categories.ref;
      const kind = group.kind === 'scatter' ? 'num' : ref !== undefined && holdsNumbers(workbook, ref) ? 'num' : series.categories.kind === 'num' && fromKind !== 'scatter' ? 'num' : 'str';
      next.categories = kind === series.categories.kind ? series.categories : ref !== undefined ? emptyData(kind, ref) : series.categories;
    }
    if (series.values) next.values = series.values;
    Object.assign(next, lookOf(group.kind, colorOf(fromKind, series)));
    if (series.dataLabels) next.dataLabels = series.dataLabels;
    return next;
  });
  if (chart.groups[0].dataLabels) group.dataLabels = chart.groups[0].dataLabels;
  const axes = axesOf(type).map((axis, i) => {
    const old = i === 0 ? before.category : before.value;
    let next = axis;
    if (old && old.title && old.title.text !== undefined) next = { ...next, title: { text: old.title.text } };
    if (i === 1 && old && old.min !== undefined) next = { ...next, min: old.min };
    if (i === 1 && old && old.max !== undefined) next = { ...next, max: old.max };
    return next;
  });
  const next: Chart = { ...chart, groups: [group], axes };
  return next;
}

/** One series changed by an edit; the series itself when the edit says nothing new. */
function editedSeries(workbook: Workbook, sheet: Sheet, kind: ChartGroup['kind'], series: ChartSeries, edit: ChartSeriesEdit): ChartSeries {
  let next = series;
  const set = (change: Partial<ChartSeries>, remove?: keyof ChartSeries) => {
    next = { ...next, ...change };
    if (remove) delete next[remove];
  };
  if (edit.nameRef !== undefined && edit.nameRef !== null) {
    const ref = checkedRef(workbook, sheet, edit.nameRef, 'name');
    if (!series.name || series.name.ref !== ref) set({ name: { ref } });
  } else if (edit.name !== undefined || edit.nameRef === null) {
    const text = checkedText(edit.name, 'name of a series');
    if (text === undefined) {
      if (series.name) set({}, 'name');
    } else if (!series.name || series.name.ref !== undefined || series.name.text !== text) set({ name: { text } });
  }
  if (edit.categories === null) {
    if (series.categories) set({}, 'categories');
  } else if (edit.categories !== undefined) {
    const ref = checkedRef(workbook, sheet, edit.categories, 'categories');
    if (!series.categories || series.categories.ref !== ref) {
      // The form the file has the categories in stays where it can: numbers stay numbers.
      const form = kind === 'scatter' ? 'num' : series.categories && series.categories.ref !== undefined && !series.categories.multiLevel ? series.categories.kind : holdsNumbers(workbook, ref) ? 'num' : 'str';
      set({ categories: emptyData(form, ref) });
    }
  }
  if (edit.values !== undefined) {
    const ref = checkedRef(workbook, sheet, edit.values, 'values');
    if (!series.values || series.values.ref !== ref) set({ values: emptyData('num', ref) });
  }
  if (edit.color !== undefined) {
    const color = checkedColor(edit.color);
    if (color !== colorOf(kind, series)) {
      if (kind === 'scatter') {
        const marker = { ...(next.marker || { symbol: 'circle', size: 5 }) };
        if (color === undefined) delete marker.fill;
        else marker.fill = color;
        set({ marker });
      } else if (kind === 'line') {
        const line: ChartLine = { ...((next.shape && next.shape.line) || {}) };
        if (color === undefined) delete line.color;
        else line.color = color;
        const shape: ChartShape = { ...(next.shape || {}), line };
        if (Object.keys(line).length === 0) delete shape.line;
        if (Object.keys(shape).length === 0) set({}, 'shape');
        else set({ shape });
      } else {
        const shape = { ...(next.shape || {}) };
        if (color === undefined) delete shape.fill;
        else shape.fill = color;
        if (Object.keys(shape).length === 0) set({}, 'shape');
        else set({ shape });
      }
    }
  }
  return next;
}

const sameValue = (a: unknown, b: unknown): boolean => JSON.stringify(a === undefined ? null : a) === JSON.stringify(b === undefined ? null : b);

/**
 * A chart of the file as its part will say it once the writer changed the part: where the part
 * has a title, a legend or data labels, what they hold beside what was changed (their look,
 * their place) is what the part holds, whatever was taken away and put back in between.
 */
function settled(chart: Chart, read: Chart): Chart {
  let next = chart;
  const title = <T extends { title?: Chart['title'] }>(holder: T, was: Chart['title']): T => {
    const now = holder.title;
    if (!now || now.ref !== undefined || now.text === undefined || !was) return holder;
    const kept = { ...was, text: now.text };
    delete kept.ref;
    return sameValue(kept, now) ? holder : { ...holder, title: kept };
  };
  next = title(next, read.title);
  if (next.legend && read.legend) {
    // Its place and the entries that are left out are the model's; its look and where it was put by hand are the part's.
    const kept = { ...read.legend, position: next.legend.position };
    if (next.legend.deleted) kept.deleted = next.legend.deleted;
    else delete kept.deleted;
    if (!sameValue(kept, next.legend)) next = { ...next, legend: kept };
  }
  const axes = next.axes.map((axis) => {
    const was = read.axes.find((other) => other.id === axis.id);
    return was ? title(axis, was.title) : axis;
  });
  if (axes.some((axis, i) => axis !== next.axes[i])) next = { ...next, axes };
  const groups = next.groups.map((group, g) => {
    const was = read.groups[g] ? read.groups[g].dataLabels : undefined;
    const now = group.dataLabels;
    if (!now || !was || was.deleted || now.deleted) return group;
    const kept: ChartDataLabels = { ...was };
    for (const key of ['showLegendKey', 'showValue', 'showCategory', 'showSeries', 'showPercent', 'showBubbleSize'] as const) kept[key] = now[key] === true;
    return sameValue(kept, now) ? group : { ...group, dataLabels: kept };
  });
  if (groups.some((group, g) => group !== next.groups[g])) next = { ...next, groups };
  return next;
}

/**
 * Changes a chart: see ChartPatch for what can be said. A chart of the file that the engine
 * did not write is changed in what CHART_PART_CHANGES lists (chartProblem() says why not).
 * Returns no operation when nothing would change. Throws RangeError for what cannot be done.
 */
export function updateChart(workbook: Workbook, sheetIndex: number, id: number, patch: ChartPatch): Op[] {
  return withCommand(updateChartPlain(workbook, sheetIndex, id, patch), 'updateChart', [id, patch]);
}

function updateChartPlain(workbook: Workbook, sheetIndex: number, id: number, patch: ChartPatch): Op[] {
  const sheet = sheetOf(workbook, sheetIndex);
  const { drawing, object } = chartAt(sheet, id);
  if (!patch || typeof patch !== 'object') return [];
  const problem = chartProblem(workbook, sheetIndex, id);
  if (problem) throw new RangeError(problem);
  const before = object.chart;
  const whole = !partBound(object) || object.known === 'all';
  if (!whole && patch.type !== undefined && checkType(patch.type) !== typeOf(before)) throw new RangeError('Cannot change the type of this chart: its part holds formatting that another type cannot keep');
  if (!whole && (patch.series !== undefined || patch.range !== undefined || patch.seriesIn !== undefined) && before.groups.length !== 1) throw new RangeError('Cannot change the series of a chart of several types');
  let chart: Chart = before;
  // Which series of the chart as it was each series is now, group by group.
  let places: number[][] = before.groups.map((group) => group.series.map((_series, i) => i));

  const nowType = typeOf(chart);
  const type = patch.type !== undefined ? checkType(patch.type) : nowType;
  const grouping = patch.grouping !== undefined ? checkGrouping(patch.grouping) : chart.groups[0].grouping;
  if (patch.grouping !== undefined && grouping !== 'standard' && (type === 'pie' || type === 'doughnut' || type === 'scatter')) throw new RangeError('Cannot change how the series of this chart are grouped');
  if (patch.type !== undefined && type !== nowType) {
    chart = retyped(chart, type as ChartType, grouping, workbook);
    // Every series is in the one chart type now, in the order they had.
    let at = 0;
    places = [before.groups.reduce((list: number[], group) => list.concat(group.series.map(() => at++)), [])];
  } else if (patch.grouping !== undefined && grouping !== chart.groups[0].grouping) {
    const group: ChartGroup = { ...chart.groups[0], grouping };
    if (!chartKindGrouped(group.kind)) throw new RangeError('Cannot change how the series of this chart are grouped');
    if (group.kind === 'bar' && !group.threeD) {
      if (grouping === 'standard') delete group.overlap;
      else group.overlap = 100;
    }
    chart = { ...chart, groups: [group].concat(chart.groups.slice(1)) };
  }

  if (patch.title !== undefined) {
    const text = checkedText(patch.title, 'title');
    const had = chart.title && chart.title.ref === undefined ? chart.title.text : undefined;
    if (text !== had || (chart.title && chart.title.ref !== undefined)) {
      chart = withTitle(chart, text);
      if (text === undefined) chart.autoTitleDeleted = true;
      else delete chart.autoTitleDeleted;
    }
  }
  if (patch.legend !== undefined) {
    const legend = checkLegend(patch.legend);
    const had: ChartLegendPlace = chart.legend ? chart.legend.position : 'none';
    if (legend !== had) {
      chart = { ...chart };
      if (legend === 'none') delete chart.legend;
      else {
        const fresh: ChartLegend = { ...(chart.legend || {}), position: legend };
        const read = object.origin ? object.origin.chart : undefined;
        if (!chart.legend && object.origin && read && read.legend && read.legend.deleted) {
          // The legend of the part comes back with the entries it leaves out, each at the series it belongs to now.
          const flat: number[] = [];
          let offset = 0;
          object.origin.series.forEach((list, g) => {
            for (const place of list) flat.push(place < 0 ? -1 : offset + place);
            offset += read.groups[g] ? read.groups[g].series.length : 0;
          });
          const deleted = remapLegendEntries(read.legend.deleted, flat);
          if (deleted) fresh.deleted = deleted;
        }
        chart.legend = fresh;
      }
    }
  }
  if (patch.axisTitles || patch.axisBounds) {
    const named = axesNamed(chart);
    const titles = patch.axisTitles || {};
    const bounds = patch.axisBounds || {};
    const axes = chart.axes.map((axis) => {
      let next = axis;
      const wanted = axis === named.category ? titles.category : axis === named.value ? titles.value : undefined;
      if (wanted !== undefined) {
        const text = checkedText(wanted, 'title of an axis');
        const had = axis.title && axis.title.ref === undefined ? axis.title.text : undefined;
        if (text !== had || (axis.title && axis.title.ref !== undefined)) next = withTitle(next, text);
      }
      if (axis === named.value) {
        for (const key of ['min', 'max'] as const) {
          const value = bounds[key];
          if (value === undefined) continue;
          if (value !== null && !(typeof value === 'number' && Number.isFinite(value))) throw new RangeError('The smallest and the largest value of an axis are numbers');
          if (value === null ? next[key] === undefined : next[key] === value) continue;
          next = { ...next };
          if (value === null) delete next[key];
          else next[key] = value;
        }
      }
      return next;
    });
    const value = axes.find((_axis, i) => chart.axes[i] === named.value);
    if (value && value.min !== undefined && value.max !== undefined && !(value.max > value.min)) throw new RangeError('The largest value of an axis is larger than its smallest');
    if ((titles.category !== undefined && titles.category !== null && !named.category) || ((titles.value !== undefined && titles.value !== null) || bounds.min != null || bounds.max != null ? !named.value : false)) {
      throw new RangeError('This chart has no such axis');
    }
    if (axes.some((axis, i) => axis !== chart.axes[i])) chart = { ...chart, axes };
  }
  if (patch.dataLabels !== undefined && !!patch.dataLabels !== labelsShown(chart)) {
    const groups = chart.groups.map((group) => {
      const next: ChartGroup = { ...group, series: group.series.map((series) => (series.dataLabels || series.pointLabels ? (({ dataLabels: _labels, pointLabels: _points, ...rest }) => rest)(series) : series)) };
      if (patch.dataLabels) next.dataLabels = { ...LABELS_ON };
      else delete next.dataLabels;
      return next;
    });
    chart = { ...chart, groups };
  }

  let edits = patch.series;
  if (patch.range !== undefined || (patch.seriesIn !== undefined && edits === undefined)) {
    // The series anew from cells; the ones there are keep their look in the order they stand.
    let range = patch.range;
    if (range === undefined) {
      const info = dataRange(workbook, sheetIndex, chart);
      if (!info) throw new RangeError('The series of this chart are not one block of cells');
      range = info.range;
    }
    const made = seriesOfRange(workbook, sheetIndex, range, patch.seriesIn, type).series;
    const had = chart.groups[0].series.length;
    edits = made.map((item, i) => {
      const edit: ChartSeriesEdit = { values: item.values, categories: item.categories === undefined ? null : item.categories };
      if (item.nameRef !== undefined) edit.nameRef = item.nameRef;
      else edit.name = null;
      if (i < had) edit.from = i;
      return edit;
    });
  }
  if (edits !== undefined) {
    if (!Array.isArray(edits) || edits.length === 0) throw new RangeError('A chart has at least one series');
    if (edits.length > MAX_CHART_SERIES) throw new RangeError(`A chart has at most ${MAX_CHART_SERIES} series`);
    if (chart.groups.length !== 1) throw new RangeError('Cannot change the series of a chart of several types');
    const group = chart.groups[0];
    const used = new Set<number>();
    let top = -1;
    for (const series of group.series) if (series.index > top) top = series.index;
    const list: ChartSeries[] = [];
    const from: number[] = [];
    edits.forEach((edit, i) => {
      if (!edit || typeof edit !== 'object') throw new RangeError('A series is an object');
      let series: ChartSeries;
      if (edit.from !== undefined) {
        const old = Math.floor(edit.from) === edit.from ? group.series[edit.from] : undefined;
        if (!old || used.has(edit.from)) throw new RangeError(`No series at index ${edit.from}, or it is named twice`);
        used.add(edit.from);
        series = editedSeries(workbook, sheet, group.kind, old, edit);
        from.push(places[0][edit.from]);
      } else {
        const spec: ChartSeriesSpec = {};
        if (edit.nameRef) spec.nameRef = edit.nameRef;
        else if (edit.name) spec.name = edit.name;
        if (edit.categories) spec.categories = edit.categories;
        else if (group.series[0] && group.series[0].categories && group.series[0].categories.ref !== undefined && edit.categories !== null) spec.categories = group.series[0].categories.ref;
        spec.values = edit.values;
        if (edit.color) spec.color = edit.color;
        series = seriesOf(workbook, sheet, group.kind, spec, ++top);
        from.push(-1);
      }
      list.push(series.order === i ? series : { ...series, order: i });
    });
    const same = list.length === group.series.length && list.every((series, i) => series === group.series[i]);
    if (!same) {
      chart = { ...chart, groups: [{ ...group, series: list }] };
      places = [from];
      // The entries of the legend that are left out are numbered by the series: each follows its own,
      // when the series are others or stand in another order. (Entries behind the series are those of
      // trend lines; a part that has any keeps its series as they are, and the entries with them.)
      const reordered = from.length !== group.series.length || from.some((was, i) => was !== i);
      if (reordered && chart.legend && chart.legend.deleted) {
        const legend = { ...chart.legend };
        const deleted = remapLegendEntries(legend.deleted, edits.map((edit) => (edit.from !== undefined ? edit.from : -1)));
        if (deleted) legend.deleted = deleted;
        else delete legend.deleted;
        chart.legend = legend;
      }
    }
  }

  if (chart === before) return [];
  if (!chart.live) chart = { ...chart, live: true };
  if (!whole) chart = settled(chart, originOf(object).chart);
  const next: DrawingChart = { ...object, chart };
  if (!whole) {
    const origin = originOf(object);
    next.origin = { chart: origin.chart, series: places.map((group, g) => group.map((place) => (place < 0 ? -1 : origin.series[g] !== undefined && origin.series[g][place] !== undefined ? origin.series[g][place] : -1))) };
    // What the writer could not write into the part is no operation: the same rule says so here.
    const unwritable = chartPartProblem(origin.chart, chart, next.origin.series, object.seriesFixed === true);
    if (unwritable) throw new RangeError(unwritable);
  } else if (!chartWritable(chart)) throw new RangeError('Cannot change this chart: its type is not one that is written from the model');
  const list = (sheet.drawings as SheetDrawing[]).slice();
  list[id] = changed(sheet, drawing, { object: next });
  return setDrawings(sheetIndex, list);
}

/** The block of cells of the sheet the series of a chart are made of, when they are exactly what seriesOfRange() makes of it. */
function dataRange(workbook: Workbook, sheetIndex: number, chart: Chart): { range: CellRange; seriesIn: 'columns' | 'rows' } | undefined {
  const sheet = workbook.sheets[sheetIndex];
  const type = typeOf(chart);
  if (!sheet || chart.groups.length !== 1) return undefined;
  const group = chart.groups[0];
  const refs: string[] = [];
  for (const series of group.series) {
    if (series.name && series.name.ref !== undefined) refs.push(series.name.ref);
    else if (series.name) return undefined;
    if (series.categories && series.categories.ref !== undefined) refs.push(series.categories.ref);
    if (!series.values || series.values.ref === undefined) return undefined;
    refs.push(series.values.ref);
  }
  let box: CellRange | undefined;
  for (const ref of refs) {
    const areas = chartRefAreas(workbook, ref, chart.live === true);
    if (!areas || areas.length !== 1 || areas[0].sheet !== sheet) return undefined;
    const range = areas[0].range;
    box = box ? { top: Math.min(box.top, range.top), left: Math.min(box.left, range.left), bottom: Math.max(box.bottom, range.bottom), right: Math.max(box.right, range.right) } : { ...range };
  }
  if (!box) return undefined;
  const same = (ref: string | undefined, other: string | undefined): boolean => {
    if (ref === undefined || other === undefined) return ref === other;
    const a = chartRefAreas(workbook, ref, chart.live === true);
    const b = chartRefAreas(workbook, other, true);
    return !!a && !!b && a.length === 1 && b.length === 1 && a[0].sheet === b[0].sheet && a[0].range.top === b[0].range.top && a[0].range.left === b[0].range.left && a[0].range.bottom === b[0].range.bottom && a[0].range.right === b[0].range.right;
  };
  for (const seriesIn of ['columns', 'rows'] as const) {
    let made: RangeSeries;
    try {
      made = seriesOfRange(workbook, sheetIndex, box, seriesIn, type);
    } catch (error) {
      continue;
    }
    if (made.series.length !== group.series.length) continue;
    const fits = made.series.every((item, i) => {
      const series = group.series[i];
      return same(series.name ? series.name.ref : undefined, item.nameRef) && same(series.categories ? series.categories.ref : undefined, item.categories) && same(series.values ? series.values.ref : undefined, item.values);
    });
    if (fits) return { range: box, seriesIn };
  }
  return undefined;
}

/** The charts of a sheet, in the order they are drawn (the one on top last). */
export function sheetCharts(workbook: Workbook, sheetIndex: number): ChartInfo[] {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet) throw new RangeError(`No sheet at index ${sheetIndex}`);
  const out: ChartInfo[] = [];
  const drawings = sheet.drawings || [];
  for (let id = 0; id < drawings.length; id++) {
    const drawing = drawings[id];
    if (!drawing || !drawing.object || drawing.object.type !== 'chart') continue;
    const object = drawing.object;
    const chart = object.chart;
    const group = chart.groups[0];
    const named = axesNamed(chart);
    const info: ChartInfo = {
      id,
      grouping: group.grouping,
      legend: chart.legend ? chart.legend.position : 'none',
      anchor: drawingAnchorNow(drawing, sheet.structuralEdits),
      series: [],
      axisTitles: {},
      axisBounds: {},
      dataLabels: labelsShown(chart),
      editable: sheet.kind !== 'worksheet' ? 'none' : !partBound(object) || object.known === 'all' ? 'all' : object.known === 'refs' ? 'parts' : 'none',
      ofFile: object.part !== undefined,
      canAddSeries: false,
      canReorderSeries: false,
    };
    if (info.editable === 'all') info.canAddSeries = info.canReorderSeries = true;
    else if (info.editable === 'parts' && chart.groups.length === 1 && group.kind !== 'stock' && object.seriesFixed !== true) {
      info.canReorderSeries = true;
      info.canAddSeries = chartKindWritten(group.kind);
    }
    const type = typeOf(chart);
    if (type !== undefined) info.type = type;
    if (chart.title && chart.title.text !== undefined) info.title = chart.title.text;
    for (const each of chart.groups) for (const series of each.series) info.series.push(specOf(each.kind, series));
    if (named.category && named.category.title && named.category.title.text !== undefined) info.axisTitles.category = named.category.title.text;
    if (named.value && named.value.title && named.value.title.text !== undefined) info.axisTitles.value = named.value.title.text;
    if (named.value && named.value.min !== undefined) info.axisBounds.min = named.value.min;
    if (named.value && named.value.max !== undefined) info.axisBounds.max = named.value.max;
    const data = sheet.kind === 'worksheet' ? dataRange(workbook, sheetIndex, chart) : undefined;
    if (data) {
      info.range = data.range;
      info.seriesIn = data.seriesIn;
    }
    out.push(info);
  }
  return out;
}

// ------------------------------------------------------------------ copies

/** How far a copy is put from a chart that stands where it would stand: so many pixels to the right and down. */
export const CHART_COPY_OFFSET = 24;

/**
 * A chart taken off a sheet to be put down again: what copyChart() and cutChart() give and
 * pasteChart() takes. It belongs to the workbook it was taken from and is not to be altered.
 * Its references name sheets and cells as they were when it was taken: once rows, columns or
 * sheets changed since, it is put down nowhere (chartCopyOutdated()).
 */
export interface ChartCopy {
  readonly workbook: Workbook;
  /** The sheet it was taken from. */
  readonly sheet: Sheet;
  /** The drawing it was taken from, as the sheet had it then: while the sheet still has that very drawing, the copy is the chart as it is. */
  readonly drawing: SheetDrawing;
  /** The chart as it is put down (its name aside). */
  readonly object: DrawingChart;
  /** Where it stood, in pixels from the corner of cell A1, and how large it was. */
  readonly rect: { x: number; y: number; width: number; height: number };
  readonly editAs?: DrawingAnchor['editAs'];
  /**
   * What of the chart no copy takes along: elements of its part that are left out ("userShapes",
   * "externalData"), the macro its frame is assigned ("macro"), relationships of its part that
   * lead to relationships ("relationships").
   */
  readonly dropped: readonly string[];
  /**
   * The sheets the copy stands on as they were: their names and the rows and columns that had
   * been inserted and deleted. These are the sheets its references name (and the sheet it was
   * taken from, for references that name none); with `strict`, every sheet of the workbook.
   */
  readonly state: readonly { sheet: Sheet; name: string; edits: readonly SheetStructuralEdit[] }[];
  /** A reference of the copy cannot be followed to cells (a sheet that is gone, another workbook): any row or column that moves anywhere outdates the copy. */
  readonly strict?: boolean;
}

/** What pasteChart() and duplicateChart() give. */
export interface PastedChart {
  /** The operations that add the chart: one undo step. */
  ops: Op[];
  /** The id the chart has on its sheet once they are applied. */
  id: number;
  /** What of the chart that was copied the new chart does not have (ChartCopy.dropped). Nothing else of it is lost. */
  dropped: string[];
  /**
   * References of the chart that name no sheet and are not cells (a name defined for the sheet
   * the chart came from): on another sheet they could not be bound to the sheet they were written
   * on, and stay as they are.
   */
  unbound: number;
}

/**
 * Why a chart cannot be copied (or cut, or duplicated), as a message; undefined when it can. A
 * chart of the file whose part holds references the model does not (a pivot chart, formulas of
 * extensions) is not copied: on another sheet its copy could not be kept on the cells the chart
 * shows, and it would follow no rows, columns or sheets.
 */
export function chartCopyProblem(workbook: Workbook, sheetIndex: number, id: number): string | undefined {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet) return `No sheet at index ${sheetIndex}`;
  if (sheet.kind !== 'worksheet') return `Cannot copy the chart of a ${sheet.kind}`;
  const drawing = sheet.drawings && Math.floor(id) === id ? sheet.drawings[id] : undefined;
  if (!drawing || !drawing.object || drawing.object.type !== 'chart') return `No chart at index ${id}`;
  const object = drawing.object;
  // A chart of an xls file is shown from what was read of it; nothing writes one.
  if (object.chart && object.chart.shownOnly === true) return object.chart.extended === true ? 'Cannot copy this chart: a chart of this type is only shown' : 'Cannot copy this chart: it was read from an xls file and is only shown';
  if (object.part !== undefined && object.known === undefined && !(object.loose && object.loose.length > 0)) return 'Cannot copy this chart: its part holds references the model does not, so a copy could not be kept on the cells it shows';
  if (!partBound(object) && !chartWritable(object.chart)) return 'Cannot copy this chart: its type is not one that is written from the model';
  return undefined;
}

/**
 * Takes a copy of a chart; nothing of the workbook is changed. The copy shows the cells the
 * chart shows. A chart of the file whose part holds more than its model (formatting, extensions,
 * styles, colours, pictures of fills) is copied with all of that: the part of the copy is made
 * from the bytes of that part when the workbook is written. What no copy takes along is named in
 * `dropped`. Throws RangeError for a chart that cannot be copied: chartCopyProblem() says why.
 */
export function copyChart(workbook: Workbook, sheetIndex: number, id: number): ChartCopy {
  const problem = chartCopyProblem(workbook, sheetIndex, id);
  if (problem) throw new RangeError(problem);
  const sheet = workbook.sheets[sheetIndex];
  const { drawing, object } = chartAt(sheet, id);
  // The references as the sheets are called now: a copy is a chart that was added.
  let chart = object.chart;
  if (!chart.live) {
    const names = sheetNameMap(workbook);
    if (names) chart = mapChartRefs(chart, (ref) => (ref.indexOf('!') < 0 || formulaNamesSheetsUnread(ref) ? ref : renameSheetsInFormula(ref, names)));
    chart = { ...chart, live: true };
  }
  const made: DrawingChart = { type: 'chart', chart };
  if (object.description !== undefined) made.description = object.description;
  if (object.hidden) made.hidden = true;
  // The frame of a chart of the file is what the frame of its copy is made from.
  if (object.part !== undefined) made.frame = object.part;
  else if (object.frame !== undefined) made.frame = object.frame;
  const dropped: string[] = [];
  if (partBound(object) && object.known !== 'all') {
    made.known = 'refs';
    if (object.seriesFixed) made.seriesFixed = true;
    made.origin = originOf(object);
    if (object.source) made.source = object.source.drops ? { part: object.source.part, drops: object.source.drops.slice() } : { part: object.source.part };
    else {
      made.source = { part: object.part as string };
      if (object.loose && object.loose.length > 0) {
        made.source.drops = object.loose.slice();
        for (const name of object.loose) dropped.push(name);
      }
    }
  }
  if (object.part !== undefined) {
    const cost = chartPartCost(originOf(object).chart);
    if (cost && cost.stray) dropped.push('relationships');
    if (object.macro !== undefined && object.macro !== '') dropped.push('macro');
  }
  // The sheets the copy stands on: those its references name.
  const named = new Set<Sheet>();
  let strict = false;
  for (const ref of chartRefs(chart)) {
    if (ref.indexOf('!') < 0) {
      // Cells of the sheet the chart is on; a name is no cells, and moves with no row.
      if (formulaReferences(ref).length > 0) named.add(sheet);
      continue;
    }
    const areas = chartRefAreas(workbook, ref, true);
    if (!areas) strict = true;
    else for (const area of areas) named.add(area.sheet);
  }
  const anchor = drawingAnchorNow(drawing, sheet.structuralEdits);
  const box = anchorRect(anchor, createSheetMetrics(sheet));
  const copy: ChartCopy = {
    workbook,
    sheet,
    drawing,
    object: made,
    rect: { x: box.x, y: box.y, width: box.width, height: box.height },
    dropped,
    state: workbook.sheets.filter((other) => strict || named.has(other)).map((other) => ({ sheet: other, name: other.name, edits: (other.structuralEdits || []).slice() })),
  };
  const more: { editAs?: DrawingAnchor['editAs']; strict?: boolean } = {};
  if (anchor.kind === 'twoCell' && anchor.editAs !== undefined) more.editAs = anchor.editAs;
  if (strict) more.strict = true;
  return { ...copy, ...more };
}

/**
 * Whether a copy can no longer be put down: it is of another workbook, or a sheet its
 * references name is gone or has another name, or rows or columns were inserted or deleted on
 * such a sheet since the copy was taken so that its references would name other cells than the
 * chart it was taken from shows now. Sheets the chart does not name are nothing to it; neither
 * are rows and columns that leave its references as they are (below its cells, or put in and
 * taken out again), nor steps that were undone.
 */
export function chartCopyOutdated(workbook: Workbook, copy: ChartCopy): boolean {
  if (!copy || copy.workbook !== workbook || !Array.isArray(copy.state)) return true;
  const sameEdit = (a: SheetStructuralEdit, b: SheetStructuralEdit): boolean => a === b || (a.kind === b.kind && a.axis === b.axis && a.at === b.at && a.count === b.count);
  const since: StructuralEdit[] = [];
  for (const was of copy.state) {
    if (workbook.sheets.indexOf(was.sheet) < 0 || was.sheet.name !== was.name) return true;
    const now = was.sheet.structuralEdits || [];
    // What was done when the copy was taken has to be what was done first: anything else is another sheet.
    if (now.length < was.edits.length) return true;
    for (let i = 0; i < was.edits.length; i++) if (!sameEdit(now[i], was.edits[i])) return true;
    if (now.length > was.edits.length && copy.strict) return true;
    for (let i = was.edits.length; i < now.length; i++) since.push({ kind: now[i].kind, axis: now[i].axis, at: now[i].at, count: now[i].count, sheetName: was.name });
  }
  if (since.length === 0) return false;
  // What the chart it was taken from says now: each reference moved the way the charts of the workbook are moved.
  for (const ref of chartRefs(copy.object.chart)) {
    let text = ref;
    for (const edit of since) text = adjustFormula(text, copy.sheet.name, edit);
    if (text !== ref) return true;
  }
  return false;
}

/**
 * A reference with the sheet it was written on named in every area that names none
 * ("$A$1:$A$5" becomes "Sheet1!$A$1:$A$5"). `bound` is false for a reference that names no
 * sheet and holds no cells to bind (a defined name): it is left as it is.
 */
function boundRef(ref: string, sheetName: string): { ref: string; bound: boolean } {
  const found = formulaReferences(ref);
  if (found.length === 0) return { ref, bound: ref.indexOf('!') >= 0 };
  const prefix = `${quoteSheetName(sheetName)}!`;
  let out = '';
  let at = 0;
  for (const item of found) {
    if (item.bodyStart !== item.start) continue;
    out += ref.slice(at, item.start) + prefix;
    at = item.start;
  }
  return { ref: out + ref.slice(at), bound: true };
}

/**
 * Puts a copy of a chart on a sheet, above its other drawings: on the sheet it was taken from
 * or on another one of the same workbook. With `at` its top left corner is the corner of that
 * cell; without, it stands where the chart stood, moved right and down (CHART_COPY_OFFSET) for as
 * long as a chart stands there already. Its size is the size the chart had.
 *
 * The copy shows the cells the chart showed: on another sheet a reference that names no sheet
 * gets the name of the sheet the chart was taken from. From then on it follows rows, columns and
 * sheets like any chart that was added.
 *
 * Throws RangeError for a copy of another workbook, one that is outdated (chartCopyOutdated()),
 * and a sheet that takes no more charts.
 *
 * While the chart the copy was taken from is still the drawing it was (it was neither changed,
 * moved nor deleted since), the paste is "a copy of that chart" and its operation says so (the
 * command "pasteChart": see pasteChartOf()): that is what travels when a workbook is edited
 * together, where every participant makes the copy from the file they read themselves. A copy
 * of a chart that is no longer what it was (a chart that was cut) is pasted from what is held,
 * and its operation names no command.
 */
export function pasteChart(workbook: Workbook, sheetIndex: number, copy: ChartCopy, at?: { row: number; col: number }): PastedChart {
  placedSheetOf(workbook, sheetIndex);
  if (!copy || typeof copy !== 'object' || !copy.object || copy.object.type !== 'chart') throw new RangeError('What is pasted is not a chart that was copied');
  if (copy.workbook !== workbook) throw new RangeError('A chart is pasted into the workbook it was copied from');
  if (chartCopyOutdated(workbook, copy)) throw new RangeError('Cannot paste the chart: rows, columns or sheets changed since it was copied');
  const home = workbook.sheets.indexOf(copy.sheet);
  const from = home >= 0 && copy.drawing ? (copy.sheet.drawings || []).indexOf(copy.drawing) : -1;
  if (from < 0) return pasteHeld(workbook, sheetIndex, copy, at);
  // The chart as it is now: what it says is what the copy says (the copy is not outdated), and where it stands is where it stands now.
  const pasted = pasteHeld(workbook, sheetIndex, copyChart(workbook, home, from), at);
  withCommand(pasted.ops, 'pasteChart', [copy.sheet.name, from, at === undefined ? null : { row: at.row, col: at.col }]);
  return pasted;
}

/**
 * The operations of the command "pasteChart", as it travels: a copy of chart `id` of the sheet
 * called `sheetName` is put on sheet `sheetIndex` (at a cell, or where pasteChart() puts it).
 * The sheet is named, not counted: the name says which sheet is meant whatever sheets were put
 * in or moved since. Throws RangeError for a sheet or a chart that is not there, and for
 * whatever pasteChart() and copyChart() refuse.
 */
export function pasteChartOf(workbook: Workbook, sheetIndex: number, sheetName: string, id: number, at?: { row: number; col: number } | null): Op[] {
  if (typeof sheetName !== 'string' || sheetName === '' || sheetName.length > 255) throw new RangeError('The chart that is copied is named by its sheet and its place');
  if (typeof id !== 'number' || !(id >= 0) || Math.floor(id) !== id) throw new RangeError(`No chart at index ${String(id).slice(0, 40)}`);
  if (at !== undefined && at !== null && (typeof at !== 'object' || Array.isArray(at))) throw new RangeError('The place of a chart is a cell of the sheet');
  const wanted = foldSheetName(sheetName);
  let home = -1;
  workbook.sheets.forEach((sheet, index) => {
    if (home < 0 && foldSheetName(sheet.name) === wanted) home = index;
  });
  if (home < 0) throw new RangeError(`No sheet called "${sheetName.slice(0, 80)}"`);
  return pasteChart(workbook, sheetIndex, copyChart(workbook, home, id), at === null || at === undefined ? undefined : { row: at.row, col: at.col }).ops;
}

function pasteHeld(workbook: Workbook, sheetIndex: number, copy: ChartCopy, at: { row: number; col: number } | undefined): PastedChart {
  const sheet = placedSheetOf(workbook, sheetIndex);
  checkRoom(workbook, sheet);
  let unbound = 0;
  let chart = copy.object.chart;
  if (sheet !== copy.sheet) {
    chart = mapChartRefs(chart, (ref) => {
      const made = boundRef(ref, copy.sheet.name);
      if (!made.bound) unbound++;
      return made.ref;
    });
  }
  // A chart of its own, whatever else was made of the same copy.
  if (chart === copy.object.chart) chart = { ...chart };
  const metrics = createSheetMetrics(sheet);
  let x = copy.rect.x;
  let y = copy.rect.y;
  if (at !== undefined) {
    const whole = (n: unknown, limit: number) => typeof n === 'number' && n >= 0 && n < limit && Math.floor(n) === n;
    if (!at || !whole(at.row, MAX_ROWS) || !whole(at.col, MAX_COLS)) throw new RangeError('The place of a chart is a cell of the sheet');
    x = metrics.cols.offset(at.col);
    y = metrics.rows.offset(at.row);
  } else {
    const taken: { x: number; y: number }[] = [];
    for (const drawing of sheet.drawings || []) {
      if (!drawing || !drawing.object || drawing.object.type !== 'chart') continue;
      const box = anchorRect(drawingAnchorNow(drawing, sheet.structuralEdits), metrics);
      taken.push({ x: box.x, y: box.y });
    }
    for (let n = 0; n < 64 && taken.some((box) => Math.abs(box.x - x) < 1 && Math.abs(box.y - y) < 1); n++) {
      x += CHART_COPY_OFFSET;
      y += CHART_COPY_OFFSET;
    }
  }
  const anchor = anchorOfRect(metrics, x, y, copy.rect.width, copy.rect.height);
  if (copy.editAs === 'oneCell' || copy.editAs === 'absolute') anchor.editAs = copy.editAs;
  const object: DrawingChart = { ...copy.object, chart, name: nextName(sheet) };
  if (copy.object.source) object.source = copy.object.source.drops ? { part: copy.object.source.part, drops: copy.object.source.drops.slice() } : { part: copy.object.source.part };
  checkBudget(workbook, object);
  const drawings = sheet.drawings || [];
  const added: SheetDrawing = { anchor, object, since: sheet.structuralEdits ? sheet.structuralEdits.length : 0 };
  return { ops: setDrawings(sheetIndex, drawings.concat([added])), id: drawings.length, dropped: copy.dropped.slice(), unbound };
}

/** Puts a copy of a chart beside it on its sheet: copyChart() and pasteChart() in one. */
export function duplicateChart(workbook: Workbook, sheetIndex: number, id: number): PastedChart {
  return pasteChart(workbook, sheetIndex, copyChart(workbook, sheetIndex, id));
}

/**
 * Takes a copy of a chart and deletes the chart: the operations are one undo step, and the copy
 * is pasted like any other. Throws RangeError for a chart that cannot be copied or cannot be
 * deleted; nothing is taken then.
 */
export function cutChart(workbook: Workbook, sheetIndex: number, id: number): { copy: ChartCopy; ops: Op[] } {
  const copy = copyChart(workbook, sheetIndex, id);
  return { copy, ops: removeChart(workbook, sheetIndex, id) };
}

// ------------------------------------------------------------------ rows, columns and sheets

/** Whether the charts of a workbook follow rows and columns that move and sheets that are renamed: see Workbook.chartsFollow. */
export function chartsFollow(workbook: Workbook): boolean {
  return workbook.chartsFollow === true;
}

/** The lists of drawings of the sheets with every chart put through `change` (undefined: left as it is); only the lists that came out different. */
function mapCharts(workbook: Workbook, skip: Sheet | undefined, change: (chart: Chart, sheet: Sheet) => Chart): RuleLists[] {
  const out: RuleLists[] = [];
  for (const sheet of workbook.sheets) {
    if (sheet === skip || !sheet.drawings || sheet.kind !== 'worksheet') continue;
    let list: SheetDrawing[] | undefined;
    sheet.drawings.forEach((drawing, i) => {
      if (!drawing || !drawing.object || drawing.object.type !== 'chart') return;
      const object = drawing.object;
      // A chart of the file whose part cannot be changed is left as it is: a workbook with one keeps its rows, columns and sheet names.
      if (object.part !== undefined && object.known === undefined) return;
      const chart = change(object.chart, sheet);
      if (chart === object.chart) return;
      const next: DrawingChart = { ...object, chart: chart.live ? chart : { ...chart, live: true } };
      // What the writer changes in the part of a chart of the file is the difference to what was read.
      if (partBound(object) && object.known !== 'all') next.origin = originOf(object);
      if (!list) list = sheet.drawings ? sheet.drawings.slice() : [];
      // Its place is not touched: the anchor is as it was, and so is the count of edits it was put after.
      const moved: SheetDrawing = { ...drawing, object: next };
      const from = fileDrawing(drawing);
      if (from) moved.from = from;
      list[i] = moved;
    });
    if (list) out.push({ sheet, drawings: { list } });
  }
  return out;
}

/**
 * What rows or columns inserted into or deleted from a sheet do to the charts of the workbook:
 * the references of their series follow the cells they name ("#REF!" for cells that are gone).
 * `edit` names the sheet as it is called now. Nothing is changed here: the lists that come out
 * different are returned, for swapRuleLists(). (Where a chart stands follows by itself: its
 * anchor is read through the structural edits of its sheet.)
 */
export function chartsMoved(workbook: Workbook, edit: StructuralEdit): RuleLists[] {
  return mapCharts(workbook, undefined, (chart, sheet) => mapChartRefs(chart, (ref) => adjustFormula(ref, sheet.name, edit)));
}

/**
 * What a change of sheet names does to the charts of the workbook (see renameSheetsInFormula();
 * a sheet that is gone leaves "#REF!"). `skip` is a sheet whose charts do not matter (it is
 * being deleted). Throws RangeError when a reference that may name a sheet cannot be read to its end.
 */
export function chartsRenamed(workbook: Workbook, map: SheetNameMap, skip?: Sheet): RuleLists[] {
  return mapCharts(workbook, skip, (chart) =>
    mapChartRefs(chart, (ref) => {
      if (ref.indexOf('!') < 0) return ref;
      if (formulaNamesSheetsUnread(ref)) throw new RangeError('Cannot change the sheet: a reference of a chart cannot be read');
      return renameSheetsInFormula(ref, map);
    }),
  );
}
