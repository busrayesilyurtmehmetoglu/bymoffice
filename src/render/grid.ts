// Draws a sheet on a Canvas 2D context: fills, grid lines, borders, text,
// merged cells, the selection, the headers and the frozen pane lines.
//
// Only what is visible is touched. Every loop runs over the row and column
// ranges of gridPanes(), never over the 1,048,576 x 16,384 extent of a sheet.
// Nothing that comes from a file reaches `ctx.font`, `fillStyle` or
// `strokeStyle` unchecked: colours go through resolveColor() ("#RRGGBB") and
// font names through sanitizeFontFamily().

import { gridPanes } from '../layout/grid-geometry';
import type { GridLayout, GridPane } from '../layout/grid-geometry';
import type { AxisMetrics } from '../layout/sheet-metrics';
import { resolveCellStyle, resolveColor } from '../model/sheet/color';
import { columnName, MAX_COLS, MAX_ROWS } from '../model/sheet/types';
import type { BorderEdge, Cell, CellRange, Fill, Font, Row, Sheet, StyleSheet, Workbook } from '../model/sheet/types';
import { builtinFormatId, formatCell } from '../format/cell-format';
import { iconGlyph, type CellConditional, type CellDataBar, type ConditionalSource, type ConditionalStyle, type IconGlyph } from '../format/conditional';
import type { FormatLocale } from '../format/locale';
import type { FormatOptions } from '../format/number-format';
import { cssFont, FONT_FALLBACK, fontPixels, TextPainter, wrapLines } from './text';

export interface GridTheme {
  background: string;
  gridLine: string;
  headerBackground: string;
  headerText: string;
  headerBorder: string;
  headerSelectedBackground: string;
  selectionBorder: string;
  selectionFill: string;
  frozenLine: string;
  defaultText: string;
  /** Row numbers of the rows a filter with criteria governs. Blue when missing. */
  filterHeaderText?: string;
}

export const DEFAULT_GRID_THEME: GridTheme = {
  background: '#FFFFFF',
  gridLine: '#E1E1E1',
  headerBackground: '#F3F3F3',
  headerText: '#444444',
  headerBorder: '#C6C6C6',
  headerSelectedBackground: '#D2E3D8',
  selectionBorder: '#1A7F45',
  selectionFill: 'rgba(26, 127, 69, 0.12)',
  frozenLine: '#9E9E9E',
  defaultText: '#000000',
};

export interface GridSelection {
  active: { row: number; col: number };
  range: CellRange;
}

/** The filter of a sheet as the grid draws it: a button in each header cell, and the rows it governs. */
export interface GridFilter {
  /** The header row is the first row of the range. */
  range: CellRange;
  /** By column, counted from the left edge of the range: 0 no button, 1 a button, 2 the button of a column with a criterion. */
  buttons: ArrayLike<number>;
  /** Some criterion is set: the numbers of the rows under the header are drawn in the filter colour. */
  filtered: boolean;
}

/** Largest size of the button a filter or a list draws inside a cell, in pixels. */
export const CELL_BUTTON_SIZE = 16;
const MIN_CELL_BUTTON = 8;

/**
 * The square a cell's button (filter, list) takes: the bottom right corner of the cell, inside
 * its grid lines. Undefined for a cell too small to hold one. The view uses the same function
 * to tell a press on the button from a press on the cell.
 */
export function cellButtonRect(x: number, y: number, width: number, height: number): { x: number; y: number; size: number } | undefined {
  const size = Math.min(CELL_BUTTON_SIZE, Math.floor(width) - 2, Math.floor(height) - 2);
  if (!(size >= MIN_CELL_BUTTON)) return undefined;
  return { x: x + width - size - 1, y: y + height - size - 1, size };
}

export interface GridScene {
  workbook: Workbook;
  sheet: Sheet;
  layout: GridLayout;
  selection?: GridSelection;
  theme?: GridTheme;
  /** Decimal separators, month names and date order for number formats; en-US when missing. */
  locale?: FormatLocale;
  /** Canvas backing store pixels per CSS pixel. */
  devicePixelRatio: number;
  /** Conditional formats of the sheet; nothing is asked of it for a sheet that has none. */
  conditional?: ConditionalSource;
  /** The filter of the sheet. */
  filter?: GridFilter;
  /** The cell that shows the button of a list to choose from (a data validation). */
  dropdown?: { row: number; col: number };
  /**
   * A mark in the top left corner of a cell, as Excel marks cells it has something to say
   * about: the colour of the mark ("#RRGGBB"), or undefined for no mark. Asked for every
   * visible cell of every frame, so it has to be a plain lookup.
   */
  cellMark?(row: number, col: number): string | undefined;
}

/** Space between a cell edge and its text, in pixels. */
const CELL_PADDING = 3;
/** Width of one indent level, in pixels (three spaces of the default font). */
const INDENT_WIDTH = 9;
const MAX_INDENT = 250;
/** Line height and the part of it below the baseline, as multiples of the font size. */
const LINE_HEIGHT = 1.2;
const DESCENT = 0.25;
/** Longest text measured for a cell that does not wrap, and for one that does (Excel's cell limit). */
const MAX_PLAIN_CHARS = 2000;
const MAX_WRAP_CHARS = 32767;
const MAX_RUNS = 4096;
/** How many columns outside the visible ones are searched for text that overflows into view. */
const OVERFLOW_SCAN = 128;
const MAX_HASHES = 512;
const HEADER_FONT = '12px ' + FONT_FALLBACK;
/** Headers narrower or lower than this get no label. */
const MIN_COL_LABEL_WIDTH = 12;
const MIN_ROW_LABEL_HEIGHT = 8;
/** Space between the edge of a cell and its data bar, and between an icon and the text after it. */
const BAR_INSET = 2;
const ICON_GAP = 3;
const MIN_ICON = 6;
const MAX_ICON = 32;
/** How far the end of a gradient data bar is faded towards white. */
const BAR_FADE = 0.85;
const AXIS_DASH = [1, 1];
/** Parts of an icon that are "off": the empty bars of a rating, the outline of a quarter circle. */
const ICON_OFF = '#C8C8C8';
const ICON_DARK = '#505050';
const ICON_LIGHT = '#FFFFFF';
const HEX_COLOR = /^#[0-9A-Fa-f]{6}$/;

interface LineStyle {
  name: string;
  width: number;
  dash: number[];
  double: boolean;
  alpha: number;
  /** Which of two edges on the same line is drawn: the higher rank. */
  rank: number;
}

const NO_DASH: number[] = [];
const line = (name: string, width: number, rank: number, dash: number[] = NO_DASH, double = false, alpha = 1): LineStyle => ({
  name,
  width,
  dash,
  double,
  alpha,
  rank,
});

/** ST_BorderStyle -> how it is stroked. A Map, so that names from a file cannot reach Object.prototype. */
const LINE_STYLES = new Map<string, LineStyle>(
  [
    line('hair', 1, 1, NO_DASH, false, 0.5),
    line('dotted', 1, 2, [1, 1]),
    line('dashDotDot', 1, 3, [6, 2, 2, 2, 2, 2]),
    line('dashDot', 1, 4, [6, 2, 2, 2]),
    line('dashed', 1, 5, [3, 2]),
    line('thin', 1, 6),
    line('mediumDashDotDot', 2, 7, [8, 3, 3, 3, 3, 3]),
    line('slantDashDot', 2, 8, [8, 3, 3, 3]),
    line('mediumDashDot', 2, 9, [8, 3, 3, 3]),
    line('mediumDashed', 2, 10, [6, 3]),
    line('medium', 2, 11),
    line('double', 1, 12, NO_DASH, true),
    line('thick', 3, 13),
  ].map((style): [string, LineStyle] => [style.name, style]),
);
const THIN = LINE_STYLES.get('thin') as LineStyle;

interface EdgeStyle {
  line: LineStyle;
  color: string;
  /** Edges with the same key are stroked in one path. */
  key: string;
}

interface TextStyle {
  font: string;
  px: number;
  color: string;
  /** Number of underline strokes: 0, 1 or 2. */
  underline: number;
  strike: boolean;
}

type HorizontalAlign = 'general' | 'left' | 'center' | 'right';
type VerticalAlign = 'top' | 'center' | 'bottom';

/** A cell format with everything a frame needs already looked up. */
interface DrawStyle {
  fill: string | undefined;
  text: TextStyle;
  horizontal: HorizontalAlign;
  vertical: VerticalAlign;
  wrap: boolean;
  indent: number;
  left: EdgeStyle | undefined;
  right: EdgeStyle | undefined;
  top: EdgeStyle | undefined;
  bottom: EdgeStyle | undefined;
  hasBorder: boolean;
  /** Number format code, and the options to apply it with. */
  numFmt: string;
  format: FormatOptions;
}

/** A piece of cell text with one font; a plain cell has one, rich text has one per run. */
interface Segment {
  text: string;
  style: TextStyle;
}

/** The visible (not hidden) items of one axis of a pane, with their pixel positions. */
interface VisibleAxis {
  index: number[];
  pos: number[];
  size: number[];
  /** Pixel position of item 0 of the axis in this pane (scrolling applied). */
  origin: number;
}

interface PaneMerge {
  range: CellRange;
  x: number;
  y: number;
  width: number;
  height: number;
}

interface PaneView {
  pane: GridPane;
  cols: VisibleAxis;
  rows: VisibleAxis;
  merges: PaneMerge[];
  /** What conditional formats give the visible cells, by visible position; only present when a cell has something. */
  conditional?: (CellConditional | undefined)[];
}

/** State of one renderGrid() call. */
interface Frame {
  ctx: CanvasRenderingContext2D;
  painter: TextPainter;
  sheet: Sheet;
  styles: StyleSheet;
  layout: GridLayout;
  theme: GridTheme;
  selection: GridSelection | undefined;
  /** Merged ranges that intersect a pane, clamped to the sheet. */
  merges: CellRange[];
  styleCache: Map<number, DrawStyle>;
  textStyleCache: Map<Font, TextStyle>;
  format: FormatOptions;
  /** Text styles recoloured by a number format section such as [Red], keyed by font and colour. */
  colorStyleCache: Map<string, TextStyle>;
  conditional: ConditionalSource | undefined;
  /** Cell formats with a conditional format on top, by the conditional format and the index of the cell format. */
  conditionalStyles: Map<ConditionalStyle, Map<number, DrawStyle>>;
  filter: GridFilter | undefined;
  dropdown: { row: number; col: number } | undefined;
  cellMark: ((row: number, col: number) => string | undefined) | undefined;
}

/** Draws the whole scene. The canvas backing store is layout.width*dpr by layout.height*dpr. */
export function renderGrid(ctx: CanvasRenderingContext2D, scene: GridScene): void {
  const dpr = scene.devicePixelRatio > 0 && Number.isFinite(scene.devicePixelRatio) ? scene.devicePixelRatio : 1;
  const { layout, sheet } = scene;
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  if (!(layout.width > 0) || !(layout.height > 0)) return;

  const theme = scene.theme || DEFAULT_GRID_THEME;
  ctx.globalAlpha = 1;
  ctx.lineWidth = 1;
  ctx.setLineDash(NO_DASH);
  ctx.textAlign = 'left';
  ctx.textBaseline = 'alphabetic';
  ctx.fillStyle = theme.background;
  ctx.fillRect(0, 0, layout.width, layout.height);

  const panes = gridPanes(layout);
  const frame: Frame = {
    ctx,
    painter: new TextPainter(ctx),
    sheet,
    styles: scene.workbook.styles,
    layout,
    theme,
    selection: validSelection(scene.selection),
    merges: visibleMerges(sheet, panes),
    styleCache: new Map(),
    textStyleCache: new Map(),
    format: { locale: scene.locale, date1904: scene.workbook.date1904 === true },
    colorStyleCache: new Map(),
    conditional: scene.conditional && typeof scene.conditional.resolve === 'function' ? scene.conditional : undefined,
    conditionalStyles: new Map(),
    filter: validFilter(scene.filter),
    dropdown: scene.dropdown && scene.dropdown.row >= 0 && scene.dropdown.col >= 0 ? scene.dropdown : undefined,
    cellMark: typeof scene.cellMark === 'function' ? scene.cellMark : undefined,
  };
  if (frame.conditional && frame.conditional.begin) frame.conditional.begin();

  const views: PaneView[] = [];
  for (const pane of panes) {
    const view = paneView(frame, pane);
    if (!view) continue;
    views.push(view);
    drawPane(frame, view);
  }
  drawHeaders(frame, views);
  drawFrozenLines(frame);
}

// ---------------------------------------------------------------------------
// Geometry of a frame

function validSelection(selection: GridSelection | undefined): GridSelection | undefined {
  if (!selection || !selection.range || !selection.active) return undefined;
  const range = clampRange(selection.range);
  if (!range) return undefined;
  const { row, col } = selection.active;
  if (!(row >= 0 && row < MAX_ROWS && col >= 0 && col < MAX_COLS)) return undefined;
  return { active: { row: Math.floor(row), col: Math.floor(col) }, range };
}

function validFilter(filter: GridFilter | undefined): GridFilter | undefined {
  if (!filter || !filter.range || !filter.buttons) return undefined;
  const range = clampRange(filter.range);
  return range === filter.range ? filter : undefined;
}

/** The range cut to the sheet; undefined when it is not a range (NaN, inverted, outside). */
function clampRange(range: CellRange): CellRange | undefined {
  const top = Math.max(Math.floor(range.top), 0);
  const left = Math.max(Math.floor(range.left), 0);
  const bottom = Math.min(Math.floor(range.bottom), MAX_ROWS - 1);
  const right = Math.min(Math.floor(range.right), MAX_COLS - 1);
  if (!(top <= bottom && left <= right)) return undefined;
  return top === range.top && left === range.left && bottom === range.bottom && right === range.right ? range : { top, left, bottom, right };
}

/** One pass over the merge list per frame: the ranges that reach into a pane. */
function visibleMerges(sheet: Sheet, panes: GridPane[]): CellRange[] {
  const found: CellRange[] = [];
  const merges = sheet.merges;
  if (!merges || merges.length === 0 || panes.length === 0) return found;
  for (let i = 0; i < merges.length; i++) {
    const merge = merges[i];
    if (!merge) continue;
    for (let p = 0; p < panes.length; p++) {
      const pane = panes[p];
      if (merge.bottom < pane.firstRow || merge.top > pane.lastRow || merge.right < pane.firstCol || merge.left > pane.lastCol) continue;
      const range = clampRange(merge);
      if (range && (range.bottom > range.top || range.right > range.left)) found.push(range);
      break;
    }
  }
  return found;
}

function visibleAxis(metrics: AxisMetrics, first: number, last: number, origin: number): VisibleAxis {
  const axis: VisibleAxis = { index: [], pos: [], size: [], origin };
  const from = Math.max(first, 0);
  const to = Math.min(last, metrics.count - 1);
  if (!(from <= to)) return axis;
  let pos = Math.round(origin + metrics.offset(from));
  // Runs of hidden items are jumped over, so the work is bounded by what is visible.
  for (let i = metrics.skipHidden(from, 1); i <= to; i = metrics.skipHidden(i + 1, 1)) {
    const size = metrics.size(i);
    axis.index.push(i);
    axis.pos.push(pos);
    axis.size.push(size);
    pos += size;
  }
  return axis;
}

/** Position in a sorted list of the first value that is >= `value`. */
function lowerBound(list: number[], value: number): number {
  let low = 0;
  let high = list.length;
  while (low < high) {
    const mid = (low + high) >> 1;
    if (list[mid] < value) low = mid + 1;
    else high = mid;
  }
  return low;
}

function paneView(frame: Frame, pane: GridPane): PaneView | undefined {
  const { layout } = frame;
  const metrics = layout.metrics;
  const cols = visibleAxis(metrics.cols, pane.firstCol, pane.lastCol, metrics.headerWidth - (pane.frozenCols ? 0 : layout.scrollX));
  const rows = visibleAxis(metrics.rows, pane.firstRow, pane.lastRow, metrics.headerHeight - (pane.frozenRows ? 0 : layout.scrollY));
  if (cols.index.length === 0 || rows.index.length === 0) return undefined;

  // A merge is laid out in the coordinates of the pane that draws it, so a
  // merge across the frozen boundary is drawn by each pane under its own clip.
  const merges: PaneMerge[] = [];
  for (const range of frame.merges) {
    if (range.bottom < pane.firstRow || range.top > pane.lastRow || range.right < pane.firstCol || range.left > pane.lastCol) continue;
    const x = Math.round(cols.origin + metrics.cols.offset(range.left));
    const y = Math.round(rows.origin + metrics.rows.offset(range.top));
    const width = metrics.cols.offset(range.right + 1) - metrics.cols.offset(range.left);
    const height = metrics.rows.offset(range.bottom + 1) - metrics.rows.offset(range.top);
    if (width > 0 && height > 0) merges.push({ range, x, y, width, height });
  }
  return { pane, cols, rows, merges };
}

/** 1-based position in `merges` of the merge that covers a cell; 0 when there is none. */
function mergeIndexAt(merges: PaneMerge[], row: number, col: number): number {
  for (let i = 0; i < merges.length; i++) {
    const range = merges[i].range;
    if (row >= range.top && row <= range.bottom && col >= range.left && col <= range.right) return i + 1;
  }
  return 0;
}

/** Marks the visible cells of each merge with its 1-based position; undefined when the pane has no merge. */
function markMerges(view: PaneView): Int32Array | undefined {
  const { cols, rows, merges } = view;
  if (merges.length === 0) return undefined;
  const colCount = cols.index.length;
  const marks = new Int32Array(rows.index.length * colCount);
  for (let i = 0; i < merges.length; i++) {
    const range = merges[i].range;
    const firstRow = lowerBound(rows.index, range.top);
    const endRow = lowerBound(rows.index, range.bottom + 1);
    const firstCol = lowerBound(cols.index, range.left);
    const endCol = lowerBound(cols.index, range.right + 1);
    for (let ri = firstRow; ri < endRow; ri++) {
      for (let ci = firstCol; ci < endCol; ci++) {
        // Overlapping merges are not valid; the first one keeps the cell.
        if (marks[ri * colCount + ci] === 0) marks[ri * colCount + ci] = i + 1;
      }
    }
  }
  return marks;
}

/** Column-level style index for each visible column; 0 where a column has none. */
function columnStyles(sheet: Sheet, cols: VisibleAxis): number[] {
  const styles: number[] = [];
  const ranges = sheet.columns;
  for (let ci = 0; ci < cols.index.length; ci++) {
    const col = cols.index[ci];
    let style = 0;
    // `columns` is sorted by `min`: find the last range that starts at or before the column.
    let low = 0;
    let high = ranges.length;
    while (low < high) {
      const mid = (low + high) >> 1;
      if (ranges[mid].min <= col) low = mid + 1;
      else high = mid;
    }
    const range = low > 0 ? ranges[low - 1] : undefined;
    if (range && range.max >= col && range.style !== undefined) style = range.style;
    styles.push(style);
  }
  return styles;
}

// ---------------------------------------------------------------------------
// Styles

function fillColor(fill: Fill, styles: StyleSheet): string | undefined {
  // Patterns other than "solid" are drawn in their foreground colour for now; gradients are skipped.
  if (fill.type !== 'pattern' || fill.pattern === 'none') return undefined;
  return resolveColor(fill.fg, styles);
}

function edgeStyle(edge: BorderEdge | undefined, styles: StyleSheet): EdgeStyle | undefined {
  if (!edge || !edge.style || edge.style === 'none') return undefined;
  const lineStyle = LINE_STYLES.get(edge.style) || THIN;
  const color = resolveColor(edge.color, styles) || '#000000';
  return { line: lineStyle, color, key: lineStyle.name + color };
}

function textStyle(frame: Frame, font: Font): TextStyle {
  let style = frame.textStyleCache.get(font);
  if (!style) {
    const underline = font.underline === 'double' || font.underline === 'doubleAccounting' ? 2 : font.underline === 'single' || font.underline === 'singleAccounting' ? 1 : 0;
    style = {
      font: cssFont(font),
      px: fontPixels(font),
      color: resolveColor(font.color, frame.styles) || frame.theme.defaultText,
      underline,
      strike: font.strike === true,
    };
    frame.textStyleCache.set(font, style);
  }
  return style;
}

function drawStyle(frame: Frame, index: number): DrawStyle {
  let style = frame.styleCache.get(index);
  if (style) return style;
  const resolved = resolveCellStyle(frame.styles, index);
  const { alignment, border } = resolved;
  const h = alignment.horizontal;
  const v = alignment.vertical;
  const indent = alignment.indent;
  const left = edgeStyle(border.left, frame.styles);
  const right = edgeStyle(border.right, frame.styles);
  const top = edgeStyle(border.top, frame.styles);
  const bottom = edgeStyle(border.bottom, frame.styles);
  style = {
    fill: fillColor(resolved.fill, frame.styles),
    text: textStyle(frame, resolved.font),
    // "fill", "justify" and "distributed" are not implemented; they read as left.
    horizontal: h === 'center' || h === 'centerContinuous' ? 'center' : h === 'right' ? 'right' : h === undefined || h === 'general' ? 'general' : 'left',
    vertical: v === 'top' || v === 'justify' ? 'top' : v === 'center' || v === 'distributed' ? 'center' : 'bottom',
    wrap: alignment.wrapText === true,
    indent: typeof indent === 'number' && indent > 0 ? Math.min(Math.floor(indent), MAX_INDENT) : 0,
    left,
    right,
    top,
    bottom,
    hasBorder: !!(left || right || top || bottom),
    numFmt: resolved.numFmt,
    format: { locale: frame.format.locale, date1904: frame.format.date1904, builtinId: builtinFormatId(frame.styles, resolved.numFmtId) },
  };
  frame.styleCache.set(index, style);
  return style;
}

/** A colour handed in from outside the model: only "#RRGGBB" reaches the canvas. */
function safeColor(color: string | null | undefined): string | undefined {
  return typeof color === 'string' && HEX_COLOR.test(color) ? color : undefined;
}

/** The format of a cell with what its conditional formats state on top. */
function conditionalStyle(frame: Frame, index: number, format: ConditionalStyle): DrawStyle {
  let byIndex = frame.conditionalStyles.get(format);
  if (!byIndex) {
    byIndex = new Map();
    frame.conditionalStyles.set(format, byIndex);
  }
  let style = byIndex.get(index);
  if (style) return style;
  const base = drawStyle(frame, index);
  style = { ...base };
  if (format.fill !== undefined) style.fill = safeColor(format.fill);

  const color = safeColor(format.color);
  if (format.bold !== undefined || format.italic !== undefined || format.strike !== undefined || format.underline !== undefined || color) {
    const font = resolveCellStyle(frame.styles, index).font;
    const changed: Font = { ...font };
    if (format.bold !== undefined) changed.bold = format.bold === true;
    if (format.italic !== undefined) changed.italic = format.italic === true;
    if (format.strike !== undefined) changed.strike = format.strike === true;
    if (format.underline !== undefined) changed.underline = format.underline;
    const text = textStyle(frame, changed);
    style.text = color ? { ...text, color } : text;
  }

  if (format.left || format.right || format.top || format.bottom) {
    style.left = edgeStyle(format.left, frame.styles) || base.left;
    style.right = edgeStyle(format.right, frame.styles) || base.right;
    style.top = edgeStyle(format.top, frame.styles) || base.top;
    style.bottom = edgeStyle(format.bottom, frame.styles) || base.bottom;
    style.hasBorder = !!(style.left || style.right || style.top || style.bottom);
  }
  if (format.numFmt && typeof format.numFmt.code === 'string') {
    const id = format.numFmt.id;
    style.numFmt = format.numFmt.code;
    style.format = { locale: frame.format.locale, date1904: frame.format.date1904, builtinId: id >= 0 ? builtinFormatId(frame.styles, id) : undefined };
  }
  byIndex.set(index, style);
  return style;
}

/** Style index of a cell: its own, else the row's, else the column's. */
function styleIndexAt(sheet: Sheet, row: number, col: number): number {
  const rowData = sheet.rows.get(row);
  const cell = rowData ? rowData.cells[col] : undefined;
  if (cell) return cell.style;
  if (rowData && rowData.style !== undefined) return rowData.style;
  for (const range of sheet.columns) {
    if (range.min > col) break;
    if (range.max >= col) return range.style !== undefined ? range.style : 0;
  }
  return 0;
}

function hasContent(cell: Cell | undefined): cell is Cell {
  if (cell === undefined) return false;
  // A formula that has not been calculated yet is drawn as its text.
  if (cell.value === null || cell.value === undefined) return cell.formula !== undefined && cell.formula.text !== '';
  return cell.value !== '';
}

// ---------------------------------------------------------------------------
// One pane

function drawPane(frame: Frame, view: PaneView): void {
  const { ctx, sheet, theme } = frame;
  const { pane, cols, rows, merges } = view;
  const rect = pane.rect;
  const colCount = cols.index.length;
  const rowCount = rows.index.length;

  ctx.save();
  ctx.beginPath();
  ctx.rect(rect.x, rect.y, rect.width, rect.height);
  ctx.clip();
  frame.painter.invalidate();

  const showGridLines = sheet.view.showGridLines !== false;
  if (showGridLines) {
    // Each line is the last pixel of its row or column. Fills are painted over them afterwards.
    ctx.strokeStyle = theme.gridLine;
    ctx.lineWidth = 1;
    ctx.beginPath();
    for (let ci = 0; ci < colCount; ci++) {
      const x = cols.pos[ci] + cols.size[ci] - 0.5;
      ctx.moveTo(x, rect.y);
      ctx.lineTo(x, rect.y + rect.height);
    }
    for (let ri = 0; ri < rowCount; ri++) {
      const y = rows.pos[ri] + rows.size[ri] - 0.5;
      ctx.moveTo(rect.x, y);
      ctx.lineTo(rect.x + rect.width, y);
    }
    ctx.stroke();
  }

  const marks = markMerges(view);
  const colStyles = columnStyles(sheet, cols);
  // Border edges by visible position: hEdges[ri * colCount + ci] is the top
  // edge of a cell (row `rowCount` holds the bottom edges of the last row),
  // vEdges[ri * (colCount + 1) + ci] its left edge. Allocated on first use.
  let hEdges: (EdgeStyle | undefined)[] | undefined;
  let vEdges: (EdgeStyle | undefined)[] | undefined;
  const textRows: number[] = [];
  const textCols: number[] = [];

  const conditional = frame.conditional;
  let found: (CellConditional | undefined)[] | undefined;
  let lastIndex = -1;
  let lastStyle: DrawStyle | undefined;
  for (let ri = 0; ri < rowCount; ri++) {
    const r = rows.index[ri];
    const row = sheet.rows.get(r);
    const rowStyle = row ? row.style : undefined;
    for (let ci = 0; ci < colCount; ci++) {
      const c = cols.index[ci];
      const cell = row ? row.cells[c] : undefined;
      const index = cell ? cell.style : rowStyle !== undefined ? rowStyle : colStyles[ci];
      if (index !== lastIndex || !lastStyle) {
        lastStyle = drawStyle(frame, index);
        lastIndex = index;
      }
      let style = lastStyle;
      const mark = marks ? marks[ri * colCount + ci] : 0;
      let fill = style.fill;
      if (conditional) {
        // Every cell of a merged area is formatted as its top-left cell is.
        const cond = mark === 0 ? conditional.resolve(r, c) : conditional.resolve(merges[mark - 1].range.top, merges[mark - 1].range.left);
        if (cond) {
          if (cond.style) {
            style = conditionalStyle(frame, index, cond.style);
            fill = style.fill;
          }
          if (cond.fill !== undefined) fill = safeColor(cond.fill) || fill;
          if (mark === 0) {
            if (!found) found = view.conditional = new Array<CellConditional | undefined>(rowCount * colCount);
            found[ri * colCount + ci] = cond;
          }
        }
      }

      if (mark === 0) {
        if (fill) {
          // One pixel further up and left: a filled cell hides the grid lines on all its sides.
          frame.painter.setFill(fill);
          ctx.fillRect(cols.pos[ci] - 1, rows.pos[ri] - 1, cols.size[ci] + 1, rows.size[ri] + 1);
        }
        if (found) {
          const cond = found[ri * colCount + ci];
          if (cond && cond.bar) drawDataBar(frame, cond.bar, cols.pos[ci], rows.pos[ri], cols.size[ci], rows.size[ri]);
        }
        if (hasContent(cell)) {
          textRows.push(ri);
          textCols.push(ci);
        }
      }

      if (style.hasBorder) {
        let { left, right, top, bottom } = style;
        if (mark !== 0) {
          // Inside a merge only the edges on its outline are drawn, each from the cell that lies there.
          const range = merges[mark - 1].range;
          if (r !== range.top) top = undefined;
          if (r !== range.bottom) bottom = undefined;
          if (c !== range.left) left = undefined;
          if (c !== range.right) right = undefined;
        }
        if (!hEdges || !vEdges) {
          hEdges = new Array<EdgeStyle | undefined>((rowCount + 1) * colCount);
          vEdges = new Array<EdgeStyle | undefined>(rowCount * (colCount + 1));
        }
        if (top) hEdges[ri * colCount + ci] = stronger(hEdges[ri * colCount + ci], top);
        if (bottom) hEdges[(ri + 1) * colCount + ci] = stronger(hEdges[(ri + 1) * colCount + ci], bottom);
        if (left) vEdges[ri * (colCount + 1) + ci] = stronger(vEdges[ri * (colCount + 1) + ci], left);
        if (right) vEdges[ri * (colCount + 1) + ci + 1] = stronger(vEdges[ri * (colCount + 1) + ci + 1], right);
      }
    }
  }

  // Merged areas: the fill of the top-left cell, or the background to take the grid lines out.
  for (let i = 0; i < merges.length; i++) {
    const merge = merges[i];
    const index = styleIndexAt(sheet, merge.range.top, merge.range.left);
    const cond = conditional ? conditional.resolve(merge.range.top, merge.range.left) : undefined;
    const style = cond && cond.style ? conditionalStyle(frame, index, cond.style) : drawStyle(frame, index);
    const fill = (cond && safeColor(cond.fill)) || style.fill;
    if (fill) {
      frame.painter.setFill(fill);
      ctx.fillRect(merge.x - 1, merge.y - 1, merge.width + 1, merge.height + 1);
    } else if (showGridLines) {
      frame.painter.setFill(theme.background);
      ctx.fillRect(merge.x, merge.y, merge.width - 1, merge.height - 1);
    }
    if (cond && cond.bar) drawDataBar(frame, cond.bar, merge.x, merge.y, merge.width, merge.height);
  }

  if (hEdges && vEdges) strokeEdges(ctx, view, hEdges, vEdges);

  drawPaneText(frame, view, textRows, textCols);
  if (frame.cellMark) drawCellMarks(frame, view, frame.cellMark);
  if (frame.filter) drawFilterButtons(frame, view, frame.filter);
  if (frame.dropdown) drawDropdown(frame, view, frame.dropdown);
  if (frame.selection) drawSelection(frame, view, frame.selection);

  ctx.restore();
  frame.painter.invalidate();
}

/** Of two borders on the same line the heavier one is drawn; the first one stays on a tie. */
function stronger(current: EdgeStyle | undefined, candidate: EdgeStyle): EdgeStyle {
  return current && current.line.rank >= candidate.line.rank ? current : candidate;
}

function strokeEdges(
  ctx: CanvasRenderingContext2D,
  view: PaneView,
  hEdges: (EdgeStyle | undefined)[],
  vEdges: (EdgeStyle | undefined)[],
): void {
  const { cols, rows } = view;
  const rect = view.pane.rect;
  const colCount = cols.index.length;
  const rowCount = rows.index.length;
  const paneX = Math.round(rect.x);
  const paneY = Math.round(rect.y);

  let current: EdgeStyle | undefined;
  // Edges that share a style go into one path.
  const use = (edge: EdgeStyle) => {
    if (current && current.key === edge.key) return;
    if (current) ctx.stroke();
    current = edge;
    ctx.strokeStyle = edge.color;
    ctx.lineWidth = edge.line.width;
    ctx.setLineDash(edge.line.dash);
    ctx.globalAlpha = edge.line.alpha;
    ctx.beginPath();
  };

  for (let ri = 0; ri <= rowCount; ri++) {
    // The edge lies on the grid line, the last pixel row of the row above. At
    // the top of the pane that pixel is outside the clip, so the edge moves in.
    let y = ri < rowCount ? rows.pos[ri] : rows.pos[rowCount - 1] + rows.size[rowCount - 1];
    if (ri === 0 && y === paneY) y += 1;
    for (let ci = 0; ci < colCount; ci++) {
      const edge = hEdges[ri * colCount + ci];
      if (!edge) continue;
      use(edge);
      const from = cols.pos[ci] - 1;
      const to = cols.pos[ci] + cols.size[ci];
      if (edge.line.double) {
        ctx.moveTo(from, y - 1.5);
        ctx.lineTo(to, y - 1.5);
        ctx.moveTo(from, y + 0.5);
        ctx.lineTo(to, y + 0.5);
      } else {
        // Odd widths are centred on a pixel, even ones on a pixel boundary: both stay sharp.
        const centre = edge.line.width % 2 === 1 ? y - 0.5 : y;
        ctx.moveTo(from, centre);
        ctx.lineTo(to, centre);
      }
    }
  }
  for (let ci = 0; ci <= colCount; ci++) {
    let x = ci < colCount ? cols.pos[ci] : cols.pos[colCount - 1] + cols.size[colCount - 1];
    if (ci === 0 && x === paneX) x += 1;
    for (let ri = 0; ri < rowCount; ri++) {
      const edge = vEdges[ri * (colCount + 1) + ci];
      if (!edge) continue;
      use(edge);
      const from = rows.pos[ri] - 1;
      const to = rows.pos[ri] + rows.size[ri];
      if (edge.line.double) {
        ctx.moveTo(x - 1.5, from);
        ctx.lineTo(x - 1.5, to);
        ctx.moveTo(x + 0.5, from);
        ctx.lineTo(x + 0.5, to);
      } else {
        const centre = edge.line.width % 2 === 1 ? x - 0.5 : x;
        ctx.moveTo(centre, from);
        ctx.lineTo(centre, to);
      }
    }
  }
  if (current) ctx.stroke();
  ctx.setLineDash(NO_DASH);
  ctx.globalAlpha = 1;
  ctx.lineWidth = 1;
}

// ---------------------------------------------------------------------------
// Text

function drawPaneText(frame: Frame, view: PaneView, textRows: number[], textCols: number[]): void {
  const { sheet, layout } = frame;
  const { pane, cols, rows, merges } = view;

  for (let i = 0; i < textRows.length; i++) {
    const ri = textRows[i];
    const ci = textCols[i];
    const row = sheet.rows.get(rows.index[ri]) as Row;
    const cell = row.cells[cols.index[ci]];
    const cond = view.conditional ? view.conditional[ri * cols.index.length + ci] : undefined;
    const style = cond && cond.style ? conditionalStyle(frame, cell.style, cond.style) : drawStyle(frame, cell.style);
    drawCellText(frame, view, cell, style, cols.pos[ci], rows.pos[ri], cols.size[ci], rows.size[ri], row, rows.index[ri], cols.index[ci], false, cond);
  }

  // A merged area shows the top-left cell, which may itself be out of view.
  for (let i = 0; i < merges.length; i++) {
    const merge = merges[i];
    const row = sheet.rows.get(merge.range.top);
    const cell = row ? row.cells[merge.range.left] : undefined;
    if (!hasContent(cell)) continue;
    const cond = frame.conditional ? frame.conditional.resolve(merge.range.top, merge.range.left) : undefined;
    const style = cond && cond.style ? conditionalStyle(frame, cell.style, cond.style) : drawStyle(frame, cell.style);
    drawCellText(frame, view, cell, style, merge.x, merge.y, merge.width, merge.height, undefined, merge.range.top, merge.range.left, false, cond);
  }

  // Text of cells just outside the pane can overflow into it.
  const firstCol = cols.index[0];
  const lastCol = cols.index[cols.index.length - 1];
  const leftLimit = Math.max(pane.frozenCols ? 0 : layout.frozenCols, firstCol - OVERFLOW_SCAN);
  // Frozen columns end at the frozen boundary: what lies beyond scrolls separately.
  const rightLimit = Math.min(pane.frozenCols ? layout.frozenCols - 1 : MAX_COLS - 1, lastCol + OVERFLOW_SCAN);
  for (let ri = 0; ri < rows.index.length; ri++) {
    const r = rows.index[ri];
    const row = sheet.rows.get(r);
    if (!row) continue;
    for (let c = firstCol - 1; c >= leftLimit; c--) {
      if (!hasContent(row.cells[c])) continue;
      drawOutsideCell(frame, view, row, r, c, ri);
      break;
    }
    for (let c = lastCol + 1; c <= rightLimit; c++) {
      if (!hasContent(row.cells[c])) continue;
      drawOutsideCell(frame, view, row, r, c, ri);
      break;
    }
  }
}

/** Draws a cell beside the pane if its text may reach into the pane. */
function drawOutsideCell(frame: Frame, view: PaneView, row: Row, r: number, c: number, ri: number): void {
  const cell = row.cells[c];
  if (cell.kind !== 'string') return;
  // A conditional format can make the text bold, and with that wider.
  const cond = frame.conditional ? frame.conditional.resolve(r, c) : undefined;
  const style = cond && cond.style ? conditionalStyle(frame, cell.style, cond.style) : drawStyle(frame, cell.style);
  if (style.wrap) return;
  const metrics = frame.layout.metrics;
  const width = metrics.cols.size(c);
  if (!(width > 0) || mergeIndexAt(view.merges, r, c) !== 0) return;
  const x = Math.round(view.cols.origin + metrics.cols.offset(c));
  drawCellText(frame, view, cell, style, x, view.rows.pos[ri], width, view.rows.size[ri], row, r, c, true, cond);
}

/** Cuts the text to the length that is measured and takes line breaks out of text that does not wrap. */
function prepareText(text: string, limit: number, wrap: boolean): string {
  let result = text.length > limit ? text.slice(0, limit) : text;
  if (!wrap && (result.indexOf('\n') >= 0 || result.indexOf('\r') >= 0)) result = result.replace(/[\r\n]+/g, '');
  return result;
}

function cellSegments(frame: Frame, cell: Cell, style: DrawStyle, wrap: boolean): Segment[] {
  const limit = wrap ? MAX_WRAP_CHARS : MAX_PLAIN_CHARS;
  const segments: Segment[] = [];
  const runs = cell.kind === 'string' ? cell.runs : undefined;
  if (runs && runs.length > 0) {
    let total = 0;
    const count = Math.min(runs.length, MAX_RUNS);
    for (let i = 0; i < count && total < limit; i++) {
      const run = runs[i];
      if (!run || typeof run.text !== 'string') continue;
      const text = prepareText(run.text, limit - total, wrap);
      if (text.length === 0) continue;
      const runStyle = run.font ? textStyle(frame, run.font) : style.text;
      const previous = segments.length > 0 ? segments[segments.length - 1] : undefined;
      // Runs that look the same are drawn as one piece.
      if (previous && previous.style === runStyle) previous.text += text;
      else segments.push({ text, style: runStyle });
      total += text.length;
    }
    return segments;
  }
  const formatted = formatCell(cell, style.numFmt, style.format);
  // A value its format cannot show (a negative date) is drawn like a number that does not fit.
  const text = formatted.overflow ? '#'.repeat(MAX_HASHES) : prepareText(formatted.text, limit, wrap);
  if (text.length === 0) return segments;
  let shown = style.text;
  if (formatted.color && formatted.color !== shown.color) {
    const key = shown.font + formatted.color + shown.underline + shown.strike;
    let coloured = frame.colorStyleCache.get(key);
    if (!coloured) {
      coloured = { font: shown.font, px: shown.px, color: formatted.color, underline: shown.underline, strike: shown.strike };
      frame.colorStyleCache.set(key, coloured);
    }
    shown = coloured;
  }
  segments.push({ text, style: shown });
  return segments;
}

/** Width of the characters [start, end) of the joined segments, each in its own font. */
function measureRange(frame: Frame, segments: Segment[], start: number, end: number): number {
  let width = 0;
  let offset = 0;
  for (let i = 0; i < segments.length && offset < end; i++) {
    const segment = segments[i];
    const length = segment.text.length;
    const from = Math.max(start - offset, 0);
    const to = Math.min(end - offset, length);
    if (from < to) width += frame.painter.measure(segment.style.font, from === 0 && to === length ? segment.text : segment.text.slice(from, to));
    offset += length;
  }
  return width;
}

/** Draws the characters [start, end) of the joined segments from `x` on, with underline and strike-through. */
function drawRange(frame: Frame, segments: Segment[], start: number, end: number, x: number, baseline: number): void {
  const { ctx, painter } = frame;
  let offset = 0;
  for (let i = 0; i < segments.length && offset < end; i++) {
    const segment = segments[i];
    const length = segment.text.length;
    const from = Math.max(start - offset, 0);
    const to = Math.min(end - offset, length);
    offset += length;
    if (from >= to) continue;
    const text = from === 0 && to === length ? segment.text : segment.text.slice(from, to);
    const style = segment.style;
    painter.setFont(style.font);
    painter.setFill(style.color);
    ctx.fillText(text, x, baseline);
    const decorated = style.underline > 0 || style.strike;
    if (!decorated && offset >= end) return;
    const width = painter.measure(style.font, text);
    if (decorated) {
      const thickness = Math.max(1, Math.round(style.px / 14));
      if (style.underline > 0) {
        const y = Math.round(baseline + Math.max(1, style.px * 0.1));
        ctx.fillRect(x, y, width, thickness);
        if (style.underline > 1) ctx.fillRect(x, y + thickness + 1, width, thickness);
      }
      if (style.strike) ctx.fillRect(x, Math.round(baseline - style.px * 0.3), width, thickness);
    }
    x += width;
  }
}

/** True when text may not be drawn over this cell: it has content or belongs to a merge. */
function blocksOverflow(view: PaneView, row: Row, r: number, c: number): boolean {
  return hasContent(row.cells[c]) || mergeIndexAt(view.merges, r, c) !== 0;
}

/**
 * Draws the text of one cell into the box (x, y, width, height). `row` is the
 * row of the cell when its text may overflow into empty neighbours, undefined
 * for a merged area, which never overflows. `outside` is set for a cell beside
 * the pane: it is drawn only if its text leaves the cell towards the pane.
 */
function drawCellText(
  frame: Frame,
  view: PaneView,
  cell: Cell,
  style: DrawStyle,
  x: number,
  y: number,
  width: number,
  height: number,
  row: Row | undefined,
  r: number,
  c: number,
  outside = false,
  cond?: CellConditional,
): void {
  const { ctx, painter } = frame;
  const kind = cell.kind;
  const numeric = kind === 'number' || kind === 'date';
  const wrap = style.wrap && kind === 'string';
  if (cond && (cond.icon || cond.hideValue)) {
    const icon = cond.icon ? iconGlyph(cond.icon) : undefined;
    // The icon is as high as a line of the cell's text, and never larger than the cell.
    const size = Math.floor(Math.min(Math.max(style.text.px * 1.05, MIN_ICON), MAX_ICON, height - 3, width - 2 * CELL_PADDING));
    const shown = icon !== undefined && size >= MIN_ICON && !outside;
    const top = style.vertical === 'top' ? y + 1 : style.vertical === 'center' ? y + (height - size) / 2 : y + height - 2 - size;
    if (cond.hideValue) {
      // Without its value an icon goes where the alignment of the cell puts it.
      const left = style.horizontal === 'center' ? x + (width - size) / 2 : style.horizontal === 'right' ? x + width - CELL_PADDING - size : x + CELL_PADDING;
      if (shown) drawIcon(frame, icon as IconGlyph, Math.round(left), Math.round(top), size);
      return;
    }
    if (shown) {
      drawIcon(frame, icon as IconGlyph, x + CELL_PADDING, Math.round(top), size);
      // The text gets what is left of the cell.
      x += size + ICON_GAP;
      width -= size + ICON_GAP;
      if (width <= 2 * CELL_PADDING) return;
    }
  }
  if (cell.value === 0 && frame.sheet.view.showZeros === false) return;
  let segments = cellSegments(frame, cell, style, wrap);
  if (segments.length === 0) return;

  const align = style.horizontal !== 'general' ? style.horizontal : kind === 'string' || kind === 'blank' ? 'left' : numeric ? 'right' : 'center';
  const indent = align === 'center' ? 0 : style.indent * INDENT_WIDTH;
  const available = width - 2 * CELL_PADDING - indent;

  let px = segments[0].style.px;
  let length = segments[0].text.length;
  for (let i = 1; i < segments.length; i++) {
    px = Math.max(px, segments[i].style.px);
    length += segments[i].text.length;
  }
  const lineHeight = px * LINE_HEIGHT;

  // Lines as [start, end) offsets into the joined text, two numbers per line.
  let lines: number[];
  let clipLeft = x;
  let clipRight = x + width;
  let clip = false;
  // Width of the only line of text that does not wrap; -1 for wrapped text.
  let singleWidth = -1;

  if (wrap) {
    // Rich text wraps as if it were all in the first font; each line is then measured run by run.
    const plain = segments.length === 1 ? segments[0].text : segments.map((segment) => segment.text).join('');
    const font = segments[0].style.font;
    lines = wrapLines(plain, Math.max(available, 1), (text) => painter.measure(font, text));
    // Lines only exceed the box when there are too many of them or a single character is wider than it.
    clip = (lines.length / 2) * lineHeight > height || available < px;
  } else {
    lines = [0, length];
    singleWidth = measureRange(frame, segments, 0, length);
    if (outside && (singleWidth <= available || (align === 'left' ? x > view.pane.rect.x : align === 'right' && x < view.pane.rect.x))) return;
    if (singleWidth > available) {
      if (numeric) {
        // A number is never cut: it shows as many "#" as fit.
        const font = segments[0].style;
        const hashWidth = painter.measure(font.font, '#');
        const count = hashWidth > 0 ? Math.min(Math.floor(available / hashWidth), MAX_HASHES) : 0;
        if (count <= 0) return;
        segments = [{ text: '#'.repeat(count), style: font }];
        length = count;
        lines = [0, length];
        singleWidth = measureRange(frame, segments, 0, length);
      } else {
        let textLeft: number;
        if (align === 'left') textLeft = x + CELL_PADDING + indent;
        else if (align === 'right') textLeft = x + width - CELL_PADDING - indent - singleWidth;
        else textLeft = x + (width - singleWidth) / 2;
        const textRight = textLeft + singleWidth;
        if (kind === 'string' && row) {
          // Text runs over empty neighbours until the first cell that is in use or the edge of the pane.
          const sizes = frame.layout.metrics.cols;
          const paneLeft = view.pane.rect.x;
          const paneRight = paneLeft + view.pane.rect.width;
          for (let next = c + 1; clipRight < textRight && clipRight < paneRight && next < MAX_COLS && !blocksOverflow(view, row, r, next); next++) {
            clipRight += sizes.size(next);
          }
          for (let next = c - 1; clipLeft > textLeft && clipLeft > paneLeft && next >= 0 && !blocksOverflow(view, row, r, next); next--) {
            clipLeft -= sizes.size(next);
          }
        }
        clip = textLeft < clipLeft || textRight > clipRight;
      }
    }
    if (lineHeight > height) clip = true;
  }

  const lineCount = lines.length / 2;
  const blockHeight = lineCount * lineHeight;
  let blockTop: number;
  if (style.vertical === 'top') blockTop = y + 1;
  else if (style.vertical === 'center') blockTop = y + (height - blockHeight) / 2;
  else blockTop = y + height - 1 - blockHeight;

  // Only the lines that can be seen are drawn; a long wrapped text has thousands.
  const paneRect = view.pane.rect;
  const visibleTop = Math.max(y, paneRect.y);
  const visibleBottom = Math.min(y + height, paneRect.y + paneRect.height);
  let first = 0;
  let last = lineCount - 1;
  if (lineCount > 1) {
    first = Math.max(Math.floor((visibleTop - blockTop) / lineHeight), 0);
    last = Math.min(Math.ceil((visibleBottom - blockTop) / lineHeight), lineCount - 1);
  }
  if (first > last) return;

  if (clip) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(clipLeft, y, clipRight - clipLeft, height);
    ctx.clip();
  }
  for (let i = first; i <= last; i++) {
    const start = lines[i * 2];
    const end = lines[i * 2 + 1];
    if (end <= start) continue;
    const lineWidth = singleWidth >= 0 ? singleWidth : measureRange(frame, segments, start, end);
    let left: number;
    if (align === 'left') left = x + CELL_PADDING + indent;
    else if (align === 'right') left = x + width - CELL_PADDING - indent - lineWidth;
    else left = x + (width - lineWidth) / 2;
    const baseline = Math.round(blockTop + i * lineHeight + lineHeight - px * DESCENT);
    drawRange(frame, segments, start, end, Math.round(left), baseline);
  }
  if (clip) {
    ctx.restore();
    painter.invalidate();
  }
}

// ---------------------------------------------------------------------------
// Data bars and icons

/** "#RRGGBB" moved towards white by `share` (0..1). */
function fade(color: string, share: number): string {
  let out = '#';
  for (let i = 1; i < 7; i += 2) {
    const channel = parseInt(color.slice(i, i + 2), 16);
    const value = Math.round(channel + (255 - channel) * share);
    out += (value < 16 ? '0' : '') + value.toString(16).toUpperCase();
  }
  return out;
}

/** Draws a data bar into the cell at (x, y, width, height): above the fill, below the text. */
function drawDataBar(frame: Frame, bar: CellDataBar, x: number, y: number, width: number, height: number): void {
  const { ctx, painter } = frame;
  const color = safeColor(bar.color);
  // The last pixel of a cell is its grid line.
  const left = x + BAR_INSET;
  const top = y + BAR_INSET;
  const room = width - 1 - 2 * BAR_INSET;
  const barHeight = height - 1 - 2 * BAR_INSET;
  if (!color || !(room > 0) || !(barHeight > 0)) return;
  const unit = (value: number) => (value > 0 ? (value < 1 ? value : 1) : 0);
  const mirrored = bar.rightToLeft === true;
  let from = unit(bar.from);
  let to = unit(bar.to);
  if (mirrored) {
    const swap = from;
    from = 1 - to;
    to = 1 - swap;
  }
  const x0 = Math.round(left + from * room);
  const x1 = Math.round(left + to * room);
  if (x1 > x0) {
    if (bar.gradient && typeof ctx.createLinearGradient === 'function') {
      // The bar is solid where it starts (at the axis for a negative one) and fades towards its end.
      const startsRight = (bar.negative === true) !== mirrored;
      const gradient = ctx.createLinearGradient(x0, 0, x1, 0);
      gradient.addColorStop(startsRight ? 1 : 0, color);
      gradient.addColorStop(startsRight ? 0 : 1, fade(color, BAR_FADE));
      ctx.fillStyle = gradient;
      painter.invalidate();
    } else painter.setFill(color);
    ctx.fillRect(x0, top, x1 - x0, barHeight);
    const border = safeColor(bar.border);
    if (border && x1 - x0 > 1 && barHeight > 1) {
      ctx.strokeStyle = border;
      ctx.lineWidth = 1;
      ctx.strokeRect(x0 + 0.5, top + 0.5, x1 - x0 - 1, barHeight - 1);
    }
  }
  if (typeof bar.axis === 'number' && bar.axis === bar.axis) {
    const axis = unit(mirrored ? 1 - bar.axis : bar.axis);
    const at = Math.round(left + axis * room) + 0.5;
    ctx.strokeStyle = safeColor(bar.axisColor) || '#000000';
    ctx.lineWidth = 1;
    ctx.setLineDash(AXIS_DASH);
    ctx.beginPath();
    ctx.moveTo(at, y);
    ctx.lineTo(at, y + height - 1);
    ctx.stroke();
    ctx.setLineDash(NO_DASH);
  }
}

/** Fills the polygon whose corners are given as fractions of a box of `size` pixels at (x, y). */
function fillPolygon(ctx: CanvasRenderingContext2D, x: number, y: number, size: number, points: number[]): void {
  ctx.beginPath();
  ctx.moveTo(x + points[0] * size, y + points[1] * size);
  for (let i = 2; i < points.length; i += 2) ctx.lineTo(x + points[i] * size, y + points[i + 1] * size);
  ctx.closePath();
  ctx.fill();
}

/** Strokes a line through the points, given as fractions of the box. */
function strokeLine(ctx: CanvasRenderingContext2D, x: number, y: number, size: number, points: number[], color: string, lineWidth: number): void {
  ctx.strokeStyle = color;
  ctx.lineWidth = lineWidth;
  ctx.beginPath();
  ctx.moveTo(x + points[0] * size, y + points[1] * size);
  for (let i = 2; i < points.length; i += 2) ctx.lineTo(x + points[i] * size, y + points[i + 1] * size);
  ctx.stroke();
  ctx.lineWidth = 1;
}

function fillCircle(ctx: CanvasRenderingContext2D, cx: number, cy: number, radius: number): void {
  ctx.beginPath();
  ctx.arc(cx, cy, radius, 0, Math.PI * 2);
  ctx.fill();
}

/** An arrow pointing up, as fractions of its box measured from the centre. */
const ARROW = [0, -0.46, 0.42, -0.02, 0.15, -0.02, 0.15, 0.46, -0.15, 0.46, -0.15, -0.02, -0.42, -0.02];
const STAR: number[] = [];
for (let i = 0; i < 10; i++) {
  const radius = i % 2 === 0 ? 0.48 : 0.2;
  const angle = (Math.PI * i) / 5 - Math.PI / 2;
  STAR.push(0.5 + radius * Math.cos(angle), 0.52 + radius * Math.sin(angle));
}
/** The left half of the star: the corners at or left of the middle, closed along the middle. */
const STAR_HALF = [STAR[0], STAR[1], 0.5, STAR[11], STAR[12], STAR[13], STAR[14], STAR[15], STAR[16], STAR[17], STAR[18], STAR[19]];
const CHECK = [0.2, 0.55, 0.42, 0.76, 0.8, 0.26];

/** A check mark, a cross or an exclamation mark, drawn into the box at (x, y). */
function drawSymbol(frame: Frame, shape: string, color: string, x: number, y: number, size: number): void {
  const { ctx, painter } = frame;
  const weight = Math.max(size * 0.16, 1.5);
  if (shape === 'check') strokeLine(ctx, x, y, size, CHECK, color, weight);
  else if (shape === 'cross') {
    strokeLine(ctx, x, y, size, [0.24, 0.24, 0.76, 0.76], color, weight);
    strokeLine(ctx, x, y, size, [0.76, 0.24, 0.24, 0.76], color, weight);
  } else {
    painter.setFill(color);
    ctx.fillRect(x + size * 0.42, y + size * 0.16, size * 0.16, size * 0.42);
    ctx.fillRect(x + size * 0.42, y + size * 0.68, size * 0.16, size * 0.16);
  }
}

/** Draws one icon of an icon set into a square of `size` pixels at (x, y). Vector shapes only. */
function drawIcon(frame: Frame, icon: IconGlyph, x: number, y: number, size: number): void {
  const { ctx, painter } = frame;
  const color = safeColor(icon.color) || ICON_DARK;
  const cx = x + size / 2;
  const cy = y + size / 2;
  const level = icon.level;
  painter.setFill(color);
  switch (icon.shape) {
    case 'arrow': {
      const angle = (level * Math.PI) / 180;
      const cos = Math.cos(angle);
      const sin = Math.sin(angle);
      const points: number[] = [];
      for (let i = 0; i < ARROW.length; i += 2) points.push(0.5 + ARROW[i] * cos - ARROW[i + 1] * sin, 0.5 + ARROW[i] * sin + ARROW[i + 1] * cos);
      fillPolygon(ctx, x, y, size, points);
      break;
    }
    case 'circle':
      fillCircle(ctx, cx, cy, size * 0.44);
      break;
    case 'rimCircle':
      painter.setFill(ICON_DARK);
      ctx.fillRect(x + size * 0.04, y + size * 0.04, size * 0.92, size * 0.92);
      painter.setFill(color);
      fillCircle(ctx, cx, cy, size * 0.34);
      break;
    case 'flag':
      fillPolygon(ctx, x, y, size, [0.26, 0.08, 0.9, 0.3, 0.26, 0.54]);
      painter.setFill(ICON_DARK);
      ctx.fillRect(x + size * 0.16, y + size * 0.06, Math.max(size * 0.1, 1), size * 0.88);
      break;
    case 'diamond':
      fillPolygon(ctx, x, y, size, [0.5, 0.04, 0.96, 0.5, 0.5, 0.96, 0.04, 0.5]);
      break;
    case 'triangle':
      fillPolygon(ctx, x, y, size, [0.5, 0.1, 0.96, 0.88, 0.04, 0.88]);
      break;
    case 'triangleDown':
      fillPolygon(ctx, x, y, size, [0.04, 0.12, 0.96, 0.12, 0.5, 0.9]);
      break;
    case 'dash':
      ctx.fillRect(x + size * 0.08, y + size * 0.4, size * 0.84, size * 0.2);
      break;
    case 'check':
    case 'cross':
    case 'exclamation':
      drawSymbol(frame, icon.shape, color, x, y, size);
      break;
    case 'checkCircle':
    case 'crossCircle':
    case 'exclamationCircle':
      fillCircle(ctx, cx, cy, size * 0.48);
      drawSymbol(frame, icon.shape.slice(0, -6), ICON_LIGHT, x + size * 0.14, y + size * 0.14, size * 0.72);
      break;
    case 'bars':
      for (let i = 0; i < 4; i++) {
        const barHeight = size * (0.3 + 0.2 * i);
        painter.setFill(i < level ? color : ICON_OFF);
        ctx.fillRect(x + size * (0.06 + 0.23 * i), y + size * 0.94 - barHeight, size * 0.18, barHeight);
      }
      break;
    case 'boxes':
      for (let i = 0; i < 4; i++) {
        // Filled from the bottom row up, left before right.
        const order = [2, 3, 0, 1][i];
        painter.setFill(order < level ? color : ICON_OFF);
        ctx.fillRect(x + size * (0.08 + 0.46 * (i % 2)), y + size * (0.08 + 0.46 * (i >> 1)), size * 0.38, size * 0.38);
      }
      break;
    case 'pie': {
      const radius = size * 0.44;
      painter.setFill(ICON_LIGHT);
      fillCircle(ctx, cx, cy, radius);
      if (level > 0) {
        painter.setFill(color);
        ctx.beginPath();
        ctx.moveTo(cx, cy);
        ctx.arc(cx, cy, radius, -Math.PI / 2, -Math.PI / 2 + (Math.min(level, 4) * Math.PI) / 2);
        ctx.closePath();
        ctx.fill();
      }
      ctx.strokeStyle = color;
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.arc(cx, cy, radius, 0, Math.PI * 2);
      ctx.stroke();
      break;
    }
    case 'star':
      painter.setFill(level >= 2 ? color : ICON_LIGHT);
      fillPolygon(ctx, x, y, size, STAR);
      if (level === 1) {
        painter.setFill(color);
        fillPolygon(ctx, x, y, size, STAR_HALF);
      }
      strokeLine(ctx, x, y, size, STAR.concat([STAR[0], STAR[1]]), color, 1);
      break;
  }
}

// ---------------------------------------------------------------------------
// Marks and buttons inside cells

const MARK_SIZE = 6;
const BUTTON_FACE = '#FFFFFF';
const BUTTON_BORDER = '#9AA0A6';
const BUTTON_GLYPH = '#444746';
const BUTTON_ACTIVE = '#1A7F45';
const DEFAULT_FILTER_TEXT = '#1A56C4';

/** A small triangle in the top left corner of every visible cell that has a mark. */
function drawCellMarks(frame: Frame, view: PaneView, markOf: (row: number, col: number) => string | undefined): void {
  const { ctx } = frame;
  const { cols, rows } = view;
  for (let ri = 0; ri < rows.index.length; ri++) {
    if (rows.size[ri] < MARK_SIZE) continue;
    for (let ci = 0; ci < cols.index.length; ci++) {
      if (cols.size[ci] < MARK_SIZE) continue;
      let color: string | undefined;
      try {
        color = markOf(rows.index[ri], cols.index[ci]);
      } catch (error) {
        return;
      }
      if (typeof color !== 'string' || !HEX_COLOR.test(color)) continue;
      const x = cols.pos[ci];
      const y = rows.pos[ri];
      frame.painter.setFill(color);
      ctx.beginPath();
      ctx.moveTo(x, y);
      ctx.lineTo(x + MARK_SIZE, y);
      ctx.lineTo(x, y + MARK_SIZE);
      ctx.closePath();
      ctx.fill();
    }
  }
}

/** A button in the bottom right corner of a cell: an arrow, or a funnel when `active`. */
function drawCellButton(frame: Frame, x: number, y: number, width: number, height: number, active: boolean): void {
  const rect = cellButtonRect(x, y, width, height);
  if (!rect) return;
  const { ctx } = frame;
  const size = rect.size;
  frame.painter.setFill(active ? BUTTON_ACTIVE : BUTTON_FACE);
  ctx.fillRect(rect.x, rect.y, size, size);
  ctx.strokeStyle = active ? BUTTON_ACTIVE : BUTTON_BORDER;
  ctx.lineWidth = 1;
  ctx.strokeRect(rect.x + 0.5, rect.y + 0.5, size - 1, size - 1);
  frame.painter.setFill(active ? BUTTON_FACE : BUTTON_GLYPH);
  const cx = rect.x + size / 2;
  const cy = rect.y + size / 2;
  const half = size / 4;
  ctx.beginPath();
  if (active) {
    // A funnel: wide at the top, with a stem.
    ctx.moveTo(cx - half - 1, cy - half);
    ctx.lineTo(cx + half + 1, cy - half);
    ctx.lineTo(cx + 1, cy);
    ctx.lineTo(cx + 1, cy + half + 1);
    ctx.lineTo(cx - 1, cy + half);
    ctx.lineTo(cx - 1, cy);
  } else {
    ctx.moveTo(cx - half, cy - half / 2);
    ctx.lineTo(cx + half, cy - half / 2);
    ctx.lineTo(cx, cy + half / 2 + 1);
  }
  ctx.closePath();
  ctx.fill();
}

function drawFilterButtons(frame: Frame, view: PaneView, filter: GridFilter): void {
  const { cols, rows } = view;
  const ri = lowerBound(rows.index, filter.range.top);
  if (ri >= rows.index.length || rows.index[ri] !== filter.range.top) return;
  for (let ci = lowerBound(cols.index, filter.range.left); ci < cols.index.length && cols.index[ci] <= filter.range.right; ci++) {
    const state = filter.buttons[cols.index[ci] - filter.range.left];
    if (state === 1 || state === 2) drawCellButton(frame, cols.pos[ci], rows.pos[ri], cols.size[ci], rows.size[ri], state === 2);
  }
}

/** The button of a list in a cell; in a merged cell it sits in the corner of the whole area. */
function drawDropdown(frame: Frame, view: PaneView, cell: { row: number; col: number }): void {
  const metrics = frame.layout.metrics;
  let area: CellRange = { top: cell.row, left: cell.col, bottom: cell.row, right: cell.col };
  for (const merge of frame.merges) {
    if (cell.row >= merge.top && cell.row <= merge.bottom && cell.col >= merge.left && cell.col <= merge.right) {
      area = merge;
      break;
    }
  }
  const pane = view.pane.rect;
  const x0 = Math.round(view.cols.origin + metrics.cols.offset(area.left));
  const x1 = Math.round(view.cols.origin + metrics.cols.offset(area.right + 1));
  const y0 = Math.round(view.rows.origin + metrics.rows.offset(area.top));
  const y1 = Math.round(view.rows.origin + metrics.rows.offset(area.bottom + 1));
  if (x1 <= pane.x || y1 <= pane.y || x1 > pane.x + pane.width + 1 || y1 > pane.y + pane.height + 1) return;
  drawCellButton(frame, x0, y0, x1 - x0, y1 - y0, false);
}

// ---------------------------------------------------------------------------
// Selection, headers, frozen lines

function drawSelection(frame: Frame, view: PaneView, selection: GridSelection): void {
  const { ctx, theme } = frame;
  const metrics = frame.layout.metrics;
  const { cols, rows } = view;
  const rect = view.pane.rect;
  const { range } = selection;

  const x0 = Math.round(cols.origin + metrics.cols.offset(range.left));
  const x1 = Math.round(cols.origin + metrics.cols.offset(range.right + 1));
  const y0 = Math.round(rows.origin + metrics.rows.offset(range.top));
  const y1 = Math.round(rows.origin + metrics.rows.offset(range.bottom + 1));
  // One pixel of slack: the frame reaches over the grid line before the range.
  if (x1 <= x0 || y1 <= y0 || x1 < rect.x || y1 < rect.y || x0 > rect.x + rect.width + 1 || y0 > rect.y + rect.height + 1) return;

  // The active cell (its whole merged area) stays unfilled.
  let active: CellRange = { top: selection.active.row, left: selection.active.col, bottom: selection.active.row, right: selection.active.col };
  for (const merge of frame.merges) {
    if (active.top >= merge.top && active.top <= merge.bottom && active.left >= merge.left && active.left <= merge.right) {
      active = merge;
      break;
    }
  }
  const ax0 = Math.min(Math.max(Math.round(cols.origin + metrics.cols.offset(active.left)), x0), x1);
  const ax1 = Math.min(Math.max(Math.round(cols.origin + metrics.cols.offset(active.right + 1)), x0), x1);
  const ay0 = Math.min(Math.max(Math.round(rows.origin + metrics.rows.offset(active.top)), y0), y1);
  const ay1 = Math.min(Math.max(Math.round(rows.origin + metrics.rows.offset(active.bottom + 1)), y0), y1);
  const single = ax0 === x0 && ax1 === x1 && ay0 === y0 && ay1 === y1;

  if (!single) {
    frame.painter.setFill(theme.selectionFill);
    if (ax1 <= ax0 || ay1 <= ay0) {
      ctx.fillRect(x0, y0, x1 - x0, y1 - y0);
    } else {
      if (ay0 > y0) ctx.fillRect(x0, y0, x1 - x0, ay0 - y0);
      if (ay1 < y1) ctx.fillRect(x0, ay1, x1 - x0, y1 - ay1);
      if (ax0 > x0) ctx.fillRect(x0, ay0, ax0 - x0, ay1 - ay0);
      if (ax1 < x1) ctx.fillRect(ax1, ay0, x1 - ax1, ay1 - ay0);
    }
  }

  ctx.strokeStyle = theme.selectionBorder;
  if (!single && ax1 > ax0 && ay1 > ay0) {
    ctx.lineWidth = 1;
    ctx.strokeRect(ax0 - 0.5, ay0 - 0.5, ax1 - ax0, ay1 - ay0);
  }
  // Two pixels: the grid line around the range and the pixel inside it.
  ctx.lineWidth = 2;
  ctx.strokeRect(x0, y0, x1 - x0 - 1, y1 - y0 - 1);
  ctx.lineWidth = 1;
}

function drawHeaders(frame: Frame, views: PaneView[]): void {
  const { ctx, layout, theme, selection } = frame;
  const headerWidth = layout.metrics.headerWidth;
  const headerHeight = layout.metrics.headerHeight;
  if (!(headerWidth > 0) && !(headerHeight > 0)) return;

  ctx.fillStyle = theme.headerBackground;
  if (headerHeight > 0) ctx.fillRect(0, 0, layout.width, headerHeight);
  if (headerWidth > 0) ctx.fillRect(0, 0, headerWidth, layout.height);
  ctx.font = HEADER_FONT;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.strokeStyle = theme.headerBorder;
  ctx.lineWidth = 1;

  // Panes that share their columns (or rows) share a header strip: one for the frozen part, one for the scrolling part.
  let frozenDone = false;
  let scrollDone = false;
  for (let i = 0; headerHeight > 0 && i < views.length; i++) {
    const view = views[i];
    if (view.pane.frozenCols ? frozenDone : scrollDone) continue;
    if (view.pane.frozenCols) frozenDone = true;
    else scrollDone = true;
    const { cols } = view;
    const count = cols.index.length;
    ctx.save();
    ctx.beginPath();
    ctx.rect(view.pane.rect.x, 0, view.pane.rect.width, headerHeight);
    ctx.clip();
    if (selection) {
      ctx.fillStyle = theme.headerSelectedBackground;
      for (let ci = 0; ci < count; ci++) {
        if (cols.index[ci] >= selection.range.left && cols.index[ci] <= selection.range.right) ctx.fillRect(cols.pos[ci], 0, cols.size[ci], headerHeight);
      }
    }
    ctx.beginPath();
    for (let ci = 0; ci < count; ci++) {
      const x = cols.pos[ci] + cols.size[ci] - 0.5;
      ctx.moveTo(x, 0);
      ctx.lineTo(x, headerHeight);
    }
    ctx.stroke();
    ctx.fillStyle = theme.headerText;
    for (let ci = 0; ci < count; ci++) {
      if (cols.size[ci] >= MIN_COL_LABEL_WIDTH) ctx.fillText(columnName(cols.index[ci]), cols.pos[ci] + cols.size[ci] / 2, headerHeight / 2);
    }
    ctx.restore();
  }

  frozenDone = false;
  scrollDone = false;
  for (let i = 0; headerWidth > 0 && i < views.length; i++) {
    const view = views[i];
    if (view.pane.frozenRows ? frozenDone : scrollDone) continue;
    if (view.pane.frozenRows) frozenDone = true;
    else scrollDone = true;
    const { rows } = view;
    const count = rows.index.length;
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, view.pane.rect.y, headerWidth, view.pane.rect.height);
    ctx.clip();
    if (selection) {
      ctx.fillStyle = theme.headerSelectedBackground;
      for (let ri = 0; ri < count; ri++) {
        if (rows.index[ri] >= selection.range.top && rows.index[ri] <= selection.range.bottom) ctx.fillRect(0, rows.pos[ri], headerWidth, rows.size[ri]);
      }
    }
    ctx.beginPath();
    for (let ri = 0; ri < count; ri++) {
      const y = rows.pos[ri] + rows.size[ri] - 0.5;
      ctx.moveTo(0, y);
      ctx.lineTo(headerWidth, y);
    }
    ctx.stroke();
    ctx.fillStyle = theme.headerText;
    // The rows a filter with criteria governs show their numbers in another colour, as Excel does.
    const filter = frame.filter && frame.filter.filtered ? frame.filter.range : undefined;
    const filterText = filter ? safeColor(theme.filterHeaderText) || DEFAULT_FILTER_TEXT : theme.headerText;
    let filterOn = false;
    for (let ri = 0; ri < count; ri++) {
      if (rows.size[ri] < MIN_ROW_LABEL_HEIGHT) continue;
      const inside = filter !== undefined && rows.index[ri] > filter.top && rows.index[ri] <= filter.bottom;
      if (inside !== filterOn) {
        filterOn = inside;
        ctx.fillStyle = inside ? filterText : theme.headerText;
      }
      ctx.fillText(String(rows.index[ri] + 1), headerWidth / 2, rows.pos[ri] + rows.size[ri] / 2);
    }
    ctx.restore();
  }

  // The lines that close the header strips; they also frame the corner box.
  ctx.beginPath();
  if (headerHeight > 0) {
    ctx.moveTo(0, headerHeight - 0.5);
    ctx.lineTo(layout.width, headerHeight - 0.5);
  }
  if (headerWidth > 0) {
    ctx.moveTo(headerWidth - 0.5, 0);
    ctx.lineTo(headerWidth - 0.5, layout.height);
  }
  ctx.stroke();
  ctx.textAlign = 'left';
  ctx.textBaseline = 'alphabetic';
  frame.painter.invalidate();
}

function drawFrozenLines(frame: Frame): void {
  const { ctx, layout, theme } = frame;
  const x = layout.metrics.headerWidth + layout.frozenWidth;
  const y = layout.metrics.headerHeight + layout.frozenHeight;
  const cols = layout.frozenCols > 0 && layout.frozenWidth > 0 && x <= layout.width;
  const rows = layout.frozenRows > 0 && layout.frozenHeight > 0 && y <= layout.height;
  if (!cols && !rows) return;
  ctx.strokeStyle = theme.frozenLine;
  ctx.lineWidth = 1;
  ctx.beginPath();
  if (cols) {
    ctx.moveTo(x - 0.5, 0);
    ctx.lineTo(x - 0.5, layout.height);
  }
  if (rows) {
    ctx.moveTo(0, y - 0.5);
    ctx.lineTo(layout.width, y - 0.5);
  }
  ctx.stroke();
}
