// Conditional formats: which format, colour, data bar and icon the rules of a
// sheet give a cell.
//
// The rules are read-only here (the writer copies their source elements). What
// they need from the cells of their ranges - the smallest and largest number,
// percentiles, the average, which values repeat - is worked out once per rule
// and kept until invalidate() is called, and the result for a cell is kept as
// well, so that scrolling does not look at a cell twice. Whole-column and
// whole-sheet ranges only cost what the sheet holds.
//
// Formulas are calculated by a CalcEngine when there is one. Without an engine
// the rules that need a formula do not apply; numbers and quoted text are read
// without it. Nothing here throws: a rule that cannot be evaluated does not
// apply.

import type { CalcEngine } from '../calc/engine';
import { adjustAnchoredFormula, adjustFormula, adjustRange, relativeExtent, shiftFormula, type StructuralEdit } from '../calc/refs';
import { compareScalars, isError, numberToText, textToNumber, type Scalar } from '../calc/values';
import { resolveColor } from '../model/sheet/color';
import { MAX_COLS, MAX_ROWS, cellName } from '../model/sheet/types';
import type {
  BorderEdge,
  Cell,
  CellRange,
  Color,
  ConditionalFormat,
  ConditionalRule,
  ConditionalValue,
  DifferentialFormat,
  Fill,
  Sheet,
  StyleSheet,
  Underline,
  Workbook,
} from '../model/sheet/types';
import { isoToSerial } from './cell-format';

/** What the differential formats of the rules that match a cell add up to. Only stated properties are present. */
export interface ConditionalStyle {
  /** "#RRGGBB", or null for "no fill". */
  fill?: string | null;
  bold?: boolean;
  italic?: boolean;
  strike?: boolean;
  underline?: Underline;
  /** Font colour as "#RRGGBB". */
  color?: string;
  left?: BorderEdge;
  right?: BorderEdge;
  top?: BorderEdge;
  bottom?: BorderEdge;
  numFmt?: { id: number; code: string };
}

/** A data bar. Positions are fractions of the width a bar may take, measured from the side the bars start at. */
export interface CellDataBar {
  from: number;
  to: number;
  /** "#RRGGBB". */
  color: string;
  /** The bar fades out towards its end instead of being one colour. */
  gradient: boolean;
  border?: string;
  /** The value is below zero and the bar points away from the axis. */
  negative: boolean;
  /** Where the axis is drawn, when the rule has one. */
  axis?: number;
  axisColor?: string;
  rightToLeft: boolean;
}

export interface CellIcon {
  /** Name of a set in ICON_SETS. */
  set: string;
  /** Position in the set, 0 for the icon of the lowest values. */
  index: number;
}

/** Everything conditional formats add to one cell. */
export interface CellConditional {
  /** The same object for every cell that gets the same combination of formats, so that it can key a cache. */
  style?: ConditionalStyle;
  /** Fill of a colour scale; when present it is the fill of the cell. */
  fill?: string;
  bar?: CellDataBar;
  icon?: CellIcon;
  /** A data bar or icon set that shows no value: the text of the cell is not drawn. */
  hideValue?: boolean;
}

/** What the grid renderer asks for. */
export interface ConditionalSource {
  /** Called once before the cells of a frame are asked for. */
  begin?(): void;
  resolve(row: number, col: number): CellConditional | undefined;
  /**
   * Runs `run` - the resolve() calls of one frame, say - so that the formulas calculated in
   * it share what the engine works out for them. The model must not change inside.
   */
  batch?<T>(run: () => T): T;
}

export interface ConditionalOptions {
  /** The engine that calculates the formulas of rules; none, or a function that returns undefined, turns such rules off. */
  engine?: () => CalcEngine | undefined;
  /** The current date and time as a serial number, for the "time period" rules. The engine's clock, else the system clock, when omitted. */
  now?: () => number;
  /**
   * Milliseconds of evaluation between two begin() calls. Cells not reached come back
   * without a format and `pending` is set, so that the caller can draw again. 40 when omitted.
   */
  budget?: number;
  /** Milliseconds from any fixed point; Date.now when omitted. */
  clock?: () => number;
}

// ---------------------------------------------------------------------------
// Icon sets

export type IconShape = 'arrow' | 'circle' | 'rimCircle' | 'flag' | 'diamond' | 'triangle' | 'triangleDown' | 'dash' | 'check' | 'cross' | 'exclamation' | 'checkCircle' | 'crossCircle' | 'exclamationCircle' | 'bars' | 'pie' | 'star' | 'boxes';

export interface IconGlyph {
  shape: IconShape;
  /** "#RRGGBB". */
  color: string;
  /**
   * arrow: direction in degrees, 0 pointing up, clockwise. bars and boxes: how many of
   * four are filled. pie: quarters filled. star: 0 empty, 1 half, 2 full.
   */
  level: number;
}

const RED = '#D65532';
const YELLOW = '#EAC282';
const GREEN = '#68A490';
const GRAY = '#8C8C8C';
const DARK = '#505050';
const BLACK = '#3A3A3A';
const PINK = '#E9A8A0';
const BLUE = '#4472C4';
const GOLD = '#EAC282';

const glyph = (shape: IconShape, color: string, level = 0): IconGlyph => ({ shape, color, level });
const arrows = (colors: string[], angles: number[]): IconGlyph[] => angles.map((angle, i) => glyph('arrow', colors[i], angle));
const levels = (shape: IconShape, color: string, values: number[]): IconGlyph[] => values.map((level) => glyph(shape, color, level));

/** The built-in icon sets, each from the icon of the lowest values to the icon of the highest. */
export const ICON_SETS: ReadonlyMap<string, readonly IconGlyph[]> = new Map<string, IconGlyph[]>([
  ['3Arrows', arrows([RED, YELLOW, GREEN], [180, 90, 0])],
  ['3ArrowsGray', arrows([GRAY, GRAY, GRAY], [180, 90, 0])],
  ['3Flags', [glyph('flag', RED), glyph('flag', YELLOW), glyph('flag', GREEN)]],
  ['3TrafficLights1', [glyph('circle', RED), glyph('circle', YELLOW), glyph('circle', GREEN)]],
  ['3TrafficLights2', [glyph('rimCircle', RED), glyph('rimCircle', YELLOW), glyph('rimCircle', GREEN)]],
  ['3Signs', [glyph('diamond', RED), glyph('triangle', YELLOW), glyph('circle', GREEN)]],
  ['3Symbols', [glyph('crossCircle', RED), glyph('exclamationCircle', YELLOW), glyph('checkCircle', GREEN)]],
  ['3Symbols2', [glyph('cross', RED), glyph('exclamation', YELLOW), glyph('check', GREEN)]],
  ['3Stars', levels('star', GOLD, [0, 1, 2])],
  ['3Triangles', [glyph('triangleDown', RED), glyph('dash', YELLOW), glyph('triangle', GREEN)]],
  ['4Arrows', arrows([RED, YELLOW, YELLOW, GREEN], [180, 135, 45, 0])],
  ['4ArrowsGray', arrows([GRAY, GRAY, GRAY, GRAY], [180, 135, 45, 0])],
  ['4RedToBlack', [glyph('circle', BLACK), glyph('circle', GRAY), glyph('circle', PINK), glyph('circle', RED)]],
  ['4Rating', levels('bars', BLUE, [1, 2, 3, 4])],
  ['4TrafficLights', [glyph('circle', BLACK), glyph('circle', RED), glyph('circle', YELLOW), glyph('circle', GREEN)]],
  ['5Arrows', arrows([RED, YELLOW, YELLOW, YELLOW, GREEN], [180, 135, 90, 45, 0])],
  ['5ArrowsGray', arrows([GRAY, GRAY, GRAY, GRAY, GRAY], [180, 135, 90, 45, 0])],
  ['5Rating', levels('bars', BLUE, [0, 1, 2, 3, 4])],
  ['5Quarters', levels('pie', DARK, [0, 1, 2, 3, 4])],
  ['5Boxes', levels('boxes', BLUE, [0, 1, 2, 3, 4])],
]);

/** The icon an evaluated cell shows; undefined for a set or position that does not exist. */
export function iconGlyph(icon: CellIcon): IconGlyph | undefined {
  const set = ICON_SETS.get(icon.set);
  return set ? set[icon.index] : undefined;
}

// ---------------------------------------------------------------------------
// Values

/** Stands for a cell that holds an error. */
const ERROR_CELL = { error: true };
type CellValue = number | string | boolean | null | typeof ERROR_CELL;

/** Returned instead of a result when the time for formulas of this frame is used up. */
const PENDING = { pending: true };
type Pending = typeof PENDING;

/** The value of a cell as rules see it: null for an empty cell, a serial number for an ISO date. */
function valueOfCell(cell: Cell | undefined, date1904: boolean): CellValue {
  if (!cell) return null;
  const value = cell.value;
  if (value === null || value === undefined) return null;
  switch (cell.kind) {
    case 'number':
      return typeof value === 'number' && isFinite(value) ? value : null;
    case 'string':
      return String(value);
    case 'boolean':
      return value === true;
    case 'error':
      return ERROR_CELL;
    case 'date': {
      const serial = isoToSerial(String(value), date1904);
      return serial === undefined ? String(value) : serial;
    }
  }
  return null;
}

function textOfValue(value: CellValue): string {
  if (value === null || value === ERROR_CELL) return '';
  if (typeof value === 'number') return numberToText(value);
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  return value as string;
}

const NOT_CONSTANT = { formula: true };

/**
 * A formula text that is only a number, a quoted text or TRUE/FALSE, read without an engine.
 * Numbers are read by textToNumber(), a hand-written scanner that looks at each character
 * once: texts come from the file and can be a megabyte of digits ("12%" is 0.12, as it is
 * in a formula).
 */
function constantOf(text: string): Scalar | typeof NOT_CONSTANT {
  const number = textToNumber(text);
  if (number !== undefined) return number;
  const trimmed = text.trim();
  if (trimmed.length >= 2 && trimmed.charCodeAt(0) === 34 && trimmed.charCodeAt(trimmed.length - 1) === 34) {
    const inner = trimmed.slice(1, -1);
    // Every quote inside has to be doubled; otherwise this is an expression such as "a"&"b".
    if (!/"/.test(inner.replace(/""/g, ''))) return inner.replace(/""/g, '"');
    return NOT_CONSTANT;
  }
  const upper = trimmed.toUpperCase();
  if (upper === 'TRUE') return true;
  if (upper === 'FALSE') return false;
  return NOT_CONSTANT;
}

/** TRIM takes away the space character only: a text of nothing else is blank, one with a tab or a no-break space is not. */
function isBlankText(text: string): boolean {
  for (let i = 0; i < text.length; i++) if (text.charCodeAt(i) !== 32) return false;
  return true;
}

/**
 * Whether `formula` is what Excel writes for a text rule of this kind and text, applied to
 * the cell `cell` ("A1"). Then the rule can be answered from the text alone.
 */
function isStandardTextFormula(kind: number, text: string, formula: string, cell: string): boolean {
  const quoted = '"' + text.replace(/"/g, '""') + '"';
  const written = formula.trim();
  switch (kind) {
    case K_CONTAINS_TEXT:
      return written === `NOT(ISERROR(SEARCH(${quoted},${cell})))`;
    case K_NOT_CONTAINS_TEXT:
      return written === `ISERROR(SEARCH(${quoted},${cell}))`;
    case K_BEGINS_WITH:
      return written === `LEFT(${cell},${text.length})=${quoted}` || written === `LEFT(${cell},LEN(${quoted}))=${quoted}`;
    case K_ENDS_WITH:
      return written === `RIGHT(${cell},${text.length})=${quoted}` || written === `RIGHT(${cell},LEN(${quoted}))=${quoted}`;
  }
  return false;
}

/** Longest text a wildcard pattern is matched against, and longest pattern. */
const MAX_MATCH_TEXT = 32767;
const MAX_PATTERN = 255;

/**
 * Finds `pattern` anywhere in `text` the way SEARCH does: without regard to case, "?" for
 * one character, "*" for any run, "~" before a character that is meant literally.
 */
function searchText(text: string, pattern: string): boolean {
  if (pattern === '') return true;
  if (text.length > MAX_MATCH_TEXT) text = text.slice(0, MAX_MATCH_TEXT);
  if (pattern.length > MAX_PATTERN) pattern = pattern.slice(0, MAX_PATTERN);
  return searchFolded(foldCase(text, true), foldCase(pattern, true)) || searchFolded(foldCase(text, false), foldCase(pattern, false));
}

/**
 * Upper or lower case, letter by letter and without a locale. A letter whose other case is
 * not a single letter ("İ" in lower case, "ß" in upper case) stays as it is, so that the
 * text keeps its length and a position in it means the same before and after.
 */
function foldCase(text: string, upper: boolean): string {
  const whole = upper ? text.toUpperCase() : text.toLowerCase();
  if (whole.length === text.length) return whole;
  let out = '';
  for (let i = 0; i < text.length; i++) {
    const letter = text[i];
    const folded = upper ? letter.toUpperCase() : letter.toLowerCase();
    out += folded.length === 1 ? folded : letter;
  }
  return out;
}

/**
 * Whether two texts are the same without regard to case.
 *
 * UNSURE: Excel compares with a simple case mapping that does not depend on the language of
 * the system. Which one exactly is not documented; here two texts are equal when their upper
 * case forms or their lower case forms are, which makes "i", "I" and the Turkish dotless "ı"
 * one letter and leaves the dotted capital "İ" on its own.
 */
function sameFolded(a: string, b: string): boolean {
  if (a.length !== b.length) return false;
  return foldCase(a, true) === foldCase(b, true) || foldCase(a, false) === foldCase(b, false);
}

/** searchText() for a text and a pattern already in one case. */
function searchFolded(haystack: string, needle: string): boolean {
  if (!/[?*~]/.test(needle)) return haystack.indexOf(needle) >= 0;
  // Pieces between the stars are looked for one after the other; that never backtracks over a star.
  const pieces: { chars: string[]; any: boolean[] }[] = [{ chars: [], any: [] }];
  for (let i = 0; i < needle.length; i++) {
    const c = needle[i];
    const piece = pieces[pieces.length - 1];
    if (c === '*') pieces.push({ chars: [], any: [] });
    else if (c === '?') {
      piece.chars.push('');
      piece.any.push(true);
    } else {
      piece.chars.push(c === '~' && i + 1 < needle.length ? needle[++i] : c);
      piece.any.push(false);
    }
  }
  let from = 0;
  for (let p = 0; p < pieces.length; p++) {
    const { chars, any } = pieces[p];
    if (chars.length === 0) continue;
    let found = -1;
    for (let start = from; start + chars.length <= haystack.length; start++) {
      let k = 0;
      while (k < chars.length && (any[k] || haystack[start + k] === chars[k])) k++;
      if (k === chars.length) {
        found = start;
        break;
      }
    }
    if (found < 0) return false;
    from = found + chars.length;
  }
  return true;
}

// ---------------------------------------------------------------------------
// Dates

/** A serial number of either date system as a serial of the 1900 system, whole days. */
function dayOf(serial: number, date1904: boolean): number {
  return Math.floor(serial) + (date1904 ? 1462 : 0);
}

/** Months since year 0 of a 1900-system day. */
function monthOf(day: number): number {
  const date = new Date((day - 25569) * 86400000);
  return date.getUTCFullYear() * 12 + date.getUTCMonth();
}

function inTimePeriod(period: string | undefined, day: number, today: number): boolean {
  // Day 1 of the 1900 system counts as a Sunday; weeks run from Sunday to Saturday.
  const weekStart = today - (((today - 1) % 7) + 7) % 7;
  switch (period) {
    case 'today':
      return day === today;
    case 'yesterday':
      return day === today - 1;
    case 'tomorrow':
      return day === today + 1;
    case 'last7Days':
      return day <= today && day >= today - 6;
    case 'thisWeek':
      return day >= weekStart && day <= weekStart + 6;
    case 'lastWeek':
      return day >= weekStart - 7 && day <= weekStart - 1;
    case 'nextWeek':
      return day >= weekStart + 7 && day <= weekStart + 13;
    case 'thisMonth':
      return monthOf(day) === monthOf(today);
    case 'lastMonth':
      return monthOf(day) === monthOf(today) - 1;
    case 'nextMonth':
      return monthOf(day) === monthOf(today) + 1;
  }
  return false;
}

// ---------------------------------------------------------------------------
// Formats and colours

/** Colour of the first stop of a gradient fill, read from the element as kept by the style reader. */
function gradientColor(fill: Extract<Fill, { type: 'gradient' }>): Color | undefined {
  const raw = fill.raw;
  if (!raw || !Array.isArray(raw.children)) return undefined;
  for (const stop of raw.children) {
    if (stop.type !== 'element' || !Array.isArray(stop.children)) continue;
    for (const child of stop.children) {
      if (child.type !== 'element' || !child.attrs) continue;
      const rgb = child.attrs.get('rgb');
      if (rgb !== undefined) return { rgb: rgb.trim().toUpperCase() };
      const theme = child.attrs.get('theme');
      if (theme !== undefined && /^\d{1,3}$/.test(theme)) {
        const tint = Number(child.attrs.get('tint'));
        return tint === tint ? { theme: Number(theme), tint: Math.max(-1, Math.min(1, tint)) } : { theme: Number(theme) };
      }
    }
  }
  return undefined;
}

/**
 * The colour a differential fill paints a cell with: "#RRGGBB", null for "no fill",
 * undefined when it says nothing usable. Unlike a cell fill, a solid differential fill
 * keeps its colour in `bg`. Patterns are drawn as their background, gradients as their
 * first colour.
 */
export function differentialFillColor(fill: Fill | undefined, styles: StyleSheet): string | null | undefined {
  if (!fill) return undefined;
  if (fill.type === 'gradient') return resolveColor(gradientColor(fill), styles);
  if (fill.pattern === 'none') return null;
  return resolveColor(fill.bg, styles) || resolveColor(fill.fg, styles);
}

function styleOfFormat(format: DifferentialFormat | undefined, styles: StyleSheet): ConditionalStyle | undefined {
  if (!format) return undefined;
  const style: ConditionalStyle = {};
  const font = format.font;
  if (font) {
    if (font.bold !== undefined) style.bold = font.bold === true;
    if (font.italic !== undefined) style.italic = font.italic === true;
    if (font.strike !== undefined) style.strike = font.strike === true;
    if (font.underline !== undefined) style.underline = font.underline;
    const color = resolveColor(font.color, styles);
    if (color !== undefined) style.color = color;
  }
  const fill = differentialFillColor(format.fill, styles);
  if (fill !== undefined) style.fill = fill;
  const border = format.border;
  if (border) {
    if (border.left) style.left = border.left;
    if (border.right) style.right = border.right;
    if (border.top) style.top = border.top;
    if (border.bottom) style.bottom = border.bottom;
  }
  if (format.numFmt && typeof format.numFmt.code === 'string') style.numFmt = format.numFmt;
  return style;
}

const STYLE_KEYS: (keyof ConditionalStyle)[] = ['fill', 'bold', 'italic', 'strike', 'underline', 'color', 'left', 'right', 'top', 'bottom', 'numFmt'];

/** Adds to `target` what `source` states and `target` does not: the first rule to state a property keeps it. */
function mergeStyle(target: ConditionalStyle, source: ConditionalStyle): void {
  const to = target as Record<string, unknown>;
  const from = source as Record<string, unknown>;
  for (let i = 0; i < STYLE_KEYS.length; i++) {
    const key = STYLE_KEYS[i];
    if (to[key] === undefined && from[key] !== undefined) to[key] = from[key];
  }
}

function channels(color: string): number[] {
  return [parseInt(color.slice(1, 3), 16), parseInt(color.slice(3, 5), 16), parseInt(color.slice(5, 7), 16)];
}

const hex2 = (value: number) => {
  const n = Math.max(0, Math.min(255, Math.round(value)));
  return (n < 16 ? '0' : '') + n.toString(16).toUpperCase();
};

function mixColors(a: number[], b: number[], t: number): string {
  return '#' + hex2(a[0] + (b[0] - a[0]) * t) + hex2(a[1] + (b[1] - a[1]) * t) + hex2(a[2] + (b[2] - a[2]) * t);
}

// ---------------------------------------------------------------------------
// Rules

const K_CELL_IS = 1;
const K_EXPRESSION = 2;
const K_COLOR_SCALE = 3;
const K_DATA_BAR = 4;
const K_ICON_SET = 5;
const K_TOP10 = 6;
const K_ABOVE_AVERAGE = 7;
const K_UNIQUE = 8;
const K_DUPLICATE = 9;
const K_CONTAINS_TEXT = 10;
const K_NOT_CONTAINS_TEXT = 11;
const K_BEGINS_WITH = 12;
const K_ENDS_WITH = 13;
const K_CONTAINS_BLANKS = 14;
const K_NOT_CONTAINS_BLANKS = 15;
const K_CONTAINS_ERRORS = 16;
const K_NOT_CONTAINS_ERRORS = 17;
const K_TIME_PERIOD = 18;

const KINDS = new Map<string, number>([
  ['cellIs', K_CELL_IS],
  ['expression', K_EXPRESSION],
  ['colorScale', K_COLOR_SCALE],
  ['dataBar', K_DATA_BAR],
  ['iconSet', K_ICON_SET],
  ['top10', K_TOP10],
  ['aboveAverage', K_ABOVE_AVERAGE],
  ['uniqueValues', K_UNIQUE],
  ['duplicateValues', K_DUPLICATE],
  ['containsText', K_CONTAINS_TEXT],
  ['notContainsText', K_NOT_CONTAINS_TEXT],
  ['beginsWith', K_BEGINS_WITH],
  ['endsWith', K_ENDS_WITH],
  ['containsBlanks', K_CONTAINS_BLANKS],
  ['notContainsBlanks', K_NOT_CONTAINS_BLANKS],
  ['containsErrors', K_CONTAINS_ERRORS],
  ['notContainsErrors', K_NOT_CONTAINS_ERRORS],
  ['timePeriod', K_TIME_PERIOD],
]);

/** Numbers of the cells a rule applies to. */
interface Stats {
  /** Ascending. */
  sorted: Float64Array;
  count: number;
  min: number;
  max: number;
  mean: number;
  /**
   * Standard deviation of the numbers as a whole population (STDEV.P).
   *
   * UNSURE: Excel does not document which deviation "1 std dev above average" uses; the
   * population one is what its results are reported to match.
   */
  deviation: number;
}

/** A rule with the ranges it applies to, after the rows and columns edited since the file was read. */
interface Entry {
  rule: ConditionalRule;
  kind: number;
  ranges: CellRange[];
  /**
   * The cell the formulas are written for: the top-left cell of the first range of the
   * list, whether or not another range starts further up or left. Excel keeps a rule
   * relative to the active cell of the selection it was made for and, when it writes the
   * file, relative to the first range of sqref; files of Excel and of openpyxl read that
   * way, and the writer moves formulas by the same corner.
   *
   * UNSURE: not stated by the standard; taken from files Excel writes.
   */
  anchorRow: number;
  anchorCol: number;
  formulas: string[];
  /** The formulas that are plain numbers or texts, already read; NOT_CONSTANT for the others. */
  constants: (Scalar | typeof NOT_CONSTANT)[];
  /** Position among the entries once sorted by priority. */
  order: number;
  /** What the rule's differential format states; undefined when it has none or names one that does not exist. */
  style: ConditionalStyle | undefined;
  /** As `style`, as the whole result for a cell only this rule formats. */
  alone: CellConditional | undefined;
  /** A text rule that is answered from its text attribute: it has no formula, or the formula Excel writes for that text. */
  textOnly: boolean;
  stats?: Stats;
  /** How often each value occurs, for the rules about unique and repeated values. */
  counts?: Map<string, number>;
  /** Thresholds of a colour scale, data bar or icon set; null when they cannot be worked out. */
  thresholds?: number[] | null;
  /** Limit of a top/bottom or average rule; NaN when there is none. */
  limit?: number;
  /** Colours of a colour scale as channels; null when the rule has too few. */
  scale?: number[][] | null;
}

/** Rows per block and columns per block of the index of rules, as shifts. */
const BLOCK_ROW_SHIFT = 7;
const BLOCK_COL_SHIFT = 4;
const BLOCK_COLS = MAX_COLS >> BLOCK_COL_SHIFT;
/** Blocks and cells kept; when one of the lists is full it is emptied. */
const MAX_BLOCKS = 4096;
const MAX_CACHED_CELLS = 200000;
const MAX_STYLE_COMBINATIONS = 4096;
const MAX_CONSTANTS = 4096;
const DEFAULT_BUDGET = 40;
/** Rules looked at between two looks at the clock. */
const CLOCK_EVERY = 4096;

const DEFAULT_BAR_COLOR = '#638EC6';
const DEFAULT_NEGATIVE_COLOR = '#FF0000';
const DEFAULT_AXIS_COLOR = '#000000';

const NO_ENTRIES: Entry[] = [];

function inRanges(ranges: CellRange[], row: number, col: number): boolean {
  for (let i = 0; i < ranges.length; i++) {
    const range = ranges[i];
    if (row >= range.top && row <= range.bottom && col >= range.left && col <= range.right) return true;
  }
  return false;
}

/**
 * What two cells have to share to be "the same value" for the rules about unique and
 * repeated values.
 *
 * UNSURE: Excel is believed to count the way COUNTIF does: a text that reads as a number is
 * that number ("1" repeats 1), and texts compare without regard to case.
 */
function countKey(value: CellValue): string | undefined {
  if (typeof value === 'number') return 'n' + value;
  if (typeof value === 'string') {
    if (value === '') return undefined;
    const number = textToNumber(value);
    return number !== undefined ? 'n' + number : 's' + value.toUpperCase();
  }
  if (typeof value === 'boolean') return value ? 'b1' : 'b0';
  return undefined;
}

/** The p-th percentile (0..100) of ascending numbers, interpolated between neighbours as PERCENTILE does. */
function percentile(sorted: Float64Array, p: number): number {
  const n = sorted.length;
  if (n === 0) return NaN;
  const position = (Math.max(0, Math.min(100, p)) / 100) * (n - 1);
  const low = Math.floor(position);
  const high = Math.min(low + 1, n - 1);
  return sorted[low] + (sorted[high] - sorted[low]) * (position - low);
}

/** Whether a threshold of the rule is given as something the engine has to calculate. */
function hasFormulaThreshold(rule: ConditionalRule): boolean {
  const calculated = (value: ConditionalValue | undefined) => !!value && typeof value.value === 'string' && value.value !== '' && constantOf(value.value) === NOT_CONSTANT;
  if (rule.colorScale && Array.isArray(rule.colorScale.values) && rule.colorScale.values.some(calculated)) return true;
  if (rule.iconSet && Array.isArray(rule.iconSet.values) && rule.iconSet.values.some(calculated)) return true;
  return !!rule.dataBar && (calculated(rule.dataBar.min) || calculated(rule.dataBar.max));
}

export class ConditionalFormats implements ConditionalSource {
  /** Set while cells were left without their format because the time for formulas ran out; cleared by begin(). */
  pending = false;

  private readonly workbook: Workbook;
  private readonly sheet: Sheet;
  private readonly options: ConditionalOptions;
  private entries: Entry[] | undefined;
  /** The structuralEdits lists of all sheets the entries were built for. */
  private editsSeen: (Sheet['structuralEdits'] | undefined)[] = [];
  private formatsSeen: ConditionalFormat[] | undefined;
  private bounds: CellRange = { top: 0, left: 0, bottom: -1, right: -1 };
  private blocks = new Map<number, Entry[]>();
  private cells = new Map<number, CellConditional | null>();
  private combined = new Map<string, CellConditional>();
  private constants = new Map<string, Scalar | typeof NOT_CONSTANT>();
  /** How far the relative references of a formula text reach, by the text. */
  private extents = new Map<string, ReturnType<typeof relativeExtent>>();
  /** Some rule needs the engine: the cells are then evaluated in a batch of it. */
  private usesFormulas = false;
  /** Formula texts the engine could not parse. */
  private unparsed = new Set<string>();
  private todayValue: number | undefined;
  private deadline = Infinity;
  /** Formulas calculated and cells finished since begin(). */
  private calculated = 0;
  private finished = 0;
  /** Rules looked at since begin(), and the count at which the clock is looked at next. */
  private work = 0;
  private nextCheck = CLOCK_EVERY;

  constructor(workbook: Workbook, sheet: Sheet, options: ConditionalOptions = {}) {
    this.workbook = workbook;
    this.sheet = sheet;
    this.options = options || {};
  }

  /** The sheet has at least one conditional format. */
  get active(): boolean {
    const formats = this.sheet.conditionalFormats;
    return Array.isArray(formats) && formats.length > 0;
  }

  /**
   * Forgets everything worked out from cells. Call after any change to the workbook: a cell
   * edit, a recalculation, rows or columns inserted or deleted, undo and redo.
   */
  invalidate(): void {
    this.cells.clear();
    this.todayValue = undefined;
    this.pending = false;
    const entries = this.entries;
    if (!entries) return;
    if (this.stale()) {
      this.entries = undefined;
      return;
    }
    for (let i = 0; i < entries.length; i++) {
      const entry = entries[i];
      entry.stats = undefined;
      entry.counts = undefined;
      entry.thresholds = undefined;
      entry.limit = undefined;
    }
  }

  /** Starts a frame: the time for formulas counts from here. */
  begin(): void {
    this.pending = false;
    this.calculated = 0;
    this.finished = 0;
    this.work = 0;
    this.nextCheck = CLOCK_EVERY;
    const budget = this.options.budget === undefined ? DEFAULT_BUDGET : Number(this.options.budget);
    this.deadline = budget > 0 && budget < Infinity ? this.clock() + budget : Infinity;
  }

  /** What the rules of the sheet give the cell; undefined when no rule applies to it. Never throws. */
  resolve(row: number, col: number): CellConditional | undefined {
    try {
      const entries = this.entries || this.build();
      if (entries.length === 0) return undefined;
      const bounds = this.bounds;
      if (!(row >= bounds.top && row <= bounds.bottom && col >= bounds.left && col <= bounds.right)) return undefined;
      const key = row * MAX_COLS + col;
      const known = this.cells.get(key);
      if (known !== undefined) return known === null ? undefined : known;
      const candidates = this.block(entries, row, col);
      if (candidates.length === 0) return undefined;
      // All the formulas of one cell are one batch for the engine; inside batch() of this
      // object the batches of the cells are part of the larger one.
      const engine = this.usesFormulas ? this.engine() : undefined;
      const result = engine && typeof engine.batch === 'function' ? engine.batch(() => this.evaluate(candidates, row, col)) : this.evaluate(candidates, row, col);
      if (result === PENDING) {
        this.pending = true;
        return undefined;
      }
      this.finished++;
      if (this.cells.size >= MAX_CACHED_CELLS) this.cells.clear();
      this.cells.set(key, result === undefined ? null : (result as CellConditional));
      return result as CellConditional | undefined;
    } catch (error) {
      return undefined;
    }
  }

  /**
   * Runs `run` as one batch of the engine: formulas calculated by the resolve() calls inside
   * share one calculation budget and what functions keep between calls (the index of a
   * lookup table). The model must not change while `run` runs. Without an engine, or with
   * one that has no batches, `run` is simply called.
   */
  batch<T>(run: () => T): T {
    const engine = this.engine();
    return engine && typeof engine.batch === 'function' ? engine.batch(run) : run();
  }

  // -------------------------------------------------------------------------
  // Rules and where they apply

  private clock(): number {
    const value = this.options.clock ? Number(this.options.clock()) : Date.now();
    return value === value ? value : 0;
  }

  private engine(): CalcEngine | undefined {
    try {
      return this.options.engine ? this.options.engine() : undefined;
    } catch (error) {
      return undefined;
    }
  }

  /** Rows or columns were inserted or deleted, or the formats replaced, since the entries were built. */
  private stale(): boolean {
    if (this.formatsSeen !== this.sheet.conditionalFormats) return true;
    const sheets = this.workbook.sheets || [];
    if (sheets.length !== this.editsSeen.length) return true;
    for (let i = 0; i < sheets.length; i++) if (sheets[i].structuralEdits !== this.editsSeen[i]) return true;
    return false;
  }

  private build(): Entry[] {
    const sheet = this.sheet;
    const styles = this.workbook.styles;
    const sheets = this.workbook.sheets || [];
    this.formatsSeen = sheet.conditionalFormats;
    this.editsSeen = sheets.map((other) => other.structuralEdits);
    this.blocks.clear();
    this.cells.clear();
    this.combined.clear();
    this.usesFormulas = false;

    const own: StructuralEdit[] = (sheet.structuralEdits || []).map((edit) => ({ ...edit, sheetName: sheet.name }));
    const foreign: StructuralEdit[] = [];
    for (const other of sheets) {
      if (other === sheet || !other.structuralEdits) continue;
      for (const edit of other.structuralEdits) foreign.push({ ...edit, sheetName: other.name });
    }

    const entries: Entry[] = [];
    const formats = Array.isArray(sheet.conditionalFormats) ? sheet.conditionalFormats : [];
    for (let f = 0; f < formats.length; f++) {
      const format = formats[f];
      if (!format || !Array.isArray(format.ranges) || !Array.isArray(format.rules)) continue;
      let ranges = format.ranges.filter((range) => !!range && range.top <= range.bottom && range.left <= range.right);
      if (ranges.length === 0) continue;
      // Formulas are carried through the edits rule by rule, the ranges once for all rules.
      const texts = format.rules.map((rule) => (rule && Array.isArray(rule.formulas) ? rule.formulas.map(String) : []));
      const adjust = (edit: StructuralEdit, from: CellRange, to: CellRange) => {
        const before = { row: from.top, col: from.left };
        const after = { row: to.top, col: to.left };
        // Relative references go round the edge of the sheet, as they do when a rule is evaluated.
        for (const list of texts) for (let i = 0; i < list.length; i++) list[i] = adjustAnchoredFormula(list[i], sheet.name, edit, before, after, true);
      };
      for (const edit of own) {
        const moved: CellRange[] = [];
        for (const range of ranges) {
          const next = adjustRange(range, edit);
          if (next) moved.push(next);
        }
        if (moved.length === 0) {
          ranges = moved;
          break;
        }
        adjust(edit, ranges[0], moved[0]);
        ranges = moved;
      }
      if (ranges.length === 0) continue;
      for (const edit of foreign) adjust(edit, ranges[0], ranges[0]);

      for (let r = 0; r < format.rules.length; r++) {
        const rule = format.rules[r];
        const kind = rule ? KINDS.get(rule.type) : undefined;
        if (!rule || kind === undefined) continue;
        const dxfs = styles.dxfFormats;
        const style = styleOfFormat(rule.dxf || (rule.dxfId !== undefined && dxfs ? dxfs[rule.dxfId] : undefined), styles);
        const moved = own.length > 0 || foreign.length > 0 ? this.movedRule(rule, own.concat(foreign)) : rule;
        const constants = texts[r].map(constantOf);
        const textRule = kind === K_CONTAINS_TEXT || kind === K_NOT_CONTAINS_TEXT || kind === K_BEGINS_WITH || kind === K_ENDS_WITH;
        const textOnly =
          textRule && typeof rule.text === 'string' && (texts[r].length === 0 || isStandardTextFormula(kind, rule.text, texts[r][0], cellName(ranges[0].top, ranges[0].left)));
        if (!this.usesFormulas && !textOnly && (constants.indexOf(NOT_CONSTANT) >= 0 || hasFormulaThreshold(moved))) this.usesFormulas = true;
        entries.push({
          rule: moved,
          kind,
          ranges,
          anchorRow: ranges[0].top,
          anchorCol: ranges[0].left,
          formulas: texts[r],
          constants,
          order: entries.length,
          style,
          alone: style ? { style } : undefined,
          textOnly,
        });
      }
    }
    // Lower priority numbers first; rules without one keep the order of the file, after the others.
    entries.sort((a, b) => {
      const pa = a.rule.priority >= 0 ? a.rule.priority : Infinity;
      const pb = b.rule.priority >= 0 ? b.rule.priority : Infinity;
      return pa === pb ? a.order - b.order : pa < pb ? -1 : 1;
    });
    const bounds: CellRange = { top: Infinity, left: Infinity, bottom: -1, right: -1 };
    for (let i = 0; i < entries.length; i++) {
      entries[i].order = i;
      for (const range of entries[i].ranges) {
        if (range.top < bounds.top) bounds.top = range.top;
        if (range.left < bounds.left) bounds.left = range.left;
        if (range.bottom > bounds.bottom) bounds.bottom = range.bottom;
        if (range.right > bounds.right) bounds.right = range.right;
      }
    }
    this.bounds = bounds;
    this.entries = entries;
    return entries;
  }

  /** A rule whose thresholds are formulas, with those formulas following the edits. */
  private movedRule(rule: ConditionalRule, edits: StructuralEdit[]): ConditionalRule {
    const move = (values: ConditionalValue[]): ConditionalValue[] =>
      values.map((value) => {
        if (!value || typeof value.value !== 'string' || textToNumber(value.value) !== undefined) return value;
        let text = value.value;
        for (const edit of edits) text = adjustFormula(text, this.sheet.name, edit);
        return text === value.value ? value : { ...value, value: text };
      });
    if (rule.colorScale && Array.isArray(rule.colorScale.values)) return { ...rule, colorScale: { ...rule.colorScale, values: move(rule.colorScale.values) } };
    if (rule.iconSet && Array.isArray(rule.iconSet.values)) return { ...rule, iconSet: { ...rule.iconSet, values: move(rule.iconSet.values) } };
    if (rule.dataBar) {
      const [min, max] = move([rule.dataBar.min, rule.dataBar.max]);
      return { ...rule, dataBar: { ...rule.dataBar, min, max } };
    }
    return rule;
  }

  /** The rules that reach into the block of cells around (row, col), in order of priority. */
  private block(entries: Entry[], row: number, col: number): Entry[] {
    const blockRow = row >> BLOCK_ROW_SHIFT;
    const blockCol = col >> BLOCK_COL_SHIFT;
    const key = blockRow * BLOCK_COLS + blockCol;
    let found = this.blocks.get(key);
    if (found) return found;
    const top = blockRow << BLOCK_ROW_SHIFT;
    const left = blockCol << BLOCK_COL_SHIFT;
    const bottom = top + (1 << BLOCK_ROW_SHIFT) - 1;
    const right = left + (1 << BLOCK_COL_SHIFT) - 1;
    found = [];
    for (let i = 0; i < entries.length; i++) {
      const ranges = entries[i].ranges;
      for (let k = 0; k < ranges.length; k++) {
        const range = ranges[k];
        if (range.bottom < top || range.top > bottom || range.right < left || range.left > right) continue;
        found.push(entries[i]);
        break;
      }
    }
    if (found.length === 0) found = NO_ENTRIES;
    if (this.blocks.size >= MAX_BLOCKS) this.blocks.clear();
    this.blocks.set(key, found);
    return found;
  }

  // -------------------------------------------------------------------------
  // One cell

  private evaluate(candidates: Entry[], row: number, col: number): CellConditional | Pending | undefined {
    const modelRow = this.sheet.rows.get(row);
    const value = valueOfCell(modelRow ? modelRow.cells[col] : undefined, this.workbook.date1904 === true);

    // Rules with a format that matched, in order of priority.
    let first: Entry | undefined;
    let more: Entry[] | undefined;
    let fill: string | undefined;
    /** A format stated a fill before a colour scale was reached. */
    let fillTaken = false;
    let bar: CellDataBar | undefined;
    let icon: CellIcon | undefined;
    let hideValue = false;
    let any = false;

    for (let i = 0; i < candidates.length; i++) {
      const entry = candidates[i];
      if (++this.work >= this.nextCheck) {
        // A sheet with a great many rules is spread over frames as well; every frame finishes a cell at least.
        this.nextCheck = this.work + CLOCK_EVERY;
        if (this.finished > 0 && this.deadline < Infinity && this.clock() > this.deadline) return PENDING;
      }
      if (!inRanges(entry.ranges, row, col)) continue;
      const kind = entry.kind;
      if (kind === K_COLOR_SCALE) {
        if (fill !== undefined || fillTaken) continue;
        const color = this.scaleColor(entry, value);
        if (color === PENDING) return PENDING;
        if (color !== undefined) {
          fill = color as string;
          any = true;
        }
        continue;
      }
      if (kind === K_DATA_BAR) {
        if (bar) continue;
        const found = this.dataBar(entry, value);
        if (found === PENDING) return PENDING;
        if (found) {
          bar = found as CellDataBar;
          if (entry.rule.dataBar && entry.rule.dataBar.showValue === false) hideValue = true;
          any = true;
        }
        continue;
      }
      if (kind === K_ICON_SET) {
        if (icon) continue;
        const found = this.icon(entry, value);
        if (found === PENDING) return PENDING;
        if (found) {
          const shown = found as CellIcon | null;
          // A cell that takes part in the set shows no value even when its own icon is "no icon".
          if (entry.rule.iconSet && entry.rule.iconSet.showValue === false) hideValue = true;
          if (shown && shown.index >= 0) icon = shown;
          any = true;
        }
        continue;
      }
      const matched = this.matches(entry, value, row, col);
      if (matched === PENDING) return PENDING;
      if (!matched) continue;
      if (entry.style) {
        if (!first) first = entry;
        else (more || (more = [first])).push(entry);
        if (entry.style.fill !== undefined && fill === undefined) fillTaken = true;
        any = true;
      }
      if (entry.rule.stopIfTrue) break;
    }
    if (!any) return undefined;

    if (first && !more && fill === undefined && !bar && !icon && !hideValue) return first.alone;
    const result: CellConditional = {};
    if (first) result.style = more ? this.combine(more) : first.style;
    if (fill !== undefined) result.fill = fill;
    if (bar) result.bar = bar;
    if (icon) result.icon = icon;
    if (hideValue) result.hideValue = true;
    return result;
  }

  /** The formats of several rules as one; the same object for the same rules. */
  private combine(entries: Entry[]): ConditionalStyle {
    let key = '';
    for (let i = 0; i < entries.length; i++) key += entries[i].order + ',';
    let found = this.combined.get(key);
    if (!found) {
      const style: ConditionalStyle = {};
      for (let i = 0; i < entries.length; i++) mergeStyle(style, entries[i].style as ConditionalStyle);
      found = { style };
      if (this.combined.size >= MAX_STYLE_COMBINATIONS) this.combined.clear();
      this.combined.set(key, found);
    }
    return found.style as ConditionalStyle;
  }

  /**
   * The value of a formula of a rule for the cell at (row, col): relative references move
   * with the cell. undefined when it cannot be calculated (no engine, not a formula).
   * `index` is the position among the rule's formulas.
   */
  private operand(entry: Entry, index: number, row: number, col: number): Scalar | Pending | undefined {
    const constant = entry.constants[index];
    if (constant !== undefined && constant !== NOT_CONSTANT) return constant as Scalar;
    return this.formula(entry, entry.formulas[index], row, col);
  }

  /** The same for any formula text of the rule (a threshold, say). */
  private formula(entry: Entry, text: string | undefined, row: number, col: number): Scalar | Pending | undefined {
    if (typeof text !== 'string' || text === '') return undefined;
    let constant = this.constants.get(text);
    if (constant === undefined) {
      constant = constantOf(text);
      if (this.constants.size >= MAX_CONSTANTS) this.constants.clear();
      this.constants.set(text, constant);
    }
    if (constant !== NOT_CONSTANT) return constant as Scalar;
    if (this.unparsed.has(text)) return undefined;
    const engine = this.engine();
    if (!engine) return undefined;
    // At least one formula is calculated per frame, so that every frame gets somewhere.
    if (this.calculated > 0 && this.deadline < Infinity && this.clock() > this.deadline) return PENDING;
    this.calculated++;
    try {
      const dRow = row - entry.anchorRow;
      const dCol = col - entry.anchorCol;
      const sheet = this.sheetIndex();
      // The engine moves the references itself, from one parse of the text, as long as none
      // of them leaves the sheet: those it would make #REF!, while in a conditional format
      // they come back in at the opposite edge (the cell above row 1 is the last row).
      // An engine without evaluateShifted() gets the moved text.
      if (typeof engine.evaluateShifted === 'function' && ((dRow === 0 && dCol === 0) || this.staysOnSheet(text, dRow, dCol))) {
        return engine.evaluateShifted(text, { sheet, row: entry.anchorRow, col: entry.anchorCol }, { row, col });
      }
      return engine.evaluateText(shiftFormula(text, dRow, dCol, true), { sheet, row, col });
    } catch (error) {
      if (this.unparsed.size >= MAX_CONSTANTS) this.unparsed.clear();
      this.unparsed.add(text);
      return undefined;
    }
  }

  /** Whether every relative reference of the formula is still on the sheet once moved by (dRow, dCol). */
  private staysOnSheet(text: string, dRow: number, dCol: number): boolean {
    let extent = this.extents.get(text);
    if (!extent) {
      extent = relativeExtent(text);
      if (this.extents.size >= MAX_CONSTANTS) this.extents.clear();
      this.extents.set(text, extent);
    }
    return extent.minRow + dRow >= 0 && extent.maxRow + dRow < MAX_ROWS && extent.minCol + dCol >= 0 && extent.maxCol + dCol < MAX_COLS;
  }

  private sheetIndex(): number {
    const index = (this.workbook.sheets || []).indexOf(this.sheet);
    return index < 0 ? 0 : index;
  }

  /** Today as a whole day of the 1900 date system. */
  private today(): number {
    if (this.todayValue === undefined) {
      const date1904 = this.workbook.date1904 === true;
      let serial: number | undefined;
      if (this.options.now) serial = Number(this.options.now());
      else {
        const engine = this.engine();
        if (engine) {
          try {
            const value = engine.evaluateText('TODAY()', { sheet: this.sheetIndex(), row: 0, col: 0 });
            if (typeof value === 'number') serial = value;
          } catch (error) {
            serial = undefined;
          }
        }
        if (serial === undefined) {
          const date = new Date();
          serial = (date.getTime() - date.getTimezoneOffset() * 60000) / 86400000 + 25569 - (date1904 ? 1462 : 0);
        }
      }
      this.todayValue = serial === serial && isFinite(serial as number) ? dayOf(serial as number, date1904) : 0;
    }
    return this.todayValue;
  }

  /** Whether a rule that carries a format applies to the cell. */
  private matches(entry: Entry, value: CellValue, row: number, col: number): boolean | Pending {
    const rule = entry.rule;
    switch (entry.kind) {
      case K_CELL_IS:
        return this.cellIs(entry, value, row, col);
      case K_EXPRESSION: {
        const result = this.operand(entry, 0, row, col);
        if (result === PENDING) return PENDING;
        return typeof result === 'number' ? result !== 0 : result === true;
      }
      case K_CONTAINS_TEXT:
      case K_NOT_CONTAINS_TEXT:
      case K_BEGINS_WITH:
      case K_ENDS_WITH: {
        if (!entry.textOnly) {
          // Excel evaluates the formula; the text attribute is what its dialog shows. They
          // differ in files other programs wrote, and in rules edited as formulas.
          const result = this.operand(entry, 0, row, col);
          if (result === PENDING) return PENDING;
          // No engine, or a formula it cannot read: the text attribute is all there is.
          if (result !== undefined || typeof rule.text !== 'string') return typeof result === 'number' ? result !== 0 : result === true;
        }
        const wanted = rule.text as string;
        if (value === ERROR_CELL) return entry.kind === K_NOT_CONTAINS_TEXT;
        const text = textOfValue(value);
        if (entry.kind === K_CONTAINS_TEXT) return searchText(text, wanted);
        if (entry.kind === K_NOT_CONTAINS_TEXT) return !searchText(text, wanted);
        if (entry.kind === K_BEGINS_WITH) return sameFolded(text.slice(0, wanted.length), wanted);
        return sameFolded(wanted.length === 0 ? '' : text.slice(-wanted.length), wanted);
      }
      // As LEN(TRIM(cell))=0, which Excel writes for these rules.
      case K_CONTAINS_BLANKS:
        return value !== ERROR_CELL && isBlankText(textOfValue(value));
      case K_NOT_CONTAINS_BLANKS:
        return value !== ERROR_CELL && !isBlankText(textOfValue(value));
      case K_CONTAINS_ERRORS:
        return value === ERROR_CELL;
      case K_NOT_CONTAINS_ERRORS:
        return value !== ERROR_CELL;
      case K_TIME_PERIOD:
        return typeof value === 'number' && inTimePeriod(rule.timePeriod, dayOf(value, this.workbook.date1904 === true), this.today());
      case K_TOP10: {
        if (typeof value !== 'number') return false;
        const limit = this.limit(entry);
        return limit === limit && (rule.bottom ? value <= limit : value >= limit);
      }
      case K_ABOVE_AVERAGE: {
        if (typeof value !== 'number') return false;
        const limit = this.limit(entry);
        if (limit !== limit) return false;
        // With a number of deviations the limit itself never counts.
        const inclusive = rule.equalAverage === true && !(rule.stdDev !== undefined && rule.stdDev > 0);
        if (rule.aboveAverage === false) return inclusive ? value <= limit : value < limit;
        return inclusive ? value >= limit : value > limit;
      }
      case K_UNIQUE:
      case K_DUPLICATE: {
        const key = countKey(value);
        if (key === undefined) return false;
        const count = this.counts(entry).get(key) || 0;
        return entry.kind === K_UNIQUE ? count === 1 : count > 1;
      }
    }
    return false;
  }

  private cellIs(entry: Entry, value: CellValue, row: number, col: number): boolean | Pending {
    if (value === ERROR_CELL) return false;
    const first = this.operand(entry, 0, row, col);
    if (first === PENDING) return PENDING;
    if (first === undefined || isError(first)) return false;
    const a = compareScalars(value as Scalar, first as Scalar);
    if (typeof a !== 'number') return false;
    const operator = entry.rule.operator;
    switch (operator) {
      case 'equal':
        return a === 0;
      case 'notEqual':
        return a !== 0;
      case 'greaterThan':
        return a > 0;
      case 'greaterThanOrEqual':
        return a >= 0;
      case 'lessThan':
        return a < 0;
      case 'lessThanOrEqual':
        return a <= 0;
      case 'between':
      case 'notBetween': {
        const second = this.operand(entry, 1, row, col);
        if (second === PENDING) return PENDING;
        if (second === undefined || isError(second)) return false;
        const b = compareScalars(value as Scalar, second as Scalar);
        if (typeof b !== 'number') return false;
        // The two limits may come in either order.
        const inside = (a >= 0 && b <= 0) || (a <= 0 && b >= 0);
        return operator === 'between' ? inside : !inside;
      }
    }
    return false;
  }

  // -------------------------------------------------------------------------
  // What a rule knows about all of its cells

  /**
   * Calls `visit` for every cell the rule applies to that exists in the sheet, once: a cell
   * that several ranges of the rule name is one cell of it.
   */
  private eachCell(entry: Entry, visit: (cell: Cell) => void): void {
    const rows = this.sheet.rows;
    // With one range no cell can come up twice.
    const seen = entry.ranges.length > 1 ? new Set<number>() : undefined;
    for (let k = 0; k < entry.ranges.length; k++) {
      const { top, left, bottom, right } = entry.ranges[k];
      const scan = (row: { cells: Cell[] }, r: number) => {
        const cells = row.cells;
        const last = Math.min(right, cells.length - 1);
        for (let c = left; c <= last; c++) {
          const cell = cells[c];
          if (!cell) continue;
          if (seen) {
            const key = r * MAX_COLS + c;
            if (seen.has(key)) continue;
            seen.add(key);
          }
          visit(cell);
        }
      };
      if (bottom - top + 1 <= rows.size) {
        for (let r = top; r <= bottom; r++) {
          const row = rows.get(r);
          if (row) scan(row, r);
        }
      } else {
        // A range taller than the sheet is full: only the rows that exist are looked at.
        rows.forEach((row, r) => {
          if (r >= top && r <= bottom) scan(row, r);
        });
      }
    }
  }

  private stats(entry: Entry): Stats {
    if (entry.stats) return entry.stats;
    const date1904 = this.workbook.date1904 === true;
    const numbers: number[] = [];
    this.eachCell(entry, (cell) => {
      const value = valueOfCell(cell, date1904);
      if (typeof value === 'number') numbers.push(value);
    });
    const sorted = Float64Array.from(numbers).sort();
    const count = sorted.length;
    let sum = 0;
    for (let i = 0; i < count; i++) sum += sorted[i];
    const mean = count > 0 ? sum / count : NaN;
    let squares = 0;
    for (let i = 0; i < count; i++) squares += (sorted[i] - mean) * (sorted[i] - mean);
    entry.stats = {
      sorted,
      count,
      min: count > 0 ? sorted[0] : NaN,
      max: count > 0 ? sorted[count - 1] : NaN,
      mean,
      deviation: count > 0 ? Math.sqrt(squares / count) : 0,
    };
    return entry.stats;
  }

  private counts(entry: Entry): Map<string, number> {
    if (entry.counts) return entry.counts;
    const date1904 = this.workbook.date1904 === true;
    const counts = new Map<string, number>();
    this.eachCell(entry, (cell) => {
      const key = countKey(valueOfCell(cell, date1904));
      if (key !== undefined) counts.set(key, (counts.get(key) || 0) + 1);
    });
    entry.counts = counts;
    return counts;
  }

  /** The value a top/bottom rule or an average rule compares against; NaN when there is nothing to compare. */
  private limit(entry: Entry): number {
    if (entry.limit !== undefined) return entry.limit;
    const rule = entry.rule;
    const stats = this.stats(entry);
    let limit = NaN;
    if (stats.count > 0) {
      if (entry.kind === K_TOP10) {
        const rank = rule.rank !== undefined && rule.rank >= 0 ? rule.rank : 10;
        let n = rule.percent ? Math.floor((stats.count * rank) / 100) : rank;
        n = Math.max(1, Math.min(n, stats.count));
        limit = rule.bottom ? stats.sorted[n - 1] : stats.sorted[stats.count - n];
      } else {
        const deviations = rule.stdDev !== undefined && rule.stdDev > 0 ? rule.stdDev : 0;
        limit = stats.mean + (rule.aboveAverage === false ? -1 : 1) * deviations * stats.deviation;
      }
    }
    entry.limit = limit;
    return limit;
  }

  /** One threshold as a number; NaN when it cannot be worked out. */
  private threshold(entry: Entry, value: ConditionalValue | undefined, stats: Stats): number | Pending {
    if (!value) return NaN;
    const type = value.type;
    if (type === 'min' || type === 'autoMin') return stats.min;
    if (type === 'max' || type === 'autoMax') return stats.max;
    // Thresholds are the same for every cell of the rule: their formulas are not moved.
    const stated = this.formula(entry, value.value, entry.anchorRow, entry.anchorCol);
    if (stated === PENDING) return PENDING;
    const number = typeof stated === 'number' ? stated : NaN;
    if (type === 'percent') return stats.min + ((stats.max - stats.min) * number) / 100;
    if (type === 'percentile') return number === number ? percentile(stats.sorted, number) : NaN;
    // "num" and "formula".
    return number;
  }

  /** The thresholds of a rule as numbers; null when one of them cannot be worked out. */
  private thresholds(entry: Entry, values: ConditionalValue[] | undefined, count: number): number[] | null | Pending {
    if (entry.thresholds !== undefined) return entry.thresholds;
    let found: number[] | null = null;
    const stats = this.stats(entry);
    if (Array.isArray(values) && values.length >= count && stats.count > 0) {
      found = [];
      for (let i = 0; i < count; i++) {
        const limit = this.threshold(entry, values[i], stats);
        if (limit === PENDING) return PENDING;
        if (limit !== limit) {
          found = null;
          break;
        }
        found.push(limit as number);
      }
    }
    entry.thresholds = found;
    return found;
  }

  private scaleColor(entry: Entry, value: CellValue): string | Pending | undefined {
    if (typeof value !== 'number') return undefined;
    const scale = entry.rule.colorScale;
    if (!scale) return undefined;
    if (entry.scale === undefined) {
      const colors: number[][] = [];
      const stated = Array.isArray(scale.colors) ? scale.colors : [];
      for (let i = 0; i < stated.length && i < 3; i++) {
        const color = resolveColor(stated[i], this.workbook.styles);
        if (color === undefined) break;
        colors.push(channels(color));
      }
      entry.scale = colors.length >= 2 && colors.length === Math.min(stated.length, 3) ? colors : null;
    }
    const colors = entry.scale;
    if (!colors) return undefined;
    const limits = this.thresholds(entry, scale.values, colors.length);
    if (limits === PENDING) return PENDING;
    if (!limits) return undefined;
    const points = limits as number[];
    const last = points.length - 1;
    if (value <= points[0]) return mixColors(colors[0], colors[0], 0);
    if (value >= points[last]) return mixColors(colors[last], colors[last], 0);
    for (let i = 0; i < last; i++) {
      if (value > points[i + 1]) continue;
      const span = points[i + 1] - points[i];
      return mixColors(colors[i], colors[i + 1], span > 0 ? (value - points[i]) / span : 0);
    }
    return undefined;
  }

  private dataBar(entry: Entry, value: CellValue): CellDataBar | Pending | undefined {
    if (typeof value !== 'number') return undefined;
    const rule = entry.rule.dataBar;
    if (!rule) return undefined;
    const limits = this.thresholds(entry, [rule.min, rule.max], 2);
    if (limits === PENDING) return PENDING;
    if (!limits) return undefined;
    let low = (limits as number[])[0];
    let high = (limits as number[])[1];
    if (low > high) {
      const swap = low;
      low = high;
      high = swap;
    }
    const styles = this.workbook.styles;
    const extended = rule.extended === true;
    // "Automatic" limits always take in zero, so that bars compare by size.
    if (extended && rule.min && rule.min.type === 'autoMin') low = Math.min(low, 0);
    if (extended && rule.max && rule.max.type === 'autoMax') high = Math.max(high, 0);

    const positive = resolveColor(rule.color, styles) || DEFAULT_BAR_COLOR;
    const border = extended ? resolveColor(rule.borderColor, styles) : undefined;
    const negative = extended && value < 0;
    const bar: CellDataBar = {
      from: 0,
      to: 0,
      color: negative && !rule.negativeFillSameAsPositive ? resolveColor(rule.negativeFillColor, styles) || DEFAULT_NEGATIVE_COLOR : positive,
      gradient: rule.gradient !== false,
      negative,
      rightToLeft: rule.direction === 'rightToLeft',
    };
    const barBorder = negative && !rule.negativeBorderSameAsPositive && border ? resolveColor(rule.negativeBorderColor, styles) || border : border;
    if (barBorder) bar.border = barBorder;

    const clamp = (x: number) => (x < 0 ? 0 : x > 1 ? 1 : x);
    const span = high - low;
    const axisPosition = extended ? rule.axisPosition : 'none';
    if (axisPosition === 'middle') {
      // UNSURE: the axis is in the middle of the cell whatever the signs of the values; with
      // none below zero the bars use the right half only. Believed to be what Excel draws.
      const reach = Math.max(Math.abs(low), Math.abs(high));
      const length = reach > 0 ? clamp(Math.abs(value) / reach) / 2 : 0;
      bar.axis = 0.5;
      bar.from = value < 0 ? 0.5 - length : 0.5;
      bar.to = value < 0 ? 0.5 : 0.5 + length;
    } else if (axisPosition === 'automatic' && low < 0) {
      // The axis sits where zero falls between the two limits; with only negative values, at the far end.
      const axis = high > 0 ? -low / span : 1;
      bar.axis = axis;
      if (value >= 0) {
        bar.from = axis;
        bar.to = axis + (high > 0 ? clamp(value / high) : 0) * (1 - axis);
      } else {
        bar.from = axis - clamp(value / low) * axis;
        bar.to = axis;
      }
    } else {
      // The defaults of both the 2007 element and its 2010 extension.
      const minLength = (rule.minLength >= 0 && rule.minLength <= 100 ? rule.minLength : 10) / 100;
      const maxLength = (rule.maxLength >= 0 && rule.maxLength <= 100 ? rule.maxLength : 90) / 100;
      const share = span > 0 ? clamp((value - low) / span) : value >= high ? 1 : 0;
      bar.to = clamp(minLength + share * (maxLength - minLength));
      bar.negative = false;
    }
    if (bar.axis !== undefined) bar.axisColor = resolveColor(rule.axisColor, styles) || DEFAULT_AXIS_COLOR;
    return bar;
  }

  /** The icon of a cell; an index of -1 for a cell that is part of the set but shows no icon. */
  private icon(entry: Entry, value: CellValue): CellIcon | Pending | undefined {
    if (typeof value !== 'number') return undefined;
    const rule = entry.rule.iconSet;
    if (!rule) return undefined;
    const custom = Array.isArray(rule.custom) && rule.custom.length > 0 ? rule.custom : undefined;
    let name = typeof rule.set === 'string' ? rule.set : '';
    let set = ICON_SETS.get(name);
    if (!set) {
      // An unknown set of a known size is drawn as the plainest set of that size.
      const size = Number(name.charAt(0));
      name = size === 4 ? '4TrafficLights' : size === 5 ? '5Quarters' : '3TrafficLights1';
      set = ICON_SETS.get(name) as readonly IconGlyph[];
    }
    const size = custom ? custom.length : set.length;
    const limits = this.thresholds(entry, rule.values, size);
    if (limits === PENDING) return PENDING;
    if (!limits) return undefined;
    const points = limits as number[];
    let index = 0;
    // The lowest icon has no threshold of its own: its <cfvo> is only a placeholder.
    for (let i = 1; i < size; i++) {
      const inclusive = !rule.values[i] || rule.values[i].gte !== false;
      if (inclusive ? value >= points[i] : value > points[i]) index = i;
    }
    if (rule.reverse) index = size - 1 - index;
    if (custom) {
      const chosen = custom[index];
      const chosenSet = chosen ? ICON_SETS.get(chosen.set) : undefined;
      if (!chosen || !chosenSet || !(chosen.id >= 0 && chosen.id < chosenSet.length)) return { set: '', index: -1 };
      return { set: chosen.set, index: chosen.id };
    }
    return { set: name, index };
  }
}
