// Shared machinery of the function library: reading arguments the way Excel
// does (scalars, arrays and references each follow their own rules), lifting a
// scalar function over arrays, criteria of the SUMIF family, wildcards, and
// the calendar arithmetic of serial numbers.
//
// Nothing here throws, whatever the arguments are.

import { serialToDate } from '../../format/number-format';
import { MAX_COLS, MAX_ROWS } from '../../model/sheet/types';
import {
  calcError,
  chargeText,
  checkNumber,
  ERROR_NA,
  ERROR_NUM,
  ERROR_REF,
  ERROR_VALUE,
  isArray,
  isError,
  isRef,
  numberSeparators,
  parseDateText,
  parseErrorCode,
  parseNumberText,
  plainNumberText,
  readRange,
  tailFrom,
  textToNumber,
  toNumber,
  type ArgSelector,
  type ArrayValue,
  type CalcContext,
  type CalcError,
  type CalcFunction,
  type FunctionSpec,
  type RefValue,
  type Scalar,
  type Value,
} from '../values';

/** Longest text a cell holds; a function whose result would be longer answers #VALUE!. */
export const MAX_TEXT_LENGTH = 32767;
/** Largest array a function builds or reads a reference into. */
export const MAX_ARRAY_CELLS = 1 << 22;
/** Last serial of each date system: 31 Dec 9999. */
export const MAX_SERIAL_1900 = 2958465;
export const MAX_SERIAL_1904 = 2957003;

export const ERROR_CALC = calcError('#CALC!');

/** A function made by mapScalars() knows which of its arguments are single values. */
type LiftedFunction = CalcFunction & { scalarArgs?: ArgSelector };

/**
 * Describes a function. `arrayArgs` names the arguments that are calculated as arrays even
 * in an ordinary formula; which arguments are single values is taken from the function
 * itself when mapScalars() or mapScalarsExcept() made it.
 */
export function spec(call: CalcFunction, min: number, max?: number, arrayArgs?: ArgSelector): FunctionSpec {
  const out: FunctionSpec = max === undefined ? { call, min } : { call, min, max };
  if (arrayArgs !== undefined) out.arrayArgs = arrayArgs;
  const scalarArgs = (call as LiftedFunction).scalarArgs;
  if (scalarArgs !== undefined) out.scalarArgs = scalarArgs;
  return out;
}

/** A value read from an array or a context; a hole in a malformed array is an empty cell. */
export function clean(value: Scalar | undefined): Scalar {
  return value === undefined ? null : value;
}

/** Removes floating point noise beyond the 15 digits Excel works with. */
export function tidy(value: number): number {
  return Number(value.toPrecision(15));
}

// ---------------------------------------------------------------------------
// Arguments

function isSingleCell(ref: RefValue): boolean {
  return ref.range.top === ref.range.bottom && ref.range.left === ref.range.right;
}

/** The scalar an argument stands for (a scalar, a one-cell reference, a 1 x 1 array), or undefined. */
export function single(value: Value, context: CalcContext): Scalar | undefined {
  if (isRef(value)) {
    return isSingleCell(value) ? clean(context.cellValue(value.sheet, value.range.top, value.range.left)) : undefined;
  }
  if (isArray(value)) return value.rows === 1 && value.cols === 1 ? clean(value.data[0]) : undefined;
  return clean(value);
}

/** The scalar of an argument where only one value makes sense: the top left one of an array or range. */
export function first(value: Value | undefined, context: CalcContext): Scalar {
  if (value === undefined) return null;
  if (isRef(value)) return clean(context.cellValue(value.sheet, value.range.top, value.range.left));
  if (isArray(value)) return value.data.length > 0 ? clean(value.data[0]) : ERROR_CALC;
  return clean(value);
}

/**
 * An argument as an array. A reference of written bounds keeps its own shape, empty cells
 * included; a whole row or column (or a range too large to hold) is cut to the content of its
 * sheet (see readRange()). #REF! for a range that is not one, #NUM! when the cells would not fit.
 */
export function toArray(value: Value, context: CalcContext): ArrayValue | CalcError {
  if (isRef(value)) {
    const range = value.range;
    if (!(range.bottom >= range.top) || !(range.right >= range.left)) return ERROR_REF;
    const read = readRange(range, value.sheet, context, MAX_ARRAY_CELLS);
    const bottom = read.bottom;
    const right = read.right;
    const rows = bottom - range.top + 1;
    const cols = right - range.left + 1;
    const size = rows * cols;
    if (!(size >= 1 && size <= MAX_ARRAY_CELLS) || Math.floor(rows) !== rows || Math.floor(cols) !== cols) {
      return ERROR_NUM;
    }
    if (context.tick !== undefined) context.tick(size >> 4);
    const data: Scalar[] = new Array(size).fill(null);
    let filled = 0;
    context.forEachCell(value.sheet, { top: range.top, left: range.left, bottom, right }, (cell, row, col) => {
      const r = row - range.top;
      const c = col - range.left;
      if (r >= 0 && r < rows && c >= 0 && c < cols && cell !== undefined) {
        const at = r * cols + c;
        data[at] = cell;
        if (cell !== null && at >= filled) filled = at + 1;
      }
    });
    const tail = tailFrom(size, filled);
    return tail === undefined ? { rows, cols, data } : { rows, cols, data, tail };
  }
  if (isArray(value)) {
    if (!(value.rows >= 1 && value.cols >= 1) || value.data.length < value.rows * value.cols) return ERROR_CALC;
    return value;
  }
  return { rows: 1, cols: 1, data: [clean(value)] };
}

/**
 * A rectangle of cells that may be far larger than what it holds: a reference
 * is read through the context, so a whole column costs only its filled cells.
 */
export interface Grid {
  readonly rows: number;
  readonly cols: number;
  /** The value at a position; null outside the rectangle. */
  get(row: number, col: number): Scalar;
  /** Calls `visit` for every cell that holds something, with its position inside the rectangle. */
  each(visit: (value: Scalar, row: number, col: number) => void): void;
  /**
   * Positions (row * cols + col) of the cells equal to a value, from an index shared by the
   * formulas of a recalculation; undefined when there is no index to ask (see rangeIndex()).
   */
  find?(key: ExactKey): number[] | undefined;
}

export function toGrid(value: Value, context: CalcContext): Grid {
  if (isRef(value)) {
    const range = value.range;
    const sheet = value.sheet;
    const rows = range.bottom >= range.top ? range.bottom - range.top + 1 : 0;
    const cols = range.right >= range.left ? range.right - range.left + 1 : 0;
    return {
      rows,
      cols,
      get(row, col) {
        if (!(row >= 0 && row < rows && col >= 0 && col < cols)) return null;
        return clean(context.cellValue(sheet, range.top + row, range.left + col));
      },
      each(visit) {
        if (rows === 0 || cols === 0) return;
        context.forEachCell(sheet, range, (cell, row, col) => {
          if (cell !== null && cell !== undefined) visit(cell, row - range.top, col - range.left);
        });
      },
      find(key) {
        if (rows === 0 || cols === 0) return undefined;
        const index = rangeIndex(value, context);
        return index === undefined ? undefined : criterionPositions(index, key);
      },
    };
  }
  if (isArray(value)) {
    const data = value.data;
    const rows = value.rows >= 1 && value.cols >= 1 && data.length >= value.rows * value.cols ? value.rows : 0;
    const cols = rows === 0 ? 0 : value.cols;
    return {
      rows,
      cols,
      get(row, col) {
        if (!(row >= 0 && row < rows && col >= 0 && col < cols)) return null;
        return clean(data[row * cols + col]);
      },
      each(visit) {
        const size = rows * cols;
        for (let i = 0; i < size; i++) {
          const cell = data[i];
          if (cell !== null && cell !== undefined) visit(cell, Math.floor(i / cols), i % cols);
        }
      },
    };
  }
  const scalar = clean(value);
  return {
    rows: 1,
    cols: 1,
    get(row, col) {
      return row === 0 && col === 0 ? scalar : null;
    },
    each(visit) {
      if (scalar !== null) visit(scalar, 0, 0);
    },
  };
}

/**
 * The rows of a grid worth looking at in a search that never matches an empty cell: a
 * reference stops at the last row of its sheet that holds something (CalcContext.usedExtent).
 * Below it every cell is empty, so the answer is the one the whole range gives; what depends on
 * how many empty cells follow (a binary search) cuts them off itself.
 */
/** Ranges smaller than this are searched whole: working out how far their sheet goes would cost more. */
const SMALL_SEARCH = 1 << 16;

export function usedRows(value: Value, grid: Grid, context: CalcContext): number {
  if (!isRef(value) || grid.rows * grid.cols <= SMALL_SEARCH) return grid.rows;
  const rows = Math.min(grid.rows, context.usedExtent(value.sheet, value.range).rows - value.range.top);
  return rows > 0 ? Math.floor(rows) : 0;
}

export function usedCols(value: Value, grid: Grid, context: CalcContext): number {
  if (!isRef(value) || grid.rows * grid.cols <= SMALL_SEARCH) return grid.cols;
  const cols = Math.min(grid.cols, context.usedExtent(value.sheet, value.range).cols - value.range.left);
  return cols > 0 ? Math.floor(cols) : 0;
}

// ---------------------------------------------------------------------------
// Lifting scalar functions over arrays

/** A function of scalars; arguments named "raw" arrive as they were given and must be read by the function. */
export type ScalarFunction = (args: Scalar[], context: CalcContext) => Scalar;
export type MixedFunction = (args: Value[], context: CalcContext) => Scalar;

function lift(fn: MixedFunction, raw: ((index: number) => boolean) | undefined): CalcFunction {
  const lifted: LiftedFunction = (args, context) => {
    const count = args.length;
    const parts: Value[] = new Array(count);
    let rows = 1;
    let cols = 1;
    let lifted = false;
    for (let i = 0; i < count; i++) {
      const arg = args[i];
      if (raw !== undefined && raw(i)) {
        parts[i] = arg;
        continue;
      }
      const scalar = single(arg, context);
      if (scalar !== undefined) {
        parts[i] = scalar;
        continue;
      }
      const array = toArray(arg, context);
      if (isError(array)) return array;
      parts[i] = array;
      lifted = true;
      if (array.rows > rows) rows = array.rows;
      if (array.cols > cols) cols = array.cols;
    }
    if (!lifted) return fn(parts, context);
    const size = rows * cols;
    if (!(size <= MAX_ARRAY_CELLS)) return ERROR_NUM;
    const data: Scalar[] = new Array(size);
    const call: Value[] = parts.slice();
    for (let r = 0; r < rows; r++) {
      // One unit of work per element; this is also where a calculation over its budget is given up.
      if (context.tick !== undefined) context.tick(cols);
      for (let c = 0; c < cols; c++) {
        for (let i = 0; i < count; i++) {
          const part = parts[i];
          if (isArray(part) && !(raw !== undefined && raw(i))) {
            // One row or one column repeats along the other direction; anything else outside is #N/A.
            const pr = part.rows === 1 ? 0 : r;
            const pc = part.cols === 1 ? 0 : c;
            call[i] = pr < part.rows && pc < part.cols ? clean(part.data[pr * part.cols + pc]) : ERROR_NA;
          }
        }
        const value = clean(fn(call, context));
        // An array of long texts can be far larger than its number of cells says.
        if (typeof value === 'string' && !chargeText(value)) return ERROR_VALUE;
        data[r * cols + c] = value;
      }
    }
    return { rows, cols, data };
  };
  lifted.scalarArgs = raw === undefined ? true : (index) => !raw(index);
  return lifted;
}

/**
 * Makes a function of scalars accept arrays and ranges: it is applied item by
 * item and the results come back as an array (`ABS({1,-2})` is `{1,2}`). A
 * one-cell reference counts as a scalar.
 */
export function mapScalars(fn: ScalarFunction): CalcFunction {
  return lift(fn as MixedFunction, undefined);
}

/** As mapScalars, but the arguments `raw` names (ranges, lists) are passed through untouched. */
export function mapScalarsExcept(raw: (index: number) => boolean, fn: MixedFunction): CalcFunction {
  return lift(fn, raw);
}

/**
 * A function of numbers: every argument is converted as arithmetic does
 * (errors pass through, text that is not a number is #VALUE!) and the result
 * is checked. Missing optional arguments are simply absent.
 */
export function numeric(fn: (...values: number[]) => number | CalcError): CalcFunction {
  return mapScalars((args, context) => {
    const values: number[] = new Array(args.length);
    for (let i = 0; i < args.length; i++) {
      const value = scalarNumber(args[i], context);
      if (typeof value !== 'number') return value;
      values[i] = value;
    }
    const result = fn.apply(undefined, values);
    // -0 is not a value a cell holds.
    return typeof result === 'number' ? (result === 0 ? 0 : checkNumber(result)) : result;
  });
}

// ---------------------------------------------------------------------------
// Text as numbers

export { parseDateText, parseNumberText };

/** A scalar as a number; text may also be a date, a time or a formatted number. The rule operators follow too. */
export function scalarNumber(value: Scalar, context: CalcContext): number | CalcError {
  return toNumber(clean(value), context);
}

// ---------------------------------------------------------------------------
// Numbers of aggregate functions

export interface NumberRules {
  /** The "A" functions: text inside ranges counts as 0, booleans as 1 and 0. */
  loose?: boolean;
  /** Errors are left out instead of becoming the result (AGGREGATE). */
  skipErrors?: boolean;
  /** Cells of references to leave out (SUBTOTAL passes over other subtotals). */
  skipCell?: (sheet: number, row: number, col: number) => boolean;
}

/**
 * Calls `visit` for every number the arguments of an aggregate function give.
 * Inside references and arrays only numbers count; text, booleans and empty
 * cells are skipped. An argument given directly is converted: a boolean is 1
 * or 0, numeric text is its number, other text is #VALUE!. The first error
 * found is returned.
 */
export function eachNumber(
  args: Value[],
  context: CalcContext,
  rules: NumberRules | undefined,
  visit: (value: number) => void,
): CalcError | undefined {
  const loose = rules !== undefined && rules.loose === true;
  const skipErrors = rules !== undefined && rules.skipErrors === true;
  const skipCell = rules !== undefined ? rules.skipCell : undefined;
  let error = undefined as CalcError | undefined;
  const cell = (value: Scalar | undefined): void => {
    if (typeof value === 'number') visit(value);
    else if (value === null || value === undefined) return;
    else if (isError(value)) {
      if (!skipErrors && error === undefined) error = value;
    } else if (loose) visit(value === true ? 1 : 0);
  };
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (isRef(arg)) {
      if (arg.range.bottom >= arg.range.top && arg.range.right >= arg.range.left) {
        if (skipCell === undefined) context.forEachCell(arg.sheet, arg.range, cell);
        else {
          const sheet = arg.sheet;
          context.forEachCell(sheet, arg.range, (value, row, col) => {
            if (!skipCell(sheet, row, col)) cell(value);
          });
        }
      }
    } else if (isArray(arg)) {
      const data = arg.data;
      for (let k = 0; k < data.length; k++) cell(data[k]);
    } else if (isError(arg)) {
      if (!skipErrors) return arg;
    } else {
      const number = scalarNumber(clean(arg), context);
      if (typeof number === 'number') visit(number);
      else if (!skipErrors) return number;
    }
    if (error !== undefined) return error;
  }
  return undefined;
}

export function collectNumbers(args: Value[], context: CalcContext, rules?: NumberRules): number[] | CalcError {
  const numbers: number[] = [];
  const error = eachNumber(args, context, rules, (value) => {
    numbers.push(value);
  });
  return error === undefined ? numbers : error;
}

/** Every value of the arguments in order, empty cells of arrays included; references give only filled cells. */
export function eachScalar(args: Value[], context: CalcContext, visit: (value: Scalar) => void): void {
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (isRef(arg)) {
      if (arg.range.bottom >= arg.range.top && arg.range.right >= arg.range.left) {
        context.forEachCell(arg.sheet, arg.range, (value) => {
          if (value !== null && value !== undefined) visit(value);
        });
      }
    } else if (isArray(arg)) {
      const data = arg.data;
      for (let k = 0; k < data.length; k++) visit(clean(data[k]));
    } else visit(clean(arg));
  }
}

// ---------------------------------------------------------------------------
// Index of a range by value
//
// An exact lookup or an "equals" criterion over a range is a scan of the range. A sheet with
// thousands of such formulas over the same table would scan it thousands of times, so the
// positions of every value are worked out once and kept by the context for the rest of the
// recalculation (CalcContext.memo). Within one recalculation every cell has one value, so
// what is kept never goes stale.

const NO_POSITIONS: number[] = [];

/** A value cells are looked up by: a number, text (folded with foldCase) or a boolean. */
export type ExactKey = number | string | boolean;

/** Text longer than this is not put into the index; a search for such text scans. */
const MAX_INDEXED_TEXT = 255;
/** The index is built when a range is searched this many times in a recalculation; before that a scan is cheaper. */
const INDEX_AFTER_USES = 2;

/** One position, or several in a list. Most values of a key column are there once; a list for each would cost far more memory. */
type Positions = number | number[];

/** Positions are row * cols + col inside the range, each list in ascending order. */
export interface RangeIndex {
  numbers: Map<number, Positions>;
  /** By folded text. */
  texts: Map<string, Positions>;
  /**
   * Text cells that read as a plain number, by that number: a criterion "equals 5" also takes
   * "5". Text that is none the way files write one is read with the separators of the
   * locale, as the criterion itself is.
   */
  numericTexts: Map<number, Positions>;
  yes: number[];
  no: number[];
  /** Whether any text was too long to index. */
  longText: boolean;
}

function addPosition<K>(map: Map<K, Positions>, key: K, position: number): void {
  const held = map.get(key);
  if (held === undefined) map.set(key, position);
  else if (typeof held === 'number') map.set(key, [held, position]);
  else held.push(position);
}

function listOf(positions: Positions | undefined): number[] {
  return positions === undefined ? NO_POSITIONS : typeof positions === 'number' ? [positions] : positions;
}

function buildIndex(ref: RefValue, context: CalcContext): RangeIndex {
  const index: RangeIndex = { numbers: new Map(), texts: new Map(), numericTexts: new Map(), yes: [], no: [], longText: false };
  const range = ref.range;
  const cols = range.right - range.left + 1;
  const separators = numberSeparators(context);
  let previous = -1;
  let ordered = true;
  context.forEachCell(ref.sheet, range, (value, row, col) => {
    const position = (row - range.top) * cols + (col - range.left);
    if (position < previous) ordered = false;
    previous = position;
    if (typeof value === 'number') addPosition(index.numbers, value, position);
    else if (typeof value === 'string') {
      if (value.length > MAX_INDEXED_TEXT) index.longText = true;
      else {
        addPosition(index.texts, foldCase(value), position);
        const number = plainNumberText(value, separators);
        if (number !== undefined) addPosition(index.numericTexts, number, position);
      }
    } else if (value === true) index.yes.push(position);
    else if (value === false) index.no.push(position);
  });
  if (!ordered) {
    // A context that does not visit in row order: put every list in order once.
    const sort = (list: Positions) => {
      if (typeof list !== 'number') list.sort((a, b) => a - b);
    };
    index.numbers.forEach(sort);
    index.texts.forEach(sort);
    index.numericTexts.forEach(sort);
    sort(index.yes);
    sort(index.no);
  }
  return index;
}

/**
 * The index of a reference, or undefined when a scan is the better way: the context keeps
 * nothing between formulas, or this is the first search of the range in this recalculation.
 */
export function rangeIndex(ref: RefValue, context: CalcContext): RangeIndex | undefined {
  if (context.memo === undefined) return undefined;
  const range = ref.range;
  if (!(range.bottom >= range.top) || !(range.right >= range.left)) return undefined;
  const key = ref.sheet + ':' + range.top + ':' + range.left + ':' + range.bottom + ':' + range.right;
  const uses = context.memo('uses:' + key, () => ({ count: 0 }));
  if (++uses.count < INDEX_AFTER_USES) return undefined;
  // Which text cells read as numbers depends on the separators of the locale.
  const separators = numberSeparators(context);
  const written = separators === undefined ? '' : ':' + separators.decimal + separators.group;
  return context.memo('index:' + key + written, () => buildIndex(ref, context));
}

/** Positions of the cells a lookup for `key` matches: the same kind of value, text without regard to case. Undefined when the index cannot tell. */
export function lookupPositions(index: RangeIndex, key: ExactKey): number[] | undefined {
  if (typeof key === 'number') return listOf(index.numbers.get(key));
  if (typeof key === 'boolean') return key ? index.yes : index.no;
  if (key.length > MAX_INDEXED_TEXT) return index.longText ? undefined : NO_POSITIONS;
  return listOf(index.texts.get(key));
}

/** Positions of the cells an "equals" criterion matches: as lookupPositions(), and a number also matches text that reads as it. */
function criterionPositions(index: RangeIndex, key: ExactKey): number[] | undefined {
  if (typeof key !== 'number') return lookupPositions(index, key);
  const texts = index.numericTexts.get(key);
  if (texts === undefined) return listOf(index.numbers.get(key));
  const numbers = listOf(index.numbers.get(key));
  if (numbers.length === 0) return listOf(texts);
  return mergePositions(numbers, listOf(texts));
}

function mergePositions(numbers: number[], texts: number[]): number[] {
  // Both lists are in order; so is what they merge into, as a scan would meet the cells.
  const merged: number[] = new Array(numbers.length + texts.length);
  let i = 0;
  let j = 0;
  for (let k = 0; k < merged.length; k++) merged[k] = j >= texts.length || (i < numbers.length && numbers[i] < texts[j]) ? numbers[i++] : texts[j++];
  return merged;
}

// ---------------------------------------------------------------------------
// Wildcards
//
// "*" stands for any run of characters, "?" for one, "~" takes the next
// character literally. Patterns are matched by hand, segment by segment, so
// the cost is at most (text length x pattern length): there is no regular
// expression to backtrack.

const ANY_ONE = '￿';

export interface Pattern {
  /** The parts between "*"; inside a part ANY_ONE marks a "?". */
  readonly segments: string[];
  /** Marks, for each segment, which of its characters are "?" (as opposed to a literal U+FFFF). */
  readonly wild: boolean[][];
}

/** Text folded for comparison without regard to case, keeping every character at its position. */
export function foldCase(text: string): string {
  const upper = text.toUpperCase();
  if (upper.length === text.length) return upper;
  let out = '';
  for (let i = 0; i < text.length; i++) {
    const c = text.charAt(i);
    const u = c.toUpperCase();
    out += u.length === 1 ? u : c;
  }
  return out;
}

export function hasWildcard(text: string): boolean {
  return text.indexOf('*') >= 0 || text.indexOf('?') >= 0 || text.indexOf('~') >= 0;
}

export function compilePattern(pattern: string): Pattern {
  const segments: string[] = [];
  const wild: boolean[][] = [];
  let segment = '';
  let marks: boolean[] = [];
  for (let i = 0; i < pattern.length; i++) {
    const c = pattern.charAt(i);
    if (c === '*') {
      segments.push(segment);
      wild.push(marks);
      segment = '';
      marks = [];
    } else if (c === '?') {
      segment += ANY_ONE;
      marks.push(true);
    } else if (c === '~' && i + 1 < pattern.length && '*?~'.indexOf(pattern.charAt(i + 1)) >= 0) {
      segment += pattern.charAt(++i);
      marks.push(false);
    } else {
      segment += c;
      marks.push(false);
    }
  }
  segments.push(segment);
  wild.push(marks);
  return { segments, wild };
}

function segmentAt(pattern: Pattern, index: number, text: string, at: number): boolean {
  const segment = pattern.segments[index];
  if (at < 0 || at + segment.length > text.length) return false;
  const marks = pattern.wild[index];
  for (let i = 0; i < segment.length; i++) {
    if (!marks[i] && segment.charCodeAt(i) !== text.charCodeAt(at + i)) return false;
  }
  return true;
}

function findSegment(pattern: Pattern, index: number, text: string, from: number, limit: number): number {
  const length = pattern.segments[index].length;
  for (let at = from; at + length <= limit; at++) {
    if (segmentAt(pattern, index, text, at)) return at;
  }
  return -1;
}

/** True when the whole text matches the pattern. Both must already be folded if case is to be ignored. */
export function patternMatches(pattern: Pattern, text: string): boolean {
  const segments = pattern.segments;
  const last = segments.length - 1;
  if (last === 0) return segments[0].length === text.length && segmentAt(pattern, 0, text, 0);
  if (!segmentAt(pattern, 0, text, 0)) return false;
  const tail = text.length - segments[last].length;
  let at = segments[0].length;
  if (tail < at || !segmentAt(pattern, last, text, tail)) return false;
  for (let i = 1; i < last; i++) {
    const found = findSegment(pattern, i, text, at, tail);
    if (found < 0) return false;
    at = found + segments[i].length;
  }
  return true;
}

/** Position of the first place at or after `from` where the pattern matches a stretch of the text, or -1. */
export function patternSearch(pattern: Pattern, text: string, from: number): number {
  const segments = pattern.segments;
  const start = findSegment(pattern, 0, text, from, text.length);
  if (start < 0) return -1;
  // The later parts are taken at their first position; if they do not fit after the first
  // possible start they do not fit after any later one either.
  let at = start + segments[0].length;
  for (let i = 1; i < segments.length; i++) {
    const found = findSegment(pattern, i, text, at, text.length);
    if (found < 0) return -1;
    at = found + segments[i].length;
  }
  return start;
}

// ---------------------------------------------------------------------------
// Criteria of SUMIF, COUNTIF and the ...IFS family

export interface Criterion {
  test(value: Scalar): boolean;
  /** Whether an empty cell meets the criterion; such criteria cannot be answered from the filled cells alone. */
  readonly blank: boolean;
  /** The value the criterion asks cells to equal, when that is all it asks: an index can answer it. */
  readonly exact?: ExactKey;
}

const OPERATORS = ['>=', '<=', '<>', '>', '<', '='];

function relation(op: string, order: number): boolean {
  switch (op) {
    case '>':
      return order > 0;
    case '<':
      return order < 0;
    case '>=':
      return order >= 0;
    default:
      return order <= 0;
  }
}

function criterion(test: (value: Scalar) => boolean, exact?: ExactKey): Criterion {
  return exact === undefined ? { test, blank: test(null) } : { test, blank: test(null), exact };
}

/**
 * Compiles a criterion: `">5"`, `"<=10"`, `"<>x"`, `"=text"`, a plain value to
 * equal, wildcards in text. Text is compared without regard to case; a number
 * equals numeric text; "<", ">" compare only values of the operand's own kind.
 * A number in the text is read by parseNumberText(): ">7.5" asks for more than 7.5 in
 * every locale, and so does ">7,5" where the locale of the context writes a decimal comma.
 */
export function compileCriterion(source: Scalar, context: CalcContext): Criterion {
  let op = '=';
  let operand: Scalar = source;
  if (typeof source === 'string') {
    let rest = source;
    let explicit = false;
    for (let i = 0; i < OPERATORS.length; i++) {
      if (source.lastIndexOf(OPERATORS[i], 0) === 0) {
        op = OPERATORS[i];
        rest = source.substring(op.length);
        explicit = true;
        break;
      }
    }
    if (rest === '') {
      // "=" asks for cells that are truly empty, "" also takes empty text, "<>" every cell that holds something.
      if (!explicit) return criterion((value) => value === null || value === '');
      if (op === '=') return criterion((value) => value === null);
      if (op === '<>') return criterion((value) => value !== null);
      return criterion(() => false);
    }
    const number = parseNumberText(rest, context);
    const upper = rest.toUpperCase();
    const code = parseErrorCode(rest);
    if (number !== undefined) operand = number;
    else if (upper === 'TRUE') operand = true;
    else if (upper === 'FALSE') operand = false;
    else if (code !== undefined) operand = calcError(code);
    else operand = rest;
  } else if (source === null || source === undefined) {
    // An empty criterion cell is read as 0.
    operand = 0;
  }

  let equals: (value: Scalar) => boolean;
  let ordered: ((value: Scalar) => boolean) | undefined;
  let exact: ExactKey | undefined;
  if (typeof operand === 'number') {
    const target = operand;
    const read = textNumbers(context);
    exact = target;
    equals = (value) => (typeof value === 'number' ? value === target : typeof value === 'string' && read(value) === target);
    ordered = (value) => typeof value === 'number' && relation(op, value < target ? -1 : value > target ? 1 : 0);
  } else if (typeof operand === 'boolean') {
    const target = operand;
    exact = target;
    equals = (value) => value === target;
    ordered = (value) => typeof value === 'boolean' && relation(op, (value ? 1 : 0) - (target ? 1 : 0));
  } else if (isError(operand)) {
    const target = operand.error;
    equals = (value) => isError(value) && value.error === target;
  } else {
    const target = foldCase(operand as string);
    if (hasWildcard(target)) {
      const pattern = compilePattern(target);
      equals = (value) => typeof value === 'string' && patternMatches(pattern, foldCase(value));
    } else {
      exact = target;
      equals = (value) => typeof value === 'string' && value.length === target.length && foldCase(value) === target;
    }
    ordered = (value) => {
      if (typeof value !== 'string') return false;
      const text = foldCase(value);
      return relation(op, text < target ? -1 : text > target ? 1 : 0);
    };
  }
  if (op === '=') return criterion(equals, exact);
  if (op === '<>') return criterion((value) => !equals(value));
  return criterion(ordered === undefined ? () => false : ordered);
}

/** Texts whose reading with the separators of a locale is kept during one recalculation. */
const MAX_KEPT_TEXTS = 100000;

/**
 * How a criterion that asks for a number reads the text cells it meets: as a plain number,
 * and where the locale writes a decimal comma also as the locale writes one. A scan meets
 * the same texts again and again (every formula of a column asks the same range), so the
 * second reading is kept for the rest of the recalculation.
 */
function textNumbers(context: CalcContext): (text: string) => number | undefined {
  const separators = numberSeparators(context);
  if (separators === undefined) return textToNumber;
  if (context.memo === undefined) return (text) => plainNumberText(text, separators);
  const kept = context.memo('texts:' + separators.decimal + separators.group, () => new Map<string, number | null>());
  return (text) => {
    const plain = textToNumber(text);
    if (plain !== undefined) return plain;
    // Longer text is no number however it is read.
    if (text.length > MAX_KEPT_TEXT_LENGTH) return undefined;
    const held = kept.get(text);
    if (held !== undefined) return held === null ? undefined : held;
    const number = plainNumberText(text, separators);
    if (kept.size < MAX_KEPT_TEXTS) kept.set(text, number === undefined ? null : number);
    return number;
  };
}

/** Longest text plainNumberText() looks at closely. */
const MAX_KEPT_TEXT_LENGTH = 64;

export interface Condition {
  grid: Grid;
  criterion: Criterion;
}

function othersHold(conditions: Condition[], skip: number, row: number, col: number): boolean {
  for (let i = 0; i < conditions.length; i++) {
    if (i !== skip && !conditions[i].criterion.test(conditions[i].grid.get(row, col))) return false;
  }
  return true;
}

function driver(conditions: Condition[]): number {
  for (let i = 0; i < conditions.length; i++) if (!conditions[i].criterion.blank) return i;
  return -1;
}

/**
 * The positions where the "equals" criterion of some condition holds, taken from the index
 * of its range; undefined when no condition can be answered that way.
 */
function indexed(conditions: Condition[]): { lead: number; positions: number[] } | undefined {
  for (let i = 0; i < conditions.length; i++) {
    const condition = conditions[i];
    const exact = condition.criterion.exact;
    if (exact === undefined || condition.criterion.blank || condition.grid.find === undefined) continue;
    const positions = condition.grid.find(exact);
    if (positions !== undefined) return { lead: i, positions };
  }
  return undefined;
}

/** Number of positions where every condition holds. All grids must have the same shape. */
export function countMatches(conditions: Condition[]): number {
  if (conditions.length === 0) return 0;
  let count = 0;
  const found = indexed(conditions);
  if (found !== undefined) {
    if (conditions.length === 1) return found.positions.length;
    const cols = conditions[found.lead].grid.cols;
    for (let i = 0; i < found.positions.length; i++) {
      const position = found.positions[i];
      const row = Math.floor(position / cols);
      if (othersHold(conditions, found.lead, row, position - row * cols)) count++;
    }
    return count;
  }
  const lead = driver(conditions);
  if (lead >= 0) {
    // Only filled cells of this range can match, so only they are visited.
    const test = conditions[lead].criterion.test;
    conditions[lead].grid.each((value, row, col) => {
      if (test(value) && othersHold(conditions, lead, row, col)) count++;
    });
    return count;
  }
  // Every criterion accepts empty cells: count the positions that fail instead.
  const cols = conditions[0].grid.cols;
  const failing = new Set<number>();
  for (let i = 0; i < conditions.length; i++) {
    const test = conditions[i].criterion.test;
    conditions[i].grid.each((value, row, col) => {
      if (!test(value)) failing.add(row * cols + col);
    });
  }
  return conditions[0].grid.rows * cols - failing.size;
}

/** Calls `visit` with every filled cell of `values` whose position meets all conditions. */
export function eachMatch(conditions: Condition[], values: Grid, visit: (value: Scalar) => void): void {
  const found = indexed(conditions);
  if (found !== undefined) {
    const cols = conditions[found.lead].grid.cols;
    for (let i = 0; i < found.positions.length; i++) {
      const position = found.positions[i];
      const row = Math.floor(position / cols);
      const col = position - row * cols;
      if (!othersHold(conditions, found.lead, row, col)) continue;
      const result = values.get(row, col);
      if (result !== null) visit(result);
    }
    return;
  }
  const lead = driver(conditions);
  if (lead >= 0) {
    const test = conditions[lead].criterion.test;
    conditions[lead].grid.each((value, row, col) => {
      if (!test(value) || !othersHold(conditions, lead, row, col)) return;
      const result = values.get(row, col);
      if (result !== null) visit(result);
    });
    return;
  }
  const rows = conditions.length > 0 ? conditions[0].grid.rows : values.rows;
  const cols = conditions.length > 0 ? conditions[0].grid.cols : values.cols;
  values.each((value, row, col) => {
    if (row < rows && col < cols && othersHold(conditions, -1, row, col)) visit(value);
  });
}

/**
 * The numbers of `values` at the positions that meet the (range, criterion)
 * pairs. With `reshape` the value range takes the shape of the criteria range
 * from its own top left cell, as SUMIF does; without it a different shape is
 * #VALUE!. Text and booleans among the values are skipped, errors returned.
 */
export function matchNumbers(
  values: Value,
  pairs: Value[],
  context: CalcContext,
  reshape: boolean,
  visit: (value: number) => void,
): CalcError | undefined {
  const conditions = compileConditions(pairs, context);
  if (isError(conditions)) return conditions;
  const rows = conditions[0].grid.rows;
  const cols = conditions[0].grid.cols;
  let target = values;
  if (reshape && isRef(values)) {
    const range = values.range;
    target = {
      ref: true,
      sheet: values.sheet,
      range: {
        top: range.top,
        left: range.left,
        bottom: Math.min(range.top + rows - 1, MAX_ROWS - 1),
        right: Math.min(range.left + cols - 1, MAX_COLS - 1),
      },
    };
  }
  const grid = toGrid(target, context);
  if (!reshape && (grid.rows !== rows || grid.cols !== cols)) return ERROR_VALUE;
  let error = undefined as CalcError | undefined;
  eachMatch(conditions, grid, (value) => {
    if (typeof value === 'number') visit(value);
    else if (isError(value) && error === undefined) error = value;
  });
  return error;
}

/** The conditions of COUNTIFS-like arguments, or an error. */
export function compileConditions(pairs: Value[], context: CalcContext): Condition[] | CalcError {
  if (pairs.length === 0 || pairs.length % 2 !== 0) return ERROR_VALUE;
  const conditions: Condition[] = [];
  for (let i = 0; i < pairs.length; i += 2) {
    const grid = toGrid(pairs[i], context);
    if (grid.rows === 0 || grid.cols === 0) return ERROR_REF;
    if (i > 0 && (grid.rows !== conditions[0].grid.rows || grid.cols !== conditions[0].grid.cols)) return ERROR_VALUE;
    conditions.push({ grid, criterion: compileCriterion(first(pairs[i + 1], context), context) });
  }
  return conditions;
}

// ---------------------------------------------------------------------------
// Rounding

/**
 * Rounds to a number of decimal digits (negative: to tens, hundreds, ...).
 * `mode` 0 rounds half away from zero, 1 always away from zero, -1 toward
 * zero. The scaled value is first cleaned of floating point noise, so
 * 1.005 rounds to 1.01 as it does in Excel.
 */
export function roundTo(value: number, digits: number, mode: number): number {
  if (value === 0 || !isFinite(value)) return value;
  const places = digits < 0 ? Math.ceil(digits) : Math.floor(digits);
  const abs = Math.abs(value);
  let scale: number;
  let scaled: number;
  if (places >= 0) {
    scale = Math.pow(10, places);
    scaled = abs * scale;
    if (!isFinite(scaled)) return value;
  } else {
    scale = Math.pow(10, -places);
    if (!isFinite(scale)) return mode === 1 ? Infinity : 0;
    scaled = abs / scale;
  }
  // Past 2^53 there are no fraction digits left to round.
  if (scaled >= 9007199254740992) return value;
  scaled = tidy(scaled);
  const rounded = mode === 0 ? Math.floor(scaled + 0.5) : mode > 0 ? Math.ceil(scaled) : Math.floor(scaled);
  const result = places >= 0 ? rounded / scale : rounded * scale;
  return value < 0 ? -result : result;
}

// ---------------------------------------------------------------------------
// Calendar

/** Days from 1 January 1970 to a date of the Gregorian calendar; `month` is 1 to 12. */
function daysFromCivil(year: number, month: number, day: number): number {
  const y = month <= 2 ? year - 1 : year;
  const era = Math.floor(y / 400);
  const yoe = y - era * 400;
  const doy = Math.floor((153 * (month + (month > 2 ? -3 : 9)) + 2) / 5) + day - 1;
  const doe = yoe * 365 + Math.floor(yoe / 4) - Math.floor(yoe / 100) + doy;
  return era * 146097 + doe - 719468;
}

/** Days between the 1900 date system's day 0 (30 December 1899) and 1 January 1970. */
const EPOCH_1900 = 25569;
const EPOCH_SHIFT_1904 = 1462;

/**
 * Serial of the first day of a month; `month` may run past 1..12 and carries
 * into the year. In the 1900 system the days before 1 March 1900 are one
 * lower, which makes room for the 29 February 1900 Excel believes in.
 */
export function monthStart(year: number, month: number, date1904: boolean): number {
  const total = year * 12 + (month - 1);
  const y = Math.floor(total / 12);
  const m = total - y * 12 + 1;
  const real = daysFromCivil(y, m, 1) + EPOCH_1900;
  if (date1904) return real - EPOCH_SHIFT_1904;
  return real < 61 ? real - 1 : real;
}

/** Serial of a date; day and month may overflow in either direction, as in DATE. */
export function dateSerial(year: number, month: number, day: number, date1904: boolean): number {
  return monthStart(year, month, date1904) + day - 1;
}

export function daysInMonth(year: number, month: number, date1904: boolean): number {
  return monthStart(year, month + 1, date1904) - monthStart(year, month, date1904);
}

/** Leap years as the date system sees them: the 1900 system counts 1900 as one. */
export function isLeapYear(year: number, date1904: boolean): boolean {
  return daysInMonth(year, 2, date1904) === 29;
}

export interface DateParts {
  year: number;
  month: number;
  day: number;
}

/** Year, month and day of a serial (its time of day is dropped), or undefined when it is not a date. */
export function dateParts(serial: number, date1904: boolean): DateParts | undefined {
  if (!(serial >= 0)) return undefined;
  const parts = serialToDate(Math.floor(serial), date1904);
  if (parts.year !== parts.year) return undefined;
  return { year: parts.year, month: parts.month, day: parts.day };
}

export function maxSerial(date1904: boolean): number {
  return date1904 ? MAX_SERIAL_1904 : MAX_SERIAL_1900;
}

/** Day of the week of a serial, 0 for Sunday, in the calendar the date system shows. */
export function weekdayOf(serial: number, date1904: boolean): number {
  const day = Math.floor(serial);
  // Serial 1 is shown as a Sunday in the 1900 system; 1 January 1904 (serial 0) was a Friday.
  const index = (day + (date1904 ? 5 : 6)) % 7;
  return index < 0 ? index + 7 : index;
}

/** A scalar as a date serial: numbers as they are, text as a typed date; #NUM! outside the calendar. */
export function scalarSerial(value: Scalar, context: CalcContext): number | CalcError {
  const number = scalarNumber(value, context);
  if (typeof number !== 'number') return number;
  if (!(number >= 0) || Math.floor(number) > maxSerial(context.date1904 === true)) return ERROR_NUM;
  return number;
}
