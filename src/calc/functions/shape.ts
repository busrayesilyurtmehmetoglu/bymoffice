// The functions of dynamic arrays that change the shape of an array: VSTACK and HSTACK put
// arrays together, TAKE, DROP, CHOOSEROWS and CHOOSECOLS cut parts out of one, EXPAND makes
// one larger, TOCOL and TOROW lay one out in a line, WRAPROWS and WRAPCOLS fold a line into
// rows or columns.
//
// What is missing where arrays of different sizes meet is #N/A, as in Excel. Every result is
// checked against MAX_ARRAY_CELLS before anything is made, so that no argument (a size of a
// million by a million) makes the engine slow or out of memory: such a result is #NUM!.

import {
  ERROR_NA,
  ERROR_NUM,
  ERROR_VALUE,
  isArray,
  isError,
  isRef,
  readRange,
  toBoolean,
  toNumber,
  type ArrayValue,
  type CalcContext,
  type CalcError,
  type CalcFunction,
  type FunctionSpec,
  type Scalar,
  type Value,
} from '../values';
import { clean, ERROR_CALC, first, MAX_ARRAY_CELLS, single, toArray } from './helpers';

/** Whether an array of this size may be made. */
const fits = (rows: number, cols: number): boolean => rows >= 1 && cols >= 1 && rows * cols <= MAX_ARRAY_CELLS;

/** Counts the work of making an array of `size` elements. */
function charge(context: CalcContext, size: number): void {
  if (context.tick !== undefined) context.tick(size >> 4);
}

/** An argument that is a whole number, cut towards zero; `fallback` when it is left out or empty. */
function wholeArg(args: Value[], index: number, fallback: number, context: CalcContext): number | CalcError {
  if (index >= args.length || args[index] === null) return fallback;
  const value = single(args[index], context);
  // A range or an array of several values where one is taken.
  if (value === undefined) return ERROR_VALUE;
  const number = toNumber(value, context);
  if (typeof number !== 'number') return number;
  if (number !== number) return ERROR_VALUE;
  return number < 0 ? Math.ceil(number) : Math.floor(number);
}

/** The value an array is filled with: #N/A when the argument is not there. */
function padArg(args: Value[], index: number, context: CalcContext): Scalar {
  return index >= args.length ? ERROR_NA : first(args[index], context);
}

/**
 * The size an argument has as an array (see toArray()), worked out without reading a cell: a whole
 * column is cut to the content of its sheet. One cell for what is no rectangle (toArray() says the error).
 */
function sizeOf(value: Value, context: CalcContext): { rows: number; cols: number } {
  if (isRef(value)) {
    const range = value.range;
    if (!(range.bottom >= range.top) || !(range.right >= range.left)) return { rows: 1, cols: 1 };
    const read = readRange(range, value.sheet, context, MAX_ARRAY_CELLS);
    return { rows: read.bottom - range.top + 1, cols: read.right - range.left + 1 };
  }
  if (isArray(value)) return { rows: value.rows, cols: value.cols };
  return { rows: 1, cols: 1 };
}

/** Every argument from `from` on as an array. */
function arraysOf(args: Value[], from: number, context: CalcContext): ArrayValue[] | CalcError {
  const out: ArrayValue[] = [];
  for (let k = from; k < args.length; k++) {
    const array = toArray(args[k], context);
    if (isError(array)) return array;
    out.push(array);
  }
  return out;
}

// ---------------------------------------------------------------------------
// VSTACK and HSTACK

function stack(vertical: boolean): CalcFunction {
  return (args, context) => {
    // Too large in all: said before any of them is read.
    let along = 0;
    let across = 0;
    for (const arg of args) {
      const size = sizeOf(arg, context);
      along += vertical ? size.rows : size.cols;
      across = Math.max(across, vertical ? size.cols : size.rows);
    }
    if (!fits(along, across)) return ERROR_NUM;
    const arrays = arraysOf(args, 0, context);
    if (isError(arrays)) return arrays;
    let rows = 0;
    let cols = 0;
    for (const array of arrays) {
      if (vertical) {
        rows += array.rows;
        if (array.cols > cols) cols = array.cols;
      } else {
        cols += array.cols;
        if (array.rows > rows) rows = array.rows;
      }
    }
    if (!fits(rows, cols)) return ERROR_NUM;
    charge(context, rows * cols);
    const data: Scalar[] = new Array(rows * cols);
    let offset = 0;
    for (const array of arrays) {
      const source = array.data;
      const width = array.cols;
      if (vertical) {
        for (let r = 0; r < array.rows; r++) {
          const at = (offset + r) * cols;
          for (let c = 0; c < cols; c++) data[at + c] = c < width ? clean(source[r * width + c]) : ERROR_NA;
        }
        offset += array.rows;
      } else {
        for (let r = 0; r < rows; r++) {
          const at = r * cols + offset;
          for (let c = 0; c < width; c++) data[at + c] = r < array.rows ? clean(source[r * width + c]) : ERROR_NA;
        }
        offset += width;
      }
    }
    return { rows, cols, data };
  };
}

// ---------------------------------------------------------------------------
// TAKE, DROP, EXPAND

/** The rows and columns `top` to `top + rows` and `left` to `left + cols` of an array. */
function part(array: ArrayValue, top: number, rows: number, left: number, cols: number): ArrayValue {
  const data: Scalar[] = new Array(rows * cols);
  for (let r = 0; r < rows; r++) {
    const from = (top + r) * array.cols + left;
    for (let c = 0; c < cols; c++) data[r * cols + c] = clean(array.data[from + c]);
  }
  return { rows, cols, data };
}

/**
 * Where a part taken along one direction starts and how long it is: `count` from the start,
 * from the end when negative, everything when it is more than there is. Undefined for none.
 */
function taken(size: number, count: number): { start: number; length: number } | undefined {
  if (count === 0) return undefined;
  const length = Math.min(Math.abs(count), size);
  return { start: count > 0 ? 0 : size - length, length };
}

/** What is left along one direction when `count` is dropped from the start, or from the end when negative. */
function left(size: number, count: number): { start: number; length: number } | undefined {
  const gone = Math.min(Math.abs(count), size);
  if (gone === size) return undefined;
  return { start: count > 0 ? gone : 0, length: size - gone };
}

function cut(keep: boolean): CalcFunction {
  return (args, context) => {
    const array = toArray(args[0], context);
    if (isError(array)) return array;
    const rows = wholeArg(args, 1, keep ? array.rows : 0, context);
    if (typeof rows !== 'number') return rows;
    const cols = wholeArg(args, 2, keep ? array.cols : 0, context);
    if (typeof cols !== 'number') return cols;
    // UNSURE: DROP of 0 rows or columns drops none; Excel's page has the same words for TAKE and DROP (#CALC! for 0).
    const down = keep ? taken(array.rows, rows) : left(array.rows, rows);
    const across = keep ? taken(array.cols, cols) : left(array.cols, cols);
    if (!down || !across) return ERROR_CALC;
    charge(context, down.length * across.length);
    return part(array, down.start, down.length, across.start, across.length);
  };
}

const expand: CalcFunction = (args, context) => {
  const array = toArray(args[0], context);
  if (isError(array)) return array;
  const rows = wholeArg(args, 1, array.rows, context);
  if (typeof rows !== 'number') return rows;
  const cols = wholeArg(args, 2, array.cols, context);
  if (typeof cols !== 'number') return cols;
  if (rows < array.rows || cols < array.cols) return ERROR_VALUE;
  // UNSURE: Excel's page says only that expanding to a large array is #NUM!.
  if (!fits(rows, cols)) return ERROR_NUM;
  const pad = padArg(args, 3, context);
  charge(context, rows * cols);
  const data: Scalar[] = new Array(rows * cols);
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) data[r * cols + c] = r < array.rows && c < array.cols ? clean(array.data[r * array.cols + c]) : pad;
  }
  return { rows, cols, data };
};

// ---------------------------------------------------------------------------
// CHOOSEROWS and CHOOSECOLS

/** The positions (from 0) every argument from the second on names, each a number or an array of them; negative counts from the end. */
function positions(args: Value[], size: number, context: CalcContext): number[] | CalcError {
  const out: number[] = [];
  for (let k = 1; k < args.length; k++) {
    const list = toArray(args[k], context);
    if (isError(list)) return list;
    const count = list.rows * list.cols;
    for (let i = 0; i < count; i++) {
      const number = toNumber(clean(list.data[i]), context);
      if (typeof number !== 'number') return number;
      if (number !== number) return ERROR_VALUE;
      const at = number < 0 ? Math.ceil(number) : Math.floor(number);
      if (at === 0 || Math.abs(at) > size) return ERROR_VALUE;
      out.push(at > 0 ? at - 1 : size + at);
      if (out.length > MAX_ARRAY_CELLS) return ERROR_NUM;
    }
  }
  return out;
}

function choose(byRows: boolean): CalcFunction {
  return (args, context) => {
    // Too many picked in all: said before anything is read.
    const whole = sizeOf(args[0], context);
    let count = 0;
    for (let k = 1; k < args.length; k++) {
      const size = sizeOf(args[k], context);
      count += size.rows * size.cols;
    }
    if (!fits(count, byRows ? whole.cols : whole.rows)) return ERROR_NUM;
    const array = toArray(args[0], context);
    if (isError(array)) return array;
    const picked = positions(args, byRows ? array.rows : array.cols, context);
    if (isError(picked)) return picked;
    const rows = byRows ? picked.length : array.rows;
    const cols = byRows ? array.cols : picked.length;
    if (!fits(rows, cols)) return ERROR_NUM;
    charge(context, rows * cols);
    const data: Scalar[] = new Array(rows * cols);
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const from = byRows ? picked[r] * array.cols + c : r * array.cols + picked[c];
        data[r * cols + c] = clean(array.data[from]);
      }
    }
    return { rows, cols, data };
  };
}

// ---------------------------------------------------------------------------
// TOCOL, TOROW, WRAPROWS, WRAPCOLS

const IGNORE_BLANKS = 1;
const IGNORE_ERRORS = 2;

function line(asColumn: boolean): CalcFunction {
  return (args, context) => {
    const array = toArray(args[0], context);
    if (isError(array)) return array;
    const ignore = wholeArg(args, 1, 0, context);
    if (typeof ignore !== 'number') return ignore;
    if (ignore < 0 || ignore > 3) return ERROR_VALUE;
    let byColumn: boolean | CalcError = false;
    if (args.length > 2 && args[2] !== null) {
      const flag = single(args[2], context);
      byColumn = flag === undefined ? ERROR_VALUE : toBoolean(flag);
    }
    if (typeof byColumn !== 'boolean') return byColumn;
    const rows = array.rows;
    const cols = array.cols;
    const size = rows * cols;
    charge(context, size);
    const out: Scalar[] = [];
    for (let i = 0; i < size; i++) {
      const at = byColumn ? (i % rows) * cols + Math.floor(i / rows) : i;
      const value = clean(array.data[at]);
      // UNSURE: only an empty cell is a blank; text of no characters ("") is kept.
      if (value === null && (ignore & IGNORE_BLANKS) !== 0) continue;
      if (isError(value) && (ignore & IGNORE_ERRORS) !== 0) continue;
      out.push(value);
    }
    if (out.length === 0) return ERROR_CALC;
    return asColumn ? { rows: out.length, cols: 1, data: out } : { rows: 1, cols: out.length, data: out };
  };
}

function wrap(intoRows: boolean): CalcFunction {
  return (args, context) => {
    const vector = toArray(args[0], context);
    if (isError(vector)) return vector;
    if (vector.rows !== 1 && vector.cols !== 1) return ERROR_VALUE;
    const count = wholeArg(args, 1, 0, context);
    if (typeof count !== 'number') return count;
    if (count < 1) return ERROR_NUM;
    const size = vector.rows * vector.cols;
    // UNSURE: a count larger than the vector still makes a line of that many values, the rest of them padding.
    const lines = Math.ceil(size / count);
    const rows = intoRows ? lines : count;
    const cols = intoRows ? count : lines;
    if (!fits(rows, cols)) return ERROR_NUM;
    const pad = padArg(args, 2, context);
    charge(context, rows * cols);
    const data: Scalar[] = new Array(rows * cols);
    for (let i = 0; i < rows * cols; i++) {
      // The i-th value of the vector goes along a row (WRAPROWS) or down a column (WRAPCOLS).
      const at = intoRows ? i : (i % rows) * cols + Math.floor(i / rows);
      data[at] = i < size ? clean(vector.data[i]) : pad;
    }
    return { rows, cols, data };
  };
}

const all = true;
const firstOnly = (index: number) => index === 0;
const notFirst = (index: number) => index >= 1;

export const SHAPE_FUNCTIONS: { [name: string]: FunctionSpec } = {
  VSTACK: { call: stack(true), min: 1, max: 254, arrayArgs: all },
  HSTACK: { call: stack(false), min: 1, max: 254, arrayArgs: all },
  TAKE: { call: cut(true), min: 2, max: 3, arrayArgs: firstOnly, scalarArgs: notFirst },
  DROP: { call: cut(false), min: 2, max: 3, arrayArgs: firstOnly, scalarArgs: notFirst },
  EXPAND: { call: expand, min: 2, max: 4, arrayArgs: firstOnly, scalarArgs: notFirst },
  CHOOSEROWS: { call: choose(true), min: 2, max: 254, arrayArgs: all },
  CHOOSECOLS: { call: choose(false), min: 2, max: 254, arrayArgs: all },
  TOCOL: { call: line(true), min: 1, max: 3, arrayArgs: firstOnly, scalarArgs: notFirst },
  TOROW: { call: line(false), min: 1, max: 3, arrayArgs: firstOnly, scalarArgs: notFirst },
  WRAPROWS: { call: wrap(true), min: 2, max: 3, arrayArgs: firstOnly, scalarArgs: notFirst },
  WRAPCOLS: { call: wrap(false), min: 2, max: 3, arrayArgs: firstOnly, scalarArgs: notFirst },
};
