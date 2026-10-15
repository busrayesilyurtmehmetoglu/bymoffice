// Lookup functions and the simple dynamic array functions.
//
// Searches read a reference through the context and stop at the last used row
// or column of the sheet, so a whole-column table costs only what is in it. An
// exact search is a scan the first time and is answered from an index of the
// column once the same range is searched again in a recalculation; an
// approximate one is a binary search that trusts the data to be sorted, as
// Excel does.

import {
  ERROR_NA,
  ERROR_NUM,
  ERROR_REF,
  ERROR_VALUE,
  isError,
  isRef,
  toBoolean,
  type ArrayValue,
  type CalcContext,
  type CalcError,
  type CalcFunction,
  type FunctionSpec,
  type RefValue,
  type Scalar,
  type Value,
} from '../values';
import {
  clean,
  compilePattern,
  ERROR_CALC,
  first,
  foldCase,
  hasWildcard,
  lookupPositions,
  mapScalarsExcept,
  patternMatches,
  rangeIndex,
  scalarNumber,
  single,
  spec,
  toArray,
  toGrid,
  usedCols,
  usedRows,
  type Grid,
} from './helpers';

type Getter = (index: number) => Scalar;

/** Most cells of the row or column XLOOKUP returns: a whole column of a sheet. */
const MAX_LOOKUP_RESULT = 1 << 20;

/** Order of two values of the same kind (text without regard to case); undefined when the kinds differ. */
function compareSame(cell: Scalar, lookup: Scalar): number | undefined {
  if (typeof cell === 'number') {
    return typeof lookup === 'number' ? (cell < lookup ? -1 : cell > lookup ? 1 : 0) : undefined;
  }
  if (typeof cell === 'string') {
    if (typeof lookup !== 'string') return undefined;
    const a = foldCase(cell);
    const b = foldCase(lookup);
    return a < b ? -1 : a > b ? 1 : 0;
  }
  if (typeof cell === 'boolean') {
    return typeof lookup === 'boolean' ? (cell === lookup ? 0 : cell ? 1 : -1) : undefined;
  }
  return undefined;
}

/**
 * Binary search in data sorted ascending (or descending): the last position
 * whose value does not come after the lookup value, or -1. Cells of another
 * kind (and empty cells) are stepped over.
 */
function bisect(get: Getter, count: number, lookup: Scalar, descending: boolean): number {
  let low = 0;
  let high = count - 1;
  let result = -1;
  while (low <= high) {
    const middle = Math.floor((low + high) / 2);
    let probe = middle;
    let order: number | undefined;
    for (; probe >= low; probe--) {
      order = compareSame(get(probe), lookup);
      if (order !== undefined) break;
    }
    if (order === undefined) {
      low = middle + 1;
      continue;
    }
    if (descending) order = -order;
    if (order <= 0) {
      result = probe;
      low = middle + 1;
    } else high = probe - 1;
  }
  return result;
}

/** The next position after `from` holding a value of the lookup's kind, or -1. */
function nextComparable(get: Getter, count: number, from: number, lookup: Scalar): number {
  for (let i = from + 1; i < count; i++) if (compareSame(get(i), lookup) !== undefined) return i;
  return -1;
}

/** The values to search, when they are the cells of one row or column of a sheet. */
interface Lane {
  ref: RefValue;
  context: CalcContext;
}

/**
 * Exact search in a row or column of a sheet: the first (or last) position holding the
 * value, -1 when there is none. Text is compared without regard to case.
 */
function findInLane(lane: Lane, lookup: number | string | boolean, backward: boolean): number {
  const folded = typeof lookup === 'string' ? foldCase(lookup) : lookup;
  const index = rangeIndex(lane.ref, lane.context);
  if (index !== undefined) {
    const positions = lookupPositions(index, folded);
    if (positions !== undefined) return positions.length === 0 ? -1 : backward ? positions[positions.length - 1] : positions[0];
  }
  // One pass over the cells that hold something, read in bulk.
  const range = lane.ref.range;
  const vertical = range.left === range.right;
  let found = -1;
  lane.context.forEachCell(lane.ref.sheet, range, (cell, row, col) => {
    if (cell !== lookup && !(typeof folded === 'string' && typeof cell === 'string' && cell.length === folded.length && foldCase(cell) === folded)) return;
    const position = vertical ? row - range.top : col - range.left;
    if (found < 0 || (backward ? position > found : position < found)) found = position;
  });
  return found;
}

const MATCH_EXACT = 0;
const MATCH_SMALLER = -1;
const MATCH_LARGER = 1;
const MATCH_WILDCARD = 2;

/**
 * Position of a lookup value among `count` values, or -1. `matchMode`: exact,
 * exact or next smaller, exact or next larger, wildcard. `searchMode`: 1 from
 * the first, -1 from the last, 2 binary in ascending data, -2 in descending.
 */
function findMatch(get: Getter, count: number, lookup: Scalar, matchMode: number, searchMode: number, lane?: Lane): number {
  if (isError(lookup)) return -1;
  // An empty lookup value is 0: it finds a cell that holds 0, never an empty one.
  if (lookup === null) lookup = 0;
  if (searchMode === 2 || searchMode === -2) {
    const descending = searchMode === -2;
    // Empty cells at the end change where a binary search looks, and how many there are depends on
    // how far the sheet was once written (usedRows()): the search ends at the last cell that holds something.
    while (count > 0 && get(count - 1) === null) count--;
    const at = bisect(get, count, lookup, descending);
    const exact = at >= 0 && compareSame(get(at), lookup) === 0;
    if (exact) return at;
    if (matchMode === MATCH_EXACT || matchMode === MATCH_WILDCARD) return -1;
    // The search stops on the last value before the lookup value in sort order.
    const wantsBefore = descending ? matchMode === MATCH_LARGER : matchMode === MATCH_SMALLER;
    return wantsBefore ? at : nextComparable(get, count, at, lookup);
  }
  const backward = searchMode === -1;
  const pattern =
    matchMode === MATCH_WILDCARD && typeof lookup === 'string' && hasWildcard(lookup)
      ? compilePattern(foldCase(lookup))
      : undefined;
  if (pattern === undefined && (matchMode === MATCH_EXACT || matchMode === MATCH_WILDCARD)) {
    if (lane !== undefined) {
      const at = findInLane(lane, lookup, backward);
      return at < count ? at : -1;
    }
    // The common case, kept tight: text is folded only when its length already fits.
    const folded = typeof lookup === 'string' ? foldCase(lookup) : '';
    for (let step = 0; step < count; step++) {
      const i = backward ? count - 1 - step : step;
      const cell = get(i);
      if (cell === lookup) return i;
      if (typeof lookup === 'string' && typeof cell === 'string' && cell.length === folded.length && foldCase(cell) === folded) {
        return i;
      }
    }
    return -1;
  }
  let best = -1;
  let bestValue: Scalar = null;
  for (let step = 0; step < count; step++) {
    const i = backward ? count - 1 - step : step;
    const cell = get(i);
    if (pattern !== undefined) {
      if (typeof cell === 'string' && patternMatches(pattern, foldCase(cell))) return i;
      continue;
    }
    const order = compareSame(cell, lookup);
    if (order === undefined) continue;
    if (order === 0) return i;
    if ((matchMode === MATCH_SMALLER && order < 0) || (matchMode === MATCH_LARGER && order > 0)) {
      // Keep the closest value seen so far; the first one found wins a tie.
      const closer = best < 0 ? undefined : compareSame(cell, bestValue);
      if (best < 0 || (closer !== undefined && (matchMode === MATCH_SMALLER ? closer > 0 : closer < 0))) {
        best = i;
        bestValue = cell;
      }
    }
  }
  return best;
}

function optionalNumber(args: Value[], index: number, fallback: number, context: CalcContext): number | CalcError {
  if (index >= args.length) return fallback;
  const value = first(args[index], context);
  if (value === null) return fallback;
  const number = scalarNumber(value, context);
  return typeof number === 'number' ? Math.floor(number) : number;
}

function tableLookup(vertical: boolean): CalcFunction {
  return mapScalarsExcept(
    (index) => index === 1,
    (args, context) => {
      const lookup = args[0] as Scalar;
      if (isError(lookup)) return lookup;
      const table = toGrid(args[1], context);
      const position = scalarNumber(args[2] as Scalar, context);
      if (typeof position !== 'number') return position;
      // Left out, the search is approximate; left empty ("VLOOKUP(x,t,2,)") it is FALSE: exact.
      const approximate = args.length > 3 ? toBoolean(args[3] as Scalar) : true;
      if (typeof approximate !== 'boolean') return approximate;
      const index = Math.floor(position);
      if (!(index >= 1)) return ERROR_VALUE;
      if (index > (vertical ? table.cols : table.rows)) return ERROR_REF;
      const count = vertical ? usedRows(args[1], table, context) : usedCols(args[1], table, context);
      const get: Getter = vertical ? (i) => table.get(i, 0) : (i) => table.get(0, i);
      const at = approximate
        ? findMatch(get, count, lookup, MATCH_SMALLER, 2)
        : findMatch(get, count, lookup, MATCH_WILDCARD, 1, firstLane(args[1], vertical, context));
      if (at < 0) return ERROR_NA;
      return vertical ? table.get(at, index - 1) : table.get(index - 1, at);
    },
  );
}

/** The first column (or row) of a table that is a reference. */
function firstLane(table: Value, vertical: boolean, context: CalcContext): Lane | undefined {
  if (!isRef(table)) return undefined;
  const range = table.range;
  if (!(range.bottom >= range.top) || !(range.right >= range.left)) return undefined;
  return {
    ref: {
      ref: true,
      sheet: table.sheet,
      range: vertical
        ? { top: range.top, left: range.left, bottom: range.bottom, right: range.left }
        : { top: range.top, left: range.left, bottom: range.top, right: range.right },
    },
    context,
  };
}

/** A grid of one row or one column as a list; undefined when it is neither. */
function vector(value: Value, grid: Grid, context: CalcContext): { get: Getter; count: number; lane?: Lane } | undefined {
  const lane = isRef(value) && grid.rows >= 1 && grid.cols >= 1 ? { ref: value, context } : undefined;
  if (grid.cols === 1) return { get: (i) => grid.get(i, 0), count: usedRows(value, grid, context), lane };
  if (grid.rows === 1) return { get: (i) => grid.get(0, i), count: usedCols(value, grid, context), lane };
  return undefined;
}

const lookupFunction = mapScalarsExcept(
  (index) => index > 0,
  (args, context) => {
    const lookup = args[0] as Scalar;
    if (isError(lookup)) return lookup;
    const source = toGrid(args[1], context);
    if (args.length > 2) {
      const keys = vector(args[1], source, context);
      if (keys === undefined) return ERROR_NA;
      const at = findMatch(keys.get, keys.count, lookup, MATCH_SMALLER, 2);
      if (at < 0) return ERROR_NA;
      const results = toGrid(args[2], context);
      if (results.cols === 1) return at < results.rows ? results.get(at, 0) : ERROR_NA;
      if (results.rows === 1) return at < results.cols ? results.get(0, at) : ERROR_NA;
      return ERROR_NA;
    }
    // Array form: search the first row of a wide table or the first column of a tall one, answer from the last.
    if (source.rows === 0 || source.cols === 0) return ERROR_NA;
    if (source.cols > source.rows) {
      const at = findMatch((i) => source.get(0, i), usedCols(args[1], source, context), lookup, MATCH_SMALLER, 2);
      return at < 0 ? ERROR_NA : source.get(source.rows - 1, at);
    }
    const at = findMatch((i) => source.get(i, 0), usedRows(args[1], source, context), lookup, MATCH_SMALLER, 2);
    return at < 0 ? ERROR_NA : source.get(at, source.cols - 1);
  },
);

const match = mapScalarsExcept(
  (index) => index === 1,
  (args, context) => {
    const lookup = args[0] as Scalar;
    if (isError(lookup)) return lookup;
    // Left out, the match type is 1; left empty ("MATCH(x,r,)") it is 0: exact.
    let type: number | CalcError = 1;
    if (args.length > 2) {
      type = scalarNumber(first(args[2], context), context);
      if (typeof type !== 'number') return type;
    }
    const keys = vector(args[1], toGrid(args[1], context), context);
    if (keys === undefined) return ERROR_NA;
    let at: number;
    if (type === 0) at = findMatch(keys.get, keys.count, lookup, MATCH_WILDCARD, 1, keys.lane);
    else if (type > 0) at = findMatch(keys.get, keys.count, lookup, MATCH_SMALLER, 2);
    else at = findMatch(keys.get, keys.count, lookup, MATCH_LARGER, -2);
    return at < 0 ? ERROR_NA : at + 1;
  },
);

function validModes(matchMode: number, searchMode: number): boolean {
  return matchMode >= -1 && matchMode <= 2 && (searchMode === 1 || searchMode === -1 || searchMode === 2 || searchMode === -2);
}

const xmatch = mapScalarsExcept(
  (index) => index === 1,
  (args, context) => {
    const lookup = args[0] as Scalar;
    if (isError(lookup)) return lookup;
    const matchMode = optionalNumber(args, 2, 0, context);
    if (typeof matchMode !== 'number') return matchMode;
    const searchMode = optionalNumber(args, 3, 1, context);
    if (typeof searchMode !== 'number') return searchMode;
    if (!validModes(matchMode, searchMode)) return ERROR_VALUE;
    const keys = vector(args[1], toGrid(args[1], context), context);
    if (keys === undefined) return ERROR_VALUE;
    const at = findMatch(keys.get, keys.count, lookup, matchMode, searchMode, keys.lane);
    return at < 0 ? ERROR_NA : at + 1;
  },
);

const xlookup: CalcFunction = (args, context) => {
  const matchMode = optionalNumber(args, 4, 0, context);
  if (typeof matchMode !== 'number') return matchMode;
  const searchMode = optionalNumber(args, 5, 1, context);
  if (typeof searchMode !== 'number') return searchMode;
  if (!validModes(matchMode, searchMode)) return ERROR_VALUE;
  const keyGrid = toGrid(args[1], context);
  const keys = vector(args[1], keyGrid, context);
  if (keys === undefined) return ERROR_VALUE;
  const results = toGrid(args[2], context);
  const byRow = keyGrid.cols === 1;
  if (byRow ? results.rows !== keyGrid.rows : results.cols !== keyGrid.cols) return ERROR_VALUE;
  const find = (lookup: Scalar): number =>
    isError(lookup) ? -1 : findMatch(keys.get, keys.count, lookup, matchMode, searchMode, keys.lane);
  const missing = (): Scalar => (args.length > 3 && args[3] !== null ? first(args[3], context) : ERROR_NA);

  const lookup = single(args[0], context);
  if (lookup === undefined) {
    // Several lookup values: one answer each, the first cell of what a single lookup would return.
    const lookups = toArray(args[0], context);
    if (isError(lookups)) return lookups;
    const data: Scalar[] = new Array(lookups.rows * lookups.cols);
    for (let i = 0; i < data.length; i++) {
      const value = clean(lookups.data[i]);
      const at = find(value);
      data[i] = isError(value) ? value : at < 0 ? missing() : byRow ? results.get(at, 0) : results.get(0, at);
    }
    return { rows: lookups.rows, cols: lookups.cols, data };
  }
  if (isError(lookup)) return lookup;
  const at = find(lookup);
  if (at < 0) return missing();
  const full = byRow ? results.cols : results.rows;
  if (full === 1) return byRow ? results.get(at, 0) : results.get(0, at);
  // The whole row (or column) of the return range, empty cells included.
  if (!(full <= MAX_LOOKUP_RESULT)) return ERROR_NUM;
  if (context.tick !== undefined) context.tick(full >> 4);
  const size = full;
  const data: Scalar[] = new Array(size);
  for (let i = 0; i < size; i++) data[i] = byRow ? results.get(at, i) : results.get(i, at);
  return byRow ? { rows: 1, cols: size, data } : { rows: size, cols: 1, data };
};

// ---------------------------------------------------------------------------
// Dynamic arrays

function transposed(array: ArrayValue): ArrayValue {
  const data: Scalar[] = new Array(array.rows * array.cols);
  for (let r = 0; r < array.rows; r++) {
    for (let c = 0; c < array.cols; c++) data[c * array.rows + r] = clean(array.data[r * array.cols + c]);
  }
  return { rows: array.cols, cols: array.rows, data };
}

const transpose: CalcFunction = (args, context) => {
  const array = toArray(args[0], context);
  return isError(array) ? array : transposed(array);
};

function optionalFlag(args: Value[], index: number, context: CalcContext): boolean | CalcError {
  if (index >= args.length) return false;
  const value = first(args[index], context);
  return value === null ? false : toBoolean(value);
}

/** The rows `picked` of an array, in that order. */
function takeRows(array: ArrayValue, picked: number[]): ArrayValue {
  const cols = array.cols;
  const data: Scalar[] = new Array(picked.length * cols);
  for (let i = 0; i < picked.length; i++) {
    for (let c = 0; c < cols; c++) data[i * cols + c] = clean(array.data[picked[i] * cols + c]);
  }
  return { rows: picked.length, cols, data };
}

function rowKey(array: ArrayValue, row: number): string {
  let key = '';
  for (let c = 0; c < array.cols; c++) {
    const value = clean(array.data[row * array.cols + c]);
    // Each cell is tagged with its kind and length so that no two different rows share a key.
    if (typeof value === 'number') key += 'n' + value + ';';
    else if (typeof value === 'string') key += 's' + value.length + ':' + foldCase(value);
    else if (typeof value === 'boolean') key += value ? 't;' : 'f;';
    else if (value === null) key += 'e;';
    else key += 'x' + value.error + ';';
  }
  return key;
}

const unique: CalcFunction = (args, context) => {
  const source = toArray(args[0], context);
  if (isError(source)) return source;
  const byCol = optionalFlag(args, 1, context);
  if (typeof byCol !== 'boolean') return byCol;
  const once = optionalFlag(args, 2, context);
  if (typeof once !== 'boolean') return once;
  const array = byCol ? transposed(source) : source;
  const counts = new Map<string, number>();
  const keys: string[] = new Array(array.rows);
  for (let r = 0; r < array.rows; r++) {
    keys[r] = rowKey(array, r);
    counts.set(keys[r], (counts.get(keys[r]) || 0) + 1);
  }
  const picked: number[] = [];
  const taken = new Set<string>();
  for (let r = 0; r < array.rows; r++) {
    if (once ? counts.get(keys[r]) === 1 : !taken.has(keys[r])) picked.push(r);
    taken.add(keys[r]);
  }
  if (picked.length === 0) return ERROR_CALC;
  const result = takeRows(array, picked);
  return byCol ? transposed(result) : result;
};

/** Sort order across kinds: numbers, text, booleans, errors; empty cells always last. */
function sortRank(value: Scalar): number {
  if (typeof value === 'number') return 0;
  if (typeof value === 'string') return 1;
  if (typeof value === 'boolean') return 2;
  return value === null ? 4 : 3;
}

function sortCompare(a: Scalar, b: Scalar, direction: number): number {
  const ra = sortRank(a);
  const rb = sortRank(b);
  if (ra === 4 || rb === 4) return ra - rb;
  if (ra !== rb) return (ra - rb) * direction;
  const order = compareSame(a, b);
  return order === undefined ? 0 : order * direction;
}

const sort: CalcFunction = (args, context) => {
  const source = toArray(args[0], context);
  if (isError(source)) return source;
  const index = optionalNumber(args, 1, 1, context);
  if (typeof index !== 'number') return index;
  const direction = optionalNumber(args, 2, 1, context);
  if (typeof direction !== 'number') return direction;
  const byCol = optionalFlag(args, 3, context);
  if (typeof byCol !== 'boolean') return byCol;
  if (direction !== 1 && direction !== -1) return ERROR_VALUE;
  const array = byCol ? transposed(source) : source;
  if (!(index >= 1 && index <= array.cols)) return ERROR_VALUE;
  const order: number[] = new Array(array.rows);
  for (let r = 0; r < array.rows; r++) order[r] = r;
  const column = index - 1;
  const cols = array.cols;
  // Array.prototype.sort is stable: rows with equal keys keep their order.
  order.sort((x, y) =>
    sortCompare(clean(array.data[x * cols + column]), clean(array.data[y * cols + column]), direction) || x - y,
  );
  const result = takeRows(array, order);
  return byCol ? transposed(result) : result;
};

/**
 * SORTBY(array, by_array1, [sort_order1], [by_array2, sort_order2], ...): the rows (or, when
 * the keys are rows, the columns) of `array` in the order of the keys, the first key first.
 * Every key is one column as tall as the array, or one row as wide.
 */
const sortBy: CalcFunction = (args, context) => {
  const source = toArray(args[0], context);
  if (isError(source)) return source;
  const keys: { values: readonly Scalar[]; direction: number }[] = [];
  let byCol: boolean | undefined;
  for (let k = 1; k < args.length; k += 2) {
    const key = toArray(args[k], context);
    if (isError(key)) return key;
    const direction = k + 1 < args.length ? optionalNumber(args, k + 1, 1, context) : 1;
    if (typeof direction !== 'number') return direction;
    if (direction !== 1 && direction !== -1) return ERROR_VALUE;
    // A key of one cell sorts an array of one row or one column, either way.
    const column = key.cols === 1 && key.rows === source.rows && (key.rows > 1 || source.rows === 1);
    const row = !column && key.rows === 1 && key.cols === source.cols;
    if (!column && !row) return ERROR_VALUE;
    if (byCol === undefined) byCol = row;
    else if (byCol !== row) return ERROR_VALUE;
    keys.push({ values: key.data, direction });
  }
  const array = byCol ? transposed(source) : source;
  const order: number[] = new Array(array.rows);
  for (let r = 0; r < array.rows; r++) order[r] = r;
  order.sort((x, y) => {
    for (const key of keys) {
      const by = sortCompare(clean(key.values[x]), clean(key.values[y]), key.direction);
      if (by !== 0) return by;
    }
    return x - y;
  });
  const result = takeRows(array, order);
  return byCol ? transposed(result) : result;
};

const filter: CalcFunction = (args, context) => {
  const array = toArray(args[0], context);
  if (isError(array)) return array;
  const include = toArray(args[1], context);
  if (isError(include)) return include;
  const flags = include.rows * include.cols;
  const byRow = include.cols === 1 && include.rows === array.rows;
  const byCol = !byRow && include.rows === 1 && include.cols === array.cols;
  if (!byRow && !byCol) return ERROR_VALUE;
  const picked: number[] = [];
  for (let i = 0; i < flags; i++) {
    const keep = toBoolean(clean(include.data[i]));
    if (typeof keep !== 'boolean') return keep;
    if (keep) picked.push(i);
  }
  if (picked.length === 0) return args.length > 2 ? first(args[2], context) : ERROR_CALC;
  return byRow ? takeRows(array, picked) : transposed(takeRows(transposed(array), picked));
};

// The arguments named here take arrays: what is written there is calculated element by
// element in an ordinary formula too (LOOKUP(2,1/(A1:A9="x"),B1:B9) needs no array formula).
const second = (index: number) => index === 1;
const afterFirst = (index: number) => index >= 1;

export const LOOKUP_FUNCTIONS: { [name: string]: FunctionSpec } = {
  VLOOKUP: spec(tableLookup(true), 3, 4, second),
  HLOOKUP: spec(tableLookup(false), 3, 4, second),
  LOOKUP: spec(lookupFunction, 2, 3, afterFirst),
  MATCH: spec(match, 2, 3, second),
  XLOOKUP: spec(xlookup, 3, 6, true),
  XMATCH: spec(xmatch, 2, 4, second),
  TRANSPOSE: spec(transpose, 1, 1, true),
  UNIQUE: spec(unique, 1, 3, true),
  SORT: spec(sort, 1, 4, true),
  SORTBY: spec(sortBy, 2, 254, true),
  FILTER: spec(filter, 2, 3, true),
};
