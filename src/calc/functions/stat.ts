// Statistical functions: averages, counts, extremes, spread, order statistics,
// the ...IF family and simple regression.

import {
  checkNumber,
  ERROR_DIV0,
  ERROR_NA,
  ERROR_NUM,
  isArray,
  isError,
  isRef,
  type CalcContext,
  type CalcError,
  type CalcFunction,
  type FunctionSpec,
  type Scalar,
  type Value,
} from '../values';
import {
  clean,
  collectNumbers,
  compileConditions,
  countMatches,
  eachNumber,
  mapScalarsExcept,
  matchNumbers,
  parseNumberText,
  scalarNumber,
  spec,
  toArray,
  toGrid,
  type NumberRules,
} from './helpers';

const LOOSE: NumberRules = { loose: true };

// ---------------------------------------------------------------------------
// Kernels: a statistic of a list of numbers. SUBTOTAL and AGGREGATE reuse them.

export type Kernel = (numbers: number[]) => number | CalcError;

export function sum(numbers: number[]): number {
  let total = 0;
  for (let i = 0; i < numbers.length; i++) total += numbers[i];
  return total;
}

export function product(numbers: number[]): number {
  if (numbers.length === 0) return 0;
  let total = 1;
  for (let i = 0; i < numbers.length; i++) total *= numbers[i];
  return total;
}

export function mean(numbers: number[]): number | CalcError {
  return numbers.length === 0 ? ERROR_DIV0 : sum(numbers) / numbers.length;
}

export function maximum(numbers: number[]): number {
  if (numbers.length === 0) return 0;
  let best = numbers[0];
  for (let i = 1; i < numbers.length; i++) if (numbers[i] > best) best = numbers[i];
  return best;
}

export function minimum(numbers: number[]): number {
  if (numbers.length === 0) return 0;
  let best = numbers[0];
  for (let i = 1; i < numbers.length; i++) if (numbers[i] < best) best = numbers[i];
  return best;
}

/** Variance: of a sample (divided by n - 1) or of a whole population (divided by n). */
export function variance(numbers: number[], sample: boolean): number | CalcError {
  const count = numbers.length;
  if (count < (sample ? 2 : 1)) return ERROR_DIV0;
  const average = sum(numbers) / count;
  let squares = 0;
  for (let i = 0; i < count; i++) {
    const d = numbers[i] - average;
    squares += d * d;
  }
  return squares / (sample ? count - 1 : count);
}

export function deviation(numbers: number[], sample: boolean): number | CalcError {
  const result = variance(numbers, sample);
  return typeof result === 'number' ? Math.sqrt(result) : result;
}

function ascending(a: number, b: number): number {
  return a - b;
}

export function median(numbers: number[]): number | CalcError {
  const count = numbers.length;
  if (count === 0) return ERROR_NUM;
  const sorted = numbers.slice().sort(ascending);
  const half = count >> 1;
  return count % 2 === 1 ? sorted[half] : (sorted[half - 1] + sorted[half]) / 2;
}

/** The value that occurs most often; the one met first when several tie; #N/A when nothing repeats. */
export function mode(numbers: number[]): number | CalcError {
  const counts = new Map<number, number>();
  let most = 0;
  for (let i = 0; i < numbers.length; i++) {
    const count = (counts.get(numbers[i]) || 0) + 1;
    counts.set(numbers[i], count);
    if (count > most) most = count;
  }
  if (most < 2) return ERROR_NA;
  for (let i = 0; i < numbers.length; i++) if (counts.get(numbers[i]) === most) return numbers[i];
  return ERROR_NA;
}

export function largest(numbers: number[], k: number): number | CalcError {
  const rank = Math.floor(k);
  if (!(rank >= 1 && rank <= numbers.length)) return ERROR_NUM;
  return numbers.slice().sort(ascending)[numbers.length - rank];
}

export function smallest(numbers: number[], k: number): number | CalcError {
  const rank = Math.floor(k);
  if (!(rank >= 1 && rank <= numbers.length)) return ERROR_NUM;
  return numbers.slice().sort(ascending)[rank - 1];
}

function interpolate(sorted: number[], position: number): number {
  const low = Math.floor(position);
  const fraction = position - low;
  if (fraction === 0 || low + 1 >= sorted.length) return sorted[low];
  return sorted[low] + fraction * (sorted[low + 1] - sorted[low]);
}

export function percentileInc(numbers: number[], k: number): number | CalcError {
  const count = numbers.length;
  if (count === 0 || !(k >= 0 && k <= 1)) return ERROR_NUM;
  return interpolate(numbers.slice().sort(ascending), k * (count - 1));
}

export function percentileExc(numbers: number[], k: number): number | CalcError {
  const count = numbers.length;
  if (count === 0 || !(k > 0 && k < 1)) return ERROR_NUM;
  const position = k * (count + 1) - 1;
  if (position < 0 || position > count - 1) return ERROR_NUM;
  return interpolate(numbers.slice().sort(ascending), position);
}

export function quartileInc(numbers: number[], quart: number): number | CalcError {
  const q = Math.floor(quart);
  if (!(q >= 0 && q <= 4)) return ERROR_NUM;
  return percentileInc(numbers, q / 4);
}

export function quartileExc(numbers: number[], quart: number): number | CalcError {
  const q = Math.floor(quart);
  if (!(q >= 1 && q <= 3)) return ERROR_NUM;
  return percentileExc(numbers, q / 4);
}

/** COUNT: numbers inside ranges; a direct argument also counts when it is a boolean or numeric text. */
export function countNumbers(args: Value[], context: CalcContext, skipCell?: NumberRules['skipCell']): number {
  let count = 0;
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (isRef(arg)) {
      if (arg.range.bottom >= arg.range.top && arg.range.right >= arg.range.left) {
        const sheet = arg.sheet;
        context.forEachCell(sheet, arg.range, (value, row, col) => {
          if (typeof value === 'number' && !(skipCell !== undefined && skipCell(sheet, row, col))) count++;
        });
      }
    } else if (isArray(arg)) {
      for (let k = 0; k < arg.data.length; k++) if (typeof arg.data[k] === 'number') count++;
    } else if (typeof arg === 'string') {
      if (parseNumberText(arg, context) !== undefined) count++;
    } else if (!isError(arg)) count++;
  }
  return count;
}

/** COUNTA: everything that is not an empty cell, errors and empty text included. */
export function countValues(args: Value[], context: CalcContext, skipCell?: NumberRules['skipCell']): number {
  let count = 0;
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (isRef(arg)) {
      if (arg.range.bottom >= arg.range.top && arg.range.right >= arg.range.left) {
        const sheet = arg.sheet;
        context.forEachCell(sheet, arg.range, (value, row, col) => {
          if (value !== null && value !== undefined && !(skipCell !== undefined && skipCell(sheet, row, col))) count++;
        });
      }
    } else if (isArray(arg)) {
      for (let k = 0; k < arg.data.length; k++) if (arg.data[k] !== null && arg.data[k] !== undefined) count++;
    } else count++;
  }
  return count;
}

// ---------------------------------------------------------------------------
// Building functions from kernels

function fromNumbers(kernel: Kernel, rules?: NumberRules): CalcFunction {
  return (args, context) => {
    const numbers = collectNumbers(args, context, rules);
    if (isError(numbers)) return numbers;
    const result = kernel(numbers);
    return typeof result === 'number' ? checkNumber(result) : result;
  };
}

/** Functions of (numbers, k): the list is taken as given, k may be an array and then the result is one too. */
function withParameter(kernel: (numbers: number[], k: number) => number | CalcError): CalcFunction {
  return mapScalarsExcept(
    (index) => index === 0,
    (args, context) => {
      const numbers = collectNumbers([args[0]], context);
      if (isError(numbers)) return numbers;
      const k = scalarNumber(args[1] as Scalar, context);
      if (typeof k !== 'number') return k;
      const result = kernel(numbers, k);
      return typeof result === 'number' ? checkNumber(result) : result;
    },
  );
}

const average: CalcFunction = (args, context) => {
  let total = 0;
  let count = 0;
  const error = eachNumber(args, context, undefined, (value) => {
    total += value;
    count++;
  });
  if (error !== undefined) return error;
  return count === 0 ? ERROR_DIV0 : checkNumber(total / count);
};

function extreme(sign: number, rules?: NumberRules): CalcFunction {
  return (args, context) => {
    let best = 0;
    let found = false;
    const error = eachNumber(args, context, rules, (value) => {
      if (!found || (sign > 0 ? value > best : value < best)) best = value;
      found = true;
    });
    return error !== undefined ? error : best;
  };
}

function rank(averageTies: boolean): CalcFunction {
  return mapScalarsExcept(
    (index) => index === 1,
    (args, context) => {
      const value = scalarNumber(args[0] as Scalar, context);
      if (typeof value !== 'number') return value;
      let descending = true;
      if (args.length > 2) {
        const order = scalarNumber(args[2] as Scalar, context);
        if (typeof order !== 'number') return order;
        descending = order === 0;
      }
      let before = 0;
      let equal = 0;
      const error = eachNumber([args[1]], context, undefined, (other) => {
        if (other === value) equal++;
        else if (descending ? other > value : other < value) before++;
      });
      if (error !== undefined) return error;
      if (equal === 0) return ERROR_NA;
      return averageTies ? before + 1 + (equal - 1) / 2 : before + 1;
    },
  );
}

// ---------------------------------------------------------------------------
// The ...IF family

const countIf = mapScalarsExcept(
  (index) => index === 0,
  (args, context) => {
    const conditions = compileConditions(args, context);
    return isError(conditions) ? conditions : countMatches(conditions);
  },
);

const countIfs = mapScalarsExcept(
  (index) => index % 2 === 0,
  (args, context) => {
    const conditions = compileConditions(args, context);
    return isError(conditions) ? conditions : countMatches(conditions);
  },
);

type Fold = (numbers: { total: number; count: number; min: number; max: number }) => number | CalcError;

function conditional(valueIndex: number, fold: Fold): CalcFunction {
  // valueIndex 2: (range, criterion, [values]) as SUMIF; 0: (values, range, criterion, ...) as SUMIFS.
  const single = valueIndex === 2;
  return mapScalarsExcept(
    (index) => (single ? index !== 1 : index === 0 || index % 2 === 1),
    (args, context) => {
      const state = { total: 0, count: 0, min: Infinity, max: -Infinity };
      const pairs = single ? args.slice(0, 2) : args.slice(1);
      const values = single ? (args.length > 2 ? args[2] : args[0]) : args[0];
      const error = matchNumbers(values, pairs, context, single, (value) => {
        state.total += value;
        state.count++;
        if (value < state.min) state.min = value;
        if (value > state.max) state.max = value;
      });
      if (error !== undefined) return error;
      const result = fold(state);
      return typeof result === 'number' ? checkNumber(result) : result;
    },
  );
}

export const sumIf = conditional(2, (s) => s.total);
export const sumIfs = conditional(0, (s) => s.total);
const averageOf: Fold = (s) => (s.count === 0 ? ERROR_DIV0 : s.total / s.count);

const countBlank: CalcFunction = (args, context) => {
  const grid = toGrid(args[0], context);
  let filled = 0;
  grid.each((value) => {
    if (value !== '') filled++;
  });
  return grid.rows * grid.cols - filled;
};

// ---------------------------------------------------------------------------
// Two lists side by side

interface Pairs {
  xs: number[];
  ys: number[];
}

/** The positions where both lists hold a number. Lists of different sizes are #N/A. */
function numberPairs(ys: Value, xs: Value, context: CalcContext): Pairs | CalcError {
  const gy = toGrid(ys, context);
  const gx = toGrid(xs, context);
  if (gy.rows * gy.cols !== gx.rows * gx.cols) return ERROR_NA;
  const ay = toArray(ys, context);
  if (isError(ay)) return ay;
  const ax = toArray(xs, context);
  if (isError(ax)) return ax;
  const result: Pairs = { xs: [], ys: [] };
  const count = Math.min(ay.rows * ay.cols, ax.rows * ax.cols);
  for (let i = 0; i < count; i++) {
    const y = clean(ay.data[i]);
    const x = clean(ax.data[i]);
    if (isError(y)) return y;
    if (isError(x)) return x;
    if (typeof y === 'number' && typeof x === 'number') {
      result.ys.push(y);
      result.xs.push(x);
    }
  }
  return result;
}

function regression(kind: 'correl' | 'slope' | 'intercept'): CalcFunction {
  return (args, context) => {
    const pairs = numberPairs(args[0], args[1], context);
    if (isError(pairs)) return pairs;
    const count = pairs.xs.length;
    if (count === 0) return ERROR_DIV0;
    const meanX = sum(pairs.xs) / count;
    const meanY = sum(pairs.ys) / count;
    let sxx = 0;
    let syy = 0;
    let sxy = 0;
    for (let i = 0; i < count; i++) {
      const dx = pairs.xs[i] - meanX;
      const dy = pairs.ys[i] - meanY;
      sxx += dx * dx;
      syy += dy * dy;
      sxy += dx * dy;
    }
    if (kind === 'correl') return sxx === 0 || syy === 0 ? ERROR_DIV0 : checkNumber(sxy / Math.sqrt(sxx * syy));
    if (sxx === 0) return ERROR_DIV0;
    const slope = sxy / sxx;
    return checkNumber(kind === 'slope' ? slope : meanY - slope * meanX);
  };
}

const geometricMean: Kernel = (numbers) => {
  if (numbers.length === 0) return ERROR_NUM;
  let logs = 0;
  for (let i = 0; i < numbers.length; i++) {
    if (!(numbers[i] > 0)) return ERROR_NUM;
    logs += Math.log(numbers[i]);
  }
  return Math.exp(logs / numbers.length);
};

// ---------------------------------------------------------------------------

// Arguments that take an array in Excel: an expression there is calculated element by element.
const firstArg = (index: number) => index === 0;

const stdevS = spec(fromNumbers((n) => deviation(n, true)), 1);
const stdevP = spec(fromNumbers((n) => deviation(n, false)), 1);
const varS = spec(fromNumbers((n) => variance(n, true)), 1);
const varP = spec(fromNumbers((n) => variance(n, false)), 1);
const modeSingle = spec(fromNumbers(mode), 1);
const percentile = spec(withParameter(percentileInc), 2, 2, firstArg);
const quartile = spec(withParameter(quartileInc), 2, 2, firstArg);
const rankEq = spec(rank(false), 2, 3);

export const STAT_FUNCTIONS: { [name: string]: FunctionSpec } = {
  AVERAGE: spec(average, 1),
  AVERAGEA: spec(fromNumbers(mean, LOOSE), 1),
  AVERAGEIF: spec(conditional(2, averageOf), 2, 3),
  AVERAGEIFS: spec(conditional(0, averageOf), 3),
  COUNT: spec((args, context) => countNumbers(args, context), 1),
  COUNTA: spec((args, context) => countValues(args, context), 1),
  COUNTBLANK: spec(countBlank, 1, 1),
  COUNTIF: spec(countIf, 2, 2),
  COUNTIFS: spec(countIfs, 2),
  MIN: spec(extreme(-1), 1),
  MAX: spec(extreme(1), 1),
  MINA: spec(extreme(-1, LOOSE), 1),
  MAXA: spec(extreme(1, LOOSE), 1),
  MINIFS: spec(conditional(0, (s) => (s.count === 0 ? 0 : s.min)), 3),
  MAXIFS: spec(conditional(0, (s) => (s.count === 0 ? 0 : s.max)), 3),
  MEDIAN: spec(fromNumbers(median), 1),
  MODE: modeSingle,
  'MODE.SNGL': modeSingle,
  STDEV: stdevS,
  'STDEV.S': stdevS,
  'STDEV.P': stdevP,
  STDEVP: stdevP,
  VAR: varS,
  'VAR.S': varS,
  'VAR.P': varP,
  VARP: varP,
  LARGE: spec(withParameter(largest), 2, 2, firstArg),
  SMALL: spec(withParameter(smallest), 2, 2, firstArg),
  RANK: rankEq,
  'RANK.EQ': rankEq,
  'RANK.AVG': spec(rank(true), 2, 3),
  PERCENTILE: percentile,
  'PERCENTILE.INC': percentile,
  'PERCENTILE.EXC': spec(withParameter(percentileExc), 2, 2, firstArg),
  QUARTILE: quartile,
  'QUARTILE.INC': quartile,
  'QUARTILE.EXC': spec(withParameter(quartileExc), 2, 2, firstArg),
  CORREL: spec(regression('correl'), 2, 2, true),
  SLOPE: spec(regression('slope'), 2, 2, true),
  INTERCEPT: spec(regression('intercept'), 2, 2, true),
  GEOMEAN: spec(fromNumbers(geometricMean), 1),
};
