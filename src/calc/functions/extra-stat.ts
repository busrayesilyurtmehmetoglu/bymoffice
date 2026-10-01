// More statistical functions: the normal distribution and its inverse, linear
// forecast, the trimmed mean and frequency counts.

import {
  checkNumber,
  ERROR_DIV0,
  ERROR_NA,
  ERROR_NUM,
  isError,
  toBoolean,
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
  mapScalars,
  mapScalarsExcept,
  scalarNumber,
  spec,
  tidy,
  toArray,
  toGrid,
} from './helpers';

// ---------------------------------------------------------------------------
// The normal distribution

const SQRT_2PI = 2.5066282746310002;

function density(z: number): number {
  return Math.exp(-0.5 * z * z) / SQRT_2PI;
}

/**
 * P(Z <= z) of the standard normal distribution, to about 15 digits; the lower
 * tail keeps its relative accuracy. Near the middle a power series is summed
 * (all terms positive), further out the continued fraction of the tail.
 */
export function normalCdf(z: number): number {
  const x = Math.abs(z);
  if (!(x <= 40)) return z < 0 ? 0 : z > 0 ? 1 : NaN;
  let tail: number;
  if (x < 2.5) {
    let term = x;
    let total = x;
    for (let n = 1; n < 200; n++) {
      term *= (x * x) / (2 * n + 1);
      total += term;
      if (term <= total * 1e-17) break;
    }
    tail = 0.5 - density(x) * total;
  } else {
    let fraction = 0;
    for (let n = 200; n >= 1; n--) fraction = n / (x + fraction);
    tail = density(x) / (x + fraction);
  }
  return z > 0 ? 1 - tail : tail;
}

// Rational approximations of the inverse (P. J. Acklam), good to about 1e-9 relative.
const INV_A = [
  -3.969683028665376e1, 2.209460984245205e2, -2.759285104469687e2, 1.38357751867269e2, -3.066479806614716e1,
  2.506628277459239,
];
const INV_B = [-5.447609879822406e1, 1.615858368580409e2, -1.556989798598866e2, 6.680131188771972e1, -1.328068155288572e1];
const INV_C = [
  -7.784894002430293e-3, -3.223964580411365e-1, -2.400758277161838, -2.549732539343734, 4.374664141464968,
  2.938163982698783,
];
const INV_D = [7.784695709041462e-3, 3.224671290700398e-1, 2.445134137142996, 3.754408661907416];
const INV_LOW = 0.02425;

/** The z with P(Z <= z) = p, for 0 < p < 1. */
export function normalInverse(p: number): number {
  // The lower half is solved; there the tail is known to full relative accuracy.
  if (p > 0.5) return -normalInverse(1 - p);
  let x: number;
  if (p < INV_LOW) {
    const q = Math.sqrt(-2 * Math.log(p));
    x =
      (((((INV_C[0] * q + INV_C[1]) * q + INV_C[2]) * q + INV_C[3]) * q + INV_C[4]) * q + INV_C[5]) /
      ((((INV_D[0] * q + INV_D[1]) * q + INV_D[2]) * q + INV_D[3]) * q + 1);
  } else {
    const q = p - 0.5;
    const r = q * q;
    x =
      ((((((INV_A[0] * r + INV_A[1]) * r + INV_A[2]) * r + INV_A[3]) * r + INV_A[4]) * r + INV_A[5]) * q) /
      (((((INV_B[0] * r + INV_B[1]) * r + INV_B[2]) * r + INV_B[3]) * r + INV_B[4]) * r + 1);
  }
  // Halley steps against the distribution itself take the estimate to full precision.
  for (let i = 0; i < 4; i++) {
    const slope = density(x);
    if (!(slope > 0)) break;
    const step = (normalCdf(x) - p) / slope;
    if (step === 0 || !isFinite(step)) break;
    const scale = 1 + (x * step) / 2;
    x -= scale > 0.5 ? step / scale : step;
  }
  return x;
}

/** Arguments as numbers, the one at `flag` as a boolean (1 or 0); the first that is neither wins as an error. */
function numbers(args: Scalar[], flag: number, context: CalcContext): number[] | CalcError {
  const values: number[] = new Array(args.length);
  for (let i = 0; i < args.length; i++) {
    if (i === flag) {
      const truth = toBoolean(args[i]);
      if (typeof truth !== 'boolean') return truth;
      values[i] = truth ? 1 : 0;
    } else {
      const value = scalarNumber(args[i], context);
      if (typeof value !== 'number') return value;
      values[i] = value;
    }
  }
  return values;
}

const normDist = mapScalars((args, context) => {
  const values = numbers(args, 3, context);
  if (isError(values)) return values;
  const deviation = values[2];
  if (!(deviation > 0)) return ERROR_NUM;
  const z = (values[0] - values[1]) / deviation;
  return checkNumber(values[3] === 1 ? normalCdf(z) : density(z) / deviation);
});

const normStandardDist = mapScalars((args, context) => {
  const values = numbers(args, 1, context);
  if (isError(values)) return values;
  return checkNumber(values[1] === 1 ? normalCdf(values[0]) : density(values[0]));
});

const normStandardCumulative = mapScalars((args, context) => {
  const z = scalarNumber(args[0], context);
  return typeof z === 'number' ? checkNumber(normalCdf(z)) : z;
});

const normInverse = mapScalars((args, context) => {
  const values = numbers(args, -1, context);
  if (isError(values)) return values;
  if (!(values[0] > 0 && values[0] < 1) || !(values[2] > 0)) return ERROR_NUM;
  return checkNumber(values[1] + values[2] * normalInverse(values[0]));
});

const normStandardInverse = mapScalars((args, context) => {
  const p = scalarNumber(args[0], context);
  if (typeof p !== 'number') return p;
  return p > 0 && p < 1 ? checkNumber(normalInverse(p)) : ERROR_NUM;
});

// ---------------------------------------------------------------------------
// Two lists side by side (as CORREL and SLOPE read them in stat.ts)

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

const forecast = mapScalarsExcept(
  (index) => index > 0,
  (args, context) => {
    const at = scalarNumber(args[0] as Scalar, context);
    if (typeof at !== 'number') return at;
    const pairs = numberPairs(args[1], args[2], context);
    if (isError(pairs)) return pairs;
    const count = pairs.xs.length;
    if (count === 0) return ERROR_DIV0;
    let meanX = 0;
    let meanY = 0;
    for (let i = 0; i < count; i++) {
      meanX += pairs.xs[i];
      meanY += pairs.ys[i];
    }
    meanX /= count;
    meanY /= count;
    let sxx = 0;
    let sxy = 0;
    for (let i = 0; i < count; i++) {
      const dx = pairs.xs[i] - meanX;
      sxx += dx * dx;
      sxy += dx * (pairs.ys[i] - meanY);
    }
    if (sxx === 0) return ERROR_DIV0;
    return checkNumber(meanY + (sxy / sxx) * (at - meanX));
  },
);

// ---------------------------------------------------------------------------

function ascending(a: number, b: number): number {
  return a - b;
}

const trimMean = mapScalarsExcept(
  (index) => index === 0,
  (args, context) => {
    const values = collectNumbers([args[0]], context);
    if (isError(values)) return values;
    const percent = scalarNumber(args[1] as Scalar, context);
    if (typeof percent !== 'number') return percent;
    if (!(percent >= 0 && percent < 1) || values.length === 0) return ERROR_NUM;
    // The points left out are split evenly between the two ends: an odd number loses one.
    const cut = Math.floor(Math.floor(tidy(values.length * percent)) / 2);
    values.sort(ascending);
    let total = 0;
    for (let i = cut; i < values.length - cut; i++) total += values[i];
    return checkNumber(total / (values.length - 2 * cut));
  },
);

/** For every bin the count of values up to it and above the bin below; one more count for the rest. */
const frequency: CalcFunction = (args, context) => {
  const values = collectNumbers([args[0]], context);
  if (isError(values)) return values;
  const bins = collectNumbers([args[1]], context);
  if (isError(bins)) return bins;
  values.sort(ascending);
  // Bins are filled in ascending order but answered in the order they were given.
  const order: number[] = new Array(bins.length);
  for (let i = 0; i < order.length; i++) order[i] = i;
  order.sort((a, b) => bins[a] - bins[b] || a - b);
  const data: Scalar[] = new Array(bins.length + 1);
  let next = 0;
  for (let i = 0; i < order.length; i++) {
    const limit = bins[order[i]];
    const from = next;
    while (next < values.length && values[next] <= limit) next++;
    data[order[i]] = next - from;
  }
  data[bins.length] = values.length - next;
  return { rows: data.length, cols: 1, data };
};

const forecastSpec = spec(forecast, 3, 3, (index) => index > 0);
const normDistSpec = spec(normDist, 4, 4);
const normInvSpec = spec(normInverse, 3, 3);

export const EXTRA_STAT_FUNCTIONS: { [name: string]: FunctionSpec } = {
  FORECAST: forecastSpec,
  'FORECAST.LINEAR': forecastSpec,
  TRIMMEAN: spec(trimMean, 2, 2, (index) => index === 0),
  FREQUENCY: spec(frequency, 2, 2, true),
  'NORM.DIST': normDistSpec,
  NORMDIST: normDistSpec,
  'NORM.INV': normInvSpec,
  NORMINV: normInvSpec,
  'NORM.S.DIST': spec(normStandardDist, 2, 2),
  NORMSDIST: spec(normStandardCumulative, 1, 1),
  'NORM.S.INV': spec(normStandardInverse, 1, 1),
  NORMSINV: spec(normStandardInverse, 1, 1),
};
