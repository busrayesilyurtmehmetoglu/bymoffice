// More functions of numbers: the rest of the trigonometric family, means and
// deviations that stat.ts leaves out, the "A" variants of deviation and
// variance, and the functions of two lists side by side.
//
// Nothing here throws, whatever the arguments are.

import {
  checkNumber,
  ERROR_DIV0,
  ERROR_NUM,
  isError,
  type CalcError,
  type CalcFunction,
  type FunctionSpec,
} from '../values';
import { numberPairs, type Pairs } from './extra-stat';
import { collectNumbers, numeric, spec, type NumberRules } from './helpers';
import { deviation, STAT_FUNCTIONS, sum, variance, type Kernel } from './stat';

// ---------------------------------------------------------------------------
// Trigonometry

/** Excel answers #NUM! for COT, CSC and SEC of an angle this large or larger (2^27). */
const MAX_ANGLE = 134217728;

/** 1 / f(x); #DIV/0! where f is 0, #NUM! for an angle Excel does not take. */
function reciprocal(f: (x: number) => number, bounded: boolean): CalcFunction {
  return numeric((x) => {
    if (bounded && Math.abs(x) >= MAX_ANGLE) return ERROR_NUM;
    const value = f(x);
    return value === 0 ? ERROR_DIV0 : 1 / value;
  });
}

// ES2015 has these; a WebView old enough to lack them gets the formulas they stand for.
function asinh(x: number): number {
  if (typeof Math.asinh === 'function') return Math.asinh(x);
  return x === -Infinity ? x : Math.log(x + Math.sqrt(x * x + 1));
}

function acosh(x: number): number {
  if (typeof Math.acosh === 'function') return Math.acosh(x);
  return Math.log(x + Math.sqrt(x * x - 1));
}

function atanh(x: number): number {
  if (typeof Math.atanh === 'function') return Math.atanh(x);
  return Math.log((1 + x) / (1 - x)) / 2;
}

function sinh(x: number): number {
  if (typeof Math.sinh === 'function') return Math.sinh(x);
  return (Math.exp(x) - Math.exp(-x)) / 2;
}

function cosh(x: number): number {
  if (typeof Math.cosh === 'function') return Math.cosh(x);
  return (Math.exp(x) + Math.exp(-x)) / 2;
}

function tanh(x: number): number {
  if (typeof Math.tanh === 'function') return Math.tanh(x);
  const a = Math.exp(x);
  const b = Math.exp(-x);
  return a === Infinity ? 1 : b === Infinity ? -1 : (a - b) / (a + b);
}

// ---------------------------------------------------------------------------
// Means and deviations

const LOOSE: NumberRules = { loose: true };

function fromNumbers(kernel: Kernel, rules?: NumberRules): CalcFunction {
  return (args, context) => {
    const numbers = collectNumbers(args, context, rules);
    if (isError(numbers)) return numbers;
    const result = kernel(numbers);
    return typeof result === 'number' ? checkNumber(result) : result;
  };
}

/**
 * The harmonic mean; #NUM! when a number is 0 or less.
 * UNSURE: without any number the answer is #NUM! here, as GEOMEAN has it; not checked in Excel.
 */
const harmonicMean: Kernel = (numbers) => {
  if (numbers.length === 0) return ERROR_NUM;
  let reciprocals = 0;
  for (let i = 0; i < numbers.length; i++) {
    if (!(numbers[i] > 0)) return ERROR_NUM;
    reciprocals += 1 / numbers[i];
  }
  return numbers.length / reciprocals;
};

/** Sum of the squared distances from the mean. UNSURE: #NUM! without any number, as Excel documents for AVEDEV. */
const squaredDeviations: Kernel = (numbers) => {
  if (numbers.length === 0) return ERROR_NUM;
  const average = sum(numbers) / numbers.length;
  let squares = 0;
  for (let i = 0; i < numbers.length; i++) {
    const d = numbers[i] - average;
    squares += d * d;
  }
  return squares;
};

const averageDeviation: Kernel = (numbers) => {
  if (numbers.length === 0) return ERROR_NUM;
  const average = sum(numbers) / numbers.length;
  let distances = 0;
  for (let i = 0; i < numbers.length; i++) distances += Math.abs(numbers[i] - average);
  return distances / numbers.length;
};

/** n!!: the product of every second number down from n. */
function doubleFactorial(x: number): number | CalcError {
  const n = x < 0 ? Math.ceil(x) : Math.floor(x);
  if (n < -1) return ERROR_NUM;
  // 301!! no longer fits a number.
  if (n > 300) return ERROR_NUM;
  let result = 1;
  for (let k = n; k > 1; k -= 2) result *= k;
  return result;
}

// ---------------------------------------------------------------------------
// Two lists side by side

/**
 * A function of the positions where both lists hold a number. Lists of different sizes are
 * #N/A, an error in either list is the result.
 */
function ofPairs(kernel: (pairs: Pairs) => number | CalcError): CalcFunction {
  return (args, context) => {
    const pairs = numberPairs(args[0], args[1], context);
    if (isError(pairs)) return pairs;
    const result = kernel(pairs);
    return typeof result === 'number' ? checkNumber(result) : result;
  };
}

/** UNSURE: lists without one pair of numbers give 0 here; Excel was not asked. */
function sumOfPairs(term: (x: number, y: number) => number): CalcFunction {
  return ofPairs((pairs) => {
    let total = 0;
    for (let i = 0; i < pairs.xs.length; i++) total += term(pairs.ys[i], pairs.xs[i]);
    return total;
  });
}

function covariance(sample: boolean): CalcFunction {
  return ofPairs((pairs) => {
    const count = pairs.xs.length;
    if (count < (sample ? 2 : 1)) return ERROR_DIV0;
    const meanX = sum(pairs.xs) / count;
    const meanY = sum(pairs.ys) / count;
    let products = 0;
    for (let i = 0; i < count; i++) products += (pairs.xs[i] - meanX) * (pairs.ys[i] - meanY);
    return products / (sample ? count - 1 : count);
  });
}

const correl = STAT_FUNCTIONS.CORREL;

/** The square of the correlation of two lists. */
const rSquared: CalcFunction = (args, context) => {
  const r = correl.call(args, context);
  return typeof r === 'number' ? checkNumber(r * r) : r;
};

const covarianceP = spec(covariance(false), 2, 2, true);

export const EXTRA_MATH_FUNCTIONS: { [name: string]: FunctionSpec } = {
  ASINH: spec(numeric(asinh), 1, 1),
  ACOSH: spec(numeric((x) => (x < 1 ? ERROR_NUM : acosh(x))), 1, 1),
  ATANH: spec(numeric((x) => (x <= -1 || x >= 1 ? ERROR_NUM : atanh(x))), 1, 1),
  COT: spec(reciprocal(Math.tan, true), 1, 1),
  CSC: spec(reciprocal(Math.sin, true), 1, 1),
  SEC: spec(reciprocal(Math.cos, true), 1, 1),
  COTH: spec(reciprocal(tanh, false), 1, 1),
  CSCH: spec(reciprocal(sinh, false), 1, 1),
  SECH: spec(numeric((x) => 1 / cosh(x)), 1, 1),
  // The angle is between 0 and pi, as Excel has it: ACOT(-1) is 3*pi/4.
  ACOT: spec(numeric((x) => Math.PI / 2 - Math.atan(x)), 1, 1),
  ACOTH: spec(numeric((x) => (x >= -1 && x <= 1 ? ERROR_NUM : Math.log((x + 1) / (x - 1)) / 2)), 1, 1),
  SQRTPI: spec(numeric((x) => (x < 0 ? ERROR_NUM : Math.sqrt(x * Math.PI))), 1, 1),
  FACTDOUBLE: spec(numeric(doubleFactorial), 1, 1),
  HARMEAN: spec(fromNumbers(harmonicMean), 1),
  AVEDEV: spec(fromNumbers(averageDeviation), 1),
  DEVSQ: spec(fromNumbers(squaredDeviations), 1),
  STDEVA: spec(fromNumbers((n) => deviation(n, true), LOOSE), 1),
  STDEVPA: spec(fromNumbers((n) => deviation(n, false), LOOSE), 1),
  VARA: spec(fromNumbers((n) => variance(n, true), LOOSE), 1),
  VARPA: spec(fromNumbers((n) => variance(n, false), LOOSE), 1),
  SUMX2MY2: spec(sumOfPairs((x, y) => x * x - y * y), 2, 2, true),
  SUMX2PY2: spec(sumOfPairs((x, y) => x * x + y * y), 2, 2, true),
  SUMXMY2: spec(sumOfPairs((x, y) => (x - y) * (x - y)), 2, 2, true),
  PEARSON: correl,
  RSQ: spec(rSquared, 2, 2, true),
  COVAR: covarianceP,
  'COVARIANCE.P': covarianceP,
  'COVARIANCE.S': spec(covariance(true), 2, 2, true),
};
