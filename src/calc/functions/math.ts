// Math functions: sums and products, rounding in all its forms, powers and
// logarithms, trigonometry, combinatorics, SUBTOTAL and AGGREGATE.

import { parseFormula, type Node } from '../parse';
import {
  checkNumber,
  ERROR_DIV0,
  ERROR_NUM,
  ERROR_VALUE,
  isError,
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
  eachNumber,
  first,
  mapScalars,
  numeric,
  roundTo,
  scalarNumber,
  spec,
  tidy,
  toArray,
  toGrid,
  type NumberRules,
} from './helpers';
import {
  countNumbers,
  countValues,
  deviation,
  largest,
  maximum,
  mean,
  median,
  minimum,
  mode,
  percentileExc,
  percentileInc,
  product,
  quartileExc,
  quartileInc,
  smallest,
  sum,
  sumIf,
  sumIfs,
  variance,
  type Kernel,
} from './stat';

/** Integers above this are not exact in a double. */
const MAX_EXACT = 9007199254740992;

function fold(start: number, step: (total: number, value: number) => number, empty?: number): CalcFunction {
  return (args, context) => {
    let total = start;
    let count = 0;
    const error = eachNumber(args, context, undefined, (value) => {
      total = step(total, value);
      count++;
    });
    if (error !== undefined) return error;
    return checkNumber(count === 0 && empty !== undefined ? empty : total);
  };
}

const sumProduct: CalcFunction = (args, context) => {
  const shape = toGrid(args[0], context);
  const arrays = [];
  for (let i = 0; i < args.length; i++) {
    const grid = toGrid(args[i], context);
    if (grid.rows !== shape.rows || grid.cols !== shape.cols) return ERROR_VALUE;
    const array = toArray(args[i], context);
    if (isError(array)) return array;
    arrays.push(array);
  }
  // Whole columns are cut to the used rows; what lies beyond is empty and adds nothing.
  let rows = arrays[0].rows;
  let cols = arrays[0].cols;
  for (let i = 1; i < arrays.length; i++) {
    rows = Math.min(rows, arrays[i].rows);
    cols = Math.min(cols, arrays[i].cols);
  }
  let total = 0;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      let term = 1;
      for (let i = 0; i < arrays.length; i++) {
        const value = clean(arrays[i].data[r * arrays[i].cols + c]);
        if (isError(value)) return value;
        // Text, booleans and empty cells count as zero.
        term *= typeof value === 'number' ? value : 0;
      }
      total += term;
    }
  }
  return checkNumber(total);
};

function gcdOf(a: number, b: number): number {
  let x = a;
  let y = b;
  // Euclid on integers below 2^53 ends within a few hundred steps; the cap is only a safety net.
  for (let i = 0; i < 4096 && y !== 0; i++) {
    const t = x % y;
    x = y;
    y = t;
  }
  return x;
}

function integers(kind: 'gcd' | 'lcm'): CalcFunction {
  return (args, context) => {
    let result = kind === 'gcd' ? 0 : 1;
    let invalid = false;
    const error = eachNumber(args, context, undefined, (value) => {
      const n = Math.floor(value);
      if (!(n >= 0 && n < MAX_EXACT)) invalid = true;
      if (invalid) return;
      if (kind === 'gcd') result = gcdOf(result, n);
      else if (n === 0 || result === 0) result = 0;
      else {
        result = (result / gcdOf(result, n)) * n;
        if (!(result < MAX_EXACT)) invalid = true;
      }
    });
    if (error !== undefined) return error;
    return invalid ? ERROR_NUM : result;
  };
}

/** Rounds x to a multiple of `step` by `round` (Math.ceil or Math.floor), free of floating point noise. */
function toMultiple(x: number, step: number, round: (value: number) => number): number {
  return tidy(round(tidy(x / step)) * step);
}

function ceiling(x: number, significance: number): number | CalcError {
  if (x === 0 || significance === 0) return 0;
  if (x > 0 && significance < 0) return ERROR_NUM;
  if (x < 0 && significance < 0) return -toMultiple(-x, -significance, Math.ceil);
  return toMultiple(x, significance, Math.ceil);
}

function floor(x: number, significance: number): number | CalcError {
  if (x === 0) return 0;
  if (significance === 0) return ERROR_DIV0;
  if (x > 0 && significance < 0) return ERROR_NUM;
  if (x < 0 && significance < 0) return -toMultiple(-x, -significance, Math.floor);
  return toMultiple(x, significance, Math.floor);
}

function factorial(n: number): number {
  let result = 1;
  for (let i = 2; i <= n && i <= 171; i++) result *= i;
  return result;
}

function combinations(n: number, k: number): number {
  const m = Math.min(k, n - k);
  let result = 1;
  // The product overflows to Infinity long before the loop gets long.
  for (let i = 1; i <= m && isFinite(result); i++) result = (result * (n - m + i)) / i;
  return Math.round(result);
}

function permutations(n: number, k: number): number {
  let result = 1;
  for (let i = 0; i < k && isFinite(result); i++) result *= n - i;
  return result;
}

function awayFromZero(x: number, parity: number): number {
  const abs = Math.abs(x);
  let n = Math.ceil(tidy(abs));
  if (n % 2 !== parity) n++;
  return x < 0 ? -n : n;
}

// ---------------------------------------------------------------------------
// SUBTOTAL and AGGREGATE

const KERNELS: { [kind: number]: Kernel } = {
  1: mean,
  4: maximum,
  5: minimum,
  6: product,
  7: (n) => deviation(n, true),
  8: (n) => deviation(n, false),
  9: sum,
  10: (n) => variance(n, true),
  11: (n) => variance(n, false),
  12: median,
  13: mode,
};

const KERNELS_K: { [kind: number]: (numbers: number[], k: number) => number | CalcError } = {
  14: largest,
  15: smallest,
  16: percentileInc,
  17: quartileInc,
  18: percentileExc,
  19: quartileExc,
};

function aggregate(kind: number, values: Value[], context: CalcContext, rules: NumberRules | undefined): Scalar {
  const skipCell = rules === undefined ? undefined : rules.skipCell;
  if (kind === 2) return countNumbers(values, context, skipCell);
  if (kind === 3) return countValues(values, context, skipCell);
  const kernel = KERNELS[kind];
  if (kernel === undefined) return ERROR_VALUE;
  const numbers = collectNumbers(values, context, rules);
  if (isError(numbers)) return numbers;
  const result = kernel(numbers);
  return typeof result === 'number' ? checkNumber(result) : result;
}

/** Whether formulas call SUBTOTAL or AGGREGATE, by their text. */
let subtotalFormulas = new Map<string, boolean>();
const MAX_REMEMBERED = 4096;

function callsSubtotal(text: string): boolean {
  const known = subtotalFormulas.get(text);
  if (known !== undefined) return known;
  let found = false;
  const upper = text.toUpperCase();
  if (upper.indexOf('SUBTOTAL') >= 0 || upper.indexOf('AGGREGATE') >= 0) {
    // The words may also stand in a text constant or a name: only a call counts.
    let root: Node | undefined;
    try {
      root = parseFormula(text);
    } catch (error) {
      root = undefined;
    }
    const nodes: Node[] = root === undefined ? [] : [root];
    while (!found && nodes.length > 0) {
      const node = nodes.pop() as Node;
      if (node.type === 'call') {
        if (node.name === 'SUBTOTAL' || node.name === 'AGGREGATE') found = true;
        for (let i = 0; i < node.args.length; i++) {
          const arg = node.args[i];
          if (arg !== null) nodes.push(arg);
        }
      } else if (node.type === 'unary') nodes.push(node.operand);
      else if (node.type === 'binary') nodes.push(node.left, node.right);
    }
  }
  if (subtotalFormulas.size >= MAX_REMEMBERED) subtotalFormulas = new Map();
  subtotalFormulas.set(text, found);
  return found;
}

/** The test for cells that hold a subtotal of their own, which SUBTOTAL and AGGREGATE pass over; undefined when the context cannot tell. */
function nestedSubtotals(context: CalcContext): ((sheet: number, row: number, col: number) => boolean) | undefined {
  if (context.cellFormula === undefined) return undefined;
  return (sheet, row, col) => {
    const text = (context.cellFormula as NonNullable<CalcContext['cellFormula']>)(sheet, row, col);
    return text !== undefined && callsSubtotal(text);
  };
}

// Both ranges of function numbers leave out the cells of the ranges that hold a SUBTOTAL
// (or AGGREGATE) formula of their own, so that subtotals can be stacked without counting
// anything twice. Function numbers 101 to 111 also leave out rows hidden by hand, and both
// ranges rows hidden by a filter; the context does not tell which rows are hidden, so
// those are included.
const subtotal: CalcFunction = (args, context) => {
  const code = scalarNumber(first(args[0], context), context);
  if (typeof code !== 'number') return code;
  let kind = Math.floor(code);
  if (kind >= 101 && kind <= 111) kind -= 100;
  if (!(kind >= 1 && kind <= 11)) return ERROR_VALUE;
  return aggregate(kind, args.slice(1), context, { skipCell: nestedSubtotals(context) });
};

// Options 0 to 7 choose what to leave out: nested SUBTOTAL and AGGREGATE results (0 to 3),
// hidden rows (1, 3, 5, 7; not known here), errors (2, 3, 6, 7).
const aggregateFunction: CalcFunction = (args, context) => {
  const code = scalarNumber(first(args[0], context), context);
  if (typeof code !== 'number') return code;
  const option = scalarNumber(first(args[1], context), context);
  if (typeof option !== 'number') return option;
  const kind = Math.floor(code);
  const choice = Math.floor(option);
  if (!(kind >= 1 && kind <= 19) || !(choice >= 0 && choice <= 7)) return ERROR_VALUE;
  const rules: NumberRules = {};
  if (choice === 2 || choice === 3 || choice === 6 || choice === 7) rules.skipErrors = true;
  if (choice <= 3) rules.skipCell = nestedSubtotals(context);
  if (kind <= 13) return aggregate(kind, args.slice(2), context, rules);
  if (args.length !== 4) return ERROR_VALUE;
  const k = scalarNumber(first(args[3], context), context);
  if (typeof k !== 'number') return k;
  const numbers = collectNumbers([args[2]], context, rules);
  if (isError(numbers)) return numbers;
  const result = KERNELS_K[kind](numbers, k);
  return typeof result === 'number' ? checkNumber(result) : result;
};

// ---------------------------------------------------------------------------

const random: CalcFunction = (_args, context) => {
  const value = context.random();
  // The contract promises [0, 1); a generator that breaks it must not leak NaN into a sheet.
  return value >= 0 && value < 1 ? value : 0;
};

const randomBetween = mapScalars((args, context) => {
  const low = scalarNumber(args[0], context);
  if (typeof low !== 'number') return low;
  const high = scalarNumber(args[1], context);
  if (typeof high !== 'number') return high;
  const bottom = Math.ceil(low);
  const top = Math.floor(high);
  if (!(bottom <= top) || !isFinite(bottom) || !isFinite(top)) return ERROR_NUM;
  const roll = context.random();
  const unit = roll >= 0 && roll < 1 ? roll : 0;
  return Math.min(top, bottom + Math.floor(unit * (top - bottom + 1)));
});

export const MATH_FUNCTIONS: { [name: string]: FunctionSpec } = {
  SUM: spec(fold(0, (total, value) => total + value), 1),
  SUMIF: spec(sumIf, 2, 3),
  SUMIFS: spec(sumIfs, 3),
  SUMPRODUCT: spec(sumProduct, 1, undefined, true),
  SUMSQ: spec(fold(0, (total, value) => total + value * value), 1),
  PRODUCT: spec(fold(1, (total, value) => total * value, 0), 1),
  ABS: spec(numeric(Math.abs), 1, 1),
  SIGN: spec(numeric((x) => (x > 0 ? 1 : x < 0 ? -1 : 0)), 1, 1),
  INT: spec(numeric(Math.floor), 1, 1),
  TRUNC: spec(numeric((x, digits) => roundTo(x, digits === undefined ? 0 : digits, -1)), 1, 2),
  ROUND: spec(numeric((x, digits) => roundTo(x, digits, 0)), 2, 2),
  ROUNDUP: spec(numeric((x, digits) => roundTo(x, digits, 1)), 2, 2),
  ROUNDDOWN: spec(numeric((x, digits) => roundTo(x, digits, -1)), 2, 2),
  MROUND: spec(
    numeric((x, multiple) => {
      if (x === 0 || multiple === 0) return 0;
      if (x < 0 !== multiple < 0) return ERROR_NUM;
      return tidy(Math.floor(tidy(x / multiple) + 0.5) * multiple);
    }),
    2,
    2,
  ),
  CEILING: spec(numeric(ceiling), 2, 2),
  'CEILING.MATH': spec(
    numeric((x, significance, direction) => {
      const step = Math.abs(significance === undefined ? 1 : significance);
      if (x === 0 || step === 0) return 0;
      // Negative numbers go toward zero unless the mode asks for away from it.
      return toMultiple(x, step, x < 0 && direction !== undefined && direction !== 0 ? Math.floor : Math.ceil);
    }),
    1,
    3,
  ),
  FLOOR: spec(numeric(floor), 2, 2),
  'FLOOR.MATH': spec(
    numeric((x, significance, direction) => {
      const step = Math.abs(significance === undefined ? 1 : significance);
      if (x === 0 || step === 0) return 0;
      return toMultiple(x, step, x < 0 && direction !== undefined && direction !== 0 ? Math.ceil : Math.floor);
    }),
    1,
    3,
  ),
  EVEN: spec(numeric((x) => awayFromZero(x, 0)), 1, 1),
  ODD: spec(numeric((x) => (x === 0 ? 1 : awayFromZero(x, 1))), 1, 1),
  MOD: spec(
    numeric((n, d) => {
      if (d === 0) return ERROR_DIV0;
      // The result takes the sign of the divisor.
      const result = n - d * Math.floor(n / d);
      return Math.abs(result) >= Math.abs(d) || (result !== 0 && result < 0 !== d < 0) ? 0 : result;
    }),
    2,
    2,
  ),
  QUOTIENT: spec(numeric((n, d) => (d === 0 ? ERROR_DIV0 : Math.trunc(n / d))), 2, 2),
  POWER: spec(
    numeric((base, exponent) => {
      if (base === 0 && exponent === 0) return ERROR_NUM;
      if (base === 0 && exponent < 0) return ERROR_DIV0;
      return Math.pow(base, exponent);
    }),
    2,
    2,
  ),
  SQRT: spec(numeric((x) => (x < 0 ? ERROR_NUM : Math.sqrt(x))), 1, 1),
  EXP: spec(numeric(Math.exp), 1, 1),
  LN: spec(numeric((x) => (x > 0 ? Math.log(x) : ERROR_NUM)), 1, 1),
  LOG: spec(
    numeric((x, base) => {
      const b = base === undefined ? 10 : base;
      if (!(x > 0) || !(b > 0)) return ERROR_NUM;
      if (b === 1) return ERROR_DIV0;
      return b === 10 ? Math.log10(x) : Math.log(x) / Math.log(b);
    }),
    1,
    2,
  ),
  LOG10: spec(numeric((x) => (x > 0 ? Math.log10(x) : ERROR_NUM)), 1, 1),
  PI: spec(() => Math.PI, 0, 0),
  SIN: spec(numeric(Math.sin), 1, 1),
  COS: spec(numeric(Math.cos), 1, 1),
  TAN: spec(numeric(Math.tan), 1, 1),
  ASIN: spec(numeric(Math.asin), 1, 1),
  ACOS: spec(numeric(Math.acos), 1, 1),
  ATAN: spec(numeric(Math.atan), 1, 1),
  ATAN2: spec(numeric((x, y) => (x === 0 && y === 0 ? ERROR_DIV0 : Math.atan2(y, x))), 2, 2),
  SINH: spec(numeric(Math.sinh), 1, 1),
  COSH: spec(numeric(Math.cosh), 1, 1),
  TANH: spec(numeric(Math.tanh), 1, 1),
  DEGREES: spec(numeric((x) => (x * 180) / Math.PI), 1, 1),
  RADIANS: spec(numeric((x) => (x * Math.PI) / 180), 1, 1),
  FACT: spec(
    numeric((x) => {
      const n = Math.floor(x);
      return n < 0 || n > 170 ? ERROR_NUM : factorial(n);
    }),
    1,
    1,
  ),
  COMBIN: spec(
    numeric((x, y) => {
      const n = Math.floor(x);
      const k = Math.floor(y);
      return n < 0 || k < 0 || n < k ? ERROR_NUM : combinations(n, k);
    }),
    2,
    2,
  ),
  PERMUT: spec(
    numeric((x, y) => {
      const n = Math.floor(x);
      const k = Math.floor(y);
      return n < 0 || k < 0 || n < k ? ERROR_NUM : permutations(n, k);
    }),
    2,
    2,
  ),
  GCD: spec(integers('gcd'), 1),
  LCM: spec(integers('lcm'), 1),
  RAND: spec(random, 0, 0),
  RANDBETWEEN: spec(randomBetween, 2, 2),
  SUBTOTAL: spec(subtotal, 2),
  // The array form (function numbers 14 to 19) takes an expression: AGGREGATE(14,6,A1:A9/(B1:B9="x"),1).
  AGGREGATE: spec(aggregateFunction, 3, undefined, (index) => index === 2),
};
