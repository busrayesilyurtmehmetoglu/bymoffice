// Engineering functions: numbers written in base 2, 8 and 16 (ten digits,
// negative numbers in two's complement) and bitwise operations.

import {
  ERROR_NUM,
  ERROR_VALUE,
  isError,
  toText,
  type CalcContext,
  type CalcFunction,
  type FunctionSpec,
  type Scalar,
} from '../values';
import { mapScalars, numeric, scalarNumber, spec } from './helpers';

/** Digits a number of another base has at most. */
const DIGITS = 10;
const DIGIT_CHARS = '0123456789ABCDEF';

/** A number in another base: `bits` is the size of one digit (1, 3 or 4). */
function fromDecimal(bits: number): CalcFunction {
  const radix = Math.pow(2, bits);
  // Ten digits hold this many values; the upper half of them stands for the negative numbers.
  const modulus = Math.pow(2, bits * DIGITS);
  const half = modulus / 2;
  return mapScalars((args: Scalar[], context: CalcContext) => {
    const source = args[0];
    if (typeof source === 'boolean') return ERROR_VALUE;
    const given = scalarNumber(source, context);
    if (typeof given !== 'number') return given;
    const padded = args.length > 1 && args[1] !== null;
    let places = 0;
    if (padded) {
      if (typeof args[1] === 'boolean') return ERROR_VALUE;
      const count = scalarNumber(args[1], context);
      if (typeof count !== 'number') return count;
      places = Math.trunc(count);
    }
    const value = Math.trunc(given);
    if (!(value >= -half && value < half)) return ERROR_NUM;
    // A negative number always takes all ten digits; `places` is not looked at.
    let rest = value < 0 ? value + modulus : value;
    let text = '';
    for (let i = 0; i < DIGITS && (rest > 0 || i === 0); i++) {
      const digit = rest % radix;
      text = DIGIT_CHARS.charAt(digit) + text;
      rest = (rest - digit) / radix;
    }
    if (value < 0 || !padded) return text;
    if (!(places >= 1 && places <= DIGITS) || text.length > places) return ERROR_NUM;
    while (text.length < places) text = '0' + text;
    return text;
  });
}

function toDecimal(bits: number): CalcFunction {
  const radix = Math.pow(2, bits);
  const modulus = Math.pow(2, bits * DIGITS);
  const half = modulus / 2;
  return mapScalars((args: Scalar[]) => {
    const source = args[0];
    if (typeof source === 'boolean') return ERROR_VALUE;
    const text = toText(source);
    if (isError(text)) return text;
    if (text.length > DIGITS) return ERROR_NUM;
    let value = 0;
    for (let i = 0; i < text.length; i++) {
      const c = text.charCodeAt(i);
      // 0-9, A-F, a-f
      const digit = c >= 48 && c <= 57 ? c - 48 : c >= 65 && c <= 70 ? c - 55 : c >= 97 && c <= 102 ? c - 87 : radix;
      if (digit >= radix) return ERROR_NUM;
      value = value * radix + digit;
    }
    return value >= half ? value - modulus : value;
  });
}

/** Bitwise operators work on 32 bits; numbers of up to 48 bits are taken in two halves of 24. */
const HALF = 16777216;
const BIT_LIMIT = 281474976710656;

function bitwise(combine: (a: number, b: number) => number): CalcFunction {
  return numeric((a, b) => {
    if (!(a >= 0 && b >= 0 && a < BIT_LIMIT && b < BIT_LIMIT)) return ERROR_NUM;
    if (Math.floor(a) !== a || Math.floor(b) !== b) return ERROR_NUM;
    const high = combine(Math.floor(a / HALF), Math.floor(b / HALF));
    const low = combine(a % HALF, b % HALF);
    return high * HALF + low;
  });
}

export const EXTRA_ENGINEERING_FUNCTIONS: { [name: string]: FunctionSpec } = {
  DEC2BIN: spec(fromDecimal(1), 1, 2),
  DEC2OCT: spec(fromDecimal(3), 1, 2),
  DEC2HEX: spec(fromDecimal(4), 1, 2),
  BIN2DEC: spec(toDecimal(1), 1, 1),
  OCT2DEC: spec(toDecimal(3), 1, 1),
  HEX2DEC: spec(toDecimal(4), 1, 1),
  BITAND: spec(bitwise((a, b) => a & b), 2, 2),
  BITOR: spec(bitwise((a, b) => a | b), 2, 2),
  BITXOR: spec(bitwise((a, b) => a ^ b), 2, 2),
};
