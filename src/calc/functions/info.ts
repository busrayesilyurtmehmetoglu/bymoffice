// Information functions: what kind of value something is.

import {
  ERROR_NA,
  ERROR_VALUE,
  isArray,
  isError,
  isRef,
  type CalcFunction,
  type ErrorCode,
  type FunctionSpec,
  type Scalar,
} from '../values';
import { mapScalars, scalarNumber, single, spec } from './helpers';

function test(check: (value: Scalar) => boolean): CalcFunction {
  return mapScalars((args) => check(args[0]));
}

function parity(odd: boolean): CalcFunction {
  return mapScalars((args, context) => {
    // Booleans are not numbers here, unlike in arithmetic.
    if (typeof args[0] === 'boolean') return ERROR_VALUE;
    const number = scalarNumber(args[0], context);
    if (typeof number !== 'number') return number;
    const whole = number < 0 ? Math.ceil(number) : Math.floor(number);
    return (Math.abs(whole) % 2 === 1) === odd;
  });
}

const ERROR_TYPES: { [code in ErrorCode]: number } = {
  '#NULL!': 1,
  '#DIV/0!': 2,
  '#VALUE!': 3,
  '#REF!': 4,
  '#NAME?': 5,
  '#NUM!': 6,
  '#N/A': 7,
  '#SPILL!': 9,
  '#CALC!': 14,
};

const errorType = mapScalars((args) => {
  const value = args[0];
  if (!isError(value)) return ERROR_NA;
  const type = ERROR_TYPES[value.error];
  return typeof type === 'number' ? type : ERROR_NA;
});

const type: CalcFunction = (args, context) => {
  const arg = args[0];
  const value = single(arg, context);
  // A range of several cells or an array is type 64.
  if (value === undefined) return isRef(arg) || isArray(arg) ? 64 : ERROR_VALUE;
  if (typeof value === 'string') return 2;
  if (typeof value === 'boolean') return 4;
  if (isError(value)) return 16;
  return 1;
};

export const INFO_FUNCTIONS: { [name: string]: FunctionSpec } = {
  ISNUMBER: spec(test((value) => typeof value === 'number'), 1, 1),
  ISTEXT: spec(test((value) => typeof value === 'string'), 1, 1),
  ISNONTEXT: spec(test((value) => typeof value !== 'string'), 1, 1),
  ISLOGICAL: spec(test((value) => typeof value === 'boolean'), 1, 1),
  ISERROR: spec(test((value) => isError(value)), 1, 1),
  ISERR: spec(test((value) => isError(value) && value.error !== '#N/A'), 1, 1),
  ISNA: spec(test((value) => isError(value) && value.error === '#N/A'), 1, 1),
  ISEVEN: spec(parity(false), 1, 1),
  ISODD: spec(parity(true), 1, 1),
  NA: spec(() => ERROR_NA, 0, 0),
  'ERROR.TYPE': spec(errorType, 1, 1),
  TYPE: spec(type, 1, 1),
  TRUE: spec(() => true, 0, 0),
  FALSE: spec(() => false, 0, 0),
};
