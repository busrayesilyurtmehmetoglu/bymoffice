// The function library: every worksheet function by its upper-case name.
//
// Functions that need more than evaluated arguments (IF and its relatives,
// ROW, OFFSET, INDIRECT, INDEX, LET, LAMBDA ...) are special forms of the
// evaluator and are not listed here.

import { ERROR_NUM, ERROR_VALUE, isArray, type FunctionSpec, type Scalar, type Value } from '../values';
import { DATE_FUNCTIONS } from './date';
import { EXTRA_FUNCTIONS } from './extra';
import { FINANCIAL_FUNCTIONS } from './financial';
import { INFO_FUNCTIONS } from './info';
import { LOOKUP_FUNCTIONS } from './lookup';
import { MATH_FUNCTIONS } from './math';
import { STAT_FUNCTIONS } from './stat';
import { TEXT_FUNCTIONS } from './text';

export const FUNCTIONS: Map<string, FunctionSpec> = new Map();

const GROUPS = [
  MATH_FUNCTIONS,
  STAT_FUNCTIONS,
  TEXT_FUNCTIONS,
  DATE_FUNCTIONS,
  LOOKUP_FUNCTIONS,
  INFO_FUNCTIONS,
  FINANCIAL_FUNCTIONS,
  EXTRA_FUNCTIONS,
];

function isSound(value: Scalar | undefined): boolean {
  return value !== undefined && (typeof value !== 'number' || isFinite(value));
}

/** A result the workbook can hold: a number is finite (anything else is #NUM!) and an array has no holes. */
function sound(value: Value): Value {
  if (typeof value === 'number') return isFinite(value) ? value : ERROR_NUM;
  if (value === undefined) return ERROR_VALUE;
  if (!isArray(value)) return value;
  const data = value.data;
  const size = value.rows * value.cols;
  let clean = data.length === size && size >= 1;
  for (let i = 0; clean && i < size; i++) clean = isSound(data[i]);
  if (clean) return value;
  if (!(size >= 1) || data.length < size) return ERROR_VALUE;
  const fixed: Scalar[] = new Array(size);
  for (let i = 0; i < size; i++) {
    const item = data[i];
    fixed[i] = item === undefined ? null : typeof item === 'number' && !isFinite(item) ? ERROR_NUM : item;
  }
  return { rows: value.rows, cols: value.cols, data: fixed };
}

/**
 * The one door every call goes through. The evaluator checks argument counts
 * before it calls; a function reached with the wrong number anyway answers
 * #VALUE! instead of reading arguments that are not there. Results are checked
 * so that no NaN or Infinity (from a cell that should never have held one)
 * leaves the library.
 */
function guarded(spec: FunctionSpec): FunctionSpec {
  const call = spec.call;
  const min = spec.min;
  const max = spec.max;
  const checked: FunctionSpec = {
    call: (args, context) =>
      args.length < min || (max !== undefined && args.length > max) ? ERROR_VALUE : sound(call(args, context)),
    min,
  };
  if (max !== undefined) checked.max = max;
  if (spec.arrayArgs !== undefined) checked.arrayArgs = spec.arrayArgs;
  if (spec.scalarArgs !== undefined) checked.scalarArgs = spec.scalarArgs;
  return checked;
}

for (let i = 0; i < GROUPS.length; i++) {
  const group = GROUPS[i];
  for (const name in group) {
    if (Object.prototype.hasOwnProperty.call(group, name)) FUNCTIONS.set(name, guarded(group[name]));
  }
}
