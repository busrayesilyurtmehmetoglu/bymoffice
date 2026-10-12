// Further worksheet functions, gathered from the files beside this one, plus a
// few that belong to no family: HYPERLINK, ADDRESS, the "precise" rounding to
// multiples, ROMAN, SEQUENCE, RANDARRAY and MMULT.
//
// Nothing here throws, whatever the arguments are.

import { MAX_COLS, MAX_ROWS } from '../../model/sheet/types';
import {
  checkNumber,
  ERROR_NUM,
  ERROR_VALUE,
  isError,
  toBoolean,
  toText,
  type CalcContext,
  type CalcError,
  type CalcFunction,
  type FunctionSpec,
  type Scalar,
  type Value,
} from '../values';
import { EXTRA_DATABASE_FUNCTIONS } from './extra-database';
import { EXTRA_DATE_FUNCTIONS } from './extra-date';
import { EXTRA_ENGINEERING_FUNCTIONS } from './extra-engineering';
import { EXTRA_FINANCIAL_FUNCTIONS } from './extra-financial';
import { EXTRA_MATH_FUNCTIONS } from './extra-math';
import { EXTRA_STAT_FUNCTIONS } from './extra-stat';
import {
  clean,
  ERROR_CALC,
  first,
  mapScalars,
  MAX_ARRAY_CELLS,
  MAX_TEXT_LENGTH,
  numeric,
  scalarNumber,
  spec,
  tidy,
  toArray,
} from './helpers';

/** Multiplications MMULT may make before it gives up with #NUM!. */
const MAX_PRODUCTS = 1 << 28;

// ---------------------------------------------------------------------------
// HYPERLINK and ADDRESS

/** The text a hyperlink cell shows: the friendly name when there is one, else the link itself. */
const hyperlink = mapScalars((args) => {
  const shown = args.length > 1 ? args[1] : args[0];
  return typeof shown === 'number' ? checkNumber(shown) : shown;
});

function columnLetters(col: number): string {
  let text = '';
  for (let n = col; n > 0; n = Math.floor((n - 1) / 26)) text = String.fromCharCode(65 + ((n - 1) % 26)) + text;
  return text;
}

function isDigit(c: number): boolean {
  return c >= 48 && c <= 57;
}

function isLetter(c: number): boolean {
  if ((c >= 65 && c <= 90) || (c >= 97 && c <= 122)) return true;
  // Past ASCII everything counts as a letter except signs and the punctuation blocks.
  return c >= 0xc0 && c !== 0xd7 && c !== 0xf7 && !(c >= 0x2000 && c <= 0x2bff) && !(c >= 0x3000 && c <= 0x303f);
}

/** True for names that could be taken for a cell: "A1", "XFD10", "R1C1", "R", "C", "RC". */
function looksLikeCell(name: string): boolean {
  const upper = name.toUpperCase();
  let i = 0;
  while (i < upper.length && i < 4 && upper.charCodeAt(i) >= 65 && upper.charCodeAt(i) <= 90) i++;
  if (i >= 1 && i <= 3 && i < upper.length) {
    let k = i;
    while (k < upper.length && isDigit(upper.charCodeAt(k))) k++;
    if (k === upper.length) return true;
  }
  // R, C, RC and the same with numbers after the letters.
  let at = 0;
  if (upper.charAt(at) === 'R') {
    at++;
    while (at < upper.length && isDigit(upper.charCodeAt(at))) at++;
  }
  if (upper.charAt(at) === 'C') {
    at++;
    while (at < upper.length && isDigit(upper.charCodeAt(at))) at++;
  }
  return at > 0 && at === upper.length;
}

/** A sheet name as it stands before "!": in apostrophes unless it is a plain word. */
function sheetPrefix(name: string): string {
  let simple = name.length > 0 && !isDigit(name.charCodeAt(0)) && name.charAt(0) !== '.';
  for (let i = 0; simple && i < name.length; i++) {
    const c = name.charCodeAt(i);
    // letters, digits, "_" and "."
    simple = isLetter(c) || isDigit(c) || c === 95 || c === 46;
  }
  if (name.length === 0 || (simple && !looksLikeCell(name))) return name + '!';
  return "'" + name.split("'").join("''") + "'!";
}

const address = mapScalars((args, context) => {
  const rowArg = scalarNumber(args[0], context);
  if (typeof rowArg !== 'number') return rowArg;
  const colArg = scalarNumber(args[1], context);
  if (typeof colArg !== 'number') return colArg;
  let kind = 1;
  if (args.length > 2 && args[2] !== null) {
    const given = scalarNumber(args[2], context);
    if (typeof given !== 'number') return given;
    kind = Math.trunc(given);
  }
  let a1: boolean | CalcError = true;
  if (args.length > 3 && args[3] !== null) a1 = toBoolean(args[3]);
  if (typeof a1 !== 'boolean') return a1;
  let prefix = '';
  if (args.length > 4 && args[4] !== null) {
    const sheet = toText(args[4]);
    if (isError(sheet)) return sheet;
    if (sheet.length > MAX_TEXT_LENGTH) return ERROR_VALUE;
    prefix = sheetPrefix(sheet);
  }
  const row = Math.trunc(rowArg);
  const col = Math.trunc(colArg);
  if (!(row >= 1 && row <= MAX_ROWS && col >= 1 && col <= MAX_COLS) || !(kind >= 1 && kind <= 4)) return ERROR_VALUE;
  // 1: both fixed, 2: row fixed, 3: column fixed, 4: neither.
  const rowFixed = kind === 1 || kind === 2;
  const colFixed = kind === 1 || kind === 3;
  const cell = a1
    ? (colFixed ? '$' : '') + columnLetters(col) + (rowFixed ? '$' : '') + row
    : 'R' + (rowFixed ? String(row) : '[' + row + ']') + 'C' + (colFixed ? String(col) : '[' + col + ']');
  const text = prefix + cell;
  return text.length > MAX_TEXT_LENGTH ? ERROR_VALUE : text;
});

// ---------------------------------------------------------------------------
// Rounding to a multiple, whatever the signs

/** Rounds x to a multiple of `step` by `round` (Math.ceil or Math.floor), free of floating point noise. */
function toMultiple(x: number, step: number, round: (value: number) => number): number {
  return tidy(round(tidy(x / step)) * step);
}

function precise(round: (value: number) => number): CalcFunction {
  return numeric((x, significance) => {
    const step = Math.abs(significance === undefined ? 1 : significance);
    if (x === 0 || step === 0) return 0;
    return toMultiple(x, step, round);
  });
}

// ---------------------------------------------------------------------------
// ROMAN

const ROMAN_CHARS = 'MDCLXVI';
const ROMAN_VALUES = [1000, 500, 100, 50, 10, 5, 1];

/**
 * A number from 0 to 3999 in Roman numerals. Form 0 is the classic one; each
 * further form lets a smaller numeral be subtracted from a larger one (499 is
 * CDXCIX, LDVLIV, XDIX, VDIV and at last ID).
 */
function romanText(value: number, form: number): string {
  let text = '';
  let rest = value;
  for (let i = 0; i <= 3; i++) {
    let index = 2 * i;
    const digit = Math.floor(rest / ROMAN_VALUES[index]);
    if (digit % 5 === 4) {
      const larger = digit === 4 ? index - 1 : index - 2;
      // The subtracted numeral may shrink while the number still reaches the difference.
      for (let steps = 0; steps < form && index < ROMAN_VALUES.length - 1; steps++) {
        if (ROMAN_VALUES[larger] - ROMAN_VALUES[index + 1] <= rest) index++;
        else break;
      }
      text += ROMAN_CHARS.charAt(index) + ROMAN_CHARS.charAt(larger);
      rest = rest + ROMAN_VALUES[index] - ROMAN_VALUES[larger];
    } else {
      if (digit > 4) text += ROMAN_CHARS.charAt(index - 1);
      for (let k = digit % 5; k > 0; k--) text += ROMAN_CHARS.charAt(index);
      rest %= ROMAN_VALUES[index];
    }
  }
  return text;
}

const roman = mapScalars((args, context) => {
  const given = scalarNumber(args[0], context);
  if (typeof given !== 'number') return given;
  let form = 0;
  if (args.length > 1 && args[1] !== null) {
    // TRUE is the classic form, FALSE the shortest one.
    if (typeof args[1] === 'boolean') form = args[1] ? 0 : 4;
    else {
      const style = scalarNumber(args[1], context);
      if (typeof style !== 'number') return style;
      form = Math.trunc(style);
    }
  }
  const value = Math.trunc(given);
  if (!(value >= 0 && value <= 3999) || !(form >= 0 && form <= 4)) return ERROR_VALUE;
  return romanText(value, form);
});

// ---------------------------------------------------------------------------
// SEQUENCE and RANDARRAY

/** An optional number: the fallback when it is left out or empty. */
function optional(args: Value[], index: number, fallback: number, context: CalcContext): number | CalcError {
  if (index >= args.length) return fallback;
  const value = first(args[index], context);
  return value === null ? fallback : scalarNumber(value, context);
}

/** Rows and columns of an array to build, or the error that says why not. */
function shape(args: Value[], context: CalcContext): [number, number] | CalcError {
  // A size left empty, as in SEQUENCE(,2), is 1.
  const rowsArg = optional(args, 0, 1, context);
  if (typeof rowsArg !== 'number') return rowsArg;
  const colsArg = optional(args, 1, 1, context);
  if (typeof colsArg !== 'number') return colsArg;
  const rows = Math.trunc(rowsArg);
  const cols = Math.trunc(colsArg);
  if (!(rows >= 0 && cols >= 0)) return ERROR_VALUE;
  if (rows === 0 || cols === 0) return ERROR_CALC;
  if (!(rows * cols <= MAX_ARRAY_CELLS)) return ERROR_NUM;
  return [rows, cols];
}

const sequence: CalcFunction = (args, context) => {
  const size = shape(args, context);
  if (isError(size)) return size;
  const start = optional(args, 2, 1, context);
  if (typeof start !== 'number') return start;
  const step = optional(args, 3, 1, context);
  if (typeof step !== 'number') return step;
  const count = size[0] * size[1];
  if (!isFinite(start + (count - 1) * step)) return ERROR_NUM;
  const data: Scalar[] = new Array(count);
  for (let i = 0; i < count; i++) data[i] = start + i * step;
  return { rows: size[0], cols: size[1], data };
};

const randomArray: CalcFunction = (args, context) => {
  const size = shape(args, context);
  if (isError(size)) return size;
  const low = optional(args, 2, 0, context);
  if (typeof low !== 'number') return low;
  const high = optional(args, 3, 1, context);
  if (typeof high !== 'number') return high;
  let whole: boolean | CalcError = false;
  if (args.length > 4) whole = toBoolean(first(args[4], context));
  if (typeof whole !== 'boolean') return whole;
  if (!(low <= high) || !isFinite(high - low)) return ERROR_VALUE;
  if (whole && (Math.floor(low) !== low || Math.floor(high) !== high)) return ERROR_VALUE;
  const count = size[0] * size[1];
  const data: Scalar[] = new Array(count);
  const span = whole ? high - low + 1 : high - low;
  for (let i = 0; i < count; i++) {
    let random = context.random();
    if (!(random >= 0 && random < 1)) random = 0;
    let value = whole ? low + Math.floor(random * span) : low + random * span;
    if (value > high) value = high;
    data[i] = value;
  }
  return { rows: size[0], cols: size[1], data };
};

/** SEQUENCE and RANDARRAY read single values: a range given to them stands for one cell. */
function ofScalars(call: CalcFunction, min: number, max: number): FunctionSpec {
  return { call, min, max, scalarArgs: true };
}

// ---------------------------------------------------------------------------
// MMULT

const matrixProduct: CalcFunction = (args, context) => {
  const a = toArray(args[0], context);
  if (isError(a)) return a;
  const b = toArray(args[1], context);
  if (isError(b)) return b;
  if (a.cols !== b.rows) return ERROR_VALUE;
  const rows = a.rows;
  const cols = b.cols;
  const inner = a.cols;
  if (!(rows * cols <= MAX_ARRAY_CELLS) || !(rows * cols * inner <= MAX_PRODUCTS)) return ERROR_NUM;
  // Every cell of both matrices must be a number; an empty cell or text spoils the product.
  const left: number[] = new Array(rows * inner);
  const right: number[] = new Array(inner * cols);
  for (let i = 0; i < left.length; i++) {
    const value = clean(a.data[i]);
    if (isError(value)) return value;
    if (typeof value !== 'number') return ERROR_VALUE;
    left[i] = value;
  }
  for (let i = 0; i < right.length; i++) {
    const value = clean(b.data[i]);
    if (isError(value)) return value;
    if (typeof value !== 'number') return ERROR_VALUE;
    right[i] = value;
  }
  const data: Scalar[] = new Array(rows * cols);
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      let total = 0;
      for (let k = 0; k < inner; k++) total += left[r * inner + k] * right[k * cols + c];
      data[r * cols + c] = isFinite(total) ? total : ERROR_NUM;
    }
  }
  return { rows, cols, data };
};

// ---------------------------------------------------------------------------

const ceilingPrecise = spec(precise(Math.ceil), 1, 2);

const OWN_FUNCTIONS: { [name: string]: FunctionSpec } = {
  HYPERLINK: spec(hyperlink, 1, 2),
  ADDRESS: spec(address, 2, 5),
  'CEILING.PRECISE': ceilingPrecise,
  'ISO.CEILING': ceilingPrecise,
  'FLOOR.PRECISE': spec(precise(Math.floor), 1, 2),
  ROMAN: spec(roman, 1, 2),
  SEQUENCE: ofScalars(sequence, 1, 4),
  RANDARRAY: ofScalars(randomArray, 0, 5),
  MMULT: spec(matrixProduct, 2, 2, true),
};

/** Every function of this file and of the extra-*.ts files beside it, by its upper-case name. */
export const EXTRA_FUNCTIONS: { [name: string]: FunctionSpec } = {};

const GROUPS = [
  OWN_FUNCTIONS,
  EXTRA_DATE_FUNCTIONS,
  EXTRA_FINANCIAL_FUNCTIONS,
  EXTRA_ENGINEERING_FUNCTIONS,
  EXTRA_STAT_FUNCTIONS,
  EXTRA_DATABASE_FUNCTIONS,
  EXTRA_MATH_FUNCTIONS,
];

for (let i = 0; i < GROUPS.length; i++) {
  const group = GROUPS[i];
  for (const name in group) {
    if (Object.prototype.hasOwnProperty.call(group, name)) EXTRA_FUNCTIONS[name] = group[name];
  }
}
