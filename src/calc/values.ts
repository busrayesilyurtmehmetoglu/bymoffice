// Values of the formula engine and the rules Excel uses to turn one kind into
// another. Everything that calculates (the evaluator, every function) goes
// through these helpers, so the rules live in one place.

import { parseCellInput } from '../edit/input';
import { serialToDate } from '../format/number-format';
import type { CellRange } from '../model/sheet/types';

export type ErrorCode = '#NULL!' | '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#NUM!' | '#N/A' | '#SPILL!' | '#CALC!';

/** An error value. Errors are values: they travel through formulas until something handles them. */
export interface CalcError {
  readonly error: ErrorCode;
}

/** What one cell can hold: a number, text, a boolean, nothing (null), or an error. */
export type Scalar = number | string | boolean | null | CalcError;

/** A rectangle of scalars: an array constant, the result of an array function, or the cells of a range. */
export interface ArrayValue {
  readonly rows: number;
  readonly cols: number;
  /** Row by row; length is rows * cols. */
  readonly data: readonly Scalar[];
}

/**
 * A reference to cells. It is a value of its own because some functions need
 * to know where the cells are (ROW, OFFSET) or want to skip what they do not
 * need (SUM over a whole column reads only the cells that exist).
 */
export interface RefValue {
  readonly ref: true;
  readonly sheet: number;
  readonly range: CellRange;
}

export type Value = Scalar | ArrayValue | RefValue;

const ERRORS: Record<ErrorCode, CalcError> = {
  '#NULL!': { error: '#NULL!' },
  '#DIV/0!': { error: '#DIV/0!' },
  '#VALUE!': { error: '#VALUE!' },
  '#REF!': { error: '#REF!' },
  '#NAME?': { error: '#NAME?' },
  '#NUM!': { error: '#NUM!' },
  '#N/A': { error: '#N/A' },
  '#SPILL!': { error: '#SPILL!' },
  '#CALC!': { error: '#CALC!' },
};

/** The shared error value for a code; errors are compared by code, never by identity. */
export function calcError(code: ErrorCode): CalcError {
  return ERRORS[code];
}

export const ERROR_VALUE = ERRORS['#VALUE!'];
export const ERROR_DIV0 = ERRORS['#DIV/0!'];
export const ERROR_REF = ERRORS['#REF!'];
export const ERROR_NAME = ERRORS['#NAME?'];
export const ERROR_NUM = ERRORS['#NUM!'];
export const ERROR_NA = ERRORS['#N/A'];

export function isError(value: unknown): value is CalcError {
  return typeof value === 'object' && value !== null && typeof (value as CalcError).error === 'string';
}

export function isRef(value: unknown): value is RefValue {
  return typeof value === 'object' && value !== null && (value as RefValue).ref === true;
}

export function isArray(value: unknown): value is ArrayValue {
  return typeof value === 'object' && value !== null && Array.isArray((value as ArrayValue).data);
}

/** The error code text of an error literal such as "#N/A", or undefined. */
export function parseErrorCode(text: string): ErrorCode | undefined {
  const upper = text.toUpperCase();
  return Object.prototype.hasOwnProperty.call(ERRORS, upper) ? (upper as ErrorCode) : undefined;
}

// ---------------------------------------------------------------------------
// Text as numbers
//
// One rule for operators and for function arguments: plain numbers (with an optional
// percent sign), then what the cell editor would read as a number: thousands separators,
// currency, dates and times (en-US conventions). Everything is scanned by hand and only
// short text is ever looked at closely, so no text can make this slow.

function isSpaceCode(c: number): boolean {
  if (c === 32 || (c >= 9 && c <= 13) || c === 160) return true;
  if (c < 0x1680) return false;
  return c === 0x1680 || (c >= 0x2000 && c <= 0x200a) || c === 0x2028 || c === 0x2029 || c === 0x202f || c === 0x205f || c === 0x3000 || c === 0xfeff;
}

/**
 * A plain number as text: digits with an optional fraction and exponent, an optional sign,
 * an optional trailing percent sign, spaces around it. Undefined when the text is anything else.
 */
export function textToNumber(text: string): number | undefined {
  const n = text.length;
  let i = 0;
  while (i < n && isSpaceCode(text.charCodeAt(i))) i++;
  const start = i;
  let c = text.charCodeAt(i);
  if (c === 43 || c === 45) i++;
  let digits = 0;
  while (i < n && (c = text.charCodeAt(i)) >= 48 && c <= 57) {
    i++;
    digits++;
  }
  if (i < n && text.charCodeAt(i) === 46) {
    i++;
    while (i < n && (c = text.charCodeAt(i)) >= 48 && c <= 57) {
      i++;
      digits++;
    }
  }
  if (digits === 0) return undefined;
  c = text.charCodeAt(i);
  if (c === 69 || c === 101) {
    i++;
    c = text.charCodeAt(i);
    if (c === 43 || c === 45) i++;
    const exponent = i;
    while (i < n && (c = text.charCodeAt(i)) >= 48 && c <= 57) i++;
    if (i === exponent) return undefined;
  }
  const end = i;
  while (i < n && isSpaceCode(text.charCodeAt(i))) i++;
  const percent = i < n && text.charCodeAt(i) === 37;
  if (percent) i++;
  while (i < n && isSpaceCode(text.charCodeAt(i))) i++;
  if (i < n) return undefined;
  const value = Number(text.substring(start, end));
  if (!isFinite(value)) return undefined;
  return percent ? value / 100 : value;
}

/** Longest text that is looked at as a date, a time or a formatted number. */
const MAX_NUMBER_TEXT = 64;

const DATE_FORMAT_MIN = 14;
const DATE_FORMAT_MAX = 22;
const FORMAT_ELAPSED = 46;
const FORMAT_SUBSECONDS = 47;

function inputOptions(context: CalcContext | undefined): { date1904: boolean; today?: Date } {
  if (!context) return { date1904: false };
  // Typed dates without a year take the year of the calculation's clock.
  const year = serialToDate(context.now(), context.date1904).year;
  if (year !== year) return { date1904: context.date1904 === true };
  // Noon on 1 July is in the same year in every time zone.
  return { date1904: context.date1904 === true, today: new Date(Date.UTC(year, 6, 1, 12)) };
}

/** Text as a date or a time: its serial number; undefined when it is neither. */
export function parseDateText(text: string, context?: CalcContext): number | undefined {
  if (text.length > MAX_NUMBER_TEXT) return undefined;
  const parsed = parseCellInput(text, inputOptions(context));
  if (parsed.kind !== 'number' || typeof parsed.value !== 'number') return undefined;
  const id = parsed.numFmtId;
  if (id === undefined) return undefined;
  if ((id >= DATE_FORMAT_MIN && id <= DATE_FORMAT_MAX) || id === FORMAT_ELAPSED || id === FORMAT_SUBSECONDS) {
    return parsed.value;
  }
  return undefined;
}

/**
 * Text as a number the way Excel reads it wherever a number is needed: plain numbers,
 * percentages, thousands separators, currency, dates and times (en-US conventions).
 * The context gives the date system and the year of dates written without one.
 */
export function parseNumberText(text: string, context?: CalcContext): number | undefined {
  const plain = textToNumber(text);
  if (plain !== undefined) return plain;
  if (text.length > MAX_NUMBER_TEXT || text.length === 0) return undefined;
  const c = text.charCodeAt(0);
  // Formulas and forced text are meanings of the cell editor, not of a value.
  if (c === 61 || c === 39) return undefined;
  const parsed = parseCellInput(text, inputOptions(context));
  return parsed.kind === 'number' && typeof parsed.value === 'number' && isFinite(parsed.value) ? parsed.value : undefined;
}

/**
 * A scalar as a number, as arithmetic operators and numeric arguments need it: empty is 0,
 * TRUE is 1, text is read by parseNumberText() and is #VALUE! when that fails; errors pass
 * through.
 */
export function toNumber(value: Scalar, context?: CalcContext): number | CalcError {
  if (typeof value === 'number') return value;
  if (value === null) return 0;
  if (typeof value === 'boolean') return value ? 1 : 0;
  if (typeof value === 'string') {
    const number = parseNumberText(value, context);
    return number === undefined ? ERROR_VALUE : number;
  }
  return value;
}

// ---------------------------------------------------------------------------
// Numbers as text

/** Longest decimal notation a small number is written in before scientific notation takes over. */
const MAX_DECIMAL_TEXT = 20;

/**
 * A number as text where a formula needs text ("&", LEN, LEFT, CONCATENATE ...): all of its
 * digits, up to the 15 Excel works with, not the 11 characters a cell shows. From 1E+15 up
 * and below 1E-9 it is written in scientific notation.
 */
export function numberToText(value: number): string {
  if (value === 0) return '0';
  if (!isFinite(value)) return '#NUM!';
  // "d.dddddddddddddde+x": 15 significant digits.
  const text = Math.abs(value).toExponential(14);
  const mark = text.indexOf('e');
  const exponent = Number(text.substring(mark + 1));
  let digits = text.charAt(0) + text.substring(2, mark);
  let end = digits.length;
  while (end > 1 && digits.charCodeAt(end - 1) === 48) end--;
  digits = digits.substring(0, end);
  const sign = value < 0 ? '-' : '';
  if (exponent >= 0 && exponent < 15) {
    const whole = exponent + 1;
    if (digits.length <= whole) {
      let out = digits;
      for (let k = digits.length; k < whole; k++) out += '0';
      return sign + out;
    }
    return sign + digits.substring(0, whole) + '.' + digits.substring(whole);
  }
  if (exponent < 0 && exponent >= -9) {
    let out = '0.';
    for (let k = 1; k < -exponent; k++) out += '0';
    out += digits;
    if (out.length <= MAX_DECIMAL_TEXT) return sign + out;
  }
  const power = String(Math.abs(exponent));
  return sign + digits.charAt(0) + (digits.length > 1 ? '.' + digits.substring(1) : '') + 'E' + (exponent < 0 ? '-' : '+') + (power.length < 2 ? '0' + power : power);
}

/** A scalar as text, as "&" needs it: numbers by numberToText(), booleans as TRUE and FALSE, empty as "". */
export function toText(value: Scalar): string | CalcError {
  if (typeof value === 'string') return value;
  if (value === null) return '';
  if (typeof value === 'number') return numberToText(value);
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  return value;
}

// ---------------------------------------------------------------------------
// The text one formula may build
//
// The number of cells of an array is limited, but each of them can hold 32,767 characters.
// Whoever fills an array with text it builds (an operator, a function applied to every
// element) reports the characters here; past the limit the result is #VALUE!.

/** Characters of text the arrays of one formula may hold in all. */
export const MAX_FORMULA_TEXT = 50000000;

let textBuilt = 0;

/** Starts counting for a formula; returns the count of the formula that was being calculated, for endTextBudget(). */
export function beginTextBudget(): number {
  const saved = textBuilt;
  textBuilt = 0;
  return saved;
}

export function endTextBudget(saved: number): void {
  textBuilt = saved;
}

/** Counts a value put into an array; false when the formula has built more text than it may. */
export function chargeText(value: Scalar): boolean {
  if (typeof value !== 'string') return true;
  textBuilt += value.length;
  return textBuilt <= MAX_FORMULA_TEXT;
}

/** A scalar as a boolean, as IF needs it: 0 and empty are FALSE, "TRUE"/"FALSE" text is read, other text is #VALUE!. */
export function toBoolean(value: Scalar): boolean | CalcError {
  if (typeof value === 'boolean') return value;
  if (value === null) return false;
  if (typeof value === 'number') return value !== 0;
  if (typeof value === 'string') {
    const upper = value.toUpperCase();
    if (upper === 'TRUE') return true;
    if (upper === 'FALSE') return false;
    return ERROR_VALUE;
  }
  return value;
}

/** Excel's result for a number that JavaScript cannot hold: #NUM! instead of Infinity or NaN. */
export function checkNumber(value: number): number | CalcError {
  return isFinite(value) ? value : ERROR_NUM;
}

// Excel orders values of different kinds: numbers before text before booleans.
// An empty cell compares as 0, "" or FALSE, whatever the other side is.
function rank(value: number | string | boolean): number {
  return typeof value === 'number' ? 0 : typeof value === 'string' ? 1 : 2;
}

/**
 * Compares two scalars the way "=", "<" and ">" do: negative, zero or positive.
 * Text is compared without regard to case. An error on either side is returned
 * (the left one first).
 */
export function compareScalars(a: Scalar, b: Scalar): number | CalcError {
  if (isError(a)) return a;
  if (isError(b)) return b;
  let left = a;
  let right = b;
  if (left === null) left = right === null ? 0 : typeof right === 'number' ? 0 : typeof right === 'string' ? '' : false;
  if (right === null) right = typeof left === 'number' ? 0 : typeof left === 'string' ? '' : false;
  const leftRank = rank(left);
  const rightRank = rank(right);
  if (leftRank !== rightRank) return leftRank - rightRank;
  if (typeof left === 'number') return left === right ? 0 : left < (right as number) ? -1 : 1;
  if (typeof left === 'boolean') return left === right ? 0 : left ? 1 : -1;
  const x = (left as string).toUpperCase();
  const y = (right as string).toUpperCase();
  return x === y ? 0 : x < y ? -1 : 1;
}

/** A scalar as an ArrayValue of one cell. */
export function scalarArray(value: Scalar): ArrayValue {
  return { rows: 1, cols: 1, data: [value] };
}

/** The value at a row and column of an array; a 1 x 1 array answers for every position, as Excel broadcasts it. */
export function arrayAt(array: ArrayValue, row: number, col: number): Scalar {
  if (array.rows === 1 && array.cols === 1) return array.data[0];
  const r = array.rows === 1 ? 0 : row;
  const c = array.cols === 1 ? 0 : col;
  if (r >= array.rows || c >= array.cols) return ERROR_NA;
  return array.data[r * array.cols + c];
}

/** What the evaluator gives every function to reach the workbook. */
export interface CalcContext {
  /** The value of a cell; null when it is empty. Calculates the cell first when it holds a formula. */
  cellValue(sheet: number, row: number, col: number): Scalar;
  /**
   * Calls `visit` for every cell of a range that holds something, in row order.
   * Whole rows and columns cost only what is in them.
   */
  forEachCell(sheet: number, range: CellRange, visit: (value: Scalar, row: number, col: number) => void): void;
  /** One past the last row and column in use on a sheet: the limit a range is cut to when it is read as an array. */
  usedExtent(sheet: number): { rows: number; cols: number };
  /** Index of a sheet by name (not case sensitive), or undefined. */
  sheetIndex(name: string): number | undefined;
  sheetName(index: number): string | undefined;
  sheetCount(): number;
  /** The cell being calculated. */
  readonly sheet: number;
  readonly row: number;
  readonly col: number;
  readonly date1904: boolean;
  /** The current date and time as a serial number, the same for a whole recalculation. */
  now(): number;
  /** A random number in [0, 1). */
  random(): number;
  /** The formula text of a cell without the leading "=", or undefined when the cell holds no formula. */
  cellFormula?(sheet: number, row: number, col: number): string | undefined;
  /**
   * Something worked out from cells that other formulas of the same recalculation can use
   * as well (the index of a lookup column). `build` is called when nothing is kept under
   * `key`; what is kept lives until the recalculation ends.
   */
  memo?<T>(key: string, build: () => T): T;
  /**
   * Counts work (cells read, elements calculated) against the budget of the recalculation.
   * Throws to give the calculation up when the budget is used up; functions let that pass.
   */
  tick?(units: number): void;
}

/** Largest array a reference is read into. */
const MAX_REF_CELLS = 5000000;

/**
 * The cells of a reference as an array, cut to the used part of its sheet. Never throws:
 * a reference that is not a rectangle gives a single #REF!, one too large to hold a single #NUM!.
 */
export function refToArray(ref: RefValue, context: CalcContext): ArrayValue {
  const range = ref.range;
  if (!range || !(range.top >= 0) || !(range.left >= 0) || !(range.bottom >= range.top) || !(range.right >= range.left)) return scalarArray(ERROR_REF);
  const extent = context.usedExtent(ref.sheet);
  const bottom = Math.min(range.bottom, Math.max(extent.rows - 1, range.top));
  const right = Math.min(range.right, Math.max(extent.cols - 1, range.left));
  const rows = bottom - range.top + 1;
  const cols = right - range.left + 1;
  if (!(rows >= 1 && cols >= 1) || Math.floor(rows) !== rows || Math.floor(cols) !== cols) return scalarArray(ERROR_REF);
  if (rows * cols > MAX_REF_CELLS) return scalarArray(ERROR_NUM);
  const data: Scalar[] = new Array(rows * cols).fill(null);
  context.forEachCell(ref.sheet, { top: range.top, left: range.left, bottom, right }, (value, row, col) => {
    const r = row - range.top;
    const c = col - range.left;
    if (r >= 0 && r < rows && c >= 0 && c < cols) data[r * cols + c] = value;
  });
  return { rows, cols, data };
}

/** A function of the library. Arguments arrive evaluated; references stay references. */
export type CalcFunction = (args: Value[], context: CalcContext) => Value;

/** Which arguments of a function something holds for: all of them, or the ones the test accepts (by position, from 0). */
export type ArgSelector = true | ((index: number) => boolean);

export interface FunctionSpec {
  call: CalcFunction;
  /** Fewest and most arguments; `max` undefined means any number. */
  min: number;
  max?: number;
  /**
   * Arguments that take an array: what is written there is calculated element by element
   * even in an ordinary formula (the arguments of SUMPRODUCT, the lookup array of MATCH).
   * Everywhere else a range that meets an operator stands for the one cell in the formula's
   * row or column (implicit intersection), unless the formula is an array formula.
   */
  arrayArgs?: ArgSelector;
  /**
   * Arguments that take a single value. In an ordinary formula a range given there is
   * reduced by implicit intersection before the function is called.
   */
  scalarArgs?: ArgSelector;
}
