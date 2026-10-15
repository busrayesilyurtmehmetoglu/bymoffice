// Values of the formula engine and the rules Excel uses to turn one kind into
// another. Everything that calculates (the evaluator, every function) goes
// through these helpers, so the rules live in one place.

import { parseCellInput } from '../edit/input';
import type { FormatLocale } from '../format/locale';
import { serialToDate } from '../format/number-format';
import { MAX_COLS, MAX_ROWS, type CellRange } from '../model/sheet/types';

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
  /**
   * A hint: from this index on, every element of `data` is the same as `data[tail]` (the empty cells
   * below what a sheet holds, in a whole column read as an array). The elements are there all the
   * same; what works element by element may work out the rest from one of them.
   */
  readonly tail?: number;
}

/** Arrays of fewer elements are not worth a tail. */
const MIN_TAIL = 1024;

/** The index from which an array's elements are all the same, when `uniform` says from where they are null; undefined when not worth it. */
export function tailFrom(size: number, filled: number): number | undefined {
  return size - filled >= MIN_TAIL ? filled : undefined;
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
//
// Text that is a number or a date the way files write them has that value in every locale:
// the engine itself writes numbers with a point ("&", CONCAT, TEXT), and what it writes it
// must read back. Only text that is neither is then read the way the locale of the context
// writes numbers, where that is with another decimal sign ("7,5" and "1.234,5" in Turkish).
// That second reading knows the decimal sign and the thousands separator, no dates or times.
// UNSURE: a Turkish Excel reads "1.234" as 1234 and "7.5" as a date; here both stay what a file means.

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

function isDateFormatId(id: number | undefined): boolean {
  return id !== undefined && ((id >= DATE_FORMAT_MIN && id <= DATE_FORMAT_MAX) || id === FORMAT_ELAPSED || id === FORMAT_SUBSECONDS);
}

/** Text as a date or a time: its serial number; undefined when it is neither. */
export function parseDateText(text: string, context?: CalcContext): number | undefined {
  if (text.length > MAX_NUMBER_TEXT) return undefined;
  const parsed = parseCellInput(text, inputOptions(context));
  if (parsed.kind !== 'number' || typeof parsed.value !== 'number') return undefined;
  return isDateFormatId(parsed.numFmtId) ? parsed.value : undefined;
}

/**
 * A whole number and a fraction, as "1 1/2" or "-0 3/4": its value; undefined when the text
 * is anything else. A fraction alone ("1/2") is not read here: it is a date.
 */
function fractionText(text: string): number | undefined {
  const n = text.length;
  let i = 0;
  while (i < n && isSpaceCode(text.charCodeAt(i))) i++;
  let c = text.charCodeAt(i);
  const negative = c === 45;
  if (c === 43 || c === 45) i++;
  const parts = [0, 0, 0];
  for (let part = 0; part < 3; part++) {
    const start = i;
    while (i < n && (c = text.charCodeAt(i)) >= 48 && c <= 57) {
      parts[part] = parts[part] * 10 + (c - 48);
      i++;
    }
    if (i === start) return undefined;
    if (part === 0) {
      // Spaces between the whole number and the fraction.
      const gap = i;
      while (i < n && text.charCodeAt(i) === 32) i++;
      if (i === gap) return undefined;
    } else if (part === 1) {
      if (text.charCodeAt(i) !== 47) return undefined;
      i++;
    }
  }
  while (i < n && isSpaceCode(text.charCodeAt(i))) i++;
  if (i < n || parts[2] === 0) return undefined;
  const value = parts[0] + parts[1] / parts[2];
  return isFinite(value) ? (negative ? -value : value) : undefined;
}

/** The decimal sign and the thousands separator of a locale that writes numbers another way than files do. */
export interface NumberSeparators {
  readonly decimal: string;
  /** Empty when the locale has none that can be told from the decimal sign. */
  readonly group: string;
}

function isDigitCode(c: number): boolean {
  return c >= 48 && c <= 57;
}

/** Whether a character can separate digits: one character that is no part of a number written the plain way. */
function isSeparator(text: unknown): text is string {
  if (typeof text !== 'string' || text.length !== 1) return false;
  const c = text.charCodeAt(0);
  // Digits, signs, the percent sign and the "E" of an exponent mean something else.
  return !((c >= 48 && c <= 57) || c === 43 || c === 45 || c === 37 || c === 69 || c === 101);
}

const separatorsOf = new WeakMap<FormatLocale, NumberSeparators | null>();

/**
 * The separators of the locale of a context, when text that is no number the way files write
 * one is to be read with them: the locale writes its decimal sign as something other than a point. Undefined
 * without a locale and for a locale that writes numbers the way files do (English).
 */
export function numberSeparators(context: CalcContext | undefined): NumberSeparators | undefined {
  const locale = context === undefined ? undefined : context.locale;
  if (locale === undefined || locale === null || typeof locale !== 'object') return undefined;
  let found = separatorsOf.get(locale);
  if (found === undefined) {
    found = null;
    if (isSeparator(locale.decimal) && locale.decimal !== '.') {
      found = { decimal: locale.decimal, group: isSeparator(locale.group) && locale.group !== locale.decimal ? locale.group : '' };
    }
    separatorsOf.set(locale, found);
  }
  return found === null ? undefined : found;
}

/**
 * Text holding a number written with the separators of a locale, rewritten the way files
 * write it: "1.234,5" becomes "1234.5", "(12,5%)" becomes "(12.5%)". Everything around the
 * digits stays. Undefined when the text holds neither separator, or when one stands where
 * the locale would not put it: thousands separators need one to three digits before them
 * and three after each, none follows the decimal sign, and digits follow a decimal sign.
 */
export function fromLocalNumber(text: string, separators: NumberSeparators): string | undefined {
  const decimal = separators.decimal;
  const group = separators.group;
  const n = text.length;
  let out = '';
  let i = 0;
  let copied = 0;
  let found = false;
  while (i < n) {
    if (!isDigitCode(text.charCodeAt(i))) {
      const ch = text.charAt(i);
      // A thousands separator opens no number; a decimal sign may (",5").
      if (group !== '' && ch === group) return undefined;
      if (ch !== decimal) {
        i++;
        continue;
      }
    }
    // A number starts here: groups of digits, then a decimal part.
    let grouped = false;
    for (;;) {
      const start = i;
      while (i < n && isDigitCode(text.charCodeAt(i))) i++;
      const run = i - start;
      // The first group has one to three digits, every later one exactly three.
      if (grouped && run !== 3) return undefined;
      if (i >= n || group === '' || text.charAt(i) !== group) break;
      if (!grouped && (run < 1 || run > 3)) return undefined;
      out += text.substring(copied, i);
      copied = i + 1;
      grouped = true;
      found = true;
      i++;
    }
    if (i < n && text.charAt(i) === decimal) {
      out += text.substring(copied, i) + '.';
      found = true;
      i++;
      copied = i;
      const start = i;
      while (i < n && isDigitCode(text.charCodeAt(i))) i++;
      // A decimal sign with nothing after it ("5,") is not how a number is written.
      if (i === start) return undefined;
      // Neither separator follows a decimal part.
      if (i < n && (text.charAt(i) === decimal || (group !== '' && text.charAt(i) === group))) return undefined;
    }
  }
  return found ? out + text.substring(copied) : undefined;
}

/**
 * A plain number as text, as textToNumber() reads one; text that is none, and that is no
 * number or date of another kind the way files write them ("1,234", "3/4/2024"), is then
 * read with the separators of a locale, when there are any ("7,5", "1.234,5").
 */
export function plainNumberText(text: string, separators: NumberSeparators | undefined): number | undefined {
  const number = textToNumber(text);
  if (number !== undefined || separators === undefined) return number;
  if (text.length > MAX_NUMBER_TEXT) return undefined;
  const local = fromLocalNumber(text, separators);
  if (local === undefined) return undefined;
  const read = textToNumber(local);
  // What files write keeps its meaning: "1,234" is 1234 with a thousands separator, no plain number.
  return read === undefined || parseFileNumberText(text, undefined) !== undefined ? undefined : read;
}

/** parseNumberText() for text written the way files write numbers. */
function parseFileNumberText(text: string, context: CalcContext | undefined): number | undefined {
  const plain = textToNumber(text);
  if (plain !== undefined) return plain;
  if (text.length > MAX_NUMBER_TEXT || text.length === 0) return undefined;
  const c = text.charCodeAt(0);
  // Formulas and forced text are meanings of the cell editor, not of a value.
  if (c === 61 || c === 39) return undefined;
  const fraction = fractionText(text);
  if (fraction !== undefined) return fraction;
  const parsed = parseCellInput(text, inputOptions(context));
  return parsed.kind === 'number' && typeof parsed.value === 'number' && isFinite(parsed.value) ? parsed.value : undefined;
}

/**
 * Text that is no number and no date the way files write them, read the way a locale writes
 * numbers: plain numbers, percentages, thousands separators, currency. Never a date or a time.
 */
function parseLocalNumberText(text: string, separators: NumberSeparators): number | undefined {
  if (text.length > MAX_NUMBER_TEXT || text.length === 0) return undefined;
  const c = text.charCodeAt(0);
  if (c === 61 || c === 39) return undefined;
  const local = fromLocalNumber(text, separators);
  if (local === undefined) return undefined;
  const plain = textToNumber(local);
  if (plain !== undefined) return plain;
  const parsed = parseCellInput(local, { date1904: false });
  if (parsed.kind !== 'number' || typeof parsed.value !== 'number' || !isFinite(parsed.value)) return undefined;
  return isDateFormatId(parsed.numFmtId) ? undefined : parsed.value;
}

/**
 * Text as a number the way Excel reads it wherever a number is needed: plain numbers,
 * percentages, thousands separators, currency, fractions, dates and times (en-US conventions).
 * The context gives the date system, the year of dates written without one, and the locale:
 * text that is no number the way files write one is read with the decimal sign and the
 * thousands separator of the locale, where they differ ("7,5", "1.234,5"). What is a number
 * or a date the way files write it has the same value in every locale.
 */
export function parseNumberText(text: string, context?: CalcContext): number | undefined {
  const number = parseFileNumberText(text, context);
  if (number !== undefined) return number;
  const separators = numberSeparators(context);
  return separators === undefined ? undefined : parseLocalNumberText(text, separators);
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
  /**
   * One past the last row and column of a sheet that hold something (a value or a formula; of spills, those
   * that reach into `range`, the range being read), whatever was written and cleared since: the limit a whole
   * column or row is cut to when it is read as an array.
   */
  usedExtent(sheet: number, range?: CellRange): { rows: number; cols: number };
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
  /**
   * The locale of whoever looks at the workbook, when there is one: UPPER, LOWER and PROPER
   * follow its letters (in Turkish "i" is the small "İ" and "ı" the small "I"), and text
   * that is no number the way files write one is read with its decimal sign ("7,5" in
   * criteria, VALUE, text met by an operator). Without it letters change case the way they
   * do in English and only the point is a decimal sign. What functions write (TEXT, FIXED,
   * "&") does not depend on it.
   */
  readonly locale?: FormatLocale;
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
  /** Whether a row is hidden, by hand or by a filter. Without it no row counts as hidden. */
  rowHidden?(sheet: number, row: number): boolean;
  /**
   * Whether a row is hidden and lies among the data rows of a filter that has a criterion:
   * such rows are left out by every SUBTOTAL, rows hidden otherwise only by SUBTOTAL 101 to 111.
   */
  rowFiltered?(sheet: number, row: number): boolean;
  /**
   * The cells the dynamic array formula of a cell spills into ("A1#"), its own cell included;
   * an error when the formula cannot spill (#SPILL!), undefined when the cell holds no such
   * formula. Calculates the formula first when needed.
   */
  spillRange?(sheet: number, row: number, col: number): CellRange | CalcError | undefined;
}

/** Largest array a reference is read into. */
const MAX_REF_CELLS = 5000000;

/** Whether a range is a whole column (or several) or a whole row: "A:A", "1:1", "A:C". */
export function isWholeLine(range: CellRange): boolean {
  return (range.top <= 0 && range.bottom >= MAX_ROWS - 1) || (range.left <= 0 && range.right >= MAX_COLS - 1);
}

/**
 * The part of a range that is read as an array. A range whose bounds are written (A1:A30) is read
 * whole, empty cells included. A whole column or row (A:A, 1:1), and a range larger than `limit`, is
 * cut to the content of its sheet (CalcContext.usedExtent: the last row and column that hold a value,
 * a formula or a spill, never how far the sheet was once written), at least its first cell kept.
 */
export function readRange(range: CellRange, sheet: number, context: CalcContext, limit: number): CellRange {
  if (!isWholeLine(range) && (range.bottom - range.top + 1) * (range.right - range.left + 1) <= limit) return range;
  const extent = context.usedExtent(sheet, range);
  // A side is looked at only when it is long: a whole column needs how far the rows go, not the columns.
  return {
    top: range.top,
    left: range.left,
    bottom: range.bottom > range.top ? Math.min(range.bottom, Math.max(extent.rows - 1, range.top)) : range.bottom,
    right: range.right > range.left ? Math.min(range.right, Math.max(extent.cols - 1, range.left)) : range.right,
  };
}

/**
 * The cells of a reference as an array, all of them. A reference that is not a rectangle
 * gives a single #REF!, one too large to hold a single #NUM!. Only the abort of a
 * calculation's budget (CalcContext.tick) is thrown.
 */
export function refToArray(ref: RefValue, context: CalcContext): ArrayValue {
  const range = ref.range;
  if (!range || !(range.top >= 0) || !(range.left >= 0) || !(range.bottom >= range.top) || !(range.right >= range.left)) return scalarArray(ERROR_REF);
  // A range of written bounds whole, empty cells included; a whole column or row to the content of its sheet.
  const read = readRange(range, ref.sheet, context, MAX_REF_CELLS);
  const bottom = read.bottom;
  const right = read.right;
  const rows = bottom - range.top + 1;
  const cols = right - range.left + 1;
  if (!(rows >= 1 && cols >= 1) || Math.floor(rows) !== rows || Math.floor(cols) !== cols) return scalarArray(ERROR_REF);
  if (rows * cols > MAX_REF_CELLS) return scalarArray(ERROR_NUM);
  if (context.tick !== undefined) context.tick((rows * cols) >> 4);
  const data: Scalar[] = new Array(rows * cols).fill(null);
  let filled = 0;
  context.forEachCell(ref.sheet, { top: range.top, left: range.left, bottom, right }, (value, row, col) => {
    const r = row - range.top;
    const c = col - range.left;
    if (r >= 0 && r < rows && c >= 0 && c < cols) {
      const at = r * cols + c;
      data[at] = value;
      if (value !== null && at >= filled) filled = at + 1;
    }
  });
  const tail = tailFrom(rows * cols, filled);
  return tail === undefined ? { rows, cols, data } : { rows, cols, data, tail };
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
