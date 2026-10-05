// Pointing at cells while a formula is typed ("point mode"): where in the text
// a click on a cell adds a reference, what that reference is written as, and
// how a reference that was just added is replaced by the next one. Nothing in
// here touches the DOM; the workbook view does the pointing itself.

import { quoteSheetName } from '../calc/refs';
import { MAX_COLS, MAX_ROWS, cellName, columnName, type CellRange } from '../model/sheet/types';

/** Longer texts are never pointed into: no formula is this long, and every press would read all of it. */
export const MAX_POINT_TEXT = 32768;

/** What a reference can follow: the "=" of the formula, an operator, a separator of arguments, an opening bracket. */
const BEFORE = '=+-*/^&,;(:<>';
/** What can follow a reference that is put before it. */
const AFTER = '+-*/^&,;):<>=';

function isSpace(code: number): boolean {
  return code === 32 || code === 9 || code === 10 || code === 13 || code === 160;
}

/** Whether a text is a formula being typed: it begins with "=". */
export function isFormulaText(text: string): boolean {
  return typeof text === 'string' && text.charCodeAt(0) === 61;
}

export interface PointContext {
  /**
   * The decimal separator of the locale the formula is typed in. With "," a comma right
   * behind a digit is taken for the comma of a number ("=1,"), not for a separator.
   */
  decimal?: string;
}

const isDigit = (code: number) => code >= 48 && code <= 57;

/** Whether the text ends in the mantissa and "E" of a number in scientific notation ("=1E", "=A1+1.5e"). */
const SCIENTIFIC = /(?:^|[^A-Za-z0-9_.$])(?:\d+\.?\d*|\.\d+)[Ee]$/;
/** Whether the text ends in the beginning of an error constant that goes on with "/" ("#N", "#DIV", "#SAYI"). */
const ERROR_HEAD = /#[^\s\/#!?()+\-*^&=<>,;:"']*$/;

/**
 * Whether a reference can be put into a formula that is being typed, in the
 * place of the characters from `start` to `end` (the caret when they are the
 * same): the text is a formula, the place is not inside a text in quotes or a
 * sheet name in apostrophes, what stands before it (spaces aside) is "=", an
 * operator, a separator or "(", and what stands after it is the end of the
 * text, an operator, a separator or ")".
 *
 * After an operand ("=A1", "=A1 ", "=SUM(A1)") the answer is no, also after a
 * space: a space between two references is the intersection operator, which
 * nobody means by a click. The same after "!", inside "{}" and inside the
 * square brackets of a reference to a table, and where the sign before the
 * place belongs to a number ("=1E+") or to an error constant ("=#N/",
 * "=#DIV/"). Where the locale writes numbers with a decimal comma, a comma
 * right behind a digit is no place either.
 */
export function canPointAt(text: string, start: number, end: number = start, context: PointContext = {}): boolean {
  if (!isFormulaText(text) || text.length > MAX_POINT_TEXT) return false;
  if (!Number.isInteger(start) || !Number.isInteger(end) || start < 1 || end < start || end > text.length) return false;
  // Quotes open and close; a doubled quote inside a text closes and opens again, which comes to the same.
  let quoted = false;
  let named = false;
  // Braces hold constants only, and square brackets the names of a table's columns.
  let braces = 0;
  let brackets = 0;
  for (let i = 1; i < end; i++) {
    if (i === start && (quoted || named || braces > 0 || brackets > 0)) return false;
    const code = text.charCodeAt(i);
    if (code === 34 && !named) quoted = !quoted;
    else if (code === 39 && !quoted && brackets === 0) named = !named;
    else if (!quoted && !named) {
      if (code === 123) braces++;
      else if (code === 125 && braces > 0) braces--;
      else if (code === 91) brackets++;
      else if (code === 93 && brackets > 0) brackets--;
    }
  }
  if (quoted || named || braces > 0 || brackets > 0) return false;

  let before = start - 1;
  while (before > 0 && isSpace(text.charCodeAt(before))) before--;
  const sign = text.charAt(before);
  if (BEFORE.indexOf(sign) < 0) return false;
  // Right behind the sign of an exponent, the "/" of an error constant or the comma of a number: not an operator.
  if ((sign === '+' || sign === '-') && SCIENTIFIC.test(text.substring(0, before))) return false;
  if (sign === '/' && ERROR_HEAD.test(text.substring(1, before))) return false;
  if (sign === ',' && context.decimal === ',' && before > 1 && isDigit(text.charCodeAt(before - 1))) return false;
  let after = end;
  while (after < text.length && isSpace(text.charCodeAt(after))) after++;
  return after >= text.length || AFTER.indexOf(text.charAt(after)) >= 0;
}

export interface PointReferenceOptions {
  /**
   * The range is one cell although it spans several: a merged cell, which is
   * written as its top-left cell.
   */
  single?: boolean;
  /** Name of the sheet the range is on, when that is not the sheet of the formula. */
  sheet?: string;
}

function index(value: number, limit: number): number {
  if (!Number.isInteger(value) || value < 0 || value >= limit) throw new RangeError('The reference is outside the sheet');
  return value;
}

/**
 * A range as a formula names it, always relative: "B5", "B5:D9", whole
 * columns "B:D", whole rows "3:7", on another sheet "Sheet2!B5" or
 * "'Sales 2026'!B5". The corners may come in any order. The text is the same
 * in every locale. Throws RangeError for a range outside the sheet.
 */
export function pointReference(range: CellRange, options: PointReferenceOptions = {}): string {
  const top = Math.min(index(range.top, MAX_ROWS), index(range.bottom, MAX_ROWS));
  const bottom = Math.max(range.top, range.bottom);
  const left = Math.min(index(range.left, MAX_COLS), index(range.right, MAX_COLS));
  const right = Math.max(range.left, range.right);
  const rows = left === 0 && right === MAX_COLS - 1;
  const cols = top === 0 && bottom === MAX_ROWS - 1;
  let text: string;
  if (rows) text = `${top + 1}:${bottom + 1}`;
  else if (cols) text = `${columnName(left)}:${columnName(right)}`;
  else if (options.single === true || (top === bottom && left === right)) text = cellName(top, left);
  else text = `${cellName(top, left)}:${cellName(bottom, right)}`;
  return options.sheet !== undefined ? `${quoteSheetName(String(options.sheet))}!${text}` : text;
}

/** Where in the text a reference was put by pointing, and what was put there. */
export interface PointSpan {
  start: number;
  end: number;
  text: string;
}

/**
 * Whether the reference a span stands for is still there and the caret right
 * behind it: the next reference pointed at then takes its place.
 */
export function pointSpanLive(text: string, caretStart: number, caretEnd: number, span: PointSpan): boolean {
  if (!isFormulaText(text) || caretStart !== caretEnd || caretEnd !== span.end) return false;
  if (!(span.start >= 1) || span.end !== span.start + span.text.length || span.end > text.length) return false;
  return text.substring(span.start, span.end) === span.text;
}

/**
 * Puts a reference in the place of the characters from `start` to `end` of a
 * text. Returns the new text and the span the reference has in it; the caret
 * belongs at the end of the span.
 */
export function putReference(text: string, start: number, end: number, reference: string): { text: string; span: PointSpan } {
  const length = text.length;
  const from = Math.min(Math.max(Number.isInteger(start) ? start : length, 0), length);
  const to = Math.min(Math.max(Number.isInteger(end) ? end : from, from), length);
  return {
    text: text.substring(0, from) + reference + text.substring(to),
    span: { start: from, end: from + reference.length, text: reference },
  };
}
