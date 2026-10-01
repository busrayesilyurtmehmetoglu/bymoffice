// What a cell shows: its value run through the number format of its style.

import { resolveCellStyle } from '../model/sheet/color';
import type { Cell, StyleSheet, Workbook } from '../model/sheet/types';
import type { FormatLocale } from './locale';
import { formatValue, type FormatOptions, type FormattedValue } from './number-format';

const EMPTY: FormattedValue = { text: '' };

const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,9}))?)?)?(?:Z|[+-]\d{2}:?\d{2})?$/;
const DAY_MS = 86400000;
/** Days between the 1900 date system's day 0 (30 December 1899) and the Unix epoch. */
const EPOCH_1900 = 25569;
const EPOCH_SHIFT_1904 = 1462;

/**
 * ISO 8601 text of a cell of type "d" -> serial number, or undefined when the
 * text is not a date. A time zone suffix is ignored: the cell states wall-clock
 * time, as every other date in a sheet does.
 */
export function isoToSerial(text: string, date1904 = false): number | undefined {
  const match = ISO_DATE.exec(text.trim());
  if (!match) return undefined;
  const [year, month, day] = [Number(match[1]), Number(match[2]), Number(match[3])];
  const utc = new Date(0);
  utc.setUTCFullYear(year, month - 1, day);
  // The Date object accepts 31 February and moves on; a date is only valid if it comes back unchanged.
  if (utc.getUTCFullYear() !== year || utc.getUTCMonth() !== month - 1 || utc.getUTCDate() !== day) return undefined;
  const [hour, minute, second] = [Number(match[4] || 0), Number(match[5] || 0), Number(match[6] || 0)];
  if (hour > 23 || minute > 59 || second > 59) return undefined;
  const fraction = match[7] ? Number(`0.${match[7]}`) : 0;
  let serial = utc.getTime() / DAY_MS + EPOCH_1900;
  if (date1904) serial -= EPOCH_SHIFT_1904;
  // The 1900 system counts a 29 February 1900 that never was: days before 1 March 1900 are one lower.
  else if (serial < 61) serial -= 1;
  return serial + (hour * 3600 + minute * 60 + second + fraction) / 86400;
}

/**
 * Formats a cell with a number format code. Errors are shown as they are.
 * ISO date cells are formatted as the date they state; with a "General" format
 * or text that is not a date they keep their text.
 */
export function formatCell(cell: Cell | undefined, code: string, options?: FormatOptions): FormattedValue {
  if (!cell || cell.value === null) return EMPTY;
  if (cell.kind === 'error') return { text: String(cell.value) };
  if (cell.kind === 'date') {
    const text = String(cell.value);
    const serial = /^general$/i.test(code) ? undefined : isoToSerial(text, options ? options.date1904 : false);
    return serial === undefined ? { text } : formatValue(serial, code, options);
  }
  return formatValue(cell.value, code, options);
}

/**
 * The id to treat as a built-in format, if any. Built-in dates (14, 22) follow
 * the locale instead of their code, but only when the file has not redefined
 * that id with a code of its own.
 */
export function builtinFormatId(styles: StyleSheet, numFmtId: number): number | undefined {
  return numFmtId < 164 && !styles.numFmts.has(numFmtId) ? numFmtId : undefined;
}

/** The text a cell shows in the grid, with the workbook's styles and date system. */
export function cellDisplayText(workbook: Workbook, cell: Cell | undefined, locale?: FormatLocale): string {
  if (!cell || cell.value === null) return '';
  const style = resolveCellStyle(workbook.styles, cell.style);
  const options = { locale, date1904: workbook.date1904, builtinId: builtinFormatId(workbook.styles, style.numFmtId) };
  return formatCell(cell, style.numFmt, options).text;
}
