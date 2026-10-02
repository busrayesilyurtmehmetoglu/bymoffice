// Text typed into a cell -> the value Excel would store, and back: the text the
// cell editor shows for a cell. Both follow the locale: "1.234,56" and
// "15.03.2024" in Turkish, "1,234.56" and "3/15/2024" in English.
//
// Nothing here throws, and nothing is slow on huge input: only short text is
// ever examined as a number or a date, by hand-written scanners.

import { LOCALE_EN, LOCALE_TR, normalizeLocale, type FormatLocale } from '../format/locale';
import { isDateFormat, serialToDate } from '../format/number-format';
import type { Cell } from '../model/sheet/types';
import { formulaFromLocal, formulaToLocal, localErrorCode } from './formula-locale';

export interface InputOptions {
  locale?: FormatLocale;
  date1904?: boolean;
  /** The current date, for typed dates without a year ("15.03"). Default: now. */
  today?: Date;
  /** Whether the workbook has a defined name of this spelling, for words of a formula that are names in it and constants elsewhere. */
  isName?(name: string): boolean;
}

export interface ParsedInput {
  kind: 'blank' | 'number' | 'string' | 'boolean' | 'error';
  value: number | string | boolean | null;
  /** Formula text without the leading "=", when the input is a formula. */
  formula?: string;
  /**
   * The input is a formula that is not written the way the locale writes formulas (a comma
   * that is not a decimal comma): `formula` is the text as typed, and must not be stored.
   */
  invalid?: boolean;
  /**
   * Number format the input implies ("0%", a date or time code, currency), as a built-in
   * format id where one exists (9, 10, 14, 20, 21, 22, 4 ...) or a format code otherwise.
   * The caller applies it only when the cell still has the General format.
   */
  numFmtId?: number;
  numFmt?: string;
}

/** Longest text a cell holds. */
const MAX_TEXT_LENGTH = 32767;
/** Longer input is never a number, a date, a boolean or an error. */
const MAX_VALUE_LENGTH = 64;

const ERRORS = [
  '#N/A', '#DIV/0!', '#REF!', '#VALUE!', '#NAME?', '#NUM!', '#NULL!',
  '#SPILL!', '#GETTING_DATA', '#CALC!', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#UNKNOWN!', '#BUSY!',
];

/** Currency symbols recognised before or after a number; compared after fold(). */
const CURRENCIES = ['₺', 'TL', '$', '€', '£'];

// Built-in number format ids.
const FMT_THOUSANDS = 3;
const FMT_THOUSANDS_2 = 4;
const FMT_PERCENT = 9;
const FMT_PERCENT_2 = 10;
const FMT_DATE = 14;
const FMT_DAY_MONTH_YEAR = 15;
const FMT_DAY_MONTH = 16;
const FMT_MONTH_YEAR = 17;
const FMT_TIME_12 = 18;
const FMT_TIME_12_SECONDS = 19;
const FMT_TIME = 20;
const FMT_TIME_SECONDS = 21;
const FMT_DATE_TIME = 22;
const FMT_ELAPSED = 46;
const FMT_SUBSECONDS = 47;

const MAX_SERIAL_1900 = 2958465;
const MAX_SERIAL_1904 = 2957003;

// ---------------------------------------------------------------------------
// Characters

function isDigit(code: number): boolean {
  return code >= 48 && code <= 57;
}

function isSpace(code: number): boolean {
  return code === 32 || code === 9 || code === 0xa0 || code === 10 || code === 13;
}

function allDigits(text: string): boolean {
  if (text.length === 0) return false;
  for (let i = 0; i < text.length; i++) if (!isDigit(text.charCodeAt(i))) return false;
  return true;
}

function trim(text: string): string {
  let start = 0;
  let end = text.length;
  while (start < end && isSpace(text.charCodeAt(start))) start++;
  while (end > start && isSpace(text.charCodeAt(end - 1))) end--;
  return start === 0 && end === text.length ? text : text.substring(start, end);
}

/**
 * Case folding that also works for Turkish: "I", "İ", "ı" and "i" all become "i", so
 * "DOĞRU" = "doğru", "YANLIŞ" = "yanlış", "MAYIS" = "Mayıs" without a locale-aware API.
 */
function fold(text: string): string {
  return text.replace(/[İIı]/g, 'i').toLowerCase();
}

function zeros(n: number): string {
  let out = '';
  for (let i = 0; i < n; i++) out += '0';
  return out;
}

// ---------------------------------------------------------------------------
// Numbers

// Results of parseNumber().
let nValue = 0;
let nDecimals = false;
let nGrouped = false;
let nExponent = false;
/** A whole number written out digit by digit that a double cannot hold exactly. */
let nInexact = false;

/**
 * An unsigned number in the locale's notation: digits with correctly placed thousands
 * separators, a decimal part, an exponent. `shift` is added to the exponent ("12,5%" is
 * read as 12.5e-2, which rounds exactly like 0.125). False when the text is anything else.
 */
function parseUnsigned(text: string, locale: FormatLocale, shift: number): boolean {
  const length = text.length;
  const decimal = locale.decimal;
  const group = locale.group === decimal ? '' : locale.group;
  let i = 0;
  let digits = '';
  let groupDigits = 0;
  let grouped = false;
  while (i < length) {
    const code = text.charCodeAt(i);
    if (isDigit(code)) {
      digits += text.charAt(i);
      groupDigits++;
      i++;
    } else if (group.length === 1 && text.charAt(i) === group) {
      // First group: one to three digits; every later group exactly three.
      if (grouped ? groupDigits !== 3 : groupDigits < 1 || groupDigits > 3) return false;
      grouped = true;
      groupDigits = 0;
      i++;
    } else {
      break;
    }
  }
  if (grouped && groupDigits !== 3) return false;
  let fraction = '';
  let hasPoint = false;
  if (i < length && decimal.length === 1 && text.charAt(i) === decimal) {
    hasPoint = true;
    i++;
    while (i < length && isDigit(text.charCodeAt(i))) {
      fraction += text.charAt(i);
      i++;
    }
  }
  if (digits.length === 0 && fraction.length === 0) return false;
  let exponent = 0;
  let hasExponent = false;
  if (i < length && (text.charAt(i) === 'e' || text.charAt(i) === 'E')) {
    i++;
    let sign = 1;
    if (i < length && (text.charAt(i) === '+' || text.charAt(i) === '-')) {
      if (text.charAt(i) === '-') sign = -1;
      i++;
    }
    const start = i;
    while (i < length && isDigit(text.charCodeAt(i))) i++;
    if (i === start || i - start > 4) return false;
    exponent = sign * Number(text.substring(start, i));
    hasExponent = true;
  }
  if (i !== length) return false;
  const value = Number((digits || '0') + '.' + (fraction || '0') + 'e' + (exponent + shift));
  if (!isFinite(value)) return false;
  nValue = value;
  nDecimals = hasPoint && fraction.length > 0;
  nGrouped = grouped;
  nExponent = hasExponent;
  nInexact = false;
  if (!hasPoint && !hasExponent && shift === 0 && digits.length > 15) {
    let start = 0;
    while (start < digits.length - 1 && digits.charCodeAt(start) === 48) start++;
    nInexact = String(value) !== digits.substring(start);
  }
  return true;
}

/** A number with an optional sign, or in parentheses for a negative: "-12,5", "(12,5)". */
function parseNumber(text: string, locale: FormatLocale, shift: number): boolean {
  const length = text.length;
  if (length === 0) return false;
  const first = text.charAt(0);
  if (first === '(' && length > 2 && text.charAt(length - 1) === ')') {
    if (!parseUnsigned(trim(text.substring(1, length - 1)), locale, shift)) return false;
    nValue = -nValue;
    return true;
  }
  if (first === '-' || first === '+') {
    if (!parseUnsigned(text.substring(1), locale, shift)) return false;
    if (first === '-') nValue = -nValue;
    return true;
  }
  return parseUnsigned(text, locale, shift);
}

function numberResult(value: number, numFmtId?: number, numFmt?: string): ParsedInput {
  // -0 is not a value a cell holds.
  const result: ParsedInput = { kind: 'number', value: value === 0 ? 0 : value };
  if (numFmtId !== undefined) result.numFmtId = numFmtId;
  if (numFmt !== undefined) result.numFmt = numFmt;
  return result;
}

function tryPercent(text: string, locale: FormatLocale): ParsedInput | null {
  const length = text.length;
  let body: string;
  if (text.charAt(length - 1) === '%') body = trim(text.substring(0, length - 1));
  else if (text.charAt(0) === '%') body = trim(text.substring(1));
  else return null;
  if (!parseNumber(body, locale, -2)) return null;
  return numberResult(nValue, nDecimals ? FMT_PERCENT_2 : FMT_PERCENT);
}

function tryPlainNumber(text: string, locale: FormatLocale): ParsedInput | null {
  if (!parseNumber(text, locale, 0)) return null;
  // Long identifiers (card, account, phone numbers) would lose digits as a number: Excel
  // rounds them to 15 digits, this engine keeps them as the text that was typed.
  if (nInexact) return null;
  // Scientific input suggests no format: retyping what the editor shows for a General
  // cell ("1E+20") must not change the cell's format.
  if (nGrouped) return numberResult(nValue, nDecimals ? FMT_THOUSANDS_2 : FMT_THOUSANDS);
  return numberResult(nValue);
}

function tryCurrency(text: string, locale: FormatLocale): ParsedInput | null {
  let body = text;
  let negative = false;
  if (body.charAt(0) === '(' && body.length > 2 && body.charAt(body.length - 1) === ')') {
    negative = true;
    body = trim(body.substring(1, body.length - 1));
  } else if (body.charAt(0) === '-') {
    negative = true;
    body = body.substring(1);
  } else if (body.charAt(0) === '+') {
    body = body.substring(1);
  }
  const folded = fold(body);
  let symbol = '';
  let prefix = false;
  for (let i = 0; i < CURRENCIES.length && symbol === ''; i++) {
    const candidate = fold(CURRENCIES[i]);
    if (folded.length <= candidate.length) continue;
    if (folded.indexOf(candidate) === 0) {
      symbol = CURRENCIES[i];
      prefix = true;
      body = trim(body.substring(candidate.length));
    } else if (folded.lastIndexOf(candidate) === folded.length - candidate.length) {
      symbol = CURRENCIES[i];
      body = trim(body.substring(0, body.length - candidate.length));
    }
  }
  if (symbol === '') return null;
  // "$-5" is negative too, but one sign is enough.
  if (prefix && !negative && body.charAt(0) === '-') {
    negative = true;
    body = body.substring(1);
  }
  if (!parseUnsigned(body, locale, 0) || nExponent || nInexact) return null;
  const digits = nDecimals ? '#,##0.00' : '#,##0';
  const code = prefix ? '"' + symbol + '"' + digits : digits + ' "' + symbol + '"';
  return numberResult(negative ? -nValue : nValue, undefined, code);
}

// ---------------------------------------------------------------------------
// Dates and times

function isLeapYear(year: number): boolean {
  return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0;
}

const MONTH_DAYS = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

/** Days from 1 January 1970 to a date of the proleptic Gregorian calendar. */
function daysFromCivil(year: number, month: number, day: number): number {
  const y = month <= 2 ? year - 1 : year;
  const era = Math.floor(y / 400);
  const yearOfEra = y - era * 400;
  const dayOfYear = Math.floor((153 * (month + (month > 2 ? -3 : 9)) + 2) / 5) + day - 1;
  const dayOfEra = yearOfEra * 365 + Math.floor(yearOfEra / 4) - Math.floor(yearOfEra / 100) + dayOfYear;
  return era * 146097 + dayOfEra - 719468;
}

/**
 * Serial of a calendar date, or -1 when there is no such date in the date system.
 * The 1900 system has a 29 February 1900 (serial 60) that never existed.
 */
function dateSerial(year: number, month: number, day: number, date1904: boolean): number {
  if (year < 1900 || year > 9999 || month < 1 || month > 12 || day < 1) return -1;
  if (!date1904 && year === 1900 && month === 2 && day === 29) return 60;
  const limit = month === 2 && isLeapYear(year) ? 29 : MONTH_DAYS[month - 1];
  if (day > limit) return -1;
  const days = daysFromCivil(year, month, day) + 25569;
  if (date1904) return days >= 1462 ? days - 1462 : -1;
  return days < 61 ? days - 1 : days;
}

/** Two-digit years: 00-29 are 2000-2029, 30-99 are 1930-1999. -1 for a year that cannot be. */
function fullYear(text: string): number {
  const value = Number(text);
  if (text.length === 4) return value;
  if (text.length > 2) return -1;
  return value < 30 ? 2000 + value : 1900 + value;
}

const foldedMonths = new WeakMap<object, string[][]>();
const ENGLISH_MONTHS = [LOCALE_EN.monthNames.map(fold), LOCALE_EN.shortMonthNames.map(fold)];

/** Month number for a name in the locale's language or in English; 0 when it is neither. */
function monthNumber(name: string, locale: FormatLocale): number {
  let lists = foldedMonths.get(locale);
  if (lists === undefined) {
    lists = [locale.monthNames.map(fold), locale.shortMonthNames.map(fold)].concat(ENGLISH_MONTHS);
    foldedMonths.set(locale, lists);
  }
  const folded = fold(name);
  for (let i = 0; i < lists.length; i++) {
    const index = lists[i].indexOf(folded);
    if (index >= 0) return index + 1;
  }
  return 0;
}

// Results of parseDate().
let dSerial = 0;
let dFormat = FMT_DATE;

function setDate(year: number, month: number, day: number, date1904: boolean, format: number): boolean {
  const serial = dateSerial(year, month, day, date1904);
  if (serial < 0) return false;
  dSerial = serial;
  dFormat = format;
  return true;
}

function isDayOrMonth(text: string): boolean {
  return text.length <= 2 && allDigits(text);
}

/** "15.03.2024", "15/3/24", "2024-03-15", "15.03", "3/2024": digits around one kind of separator. */
function parseNumericDate(text: string, locale: FormatLocale, date1904: boolean, thisYear: number): boolean {
  let i = 0;
  while (i < text.length && isDigit(text.charCodeAt(i))) i++;
  if (i === 0 || i === text.length) return false;
  const separator = text.charAt(i);
  if (separator !== '.' && separator !== '/' && separator !== '-') return false;
  // Dots separate a date only where the locale writes dates with dots ("15.03.2024").
  if (separator === '.' && locale.dateSeparator !== '.') return false;
  const parts = text.split(separator);
  if (parts.length < 2 || parts.length > 3) return false;
  for (let p = 0; p < parts.length; p++) if (!allDigits(parts[p]) || parts[p].length > 4) return false;
  const order = locale.dateOrder || 'mdy';

  if (parts.length === 2) {
    // "3/2024" and "2024-03": a month of a year.
    if (parts[1].length === 4 && isDayOrMonth(parts[0])) {
      return setDate(Number(parts[1]), Number(parts[0]), 1, date1904, FMT_MONTH_YEAR);
    }
    if (parts[0].length === 4 && isDayOrMonth(parts[1])) {
      return setDate(Number(parts[0]), Number(parts[1]), 1, date1904, FMT_MONTH_YEAR);
    }
    if (!isDayOrMonth(parts[0]) || !isDayOrMonth(parts[1])) return false;
    const day = Number(order === 'dmy' ? parts[0] : parts[1]);
    const month = Number(order === 'dmy' ? parts[1] : parts[0]);
    return setDate(thisYear, month, day, date1904, FMT_DAY_MONTH);
  }

  let yearText: string;
  let monthText: string;
  let dayText: string;
  if (parts[0].length === 4 || order === 'ymd') {
    // A four-digit year in front is year-month-day in every locale.
    yearText = parts[0];
    monthText = parts[1];
    dayText = parts[2];
  } else if (order === 'dmy') {
    dayText = parts[0];
    monthText = parts[1];
    yearText = parts[2];
  } else {
    monthText = parts[0];
    dayText = parts[1];
    yearText = parts[2];
  }
  if (!isDayOrMonth(monthText) || !isDayOrMonth(dayText)) return false;
  const year = fullYear(yearText);
  if (year < 0) return false;
  return setDate(year, Number(monthText), Number(dayText), date1904, FMT_DATE);
}

/** "15 Mart 2024", "15-Mar-24", "March 15, 2024", "15 Mar", "Mart 2024": a month by name. */
function parseNamedDate(text: string, locale: FormatLocale, date1904: boolean, thisYear: number): boolean {
  const tokens: string[] = [];
  let start = -1;
  for (let i = 0; i <= text.length; i++) {
    const c = i < text.length ? text.charAt(i) : ' ';
    const separator = c === ' ' || c === ',' || c === '.' || c === '-' || c === '/' || c === ' ';
    if (separator) {
      if (start >= 0) tokens.push(text.substring(start, i));
      start = -1;
    } else if (start < 0) {
      start = i;
    }
  }
  if (tokens.length < 2 || tokens.length > 3) return false;
  let month = 0;
  let monthAt = -1;
  const numbers: string[] = [];
  for (let i = 0; i < tokens.length; i++) {
    if (allDigits(tokens[i])) {
      if (tokens[i].length > 4) return false;
      numbers.push(tokens[i]);
    } else {
      if (monthAt >= 0) return false;
      month = monthNumber(tokens[i], locale);
      if (month === 0) return false;
      monthAt = i;
    }
  }
  if (monthAt < 0 || monthAt === 2) return false;

  if (tokens.length === 3) {
    // "15 Mart 2024" or "March 15, 2024": the day, then the year.
    if (!isDayOrMonth(numbers[0])) return false;
    const year = fullYear(numbers[1]);
    if (year < 0) return false;
    return setDate(year, month, Number(numbers[0]), date1904, FMT_DAY_MONTH_YEAR);
  }
  const number = numbers[0];
  if (number.length === 4) return setDate(Number(number), month, 1, date1904, FMT_MONTH_YEAR);
  if (!isDayOrMonth(number)) return false;
  return setDate(thisYear, month, Number(number), date1904, FMT_DAY_MONTH);
}

function parseDate(text: string, locale: FormatLocale, date1904: boolean, options: InputOptions | undefined): boolean {
  if (text.length < 3) return false;
  let thisYear = NaN;
  const today = options !== undefined ? options.today : undefined;
  if (today instanceof Date) thisYear = today.getFullYear();
  if (thisYear !== thisYear) thisYear = new Date().getFullYear();
  if (isDigit(text.charCodeAt(0)) && parseNumericDate(text, locale, date1904, thisYear)) return true;
  return parseNamedDate(text, locale, date1904, thisYear);
}

// Results of parseTime().
let tValue = 0;
let tFormat = FMT_TIME;
let tElapsed = false;

// Result of stripMeridiem(): the text without the marker.
let stripped = '';

/** Removes an AM/PM marker from the end of the text: 1 for AM, 2 for PM, 0 when there is none. */
function stripMeridiem(text: string, locale: FormatLocale): number {
  const folded = fold(text);
  stripped = text;
  // Positions are taken from the folded text; they only apply if folding kept the length.
  if (folded.length !== text.length) return 0;
  const markers = [fold(locale.am), fold(locale.pm), 'am', 'pm', 'a', 'p'];
  for (let i = 0; i < markers.length; i++) {
    const marker = markers[i];
    const at = folded.length - marker.length;
    if (marker.length === 0 || at < 1 || folded.lastIndexOf(marker) !== at) continue;
    const before = folded.charCodeAt(at - 1);
    if (!isSpace(before)) {
      // "1:30PM" is a time; "5a", "12A", "9am" are codes someone typed, not times.
      if (!isDigit(before) || text.indexOf(':') < 0) continue;
    }
    stripped = trim(text.substring(0, at));
    return i % 2 === 0 ? 1 : 2;
  }
  stripped = text;
  return 0;
}

/** "13:45", "13:45:30", "13:45:30,5", "1:30 PM", "1 PM", "25:00". */
function parseTime(text: string, locale: FormatLocale): boolean {
  const meridiem = stripMeridiem(text, locale);
  const parts = stripped.split(':');
  if (parts.length > 3 || (parts.length === 1 && meridiem === 0)) return false;
  if (!allDigits(parts[0]) || parts[0].length > 4) return false;
  let hours = Number(parts[0]);
  let minutes = 0;
  let seconds = 0;
  let subseconds = false;
  if (parts.length > 1) {
    if (!isDayOrMonth(parts[1])) return false;
    minutes = Number(parts[1]);
  }
  if (parts.length > 2) {
    let whole = parts[2];
    let fraction = '';
    let point = whole.indexOf('.');
    if (point < 0 && locale.decimal.length === 1) point = whole.indexOf(locale.decimal);
    if (point >= 0) {
      fraction = whole.substring(point + 1);
      whole = whole.substring(0, point);
      if (!allDigits(fraction) || fraction.length > 9) return false;
      subseconds = true;
    }
    if (!isDayOrMonth(whole)) return false;
    seconds = Number(whole + (subseconds ? '.' + fraction : ''));
  }
  if (minutes > 59 || seconds >= 60) return false;
  if (meridiem !== 0) {
    if (hours > 12) return false;
    if (hours === 12) hours = 0;
    if (meridiem === 2) hours += 12;
  }
  tValue = (hours * 3600 + minutes * 60 + seconds) / 86400;
  tElapsed = hours >= 24;
  if (subseconds) tFormat = FMT_SUBSECONDS;
  else if (tElapsed) tFormat = FMT_ELAPSED;
  else if (meridiem !== 0) tFormat = parts.length > 2 ? FMT_TIME_12_SECONDS : FMT_TIME_12;
  else tFormat = parts.length > 2 ? FMT_TIME_SECONDS : FMT_TIME;
  return true;
}

function tryDateTime(text: string, locale: FormatLocale, options: InputOptions | undefined): ParsedInput | null {
  const date1904 = options !== undefined && options.date1904 === true;
  const colon = text.indexOf(':');
  if (colon < 0) {
    if (parseDate(text, locale, date1904, options)) return numberResult(dSerial, dFormat);
    // "1 PM"
    if (parseTime(text, locale)) return numberResult(tValue, tFormat);
    return null;
  }
  // The time starts at the last space before the first colon; "T" joins an ISO date and time.
  let split = text.lastIndexOf(' ', colon);
  if (split < 0) {
    const iso = text.toUpperCase().lastIndexOf('T', colon);
    if (iso > 0 && isDigit(text.charCodeAt(iso - 1))) split = iso;
  }
  if (split < 0) return parseTime(text, locale) ? numberResult(tValue, tFormat) : null;
  const datePart = trim(text.substring(0, split));
  if (datePart.length === 0) return parseTime(text, locale) ? numberResult(tValue, tFormat) : null;
  if (!parseTime(trim(text.substring(split + 1)), locale) || tElapsed) return null;
  if (!parseDate(datePart, locale, date1904, options)) return null;
  if (dSerial + tValue > (date1904 ? MAX_SERIAL_1904 : MAX_SERIAL_1900) + 1) return null;
  return numberResult(dSerial + tValue, FMT_DATE_TIME);
}

// ---------------------------------------------------------------------------
// parseCellInput

function stringResult(text: string): ParsedInput {
  return { kind: 'string', value: text.length > MAX_TEXT_LENGTH ? text.substring(0, MAX_TEXT_LENGTH) : text };
}

function resolveLocale(options: InputOptions | undefined): FormatLocale {
  if (options === undefined || options === null) return LOCALE_EN;
  const locale = options.locale as FormatLocale;
  return locale === LOCALE_EN || locale === LOCALE_TR ? locale : normalizeLocale(locale);
}

/** What Excel would store for text typed into a cell. */
export function parseCellInput(text: string, options?: InputOptions): ParsedInput {
  if (typeof text !== 'string' || text.length === 0) return { kind: 'blank', value: null };
  const first = text.charAt(0);
  if (first === '=') {
    if (text.length === 1) return { kind: 'string', value: text };
    // Typed the way the locale writes formulas (";" between arguments, "1,5"); stored the one way files have them.
    const typed = text.substring(1);
    try {
      return { kind: 'blank', value: null, formula: formulaFromLocal(typed, resolveLocale(options), options) };
    } catch (error) {
      return { kind: 'blank', value: null, formula: typed, invalid: true };
    }
  }
  // A leading apostrophe means "text, exactly as typed".
  if (first === "'") return stringResult(text.substring(1));
  if (text.length > MAX_VALUE_LENGTH * 4) return stringResult(text);
  const body = trim(text);
  if (body.length === 0 || body.length > MAX_VALUE_LENGTH) return stringResult(text);

  const locale = resolveLocale(options);
  const folded = fold(body);
  if (folded === 'true' || folded === fold(locale.trueText)) return { kind: 'boolean', value: true };
  if (folded === 'false' || folded === fold(locale.falseText)) return { kind: 'boolean', value: false };
  if (first === '#' || body.charAt(0) === '#') {
    const index = ERRORS.indexOf(body.toUpperCase());
    if (index >= 0) return { kind: 'error', value: ERRORS[index] };
    // The locale's own name of an error ("#YOK"); what is stored and shown is the English one.
    const local = localErrorCode(body, locale);
    return local !== undefined ? { kind: 'error', value: local } : stringResult(text);
  }
  const result =
    tryPlainNumber(body, locale) ||
    tryPercent(body, locale) ||
    tryCurrency(body, locale) ||
    tryDateTime(body, locale, options);
  return result || stringResult(text);
}

// ---------------------------------------------------------------------------
// cellEditText

/**
 * A number the way it is typed: the shortest digits that give exactly this number back,
 * the locale's decimal separator, no thousands separators; "1E+20" notation outside
 * 1E-9 .. 1E+15. `shift` moves the decimal point (2 for a percentage).
 */
function numberText(value: number, shift: number, locale: FormatLocale): string {
  if (value !== value || value === Infinity || value === -Infinity) return '#NUM!';
  if (value === 0) return '0';
  // Shortest digits that give the number back: "d.ddde+x".
  const text = Math.abs(value).toExponential();
  const mark = text.indexOf('e');
  const digits = text.charAt(0) + (mark > 2 ? text.substring(2, mark) : '');
  const exponent = Number(text.substring(mark + 1)) + shift;
  let out = value < 0 ? '-' : '';
  if (exponent >= 15 || exponent < -9) {
    out += digits.charAt(0);
    if (digits.length > 1) out += locale.decimal + digits.substring(1);
    const magnitude = String(Math.abs(exponent));
    return out + 'E' + (exponent < 0 ? '-' : '+') + (magnitude.length < 2 ? '0' + magnitude : magnitude);
  }
  if (exponent < 0) return out + '0' + locale.decimal + zeros(-exponent - 1) + digits;
  const whole = exponent + 1;
  if (digits.length <= whole) return out + digits + zeros(whole - digits.length);
  return out + digits.substring(0, whole) + locale.decimal + digits.substring(whole);
}

/** True when the format code has a "%" outside quotes, escapes and brackets. */
function isPercentFormat(code: string): boolean {
  if (typeof code !== 'string' || code.length > 1024) return false;
  for (let i = 0; i < code.length; i++) {
    const c = code.charAt(i);
    if (c === '%') return true;
    if (c === '"') {
      const end = code.indexOf('"', i + 1);
      if (end < 0) return false;
      i = end;
    } else if (c === '[') {
      const end = code.indexOf(']', i + 1);
      if (end < 0) return false;
      i = end;
    } else if (c === '\\' || c === '_' || c === '*') {
      i++;
    }
  }
  return false;
}

function pad2(value: number): string {
  return value < 10 ? '0' + value : String(value);
}

/** A date serial the way it is typed in the locale; '' when the serial is not a date. */
function dateText(serial: number, locale: FormatLocale, date1904: boolean): string {
  const parts = serialToDate(serial, date1904);
  if (parts.year !== parts.year) return '';
  // serialToDate() rounds to the millisecond, which may carry into the next day.
  const timeOnly = !date1904 && serial < 1 && parts.day === 0;
  const hasTime = parts.hour !== 0 || parts.minute !== 0 || parts.second !== 0 || parts.millisecond !== 0;
  let time = '';
  if (hasTime || timeOnly) {
    time = parts.hour + ':' + pad2(parts.minute) + ':' + pad2(parts.second);
    if (parts.millisecond !== 0) {
      const text = String(parts.millisecond);
      time += locale.decimal + zeros(3 - text.length) + text;
    }
  }
  if (timeOnly) return time;
  const order = locale.dateOrder || 'mdy';
  const custom = locale.dateSeparator;
  const separator = custom === '.' || custom === '/' || custom === '-' ? custom : '/';
  let date: string;
  if (order === 'ymd') date = parts.year + '-' + pad2(parts.month) + '-' + pad2(parts.day);
  else if (order === 'dmy') date = pad2(parts.day) + separator + pad2(parts.month) + separator + parts.year;
  else date = parts.month + separator + parts.day + separator + parts.year;
  return hasTime ? date + ' ' + time : date;
}

const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,9}))?)?)?/;

/** Serial for the ISO 8601 text of a cell of kind "date"; -1 when it is not a date. */
function isoSerial(text: string, date1904: boolean): number {
  const match = ISO_DATE.exec(text.length > 64 ? text.substring(0, 64) : text);
  if (!match) return -1;
  const serial = dateSerial(Number(match[1]), Number(match[2]), Number(match[3]), date1904);
  if (serial < 0) return -1;
  const seconds = Number(match[4] || 0) * 3600 + Number(match[5] || 0) * 60 + Number((match[6] || '0') + '.' + (match[7] || '0'));
  return serial + seconds / 86400;
}

/** The text the cell editor shows for a cell: what the user would type to get this value. */
export function cellEditText(cell: Cell | undefined, numFmt: string, options?: InputOptions): string {
  if (cell === undefined || cell === null) return '';
  if (cell.formula && typeof cell.formula.text === 'string' && cell.formula.text.length > 0) {
    return '=' + formulaToLocal(cell.formula.text, resolveLocale(options));
  }
  const value = cell.value;
  if (value === null || value === undefined) return '';
  const locale = resolveLocale(options);
  const date1904 = options !== undefined && options !== null && options.date1904 === true;
  if (typeof value === 'boolean') return value ? locale.trueText : locale.falseText;
  if (typeof value === 'number') {
    if (isDateFormat(numFmt)) {
      const text = dateText(value, locale, date1904);
      if (text !== '') return text;
    } else if (isPercentFormat(numFmt)) {
      return numberText(value, 2, locale) + '%';
    }
    return numberText(value, 0, locale);
  }
  if (typeof value !== 'string') return String(value);
  if (cell.kind === 'error') return value;
  if (cell.kind === 'date') {
    const serial = isoSerial(value, date1904);
    const text = serial < 0 ? '' : dateText(serial, locale, date1904);
    return text !== '' ? text : value;
  }
  // Text that would be read back as something else ("123", "=x", "TRUE", "15.03.2024",
  // or text that itself starts with an apostrophe) needs the apostrophe that keeps it text.
  const parsed = parseCellInput(value, options);
  if (parsed.kind === 'string' && parsed.value === value && parsed.formula === undefined) return value;
  return "'" + value;
}
