// Text functions: joining and cutting, searching and replacing, case, and the
// conversions between text and numbers.
//
// Lengths and positions count UTF-16 code units, as Excel does. UPPER, LOWER
// and PROPER use the language-neutral case mapping of JavaScript; Turkish
// dotted and dotless i are not treated specially.

import { formatValue } from '../../format/number-format';
import {
  ERROR_NA,
  ERROR_VALUE,
  isError,
  textToNumber,
  toBoolean,
  toText,
  type CalcContext,
  type CalcError,
  type CalcFunction,
  type FunctionSpec,
  type Scalar,
} from '../values';
import {
  compilePattern,
  eachScalar,
  foldCase,
  hasWildcard,
  mapScalars,
  MAX_TEXT_LENGTH,
  parseNumberText,
  patternSearch,
  roundTo,
  scalarNumber,
  spec,
} from './helpers';

function checkText(text: string): string | CalcError {
  return text.length > MAX_TEXT_LENGTH ? ERROR_VALUE : text;
}

/** An optional numeric argument, truncated to an integer. */
function integer(args: Scalar[], index: number, fallback: number, context: CalcContext): number | CalcError {
  if (index >= args.length) return fallback;
  const value = scalarNumber(args[index], context);
  if (typeof value !== 'number') return value;
  if (value !== value) return ERROR_VALUE;
  return value < 0 ? Math.ceil(value) : Math.floor(value);
}

function flag(args: Scalar[], index: number, fallback: boolean): boolean | CalcError {
  // Given but empty is FALSE, as an empty cell is.
  if (index >= args.length) return fallback;
  return toBoolean(args[index]);
}

/** A function of one text argument. */
function textFunction(fn: (text: string) => Scalar): CalcFunction {
  return mapScalars((args) => {
    const text = toText(args[0]);
    return typeof text === 'string' ? fn(text) : text;
  });
}

const concatenate = mapScalars((args) => {
  let out = '';
  for (let i = 0; i < args.length; i++) {
    const text = toText(args[i]);
    if (typeof text !== 'string') return text;
    out += text;
    if (out.length > MAX_TEXT_LENGTH) return ERROR_VALUE;
  }
  return out;
});

const concat: CalcFunction = (args, context) => {
  const parts: string[] = [];
  let length = 0;
  let error = undefined as CalcError | undefined;
  eachScalar(args, context, (value) => {
    if (error !== undefined || length > MAX_TEXT_LENGTH) return;
    const text = toText(value);
    if (typeof text !== 'string') {
      error = text;
      return;
    }
    length += text.length;
    parts.push(text);
  });
  if (error !== undefined) return error;
  return length > MAX_TEXT_LENGTH ? ERROR_VALUE : parts.join('');
};

const textJoin: CalcFunction = (args, context) => {
  const delimiters: string[] = [];
  let error = undefined as CalcError | undefined;
  eachScalar([args[0]], context, (value) => {
    const text = toText(value);
    if (typeof text === 'string') delimiters.push(text);
    else if (error === undefined) error = text;
  });
  if (error !== undefined) return error;
  if (delimiters.length === 0) delimiters.push('');
  let skipEmpty: boolean | CalcError = true;
  let seen = false;
  eachScalar([args[1]], context, (value) => {
    if (!seen) skipEmpty = toBoolean(value);
    seen = true;
  });
  if (typeof skipEmpty !== 'boolean') return skipEmpty;
  const skip = skipEmpty;
  let out = '';
  let count = 0;
  eachScalar(args.slice(2), context, (value) => {
    if (error !== undefined || out.length > MAX_TEXT_LENGTH) return;
    const text = toText(value);
    if (typeof text !== 'string') {
      error = text;
      return;
    }
    if (skip && text === '') return;
    // Several delimiters are used in turn.
    if (count > 0) out += delimiters[(count - 1) % delimiters.length];
    out += text;
    count++;
  });
  if (error !== undefined) return error;
  return checkText(out);
};

function side(right: boolean): CalcFunction {
  return mapScalars((args, context) => {
    const text = toText(args[0]);
    if (typeof text !== 'string') return text;
    const count = integer(args, 1, 1, context);
    if (typeof count !== 'number') return count;
    if (count < 0) return ERROR_VALUE;
    if (count >= text.length) return text;
    return right ? text.substring(text.length - count) : text.substring(0, count);
  });
}

const mid = mapScalars((args, context) => {
  const text = toText(args[0]);
  if (typeof text !== 'string') return text;
  const start = integer(args, 1, 1, context);
  if (typeof start !== 'number') return start;
  const count = integer(args, 2, 0, context);
  if (typeof count !== 'number') return count;
  if (start < 1 || count < 0) return ERROR_VALUE;
  if (start > text.length) return '';
  return text.substring(start - 1, Math.min(text.length, start - 1 + count));
});

function isLetter(c: string): boolean {
  return c.toLowerCase() !== c.toUpperCase();
}

function proper(text: string): string {
  let out = '';
  let inWord = false;
  for (let i = 0; i < text.length; i++) {
    const c = text.charAt(i);
    out += inWord ? c.toLowerCase() : c.toUpperCase();
    inWord = isLetter(c);
  }
  return out;
}

/** Removes leading and trailing spaces and leaves one space between words. Only U+0020 counts. */
function trimSpaces(text: string): string {
  let out = '';
  let pending = false;
  for (let i = 0; i < text.length; i++) {
    const c = text.charAt(i);
    if (c === ' ') {
      pending = out.length > 0;
    } else {
      if (pending) out += ' ';
      pending = false;
      out += c;
    }
  }
  return out;
}

function cleanControls(text: string): string {
  let out = '';
  for (let i = 0; i < text.length; i++) {
    if (text.charCodeAt(i) >= 32) out += text.charAt(i);
  }
  return out;
}

function finder(search: boolean): CalcFunction {
  return mapScalars((args, context) => {
    const needle = toText(args[0]);
    if (typeof needle !== 'string') return needle;
    const within = toText(args[1]);
    if (typeof within !== 'string') return within;
    const start = integer(args, 2, 1, context);
    if (typeof start !== 'number') return start;
    if (start < 1 || start > within.length + (needle === '' ? 1 : 0)) return ERROR_VALUE;
    let at: number;
    if (!search) at = within.indexOf(needle, start - 1);
    else {
      const folded = foldCase(needle);
      at = hasWildcard(folded)
        ? patternSearch(compilePattern(folded), foldCase(within), start - 1)
        : foldCase(within).indexOf(folded, start - 1);
    }
    return at < 0 ? ERROR_VALUE : at + 1;
  });
}

const substitute = mapScalars((args, context) => {
  const text = toText(args[0]);
  if (typeof text !== 'string') return text;
  const from = toText(args[1]);
  if (typeof from !== 'string') return from;
  const to = toText(args[2]);
  if (typeof to !== 'string') return to;
  let instance = 0;
  if (args.length > 3) {
    const n = integer(args, 3, 0, context);
    if (typeof n !== 'number') return n;
    if (n < 1) return ERROR_VALUE;
    instance = n;
  }
  if (from === '') return text;
  let out = '';
  let at = 0;
  let seen = 0;
  for (;;) {
    const found = text.indexOf(from, at);
    if (found < 0) break;
    seen++;
    if (instance === 0 || seen === instance) {
      out += text.substring(at, found) + to;
      if (out.length > MAX_TEXT_LENGTH) return ERROR_VALUE;
    } else out += text.substring(at, found + from.length);
    at = found + from.length;
    if (seen === instance) break;
  }
  return checkText(out + text.substring(at));
});

const replace = mapScalars((args, context) => {
  const text = toText(args[0]);
  if (typeof text !== 'string') return text;
  const start = integer(args, 1, 1, context);
  if (typeof start !== 'number') return start;
  const count = integer(args, 2, 0, context);
  if (typeof count !== 'number') return count;
  const insert = toText(args[3]);
  if (typeof insert !== 'string') return insert;
  if (start < 1 || count < 0) return ERROR_VALUE;
  const from = Math.min(start - 1, text.length);
  return checkText(text.substring(0, from) + insert + text.substring(Math.min(text.length, from + count)));
});

const repeat = mapScalars((args, context) => {
  const text = toText(args[0]);
  if (typeof text !== 'string') return text;
  const count = integer(args, 1, 0, context);
  if (typeof count !== 'number') return count;
  if (count < 0 || text.length * count > MAX_TEXT_LENGTH) return ERROR_VALUE;
  if (count === 0 || text === '') return '';
  let out = '';
  for (let i = 0; i < count; i++) out += text;
  return out;
});

const exact = mapScalars((args) => {
  const a = toText(args[0]);
  if (typeof a !== 'string') return a;
  const b = toText(args[1]);
  if (typeof b !== 'string') return b;
  return a === b;
});

const value = mapScalars((args, context) => {
  const arg = args[0];
  if (typeof arg === 'number' || isError(arg)) return arg;
  if (arg === null) return 0;
  if (typeof arg !== 'string') return ERROR_VALUE;
  const number = parseNumberText(arg, context);
  return number === undefined ? ERROR_VALUE : number;
});

const text = mapScalars((args, context) => {
  let subject = args[0];
  if (isError(subject)) return subject;
  const code = toText(args[1]);
  if (typeof code !== 'string') return code;
  if (subject === null) subject = 0;
  if (typeof subject === 'boolean') return subject ? 'TRUE' : 'FALSE';
  if (typeof subject === 'string') {
    const number = parseNumberText(subject, context);
    if (number !== undefined) subject = number;
  }
  if (code === '') return typeof subject === 'string' ? subject : '';
  const result = formatValue(subject as number | string, code, { date1904: context.date1904 === true });
  if (result.overflow === true) return ERROR_VALUE;
  return checkText(result.text);
});

/** Digits of a non-negative number with a fixed count of decimals, optionally grouped in thousands. */
function fixedDigits(abs: number, decimals: number, group: boolean): string | undefined {
  // toFixed() gives up on fixed notation at 1e21 and takes at most 100 decimals.
  if (!(abs < 1e21)) return undefined;
  const places = Math.max(0, decimals);
  let digits = abs.toFixed(Math.min(places, 100));
  for (let i = 100; i < places; i++) digits += '0';
  if (!group) return digits;
  const point = digits.indexOf('.');
  const whole = point < 0 ? digits : digits.substring(0, point);
  let out = '';
  for (let i = 0; i < whole.length; i++) {
    if (i > 0 && (whole.length - i) % 3 === 0) out += ',';
    out += whole.charAt(i);
  }
  return point < 0 ? out : out + digits.substring(point);
}

const fixed = mapScalars((args, context) => {
  const number = scalarNumber(args[0], context);
  if (typeof number !== 'number') return number;
  // Left out, two decimals; left empty ("FIXED(x,,TRUE)"), none.
  const decimals = integer(args, 1, 2, context);
  if (typeof decimals !== 'number') return decimals;
  const plain = flag(args, 2, false);
  if (typeof plain !== 'boolean') return plain;
  if (decimals > 127) return ERROR_VALUE;
  const rounded = roundTo(number, decimals, 0);
  const digits = fixedDigits(Math.abs(rounded), decimals, !plain);
  if (digits === undefined) return ERROR_VALUE;
  // A negative number that rounds to zero keeps its sign, as it does under a number format: "-0.00".
  return (number < 0 ? '-' : '') + digits;
});

const dollar = mapScalars((args, context) => {
  const number = scalarNumber(args[0], context);
  if (typeof number !== 'number') return number;
  const decimals = integer(args, 1, 2, context);
  if (typeof decimals !== 'number') return decimals;
  if (decimals > 127) return ERROR_VALUE;
  const rounded = roundTo(number, decimals, 0);
  const digits = fixedDigits(Math.abs(rounded), decimals, true);
  if (digits === undefined) return ERROR_VALUE;
  // The currency format "$#,##0.00_);($#,##0.00)": negatives go in parentheses, also when they round to zero.
  return number < 0 ? '($' + digits + ')' : '$' + digits;
});

// Windows-1252 code points 128 to 159, the part that differs from Latin-1.
const CP1252 = [
  0x20ac, 0x81, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021, 0x02c6, 0x2030, 0x0160, 0x2039, 0x0152, 0x8d, 0x017d,
  0x8f, 0x90, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x02dc, 0x2122, 0x0161, 0x203a, 0x0153, 0x9d,
  0x017e, 0x0178,
];

const char = mapScalars((args, context) => {
  const code = integer(args, 0, 0, context);
  if (typeof code !== 'number') return code;
  if (code < 1 || code > 255) return ERROR_VALUE;
  return String.fromCharCode(code >= 128 && code < 160 ? CP1252[code - 128] : code);
});

const code = textFunction((subject) => {
  if (subject === '') return ERROR_VALUE;
  const unit = subject.charCodeAt(0);
  if (unit < 128 || (unit >= 160 && unit < 256)) return unit;
  const index = CP1252.indexOf(unit);
  // Characters outside the code page read as a question mark.
  return index >= 0 ? index + 128 : 63;
});

const unichar = mapScalars((args, context) => {
  const point = integer(args, 0, 0, context);
  if (typeof point !== 'number') return point;
  if (point < 1 || point > 0x10ffff) return ERROR_VALUE;
  if (point >= 0xd800 && point <= 0xdfff) return ERROR_NA;
  return String.fromCodePoint(point);
});

const unicode = textFunction((subject) => {
  const point = subject.codePointAt(0);
  return point === undefined ? ERROR_VALUE : point;
});

const numberValue = mapScalars((args) => {
  const source = toText(args[0]);
  if (typeof source !== 'string') return source;
  let decimal = '.';
  let group = ',';
  if (args.length > 1 && args[1] !== null) {
    const given = toText(args[1]);
    if (typeof given !== 'string') return given;
    if (given === '') return ERROR_VALUE;
    decimal = given.charAt(0);
  }
  if (args.length > 2 && args[2] !== null) {
    const given = toText(args[2]);
    if (typeof given !== 'string') return given;
    if (given === '') return ERROR_VALUE;
    group = given.charAt(0);
  }
  if (decimal === group || source.length > 1024) return ERROR_VALUE;
  let body = '';
  let percents = 0;
  let seenDecimal = false;
  for (let i = 0; i < source.length; i++) {
    const c = source.charAt(i);
    if (c === ' ' || c === '\t' || c === '\n' || c === '\r' || c === ' ') continue;
    if (c === '%') {
      percents++;
      continue;
    }
    // Percent signs only close the number.
    if (percents > 0) return ERROR_VALUE;
    if (c === decimal) {
      if (seenDecimal) return ERROR_VALUE;
      seenDecimal = true;
      body += '.';
    } else if (c === group) {
      if (seenDecimal) return ERROR_VALUE;
    } else body += c;
  }
  if (body === '') return percents > 0 ? ERROR_VALUE : 0;
  // What is left must be a plain number; the spaces and percent signs textToNumber() would allow are gone already.
  if (textToNumber(body) === undefined) return ERROR_VALUE;
  let result = Number(body);
  for (let i = 0; i < percents && result !== 0; i++) result /= 100;
  return isFinite(result) ? result : ERROR_VALUE;
});

/** Where the n-th delimiter is: its start and end, or undefined. Negative n counts from the end. */
function locate(
  subject: string,
  delimiter: string,
  instance: number,
  matchEnd: boolean,
): { start: number; end: number } | undefined {
  const length = subject.length;
  if (delimiter === '') return instance > 0 ? { start: 0, end: 0 } : { start: length, end: length };
  let seen = 0;
  if (instance > 0) {
    let at = 0;
    for (;;) {
      const found = subject.indexOf(delimiter, at);
      if (found < 0) break;
      if (++seen === instance) return { start: found, end: found + delimiter.length };
      at = found + delimiter.length;
    }
    // With match_end the end of the text stands in for one more delimiter.
    return matchEnd && seen === instance - 1 ? { start: length, end: length } : undefined;
  }
  let at = length - delimiter.length;
  while (at >= 0) {
    const found = subject.lastIndexOf(delimiter, at);
    if (found < 0) break;
    if (++seen === -instance) return { start: found, end: found + delimiter.length };
    at = found - delimiter.length;
  }
  return matchEnd && seen === -instance - 1 ? { start: 0, end: 0 } : undefined;
}

function around(after: boolean): CalcFunction {
  return mapScalars((args, context) => {
    const subject = toText(args[0]);
    if (typeof subject !== 'string') return subject;
    const delimiter = toText(args[1]);
    if (typeof delimiter !== 'string') return delimiter;
    const instance = args.length > 2 && args[2] !== null ? integer(args, 2, 1, context) : 1;
    if (typeof instance !== 'number') return instance;
    const mode = args.length > 3 ? integer(args, 3, 0, context) : 0;
    if (typeof mode !== 'number') return mode;
    const matchEnd = args.length > 4 ? integer(args, 4, 0, context) : 0;
    if (typeof matchEnd !== 'number') return matchEnd;
    if (instance === 0 || Math.abs(instance) > subject.length || mode < 0 || mode > 1) return ERROR_VALUE;
    const insensitive = mode === 1;
    const place = locate(
      insensitive ? foldCase(subject) : subject,
      insensitive ? foldCase(delimiter) : delimiter,
      instance,
      matchEnd !== 0,
    );
    if (place === undefined) return args.length > 5 ? args[5] : ERROR_NA;
    return after ? subject.substring(place.end) : subject.substring(0, place.start);
  });
}

export const TEXT_FUNCTIONS: { [name: string]: FunctionSpec } = {
  CONCATENATE: spec(concatenate, 1),
  CONCAT: spec(concat, 1),
  TEXTJOIN: spec(textJoin, 3),
  LEFT: spec(side(false), 1, 2),
  RIGHT: spec(side(true), 1, 2),
  MID: spec(mid, 3, 3),
  LEN: spec(textFunction((subject) => subject.length), 1, 1),
  LOWER: spec(textFunction((subject) => subject.toLowerCase()), 1, 1),
  UPPER: spec(textFunction((subject) => subject.toUpperCase()), 1, 1),
  PROPER: spec(textFunction(proper), 1, 1),
  TRIM: spec(textFunction(trimSpaces), 1, 1),
  CLEAN: spec(textFunction(cleanControls), 1, 1),
  FIND: spec(finder(false), 2, 3),
  SEARCH: spec(finder(true), 2, 3),
  SUBSTITUTE: spec(substitute, 3, 4),
  REPLACE: spec(replace, 4, 4),
  REPT: spec(repeat, 2, 2),
  EXACT: spec(exact, 2, 2),
  VALUE: spec(value, 1, 1),
  TEXT: spec(text, 2, 2),
  FIXED: spec(fixed, 1, 3),
  DOLLAR: spec(dollar, 1, 2),
  CHAR: spec(char, 1, 1),
  CODE: spec(code, 1, 1),
  UNICHAR: spec(unichar, 1, 1),
  UNICODE: spec(unicode, 1, 1),
  NUMBERVALUE: spec(numberValue, 1, 3),
  TEXTBEFORE: spec(around(false), 2, 6),
  TEXTAFTER: spec(around(true), 2, 6),
};
