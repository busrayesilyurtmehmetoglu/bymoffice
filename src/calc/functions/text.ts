// Text functions: joining and cutting, searching and replacing, case, and the
// conversions between text and numbers.
//
// Lengths and positions count UTF-16 code units, as Excel does. UPPER, LOWER
// and PROPER use the language-neutral case mapping of JavaScript, except in a
// Turkish locale (CalcContext.locale), where the dotted and the dotless i are
// two letters: "i" and "İ", "ı" and "I". The pairs are written out here;
// toLocaleUpperCase() is not used, since what it does depends on the engine.

import { formatValue } from '../../format/number-format';
import {
  chargeText,
  ERROR_NA,
  ERROR_NUM,
  ERROR_VALUE,
  isError,
  numberSeparators,
  textToNumber,
  toBoolean,
  toText,
  type CalcContext,
  type CalcError,
  type CalcFunction,
  type FunctionSpec,
  type Scalar,
  type Value,
} from '../values';
import {
  clean,
  compilePattern,
  eachScalar,
  ERROR_CALC,
  first,
  foldCase,
  hasWildcard,
  mapScalars,
  MAX_ARRAY_CELLS,
  MAX_TEXT_LENGTH,
  single,
  toArray,
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
function textFunction(fn: (text: string, context: CalcContext) => Scalar): CalcFunction {
  return mapScalars((args, context) => {
    const text = toText(args[0]);
    return typeof text === 'string' ? fn(text, context) : text;
  });
}

// ---------------------------------------------------------------------------
// Case

const DOTTED_CAPITAL_I = '\u0130';
const DOTLESS_SMALL_I = '\u0131';
/** U+0307, the dot written after a letter: with "I" before it, the dotted capital in two characters. */
const COMBINING_DOT = 0x307;

/**
 * Whether letters change case the Turkish way. A locale carries no language tag; the
 * Turkish one is known by its word for TRUE, as it is where formulas are typed.
 * UNSURE: Excel follows the language of the system here; not checked on a Turkish one.
 */
function turkishCase(context: CalcContext): boolean {
  const locale = context.locale;
  return locale !== undefined && locale !== null && locale.trueText === 'DOĞRU';
}

function upperChar(c: string, turkish: boolean): string {
  if (turkish) {
    if (c === 'i') return DOTTED_CAPITAL_I;
    if (c === DOTLESS_SMALL_I) return 'I';
  }
  return c.toUpperCase();
}

function lowerChar(c: string, turkish: boolean): string {
  if (turkish) {
    if (c === 'I') return DOTLESS_SMALL_I;
    if (c === DOTTED_CAPITAL_I) return 'i';
  }
  return c.toLowerCase();
}

function upper(text: string, context: CalcContext): string {
  if (!turkishCase(context) || (text.indexOf('i') < 0 && text.indexOf(DOTLESS_SMALL_I) < 0)) return text.toUpperCase();
  let out = '';
  let from = 0;
  for (let i = 0; i < text.length; i++) {
    const c = text.charAt(i);
    if (c !== 'i' && c !== DOTLESS_SMALL_I) continue;
    out += text.substring(from, i).toUpperCase() + upperChar(c, true);
    // A combining dot written after "i" is the dot the capital already has.
    if (c === 'i' && text.charCodeAt(i + 1) === COMBINING_DOT) i++;
    from = i + 1;
  }
  return out + text.substring(from).toUpperCase();
}

function lower(text: string, context: CalcContext): string {
  if (!turkishCase(context) || (text.indexOf('I') < 0 && text.indexOf(DOTTED_CAPITAL_I) < 0)) return text.toLowerCase();
  let out = '';
  let from = 0;
  for (let i = 0; i < text.length; i++) {
    const c = text.charAt(i);
    if (c !== 'I' && c !== DOTTED_CAPITAL_I) continue;
    // A combining dot written after "I" belongs to it: the two together are one dotted capital.
    const dotted = c === 'I' && text.charCodeAt(i + 1) === COMBINING_DOT;
    out += text.substring(from, i).toLowerCase() + (dotted ? 'i' : lowerChar(c, true));
    if (dotted) i++;
    from = i + 1;
  }
  return out + text.substring(from).toLowerCase();
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

/** A mark written after the letter it belongs to (U+0300 to U+036F): part of that letter, not the end of a word. */
function isCombiningMark(code: number): boolean {
  return code >= 0x300 && code <= 0x36f;
}

function proper(text: string, context: CalcContext): string {
  const turkish = turkishCase(context);
  let out = '';
  let inWord = false;
  for (let i = 0; i < text.length; i++) {
    const c = text.charAt(i);
    if (isCombiningMark(text.charCodeAt(i))) {
      out += c;
      continue;
    }
    if (turkish && (c === 'i' || c === 'I') && text.charCodeAt(i + 1) === COMBINING_DOT) {
      // The dotted letter in two characters: one letter, written as one.
      out += inWord ? 'i' : DOTTED_CAPITAL_I;
      i++;
    } else out += inWord ? lowerChar(c, turkish) : upperChar(c, turkish);
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

/** NUMBERVALUE with both separators known: every thousands separator before the decimal sign is dropped, wherever it stands. */
function readNumberValue(source: string, decimal: string, group: string): number | CalcError {
  let body = '';
  let percents = 0;
  let seenDecimal = false;
  for (let i = 0; i < source.length; i++) {
    const c = source.charAt(i);
    if (c === ' ' || c === '\t' || c === '\n' || c === '\r' || c === '\u00a0') continue;
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
}

const numberValue = mapScalars((args, context) => {
  const source = toText(args[0]);
  if (typeof source !== 'string') return source;
  let decimal = '.';
  let group = ',';
  let given = false;
  if (args.length > 1 && args[1] !== null) {
    const text = toText(args[1]);
    if (typeof text !== 'string') return text;
    if (text === '') return ERROR_VALUE;
    decimal = text.charAt(0);
    given = true;
  }
  if (args.length > 2 && args[2] !== null) {
    const text = toText(args[2]);
    if (typeof text !== 'string') return text;
    if (text === '') return ERROR_VALUE;
    group = text.charAt(0);
    given = true;
  }
  if (decimal === group || source.length > 1024) return ERROR_VALUE;
  const number = readNumberValue(source, decimal, group);
  if (given || !isError(number)) return number;
  // No separator was given and the text is no number the way files write one: as everywhere
  // text is read as a number, the separators of the locale are tried second ("1.234,5"),
  // when its decimal sign is there. Points alone ("1.234.5") stay the error they were.
  const local = numberSeparators(context);
  return local === undefined || source.indexOf(local.decimal) < 0 ? number : readNumberValue(source, local.decimal, local.group);
});

// ---------------------------------------------------------------------------
// TEXTBEFORE, TEXTAFTER and TEXTSPLIT: text cut at delimiters, of which there may be several
// (an array constant, a range). Where two of them are found at the same place the longer one
// is taken (UNSURE: not checked against Excel).

/** Most delimiters one argument may give, those spelled alike counted once. */
const MAX_DELIMITERS = 1024;

/** The delimiters an argument gives: a text, or an array or range of them; none when it is left empty. */
function delimiterList(value: Value | undefined, context: CalcContext): string[] | CalcError {
  if (value === undefined || value === null) return [];
  const list = toArray(value, context);
  if (isError(list)) return list;
  const count = list.rows * list.cols;
  const out: string[] = [];
  const seen = new Set<string>();
  for (let i = 0; i < count; i++) {
    const text = toText(clean(list.data[i]));
    if (typeof text !== 'string') return text;
    if (seen.has(text)) continue;
    if (seen.size >= MAX_DELIMITERS) return ERROR_VALUE;
    seen.add(text);
    out.push(text);
  }
  return out;
}

interface Found {
  start: number;
  end: number;
  /** Which delimiter, by its place in the list the index was made with. */
  which: number;
}

/**
 * Delimiters looked up by their first character: for each first character the delimiters that
 * begin with it, the longest first (and of those as long, the preferred first). Made once for a
 * call, whatever number of texts it is applied to.
 */
class Delimiters {
  readonly byFirst = new Map<number, number[]>();
  /** A delimiter of no characters is among them. */
  readonly empty: boolean;

  constructor(
    readonly texts: readonly string[],
    /** Of two found at the same place and as long, the one this prefers (a row delimiter of TEXTSPLIT). */
    preferred?: (which: number) => boolean,
  ) {
    this.empty = texts.indexOf('') >= 0;
    for (let k = 0; k < texts.length; k++) {
      if (texts[k] === '') continue;
      const code = texts[k].charCodeAt(0);
      const list = this.byFirst.get(code);
      if (list) list.push(k);
      else this.byFirst.set(code, [k]);
    }
    const rank = (k: number) => (preferred && preferred(k) ? 0 : 1);
    this.byFirst.forEach((list) => list.sort((a, b) => texts[b].length - texts[a].length || rank(a) - rank(b) || a - b));
  }
}

/** Units of work counted before they are handed to the budget of the calculation. */
const WORK_BATCH = 256;

/**
 * Finds delimiters in a text one after another, forwards or backwards. One delimiter is searched
 * for by the engine's own search; several are found by walking the text once, looking at each
 * place only at the delimiters that begin with its character, so the work grows with the length
 * of the text and with how many delimiters share a first character, not with the number of
 * delimiters times the number of pieces. What it does is counted in the budget of the
 * calculation (CalcContext.tick), so that a call with much to do is cut off by it.
 */
class Matcher {
  private work = 0;

  constructor(
    private readonly subject: string,
    private readonly delimiters: Delimiters,
    private readonly context: CalcContext,
  ) {}

  private spend(units: number): void {
    this.work += units;
    if (this.work >= WORK_BATCH) {
      if (this.context.tick !== undefined) this.context.tick(this.work);
      this.work = 0;
    }
  }

  /** The first delimiter that starts at `from` or after it. */
  next(from: number): Found | undefined {
    const subject = this.subject;
    const texts = this.delimiters.texts;
    const n = subject.length;
    if (texts.length === 1) {
      const at = subject.indexOf(texts[0], from);
      this.spend(1 + (((at < 0 ? n : at) - from) >> 6));
      return at < 0 ? undefined : { start: at, end: at + texts[0].length, which: 0 };
    }
    const byFirst = this.delimiters.byFirst;
    let looked = 0;
    for (let i = from; i < n; i++) {
      looked++;
      const list = byFirst.get(subject.charCodeAt(i));
      if (list !== undefined) {
        for (let j = 0; j < list.length; j++) {
          const delimiter = texts[list[j]];
          looked += 1 + (delimiter.length >> 4);
          if (i + delimiter.length <= n && subject.startsWith(delimiter, i)) {
            this.spend(looked);
            return { start: i, end: i + delimiter.length, which: list[j] };
          }
        }
      }
      if (looked >= WORK_BATCH) {
        this.spend(looked);
        looked = 0;
      }
    }
    this.spend(looked);
    return undefined;
  }

  /** The last delimiter that ends at `limit` or before it. */
  previous(limit: number): Found | undefined {
    const subject = this.subject;
    const texts = this.delimiters.texts;
    if (texts.length === 1) {
      const length = texts[0].length;
      const at = limit - length < 0 ? -1 : subject.lastIndexOf(texts[0], limit - length);
      this.spend(1 + ((limit - (at < 0 ? 0 : at)) >> 6));
      return at < 0 ? undefined : { start: at, end: at + length, which: 0 };
    }
    const byFirst = this.delimiters.byFirst;
    let looked = 0;
    for (let i = limit - 1; i >= 0; i--) {
      looked++;
      const list = byFirst.get(subject.charCodeAt(i));
      if (list !== undefined) {
        for (let j = 0; j < list.length; j++) {
          const delimiter = texts[list[j]];
          looked += 1 + (delimiter.length >> 4);
          if (i + delimiter.length <= limit && subject.startsWith(delimiter, i)) {
            this.spend(looked);
            return { start: i, end: i + delimiter.length, which: list[j] };
          }
        }
      }
      if (looked >= WORK_BATCH) {
        this.spend(looked);
        looked = 0;
      }
    }
    this.spend(looked);
    return undefined;
  }

  /** Hands what is left of the work counted to the budget. */
  done(): void {
    if (this.work > 0 && this.context.tick !== undefined) this.context.tick(this.work);
    this.work = 0;
  }
}

/** Where the n-th delimiter is: its start and end, or undefined. Negative n counts from the end. */
function locate(
  subject: string,
  delimiters: Delimiters,
  instance: number,
  matchEnd: boolean,
  context: CalcContext,
): { start: number; end: number } | undefined {
  const length = subject.length;
  // A delimiter of no characters is found at once.
  if (delimiters.empty) return instance > 0 ? { start: 0, end: 0 } : { start: length, end: length };
  const matcher = new Matcher(subject, delimiters, context);
  let seen = 0;
  let found: Found | undefined;
  if (instance > 0) {
    let at = 0;
    for (;;) {
      found = matcher.next(at);
      if (!found || ++seen === instance) break;
      at = found.end;
    }
    matcher.done();
    if (found) return found;
    // With match_end the end of the text stands in for one more delimiter.
    return matchEnd && seen === instance - 1 ? { start: length, end: length } : undefined;
  }
  let limit = length;
  for (;;) {
    found = matcher.previous(limit);
    if (!found || ++seen === -instance) break;
    limit = found.start;
  }
  matcher.done();
  if (found) return found;
  return matchEnd && seen === -instance - 1 ? { start: 0, end: 0 } : undefined;
}

/** TEXTBEFORE and TEXTAFTER: applied to each text (and each instance, mode ...) of arrays; the delimiters are one list. */
function around(after: boolean): CalcFunction {
  return (args, context) => {
    // A delimiter left empty is the text of no characters, as an empty cell is. UNSURE: Excel not checked.
    const given = args[1] === null ? [''] : delimiterList(args[1], context);
    if (isError(given)) return given;
    if (given.length === 0) return ERROR_VALUE;
    const exact = new Delimiters(given);
    let folded: Delimiters | undefined;
    const rest = args.slice();
    rest[1] = null;
    // Made for each call, so that what it closes over belongs to this call alone.
    const lifted = mapScalars((values, context) => {
      const subject = toText(values[0]);
      if (typeof subject !== 'string') return subject;
      const instance = values.length > 2 && values[2] !== null ? integer(values, 2, 1, context) : 1;
      if (typeof instance !== 'number') return instance;
      const mode = values.length > 3 ? integer(values, 3, 0, context) : 0;
      if (typeof mode !== 'number') return mode;
      const matchEnd = values.length > 4 ? integer(values, 4, 0, context) : 0;
      if (typeof matchEnd !== 'number') return matchEnd;
      if (instance === 0 || Math.abs(instance) > subject.length || mode < 0 || mode > 1) return ERROR_VALUE;
      const insensitive = mode === 1;
      if (insensitive && !folded) folded = new Delimiters(given.map(foldCase));
      const place = locate(insensitive ? foldCase(subject) : subject, insensitive ? (folded as Delimiters) : exact, instance, matchEnd !== 0, context);
      if (place === undefined) return values.length > 5 ? values[5] : ERROR_NA;
      return after ? subject.substring(place.end) : subject.substring(0, place.start);
    });
    return lifted(rest, context);
  };
}

/**
 * The rows of a text split at row delimiters, each split at column delimiters; #VALUE! when the
 * pieces are too much text. With `firstOnly` the search stops at the first piece kept: a single
 * row of it (no row when no piece is kept).
 */
function splitText(
  subject: string,
  delimiters: Delimiters,
  rowDelimiters: number,
  insensitive: boolean,
  ignoreEmpty: boolean,
  firstOnly: boolean,
  context: CalcContext,
): string[][] | CalcError {
  // The row delimiters come first in the list.
  const matcher = new Matcher(insensitive ? foldCase(subject) : subject, delimiters, context);
  const rows: string[][] = [];
  let row: string[] = [];
  let at = 0;
  for (;;) {
    const found = matcher.next(at);
    const piece = subject.substring(at, found ? found.start : subject.length);
    if (!(ignoreEmpty && piece === '')) {
      if (!chargeText(piece)) {
        matcher.done();
        return ERROR_VALUE;
      }
      row.push(piece);
      if (firstOnly) {
        rows.push(row);
        break;
      }
    }
    if (!found || found.which < rowDelimiters) {
      if (!(ignoreEmpty && row.length === 0)) rows.push(row);
      row = [];
    }
    if (!found) break;
    // Every delimiter has characters: the search moves on.
    at = found.end;
  }
  matcher.done();
  return rows;
}

/** TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with]). */
const textSplit: CalcFunction = (args, context) => {
  const columns = delimiterList(args[1], context);
  if (isError(columns)) return columns;
  const lines = delimiterList(args.length > 2 ? args[2] : undefined, context);
  if (isError(lines)) return lines;
  let ignoreEmpty: boolean | CalcError = false;
  if (args.length > 3 && args[3] !== null) {
    const value = single(args[3], context);
    ignoreEmpty = value === undefined ? ERROR_VALUE : toBoolean(value);
  }
  if (typeof ignoreEmpty !== 'boolean') return ignoreEmpty;
  let mode = 0;
  if (args.length > 4 && args[4] !== null) {
    const value = single(args[4], context);
    const number = value === undefined ? ERROR_VALUE : integer([value], 0, 0, context);
    if (typeof number !== 'number') return number;
    mode = number;
  }
  if (mode !== 0 && mode !== 1) return ERROR_VALUE;
  const pad: Scalar = args.length > 5 ? first(args[5], context) : ERROR_NA;
  // A delimiter of no characters splits nothing. UNSURE: what Excel does with one was not checked.
  const fold = (text: string) => (mode === 1 ? foldCase(text) : text);
  const rowList = lines.filter((text) => text !== '').map(fold);
  const columnList = columns.filter((text) => text !== '').map(fold);
  if (rowList.length + columnList.length === 0) return ERROR_VALUE;
  const rowCount = rowList.length;
  const delimiters = new Delimiters(rowList.concat(columnList), (which) => which < rowCount);
  const ignoring = ignoreEmpty;
  const split = (value: Scalar, firstOnly: boolean): string[][] | CalcError => {
    const subject = toText(value);
    if (typeof subject !== 'string') return subject;
    return splitText(subject, delimiters, rowCount, mode === 1, ignoring, firstOnly, context);
  };
  const one = single(args[0], context);
  if (one !== undefined) {
    const rows = split(one, false);
    if (isError(rows)) return rows;
    if (rows.length === 0) return ERROR_CALC;
    let cols = 0;
    for (const row of rows) if (row.length > cols) cols = row.length;
    if (!(rows.length * cols <= MAX_ARRAY_CELLS)) return ERROR_NUM;
    const data: Scalar[] = new Array(rows.length * cols);
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      for (let c = 0; c < cols; c++) data[r * cols + c] = c < row.length ? row[c] : pad;
    }
    return { rows: rows.length, cols, data };
  }
  // Several texts: each gives the first of its pieces, as an array of arrays does in Excel. UNSURE.
  const texts = toArray(args[0], context);
  if (isError(texts)) return texts;
  const size = texts.rows * texts.cols;
  const data: Scalar[] = new Array(size);
  for (let i = 0; i < size; i++) {
    const rows = split(clean(texts.data[i]), true);
    data[i] = isError(rows) ? rows : rows.length === 0 ? ERROR_CALC : rows[0].length > 0 ? rows[0][0] : pad;
  }
  return { rows: texts.rows, cols: texts.cols, data };
};

const delimiterArg = (index: number) => index === 1;
const notDelimiter = (index: number) => index !== 1;

export const TEXT_FUNCTIONS: { [name: string]: FunctionSpec } = {
  CONCATENATE: spec(concatenate, 1),
  CONCAT: spec(concat, 1),
  TEXTJOIN: spec(textJoin, 3),
  LEFT: spec(side(false), 1, 2),
  RIGHT: spec(side(true), 1, 2),
  MID: spec(mid, 3, 3),
  LEN: spec(textFunction((subject) => subject.length), 1, 1),
  LOWER: spec(textFunction(lower), 1, 1),
  UPPER: spec(textFunction(upper), 1, 1),
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
  TEXTBEFORE: { call: around(false), min: 2, max: 6, arrayArgs: delimiterArg, scalarArgs: notDelimiter },
  TEXTAFTER: { call: around(true), min: 2, max: 6, arrayArgs: delimiterArg, scalarArgs: notDelimiter },
  TEXTSPLIT: { call: textSplit, min: 2, max: 6, arrayArgs: (index) => index <= 2, scalarArgs: (index) => index >= 3 },
};
