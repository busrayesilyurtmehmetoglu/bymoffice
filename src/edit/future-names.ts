// The prefixes a file gives to what is newer than its format.
//
// A file stores a function added to Excel after 2007 with "_xlfn." in front of its name
// ("_xlfn.XLOOKUP("), two of them with "_xlfn._xlws." ("_xlfn._xlws.FILTER("), the names of
// LET and the parameters of LAMBDA with "_xlpm.", "A1#" as "_xlfn.ANCHORARRAY(A1)" and "@A1:A3"
// as "_xlfn.SINGLE(A1:A3)". A formula is shown and typed without any of that; the model holds
// it the way files do. Like the rest of ./formula-locale.ts, nothing here parses a formula: text
// it cannot make sense of comes out as it went in, and the parser has the last word.

import { parseFormula } from '../calc/parse';

/**
 * Functions a file stores with "_xlfn.". UNSURE: taken from the list of Excel's "future
 * functions" as remembered, for the functions this engine has; not checked against a file
 * Excel wrote for every one of them.
 */
const FUTURE = new Set([
  'ACOT', 'ACOTH', 'AGGREGATE', 'BITAND', 'BITOR', 'BITXOR', 'BYCOL', 'BYROW', 'CEILING.MATH', 'CEILING.PRECISE', 'CONCAT',
  'COT', 'COTH', 'COVARIANCE.P', 'COVARIANCE.S', 'CSC', 'CSCH', 'DAYS', 'FLOOR.MATH', 'FLOOR.PRECISE', 'FORECAST.LINEAR',
  'FORMULATEXT', 'IFNA', 'IFS', 'ISFORMULA', 'ISO.CEILING', 'ISOMITTED', 'ISOWEEKNUM', 'LAMBDA', 'LET', 'MAKEARRAY', 'MAP',
  'MAXIFS', 'MINIFS', 'MODE.SNGL', 'NETWORKDAYS.INTL', 'NORM.DIST', 'NORM.INV', 'NORM.S.DIST', 'NORM.S.INV', 'NUMBERVALUE',
  'PERCENTILE.EXC', 'PERCENTILE.INC', 'QUARTILE.EXC', 'QUARTILE.INC', 'RANDARRAY', 'RANK.AVG', 'RANK.EQ', 'REDUCE', 'SCAN',
  'SEC', 'SECH', 'SEQUENCE', 'SHEET', 'SHEETS', 'SORTBY', 'STDEV.P', 'STDEV.S', 'SWITCH', 'TEXTAFTER', 'TEXTBEFORE', 'TEXTJOIN',
  'UNICHAR', 'UNICODE', 'UNIQUE', 'VAR.P', 'VAR.S', 'WORKDAY.INTL', 'XLOOKUP', 'XMATCH', 'XOR', 'ANCHORARRAY', 'SINGLE',
]);

/** Functions a file stores with "_xlfn._xlws.". */
const WORKSHEET_ONLY = new Set(['FILTER', 'SORT']);

/** Longer text is not a formula and is left alone. */
const MAX_LENGTH = 32768;

/** Characters of names and references, as the formula parser has them. */
function isWordChar(c: number): boolean {
  if (c >= 128) return true;
  if (c >= 97) return c <= 122;
  if (c >= 65) return c <= 90 || c === 95 || c === 92;
  if (c >= 48) return c <= 57 || c === 63;
  return c === 36 || c === 46;
}

const isSpace = (c: number) => c === 32 || c === 9 || c === 10 || c === 13;

/** End of the quoted run that starts at `i`; the length of the text when it never ends. */
function quotedEnd(s: string, i: number): number {
  const quote = s.charCodeAt(i);
  for (let j = i + 1; j < s.length; j++) {
    if (s.charCodeAt(j) !== quote) continue;
    if (s.charCodeAt(j + 1) === quote) j++;
    else return j + 1;
  }
  return s.length;
}

/** End of the bracketed run that starts at `i`, brackets inside it included. */
function bracketEnd(s: string, i: number): number {
  let depth = 0;
  for (let j = i; j < s.length; j++) {
    const c = s.charCodeAt(j);
    if (c === 91) depth++;
    else if (c === 93 && --depth === 0) return j + 1;
  }
  return s.length;
}

function wordEnd(s: string, i: number): number {
  let j = i;
  while (j < s.length && isWordChar(s.charCodeAt(j))) j++;
  return j;
}

const PREFIXES = ['_XLFN.', '_XLWS.', '_XLPM.'];

/** Length of the prefixes a word starts with ("_xlfn._xlws.FILTER" has two). */
function prefixLength(word: string): number {
  let at = 0;
  for (let guard = 0; guard < 3; guard++) {
    const head = word.substr(at, 6).toUpperCase();
    if (PREFIXES.indexOf(head) < 0) break;
    at += 6;
  }
  return at;
}

/** Whether a word spells a cell or a range corner ("A1", "$B$7", "XFD1048576"): never a name of LET. */
function isCellWord(word: string): boolean {
  return /^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}$/.test(word);
}

/** A call being read: which function, which argument, and the names it binds so far. */
interface Frame {
  name: string;
  arg: number;
  /** Characters of the current argument that are not spaces, and whether it is one plain word so far. */
  start: number;
  bound: Set<string>;
}

/** The names bound by the calls around, the innermost last. */
function isBound(frames: Frame[], upper: string): boolean {
  for (let k = frames.length - 1; k >= 0; k--) if (frames[k].bound.has(upper)) return true;
  return false;
}

/**
 * A formula as files store it: "_xlfn." in front of the functions that need it, "_xlpm." in
 * front of the names of LET and the parameters of LAMBDA, "A1#" and "@" written as the calls
 * files have for them. The text is the formula without "=", with English names and "," between
 * arguments; what has a prefix already keeps it.
 */
export function withFuturePrefixes(text: string): string {
  if (typeof text !== 'string' || text.length > MAX_LENGTH) return text;
  const n = text.length;
  let out = '';
  let i = 0;
  /** The calls being read, innermost last. */
  const frames: Frame[] = [];
  /** Every open parenthesis, with the call it belongs to (none for a parenthesis of its own). */
  const opens: (Frame | undefined)[] = [];
  /** Where in `out` the reference or name written last starts (its sheet prefix included); -1 after anything else. */
  let operandStart = -1;
  /** Where in `out` the sheet prefix written last starts. */
  let prefixStart = -1;
  /** Each "@" waits for the end of what it takes: the number of open parentheses it was met at. */
  const singles: number[] = [];
  /** After an operand: what an "@" at this depth takes ends where an operator, a separator or a ")" follows. */
  const closeSingles = () => {
    while (singles.length > 0 && singles[singles.length - 1] === opens.length) {
      let j = i;
      while (j < n && isSpace(text.charCodeAt(j))) j++;
      if (j < n) {
        const c = text.charCodeAt(j);
        // A range goes on ("@A1:A3"), and so does the spill of what was written ("@A1#"); anything else ends
        // what "@" takes, a space too ("@A:A 1:1" is the intersection of "@A:A" and "1:1"), as the parser reads it.
        if (j === i && (c === 58 || c === 35)) return;
      }
      singles.pop();
      out += ')';
    }
  };
  while (i < n) {
    const c = text.charCodeAt(i);
    if (c === 34) {
      const end = quotedEnd(text, i);
      out += text.slice(i, end);
      i = end;
      operandStart = -1;
      closeSingles();
    } else if (c === 39) {
      // A sheet name in quotes.
      const end = quotedEnd(text, i);
      prefixStart = out.length;
      out += text.slice(i, end);
      i = end;
    } else if (c === 91) {
      const end = bracketEnd(text, i);
      out += text.slice(i, end);
      i = end;
    } else if (c === 35) {
      // "A1#": the spill of the reference just written. Anything else is an error constant.
      const next = text.charCodeAt(i + 1);
      if (operandStart >= 0 && !(i + 1 < n && (isWordChar(next) || next === 47 || next === 33))) {
        out = `${out.slice(0, operandStart)}_xlfn.ANCHORARRAY(${out.slice(operandStart)})`;
        i++;
      } else {
        let j = i + 1;
        while (j < n && (isWordChar(text.charCodeAt(j)) || text.charCodeAt(j) === 47 || text.charCodeAt(j) === 33)) j++;
        out += text.slice(i, j);
        i = j;
        operandStart = -1;
      }
      closeSingles();
    } else if (c === 64) {
      out += '_xlfn.SINGLE(';
      i++;
      while (i < n && isSpace(text.charCodeAt(i))) i++;
      if (text.charCodeAt(i) === 40) {
        // "@(...)": the parentheses are those of the call, and its ")" closes it.
        opens.push(undefined);
        i++;
        operandStart = -1;
      } else singles.push(opens.length);
    } else if (c >= 48 && c <= 57) {
      // A number, or the rows of "1:3"; an exponent's sign is copied as it comes.
      const end = wordEnd(text, i);
      out += text.slice(i, end);
      operandStart = -1;
      i = end;
      closeSingles();
    } else if (isWordChar(c)) {
      const end = wordEnd(text, i);
      let word = text.slice(i, end);
      const next = text.charCodeAt(end);
      let j = end;
      while (j < n && isSpace(text.charCodeAt(j))) j++;
      const afterSheet = out.charCodeAt(out.length - 1) === 33;
      const start = afterSheet && prefixStart >= 0 ? prefixStart : out.length;
      const drop = prefixLength(word);
      const upper = word.slice(drop).toUpperCase();
      if (next === 40) {
        // A call: of a newer function, of a LAMBDA a name of LET stands for, or of anything else.
        if (drop === 0 && !afterSheet) {
          if (isBound(frames, upper)) word = `_xlpm.${word}`;
          else if (WORKSHEET_ONLY.has(upper)) word = `_xlfn._xlws.${word}`;
          else if (FUTURE.has(upper)) word = `_xlfn.${word}`;
        }
        out += word + '(';
        i = end + 1;
        const frame: Frame = { name: upper, arg: 0, start: out.length, bound: new Set() };
        frames.push(frame);
        opens.push(frame);
        operandStart = -1;
        continue;
      }
      const sheetPrefix = next === 33 || (next === 58 && wordEnd(text, end + 1) > end + 1 && text.charCodeAt(wordEnd(text, end + 1)) === 33);
      if (sheetPrefix) {
        prefixStart = out.length;
        out += word;
        i = end;
        continue;
      }
      const frame = frames.length > 0 ? frames[frames.length - 1] : undefined;
      const plain = drop === 0 && !afterSheet && !isCellWord(word) && out.charCodeAt(out.length - 1) !== 58;
      // The whole of an argument of LET or LAMBDA that more arguments follow is a name the call binds.
      // (Only spaces may stand before it in the argument: a longer stretch is not looked at, so that no text makes this slow.)
      const before = frame ? out.length - frame.start : 0;
      const whole = !!frame && opens[opens.length - 1] === frame && before <= 64 && out.slice(frame.start).trim() === '' && text.charCodeAt(j) === 44;
      if (plain && whole && frame && ((frame.name === 'LET' && frame.arg % 2 === 0) || frame.name === 'LAMBDA')) {
        frame.bound.add(upper);
        word = `_xlpm.${word}`;
      } else if (plain && next !== 58 && isBound(frames, upper)) word = `_xlpm.${word}`;
      operandStart = start;
      out += word;
      i = end;
      closeSingles();
    } else if (c === 40) {
      out += '(';
      opens.push(undefined);
      i++;
      operandStart = -1;
    } else if (c === 41) {
      if (opens.pop()) frames.pop();
      out += ')';
      i++;
      operandStart = -1;
      closeSingles();
    } else if (c === 44) {
      const top = opens.length > 0 ? opens[opens.length - 1] : undefined;
      out += ',';
      i++;
      if (top) {
        top.arg++;
        top.start = out.length;
      }
      operandStart = -1;
    } else if (c === 123) {
      // An array constant holds no names.
      let j = i + 1;
      while (j < n && text.charCodeAt(j) !== 125) j = text.charCodeAt(j) === 34 ? quotedEnd(text, j) : j + 1;
      out += text.slice(i, Math.min(j + 1, n));
      i = j + 1;
      operandStart = -1;
      closeSingles();
    } else {
      // ":" and spaces keep the reference going ("A1:B2#" is the range to the spill of B2: the "#" takes B2 alone).
      operandStart = -1;
      out += text.charAt(i);
      i++;
      if (c === 37) closeSingles();
    }
  }
  for (let k = 0; k < singles.length; k++) out += ')';
  return out;
}
/**
 * How much of a word is a prefix that is not shown: "_xlfn." before a function this engine
 * knows as a newer one ("_xlfn._xlws." before FILTER and SORT), spelled as withFuturePrefixes()
 * writes it, so that what is shown and typed again is stored as it was; "_xlpm." before a name
 * of LET or LAMBDA. Any other prefix ("_xlfn.TEXTSPLIT", of a function the engine does not
 * have) is shown as it is, and kept so.
 */
function shownPrefix(word: string, next: number): number {
  if (word.lastIndexOf('_xlpm.', 0) === 0) return next === 33 || word.length <= 6 ? 0 : 6;
  if (next !== 40) return 0;
  if (word.lastIndexOf('_xlfn._xlws.', 0) === 0) return WORKSHEET_ONLY.has(word.slice(12).toUpperCase()) ? 12 : 0;
  if (word.lastIndexOf('_xlfn.', 0) === 0) return FUTURE.has(word.slice(6).toUpperCase()) ? 6 : 0;
  return 0;
}

/**
 * A stored formula the way it is shown and typed: without the prefixes of newer functions and
 * of the names of LET and LAMBDA, with "A1#" and "@" for the calls files have for them.
 *
 * Whatever is shown is typed back as the very text stored (withFuturePrefixes() gives it again,
 * parentheses and spelling included), or it is not shown that way: the short forms "#" and "@"
 * are used where the whole formula comes back so; where not ("_xlfn.SINGLE(A1):B5" would come
 * back as "@A1:B5", the intersection of the whole range), the calls are shown as calls, and a
 * formula that does not come back even then is shown with its prefixes, as stored.
 */
export function withoutFuturePrefixes(text: string): string {
  if (typeof text !== 'string' || text.length > MAX_LENGTH || text.indexOf('_') < 0) return text;
  const short = stripPrefixes(text, false);
  if (withFuturePrefixes(short) === text) return short;
  const long = stripPrefixes(text, true);
  if (long !== short && withFuturePrefixes(long) === text) return long;
  return text;
}

/**
 * One pass over a stored formula: the prefixes go, and "_xlfn.ANCHORARRAY(...)" and
 * "_xlfn.SINGLE(...)" become "...#" and "@..." (`long`: "ANCHORARRAY(...)" and "@(...)", what
 * they hold as it is). What a call of those holds is known when its ")" is met.
 */
/** Whether a text reads as a single operand that "#" or "@" may stand next to: no operator joins it. */
function isOperand(inner: string): boolean {
  if (inner.length === 0 || inner.length > 512 || isSpace(inner.charCodeAt(0)) || isSpace(inner.charCodeAt(inner.length - 1))) return false;
  try {
    const node = parseFormula(inner);
    return node.type !== 'binary' && node.type !== 'unary';
  } catch (error) {
    return false;
  }
}

function stripPrefixes(text: string, long: boolean): string {
  let out = '';
  /** Every open parenthesis; for "_xlfn.ANCHORARRAY(" and "_xlfn.SINGLE(", which one and where in `out` what it holds starts. */
  const opens: ({ spill: boolean; start: number } | undefined)[] = [];
  const n = text.length;
  let i = 0;
  while (i < n) {
    const c = text.charCodeAt(i);
    if (c === 34 || c === 39) {
      const end = quotedEnd(text, i);
      out += text.slice(i, end);
      i = end;
    } else if (c === 91) {
      const end = bracketEnd(text, i);
      out += text.slice(i, end);
      i = end;
    } else if (isWordChar(c)) {
      const end = wordEnd(text, i);
      const word = text.slice(i, end);
      const drop = shownPrefix(word, text.charCodeAt(end));
      const name = word.slice(drop).toUpperCase();
      if (drop > 0 && text.charCodeAt(end) === 40 && (name === 'ANCHORARRAY' || name === 'SINGLE')) {
        opens.push({ spill: name === 'ANCHORARRAY', start: out.length });
        i = end + 1;
        continue;
      }
      out += word.slice(drop);
      i = end;
    } else if (c === 40) {
      opens.push(undefined);
      out += '(';
      i++;
    } else if (c === 41) {
      const open = opens.pop();
      if (open) {
        const inner = out.slice(open.start);
        // The short form where what is inside reads as one operand of it ("A1", "$A$1:B2", "'S'!A1"); else "ANCHORARRAY(...)", "@(...)".
        const plain = !long && isOperand(inner);
        out = out.slice(0, open.start) + (plain ? (open.spill ? `${inner}#` : `@${inner}`) : open.spill ? `ANCHORARRAY(${inner})` : `@(${inner})`);
      } else out += ')';
      i++;
    } else {
      out += text.charAt(i);
      i++;
    }
  }
  // A call that is never closed (a formula that cannot be read) is written as it came, without its prefix.
  for (let k = opens.length - 1; k >= 0; k--) {
    const open = opens[k];
    if (open) out = out.slice(0, open.start) + (open.spill ? 'ANCHORARRAY(' : 'SINGLE(') + out.slice(open.start);
  }
  return out;
}
