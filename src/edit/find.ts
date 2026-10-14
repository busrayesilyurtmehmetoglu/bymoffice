// Finding text in text, for Find and Replace in workbooks and documents. What is
// looked for is plain text, never a pattern: no character of it has a meaning of
// its own. Without regard to case the text is folded the way the locale writes
// it (Turkish pairs "I" with "ı" and "İ" with "i"), one character staying one
// character, so that what is found in the folded text is found at the same
// place in the text itself.

import { getLocale, type FormatLocale } from '../format/locale';
import { foldCase } from '../ops/sort';

/** Longest text that is looked for, and longest replacement. */
export const MAX_FIND_TEXT = 1000;

export interface FindQuery {
  /** The text to look for, as it is. */
  text: string;
  /** Letters must have the same case. Off by default. */
  matchCase?: boolean;
  /**
   * 'word': only where no letter or digit goes on right before or right after
   * the match (Word's "whole words only"); 'all': the whole text must be the
   * match (Excel's "match entire cell contents").
   */
  whole?: 'word' | 'all';
  /** Case is folded the way this locale writes (Turkish "I" and "ı", "İ" and "i"); 'tr' and 'en' stand for those languages. */
  locale?: FormatLocale | 'tr' | 'en';
}

/** A query made ready for looking: see compileFind(). */
export interface CompiledFind {
  readonly needle: string;
  readonly matchCase: boolean;
  readonly whole: 'word' | 'all' | undefined;
  readonly locale: FormatLocale | undefined;
}

const TURKISH = getLocale('tr');

/**
 * Checks and prepares a query. Undefined for an empty text: nothing is looked
 * for. Throws RangeError for a text that is not a string or is longer than
 * MAX_FIND_TEXT.
 */
export function compileFind(query: FindQuery): CompiledFind | undefined {
  if (!query || typeof query.text !== 'string') throw new RangeError('The text to find must be a string');
  if (query.text.length > MAX_FIND_TEXT) throw new RangeError(`The text to find is longer than ${MAX_FIND_TEXT} characters`);
  if (query.text === '') return undefined;
  const matchCase = query.matchCase === true;
  const whole = query.whole === 'word' || query.whole === 'all' ? query.whole : undefined;
  const locale = query.locale === 'tr' ? TURKISH : query.locale === 'en' || typeof query.locale !== 'object' ? undefined : query.locale;
  const needle = matchCase ? query.text : foldCase(query.text, locale);
  return { needle, matchCase, whole, locale };
}

/** The text as it is compared: folded unless case matters. Always as long as the text. */
export function foldFor(find: CompiledFind, text: string): string {
  if (find.matchCase) return text;
  const folded = foldCase(text, find.locale);
  // One character stays one character (see foldCase); should a case mapping ever break that, case is kept.
  return folded.length === text.length ? folded : text;
}

let wordChar: RegExp;
try {
  wordChar = new RegExp('[\\p{L}\\p{N}\\p{M}_]', 'u');
} catch {
  // A runtime without Unicode property escapes: letters of the Basic Multilingual Plane beyond ASCII count as letters.
  wordChar = /[0-9A-Za-z_À-ɏͰ-ϿЀ-ӿ֐-ۿ฀-๿぀-鿿가-힯]/;
}

/** Whether the character at `index` belongs to a word (a letter, a digit, a mark or "_"). A surrogate pair is read whole. */
function isWordAt(text: string, index: number): boolean {
  if (index < 0 || index >= text.length) return false;
  const code = text.charCodeAt(index);
  let char = text.charAt(index);
  if (code >= 0xd800 && code <= 0xdbff && index + 1 < text.length) char = text.slice(index, index + 2);
  else if (code >= 0xdc00 && code <= 0xdfff && index > 0) char = text.slice(index - 1, index + 1);
  return wordChar.test(char);
}

/** Whether a match at `start` in `text` stands as the query wants it to (a word of its own, the whole text). */
function stands(find: CompiledFind, text: string, start: number): boolean {
  const end = start + find.needle.length;
  if (find.whole === 'all') return start === 0 && end === text.length;
  if (find.whole === 'word') return !isWordAt(text, start - 1) && !isWordAt(text, end);
  return true;
}

/** Where the first match at or after `from` starts; -1 when there is none. `folded`: the text as foldFor() gives it, when it is at hand. */
export function findIn(find: CompiledFind, text: string, from = 0, folded?: string): number {
  if (find.whole === 'all') return from <= 0 && text.length === find.needle.length && (folded !== undefined ? folded : foldFor(find, text)) === find.needle ? 0 : -1;
  const hay = folded !== undefined ? folded : foldFor(find, text);
  let at = Math.max(0, from);
  // A text has at most as many places to try as it has characters.
  for (let guard = 0; guard <= hay.length; guard++) {
    const index = hay.indexOf(find.needle, at);
    if (index < 0) return -1;
    if (stands(find, text, index)) return index;
    at = index + 1;
  }
  return -1;
}

/** Where the last match that ends at or before `before` starts; -1 when there is none. */
export function findLastIn(find: CompiledFind, text: string, before = text.length, folded?: string): number {
  const hay = folded !== undefined ? folded : foldFor(find, text);
  if (find.whole === 'all') return before >= text.length && hay.length === find.needle.length && hay === find.needle ? 0 : -1;
  let at = Math.min(before, hay.length) - find.needle.length;
  for (let guard = 0; guard <= hay.length && at >= 0; guard++) {
    const index = hay.lastIndexOf(find.needle, at);
    if (index < 0) return -1;
    if (stands(find, text, index)) return index;
    at = index - 1;
  }
  return -1;
}

/** The starts of the matches in a text, one after the other and not overlapping; at most `limit` of them. */
export function matchesIn(find: CompiledFind, text: string, limit = Infinity): number[] {
  const out: number[] = [];
  const folded = foldFor(find, text);
  let from = 0;
  while (out.length < limit) {
    const index = findIn(find, text, from, folded);
    if (index < 0) break;
    out.push(index);
    from = index + find.needle.length;
  }
  return out;
}

/** The text with every match replaced. `count`: how many were. */
export function replaceIn(find: CompiledFind, text: string, replacement: string): { text: string; count: number } {
  const starts = matchesIn(find, text);
  if (starts.length === 0) return { text, count: 0 };
  let out = '';
  let from = 0;
  for (let i = 0; i < starts.length; i++) {
    out += text.slice(from, starts[i]) + replacement;
    from = starts[i] + find.needle.length;
  }
  return { text: out + text.slice(from), count: starts.length };
}

/** Checks a replacement: a string no longer than MAX_FIND_TEXT. Throws RangeError otherwise. */
export function checkReplacement(replacement: unknown): string {
  if (typeof replacement !== 'string') throw new RangeError('The replacement must be a string');
  if (replacement.length > MAX_FIND_TEXT) throw new RangeError(`The replacement is longer than ${MAX_FIND_TEXT} characters`);
  return replacement;
}
