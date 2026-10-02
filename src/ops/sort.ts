// Sorting: the order Excel gives the values of a column, and the command that
// sorts the rows of a range by one or more of its columns. The command only
// works out where each row goes; the operation it returns ("sortRange", see
// ops.ts) moves the cells and is what undo and redo see.

import { isoToSerial } from '../format/cell-format';
import type { FormatLocale } from '../format/locale';
import { resolveCellStyle } from '../model/sheet/color';
import { MAX_COLS, MAX_ROWS, getCell, type AutoFilter, type Cell, type CellRange, type Row, type Sheet, type Workbook } from '../model/sheet/types';
import type { EditOptions } from './commands';
import type { Op } from './ops';

/** One column to sort by; the first key decides, the next ones break its ties. */
export interface SortKey {
  /** Column index in the sheet; it has to lie inside the sorted range. */
  col: number;
  descending?: boolean;
}

export interface SortOptions extends EditOptions {
  /** The first row of the range holds titles and stays where it is. */
  header?: boolean;
}

// Kinds of values in the order an ascending sort puts them. Empty cells come last in either direction.
export const SORT_NUMBER = 0;
export const SORT_TEXT = 1;
export const SORT_BOOLEAN = 2;
export const SORT_ERROR = 3;
export const SORT_BLANK = 4;

/**
 * Whether a locale writes Turkish (or Azerbaijani): its alphabet has letters of its own and
 * two pairs of "i". A locale has no name, so it is known by its first month.
 */
function isTurkish(locale: FormatLocale | undefined): boolean {
  if (!locale || !Array.isArray(locale.monthNames)) return false;
  const first = locale.monthNames[0];
  return first === 'Ocak' || first === 'Yanvar';
}

/**
 * Text without regard to case, for comparing what cells show: one character stays one
 * character. Turkish pairs "I" with "ı" and "İ" with "i"; everywhere else "I" is the capital
 * of "i", and "İ" and "ı" are letters of their own.
 */
export function foldCase(text: string, locale?: FormatLocale): string {
  if (isTurkish(locale)) return text.replace(/I/g, 'ı').replace(/İ/g, 'i').toLowerCase();
  // toLowerCase() makes two characters of "İ" ("i" and a combining dot).
  if (text.indexOf('İ') < 0) return text.toLowerCase();
  return text
    .split('İ')
    .map((part) => part.toLowerCase())
    .join('İ');
}

/** The Turkish alphabet with the letters it lacks (q, w, x) where dictionaries put them. */
const TURKISH_ALPHABET = 'abcçdefgğhıijklmnoöpqrsştuüvwxyz';
const LETTER_BASE = 0x200;
const OTHER_BASE = 0x400;

// What a letter carries on top of its base letter, in the order a collator ranks them
// (the plain letter first). The values end up as characters of the key.
const PLAIN = 0;
const ACUTE = 1;
const GRAVE = 2;
const BREVE = 3;
const CIRCUMFLEX = 4;
const CARON = 5;
const RING = 6;
const DIAERESIS = 7;
const DOUBLE_ACUTE = 8;
const TILDE = 9;
const DOT_ABOVE = 10;
const CEDILLA = 11;
const OGONEK = 12;
const STROKE = 13;
const MIDDLE_DOT = 14;
const MACRON = 15;
/** A letter that is not its base letter with a mark: eth, long s, dotless i, eng, kra. */
const VARIANT = 16;
/** Two letters written as one (æ, œ, ß): after every spelling of the two with marks. */
const LIGATURE = 17;

/**
 * The letters of Latin-1 Supplement and Latin Extended-A, lower case, by what they carry:
 * each entry is the letter followed by its base letter (or letters). `Intl` is not used, so
 * the table is written out; it follows the order of the Unicode collation algorithm, except
 * that letters with a place of their own there (þ, ŋ, ŧ, ĸ) are folded to a base letter too.
 */
const MARKED: [mark: number, letters: string][] = [
  [ACUTE, 'áa ée íi óo úu ýy ćc ĺl ńn ŕr śs źz'],
  [GRAVE, 'àa èe ìi òo ùu'],
  [BREVE, 'ăa ĕe ğg ĭi ŏo ŭu'],
  [CIRCUMFLEX, 'âa êe îi ôo ûu ĉc ĝg ĥh ĵj ŝs ŵw ŷy'],
  [CARON, 'čc ďd ěe ľl ňn řr šs ťt žz'],
  [RING, 'åa ůu'],
  [DIAERESIS, 'äa ëe ïi öo üu ÿy'],
  [DOUBLE_ACUTE, 'őo űu'],
  [TILDE, 'ãa ñn õo ĩi ũu'],
  [DOT_ABOVE, 'ċc ėe ġg żz'],
  [CEDILLA, 'çc ģg ķk ļl ņn ŗr şs ţt'],
  [OGONEK, 'ąa ęe įi ųu'],
  [STROKE, 'đd ħh łl øo ŧt'],
  [MIDDLE_DOT, 'ŀl'],
  [MACRON, 'āa ēe īi ōo ūu'],
  [VARIANT, 'ðd ıi ĸk ŋn ſs ŉn'],
  [LIGATURE, 'æae œoe ßss þth'],
  // Two letters that only look like one.
  [PLAIN, 'ĳij'],
];

/** The Turkish letters among them: in a Turkish text they are letters of the alphabet, not marked ones. */
const TURKISH_OWN = 'çğıöşü';

interface Folded {
  base: string;
  mark: number;
}

let foldedLetters: Map<number, Folded> | undefined;

/** The base letters and the mark of a letter outside ASCII, lower or upper case; undefined for any other character. */
function foldedLetter(code: number): Folded | undefined {
  if (!foldedLetters) {
    foldedLetters = new Map();
    for (const [mark, letters] of MARKED) {
      for (const entry of letters.split(' ')) {
        const folded = { base: entry.slice(1), mark };
        const lower = entry.charAt(0);
        foldedLetters.set(lower.charCodeAt(0), folded);
        const upper = lower.toUpperCase();
        // "ß" and "ŉ" have no capital of one character; "ı" has "I", which is a letter of ASCII.
        if (upper.length === 1 && upper !== lower && upper.charCodeAt(0) >= 128) foldedLetters.set(upper.charCodeAt(0), folded);
      }
    }
    // The capital of "i" in Turkish: elsewhere an "i" that kept its dot.
    foldedLetters.set(0x130, { base: 'i', mark: DOT_ABOVE });
  }
  return foldedLetters.get(code);
}

let turkishWeights: Map<number, number> | undefined;

/** The place of a character in Turkish text; undefined for a letter that is folded to a base letter. */
function turkishWeight(code: number): number | undefined {
  if (!turkishWeights) {
    turkishWeights = new Map();
    for (let i = 0; i < TURKISH_ALPHABET.length; i++) {
      const lower = TURKISH_ALPHABET.charCodeAt(i);
      turkishWeights.set(lower, LETTER_BASE + i);
      // Upper case by hand: "ı" pairs with "I" and "i" with "İ".
      const upper = lower === 0x131 ? 0x49 : lower === 0x69 ? 0x130 : TURKISH_ALPHABET.charAt(i).toUpperCase().charCodeAt(0);
      turkishWeights.set(upper, LETTER_BASE + i);
    }
  }
  const weight = turkishWeights.get(code);
  if (weight !== undefined) return weight;
  // Digits and ASCII punctuation before the letters.
  if (code < 128) return code;
  // Everything that is not a letter with a base letter after them, in code order.
  return foldedLetter(code) ? undefined : Math.min(OTHER_BASE + code, 0xffff);
}

/** Separates the letters of a key from their marks; lower than any letter, so that a shorter word comes first. */
const MARKS_FOLLOW = '\u0000';
const MARK_BASE = 0x20;

/**
 * A text whose order by code units is the order the text sorts in, without regard to
 * case. Letters are compared first, by their base letters: in English every letter outside
 * ASCII is its base letter with a mark ("ç" is a "c", "ı" an "i"), in Turkish the letters of
 * the Turkish alphabet have their own places (ç after c, ı before i, ...) and the others are
 * folded ("â" is an "a"). Words with the same letters are told apart by their marks, the
 * plain letter first: a < â < b, "kar" < "kâr" < "kara". `Intl` is not used: its results
 * differ between engines, and old ones lack it.
 */
export function collationKey(text: string, locale?: FormatLocale): string {
  const turkish = isTurkish(locale);
  // Nothing but ASCII: the lower-case text is the key.
  if (!turkish && !/[^\u0000-\u007f]/.test(text)) return text.toLowerCase();
  let letters = '';
  let marks = '';
  // How many marks are worth writing: the plain ones at the end say nothing.
  let marked = 0;
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    if (turkish) {
      const weight = turkishWeight(code);
      if (weight !== undefined) {
        letters += String.fromCharCode(weight);
        marks += String.fromCharCode(MARK_BASE);
        continue;
      }
    } else if (code < 128) {
      letters += code >= 65 && code <= 90 ? String.fromCharCode(code + 32) : text.charAt(i);
      marks += String.fromCharCode(MARK_BASE);
      continue;
    }
    const folded = foldedLetter(code);
    if (!folded) {
      letters += text.charAt(i).toLowerCase();
      marks += String.fromCharCode(MARK_BASE);
      continue;
    }
    for (let k = 0; k < folded.base.length; k++) {
      let base = folded.base.charCodeAt(k);
      // A capital "I" with a mark is, in Turkish, the capital of the dotless "ı".
      if (turkish && base === 0x69 && folded.base.length === 1 && code !== 0x130 && text.charAt(i).toLowerCase() !== text.charAt(i)) base = 0x131;
      letters += String.fromCharCode(turkish ? (turkishWeight(base) as number) : base);
      // The mark stands with the first letter.
      marks += String.fromCharCode(MARK_BASE + (k === 0 ? folded.mark : PLAIN));
      if (k === 0 && folded.mark !== PLAIN) marked = marks.length;
    }
  }
  return marked === 0 ? letters : letters + MARKS_FOLLOW + marks.slice(0, marked);
}

/** What a cell is worth to a sort: its kind, and the number or text to compare within the kind. */
export interface SortValue {
  rank: number;
  number: number;
  text: string;
}

export function sortValueOf(cell: Cell | undefined, date1904: boolean): SortValue {
  if (!cell || cell.value === null || cell.value === undefined || cell.kind === 'blank') return { rank: SORT_BLANK, number: 0, text: '' };
  const value = cell.value;
  if (cell.kind === 'error') return { rank: SORT_ERROR, number: 0, text: '' };
  if (typeof value === 'number') return { rank: SORT_NUMBER, number: value, text: '' };
  if (typeof value === 'boolean') return { rank: SORT_BOOLEAN, number: value ? 1 : 0, text: '' };
  if (cell.kind === 'date') {
    const serial = isoToSerial(String(value), date1904);
    if (serial !== undefined) return { rank: SORT_NUMBER, number: serial, text: '' };
  }
  return { rank: SORT_TEXT, number: 0, text: String(value) };
}

const isFilled = (cell: Cell | undefined) => cell !== undefined && ((cell.value !== null && cell.value !== '') || cell.formula !== undefined);

/**
 * The block of cells around a cell that empty rows and columns (or the edges of the
 * sheet) close in on all sides: what Excel sorts and filters when one cell is selected.
 * A cell with nothing around it is its own region.
 */
export function currentRegion(sheet: Sheet, row: number, col: number): CellRange {
  const maxRow = Math.min(sheet.rowCount, MAX_ROWS) - 1;
  const maxCol = Math.min(sheet.colCount, MAX_COLS) - 1;
  let top = row;
  let bottom = row;
  let left = col;
  let right = col;
  const rowHas = (r: number): boolean => {
    const data = sheet.rows.get(r);
    if (!data) return false;
    const last = Math.min(right + 1, maxCol, data.cells.length - 1);
    for (let c = Math.max(left - 1, 0); c <= last; c++) if (isFilled(data.cells[c])) return true;
    return false;
  };
  const colHas = (c: number): boolean => {
    const last = Math.min(bottom + 1, maxRow);
    for (let r = Math.max(top - 1, 0); r <= last; r++) if (isFilled(getCell(sheet, r, c))) return true;
    return false;
  };
  // Every pass adds a row or a column inside the used area, which bounds the loop; rows are
  // added in runs, so that a tall block costs one look at each of its columns per column added.
  for (;;) {
    let grew = false;
    while (bottom < maxRow && rowHas(bottom + 1)) {
      bottom++;
      grew = true;
    }
    while (top > 0 && top - 1 <= maxRow && rowHas(top - 1)) {
      top--;
      grew = true;
    }
    if (left > 0 && left - 1 <= maxCol && colHas(left - 1)) {
      left--;
      grew = true;
    }
    if (right < maxCol && colHas(right + 1)) {
      right++;
      grew = true;
    }
    if (!grew) break;
  }
  return { top, left, bottom, right };
}

/** The part of a range that can hold cells; undefined when there is none. */
function usedPart(sheet: Sheet, range: CellRange): CellRange | undefined {
  const top = Math.max(Math.floor(range.top), 0);
  const left = Math.max(Math.floor(range.left), 0);
  const bottom = Math.min(Math.floor(range.bottom), sheet.rowCount - 1, MAX_ROWS - 1);
  const right = Math.min(Math.floor(range.right), sheet.colCount - 1, MAX_COLS - 1);
  return top <= bottom && left <= right ? { top, left, bottom, right } : undefined;
}

/**
 * The filter with the sort that was just made noted in it, when that sort was one of its data
 * rows from the first one on. What is noted is the range that was sorted, which may be less
 * than all the rows of the filter.
 */
function withSortState(filter: AutoFilter | undefined, sorted: CellRange, keys: SortKey[]): AutoFilter | undefined {
  if (!filter) return undefined;
  const range = filter.range;
  if (sorted.top !== range.top + 1 || sorted.left < range.left || sorted.right > range.right || sorted.bottom > range.bottom) return undefined;
  return {
    ...filter,
    sortState: {
      range: { top: sorted.top, left: sorted.left, bottom: sorted.bottom, right: sorted.right },
      conditions: keys.map((key) => ({ range: { top: sorted.top, left: key.col, bottom: sorted.bottom, right: key.col }, descending: key.descending === true })),
    },
  };
}

/** Rows under the first one that are looked at to tell whether the first one is a header. */
const HEADER_SAMPLE_ROWS = 100;
/** Columns of a block that are looked at for the same. */
const HEADER_SAMPLE_COLS = 256;

/**
 * Whether the first row of a block of cells looks like a header, for a sort that was not
 * told: in some column the first cell is of another kind than most of the cells under it
 * (text above numbers, say), or the first row is in bold type and the row under it is not.
 * A block whose first row is like the rest has no header, and is sorted from its first row.
 *
 * UNSURE: Excel's own guess is not documented; it looks at kinds and formats in much this way.
 */
export function hasHeaderRow(workbook: Workbook, sheet: Sheet, range: CellRange): boolean {
  if (!(range.bottom > range.top)) return false;
  const date1904 = workbook.date1904 === true;
  const first = sheet.rows.get(range.top);
  if (!first) return false;
  const bottom = Math.min(range.bottom, range.top + HEADER_SAMPLE_ROWS, sheet.rowCount - 1);
  const right = Math.min(range.right, range.left + HEADER_SAMPLE_COLS - 1, first.cells.length - 1);
  const counts = [0, 0, 0, 0];
  for (let col = range.left; col <= right; col++) {
    const head = sortValueOf(first.cells[col], date1904).rank;
    if (head === SORT_BLANK) continue;
    counts[0] = counts[1] = counts[2] = counts[3] = 0;
    for (let r = range.top + 1; r <= bottom; r++) {
      const row = sheet.rows.get(r);
      const rank = row ? sortValueOf(row.cells[col], date1904).rank : SORT_BLANK;
      if (rank !== SORT_BLANK) counts[rank]++;
    }
    // The kind most cells have; where two kinds are as frequent, the one of the first cell counts as it.
    let most = head;
    for (let rank = 0; rank < counts.length; rank++) if (counts[rank] > counts[most]) most = rank;
    if (counts[most] > 0 && most !== head) return true;
  }
  const bold = (row: Row | undefined): boolean | undefined => {
    let any: boolean | undefined;
    for (let col = range.left; row && col <= Math.min(right, row.cells.length - 1); col++) {
      const cell = row.cells[col];
      if (!cell || sortValueOf(cell, date1904).rank === SORT_BLANK) continue;
      if (resolveCellStyle(workbook.styles, cell.style).font.bold !== true) return false;
      any = true;
    }
    return any;
  };
  return bold(first) === true && bold(sheet.rows.get(range.top + 1)) === false;
}

/**
 * Sorts the rows of a range by one or more of its columns. Numbers come before text,
 * text before TRUE and FALSE, those before errors; a descending sort turns that round.
 * Empty cells come last either way. Text is compared without regard to case, in the
 * order of the locale's alphabet. Rows that compare equal keep their order. Hidden rows
 * stay where they are, as in Excel: a sort inside a filter moves what the filter shows.
 *
 * Cells move with their formats and formulas; relative references move along. Row
 * heights stay. A range with merged cells, or one that cuts an array formula, is
 * refused when the operation is applied.
 */
export function sortRange(workbook: Workbook, sheetIndex: number, range: CellRange, keys: SortKey[], options: SortOptions = {}): Op[] {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet) throw new RangeError(`No sheet at index ${sheetIndex}`);
  if (!Array.isArray(keys) || keys.length === 0) throw new RangeError('A sort needs at least one key');
  let used = usedPart(sheet, range);
  if (used && options.header) used = used.top < used.bottom ? { ...used, top: used.top + 1 } : undefined;
  for (const key of keys) {
    if (!key || !(key.col >= range.left && key.col <= range.right) || Math.floor(key.col) !== key.col) throw new RangeError('The sort key is outside the range');
  }
  if (!used) return [];
  const data = used;
  const height = data.bottom - data.top + 1;
  const date1904 = workbook.date1904 === true;

  // The rows that take part, and for each key what every one of them is worth.
  const visible: number[] = [];
  for (let i = 0; i < height; i++) {
    const row = sheet.rows.get(data.top + i);
    if (!row || !row.hidden) visible.push(i);
  }
  const count = visible.length;
  const ranks: Uint8Array[] = [];
  const numbers: Float64Array[] = [];
  const texts: (string[] | undefined)[] = [];
  for (const key of keys) {
    const rank = new Uint8Array(count);
    const number = new Float64Array(count);
    let text: string[] | undefined;
    for (let k = 0; k < count; k++) {
      const row = sheet.rows.get(data.top + visible[k]);
      const value = sortValueOf(row ? row.cells[key.col] : undefined, date1904);
      rank[k] = value.rank;
      if (value.rank === SORT_TEXT) {
        if (!text) text = new Array<string>(count);
        text[k] = collationKey(value.text, options.locale);
      } else number[k] = value.number;
    }
    ranks.push(rank);
    numbers.push(number);
    texts.push(text);
  }

  const picks: number[] = new Array(count);
  for (let k = 0; k < count; k++) picks[k] = k;
  picks.sort((a, b) => {
    for (let i = 0; i < keys.length; i++) {
      const rank = ranks[i];
      const ra = rank[a];
      const rb = rank[b];
      if (ra === SORT_BLANK || rb === SORT_BLANK) {
        if (ra === rb) continue;
        return ra === SORT_BLANK ? 1 : -1;
      }
      let order = ra - rb;
      if (order === 0) {
        if (ra === SORT_TEXT) {
          const text = texts[i] as string[];
          order = text[a] < text[b] ? -1 : text[a] > text[b] ? 1 : 0;
        } else if (ra !== SORT_ERROR) {
          const number = numbers[i];
          order = number[a] < number[b] ? -1 : number[a] > number[b] ? 1 : 0;
        }
      }
      if (order !== 0) return keys[i].descending ? -order : order;
    }
    // The order the rows had: not every engine this runs in sorts stably by itself.
    return a - b;
  });

  const order: number[] = new Array(height);
  for (let i = 0; i < height; i++) order[i] = i;
  let moved = false;
  for (let k = 0; k < count; k++) {
    order[visible[k]] = visible[picks[k]];
    if (picks[k] !== k) moved = true;
  }

  const ops: Op[] = [];
  if (moved) ops.push({ type: 'sortRange', sheet: sheetIndex, range: data, order });
  const filter = withSortState(sheet.autoFilter, data, keys);
  if (filter) ops.push({ type: 'setAutoFilter', sheet: sheetIndex, autoFilter: filter });
  return ops;
}
