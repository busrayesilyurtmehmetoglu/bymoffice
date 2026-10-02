// Sorting: the order Excel gives the values of a column, and the command that
// sorts the rows of a range by one or more of its columns. The command only
// works out where each row goes; the operation it returns ("sortRange", see
// ops.ts) moves the cells and is what undo and redo see.

import { isoToSerial } from '../format/cell-format';
import type { FormatLocale } from '../format/locale';
import { MAX_COLS, MAX_ROWS, getCell, type AutoFilter, type Cell, type CellRange, type Sheet, type Workbook } from '../model/sheet/types';
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

/** Whether a locale is the Turkish one: its alphabet has letters of its own and two pairs of "i". */
function isTurkish(locale: FormatLocale | undefined): boolean {
  return !!locale && Array.isArray(locale.monthNames) && locale.monthNames[0] === 'Ocak';
}

/** The Turkish alphabet with the letters it lacks (q, w, x) where dictionaries put them. */
const TURKISH_ALPHABET = 'abcçdefgğhıijklmnoöpqrsştuüvwxyz';
const LETTER_BASE = 0x200;
const OTHER_BASE = 0x400;
let turkishWeights: Map<number, number> | undefined;

function turkishWeight(code: number): number {
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
  // Digits and ASCII punctuation before the letters, everything else after them, in code order.
  return code < 128 ? code : Math.min(OTHER_BASE + code, 0xffff);
}

/**
 * A text whose order by code units is the order the text sorts in, without regard to
 * case. English: the lower-case text. Turkish: by the Turkish alphabet (ç after c, ı
 * before i, ...). `Intl` is not used: its results differ between engines.
 */
export function collationKey(text: string, locale?: FormatLocale): string {
  if (!isTurkish(locale)) return text.toLowerCase();
  let out = '';
  for (let i = 0; i < text.length; i++) out += String.fromCharCode(turkishWeight(text.charCodeAt(i)));
  return out;
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

/** The filter with the sort that was just made noted in it, when that sort was one of its data rows. */
function withSortState(filter: AutoFilter | undefined, sorted: CellRange, keys: SortKey[]): AutoFilter | undefined {
  if (!filter) return undefined;
  const range = filter.range;
  if (sorted.top !== range.top + 1 || sorted.left < range.left || sorted.right > range.right || sorted.bottom > range.bottom) return undefined;
  const data = { top: range.top + 1, left: range.left, bottom: range.bottom, right: range.right };
  return {
    ...filter,
    sortState: {
      range: data,
      conditions: keys.map((key) => ({ range: { top: data.top, left: key.col, bottom: data.bottom, right: key.col }, descending: key.descending === true })),
    },
  };
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
