// The filter of a sheet: which rows its criteria show, the list of values a
// column offers, and the commands that switch the filter on and off and change
// the criterion of a column. Like the other commands these only build
// operations; the rows to show and hide are worked out here, because what a
// cell shows depends on the locale, and stated in the operation.
//
// A filter hides rows by marking them hidden, exactly as hiding by hand does:
// the file has one attribute for both. Clearing the criteria therefore shows
// every row under the header.

import { builtinFormatId, formatCell, isoToSerial } from '../format/cell-format';
import { LOCALE_EN, type FormatLocale } from '../format/locale';
import { serialToDate } from '../format/number-format';
import { resolveCellStyle } from '../model/sheet/color';
import { MAX_COLS, MAX_ROWS, type AutoFilter, type Cell, type CellRange, type FilterColumn, type FilterCondition, type FilterDateGroup, type Row, type Sheet, type Workbook } from '../model/sheet/types';
import type { EditOptions } from './commands';
import type { Op } from './ops';
import { SORT_BLANK, SORT_TEXT, collationKey, foldCase, sortValueOf } from './sort';

/** Most distinct values the list of a column offers. */
export const MAX_FILTER_VALUES = 10000;

/** Excel's limit on the text of a custom condition; longer text from a file is cut to it. */
const MAX_CONDITION_TEXT = 255;
/** How much of a cell's text a wildcard is matched against: bounds the work a hostile file can ask for. */
const MAX_MATCH_TEXT = 2048;

/** A criterion for a column: a FilterColumn without its position, which the command gives it. */
export type FilterCriterion = Omit<FilterColumn, 'colId' | 'raw'>;

export interface FilterValue {
  /** What the cells show. */
  text: string;
  /** The criterion of the column shows rows with this value now. */
  checked: boolean;
}

export interface FilterValueList {
  /** Distinct values of the column, in sort order; at most MAX_FILTER_VALUES. */
  items: FilterValue[];
  /** The column has empty cells, and whether they are shown now. */
  blanks: boolean;
  blanksChecked: boolean;
  /** There were more distinct values than the list holds. */
  truncated: boolean;
  /**
   * The criterion of the column is one this engine cannot apply (a colour, an icon): which
   * values are checked only says which rows are shown now, and choosing values replaces it.
   */
  opaque: boolean;
}

const isBlank = (cell: Cell | undefined) => !cell || cell.value === null || cell.value === undefined || cell.value === '' || cell.kind === 'blank';

/** The number a cell holds, dates given as text included; undefined for anything else. */
function numberOf(cell: Cell | undefined, date1904: boolean): number | undefined {
  if (!cell) return undefined;
  if (cell.kind === 'number' && typeof cell.value === 'number') return cell.value;
  if (cell.kind === 'date' && typeof cell.value === 'string') return isoToSerial(cell.value, date1904);
  return undefined;
}

/** Reads what cells show, with the format of each style looked up once. */
class TextReader {
  private readonly formats = new Map<number, { code: string; builtinId: number | undefined }>();

  constructor(
    private readonly workbook: Workbook,
    private readonly locale: FormatLocale | undefined,
  ) {}

  text(cell: Cell): string {
    let format = this.formats.get(cell.style);
    if (!format) {
      const resolved = resolveCellStyle(this.workbook.styles, cell.style);
      format = { code: resolved.numFmt, builtinId: builtinFormatId(this.workbook.styles, resolved.numFmtId) };
      this.formats.set(cell.style, format);
    }
    return formatCell(cell, format.code, { locale: this.locale, date1904: this.workbook.date1904, builtinId: format.builtinId }).text;
  }
}

/**
 * "*" stands for any run of characters, "?" for one, "~" takes the next character
 * literally. Both texts are already folded. Linear in practice and never worse than
 * the product of the two lengths: on a mismatch only the last "*" is tried again.
 */
export function wildcardMatch(text: string, pattern: string): boolean {
  let t = 0;
  let p = 0;
  let star = -1;
  let mark = 0;
  const n = text.length;
  const m = pattern.length;
  while (t < n) {
    const c = p < m ? pattern.charAt(p) : '';
    if (c === '*') {
      star = p++;
      mark = t;
    } else if (c === '?') {
      p++;
      t++;
    } else if (c !== '' && (c === '~' && p + 1 < m ? pattern.charAt(p + 1) === text.charAt(t) : c === text.charAt(t))) {
      p += c === '~' && p + 1 < m ? 2 : 1;
      t++;
    } else if (star >= 0) {
      p = star + 1;
      t = ++mark;
    } else return false;
  }
  while (p < m && pattern.charAt(p) === '*') p++;
  return p === m;
}

const hasWildcard = (text: string) => text.indexOf('*') >= 0 || text.indexOf('?') >= 0 || text.indexOf('~') >= 0;

/** A number as files write it ("12", "-1.5", "2E3"); undefined for any other text. */
function plainNumber(text: string): number | undefined {
  if (text.length === 0 || text.length > 32 || !/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(text)) return undefined;
  const value = Number(text);
  return isFinite(value) ? value : undefined;
}

/** Says whether the cell of a row passes the criterion of a column. */
type Matcher = (cell: Cell | undefined) => boolean;

interface Scope {
  workbook: Workbook;
  sheet: Sheet;
  locale: FormatLocale | undefined;
  /** Text without regard to case, the way the locale pairs its letters. */
  fold(text: string): string;
  reader: TextReader;
  /** The same cells read in English, for values a file states that way. */
  english: TextReader;
  top: number;
  bottom: number;
}

function matchesDateGroup(group: FilterDateGroup, serial: number, date1904: boolean): boolean {
  const date = serialToDate(serial, date1904);
  if (date.year !== group.year) return false;
  const grouping = group.grouping;
  if (grouping === 'year') return true;
  if (group.month !== undefined && date.month !== group.month) return false;
  if (grouping === 'month') return true;
  if (group.day !== undefined && date.day !== group.day) return false;
  if (grouping === 'day') return true;
  if (group.hour !== undefined && date.hour !== group.hour) return false;
  if (grouping === 'hour') return true;
  if (group.minute !== undefined && date.minute !== group.minute) return false;
  if (grouping === 'minute') return true;
  return group.second === undefined || date.second === group.second;
}

function conditionMatcher(scope: Scope, condition: FilterCondition): Matcher {
  const { reader, english, locale, fold } = scope;
  const date1904 = scope.workbook.date1904 === true;
  const operator = condition.operator;
  const value = String(condition.value).slice(0, MAX_CONDITION_TEXT);
  const number = plainNumber(value);
  if (operator === 'equal' || operator === 'notEqual') {
    const negate = operator === 'notEqual';
    // Excel writes "is not blank" as: not equal to a space.
    if (value === ' ' || value === '') return (cell) => isBlank(cell) !== negate;
    const pattern = fold(value);
    const wild = hasWildcard(value);
    const matches = (shown: string) => {
      const text = fold(shown);
      return wild ? wildcardMatch(text.length > MAX_MATCH_TEXT ? text.slice(0, MAX_MATCH_TEXT) : text, pattern) : text === pattern;
    };
    return (cell) => {
      let same: boolean;
      if (isBlank(cell)) same = false;
      else {
        const held = numberOf(cell, date1904);
        if (number !== undefined && !wild && held !== undefined && held === number) same = true;
        else {
          same = matches(reader.text(cell as Cell));
          // A file states TRUE and FALSE in English whatever the language its author saw them in.
          if (!same && (cell as Cell).kind === 'boolean' && english !== reader) same = matches(english.text(cell as Cell));
        }
      }
      return same !== negate;
    };
  }
  const compare = (order: number) =>
    operator === 'greaterThan' ? order > 0 : operator === 'greaterThanOrEqual' ? order >= 0 : operator === 'lessThan' ? order < 0 : operator === 'lessThanOrEqual' ? order <= 0 : false;
  if (number !== undefined) {
    // UNSURE: against a number only numbers are compared; text is taken not to pass.
    return (cell) => {
      const held = numberOf(cell, date1904);
      return held !== undefined && compare(held < number ? -1 : held > number ? 1 : 0);
    };
  }
  const key = collationKey(value, locale);
  return (cell) => {
    if (!cell || cell.kind !== 'string' || typeof cell.value !== 'string') return false;
    const held = collationKey(cell.value, locale);
    return compare(held < key ? -1 : held > key ? 1 : 0);
  };
}

/** The numbers a column holds under the header, whatever the other criteria say of their rows. */
function columnNumbers(scope: Scope, col: number): number[] {
  const out: number[] = [];
  const date1904 = scope.workbook.date1904 === true;
  for (let r = scope.top; r <= scope.bottom; r++) {
    const row = scope.sheet.rows.get(r);
    const value = row ? numberOf(row.cells[col], date1904) : undefined;
    if (value !== undefined) out.push(value);
  }
  return out;
}

/** The test a criterion makes; null for a criterion this engine cannot apply, undefined for none. */
function matcherOf(scope: Scope, column: FilterColumn, col: number): Matcher | null | undefined {
  const date1904 = scope.workbook.date1904 === true;
  switch (column.kind) {
    case 'none':
      return undefined;
    case 'values': {
      const wanted = new Set<string>();
      const { reader, english, fold } = scope;
      for (const value of column.values || []) wanted.add(fold(String(value)));
      const blank = column.blank === true;
      const groups = column.dateGroups || [];
      // UNSURE: whether a file states numbers (and TRUE and FALSE) the way its author's locale shows them or in English.
      const inEnglish = (held: Cell) => english !== reader && wanted.size > 0 && wanted.has(fold(english.text(held)));
      return (cell) => {
        if (isBlank(cell)) return blank;
        const held = cell as Cell;
        if (wanted.has(fold(reader.text(held)))) return true;
        if (held.kind === 'boolean') return inEnglish(held);
        const number = numberOf(held, date1904);
        if (number === undefined) return false;
        for (let i = 0; i < groups.length; i++) if (matchesDateGroup(groups[i], number, date1904)) return true;
        return inEnglish(held);
      };
    }
    case 'custom': {
      const tests = (column.conditions || []).map((condition) => conditionMatcher(scope, condition));
      if (tests.length === 0) return undefined;
      const and = column.and === true;
      return (cell) => {
        for (let i = 0; i < tests.length; i++) {
          const ok = tests[i](cell);
          if (and ? !ok : ok) return !and;
        }
        return and;
      };
    }
    case 'top10': {
      const value = column.value;
      if (value === undefined || !(value > 0)) return null;
      const numbers = columnNumbers(scope, col);
      if (numbers.length === 0) return () => false;
      const top = column.top !== false;
      numbers.sort((a, b) => (top ? b - a : a - b));
      const wanted = column.percent ? Math.floor((numbers.length * Math.min(value, 100)) / 100) : Math.floor(value);
      const limit = numbers[Math.min(Math.max(wanted, 1), numbers.length) - 1];
      return (cell) => {
        const held = numberOf(cell, date1904);
        return held !== undefined && (top ? held >= limit : held <= limit);
      };
    }
    case 'dynamic': {
      const above = column.dynamicType === 'aboveAverage';
      if (!above && column.dynamicType !== 'belowAverage') return null;
      const numbers = columnNumbers(scope, col);
      let sum = 0;
      for (let i = 0; i < numbers.length; i++) sum += numbers[i];
      const average = numbers.length > 0 ? sum / numbers.length : 0;
      return (cell) => {
        const held = numberOf(cell, date1904);
        return held !== undefined && (above ? held > average : held < average);
      };
    }
    default:
      // A colour, an icon, or something unknown.
      return null;
  }
}

const hasCell = (row: Row | undefined): boolean => row !== undefined && row.cells.some(() => true);

/** Rows looked at one by one from the bottom before the rows of the sheet are gone through instead. */
const DESCENT_LIMIT = 1024;

/**
 * The index of the last row that holds a cell, at or above `from`; -1 when there is none.
 * Sheet.rowCount is not that: it is where the cells of the sheet once ended, and it does not
 * shrink when cells are cleared, while the file saved from the sheet has no rows there any
 * more. What a filter does must not depend on which of the two it is asked of, so it measures
 * by the cells there are. A row without cells (one that only has a height) holds nothing;
 * a cell that only carries a format is a cell, as in the file.
 */
export function lastFilledRow(sheet: Sheet, from: number = MAX_ROWS - 1): number {
  const start = Math.min(Math.floor(from), sheet.rowCount - 1, MAX_ROWS - 1);
  if (!(start >= 0)) return -1;
  for (let r = start, steps = 0; r >= 0 && steps < DESCENT_LIMIT; r--, steps++) if (hasCell(sheet.rows.get(r))) return r;
  if (start < DESCENT_LIMIT) return -1;
  // A long empty stretch (a sheet that was emptied): the rows there are, whatever their order.
  let last = -1;
  for (const [index, row] of sheet.rows) if (index > last && index <= start && hasCell(row)) last = index;
  return last;
}

/** The rows under the header of a filter that can hold cells; undefined when there are none. */
export function filterDataRows(sheet: Sheet, filter: AutoFilter): { top: number; bottom: number } | undefined {
  const top = filter.range.top + 1;
  const bottom = lastFilledRow(sheet, filter.range.bottom);
  return top <= bottom ? { top, bottom } : undefined;
}

interface Evaluation {
  top: number;
  bottom: number;
  /** One entry per row from `top`: 1 when the row passes every criterion that can be applied. */
  pass: Uint8Array;
  /** Some criterion cannot be applied (a colour, an icon, an unknown kind). */
  opaque: boolean;
}

function scopeOf(workbook: Workbook, sheet: Sheet, rows: { top: number; bottom: number }, locale: FormatLocale | undefined): Scope {
  const reader = new TextReader(workbook, locale);
  // In English (and without a locale, which reads as English) there is one way to show a cell.
  const english = locale === undefined || locale === LOCALE_EN ? reader : new TextReader(workbook, LOCALE_EN);
  return { workbook, sheet, locale, fold: (text) => foldCase(text, locale), reader, english, top: rows.top, bottom: rows.bottom };
}

/** Which rows pass the criteria of a filter; `skip` leaves the criterion of one column out. */
function evaluate(workbook: Workbook, sheet: Sheet, filter: AutoFilter, locale: FormatLocale | undefined, skip = -1): Evaluation | undefined {
  const rows = filterDataRows(sheet, filter);
  if (!rows) return undefined;
  const scope = scopeOf(workbook, sheet, rows, locale);
  const cols: number[] = [];
  const tests: Matcher[] = [];
  let opaque = false;
  for (const column of filter.columns) {
    const col = filter.range.left + column.colId;
    if (column.colId === skip || col > filter.range.right || col >= MAX_COLS) continue;
    const test = matcherOf(scope, column, col);
    if (test === null) opaque = true;
    else if (test) {
      cols.push(col);
      tests.push(test);
    }
  }
  const pass = new Uint8Array(rows.bottom - rows.top + 1);
  for (let r = rows.top; r <= rows.bottom; r++) {
    const row = sheet.rows.get(r);
    let ok = 1;
    for (let i = 0; i < tests.length; i++) {
      if (!tests[i](row ? row.cells[cols[i]] : undefined)) {
        ok = 0;
        break;
      }
    }
    pass[r - rows.top] = ok;
  }
  return { top: rows.top, bottom: rows.bottom, pass, opaque };
}

/** The hidden rows under the header of a filter, wherever they are. */
function hiddenRows(sheet: Sheet, filter: AutoFilter): number[] {
  const out: number[] = [];
  const { top, bottom } = filter.range;
  for (const [index, row] of sheet.rows) if (row.hidden && index > top && index <= bottom) out.push(index);
  return out.sort((a, b) => a - b);
}

function sheetOf(workbook: Workbook, sheetIndex: number): Sheet {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet) throw new RangeError(`No sheet at index ${sheetIndex}`);
  return sheet;
}

/**
 * Switches the filter of a sheet on over a range (its first row is the header) or, with
 * undefined, off. An earlier filter is replaced; the rows it hid are shown again.
 */
export function setAutoFilter(workbook: Workbook, sheetIndex: number, range: CellRange | undefined): Op[] {
  const sheet = sheetOf(workbook, sheetIndex);
  const current = sheet.autoFilter;
  if (!range && !current) return [];
  const show = current ? hiddenRows(sheet, current) : [];
  let next: AutoFilter | undefined;
  if (range) {
    // A table brings a filter of its own, kept as the file has it; a second one over its cells would not be valid.
    if ((sheet.anchored || []).indexOf('table') >= 0) throw new RangeError('Cannot filter in a sheet with table: not supported yet');
    const top = Math.max(Math.floor(range.top), 0);
    const left = Math.max(Math.floor(range.left), 0);
    // Whole columns are cut down to the rows that hold something.
    const bottom = Math.min(Math.floor(range.bottom), Math.max(lastFilledRow(sheet), top), MAX_ROWS - 1);
    const right = Math.min(Math.floor(range.right), MAX_COLS - 1);
    if (!(top <= bottom && left <= right)) throw new RangeError('Invalid filter range');
    next = { range: { top, left, bottom, right }, columns: [] };
  }
  return [{ type: 'setAutoFilter', sheet: sheetIndex, autoFilter: next, show }];
}

/** The operation that puts a changed filter in place, with the rows that change with it. */
function applyColumns(workbook: Workbook, sheetIndex: number, sheet: Sheet, current: AutoFilter, next: AutoFilter, locale: FormatLocale | undefined): Op[] {
  const show: number[] = [];
  const hide: number[] = [];
  const after = evaluate(workbook, sheet, next, locale);
  delete next.held;
  if (after) {
    // Rows that are hidden although the criteria that can be applied let them pass were hidden by
    // one that cannot (or by hand). While such a criterion is still there they stay hidden; they
    // are noted in the filter, because once another criterion hides them too they cannot be told apart.
    let held: Set<number> | undefined;
    if (after.opaque) {
      held = new Set(current.held || []);
      const before = evaluate(workbook, sheet, current, locale);
      for (let r = after.top; before && r <= after.bottom; r++) {
        const row = sheet.rows.get(r);
        if (row && row.hidden && before.pass[r - before.top] === 1) held.add(r);
      }
      next.held = Array.from(held).sort((a, b) => a - b);
    }
    for (let r = after.top; r <= after.bottom; r++) {
      const row = sheet.rows.get(r);
      const hidden = row ? row.hidden : false;
      const wanted = after.pass[r - after.top] === 0 || (held !== undefined && held.has(r));
      if (wanted !== hidden) (wanted ? hide : show).push(r);
    }
  }
  return [{ type: 'setAutoFilter', sheet: sheetIndex, autoFilter: next, show, hide }];
}

/**
 * Sets what the filter asks of one column (counted from the left edge of the filter
 * range); undefined removes the criterion. The rows under the header are shown and
 * hidden to match all criteria: one undo step.
 */
export function setFilterColumn(workbook: Workbook, sheetIndex: number, colId: number, criterion: FilterCriterion | undefined, options: EditOptions = {}): Op[] {
  const sheet = sheetOf(workbook, sheetIndex);
  const current = sheet.autoFilter;
  if (!current) throw new RangeError('The sheet has no filter');
  if (!(colId >= 0 && colId <= current.range.right - current.range.left) || Math.floor(colId) !== colId) throw new RangeError('The column is outside the filter');
  const previous = current.columns.find((column) => column.colId === colId);
  const columns = current.columns.filter((column) => column.colId !== colId);
  if (criterion && criterion.kind !== 'none') {
    const column: FilterColumn = { ...criterion, colId };
    if (previous && previous.hiddenButton && column.hiddenButton === undefined) column.hiddenButton = true;
    columns.push(column);
  } else if ((criterion && criterion.hiddenButton) || (!criterion && previous && previous.hiddenButton)) {
    columns.push({ colId, kind: 'none', hiddenButton: true });
  } else if (!previous) return [];
  columns.sort((a, b) => a.colId - b.colId);
  return applyColumns(workbook, sheetIndex, sheet, current, { ...current, columns }, options.locale);
}

/** Removes every criterion and keeps the filter: all rows under the header are shown. */
export function clearFilter(workbook: Workbook, sheetIndex: number, options: EditOptions = {}): Op[] {
  const sheet = sheetOf(workbook, sheetIndex);
  const current = sheet.autoFilter;
  if (!current || !current.columns.some((column) => column.kind !== 'none')) return [];
  const columns = current.columns.filter((column) => column.hiddenButton).map((column): FilterColumn => ({ colId: column.colId, kind: 'none', hiddenButton: true }));
  return applyColumns(workbook, sheetIndex, sheet, current, { ...current, columns }, options.locale);
}

/**
 * The distinct values a column of the filter offers, as its cells show them, with
 * whether the column's criterion shows them now. Only rows that the criteria of the
 * other columns let pass are looked at, as in Excel.
 */
export function filterValues(workbook: Workbook, sheetIndex: number, colId: number, options: EditOptions = {}): FilterValueList {
  const sheet = sheetOf(workbook, sheetIndex);
  const filter = sheet.autoFilter;
  const list: FilterValueList = { items: [], blanks: false, blanksChecked: false, truncated: false, opaque: false };
  if (!filter) return list;
  const col = filter.range.left + colId;
  const others = evaluate(workbook, sheet, filter, options.locale, colId);
  if (!others || !(col >= filter.range.left && col <= filter.range.right)) return list;
  const scope = scopeOf(workbook, sheet, others, options.locale);
  const own = filter.columns.find((column) => column.colId === colId);
  const test = own ? matcherOf(scope, own, col) : undefined;
  list.opaque = test === null;
  const date1904 = workbook.date1904 === true;

  const found = new Map<string, { text: string; checked: boolean; rank: number; number: number; key: string }>();
  for (let r = others.top; r <= others.bottom; r++) {
    if (others.pass[r - others.top] === 0) continue;
    const row = sheet.rows.get(r);
    const cell = row ? row.cells[col] : undefined;
    // A criterion that cannot be applied says nothing about a value: what is shown counts as chosen.
    const checked = test ? test(cell) : test === null ? !(row && row.hidden) : true;
    if (isBlank(cell)) {
      list.blanks = true;
      if (checked) list.blanksChecked = true;
      continue;
    }
    const text = scope.reader.text(cell as Cell);
    const id = scope.fold(text);
    const known = found.get(id);
    if (known) {
      if (checked) known.checked = true;
      continue;
    }
    if (found.size >= MAX_FILTER_VALUES) {
      list.truncated = true;
      continue;
    }
    const value = sortValueOf(cell, date1904);
    found.set(id, { text, checked, rank: value.rank === SORT_BLANK ? SORT_TEXT : value.rank, number: value.number, key: value.rank === SORT_TEXT ? collationKey(value.text, options.locale) : '' });
  }
  const entries = Array.from(found.values());
  entries.sort((a, b) => a.rank - b.rank || (a.number < b.number ? -1 : a.number > b.number ? 1 : 0) || (a.key < b.key ? -1 : a.key > b.key ? 1 : 0) || (a.text < b.text ? -1 : a.text > b.text ? 1 : 0));
  list.items = entries.map((entry) => ({ text: entry.text, checked: entry.checked }));
  return list;
}

/**
 * The rows of a span that the filter of the sheet hides: hidden rows under the header of a
 * filter that has a criterion. As in Excel, what is done to a selection (clearing, copying,
 * formatting, deleting rows) leaves these rows out, while rows hidden by hand stay part of
 * it. The model has one flag for both ways of hiding, so in a filtered list every hidden row
 * counts as filtered. Sorted; empty for a sheet without a filter in use.
 */
export function rowsHiddenByFilter(sheet: Sheet, top: number, bottom: number): number[] {
  const out: number[] = [];
  const filter = sheet.autoFilter;
  if (!filter || !filter.columns.some(isFilterActive)) return out;
  const first = Math.max(top, filter.range.top + 1);
  const last = Math.min(bottom, filter.range.bottom);
  if (first > last) return out;
  if (last - first < sheet.rows.size) {
    for (let r = first; r <= last; r++) {
      const row = sheet.rows.get(r);
      if (row && row.hidden) out.push(r);
    }
    return out;
  }
  for (const [index, row] of sheet.rows) if (row.hidden && index >= first && index <= last) out.push(index);
  return out.sort((a, b) => a - b);
}

/** The runs of rows of a span that are left when some rows (sorted) are taken out of it. */
export function rowRunsWithout(top: number, bottom: number, without: number[]): { top: number; bottom: number }[] {
  const runs: { top: number; bottom: number }[] = [];
  let next = top;
  for (let i = 0; i < without.length; i++) {
    const row = without[i];
    if (row < next || row > bottom) continue;
    if (row > next) runs.push({ top: next, bottom: row - 1 });
    next = row + 1;
  }
  if (next <= bottom) runs.push({ top: next, bottom });
  return runs;
}

/**
 * Deletes rows as Excel does in a filtered list: the rows of the span that the filter hides
 * stay. What is left to delete may be several runs of rows; they are deleted from the last
 * one up, and applied together they are one undo step. Without filtered rows in the span
 * this is the one "delete" operation.
 */
export function deleteRows(workbook: Workbook, sheetIndex: number, at: number, count: number): Op[] {
  const sheet = sheetOf(workbook, sheetIndex);
  const kept = count > 0 ? rowsHiddenByFilter(sheet, at, at + count - 1) : [];
  if (kept.length === 0) return [{ type: 'delete', sheet: sheetIndex, axis: 'row', at, count }];
  const ops: Op[] = [];
  const runs = rowRunsWithout(at, at + count - 1, kept);
  for (let i = runs.length - 1; i >= 0; i--) ops.push({ type: 'delete', sheet: sheetIndex, axis: 'row', at: runs[i].top, count: runs[i].bottom - runs[i].top + 1 });
  return ops;
}

/**
 * Deletes columns. A criterion of the filter on one of them goes with its column, so the
 * rows are first shown and hidden by the criteria that stay: rows that only the deleted
 * column's criterion hid come back. Applied together the operations are one undo step.
 */
export function deleteColumns(workbook: Workbook, sheetIndex: number, at: number, count: number, options: EditOptions = {}): Op[] {
  const sheet = sheetOf(workbook, sheetIndex);
  const ops: Op[] = [];
  const current = sheet.autoFilter;
  if (current && count > 0) {
    const gone = (column: FilterColumn) => {
      const col = current.range.left + column.colId;
      return col >= at && col < at + count;
    };
    if (current.columns.some((column) => gone(column) && isFilterActive(column))) {
      const columns = current.columns.filter((column) => !gone(column));
      ops.push(...applyColumns(workbook, sheetIndex, sheet, current, { ...current, columns }, options.locale));
    }
  }
  ops.push({ type: 'delete', sheet: sheetIndex, axis: 'col', at, count });
  return ops;
}

/** Whether the filter asks anything of a column, for the look of its button. */
export function isFilterActive(column: FilterColumn | undefined): boolean {
  return column !== undefined && column.kind !== 'none';
}

/** How many rows under the header a filter with criteria shows; undefined for a filter without any. */
export function filterSummary(sheet: Sheet): { shown: number; total: number } | undefined {
  const filter = sheet.autoFilter;
  if (!filter || !filter.columns.some(isFilterActive)) return undefined;
  const rows = filterDataRows(sheet, filter);
  if (!rows) return { shown: 0, total: 0 };
  const total = rows.bottom - rows.top + 1;
  let hidden = 0;
  for (const [index, row] of sheet.rows) if (row.hidden && index >= rows.top && index <= rows.bottom) hidden++;
  return { shown: total - hidden, total };
}
