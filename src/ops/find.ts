// Find and Replace in a workbook. A search walks the cells in Excel's order (by
// rows: left to right, then down) from a cell, round the end of the sheet (and
// of the workbook) back to where it started, a few cells at a time: the caller
// says how many per call (SheetSearch.step), so that a search through a large
// workbook never holds the page up.
//
// What a cell is searched as: with "formulas" (the default, and always for
// replacing, as in Excel) what the formula bar shows - the formula, or the
// value as it is typed (1234.5, not "1,234.50") - and text as it is, without
// the apostrophe that keeps it text; with "values" the text the cell shows.
// Hidden rows and columns are searched for formulas and passed over for
// values (Excel does the same).
//
// Replacing writes the cell anew the way typing would (setCellInput), except
// that text stays text: a text cell is given its new text as text, and the
// formatting of rich text stays with the characters that were not replaced
// (the replacement takes that of the first character it replaces). Cells that
// cannot be written so are passed over and counted: a part of an array
// formula, a formula that cannot be read once replaced, text that grows too
// long. "Replace all" leaves rows the filter hides alone, as clearing does.

import { compileFind, findIn, foldFor, matchesIn, replaceIn, type CompiledFind, type FindQuery } from '../edit/find';
import { cellEditText } from '../edit/input';
import { formulaToLocal } from '../edit/formula-locale';
import { cellDisplayText } from '../format/cell-format';
import type { FormatLocale } from '../format/locale';
import { resolveCellStyle } from '../model/sheet/color';
import { MAX_COLS, MAX_ROWS, fillsRange, parseRange, type Cell, type CellRange, type RichRun, type Sheet, type Workbook } from '../model/sheet/types';
import { setCellInput } from './commands';
import { rowsHiddenByFilter } from './filter';
import { formulaHidden } from './protection';
import { explicitFormula, sharedMasters, type CellChange, type Op, type SharedMasters } from './ops';

export interface SheetFindOptions {
  matchCase?: boolean;
  /** The text must be the whole content of the cell. */
  entireCell?: boolean;
  /** 'formulas' (the default): what the formula bar shows; 'values': the text the cell shows. Replacing always looks at formulas. */
  lookIn?: 'formulas' | 'values';
  /** 'sheet' (the default): the sheet the search starts on; 'workbook': every worksheet, from that one on. */
  within?: 'sheet' | 'workbook';
  /** Only these cells of the sheet the search starts on (a selection of several cells, as in Excel); `within` is then 'sheet'. */
  range?: CellRange;
  /** How numbers, dates and formulas are written, and how case is folded. */
  locale?: FormatLocale;
}

export interface CellAddress {
  sheet: number;
  row: number;
  col: number;
}

/** Most cells "Replace all" writes in one step. */
export const MAX_REPLACE_CELLS = 200000;
/** Longest text a cell holds. */
const MAX_CELL_TEXT = 32767;

/** A query for cells: undefined for an empty text. Throws RangeError for one that is not valid (see compileFind). */
export function compileSheetFind(text: string, options: SheetFindOptions = {}): CompiledFind | undefined {
  const query: FindQuery = { text, matchCase: options.matchCase === true, locale: options.locale };
  if (options.entireCell === true) query.whole = 'all';
  return compileFind(query);
}

/** What is known of a sheet while it is searched. */
interface SheetContext {
  sheet: Sheet;
  masters?: SharedMasters;
  hiddenCols?: (col: number) => boolean;
}

function hiddenColumns(sheet: Sheet): (col: number) => boolean {
  const hidden = sheet.columns.filter((range) => range.hidden);
  if (hidden.length === 0) return () => false;
  return (col: number) => {
    for (let i = 0; i < hidden.length; i++) if (col >= hidden[i].min - 1 && col <= hidden[i].max - 1) return true;
    return false;
  };
}

/** The text a cell is searched as (see the head of this file); '' for a cell without any. */
export function cellSearchText(workbook: Workbook, sheet: Sheet, row: number, col: number, cell: Cell | undefined, lookIn: 'formulas' | 'values' = 'formulas', locale?: FormatLocale, masters?: SharedMasters): string {
  if (!cell) return '';
  if (lookIn === 'values') return cellDisplayText(workbook, cell, locale);
  const formula = cell.formula;
  // A formula the protection of the sheet hides is no text of the formula bar (which shows nothing): it is not
  // searched, and so not replaced either. Its value is still found among the values.
  if (formula && formulaHidden(workbook, sheet, row, col)) return '';
  if (formula) {
    let text = formula.text;
    if (text === '' && formula.type === 'shared') {
      const explicit = explicitFormula(cell, row, col, masters || sharedMasters(sheet));
      text = explicit.formula ? explicit.formula.text : '';
    }
    if (text !== '') return '=' + formulaToLocal(text, locale);
  }
  if (cell.value === null || cell.value === undefined) return '';
  // Text as it is: the apostrophe that keeps "123" text in the editor is not part of it.
  if (cell.kind === 'string' && typeof cell.value === 'string') return cell.value;
  if (cell.kind === 'error') return String(cell.value);
  return cellEditText(cell, resolveCellStyle(workbook.styles, cell.style).numFmt, { locale, date1904: workbook.date1904 });
}

/** The row indices of a sheet that hold cells, in order (rows are usually read in order: then nothing is sorted). */
function rowsInOrder(sheet: Sheet, top: number, bottom: number): number[] {
  const out: number[] = [];
  let sorted = true;
  for (const index of sheet.rows.keys()) {
    if (index < top || index > bottom) continue;
    if (out.length > 0 && index < out[out.length - 1]) sorted = false;
    out.push(index);
  }
  if (!sorted) out.sort((a, b) => a - b);
  return out;
}

const before = (r1: number, c1: number, r2: number, c2: number): boolean => r1 < r2 || (r1 === r2 && c1 < c2);

/** One cell the walk comes to. */
interface Visit {
  sheet: number;
  row: number;
  col: number;
  cell: Cell;
}

/**
 * The cells of one sheet in order (or in reverse), within `range`, and only those after
 * `after` (in the direction of the walk) and up to and including `upTo`, when given.
 */
function* sheetCells(workbook: Workbook, index: number, range: CellRange | undefined, backward: boolean, after?: { row: number; col: number }, upTo?: { row: number; col: number }): Generator<Visit | number> {
  const sheet = workbook.sheets[index];
  if (!sheet || sheet.kind !== 'worksheet') return;
  const area = range || { top: 0, left: 0, bottom: MAX_ROWS - 1, right: MAX_COLS - 1 };
  const rows = rowsInOrder(sheet, area.top, area.bottom);
  // Sorting a long list costs about as much as looking at its cells: it counts as work done.
  yield rows.length;
  for (let k = 0; k < rows.length; k++) {
    const r = rows[backward ? rows.length - 1 - k : k];
    if (after && (backward ? r > after.row : r < after.row)) continue;
    if (upTo && (backward ? r < upTo.row : r > upTo.row)) break;
    const row = sheet.rows.get(r);
    if (!row) continue;
    const cells = row.cells;
    const last = Math.min(cells.length - 1, area.right);
    for (let j = 0; j <= last - area.left; j++) {
      const c = backward ? last - j : area.left + j;
      const cell = cells[c];
      if (cell === undefined) continue;
      if (after && !(backward ? before(r, c, after.row, after.col) : before(after.row, after.col, r, c))) continue;
      if (upTo && (backward ? before(r, c, upTo.row, upTo.col) : before(upTo.row, upTo.col, r, c))) {
        if (r === upTo.row) break;
        continue;
      }
      yield { sheet: index, row: r, col: c, cell };
    }
  }
}

/**
 * Every cell of the search in order, starting right after `from` and ending with `from`
 * itself, round the end of the sheet (and of the workbook). A number stands for work done
 * without a cell to show for it.
 */
function* walk(workbook: Workbook, from: CellAddress, backward: boolean, options: SheetFindOptions): Generator<Visit | number> {
  const range = options.range;
  const whole = options.within === 'workbook' && !range;
  const at = { row: from.row, col: from.col };
  yield* sheetCells(workbook, from.sheet, range, backward, at);
  if (whole) {
    const count = workbook.sheets.length;
    for (let k = 1; k < count; k++) {
      const index = (from.sheet + (backward ? count - k : k)) % count;
      // Hidden sheets are not searched from another sheet (the one the search starts on is, whatever it is).
      const sheet = workbook.sheets[index];
      if (sheet && sheet.state === 'visible') yield* sheetCells(workbook, index, undefined, backward);
    }
  }
  yield* sheetCells(workbook, from.sheet, range, backward, undefined, at);
}

/**
 * A search that goes on where it stopped: from a cell, in one direction, round to that
 * cell again. Made for a workbook as it is; the caller starts a new one after a change.
 */
export class SheetSearch {
  private readonly cells: Generator<Visit | number>;
  private readonly contexts = new Map<number, SheetContext>();
  private done = false;
  /** About how many cells the last call of step() looked at. */
  used = 0;

  constructor(
    private readonly workbook: Workbook,
    private readonly find: CompiledFind,
    private readonly options: SheetFindOptions,
    from: CellAddress,
    backward = false,
  ) {
    this.cells = walk(workbook, from, backward, options);
  }

  /** Whether the search went all the way round. */
  get finished(): boolean {
    return this.done;
  }

  /**
   * Looks at about `budget` cells more. The next cell that matches; null when the search went
   * all the way round without another; undefined when the budget was used up first.
   */
  step(budget: number): CellAddress | null | undefined {
    this.used = 0;
    if (this.done) return null;
    const limit = Math.max(1, budget);
    while (this.used < limit) {
      const next = this.cells.next();
      if (next.done) {
        this.done = true;
        return null;
      }
      const visit = next.value;
      if (typeof visit === 'number') {
        this.used += Math.max(1, visit);
        continue;
      }
      this.used++;
      if (this.matches(visit)) return { sheet: visit.sheet, row: visit.row, col: visit.col };
    }
    return undefined;
  }

  private context(index: number): SheetContext {
    let context = this.contexts.get(index);
    if (!context) {
      context = { sheet: this.workbook.sheets[index] };
      this.contexts.set(index, context);
    }
    return context;
  }

  private matches(visit: Visit): boolean {
    const context = this.context(visit.sheet);
    const sheet = context.sheet;
    const lookIn = this.options.lookIn === 'values' ? 'values' : 'formulas';
    if (lookIn === 'values') {
      // What does not show is not a value that is seen.
      const row = sheet.rows.get(visit.row);
      if (row && row.hidden) return false;
      if (!context.hiddenCols) context.hiddenCols = hiddenColumns(sheet);
      if (context.hiddenCols(visit.col)) return false;
    } else if (!context.masters && visit.cell.formula && visit.cell.formula.type === 'shared') context.masters = sharedMasters(sheet);
    const text = cellSearchText(this.workbook, sheet, visit.row, visit.col, visit.cell, lookIn, this.options.locale, context.masters);
    return text !== '' && findIn(this.find, text, 0, foldFor(this.find, text)) >= 0;
  }
}

/** What replacing does to one cell. */
export type CellReplacement = { change: CellChange; count: number } | { skipped: true; count: number } | undefined;

/** The array formulas of a sheet that cover more than their own cell. */
function arrayAreas(sheet: Sheet): CellRange[] {
  const out: CellRange[] = [];
  for (const row of sheet.rows.values()) {
    row.cells.forEach((cell) => {
      const formula = cell.formula;
      if (!fillsRange(formula)) return;
      const area = parseRange(formula.ref);
      if (area && (area.top !== area.bottom || area.left !== area.right)) out.push(area);
    });
  }
  return out;
}

/** Rich text with every match replaced: what is kept keeps its formatting, the replacement takes that of the run its match starts in. */
function replaceRuns(runs: RichRun[], starts: number[], length: number, replacement: string): RichRun[] {
  const out: RichRun[] = [];
  let m = 0;
  let cut = 0;
  let pos = 0;
  for (let i = 0; i < runs.length; i++) {
    const run = runs[i];
    const end = pos + run.text.length;
    let text = '';
    let at = Math.max(pos, cut);
    while (at < end) {
      if (m < starts.length && starts[m] < end) {
        text += run.text.slice(at - pos, Math.max(at, starts[m]) - pos) + replacement;
        at = starts[m] + length;
        cut = at;
        m++;
      } else {
        text += run.text.slice(at - pos);
        at = end;
      }
    }
    if (text !== '') out.push({ ...run, text });
    pos = end;
  }
  return out.length > 0 ? out : [{ ...runs[0], text: '' }];
}

/**
 * The change that replaces every match in one cell; undefined when the cell has none, and
 * `skipped` when it has some but cannot be written so (see the head of this file). The cell
 * is read the way the search reads formulas.
 */
export function replaceInCell(workbook: Workbook, sheetIndex: number, row: number, col: number, find: CompiledFind, replacement: string, locale?: FormatLocale, areas?: CellRange[], masters?: SharedMasters): CellReplacement {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet || sheet.kind !== 'worksheet') return undefined;
  const line = sheet.rows.get(row);
  const cell = line ? line.cells[col] : undefined;
  if (!cell) return undefined;
  const text = cellSearchText(workbook, sheet, row, col, cell, 'formulas', locale, masters);
  if (text === '') return undefined;
  const replaced = replaceIn(find, text, replacement);
  if (replaced.count === 0) return undefined;
  const count = replaced.count;
  const formula = cell.formula;
  if (formula && (formula.type === 'array' || formula.type === 'dataTable')) return { skipped: true, count };
  // The cells a dynamic array formula spills into are its result, no content: its own formula can be changed.
  const covered = formula && formula.dynamic ? [] : areas || arrayAreas(sheet);
  for (let i = 0; i < covered.length; i++) {
    const area = covered[i];
    if (row >= area.top && row <= area.bottom && col >= area.left && col <= area.right) return { skipped: true, count };
  }
  if (replaced.text.length > MAX_CELL_TEXT) return { skipped: true, count };
  const isText = !formula && cell.kind === 'string' && typeof cell.value === 'string';
  if (isText && replaced.text !== '') {
    // Text stays text, with the formatting of its characters; what else the cell carries (cell and value metadata,
    // whether it shows phonetic text) stays too. Only the shared string it was read from is another one now.
    const next: Cell = { ...cell, value: replaced.text };
    delete next.si;
    delete next.runs;
    if (cell.runs && cell.runs.length > 0) next.runs = replaceRuns(cell.runs, matchesIn(find, text), find.needle.length, replacement);
    return { change: { row, col, cell: next }, count };
  }
  let ops: Op[];
  try {
    ops = setCellInput(workbook, sheetIndex, row, col, replaced.text, { locale, masters });
  } catch {
    return { skipped: true, count };
  }
  const op = ops[0];
  if (!op || op.type !== 'setCells' || op.cells.length !== 1) return { skipped: true, count };
  return { change: op.cells[0], count };
}

/** What replaceAllOps() found and made. */
export interface ReplaceAllResult {
  ops: Op[];
  /** Matches replaced, and cells written. */
  replaced: number;
  cells: number;
  /** Matches in cells that could not be written (see replaceInCell). */
  skipped: number;
  /** More than MAX_REPLACE_CELLS cells would be written: nothing is (`ops` is empty). */
  tooMany?: true;
}

/**
 * Collects the changes of "Replace all", a few cells at a time: step() goes on until it is
 * done (true) or `budget` cells were looked at (false); result() then has the operations,
 * one setCells per sheet. Rows the filter of a sheet hides are left alone.
 */
export class ReplaceAll {
  private readonly search: SheetSearch;
  private readonly bySheet = new Map<number, CellChange[]>();
  private readonly areas = new Map<number, CellRange[]>();
  private readonly filtered = new Map<number, Set<number>>();
  private readonly masters = new Map<number, SharedMasters>();
  private replaced = 0;
  private cells = 0;
  private skipped = 0;
  private over = false;
  private finished = false;

  constructor(
    private readonly workbook: Workbook,
    private readonly find: CompiledFind,
    private readonly replacement: string,
    private readonly options: SheetFindOptions,
    from: CellAddress,
  ) {
    // Replacing always reads formulas. Starting before the first cell, the walk goes once through everything.
    this.search = new SheetSearch(workbook, find, { ...options, lookIn: 'formulas' }, from);
  }

  step(budget: number): boolean {
    if (this.finished) return true;
    let left = Math.max(1, budget);
    while (left > 0) {
      const hit = this.search.step(left);
      left -= this.search.used;
      if (hit === null) {
        this.finished = true;
        return true;
      }
      if (hit === undefined) return false;
      this.take(hit);
      if (this.over) {
        this.finished = true;
        return true;
      }
    }
    return false;
  }

  private take(hit: CellAddress): void {
    const sheet = this.workbook.sheets[hit.sheet];
    let filtered = this.filtered.get(hit.sheet);
    if (!filtered) {
      filtered = new Set(rowsHiddenByFilter(sheet, 0, MAX_ROWS - 1));
      this.filtered.set(hit.sheet, filtered);
    }
    if (filtered.has(hit.row)) return;
    let areas = this.areas.get(hit.sheet);
    if (!areas) {
      areas = arrayAreas(sheet);
      this.areas.set(hit.sheet, areas);
    }
    let masters = this.masters.get(hit.sheet);
    if (!masters) {
      masters = sharedMasters(sheet);
      this.masters.set(hit.sheet, masters);
    }
    const result = replaceInCell(this.workbook, hit.sheet, hit.row, hit.col, this.find, this.replacement, this.options.locale, areas, masters);
    if (!result) return;
    if ('skipped' in result) {
      this.skipped += result.count;
      return;
    }
    if (this.cells >= MAX_REPLACE_CELLS) {
      this.over = true;
      return;
    }
    let list = this.bySheet.get(hit.sheet);
    if (!list) {
      list = [];
      this.bySheet.set(hit.sheet, list);
    }
    list.push(result.change);
    this.cells++;
    this.replaced += result.count;
  }

  result(): ReplaceAllResult {
    if (this.over) return { ops: [], replaced: 0, cells: 0, skipped: 0, tooMany: true };
    const ops: Op[] = [];
    this.bySheet.forEach((cells, sheet) => ops.push({ type: 'setCells', sheet, cells }));
    return { ops, replaced: this.replaced, cells: this.cells, skipped: this.skipped };
  }
}

/** The address before the first cell of a sheet (or after its last, for a backward walk): a walk from it goes through everything. */
export function searchStart(sheet: number, backward = false): CellAddress {
  return backward ? { sheet, row: MAX_ROWS, col: MAX_COLS } : { sheet, row: -1, col: -1 };
}

/** Finds the next match in one go, for a workbook small enough not to need steps (and for Node). */
export function findNextCell(workbook: Workbook, text: string, from: CellAddress, options: SheetFindOptions = {}, backward = false): CellAddress | undefined {
  const find = compileSheetFind(text, options);
  if (!find) return undefined;
  const search = new SheetSearch(workbook, find, options, from, backward);
  for (;;) {
    const hit = search.step(1 << 20);
    if (hit !== undefined) return hit === null ? undefined : hit;
  }
}

/** "Replace all" in one go (see ReplaceAll); `from` defaults to before the first cell of sheet 0. */
export function replaceAllOps(workbook: Workbook, text: string, replacement: string, options: SheetFindOptions = {}, from?: CellAddress): ReplaceAllResult {
  const find = compileSheetFind(text, options);
  if (!find) return { ops: [], replaced: 0, cells: 0, skipped: 0 };
  const job = new ReplaceAll(workbook, find, replacement, options, from || searchStart(0));
  while (!job.step(1 << 20));
  return job.result();
}
