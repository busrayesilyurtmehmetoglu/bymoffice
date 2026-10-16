// The protection of sheets and of the structure of a workbook, applied: what a
// protected sheet lets the user do (ECMA-376 Part 1, 18.3.1.85 CT_SheetProtection)
// and what a workbook with a locked structure does not (18.2.29 lockStructure).
//
// One rule decides it for every change: protectionProblem() looks at the
// operations a command made, before they are applied, and says why they may
// not be. The view asks it for every step it takes (typing, paste, Delete,
// Find and Replace, sort, filter, formats, rows and columns, notes, rules,
// drawings, sheets), so that what reaches the workbook through any of them is
// refused the same way. Steps of others, edited together, are not asked about:
// see the note on collaborative editing in docs/kullanim-kilavuzu.md.
//
// Also here: protecting a sheet and removing its protection ("setProtection"),
// with the password checked as Excel checks it (the iterated SHA hash of
// algorithmName/hashValue/saltValue/spinCount, or the 16-bit hash of the
// legacy `password` attribute).

import { protectionNow } from '../formats/xlsx/build';
import { fromBase64, toBase64, randomBytes, utf16le, constantTimeEqual } from '../crypto/bytes';
import { HASH_NAMES, type HashName } from '../crypto/hash';
import { MAX_SPIN_COUNT, spinHash } from '../crypto/work';
import { createHash } from '../crypto/hash';
import { MAX_COLS, MAX_ROWS, getCell, parseRange, type Cell, type CellRange, type ColumnRange, type Sheet, type SheetProtection, type Workbook } from '../model/sheet/types';
import type { Op, RowProps } from './ops';

/**
 * The attributes of <sheetProtection> that name something a protected sheet may forbid. Each
 * one set to true forbids it; Excel writes "1" for every box left unticked.
 */
export type ProtectionFlag =
  | 'objects'
  | 'scenarios'
  | 'formatCells'
  | 'formatColumns'
  | 'formatRows'
  | 'insertColumns'
  | 'insertRows'
  | 'insertHyperlinks'
  | 'deleteColumns'
  | 'deleteRows'
  | 'selectLockedCells'
  | 'sort'
  | 'autoFilter'
  | 'pivotTables'
  | 'selectUnlockedCells';

export const PROTECTION_FLAGS: readonly ProtectionFlag[] = [
  'objects',
  'scenarios',
  'formatCells',
  'formatColumns',
  'formatRows',
  'insertColumns',
  'insertRows',
  'insertHyperlinks',
  'deleteColumns',
  'deleteRows',
  'selectLockedCells',
  'sort',
  'autoFilter',
  'pivotTables',
  'selectUnlockedCells',
];

/**
 * What an attribute that is not there says, from the schema (the default of each attribute of
 * CT_SheetProtection): true forbids. `objects` and `scenarios` default to false - a sheet
 * protected without them leaves its objects free - and so do the two kinds of selection; every
 * other one defaults to true. openpyxl has the same defaults. UNSURE: how Excel itself reads a
 * file that leaves out `objects` was not checked (Excel always writes it).
 */
export const FLAG_DEFAULTS: { readonly [flag in ProtectionFlag]: boolean } = {
  objects: false,
  scenarios: false,
  formatCells: true,
  formatColumns: true,
  formatRows: true,
  insertColumns: true,
  insertRows: true,
  insertHyperlinks: true,
  deleteColumns: true,
  deleteRows: true,
  selectLockedCells: false,
  sort: true,
  autoFilter: true,
  pivotTables: true,
  selectUnlockedCells: false,
};

/** Why a step is refused: what the protection forbids, as protectionProblem() finds it. */
export type ProtectionReason =
  /** A locked cell would change (its value, its formula), or one that a sort or a deletion moves. */
  | 'cells'
  | 'formatCells'
  | 'formatColumns'
  | 'formatRows'
  | 'insertColumns'
  | 'insertRows'
  | 'deleteColumns'
  | 'deleteRows'
  | 'sort'
  | 'autoFilter'
  | 'pivotTables'
  | 'objects'
  /** Data validations and conditional formats (other than those a paste or a cut carries). */
  | 'rules'
  | 'merge'
  /** The format of a cell would be locked or unlocked, or its formula hidden or shown. */
  | 'lock'
  /** The structure of the workbook is locked: sheets are not added, deleted, renamed, moved, copied, hidden or shown. */
  | 'structure';

/** A refusal of the protection; `code` is "protected" for whoever tells refusals apart (the bridge). */
export class ProtectedError extends RangeError {
  readonly code = 'protected';
  constructor(
    readonly reason: ProtectionReason,
    /** The sheet the step was refused on (its index then); undefined for the structure of the workbook. */
    readonly sheet?: number,
  ) {
    super(PROTECTION_MESSAGES[reason]);
    this.name = 'ProtectedError';
  }
}

export const PROTECTION_MESSAGES: { readonly [reason in ProtectionReason]: string } = {
  cells: 'The cell or range is on a protected sheet: to change it, remove the protection of the sheet',
  formatCells: 'The sheet is protected: its cells cannot be formatted',
  formatColumns: 'The sheet is protected: its columns cannot be resized, hidden or shown',
  formatRows: 'The sheet is protected: its rows cannot be resized, hidden or shown',
  insertColumns: 'The sheet is protected: columns cannot be inserted',
  insertRows: 'The sheet is protected: rows cannot be inserted',
  deleteColumns: 'The sheet is protected: these columns cannot be deleted',
  deleteRows: 'The sheet is protected: these rows cannot be deleted',
  sort: 'The sheet is protected: this range cannot be sorted',
  autoFilter: 'The sheet is protected: its filter cannot be changed',
  pivotTables: 'The sheet is protected: its pivot tables cannot be changed',
  objects: 'The sheet is protected: its notes, pictures, shapes and charts cannot be changed',
  rules: 'The sheet is protected: its data validations and conditional formats cannot be changed',
  merge: 'The sheet is protected: cells cannot be merged or unmerged',
  lock: 'The sheet is protected: cells cannot be locked or unlocked, nor their formulas hidden',
  structure: 'The structure of the workbook is protected: sheets cannot be added, deleted, renamed, moved, copied, hidden or shown',
};

// ------------------------------------------------------------------ reading the flags

const trueValue = (value: string) => value === '1' || value === 'true';
const falseValue = (value: string) => value === '0' || value === 'false';

function attrOf(attrs: readonly [string, string][] | undefined, name: string): string | undefined {
  if (!attrs) return undefined;
  for (let i = 0; i < attrs.length; i++) if (attrs[i][0] === name) return attrs[i][1];
  return undefined;
}

/** Whether a sheet is protected: <sheetProtection sheet="1">. */
export function sheetProtected(sheet: Sheet | undefined): boolean {
  if (!sheet || !sheet.protection) return false;
  const value = attrOf(sheet.protection.attrs, 'sheet');
  return value !== undefined && trueValue(value);
}

/**
 * Whether a protected sheet forbids something; false for a sheet that is not protected. A value
 * that is neither true nor false ("yes") is read as the default.
 */
export function protectionForbids(sheet: Sheet | undefined, flag: ProtectionFlag): boolean {
  if (!sheetProtected(sheet)) return false;
  const value = attrOf((sheet as Sheet).protection!.attrs, flag);
  if (value === undefined) return FLAG_DEFAULTS[flag];
  if (trueValue(value)) return true;
  if (falseValue(value)) return false;
  return FLAG_DEFAULTS[flag];
}

/** Whether the protection of a sheet asks for a password to be removed. */
export function protectionHasPassword(sheet: Sheet | undefined): boolean {
  const attrs = sheet && sheet.protection && sheet.protection.attrs;
  if (!attrs) return false;
  const legacy = attrOf(attrs, 'password');
  return (legacy !== undefined && legacy !== '' && !/^0+$/.test(legacy)) || attrOf(attrs, 'hashValue') !== undefined;
}

/** Whether the structure of a workbook is locked: <workbookProtection lockStructure="1">. */
export function structureLocked(workbook: Workbook): boolean {
  const value = attrOf(workbook.protection && workbook.protection.attrs, 'lockStructure');
  return value !== undefined && trueValue(value);
}

/** Why the list of sheets of a workbook cannot be changed; undefined when it can. */
export function structureProblem(workbook: Workbook): string | undefined {
  return structureLocked(workbook) ? PROTECTION_MESSAGES.structure : undefined;
}

// ------------------------------------------------------------------ locked cells

/** Whether a format locks its cells: unless its <protection> says locked="0", it does. */
export function formatLocked(workbook: Workbook, style: number): boolean {
  const format = workbook.styles.cellXfs[style];
  return !(format && format.protection && format.protection.locked === false);
}

/** Whether a format hides the formulas of its cells (on a protected sheet). */
export function formatHidden(workbook: Workbook, style: number): boolean {
  const format = workbook.styles.cellXfs[style];
  return !!(format && format.protection && format.protection.hidden === true);
}

/** The format a position has without a cell of its own: its row's, else its column's, else the default. */
function inheritedStyle(sheet: Sheet, row: number, col: number): number {
  const rowData = sheet.rows.get(row);
  if (rowData && rowData.style !== undefined) return rowData.style;
  return columnStyle(sheet.columns, col);
}

function columnStyle(columns: readonly ColumnRange[], col: number): number {
  for (let i = 0; i < columns.length; i++) {
    const range = columns[i];
    if (range.min > col) break;
    if (range.max >= col) return range.style !== undefined ? range.style : 0;
  }
  return 0;
}

const styleOfPosition = (sheet: Sheet, row: number, col: number): number => {
  const cell = getCell(sheet, row, col);
  return cell ? cell.style : inheritedStyle(sheet, row, col);
};

/** The most open ranges looked at: a file with more has the rest read as not open. */
const MAX_OPEN_RANGES = 1024;

/**
 * The ranges of a protected sheet that are open to everyone (<protectedRange> without a
 * password and without security descriptors), where rows and columns moved since put them.
 * UNSURE: Excel asks for the password of a range that has one; such a range is read as closed.
 */
function openRanges(sheet: Sheet): CellRange[] {
  const protection = sheet.protection;
  if (!protection || !protection.ranges) return [];
  const now = protectionNow(protection, sheet.structuralEdits, sheet.name);
  const out: CellRange[] = [];
  for (const range of now.ranges || []) {
    if (range.descriptors && range.descriptors.length > 0) continue;
    if (attrOf(range.attrs, 'password') !== undefined || attrOf(range.attrs, 'hashValue') !== undefined) continue;
    if (attrOf(range.attrs, 'securityDescriptor') !== undefined) continue;
    const sqref = attrOf(range.attrs, 'sqref');
    if (!sqref) continue;
    const refs = sqref.split(/\s+/);
    for (let i = 0; i < refs.length && out.length < MAX_OPEN_RANGES; i++) {
      const parsed = refs[i] ? parseRange(refs[i]) : undefined;
      if (parsed) out.push(parsed);
    }
  }
  return out;
}

const inside = (range: CellRange, row: number, col: number) => row >= range.top && row <= range.bottom && col >= range.left && col <= range.right;
const covers = (outer: CellRange, top: number, left: number, bottom: number, right: number) => outer.top <= top && outer.bottom >= bottom && outer.left <= left && outer.right >= right;

/** What the checks of one step share: the open ranges of each sheet, worked out once. */
class LockReader {
  private readonly opens = new Map<Sheet, CellRange[]>();

  constructor(private readonly workbook: Workbook) {}

  private open(sheet: Sheet): CellRange[] {
    let list = this.opens.get(sheet);
    if (!list) {
      list = openRanges(sheet);
      this.opens.set(sheet, list);
    }
    return list;
  }

  /** Whether the user may not change what is at a position of a protected sheet: it is locked and in no open range. */
  lockedAt(sheet: Sheet, row: number, col: number, style = styleOfPosition(sheet, row, col)): boolean {
    if (!formatLocked(this.workbook, style)) return false;
    const open = this.open(sheet);
    for (let i = 0; i < open.length; i++) if (inside(open[i], row, col)) return false;
    return true;
  }

  /**
   * The first column of [a, b] in row `row` that is in no open range and (with `cells`) has no cell
   * of its own; undefined when there is none. Jumps over each open range it meets.
   */
  private freeColumn(sheet: Sheet, row: number, a: number, b: number, cells?: readonly (Cell | undefined)[]): number | undefined {
    const open = this.open(sheet);
    let c = a;
    while (c <= b) {
      if (cells && cells[c]) {
        c++;
        continue;
      }
      let jump = -1;
      for (let i = 0; i < open.length; i++) if (inside(open[i], row, c) && open[i].right > jump) jump = open[i].right;
      if (jump < 0) return c;
      c = jump + 1;
    }
    return undefined;
  }

  private coveredByOpen(sheet: Sheet, top: number, left: number, bottom: number, right: number): boolean {
    const open = this.open(sheet);
    for (let i = 0; i < open.length; i++) if (covers(open[i], top, left, bottom, right)) return true;
    return false;
  }

  /**
   * A locked position of a range, the empty ones included (they take the format of their row or
   * column); undefined when every position is free. Looks at the rows and cells the sheet has,
   * not at every position: a range of whole columns costs what the sheet holds.
   */
  lockedIn(sheet: Sheet, range: CellRange): { row: number; col: number } | undefined {
    const top = Math.max(0, range.top);
    const bottom = Math.min(MAX_ROWS - 1, range.bottom);
    const left = Math.max(0, range.left);
    const right = Math.min(MAX_COLS - 1, range.right);
    if (top > bottom || left > right) return undefined;
    if (this.coveredByOpen(sheet, top, left, bottom, right)) return undefined;
    const segments = columnSegments(sheet.columns, left, right);
    let rowsWithRecord = 0;
    const visit = (r: number): { row: number; col: number } | undefined => {
      const row = sheet.rows.get(r);
      if (!row) return undefined;
      rowsWithRecord++;
      let found: { row: number; col: number } | undefined;
      const filled: number[] = new Array(segments.length).fill(0);
      let filledRow = 0;
      // The cells come in column order, and so do the segments: one pointer walks them.
      let k = 0;
      row.cells.forEach((cell, c) => {
        if (found || c < left || c > right) return;
        filledRow++;
        if (row.style === undefined) {
          while (k < segments.length && segments[k].max < c) k++;
          if (k < segments.length && segments[k].min <= c) filled[k]++;
        }
        if (this.lockedAt(sheet, r, c, cell.style)) found = { row: r, col: c };
      });
      if (found) return found;
      // The positions of the row without a cell of their own.
      // A position is named only when it is locked and in no open range.
      if (row.style !== undefined) {
        if (filledRow < right - left + 1 && formatLocked(this.workbook, row.style)) {
          const c = this.freeColumn(sheet, r, left, right, row.cells);
          if (c !== undefined) return { row: r, col: c };
        }
        return undefined;
      }
      for (let k = 0; k < segments.length; k++) {
        const segment = segments[k];
        if (filled[k] < segment.max - segment.min + 1 && formatLocked(this.workbook, segment.style)) {
          const c = this.freeColumn(sheet, r, segment.min, segment.max, row.cells);
          if (c !== undefined) return { row: r, col: c };
        }
      }
      return undefined;
    };
    if (bottom - top + 1 <= sheet.rows.size) {
      for (let r = top; r <= bottom; r++) {
        const found = visit(r);
        if (found) return found;
      }
    } else {
      for (const r of sheet.rows.keys()) {
        if (r < top || r > bottom) continue;
        const found = visit(r);
        if (found) return found;
      }
    }
    // Rows the sheet has no record for: each position takes the format of its column.
    // Which open ranges hold such a row changes only where one of them starts or ends: one row
    // without a record from each band of rows between those edges is enough.
    if (rowsWithRecord < bottom - top + 1) {
      const locked = segments.filter((segment) => formatLocked(this.workbook, segment.style));
      if (locked.length === 0) return undefined;
      const edges = [top];
      for (const range of this.open(sheet)) {
        if (range.top > top && range.top <= bottom) edges.push(range.top);
        if (range.bottom + 1 > top && range.bottom + 1 <= bottom) edges.push(range.bottom + 1);
      }
      edges.sort((x, y) => x - y);
      for (let i = 0; i < edges.length; i++) {
        if (i > 0 && edges[i] === edges[i - 1]) continue;
        const end = i + 1 < edges.length ? edges[i + 1] - 1 : bottom;
        let r = edges[i];
        while (r <= end && sheet.rows.has(r)) r++;
        if (r > end) continue;
        for (const segment of locked) {
          const c = this.freeColumn(sheet, r, segment.min, segment.max);
          if (c !== undefined) return { row: r, col: c };
        }
      }
    }
    return undefined;
  }
}

/** The columns of [left, right] in runs of one format: the column records, and the default format between them. */
function columnSegments(columns: readonly ColumnRange[], left: number, right: number): { min: number; max: number; style: number }[] {
  const out: { min: number; max: number; style: number }[] = [];
  let at = left;
  for (let i = 0; i < columns.length && at <= right; i++) {
    const range = columns[i];
    if (range.max < at) continue;
    if (range.min > right) break;
    if (range.min > at) {
      out.push({ min: at, max: range.min - 1, style: 0 });
      at = range.min;
    }
    const max = Math.min(range.max, right);
    out.push({ min: at, max, style: range.style !== undefined ? range.style : 0 });
    at = max + 1;
  }
  if (at <= right) out.push({ min: at, max: right, style: 0 });
  return out;
}

/** Whether the user may not change a cell: its sheet is protected and the cell locked (and in no open range). */
export function cellProtected(workbook: Workbook, sheet: Sheet, row: number, col: number): boolean {
  return sheetProtected(sheet) && new LockReader(workbook).lockedAt(sheet, row, col);
}

/** A position of a range the user may not change on a protected sheet; undefined when there is none or the sheet is not protected. */
export function protectedCellIn(workbook: Workbook, sheet: Sheet, range: CellRange): { row: number; col: number } | undefined {
  return sheetProtected(sheet) ? new LockReader(workbook).lockedIn(sheet, range) : undefined;
}

/** Whether the formula of a cell is hidden: its sheet is protected and its format hides it. Its value is still shown. */
export function formulaHidden(workbook: Workbook, sheet: Sheet, row: number, col: number): boolean {
  if (!sheetProtected(sheet)) return false;
  const cell = getCell(sheet, row, col);
  return !!cell && !!cell.formula && formatHidden(workbook, cell.style);
}

// ------------------------------------------------------------------ what a command may do

/** What the commands of the view ask about before they make any operation: see actionProblem(). */
export type ProtectionAction =
  | 'formatCells'
  | 'formatColumns'
  | 'formatRows'
  | 'insertColumns'
  | 'insertRows'
  | 'deleteColumns'
  | 'deleteRows'
  | 'sort'
  | 'autoFilter'
  | 'filterOnOff'
  | 'pivotTables'
  | 'objects'
  | 'rules'
  | 'merge';

/**
 * Why a command of a kind cannot be used on a sheet as it is protected; undefined when it can,
 * also for a sheet that is not protected. For the controls (a button that is off says why);
 * what a command actually changes is judged again by protectionProblem() when it is applied.
 * With a range, deletions and sorts look at its cells too: they move locked cells.
 */
export function actionProblem(workbook: Workbook, sheetIndex: number, action: ProtectionAction, range?: CellRange): string | undefined {
  const reason = actionReason(workbook, sheetIndex, action, range);
  return reason === undefined ? undefined : PROTECTION_MESSAGES[reason];
}

/** As actionProblem(), with the reason: for the words of the view. */
export function actionReason(workbook: Workbook, sheetIndex: number, action: ProtectionAction, range?: CellRange): ProtectionReason | undefined {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheetProtected(sheet)) return undefined;
  switch (action) {
    case 'formatCells':
      return protectionForbids(sheet, 'formatCells') ? 'formatCells' : undefined;
    case 'merge':
      return protectionForbids(sheet, 'formatCells') ? 'merge' : undefined;
    case 'formatColumns':
    case 'formatRows':
    case 'insertColumns':
    case 'insertRows':
    case 'pivotTables':
    case 'objects':
      return protectionForbids(sheet, action) ? action : undefined;
    case 'autoFilter':
      return protectionForbids(sheet, 'autoFilter') ? 'autoFilter' : undefined;
    case 'filterOnOff':
      // A filter is not turned on or off on a protected sheet, whatever it allows. UNSURE: from how Excel greys its button.
      return 'autoFilter';
    case 'rules':
      return 'rules';
    case 'deleteColumns':
    case 'deleteRows':
    case 'sort': {
      if (protectionForbids(sheet, action)) return action;
      if (range && new LockReader(workbook).lockedIn(sheet, wholeOf(range, action))) return action === 'sort' ? 'sort' : action;
      return undefined;
    }
  }
  return undefined;
}

function wholeOf(range: CellRange, action: ProtectionAction): CellRange {
  if (action === 'deleteRows') return { top: range.top, bottom: range.bottom, left: 0, right: MAX_COLS - 1 };
  if (action === 'deleteColumns') return { top: 0, bottom: MAX_ROWS - 1, left: range.left, right: range.right };
  return range;
}

// ------------------------------------------------------------------ the rule for a step

const isEmpty = (cell: Cell | undefined) => !cell || cell.kind === 'blank';

/** Whether what a cell holds differs: its kind, value, formula or rich text. A format is not content. */
function contentDiffers(a: Cell | undefined, b: Cell | undefined): boolean {
  if (isEmpty(a) && isEmpty(b)) return false;
  if (!a || !b) return true;
  if (a.kind !== b.kind || a.value !== b.value) return true;
  const fa = a.formula;
  const fb = b.formula;
  if (!fa !== !fb) return true;
  if (fa && fb && (fa.text !== fb.text || fa.type !== fb.type || fa.ref !== fb.ref)) return true;
  if (a.runs !== b.runs) {
    if (!a.runs || !b.runs || a.runs.length !== b.runs.length) return true;
    for (let i = 0; i < a.runs.length; i++) if (a.runs[i].text !== b.runs[i].text) return true;
  }
  return false;
}

/** Whether two formats lock and hide alike. */
function sameLock(workbook: Workbook, a: number, b: number): boolean {
  return a === b || (formatLocked(workbook, a) === formatLocked(workbook, b) && formatHidden(workbook, a) === formatHidden(workbook, b));
}

/** The changes of the columns of a sheet between two lists: of their sizes (width, hidden) and of their formats. */
function columnChanges(before: readonly ColumnRange[], after: readonly ColumnRange[]): { sized: boolean; styles: [number, number][] } {
  const cuts = new Set<number>([0, MAX_COLS]);
  for (const list of [before, after]) for (const range of list) {
    cuts.add(Math.max(0, range.min));
    cuts.add(Math.min(MAX_COLS, range.max + 1));
  }
  const points = [...cuts].sort((x, y) => x - y);
  let sized = false;
  const styles: [number, number][] = [];
  // Both lists are in column order: one walk through each, however many records they have.
  const walker = (list: readonly ColumnRange[]) => {
    let at = 0;
    return (col: number): ColumnRange | undefined => {
      while (at < list.length && list[at].max < col) at++;
      return at < list.length && list[at].min <= col ? list[at] : undefined;
    };
  };
  const findBefore = walker(before);
  const findAfter = walker(after);
  for (let i = 0; i + 1 < points.length; i++) {
    const col = points[i];
    const a = findBefore(col);
    const b = findAfter(col);
    const widthA = a && a.customWidth ? a.width : undefined;
    const widthB = b && b.customWidth ? b.width : undefined;
    if (widthA !== widthB || !!(a && a.hidden) !== !!(b && b.hidden)) sized = true;
    const styleA = a && a.style !== undefined ? a.style : 0;
    const styleB = b && b.style !== undefined ? b.style : 0;
    if (styleA !== styleB) styles.push([styleA, styleB]);
  }
  return { sized, styles };
}

const SIZE_PROPS: (keyof RowProps)[] = ['height', 'customHeight', 'hidden'];

/**
 * Why a step (the operations of one command) may not be applied to the workbook as it is now, as
 * its protection says; undefined when it may. Each operation is looked at against the workbook
 * before the step (UNSURE: a step that moves rows and then writes cells is judged by the cells
 * where they were; no command of the view makes one):
 *
 * - Cells: a step that changes what a cell holds (a value, a formula: typing, paste, Delete,
 *   Replace, a cut) may only touch cells that are not locked, and then may change their formats
 *   too (a paste brings its formats, Excel does the same); one locked cell refuses the whole
 *   step (UNSURE: Excel refuses a paste into a range that is partly locked). A step that changes
 *   formats alone needs formatCells, and never locks, unlocks or hides (`lock`).
 * - Columns and rows: a width, a height, hidden or shown needs formatColumns / formatRows; their
 *   format, formatCells as above.
 * - Inserted rows and columns need insertRows / insertColumns; deleted ones need deleteRows /
 *   deleteColumns and no locked position in them (as Excel: whole rows hold locked cells
 *   unless their format is unlocked).
 * - A sort needs sort and no locked cell in its range; a change of the criteria of the filter
 *   needs autoFilter; a filter is not added or removed at all.
 * - Merging and unmerging needs formatCells (UNSURE).
 * - Notes and drawings need objects (pictures, shapes and charts ask it themselves too).
 * - Data validations and conditional formats are not changed (UNSURE: Excel's buttons are off),
 *   but a paste or a cut that changes cells may carry them along, as it carries formats.
 * - A refresh of a pivot table needs pivotTables, and then writes its cells (UNSURE).
 * - The list of sheets and tab colours, with a locked structure: refused (`structure`).
 *
 * Frozen panes, the page setup, the names of print areas and the protection itself are free.
 */
export function protectionProblem(workbook: Workbook, ops: readonly Op[]): ProtectedError | undefined {
  const locked = structureLocked(workbook);
  let anyProtected = false;
  for (let i = 0; i < ops.length; i++) {
    const op = ops[i];
    switch (op.type) {
      case 'insertSheet':
      case 'deleteSheet':
      case 'renameSheet':
      case 'moveSheet':
      case 'setSheetState':
      case 'copySheet':
      case 'setTabColor':
        if (locked) return new ProtectedError('structure');
        break;
      default:
        if (sheetProtected(workbook.sheets[op.sheet])) anyProtected = true;
    }
  }
  if (!anyProtected) return undefined;

  const reader = new LockReader(workbook);
  // What kind of step it is, per sheet: one that changes what cells hold, one that refreshes a pivot table, one that sorts.
  const content = new Set<number>();
  const pivot = new Set<number>();
  const sorting = new Set<number>();
  for (let i = 0; i < ops.length; i++) {
    const op = ops[i];
    if (op.type === 'setCells') {
      const sheet = workbook.sheets[op.sheet];
      if (!sheet || content.has(op.sheet)) continue;
      for (let k = 0; k < op.cells.length; k++) {
        const change = op.cells[k];
        if (contentDiffers(getCell(sheet, change.row, change.col), change.cell)) {
          content.add(op.sheet);
          break;
        }
      }
    } else if (op.type === 'setPivot' || op.type === 'addPivot' || op.type === 'removePivot' || op.type === 'deletePivot' || op.type === 'restorePivot') {
      pivot.add(op.sheet);
      if (op.type === 'setPivot' && workbook.pivots) {
        for (const entry of op.tables) {
          const table = workbook.pivots.tables[entry.table];
          const at = table ? workbook.sheets.indexOf(table.sheet) : -1;
          if (at >= 0) pivot.add(at);
        }
      }
    } else if (op.type === 'sortRange') sorting.add(op.sheet);
  }

  for (let i = 0; i < ops.length; i++) {
    const op = ops[i];
    if (op.type === 'insertSheet' || op.type === 'deleteSheet' || op.type === 'renameSheet' || op.type === 'moveSheet' || op.type === 'setSheetState' || op.type === 'copySheet' || op.type === 'setTabColor') continue;
    const sheet = workbook.sheets[op.sheet];
    if (!sheetProtected(sheet)) continue;
    const forbids = (flag: ProtectionFlag) => protectionForbids(sheet, flag);
    const refuse = (reason: ProtectionReason) => new ProtectedError(reason, op.sheet);
    switch (op.type) {
      case 'setCells': {
        // A refresh writes the cells of its table.
        if (pivot.has(op.sheet)) break;
        const writes = content.has(op.sheet);
        for (let k = 0; k < op.cells.length; k++) {
          const change = op.cells[k];
          const old = getCell(sheet, change.row, change.col);
          const inherited = old ? -1 : inheritedStyle(sheet, change.row, change.col);
          const before = old ? old.style : inherited;
          const after = change.cell ? change.cell.style : old ? inheritedStyle(sheet, change.row, change.col) : inherited;
          const changed = contentDiffers(old, change.cell);
          if (!changed && before === after) continue;
          if (writes) {
            if (reader.lockedAt(sheet, change.row, change.col, before)) return refuse('cells');
          } else {
            if (forbids('formatCells')) return refuse('formatCells');
            if (!sameLock(workbook, before, after)) return refuse('lock');
          }
        }
        break;
      }
      case 'setColumns': {
        const changes = columnChanges(sheet.columns, op.columns);
        if (changes.sized && forbids('formatColumns')) return refuse('formatColumns');
        if (changes.styles.length > 0) {
          if (forbids('formatCells')) return refuse('formatCells');
          for (const [a, b] of changes.styles) if (!sameLock(workbook, a, b)) return refuse('lock');
        }
        break;
      }
      case 'setRows': {
        for (let k = 0; k < op.rows.length; k++) {
          const change = op.rows[k];
          const row = sheet.rows.get(change.row);
          for (const prop of SIZE_PROPS) {
            if (!(prop in change.props)) continue;
            const was = row ? row[prop] : prop === 'height' ? undefined : false;
            const now = change.props[prop];
            if ((was || undefined) !== (now || undefined) && forbids('formatRows')) return refuse('formatRows');
          }
          if ('style' in change.props) {
            const was = row && row.style !== undefined ? row.style : 0;
            const now = change.props.style !== undefined ? change.props.style : 0;
            if (was !== now) {
              if (forbids('formatCells')) return refuse('formatCells');
              if (!sameLock(workbook, was, now)) return refuse('lock');
            }
          }
        }
        break;
      }
      case 'setMerges':
        if (forbids('formatCells')) return refuse('merge');
        break;
      case 'setAutoFilter': {
        const was = sheet.autoFilter;
        const now = op.autoFilter;
        if (!was !== !now || (was && now && (was.range.top !== now.range.top || was.range.left !== now.range.left || was.range.bottom !== now.range.bottom || was.range.right !== now.range.right))) {
          return refuse('autoFilter');
        }
        // The state of a sort comes with the sort.
        if (!sorting.has(op.sheet) && forbids('autoFilter')) return refuse('autoFilter');
        break;
      }
      case 'sortRange':
        if (forbids('sort') || reader.lockedIn(sheet, op.range)) return refuse('sort');
        break;
      case 'insert':
        if (forbids(op.axis === 'row' ? 'insertRows' : 'insertColumns')) return refuse(op.axis === 'row' ? 'insertRows' : 'insertColumns');
        break;
      case 'delete': {
        const flag = op.axis === 'row' ? 'deleteRows' : 'deleteColumns';
        if (forbids(flag)) return refuse(flag);
        const range: CellRange = op.axis === 'row' ? { top: op.at, bottom: op.at + op.count - 1, left: 0, right: MAX_COLS - 1 } : { top: 0, bottom: MAX_ROWS - 1, left: op.at, right: op.at + op.count - 1 };
        if (reader.lockedIn(sheet, range)) return refuse(flag);
        break;
      }
      case 'setComments':
      case 'setDrawings':
        if (forbids('objects')) return refuse('objects');
        break;
      case 'setDataValidations':
      case 'setConditionalFormats':
        if (!content.has(op.sheet)) return refuse('rules');
        break;
      case 'setPivot':
      case 'addPivot':
      case 'removePivot':
      case 'deletePivot':
      case 'restorePivot':
        if (forbids('pivotTables')) return refuse('pivotTables');
        break;
      default:
        // Panes, the page setup, names, styles, the protection itself, and the inverses of structural steps.
        break;
    }
  }
  return undefined;
}

// ------------------------------------------------------------------ protecting and unprotecting

/** The attributes of a protection Excel writes for a sheet protected with the boxes it ticks by default. */
export function defaultProtectionAttrs(allow: Partial<{ [flag in ProtectionFlag]: boolean }> = {}): [string, string][] {
  const attrs: [string, string][] = [['sheet', '1']];
  // As Excel writes them: objects and scenarios always, every other flag only where it differs from its default.
  const objects = allow.objects === true ? '0' : '1';
  const scenarios = allow.scenarios === true ? '0' : '1';
  attrs.push(['objects', objects], ['scenarios', scenarios]);
  for (const flag of PROTECTION_FLAGS) {
    if (flag === 'objects' || flag === 'scenarios') continue;
    const wanted = allow[flag];
    if (wanted === undefined) continue;
    const forbid = !wanted;
    if (forbid !== FLAG_DEFAULTS[flag]) attrs.push([flag, forbid ? '1' : '0']);
  }
  return attrs;
}

/** The most rounds a password is hashed with here when a sheet is protected with one: what Excel uses. */
export const PROTECTION_SPIN_COUNT = 100000;

/** The attributes of a hashed password, as Excel writes them for a sheet: SHA-512, a salt of 16 bytes, 100,000 rounds. */
export async function passwordAttrs(password: string, salt: Uint8Array = randomBytes(16), spinCount = PROTECTION_SPIN_COUNT): Promise<[string, string][]> {
  const hash = await protectionHash('SHA512', salt, password, spinCount, 'append');
  return [
    ['algorithmName', 'SHA-512'],
    ['hashValue', toBase64(hash)],
    ['saltValue', toBase64(salt)],
    ['spinCount', String(spinCount)],
  ];
}

/**
 * The hash of a password as ISO/IEC 29500 states it for sheet and workbook protection: H(salt +
 * password as UTF-16LE), then `spinCount` times H(previous + round number) - the round number
 * after the digest ("append"), unlike the key derivation of encrypted files, which puts it
 * before ("prepend"). UNSURE: no file of Excel with a known password was at hand to compare
 * with; checkProtectionPassword() takes either order.
 */
async function protectionHash(name: HashName, salt: Uint8Array, password: string, spinCount: number, order: 'append' | 'prepend'): Promise<Uint8Array> {
  const bytes = utf16le(password);
  if (order === 'prepend') return spinHash(name, salt, bytes, spinCount);
  const hash = createHash(name);
  const size = hash.digestSize;
  const buffer = new Uint8Array(size + 4);
  const value = buffer.subarray(0, size);
  hash.update(salt).update(bytes).digest(value);
  for (let i = 0; i < spinCount; i++) {
    buffer[size] = i & 0xff;
    buffer[size + 1] = (i >>> 8) & 0xff;
    buffer[size + 2] = (i >>> 16) & 0xff;
    buffer[size + 3] = i >>> 24;
    hash.update(buffer).digest(value);
    // A turn for the page now and then: a hundred thousand rounds take a moment.
    if ((i & 0x3fff) === 0x3fff) await new Promise<void>((resolve) => setTimeout(resolve, 0));
  }
  const out = value.slice();
  buffer.fill(0);
  hash.wipe();
  return out;
}

/**
 * The hash of the `password` attribute of the legacy protection (and of the PASSWORD record of xls
 * files) as hexadecimal digits (at least four), as openpyxl computes it (`hash_password`): from every
 * character of the password. For the up to 15 characters Excel's own dialog takes this is the 16-bit
 * hash Excel writes; for longer passwords openpyxl's numbers grow past 16 bits, and so do these (no
 * rounding: the bits are kept in 16-bit pieces). `characters` cuts the password first: the checking
 * of a file tries the whole password and its first 15 characters, since which of the two a writer of
 * a long password hashed is not known. UNSURE: characters beyond ASCII (Excel works from the bytes of
 * the system code page); no Excel file with a long legacy password was at hand.
 */
export function legacyPasswordHash(password: string, characters = password.length): string {
  const length = Math.max(0, Math.min(password.length, characters, 1024));
  // Little-endian pieces of 16 bits: room for the character shifted by up to 1024 places.
  const pieces = new Array<number>(Math.ceil((length + 32) / 16) + 2).fill(0);
  const xorAt = (value: number, shift: number) => {
    // value < 2^16, put at bit `shift`.
    let at = shift >> 4;
    let bits = value * Math.pow(2, shift & 15);
    while (bits > 0) {
      pieces[at] ^= bits % 0x10000;
      bits = Math.floor(bits / 0x10000);
      at++;
    }
  };
  for (let i = 0; i < length; i++) {
    const code = password.charCodeAt(i);
    const shift = i + 1;
    if (shift < 15) {
      // Below 2^31: the bit operations of numbers hold it. The two halves are OR'ed, as openpyxl does.
      const value = code << shift;
      xorAt((value & 0x7fff) | (value >>> 15), 0);
    } else {
      // The low 15 bits are all zero; what is above them comes down by 15.
      xorAt(code, shift - 15);
    }
  }
  xorAt(length, 0);
  xorAt(0xce4b, 0);
  let text = '';
  for (let i = pieces.length - 1; i >= 0; i--) text += pieces[i].toString(16).padStart(4, '0');
  text = text.replace(/^0+/, '');
  return text.toUpperCase().padStart(4, '0');
}

/** Hash names of the protection, as the file writes them ("SHA-512"), to those of the engine. */
function hashNameOf(algorithm: string): HashName | undefined {
  const plain = algorithm.replace(/-/g, '').toUpperCase();
  return (HASH_NAMES as readonly string[]).indexOf(plain) >= 0 ? (plain as HashName) : undefined;
}

/**
 * Whether a password removes a protection (of a sheet or a workbook: the attributes of its element).
 * A protection without a password is removed by any. Throws RangeError for a hash of a kind that
 * cannot be checked here (an unknown algorithm, more rounds than Office allows).
 */
export async function checkProtectionPassword(attrs: readonly [string, string][] | undefined, password: string): Promise<boolean> {
  const value = (name: string) => attrOf(attrs, name);
  const hashValue = value('hashValue') !== undefined ? value('hashValue') : value('workbookHashValue');
  if (hashValue !== undefined) {
    const prefix = value('hashValue') !== undefined ? '' : 'workbook';
    const algorithm = value(prefix ? 'workbookAlgorithmName' : 'algorithmName') || '';
    const name = hashNameOf(algorithm);
    if (!name) throw new RangeError(`The protection is hashed with ${algorithm.slice(0, 40) || 'an unnamed algorithm'}, which cannot be checked here`);
    const salt = fromBase64(value(prefix ? 'workbookSaltValue' : 'saltValue') || '');
    const expected = fromBase64(hashValue);
    const spin = Number(value(prefix ? 'workbookSpinCount' : 'spinCount') || '0');
    if (!salt || !expected) throw new RangeError('The hash of the protection cannot be read');
    if (!(spin >= 0 && spin <= MAX_SPIN_COUNT) || Math.floor(spin) !== spin) throw new RangeError('The protection asks for more rounds of hashing than are done here');
    if (password.length > 255) return false;
    for (const order of ['append', 'prepend'] as const) {
      const hash = await protectionHash(name, salt, password, spin, order);
      if (hash.length === expected.length && constantTimeEqual(hash, expected)) return true;
    }
    return false;
  }
  const legacy = value('password') !== undefined ? value('password') : value('workbookPassword');
  if (legacy !== undefined && legacy !== '' && !/^0+$/.test(legacy)) {
    const expected = legacy.toUpperCase().replace(/^0+/, '').padStart(4, '0');
    // The whole password (openpyxl), or its first 15 characters (Excel's limit): see legacyPasswordHash().
    return legacyPasswordHash(password) === expected || legacyPasswordHash(password, 15) === expected;
  }
  return true;
}

/**
 * The operation that gives a sheet a protection (undefined: none). The model's object is replaced,
 * never changed; the ranges open in the protected sheet (<protectedRanges>) stay with the sheet.
 */
export function setProtectionOps(workbook: Workbook, sheetIndex: number, attrs: [string, string][] | undefined): Op[] {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet) throw new RangeError(`No sheet at index ${sheetIndex}`);
  if (sheet.kind !== 'worksheet') throw new RangeError(`Cannot protect a ${sheet.kind}: not supported yet`);
  const was = sheet.protection;
  let protection: SheetProtection | undefined;
  if (attrs) protection = { ...(was && was.ranges ? { ranges: was.ranges } : {}), attrs };
  else if (was && was.ranges) protection = { ranges: was.ranges };
  return [{ type: 'setProtection', sheet: sheetIndex, protection }];
}

/** The operation that removes the protection of a sheet, the password checked first (see checkProtectionPassword()). */
export async function unprotectSheetOps(workbook: Workbook, sheetIndex: number, password = ''): Promise<Op[]> {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet) throw new RangeError(`No sheet at index ${sheetIndex}`);
  if (!sheetProtected(sheet)) return [];
  if (!(await checkProtectionPassword(sheet.protection!.attrs, String(password)))) throw new PasswordMismatch();
  return setProtectionOps(workbook, sheetIndex, undefined);
}

/** The password given does not remove the protection. */
export class PasswordMismatch extends RangeError {
  readonly code = 'wrong-password';
  constructor() {
    super('The password is not the one of the protection');
    this.name = 'PasswordMismatch';
  }
}
