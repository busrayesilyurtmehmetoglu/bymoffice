// Every change to a workbook is an operation. Applying one returns the
// operation that undoes it, which is all undo and redo need; later the same
// operations are what collaborative editing will send over the wire.
//
// Cells and rows are treated as immutable here: a change puts a new object in
// place instead of altering the old one. That is what lets a structural
// operation keep the previous state for undo by reference, without copying it.

import { touchModel } from '../model/edits';
import { adjustFormula, adjustRange, adjustReferenceText, formulaNamesSheetsUnread, renameSheetsInFormula, shiftFormula, type StructuralEdit } from '../calc/refs';
import { adjustDrawingAnchor, commentCellNow, commentEdits } from '../render/sheet-drawings';
import { chartsFollow, chartsMoved, chartsRenamed } from './charts';
import { liveRulesMoved, liveRulesRenamed, swapRuleLists, type RuleLists } from './rules';
import {
  FILTER_NAME,
  MAX_COLS,
  MAX_ROWS,
  cellName,
  checkSheetName,
  createSheet,
  isFileSheet,
  foldSheetName,
  parseCellRef,
  parseRange,
  type AutoFilter,
  type Cell,
  type CellFormat,
  type CellRange,
  type Color,
  type ColumnRange,
  type ConditionalFormat,
  type DataValidation,
  type DefinedName,
  type DrawingAnchor,
  type Hyperlink,
  type Row,
  type Sheet,
  type SheetComment,
  type SheetDrawing,
  type SheetNameMap,
  type SheetNeighbours,
  type SheetState,
  type SheetStructuralEdit,
  type StyleSheet,
  type Workbook,
} from '../model/sheet/types';

export interface CellChange {
  row: number;
  col: number;
  /** The new cell; undefined removes the cell. */
  cell: Cell | undefined;
}

/** Everything a structural operation can alter, held by reference. */
export interface SheetSnapshot {
  rows: Map<number, Row>;
  columns: ColumnRange[];
  merges: CellRange[];
  hyperlinks: Hyperlink[];
  dimension: CellRange | undefined;
  rowCount: number;
  colCount: number;
  structuralEdits: SheetStructuralEdit[] | undefined;
  autoFilter?: AutoFilter;
  /** The comments, when the operation moved notes (a sort): the list to put back. Missing when it left them alone. */
  comments?: { list: SheetComment[] | undefined };
}

/**
 * The list of sheets and everything an operation on it alters, held by reference: what the
 * workbook is put back to by "restoreSheets".
 */
export interface SheetsSnapshot {
  /** The sheets, in order; the very objects, so that a sheet that comes back is the one that went. */
  sheets: Sheet[];
  /** The name of a sheet (and the name the file has for it) to put back. */
  renamed?: { sheet: Sheet; name: string; nameInFile: string | undefined };
  /** Whether a sheet is shown, to put back. */
  shown?: { sheet: Sheet; state: SheetState };
  /** Cells whose formulas named a sheet, as they were. Shared formulas stay shared. */
  cells: { sheet: Sheet; cells: CellChange[] }[];
  /** Links that pointed into a sheet, as they were. */
  links: { sheet: Sheet; hyperlinks: Hyperlink[] }[];
  /** Lists of live rules (see rules.ts) whose formulas named a sheet, as they were. */
  rules?: RuleLists[];
  definedNames: DefinedName[];
  activeSheet: number;
  removedSheetNames: string[] | undefined;
  removedSheetNeighbours: SheetNeighbours[] | undefined;
}

export type Op =
  | {
      type: 'setCells';
      sheet: number;
      cells: CellChange[];
      /** Used extent to restore; set on the inverse of an operation that grew the sheet. */
      extent?: { rowCount: number; colCount: number };
    }
  | {
      /** Replaces the column settings (widths, visibility, formats) of a sheet. */
      type: 'setColumns';
      sheet: number;
      columns: ColumnRange[];
    }
  | {
      /** Changes properties of rows; a property given as undefined is removed. */
      type: 'setRows';
      sheet: number;
      rows: { row: number; props: RowProps }[];
    }
  | { type: 'setMerges'; sheet: number; merges: CellRange[] }
  | {
      /**
       * Replaces the filter of a sheet (undefined: no filter) and shows and hides rows with it.
       * Which rows a criterion hides depends on what the cells show, so it is worked out by the
       * commands (see filter.ts) and stated here: applying the operation needs no locale.
       */
      type: 'setAutoFilter';
      sheet: number;
      autoFilter: AutoFilter | undefined;
      show?: number[];
      hide?: number[];
    }
  | {
      /**
       * Moves the cells of the rows of a range, as sorting does: row `range.top + i` takes the
       * cells that row `range.top + order[i]` has within the columns of the range. `order` is a
       * permutation of 0 .. height - 1. Relative references of moved formulas move with them.
       */
      type: 'sortRange';
      sheet: number;
      range: CellRange;
      order: number[];
    }
  | {
      /**
       * Replaces the comments of a sheet (undefined: none). The list is held by reference and
       * never altered, so that the inverse puts the very list back. What the list may hold is
       * the business of the commands (see notes.ts).
       */
      type: 'setComments';
      sheet: number;
      comments: SheetComment[] | undefined;
    }
  | {
      /**
       * Replaces the data validations of a sheet (undefined: none). The list is held by reference
       * and never altered, so that the inverse puts the very list back. What the list may hold is
       * the business of the commands (see rules.ts).
       */
      type: 'setDataValidations';
      sheet: number;
      validations: DataValidation[] | undefined;
      by?: OpCommand;
    }
  | {
      /** Replaces the conditional formats of a sheet (undefined: none); as "setDataValidations". */
      type: 'setConditionalFormats';
      sheet: number;
      formats: ConditionalFormat[] | undefined;
      by?: OpCommand;
    }
  | {
      /**
       * Replaces the drawings of a sheet (undefined: none). The list is held by reference and
       * never altered, so that the inverse puts the very list back. What the list may hold is
       * the business of the commands (see charts.ts).
       */
      type: 'setDrawings';
      sheet: number;
      drawings: SheetDrawing[] | undefined;
      by?: OpCommand;
    }
  | { type: 'insert'; sheet: number; axis: 'row' | 'col'; at: number; count: number }
  | { type: 'delete'; sheet: number; axis: 'row' | 'col'; at: number; count: number }
  | {
      /** Puts a sheet back the way a snapshot recorded it; the inverse of a structural operation. */
      type: 'restore';
      sheet: number;
      snapshot: SheetSnapshot;
      /** Cells of other sheets whose formulas referred to the sheet. */
      others: { sheet: number; cells: CellChange[] }[];
      /** Links of other sheets that pointed into the sheet. */
      otherLinks?: { sheet: number; hyperlinks: Hyperlink[] }[];
      definedNames: DefinedName[];
      /** Lists of rules that were added or changed since the file was read (live rules), as they were: of this sheet, and of sheets whose rules name it. */
      rules?: RuleLists[];
      /** Roughly how many rows this operation keeps alive for undo; History uses it to bound memory. */
      weight?: number;
    }
  | {
      /**
       * Adds an empty worksheet named `name` so that it becomes sheet number `sheet` (0 puts it
       * first, the number of sheets last). Sheets after it move up by one, and with them the
       * active sheet and the sheets that defined names are local to.
       */
      type: 'insertSheet';
      sheet: number;
      name: string;
    }
  | {
      /**
       * Removes a worksheet. What other sheets, defined names and links read from it becomes
       * #REF!, the names local to it go with it, and another visible sheet becomes the active
       * one when it was. Refused for the last visible sheet.
       */
      type: 'deleteSheet';
      sheet: number;
    }
  | {
      /** Gives a worksheet another name, and writes that name wherever a formula, a defined name or a link names the sheet. */
      type: 'renameSheet';
      sheet: number;
      name: string;
    }
  | {
      /**
       * Moves a sheet so that it becomes sheet number `to`. The active sheet and the sheets that
       * defined names are local to follow their sheets. No formula changes: a 3D reference
       * ("Sheet1:Sheet3!A1") is the text it was and reads the sheets that stand between its ends
       * now, as in Excel (this engine calculates no 3D reference).
       */
      type: 'moveSheet';
      sheet: number;
      to: number;
    }
  | {
      /**
       * Shows or hides a sheet. Refused for the last visible sheet; when the active sheet is
       * hidden, another visible one becomes the active one.
       */
      type: 'setSheetState';
      sheet: number;
      state: SheetState;
    }
  | {
      /**
       * Adds `copy` (a sheet that is in no file: see copySheet() in sheets.ts, which makes it)
       * so that it becomes sheet number `sheet`, with the defined names that are local to it
       * (`names`, whose `localSheet` is set here). Otherwise like "insertSheet".
       */
      type: 'copySheet';
      sheet: number;
      copy: Sheet;
      names: DefinedName[];
    }
  | {
      /** Sets the colour of the tab of a sheet; undefined removes it. Its own inverse. */
      type: 'setTabColor';
      sheet: number;
      color: Color | undefined;
    }
  | {
      /**
       * Puts the list of sheets back the way a snapshot recorded it; the inverse of the
       * operations on sheets, and its own inverse. `sheet` is the index the operation it undoes
       * was about (where a sheet comes back, goes away or changes its name).
       */
      type: 'restoreSheets';
      sheet: number;
      state: SheetsSnapshot;
      /** As for "restore". */
      weight?: number;
    };

/**
 * The command that made an operation which replaces a whole list (rules, drawings): its name and
 * what it was called with after the workbook and the sheet. Applying the operation never looks at
 * it. Collaborative editing sends the command in place of the list, which names things of the
 * file each participant read (see src/collab/wire.ts); an operation without it is not sent.
 */
export interface OpCommand {
  name: string;
  args: unknown[];
}

/** Notes on the one operation a command made which command that was (see OpCommand). */
export function withCommand(ops: Op[], name: string, args: unknown[]): Op[] {
  if (ops.length === 1) {
    const op = ops[0];
    if (op.type === 'setDataValidations' || op.type === 'setConditionalFormats' || op.type === 'setDrawings') op.by = { name, args };
  }
  return ops;
}

/** The row properties an operation may set. */
export type RowProps = Partial<Pick<Row, 'height' | 'customHeight' | 'hidden' | 'style'>>;

const sheetAt = (workbook: Workbook, index: number): Sheet => {
  const sheet = workbook.sheets[index];
  if (!sheet) throw new RangeError(`No sheet at index ${index}`);
  return sheet;
};

const newRow = (cells: Cell[] = []): Row => ({ customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells });

const isPlainRow = (row: Row) =>
  row.height === undefined && !row.customHeight && !row.hidden && row.style === undefined && row.outlineLevel === 0 && !row.collapsed && row.attrs === undefined;

function checkCell(row: number, col: number): void {
  if (!(row >= 0 && row < MAX_ROWS && col >= 0 && col < MAX_COLS) || Math.floor(row) !== row || Math.floor(col) !== col) {
    throw new RangeError(`Cell (${row}, ${col}) is outside the sheet`);
  }
}

/**
 * Applies cell changes and returns the changes that undo them, in reverse order. With
 * `keepShared` the cells that depend on a changed shared formula are left as they are: for
 * changes that are the same for every cell of the group (another name for a sheet).
 */
function setCells(sheet: Sheet, changes: CellChange[], keepShared = false): CellChange[] {
  // Everything is checked before anything is changed: an operation applies as a whole or not at all.
  for (const change of changes) checkCell(change.row, change.col);
  const all = keepShared ? changes : withDependants(sheet, changes);
  const undo: CellChange[] = [];
  // Rows are copied once per operation, on their first change.
  const copied = new Set<number>();
  for (const change of all) {
    let row = sheet.rows.get(change.row);
    const previous = row ? row.cells[change.col] : undefined;
    if (previous === change.cell) continue;
    if (!row) {
      if (!change.cell) continue;
      row = newRow();
      sheet.rows.set(change.row, row);
      copied.add(change.row);
    } else if (!copied.has(change.row)) {
      row = { ...row, cells: row.cells.slice() };
      sheet.rows.set(change.row, row);
      copied.add(change.row);
    }
    if (change.cell) row.cells[change.col] = change.cell;
    else {
      delete row.cells[change.col];
      // A row left with neither cells nor properties of its own is not a row any more.
      if (isPlainRow(row) && !row.cells.some(() => true)) sheet.rows.delete(change.row);
    }
    undo.push({ row: change.row, col: change.col, cell: previous });
    (sheet.changedRows || (sheet.changedRows = new Set())).add(change.row);
    if (change.cell) {
      if (change.row >= sheet.rowCount) sheet.rowCount = change.row + 1;
      if (change.col >= sheet.colCount) sheet.colCount = change.col + 1;
    }
  }
  return undo.reverse();
}

/**
 * A shared formula keeps its text in one cell only. Before that cell is
 * overwritten, the cells that depend on it get the formula written out, or
 * they would be left pointing at nothing.
 */
function withDependants(sheet: Sheet, changes: CellChange[]): CellChange[] {
  let groups: Set<number> | undefined;
  for (const change of changes) {
    const row = sheet.rows.get(change.row);
    const formula = row && row.cells[change.col] ? row.cells[change.col].formula : undefined;
    // A change that leaves the formula itself alone (a new format, say) keeps the group as it is.
    const kept = change.cell !== undefined && change.cell.formula === formula;
    if (formula && formula.type === 'shared' && formula.text !== '' && formula.si !== undefined && !kept) {
      if (!groups) groups = new Set();
      groups.add(formula.si);
    }
  }
  if (!groups) return changes;
  const masters = sharedMasters(sheet);
  const targeted = new Set(changes.map((change) => change.row * MAX_COLS + change.col));
  const extra: CellChange[] = [];
  for (const [index, row] of sheet.rows) {
    row.cells.forEach((cell, col) => {
      const formula = cell.formula;
      if (!formula || formula.type !== 'shared' || formula.si === undefined || !groups!.has(formula.si)) return;
      if (targeted.has(index * MAX_COLS + col)) return;
      extra.push({ row: index, col, cell: explicitFormula(cell, index, col, masters) });
    });
  }
  return extra.concat(changes);
}

function snapshotOf(sheet: Sheet): SheetSnapshot {
  return {
    rows: sheet.rows,
    columns: sheet.columns,
    merges: sheet.merges,
    hyperlinks: sheet.hyperlinks,
    dimension: sheet.dimension,
    rowCount: sheet.rowCount,
    colCount: sheet.colCount,
    structuralEdits: sheet.structuralEdits,
    autoFilter: sheet.autoFilter,
  };
}

function restoreSnapshot(sheet: Sheet, snapshot: SheetSnapshot): void {
  sheet.rows = snapshot.rows;
  sheet.columns = snapshot.columns;
  sheet.merges = snapshot.merges;
  sheet.hyperlinks = snapshot.hyperlinks;
  sheet.dimension = snapshot.dimension;
  sheet.rowCount = snapshot.rowCount;
  sheet.colCount = snapshot.colCount;
  sheet.structuralEdits = snapshot.structuralEdits;
  if (snapshot.autoFilter) sheet.autoFilter = snapshot.autoFilter;
  else delete sheet.autoFilter;
  if (snapshot.comments) {
    if (snapshot.comments.list && snapshot.comments.list.length > 0) sheet.comments = snapshot.comments.list;
    else delete sheet.comments;
    sheet.commentsModified = true;
  }
}

export type SharedMasters = Map<number, { text: string; row: number; col: number }>;

/** The cells that hold the text of each shared formula group of a sheet. */
export function sharedMasters(sheet: Sheet): SharedMasters {
  const masters: SharedMasters = new Map();
  for (const [index, row] of sheet.rows) {
    row.cells.forEach((cell, col) => {
      const formula = cell.formula;
      if (formula && formula.type === 'shared' && formula.text !== '' && formula.si !== undefined && !masters.has(formula.si)) {
        masters.set(formula.si, { text: formula.text, row: index, col });
      }
    });
  }
  return masters;
}

/** A shared formula written out for the cell it applies to. */
export function explicitFormula(cell: Cell, row: number, col: number, masters: SharedMasters): Cell {
  const formula = cell.formula;
  if (!formula || formula.type !== 'shared') return cell;
  let text = formula.text;
  if (text === '' && formula.si !== undefined) {
    const master = masters.get(formula.si);
    // A dependant without its master cannot be rebuilt; it keeps its last value and loses the formula.
    if (!master) return { kind: cell.kind, value: cell.value, style: cell.style, ...(cell.runs ? { runs: cell.runs } : {}) };
    text = shiftFormula(master.text, row - master.row, col - master.col);
  }
  return { ...cell, formula: { text, type: 'normal', ...(formula.attrs ? { attrs: formula.attrs } : {}) } };
}

/**
 * Rewrites the formulas of a sheet for a structural edit and returns the rows
 * with the changes applied, leaving the sheet's own rows untouched. Shared
 * formulas are written out: the group no longer describes its cells once
 * rows or columns move through it.
 */
function adjustSheetFormulas(sheet: Sheet, edit: StructuralEdit, unshareAll: boolean): Map<number, Row> {
  const masters = sharedMasters(sheet);
  let shared = masters.size > 0;
  if (!shared && unshareAll) {
    // Dependants whose master is gone still have to lose their formula.
    for (const row of sheet.rows.values()) {
      if (row.cells.some((cell) => cell.formula !== undefined && cell.formula.type === 'shared')) {
        shared = true;
        break;
      }
    }
  }
  // In a sheet whose cells stay where they are, only the groups whose text changes are written out.
  let affected: Set<number> | undefined;
  if (!unshareAll) {
    // A group is affected when the formula of any of its cells changes, not only that of the master:
    // the dependants refer to other rows than the master does.
    const groups = new Set<number>();
    for (const [index, row] of sheet.rows) {
      row.cells.forEach((cell, col) => {
        const formula = cell.formula;
        if (!formula || formula.type !== 'shared' || formula.si === undefined || groups.has(formula.si)) return;
        const explicit = explicitFormula(cell, index, col, masters).formula;
        if (explicit && adjustFormula(explicit.text, sheet.name, edit) !== explicit.text) groups.add(formula.si);
      });
    }
    affected = groups;
    shared = groups.size > 0;
  }

  const changed = new Map<number, Row>();
  for (const [index, row] of sheet.rows) {
    let cells: Cell[] | undefined;
    row.cells.forEach((cell, col) => {
      if (!cell.formula) return;
      const group = cell.formula.type === 'shared' ? cell.formula.si : undefined;
      let next = shared && (!affected || (group !== undefined && affected.has(group))) ? explicitFormula(cell, index, col, masters) : cell;
      const formula = next.formula;
      if (formula) {
        // A dependant that keeps pointing at its group has no text of its own to adjust.
        const text = formula.text === '' ? '' : adjustFormula(formula.text, sheet.name, edit);
        const ref = formula.ref === undefined ? undefined : adjustReferenceText(formula.ref, sheet.name, edit);
        if (text !== formula.text || ref !== formula.ref) {
          next = { ...next, formula: { ...formula, text, ...(ref === undefined || ref === '' ? {} : { ref }) } };
          if (ref === '' && next.formula) delete next.formula.ref;
        }
      }
      if (next !== cell) {
        if (!cells) cells = row.cells.slice();
        cells[col] = next;
      }
    });
    if (cells) changed.set(index, { ...row, cells });
  }
  return changed;
}

function shiftColumns(columns: ColumnRange[], edit: SheetStructuralEdit): ColumnRange[] {
  const out: ColumnRange[] = [];
  for (const range of columns) {
    const moved = adjustRange({ top: 0, bottom: 0, left: range.min, right: range.max }, edit);
    if (moved) out.push({ ...range, min: moved.left, max: Math.min(moved.right, MAX_COLS - 1) });
  }
  return out;
}

/**
 * An array formula covers its range as one unit. Like Excel, an edit that
 * would cut through such a range is refused; moving or removing all of it is fine.
 */
function checkArrayFormulas(sheet: Sheet, edit: SheetStructuralEdit): void {
  const end = edit.at + edit.count - 1;
  for (const row of sheet.rows.values()) {
    row.cells.forEach((cell) => {
      const formula = cell.formula;
      if (!formula || formula.type !== 'array' || formula.ref === undefined) return;
      const range = parseRange(formula.ref);
      if (!range) return;
      const first = edit.axis === 'row' ? range.top : range.left;
      const last = edit.axis === 'row' ? range.bottom : range.right;
      if (first === last && edit.kind === 'insert') return;
      const cuts = edit.kind === 'insert' ? edit.at > first && edit.at <= last : edit.at <= last && end >= first && !(edit.at <= first && end >= last);
      if (cuts) throw new RangeError(`Cannot ${edit.kind} ${edit.axis}s through the array formula in ${formula.ref}`);
    });
  }
}

/** Kinds of anchored parts whose positions the writer adjusts after a structural edit. */
const MOVABLE_PARTS = new Set(['drawing', 'comments', 'vmlDrawing', 'image']);

/** Inserts or deletes rows or columns and fixes up everything that refers to them. */
function structural(workbook: Workbook, index: number, edit: SheetStructuralEdit): Op {
  const sheet = sheetAt(workbook, index);
  const { kind, axis, at, count } = edit;
  const limit = axis === 'row' ? MAX_ROWS : MAX_COLS;
  if (!(at >= 0 && count > 0 && at + count <= limit) || Math.floor(at) !== at || Math.floor(count) !== count) {
    throw new RangeError(`Cannot ${kind} ${count} ${axis}(s) at ${at}`);
  }
  if (kind === 'insert') {
    // Excel refuses an insert that would push cells off the sheet rather than lose them.
    // Rows count even without cells: a height or a format is content too.
    let used = axis === 'row' ? sheet.rowCount : sheet.colCount;
    if (axis === 'row') for (const key of sheet.rows.keys()) if (key >= used) used = key + 1;
    if (used + count > limit && used > at) throw new RangeError(`Cannot insert ${count} ${axis}(s): cells would be pushed off the sheet`);
  }
  // Content placed by cell position but stored in other parts has to move with the cells. The writer
  // moves pictures and comments; for anything else the edit is refused rather than leave it behind.
  // Charts follow when the model holds every reference of every chart of the workbook (see charts.ts).
  const follow = chartsFollow(workbook);
  const stuck = (sheet.anchored || []).filter((part) => !MOVABLE_PARTS.has(part) && !(part === 'chart' && follow));
  if (stuck.length > 0) throw new RangeError(`Cannot ${kind} ${axis}s in a sheet with ${stuck.join(', ')}: not supported yet`);
  checkArrayFormulas(sheet, edit);

  const snapshot = snapshotOf(sheet);
  const refEdit: StructuralEdit = { ...edit, sheetName: sheet.name };
  const adjusted = adjustSheetFormulas(sheet, refEdit, true);
  let weight = adjusted.size + (axis === 'col' ? sheet.rows.size : 0);

  // A filter that loses its header row, or all of its range, is gone, as in Excel; the rows it hid are shown.
  const filter = sheet.autoFilter ? shiftAutoFilter(sheet.autoFilter, edit) : undefined;
  const dropped = sheet.autoFilter && !filter ? sheet.autoFilter.range : undefined;
  const shown = (i: number, row: Row): Row | undefined => {
    if (!dropped || !row.hidden || i <= dropped.top || i > dropped.bottom) return row;
    const next: Row = { ...row, hidden: false };
    return isPlainRow(next) && !next.cells.some(() => true) ? undefined : next;
  };

  const rows = new Map<number, Row>();
  if (axis === 'row') {
    for (const [i, original] of sheet.rows) {
      const row = shown(i, adjusted.get(i) || original);
      if (!row) continue;
      if (i < at) rows.set(i, row);
      else if (kind === 'insert') rows.set(i + count, row);
      else if (i >= at + count) rows.set(i - count, row);
    }
    sheet.rowCount = kind === 'insert' ? (sheet.rowCount > at ? sheet.rowCount + count : sheet.rowCount) : sheet.rowCount <= at ? sheet.rowCount : Math.max(sheet.rowCount - count, at);
  } else {
    for (const [i, original] of sheet.rows) {
      const row = shown(i, adjusted.get(i) || original);
      if (!row) continue;
      if (row.cells.length <= at) {
        rows.set(i, row);
        continue;
      }
      const cells: Cell[] = [];
      row.cells.forEach((cell, col) => {
        if (col < at) cells[col] = cell;
        else if (kind === 'insert') cells[col + count] = cell;
        else if (col >= at + count) cells[col - count] = cell;
      });
      rows.set(i, { ...row, cells });
    }
    sheet.colCount = kind === 'insert' ? (sheet.colCount > at ? sheet.colCount + count : sheet.colCount) : sheet.colCount <= at ? sheet.colCount : Math.max(sheet.colCount - count, at);
    sheet.columns = shiftColumns(sheet.columns, edit);
  }
  sheet.rows = rows;

  const moveRange = (range: CellRange) => adjustRange(range, edit);
  sheet.merges = sheet.merges.map(moveRange).filter((range): range is CellRange => !!range && (range.bottom > range.top || range.right > range.left));
  const hyperlinks: Hyperlink[] = [];
  for (const link of sheet.hyperlinks) {
    const range = moveRange(link.range);
    if (!range) continue;
    // A link into this workbook names a cell that may have moved as well.
    const location = link.location === undefined ? undefined : adjustFormula(link.location, sheet.name, refEdit);
    hyperlinks.push(location === link.location ? { ...link, range } : { ...link, range, location });
  }
  sheet.hyperlinks = hyperlinks;
  if (sheet.dimension) sheet.dimension = moveRange(sheet.dimension);
  if (filter) sheet.autoFilter = filter;
  else if (dropped) {
    delete sheet.autoFilter;
    sheet.autoFilterModified = true;
  }
  sheet.structuralEdits = (sheet.structuralEdits || []).concat([{ kind, axis, at, count }]);
  sheet.modified = true;

  // Formulas elsewhere that point into this sheet, and names defined on it.
  const others: { sheet: number; cells: CellChange[] }[] = [];
  const otherLinks: { sheet: number; hyperlinks: Hyperlink[] }[] = [];
  workbook.sheets.forEach((other, otherIndex) => {
    if (other === sheet) return;
    const changes: CellChange[] = [];
    for (const [rowIndex, row] of adjustSheetFormulas(other, refEdit, false)) {
      const previous = other.rows.get(rowIndex);
      row.cells.forEach((cell, col) => {
        if (!previous || previous.cells[col] !== cell) changes.push({ row: rowIndex, col, cell });
      });
    }
    weight += changes.length;
    if (changes.length > 0) {
      others.push({ sheet: otherIndex, cells: setCells(other, changes) });
      other.modified = true;
    }
    // Links elsewhere that point at a cell of this sheet follow it.
    let links: Hyperlink[] | undefined;
    other.hyperlinks.forEach((link, i) => {
      if (link.location === undefined) return;
      const location = adjustFormula(link.location, other.name, refEdit);
      if (location === link.location) return;
      if (!links) links = other.hyperlinks.slice();
      links[i] = { ...link, location };
    });
    if (links) {
      otherLinks.push({ sheet: otherIndex, hyperlinks: other.hyperlinks });
      other.hyperlinks = links;
      other.hyperlinksModified = true;
      other.modified = true;
    }
  });
  const definedNames = workbook.definedNames;
  workbook.definedNames = definedNames.map((name) => {
    const home = name.localSheet !== undefined && workbook.sheets[name.localSheet] ? workbook.sheets[name.localSheet].name : '';
    const value = adjustFormula(name.value, home, refEdit);
    return value === name.value ? name : { ...name, value };
  });
  // The name that went with the filter would be left naming cells that are gone.
  if (dropped) syncFilterName(workbook, index);
  // Rules added or changed since the file was read are kept where their cells are; those of the file are
  // left as the file has them, and are read through the edit from now on.
  // So are the references of charts: wherever one names cells that moved, it is written anew.
  const rules = swapRuleLists(liveRulesMoved(workbook, sheet, refEdit).concat(chartsMoved(workbook, refEdit)));

  return { type: 'restore', sheet: index, snapshot, others, otherLinks, definedNames, ...(rules.length > 0 ? { rules } : {}), weight };
}

/**
 * A filter after rows or columns moved: its range follows, and its criteria stay with their
 * columns (a criterion on a deleted column goes with it; the command that deletes the column
 * shows the rows only that criterion hid, see deleteColumns() in filter.ts). Undefined when
 * the range is gone, or its header row is: a filter cannot do without its header, and the
 * row under it must not silently become one.
 * The writer does the same to the source element of a filter that was not changed otherwise.
 */
export function shiftAutoFilter(filter: AutoFilter, edit: SheetStructuralEdit): AutoFilter | undefined {
  const header = filter.range.top;
  if (edit.kind === 'delete' && edit.axis === 'row' && edit.at <= header && edit.at + edit.count > header) return undefined;
  const range = adjustRange(filter.range, edit);
  if (!range) return undefined;
  const next: AutoFilter = { ...filter, range };
  if (edit.axis === 'col') {
    next.columns = [];
    for (const column of filter.columns) {
      const at = filter.range.left + column.colId;
      const moved = adjustRange({ top: 0, bottom: 0, left: at, right: at }, edit);
      if (!moved || moved.left < range.left || moved.left > range.right) continue;
      const colId = moved.left - range.left;
      next.columns.push(colId === column.colId ? column : { ...column, colId });
    }
  } else if (filter.held) {
    next.held = [];
    for (const row of filter.held) {
      const moved = adjustRange({ top: row, bottom: row, left: 0, right: 0 }, edit);
      if (moved) next.held.push(moved.top);
    }
  }
  const state = filter.sortState;
  if (state) {
    const sorted = adjustRange(state.range, edit);
    if (!sorted) delete next.sortState;
    else {
      const conditions = [];
      for (const condition of state.conditions) {
        const key = adjustRange(condition.range, edit);
        if (key) conditions.push({ ...condition, range: key });
      }
      next.sortState = { ...state, range: sorted, conditions };
    }
  }
  return next;
}

/** A sheet name as it stands before "!" in a reference. */
function quotedSheetName(name: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name) && !parseCellRef(name) && !/^[Rr][Cc]?\d*$/.test(name)) return name;
  return `'${name.replace(/'/g, "''")}'`;
}

/**
 * Brings the hidden defined name of a sheet's filter in line with the filter: applications
 * expect "_xlnm._FilterDatabase", local to the sheet, to name the filter range.
 */
function syncFilterName(workbook: Workbook, index: number): void {
  const sheet = workbook.sheets[index];
  const names = workbook.definedNames;
  const at = names.findIndex((name) => name.localSheet === index && name.name.toLowerCase() === FILTER_NAME.toLowerCase());
  if (!sheet.autoFilter) {
    if (at >= 0) workbook.definedNames = names.filter((_, i) => i !== at);
    return;
  }
  const { top, left, bottom, right } = sheet.autoFilter.range;
  const value = `${quotedSheetName(sheet.name)}!${absoluteName(top, left)}${top === bottom && left === right ? '' : `:${absoluteName(bottom, right)}`}`;
  if (at >= 0) {
    if (names[at].value === value) return;
    workbook.definedNames = names.map((name, i) => (i === at ? { ...name, value } : name));
  } else workbook.definedNames = names.concat([{ name: FILTER_NAME, value, localSheet: index, hidden: true }]);
}

function absoluteName(row: number, col: number): string {
  const name = cellName(row, col);
  const digits = name.search(/\d/);
  return `$${name.slice(0, digits)}$${name.slice(digits)}`;
}

/** Hides or shows rows; returns the rows whose state changed. */
function setRowsHidden(sheet: Sheet, rows: number[] | undefined, hidden: boolean): number[] {
  const changed: number[] = [];
  if (!rows) return changed;
  for (let i = 0; i < rows.length; i++) {
    const index = rows[i];
    const current = sheet.rows.get(index);
    if ((current ? current.hidden : false) === hidden) continue;
    const next: Row = current ? { ...current, hidden } : { ...newRow(), hidden };
    if (isPlainRow(next) && !next.cells.some(() => true)) sheet.rows.delete(index);
    else sheet.rows.set(index, next);
    (sheet.changedRows || (sheet.changedRows = new Set())).add(index);
    changed.push(index);
  }
  return changed;
}

/** Kinds of anchored parts that hold positions of cells which a sort cannot take along. */
const SORT_BLOCKERS = ['comments', 'pivotTable', 'table'];

const intersects = (a: CellRange, b: CellRange) => a.top <= b.bottom && a.bottom >= b.top && a.left <= b.right && a.right >= b.left;

/** The cells a reference names after "!": a cell, a range, whole columns or whole rows. */
const QUALIFIED_CELLS = /^(?:\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d{1,7}:\$?\d{1,7})(?![A-Za-z0-9_.])/;
/** Longest text QUALIFIED_CELLS can match, and one character more for what must not follow. */
const QUALIFIED_LENGTH = 24;

/**
 * Moves the relative references of a formula by rows, as shiftFormula() does, but leaves
 * references that name a sheet ("Other!A1", "'My data'!A1:B2", also the formula's own sheet
 * by name) as they are. That is what Excel does to the formulas of rows a sort moves: what a
 * row reads beside itself goes along, what it reads by sheet name stays put.
 *
 * The text is cut at the cells of every reference with a sheet name; the pieces between
 * them are shifted one by one. Text in quotes and in brackets is never cut.
 */
export function shiftUnqualified(text: string, dRow: number): string {
  if (dRow === 0 || text.indexOf('!') < 0) return shiftFormula(text, dRow, 0);
  const n = text.length;
  let out = '';
  let start = 0;
  let i = 0;
  while (i < n) {
    const c = text.charCodeAt(i);
    if (c === 34 || c === 39) {
      // A string, or a sheet name in apostrophes: a doubled quote stays inside.
      let j = i + 1;
      for (; j < n; j++) {
        if (text.charCodeAt(j) !== c) continue;
        if (text.charCodeAt(j + 1) === c) j++;
        else break;
      }
      i = j + 1;
    } else if (c === 91) {
      // A workbook index or a table reference, brackets inside it included.
      let depth = 0;
      let j = i;
      for (; j < n; j++) {
        const d = text.charCodeAt(j);
        if (d === 91) depth++;
        else if (d === 93 && --depth === 0) break;
      }
      i = j + 1;
    } else if (c === 33) {
      const cells = QUALIFIED_CELLS.exec(text.substr(i + 1, QUALIFIED_LENGTH));
      if (cells) {
        out += shiftFormula(text.slice(start, i + 1), dRow, 0) + cells[0];
        i += 1 + cells[0].length;
        start = i;
      } else i++;
    } else i++;
  }
  return start === 0 ? shiftFormula(text, dRow, 0) : out + shiftFormula(text.slice(start), dRow, 0);
}

/** The box of a note moved by rows with its cell; it stays on the sheet, as tall as it was where it can. */
function shiftAnchorRows(anchor: DrawingAnchor, by: number): DrawingAnchor {
  const from = anchor.from;
  if (by === 0 || anchor.kind === 'absolute' || !from) return anchor;
  const to = anchor.kind === 'twoCell' ? anchor.to : undefined;
  // Not above the first row, and not below the last one.
  let delta = Math.max(by, -from.row);
  delta = Math.min(delta, MAX_ROWS - 1 - (to ? Math.max(to.row, from.row) : from.row));
  const moved: DrawingAnchor = { ...anchor, from: { ...from, row: from.row + delta } };
  if (to) moved.to = { ...to, row: to.row + delta };
  return moved;
}

/** Moves the cells of rows within a range (see the operation "sortRange"). */
function sortRows(workbook: Workbook, index: number, range: CellRange, order: number[]): Op {
  const sheet = sheetAt(workbook, index);
  checkCell(range.top, range.left);
  checkCell(range.bottom, range.right);
  const height = range.bottom - range.top + 1;
  const width = range.right - range.left + 1;
  if (height < 1 || width < 1 || !Array.isArray(order) || order.length !== height) throw new RangeError('Invalid sort order');
  const seen = new Uint8Array(height);
  for (let i = 0; i < height; i++) {
    const from = order[i];
    if (!(from >= 0 && from < height) || Math.floor(from) !== from || seen[from]) throw new RangeError('Invalid sort order');
    seen[from] = 1;
  }
  // Where each source row goes; rows that stay need no work.
  const target = new Int32Array(height);
  let moved = 0;
  for (let i = 0; i < height; i++) {
    target[order[i]] = i;
    if (order[i] !== i) moved++;
  }

  // What is placed by cell position in other parts would stay behind while its cell moves.
  let stuck = (sheet.anchored || []).filter((part) => SORT_BLOCKERS.indexOf(part) >= 0);
  // Notes go with their rows (see below). Comments are in the way only where the model does not hold all
  // of them, or cannot change them: the file has comments the reader could not show (a second one on a
  // cell, one too many), or its comment parts are not the sheet's alone.
  const unread = (sheet.warnings || []).some((warning) => /comment/i.test(warning));
  const comments = stuck.indexOf('comments') >= 0 && (unread || sheet.notesBlocked !== undefined);
  stuck = stuck.filter((part) => part !== 'comments');
  if (comments) stuck.unshift('comments');
  if (stuck.length > 0) throw new RangeError(`Cannot sort in a sheet with ${stuck.join(', ')}: not supported yet`);
  // The notes of the rows that move, each with the cell it is on now. A conversation is kept in parts
  // that are written back as they are, by cell: it cannot follow its row, so its range is not sorted.
  const edits = sheet.structuralEdits;
  const editCount = edits ? edits.length : 0;
  let movedNotes: Map<number, { row: number; col: number }> | undefined;
  const sheetComments = sheet.comments;
  if (sheetComments) {
    for (let i = 0; i < sheetComments.length; i++) {
      const comment = sheetComments[i];
      if (!comment || !(comment.row >= 0 && comment.row < MAX_ROWS && comment.col >= 0 && comment.col < MAX_COLS)) continue;
      const cell = commentCellNow(comment, edits);
      if (!cell || cell.row < range.top || cell.row > range.bottom || cell.col < range.left || cell.col > range.right) continue;
      if (comment.threaded) throw new RangeError('Cannot sort a range that contains threaded comments');
      if (target[cell.row - range.top] !== cell.row - range.top) (movedNotes || (movedNotes = new Map())).set(i, cell);
    }
  }
  // Like Excel: merged cells cannot be told apart from their rows.
  for (const merge of sheet.merges) if (intersects(merge, range)) throw new RangeError('Cannot sort a range that contains merged cells');
  // An array formula covers its range as one unit; one that fills a single cell moves with it.
  for (const [rowIndex, row] of sheet.rows) {
    row.cells.forEach((cell, col) => {
      const formula = cell.formula;
      if (!formula || (formula.type !== 'array' && formula.type !== 'dataTable')) return;
      const area = (formula.ref !== undefined && parseRange(formula.ref)) || { top: rowIndex, left: col, bottom: rowIndex, right: col };
      if (!intersects(area, range)) return;
      const single = formula.type === 'array' && area.top === area.bottom && area.left === area.right;
      if (!single) throw new RangeError(`Cannot sort through the array formula in ${formula.ref === undefined ? cellName(rowIndex, col) : formula.ref}`);
    });
  }

  const snapshot = snapshotOf(sheet);
  const rows = new Map(sheet.rows);
  let weight = moved;

  // A shared formula describes its cells by where they are; the groups a moved cell belongs to are written out.
  let groups: Set<number> | undefined;
  for (let i = 0; i < height; i++) {
    if (order[i] === i) continue;
    const row = rows.get(range.top + i);
    if (!row) continue;
    const last = Math.min(range.right, row.cells.length - 1);
    for (let col = range.left; col <= last; col++) {
      const cell = row.cells[col];
      const formula = cell ? cell.formula : undefined;
      if (formula && formula.type === 'shared' && formula.si !== undefined) (groups || (groups = new Set())).add(formula.si);
    }
  }
  if (groups) {
    const masters = sharedMasters(sheet);
    const wanted = groups;
    for (const [rowIndex, row] of sheet.rows) {
      let cells: Cell[] | undefined;
      row.cells.forEach((cell, col) => {
        const formula = cell.formula;
        if (!formula || formula.type !== 'shared' || formula.si === undefined || !wanted.has(formula.si)) return;
        if (!cells) cells = row.cells.slice();
        cells[col] = explicitFormula(cell, rowIndex, col, masters);
      });
      if (cells) {
        rows.set(rowIndex, { ...row, cells });
        (sheet.changedRows || (sheet.changedRows = new Set())).add(rowIndex);
        weight++;
      }
    }
  }

  // The cells of every row that moves, taken before any row is written.
  const taken: (Cell | undefined)[][] = new Array(height);
  for (let i = 0; i < height; i++) {
    if (target[i] === i) continue;
    const row = rows.get(range.top + i);
    taken[i] = row ? row.cells.slice(range.left, range.right + 1) : [];
  }
  const changedRows = sheet.changedRows || (sheet.changedRows = new Set());
  for (let i = 0; i < height; i++) {
    const from = order[i];
    if (from === i) continue;
    const at = range.top + i;
    const current = rows.get(at);
    const cells = current ? current.cells.slice() : [];
    for (let col = range.left; col <= range.right && col < cells.length; col++) delete cells[col];
    const source = taken[from];
    for (let k = 0; k < source.length; k++) {
      let cell = source[k];
      if (!cell) continue;
      const formula = cell.formula;
      if (formula) {
        // As when a cell is copied: what it refers to relative to itself is still that far from it.
        // What it refers to by sheet name stays where it is.
        const text = shiftUnqualified(formula.text, i - from);
        const single = formula.type === 'array' && formula.ref !== undefined;
        if (text !== formula.text || single) cell = { ...cell, formula: { ...formula, text, ...(single ? { ref: cellName(at, range.left + k) } : {}) } };
      }
      cells[range.left + k] = cell;
    }
    while (cells.length > 0 && !(cells.length - 1 in cells)) cells.length--;
    const next: Row = current ? { ...current, cells } : newRow(cells);
    if (isPlainRow(next) && cells.length === 0) rows.delete(at);
    else rows.set(at, next);
    changedRows.add(at);
  }
  sheet.rows = rows;

  // A link on one cell goes where its cell goes.
  let links: Hyperlink[] | undefined;
  sheet.hyperlinks.forEach((link, i) => {
    const at = link.range;
    if (at.top !== at.bottom || at.top < range.top || at.top > range.bottom || at.left < range.left || at.right > range.right) return;
    const row = range.top + target[at.top - range.top];
    if (row === at.top) return;
    if (!links) links = sheet.hyperlinks.slice();
    links[i] = { ...link, range: { ...at, top: row, bottom: row } };
  });
  if (links) {
    sheet.hyperlinks = links;
    sheet.hyperlinksModified = true;
  }
  // A note goes where its cell goes. It is held from now on like a note that was put on that cell (see
  // SheetComment.since); a comment of the file remembers the cell the file has it on (SheetComment.from).
  if (movedNotes && sheetComments) {
    snapshot.comments = { list: sheetComments };
    const next = sheetComments.slice();
    for (const [i, cell] of movedNotes) {
      const comment = sheetComments[i];
      const row = range.top + target[cell.row - range.top];
      const placed: SheetComment = { ...comment, row, col: cell.col, since: editCount };
      if (comment.since === undefined && comment.from === undefined) placed.from = { row: comment.row, col: comment.col };
      if (comment.box) placed.box = shiftAnchorRows(adjustDrawingAnchor(comment.box, commentEdits(comment, edits)), row - cell.row);
      next[i] = placed;
    }
    sheet.comments = next;
    sheet.commentsModified = true;
    weight += movedNotes.size;
  }
  sheet.modified = true;
  return { type: 'restore', sheet: index, snapshot, others: [], definedNames: workbook.definedNames, weight };
}

// ---------------------------------------------------------------------------
// Sheets

/**
 * Kinds of parts that name sheets in text the engine does not rewrite: the ranges of a chart
 * and the source of a pivot table. A workbook with any keeps its sheets and their names.
 */
const NAME_HOLDERS = ['chart', 'pivotTable'];

/**
 * Form controls are linked to cells by formulas in parts that are written back as they are
 * (the legacy drawing, control properties, <controlPr>): those would go on naming a sheet by
 * the name it had. The kinds are the ones the reader notes in Sheet.anchored.
 */
const CONTROL_KINDS = ['formControl', 'ctrlProp', 'control'];

function nameHolder(workbook: Workbook): string | undefined {
  for (const sheet of workbook.sheets) {
    const anchored = sheet.anchored;
    if (!anchored) continue;
    for (const kind of NAME_HOLDERS) if (anchored.indexOf(kind) >= 0 && !(kind === 'chart' && chartsFollow(workbook))) return kind;
    for (const kind of CONTROL_KINDS) if (anchored.indexOf(kind) >= 0) return 'form controls';
  }
  return undefined;
}

/**
 * Sheet names are told apart without regard to case (see foldSheetName()). In a workbook
 * that has two sheets whose names differ by no more than that, a reference cannot be told to
 * belong to one of them: the sheets of such a workbook are left as they are. Returns why,
 * undefined for every other workbook.
 */
function sheetNameClash(workbook: Workbook): string | undefined {
  const seen = new Map<string, string>();
  for (const sheet of workbook.sheets) {
    const key = foldSheetName(sheet.name);
    const other = seen.get(key);
    if (other !== undefined) return `Cannot change the sheets: "${other}" and "${sheet.name}" differ only by case`;
    seen.set(key, sheet.name);
  }
  return undefined;
}

/**
 * Why a sheet cannot be called `name`; undefined when it can. `except` is the index of the
 * sheet that is being renamed: a sheet may take its own name written in other letter case.
 */
export function sheetNameProblem(workbook: Workbook, name: string, except?: number): string | undefined {
  try {
    checkSheetName(name);
  } catch (error) {
    return (error as Error).message;
  }
  const wanted = foldSheetName(name);
  for (let i = 0; i < workbook.sheets.length; i++) {
    if (i !== except && foldSheetName(workbook.sheets[i].name) === wanted) return `Cannot name the sheet "${name}": another sheet has that name`;
  }
  return undefined;
}

/** Why a sheet cannot be renamed (to `name`, when one is given); undefined when it can. Looks at no cell. */
export function renameSheetProblem(workbook: Workbook, index: number, name?: string): string | undefined {
  const sheet = workbook.sheets[index];
  if (!sheet) return `No sheet at index ${index}`;
  if (sheet.kind !== 'worksheet') return `Cannot rename a ${sheet.kind}: not supported yet`;
  const clash = sheetNameClash(workbook);
  if (clash) return clash;
  const holder = nameHolder(workbook);
  if (holder) return `Cannot rename a sheet in a workbook with ${holder}: not supported yet`;
  return name === undefined ? undefined : sheetNameProblem(workbook, name, index);
}

/** Why a sheet cannot be deleted; undefined when it can. Looks at no cell. */
export function deleteSheetProblem(workbook: Workbook, index: number): string | undefined {
  const sheet = workbook.sheets[index];
  if (!sheet) return `No sheet at index ${index}`;
  if (sheet.kind !== 'worksheet') return `Cannot delete a ${sheet.kind}: not supported yet`;
  const clash = sheetNameClash(workbook);
  if (clash) return clash;
  const holder = nameHolder(workbook);
  if (holder) return `Cannot delete a sheet in a workbook with ${holder}: not supported yet`;
  // Formulas elsewhere name the columns of a table by the table's name, which is not rewritten.
  if ((sheet.anchored || []).indexOf('table') >= 0) return 'Cannot delete a sheet with table: not supported yet';
  if (!workbook.sheets.some((other) => other !== sheet && other.state === 'visible')) return 'Cannot delete the last visible sheet';
  return undefined;
}

/** Why a sheet cannot be inserted at an index under a name; undefined when it can. */
export function insertSheetProblem(workbook: Workbook, at: number, name: string): string | undefined {
  if (!(at >= 0 && at <= workbook.sheets.length) || Math.floor(at) !== at) return `Cannot insert a sheet at index ${at}`;
  return sheetNameClash(workbook) || sheetNameProblem(workbook, name);
}

/** What giving sheets other names (see renameSheetsInFormula()) changes outside the list of sheets. */
interface SheetNameChanges {
  cells: { sheet: Sheet; cells: CellChange[] }[];
  links: { sheet: Sheet; hyperlinks: Hyperlink[] }[];
  /** The workbook's own list when no name changed. */
  definedNames: DefinedName[];
  /** Cells and names changed. */
  count: number;
}

/**
 * Works out what a change of sheet names does to the formulas, defined names and links of a
 * workbook, without changing anything. `skip` is a sheet whose content does not matter (it
 * is being deleted). Throws when a formula that may name a sheet cannot be read to its end:
 * it would be left naming the sheet as it was.
 */
function sheetNameChanges(workbook: Workbook, map: SheetNameMap, what: string, skip?: Sheet): SheetNameChanges {
  const out: SheetNameChanges = { cells: [], links: [], definedNames: workbook.definedNames, count: 0 };
  const rewritten = (text: string, where: () => string): string => {
    // No sheet is named without a "!"; most formulas are done with here.
    if (text.indexOf('!') < 0) return text;
    if (formulaNamesSheetsUnread(text)) throw new RangeError(`Cannot ${what} the sheet: the formula of ${where()} cannot be read`);
    return renameSheetsInFormula(text, map);
  };
  for (const sheet of workbook.sheets) {
    if (sheet === skip) continue;
    const changes: CellChange[] = [];
    for (const [index, row] of sheet.rows) {
      row.cells.forEach((cell, col) => {
        const formula = cell.formula;
        // A cell that only points at its shared formula has no text; the cell that holds the text is rewritten.
        if (!formula || formula.text === '') return;
        const text = rewritten(formula.text, () => `${sheet.name}!${cellName(index, col)}`);
        if (text !== formula.text) changes.push({ row: index, col, cell: { ...cell, formula: { ...formula, text } } });
      });
    }
    if (changes.length > 0) {
      out.cells.push({ sheet, cells: changes });
      out.count += changes.length;
    }
    let links: Hyperlink[] | undefined;
    sheet.hyperlinks.forEach((link, i) => {
      if (link.location === undefined) return;
      // The place a link goes to is not a formula: one that cannot be read is left alone.
      const location = renameSheetsInFormula(link.location, map);
      if (location === link.location) return;
      if (!links) links = sheet.hyperlinks.slice();
      links[i] = { ...link, location };
    });
    if (links) out.links.push({ sheet, hyperlinks: links });
  }
  let names: DefinedName[] | undefined;
  workbook.definedNames.forEach((name, i) => {
    if (skip && name.localSheet !== undefined && workbook.sheets[name.localSheet] === skip) return;
    const value = rewritten(name.value, () => `the name ${name.name}`);
    if (value === name.value) return;
    if (!names) names = workbook.definedNames.slice();
    names[i] = { ...name, value };
    out.count++;
  });
  if (names) out.definedNames = names;
  return out;
}

function sheetsSnapshot(workbook: Workbook): SheetsSnapshot {
  return {
    sheets: workbook.sheets,
    cells: [],
    links: [],
    definedNames: workbook.definedNames,
    activeSheet: workbook.activeSheet,
    removedSheetNames: workbook.removedSheetNames,
    removedSheetNeighbours: workbook.removedSheetNeighbours,
  };
}

/** Puts changed cells and links in place and notes in `undo` what was there. */
function applySheetNameChanges(changes: Pick<SheetNameChanges, 'cells' | 'links'>, undo: SheetsSnapshot): void {
  for (const change of changes.cells) {
    undo.cells.push({ sheet: change.sheet, cells: setCells(change.sheet, change.cells, true) });
    change.sheet.modified = true;
  }
  for (const change of changes.links) {
    undo.links.push({ sheet: change.sheet, hyperlinks: change.sheet.hyperlinks });
    change.sheet.hyperlinks = change.hyperlinks;
    change.sheet.hyperlinksModified = true;
    change.sheet.modified = true;
  }
}

function renameSheetOp(workbook: Workbook, index: number, name: string): Op {
  const problem = renameSheetProblem(workbook, index, name);
  if (problem) throw new RangeError(problem);
  const sheet = workbook.sheets[index];
  const old = foldSheetName(sheet.name);
  const renames: SheetNameMap = (written) => (foldSheetName(written) === old ? name : undefined);
  const changes = sheetNameChanges(workbook, renames, 'rename');
  // Worked out before anything is changed: a chart whose reference cannot be read refuses the rename.
  const charts = chartsRenamed(workbook, renames);

  const undo = sheetsSnapshot(workbook);
  undo.renamed = { sheet, name: sheet.name, nameInFile: sheet.nameInFile };
  applySheetNameChanges(changes, undo);
  // Rules of the file name sheets as the file does (see nameInFile below); those changed since name them as they are called.
  const rules = swapRuleLists(liveRulesRenamed(workbook, renames).concat(charts));
  if (rules.length > 0) undo.rules = rules;
  workbook.definedNames = changes.definedNames;
  // What is kept the way the file has it goes on naming the sheet as the file does.
  if (isFileSheet(sheet) && sheet.nameInFile === undefined && sheet.name !== name) sheet.nameInFile = sheet.name;
  sheet.name = name;
  workbook.sheetsModified = true;
  return { type: 'restoreSheets', sheet: index, state: undo, weight: changes.count };
}

function deleteSheetOp(workbook: Workbook, index: number): Op {
  const problem = deleteSheetProblem(workbook, index);
  if (problem) throw new RangeError(problem);
  const sheet = workbook.sheets[index];
  const gone = foldSheetName(sheet.name);
  const before = workbook.sheets;
  const map: SheetNameMap = (written) => (foldSheetName(written) === gone ? null : undefined);
  // A 3D reference that ends at the sheet ends at the sheet beside it from now on.
  map.runs = {
    neighbour: (_name, step) => {
      const beside = before[index + step];
      return beside ? beside.name : undefined;
    },
    index: (name) => {
      const wanted = foldSheetName(name);
      for (let i = 0; i < before.length; i++) if (foldSheetName(before[i].name) === wanted) return i;
      return -1;
    },
  };
  const changes = sheetNameChanges(workbook, map, 'delete', sheet);
  // The charts of the other sheets that show cells of this one show "#REF!" from now on.
  const charts = chartsRenamed(workbook, map, sheet);

  const undo = sheetsSnapshot(workbook);
  applySheetNameChanges(changes, undo);
  const rules = swapRuleLists(liveRulesRenamed(workbook, map, sheet).concat(charts));
  if (rules.length > 0) undo.rules = rules;
  // The names of the sheet go with it; those of the sheets after it follow their sheets.
  const names: DefinedName[] = [];
  let namesChanged = changes.definedNames !== undo.definedNames;
  for (const name of changes.definedNames) {
    if (name.localSheet === undefined || name.localSheet < index) names.push(name);
    else {
      namesChanged = true;
      if (name.localSheet > index) names.push({ ...name, localSheet: name.localSheet - 1 });
    }
  }
  if (namesChanged) workbook.definedNames = names;
  const sheets = workbook.sheets.filter((other) => other !== sheet);
  workbook.sheets = sheets;
  let active = workbook.activeSheet;
  if (active > index) active--;
  else if (active === index || !sheets[active]) {
    // The next visible sheet takes its place, or the one before when it was the last.
    active = -1;
    for (let i = index; i < sheets.length && active < 0; i++) if (sheets[i].state === 'visible') active = i;
    for (let i = Math.min(index, sheets.length) - 1; i >= 0 && active < 0; i--) if (sheets[i].state === 'visible') active = i;
  }
  workbook.activeSheet = active;
  // What stood beside a sheet deleted earlier and was this sheet is now the sheet beside this one.
  const beside: SheetNeighbours = { previous: before[index - 1], next: before[index + 1] };
  let neighbours = (workbook.removedSheetNeighbours || []).map((entry) =>
    entry.previous === sheet || entry.next === sheet
      ? { previous: entry.previous === sheet ? beside.previous : entry.previous, next: entry.next === sheet ? beside.next : entry.next }
      : entry,
  );
  if (isFileSheet(sheet)) {
    workbook.removedSheetNames = (workbook.removedSheetNames || []).concat([sheet.nameInFile !== undefined ? sheet.nameInFile : sheet.name]);
    neighbours = neighbours.concat([beside]);
  }
  if (neighbours.length > 0) workbook.removedSheetNeighbours = neighbours;
  workbook.sheetsModified = true;
  return { type: 'restoreSheets', sheet: index, state: undo, weight: changes.count + sheet.rows.size };
}

function insertSheetOp(workbook: Workbook, at: number, name: string): Op {
  const problem = insertSheetProblem(workbook, at, name);
  if (problem) throw new RangeError(problem);
  let sheetId = 0;
  for (const other of workbook.sheets) if (other.sheetId > sheetId) sheetId = other.sheetId;
  // An empty part name: the sheet is not in the file, the writer gives it a part.
  const sheet = createSheet({ name, sheetId: sheetId + 1, state: 'visible', kind: 'worksheet', part: '' });
  sheet.modified = true;

  const undo = sheetsSnapshot(workbook);
  const sheets = workbook.sheets.slice();
  sheets.splice(at, 0, sheet);
  workbook.sheets = sheets;
  if (workbook.definedNames.some((defined) => defined.localSheet !== undefined && defined.localSheet >= at)) {
    workbook.definedNames = workbook.definedNames.map((defined) => (defined.localSheet !== undefined && defined.localSheet >= at ? { ...defined, localSheet: defined.localSheet + 1 } : defined));
  }
  // The sheet that was shown stays the one that is.
  if (workbook.activeSheet >= at) workbook.activeSheet++;
  workbook.sheetsModified = true;
  return { type: 'restoreSheets', sheet: at, state: undo };
}

/** Why a sheet cannot be moved to an index; undefined when it can. */
export function moveSheetProblem(workbook: Workbook, from: number, to: number): string | undefined {
  if (!workbook.sheets[from]) return `No sheet at index ${from}`;
  if (!(to >= 0 && to < workbook.sheets.length) || Math.floor(to) !== to) return `Cannot move a sheet to index ${to}`;
  return sheetNameClash(workbook);
}

/** Why a sheet cannot be given a state; undefined when it can. */
export function sheetStateProblem(workbook: Workbook, index: number, state: SheetState): string | undefined {
  const sheet = workbook.sheets[index];
  if (!sheet) return `No sheet at index ${index}`;
  if (state !== 'visible' && state !== 'hidden' && state !== 'veryHidden') return 'A sheet is visible, hidden or veryHidden';
  const clash = sheetNameClash(workbook);
  if (clash) return clash;
  if (state !== 'visible' && sheet.state === 'visible' && !workbook.sheets.some((other) => other !== sheet && other.state === 'visible')) return 'Cannot hide the last visible sheet';
  return undefined;
}

/** Why a sheet cannot be copied (under `name`, when one is given, to `at`); undefined when it can. */
export function copySheetProblem(workbook: Workbook, index: number, at?: number, name?: string): string | undefined {
  const sheet = workbook.sheets[index];
  if (!sheet) return `No sheet at index ${index}`;
  if (sheet.kind !== 'worksheet') return `Cannot copy a ${sheet.kind}: not supported yet`;
  if (at !== undefined && (!(at >= 0 && at <= workbook.sheets.length) || Math.floor(at) !== at)) return `Cannot insert a sheet at index ${at}`;
  const clash = sheetNameClash(workbook);
  if (clash) return clash;
  return name === undefined ? undefined : sheetNameProblem(workbook, name);
}

/** Why the tab of a sheet cannot be coloured; undefined when it can. */
export function tabColorProblem(workbook: Workbook, index: number): string | undefined {
  const sheet = workbook.sheets[index];
  if (!sheet) return `No sheet at index ${index}`;
  if (sheet.kind !== 'worksheet') return `Cannot colour the tab of a ${sheet.kind}: not supported yet`;
  return sheetNameClash(workbook);
}

/** The defined names with the sheets they are local to looked up again after the list of sheets changed. */
function rescopeNames(workbook: Workbook, before: Sheet[]): void {
  const after = workbook.sheets;
  let names: DefinedName[] | undefined;
  workbook.definedNames.forEach((name, i) => {
    if (name.localSheet === undefined) return;
    const sheet = before[name.localSheet];
    const now = sheet ? after.indexOf(sheet) : -1;
    if (now < 0 || now === name.localSheet) return;
    if (!names) names = workbook.definedNames.slice();
    names[i] = { ...name, localSheet: now };
  });
  if (names) workbook.definedNames = names;
}

function moveSheetOp(workbook: Workbook, from: number, to: number): Op {
  const problem = moveSheetProblem(workbook, from, to);
  if (problem) throw new RangeError(problem);
  const undo = sheetsSnapshot(workbook);
  const before = workbook.sheets;
  const active = before[workbook.activeSheet];
  const sheets = before.slice();
  const moved = sheets.splice(from, 1)[0];
  sheets.splice(to, 0, moved);
  workbook.sheets = sheets;
  rescopeNames(workbook, before);
  if (active) workbook.activeSheet = sheets.indexOf(active);
  workbook.sheetsModified = true;
  return { type: 'restoreSheets', sheet: from, state: undo };
}

function setSheetStateOp(workbook: Workbook, index: number, state: SheetState): Op {
  const problem = sheetStateProblem(workbook, index, state);
  if (problem) throw new RangeError(problem);
  const sheet = workbook.sheets[index];
  const undo = sheetsSnapshot(workbook);
  undo.shown = { sheet, state: sheet.state };
  sheet.state = state;
  if (state !== 'visible' && workbook.activeSheet === index) {
    // The next visible sheet takes its place, or the one before when there is none after it.
    const sheets = workbook.sheets;
    let active = -1;
    for (let i = index + 1; i < sheets.length && active < 0; i++) if (sheets[i].state === 'visible') active = i;
    for (let i = index - 1; i >= 0 && active < 0; i--) if (sheets[i].state === 'visible') active = i;
    if (active >= 0) workbook.activeSheet = active;
  }
  workbook.sheetsModified = true;
  return { type: 'restoreSheets', sheet: index, state: undo };
}

function copySheetOp(workbook: Workbook, at: number, copy: Sheet, names: DefinedName[]): Op {
  if (!copy || typeof copy !== 'object' || !(copy.rows instanceof Map) || copy.part !== '' || copy.kind !== 'worksheet') throw new RangeError('The copy of a sheet is a worksheet that is in no file');
  if (workbook.sheets.indexOf(copy) >= 0) throw new RangeError('The sheet is in the workbook already');
  const problem = insertSheetProblem(workbook, at, copy.name);
  if (problem) throw new RangeError(problem);
  const undo = sheetsSnapshot(workbook);
  const sheets = workbook.sheets.slice();
  sheets.splice(at, 0, copy);
  workbook.sheets = sheets;
  const shifted = workbook.definedNames.map((defined) => (defined.localSheet !== undefined && defined.localSheet >= at ? { ...defined, localSheet: defined.localSheet + 1 } : defined));
  workbook.definedNames = shifted.concat((Array.isArray(names) ? names : []).map((defined) => ({ ...defined, localSheet: at })));
  // The sheet that was shown stays the one that is.
  if (workbook.activeSheet >= at) workbook.activeSheet++;
  workbook.sheetsModified = true;
  // The step holds the list of sheets as it was, which the copy is not in: it weighs nothing. (Undone, the
  // step that brings the copy back holds it, on the stack of what can be done again, which is not weighed.)
  return { type: 'restoreSheets', sheet: at, state: undo };
}

function restoreSheets(workbook: Workbook, index: number, state: SheetsSnapshot, weight: number | undefined): Op {
  const undo = sheetsSnapshot(workbook);
  if (state.shown) {
    undo.shown = { sheet: state.shown.sheet, state: state.shown.sheet.state };
    state.shown.sheet.state = state.shown.state;
  }
  if (state.renamed) {
    const sheet = state.renamed.sheet;
    undo.renamed = { sheet, name: sheet.name, nameInFile: sheet.nameInFile };
    sheet.name = state.renamed.name;
    if (state.renamed.nameInFile === undefined) delete sheet.nameInFile;
    else sheet.nameInFile = state.renamed.nameInFile;
  }
  // Last first: the order the changes are undone in.
  for (let i = state.cells.length - 1; i >= 0; i--) {
    const change = state.cells[i];
    undo.cells.push({ sheet: change.sheet, cells: setCells(change.sheet, change.cells, true) });
    change.sheet.modified = true;
  }
  for (let i = state.links.length - 1; i >= 0; i--) {
    const change = state.links[i];
    undo.links.push({ sheet: change.sheet, hyperlinks: change.sheet.hyperlinks });
    change.sheet.hyperlinks = change.hyperlinks;
    change.sheet.hyperlinksModified = true;
    change.sheet.modified = true;
  }
  const rules = swapRuleLists(state.rules);
  if (rules.length > 0) undo.rules = rules;
  workbook.sheets = state.sheets;
  workbook.definedNames = state.definedNames;
  workbook.activeSheet = state.activeSheet;
  if (state.removedSheetNames === undefined) delete workbook.removedSheetNames;
  else workbook.removedSheetNames = state.removedSheetNames;
  if (state.removedSheetNeighbours === undefined) delete workbook.removedSheetNeighbours;
  else workbook.removedSheetNeighbours = state.removedSheetNeighbours;
  workbook.sheetsModified = true;
  return { type: 'restoreSheets', sheet: index, state: undo, weight };
}

/** Column settings sorted, clipped to the sheet and free of overlaps (an earlier range wins). */
function normalizeColumns(columns: ColumnRange[]): ColumnRange[] {
  const sorted = columns
    .map((range) => ({ ...range, min: Math.max(Math.floor(range.min), 0), max: Math.min(Math.floor(range.max), MAX_COLS - 1) }))
    .filter((range) => range.min <= range.max)
    .sort((a, b) => a.min - b.min);
  const out: ColumnRange[] = [];
  for (const range of sorted) {
    const last = out[out.length - 1];
    if (last && range.min <= last.max) {
      if (range.max <= last.max) continue;
      range.min = last.max + 1;
    }
    out.push(range);
  }
  return out;
}

/** Applies an operation to the workbook and returns the operation that undoes it. */
export function applyOp(workbook: Workbook, op: Op): Op {
  // Counted before anything is changed: a file that is being written from this workbook is of no one moment now.
  touchModel(workbook);
  switch (op.type) {
    case 'setCells': {
      const sheet = sheetAt(workbook, op.sheet);
      const extent = { rowCount: sheet.rowCount, colCount: sheet.colCount };
      const undo = setCells(sheet, op.cells);
      if (op.extent) {
        sheet.rowCount = op.extent.rowCount;
        sheet.colCount = op.extent.colCount;
      }
      sheet.modified = true;
      const grew = sheet.rowCount !== extent.rowCount || sheet.colCount !== extent.colCount;
      return { type: 'setCells', sheet: op.sheet, cells: undo, ...(grew ? { extent } : {}) };
    }
    case 'setColumns': {
      const sheet = sheetAt(workbook, op.sheet);
      const previous = sheet.columns;
      sheet.columns = normalizeColumns(op.columns);
      sheet.columnsModified = true;
      sheet.modified = true;
      return { type: 'setColumns', sheet: op.sheet, columns: previous };
    }
    case 'setRows': {
      const sheet = sheetAt(workbook, op.sheet);
      for (const change of op.rows) checkCell(change.row, 0);
      const undo: { row: number; props: RowProps }[] = [];
      for (const change of op.rows) {
        const current = sheet.rows.get(change.row);
        const before: RowProps = {};
        const next: Row = current ? { ...current } : newRow();
        for (const key of Object.keys(change.props) as (keyof RowProps)[]) {
          (before as Record<string, unknown>)[key] = current ? current[key] : key === 'height' || key === 'style' ? undefined : false;
          const value = change.props[key];
          if (value === undefined && (key === 'height' || key === 'style')) delete next[key];
          else (next as unknown as Record<string, unknown>)[key] = value === undefined ? false : value;
        }
        // A row brought back to having nothing of its own is not kept.
        if (isPlainRow(next) && !next.cells.some(() => true)) sheet.rows.delete(change.row);
        else sheet.rows.set(change.row, next);
        (sheet.changedRows || (sheet.changedRows = new Set())).add(change.row);
        undo.push({ row: change.row, props: before });
      }
      sheet.modified = true;
      return { type: 'setRows', sheet: op.sheet, rows: undo.reverse() };
    }
    case 'setMerges': {
      const sheet = sheetAt(workbook, op.sheet);
      const previous = sheet.merges;
      sheet.merges = op.merges.map((range) => ({ ...range }));
      sheet.mergesModified = true;
      sheet.modified = true;
      return { type: 'setMerges', sheet: op.sheet, merges: previous };
    }
    case 'setAutoFilter': {
      const sheet = sheetAt(workbook, op.sheet);
      for (const list of [op.show, op.hide]) if (list) for (let i = 0; i < list.length; i++) checkCell(list[i], 0);
      const previous = sheet.autoFilter;
      const shown = setRowsHidden(sheet, op.show, false);
      const hidden = setRowsHidden(sheet, op.hide, true);
      if (op.autoFilter) sheet.autoFilter = op.autoFilter;
      else delete sheet.autoFilter;
      sheet.autoFilterModified = true;
      sheet.modified = true;
      syncFilterName(workbook, op.sheet);
      return { type: 'setAutoFilter', sheet: op.sheet, autoFilter: previous, show: hidden, hide: shown };
    }
    case 'setComments': {
      const sheet = sheetAt(workbook, op.sheet);
      if (op.comments !== undefined && !Array.isArray(op.comments)) throw new RangeError('The comments of a sheet are a list');
      const previous = sheet.comments;
      if (op.comments && op.comments.length > 0) sheet.comments = op.comments;
      else delete sheet.comments;
      // Not `modified`: no cell changed, so the sheet part stays as it is and nothing has to be calculated again.
      sheet.commentsModified = true;
      return { type: 'setComments', sheet: op.sheet, comments: previous };
    }
    case 'setDrawings': {
      const sheet = sheetAt(workbook, op.sheet);
      if (op.drawings !== undefined && !Array.isArray(op.drawings)) throw new RangeError('The drawings of a sheet are a list');
      const previous = sheet.drawings;
      // What the sheet came with, noted once: while the list is that one, the writer leaves the parts alone.
      if (!sheet.drawingsRead) sheet.drawingsRead = { list: previous };
      if (op.drawings && op.drawings.length > 0) sheet.drawings = op.drawings;
      else delete sheet.drawings;
      // Not `modified`: no cell changed, so the sheet part stays as it is and nothing has to be calculated again.
      sheet.drawingsModified = true;
      return { type: 'setDrawings', sheet: op.sheet, drawings: previous };
    }
    case 'setDataValidations': {
      const sheet = sheetAt(workbook, op.sheet);
      if (op.validations !== undefined && !Array.isArray(op.validations)) throw new RangeError('The data validations of a sheet are a list');
      const previous = sheet.dataValidations;
      // What the sheet came with, noted once: while the list is that one, the writer leaves the part alone.
      if (!sheet.validationsRead) sheet.validationsRead = { list: previous };
      if (op.validations && op.validations.length > 0) sheet.dataValidations = op.validations;
      else delete sheet.dataValidations;
      // Not `modified`: no cell changed, so nothing has to be calculated again.
      sheet.validationsModified = true;
      return { type: 'setDataValidations', sheet: op.sheet, validations: previous };
    }
    case 'setConditionalFormats': {
      const sheet = sheetAt(workbook, op.sheet);
      if (op.formats !== undefined && !Array.isArray(op.formats)) throw new RangeError('The conditional formats of a sheet are a list');
      const previous = sheet.conditionalFormats;
      // What the sheet came with, noted once: while the list is that one, the writer leaves the part alone.
      if (!sheet.conditionalRead) sheet.conditionalRead = { list: previous };
      if (op.formats && op.formats.length > 0) sheet.conditionalFormats = op.formats;
      else delete sheet.conditionalFormats;
      // Not `modified`: no cell changed, so nothing has to be calculated again.
      sheet.conditionalModified = true;
      return { type: 'setConditionalFormats', sheet: op.sheet, formats: previous };
    }
    case 'sortRange':
      return sortRows(workbook, op.sheet, op.range, op.order);
    case 'insert':
    case 'delete':
      return structural(workbook, op.sheet, { kind: op.type, axis: op.axis, at: op.at, count: op.count });
    case 'restore': {
      const sheet = sheetAt(workbook, op.sheet);
      const snapshot = snapshotOf(sheet);
      if (op.snapshot.comments) snapshot.comments = { list: sheet.comments };
      restoreSnapshot(sheet, op.snapshot);
      sheet.modified = true;
      const others = op.others.map((other) => ({ sheet: other.sheet, cells: setCells(sheetAt(workbook, other.sheet), other.cells) }));
      const otherLinks = (op.otherLinks || []).map((other) => {
        const target = sheetAt(workbook, other.sheet);
        const previous = target.hyperlinks;
        target.hyperlinks = other.hyperlinks;
        target.hyperlinksModified = true;
        target.modified = true;
        return { sheet: other.sheet, hyperlinks: previous };
      });
      const definedNames = workbook.definedNames;
      workbook.definedNames = op.definedNames;
      const rules = swapRuleLists(op.rules);
      return { type: 'restore', sheet: op.sheet, snapshot, others, otherLinks, definedNames, ...(rules.length > 0 ? { rules } : {}), weight: op.weight };
    }
    case 'insertSheet':
      return insertSheetOp(workbook, op.sheet, op.name);
    case 'deleteSheet':
      return deleteSheetOp(workbook, op.sheet);
    case 'renameSheet':
      return renameSheetOp(workbook, op.sheet, op.name);
    case 'moveSheet':
      return moveSheetOp(workbook, op.sheet, op.to);
    case 'setSheetState':
      return setSheetStateOp(workbook, op.sheet, op.state);
    case 'copySheet':
      return copySheetOp(workbook, op.sheet, op.copy, op.names);
    case 'setTabColor': {
      const problem = tabColorProblem(workbook, op.sheet);
      if (problem) throw new RangeError(problem);
      const sheet = sheetAt(workbook, op.sheet);
      const previous = sheet.tabColor;
      // What the sheet came with, noted once: the writer leaves the part alone while the colour is that one.
      if (!sheet.tabColorRead) sheet.tabColorRead = { color: previous };
      if (op.color) sheet.tabColor = op.color;
      else delete sheet.tabColor;
      // Not `modified`: no cell changed, and nothing has to be calculated again.
      sheet.tabColorModified = true;
      workbook.sheetsModified = true;
      return { type: 'setTabColor', sheet: op.sheet, color: previous };
    }
    case 'restoreSheets':
      return restoreSheets(workbook, op.sheet, op.state, op.weight);
  }
}

/**
 * A structural step keeps the rows it replaced. A sheet of 100,000 formula
 * rows costs about 40 MB per step, so steps are dropped by what they hold as
 * well as by their number.
 */
const MAX_UNDO_WEIGHT = 300000;

function weightOf(ops: Op[]): number {
  let weight = 0;
  for (const op of ops) {
    if (op.type === 'restore' || op.type === 'restoreSheets') weight += op.weight || 0;
    else if (op.type === 'setCells') weight += op.cells.length / 16;
    else if (op.type === 'setAutoFilter') weight += ((op.show ? op.show.length : 0) + (op.hide ? op.hide.length : 0)) / 16;
    // The list of comments a step replaced is kept, not the comments themselves.
    else if (op.type === 'setComments') weight += (op.comments ? op.comments.length : 0) / 16;
    else if (op.type === 'setDrawings') weight += (op.drawings ? op.drawings.length : 0) / 16;
    else if (op.type === 'setDataValidations') weight += (op.validations ? op.validations.length : 0) / 16;
    else if (op.type === 'setConditionalFormats') weight += (op.formats ? op.formats.length : 0) / 16;
  }
  return weight;
}

/** Undo and redo over applyOp(). `M` is whatever the caller wants back with each step (a selection, say). */
export class History<M = undefined> {
  private undoStack: { ops: Op[]; before: M; after: M }[] = [];
  private redoStack: { ops: Op[]; before: M; after: M }[] = [];
  /** Rows kept alive by the steps on the undo stack; see MAX_UNDO_WEIGHT. */
  private weight = 0;

  constructor(
    readonly workbook: Workbook,
    private readonly limit = 200,
  ) {}

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  /**
   * Applies operations as one undoable step. If one of them throws, the ones
   * already applied are undone and the error is passed on.
   */
  apply(ops: Op[], before: M, after: M): Op[] {
    if (ops.length === 0) return [];
    const inverse = this.run(ops);
    this.undoStack.push({ ops: inverse, before, after });
    this.weight += weightOf(inverse);
    // The oldest steps go when there are too many of them or they hold too much; the newest always stays.
    while (this.undoStack.length > 1 && (this.undoStack.length > this.limit || this.weight > MAX_UNDO_WEIGHT)) {
      this.weight -= weightOf((this.undoStack.shift() as { ops: Op[] }).ops);
    }
    this.redoStack = [];
    return inverse;
  }

  /** Undoes the last step and returns what was passed as `before` for it. */
  undo(): { meta: M } | undefined {
    const step = this.undoStack.pop();
    if (!step) return undefined;
    this.weight -= weightOf(step.ops);
    this.redoStack.push({ ops: this.run(step.ops), before: step.before, after: step.after });
    return { meta: step.before };
  }

  redo(): { meta: M } | undefined {
    const step = this.redoStack.pop();
    if (!step) return undefined;
    const inverse = this.run(step.ops);
    this.weight += weightOf(inverse);
    this.undoStack.push({ ops: inverse, before: step.before, after: step.after });
    return { meta: step.after };
  }

  /** Applies operations in order and returns their inverses, last first. */
  private run(ops: Op[]): Op[] {
    const inverse: Op[] = [];
    try {
      for (const op of ops) inverse.push(applyOp(this.workbook, op));
    } catch (error) {
      for (let i = inverse.length - 1; i >= 0; i--) applyOp(this.workbook, inverse[i]);
      throw error;
    }
    return inverse.reverse();
  }
}

/** Compares two flat records field by field; undefined and a missing record are the same. */
function sameRecord(a: object | undefined, b: object | undefined): boolean {
  const left = (a || {}) as Record<string, unknown>;
  const right = (b || {}) as Record<string, unknown>;
  for (const key of Object.keys(left)) if (left[key] !== right[key]) return false;
  for (const key of Object.keys(right)) if (left[key] !== right[key]) return false;
  return true;
}

export const sameFormat = (a: CellFormat, b: CellFormat) =>
  a.numFmtId === b.numFmtId &&
  a.fontId === b.fontId &&
  a.fillId === b.fillId &&
  a.borderId === b.borderId &&
  a.xfId === b.xfId &&
  a.quotePrefix === b.quotePrefix &&
  sameRecord(a.alignment, b.alignment) &&
  sameRecord(a.protection, b.protection);

/** How much of a style sheet there was at some moment (see markStyles()). */
export interface StylesMark {
  numFmts: number;
  fonts: number;
  fills: number;
  borders: number;
  cellXfs: number;
  modified: boolean | undefined;
}

/**
 * Notes how far the tables of a style sheet reach. Formats are only ever added at the end
 * (reading typed text may add a date format before anything is known to be written), so
 * resetStyles() can take back whatever was added since, as if it had never been asked for.
 */
export function markStyles(styles: StyleSheet): StylesMark {
  return { numFmts: styles.numFmts.size, fonts: styles.fonts.length, fills: styles.fills.length, borders: styles.borders.length, cellXfs: styles.cellXfs.length, modified: styles.modified };
}

/** Removes the formats added since markStyles(); nothing may refer to them any more. */
export function resetStyles(styles: StyleSheet, mark: StylesMark): void {
  if (styles.numFmts.size > mark.numFmts) {
    // A map keeps its keys in the order they were added.
    const added = Array.from(styles.numFmts.keys()).slice(mark.numFmts);
    for (const id of added) styles.numFmts.delete(id);
  }
  if (styles.fonts.length > mark.fonts) styles.fonts.length = mark.fonts;
  if (styles.fills.length > mark.fills) styles.fills.length = mark.fills;
  if (styles.borders.length > mark.borders) styles.borders.length = mark.borders;
  if (styles.cellXfs.length > mark.cellXfs) styles.cellXfs.length = mark.cellXfs;
  if (mark.modified === undefined) delete styles.modified;
  else styles.modified = mark.modified;
}

/** Excel's limit on cell formats in one workbook. */
export const MAX_CELL_FORMATS = 65430;

/** First id a file may use for its own number formats. */
const FIRST_CUSTOM_FORMAT = 164;

/** Id of a number format code in the style sheet, adding it when it is not there yet. */
export function ensureNumberFormat(styles: StyleSheet, code: string): number {
  if (/^general$/i.test(code)) return 0;
  let highest = FIRST_CUSTOM_FORMAT - 1;
  for (const [id, existing] of styles.numFmts) {
    if (existing === code) return id;
    if (id > highest) highest = id;
  }
  styles.numFmts.set(highest + 1, code);
  styles.modified = true;
  return highest + 1;
}

/** Index of the cell format that is `base` with another number format, adding it when it is not there yet. */
export function ensureCellFormat(styles: StyleSheet, base: number, numFmtId: number): number {
  const source = styles.cellXfs[base] || styles.cellXfs[0];
  if (!source) throw new Error('The style sheet has no cell formats');
  if (source.numFmtId === numFmtId) return styles.cellXfs[base] ? base : 0;
  const wanted: CellFormat = { ...source, numFmtId };
  const found = styles.cellXfs.findIndex((format) => sameFormat(format, wanted));
  if (found >= 0) return found;
  if (styles.cellXfs.length >= MAX_CELL_FORMATS) throw new RangeError('The workbook has too many cell formats');
  styles.modified = true;
  return styles.cellXfs.push(wanted) - 1;
}
