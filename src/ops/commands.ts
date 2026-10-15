// What the editor does, expressed as operations: typing into a cell, clearing
// a range, copying and pasting. These functions only build operations; the
// caller applies them through a History so that each is one undo step.

import { needsDynamic } from '../calc/dynamic';
import { parseFormula } from '../calc/parse';
import { shiftFormula } from '../calc/refs';
import { MAX_PASTE_CELLS, parseTsvLimited, toHtmlTable, toTsv } from '../edit/clipboard';
import { cellEditText, parseCellInput } from '../edit/input';
import { builtinFormatId, formatCell } from '../format/cell-format';
import type { FormatLocale } from '../format/locale';
import { resolveCellStyle } from '../model/sheet/color';
import { MAX_COLS, MAX_ROWS, getCell, parseRange, type Cell, type CellRange, type Sheet, type SheetComment, type Workbook } from '../model/sheet/types';
import { rowRunsWithout, rowsHiddenByFilter } from './filter';
import { DEFAULT_NOTE_AUTHOR, MAX_SHEET_NOTES, cellNow, clearNotes, notesBlocked } from './notes';
import { copiedRulesExtent, copyRules, cutRules, pasteRules, type CopiedRules } from './rules';
import { ensureCellFormat, ensureNumberFormat, explicitFormula, sharedMasters, type CellChange, type Op, type SharedMasters } from './ops';

export interface EditOptions {
  locale?: FormatLocale;
  /**
   * The shared formulas of the sheet (sharedMasters()), for a caller that writes many cells of
   * it at once and works them out once; looked for in the sheet when needed otherwise.
   */
  masters?: SharedMasters;
}

/** Format index of a cell position: the cell's own, else its row's, else its column's. */
export function styleAt(sheet: Sheet, row: number, col: number): number {
  const rowData = sheet.rows.get(row);
  const cell = rowData ? rowData.cells[col] : undefined;
  if (cell) return cell.style;
  if (rowData && rowData.style !== undefined) return rowData.style;
  for (const range of sheet.columns) {
    if (range.min > col) break;
    if (range.max >= col) return range.style !== undefined ? range.style : 0;
  }
  return 0;
}

/** Format a position has without a cell of its own: its row's, else its column's. */
function inheritedStyle(sheet: Sheet, row: number, col: number): number {
  const rowData = sheet.rows.get(row);
  if (rowData && rowData.style !== undefined) return rowData.style;
  for (const range of sheet.columns) {
    if (range.min > col) break;
    if (range.max >= col) return range.style !== undefined ? range.style : 0;
  }
  return 0;
}

/** An emptied cell: gone when the position would look the same without it, otherwise blank with its format. */
export function emptied(sheet: Sheet, row: number, col: number, style: number): Cell | undefined {
  return style === inheritedStyle(sheet, row, col) ? undefined : { kind: 'blank', value: null, style };
}

const isTextFormat = (code: string) => code === '@';

/** Longest text a cell holds. */
const MAX_CELL_TEXT = 32767;

/** Whether the workbook has a defined name of this spelling, whatever its case. */
export function hasDefinedName(workbook: Workbook, word: string): boolean {
  const names = workbook.definedNames;
  if (!names || names.length === 0) return false;
  const wanted = word.toLowerCase();
  const upper = word.toUpperCase();
  for (let i = 0; i < names.length; i++) {
    const name = names[i].name;
    if (name.length === word.length && (name === word || name.toLowerCase() === wanted || name.toUpperCase() === upper)) return true;
  }
  return false;
}

/**
 * The cell that typed text turns into at a position, or undefined when the position ends up empty.
 * Throws RangeError("The formula cannot be read") for a formula that is not written the way the
 * locale writes formulas.
 */
export function cellFromInput(workbook: Workbook, sheet: Sheet, row: number, col: number, text: string, options: EditOptions = {}): Cell | undefined {
  let style = styleAt(sheet, row, col);
  if (text === '') return emptied(sheet, row, col, style);
  const format = resolveCellStyle(workbook.styles, style).numFmt;
  // In a cell formatted as Text, what is typed is the text: "0123" keeps its zero and "=A1" is not a formula.
  if (isTextFormat(format)) return { kind: 'string', value: text.length > MAX_CELL_TEXT ? text.slice(0, MAX_CELL_TEXT) : text, style };
  // The formula that is there now tells which names are functions of the file's own: typed again, they stay those.
  const held = getCell(sheet, row, col);
  const stored = held && held.formula ? held.formula.text : undefined;
  const parsed = parseCellInput(text, { locale: options.locale, date1904: workbook.date1904, isName: (name) => hasDefinedName(workbook, name), stored });
  if (parsed.invalid) throw new RangeError('The formula cannot be read');
  if (parsed.kind === 'blank' && parsed.formula === undefined) return emptied(sheet, row, col, style);
  // Typing "15.03.2024" or "12%" into a cell without a format of its own gives it one.
  const implied = parsed.numFmtId !== undefined ? parsed.numFmtId : parsed.numFmt !== undefined ? ensureNumberFormat(workbook.styles, parsed.numFmt) : undefined;
  if (implied !== undefined && /^general$/i.test(format)) {
    style = ensureCellFormat(workbook.styles, style, implied);
  }
  const cell: Cell = { kind: parsed.kind, value: parsed.value, style };
  if (parsed.formula !== undefined) {
    cell.formula = { text: parsed.formula, type: 'normal' };
    // A formula typed in is a dynamic array formula, as in Excel 365, where it means something else written as an
    // ordinary one (see needsDynamic(): by its text alone). One typed again as it was, an ordinary formula of the file,
    // stays one, with the implicit intersection it was written for.
    // UNSURE: Excel shows such a formula with "@" where it intersects, and keeps the "@" when it is typed again.
    const before = held && held.formula && held.formula.dynamic !== true && (held.formula.type === 'normal' || held.formula.type === 'shared') ? (held.formula.type === 'shared' ? explicitFormula(held, row, col, options.masters || sharedMasters(sheet)).formula : held.formula)
        : undefined;
    const legacy = !!before && before.text === parsed.formula;
    if (!legacy && needsDynamic(parsed.formula)) cell.formula.dynamic = true;
  }
  return cell;
}

/** A copy of a cell whose value is no longer what a spill put there. */
/** Dynamic array formulas of a copy whose spills are looked at when it is pasted (see pasteAt()); past it, their cells are pasted as values. */
const MAX_PASTED_ANCHORS = 4096;

function withoutSpill(cell: Cell): Cell {
  const out = { ...cell };
  delete out.spilled;
  return out;
}

function readable(formula: string): boolean {
  try {
    parseFormula(formula);
    return true;
  } catch (error) {
    return false;
  }
}

export function setCellInput(workbook: Workbook, sheetIndex: number, row: number, col: number, text: string, options: EditOptions = {}): Op[] {
  const sheet = workbook.sheets[sheetIndex];
  const cell = cellFromInput(workbook, sheet, row, col, text, options);
  if (cell === undefined && getCell(sheet, row, col) === undefined) return [];
  // A formula that cannot be read is refused, as Excel refuses it: stored, it would calculate to
  // nothing here and make a file that Excel has to repair.
  if (cell && cell.formula) {
    try {
      parseFormula(cell.formula.text);
    } catch (error) {
      throw new RangeError('The formula cannot be read');
    }
  }
  return [{ type: 'setCells', sheet: sheetIndex, cells: [{ row, col, cell }] }];
}

/** The text the cell editor starts with for a position. */
export function editTextAt(workbook: Workbook, sheet: Sheet, row: number, col: number, options: EditOptions = {}): string {
  const cell = getCell(sheet, row, col);
  if (!cell) return '';
  const format = resolveCellStyle(workbook.styles, cell.style).numFmt;
  // Text typed into a Text cell goes back in as it is; no apostrophe is needed to keep it text.
  if (isTextFormat(format) && cell.kind === 'string' && !cell.formula) return String(cell.value);
  // A cell that only points at its shared formula is edited as the formula it stands for.
  const shown = cell.formula && cell.formula.type === 'shared' && cell.formula.text === '' ? explicitFormula(cell, row, col, sharedMasters(sheet)) : cell;
  return cellEditText(shown, format, { locale: options.locale, date1904: workbook.date1904 });
}

/** Clips a range to the cells that exist, so that clearing whole rows or columns stays cheap. */
function usedPart(sheet: Sheet, range: CellRange): CellRange {
  return { top: range.top, left: range.left, bottom: Math.min(range.bottom, sheet.rowCount - 1), right: Math.min(range.right, sheet.colCount - 1) };
}

/**
 * Removes the contents of a range and keeps its formatting. Rows that the filter of the
 * sheet hides are left as they are, as in Excel; rows hidden by hand are emptied.
 */
export function clearRange(workbook: Workbook, sheetIndex: number, range: CellRange): Op[] {
  const sheet = workbook.sheets[sheetIndex];
  const used = usedPart(sheet, range);
  const filtered = rowsHiddenByFilter(sheet, used.top, used.bottom);
  const kept = filtered.length > 0 ? new Set(filtered) : undefined;
  const cells: CellChange[] = [];
  for (const [rowIndex, row] of sheet.rows) {
    if (rowIndex < used.top || rowIndex > used.bottom || (kept && kept.has(rowIndex))) continue;
    row.cells.forEach((cell, col) => {
      if (col < used.left || col > used.right) return;
      if (cell.kind === 'blank' && !cell.formula) return;
      cells.push({ row: rowIndex, col, cell: emptied(sheet, rowIndex, col, cell.style) });
    });
  }
  return cells.length > 0 ? [{ type: 'setCells', sheet: sheetIndex, cells }] : [];
}

/** A copied range: the text other programs get, and the cells themselves for pasting back here. */
export interface CopiedRange {
  text: string;
  html: string;
  sheet: number;
  /** The range that was selected. */
  range: CellRange;
  /** Rows of the range; a missing entry is an empty cell. */
  cells: (Cell | undefined)[][];
  /**
   * The sheet rows the rows of `cells` came from, when rows of the range were left out
   * (rows that a filter hides are not copied): the copy is one block, and a pasted formula
   * moves by the distance from the row it really came from.
   */
  rows?: number[];
  /**
   * The notes of the copied cells: `row` and `col` count from the first row of `cells` and its
   * first cell. Text and author only: a pasted note is a new one, with plain text and a box
   * where a new note gets it. Threaded comments are not copied. Missing when there is none.
   */
  notes?: CopiedNote[];
  /**
   * The data validations and conditional formats of the copied cells, as they applied when
   * the cells were copied (see copyRules()). Missing when the cells have none: a paste then
   * leaves the cells it writes without a rule, as it does every cell the copy has no rule for.
   */
  rules?: CopiedRules;
}

/** A note of a copied cell. */
export interface CopiedNote {
  row: number;
  col: number;
  text: string;
  author?: string;
}

/** Largest selection that a smaller copied block is repeated over. */
const MAX_TILE_CELLS = 200000;

/** Largest range that is copied cell by cell; whole rows and columns are cut down to the used part first. */
const MAX_COPY_CELLS = 2000000;

export function copyRange(workbook: Workbook, sheetIndex: number, range: CellRange, options: EditOptions = {}): CopiedRange {
  const sheet = workbook.sheets[sheetIndex];
  const used = usedPart(sheet, range);
  const bottom = Math.max(used.bottom, range.top);
  const right = Math.max(used.right, range.left);
  if ((bottom - range.top + 1) * (right - range.left + 1) > MAX_COPY_CELLS) throw new RangeError('The range is too large to copy');
  const cells: (Cell | undefined)[][] = [];
  const grid: string[][] = [];
  const formatCache = new Map<number, { code: string; builtinId: number | undefined }>();
  let masters: SharedMasters | undefined;
  // As in Excel, a copy out of a filtered list takes the rows that are shown.
  const filtered = rowsHiddenByFilter(sheet, range.top, bottom);
  const rows: number[] = [];
  for (const run of filtered.length > 0 ? rowRunsWithout(range.top, bottom, filtered) : [{ top: range.top, bottom }]) {
    for (let r = run.top; r <= run.bottom; r++) rows.push(r);
  }
  for (const r of rows) {
    const rowCells: (Cell | undefined)[] = [];
    const rowText: string[] = [];
    for (let c = range.left; c <= right; c++) {
      let cell = getCell(sheet, r, c);
      if (cell && cell.formula && cell.formula.type === 'shared') {
        // A copied formula has to carry its own text; a shared one may only point at its group.
        if (!masters) masters = sharedMasters(sheet);
        cell = explicitFormula(cell, r, c, masters);
      }
      rowCells.push(cell);
      if (!cell) {
        rowText.push('');
        continue;
      }
      let format = formatCache.get(cell.style);
      if (!format) {
        const resolved = resolveCellStyle(workbook.styles, cell.style);
        format = { code: resolved.numFmt, builtinId: builtinFormatId(workbook.styles, resolved.numFmtId) };
        formatCache.set(cell.style, format);
      }
      rowText.push(formatCell(cell, format.code, { locale: options.locale, date1904: workbook.date1904, builtinId: format.builtinId }).text);
    }
    cells.push(rowCells);
    grid.push(rowText);
  }
  const copied: CopiedRange = { text: toTsv(grid), html: toHtmlTable(grid), sheet: sheetIndex, range: { top: range.top, left: range.left, bottom, right }, cells };
  if (filtered.length > 0) copied.rows = rows;
  // The notes of the selection itself: a note may be on an empty cell beyond the cells that were copied.
  const notes = copiedNotes(sheet, range.top, range.left, bottom, Math.min(range.bottom, MAX_ROWS - 1), Math.min(range.right, MAX_COLS - 1), rows);
  if (notes.length > 0) copied.notes = notes;
  // The rules of the selection itself as well: cells that are still empty are often the ones with a rule.
  const last = Math.min(range.bottom, MAX_ROWS - 1);
  const rules = copyRules(workbook, sheetIndex, { top: range.top, left: range.left, bottom: Math.max(last, bottom), right: Math.max(right, Math.min(range.right, MAX_COLS - 1)) }, filtered.length > 0 ? rowsHiddenByFilter(sheet, range.top, Math.max(last, bottom)) : filtered);
  if (rules) copied.rules = rules;
  return copied;
}

/**
 * The notes on the cells of a selection, in the order of the sheet's comments; the first one
 * of a cell counts. `bottom` is the last row of the cells that were copied and `rows` the
 * rows that were copied down to it; the selection goes on to `last` and to the column `right`.
 */
function copiedNotes(sheet: Sheet, top: number, left: number, bottom: number, last: number, right: number, rows: number[]): CopiedNote[] {
  const out: CopiedNote[] = [];
  const comments = sheet.comments;
  if (!comments || comments.length === 0) return out;
  // Rows a filter hides are not in the copy: where each row of the sheet stands in it.
  const contiguous = rows.length === bottom - top + 1;
  let offsets: Map<number, number> | undefined;
  let hiddenBelow: Set<number> | undefined;
  const seen = new Set<number>();
  for (let i = 0; i < comments.length; i++) {
    const comment = comments[i];
    const cell = comment ? cellNow(sheet, comment) : undefined;
    if (!cell || cell.row < top || cell.row > Math.max(last, bottom) || cell.col < left || cell.col > right) continue;
    const key = cell.row * MAX_COLS + cell.col;
    if (seen.has(key)) continue;
    seen.add(key);
    // A conversation stays where it is: its messages and the people who wrote them are not a note.
    if (comment.threaded) continue;
    let row = cell.row - top;
    if (cell.row > bottom) {
      // Below the cells that were copied: the rows of the selection follow those of the copy one by one.
      if (!hiddenBelow) hiddenBelow = new Set(rowsHiddenByFilter(sheet, bottom + 1, last));
      if (hiddenBelow.has(cell.row)) continue;
      row = rows.length + (cell.row - bottom - 1);
    } else if (!contiguous) {
      if (!offsets) {
        offsets = new Map();
        for (let k = 0; k < rows.length; k++) offsets.set(rows[k], k);
      }
      const at = offsets.get(cell.row);
      if (at === undefined) continue;
      row = at;
    }
    const note: CopiedNote = { row, col: cell.col - left, text: String(comment.text) };
    if (typeof comment.author === 'string' && comment.author !== '') note.author = comment.author;
    out.push(note);
  }
  return out;
}

/**
 * Cuts a range: its contents are removed as clearRange() removes them, and the notes on its
 * cells go with them (a note is cut and pasted with its cell), and so do their data validation
 * and their conditional formats (see cutRules()). Formats stay. Like clearRange() it leaves the
 * rows alone that the filter of the sheet hides.
 */
export function cutRange(workbook: Workbook, sheetIndex: number, range: CellRange): Op[] {
  const sheet = workbook.sheets[sheetIndex];
  const ops = clearRange(workbook, sheetIndex, range);
  const used = usedPart(sheet, range);
  const filtered = rowsHiddenByFilter(sheet, used.top, used.bottom);
  let rows: number[] | undefined;
  if (filtered.length > 0) {
    rows = [];
    for (const run of rowRunsWithout(range.top, Math.max(used.bottom, range.top), filtered)) for (let r = run.top; r <= run.bottom; r++) rows.push(r);
  }
  // Notes may be on cells beyond the used part: the range itself is looked at, without the rows the filter hides.
  const hidden = filtered.length > 0 ? new Set(filtered) : undefined;
  const notes = hidden ? clearNotesExcept(workbook, sheetIndex, range, hidden) : clearNotes(workbook, sheetIndex, range);
  // The rules likewise: those of the range itself, without the rows the filter hides.
  const top = Math.max(0, range.top);
  const last = Math.min(range.bottom, MAX_ROWS - 1);
  const across = { left: Math.max(0, range.left), right: Math.min(range.right, MAX_COLS - 1) };
  const shown = filtered.length > 0 ? rowRunsWithout(top, last, rowsHiddenByFilter(sheet, top, last)) : [{ top, bottom: last }];
  const rules = top <= last && across.left <= across.right ? cutRules(workbook, sheetIndex, shown.map((run) => ({ top: run.top, bottom: run.bottom, ...across }))) : [];
  return ops.concat(notes, rules);
}

function clearNotesExcept(workbook: Workbook, sheetIndex: number, range: CellRange, hidden: Set<number>): Op[] {
  const sheet = workbook.sheets[sheetIndex];
  const rows: number[] = [];
  for (const comment of sheet.comments || []) {
    const cell = comment ? cellNow(sheet, comment) : undefined;
    if (cell && cell.row >= range.top && cell.row <= range.bottom && !hidden.has(cell.row)) rows.push(cell.row);
  }
  return clearNotes(workbook, sheetIndex, range, rows);
}

export interface PasteResult {
  ops: Op[];
  /** The range that was written, for the caller to select. */
  range: CellRange;
  /**
   * Notes of the copy that were not pasted: their cell has a threaded comment, or the notes
   * of the sheet cannot be changed. The cells were pasted all the same. Missing when none.
   */
  skippedNotes?: number;
  /**
   * How many rules (data validations, conditional formats) the paste could not do as they are
   * where the cells come from: rules of the copy that did not come, or came without what the
   * extension list of the file says of them, and rules that stayed on the cells though the
   * copy has none there (see pasteRules()). The cells were pasted all the same. Missing when none.
   */
  skippedRules?: number;
}

/**
 * Pastes at a selection. When `copied` is given and `text` is what it put on
 * the clipboard, the cells are pasted with their formats and their formulas
 * moved along; otherwise the text is parsed as typed input, field by field.
 * A selection that is a whole multiple of the copied block is filled with
 * repeats of it, as in Excel. Copied cells bring their data validation and
 * their conditional formats, which take the place of those of the cells that
 * are written (see pasteRules()); pasted text changes no rule.
 */
export function pasteAt(workbook: Workbook, sheetIndex: number, selection: CellRange, text: string, copied?: CopiedRange, options: EditOptions = {}): PasteResult {
  const sheet = workbook.sheets[sheetIndex];
  // Line breaks are compared loosely: the clipboard may hand back \r\n for \n.
  const internal = copied !== undefined && copied.text.replace(/\r\n?/g, '\n') === text.replace(/\r\n?/g, '\n');
  let source: (Cell | string | undefined)[][];
  if (internal && copied) source = copied.cells;
  else {
    const parsed = parseTsvLimited(text);
    // Pasting part of what was copied would look like success; it is refused instead.
    if (parsed.truncated) throw new RangeError(`Cannot paste more than ${MAX_PASTE_CELLS} cells at once`);
    source = parsed.rows;
  }
  const cellHeight = source.length;
  const cellWidth = source.reduce((max, row) => Math.max(max, row.length), 0);
  // An empty clipboard pastes nothing; it does not clear the target.
  if (cellHeight === 0 || cellWidth === 0 || (!internal && text === '')) return { ops: [], range: selection };
  // What is pasted is as large as what was selected: the notes of the copy may reach beyond its cells.
  let height = cellHeight;
  let width = cellWidth;
  if (internal && copied && copied.notes) {
    for (const note of copied.notes) {
      if (note.row >= height) height = note.row + 1;
      if (note.col >= width) width = note.col + 1;
    }
  }
  // And so may its rules: cells that are still empty are often the ones with a rule.
  if (internal && copied && copied.rules) {
    const extent = copiedRulesExtent(copied.rules);
    height = Math.max(height, extent.height);
    width = Math.max(width, extent.width);
  }

  const selHeight = selection.bottom - selection.top + 1;
  const selWidth = selection.right - selection.left + 1;
  // Repeating a block over whole rows or columns would write millions of cells; that selection gets one copy.
  const tiles = selHeight % height === 0 && selWidth % width === 0 && selHeight * selWidth <= MAX_TILE_CELLS;
  const repeatY = tiles ? selHeight / height : 1;
  const repeatX = tiles ? selWidth / width : 1;
  const bottom = Math.min(selection.top + height * repeatY - 1, MAX_ROWS - 1);
  const right = Math.min(selection.left + width * repeatX - 1, MAX_COLS - 1);

  // Cells hidden under a merged range are left alone: what is written there would never be seen.
  const merges = sheet.merges.filter((m) => m.bottom >= selection.top && m.top <= bottom && m.right >= selection.left && m.left <= right);
  const covered = (r: number, c: number) => merges.some((m) => r >= m.top && r <= m.bottom && c >= m.left && c <= m.right && (r !== m.top || c !== m.left));

  // The spills of the dynamic array formulas the copy has, where they were copied from: their cells are pasted
  // empty, and the pasted formula spills into them again. A spilled cell copied without its formula is a value.
  const anchors: CellRange[] = [];
  if (internal && copied) {
    copied.cells.forEach((row, offset) => {
      if (!row) return;
      const fromRow = copied.rows ? copied.rows[offset] : copied.range.top + offset;
      row.forEach((item, k) => {
        if (!item || !item.formula || item.formula.dynamic !== true || typeof item.formula.ref !== 'string') return;
        const area = parseRange(item.formula.ref);
        if (area && area.top === fromRow && area.left === copied.range.left + k && anchors.length < MAX_PASTED_ANCHORS) anchors.push(area);
      });
    });
  }
  const withAnchor = (row: number, col: number) => anchors.some((a) => row >= a.top && row <= a.bottom && col >= a.left && col <= a.right);

  const cells: CellChange[] = [];
  for (let r = selection.top; r <= bottom; r++) {
    const sourceRow = source[(r - selection.top) % height];
    // Beyond the cells of the copy (where only a note of it stands) no cell is written.
    if (!sourceRow) continue;
    for (let c = selection.left; c <= right; c++) {
      if ((c - selection.left) % width >= cellWidth) continue;
      if (merges.length > 0 && covered(r, c)) continue;
      const item = sourceRow[(c - selection.left) % width];
      let cell: Cell | undefined;
      // A row of pasted text that is shorter than the others leaves the cells beyond its end alone.
      if (item === undefined && !internal) continue;
      if (typeof item === 'string') {
        // Pasted text that looks like a formula and cannot be read as one stays the text it is.
        const asText = (style: number): Cell => ({ kind: 'string', value: item.length > MAX_CELL_TEXT ? item.slice(0, MAX_CELL_TEXT) : item, style });
        try {
          cell = cellFromInput(workbook, sheet, r, c, item, options);
          if (cell && cell.formula && !readable(cell.formula.text)) cell = asText(cell.style);
        } catch (error) {
          // Not written the way the locale writes formulas.
          cell = asText(styleAt(sheet, r, c));
        }
      }
      else if (item && copied) {
        const offset = (r - selection.top) % height;
        const fromRow = copied.rows ? copied.rows[offset] : copied.range.top + offset;
        const fromCol = copied.range.left + ((c - selection.left) % width);
        cell = item.formula
          ? {
              // A pasted formula stands alone: it is not part of the shared or array group it came from.
              // The result of the cell it was copied from is not its result, so it has none until it is calculated.
              kind: 'blank',
              value: null,
              style: item.style,
              formula: { text: shiftFormula(item.formula.text, r - fromRow, c - fromCol), type: 'normal', ...(item.formula.dynamic ? { dynamic: true } : {}) },
            }
          : item.spilled
            ? withAnchor(fromRow, fromCol)
              ? // Its formula is pasted as well, and fills it again.
                item.style
                ? { kind: 'blank', value: null, style: item.style }
                : undefined
              : // What a spill left there is a value of its own where it is pasted.
                withoutSpill(item)
            : item;
      }
      if (cell === undefined && getCell(sheet, r, c) === undefined) continue;
      cells.push({ row: r, col: c, cell });
    }
  }
  const result: PasteResult = { ops: cells.length > 0 ? [{ type: 'setCells', sheet: sheetIndex, cells }] : [], range: { top: selection.top, left: selection.left, bottom, right } };
  if (internal && copied && copied.notes && copied.notes.length > 0 && sheet.kind === 'worksheet') {
    const pasted = pasteNotes(sheet, sheetIndex, copied.notes, selection, height, width, bottom, right, merges.length > 0 ? covered : undefined);
    if (pasted.op) result.ops.push(pasted.op);
    if (pasted.skipped > 0) result.skippedNotes = pasted.skipped;
  }
  if (internal && copied) {
    const pasted = pasteRules(workbook, sheetIndex, copied.rules, result.range, height, width);
    for (const op of pasted.ops) result.ops.push(op);
    if (pasted.skipped > 0) result.skippedRules = pasted.skipped;
  }
  return result;
}

/**
 * The notes of a copy put on the cells it was pasted to, once for each repeat of the block.
 * A note of the copy takes the place of the note its cell has; a cell the copy has no note
 * for keeps its own (UNSURE: Excel's "paste all" removes that one as well; keeping it loses
 * nothing). A cell with a threaded comment keeps that and gets no note.
 */
function pasteNotes(
  sheet: Sheet,
  sheetIndex: number,
  notes: CopiedNote[],
  selection: CellRange,
  height: number,
  width: number,
  bottom: number,
  right: number,
  covered: ((row: number, col: number) => boolean) | undefined,
): { op: Op | undefined; skipped: number } {
  const targets: { row: number; col: number; note: CopiedNote }[] = [];
  for (let top = selection.top; top <= bottom; top += height) {
    for (let left = selection.left; left <= right; left += width) {
      for (const note of notes) {
        const row = top + note.row;
        const col = left + note.col;
        if (row > bottom || col > right) continue;
        targets.push({ row, col, note });
      }
    }
  }
  if (targets.length === 0) return { op: undefined, skipped: 0 };
  if (notesBlocked(sheet)) return { op: undefined, skipped: targets.length };

  const comments = sheet.comments || [];
  // What is on each cell now: the comments of the cell, and whether one of them is a conversation.
  const here = new Map<number, { indexes: number[]; threaded: boolean }>();
  let live = 0;
  for (let i = 0; i < comments.length; i++) {
    const cell = comments[i] ? cellNow(sheet, comments[i]) : undefined;
    if (!cell) continue;
    live++;
    const key = cell.row * MAX_COLS + cell.col;
    let entry = here.get(key);
    if (!entry) here.set(key, (entry = { indexes: [], threaded: false }));
    entry.indexes.push(i);
    if (comments[i].threaded) entry.threaded = true;
  }
  const since = sheet.structuralEdits ? sheet.structuralEdits.length : 0;
  const gone = new Set<number>();
  const added: SheetComment[] = [];
  let skipped = 0;
  for (const target of targets) {
    // A cell under a merged range is not written, and neither is its note.
    if (covered && covered(target.row, target.col)) continue;
    const entry = here.get(target.row * MAX_COLS + target.col);
    if (entry && entry.threaded) {
      skipped++;
      continue;
    }
    if (entry) for (const index of entry.indexes) gone.add(index);
    added.push({ row: target.row, col: target.col, text: target.note.text, author: target.note.author !== undefined ? target.note.author : DEFAULT_NOTE_AUTHOR, since });
  }
  if (added.length === 0) return { op: undefined, skipped };
  if (live - gone.size + added.length > MAX_SHEET_NOTES) throw new RangeError('The sheet has too many comments');
  const next = comments.filter((_, index) => !gone.has(index)).concat(added);
  return { op: { type: 'setComments', sheet: sheetIndex, comments: next }, skipped };
}
