// What the editor does, expressed as operations: typing into a cell, clearing
// a range, copying and pasting. These functions only build operations; the
// caller applies them through a History so that each is one undo step.

import { shiftFormula } from '../calc/refs';
import { MAX_PASTE_CELLS, parseTsvLimited, toHtmlTable, toTsv } from '../edit/clipboard';
import { cellEditText, parseCellInput } from '../edit/input';
import { builtinFormatId, formatCell } from '../format/cell-format';
import type { FormatLocale } from '../format/locale';
import { resolveCellStyle } from '../model/sheet/color';
import { MAX_COLS, MAX_ROWS, getCell, type Cell, type CellRange, type Sheet, type Workbook } from '../model/sheet/types';
import { ensureCellFormat, ensureNumberFormat, explicitFormula, sharedMasters, type CellChange, type Op, type SharedMasters } from './ops';

export interface EditOptions {
  locale?: FormatLocale;
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
function emptied(sheet: Sheet, row: number, col: number, style: number): Cell | undefined {
  return style === inheritedStyle(sheet, row, col) ? undefined : { kind: 'blank', value: null, style };
}

const isTextFormat = (code: string) => code === '@';

/** Longest text a cell holds. */
const MAX_CELL_TEXT = 32767;

/** The cell that typed text turns into at a position, or undefined when the position ends up empty. */
export function cellFromInput(workbook: Workbook, sheet: Sheet, row: number, col: number, text: string, options: EditOptions = {}): Cell | undefined {
  let style = styleAt(sheet, row, col);
  if (text === '') return emptied(sheet, row, col, style);
  const format = resolveCellStyle(workbook.styles, style).numFmt;
  // In a cell formatted as Text, what is typed is the text: "0123" keeps its zero and "=A1" is not a formula.
  if (isTextFormat(format)) return { kind: 'string', value: text.length > MAX_CELL_TEXT ? text.slice(0, MAX_CELL_TEXT) : text, style };
  const parsed = parseCellInput(text, { locale: options.locale, date1904: workbook.date1904 });
  if (parsed.kind === 'blank' && parsed.formula === undefined) return emptied(sheet, row, col, style);
  // Typing "15.03.2024" or "12%" into a cell without a format of its own gives it one.
  const implied = parsed.numFmtId !== undefined ? parsed.numFmtId : parsed.numFmt !== undefined ? ensureNumberFormat(workbook.styles, parsed.numFmt) : undefined;
  if (implied !== undefined && /^general$/i.test(format)) {
    style = ensureCellFormat(workbook.styles, style, implied);
  }
  const cell: Cell = { kind: parsed.kind, value: parsed.value, style };
  if (parsed.formula !== undefined) cell.formula = { text: parsed.formula, type: 'normal' };
  return cell;
}

export function setCellInput(workbook: Workbook, sheetIndex: number, row: number, col: number, text: string, options: EditOptions = {}): Op[] {
  const sheet = workbook.sheets[sheetIndex];
  const cell = cellFromInput(workbook, sheet, row, col, text, options);
  if (cell === undefined && getCell(sheet, row, col) === undefined) return [];
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

/** Removes the contents of a range and keeps its formatting. */
export function clearRange(workbook: Workbook, sheetIndex: number, range: CellRange): Op[] {
  const sheet = workbook.sheets[sheetIndex];
  const used = usedPart(sheet, range);
  const cells: CellChange[] = [];
  for (const [rowIndex, row] of sheet.rows) {
    if (rowIndex < used.top || rowIndex > used.bottom) continue;
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
  range: CellRange;
  /** Rows of the range; a missing entry is an empty cell. */
  cells: (Cell | undefined)[][];
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
  for (let r = range.top; r <= bottom; r++) {
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
  return { text: toTsv(grid), html: toHtmlTable(grid), sheet: sheetIndex, range: { top: range.top, left: range.left, bottom, right }, cells };
}

export interface PasteResult {
  ops: Op[];
  /** The range that was written, for the caller to select. */
  range: CellRange;
}

/**
 * Pastes at a selection. When `copied` is given and `text` is what it put on
 * the clipboard, the cells are pasted with their formats and their formulas
 * moved along; otherwise the text is parsed as typed input, field by field.
 * A selection that is a whole multiple of the copied block is filled with
 * repeats of it, as in Excel.
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
  const height = source.length;
  const width = source.reduce((max, row) => Math.max(max, row.length), 0);
  // An empty clipboard pastes nothing; it does not clear the target.
  if (height === 0 || width === 0 || (!internal && text === '')) return { ops: [], range: selection };

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

  const cells: CellChange[] = [];
  for (let r = selection.top; r <= bottom; r++) {
    const sourceRow = source[(r - selection.top) % height];
    for (let c = selection.left; c <= right; c++) {
      if (merges.length > 0 && covered(r, c)) continue;
      const item = sourceRow[(c - selection.left) % width];
      let cell: Cell | undefined;
      // A row of pasted text that is shorter than the others leaves the cells beyond its end alone.
      if (item === undefined && !internal) continue;
      if (typeof item === 'string') cell = cellFromInput(workbook, sheet, r, c, item, options);
      else if (item && copied) {
        const fromRow = copied.range.top + ((r - selection.top) % height);
        const fromCol = copied.range.left + ((c - selection.left) % width);
        cell = item.formula
          ? {
              // A pasted formula stands alone: it is not part of the shared or array group it came from.
              // The result of the cell it was copied from is not its result, so it has none until it is calculated.
              kind: 'blank',
              value: null,
              style: item.style,
              formula: { text: shiftFormula(item.formula.text, r - fromRow, c - fromCol), type: 'normal' },
            }
          : item;
      }
      if (cell === undefined && getCell(sheet, r, c) === undefined) continue;
      cells.push({ row: r, col: c, cell });
    }
  }
  return { ops: cells.length > 0 ? [{ type: 'setCells', sheet: sheetIndex, cells }] : [], range: { top: selection.top, left: selection.left, bottom, right } };
}
