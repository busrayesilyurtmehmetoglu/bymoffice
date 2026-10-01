// Formatting commands: fonts, fills, borders, alignment, number formats,
// column widths, row heights and merged cells. Like the other commands they
// only build operations. The style tables grow by appending: an existing
// record is never changed, since other cells share it.

import { resolveCellStyle } from '../model/sheet/color';
import {
  MAX_COLS,
  MAX_ROWS,
  getCell,
  type Alignment,
  type Border,
  type BorderEdge,
  type Cell,
  type CellFormat,
  type CellRange,
  type Color,
  type ColumnRange,
  type Fill,
  type Font,
  type Sheet,
  type StyleSheet,
  type Underline,
  type Workbook,
} from '../model/sheet/types';
import { emptied, styleAt } from './commands';
import { MAX_CELL_FORMATS, ensureNumberFormat, sameFormat, type CellChange, type Op, type RowProps } from './ops';

/** Which edges of a range get a border; "all" includes the lines between its cells. */
export type BorderPreset = 'all' | 'outline' | 'none' | 'top' | 'bottom' | 'left' | 'right';

/** What to change in a format. A property that is missing stays as it is; null removes it. */
export interface StylePatch {
  bold?: boolean;
  italic?: boolean;
  strike?: boolean;
  underline?: Underline;
  fontName?: string;
  /** Points. */
  fontSize?: number;
  /** Text colour; null for the automatic colour. */
  color?: Color | null;
  /** Solid fill colour; null for no fill. */
  fill?: Color | null;
  horizontal?: Alignment['horizontal'] | null;
  vertical?: Alignment['vertical'] | null;
  wrapText?: boolean;
  /** Number format code, e.g. "0.00%". */
  numFmt?: string;
  /** Borders by position within the formatted range. */
  border?: { preset: BorderPreset; edge?: BorderEdge };
}

/** Largest number of cells one formatting command touches one by one. */
const MAX_FORMAT_CELLS = 200000;

function sameFlat(a: object | undefined, b: object | undefined): boolean {
  const left = (a || {}) as Record<string, unknown>;
  const right = (b || {}) as Record<string, unknown>;
  for (const key of Object.keys(left)) if (!sameValue(left[key], right[key])) return false;
  for (const key of Object.keys(right)) if (!sameValue(left[key], right[key])) return false;
  return true;
}

/** Equality for style records: plain values, and one level of nested records (colours, edges). */
function sameValue(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (typeof a === 'object' && typeof b === 'object' && a !== null && b !== null) return sameFlat(a, b);
  return false;
}

function findOrAdd<T extends object>(list: T[], wanted: T, styles: StyleSheet): number {
  const found = list.findIndex((item) => sameFlat(item, wanted));
  if (found >= 0) return found;
  styles.modified = true;
  return list.push(wanted) - 1;
}

/** Where a cell sits within the range being formatted; borders depend on it. */
interface Position {
  top: boolean;
  bottom: boolean;
  left: boolean;
  right: boolean;
}

const INSIDE: Position = { top: false, bottom: false, left: false, right: false };

function patchBorder(border: Border, preset: BorderPreset, edge: BorderEdge | undefined, at: Position): Border {
  const line = edge || { style: 'thin' };
  const next: Border = { ...border };
  const set = (side: 'top' | 'bottom' | 'left' | 'right', on: boolean) => {
    if (on) next[side] = line;
  };
  switch (preset) {
    case 'none':
      delete next.top;
      delete next.bottom;
      delete next.left;
      delete next.right;
      break;
    case 'all':
      next.top = next.bottom = next.left = next.right = line;
      break;
    case 'outline':
      set('top', at.top);
      set('bottom', at.bottom);
      set('left', at.left);
      set('right', at.right);
      break;
    default:
      set(preset, at[preset]);
  }
  return next;
}

/** Index of the cell format that is `base` with a patch applied, adding style records as needed. */
export function ensureStyle(styles: StyleSheet, base: number, patch: StylePatch, at: Position = INSIDE): number {
  const baseIndex = styles.cellXfs[base] ? base : 0;
  const source = styles.cellXfs[baseIndex];
  if (!source) throw new Error('The style sheet has no cell formats');
  const wanted: CellFormat = { ...source };

  const fontPatch =
    patch.bold !== undefined ||
    patch.italic !== undefined ||
    patch.strike !== undefined ||
    patch.underline !== undefined ||
    patch.fontName !== undefined ||
    patch.fontSize !== undefined ||
    patch.color !== undefined;
  if (fontPatch) {
    const font: Font = { ...(styles.fonts[source.fontId] || styles.fonts[0]) };
    if (patch.bold !== undefined) font.bold = patch.bold;
    if (patch.italic !== undefined) font.italic = patch.italic;
    if (patch.strike !== undefined) font.strike = patch.strike;
    if (patch.underline !== undefined) font.underline = patch.underline;
    if (patch.fontName !== undefined) {
      font.name = patch.fontName;
      // A named font is no longer the theme's major or minor font.
      delete font.scheme;
    }
    if (patch.fontSize !== undefined && patch.fontSize > 0) font.size = Math.min(patch.fontSize, 409);
    if (patch.color === null) delete font.color;
    else if (patch.color !== undefined) font.color = patch.color;
    wanted.fontId = findOrAdd(styles.fonts, font, styles);
  }

  if (patch.fill !== undefined) {
    const fill: Fill = patch.fill === null ? { type: 'pattern', pattern: 'none' } : { type: 'pattern', pattern: 'solid', fg: patch.fill, bg: { indexed: 64 } };
    // Fills 0 and 1 are reserved ("none" and "gray125"): Excel reads them that way whatever the file
    // says, so a style sheet without them gets them before a fill of its own is added.
    if (styles.fills.length === 0) styles.fills.push({ type: 'pattern', pattern: 'none' });
    if (styles.fills.length === 1) styles.fills.push({ type: 'pattern', pattern: 'gray125' });
    wanted.fillId = findOrAdd(styles.fills, fill, styles);
  }

  if (patch.border) {
    const border = patchBorder(styles.borders[source.borderId] || styles.borders[0], patch.border.preset, patch.border.edge, at);
    wanted.borderId = findOrAdd(styles.borders, border, styles);
  }

  if (patch.horizontal !== undefined || patch.vertical !== undefined || patch.wrapText !== undefined) {
    const alignment: Alignment = { ...(source.alignment || {}) };
    if (patch.horizontal === null) delete alignment.horizontal;
    else if (patch.horizontal !== undefined) alignment.horizontal = patch.horizontal;
    if (patch.vertical === null) delete alignment.vertical;
    else if (patch.vertical !== undefined) alignment.vertical = patch.vertical;
    if (patch.wrapText === false) delete alignment.wrapText;
    else if (patch.wrapText) alignment.wrapText = true;
    if (Object.keys(alignment).length > 0) wanted.alignment = alignment;
    else delete wanted.alignment;
  }

  if (patch.numFmt !== undefined) wanted.numFmtId = ensureNumberFormat(styles, patch.numFmt);

  if (sameFormat(source, wanted)) return baseIndex;
  const found = styles.cellXfs.findIndex((format) => sameFormat(format, wanted));
  if (found >= 0) return found;
  if (styles.cellXfs.length >= MAX_CELL_FORMATS) throw new RangeError('The workbook has too many cell formats');
  styles.modified = true;
  return styles.cellXfs.push(wanted) - 1;
}

const fullHeight = (range: CellRange) => range.top <= 0 && range.bottom >= MAX_ROWS - 1;
const fullWidth = (range: CellRange) => range.left <= 0 && range.right >= MAX_COLS - 1;

/** Column settings with one more property set on [first, last]; ranges are split where they cross the span. */
function patchColumns(columns: ColumnRange[], first: number, last: number, change: (range: ColumnRange) => ColumnRange): ColumnRange[] {
  const blank = (min: number, max: number): ColumnRange => ({ min, max, customWidth: false, hidden: false, outlineLevel: 0, collapsed: false });
  const out: ColumnRange[] = [];
  let next = first;
  for (const range of columns) {
    if (range.max < first || range.min > last) {
      out.push(range);
      continue;
    }
    // Columns of the span before this range had no settings yet.
    if (range.min > next) out.push(change(blank(next, range.min - 1)));
    if (range.min < first) out.push({ ...range, max: first - 1 });
    out.push(change({ ...range, min: Math.max(range.min, first), max: Math.min(range.max, last) }));
    if (range.max > last) out.push({ ...range, min: last + 1 });
    next = Math.max(next, range.max + 1);
  }
  if (next <= last) out.push(change(blank(next, last)));
  return out.sort((a, b) => a.min - b.min);
}

/** Rows with a format of their own get cells where a formatted column crosses them, up to this many columns. */
const MAX_CROSSING_COLUMNS = 256;

/** Column settings in which the first and the last column of a span have records of their own. */
function splitColumns(columns: ColumnRange[], first: number, last: number): ColumnRange[] {
  let out = patchColumns(columns, first, first, (column) => column);
  if (last > first) out = patchColumns(out, last, last, (column) => column);
  return out;
}

/** Applies a style patch to a range. Whole rows and columns are formatted as such, not cell by cell. */
export function formatRange(workbook: Workbook, sheetIndex: number, range: CellRange, patch: StylePatch): Op[] {
  const sheet = workbook.sheets[sheetIndex];
  const styles = workbook.styles;
  const cache = new Map<number, number>();
  const styled = (base: number, row: number, col: number): number => {
    let at = INSIDE;
    let key = base;
    if (patch.border) {
      at = { top: row === range.top, bottom: row === range.bottom, left: col === range.left, right: col === range.right };
      key = base * 16 + (at.top ? 1 : 0) + (at.bottom ? 2 : 0) + (at.left ? 4 : 0) + (at.right ? 8 : 0);
    }
    let index = cache.get(key);
    if (index === undefined) {
      index = ensureStyle(styles, base, patch, at);
      cache.set(key, index);
    }
    return index;
  };

  const ops: Op[] = [];
  const cells: CellChange[] = [];
  const restyle = (row: number, col: number, cell: Cell | undefined, inherited: number) => {
    const style = styled(cell ? cell.style : inherited, row, col);
    if (cell) {
      if (cell.style !== style) cells.push({ row, col, cell: { ...cell, style } });
    } else if (style !== inherited) cells.push({ row, col, cell: { kind: 'blank', value: null, style } });
  };

  if (fullHeight(range)) {
    const crossedRows: { row: number; props: RowProps }[] = [];
    // Whole columns: the column format changes, and every cell already in those columns.
    const last = Math.min(range.right, MAX_COLS - 1);
    // With borders the first and the last column differ from those between them, so they get records of their own.
    const split = patch.border ? splitColumns(sheet.columns, range.left, last) : sheet.columns;
    const columns = patchColumns(split, range.left, last, (column) => ({
      ...column,
      style: styled(column.style === undefined ? 0 : column.style, -1, column.min === range.left ? range.left : column.max === last ? last : column.min),
    }));
    ops.push({ type: 'setColumns', sheet: sheetIndex, columns });
    for (const [rowIndex, row] of sheet.rows) {
      row.cells.forEach((cell, col) => {
        if (col >= range.left && col <= range.right) restyle(rowIndex, col, cell, cell.style);
      });
      // A row with a format of its own hides the column's; its crossing cells are formatted directly.
      // Across more columns than that the row's own format is changed instead, as for a whole row.
      if (row.style !== undefined) {
        if (last - range.left < MAX_CROSSING_COLUMNS) {
          for (let col = range.left; col <= last; col++) {
            if (!row.cells[col]) restyle(rowIndex, col, undefined, row.style);
          }
        } else crossedRows.push({ row: rowIndex, props: { style: styled(row.style, rowIndex, -1) } });
      }
    }
    if (crossedRows.length > 0) ops.push({ type: 'setRows', sheet: sheetIndex, rows: crossedRows });
  } else if (fullWidth(range)) {
    const count = range.bottom - range.top + 1;
    if (count > MAX_FORMAT_CELLS) throw new RangeError('Too many rows to format at once');
    const rows: { row: number; props: RowProps }[] = [];
    for (let r = range.top; r <= range.bottom; r++) {
      const row = sheet.rows.get(r);
      rows.push({ row: r, props: { style: styled(row && row.style !== undefined ? row.style : 0, r, -1) } });
      if (row) row.cells.forEach((cell, col) => restyle(r, col, cell, cell.style));
    }
    ops.push({ type: 'setRows', sheet: sheetIndex, rows });
  } else {
    const count = (range.bottom - range.top + 1) * (range.right - range.left + 1);
    if (count > MAX_FORMAT_CELLS) throw new RangeError('Too many cells to format at once');
    for (let r = range.top; r <= range.bottom; r++) {
      for (let c = range.left; c <= range.right; c++) restyle(r, c, getCell(sheet, r, c), styleAt(sheet, r, c));
    }
  }
  if (cells.length > 0) ops.push({ type: 'setCells', sheet: sheetIndex, cells });
  return ops;
}

/** Sets the width of columns in characters; undefined goes back to the default width, 0 hides them. */
export function setColumnWidth(workbook: Workbook, sheetIndex: number, first: number, last: number, width: number | undefined): Op[] {
  const sheet = workbook.sheets[sheetIndex];
  const columns = patchColumns(sheet.columns, first, last, (column) => {
    // A width of nothing is a hidden column: it keeps the width it had, to come back with.
    if (width !== undefined && width <= 0) return { ...column, hidden: true };
    const next: ColumnRange = { ...column, customWidth: width !== undefined, hidden: false };
    if (width === undefined) delete next.width;
    else next.width = Math.min(Math.max(width, 0), 255);
    return next;
  });
  return [{ type: 'setColumns', sheet: sheetIndex, columns }];
}

export function setColumnsHidden(workbook: Workbook, sheetIndex: number, first: number, last: number, hidden: boolean): Op[] {
  const sheet = workbook.sheets[sheetIndex];
  return [{ type: 'setColumns', sheet: sheetIndex, columns: patchColumns(sheet.columns, first, last, (column) => ({ ...column, hidden })) }];
}

/** Sets the height of rows in points; undefined goes back to the default height. */
export function setRowHeight(sheetIndex: number, first: number, last: number, height: number | undefined): Op[] {
  if (last - first + 1 > MAX_FORMAT_CELLS) throw new RangeError('Too many rows to resize at once');
  const rows: { row: number; props: RowProps }[] = [];
  // A height of nothing hides the rows; they keep the height they had.
  const props: RowProps =
    height === undefined ? { height: undefined, customHeight: false } : height <= 0 ? { hidden: true } : { height: Math.min(height, 409.5), customHeight: true, hidden: false };
  for (let r = first; r <= last; r++) rows.push({ row: r, props });
  return [{ type: 'setRows', sheet: sheetIndex, rows }];
}

export function setRowsHidden(sheetIndex: number, first: number, last: number, hidden: boolean): Op[] {
  if (last - first + 1 > MAX_FORMAT_CELLS) throw new RangeError('Too many rows to change at once');
  const rows: { row: number; props: RowProps }[] = [];
  for (let r = first; r <= last; r++) rows.push({ row: r, props: { hidden } });
  return [{ type: 'setRows', sheet: sheetIndex, rows }];
}

const intersects = (a: CellRange, b: CellRange) => a.top <= b.bottom && a.bottom >= b.top && a.left <= b.right && a.right >= b.left;

/**
 * Merges a range into one cell. As in Excel, only the content of its top-left
 * cell is kept; merges the range overlaps are replaced by it.
 */
export function mergeRange(workbook: Workbook, sheetIndex: number, range: CellRange): Op[] {
  const sheet = workbook.sheets[sheetIndex];
  if (range.top === range.bottom && range.left === range.right) return [];
  if (fullHeight(range) || fullWidth(range)) throw new RangeError('Whole rows or columns cannot be merged');
  const ops: Op[] = [];
  const cells: CellChange[] = [];
  for (const [rowIndex, row] of sheet.rows) {
    if (rowIndex < range.top || rowIndex > range.bottom) continue;
    row.cells.forEach((cell, col) => {
      if (col < range.left || col > range.right || (rowIndex === range.top && col === range.left)) return;
      if (cell.kind === 'blank' && !cell.formula) return;
      cells.push({ row: rowIndex, col, cell: emptied(sheet, rowIndex, col, cell.style) });
    });
  }
  if (cells.length > 0) ops.push({ type: 'setCells', sheet: sheetIndex, cells });
  ops.push({ type: 'setMerges', sheet: sheetIndex, merges: sheet.merges.filter((merge) => !intersects(merge, range)).concat([{ ...range }]) });
  return ops;
}

/** Splits every merged range that a range touches. */
export function unmergeRange(workbook: Workbook, sheetIndex: number, range: CellRange): Op[] {
  const sheet = workbook.sheets[sheetIndex];
  const merges = sheet.merges.filter((merge) => !intersects(merge, range));
  return merges.length === sheet.merges.length ? [] : [{ type: 'setMerges', sheet: sheetIndex, merges }];
}

/** The format a position shows with, for a toolbar to reflect. */
export function formatAt(workbook: Workbook, sheet: Sheet, row: number, col: number): ReturnType<typeof resolveCellStyle> {
  return resolveCellStyle(workbook.styles, styleAt(sheet, row, col));
}
