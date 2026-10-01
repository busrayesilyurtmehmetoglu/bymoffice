// Spreadsheet model shared by every spreadsheet format (xlsx now, xls later).
//
// Rows and columns are 0-based everywhere in the model; "A1" is row 0, col 0.
// The model holds what the engine understands. Everything else stays in the
// source package and is written back from there (see the lossless principle).

import type { XmlElement } from '../../xml';

export const MAX_ROWS = 1048576;
export const MAX_COLS = 16384;

/** A colour as the file states it; resolveColor() turns it into "#RRGGBB". */
export interface Color {
  /** "AARRGGBB" or "RRGGBB", upper case. */
  rgb?: string;
  /** Index into the theme colour list. */
  theme?: number;
  /** -1..1, lightens (positive) or darkens (negative) the base colour. */
  tint?: number;
  /** Index into the legacy indexed palette. */
  indexed?: number;
  auto?: boolean;
}

export type Underline = 'none' | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';

export interface Font {
  name?: string;
  /** Size in points. */
  size?: number;
  bold: boolean;
  italic: boolean;
  strike: boolean;
  underline: Underline;
  color?: Color;
  vertAlign?: 'superscript' | 'subscript';
  family?: number;
  charset?: number;
  scheme?: 'major' | 'minor' | 'none';
}

export type Fill =
  | {
      type: 'pattern';
      /** "none", "solid", "gray125", ... as in ST_PatternType. */
      pattern: string;
      /** Foreground colour; this is the visible colour of a solid fill. */
      fg?: Color;
      bg?: Color;
    }
  | {
      type: 'gradient';
      /** The <gradientFill> element as read; not interpreted yet. */
      raw: XmlElement;
    };

export interface BorderEdge {
  /** "thin", "medium", "dashed", ... as in ST_BorderStyle. */
  style: string;
  color?: Color;
}

export interface Border {
  left?: BorderEdge;
  right?: BorderEdge;
  top?: BorderEdge;
  bottom?: BorderEdge;
  diagonal?: BorderEdge;
  diagonalUp: boolean;
  diagonalDown: boolean;
}

export interface Alignment {
  horizontal?: 'general' | 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
  vertical?: 'top' | 'center' | 'bottom' | 'justify' | 'distributed';
  wrapText?: boolean;
  shrinkToFit?: boolean;
  /** 0..180 degrees, or 255 for stacked vertical text. */
  textRotation?: number;
  indent?: number;
  readingOrder?: number;
}

export interface Protection {
  locked?: boolean;
  hidden?: boolean;
}

/** One <xf> record: indexes into the other style tables. */
export interface CellFormat {
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
  /** Index into cellStyleXfs; only present on cellXfs entries. */
  xfId?: number;
  alignment?: Alignment;
  protection?: Protection;
  quotePrefix: boolean;
}

export interface NamedCellStyle {
  name: string;
  xfId: number;
  builtinId?: number;
}

export interface StyleSheet {
  /** Custom number formats by id; built-in ids are answered by formatCodeOf(). */
  numFmts: Map<number, string>;
  fonts: Font[];
  fills: Fill[];
  borders: Border[];
  cellStyleXfs: CellFormat[];
  /** Cell.style is an index into this list. */
  cellXfs: CellFormat[];
  cellStyles: NamedCellStyle[];
  /** Differential formats (conditional formatting, tables); kept as read. */
  dxfs: XmlElement[];
  /** Custom indexed palette ("RRGGBB") when the file overrides the default one. */
  indexedColors?: string[];
  /** Theme colours as "RRGGBB", in the order theme indexes refer to them. */
  themeColors: string[];
  /** Set once a format was added after the styles were read. */
  modified?: boolean;
}

export interface RichRun {
  text: string;
  font?: Font;
}

export interface SharedString {
  /** Plain text of the string (all runs joined, phonetic text excluded). */
  text: string;
  /** Present only for rich text. */
  runs?: RichRun[];
}

export interface CellFormula {
  /** Formula text without the leading "=". Empty for a shared formula that only refers to its master. */
  text: string;
  type: 'normal' | 'array' | 'shared' | 'dataTable';
  /** Range the array or shared formula applies to, e.g. "A1:A10". */
  ref?: string;
  /** Shared formula group index. */
  si?: number;
  /** Attributes of <f> the model has no field for, kept to be written back. */
  attrs?: Map<string, string>;
}

export type CellKind = 'blank' | 'number' | 'string' | 'boolean' | 'error' | 'date';

export interface Cell {
  kind: CellKind;
  /**
   * number -> number, boolean -> boolean, string -> string, error -> the error
   * text ("#DIV/0!"), date -> ISO 8601 text (cells of type "d"), blank -> null.
   * Dates are normally plain numbers with a date number format.
   */
  value: number | string | boolean | null;
  /** Index into StyleSheet.cellXfs. */
  style: number;
  formula?: CellFormula;
  /** Rich text runs, when the text has mixed formatting. Shared between cells that use the same string. */
  runs?: RichRun[];
}

export interface Row {
  /** Height in points when the row states one. */
  height?: number;
  customHeight: boolean;
  hidden: boolean;
  /** Row-level format (cellXfs index) applied to cells without their own. */
  style?: number;
  outlineLevel: number;
  collapsed: boolean;
  /** Sparse, indexed by column. */
  cells: Cell[];
}

/** Properties shared by a run of columns, as in <col min max>. */
export interface ColumnRange {
  /** First and last column, inclusive. */
  min: number;
  max: number;
  /** Width in characters of the default font's widest digit. */
  width?: number;
  customWidth: boolean;
  hidden: boolean;
  style?: number;
  outlineLevel: number;
  collapsed: boolean;
}

/** Inclusive on all four sides. */
export interface CellRange {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

export interface Hyperlink {
  range: CellRange;
  /** Relationship id of an external target; `target` is filled in from the relationships. */
  rId?: string;
  target?: string;
  /** Location inside the workbook, e.g. "Sheet2!A1". */
  location?: string;
  display?: string;
  tooltip?: string;
}

export interface SheetView {
  showGridLines: boolean;
  showRowColHeaders: boolean;
  showZeros: boolean;
  rightToLeft: boolean;
  tabSelected: boolean;
  /** Percent, 100 = normal. */
  zoomScale: number;
  /** Rows and columns kept in place by frozen panes; 0 when not frozen. */
  frozenRows: number;
  frozenCols: number;
  /** First cell of the scrollable area of frozen panes, e.g. "B2". */
  topLeftCell?: string;
  /**
   * Top-left cell of the view as a whole. With frozen panes this is where the
   * frozen rows and columns start: panes frozen while scrolled to C5 keep row 5
   * and column C in place, not row 1 and column A.
   */
  firstVisibleCell?: string;
  activeCell?: string;
}

export type SheetState = 'visible' | 'hidden' | 'veryHidden';
export type SheetKind = 'worksheet' | 'chartsheet' | 'dialogsheet' | 'macrosheet';

export interface Sheet {
  name: string;
  sheetId: number;
  state: SheetState;
  kind: SheetKind;
  /** Part this sheet was read from, e.g. "/xl/worksheets/sheet1.xml". */
  part: string;
  /** Sparse, keyed by row index. */
  rows: Map<number, Row>;
  /** Sorted by `min`, not overlapping. */
  columns: ColumnRange[];
  merges: CellRange[];
  hyperlinks: Hyperlink[];
  /** Used range as the file declares it; may be missing or wrong. */
  dimension?: CellRange;
  /** One past the last row and column that hold a cell (0 for an empty sheet). */
  rowCount: number;
  colCount: number;
  /** Default row height in points. */
  defaultRowHeight?: number;
  /** Default column width in characters. */
  defaultColWidth?: number;
  baseColWidth?: number;
  view: SheetView;
  tabColor?: Color;
  /** Set once anything in the sheet was changed after it was read; the writer regenerates only such sheets. */
  modified?: boolean;
  /**
   * Rows and columns inserted or deleted since the sheet was read, in order.
   * The writer replays them on the parts of the sheet the model does not hold
   * (conditional formats, validations, ...), so that their ranges follow.
   */
  structuralEdits?: SheetStructuralEdit[];
}

export interface SheetStructuralEdit {
  kind: 'insert' | 'delete';
  axis: 'row' | 'col';
  at: number;
  count: number;
}

export interface DefinedName {
  name: string;
  /** The formula or reference text. */
  value: string;
  /** Index of the sheet the name is local to; undefined for workbook scope. */
  localSheet?: number;
  hidden: boolean;
}

export interface Workbook {
  sheets: Sheet[];
  styles: StyleSheet;
  sharedStrings: SharedString[];
  definedNames: DefinedName[];
  /** Dates count from 1904-01-01 instead of 1899-12-30. */
  date1904: boolean;
  /** Index of the sheet shown when the workbook opens. */
  activeSheet: number;
}

export function createSheet(init: Pick<Sheet, 'name' | 'sheetId' | 'state' | 'kind' | 'part'>): Sheet {
  return {
    ...init,
    rows: new Map(),
    columns: [],
    merges: [],
    hyperlinks: [],
    rowCount: 0,
    colCount: 0,
    view: {
      showGridLines: true,
      showRowColHeaders: true,
      showZeros: true,
      rightToLeft: false,
      tabSelected: false,
      zoomScale: 100,
      frozenRows: 0,
      frozenCols: 0,
    },
  };
}

export function emptyStyleSheet(): StyleSheet {
  const font: Font = { name: 'Calibri', size: 11, bold: false, italic: false, strike: false, underline: 'none' };
  const border: Border = { diagonalUp: false, diagonalDown: false };
  const format: CellFormat = { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, quotePrefix: false };
  return {
    numFmts: new Map(),
    fonts: [font],
    fills: [{ type: 'pattern', pattern: 'none' }],
    borders: [border],
    cellStyleXfs: [{ ...format }],
    cellXfs: [{ ...format, xfId: 0 }],
    cellStyles: [],
    dxfs: [],
    themeColors: [],
  };
}

/** Column index -> letters: 0 -> "A", 25 -> "Z", 26 -> "AA". */
export function columnName(col: number): string {
  let name = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
}

export function cellName(row: number, col: number): string {
  return columnName(col) + (row + 1);
}

/**
 * Parses "B7" (also "$B$7") into 0-based indexes. Returns undefined for
 * anything that is not a cell reference inside the sheet limits.
 */
export function parseCellRef(ref: string): { row: number; col: number } | undefined {
  let i = 0;
  const n = ref.length;
  if (ref.charCodeAt(i) === 36) i++;
  let col = 0;
  const lettersStart = i;
  for (; i < n; i++) {
    const c = ref.charCodeAt(i) & ~32;
    if (c < 65 || c > 90) break;
    col = col * 26 + (c - 64);
    if (col > MAX_COLS) return undefined;
  }
  if (i === lettersStart) return undefined;
  if (ref.charCodeAt(i) === 36) i++;
  let row = 0;
  const digitsStart = i;
  for (; i < n; i++) {
    const c = ref.charCodeAt(i);
    if (c < 48 || c > 57) return undefined;
    row = row * 10 + (c - 48);
    if (row > MAX_ROWS) return undefined;
  }
  if (i === digitsStart || row === 0) return undefined;
  return { row: row - 1, col: col - 1 };
}

/** Parses "A1:C3" or a single "A1"; corners may come in any order. */
export function parseRange(ref: string): CellRange | undefined {
  const colon = ref.indexOf(':');
  const first = parseCellRef(colon < 0 ? ref : ref.slice(0, colon));
  const second = colon < 0 ? first : parseCellRef(ref.slice(colon + 1));
  if (!first || !second) return undefined;
  return {
    top: Math.min(first.row, second.row),
    left: Math.min(first.col, second.col),
    bottom: Math.max(first.row, second.row),
    right: Math.max(first.col, second.col),
  };
}

export function rangeName(range: CellRange): string {
  const first = cellName(range.top, range.left);
  const last = cellName(range.bottom, range.right);
  return first === last ? first : `${first}:${last}`;
}

export function getCell(sheet: Sheet, row: number, col: number): Cell | undefined {
  return sheet.rows.get(row)?.cells[col];
}
