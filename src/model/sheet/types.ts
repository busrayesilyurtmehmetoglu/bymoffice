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
  /** The same records as far as the engine draws with them; index for index with `dxfs`. */
  dxfFormats?: DifferentialFormat[];
  /** Custom indexed palette ("RRGGBB") when the file overrides the default one. */
  indexedColors?: string[];
  /** Theme colours as "RRGGBB", in the order theme indexes refer to them. */
  themeColors: string[];
  /** Set once a format was added after the styles were read. */
  modified?: boolean;
}

/**
 * A differential format (<dxf>): only what it states is applied on top of the
 * format of a cell. A view for reading; the record itself stays in the package.
 */
export interface DifferentialFormat {
  font?: {
    bold?: boolean;
    italic?: boolean;
    strike?: boolean;
    underline?: Underline;
    color?: Color;
  };
  /**
   * As the file states it, with one difference from a cell fill: a <patternFill>
   * without a patternType is a solid fill, and the colour of a solid fill is `bg`
   * (see differentialFillColor()).
   */
  fill?: Fill;
  border?: {
    left?: BorderEdge;
    right?: BorderEdge;
    top?: BorderEdge;
    bottom?: BorderEdge;
  };
  numFmt?: { id: number; code: string };
}

/** A threshold of a colour scale, a data bar or an icon set (<cfvo>). */
export interface ConditionalValue {
  /** "min", "max", "num", "percent", "percentile", "formula"; "autoMin" and "autoMax" in the 2010 extension. */
  type: string;
  /** A number or a formula (without "="), as written. */
  value?: string;
  /** Icon sets: the threshold itself belongs to the icon above it. Default true. */
  gte: boolean;
}

export interface ConditionalColorScale {
  values: ConditionalValue[];
  colors: Color[];
}

export interface ConditionalDataBar {
  min: ConditionalValue;
  max: ConditionalValue;
  color?: Color;
  showValue: boolean;
  /** Shortest and longest bar, in percent of the cell width. */
  minLength: number;
  maxLength: number;
  /** The rest comes from the 2010 extension (<x14:dataBar>); `extended` says that one was found. */
  extended: boolean;
  gradient: boolean;
  borderColor?: Color;
  negativeFillColor?: Color;
  negativeBorderColor?: Color;
  negativeFillSameAsPositive: boolean;
  negativeBorderSameAsPositive: boolean;
  axisColor?: Color;
  axisPosition: 'automatic' | 'middle' | 'none';
  direction: 'context' | 'leftToRight' | 'rightToLeft';
}

export interface ConditionalIconSet {
  /** "3Arrows", "4Rating", "5Quarters", ...; "3TrafficLights1" when the file names none. */
  set: string;
  values: ConditionalValue[];
  showValue: boolean;
  reverse: boolean;
  /** Icons chosen one by one (2010 extension), lowest first; a set of "NoIcons" shows nothing. */
  custom?: { set: string; id: number }[];
}

/** One <cfRule>. Attributes that only some rule types use are present when the file states them. */
export interface ConditionalRule {
  /** "cellIs", "expression", "colorScale", "dataBar", "iconSet", "top10", ... as in ST_CfType; unknown types are kept and never match. */
  type: string;
  /** Lower numbers are applied first; Infinity when the file states none. */
  priority: number;
  stopIfTrue: boolean;
  /** Index into StyleSheet.dxfFormats. */
  dxfId?: number;
  /** A format stated in the rule itself (rules of the 2010 extension). */
  dxf?: DifferentialFormat;
  /** cellIs: "between", "notBetween", "equal", "notEqual", "greaterThan", "lessThan", "greaterThanOrEqual", "lessThanOrEqual". */
  operator?: string;
  /** Formula texts without "=", written for the top-left cell of the first range. */
  formulas: string[];
  /** containsText, notContainsText, beginsWith, endsWith. */
  text?: string;
  timePeriod?: string;
  /** top10. */
  rank?: number;
  percent: boolean;
  bottom: boolean;
  /** aboveAverage: false means "below average". */
  aboveAverage: boolean;
  equalAverage: boolean;
  stdDev?: number;
  colorScale?: ConditionalColorScale;
  dataBar?: ConditionalDataBar;
  iconSet?: ConditionalIconSet;
  /** Id that ties a rule to its counterpart in the 2010 extension list. */
  extId?: string;
  /** The rule exists only in the 2010 extension list. */
  extended?: boolean;
}

/** One <conditionalFormatting>: rules and the cells they apply to. A view for reading; the writer copies the source element. */
export interface ConditionalFormat {
  ranges: CellRange[];
  rules: ConditionalRule[];
}

/** One condition of a custom filter ("greater than 5", "begins with a"). */
export interface FilterCondition {
  /** "equal", "notEqual", "greaterThan", "greaterThanOrEqual", "lessThan", "lessThanOrEqual". */
  operator: string;
  /** As written; "*" and "?" are wildcards of "equal" and "notEqual". */
  value: string;
}

/** A date or a part of one chosen in a filter list (<dateGroupItem>): everything up to `grouping` has to match. */
export interface FilterDateGroup {
  /** "year", "month", "day", "hour", "minute", "second". */
  grouping: string;
  year: number;
  month?: number;
  day?: number;
  hour?: number;
  minute?: number;
  second?: number;
}

/**
 * What a filter asks of one of its columns (<filterColumn>). Exactly one kind
 * of criterion is stated; a column without any has only button settings.
 */
export interface FilterColumn {
  /** Offset of the column from the left edge of the filter range. */
  colId: number;
  /**
   * "values": the texts in `values` (compared with what the cells show), blanks and dates;
   * "custom": one or two conditions; "top10"; "dynamic" (above average, today, ...);
   * "color" and "icon" are read and kept but never applied; "none": no criterion;
   * "unknown": something this engine does not know, kept as it is.
   */
  kind: 'values' | 'custom' | 'top10' | 'dynamic' | 'color' | 'icon' | 'none' | 'unknown';
  /** values: the texts that are shown. */
  values?: string[];
  /** values: empty cells are shown as well. */
  blank?: boolean;
  dateGroups?: FilterDateGroup[];
  /** custom: both conditions have to hold (`and`) or one of them. */
  and?: boolean;
  conditions?: FilterCondition[];
  /** top10: the `value` largest (or smallest) items, or that percentage of them. */
  top?: boolean;
  percent?: boolean;
  value?: number;
  /** dynamic: "aboveAverage", "belowAverage", "today", ... as in ST_DynamicFilterType. */
  dynamicType?: string;
  /** The column shows no button (hiddenButton="1" or showButton="0"). */
  hiddenButton?: boolean;
  /**
   * The source element. While it is there the column is written back exactly as it
   * was read (attributes and extensions the model has no field for included);
   * a column changed by an edit has none.
   */
  raw?: string;
}

export interface SortCondition {
  /** The cells of the key column or row. */
  range: CellRange;
  descending: boolean;
  /** The source element, for what the model has no field for (sortBy, customList, dxfId, iconSet). */
  raw?: string;
}

/** How a range was last sorted (<sortState>); kept so that the application can sort again. */
export interface SortState {
  range: CellRange;
  conditions: SortCondition[];
  caseSensitive?: boolean;
  columnSort?: boolean;
  /** Attributes and children the model has no field for, as they go into the element. */
  attrs?: [name: string, value: string][];
  extra?: string;
}

/** Defined name, local to a sheet and hidden, under which applications keep the range of the sheet's filter. */
export const FILTER_NAME = '_xlnm._FilterDatabase';

/** The filter of a sheet (<autoFilter>): a header row with a button per column, and the rows under it. */
export interface AutoFilter {
  /** Header row and data rows. */
  range: CellRange;
  /** Columns that state something, by colId; at most one per column. */
  columns: FilterColumn[];
  sortState?: SortState;
  /** Attributes of the element the model has no field for (xr:uid, ...), kept to be written back. */
  attrs?: [name: string, value: string][];
  /** Child elements the model has no field for (<extLst>), as source text. */
  extra?: string;
  /**
   * Rows that a criterion this engine cannot apply (a colour, an icon) keeps hidden, noted
   * when another criterion is changed, so that they stay hidden whatever that one shows.
   * Not part of the file.
   */
  held?: number[];
}

/**
 * One <dataValidation>: what may be typed into its cells. A view for reading;
 * the writer copies the source element.
 */
export interface DataValidation {
  ranges: CellRange[];
  /** "none", "whole", "decimal", "list", "date", "time", "textLength", "custom". */
  type: string;
  /** "between" (the default), "notBetween", "equal", "notEqual", "greaterThan", "lessThan", "greaterThanOrEqual", "lessThanOrEqual". */
  operator: string;
  /** Formula texts without "=", written for the top-left cell of the first range. */
  formula1?: string;
  formula2?: string;
  allowBlank: boolean;
  /** A list shows no arrow in the cell. The file calls this showDropDown="1": the attribute means the opposite of its name. */
  hideDropDown: boolean;
  showInputMessage: boolean;
  promptTitle?: string;
  prompt?: string;
  showErrorMessage: boolean;
  errorStyle: 'stop' | 'warning' | 'information';
  errorTitle?: string;
  error?: string;
  /** The rule comes from the 2010 extension list (its list may be on another sheet). */
  extended?: boolean;
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
  /**
   * Index of the shared string this cell was read from. Only cells that come
   * from the file have it; the writer uses it to put the cell back exactly as
   * it was, phonetic text included.
   */
  si?: number;
  /** Cell metadata index (dynamic array formulas) and value metadata index (rich values), kept to be written back. */
  cm?: number;
  vm?: number;
  /** The cell shows phonetic text. */
  ph?: boolean;
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
  /**
   * Attributes of <row> the model has no field for (thickTop, ph, x14ac:dyDescent, ...),
   * kept to be written back. "spans" is not kept: it is a hint that is stale after any edit.
   */
  attrs?: [name: string, value: string][];
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
  /**
   * Conditional formats as the file has them. They are not changed by edits: rows and
   * columns inserted or deleted since (structuralEdits) are applied where they are evaluated.
   */
  conditionalFormats?: ConditionalFormat[];
  /** The filter of the sheet, kept up to date by edits. Rows it hides are `hidden`, like rows hidden by hand. */
  autoFilter?: AutoFilter;
  /** The filter was changed (switched on or off, a criterion, a sort); the writer regenerates <autoFilter> from the model. */
  autoFilterModified?: boolean;
  /**
   * Data validations as the file has them. Like conditional formats they are not changed
   * by edits: structuralEdits are applied where they are evaluated (see DataValidations).
   */
  dataValidations?: DataValidation[];
  /** Set once anything in the sheet was changed after it was read; the writer regenerates only such sheets. */
  modified?: boolean;
  /**
   * Rows in which a cell was changed since the sheet was read (indexes as they
   * were when the change happened). While there are no structural edits, every
   * other row is written back byte for byte from the original part.
   */
  changedRows?: Set<number>;
  /** Column widths, visibility or formats were changed; the writer regenerates <cols>. */
  columnsModified?: boolean;
  /** Cells were merged or unmerged; the writer regenerates <mergeCells>. */
  mergesModified?: boolean;
  /** Links changed without the sheet's own cells moving; the writer regenerates <hyperlinks>. */
  hyperlinksModified?: boolean;
  /**
   * Kinds of content anchored to cells but stored outside the sheet part
   * ("drawing", "comments", "table", "pivotTable", ...). Their positions are
   * not adjusted yet, so rows and columns cannot be inserted or deleted in a
   * sheet that has any.
   */
  anchored?: string[];
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
