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

/** A corner of a drawing: a cell and how far into it, in EMU (914,400 per inch; 9,525 per pixel). */
export interface DrawingMarker {
  col: number;
  colOff: number;
  row: number;
  rowOff: number;
  /**
   * How far into the cell as a share (0..1) of the column's width and the row's height, for
   * files that state it so (xls). Where present they take the place of colOff and rowOff.
   */
  colShare?: number;
  rowShare?: number;
}

/**
 * Where a drawing is on its sheet. "twoCell": from one cell to another; "oneCell": from a
 * cell, with a size; "absolute": a place and a size on the sheet, whatever its cells do.
 * Sizes and absolute positions are in EMU.
 */
export interface DrawingAnchor {
  kind: 'twoCell' | 'oneCell' | 'absolute';
  /** twoCell only: what rows and columns inserted or deleted do to it. "twoCell" when the file states nothing. */
  editAs?: 'twoCell' | 'oneCell' | 'absolute';
  from?: DrawingMarker;
  to?: DrawingMarker;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
}

export interface DrawingTextRun {
  text: string;
  /** Font name as the file states it; never trusted. */
  font?: string;
  /** Points. */
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** "#RRGGBB". */
  color?: string;
}

export interface DrawingParagraph {
  runs: DrawingTextRun[];
  align?: 'left' | 'center' | 'right' | 'justify';
}

/** The text of a shape. */
export interface DrawingText {
  paragraphs: DrawingParagraph[];
  /** Where the text stands in the height of the shape. Default "top". */
  anchor?: 'top' | 'middle' | 'bottom';
  /** Space between the shape and its text: left, top, right, bottom, in pixels. */
  insets?: [number, number, number, number];
  /** Lines are broken at the width of the shape. Default true. */
  wrap?: boolean;
}

/** What every object of a drawing has, whatever it shows. */
export interface DrawingObjectBase {
  name?: string;
  /** Alternative text. */
  description?: string;
  hidden?: boolean;
  /** Degrees, clockwise. */
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
  /**
   * For an object inside a group: its place in the group, as shares (0..1) of the
   * group's width and height. The object of an anchor fills the anchor and has none.
   */
  box?: { x: number; y: number; width: number; height: number };
}

export interface DrawingPicture extends DrawingObjectBase {
  type: 'picture';
  /** Key of the picture in Workbook.media; missing for a picture that is not in the file. */
  media?: string;
  /** The picture is linked from outside the file. It is never fetched. */
  external?: boolean;
  /** Shares (0..1) cut off each side of the picture. */
  crop?: { left: number; top: number; right: number; bottom: number };
}

export interface DrawingShape extends DrawingObjectBase {
  type: 'shape';
  /** Preset geometry name ("rect", "ellipse", "line", ...); "rect" when the file states none. */
  geometry: string;
  /** A connector or a line: drawn from corner to corner. */
  connector?: boolean;
  /** "#RRGGBB"; missing for a shape without a fill (or with one that is not drawn: gradients, patterns). */
  fill?: string;
  /** Outline: colour ("#RRGGBB") and width in pixels. */
  line?: { color: string; width: number; dash?: boolean };
  text?: DrawingText;
}

export interface DrawingGroup extends DrawingObjectBase {
  type: 'group';
  children: DrawingObject[];
}

/** Something that is not drawn as it is: a chart, a diagram, an embedded object. */
export interface DrawingPlaceholder extends DrawingObjectBase {
  type: 'placeholder';
  what: 'chart' | 'diagram' | 'object';
}

export type DrawingObject = DrawingPicture | DrawingShape | DrawingGroup | DrawingPlaceholder;

/**
 * One anchor of a sheet's drawing with what it holds. A view for reading: the drawing
 * part stays in the package and is written back from there.
 */
export interface SheetDrawing {
  anchor: DrawingAnchor;
  object: DrawingObject;
}

/** A comment (note) on a cell. A view for reading, like SheetDrawing. */
export interface SheetComment {
  row: number;
  col: number;
  author?: string;
  /** The whole text; for a legacy note that starts with the author's name, the name is part of it. */
  text: string;
  /** Present when the text has mixed formatting. */
  runs?: RichRun[];
  /** The comment is a threaded one (a conversation); `text` is then the first message and `replies` the rest. */
  threaded?: boolean;
  replies?: { author?: string; text: string }[];
  /** The box of the comment is shown all the time, not only when the cell is pointed at. */
  visible?: boolean;
  /** Where the box is, when the file says so. Offsets are in EMU, as in every anchor. */
  box?: DrawingAnchor;
}

/** The margins of the printed page, in inches (<pageMargins>). `header` and `footer` are the distances of the header and the footer from the edge of the paper. */
export interface PageMargins {
  left: number;
  right: number;
  top: number;
  bottom: number;
  header: number;
  footer: number;
}

/**
 * Paper, orientation and scaling of the printed page (<pageSetup>). A field is present when the
 * file states it; a missing one has the default named with it.
 */
export interface PageSettings {
  /** Paper size code (1 Letter, 5 Legal, 8 A3, 9 A4, 11 A5, ...). Default 1; with `usePrinterDefaults` the printer's paper. */
  paperSize?: number;
  /** Size of a paper that has no code, as a measure with its unit ("210mm"); replaces `paperSize`. */
  paperWidth?: string;
  paperHeight?: string;
  /** "default" (the default) leaves the choice to the application, which prints upright. */
  orientation?: 'default' | 'portrait' | 'landscape';
  /** Percent, 10..400. Default 100. Not used while the sheet is fitted to pages (SheetPageSetup.fitToPage). */
  scale?: number;
  /** Pages the sheet is fitted to, across and down, while `fitToPage` is set; 0 is "as many as needed". Default 1. */
  fitToWidth?: number;
  fitToHeight?: number;
  /** Number of the first page; counts only with `useFirstPageNumber`. Default 1. */
  firstPageNumber?: number;
  /** Default false: pages are numbered from 1. */
  useFirstPageNumber?: boolean;
  /** The order pages are printed and numbered in. Default "downThenOver". */
  pageOrder?: 'downThenOver' | 'overThenDown';
  /** Default false. */
  blackAndWhite?: boolean;
  /** Draft quality, without graphics. Default false. */
  draft?: boolean;
  /** Whether comments are printed, and where. Default "none". */
  cellComments?: 'none' | 'asDisplayed' | 'atEnd';
  /** What is printed for a cell that holds an error. Default "displayed". */
  errors?: 'displayed' | 'blank' | 'dash' | 'NA';
  /** Print resolution in dots per inch; values near 4,294,967,295 are the printer driver's quality levels (draft, low, medium, high). Default 600. */
  horizontalDpi?: number;
  verticalDpi?: number;
  /** Default 1. */
  copies?: number;
  /** What the element does not state is the printer's own setting and not the default of the format. Default true. */
  usePrinterDefaults?: boolean;
}

/** What is printed beside the cells, and where on the page (<printOptions>). Present when the file states it, like PageSettings. */
export interface PrintOptions {
  /** The cells are centred between the left and the right margin. Default false. */
  horizontalCentered?: boolean;
  /** The cells are centred between the top and the bottom margin. Default false. */
  verticalCentered?: boolean;
  /** Row numbers and column letters are printed. Default false. */
  headings?: boolean;
  /** Grid lines are printed; only together with `gridLinesSet`. Default false. */
  gridLines?: boolean;
  /** Default true. */
  gridLinesSet?: boolean;
}

/**
 * Headers and footers of the printed page (<headerFooter>). A text is in the code of the
 * format: "&L", "&C" and "&R" start the left, centre and right section, "&P" is the page
 * number, "&N" the page count, "&D" the date, "&T" the time, "&F" the file name, "&A" the
 * sheet name, '&"font,style"' and "&12" set a font and a size, "&&" is an ampersand.
 */
export interface HeaderFooter {
  /** Header and footer of every page, or of the odd pages only while `differentOddEven` is set. */
  oddHeader?: string;
  oddFooter?: string;
  /** Used while `differentOddEven` is set. */
  evenHeader?: string;
  evenFooter?: string;
  /** Used while `differentFirst` is set. */
  firstHeader?: string;
  firstFooter?: string;
  /** Default false. */
  differentOddEven?: boolean;
  /** Default false. */
  differentFirst?: boolean;
  /** Headers and footers are scaled with the sheet. Default true. */
  scaleWithDoc?: boolean;
  /** Headers and footers start and end at the page margins. Default true. */
  alignWithMargins?: boolean;
}

/** A page break (<brk>). */
export interface PageBreak {
  /**
   * The first row (of rowBreaks) or column (of colBreaks) of the new page: the break is above
   * that row, left of that column. Never 0.
   */
  at: number;
  /** The columns (of a row break) or rows (of a column break) the break spans, first and last. */
  min: number;
  max: number;
  /** Set by hand, not by the application while fitting pages. */
  manual: boolean;
  /** Made by a pivot table. */
  pivot?: boolean;
}

/**
 * How a sheet is printed: one member per element of a worksheet part (<pageMargins>,
 * <pageSetup>, <printOptions>, <headerFooter>, <rowBreaks>, <colBreaks>, and <pageSetUpPr> of
 * <sheetPr>). A view for reading, like conditional formats: a save of the file it was read
 * from copies the source elements, and edits do not change it (structuralEdits are applied
 * to the breaks where a new file is written). Print areas and print titles are defined
 * names ("_xlnm.Print_Area", "_xlnm.Print_Titles"). Data of a printer driver is not held.
 */
export interface SheetPageSetup {
  margins?: PageMargins;
  page?: PageSettings;
  print?: PrintOptions;
  headerFooter?: HeaderFooter;
  /** The sheet is scaled to `page.fitToWidth` by `page.fitToHeight` pages, not by `page.scale`. Default false. */
  fitToPage?: boolean;
  /** Automatic page breaks are shown on the sheet. Default true. */
  autoPageBreaks?: boolean;
  /** In ascending order, one per row. */
  rowBreaks?: PageBreak[];
  /** In ascending order, one per column. */
  colBreaks?: PageBreak[];
}

/** Where the bytes of pictures come from; read only when a picture is first shown. */
export interface WorkbookMedia {
  readPart(key: string): Promise<Uint8Array>;
}

export type SheetState = 'visible' | 'hidden' | 'veryHidden';
export type SheetKind = 'worksheet' | 'chartsheet' | 'dialogsheet' | 'macrosheet';

export interface Sheet {
  name: string;
  sheetId: number;
  state: SheetState;
  kind: SheetKind;
  /**
   * Part this sheet was read from, e.g. "/xl/worksheets/sheet1.xml". Empty for a sheet that is
   * not in the file (one added since it was read): the writer gives such a sheet its part.
   */
  part: string;
  /**
   * The name the sheet has in the file, set when a sheet that was read from the file is first
   * given another name. Content kept the way the file has it (conditional formats, data
   * validations, the source XML) names sheets as the file does: see fileSheetName() and
   * sheetNameMap(). Undoing the rename removes it again; renaming the sheet back to the name
   * of the file by another rename leaves it, equal to `name`.
   */
  nameInFile?: string;
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
  /** Margins, paper, scaling, headers and footers and page breaks of the printed sheet, as the file has them. */
  pageSetup?: SheetPageSetup;
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
  /**
   * Pictures, shapes and placeholders of charts as the file has them, bottom one first. Like
   * conditional formats they are not changed by edits: structuralEdits are applied where they
   * are shown (see SheetDrawings), by the rules the writer moves the anchors in the file with.
   */
  drawings?: SheetDrawing[];
  /** Comments as the file has them; structuralEdits are applied where they are shown. */
  comments?: SheetComment[];
  /** What of the sheet's drawings and comments could not be read and is not shown (it is still written back). */
  warnings?: string[];
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
  /** The pictures of the workbook's drawings, by DrawingPicture.media. */
  media?: WorkbookMedia;
  /**
   * Set once a sheet was inserted, deleted or renamed after the workbook was read (and never
   * cleared, not by undo either): the list of sheets, their names or the defined names may
   * differ from the file, and the writer regenerates what holds them.
   */
  sheetsModified?: boolean;
  /**
   * The names, as the file has them, of sheets of the file that were deleted. Undoing the
   * deletion takes the name out again. A sheet added after the file was read and deleted
   * again is not listed. Missing while nothing was deleted.
   */
  removedSheetNames?: string[];
  /**
   * One entry per name of removedSheetNames, in the same order: the sheets that stood before
   * and after the deleted sheet. A 3D reference that ended at the deleted sheet ends at one of
   * them (see SheetRuns). Kept pointing at sheets of the workbook: when such a neighbour is
   * deleted in turn, its own neighbour takes its place here.
   */
  removedSheetNeighbours?: SheetNeighbours[];
}

/** The sheets on either side of a deleted sheet; undefined where it was the first or the last. */
export interface SheetNeighbours {
  previous?: Sheet;
  next?: Sheet;
}

/**
 * Sheet names compare without regard to case. The four Turkish letters I, İ, ı, i count as
 * one, without relying on locale data: "İ" lower-cases to "i" plus a combining dot (U+0307).
 */
export function foldSheetName(name: string): string {
  return String(name).toLowerCase().replace(/\u0307/g, '').replace(/\u0131/g, 'i');
}

/**
 * What a sheet may be called: 1 to 31 characters that are not all spaces,
 * none of []:*?/\ and no control character, no apostrophe at either end,
 * not "History" (which Excel keeps for itself), and only characters an XML
 * file can hold. Returns the name; throws RangeError for a name that is none.
 */
export function checkSheetName(name: string): string {
  const ok =
    typeof name === 'string' &&
    name.length >= 1 &&
    name.length <= 31 &&
    name.trim() !== '' &&
    !/[[\]:*?/\\\u0000-\u001f\ufffe\uffff]/.test(name) &&
    // A surrogate without its other half.
    !/[\ud800-\udbff](?![\udc00-\udfff])|(?:^|[^\ud800-\udbff])[\udc00-\udfff]/.test(name) &&
    name.charAt(0) !== "'" &&
    name.charAt(name.length - 1) !== "'" &&
    name.toLowerCase() !== 'history';
  if (!ok) throw new RangeError('A sheet name has 1 to 31 characters, none of [ ] : * ? / \\ and no control character, no apostrophe at its start or end, and is not "History"');
  return name;
}

/** The name a sheet has in the file it was read from; undefined for a sheet that is not in the file. */
export function fileSheetName(sheet: Sheet): string | undefined {
  if (sheet.part === '') return undefined;
  return sheet.nameInFile !== undefined ? sheet.nameInFile : sheet.name;
}

/**
 * What a sheet name of the file stands for now: the name the sheet has today, null when the
 * sheet was deleted, undefined when the name is to be left as it is (the sheet still has it,
 * or the file has no sheet of that name).
 */
export interface SheetNameMap {
  (name: string): string | null | undefined;
  /** How a 3D reference goes on when the sheet at one of its ends was deleted; without it such a reference is an error. */
  runs?: SheetRuns;
}

/**
 * What a 3D reference ("Sheet2:Sheet4!A1") needs to know when a sheet at one of its ends was
 * deleted: the run then ends at the sheet that stood next to the deleted one, inside the run.
 */
export interface SheetRuns {
  /**
   * The name, as it is now, of the sheet that stood after (`step` 1) or before (`step` -1) the
   * deleted sheet the map answers null for; undefined when there is none.
   */
  neighbour(name: string, step: 1 | -1): string | undefined;
  /** Where the sheet of this name (as it is now) stands among the sheets; -1 for no sheet. */
  index(name: string): number;
}

/**
 * The map from the sheet names of the file to the sheets as they are now, for content that
 * is kept the way the file has it. Undefined while every sheet of the file is there under
 * its own name, which is the usual case and needs no work. All names are looked up in the
 * state of the file at once, never one rename after another: with "Sheet2" deleted and
 * "Sheet3" renamed to "Sheet2", the file's "Sheet2" is the deleted sheet (null) and the
 * file's "Sheet3" is today's "Sheet2". The map reflects the workbook at the time of the
 * call; ask again after the sheets changed.
 */
export function sheetNameMap(workbook: Workbook): SheetNameMap | undefined {
  const removed = workbook.removedSheetNames;
  let names: Map<string, string | null> | undefined;
  if (removed && removed.length > 0) {
    names = new Map();
    for (const name of removed) names.set(foldSheetName(name), null);
  }
  for (const sheet of workbook.sheets || []) {
    if (sheet.part === '' || sheet.nameInFile === undefined || sheet.nameInFile === sheet.name) continue;
    if (!names) names = new Map();
    names.set(foldSheetName(sheet.nameInFile), sheet.name);
  }
  if (!names) return undefined;
  const known = names;
  const map: SheetNameMap = (name) => known.get(foldSheetName(name));
  const neighbours = workbook.removedSheetNeighbours;
  if (removed && removed.length > 0 && neighbours) {
    const sheets = workbook.sheets;
    const beside = new Map<string, SheetNeighbours>();
    removed.forEach((name, i) => {
      if (neighbours[i]) beside.set(foldSheetName(name), neighbours[i]);
    });
    map.runs = {
      neighbour: (name, step) => {
        const entry = beside.get(foldSheetName(name));
        const sheet = entry && (step > 0 ? entry.next : entry.previous);
        return sheet && sheets.indexOf(sheet) >= 0 ? sheet.name : undefined;
      },
      index: (name) => {
        const wanted = foldSheetName(name);
        for (let i = 0; i < sheets.length; i++) if (foldSheetName(sheets[i].name) === wanted) return i;
        return -1;
      },
    };
  }
  return map;
}

/** What one sheet name of the file stands for now; sheetNameMap() is the one to use for many names. */
export function currentSheetName(workbook: Workbook, nameInFile: string): string | null | undefined {
  const map = sheetNameMap(workbook);
  return map ? map(nameInFile) : undefined;
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
