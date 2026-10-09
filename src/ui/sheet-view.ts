// Workbook view: shows one sheet of a workbook inside a container element and
// handles scrolling, selection, keyboard, touch and the sheet tabs.
//
// Scrolling is native: a scroll container holds the canvas (kept in place with
// sticky positioning) and a spacer that gives the container its scrollable
// size. The canvas only ever covers the visible area; drawing is done by
// renderGrid(). Nothing in this module touches the DOM at import time, so it
// can be imported where there is no document.

import { colX, createLayout, hitTest, maxScroll, rangeRect, rowY, scrollAreaSize, scrollToReveal, type GridLayout } from '../layout/grid-geometry';
import { DEFAULT_DIGIT_WIDTH, createSheetMetrics, type AxisMetrics, type SheetMetrics } from '../layout/sheet-metrics';
import { MAX_COLS, MAX_ROWS, cellName, columnName, foldSheetName, getCell, parseCellRef, rangeName, type AutoFilter, type CellRange, type DataValidation, type Sheet, type SheetComment, type Workbook } from '../model/sheet/types';
import { cellDisplayText } from '../format/cell-format';
import { getLocale, normalizeLocale, type FormatLocale } from '../format/locale';
import { resolveUiStrings, type UiStrings } from './strings';
import { resolveCellStyle, resolveColor, type ResolvedCellStyle } from '../model/sheet/color';
import { clearRange, copyRange, cutRange, editTextAt, pasteAt, setCellInput, styleAt, type CopiedRange } from '../ops/commands';
import { formulaReferences } from '../calc/refs';
import { cycleAbsolute } from '../edit/absolute';
import { canPointAt, isFormulaText, pointReference, pointSpanLive } from '../edit/point';
import {
  formatAt,
  formatRange,
  mergeRange,
  setColumnWidth as columnWidthOps,
  setRowHeight as rowHeightOps,
  unmergeRange,
  type StylePatch,
} from '../ops/format';
import { CalcEngine } from '../calc/engine';
import { ConditionalFormats } from '../format/conditional';
import { DataValidations, MAX_LIST_ITEMS, ValidationError, tryOps } from '../format/validation';
import {
  MAX_FILTER_VALUES,
  deleteColumns as deleteColumnOps,
  deleteRows as deleteRowOps,
  filterSummary,
  filterValues,
  isFilterActive,
  rowRunsWithout,
  rowsHiddenByFilter,
  setAutoFilter,
  setFilterColumn,
} from '../ops/filter';
import { History, markStyles, resetStyles, type Op } from '../ops/ops';
import { MAX_NOTE_TEXT, commentAtCell, noteProblem, setComment as setCommentOps } from '../ops/notes';
import { hasDefinedName } from '../ops/commands';
import {
  CONDITIONAL_TYPES,
  addConditionalRule as addConditionalRuleOps,
  clearConditionalRules as clearConditionalRulesOps,
  conditionalRules as conditionalRulesOf,
  moveConditionalRule as moveConditionalRuleOps,
  removeConditionalRule as removeConditionalRuleOps,
  setValidation as setValidationOps,
  subtractRanges,
  updateConditionalRule as updateConditionalRuleOps,
  validationProblem,
  validationsIn,
  type ConditionalPatch,
  type ConditionalRuleInfo,
  type ConditionalSettings,
  type ValidationSettings,
} from '../ops/rules';
import {
  ConditionalDialog,
  ConditionalManager,
  DIALOG_STYLE,
  FieldError,
  managerPage,
  ValidationDialog,
  validationChange,
  conditionalFields,
  conditionalFromFields,
  conditionalPatch,
  describeRule,
  parseRangeList,
  rangeListText,
  validationFields,
  validationFromFields,
  type ManagerRow,
} from './rules-dialog';
import {
  canCopySheet,
  canDeleteSheet,
  canHideSheet,
  canRenameSheet,
  copySheet as copySheetOps,
  copySheetLosses,
  defaultSheetName,
  deleteSheet as deleteSheetOps,
  insertSheet as insertSheetOps,
  moveSheet as moveSheetOps,
  renameSheet as renameSheetOps,
  setSheetState as setSheetStateOps,
  setTabColor as setTabColorOps,
  type CopySheetOptions,
} from '../ops/sheets';
import { Popover, type PopoverItem } from './popover';
import { PALETTE, PALETTE_COLUMNS } from './toolbar';
import { currentRegion, hasHeaderRow, sortRange } from '../ops/sort';
import { DEFAULT_GRID_THEME, cellButtonRect, renderGrid, type GridFilter, type GridSelection } from '../render/grid';
import { IMAGE_BUDGET_BYTES, IMAGE_BUDGET_BYTES_COARSE, ImageStore } from '../render/images';
import { SheetDrawings } from '../render/sheet-drawings';
import { onFontsChanged } from '../render/fonts';
import { onEmbeddedFontNotice } from '../render/embedded-fonts';
import { cssFont } from '../render/text';
import { ContextMenu, type MenuItem } from './context-menu';
import {
  DocumentPrintError,
  MAX_PRINT_PAGES,
  PrintFrame,
  canvasPicture,
  openPrintDialog,
  planPrint,
  printLimits,
  revokePicture,
  supportsNamedPages,
  type DocumentPrintResult,
  type PrintDialog,
  type PrintLimits,
  type PrintPlan,
} from './doc-print';
import { defaultPaperFor, paintSheetPage, planSheetPages, type SheetJobPlan, type SheetPrintOptions } from './sheet-print';
import { FilterMenu, ListMenu } from './filter-menu';
import { allSheets, dropIndex, recalculateForSheets, sheetErrorText, sheetToShow, type SheetListChange, type SheetOpKind } from './sheet-tabs';

export interface WorkbookViewOptions {
  /**
   * Language for number and date formats: a tag such as "tr" or "en-US", or a
   * complete FormatLocale. Defaults to the browser's language.
   */
  locale?: string | FormatLocale;
  /** Sheet shown first; defaults to workbook.activeSheet, or the first visible sheet when that one is hidden. */
  sheet?: number;
  onSelectionChange?(selection: GridSelection, sheetIndex: number): void;
  onSheetChange?(sheetIndex: number): void;
  /** No editing: no cell editor, no paste, cut, delete, undo or structural commands. Copy still works. Default false. */
  readOnly?: boolean;
  /** Recalculate formulas after every change. Default true. */
  calculate?: boolean;
  /** Who the notes added in this view are by. Default "Author". A note that is edited keeps its author. */
  author?: string;
  /** Called after every change to the workbook, including undo and redo. */
  onChange?(change: { sheets: number[] }): void;
  /**
   * Called with every message the view has for the user, in the language of
   * the view: what an action was refused for ('error'), or a hint ('info').
   * The same text goes to the view's live region for assistive technology.
   */
  onNotice?(notice: { message: string; kind: 'error' | 'info' }): void;
  /** Called when cell editing starts, on every change of the text being edited, and when it ends. */
  onEditingChange?(state: { editing: boolean; text: string }): void;
  /**
   * An element around the view (an editor shell with a formula bar, say).
   * While the focus is inside it, an open edit is not written just because
   * the focus left the cell editor.
   */
  focusScope?: HTMLElement;
  /**
   * Called while print() prepares its pages, with how many of them are ready,
   * and with undefined once that is over: the dialog is open, or the job ended.
   */
  onPrintProgress?(progress: { done: number; total: number } | undefined): void;
}

/**
 * What the view stores with every undo step: the sheet shown and its
 * selection, which undo and redo bring back, and the sheets the step touches.
 */
export interface HistoryMark {
  sheet: number;
  active: { row: number; col: number };
  range: CellRange;
  /** Indexes of the sheets the step may have changed. */
  sheets: number[];
  /** Rows, columns or merged cells changed (inserted, deleted, resized, hidden), so sizes and positions changed too. */
  structural: boolean;
  /**
   * The cells the step wrote, for recalculating what depends on them; missing
   * when rows or columns moved or too many cells changed, which recalculates everything.
   */
  cells?: { sheet: number; row: number; col: number }[];
  /** Sheets on which the step hid or showed rows (by hand, by a filter): SUBTOTAL and AGGREGATE read that. */
  rows?: number[];
  /**
   * The step changed the list of sheets: one was inserted, deleted or renamed. The indexes of
   * the mark are those of the list the step leads to (as every mark's are): `sheet` is the
   * sheet shown then, and `sheets` are all of them.
   */
  sheetOp?: SheetOpKind;
  /**
   * With a sheet deleted: the cells of other sheets whose formulas read it and were rewritten,
   * by their sheets (indexes differ before and after the step). Missing when there are too many.
   */
  sheetCells?: { sheet: Sheet; row: number; col: number }[];
}

/** Longest a recalculation holds the page before it leaves the rest for the next turn. */
const CALC_SLICE_MS = 250;

/** Most turns a calculation is given at once when something has to wait for its end. */
const MAX_FINISH_SLICES = 8;

/** Above this many written cells a step recalculates every formula instead of tracking what depends on them. */
const MAX_RECALC_CELLS = 100000;

/** The cell being edited. */
interface Edit {
  /** The sheet of the cell: another one may be shown while a formula is pointed at its cells. */
  home: Sheet;
  row: number;
  col: number;
  /**
   * The arrow keys point at cells where the formula takes a reference (the edit was begun by
   * typing); otherwise they move the caret in the text (it was begun with F2 or a double click).
   * F2 goes from one to the other.
   */
  keys: boolean;
  /** Cells were pointed at during this edit, which may have scrolled the cell out of view. */
  pointed: boolean;
  /** The text the cell had for editing when the edit began; committing the same text changes nothing. */
  original: string;
  /**
   * The edit began (through an input method) with the idle text still in the
   * input. If that is all there is when it ends, nothing was entered.
   */
  idle: string | undefined;
  font: string;
  wrap: boolean;
}

interface CellPos {
  row: number;
  col: number;
}

/** The reference a formula that is being typed was last pointed at, and where it stands in the text. */
interface Point {
  sheet: Sheet;
  active: CellPos;
  focus: CellPos;
  range: CellRange;
  /** The field the text is typed in: the editor of the grid, or the field of a formula bar. */
  field: HTMLTextAreaElement;
  start: number;
  end: number;
  text: string;
}

/** Marks the field of a shell around the view (a formula bar) that edits the cell together with the grid's editor. */
export const EDIT_FIELD_ATTRIBUTE = 'data-bym-edit-field';

/** What is remembered for a sheet while another one is shown. */
interface SheetState {
  sheet: Sheet;
  metrics: SheetMetrics;
  /** Logical scroll offsets in sheet pixels, after the frozen rows and columns. */
  scrollX: number;
  scrollY: number;
  /** Rows and columns the scrollable area currently reaches; only ever grows. */
  extentRows: number;
  extentCols: number;
  active: CellPos;
  /** The corner of the range that moves when the selection is extended. */
  focus: CellPos;
  range: CellRange;
  merges: MergeIndex;
}

/**
 * The merged ranges of a sheet arranged for lookup by row. `byTop` is sorted
 * by first row; `maxBottom[i]` is the largest last row among byTop[0..i], so
 * it never decreases and the first range that can reach a row is found by
 * binary search.
 */
interface MergeIndex {
  /** The array the index was built from and its length then; a change means the index is stale. */
  source: CellRange[];
  size: number;
  byTop: CellRange[];
  maxBottom: number[];
}

/** A column or row whose size is being dragged. */
interface Resize {
  axis: 'row' | 'col';
  index: number;
  /** Canvas coordinate of the edge the column or row starts at. */
  origin: number;
}

interface Drag {
  pointerId: number;
  mode: 'cell' | 'row' | 'col' | 'all' | 'resize';
  resize?: Resize;
  /** Last pointer position on the canvas. */
  x: number;
  y: number;
  originX: number;
  originY: number;
  moved: boolean;
  /** The drag points at cells for the formula that is being typed instead of selecting them. */
  point?: boolean;
}

interface Touch {
  pointerId: number;
  x: number;
  y: number;
  domX: number;
  domY: number;
  moved: boolean;
  /** The touch opened the menu (a long press), so lifting the finger is not a tap. */
  consumed: boolean;
  /** The touch began on a border between headers: it resizes once it moves, and is a tap on the header if it does not. */
  resize?: Resize;
}

/**
 * Browsers limit the size of an element (Firefox to about 17.8 million pixels).
 * A sheet can be taller than that, so beyond this size the native scroll
 * position is mapped proportionally onto the logical one.
 */
const MAX_SPACER = 8000000;
/** Rows and columns kept scrollable after the used area. */
const EXTRA_ROWS = 100;
const EXTRA_COLS = 26;
/** Canvas backing stores larger than this fail on some mobile browsers. */
const MAX_CANVAS_PIXELS = 16777216;
/** Pointer travel before a press becomes a drag (mouse) or stops being a tap (touch). */
const DRAG_SLOP = 3;
const TAP_SLOP = 10;
/** Slowest and fastest scroll per frame while a selection is dragged beyond the cells. */
const AUTO_SCROLL_MIN = 4;
const AUTO_SCROLL_MAX = 48;
/** Upper bound for the steps one Shift+arrow takes to get out of a merged range. */
const MAX_EXTEND_STEPS = 4096;

/** Colours of the mark in the corner of a formula cell the engine has something to say about, by what it says. */
const MARK_COLORS: { [kind: string]: string } = { file: '#E8710A', circular: '#D93025', stale: '#5F6368' };

/** What is known about the active cell beyond its content: worked out once per selection and change. */
interface CellInfo {
  sheet: Sheet;
  row: number;
  col: number;
  /** The cell offers a list to choose from. */
  list: boolean;
  /** The message its data validation shows while it is selected. */
  prompt: { title: string; text: string } | undefined;
}

/** How long a touch has to rest on a header to open its menu, and how close in time and place two taps make a double tap. */
const LONG_PRESS_MS = 550;
/** How far a pointer goes on a tab before the tab is dragged. */
const TAB_DRAG_SLOP = 6;
/** How near the ends of the strip a dragged tab scrolls it, by how much and how often. */
const TAB_SCROLL_EDGE = 24;
const TAB_SCROLL_STEP = 12;
const TAB_SCROLL_MS = 30;
/** How near a border between two headers the pointer has to be to resize, for a mouse and for a finger. */
const RESIZE_REACH_MOUSE = 4;
const RESIZE_REACH_TOUCH = 10;
/** Largest sizes a file can state: 255 characters and 409.5 points. */
const MAX_COLUMN_PIXELS = 1785;
const MAX_ROW_PIXELS = 546;
/** Rows looked at to fit a column to its content. */
const AUTO_FIT_ROWS = 2000;
/** Space the grid leaves around the text of a cell, both sides together. */
const CELL_TEXT_MARGIN = 8;
const DOUBLE_TAP_MS = 500;
/** Width of the button that adds a sheet, for a mouse and for a finger. */
const ADD_SHEET_WIDTH = 36;
const ADD_SHEET_WIDTH_COARSE = 44;
/** Longest name a sheet can have, in UTF-16 units as the field counts them. */
const MAX_SHEET_NAME = 31;

/** Language of the view's own words: the tag given as `locale`, else the browser's. */
const resolveStrings = resolveUiStrings;

// Cosmetic rules only; everything the layout depends on is set inline so the
// view still works where a style element is not allowed.
const STYLE_TEXT =
  '.bym-view{font:12px/1.2 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;color:#202124;background:#fff}' +
  '.bym-focus-ring .bym-scroller,.bym-tab:focus{outline:2px solid #1a7f45;outline-offset:-2px}' +
  '.bym-tab:focus:not(:focus-visible){outline:none}' +
  '.bym-input{color:#000;background:#fff;box-shadow:none}' +
  '.bym-menu{min-width:160px;padding:4px 0;background:#fff;border:1px solid #dadce0;border-radius:4px;box-shadow:0 2px 8px rgba(0,0,0,.25)}' +
  // Sizes are stated in full: a page rule for every button or textarea must not reach into the view.
  '.bym-menu-item{-webkit-appearance:none;appearance:none;display:block;float:none;box-sizing:border-box;width:auto;height:auto;min-width:0;min-height:0;' +
  'max-width:none;max-height:none;margin:0;border:0;border-radius:0;background:transparent;color:inherit;font:inherit;line-height:1.2;' +
  'text-align:left;padding:8px 16px;white-space:nowrap;cursor:pointer}' +
  '.bym-menu-item:hover,.bym-menu-item:focus{background:#e8eaed;outline:none}' +
  '.bym-menu-item:disabled,.bym-menu-item[aria-disabled="true"]{color:#9aa0a6;background:transparent;cursor:default}' +
  '.bym-menu-item[aria-disabled="true"]:focus{outline:1px dotted #9aa0a6;outline-offset:-3px}' +
  '.bym-menu-separator{height:1px;margin:4px 0;background:#dadce0}' +
  '.bym-resize-guide{background:#1a7f45}' +
  // What drops down from a cell: the menu of a filter button and the list of a data validation.
  '.bym-panel{padding:6px;background:#fff;border:1px solid #dadce0;border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,.18);color:#202124;' +
  'font:13px/1.2 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;letter-spacing:normal;text-transform:none;text-align:left;overflow-y:auto;overscroll-behavior:contain}' +
  '.bym-panel-item,.bym-panel-button,.bym-panel-search{-webkit-appearance:none;appearance:none;box-sizing:border-box;float:none;margin:0;min-width:0;min-height:0;' +
  'max-width:none;max-height:none;font:inherit;line-height:1.2;letter-spacing:normal;text-transform:none;text-indent:0;text-shadow:none;opacity:1}' +
  '.bym-panel-item{display:block;flex:none;width:100%;height:32px;padding:0 10px;border:0;border-radius:6px;background:transparent;color:inherit;' +
  'text-align:left;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;cursor:pointer}' +
  '.bym-panel-item:hover,.bym-panel-item:focus{background:#e8eaed;outline:none}' +
  '.bym-panel-item:disabled{color:#9aa0a6;background:transparent;cursor:default}' +
  '.bym-panel-separator{flex:none;height:1px;margin:4px 0;background:#dadce0}' +
  '.bym-panel-search{display:block;flex:none;width:100%;height:32px;margin:6px 0;padding:0 8px;border:1px solid #dadce0;border-radius:6px;background:#fff;color:inherit;outline:none}' +
  '.bym-panel-search:focus,.bym-list:focus{border-color:#1a7f45;box-shadow:0 0 0 1px #1a7f45}' +
  '.bym-list{border:1px solid #dadce0;border-radius:6px;outline:none}' +
  '.bym-dropdown .bym-list{border:0}' +
  '.bym-list-row{display:flex;align-items:center;padding:0 8px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;cursor:pointer}' +
  '.bym-list-row:hover,.bym-list:focus .bym-list-row[data-active]{background:#e8eaed}' +
  '.bym-dropdown .bym-list-row{display:block;line-height:28px}' +
  '.bym-dropdown .bym-list-row[aria-selected="true"]{font-weight:600;color:#0f5c31}' +
  '.bym-list-text{overflow:hidden;text-overflow:ellipsis}' +
  '.bym-check{flex:none;box-sizing:border-box;width:16px;height:16px;margin-right:8px;border:1px solid #5f6368;border-radius:3px;background:#fff}' +
  '.bym-check[data-state="on"],.bym-check[data-state="mixed"]{background:#1a7f45;border-color:#1a7f45}' +
  '.bym-check[data-state="on"]::after{content:"";display:block;width:4px;height:8px;margin:1px 0 0 4px;border:solid #fff;border-width:0 2px 2px 0;transform:rotate(45deg)}' +
  '.bym-check[data-state="mixed"]::after{content:"";display:block;width:8px;height:2px;margin:6px 0 0 3px;background:#fff}' +
  '.bym-panel-note{flex:none;padding:6px 4px 0;color:#5f6368;font-size:12px;white-space:normal}' +
  '.bym-panel-actions{display:flex;flex:none;justify-content:flex-end;margin-top:8px}' +
  '.bym-panel-button{display:block;flex:none;width:auto;height:32px;margin-left:8px;padding:0 14px;border:1px solid #dadce0;border-radius:6px;background:#fff;color:inherit;cursor:pointer}' +
  '.bym-panel-primary{background:#1a7f45;border-color:#1a7f45;color:#fff}' +
  '.bym-panel-button:disabled{opacity:.45;cursor:default}' +
  '.bym-panel-button:focus{outline:2px solid #1a7f45;outline-offset:1px}' +
  // A finger needs more to hit; text smaller than 16 px makes iOS zoom the page when a field takes the focus.
  '@media (pointer:coarse){.bym-panel-item,.bym-panel-button,.bym-panel-search{height:40px}.bym-panel-search{font-size:16px}.bym-dropdown .bym-list-row{line-height:40px}}' +
  '.bym-tip{max-width:240px;padding:6px 8px;background:#fffbe6;border:1px solid #e0d9a8;border-radius:4px;box-shadow:0 1px 4px rgba(0,0,0,.2);color:#202124;' +
  'font:12px/1.3 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;white-space:pre-wrap;overflow-wrap:break-word}' +
  '.bym-tip-title{font-weight:600}' +
  // The comment of a cell, beside it.
  '.bym-comment{padding:6px 8px;background:#ffffe1;border:1px solid #b8b28a;border-radius:2px;box-shadow:0 1px 4px rgba(0,0,0,.25);color:#202124;' +
  'font:12px/1.35 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;letter-spacing:normal;text-transform:none;text-align:left;' +
  'white-space:pre-wrap;overflow-wrap:break-word;word-break:break-word;overflow-x:hidden;overflow-y:auto;overscroll-behavior:contain}' +
  '.bym-comment-author{font-weight:700}' +
  // The field a note is typed in, in the place of the box that shows it.
  '.bym-note-edit{background:#ffffe1;border:1px solid #b8b28a;border-radius:2px;box-shadow:0 1px 4px rgba(0,0,0,.25);overflow:hidden}' +
  '.bym-note-input{-webkit-appearance:none;appearance:none;display:block;box-sizing:border-box;width:100%;height:100%;min-width:0;min-height:0;margin:0;padding:6px 8px;border:0;border-radius:0;' +
  'outline:0;box-shadow:none;background:transparent;color:#202124;font:12px/1.35 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;letter-spacing:normal;text-transform:none;' +
  'text-align:left;white-space:pre-wrap;overflow-wrap:break-word;resize:none;overflow-x:hidden;overflow-y:auto;overscroll-behavior:contain}' +
  '.bym-note-input:focus-visible{outline:2px solid #1a73e8;outline-offset:-2px}' +
  '.bym-note-input[aria-invalid="true"]{outline:2px solid #d93025;outline-offset:-2px}' +
  '.bym-comment-reply{margin-top:6px;padding-top:6px;border-top:1px solid #d8d2a8}' +
  '.bym-tabs{background:#f1f3f4;border-top:1px solid #dadce0}' +
  '.bym-tab{-webkit-appearance:none;appearance:none;display:block;float:none;flex:none;box-sizing:border-box;width:auto;height:auto;min-width:0;' +
  'max-width:none;max-height:none;margin:0;border:0;border-right:1px solid #dadce0;border-radius:0;' +
  'background:transparent;color:#444746;font:inherit;line-height:1.2;padding:0 14px;min-height:32px;white-space:nowrap;cursor:pointer;' +
  // A finger that rests on a tab opens its menu, not the browser's selection or its preview; two quick taps are two taps.
  '-webkit-user-select:none;user-select:none;-webkit-touch-callout:none;touch-action:manipulation}' +
  // A finger needs more to hit than a mouse pointer.
  '@media (pointer:coarse){.bym-tab{min-height:40px}}' +
  '.bym-tab:hover{background:#e8eaed}' +
  '.bym-tab[aria-selected="true"]{background:#fff;color:#1a7f45;font-weight:600}' +
  '.bym-tab-add{-webkit-appearance:none;appearance:none;display:block;float:none;box-sizing:border-box;min-width:0;min-height:0;max-width:none;max-height:none;' +
  'margin:0;padding:0;border:0;border-top:1px solid #dadce0;border-right:1px solid #dadce0;border-radius:0;background:#f1f3f4;color:#444746;' +
  'font:inherit;font-size:18px;line-height:1;text-align:center;cursor:pointer}' +
  '.bym-tab-add:hover{background:#e8eaed}' +
  '.bym-tab-add[aria-disabled="true"]{opacity:.38;cursor:default;background:transparent}' +
  '.bym-tab-add:focus{outline:2px solid #1a7f45;outline-offset:-2px}' +
  '.bym-tab-add:focus:not(:focus-visible){outline:none}' +
  '.bym-tab-input{-webkit-appearance:none;appearance:none;display:block;float:none;flex:none;box-sizing:border-box;height:auto;min-width:0;max-width:none;max-height:none;' +
  'margin:0;padding:0 8px;border:2px solid #1a7f45;border-radius:0;outline:none;background:#fff;color:#202124;font:inherit;line-height:1.2;min-height:32px}' +
  '.bym-tab-input[aria-invalid="true"]{border-color:#d93025}' +
  // Text smaller than 16 px makes iOS zoom the page when a field takes the focus.
  '@media (pointer:coarse){.bym-tab-input{min-height:40px;font-size:16px}}' +
  '.bym-note{color:#5f6368;font-size:14px}' +
  DIALOG_STYLE;

/** One style element per document or shadow root, shared by the views in it. */
const styleUsers = new WeakMap<Node, { element: HTMLStyleElement; count: number }>();

function acquireStyle(container: HTMLElement): Node {
  const root = container.getRootNode();
  const host: Node = root.nodeType === 9 ? (root as Document).head || (root as Document).documentElement : root;
  const entry = styleUsers.get(host);
  if (entry) {
    entry.count++;
    return host;
  }
  const element = document.createElement('style');
  element.setAttribute('data-bymoffice', '');
  element.textContent = STYLE_TEXT;
  host.appendChild(element);
  styleUsers.set(host, { element, count: 1 });
  return host;
}

function releaseStyle(host: Node): void {
  const entry = styleUsers.get(host);
  if (!entry || --entry.count > 0) return;
  entry.element.remove();
  styleUsers.delete(host);
}

const clamp = (value: number, min: number, max: number) => Math.min(Math.max(value, min), max);

/** Ctrl+S or Cmd+S: the key whoever holds the view saves by. */
const isSaveKey = (event: KeyboardEvent): boolean => (event.ctrlKey || event.metaKey) && !event.altKey && (event.key === 's' || event.key === 'S');

const coarsePointer = (): boolean => typeof window.matchMedia === 'function' && window.matchMedia('(pointer:coarse)').matches;

/** How often the pages are prepared for printing anew because the workbook changed meanwhile. */
const MAX_PRINT_RESTARTS = 3;
/** How long one page waits for its pictures and its conditional formats before it is printed as it is. */
const PRINT_PAGE_WAIT_MS = 30000;
/** A sheet of A4: stands for the pages while only the options of print() are checked. */
const PRINT_CHECK_PAGE = { width: 794, height: 1123 };

const sleep = (ms: number): Promise<void> => new Promise<void>((resolve) => setTimeout(resolve, ms));
/** How long printing waits for a calculation under way before it goes on with what is there. */
const PRINT_CALC_MS = 30000;

/** One call of print(): what it holds while it runs. */
interface PrintJob {
  promise: Promise<DocumentPrintResult>;
  frame?: PrintFrame;
  /** Decoded pictures of the pages that are being painted. */
  store?: ImageStore;
  /** Set once print() of the window was called. */
  dialog?: PrintDialog;
  progress?: { done: number; total: number };
}

/** Largest size of the box that shows a comment beside its cell, in pixels; longer text scrolls. */
/** Size of the field a note is typed in, in pixels; larger where a finger has to reach it. */
const NOTE_FIELD_WIDTH = 220;
const NOTE_FIELD_HEIGHT = 96;
const NOTE_FIELD_WIDTH_COARSE = 260;
const NOTE_FIELD_HEIGHT_COARSE = 132;
const COMMENT_MAX_WIDTH = 280;
const COMMENT_MAX_HEIGHT = 240;
/** Runs of a comment that keep their own weight in the box; the rest follows as plain text of the last one. */
const MAX_COMMENT_SPANS = 256;
/** Characters of a comment said after the cell it belongs to. */
const MAX_SPOKEN_COMMENT = 500;

/** A count read from the file, made safe to use as one. */
const count = (value: number) => (Number.isFinite(value) && value > 0 ? Math.floor(value) : 0);

function toIndex(value: number, limit: number): number {
  if (typeof value !== 'number' || Number.isNaN(value)) throw new RangeError('Invalid cell position');
  return clamp(Math.floor(value), 0, limit - 1);
}

const sameCell = (a: CellPos, b: CellPos) => a.row === b.row && a.col === b.col;
const sameRange = (a: CellRange, b: CellRange) => a.top === b.top && a.left === b.left && a.bottom === b.bottom && a.right === b.right;

const rangeBetween = (a: CellPos, b: CellPos): CellRange => ({
  top: Math.min(a.row, b.row),
  left: Math.min(a.col, b.col),
  bottom: Math.max(a.row, b.row),
  right: Math.max(a.col, b.col),
});

/** The nearest index after `index` in the direction of `step` that is not hidden, or -1. */
function stepVisible(axis: AxisMetrics, index: number, step: number): number {
  const i = axis.skipHidden(index + step, step > 0 ? 1 : -1);
  return i >= 0 && i < axis.count ? i : -1;
}

function buildMergeIndex(source: CellRange[]): MergeIndex {
  const byTop = source.slice().sort((a, b) => a.top - b.top);
  const maxBottom: number[] = [];
  let max = -1;
  for (let i = 0; i < byTop.length; i++) {
    if (byTop[i].bottom > max) max = byTop[i].bottom;
    maxBottom.push(max);
  }
  return { source, size: source.length, byTop, maxBottom };
}

/** Index of the first range in `byTop` order that can reach `row` or a row after it. */
function firstReaching(index: MergeIndex, row: number): number {
  let low = 0;
  let high = index.maxBottom.length;
  while (low < high) {
    const mid = (low + high) >> 1;
    if (index.maxBottom[mid] < row) low = mid + 1;
    else high = mid;
  }
  return low;
}

/** The merged range that contains a cell, if any. */
function mergeIn(index: MergeIndex, row: number, col: number): CellRange | undefined {
  const { byTop } = index;
  for (let i = firstReaching(index, row); i < byTop.length && byTop[i].top <= row; i++) {
    const m = byTop[i];
    if (m.bottom >= row && m.left <= col && m.right >= col) return m;
  }
  return undefined;
}

/** Takes a merged range into `out` if it crosses its border. Returns whether `out` grew. */
function absorb(out: CellRange, m: CellRange): boolean {
  if (m.top > out.bottom || m.bottom < out.top || m.left > out.right || m.right < out.left) return false;
  if (m.top >= out.top && m.bottom <= out.bottom && m.left >= out.left && m.right <= out.right) return false;
  out.top = Math.min(out.top, m.top);
  out.left = Math.min(out.left, m.left);
  out.bottom = Math.max(out.bottom, m.bottom);
  out.right = Math.max(out.right, m.right);
  return true;
}

/** Grows a range until no merged range crosses its border. */
function expandByMerges(index: MergeIndex, range: CellRange): CellRange {
  const { byTop, maxBottom } = index;
  if (byTop.length === 0) return range;
  const out = { ...range };
  // A pass that grows the range takes in at least one more merged range, so
  // this ends after at most one pass per merged range. Chains that lead down
  // are followed by the forward sweep and chains that lead up by the backward
  // one, each in a single pass.
  for (let pass = 0; pass <= byTop.length; pass++) {
    let grew = false;
    let i = firstReaching(index, out.top);
    for (; i < byTop.length && byTop[i].top <= out.bottom; i++) if (absorb(out, byTop[i])) grew = true;
    for (i--; i >= 0 && maxBottom[i] >= out.top; i--) if (absorb(out, byTop[i])) grew = true;
    if (!grew) break;
  }
  return out;
}

/**
 * Puts text in the place of a part of a field, the caret behind it. As typing would where the
 * browser lets a script do that: the field's own undo (Ctrl+Z) then takes the text back in one step.
 */
function replaceFieldText(field: HTMLTextAreaElement, start: number, end: number, text: string): void {
  const before = field.value;
  const wanted = before.substring(0, start) + text + before.substring(end);
  const caret = start + text.length;
  const root = field.getRootNode() as Document | ShadowRoot;
  if (root.activeElement === field) {
    try {
      // Away and back: Chromium takes what a script inserts right after typing for part of that typing, and
      // ends the step where the selection moves; its undo then takes back the reference alone. Firefox does
      // so anyway. WebKit undoes everything typed in a row at once whatever is done here.
      field.setSelectionRange(0, 0);
      field.setSelectionRange(start, end);
      const doc = field.ownerDocument;
      if (doc && typeof doc.execCommand === 'function') doc.execCommand('insertText', false, text);
    } catch {
      // Not available here; the text is set below.
    }
  }
  if (field.value !== wanted) {
    if (field.value === before && typeof field.setRangeText === 'function') field.setRangeText(text, start, end, 'end');
    if (field.value !== wanted) field.value = wanted;
  }
  try {
    if (field.selectionStart !== caret || field.selectionEnd !== caret) field.setSelectionRange(caret, caret);
  } catch {
    // Only fails on an element that is not rendered; the caret position is cosmetic.
  }
}

const fullWidth = (range: CellRange) => range.left === 0 && range.right === MAX_COLS - 1;
const fullHeight = (range: CellRange) => range.top === 0 && range.bottom === MAX_ROWS - 1;

/** Largest native scroll offset used for a logical scroll range. */
function nativeRange(limit: number, client: number): number {
  return Math.max(limit <= MAX_SPACER - client ? limit : MAX_SPACER - client, 1);
}

function toNative(value: number, limit: number, range: number): number {
  if (limit <= range) return value;
  return value >= limit ? range : (value * range) / limit;
}

function fromNative(offset: number, limit: number, range: number): number {
  if (limit <= 0) return 0;
  if (limit <= range) return clamp(Math.round(offset), 0, limit);
  if (offset >= range - 0.5) return limit;
  return clamp(Math.round((offset * limit) / range), 0, limit);
}

function resolveLocale(locale: string | FormatLocale | undefined): FormatLocale {
  // A locale object from the host page may be incomplete; missing parts fall back to en-US.
  if (typeof locale === 'object' && locale !== null) return normalizeLocale(locale);
  return getLocale(locale !== undefined ? locale : typeof navigator !== 'undefined' ? navigator.language : undefined);
}


/**
 * Whether a copy of cells taken when the sheets had the names `namesThen` (folded, see
 * foldSheetName()) no longer fits the workbook: a sheet of then is gone or has another name,
 * so a formula of the copy may name what is not there any more. Sheets that were added, moved,
 * hidden or coloured since change nothing the copy says. True for a copy nothing is known of.
 */
export function copyOutdated(namesThen: readonly string[] | undefined, sheets: readonly { name: string }[]): boolean {
  if (!namesThen) return true;
  const now = new Set<string>();
  for (const sheet of sheets) now.add(foldSheetName(sheet.name));
  for (const name of namesThen) if (!now.has(name)) return true;
  return false;
}

/**
 * Column width in pixels -> the width in characters a file stores for it: the
 * inverse of columnWidthToPixels(), in the 1/256 steps files use.
 */
function pixelsToColumnWidth(pixels: number): number {
  if (!(pixels > 0)) return 0;
  const digit = DEFAULT_DIGIT_WIDTH;
  // Aimed at the middle of the pixel, so that rounding to 1/256 cannot leave it.
  return Math.max(Math.round(((pixels + 0.5) * 256) / digit - Math.floor(128 / digit)), 0) / 256;
}

export class WorkbookView {
  readonly workbook: Workbook;
  /** Undo and redo. Each step carries the sheet and selection to go back to. */
  readonly history: History<HistoryMark>;
  private locale: FormatLocale;
  private engine: CalcEngine | undefined;
  private calcTimer: ReturnType<typeof setTimeout> | undefined;
  private missingTimer: ReturnType<typeof setTimeout> | undefined;
  private noticeTimer: ReturnType<typeof setTimeout> | undefined;
  /** Conditional formats of the sheets that have any, made when a sheet is first drawn. */
  private conditional = new Map<Sheet, ConditionalFormats>();
  private strings: UiStrings;
  private locked: boolean;

  private readonly container: HTMLElement;
  private readonly options: WorkbookViewOptions;
  private readonly root: HTMLDivElement;
  private readonly scroller: HTMLDivElement;
  private readonly canvas: HTMLCanvasElement;
  private readonly spacer: HTMLDivElement;
  private readonly tabs: HTMLDivElement;
  /** The button beside the tabs that adds a sheet. */
  private readonly addButton: HTMLButtonElement;
  private readonly note: HTMLDivElement;
  private readonly live: HTMLDivElement;
  /**
   * The one element that takes the keyboard focus and all text input. While no
   * cell is edited it is a transparent pixel over the active cell (so that an
   * input method opens its window there); while one is, it is the editor.
   */
  private readonly input: HTMLTextAreaElement;
  private readonly styleHost: Node;
  private readonly pointerEvents: boolean;

  /**
   * What is remembered for each sheet, by the sheet itself: indexes move when a sheet is
   * inserted or deleted, and a deleted sheet that undo brings back is the same object.
   */
  private readonly states = new Map<Sheet, SheetState>();
  private sheetIndex: number;
  /** The sheet at `sheetIndex` when it was last shown: tells a change of the list of sheets under the view. */
  private shown: Sheet;
  /** The field a sheet is being renamed in, in the place of its tab. */
  private renaming: { sheet: Sheet; tab: HTMLElement; input: HTMLInputElement; box: HTMLElement; busy: boolean } | undefined;
  /** The sheet whose tab was pressed while a name was being typed: shown once the field has closed. */
  private renameTarget: Sheet | undefined;
  /** A finger resting on a tab: a long press opens the menu of the tab. */
  private tabPress: { pointerId: number; x: number; y: number; sheet: Sheet; scroll: number } | undefined;
  private tabPressTimer = 0;
  private lastTabTap: { time: number; sheet: Sheet } | undefined;
  /** The click that follows the long press that opened the menu of a tab is not a tap on the tab. */
  private swallowTabClick = false;
  /**
   * A tab that may be dragged to another place: pressed with the mouse, or held with a finger
   * until the long press was seen. `started` once the pointer went far enough; `gap` is the
   * gap between the tabs it would be dropped in.
   */
  private tabDrag: { pointerId: number; sheet: Sheet; x: number; touch: boolean; started: boolean; gap: number; lastX: number } | undefined;
  /** The line that shows where a dragged tab would go. */
  private tabDropMark: HTMLElement | undefined;
  private tabDragTimer = 0;
  /** The window is listened to for the end of the pointer of `tabDrag` (see watchTabDrag()). */
  private tabDragWatched = false;
  /** The palette of the colour of a tab. */
  private tabPopover: Popover | undefined;

  /** Canvas size in CSS pixels and the ratio of its backing store to it. */
  private width = 0;
  private height = 0;
  private dpr = 1;

  /** Logical scroll limits of the shown sheet and the native ranges they map onto. */
  private limitX = 0;
  private limitY = 0;
  private rangeX = 1;
  private rangeY = 1;
  /** Native scroll offsets as last written or read; a scroll event that reports these is an echo. */
  private domX = 0;
  private domY = 0;

  private frame = 0;
  private autoScrollFrame = 0;
  private drag: Drag | undefined;
  private touch: Touch | undefined;
  private edit: Edit | undefined;
  /** The reference last pointed at for the formula being typed; the next one pointed at takes its place. */
  private point: Point | undefined;
  /** Set while a selection command runs for pointing: it moves the cells pointed at, not the selection. */
  private pointing = false;
  /** Set while the view itself writes a reference into the text, which is not the user typing. */
  private pointWriting = false;
  /** An input method is composing in the editor or in a field of the shell. */
  private composing = false;
  private menu: ContextMenu | undefined;
  /** What dropped down from a cell: the menu of a filter button or the list of a data validation. */
  private panel: FilterMenu | ListMenu | undefined;
  /** What the grid draws for each filter; a filter is replaced, never changed, so one entry serves until then. */
  private filterViews = new WeakMap<AutoFilter, GridFilter>();
  /** Data validations of the sheets that have any, made when first asked for. */
  private validation = new Map<Sheet, DataValidations>();
  private cellInfo: CellInfo | undefined;
  /** The message of the active cell's data validation, shown beside the cell. */
  private tip: HTMLDivElement | undefined;
  /** Pictures, shapes and comments of the sheets that have any, made when a sheet is first drawn. */
  private drawings = new Map<Sheet, SheetDrawings>();
  /** Decoded pictures of the workbook; made when a sheet with a picture is first drawn. */
  private images: ImageStore | undefined;
  /** The comment of a cell, shown beside it while the cell is pointed at or active. */
  private commentBox: HTMLDivElement | undefined;
  /** The cell whose comment the box shows now. */
  private commentShown: { sheet: Sheet; row: number; col: number; comment: SheetComment } | undefined;
  /** The cell the mouse pointer is over, while it carries a comment. */
  private hoverComment: CellPos | undefined;
  /** The note being typed: a field in the place of the box that shows the note of its cell. */
  private noteEdit: { sheet: Sheet; row: number; col: number; box: HTMLDivElement; input: HTMLTextAreaElement; original: string; busy: boolean } | undefined;
  /** The dialog that is open over the view (data validation, conditional formats), with the sheet it is about. */
  /**
   * The dialog of rules that is open: `changed` says whether the user changed anything in it since it was
   * opened (for the list of rules: typed cells for a rule that are not applied yet), `say` shows a message in it.
   */
  private rulesDialog: { sheet: Sheet; submit(): void; close(): void; changed(): boolean; say(message: string): void } | undefined;
  /** Set while a dialog of rules applies its own change: the one change that does not have to settle the dialog first. */
  private applyingRules = false;
  /** The names (folded) the sheets had when the selection was last copied: see copyOutdated(). */
  private copiedNames: string[] | undefined;
  /** Set while a dialog is closed by the view itself: it then does not hand over to the dialog it was opened from. */
  private leavingDialog = false;
  /** Sheets about which it was said that some of their drawings or comments are not shown. */
  private warned = new WeakSet<Sheet>();
  /** The comment of this cell was put away with Escape; it stays away until another cell is active. */
  private commentDismissed: { sheet: Sheet; row: number; col: number } | undefined;
  /**
   * Formula cells the engine has something to say about, per sheet, by row * MAX_COLS + column:
   * "file" (not calculated here, the value of the file is shown), "circular", "stale".
   * Built again after a recalculation, when it is next asked for.
   */
  private marks: (Map<number, string> | undefined)[] | undefined;
  private marksDirty = true;
  /** What copy() last put on the clipboard; pasting the same text back pastes these cells. */
  private copied: CopiedRange | undefined;
  private longPressTimer = 0;
  /** The cell the last tap selected, for telling a double tap: the second tap has to select the same cell. */
  private lastTap: { time: number; sheet: Sheet; row: number; col: number } | undefined;
  /** Names the last copy on the clipboard, so that a paste can tell this view's own cells from any other text. */
  private copyId = '';
  private committing = false;
  /** Set by the copy and cut handlers, so that the menu can tell whether execCommand reached them. */
  private clipboardHandled = false;
  private guide: HTMLDivElement | undefined;
  private measureCanvas: CanvasRenderingContext2D | null | undefined;
  /** The last pointer was a finger: editing has to ask for the on-screen keyboard. */
  private touchInput = false;
  private keyboardShown = false;
  /** Set around focus calls of the view's own, which must not count as the user arriving or leaving. */
  private quietFocus = false;
  private refocusing = false;
  private resizeObserver: ResizeObserver | undefined;
  private unwatchFonts: (() => void) | undefined;
  private unwatchEmbeddedFont: (() => void) | undefined;
  private dprQuery: MediaQueryList | undefined;
  private destroyed = false;
  private printJob: PrintJob | undefined;
  /** Counts the changes of the workbook and of how it is shown: pages prepared for printing before are out of date. */
  private printEpoch = 0;
  /** Language tag the paper of a sheet without one is chosen by. */
  private languageTag: string | undefined;

  constructor(container: HTMLElement, workbook: Workbook, options: WorkbookViewOptions = {}) {
    if (workbook.sheets.length === 0) throw new Error('The workbook has no sheets');
    this.container = container;
    this.workbook = workbook;
    this.options = options;
    this.locale = resolveLocale(options.locale);
    this.strings = resolveStrings(options.locale);
    this.languageTag = typeof options.locale === 'string' ? options.locale : typeof navigator !== 'undefined' ? navigator.language : undefined;
    this.locked = options.readOnly === true;
    this.history = new History<HistoryMark>(workbook);
    this.sheetIndex = this.initialSheet();
    this.shown = workbook.sheets[this.sheetIndex];
    this.pointerEvents = 'PointerEvent' in window;

    const root = document.createElement('div');
    root.className = 'bym-view';
    root.setAttribute('role', 'group');
    root.setAttribute('aria-label', this.strings.spreadsheet);
    // The scroll mapping assumes offsets that grow to the right, whatever the page direction is.
    root.style.cssText = 'position:relative;display:flex;flex-direction:column;width:100%;height:100%;box-sizing:border-box;overflow:hidden;direction:ltr';

    const scroller = document.createElement('div');
    scroller.className = 'bym-scroller';
    // The focus lives in the input element; some browsers would make a scrolling element a Tab stop of its own.
    scroller.tabIndex = -1;
    scroller.setAttribute('role', 'grid');
    scroller.style.cssText = 'position:relative;flex:1 1 0;min-height:0;overflow:auto;overscroll-behavior:contain';

    const canvas = document.createElement('canvas');
    canvas.className = 'bym-canvas';
    canvas.style.cssText =
      'display:block;position:-webkit-sticky;top:0;left:0;float:none;box-sizing:content-box;margin:0;padding:0;border:0;' +
      'min-width:0;min-height:0;max-width:none;max-height:none';
    canvas.style.position = 'sticky';

    const spacer = document.createElement('div');
    spacer.className = 'bym-spacer';
    spacer.style.cssText = 'position:absolute;top:0;left:0;width:1px;height:1px;visibility:hidden;pointer-events:none';

    const note = document.createElement('div');
    note.className = 'bym-note';
    note.style.cssText = 'position:absolute;left:0;right:0;top:45%;padding:0 16px;text-align:center;pointer-events:none;display:none';

    const tabs = document.createElement('div');
    tabs.className = 'bym-tabs';
    tabs.setAttribute('role', 'tablist');
    tabs.style.cssText = 'position:relative;display:flex;flex:none;overflow-x:auto;overflow-y:hidden';

    // Beside the tabs, not among them: the tab list holds tabs only, and stays in place while they scroll.
    const addButton = document.createElement('button');
    addButton.type = 'button';
    addButton.className = 'bym-tab-add';
    addButton.textContent = '+';
    addButton.style.cssText = 'position:absolute;z-index:1;left:0;bottom:0';

    const live = document.createElement('div');
    live.className = 'bym-live';
    live.setAttribute('aria-live', 'polite');
    live.setAttribute('aria-atomic', 'true');
    live.style.cssText = 'position:absolute;width:1px;height:1px;margin:-1px;padding:0;overflow:hidden;clip:rect(0 0 0 0);white-space:nowrap;border:0';

    const input = document.createElement('textarea');
    input.className = 'bym-input';
    input.setAttribute('aria-label', this.strings.cellEditor);
    // No on-screen keyboard until a cell is edited.
    input.setAttribute('inputmode', 'none');
    input.setAttribute('autocomplete', 'off');
    input.setAttribute('autocorrect', 'off');
    input.setAttribute('autocapitalize', 'off');
    input.setAttribute('wrap', 'off');
    input.spellcheck = false;
    input.rows = 1;
    input.readOnly = this.locked;
    input.style.cssText =
      'position:absolute;z-index:2;left:0;top:0;display:block;float:none;width:1px;height:1px;min-width:0;min-height:0;max-width:none;max-height:none;' +
      'margin:0;padding:0;border:0;outline:none;resize:none;overflow:hidden;-webkit-appearance:none;appearance:none;' +
      'box-sizing:border-box;opacity:0;pointer-events:none;white-space:pre;font:12px/1.2 sans-serif;border-radius:0';

    this.root = root;
    this.input = input;
    this.scroller = scroller;
    this.canvas = canvas;
    this.spacer = spacer;
    this.tabs = tabs;
    this.addButton = addButton;
    this.note = note;
    this.live = live;

    this.buildTabs();
    scroller.appendChild(canvas);
    scroller.appendChild(spacer);
    root.appendChild(scroller);
    root.appendChild(input);
    root.appendChild(note);
    root.appendChild(addButton);
    root.appendChild(tabs);
    root.appendChild(live);

    this.styleHost = acquireStyle(container);
    container.appendChild(root);
    this.placeAddButton();

    scroller.addEventListener('scroll', this.onScroll, { passive: true });
    input.addEventListener('keydown', this.onKeyDown);
    input.addEventListener('input', this.onInput);
    input.addEventListener('compositionstart', this.onCompositionStart);
    input.addEventListener('mousedown', this.onInputMouseDown);
    input.addEventListener('focus', this.onFocus);
    input.addEventListener('blur', this.onBlur);
    input.addEventListener('copy', this.onCopy);
    input.addEventListener('cut', this.onCut);
    input.addEventListener('paste', this.onPaste);
    canvas.addEventListener('dblclick', this.onDoubleClick);
    canvas.addEventListener('contextmenu', this.onContextMenu);
    // Panning and pinching stay with the browser; its double-tap zoom would swallow the double tap that edits.
    canvas.style.touchAction = 'manipulation';
    tabs.addEventListener('click', this.onTabClick);
    tabs.addEventListener('keydown', this.onTabKeyDown);
    tabs.addEventListener('contextmenu', this.onTabContextMenu);
    if (this.pointerEvents) {
      tabs.addEventListener('pointerdown', this.onTabPointerDown);
      tabs.addEventListener('pointermove', this.onTabPointerMove);
      tabs.addEventListener('pointerup', this.onTabPointerEnd);
      tabs.addEventListener('pointercancel', this.onTabPointerEnd);
      tabs.addEventListener('touchmove', this.onTabTouchMove, { passive: false });
    }
    addButton.addEventListener('click', this.onAddClick);
    root.addEventListener('mousedown', this.onRootMouseDown);
    root.addEventListener('focusout', this.onFocusOut);
    root.addEventListener('compositionstart', this.onComposition);
    root.addEventListener('compositionend', this.onComposition);
    if (options.focusScope) {
      options.focusScope.addEventListener('focusout', this.onFocusOut);
      options.focusScope.addEventListener('compositionstart', this.onComposition);
      options.focusScope.addEventListener('compositionend', this.onComposition);
    }
    // Not passive: a finger on the border between two headers resizes instead of scrolling.
    canvas.addEventListener('touchstart', this.onTouchStart, { passive: false });
    scroller.addEventListener('focus', this.onScrollerFocus);
    if (this.pointerEvents) {
      canvas.addEventListener('pointerdown', this.onPointerDown);
      canvas.addEventListener('pointermove', this.onPointerMove);
      canvas.addEventListener('pointerup', this.onPointerUp);
      canvas.addEventListener('pointercancel', this.onPointerCancel);
      canvas.addEventListener('lostpointercapture', this.onPointerCancel);
      // After a tap the browser also sends mouse events; their default would take the focus away from the input.
      canvas.addEventListener('mousedown', this.onMouseDown);
      canvas.addEventListener('pointerleave', this.onPointerLeave);
    } else {
      // Browsers without pointer events: a tap arrives as the mouse events the browser makes up for it.
      canvas.addEventListener('mousedown', this.onPointerDown);
      canvas.addEventListener('mousemove', this.onHoverMove);
      canvas.addEventListener('mouseleave', this.onPointerLeave);
    }
    window.addEventListener('resize', this.onResize);
    if (typeof ResizeObserver === 'function') {
      this.resizeObserver = new ResizeObserver(this.onResize);
      this.resizeObserver.observe(scroller);
    }
    this.watchPixelRatio();

    this.measure();
    this.showSheet();
    this.draw();
    this.calculateMissing();
    // A sheet that was asked for is the one a saved file opens on, like one the user chose. Without
    // the option the model is left as the file made it (a hidden active sheet, an index that is none).
    if (options.sheet !== undefined) this.noteActiveSheet();
    // A font the application registered or took out: the cells are painted anew (what was measured is forgotten by
    // the measuring itself). Last of all: a constructor that throws leaves no listener behind.
    this.unwatchFonts = onFontsChanged(() => this.schedule());
    // The font of the package was needed and could not be loaded: said once, by one view of the page.
    this.unwatchEmbeddedFont = onEmbeddedFontNotice((notice) => this.notify(this.strings.embeddedFontNotLoaded(notice.family), 'info'));
  }

  get activeSheet(): number {
    return this.sheetIndex;
  }

  /** Shows a sheet. Hidden sheets have no tab but can be shown this way. */
  setActiveSheet(index: number): void {
    if (!Number.isInteger(index) || index < 0 || index >= this.workbook.sheets.length) {
      throw new RangeError(`No sheet with index ${index}`);
    }
    if (this.destroyed) return;
    // What is being typed belongs to the sheet of its cell, and is written before another sheet is shown for
    // good - also when that is the sheet shown now, for a formula to point at.
    if (this.edit && (index !== this.sheetIndex || this.editAway())) this.commitEditing();
    if (index === this.sheetIndex) return;
    // So does a name that is being typed for it, and a note that is being typed on it.
    this.finishRename(true, true);
    this.finishNote(true, true);
    this.show(index);
  }

  /** Shows another sheet and leaves what is being typed alone: see setActiveSheet(). */
  private show(index: number): void {
    if (index === this.sheetIndex) return;
    this.closeMenu();
    this.endDrag();
    this.endTouch();
    this.lastTap = undefined;
    this.hoverComment = undefined;
    this.finishNote(true, true);
    this.sheetIndex = index;
    this.shown = this.workbook.sheets[index];
    this.showSheet();
    this.schedule();
    this.noteActiveSheet();
    if (this.options.onSheetChange) this.options.onSheetChange(index);
  }

  /**
   * Tells the model which sheet is shown, so that a saved file opens on it (the writer says so
   * only along with another change: see writeXlsx()). Not a sheet that is only looked at for a
   * formula to point at - the view goes back to the sheet of the cell when the input ends - and
   * not a hidden sheet, which setActiveSheet() can show: a file does not open on one.
   */
  private noteActiveSheet(): void {
    if (this.destroyed || this.editAway()) return;
    const sheet = this.workbook.sheets[this.sheetIndex];
    if (!sheet || sheet !== this.shown || sheet.state !== 'visible') return;
    if (this.workbook.activeSheet !== this.sheetIndex) this.workbook.activeSheet = this.sheetIndex;
  }

  get selection(): GridSelection {
    const state = this.state();
    return { active: { ...state.active }, range: { ...state.range } };
  }

  /** Selects a range and scrolls its active cell (the top-left one unless given) into view. */
  select(range: CellRange, active?: { row: number; col: number }): void {
    if (this.destroyed) return;
    // Like a click on another cell: what is being typed is written first.
    if (this.edit) this.commitEditing();
    const a = { row: toIndex(range.top, MAX_ROWS), col: toIndex(range.left, MAX_COLS) };
    const b = { row: toIndex(range.bottom, MAX_ROWS), col: toIndex(range.right, MAX_COLS) };
    // The same rule as for clicks and keys: a merged cell is selected whole, with its top-left cell active.
    const merges = this.merges();
    const next = expandByMerges(merges, rangeBetween(a, b));
    let cell = active
      ? { row: clamp(toIndex(active.row, MAX_ROWS), next.top, next.bottom), col: clamp(toIndex(active.col, MAX_COLS), next.left, next.right) }
      : { row: next.top, col: next.left };
    const merge = mergeIn(merges, cell.row, cell.col);
    if (merge) cell = { row: merge.top, col: merge.left };
    const focus = { row: cell.row === next.top ? next.bottom : next.top, col: cell.col === next.left ? next.right : next.left };
    this.setSelection(cell, next, focus, cell);
  }

  /** Scrolls a cell into view without changing the selection. */
  scrollTo(row: number, col: number): void {
    if (this.destroyed) return;
    this.reveal({ row: toIndex(row, MAX_ROWS), col: toIndex(col, MAX_COLS) });
  }

  /** Current scroll offsets in sheet pixels and the canvas size in CSS pixels (for tests and embedding). */
  getViewState(): { scrollX: number; scrollY: number; width: number; height: number; devicePixelRatio: number } {
    if (!this.destroyed) this.readScroll();
    const state = this.state();
    return { scrollX: state.scrollX, scrollY: state.scrollY, width: this.width, height: this.height, devicePixelRatio: this.dpr };
  }

  /**
   * What the shown sheet has floating above its cells (for tests and embedding): the number of
   * objects, the bytes of decoded pictures held right now, and the most the view ever holds.
   */
  getDrawingState(): { objects: number; imageBytes: number; imageBudget: number } {
    const layer = this.destroyed ? undefined : this.drawingsOf(this.state().sheet);
    return {
      objects: layer ? layer.rects(this.state().metrics).length : 0,
      imageBytes: this.images ? this.images.bytes : 0,
      imageBudget: this.images ? this.images.budget : 0,
    };
  }

  /** Changes the language of number and date formats, and of the view's own labels, and redraws. */
  setLocale(locale: string | FormatLocale): void {
    this.locale = resolveLocale(locale);
    // What is typed into a dialog is typed the way the locale writes it: it is not read by another one.
    if (this.rulesDialog) this.rulesDialog.close();
    this.validation.forEach((rules) => rules.setLocale(this.locale));
    // The view's own words follow a language tag; a locale object only changes formats.
    if (typeof locale === 'string') {
      this.strings = resolveStrings(locale);
      this.languageTag = locale;
    }
    this.printEpoch++;
    this.drawings.forEach((layer) => layer.setLabels(this.drawingLabels()));
    this.root.setAttribute('aria-label', this.strings.spreadsheet);
    this.input.setAttribute('aria-label', this.strings.cellEditor);
    this.placeAddButton();
    if (this.renaming) this.renaming.input.setAttribute('aria-label', this.strings.sheetName);
    this.showSheet();
    this.render();
    this.announce();
  }

  /** Forces a synchronous redraw. */
  render(): void {
    if (this.destroyed) return;
    this.measure();
    this.readScroll();
    this.growExtent(this.state().scrollX, this.state().scrollY);
    this.commitScroll();
    this.draw();
  }

  /** Removes every listener, observer, pending frame and the view's DOM. Safe to call twice. */
  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    if (this.unwatchFonts) this.unwatchFonts();
    this.unwatchFonts = undefined;
    if (this.unwatchEmbeddedFont) this.unwatchEmbeddedFont();
    this.unwatchEmbeddedFont = undefined;
    // A print job under way ends here: its pages leave the document at once.
    if (this.printJob) this.endPrint(this.printJob);
    if (this.calcTimer !== undefined) clearTimeout(this.calcTimer);
    if (this.missingTimer !== undefined) clearTimeout(this.missingTimer);
    if (this.noticeTimer !== undefined) clearTimeout(this.noticeTimer);
    this.endDrag();
    this.endTouch();
    this.clearTabPress();
    this.edit = undefined;
    this.point = undefined;
    this.renaming = undefined;
    this.closeNote(false);
    this.closeMenu();
    const input = this.input;
    input.removeEventListener('keydown', this.onKeyDown);
    input.removeEventListener('input', this.onInput);
    input.removeEventListener('compositionstart', this.onCompositionStart);
    input.removeEventListener('mousedown', this.onInputMouseDown);
    input.removeEventListener('focus', this.onFocus);
    input.removeEventListener('blur', this.onBlur);
    input.removeEventListener('copy', this.onCopy);
    input.removeEventListener('cut', this.onCut);
    input.removeEventListener('paste', this.onPaste);
    this.canvas.removeEventListener('dblclick', this.onDoubleClick);
    this.canvas.removeEventListener('contextmenu', this.onContextMenu);
    if (this.frame) cancelAnimationFrame(this.frame);
    this.frame = 0;
    if (this.resizeObserver) this.resizeObserver.disconnect();
    this.resizeObserver = undefined;
    this.unwatchPixelRatio();
    window.removeEventListener('resize', this.onResize);
    this.scroller.removeEventListener('scroll', this.onScroll);
    this.tabs.removeEventListener('click', this.onTabClick);
    this.tabs.removeEventListener('keydown', this.onTabKeyDown);
    this.tabs.removeEventListener('contextmenu', this.onTabContextMenu);
    this.tabs.removeEventListener('pointerdown', this.onTabPointerDown);
    this.tabs.removeEventListener('pointermove', this.onTabPointerMove);
    this.tabs.removeEventListener('pointerup', this.onTabPointerEnd);
    this.tabs.removeEventListener('pointercancel', this.onTabPointerEnd);
    this.tabs.removeEventListener('touchmove', this.onTabTouchMove);
    this.endTabDrag();
    if (this.tabPopover) this.tabPopover.close('owner');
    this.addButton.removeEventListener('click', this.onAddClick);
    this.root.removeEventListener('mousedown', this.onRootMouseDown);
    this.root.removeEventListener('focusout', this.onFocusOut);
    this.root.removeEventListener('compositionstart', this.onComposition);
    this.root.removeEventListener('compositionend', this.onComposition);
    if (this.options.focusScope) {
      this.options.focusScope.removeEventListener('focusout', this.onFocusOut);
      this.options.focusScope.removeEventListener('compositionstart', this.onComposition);
      this.options.focusScope.removeEventListener('compositionend', this.onComposition);
    }
    this.canvas.removeEventListener('touchstart', this.onTouchStart);
    this.scroller.removeEventListener('focus', this.onScrollerFocus);
    const canvas = this.canvas;
    canvas.removeEventListener('pointerdown', this.onPointerDown);
    canvas.removeEventListener('pointermove', this.onPointerMove);
    canvas.removeEventListener('pointerup', this.onPointerUp);
    canvas.removeEventListener('pointercancel', this.onPointerCancel);
    canvas.removeEventListener('lostpointercapture', this.onPointerCancel);
    canvas.removeEventListener('mousedown', this.onPointerDown);
    canvas.removeEventListener('mousedown', this.onMouseDown);
    canvas.removeEventListener('pointerleave', this.onPointerLeave);
    canvas.removeEventListener('mouseleave', this.onPointerLeave);
    if (!this.pointerEvents) canvas.removeEventListener('mousemove', this.onHoverMove);
    if (this.commentBox) this.commentBox.removeEventListener('mouseleave', this.onCommentLeave);
    this.commentBox = undefined;
    this.commentShown = undefined;
    // Decoded pictures are given back at once.
    if (this.images) this.images.destroy();
    this.images = undefined;
    this.drawings.clear();
    if (this.root.parentNode) this.root.parentNode.removeChild(this.root);
    releaseStyle(this.styleHost);
    // Give the backing store back right away; some browsers keep it until collection otherwise.
    canvas.width = 0;
    canvas.height = 0;
  }

  // -------------------------------------------------------------- printing

  /** Whether print() is at work: its pages are being prepared or its dialog is open. */
  get printing(): boolean {
    return this.printJob !== undefined;
  }

  /**
   * Prints the sheet that is shown, or with `sheets: 'all'` every visible
   * worksheet: lays the cells out on pages by the page setup of each sheet
   * (paper, margins, scale or fit to pages, print area and titles, page
   * breaks, header and footer), paints every page at the resolution of a
   * printer, and opens the print dialog of the window the view is in. The
   * pages are prepared one per turn, so the page stays responsive; when the
   * workbook is changed meanwhile, they are prepared anew. What belongs to
   * the screen is not printed: the selection, filter buttons, comments.
   *
   * Resolves when the dialog was closed (whether the user printed or
   * cancelled cannot be told). A call while the pages of an earlier one are
   * being prepared gets the promise of that call and opens no second dialog.
   * Rejects with a DocumentPrintError: 'unavailable' where the window cannot
   * print, 'empty' when there is nothing to print, 'too-long' beyond the
   * limits, 'cancelled' when the view is destroyed meanwhile; with a
   * RangeError for a `dpi` that is no positive number or a `sheets` that is
   * neither 'active' nor 'all'.
   */
  print(options: SheetPrintOptions = {}): Promise<DocumentPrintResult> {
    if (this.destroyed) return Promise.reject(new DocumentPrintError('cancelled', 'The view was destroyed'));
    // A wrong option is refused before any work is done, whatever else is under way.
    try {
      planPrint([PRINT_CHECK_PAGE], options);
      if (options.sheets !== undefined && options.sheets !== 'active' && options.sheets !== 'all') throw new RangeError('"sheets" must be "active" or "all"');
    } catch (error) {
      return Promise.reject(error);
    }
    const previous = this.printJob;
    if (previous) {
      if (!previous.dialog) return previous.promise;
      // print() came back and nothing said that the dialog was closed: a new call says so.
      previous.dialog.finish();
    }
    const job = {} as PrintJob;
    const nothing = () => undefined;
    const turn: Promise<unknown> = previous ? previous.promise.then(nothing, nothing) : Promise.resolve();
    job.promise = turn
      .then(() => this.runPrint(job, options))
      .then(
        (result) => {
          this.endPrint(job);
          return result;
        },
        (error) => {
          this.endPrint(job);
          this.printFailed(error);
          throw error;
        },
      );
    this.printJob = job;
    return job.promise;
  }

  private setPrintProgress(job: PrintJob, progress: { done: number; total: number } | undefined): void {
    job.progress = progress;
    if (this.printJob === job && !this.destroyed && this.options.onPrintProgress) this.options.onPrintProgress(progress ? { done: progress.done, total: progress.total } : undefined);
  }

  private async runPrint(job: PrintJob, options: SheetPrintOptions): Promise<DocumentPrintResult> {
    const cancelled = () => new DocumentPrintError('cancelled', 'Printing was cancelled: the view was destroyed');
    if (this.destroyed) throw cancelled();
    // The window of the document the view is in: inside a frame, the frame is what is printed.
    const doc = this.container.ownerDocument || document;
    const win = doc.defaultView;
    if (!win || typeof win.print !== 'function') throw new DocumentPrintError('unavailable', 'Printing is not available here: the window has no print()');
    // What is being typed is part of what is printed, as leaving the cell (the note, the tab) would make it.
    // A note or a name that cannot be taken was said, and its field is open: nothing is printed around it.
    if (!this.commitPending()) throw new DocumentPrintError('cancelled', 'Printing was cancelled: what is being typed cannot be applied');
    this.notify(this.strings.printPreparing, 'info');
    this.setPrintProgress(job, { done: 0, total: 0 });
    const limits = printLimits(coarsePointer());
    const urls = (win as unknown as { URL?: typeof URL }).URL || URL;
    let plan: PrintPlan;
    for (let attempt = 0; ; attempt++) {
      // A calculation still under way writes cells turn after turn; the pages wait for its end
      // rather than start anew at each of its turns. One that does not end is printed as it stands.
      const waitUntil = Date.now() + PRINT_CALC_MS;
      while (this.engine && this.engine.incomplete && Date.now() < waitUntil) {
        await sleep(30);
        if (this.destroyed) throw new DocumentPrintError('cancelled', 'Printing was cancelled: the view was destroyed');
      }
      const epoch = this.printEpoch;
      const media = this.workbook.media;
      // Pictures of its own: those of the view are sized for the screen and come and go with it.
      const store = media && typeof media.readPart === 'function' ? new ImageStore(media, coarsePointer() ? IMAGE_BUDGET_BYTES_COARSE : IMAGE_BUDGET_BYTES, () => undefined) : undefined;
      job.store = store;
      const layers = new Map<Sheet, SheetDrawings | undefined>();
      const metricsOf = new Map<Sheet, SheetMetrics>();
      const layerOf = (sheet: Sheet): SheetDrawings | undefined => {
        if (!layers.has(sheet)) {
          // Comments are not printed, so their boxes are left out of the layer.
          layers.set(sheet, sheet.drawings && sheet.drawings.length > 0 ? new SheetDrawings(sheet, { images: store, labels: this.drawingLabels(), notes: false }) : undefined);
        }
        return layers.get(sheet);
      };
      const sheets: number[] = [];
      if (options.sheets === 'all') {
        for (let i = 0; i < this.workbook.sheets.length; i++) if (this.workbook.sheets[i].state === 'visible') sheets.push(i);
      } else sheets.push(this.sheetIndex);
      const pages: SheetJobPlan = planSheetPages(this.workbook, sheets, {
        defaultPaper: defaultPaperFor(this.languageTag),
        maxPages: MAX_PRINT_PAGES,
        metrics: (sheet) => {
          let metrics = metricsOf.get(sheet);
          if (!metrics) {
            metrics = createSheetMetrics(sheet);
            metricsOf.set(sheet, metrics);
          }
          return metrics;
        },
        drawings: (sheet, _index, metrics) => {
          const layer = layerOf(sheet);
          return layer ? layer.rects(metrics) : undefined;
        },
      });
      plan = planPrint(
        pages.pages.map((entry) => entry.layout.paper),
        options,
        limits,
      );
      const frame = new PrintFrame(doc, plan, supportsNamedPages(win as unknown as { CSS?: { supports?(property: string, value: string): boolean } }), (url) => revokePicture(url, urls));
      job.frame = frame;
      this.setPrintProgress(job, { done: 0, total: plan.pages.length });
      const complete = await this.paintForPrint(job, epoch, pages, plan, frame, store, limits, doc, urls, options, (sheet) => metricsOf.get(sheet) as SheetMetrics, layerOf);
      if (store) store.destroy();
      job.store = undefined;
      if (this.destroyed) throw cancelled();
      if (complete) break;
      // The workbook changed under the pages: they are prepared again from the start.
      frame.dispose();
      job.frame = undefined;
      if (attempt >= MAX_PRINT_RESTARTS) throw new DocumentPrintError('failed', 'The workbook kept changing while it was prepared for printing');
    }
    this.setPrintProgress(job, undefined);
    const dialog = openPrintDialog(win as unknown as Parameters<typeof openPrintDialog>[0]);
    job.dialog = dialog;
    await dialog.closed;
    // Takes the place of "preparing": whether anything was printed cannot be told.
    if (!this.destroyed) this.notify(this.strings.printClosed, 'info');
    return { pages: plan.pages.length, dpi: plan.dpi };
  }

  /** Paints the pages of a plan and adds them to the frame; false when the workbook was changed meanwhile. */
  private async paintForPrint(
    job: PrintJob,
    epoch: number,
    pages: SheetJobPlan,
    plan: PrintPlan,
    frame: PrintFrame,
    store: ImageStore | undefined,
    limits: PrintLimits,
    doc: Document,
    urls: { createObjectURL(blob: Blob): string; revokeObjectURL(url: string): void },
    options: SheetPrintOptions,
    metricsOf: (sheet: Sheet) => SheetMetrics,
    layerOf: (sheet: Sheet) => SheetDrawings | undefined,
  ): Promise<boolean> {
    const stale = () => this.destroyed || this.printEpoch !== epoch;
    // One canvas for all pages: a page is held as pixels only while it is painted.
    const canvas = doc.createElement('canvas');
    const win = doc.defaultView;
    const nextFrame = win && typeof win.requestAnimationFrame === 'function' ? (callback: () => void) => win.requestAnimationFrame(callback) : undefined;
    const now = new Date();
    let bytes = 0;
    try {
      for (let i = 0; i < plan.pages.length; i++) {
        if (stale()) return false;
        const entry = pages.pages[i];
        const planned = plan.pages[i];
        const sheet = this.workbook.sheets[entry.sheet];
        if (!sheet) return false;
        if (canvas.width !== planned.pixelWidth) canvas.width = planned.pixelWidth;
        if (canvas.height !== planned.pixelHeight) canvas.height = planned.pixelHeight;
        const context = canvas.getContext('2d');
        if (!context) throw new DocumentPrintError('failed', 'No canvas to paint the pages on');
        const conditional = this.conditionalOf(sheet);
        const paint = () =>
          paintSheetPage(context, {
            workbook: this.workbook,
            job: entry,
            total: pages.pages.length,
            metrics: metricsOf(sheet),
            pixelWidth: planned.pixelWidth,
            locale: this.locale,
            conditional,
            drawings: layerOf(sheet),
            fileName: options.fileName,
            now,
          });
        // Pictures of the pages before this one may make room for those of this one.
        if (store) store.pass++;
        try {
          paint();
          // Painting asked for the pictures and started the rules whose formulas take long; once both are there the page is painted with them.
          const deadline = Date.now() + PRINT_PAGE_WAIT_MS;
          while (((store && store.busy) || (conditional && conditional.pending)) && Date.now() < deadline) {
            await sleep(20);
            if (stale()) return false;
            paint();
          }
        } catch (error) {
          if (error instanceof DocumentPrintError) throw error;
          throw new DocumentPrintError('failed', `Page ${i + 1} could not be painted`);
        }
        const picture = await canvasPicture(canvas, urls, nextFrame);
        bytes += picture.bytes;
        if (stale() || bytes > limits.maxBytes) {
          revokePicture(picture.url, urls);
          if (this.destroyed || !(bytes > limits.maxBytes)) return false;
          throw new DocumentPrintError('too-long', 'The pages of the workbook take too much memory to print at once');
        }
        await frame.add(picture.url, this.strings.page(i + 1));
        this.setPrintProgress(job, { done: i + 1, total: plan.pages.length });
      }
      return !stale();
    } finally {
      // Give the backing store back right away.
      canvas.width = 0;
      canvas.height = 0;
    }
  }

  /** Takes what a print job holds out of the document. Safe to call twice. */
  private endPrint(job: PrintJob): void {
    if (job.frame) job.frame.dispose();
    if (job.store) job.store.destroy();
    if (job.dialog) job.dialog.finish();
    job.frame = undefined;
    job.store = undefined;
    const reported = job.progress !== undefined;
    job.progress = undefined;
    if (this.printJob !== job) return;
    if (reported && !this.destroyed && this.options.onPrintProgress) this.options.onPrintProgress(undefined);
    this.printJob = undefined;
  }

  private printFailed(error: unknown): void {
    // A wrong option is the caller's to hear about, not the user's; so is the end of a view.
    if (this.destroyed || error instanceof RangeError) return;
    const code = error instanceof DocumentPrintError ? error.code : 'failed';
    if (code === 'cancelled') return;
    const strings = this.strings;
    this.notify(
      code === 'unavailable' ? strings.printNotAvailable : code === 'too-long' ? strings.printSheetTooLong : code === 'empty' ? strings.printNothing : strings.printSheetFailed,
      'error',
    );
  }

  // ---------------------------------------------------------------- sheets

  private initialSheet(): number {
    const sheets = this.workbook.sheets;
    const wanted = this.options.sheet;
    if (wanted !== undefined) {
      if (!Number.isInteger(wanted) || wanted < 0 || wanted >= sheets.length) throw new RangeError(`No sheet with index ${wanted}`);
      return wanted;
    }
    const active = this.workbook.activeSheet;
    if (Number.isInteger(active) && active >= 0 && active < sheets.length && sheets[active].state === 'visible') return active;
    for (let i = 0; i < sheets.length; i++) if (sheets[i].state === 'visible') return i;
    return 0;
  }

  private state(): SheetState {
    return this.stateOf(this.workbook.sheets[this.sheetIndex] || this.shown);
  }

  /** What is remembered for a sheet, made when it is first asked for. */
  private stateOf(sheet: Sheet): SheetState {
    let state = this.states.get(sheet);
    if (!state) {
      state = this.createState(sheet);
      this.states.set(sheet, state);
    }
    return state;
  }

  private createState(sheet: Sheet): SheetState {
    const metrics = createSheetMetrics(sheet);
    const view = sheet.view;
    const cell = (view.activeCell !== undefined && parseCellRef(view.activeCell)) || { row: 0, col: 0 };
    const merges = buildMergeIndex(sheet.merges);
    const merge = mergeIn(merges, cell.row, cell.col);
    const active = merge ? { row: merge.top, col: merge.left } : cell;

    // Open where the file was left: the first cell of the scrolling pane.
    const frozenRows = Math.min(count(view.frozenRows), metrics.rows.count - 1);
    const frozenCols = Math.min(count(view.frozenCols), metrics.cols.count - 1);
    const firstRef = frozenRows > 0 || frozenCols > 0 ? view.topLeftCell : view.firstVisibleCell;
    const first = firstRef !== undefined ? parseCellRef(firstRef) : undefined;

    return {
      sheet,
      metrics,
      scrollX: first ? Math.max(metrics.cols.offset(first.col) - metrics.cols.offset(frozenCols), 0) : 0,
      scrollY: first ? Math.max(metrics.rows.offset(first.row) - metrics.rows.offset(frozenRows), 0) : 0,
      extentRows: 0,
      extentCols: 0,
      active,
      focus: active,
      range: merge ? { ...merge } : { top: active.row, left: active.col, bottom: active.row, right: active.col },
      merges,
    };
  }

  /** The merged ranges of the shown sheet, indexed again when the model's list changed. */
  private merges(): MergeIndex {
    const state = this.state();
    const source = state.sheet.merges;
    if (state.merges.source !== source || state.merges.size !== source.length) state.merges = buildMergeIndex(source);
    return state.merges;
  }

  /** Brings tabs, note, scroll area and announcement in line with the shown sheet. */
  private showSheet(): void {
    const state = this.state();
    const tabs = this.tabs.children;
    let selected: HTMLElement | undefined;
    for (let i = 0; i < tabs.length; i++) {
      const tab = tabs[i] as HTMLElement;
      // The field a sheet is renamed in stands among the tabs and is none.
      if (!tab.hasAttribute('data-sheet')) continue;
      const on = tab.getAttribute('data-sheet') === String(this.sheetIndex);
      tab.setAttribute('aria-selected', on ? 'true' : 'false');
      if (on) selected = tab;
    }
    if (selected) {
      const left = selected.offsetLeft;
      const right = left + selected.offsetWidth;
      if (left < this.tabs.scrollLeft) this.tabs.scrollLeft = left;
      else if (right > this.tabs.scrollLeft + this.tabs.clientWidth) this.tabs.scrollLeft = right - this.tabs.clientWidth;
    }

    this.scroller.setAttribute('aria-label', this.strings.sheet(state.sheet.name));
    const kind = state.sheet.kind;
    this.note.textContent = kind === 'worksheet' ? '' : kind === 'chartsheet' ? this.strings.chartSheetNote : this.strings.otherSheetNote;
    this.note.style.display = kind === 'worksheet' ? 'none' : 'block';

    this.growExtent(state.scrollX, state.scrollY);
    this.commitScroll();
    this.announce();
    // Said once for a sheet: something of its drawings or comments could not be read and is missing from what is shown.
    if (state.sheet.warnings && state.sheet.warnings.length > 0 && !this.warned.has(state.sheet)) {
      this.warned.add(state.sheet);
      this.notifyLater(this.strings.drawingsNotShown);
    }
  }

  private buildTabs(): void {
    const sheets = this.workbook.sheets;
    for (let i = 0; i < sheets.length; i++) {
      const sheet = sheets[i];
      if (sheet.state !== 'visible') continue;
      const tab = document.createElement('button');
      tab.type = 'button';
      tab.className = 'bym-tab';
      tab.setAttribute('role', 'tab');
      tab.setAttribute('data-sheet', String(i));
      tab.setAttribute('aria-selected', 'false');
      // Buttons: reachable with Tab and chosen with Enter or Space without further code.
      tab.tabIndex = 0;
      tab.textContent = sheet.name;
      const color = resolveColor(sheet.tabColor, this.workbook.styles);
      if (color) tab.style.boxShadow = `inset 0 -3px 0 ${color}`;
      this.tabs.appendChild(tab);
    }
  }

  /** Moves to the next (1) or previous (-1) sheet that has a tab. */
  private stepSheet(step: number): void {
    const sheets = this.workbook.sheets;
    for (let i = this.sheetIndex + step; i >= 0 && i < sheets.length; i += step) {
      if (sheets[i].state === 'visible') {
        // While a formula is typed the sheet is shown for pointing at its cells, and the edit goes on.
        this.attempt(() => (this.formulaEdit() ? this.show(i) : this.setActiveSheet(i)));
        return;
      }
    }
  }

  /** The index of the sheet whose tab an event was on; -1 when it was on none. */
  private tabSheetOf(target: EventTarget | null): number {
    let node = target as HTMLElement | null;
    while (node && node !== this.tabs && !(typeof node.hasAttribute === 'function' && node.hasAttribute('data-sheet'))) node = node.parentElement;
    if (!node || node === this.tabs) return -1;
    const index = Number(node.getAttribute('data-sheet'));
    return Number.isInteger(index) && index >= 0 && index < this.workbook.sheets.length ? index : -1;
  }

  /** The tab of a sheet; undefined for a sheet without one (a hidden sheet). */
  private tabOf(index: number): HTMLElement | undefined {
    const tabs = this.tabs.children;
    for (let i = 0; i < tabs.length; i++) if (tabs[i].getAttribute('data-sheet') === String(index)) return tabs[i] as HTMLElement;
    return undefined;
  }

  /** A point of the window as a point of the view's root, where menus are placed. */
  private rootPoint(clientX: number, clientY: number): { x: number; y: number } {
    const rect = this.root.getBoundingClientRect();
    const scaleX = rect.width > 0 ? this.root.offsetWidth / rect.width : 1;
    const scaleY = rect.height > 0 ? this.root.offsetHeight / rect.height : 1;
    return { x: (clientX - rect.left) * scaleX, y: (clientY - rect.top) * scaleY };
  }

  /** Sizes the button that adds a sheet to the strip of the tabs, and takes it away in a read-only view. */
  private placeAddButton(): void {
    const button = this.addButton;
    const width = this.locked ? 0 : coarsePointer() ? ADD_SHEET_WIDTH_COARSE : ADD_SHEET_WIDTH;
    button.style.display = this.locked ? 'none' : 'block';
    button.style.width = `${width}px`;
    button.setAttribute('aria-label', this.strings.addSheet);
    button.title = this.strings.addSheet;
    // The tabs begin where the button ends, so that none is ever under it.
    this.tabs.style.marginLeft = `${width}px`;
    button.style.height = `${this.tabs.offsetHeight}px`;
  }

  /**
   * Builds the tabs again after the list of sheets changed. Returns whether the keyboard
   * focus was on the strip (a tab, the field a sheet is renamed in), which the caller gives back to it.
   */
  private rebuildTabs(): boolean {
    const root = this.tabs.getRootNode() as Document | ShadowRoot;
    const focused = root.activeElement;
    const within = focused !== null && this.tabs.contains(focused);
    // A name that was being typed goes with its tab. Forgotten first: taking the field away may report that it lost the focus.
    this.renaming = undefined;
    this.endTabDrag();
    this.clearTabPress();
    while (this.tabs.firstChild) this.tabs.removeChild(this.tabs.firstChild);
    this.buildTabs();
    this.placeAddButton();
    return within;
  }

  /**
   * Brings the view in line with a list of sheets that changed under it (a sheet was
   * inserted, deleted or renamed, or that was undone): the tabs, and the sheet that is shown,
   * which is the one at `show`.
   */
  private syncSheets(show: number, kind: SheetOpKind): void {
    const sheets = this.workbook.sheets;
    const index = Number.isInteger(show) && show >= 0 && show < sheets.length ? show : Math.max(sheetToShow(sheets, this.sheetIndex), 0);
    const moved = index !== this.sheetIndex || sheets[index] !== this.shown;
    // A drag or a touch that is under way was begun on the list of sheets as it was (undo by a key in the middle of a resize).
    this.endTabDrag();
    this.endDrag();
    this.endTouch();
    this.lastTap = undefined;
    this.lastTabTap = undefined;
    this.hoverComment = undefined;
    // The marks are kept by the index of their sheet.
    this.marksDirty = true;
    // What was copied holds formulas that name sheets as they were: once a sheet of then is gone or has another
    // name, a paste takes the text instead. A sheet that only moved, was hidden, coloured, added or copied changes
    // nothing a copied formula says, and the copy stays what it is.
    if (copyOutdated(this.copiedNames, sheets)) {
      this.copied = undefined;
      this.copiedNames = undefined;
      this.copyId = '';
    }
    this.sheetIndex = index;
    this.shown = sheets[index];
    if (kind === 'rename') {
      // The same tabs under other names: a press that is under way on one of them still finds it.
      const tabs = this.tabs.children;
      for (let i = 0; i < tabs.length; i++) {
        const at = this.tabSheetOf(tabs[i]);
        if (at >= 0 && tabs[i].textContent !== sheets[at].name) tabs[i].textContent = sheets[at].name;
      }
      this.placeAddButton();
      this.showSheet();
    } else {
      const within = this.rebuildTabs();
      this.showSheet();
      if (within) {
        const tab = this.tabOf(index);
        if (tab) tab.focus({ preventScroll: true });
        else this.focusInput(true);
      }
    }
    // The operation, or its undoing, left the model with the active sheet it worked out; the one shown is it.
    this.noteActiveSheet();
    if (moved && this.options.onSheetChange) this.options.onSheetChange(index);
  }

  /**
   * The formulas after the list of sheets changed: see recalculateForSheets(). `cells` are
   * the rewritten cells of a deleted sheet's readers, and `list` the sheets by whose
   * indexes the engine is told about them.
   */
  private recalculateSheets(kind: SheetListChange['kind'], cells: HistoryMark['sheetCells'], list: Sheet[], unsure: boolean): void {
    this.marksDirty = true;
    if (this.options.calculate === false) return;
    try {
      // Without an engine nothing was calculated here yet; only a sheet that went with readers gives a reason to start.
      if (!this.engine && (kind === 'names' || (cells !== undefined && cells.length === 0 && !unsure))) return;
      if (!this.engine) this.engine = new CalcEngine(this.workbook, { budget: { milliseconds: CALC_SLICE_MS } });
      const engine = this.engine;
      let addresses: { sheet: number; row: number; col: number }[] | undefined;
      if (cells) {
        addresses = [];
        for (let i = 0; i < cells.length; i++) {
          const sheet = list.indexOf(cells[i].sheet);
          if (sheet >= 0) addresses.push({ sheet, row: cells[i].row, col: cells[i].col });
        }
      }
      recalculateForSheets(engine, { kind, cells: addresses, unsure });
      if (engine.incomplete) this.continueCalculation();
    } catch (error) {
      this.engine = undefined;
    }
  }

  /**
   * Applies an operation on the list of sheets as one undo step and brings the view up to
   * date. `index` is the sheet the operation is about. An operation that throws leaves the
   * workbook and the view as they were.
   */
  private applySheetOps(kind: SheetOpKind, ops: Op[], index: number): void {
    this.closeMenu();
    this.finishNote(true, true);
    this.endDrag();
    this.endTouch();
    // A calculation under way is finished first: the graph it works on does not outlive the step.
    if (this.engine && this.engine.incomplete) this.finishCalculation(this.engine);
    const sheets = this.workbook.sheets;
    const names = this.workbook.definedNames;
    if (kind === 'delete') this.buildGraph();
    const shown = this.state().sheet;
    // Undoing shows the sheet the step was about when it has a tab: the one that comes back. Else the one shown now.
    const back = kind === 'delete' && sheets[index].state === 'visible';
    const from = back ? this.stateOf(sheets[index]) : this.state();
    const before: HistoryMark = {
      sheet: back ? index : this.sheetIndex,
      active: { ...from.active },
      range: { ...from.range },
      sheets: allSheets(sheets.length),
      structural: false,
      sheetOp: kind,
    };
    const done: HistoryMark = { ...before };
    const inverse = this.history.apply(ops, before, done);
    const now = this.workbook.sheets;
    // A new sheet is shown, and so is one that was hidden and is shown again; a sheet that stays is still shown,
    // wherever it is now; else the nearest one with a tab.
    let show = kind === 'insert' || kind === 'copy' || (kind === 'state' && now[index] && now[index].state === 'visible') ? index : now.indexOf(shown);
    if (show < 0 || show >= now.length) show = Math.max(sheetToShow(now, index), 0);
    // The sheet that was shown is hidden now: the one the operation made the active one is shown.
    if (now[show].state !== 'visible' && now[this.workbook.activeSheet] && now[this.workbook.activeSheet].state === 'visible') show = this.workbook.activeSheet;
    const to = this.stateOf(now[show]);
    done.sheet = show;
    done.active = { ...to.active };
    done.range = { ...to.range };
    done.sheets = allSheets(now.length);
    let cells: HistoryMark['sheetCells'];
    if (kind === 'delete') {
      // What the step rewrote is what its inverse puts back.
      cells = [];
      for (let i = 0; cells && i < inverse.length; i++) {
        const op = inverse[i];
        if (op.type !== 'restoreSheets') continue;
        for (let k = 0; cells && k < op.state.cells.length; k++) {
          const entry = op.state.cells[k];
          if (cells.length + entry.cells.length > MAX_RECALC_CELLS) cells = undefined;
          else for (let c = 0; c < entry.cells.length; c++) cells.push({ sheet: entry.sheet, row: entry.cells[c].row, col: entry.cells[c].col });
        }
      }
      before.sheetCells = cells;
      done.sheetCells = cells;
    }
    // The engine still has the sheets of before: a sheet that went is told by their indexes.
    this.recalculateSheets(kind === 'delete' ? 'gone' : 'names', cells, sheets, kind === 'delete' && names !== this.workbook.definedNames);
    this.syncSheets(show, kind);
    this.afterChange(done.sheets.slice(), false);
  }

  /**
   * Builds the dependency graph when there is none, before a sheet is deleted: only an engine
   * that knows the formulas of before sees which of them the step rewrites.
   */
  private buildGraph(): void {
    if (this.options.calculate === false) return;
    try {
      if (!this.engine) this.engine = new CalcEngine(this.workbook, { budget: { milliseconds: CALC_SLICE_MS } });
      if (this.engine.graphReady) return;
      this.marksDirty = true;
      this.engine.recalculate([]);
      if (this.engine.incomplete) this.finishCalculation(this.engine);
    } catch (error) {
      this.engine = undefined;
    }
  }

  /**
   * Adds an empty worksheet as sheet number `at` (after the last one when left out) and
   * shows it: one undo step. `name` is its name; without one it gets the next free default
   * name of the view's language ("Sheet1", "Sayfa1"). Returns the index of the new sheet,
   * -1 in a read-only view. Throws RangeError for an index outside the list and for a name
   * a sheet cannot have or another sheet has.
   */
  insertSheet(at?: number, name?: string): number {
    if (!this.beginSheetChange()) return -1;
    const index = at === undefined ? this.workbook.sheets.length : at;
    const ops = insertSheetOps(this.workbook, index, { name: name === undefined ? defaultSheetName(this.workbook, this.languageTag) : String(name) });
    this.applySheetOps('insert', ops, index);
    return index;
  }

  /**
   * Renames a sheet: one undo step. Formulas, defined names and links that name the sheet
   * are written with the new name; no value changes. Returns false when nothing was done
   * (a read-only view, the name the sheet has). Throws RangeError for a name a sheet cannot
   * have or another sheet has, and for a sheet that cannot be renamed (see canRenameSheet()).
   */
  renameSheet(index: number, name: string): boolean {
    if (this.destroyed || this.locked) return false;
    this.endTabDrag();
    this.commitEditing();
    // A name being typed in the tab gives way to this one.
    this.closeRename(false);
    // A note being typed is applied first; one that cannot be stays in its field, and the sheet keeps its name.
    if (!this.commitPending()) return false;
    const ops = renameSheetOps(this.workbook, index, String(name));
    if (ops.length === 0) return false;
    this.applySheetOps('rename', ops, index);
    return true;
  }

  /**
   * Deletes a sheet: one undo step, without asking. What other sheets read from it becomes
   * #REF!. When it was the sheet shown, the next sheet with a tab is (the one before for the
   * last). Returns false in a read-only view. Throws RangeError for a sheet that cannot be
   * deleted (see canDeleteSheet()): the last visible one, say.
   */
  deleteSheet(index: number): boolean {
    if (!this.beginSheetChange()) return false;
    const ops = deleteSheetOps(this.workbook, index);
    const name = this.workbook.sheets[index].name;
    this.applySheetOps('delete', ops, index);
    // After what is said about the cell that is active now.
    this.notifyLater(this.strings.sheetDeleted(name));
    return true;
  }

  /** What every change of the list of sheets through the view starts with; false when the view takes none. */
  private beginSheetChange(): boolean {
    if (this.destroyed || this.locked) return false;
    // A drag of a tab was begun on the list of sheets as it is.
    this.endTabDrag();
    // Whatever is being typed (a cell, a note, the name of a sheet) is applied first, as leaving its field
    // would. What cannot be applied was said, stays in its field, and nothing is done to the sheets.
    return this.commitPending();
  }

  /**
   * Moves a sheet so that it is sheet number `to` (0: first): one undo step. No formula
   * changes; the sheet that is shown stays shown. Returns false when nothing was done (a
   * read-only view, the place the sheet has). Throws RangeError for an index outside the list.
   */
  moveSheet(from: number, to: number): boolean {
    if (!this.beginSheetChange()) return false;
    const ops = moveSheetOps(this.workbook, from, to);
    if (ops.length === 0) return false;
    const name = this.workbook.sheets[from].name;
    this.applySheetOps('move', ops, from);
    let place = 0;
    let count = 0;
    this.workbook.sheets.forEach((sheet, i) => {
      if (sheet.state !== 'visible') return;
      count++;
      if (i <= to) place = count;
    });
    this.notifyLater(this.strings.sheetMoved(name, place, count));
    return true;
  }

  /**
   * Hides a sheet ("hidden"; "veryHidden" so that applications do not offer to show it) or
   * shows it ("visible"): one undo step. When the sheet that is shown is hidden, the next one
   * with a tab is shown; a sheet that is shown again is the one shown. Returns false when
   * nothing was done. Throws RangeError for the last visible sheet.
   */
  setSheetState(index: number, state: 'visible' | 'hidden' | 'veryHidden'): boolean {
    if (!this.beginSheetChange()) return false;
    const ops = setSheetStateOps(this.workbook, index, state);
    if (ops.length === 0) return false;
    this.applySheetOps('state', ops, index);
    return true;
  }

  /**
   * Copies a sheet into a new one, by default right behind it and named like it with a
   * number ("Name (2)"), and shows the copy: one undo step. What the copy does not get
   * (conditional formats, data validations, pictures, tables, ...: see copySheetLosses()) is
   * said to onNotice. Returns the index of the copy, -1 in a read-only view. Throws
   * RangeError for a sheet that cannot be copied and a name that cannot be used.
   */
  copySheet(index: number, options: CopySheetOptions = {}): number {
    if (!this.beginSheetChange()) return -1;
    const ops = copySheetOps(this.workbook, index, options);
    const losses = copySheetLosses(this.workbook, index);
    const name = this.workbook.sheets[index].name;
    const at = options.at !== undefined ? options.at : index + 1;
    this.applySheetOps('copy', ops, at);
    if (losses.length > 0) this.notifyLater(this.strings.sheetCopiedWithout(name, losses.map((loss) => this.strings.copyLoss(loss.what, loss.count)).join(', ')));
    return at;
  }

  /**
   * Sets the colour of the tab of a sheet ("#RRGGBB"); undefined removes it: one undo step.
   * Returns false when nothing was done. Throws RangeError for a colour written otherwise.
   */
  setTabColor(index: number, color: string | undefined): boolean {
    if (!this.beginSheetChange()) return false;
    const ops = setTabColorOps(this.workbook, index, color);
    if (ops.length === 0) return false;
    this.applySheetOps('color', ops, index);
    return true;
  }

  /** The sheets that are hidden and can be shown from the menu of a tab (not the very hidden ones), by index. */
  get hiddenSheets(): number[] {
    const out: number[] = [];
    this.workbook.sheets.forEach((sheet, index) => {
      if (sheet.state === 'hidden') out.push(index);
    });
    return out;
  }

  /** Opens the list of the hidden sheets at a point of the root: choosing one shows it. */
  private openHiddenSheets(x: number, y: number, back: () => void): void {
    const items: MenuItem[] = this.hiddenSheets.map((index) => {
      const sheet = this.workbook.sheets[index];
      // The name of a sheet is text of the file: the menu only ever shows it as text.
      return { label: sheet.name, action: () => this.attempt(() => void (this.workbook.sheets.indexOf(sheet) >= 0 && this.setSheetState(this.workbook.sheets.indexOf(sheet), 'visible'))) };
    });
    if (items.length === 0) return;
    this.menu = new ContextMenu(this.root, items, x, y, () => {
      this.menu = undefined;
      if (!this.destroyed) back();
    });
    this.menu.element.setAttribute('aria-label', this.strings.hiddenSheets);
  }

  /** Opens the palette of the colour of a tab under the tab. */
  private openTabColors(sheet: Sheet, back: () => void): void {
    const tab = this.tabOf(this.workbook.sheets.indexOf(sheet));
    if (!tab) return;
    const current = resolveColor(sheet.tabColor, this.workbook.styles);
    const apply = (color: string | undefined) => () => {
      this.attempt(() => {
        const index = this.workbook.sheets.indexOf(sheet);
        if (index >= 0) this.setTabColor(index, color);
      });
      // The tabs were built again: back to the tab the palette belonged to, or to the cells.
      back();
    };
    const swatches: PopoverItem[] = [];
    for (let i = 0; i < PALETTE.length; i++) {
      const hex = `#${PALETTE[i]}`;
      swatches.push({ label: hex, swatch: hex, value: hex, selected: current !== undefined && current.toUpperCase() === hex, action: apply(hex) });
    }
    this.tabPopover = new Popover({
      parent: this.root,
      anchor: tab,
      label: this.strings.tabColor,
      items: [{ label: this.strings.noTabColor, action: apply(undefined) }],
      swatches,
      columns: PALETTE_COLUMNS,
      onClose: (reason) => {
        this.tabPopover = undefined;
        // A colour that was chosen gives the focus back itself, once the tabs are built again; a press elsewhere took it along.
        if (!this.destroyed && (reason === 'escape' || reason === 'tab')) back();
      },
    });
  }

  /**
   * Opens the field a sheet is renamed in, in the place of its tab, and shows the sheet.
   * Enter or leaving the field renames, Escape gives up. Returns false when there is no
   * field (a read-only view, a hidden sheet). Throws RangeError for a sheet that cannot be renamed.
   */
  startRename(index: number): boolean {
    if (this.destroyed || this.locked) return false;
    const sheet = this.workbook.sheets[index];
    if (!sheet) throw new RangeError(`No sheet with index ${index}`);
    if (this.renaming && this.renaming.sheet === sheet) return true;
    // A sheet without a tab has no place for the field; nothing else happens either.
    if (sheet.state !== 'visible') return false;
    this.finishRename(true, true);
    const problem = canRenameSheet(this.workbook, index);
    if (problem) throw new RangeError(problem);
    this.commitEditing();
    this.closeMenu();
    if (index !== this.sheetIndex) this.setActiveSheet(index);
    const tab = this.tabOf(index);
    if (!tab) return false;
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'bym-tab-input';
    input.maxLength = MAX_SHEET_NAME;
    input.setAttribute('aria-label', this.strings.sheetName);
    input.setAttribute('autocomplete', 'off');
    input.setAttribute('autocorrect', 'off');
    input.setAttribute('autocapitalize', 'off');
    input.spellcheck = false;
    // The name of a sheet comes from the file: it is only ever the value of the field.
    input.value = sheet.name;
    input.style.width = `${Math.max(tab.offsetWidth + 24, 96)}px`;
    tab.style.display = 'none';
    // The tab list holds tabs: the field stands in it inside an element that is nothing to assistive technology.
    const box = document.createElement('div');
    box.className = 'bym-tab-edit';
    box.setAttribute('role', 'none');
    box.style.cssText = 'display:block;flex:none';
    box.appendChild(input);
    this.tabs.insertBefore(box, tab);
    input.addEventListener('keydown', this.onRenameKey);
    input.addEventListener('input', this.onRenameInput);
    input.addEventListener('blur', this.onRenameBlur);
    input.addEventListener('compositionstart', this.onFieldComposition);
    input.addEventListener('compositionend', this.onFieldComposition);
    this.renaming = { sheet, tab, input, box, busy: false };
    this.emitEditing();
    input.focus({ preventScroll: true });
    input.select();
    const left = box.offsetLeft;
    const right = left + box.offsetWidth;
    if (left < this.tabs.scrollLeft) this.tabs.scrollLeft = left;
    else if (right > this.tabs.scrollLeft + this.tabs.clientWidth) this.tabs.scrollLeft = right - this.tabs.clientWidth;
    return true;
  }

  /** Whether a sheet is being renamed in its tab. */
  get isRenaming(): boolean {
    return this.renaming !== undefined;
  }

  /**
   * Ends renaming in the tab. With `apply` the sheet takes the name in the field; a name it
   * cannot take is said, and the field stays open with it - unless the field is being left
   * (`leaving`: the focus went elsewhere, another sheet is shown), where the sheet keeps its name.
   */
  private finishRename(apply: boolean, leaving = false): void {
    const renaming = this.renaming;
    if (!renaming || renaming.busy) return;
    const { sheet, input } = renaming;
    const focused = (input.getRootNode() as Document | ShadowRoot).activeElement === input;
    const index = this.workbook.sheets.indexOf(sheet);
    if (apply && index >= 0 && input.value !== sheet.name) {
      renaming.busy = true;
      try {
        this.applySheetOps('rename', renameSheetOps(this.workbook, index, input.value), index);
        renaming.busy = false;
        this.closeRename(!leaving && focused);
        return;
      } catch (error) {
        renaming.busy = false;
        const message = this.localizeError(error);
        if (!leaving) {
          input.setAttribute('aria-invalid', 'true');
          input.title = message;
          input.select();
          this.notify(message, 'error');
          return;
        }
        this.closeRename(false);
        // What made the field be left (a press on a cell, on a tab) says something of its own right after.
        this.notifyLater(message, 'error');
        return;
      }
    }
    this.closeRename(!leaving && focused);
  }

  /** Takes the field away and shows the tab again, with the focus when asked to. */
  private closeRename(focusTab: boolean): void {
    const renaming = this.renaming;
    if (!renaming) return;
    // Forgotten first: taking the field away may report that it lost the focus.
    this.renaming = undefined;
    const { input, tab, box } = renaming;
    input.removeEventListener('keydown', this.onRenameKey);
    input.removeEventListener('input', this.onRenameInput);
    input.removeEventListener('blur', this.onRenameBlur);
    input.removeEventListener('compositionstart', this.onFieldComposition);
    input.removeEventListener('compositionend', this.onFieldComposition);
    if (box.parentNode) box.parentNode.removeChild(box);
    tab.style.display = '';
    if (this.destroyed) return;
    if (focusTab) tab.focus({ preventScroll: true });
    this.emitEditing();
  }

  private onRenameKey = (event: KeyboardEvent): void => {
    // Saving is the business of whoever holds the view: the key goes up to it (mount() applies the name first).
    if (isSaveKey(event)) return;
    // The other keys of the field are its own (Ctrl+Z undoes what was typed in it): none reaches the shortcuts around it.
    event.stopPropagation();
    // Enter that confirms an input method's candidate is not Enter for the name.
    if (this.destroyed || event.isComposing || event.keyCode === 229) return;
    if (event.key === 'Enter') {
      event.preventDefault();
      this.finishRename(true);
    } else if (event.key === 'Escape') {
      event.preventDefault();
      this.finishRename(false);
    } else if ((event.ctrlKey || event.metaKey) && !event.altKey && !event.shiftKey && (event.key === 'p' || event.key === 'P')) {
      // As in the grid: the browser would print the host page.
      event.preventDefault();
      this.print().catch(() => undefined);
    }
  };

  private onRenameInput = (event: Event): void => {
    const input = event.target as HTMLInputElement;
    if (!input.hasAttribute('aria-invalid')) return;
    // Another name is being typed: what was said about the last one no longer holds.
    input.removeAttribute('aria-invalid');
    input.title = '';
  };

  private onRenameBlur = (event: FocusEvent): void => {
    const renaming = this.renaming;
    if (this.destroyed || !renaming || renaming.input !== event.target) return;
    const target = this.renameTarget;
    this.renameTarget = undefined;
    this.finishRename(true, true);
    // The press that took the focus away was on a tab. The tabs moved when the field closed, so the click
    // that follows may land on another one, or on none: the sheet that was pressed is shown here instead.
    const index = target ? this.workbook.sheets.indexOf(target) : -1;
    if (index < 0 || this.renaming) return;
    this.swallowTabClick = true;
    this.attempt(() => this.setActiveSheet(index));
  };

  /**
   * Opens the menu of a tab at a point of the root: a sheet after this one, another name,
   * deleting. The tab is chosen first, as a click would. What cannot be done is shown but
   * not available, and says why. Returns false when no menu was opened (a read-only view).
   */
  private openTabMenu(index: number, x: number, y: number, keyboard: boolean): boolean {
    // The list of sheets stays as it is while a formula is typed.
    if (this.destroyed || this.locked || this.formulaEdit()) return false;
    const sheet = this.workbook.sheets[index];
    if (!sheet) return false;
    this.finishRename(true, true);
    this.closeMenu();
    if (!this.tryCommit()) return false;
    if (index !== this.sheetIndex) this.setActiveSheet(index);
    const strings = this.strings;
    const renameProblem = canRenameSheet(this.workbook, index);
    const deleteProblem = canDeleteSheet(this.workbook, index);
    const copyProblem = canCopySheet(this.workbook, index);
    const hideProblem = canHideSheet(this.workbook, index);
    // The menu closes before an item acts, but the list of sheets is looked at again all the same.
    const at = () => this.workbook.sheets.indexOf(sheet);
    const blocked = (problem: string | undefined) =>
      problem === undefined ? {} : { disabled: true, title: this.localizeError(problem), explain: () => this.notify(this.localizeError(problem), 'error') };
    // The tab beside this one on either side, among the sheets that have a tab.
    const beside = (step: number): number => {
      const sheets = this.workbook.sheets;
      for (let i = at() + step; at() >= 0 && i >= 0 && i < sheets.length; i += step) if (sheets[i].state === 'visible') return i;
      return -1;
    };
    const backToTab = () => {
      const tab = at() >= 0 ? this.tabOf(at()) : undefined;
      if (tab && keyboard) tab.focus({ preventScroll: true });
      else this.focusInput(true);
    };
    const more: MenuItem[] = [
      {
        label: strings.duplicateSheet,
        ...blocked(copyProblem),
        action: () =>
          this.attempt(() => {
            if (at() >= 0) this.copySheet(at());
          }),
      },
      { label: '', separator: true },
      {
        label: strings.moveSheetLeft,
        disabled: beside(-1) < 0,
        action: () =>
          this.attempt(() => {
            if (beside(-1) >= 0) this.moveSheet(at(), beside(-1));
          }),
      },
      {
        label: strings.moveSheetRight,
        disabled: beside(1) < 0,
        action: () =>
          this.attempt(() => {
            if (beside(1) >= 0) this.moveSheet(at(), beside(1));
          }),
      },
      { label: '', separator: true },
      {
        label: strings.hideSheet,
        ...blocked(hideProblem),
        action: () =>
          this.attempt(() => {
            if (at() >= 0) this.setSheetState(at(), 'hidden');
          }),
      },
    ];
    if (this.hiddenSheets.length > 0) more.push({ label: strings.unhideSheet, action: () => this.openHiddenSheets(x, y, backToTab) });
    if (sheet.kind === 'worksheet') more.push({ label: strings.tabColor, action: () => this.openTabColors(sheet, backToTab) });
    const items: MenuItem[] = [
      {
        label: strings.insertSheet,
        action: () =>
          this.attempt(() => {
            if (at() >= 0) this.insertSheet(at() + 1);
          }),
      },
      {
        label: strings.renameSheet,
        disabled: renameProblem !== undefined,
        title: renameProblem === undefined ? undefined : this.localizeError(renameProblem),
        explain: renameProblem === undefined ? undefined : () => this.notify(this.localizeError(renameProblem), 'error'),
        action: () =>
          this.attempt(() => {
            if (at() >= 0) this.startRename(at());
          }),
      },
      {
        label: strings.deleteSheet,
        disabled: deleteProblem !== undefined,
        title: deleteProblem === undefined ? undefined : this.localizeError(deleteProblem),
        explain: deleteProblem === undefined ? undefined : () => this.notify(this.localizeError(deleteProblem), 'error'),
        action: () =>
          this.attempt(() => {
            if (at() >= 0) this.deleteSheet(at());
          }),
      },
      ...more,
    ];
    this.menu = new ContextMenu(
      this.root,
      items,
      x,
      y,
      () => {
        this.menu = undefined;
        if (this.destroyed) return;
        // Opened from the keyboard on a tab: back to that tab. Otherwise to the cells. (An item that opens a
        // list of its own - the hidden sheets, the colours - takes the focus from there right after.)
        const tab = keyboard && at() >= 0 ? this.tabOf(at()) : undefined;
        if (tab) tab.focus({ preventScroll: true });
        else this.focusInput(true);
      },
      // Opened under a finger that is still down.
      this.tabPress !== undefined,
    );
    return true;
  }

  private onTabClick = (event: MouseEvent): void => {
    if (this.destroyed) return;
    const index = this.tabSheetOf(event.target);
    if (index < 0) return;
    // Lifting the finger of a long press is not a tap on the tab, and neither is the click after a press that
    // was dealt with when the field of a name closed. A tab chosen with a key is never either.
    const lifted = this.swallowTabClick && event.detail > 0;
    this.swallowTabClick = false;
    if (lifted) return;
    // From here on the click itself says which tab was pressed.
    this.renameTarget = undefined;
    const sheet = this.workbook.sheets[index];
    const pointer = event.detail > 0;
    if (this.formulaEdit()) {
      // A formula is being typed: the sheet is shown for pointing at its cells. The edit goes on, the focus
      // stays in the field it is typed in, and a second click renames nothing.
      this.lastTabTap = undefined;
      this.show(index);
      return;
    }
    // A pointer click hands the keyboard back to the cells; a tab chosen with the keyboard keeps the focus.
    if (pointer) this.focusInput(true);
    // A double click, or two taps of a finger, on the same tab renames the sheet. The clicks are counted here:
    // the count of the browser goes on after the second, and a finger has none everywhere.
    const last = this.lastTabTap;
    const twice = pointer && !this.locked && last !== undefined && last.sheet === sheet && event.timeStamp - last.time < DOUBLE_TAP_MS;
    this.lastTabTap = pointer && !twice ? { time: event.timeStamp, sheet } : undefined;
    this.attempt(() => {
      if (twice) this.startRename(index);
      else this.setActiveSheet(index);
    });
  };

  private onTabKeyDown = (event: KeyboardEvent): void => {
    if (this.destroyed || this.locked || event.defaultPrevented) return;
    const target = event.target as HTMLElement | null;
    // Only the keys of a tab: the field a sheet is renamed in keeps its own.
    if (!target || !target.classList || !target.classList.contains('bym-tab')) return;
    const index = this.tabSheetOf(target);
    if (index < 0) return;
    const locked = this.formulaEdit();
    if (event.key === 'F2' && !event.ctrlKey && !event.metaKey && !event.altKey && !event.shiftKey) {
      event.preventDefault();
      if (locked) return;
      this.attempt(() => {
        this.startRename(index);
      });
    } else if (event.key === 'ContextMenu' || (event.key === 'F10' && event.shiftKey && !event.ctrlKey && !event.metaKey && !event.altKey)) {
      event.preventDefault();
      if (locked) return;
      const rect = target.getBoundingClientRect();
      const point = this.rootPoint(rect.left, rect.top);
      this.attempt(() => {
        this.openTabMenu(index, point.x, point.y, true);
      });
    }
  };

  private onTabContextMenu = (event: MouseEvent): void => {
    // A read-only view has no menu of its own; the browser's is left alone.
    if (this.destroyed || this.locked) return;
    const index = this.tabSheetOf(event.target);
    if (index < 0) return;
    event.preventDefault();
    // The long press or the key that this event follows opened the menu already. None while a formula is typed.
    if (this.menu || this.formulaEdit()) return;
    const tab = this.tabOf(index);
    const rect = tab ? tab.getBoundingClientRect() : undefined;
    // From the keyboard the event carries no place of its own.
    const inside = rect !== undefined && event.clientX >= rect.left && event.clientX <= rect.right && event.clientY >= rect.top && event.clientY <= rect.bottom;
    const point = inside || !rect ? this.rootPoint(event.clientX, event.clientY) : this.rootPoint(rect.left, rect.top);
    const held = this.tabPress !== undefined;
    this.clearTabPressTimer();
    this.attempt(() => {
      if (this.openTabMenu(index, point.x, point.y, !inside) && held) this.swallowTabClick = true;
    });
  };

  private onTabPointerDown = (event: PointerEvent): void => {
    // Whatever an earlier press left behind (its end was never heard of) is over.
    this.endTabDrag();
    this.clearTabPress();
    this.swallowTabClick = false;
    this.renameTarget = undefined;
    if (this.destroyed || this.locked) return;
    const index = this.tabSheetOf(event.target);
    if (index < 0) return;
    // A name is being typed: the field closes when this press takes the focus, and the tabs move under the pointer.
    if (this.renaming) this.renameTarget = this.workbook.sheets[index];
    // The list of sheets stays as it is while a formula is typed.
    if (this.formulaEdit()) return;
    if (event.pointerType !== 'touch') {
      // The main button on a tab: the tab is dragged once the pointer goes far enough.
      if (event.button === 0 && !this.renaming && (event.target as HTMLElement).classList.contains('bym-tab')) {
        this.tabDrag = { pointerId: event.pointerId, sheet: this.workbook.sheets[index], x: event.clientX, touch: false, started: false, gap: -1, lastX: event.clientX };
        this.watchTabDrag();
      }
      return;
    }
    this.tabPress = { pointerId: event.pointerId, x: event.clientX, y: event.clientY, sheet: this.workbook.sheets[index], scroll: this.tabs.scrollLeft };
    this.tabPressTimer = window.setTimeout(this.onTabLongPress, LONG_PRESS_MS);
  };

  private onTabPointerMove = (event: PointerEvent): void => {
    const drag = this.tabDrag;
    if (drag && drag.pointerId === event.pointerId) {
      // The button was let go where the strip did not hear of it (outside the window, say).
      if (!drag.touch && (event.buttons & 1) === 0) {
        this.endTabDrag();
        return;
      }
      if (!drag.started && Math.abs(event.clientX - drag.x) > TAB_DRAG_SLOP) this.startTabDrag(drag);
      if (drag.started) {
        event.preventDefault();
        drag.lastX = event.clientX;
        this.moveTabDrag(drag);
      }
      return;
    }
    const press = this.tabPress;
    if (!press || press.pointerId !== event.pointerId) return;
    // A finger that moves scrolls the tabs.
    if (Math.abs(event.clientX - press.x) > TAP_SLOP || Math.abs(event.clientY - press.y) > TAP_SLOP) this.clearTabPress();
  };

  private onTabPointerEnd = (event: PointerEvent): void => {
    const drag = this.tabDrag;
    if (drag && drag.pointerId === event.pointerId) {
      const started = drag.started;
      const gap = drag.gap;
      this.endTabDrag();
      if (started) {
        // Letting go is not a click on whatever tab is under the pointer now.
        this.swallowTabClick = true;
        this.clearTabPress();
        const from = this.workbook.sheets.indexOf(drag.sheet);
        const to = event.type === 'pointerup' && from >= 0 ? dropIndex(this.visibleSheets(), from, gap) : undefined;
        if (to !== undefined) this.attempt(() => void this.moveSheet(from, to));
        return;
      }
    }
    const press = this.tabPress;
    if (press && press.pointerId === event.pointerId) this.clearTabPress();
  };

  /** A finger that drags a tab must not scroll the strip: the browser asks here before it does. */
  private onTabTouchMove = (event: TouchEvent): void => {
    if (this.tabDrag && this.tabDrag.touch && event.cancelable) event.preventDefault();
  };

  /** The indexes of the sheets that have a tab, in order. */
  private visibleSheets(): number[] {
    const out: number[] = [];
    this.workbook.sheets.forEach((sheet, index) => {
      if (sheet.state === 'visible') out.push(index);
    });
    return out;
  }

  private startTabDrag(drag: NonNullable<WorkbookView['tabDrag']>): void {
    // The menu a long press opened gives way: the finger went on to drag the tab.
    if (drag.touch) this.closeMenu();
    if (this.destroyed || this.locked || this.formulaEdit() || this.workbook.sheets.indexOf(drag.sheet) < 0 || !this.tryCommit()) {
      this.endTabDrag();
      return;
    }
    this.finishRename(true, true);
    drag.started = true;
    try {
      // The strip goes on hearing of the pointer wherever it goes.
      this.tabs.setPointerCapture(drag.pointerId);
    } catch (error) {
      // A pointer that is gone already; its end is all that is left to come.
    }
    const mark = document.createElement('div');
    mark.className = 'bym-tab-drop';
    mark.setAttribute('role', 'none');
    mark.style.cssText = 'position:absolute;top:0;bottom:0;width:2px;margin-left:-1px;background:#1a73e8;pointer-events:none;z-index:1';
    if (getComputedStyle(this.tabs).position === 'static') this.tabs.style.position = 'relative';
    this.tabs.appendChild(mark);
    this.tabDropMark = mark;
    const tab = this.tabOf(this.workbook.sheets.indexOf(drag.sheet));
    if (tab) tab.setAttribute('data-dragging', 'true');
    this.tabs.style.cursor = 'grabbing';
  }

  /** Places the mark of a dragged tab by where its pointer is, and scrolls the strip while the pointer is at one of its ends. */
  private moveTabDrag(drag: NonNullable<WorkbookView['tabDrag']>): void {
    const strip = this.tabs.getBoundingClientRect();
    const visible = this.visibleSheets();
    let gap = visible.length;
    let left = 0;
    for (let i = 0; i < visible.length; i++) {
      const tab = this.tabOf(visible[i]);
      if (!tab) continue;
      const rect = tab.getBoundingClientRect();
      left = tab.offsetLeft + tab.offsetWidth;
      if (drag.lastX < rect.left + rect.width / 2) {
        gap = i;
        left = tab.offsetLeft;
        break;
      }
    }
    drag.gap = gap;
    if (this.tabDropMark) this.tabDropMark.style.left = `${left}px`;
    const edge = drag.lastX < strip.left + TAB_SCROLL_EDGE ? -1 : drag.lastX > strip.right - TAB_SCROLL_EDGE ? 1 : 0;
    if (edge === 0) {
      if (this.tabDragTimer) window.clearInterval(this.tabDragTimer);
      this.tabDragTimer = 0;
    } else if (!this.tabDragTimer) {
      this.tabDragTimer = window.setInterval(() => {
        const current = this.tabDrag;
        if (this.destroyed || !current || !current.started) return;
        const before = this.tabs.scrollLeft;
        this.tabs.scrollLeft = before + (current.lastX < this.tabs.getBoundingClientRect().left + TAB_SCROLL_EDGE ? -TAB_SCROLL_STEP : TAB_SCROLL_STEP);
        if (this.tabs.scrollLeft !== before) this.moveTabDrag(current);
      }, TAB_SCROLL_MS);
    }
  }

  /**
   * While a tab may be dragged, the end of its pointer is listened for on the window as well (a button
   * let go beside the strip is not heard of by the strip), and so is Escape, which gives the drag up.
   */
  private watchTabDrag(): void {
    if (this.tabDragWatched) return;
    this.tabDragWatched = true;
    window.addEventListener('pointerup', this.onTabDragEnd, true);
    window.addEventListener('pointercancel', this.onTabDragEnd, true);
    window.addEventListener('keydown', this.onTabDragKey, true);
    window.addEventListener('blur', this.onTabDragBlur);
  }

  /** The end of a pointer anywhere in the window: one the strip does not hear of itself ends the drag, and moves nothing. */
  private onTabDragEnd = (event: PointerEvent): void => {
    const drag = this.tabDrag;
    if (!drag || drag.pointerId !== event.pointerId) return;
    const target = event.target as Node | null;
    // On the strip (always so once the pointer is captured): its own listener decides what the end means.
    if (target && (target === this.tabs || this.tabs.contains(target))) return;
    this.endTabDrag();
  };

  private onTabDragKey = (event: KeyboardEvent): void => {
    const drag = this.tabDrag;
    if (!drag || !drag.started || event.key !== 'Escape') return;
    // The key gives the drag up and does nothing else.
    event.preventDefault();
    event.stopPropagation();
    this.endTabDrag();
    // Letting go afterwards is not a click on a tab.
    this.swallowTabClick = true;
  };

  private onTabDragBlur = (): void => {
    this.endTabDrag();
  };

  /** Ends the drag of a tab, whichever way: the mark goes, the strip is as it was. Moves nothing. */
  private endTabDrag(): void {
    const drag = this.tabDrag;
    this.tabDrag = undefined;
    if (this.tabDragTimer) window.clearInterval(this.tabDragTimer);
    this.tabDragTimer = 0;
    if (this.tabDragWatched) {
      this.tabDragWatched = false;
      window.removeEventListener('pointerup', this.onTabDragEnd, true);
      window.removeEventListener('pointercancel', this.onTabDragEnd, true);
      window.removeEventListener('keydown', this.onTabDragKey, true);
      window.removeEventListener('blur', this.onTabDragBlur);
    }
    if (this.tabDropMark && this.tabDropMark.parentNode) this.tabDropMark.parentNode.removeChild(this.tabDropMark);
    this.tabDropMark = undefined;
    // Whatever a drag set, whether or not this call knows of one.
    if (this.tabs.style.cursor !== '') this.tabs.style.cursor = '';
    const tabs = this.tabs.children;
    for (let i = 0; i < tabs.length; i++) if (tabs[i].hasAttribute('data-dragging')) tabs[i].removeAttribute('data-dragging');
    if (!drag || !drag.started) return;
    try {
      this.tabs.releasePointerCapture(drag.pointerId);
    } catch (error) {
      // Released already.
    }
  }

  private onTabLongPress = (): void => {
    this.tabPressTimer = 0;
    const press = this.tabPress;
    if (this.destroyed || !press || this.menu || this.tabs.scrollLeft !== press.scroll) return;
    const index = this.workbook.sheets.indexOf(press.sheet);
    if (index < 0) return;
    const point = this.rootPoint(press.x, press.y);
    this.attempt(() => {
      if (this.openTabMenu(index, point.x, point.y, false)) this.swallowTabClick = true;
    });
    // The finger is still down: lifted where it is, the menu stays; moved on from here, it drags the tab
    // and the menu gives way (see startTabDrag()).
    if (!this.locked && this.workbook.sheets.indexOf(press.sheet) >= 0) {
      this.tabDrag = { pointerId: press.pointerId, sheet: press.sheet, x: press.x, touch: true, started: false, gap: -1, lastX: press.x };
      this.watchTabDrag();
    }
  };

  private clearTabPressTimer(): void {
    if (this.tabPressTimer) window.clearTimeout(this.tabPressTimer);
    this.tabPressTimer = 0;
  }

  private clearTabPress(): void {
    this.clearTabPressTimer();
    this.tabPress = undefined;
    // A drag that has not begun goes with the press it would have come from.
    if (this.tabDrag && !this.tabDrag.started) this.tabDrag = undefined;
  }

  private onAddClick = (event: MouseEvent): void => {
    // The list of sheets stays as it is while a formula is typed.
    if (this.destroyed || this.locked || this.formulaEdit()) return;
    this.attempt(() => {
      this.insertSheet();
    });
    // A pointer click hands the keyboard to the cells of the new sheet; a key leaves the focus on the button.
    if (event.detail > 0) this.focusInput(true);
  };

  // ------------------------------------------------------- size and scroll

  private layout(state: SheetState = this.state()): GridLayout {
    const view = { frozenRows: count(state.sheet.view.frozenRows), frozenCols: count(state.sheet.view.frozenCols) };
    return createLayout(state.metrics, view, { width: this.width, height: this.height, scrollX: state.scrollX, scrollY: state.scrollY });
  }

  /** Reads the size of the visible area; resizes the canvas when it or the pixel ratio changed. */
  private measure(): boolean {
    const width = this.scroller.clientWidth;
    const height = this.scroller.clientHeight;
    let dpr = window.devicePixelRatio || 1;
    if (width * height * dpr * dpr > MAX_CANVAS_PIXELS) dpr = Math.sqrt(MAX_CANVAS_PIXELS / (width * height));
    if (width === this.width && height === this.height && dpr === this.dpr) return false;
    this.width = width;
    this.height = height;
    this.dpr = dpr;
    this.canvas.style.width = `${width}px`;
    this.canvas.style.height = `${height}px`;
    this.canvas.width = Math.round(width * dpr);
    this.canvas.height = Math.round(height * dpr);
    return true;
  }

  private onResize = (): void => {
    if (this.destroyed) return;
    // The strip of the tabs is higher while it shows a scroll bar.
    this.placeAddButton();
    if (!this.measure()) return;
    const state = this.state();
    this.growExtent(state.scrollX, state.scrollY);
    this.commitScroll();
    // Resizing cleared the canvas; draw before the browser paints.
    this.draw();
  };

  private watchPixelRatio(): void {
    this.unwatchPixelRatio();
    if (typeof window.matchMedia !== 'function') return;
    const query = window.matchMedia(`(resolution: ${window.devicePixelRatio || 1}dppx)`);
    if (query.addEventListener) query.addEventListener('change', this.onPixelRatioChange);
    else if (query.addListener) query.addListener(this.onPixelRatioChange);
    this.dprQuery = query;
  }

  private unwatchPixelRatio(): void {
    const query = this.dprQuery;
    if (!query) return;
    if (query.removeEventListener) query.removeEventListener('change', this.onPixelRatioChange);
    else if (query.removeListener) query.removeListener(this.onPixelRatioChange);
    this.dprQuery = undefined;
  }

  private onPixelRatioChange = (): void => {
    if (this.destroyed) return;
    // The query matched the old ratio only; watch the new one.
    this.watchPixelRatio();
    this.onResize();
  };

  /**
   * Makes the scrollable area reach at least two visible areas past the given
   * offsets and past the used cells. Returns whether it grew.
   */
  private growExtent(scrollX: number, scrollY: number): boolean {
    const state = this.state();
    const layout = this.layout(state);
    const area = scrollAreaSize(layout);
    const { rows, cols } = state.metrics;
    const wantRows = Math.min(
      rows.count,
      Math.max(count(state.sheet.rowCount) + EXTRA_ROWS, rows.indexAt(layout.frozenHeight + scrollY + 2 * area.height) + 1),
    );
    const wantCols = Math.min(
      cols.count,
      Math.max(count(state.sheet.colCount) + EXTRA_COLS, cols.indexAt(layout.frozenWidth + scrollX + 2 * area.width) + 1),
    );
    let grew = false;
    if (wantRows > state.extentRows) {
      state.extentRows = wantRows;
      grew = true;
    }
    if (wantCols > state.extentCols) {
      state.extentCols = wantCols;
      grew = true;
    }
    return grew;
  }

  /**
   * Clamps the logical scroll offsets, sizes the spacer and moves the native
   * scroll position to match. The native position is left alone when it
   * already matches, so a scroll the user is making is never interrupted.
   */
  private commitScroll(): void {
    const state = this.state();
    const layout = this.layout(state);
    const area = scrollAreaSize(layout);
    const max = maxScroll(layout);
    const { rows, cols } = state.metrics;
    this.limitX = clamp(cols.offset(state.extentCols) - layout.frozenWidth - area.width, 0, max.x);
    this.limitY = clamp(rows.offset(state.extentRows) - layout.frozenHeight - area.height, 0, max.y);
    state.scrollX = Math.min(layout.scrollX, this.limitX);
    state.scrollY = Math.min(layout.scrollY, this.limitY);
    this.rangeX = nativeRange(this.limitX, this.width);
    this.rangeY = nativeRange(this.limitY, this.height);
    this.spacer.style.width = `${this.width + this.rangeX}px`;
    this.spacer.style.height = `${this.height + this.rangeY}px`;

    const left = toNative(state.scrollX, this.limitX, this.rangeX);
    const top = toNative(state.scrollY, this.limitY, this.rangeY);
    const scroller = this.scroller;
    if (Math.abs(scroller.scrollLeft - left) >= 1) scroller.scrollLeft = left;
    if (Math.abs(scroller.scrollTop - top) >= 1) scroller.scrollTop = top;
    this.domX = scroller.scrollLeft;
    this.domY = scroller.scrollTop;
  }

  /** Takes over the native scroll position if the user moved it. Returns whether the logical offsets changed. */
  private readScroll(): boolean {
    const state = this.state();
    const left = this.scroller.scrollLeft;
    const top = this.scroller.scrollTop;
    let changed = false;
    if (left !== this.domX) {
      this.domX = left;
      const x = fromNative(left, this.limitX, this.rangeX);
      if (x !== state.scrollX) {
        state.scrollX = x;
        changed = true;
      }
    }
    if (top !== this.domY) {
      this.domY = top;
      const y = fromNative(top, this.limitY, this.rangeY);
      if (y !== state.scrollY) {
        state.scrollY = y;
        changed = true;
      }
    }
    // What was tapped a moment ago is somewhere else now.
    if (changed) this.lastTap = undefined;
    return changed;
  }

  private onScroll = (): void => {
    if (this.destroyed || !this.readScroll()) return;
    this.closeMenu();
    const state = this.state();
    if (this.growExtent(state.scrollX, state.scrollY)) this.commitScroll();
    this.schedule();
  };

  /** Scrolls to logical offsets, growing the scrollable area when they lie beyond it. */
  private scrollToOffsets(scrollX: number, scrollY: number): boolean {
    const state = this.state();
    const max = maxScroll(this.layout(state));
    const x = clamp(scrollX, 0, max.x);
    const y = clamp(scrollY, 0, max.y);
    if (x === state.scrollX && y === state.scrollY) return false;
    this.growExtent(x, y);
    state.scrollX = x;
    state.scrollY = y;
    this.commitScroll();
    this.schedule();
    return true;
  }

  /**
   * Scrolls a cell into view. For a selection of whole rows only the row is
   * brought into view, for whole columns only the column: the far end of such
   * a selection is the last column or row of the sheet, which nobody wants to see.
   */
  private reveal(cell: CellPos, selection?: CellRange): void {
    this.readScroll();
    const state = this.state();
    const target = scrollToReveal(this.layout(state), cell.row, cell.col);
    this.scrollToOffsets(
      selection && fullWidth(selection) ? state.scrollX : target.scrollX,
      selection && fullHeight(selection) ? state.scrollY : target.scrollY,
    );
  }

  // --------------------------------------------------------------- drawing

  private schedule(): void {
    if (this.destroyed || this.frame) return;
    this.frame = requestAnimationFrame(() => {
      this.frame = 0;
      this.draw();
    });
  }

  private draw(): void {
    if (this.frame) {
      cancelAnimationFrame(this.frame);
      this.frame = 0;
    }
    if (this.destroyed || this.width <= 0 || this.height <= 0) return;
    const context = this.canvas.getContext('2d');
    if (!context) return;
    const state = this.state();
    const conditional = this.conditionalOf(state.sheet);
    const floating = this.drawingsOf(state.sheet);
    // A picture asked for in this frame is never dropped to make room for another.
    if (this.images) this.images.pass++;
    renderGrid(context, {
      workbook: this.workbook,
      sheet: state.sheet,
      layout: this.layout(state),
      selection: { active: state.active, range: state.range },
      locale: this.locale,
      devicePixelRatio: this.dpr,
      conditional,
      filter: this.gridFilter(state.sheet),
      dropdown: this.activeInfo().list ? state.active : undefined,
      cellMark: this.markLookup(),
      drawings: floating,
      // Asked for every visible cell: only for a sheet that has comments.
      comments: floating && state.sheet.comments && state.sheet.comments.length > 0 ? floating : undefined,
      // What the formula being typed was pointed at, on the sheet it is on.
      reference: this.edit && this.point && this.point.sheet === state.sheet ? this.point.range : undefined,
    });
    // Rules whose formulas took too long are finished in the next frames.
    if (conditional && conditional.pending) this.schedule();
    // The input follows its cell through every scroll and resize.
    this.placeInput();
    this.placeTip();
    this.placeComment();
    this.placeNote();
  }

  // -------------------------------------------------- drawings and comments

  /** What floats above the cells of a sheet; undefined for a sheet without drawings and comments, which then costs nothing to draw. */
  private drawingsOf(sheet: Sheet): SheetDrawings | undefined {
    if (sheet.kind !== 'worksheet' || SheetDrawings.isEmpty(sheet)) return undefined;
    let found = this.drawings.get(sheet);
    if (!found) {
      const media = this.workbook.media;
      if (!this.images && media && typeof media.readPart === 'function') {
        // Pictures are read from the file and decoded when they first come into view.
        this.images = new ImageStore(media, coarsePointer() ? IMAGE_BUDGET_BYTES_COARSE : IMAGE_BUDGET_BYTES, () => this.schedule());
      }
      found = new SheetDrawings(sheet, { images: this.images, labels: this.drawingLabels() });
      this.drawings.set(sheet, found);
    }
    return found;
  }

  private drawingLabels(): { chart: string; diagram: string; object: string; picture: string } {
    const name = this.strings.placeholder;
    return { chart: name('chart'), diagram: name('diagram'), object: name('object'), picture: name('picture') };
  }

  /** The comment of a cell (of the merged range it belongs to), with the cell that carries it. */
  private commentOf(row: number, col: number): { row: number; col: number; comment: SheetComment } | undefined {
    const state = this.state();
    const layer = this.drawingsOf(state.sheet);
    if (!layer) return undefined;
    const merge = mergeIn(this.merges(), row, col);
    const at = merge ? { row: merge.top, col: merge.left } : { row, col };
    const comment = layer.commentAt(at.row, at.col);
    return comment ? { row: at.row, col: at.col, comment } : undefined;
  }

  /** What a comment says, as one line for assistive technology. */
  private commentText(comment: SheetComment): string {
    let text = comment.text;
    if (comment.author !== undefined && comment.author !== '' && text.indexOf(comment.author) !== 0) text = `${comment.author}: ${text}`;
    for (const reply of comment.replies || []) text += ` ${reply.author !== undefined && reply.author !== '' ? `${reply.author}: ` : ''}${reply.text}`;
    text = text.replace(/\s+/g, ' ').trim();
    return text.length > MAX_SPOKEN_COMMENT ? `${text.slice(0, MAX_SPOKEN_COMMENT)}\u2026` : text;
  }

  /** Fills the box with a comment. Everything from the file goes in as text, never as markup. */
  private fillComment(box: HTMLDivElement, comment: SheetComment): void {
    while (box.firstChild) box.removeChild(box.firstChild);
    const line = (className: string, text: string, parent: HTMLElement = box): HTMLElement => {
      const el = document.createElement(className === 'bym-comment-reply' ? 'div' : 'span');
      if (className !== '') el.className = className;
      el.textContent = text;
      parent.appendChild(el);
      return el;
    };
    const message = (author: string | undefined, text: string, runs: SheetComment['runs'], parent: HTMLElement) => {
      // A legacy note usually starts with the name of who wrote it; it is not said twice.
      if (author !== undefined && author !== '' && text.indexOf(author) !== 0) line('bym-comment-author', `${author}:\n`, parent);
      if (runs && runs.length > 0) {
        let rest = '';
        for (let i = 0; i < runs.length; i++) {
          const font = runs[i].font;
          // Bold runs stay bold; beyond the limit the text goes on as one plain piece.
          if (i < MAX_COMMENT_SPANS) line(font && font.bold ? 'bym-comment-author' : '', runs[i].text, parent);
          else rest += runs[i].text;
        }
        if (rest !== '') line('', rest, parent);
      } else line('', text, parent);
    };
    message(comment.author, comment.text, comment.threaded ? undefined : comment.runs, box);
    for (const reply of comment.replies || []) message(reply.author, reply.text, undefined, line('bym-comment-reply', ''));
  }

  /** Puts the comment of the cell the pointer is over, or else of the active cell, beside that cell; takes the box away when there is none. */
  private placeComment(): void {
    const state = this.state();
    // A sheet without comments costs nothing here.
    if (!this.commentShown && !(state.sheet.comments && state.sheet.comments.length > 0)) return;
    let target: { row: number; col: number; comment: SheetComment } | undefined;
    // Not in the way of an editor, a menu or a list.
    if (!this.edit && !this.panel && !this.menu && !this.drag && !this.noteEdit) {
      const hover = this.hoverComment;
      if (hover) target = this.commentOf(hover.row, hover.col);
      if (!target) {
        const dismissed = this.commentDismissed;
        const active = this.commentOf(state.active.row, state.active.col);
        if (active && !(dismissed && dismissed.sheet === state.sheet && dismissed.row === active.row && dismissed.col === active.col)) target = active;
      }
      // A box that is always shown is drawn on the sheet already.
      if (target && target.comment.visible) target = undefined;
    }
    let rect: { x: number; y: number; width: number; height: number } | undefined;
    if (target) {
      const layout = this.layout(state);
      const merge = mergeIn(this.merges(), target.row, target.col);
      rect = rangeRect(layout, merge || { top: target.row, left: target.col, bottom: target.row, right: target.col });
      const paneLeft = layout.metrics.headerWidth + (target.col < layout.frozenCols ? 0 : layout.frozenWidth);
      const paneTop = layout.metrics.headerHeight + (target.row < layout.frozenRows ? 0 : layout.frozenHeight);
      const inView = rect.width > 0 && rect.height > 0 && rect.x + rect.width > paneLeft && rect.y + rect.height > paneTop && rect.x < this.width && rect.y < this.height;
      if (!inView) rect = undefined;
    }
    if (!target || !rect) {
      if (this.commentBox && this.commentBox.parentNode) this.commentBox.parentNode.removeChild(this.commentBox);
      this.commentShown = undefined;
      return;
    }
    let box = this.commentBox;
    if (!box) {
      box = document.createElement('div');
      box.className = 'bym-comment';
      box.setAttribute('role', 'note');
      box.addEventListener('mouseleave', this.onCommentLeave);
      this.commentBox = box;
    }
    if (!box.parentNode) {
      box.style.cssText = 'position:absolute;z-index:1;left:0;top:0;box-sizing:border-box;pointer-events:none';
      this.root.appendChild(box);
    }
    const shown = this.commentShown;
    if (!shown || shown.comment !== target.comment) {
      this.fillComment(box, target.comment);
      box.setAttribute('aria-label', this.strings.comment);
      box.scrollTop = 0;
    }
    this.commentShown = { sheet: state.sheet, row: target.row, col: target.col, comment: target.comment };
    // Inside the view whatever the text is: long words break, long text scrolls.
    box.style.maxWidth = `${Math.max(Math.min(COMMENT_MAX_WIDTH, this.width - 8), 0)}px`;
    box.style.maxHeight = `${Math.max(Math.min(COMMENT_MAX_HEIGHT, this.height - 8), 0)}px`;
    const width = box.offsetWidth;
    const height = box.offsetHeight;
    // Text that does not fit can be scrolled; a box that shows everything never takes a press from the cells under it.
    box.style.pointerEvents = box.scrollHeight > box.clientHeight + 1 ? 'auto' : 'none';
    // To the right of the cell as Excel has it; to its left when there is no room; else wherever it fits.
    const right = rect.x + rect.width + 8;
    const leftSide = rect.x - width - 8;
    const left = right + width <= this.width - 4 ? right : leftSide >= 4 ? leftSide : clamp(right, 4, Math.max(this.width - width - 4, 4));
    const top = clamp(rect.y - 4, 4, Math.max(this.height - height - 4, 4));
    box.style.left = `${Math.round(left)}px`;
    box.style.top = `${Math.round(top)}px`;
  }

  // ------------------------------------------------------------------- notes

  /**
   * Sets the note of a cell of the shown sheet, as one undo step: `text` becomes its text (a
   * cell without a note gets one by the `author` of the options), and without text (undefined,
   * empty, only white space) the note is deleted. In a merged cell the note is the one of its
   * top-left cell. Does nothing in a read-only view. Throws RangeError for a text that is too
   * long and for a cell with a threaded comment, which is not edited here.
   */
  setComment(row: number, col: number, text: string | undefined): void {
    if (this.destroyed || this.locked) return;
    const r = toIndex(row, MAX_ROWS);
    const c = toIndex(col, MAX_COLS);
    this.commitEditing();
    this.finishNote(true, true);
    this.applyNote(r, c, text);
  }

  /**
   * Opens the field the note of a cell is typed in, beside the cell (the active one without
   * arguments), with the text of the note the cell has. Ctrl+Enter, or leaving the field,
   * applies what was typed; Escape gives it up. Returns false when there is no field (a
   * read-only view, a sheet without cells). Throws RangeError for a cell with a threaded comment.
   */
  editComment(row?: number, col?: number): boolean {
    if (this.destroyed || this.locked) return false;
    if (this.state().sheet.kind !== 'worksheet') return false;
    this.commitEditing();
    this.finishRename(true, true);
    this.finishNote(true, true);
    this.closeMenu();
    const state = this.state();
    const r = row === undefined ? state.active.row : toIndex(row, MAX_ROWS);
    const c = col === undefined ? state.active.col : toIndex(col, MAX_COLS);
    const problem = noteProblem(this.workbook, this.sheetIndex, r, c);
    if (problem) throw new RangeError(problem);
    // The cell is selected and brought into view: the field stands beside it.
    this.collapseTo({ row: r, col: c }, true);
    const cell = this.state().active;
    const found = commentAtCell(state.sheet, cell.row, cell.col);
    const strings = this.strings;
    const coarse = coarsePointer();
    const box = document.createElement('div');
    box.className = 'bym-note-edit';
    box.style.cssText = 'position:absolute;z-index:2;left:0;top:0;box-sizing:border-box';
    const input = document.createElement('textarea');
    input.className = 'bym-note-input';
    input.setAttribute('aria-label', `${strings.noteField} ${cellName(cell.row, cell.col)}`);
    input.title = strings.noteHint;
    input.maxLength = MAX_NOTE_TEXT;
    input.setAttribute('autocomplete', 'off');
    // A field with smaller text makes a phone zoom into the page.
    if (coarse) input.style.fontSize = '16px';
    // The text of a note comes from the file: it is only ever the value of the field.
    input.value = found ? found.comment.text.replace(/\r\n?/g, '\n') : '';
    box.appendChild(input);
    this.root.appendChild(box);
    input.addEventListener('keydown', this.onNoteKey);
    input.addEventListener('input', this.onNoteInput);
    input.addEventListener('blur', this.onNoteBlur);
    input.addEventListener('compositionstart', this.onFieldComposition);
    input.addEventListener('compositionend', this.onFieldComposition);
    // As the field holds it: a text field turns every kind of line break into "\n".
    this.noteEdit = { sheet: state.sheet, row: cell.row, col: cell.col, box, input, original: input.value, busy: false };
    this.placeNote();
    this.emitEditing();
    input.focus({ preventScroll: true });
    try {
      input.setSelectionRange(input.value.length, input.value.length);
    } catch {
      // Not rendered; the caret is wherever the browser put it.
    }
    this.schedule();
    return true;
  }

  /** Whether a note is being typed in its field. */
  get isEditingComment(): boolean {
    return this.noteEdit !== undefined;
  }

  /** Applies a note to a cell of the shown sheet as one undo step, and says so when that deleted it. */
  private applyNote(row: number, col: number, text: string | undefined): void {
    const sheet = this.state().sheet;
    const had = commentAtCell(sheet, row, col) !== undefined;
    const done = this.applyOps(setCommentOps(this.workbook, this.sheetIndex, row, col, text, { author: this.options.author }));
    // Nothing is left on the cell to be read out: that the note went is said in its place.
    if (done && had && !commentAtCell(sheet, row, col)) this.live.textContent = this.strings.noteDeleted;
  }

  /**
   * Ends typing a note. With `apply` the cell gets the text of the field; a text it cannot
   * take is said, and the field stays open with it - unless the field is being left
   * (`leaving`: the focus went elsewhere, another sheet is shown, another change is made).
   */
  private finishNote(apply: boolean, leaving = false): void {
    const note = this.noteEdit;
    if (!note || note.busy) return;
    const focused = (note.input.getRootNode() as Document | ShadowRoot).activeElement === note.input;
    // The sheet of the note left the workbook under the view, or is no longer the one shown: there is no cell to write.
    const here = this.workbook.sheets[this.sheetIndex] === note.sheet;
    // Text that was only looked at is no change, whatever the note holds (runs, other line breaks, no text at all).
    if (apply && here && !this.locked && note.input.value !== note.original) {
      note.busy = true;
      try {
        this.applyNote(note.row, note.col, note.input.value);
      } catch (error) {
        note.busy = false;
        const message = this.localizeError(error);
        if (!leaving) {
          note.input.setAttribute('aria-invalid', 'true');
          this.notify(message, 'error');
          return;
        }
        this.closeNote(false);
        // What made the field be left (a press on another cell) says something of its own right after.
        this.notifyLater(message, 'error');
        return;
      }
      note.busy = false;
    }
    this.closeNote(!leaving && focused);
  }

  /** Takes the field of the note away, and gives the focus back to the cells when asked to. */
  private closeNote(focusGrid: boolean): void {
    const note = this.noteEdit;
    if (!note) return;
    // Forgotten first: taking the field away may report that it lost the focus.
    this.noteEdit = undefined;
    const { input, box } = note;
    input.removeEventListener('keydown', this.onNoteKey);
    input.removeEventListener('input', this.onNoteInput);
    input.removeEventListener('blur', this.onNoteBlur);
    input.removeEventListener('compositionstart', this.onFieldComposition);
    input.removeEventListener('compositionend', this.onFieldComposition);
    if (box.parentNode) box.parentNode.removeChild(box);
    if (this.destroyed) return;
    if (focusGrid) this.focusInput(false);
    this.schedule();
    this.emitEditing();
  }

  /** Keeps the field of the note beside its cell, inside the view, through scrolling and resizing. */
  private placeNote(): void {
    const note = this.noteEdit;
    if (!note) return;
    const state = this.state();
    if (state.sheet !== note.sheet) return;
    const coarse = coarsePointer();
    const width = Math.max(Math.min(coarse ? NOTE_FIELD_WIDTH_COARSE : NOTE_FIELD_WIDTH, this.width - 8), 0);
    const height = Math.max(Math.min(coarse ? NOTE_FIELD_HEIGHT_COARSE : NOTE_FIELD_HEIGHT, this.height - 8), 0);
    const merge = mergeIn(this.merges(), note.row, note.col);
    const rect = rangeRect(this.layout(state), merge || { top: note.row, left: note.col, bottom: note.row, right: note.col });
    // To the right of the cell as the box of the note is; to its left when there is no room; else wherever it fits.
    const right = rect.x + rect.width + 8;
    const leftSide = rect.x - width - 8;
    const left = right + width <= this.width - 4 ? right : leftSide >= 4 ? leftSide : clamp(right, 4, Math.max(this.width - width - 4, 4));
    const top = clamp(rect.y - 4, 4, Math.max(this.height - height - 4, 4));
    const style = note.box.style;
    style.width = `${width}px`;
    style.height = `${height}px`;
    style.left = `${Math.round(left)}px`;
    style.top = `${Math.round(top)}px`;
  }

  private onNoteKey = (event: KeyboardEvent): void => {
    // Saving is the business of whoever holds the view: the key goes up to it (mount() applies the note first).
    if (isSaveKey(event)) return;
    // The other keys of the field are its own (Enter is a new line, Ctrl+Z undoes what was typed in it): none reaches the grid.
    event.stopPropagation();
    // Enter or Escape that belongs to an input method is not for the note.
    if (this.destroyed || event.isComposing || event.keyCode === 229) return;
    if (event.key === 'Enter' && (event.ctrlKey || event.metaKey) && !event.altKey) {
      event.preventDefault();
      this.finishNote(true);
    } else if (event.key === 'Escape') {
      event.preventDefault();
      this.finishNote(false);
    } else if ((event.ctrlKey || event.metaKey) && !event.altKey && !event.shiftKey && (event.key === 'p' || event.key === 'P')) {
      // As in the grid: the browser would print the host page.
      event.preventDefault();
      this.print().catch(() => undefined);
    }
  };

  private onNoteInput = (event: Event): void => {
    // Another text is being typed: what was said about the last one no longer holds.
    (event.target as HTMLTextAreaElement).removeAttribute('aria-invalid');
  };

  private onNoteBlur = (event: FocusEvent): void => {
    const note = this.noteEdit;
    if (this.destroyed || !note || note.input !== event.target) return;
    // The focus only left the page (another window, another tab): the field is still there when it comes back.
    const root = note.input.getRootNode() as Document | ShadowRoot;
    if (root.activeElement === note.input) return;
    this.finishNote(true, true);
  };

  /**
   * A composition in the field of a note or of a sheet name is not one of the grid: it stays in
   * the field. (The grid would note that an input method is composing, and a field that is taken
   * away in the middle of one never reports its end.)
   */
  private onFieldComposition = (event: Event): void => {
    event.stopPropagation();
  };

  /** Notes the cell under the mouse pointer when it carries a comment; the box follows in the next frame. */
  private hoverAt(x: number, y: number): void {
    let next: CellPos | undefined;
    const state = this.state();
    if (this.drawingsOf(state.sheet)) {
      const hit = hitTest(this.layout(state), x, y);
      if (hit.area === 'cell') {
        const found = this.commentOf(hit.row, hit.col);
        if (found) next = { row: found.row, col: found.col };
      }
    }
    const before = this.hoverComment;
    if ((before === undefined) === (next === undefined) && (!before || !next || (before.row === next.row && before.col === next.col))) return;
    this.hoverComment = next;
    this.schedule();
  }

  private onHoverMove = (event: MouseEvent): void => {
    if (this.destroyed || this.drag || this.touch) return;
    const point = this.canvasPoint(event);
    this.hoverAt(point.x, point.y);
  };

  private onPointerLeave = (event: MouseEvent): void => {
    if (!this.hoverComment) return;
    // Onto the box of the comment itself (to scroll it): it stays.
    const to = event.relatedTarget as Node | null;
    if (to && this.commentBox && this.commentBox.contains(to)) return;
    this.hoverComment = undefined;
    this.schedule();
  };

  private onCommentLeave = (event: MouseEvent): void => {
    if (!this.hoverComment || event.relatedTarget === this.canvas) return;
    this.hoverComment = undefined;
    this.schedule();
  };

  // ------------------------------------------------- validation and marks

  /** The data validations of a sheet; undefined for a sheet without any, which then costs nothing. */
  private validationOf(sheet: Sheet): DataValidations | undefined {
    if (!sheet.dataValidations || sheet.dataValidations.length === 0) return undefined;
    let found = this.validation.get(sheet);
    if (!found) {
      found = new DataValidations(this.workbook, sheet, { engine: () => this.conditionalEngine(), locale: this.locale });
      this.validation.set(sheet, found);
    }
    return found;
  }

  /** What the active cell's data validation gives it: a list to choose from, a message. */
  private activeInfo(): CellInfo {
    const state = this.state();
    const { row, col } = state.active;
    let info = this.cellInfo;
    if (info && info.sheet === state.sheet && info.row === row && info.col === col) return info;
    info = { sheet: state.sheet, row, col, list: false, prompt: undefined };
    const rules = state.sheet.kind === 'worksheet' ? this.validationOf(state.sheet) : undefined;
    if (rules) {
      try {
        // A list cannot be chosen from in a read-only view; its message can still be read.
        info.list = !this.locked && rules.hasList(row, col);
        info.prompt = rules.prompt(row, col);
      } catch (error) {
        // A rule that cannot be read gives the cell nothing.
      }
    }
    this.cellInfo = info;
    return info;
  }

  /** Whether a canvas point is on the list button of the active cell. */
  private dropdownAt(x: number, y: number, reach: number): boolean {
    if (!this.activeInfo().list) return false;
    const state = this.state();
    const { active } = state;
    const merge = mergeIn(this.merges(), active.row, active.col);
    const cell = rangeRect(this.layout(state), merge || { top: active.row, left: active.col, bottom: active.row, right: active.col });
    const box = cellButtonRect(cell.x, cell.y, cell.width, cell.height);
    return box !== undefined && x >= box.x - reach && y >= box.y - reach && x <= box.x + box.size + reach && y <= box.y + box.size + reach;
  }

  /**
   * Opens the list the active cell's data validation offers; choosing a value writes it
   * to the cell as one undo step. Returns false when the cell has no list.
   */
  openValidationList(touch = false): boolean {
    if (this.destroyed || this.locked) return false;
    // First of all: what is being typed is written, which shows the sheet of its cell again.
    if (!this.tryCommit()) return false;
    const state = this.state();
    const sheet = state.sheet;
    const rules = sheet.kind === 'worksheet' ? this.validationOf(sheet) : undefined;
    const { row, col } = state.active;
    if (!rules || !rules.hasList(row, col)) return false;
    const list = rules.list(row, col);
    if (!list) return false;
    this.closeMenu();
    if (!this.tryCommit()) return false;
    const strings = this.strings;
    const merge = mergeIn(this.merges(), row, col);
    const cell = rangeRect(this.layout(state), merge || { top: row, left: col, bottom: row, right: col });
    const x = clamp(cell.x, 0, Math.max(this.width - 1, 0));
    const y = clamp(cell.y, 0, Math.max(this.height - 1, 0));
    const anchor = { x, y, width: clamp(cell.x + cell.width - x, 0, this.width - x), height: clamp(cell.y + cell.height - y, 0, this.height - y) };
    const shown = cellDisplayText(this.workbook, getCell(sheet, row, col), this.locale);
    this.panel = new ListMenu({
      parent: this.root,
      anchor,
      label: strings.chooseFromList,
      items: list.items,
      selected: list.items.indexOf(shown),
      note: list.truncated ? strings.tooManyValues(MAX_LIST_ITEMS) : undefined,
      openedByTouch: touch,
      onChoose: (index) =>
        this.attempt(() => {
          // The list belongs to the sheet it was opened on, wherever that sheet is by now.
          if (this.state().sheet === sheet) this.applyOps(setCellInput(this.workbook, this.sheetIndex, row, col, list.items[index], { locale: this.locale }));
        }),
      onClose: () => this.panelClosed(),
    });
    // The message beside the cell makes room for the list.
    this.placeTip();
    return true;
  }

  /**
   * The data validation of a cell of the shown sheet as it applies now: its ranges where they
   * are, its formulas (as files store them) written for the first cell of them. Undefined for
   * a cell without a rule.
   */
  validationAt(row: number, col: number): DataValidation | undefined {
    if (this.destroyed) return undefined;
    const sheet = this.state().sheet;
    const rules = sheet.kind === 'worksheet' ? this.validationOf(sheet) : undefined;
    return rules ? rules.ruleAt(toIndex(row, MAX_ROWS), toIndex(col, MAX_COLS)) : undefined;
  }

  /**
   * Sets the data validation of a range of the shown sheet, as one undo step: afterwards
   * `rule` is the rule of every cell of the range. The cells are taken out of the rules they
   * were in, which stay for their other cells. Without a rule (undefined, or "any value" with
   * no input message) the cells are left without one. What is being typed is applied first;
   * nothing is done while that cannot be applied, or while a dialog of rules is open in which
   * something was changed (see commitPending()). Does nothing in a read-only view. Throws RangeError for a rule that cannot be stated (see
   * setValidation() of the operations) and for cells under a rule of the extension list.
   */
  setValidation(range: CellRange, rule: ValidationSettings | undefined): void {
    if (this.destroyed || this.locked) return;
    if (!this.commitPending()) return;
    this.applyOps(setValidationOps(this.workbook, this.sheetIndex, range, rule));
  }

  /**
   * Opens the dialog that sets the data validation of the selection. It shows the rule the
   * active cell has; OK makes what it says the rule of every selected cell, as one undo step.
   * What is being typed is applied first. Returns false when there is no dialog: a read-only
   * view, a sheet without cells, a formula that is being typed, something typed that cannot
   * be applied, or cells whose rule cannot be changed (said in the live region and to onNotice).
   */
  openValidationDialog(): boolean {
    if (this.destroyed || this.locked) return false;
    if (this.state().sheet.kind !== 'worksheet') return false;
    if (this.formulaEdit()) {
      this.notify(this.strings.rules.formulaBeingTyped, 'error');
      return false;
    }
    // A tab that is being dragged is let go where it was: the dialog takes every pointer from here on.
    this.endTabDrag();
    this.closeMenu();
    if (!this.tryCommit()) return false;
    this.finishRename(true);
    this.finishNote(true);
    if (this.edit || this.renaming || this.noteEdit) return false;
    const state = this.state();
    const sheet = state.sheet;
    const range = { ...state.range };
    const strings = this.strings;
    try {
      const problem = validationProblem(this.workbook, this.sheetIndex, range);
      if (problem) throw new RangeError(problem);
    } catch (error) {
      this.notify(this.localizeError(error), 'error');
      return false;
    }
    const date1904 = this.workbook.date1904 === true;
    const found = validationsIn(this.workbook, this.sheetIndex, range);
    // One rule that has every cell of the selection is what the dialog shows; anything else is said.
    const mixed = found.length > 1 || (found.length === 1 && subtractRanges([range], found[0].ranges).length > 0);
    const shown = this.validationAt(state.active.row, state.active.col) || found[0];
    const initial = validationFields(shown, this.locale, date1904);
    const apply = (rule: ValidationSettings | undefined) => {
      // The dialog belongs to the sheet it was opened on.
      if (this.locked || this.workbook.sheets[this.sheetIndex] !== sheet) throw new RangeError('The sheet of the dialog is not shown');
      const had = validationsIn(this.workbook, this.sheetIndex, range).length > 0;
      const done = this.applyFromDialog(setValidationOps(this.workbook, this.sheetIndex, range, rule));
      if (!done) return;
      const left = validationsIn(this.workbook, this.sheetIndex, range).length > 0;
      this.live.textContent = left ? strings.rules.validationSet : had ? strings.rules.validationCleared : '';
    };
    const dialog = new ValidationDialog({
      parent: this.options.focusScope || this.root,
      strings: strings.rules,
      locale: this.locale,
      fields: initial,
      mixed,
      coarse: coarsePointer(),
      onApply: (fields) => {
        const reading = { date1904, isName: (name: string) => hasDefinedName(this.workbook, name) };
        // What was only looked at is not changed: no step, and the rule stays the element of the file it is.
        const changed = validationChange(fields, { fields: initial, rule: shown }, this.locale, strings.rules, reading);
        if (changed) apply(changed);
        // A selection of several rules gets this one, whether a field was touched or not.
        else if (mixed) apply(validationFromFields(fields, this.locale, strings.rules, reading));
      },
      onClear: () => apply(undefined),
      explain: (error) => this.localizeError(error),
      onClose: () => {
        this.rulesDialog = undefined;
        // Back to the cells, unless the view is going away.
        if (!this.destroyed) this.focusInput(true);
      },
    });
    this.rulesDialog = { sheet, submit: () => dialog.dialog.submit(), close: () => dialog.dialog.close(), changed: () => dialog.dialog.changed, say: (message) => dialog.dialog.showError(message) };
    return true;
  }

  /** The words of the dialogs and menus of rules, in the language of the view: for a shell that offers their commands. */
  get ruleStrings(): UiStrings['rules'] {
    return this.strings.rules;
  }

  /** Whether a dialog of rules is open over the view. */
  get isDialogOpen(): boolean {
    return this.rulesDialog !== undefined;
  }

  // ------------------------------------------------- conditional formats

  /**
   * The conditional rules of the shown sheet in the order they are applied in, each as it
   * applies now; with a range, only those that apply to a cell of it. `index` is what the
   * methods below name a rule by.
   */
  conditionalRules(range?: CellRange): ConditionalRuleInfo[] {
    if (this.destroyed) return [];
    return conditionalRulesOf(this.workbook, this.sheetIndex, range);
  }

  /**
   * Adds a conditional rule for cells of the shown sheet, as one undo step. The new rule is
   * the first to be applied (in a sheet with a rule only a newer Excel states: the last).
   * What is being typed is applied first. Does nothing in a read-only view. Throws RangeError
   * for a rule that cannot be stated.
   */
  addConditionalRule(range: CellRange | CellRange[], rule: ConditionalSettings): void {
    if (!this.beginRules()) return;
    this.applyOps(addConditionalRuleOps(this.workbook, this.sheetIndex, range, rule));
  }

  /**
   * Changes the rule at `index`: the cells it applies to, whether it stops the rules behind
   * it, what it says and the format it gives. Throws RangeError for an index that names no
   * rule and for a rule that is kept as it is.
   */
  updateConditionalRule(index: number, patch: ConditionalPatch): void {
    if (!this.beginRules()) return;
    this.applyOps(updateConditionalRuleOps(this.workbook, this.sheetIndex, index, patch));
  }

  /** Deletes the rule at `index`. Throws RangeError for an index that names no rule and for a rule that is kept as it is. */
  removeConditionalRule(index: number): void {
    if (!this.beginRules()) return;
    this.applyOps(removeConditionalRuleOps(this.workbook, this.sheetIndex, index));
  }

  /** Moves the rule at `index` so that it is applied as number `to` (0: first). */
  moveConditionalRule(index: number, to: number): void {
    if (!this.beginRules()) return;
    this.applyOps(moveConditionalRuleOps(this.workbook, this.sheetIndex, index, to));
  }

  /** Takes the cells of a range out of every conditional rule of the shown sheet; without a range, deletes all its rules. */
  clearConditionalRules(range?: CellRange): void {
    if (!this.beginRules()) return;
    this.applyOps(clearConditionalRulesOps(this.workbook, this.sheetIndex, range));
  }

  /**
   * What every command on rules starts with: nothing in a read-only view, what is typed applied, and no dialog
   * of rules left open - one in which nothing was changed is closed, one with changes refuses the command.
   */
  private beginRules(): boolean {
    if (this.destroyed || this.locked) return false;
    return this.commitPending();
  }

  /**
   * Opens a dialog of conditional formats over the view. `what` says which: "manage" the list
   * of rules (of the selection, or of the sheet); a number the dialog of that rule (see
   * conditionalRules()); anything else the dialog of a new rule for the selection - a kind of
   * rule ("cellIs", "containsText", "top10", "colorScale", "dataBar", "iconSet", "timePeriod", ...) or one of the
   * comparisons "greaterThan", "lessThan", "between", "equal". What is being typed is applied
   * first. Returns false when there is no dialog: a read-only view, a sheet without cells, a
   * formula that is being typed, something typed that cannot be applied, a rule that cannot
   * be changed (said in the live region and to onNotice).
   */
  openConditionalDialog(what?: string | number): boolean {
    if (!this.beginDialog()) return false;
    if (what === 'manage') return this.openRulesManager('selection');
    return this.openRuleDialog(what, undefined);
  }

  /** What every dialog of rules starts with; false when none can be opened. */
  private beginDialog(): boolean {
    if (this.destroyed || this.locked) return false;
    if (this.state().sheet.kind !== 'worksheet') return false;
    if (this.formulaEdit()) {
      this.notify(this.strings.rules.formulaBeingTyped, 'error');
      return false;
    }
    // A tab that is being dragged is let go where it was: the dialog takes every pointer from here on.
    this.endTabDrag();
    this.closeMenu();
    if (!this.tryCommit()) return false;
    this.finishRename(true);
    this.finishNote(true);
    return !this.edit && !this.renaming && !this.noteEdit;
  }

  /** Applies operations on rules for a dialog that stays open, and says what happened. */
  private applyRules(sheet: Sheet, ops: Op[], said: string): void {
    // A dialog belongs to the sheet it was opened on.
    if (this.locked || this.workbook.sheets[this.sheetIndex] !== sheet) throw new RangeError('The sheet of the dialog is not shown');
    if (this.applyFromDialog(ops)) this.live.textContent = said;
  }

  /** applyOps() for the dialog of rules that is open: its own change is the one that does not have to settle it first. */
  private applyFromDialog(ops: Op[]): boolean {
    const nested = this.applyingRules;
    this.applyingRules = true;
    try {
      return this.applyOps(ops);
    } finally {
      this.applyingRules = nested;
    }
  }

  /** The dialog of one rule: a new one of the kind `what` names, or the rule at that index. `back` is the list to come back to. */
  private openRuleDialog(what: string | number | undefined, back: 'selection' | 'sheet' | undefined): boolean {
    const state = this.state();
    const sheet = state.sheet;
    const range = { ...state.range };
    const strings = this.strings;
    const date1904 = this.workbook.date1904 === true;
    const styles = this.workbook.styles;
    let editing: ConditionalRuleInfo | undefined;
    let type: string | undefined;
    let operator: string | undefined;
    if (typeof what === 'number') {
      editing = conditionalRulesOf(this.workbook, this.sheetIndex)[what];
      if (!editing || !editing.editable) {
        this.notify(editing && editing.rule.extended ? strings.rules.extensionRuleKept : editing && editing.rule.extId !== undefined ? strings.rules.extensionRuleSays : strings.rules.ruleKindKept, 'error');
        return false;
      }
    } else if (what === 'greaterThan' || what === 'lessThan' || what === 'between' || what === 'equal') {
      type = 'cellIs';
      operator = what;
    } else if (typeof what === 'string' && (CONDITIONAL_TYPES as readonly string[]).indexOf(what) >= 0) type = what;
    const fields = conditionalFields(editing, this.locale, date1904, type, (color) => resolveColor(color, styles));
    if (operator !== undefined) fields.operator = operator;
    const index = editing ? editing.index : -1;
    const cells = rangeListText(editing ? editing.ranges : [range], this.locale);
    const dialog = new ConditionalDialog({
      parent: this.options.focusScope || this.root,
      strings: strings.rules,
      locale: this.locale,
      fields,
      ranges: cells,
      // A rule that is edited is looked at first; a new one starts where its value is typed.
      focus: editing ? 'type' : undefined,
      keepFormat: fields.preset === 'keep',
      coarse: coarsePointer(),
      onApply: (typed, where) => {
        const reading = { date1904, isName: (name: string) => hasDefinedName(this.workbook, name) };
        if (editing) {
          // Only what was changed in the dialog is changed in the rule; nothing changed is no step.
          const patch = conditionalPatch(typed, where, { fields, ranges: cells }, editing, this.locale, strings.rules, reading);
          if (Object.keys(patch).length > 0) this.applyRules(sheet, updateConditionalRuleOps(this.workbook, this.sheetIndex, index, patch), strings.rules.ruleChanged);
          return;
        }
        const ranges = this.rangesOf(where);
        this.applyRules(sheet, addConditionalRuleOps(this.workbook, this.sheetIndex, ranges, conditionalFromFields(typed, this.locale, strings.rules, reading)), strings.rules.ruleAdded);
      },
      explain: (error) => this.localizeError(error),
      onClose: () => {
        this.rulesDialog = undefined;
        if (this.destroyed) return;
        // Back to the list it was opened from, or to the cells.
        if (back !== undefined && !this.locked && this.workbook.sheets[this.sheetIndex] === sheet && !this.leavingDialog) this.openRulesManager(back);
        else this.focusInput(true);
      },
    });
    this.rulesDialog = { sheet, submit: () => dialog.dialog.submit(), close: () => this.dismiss(() => dialog.dialog.close()), changed: () => dialog.dialog.changed, say: (message) => dialog.dialog.showError(message) };
    return true;
  }

  /** Closes a dialog for good: it does not hand over to the list it was opened from. */
  private dismiss(close: () => void): void {
    this.leavingDialog = true;
    try {
      close();
    } finally {
      this.leavingDialog = false;
    }
  }

  /** The ranges typed into a field of a dialog; throws FieldError for what names none. */
  private rangesOf(text: string): CellRange[] {
    if (text.trim() === '') throw new FieldError(this.strings.rules.needRange, 'ranges');
    const ranges = parseRangeList(text, this.locale);
    if (!ranges) throw new FieldError(this.strings.rules.badRange, 'ranges');
    return ranges;
  }

  /** The list of the conditional rules of the selection or the sheet. */
  private openRulesManager(scope: 'selection' | 'sheet'): boolean {
    const state = this.state();
    const sheet = state.sheet;
    const range = { ...state.range };
    const strings = this.strings;
    const date1904 = this.workbook.date1904 === true;
    const rows = (shown: 'selection' | 'sheet', from: number): { rows: ManagerRow[]; total: number; before?: ManagerRow; after?: ManagerRow } => {
      if (this.workbook.sheets[this.sheetIndex] !== sheet) return { rows: [], total: 0 };
      const listed = conditionalRulesOf(this.workbook, this.sheetIndex, shown === 'selection' ? range : undefined);
      // No more rows than a list can be worked with at a time: the list shows them in runs of that many.
      const page = managerPage(listed.length, from);
      const row = (info: ConditionalRuleInfo): ManagerRow => {
        // A rule only a newer Excel states is not touched; one it states more about keeps what it says, and can go or go elsewhere.
        const kept = info.rule.extended === true;
        const paired = !kept && info.rule.extId !== undefined;
        return {
          index: info.index,
          description: describeRule(info.rule, strings.rules, this.locale, date1904),
          ranges: rangeListText(info.ranges, this.locale),
          stopIfTrue: info.rule.stopIfTrue === true,
          editable: info.editable,
          removable: info.removable,
          movable: info.movable,
          adjustable: !kept,
          reason: kept ? strings.rules.extensionRuleKept : paired ? strings.rules.extensionRuleSays : !info.movable ? strings.rules.orderFixed : !info.editable ? strings.rules.ruleKindKept : undefined,
        };
      };
      const out: { rows: ManagerRow[]; total: number; before?: ManagerRow; after?: ManagerRow } = { rows: listed.slice(page.from, page.to).map(row), total: listed.length };
      // The neighbours of the first and the last rule shown: what they can be moved above and below.
      if (page.from > 0) out.before = row(listed[page.from - 1]);
      if (page.to < listed.length) out.after = row(listed[page.to]);
      return out;
    };
    const manager = new ConditionalManager({
      parent: this.options.focusScope || this.root,
      strings: strings.rules,
      scope,
      coarse: coarsePointer(),
      rows,
      onMove: (index, to) => this.applyRules(sheet, moveConditionalRuleOps(this.workbook, this.sheetIndex, index, to), strings.rules.ruleChanged),
      onRanges: (index, text) => this.applyRules(sheet, updateConditionalRuleOps(this.workbook, this.sheetIndex, index, { ranges: this.rangesOf(text) }), strings.rules.ruleChanged),
      onStop: (index, stop) => this.applyRules(sheet, updateConditionalRuleOps(this.workbook, this.sheetIndex, index, { stopIfTrue: stop }), strings.rules.ruleChanged),
      onDelete: (index) => this.applyRules(sheet, removeConditionalRuleOps(this.workbook, this.sheetIndex, index), strings.rules.ruleDeleted),
      onEdit: (index, shown) => {
        this.openRuleDialog(index, shown);
      },
      onNew: (shown) => {
        this.openRuleDialog(undefined, shown);
      },
      explain: (error) => this.localizeError(error),
      onClose: () => {
        this.rulesDialog = undefined;
        // Back to the cells, unless the view is going away; a dialog that takes the place of the list takes the focus next.
        if (!this.destroyed) this.focusInput(true);
      },
    });
    this.rulesDialog = { sheet, submit: () => manager.submit(), close: () => manager.dialog.close(), changed: () => manager.pending, say: (message) => manager.dialog.showError(message) };
    return true;
  }

  /**
   * Checks what is about to be typed into a cell against the cell's data validation.
   * A rule that stops wrong values throws a ValidationError with the rule's message;
   * a rule that only warns gives back what to say once the value is written.
   */
  private validateInput(sheetIndex: number, row: number, col: number, text: string, ops: Op[]): string | undefined {
    const sheet = this.workbook.sheets[sheetIndex];
    const rules = sheet ? this.validationOf(sheet) : undefined;
    if (!rules) return undefined;
    // The cell the text turns into; without an operation the cell stays what it is.
    let cell = getCell(sheet, row, col);
    for (const op of ops) {
      if (op.type !== 'setCells' || op.sheet !== sheetIndex) continue;
      for (const change of op.cells) if (change.row === row && change.col === col) cell = change.cell;
    }
    if (cell && cell.formula) {
      // A formula is judged by what it gives: a calculation under way is finished first, so that the
      // cells it reads hold what they come to. One that cannot be finished leaves the formula unjudged.
      const engine = this.conditionalEngine();
      if (engine && engine.incomplete) this.finishCalculation(engine);
    }
    const verdict = rules.check(row, col, cell, text, (run) => this.tryOut(ops, run));
    if (verdict.ok) return undefined;
    const message = verdict.message || this.strings.validationFailed;
    const display = verdict.title ? `${verdict.title}: ${message}` : message;
    if (verdict.style === 'stop' || verdict.style === undefined) throw new ValidationError(display);
    return this.strings.validationWarning(display);
  }

  /**
   * Runs `run` as if operations had been applied and everything that depends on them
   * calculated, and takes it all back afterwards: the formula of a custom rule reads the cell
   * it is asked about, and often formulas that read that cell. Nothing of it is drawn or
   * reported; undefined when the value could not be tried (see tryOps()).
   */
  private tryOut<T>(ops: Op[], run: () => T): T | undefined {
    const engine = this.conditionalEngine();
    // A calculation that is still going on is finished first: what it writes is real, and
    // would otherwise be taken back with the trial.
    if (engine && engine.incomplete) this.finishCalculation(engine);
    try {
      return tryOps(this.workbook, engine, ops, run);
    } finally {
      this.marksDirty = true;
      // A calculation the trial could not finish goes on in the background.
      if (engine && engine.incomplete) this.continueCalculation();
    }
  }

  /** Puts the message of the active cell's data validation beside the cell, or takes it away. */
  private placeTip(): void {
    const prompt = this.activeInfo().prompt;
    const state = this.state();
    const { active } = state;
    const layout = this.layout(state);
    const merge = prompt ? mergeIn(this.merges(), active.row, active.col) : undefined;
    const rect = prompt ? rangeRect(layout, merge || { top: active.row, left: active.col, bottom: active.row, right: active.col }) : undefined;
    const paneLeft = layout.metrics.headerWidth + (active.col < layout.frozenCols ? 0 : layout.frozenWidth);
    const paneTop = layout.metrics.headerHeight + (active.row < layout.frozenRows ? 0 : layout.frozenHeight);
    const inView = rect !== undefined && rect.width > 0 && rect.height > 0 && rect.x + rect.width > paneLeft && rect.y + rect.height > paneTop && rect.x < this.width && rect.y < this.height;
    if (!prompt || !rect || !inView || this.panel) {
      if (this.tip && this.tip.parentNode) this.tip.parentNode.removeChild(this.tip);
      this.tip = undefined;
      return;
    }
    let tip = this.tip;
    if (!tip) {
      tip = document.createElement('div');
      tip.className = 'bym-tip';
      // Never in the way of a press: it is a note, not a control.
      tip.style.cssText = 'position:absolute;z-index:1;left:0;top:0;box-sizing:border-box;pointer-events:none';
      const title = document.createElement('div');
      title.className = 'bym-tip-title';
      const text = document.createElement('div');
      text.className = 'bym-tip-text';
      tip.appendChild(title);
      tip.appendChild(text);
      this.root.appendChild(tip);
      this.tip = tip;
    }
    const title = tip.firstChild as HTMLDivElement;
    const text = tip.lastChild as HTMLDivElement;
    if (title.textContent !== prompt.title) title.textContent = prompt.title;
    if (text.textContent !== prompt.text) text.textContent = prompt.text;
    title.style.display = prompt.title === '' ? 'none' : 'block';
    text.style.display = prompt.text === '' ? 'none' : 'block';
    // Under the cell, a little to the right as Excel has it; above the cell when there is no room below.
    tip.style.maxWidth = `${Math.max(Math.min(240, this.width - 8), 0)}px`;
    const width = tip.offsetWidth;
    const height = tip.offsetHeight;
    const left = clamp(rect.x + Math.min(rect.width / 2, 24), 4, Math.max(this.width - width - 4, 4));
    const below = rect.y + rect.height + 6;
    const top = below + height <= this.height || rect.y - height - 6 < 0 ? below : rect.y - height - 6;
    tip.style.left = `${Math.round(left)}px`;
    tip.style.top = `${Math.round(top)}px`;
  }

  /** The marks of every sheet, built again after the engine calculated. Without an engine there are none. */
  private cellMarks(): (Map<number, string> | undefined)[] | undefined {
    const engine = this.engine;
    if (!engine) return undefined;
    if (!this.marksDirty && this.marks) return this.marks;
    const marks: (Map<number, string> | undefined)[] = [];
    const add = (cells: { sheet: number; row: number; col: number }[], kind: string) => {
      for (let i = 0; i < cells.length; i++) {
        const cell = cells[i];
        let map = marks[cell.sheet];
        if (!map) map = marks[cell.sheet] = new Map();
        map.set(cell.row * MAX_COLS + cell.col, kind);
      }
    };
    try {
      // The later a kind is added, the more it counts where a cell is of several.
      add(engine.stale, 'stale');
      add(engine.mismatched, 'file');
      add(engine.unsupported, 'file');
      add(engine.unparsed, 'file');
      add(engine.cycles(), 'circular');
      add(engine.circular, 'circular');
    } catch (error) {
      // Lists that cannot be read give no marks.
    }
    this.marks = marks;
    this.marksDirty = false;
    return marks;
  }

  /** What the grid asks for the mark of each cell; undefined while the shown sheet has no marks. */
  private markLookup(): ((row: number, col: number) => string | undefined) | undefined {
    const marks = this.cellMarks();
    const map = marks ? marks[this.sheetIndex] : undefined;
    if (!map || map.size === 0) return undefined;
    return (row, col) => {
      const kind = map.get(row * MAX_COLS + col);
      return kind === undefined ? undefined : MARK_COLORS[kind];
    };
  }

  /** What there is to say about the formula of a cell that has a mark; '' for a cell without one. */
  private formulaNote(row: number, col: number): string {
    const marks = this.cellMarks();
    const map = marks ? marks[this.sheetIndex] : undefined;
    const kind = map ? map.get(row * MAX_COLS + col) : undefined;
    if (kind === undefined) return '';
    return kind === 'circular' ? this.strings.formulaCircular : kind === 'stale' ? this.strings.formulaStale : this.strings.formulaNotCalculated;
  }

  /** The conditional formats of a sheet; undefined for a sheet without any, which then costs nothing to draw. */
  private conditionalOf(sheet: Sheet): ConditionalFormats | undefined {
    if (!sheet.conditionalFormats || sheet.conditionalFormats.length === 0) return undefined;
    let found = this.conditional.get(sheet);
    if (!found) {
      found = new ConditionalFormats(this.workbook, sheet, { engine: () => this.conditionalEngine() });
      this.conditional.set(sheet, found);
    }
    return found;
  }

  /**
   * Formulas that came without a result (files written by libraries that do not
   * calculate) would show nothing until something they read is edited; they are
   * calculated once the view is up.
   */
  private calculateMissing(): void {
    if (this.options.calculate === false) return;
    this.missingTimer = setTimeout(() => {
      this.missingTimer = undefined;
      if (this.destroyed) return;
      try {
        // A conditional format or an early edit may have made the engine already.
        const engine = this.engine || new CalcEngine(this.workbook, { budget: { milliseconds: CALC_SLICE_MS } });
        const missing = engine.uncalculated();
        if (missing.length === 0) return;
        this.engine = engine;
        this.marksDirty = true;
        engine.recalculate(missing);
        this.printEpoch++;
        if (engine.incomplete) this.continueCalculation();
        this.conditional.forEach((formats) => formats.invalidate());
        this.schedule();
      } catch (error) {
        this.engine = undefined;
      }
    }, 0);
  }

  /** The engine for the formulas of conditional formats: made on first use, also in a read-only view. */
  private conditionalEngine(): CalcEngine | undefined {
    if (this.options.calculate === false) return undefined;
    if (!this.engine) this.engine = new CalcEngine(this.workbook, { budget: { milliseconds: CALC_SLICE_MS } });
    return this.engine;
  }

  // ------------------------------------------------------- filter and sort

  /** What the grid draws for the filter of a sheet; undefined for a sheet without one. */
  private gridFilter(sheet: Sheet): GridFilter | undefined {
    const filter = sheet.autoFilter;
    // A table brings its own filter, which is only kept; no buttons are shown in such a sheet.
    if (!filter || sheet.kind !== 'worksheet' || (sheet.anchored !== undefined && sheet.anchored.indexOf('table') >= 0)) return undefined;
    let view = this.filterViews.get(filter);
    if (!view) {
      const width = Math.min(Math.max(filter.range.right - filter.range.left + 1, 0), MAX_COLS);
      const buttons = new Uint8Array(width);
      for (let i = 0; i < width; i++) buttons[i] = 1;
      let filtered = false;
      for (const column of filter.columns) {
        if (!(column.colId >= 0 && column.colId < width)) continue;
        if (isFilterActive(column)) filtered = true;
        buttons[column.colId] = column.hiddenButton ? 0 : isFilterActive(column) ? 2 : 1;
      }
      view = { range: filter.range, buttons, filtered };
      this.filterViews.set(filter, view);
    }
    return view;
  }

  /**
   * The column whose filter button a canvas point is on, or -1. A finger gets some
   * room around the button, but never beyond its cell.
   */
  private filterButtonAt(x: number, y: number, reach: number): number {
    const state = this.state();
    const filter = this.gridFilter(state.sheet);
    if (!filter) return -1;
    const layout = this.layout(state);
    const hit = hitTest(layout, x, y);
    if (hit.area !== 'cell' || hit.row !== filter.range.top || hit.col < filter.range.left || hit.col > filter.range.right) return -1;
    if (!filter.buttons[hit.col - filter.range.left]) return -1;
    const cell = rangeRect(layout, { top: hit.row, left: hit.col, bottom: hit.row, right: hit.col });
    const box = cellButtonRect(cell.x, cell.y, cell.width, cell.height);
    if (!box) return -1;
    return x >= box.x - reach && y >= box.y - reach && x <= box.x + box.size + reach && y <= box.y + box.size + reach ? hit.col : -1;
  }

  /** A press on a button inside a cell opens what the button stands for instead of selecting the cell. */
  private pressCellButton(x: number, y: number, touch: boolean): boolean {
    if (this.locked) return false;
    const col = this.filterButtonAt(x, y, touch ? 8 : 0);
    if (col >= 0) return this.openFilterMenu(col, touch);
    if (this.dropdownAt(x, y, touch ? 8 : 0)) return this.openValidationList(touch);
    return false;
  }

  /** Whether the shown sheet has a filter. */
  get hasFilter(): boolean {
    return this.state().sheet.autoFilter !== undefined;
  }

  /** How many rows under its header the filter of the shown sheet shows; undefined without a filter that has criteria. */
  get filterSummary(): { shown: number; total: number } | undefined {
    return filterSummary(this.state().sheet);
  }

  /**
   * What there is to say about the active cell and the sheet while nothing else is
   * being said, in the language of the view; an editor shell shows it in its status bar.
   */
  get statusText(): string {
    const state = this.state();
    const note = this.formulaNote(state.active.row, state.active.col);
    if (note !== '') return note;
    const summary = this.filterSummary;
    return summary ? this.strings.recordsShown(summary.shown, summary.total) : '';
  }

  /** Whether a selection is one cell: a single cell, or exactly one merged cell. */
  private singleCell(): boolean {
    const { range, active } = this.state();
    if (range.top === range.bottom && range.left === range.right) return true;
    const merge = mergeIn(this.merges(), active.row, active.col);
    return merge !== undefined && sameRange(merge, range);
  }

  /**
   * Switches the filter of the sheet on or off. It goes over the selection, or, with one
   * cell selected, over the block of cells around it; its first row becomes the header.
   */
  toggleFilter(): void {
    if (this.destroyed || this.locked) return;
    this.commitEditing();
    const state = this.state();
    const sheet = state.sheet;
    if (sheet.kind !== 'worksheet') return;
    // A dialog of rules with changes is not given up for this: the command is refused (see settleDialog()).
    if (!this.settleDialog()) return;
    this.closeMenu();
    if (sheet.autoFilter) {
      this.applyOps(setAutoFilter(this.workbook, this.sheetIndex, undefined));
      return;
    }
    const range = this.singleCell() ? currentRegion(sheet, state.active.row, state.active.col) : state.range;
    const first = getCell(sheet, range.top, range.left);
    if (range.top === range.bottom && range.left === range.right && (!first || first.value === null || first.value === '')) {
      throw new RangeError('There is no data to filter');
    }
    this.applyOps(setAutoFilter(this.workbook, this.sheetIndex, range));
  }

  /**
   * Sorts by the column of the active cell: inside the filter its rows, with one cell
   * selected the block of cells around it (its first row stays when it looks like a
   * header, see hasHeaderRow()), otherwise exactly the selection.
   */
  sort(descending: boolean): void {
    if (this.destroyed || this.locked) return;
    this.commitEditing();
    const state = this.state();
    const sheet = state.sheet;
    if (sheet.kind !== 'worksheet') return;
    // A dialog of rules with changes is not given up for this: the command is refused (see settleDialog()).
    if (!this.settleDialog()) return;
    this.closeMenu();
    const { active } = state;
    const filter = sheet.autoFilter;
    let range: CellRange;
    let header: boolean;
    if (filter && active.row >= filter.range.top && active.row <= filter.range.bottom && active.col >= filter.range.left && active.col <= filter.range.right) {
      range = filter.range;
      header = true;
    } else if (this.singleCell()) {
      range = currentRegion(sheet, active.row, active.col);
      header = hasHeaderRow(this.workbook, sheet, range);
    } else {
      range = state.range;
      header = false;
    }
    this.applyOps(sortRange(this.workbook, this.sheetIndex, range, [{ col: active.col, descending: descending === true }], { locale: this.locale, header }));
  }

  /**
   * Sets which values a column of the filter shows (as its cells show them), and whether
   * empty cells are shown; undefined removes the criterion of the column. One undo step.
   */
  setColumnFilter(col: number, choice: { values: string[]; blank?: boolean } | undefined): void {
    if (this.destroyed || this.locked) return;
    this.commitEditing();
    const sheet = this.state().sheet;
    const filter = sheet.autoFilter;
    if (!filter) throw new RangeError('The sheet has no filter');
    const criterion = choice ? { kind: 'values' as const, values: choice.values.map(String), ...(choice.blank ? { blank: true } : {}) } : undefined;
    this.applyOps(setFilterColumn(this.workbook, this.sheetIndex, toIndex(col, MAX_COLS) - filter.range.left, criterion, { locale: this.locale }));
    const summary = filterSummary(sheet);
    if (summary) this.notify(this.strings.recordsShown(summary.shown, summary.total), 'info');
  }

  /**
   * Opens the menu of the filter button of a column (sort, search, the values to show).
   * Returns false when the column has no button, or the view is read-only.
   */
  openFilterMenu(col: number, touch = false): boolean {
    if (this.destroyed || this.locked) return false;
    // First of all: what is being typed is written, which shows the sheet of its cell again.
    if (!this.tryCommit()) return false;
    const state = this.state();
    const sheet = state.sheet;
    const filter = this.gridFilter(sheet);
    if (!filter || !(col >= filter.range.left && col <= filter.range.right) || !filter.buttons[col - filter.range.left]) return false;
    this.closeMenu();
    if (!this.tryCommit()) return false;
    const colId = col - filter.range.left;
    const strings = this.strings;
    const list = filterValues(this.workbook, this.sheetIndex, colId, { locale: this.locale });
    const model = sheet.autoFilter as AutoFilter;
    const cell = rangeRect(this.layout(state), { top: filter.range.top, left: col, bottom: filter.range.top, right: col });
    const x = clamp(cell.x, 0, Math.max(this.width - 1, 0));
    const y = clamp(cell.y, 0, Math.max(this.height - 1, 0));
    const anchor = { x, y, width: clamp(cell.x + cell.width - x, 0, this.width - x), height: clamp(cell.y + cell.height - y, 0, this.height - y) };
    this.panel = new FilterMenu({
      parent: this.root,
      anchor,
      label: strings.filterMenu(columnName(col)),
      strings,
      values: list.items,
      blanks: list.blanks,
      blanksChecked: list.blanksChecked,
      note: list.truncated ? strings.tooManyValues(MAX_FILTER_VALUES) : undefined,
      canClear: isFilterActive(model.columns.find((column) => column.colId === colId)),
      unchangedKeeps: list.opaque,
      openedByTouch: touch,
      onSort: (descending) =>
        this.attempt(() => {
          const current = this.state().sheet.autoFilter;
          if (current) this.applyOps(sortRange(this.workbook, this.sheetIndex, current.range, [{ col, descending }], { locale: this.locale, header: true }));
        }),
      onApply: (choice) =>
        this.attempt(() => {
          this.setColumnFilter(col, choice);
          // Choosing values takes the place of a criterion that could not be shown as values: say so.
          if (list.opaque && choice) {
            const summary = filterSummary(this.state().sheet);
            this.notify(summary ? `${strings.filterReplaced} ${strings.recordsShown(summary.shown, summary.total)}` : strings.filterReplaced, 'info');
          }
        }),
      onClose: () => this.panelClosed(),
    });
    this.placeTip();
    return true;
  }

  private panelClosed(): void {
    this.panel = undefined;
    if (this.destroyed) return;
    // Back to the cells; the message beside the active cell, hidden while the panel was open, comes back.
    this.focusInput(true);
    this.schedule();
  }

  /** Alt+Down on a cell with a button: opens what the button stands for. */
  private openCellList(): boolean {
    const state = this.state();
    const { active } = state;
    const filter = this.gridFilter(state.sheet);
    if (filter && active.row === filter.range.top && active.col >= filter.range.left && active.col <= filter.range.right) return this.openFilterMenu(active.col);
    return this.openValidationList();
  }

  // ------------------------------------------------------------- selection

  private expandRange(range: CellRange): CellRange {
    return expandByMerges(this.merges(), range);
  }

  private setSelection(active: CellPos, range: CellRange, focus: CellPos, revealCell?: CellPos): void {
    const state = this.state();
    if (this.pointing) {
      // The command moves what a formula is pointed at: pointBy() reads it from here and puts the selection back.
      state.active = { row: active.row, col: active.col };
      state.range = { ...range };
      state.focus = { row: focus.row, col: focus.col };
      if (revealCell) this.reveal(revealCell, range);
      return;
    }
    const changed = !sameCell(active, state.active) || !sameRange(range, state.range);
    // A list that belongs to another cell has nothing to choose for this one.
    if (changed && this.panel) this.panel.close();
    if (!sameCell(active, state.active)) this.commentDismissed = undefined;
    state.active = { row: active.row, col: active.col };
    state.range = { ...range };
    state.focus = { row: focus.row, col: focus.col };
    if (revealCell) this.reveal(revealCell, range);
    if (!changed) return;
    this.announce();
    this.schedule();
    if (this.options.onSelectionChange) this.options.onSelectionChange(this.selection, this.sheetIndex);
  }

  /** Selects one cell; a merged cell is selected as a whole, with its top-left cell active. */
  private collapseTo(cell: CellPos, revealIt: boolean): void {
    const merge = mergeIn(this.merges(), cell.row, cell.col);
    const active = merge ? { row: merge.top, col: merge.left } : cell;
    const range = merge ? merge : { top: cell.row, left: cell.col, bottom: cell.row, right: cell.col };
    this.setSelection(active, range, active, revealIt ? active : undefined);
  }

  /** Stretches the range from the active cell to `cell`. */
  private extendTo(cell: CellPos, revealIt: boolean): void {
    const state = this.state();
    this.setSelection(state.active, this.expandRange(rangeBetween(state.active, cell)), cell, revealIt ? cell : undefined);
  }

  private selectAll(): void {
    const state = this.state();
    this.setSelection(state.active, { top: 0, left: 0, bottom: MAX_ROWS - 1, right: MAX_COLS - 1 }, { row: MAX_ROWS - 1, col: MAX_COLS - 1 });
  }

  /** Puts the address and text of the active cell where assistive technology reads it. */
  private announce(): void {
    const state = this.state();
    const { active, range } = state;
    let text = cellName(active.row, active.col);
    const value = cellDisplayText(this.workbook, getCell(state.sheet, active.row, active.col), this.locale);
    if (value !== '') text += ` ${value}`;
    const merge = mergeIn(this.merges(), active.row, active.col);
    const single = range.top === range.bottom && range.left === range.right;
    if (!single && !(merge && sameRange(merge, range))) text += `, ${this.strings.selected(rangeName(range))}`;
    // What a sighted user sees beside the cell: the message of its validation, and what its mark says.
    const prompt = this.activeInfo().prompt;
    if (prompt) text += `. ${prompt.title}${prompt.title !== '' && prompt.text !== '' ? ': ' : ''}${prompt.text}`;
    const note = this.formulaNote(active.row, active.col);
    if (note !== '') text += `. ${note}`;
    const comment = this.commentOf(active.row, active.col);
    if (comment) text += `. ${this.strings.comment}: ${this.commentText(comment.comment)}`;
    this.live.textContent = text;
    this.syncInput();
  }

  // -------------------------------------------------------------- keyboard

  private onKeyDown = (event: KeyboardEvent): void => {
    if (this.destroyed || event.defaultPrevented) return;
    if (event.key === 'Escape' && this.drag && this.drag.resize) {
      // Gives up a resize that is being dragged: nothing changes.
      this.endDrag();
      event.preventDefault();
      event.stopPropagation();
      return;
    }
    // Ctrl+P, while a cell is edited too: the browser would print the host page with a part of the grid in it.
    if ((event.ctrlKey || event.metaKey) && !event.altKey && !event.shiftKey && (event.key === 'p' || event.key === 'P')) {
      event.preventDefault();
      this.print().catch(() => undefined);
      return;
    }
    if (this.edit) {
      this.onEditKey(event);
      return;
    }
    // 229: the key went to an input method, also where isComposing is not reported.
    if (event.isComposing || event.keyCode === 229) return;
    if (event.key === 'Escape' && this.commentShown && !this.hoverComment) {
      // Puts the comment of the active cell away; it comes back when the cell is entered again.
      const shown = this.commentShown;
      this.commentDismissed = { sheet: shown.sheet, row: shown.row, col: shown.col };
      this.schedule();
      event.preventDefault();
      event.stopPropagation();
      return;
    }
    // Whatever is typed next replaces the idle text, and copy needs a selection in some browsers.
    this.selectIdleText();
    const ctrl = event.ctrlKey || event.metaKey;
    const shift = event.shiftKey;
    if (event.altKey) {
      // Alt+Down opens the list of the cell: the menu of a filter button, the values of a validation.
      if (event.key === 'ArrowDown' && !ctrl && !shift) {
        if (this.openCellList()) event.preventDefault();
        return;
      }
      // Browsers keep Ctrl+PageUp/PageDown for their own tabs; Alt does the same here.
      if (event.key !== 'Delete' && event.key !== 'Backspace') {
        if (ctrl || shift || (event.key !== 'PageUp' && event.key !== 'PageDown')) return;
        this.stepSheet(event.key === 'PageDown' ? 1 : -1);
        event.preventDefault();
        return;
      }
    }
    const state = this.state();
    const { rows, cols } = state.metrics;
    let handled = true;
    switch (event.key) {
      case 'ArrowUp':
        this.moveBy(-1, 0, ctrl, shift);
        break;
      case 'ArrowDown':
        this.moveBy(1, 0, ctrl, shift);
        break;
      case 'ArrowLeft':
        this.moveBy(0, -1, ctrl, shift);
        break;
      case 'ArrowRight':
        this.moveBy(0, 1, ctrl, shift);
        break;
      case 'Enter':
        if (ctrl) handled = false;
        else this.moveBy(shift ? -1 : 1, 0, false, false);
        break;
      case 'Home': {
        const layout = this.layout(state);
        const from = shift ? state.focus : state.active;
        const target = ctrl
          ? { row: rows.next(layout.frozenRows, 1), col: cols.next(layout.frozenCols, 1) }
          : { row: from.row, col: cols.next(0, 1) };
        this.moveTo(target, shift);
        break;
      }
      case 'End':
        if (!ctrl) handled = false;
        else {
          const lastRow = clamp(count(state.sheet.rowCount) - 1, 0, MAX_ROWS - 1);
          const lastCol = clamp(count(state.sheet.colCount) - 1, 0, MAX_COLS - 1);
          this.moveTo({ row: rows.next(lastRow, -1), col: cols.next(lastCol, -1) }, shift);
        }
        break;
      case 'PageUp':
        if (ctrl) this.stepSheet(-1);
        else this.page(-1, shift);
        break;
      case 'PageDown':
        if (ctrl) this.stepSheet(1);
        else this.page(1, shift);
        break;
      case 'a':
      case 'A':
        if (ctrl && !shift) this.selectAll();
        else handled = false;
        break;
      case 'Delete':
      case 'Backspace':
        // With Ctrl, Alt or Cmd the browser would delete a word of the idle text; here every variant clears
        // the selection. Shift+Delete is "cut" on Windows: left alone, it arrives as a cut event or does nothing.
        if (shift) handled = false;
        else if (!this.locked) this.attempt(() => this.clearSelection());
        break;
      case 'F2':
        // Shift+F2 is the note of the cell, as in Excel.
        if (shift && !ctrl) {
          if (!this.locked) this.attempt(() => void this.editComment());
        } else this.startEditing();
        break;
      case 'b':
      case 'B':
      case 'i':
      case 'I':
      case 'u':
      case 'U':
        if (!ctrl || shift || this.locked) handled = false;
        else {
          // Switches on what the active cell shows.
          const font = this.selectionFormat.font;
          const key = event.key.toLowerCase();
          const patch: StylePatch = key === 'b' ? { bold: !font.bold } : key === 'i' ? { italic: !font.italic } : { underline: font.underline === 'none' ? 'single' : 'none' };
          this.attempt(() => this.format(patch));
        }
        break;
      case 'z':
      case 'Z':
        if (!ctrl) handled = false;
        else if (shift) this.redo();
        else this.undo();
        break;
      case 'y':
      case 'Y':
        if (ctrl && !shift) this.redo();
        else handled = false;
        break;
      case '+':
      case '=':
        handled = ctrl && this.structuralKey('insert');
        break;
      case '-':
      case '_':
        handled = ctrl && this.structuralKey('delete');
        break;
      default:
        handled = false;
    }
    if (handled) event.preventDefault();
  };

  private moveTo(cell: CellPos, extend: boolean): void {
    if (extend) this.extendTo(cell, true);
    else this.collapseTo(cell, true);
  }

  /** One step (or, with `jump`, to the edge of the data block) in a direction; exactly one of dRow and dCol is not 0. */
  private moveBy(dRow: number, dCol: number, jump: boolean, extend: boolean): void {
    const state = this.state();
    const { rows, cols } = state.metrics;
    const vertical = dRow !== 0;
    const step = vertical ? dRow : dCol;
    const axis = vertical ? rows : cols;
    const next = (from: number, line: number) => (jump ? this.jumpIndex(vertical, line, from, step) : stepVisible(axis, from, step));

    if (!extend) {
      // Leave a merged cell from its edge, not from its top-left cell.
      const active = state.active;
      const merge = mergeIn(this.merges(), active.row, active.col);
      const from = vertical ? (merge ? (step > 0 ? merge.bottom : merge.top) : active.row) : merge ? (step > 0 ? merge.right : merge.left) : active.col;
      const to = next(from, vertical ? active.col : active.row);
      if (to < 0) this.collapseTo(active, true);
      else this.collapseTo(vertical ? { row: to, col: active.col } : { row: active.row, col: to }, true);
      return;
    }

    // Extending moves the edge of the range that lies opposite the active
    // cell. A merged range can swallow a step, so step until the range changes.
    const anchor = state.active;
    let focus = state.focus;
    let range = state.range;
    for (let guard = 0; guard < MAX_EXTEND_STEPS; guard++) {
      const at = vertical ? focus.row : focus.col;
      const anchorAt = vertical ? anchor.row : anchor.col;
      const far = at > anchorAt || (at === anchorAt && step > 0);
      const edge = vertical ? (far ? range.bottom : range.top) : far ? range.right : range.left;
      const to = next(edge, vertical ? focus.col : focus.row);
      if (to < 0 || to === at) break;
      focus = vertical ? { row: to, col: focus.col } : { row: focus.row, col: to };
      const grown = this.expandRange(rangeBetween(anchor, focus));
      // Whole rows stay whole rows and whole columns whole columns.
      if (vertical && fullWidth(range)) {
        grown.left = 0;
        grown.right = MAX_COLS - 1;
      } else if (!vertical && fullHeight(range)) {
        grown.top = 0;
        grown.bottom = MAX_ROWS - 1;
      }
      if (!sameRange(grown, range)) {
        range = grown;
        break;
      }
    }
    this.setSelection(anchor, range, focus, focus);
  }

  /**
   * Where Ctrl+arrow lands along one row or column: the end of the block of
   * filled cells the start is in, otherwise the next filled cell, otherwise
   * the edge of the sheet. Hidden rows and columns are passed over.
   */
  private jumpIndex(vertical: boolean, line: number, from: number, step: number): number {
    const state = this.state();
    const sheet = state.sheet;
    const axis = vertical ? state.metrics.rows : state.metrics.cols;
    const used = Math.min(count(vertical ? sheet.rowCount : sheet.colCount), axis.count);
    const filled = (index: number) => {
      if (index < 0 || index >= used) return false;
      const cell = vertical ? getCell(sheet, index, line) : getCell(sheet, line, index);
      // A formula typed in but not calculated yet has no value and still fills its cell.
      return cell !== undefined && ((cell.value !== null && cell.value !== '') || cell.formula !== undefined);
    };

    let next = stepVisible(axis, from, step);
    if (next < 0) return from;
    if (filled(from) && filled(next)) {
      // Every pass moves one cell further inside the used area, which bounds the loop.
      for (;;) {
        const after = stepVisible(axis, next, step);
        if (after < 0 || !filled(after)) return next;
        next = after;
      }
    }
    if (step < 0 && next >= used) {
      next = used - 1;
      if (next >= 0 && axis.size(next) === 0) next = stepVisible(axis, next, -1);
    }
    while (next >= 0 && next < used) {
      if (filled(next)) return next;
      next = stepVisible(axis, next, step);
    }
    return step > 0 ? axis.next(axis.count - 1, -1) : axis.next(0, 1);
  }

  /** Moves by the height of the visible area and scrolls by the same amount. */
  private page(direction: number, extend: boolean): void {
    const state = this.state();
    this.readScroll();
    const rows = state.metrics.rows;
    const area = scrollAreaSize(this.layout(state));
    const from = extend ? state.focus : state.active;
    // Up measures from the bottom of the row so that it undoes a step down exactly.
    const start = rows.offset(from.row) + (direction < 0 ? Math.max(rows.size(from.row) - 1, 0) : 0);
    let row = rows.indexAt(start + direction * area.height);
    if (direction > 0 ? row <= from.row : row >= from.row) {
      const next = stepVisible(rows, from.row, direction);
      row = next < 0 ? from.row : next;
    }
    this.scrollToOffsets(state.scrollX, state.scrollY + direction * area.height);
    this.moveTo({ row, col: from.col }, extend);
  }

  // --------------------------------------------------------------- editing

  get isEditing(): boolean {
    return this.edit !== undefined;
  }

  /**
   * Opens the editor on the active cell. Without an argument it holds the
   * cell's content as the user would type it; with one, that text instead.
   */
  startEditing(initialText?: string): void {
    if (this.destroyed || this.locked || this.edit) return;
    const state = this.state();
    if (state.sheet.kind !== 'worksheet') return;
    this.closeMenu();
    this.beginEdit();
    const edit = this.edit as Edit | undefined;
    if (!edit) return;
    const text = initialText !== undefined ? String(initialText) : edit.original;
    this.input.value = text;
    this.focusInput(true);
    this.showKeyboard();
    try {
      this.input.setSelectionRange(text.length, text.length);
    } catch {
      // Only fails on an element that is not rendered; the caret position is cosmetic.
    }
    this.placeInput();
    this.emitEditing();
  }

  /** While editing: the text in the editor. Otherwise: the edit text of the active cell. */
  get editText(): string {
    if (this.edit) return this.input.value;
    const state = this.state();
    if (state.sheet.kind !== 'worksheet') return '';
    return editTextAt(this.workbook, state.sheet, state.active.row, state.active.col, { locale: this.locale });
  }

  /** Replaces the text being edited; starts editing the active cell when nothing is being edited. Does not move focus. */
  setEditText(text: string): void {
    if (this.destroyed || this.locked) return;
    if (!this.edit) {
      if (this.state().sheet.kind !== 'worksheet') return;
      this.closeMenu();
      this.beginEdit();
    }
    this.input.value = String(text);
    // The echo of a reference the view itself wrote into a field of the shell: writePoint() goes on from here.
    if (this.pointWriting) return;
    // Typed in a field of the shell: what was pointed at before is part of the text now.
    this.typed();
    this.placeInput();
    this.emitEditing();
  }

  /**
   * Gives keyboard focus to the grid. The focus ring is for people who arrive
   * with the keyboard; code that hands the focus back after a click on a
   * toolbar leaves it off, which is the default.
   */
  focus(options?: { ring?: boolean }): void {
    if (this.destroyed) return;
    const ring = options !== undefined && options.ring === true;
    this.focusInput(!ring);
    if (ring) this.root.classList.add('bym-focus-ring');
  }

  /**
   * What an error of the operations layer says, in the language of the view.
   * Errors it does not know give the general "could not be completed".
   */
  localizeError(error: unknown): string {
    const strings = this.strings;
    // A data validation speaks for itself: the words are those of the file.
    if (error instanceof ValidationError) return error.display;
    const message = error instanceof Error ? error.message : typeof error === 'string' ? error : '';
    let match: RegExpExecArray | null;
    if (message === 'Too many cells to format at once') return strings.tooManyCells;
    if ((match = /^Too many rows to (format|resize|change) at once$/.exec(message))) return strings.tooManyRows(match[1] as 'format' | 'resize' | 'change');
    if (message === 'Whole rows or columns cannot be merged') return strings.cannotMergeWhole;
    if ((match = /^Cannot insert (\d+) (row|col)\(s\): cells would be pushed off the sheet$/.exec(message))) return strings.insertPushesOff(Number(match[1]), match[2]);
    if ((match = /^Cannot (insert|delete) (row|col)s through the array formula in (.+)$/.exec(message))) {
      return strings.arrayFormulaBlocks(match[1] as 'insert' | 'delete', match[2], match[3]);
    }
    if ((match = /^Cannot (insert|delete) (row|col)s in a sheet with (.+): not supported yet$/.exec(message))) {
      return strings.notSupportedWith(match[1] as 'insert' | 'delete', match[2], match[3]);
    }
    if ((match = /^Cannot paste more than (\d+) cells at once$/.exec(message))) return strings.pasteTooLarge(match[1]);
    if (message === 'The workbook has too many cell formats') return strings.tooManyFormats;
    if (message === 'The formula cannot be read') return strings.badFormula;
    if (message === 'The range is too large to copy') return strings.copyTooLarge;
    if (message === 'There is no data to filter') return strings.nothingToFilter;
    if (message === 'Cannot sort a range that contains merged cells') return strings.cannotSortMerged;
    if (message === 'Cannot sort a range that contains threaded comments') return strings.cannotSortThreaded;
    if ((match = /^Cannot sort through the array formula in (.+)$/.exec(message))) return strings.cannotSortArray(match[1]);
    if ((match = /^Cannot (sort|filter) in a sheet with (.+): not supported yet$/.exec(message))) return strings.sortNotSupportedWith(match[1] as 'sort' | 'filter', match[2]);
    if ((match = /^Cannot change the comment in (.+): it is a threaded comment$/.exec(message))) return strings.noteThreaded(match[1]);
    if ((match = /^The text of a note cannot be longer than (\d+) characters$/.exec(message))) return strings.noteTooLong(Number(match[1]));
    if (message === 'The sheet has too many comments') return strings.tooManyNotes;
    if ((match = /^Cannot change the comments of this sheet: (.+)$/.exec(message))) {
      return strings.notesBlocked(/legacy drawing/.test(match[1]) ? 'drawing' : /shared/.test(match[1]) ? 'shared' : 'parts');
    }
    const rules = strings.rules;
    if (/^Cannot (delete|clear) this rule: the extension list/.test(message)) return rules.extensionRuleKept;
    if (/^Cannot change this rule: the extension list/.test(message)) return rules.extensionRuleSays;
    if (/^Cannot change the order of the rules of this sheet/.test(message)) return rules.orderFixed;
    if (/^A rule of the kind ".*" cannot be changed here/.test(message)) return rules.ruleKindKept;
    if (message === 'The sheet has too many conditional formats') return rules.tooManyConditional;
    if (message === 'The rule needs a text to look for') return rules.needValue(rules.text);
    if (message === 'The rule needs a period' || message === 'This period is not known') return rules.needValue(rules.period);
    if (message === 'This icon set is not known, or is not made here') return rules.needValue(rules.iconStyle);
    if ((match = /^The number of a top or bottom rule is a whole number between 1 and (\d+)$/.exec(message))) return rules.needWholeUpTo(rules.rank, Number(match[1]));
    if (/: the rule there is one of the extension list/.test(message)) return rules.extensionRule;
    if (message === 'The sheet has too many data validations') return rules.tooManyValidations;
    if (message === 'The rule would be left with too many separate ranges') return rules.tooManyRanges;
    if ((match = /^The (.+) of the rule cannot be read as a formula$/.exec(message))) return rules.badRuleFormula(match[1] === 'highest value' ? rules.maximum : match[1] === 'source' ? rules.source : match[1] === 'formula' ? rules.formula : rules.value);
    if (/^The rule needs|^A list needs a value|^A custom rule needs a value/.test(message)) return rules.needValue(rules.value);
    if ((match = /^The items of a list written into the rule cannot be longer than (\d+) characters together$/.exec(message))) return rules.listTooLong(Number(match[1]));
    if ((match = /^The (.+) of a rule cannot be longer than (\d+) characters$/.exec(message))) return rules.textTooLong(/title/.test(match[1]) ? rules.title : rules.message, Number(match[2]));
    return sheetErrorText(message, strings) || strings.actionFailed;
  }

  /** Says something to the user: in the live region, and to the host page if it listens. */
  private notify(message: string, kind: 'error' | 'info'): void {
    this.live.textContent = message;
    if (this.options.onNotice) this.options.onNotice({ message, kind });
  }

  /**
   * Says something once the key or press that is being handled is done with: the selection
   * usually moves on right after a cell is written, and what is said about the new cell
   * would take the place of this.
   */
  private notifyLater(message: string, kind: 'error' | 'info' = 'info'): void {
    if (this.noticeTimer !== undefined) clearTimeout(this.noticeTimer);
    this.noticeTimer = setTimeout(() => {
      this.noticeTimer = undefined;
      if (!this.destroyed) this.notify(message, kind);
    }, 0);
  }

  private emitEditing(): void {
    this.syncSheetLock();
    if (this.options.onEditingChange) this.options.onEditingChange({ editing: this.edit !== undefined, text: this.editText });
  }

  /**
   * Applies a style patch to the selection as one undo step. No-op when read-only.
   * Rows that the filter hides are left out, as in Excel; with whole columns selected the
   * format is that of the columns, which those rows share.
   */
  format(patch: StylePatch): void {
    if (this.destroyed || this.locked) return;
    this.commitEditing();
    const state = this.state();
    const range = state.range;
    const filtered = fullHeight(range) ? [] : rowsHiddenByFilter(state.sheet, range.top, range.bottom);
    if (filtered.length === 0) {
      this.applyOps(formatRange(this.workbook, this.sheetIndex, range, patch));
      return;
    }
    const ops: Op[] = [];
    for (const run of rowRunsWithout(range.top, range.bottom, filtered)) {
      const part = formatRange(this.workbook, this.sheetIndex, { top: run.top, left: range.left, bottom: run.bottom, right: range.right }, patch);
      for (let i = 0; i < part.length; i++) ops.push(part[i]);
    }
    this.applyOps(ops);
  }

  /** Merges the selection into one cell, or splits the merges it touches. */
  merge(): void {
    if (this.destroyed || this.locked) return;
    this.commitEditing();
    const range = { ...this.state().range };
    this.applyOps(mergeRange(this.workbook, this.sheetIndex, range), { active: { row: range.top, col: range.left }, range });
  }

  unmerge(): void {
    if (this.destroyed || this.locked) return;
    this.commitEditing();
    this.applyOps(unmergeRange(this.workbook, this.sheetIndex, this.state().range));
  }

  /** The format of the active cell (font, fill, border, alignment, numFmt), for a toolbar to reflect. */
  get selectionFormat(): ResolvedCellStyle {
    const state = this.state();
    return formatAt(this.workbook, state.sheet, state.active.row, state.active.col);
  }

  /** Whether the active cell is part of a merged range. */
  get selectionMerged(): boolean {
    const active = this.state().active;
    return mergeIn(this.merges(), active.row, active.col) !== undefined;
  }

  /** Width in characters / height in points; undefined restores the default. One undo step. */
  setColumnWidth(first: number, last: number, width: number | undefined): void {
    if (this.destroyed || this.locked) return;
    const a = toIndex(first, MAX_COLS);
    const b = toIndex(last, MAX_COLS);
    if (width !== undefined && (typeof width !== 'number' || Number.isNaN(width))) throw new RangeError('Invalid column width');
    this.commitEditing();
    this.applyOps(columnWidthOps(this.workbook, this.sheetIndex, Math.min(a, b), Math.max(a, b), width));
  }

  setRowHeight(first: number, last: number, height: number | undefined): void {
    if (this.destroyed || this.locked) return;
    const a = toIndex(first, MAX_ROWS);
    const b = toIndex(last, MAX_ROWS);
    if (height !== undefined && (typeof height !== 'number' || Number.isNaN(height))) throw new RangeError('Invalid row height');
    this.commitEditing();
    this.applyOps(rowHeightOps(this.sheetIndex, Math.min(a, b), Math.max(a, b), height));
  }

  /** Writes the editor's text to the cell and closes the editor. Returns whether a cell was being edited. */
  commitEditing(): boolean {
    const edit = this.edit;
    if (!edit) return false;
    if (this.committing) return false;
    const text = this.input.value;
    // Text that was only looked at, or the idle text an abandoned composition left behind, is not a change.
    const unchanged = text === edit.original || (edit.idle !== undefined && text === edit.idle);
    // The cell is written on its own sheet, shown as it is for every other change: undo comes back to it.
    this.returnHome(edit);
    const home = this.workbook.sheets.indexOf(edit.home);
    if (home < 0) {
      // The sheet of the cell left the workbook under the view: there is no cell to write. Said, not passed over.
      this.endEdit();
      if (!unchanged) this.notify(this.strings.editSheetGone, 'error');
      return true;
    }
    if (!unchanged) {
      // The editor stays open until the cell is written: if that fails, what was typed is still there.
      this.committing = true;
      let warning: string | undefined;
      try {
        // Reading the text may add a format for it (a date, a percentage) before it is known to be written.
        const styles = markStyles(this.workbook.styles);
        let ops: Op[];
        try {
          ops = setCellInput(this.workbook, home, edit.row, edit.col, text, { locale: this.locale });
          // A value the cell's data validation refuses is not written, and the editor stays open with it.
          warning = this.validateInput(home, edit.row, edit.col, text, ops);
        } catch (error) {
          // What is refused leaves nothing behind: saved, the file is byte for byte what it was.
          resetStyles(this.workbook.styles, styles);
          throw error;
        }
        this.applyOps(ops);
      } finally {
        this.committing = false;
      }
      this.endEdit();
      // UNSURE: Excel asks before it keeps a value that breaks a rule that only warns; here it is kept and said.
      if (warning !== undefined) this.notifyLater(warning);
      return true;
    }
    this.endEdit();
    return true;
  }

  /**
   * Applies everything that is being typed in the view, as leaving it would: the cell that is
   * edited, the note in its field, the name of a sheet in its tab. For whoever is about to take
   * the workbook as it stands (to save it, to print it) or to change it from outside:
   * mount().save() and the bridge call it, and so does every method of the view that changes
   * the workbook.
   *
   * Returns true when nothing is left open. A cell whose text cannot be written throws, as
   * commitEditing() does, and the editor stays open. A note or a name that cannot be taken is
   * said (in the live region and to onNotice), its field stays open with what was typed, and
   * the answer is false: the workbook is then not what the user sees, and is not to be saved.
   *
   * A dialog of rules (data validation, a conditional format, the list of rules) is no field
   * that is being typed: it has OK and Cancel, and what it holds is the user's to confirm. One
   * in which nothing was changed since it was opened is closed, as Cancel closes it. One with
   * changes stays open, nothing of it is applied, that is said (in the dialog, in the live
   * region and to onNotice), and the answer is false.
   */
  commitPending(): boolean {
    if (this.destroyed) return true;
    if (this.edit) this.commitEditing();
    if (this.renaming) this.finishRename(true);
    if (this.noteEdit) this.finishNote(true);
    const settled = this.settleDialog();
    return settled && !this.edit && !this.renaming && !this.noteEdit;
  }

  /**
   * What is done about a dialog of rules that is open when the workbook is about to be taken or changed from
   * outside the dialog: closed when nothing was changed in it, left open and said otherwise (see commitPending()).
   * Returns whether no dialog is open afterwards.
   */
  private settleDialog(): boolean {
    const open = this.rulesDialog;
    if (!open) return true;
    if (open.changed()) {
      const message = this.strings.rules.dialogOpen;
      open.say(message);
      this.notify(message, 'error');
      return false;
    }
    // It does not hand back to the list it was opened from.
    this.dismiss(() => open.close());
    return !this.rulesDialog;
  }

  /** commitEditing() for keys, pointers and focus changes: a failure is announced and reported as false. */
  private tryCommit(): boolean {
    if (!this.edit) return true;
    try {
      this.commitEditing();
      return true;
    } catch (error) {
      this.notify(this.localizeError(error), 'error');
      return false;
    }
  }

  get readOnly(): boolean {
    return this.locked;
  }

  /** Turns editing off or on. Turning it off writes an open edit first; the undo history is kept either way. */
  setReadOnly(readOnly: boolean): void {
    const next = readOnly === true;
    if (this.destroyed || next === this.locked) return;
    if (next) {
      this.endTabDrag();
      this.commitEditing();
      this.finishRename(true, true);
      this.finishNote(true, true);
      this.closeMenu();
    }
    this.locked = next;
    this.input.readOnly = next;
    this.cellInfo = undefined;
    this.placeAddButton();
    this.schedule();
  }

  /** Closes the editor and leaves the cell as it was. */
  cancelEditing(): void {
    if (this.edit) this.endEdit();
  }

  /** Sets a cell from text, read the way typed input is: one undo step. */
  setCellInput(row: number, col: number, text: string): void {
    if (this.destroyed || this.locked) return;
    const r = toIndex(row, MAX_ROWS);
    const c = toIndex(col, MAX_COLS);
    this.commitEditing();
    this.applyOps(setCellInput(this.workbook, this.sheetIndex, r, c, String(text), { locale: this.locale }));
  }

  /** Empties the selected cells and keeps their formatting, like the Delete key. */
  clearSelection(): void {
    if (this.destroyed || this.locked) return;
    this.commitEditing();
    this.applyOps(clearRange(this.workbook, this.sheetIndex, this.state().range));
  }

  /**
   * Undoes the last change and returns whether there was one. While a cell is
   * edited it only abandons the edit (and returns true): as in Excel, the
   * steps before it are out of reach until the edit is over. While a note or
   * the name of a sheet is typed in its field it applies what was typed and
   * does nothing else (true; the next call undoes that); false when it cannot
   * be applied, and the field stays open. While a dialog of rules is open it
   * does nothing at all (false). redo() does the same.
   */
  undo(): boolean {
    return this.travel(false);
  }

  redo(): boolean {
    return this.travel(true);
  }

  /** Copies the selection: returns what belongs on the clipboard and remembers the cells for a later paste. */
  copy(): CopiedRange {
    const plain = copyRange(this.workbook, this.sheetIndex, this.state().range, { locale: this.locale });
    // The same text can come back from anywhere (another program, another view, another cell that shows the
    // same); only HTML that carries this mark is this copy.
    this.copyId = Math.random().toString(36).slice(2) + Date.now().toString(36);
    const copied = { ...plain, html: plain.html.replace('<table>', `<table data-bym-copy="${this.copyId}">`) };
    this.copied = copied;
    this.copiedNames = this.workbook.sheets.map((sheet) => foldSheetName(sheet.name));
    return copied;
  }

  /** Copies the selection and empties it, as one undo step. Read-only views only copy. */
  cut(): CopiedRange {
    if (!this.destroyed && !this.locked) this.commitEditing();
    const copied = this.copy();
    if (!this.destroyed && !this.locked) {
      // The notes of the cells go with them, as one step: a paste of this copy puts them where the cells go.
      this.applyOps(cutRange(this.workbook, this.sheetIndex, this.state().range));
    }
    return copied;
  }

  /**
   * Pastes clipboard text at the selection and selects what was written.
   * `html` is the clipboard's text/html, if any: when it is what the last
   * copy() of this view produced, the copied cells themselves are pasted, with
   * their formats and with formulas moved along. Without it, or with any other
   * HTML, the text is read field by field as if typed.
   */
  paste(text: string, html?: string): void {
    if (this.destroyed || this.locked) return;
    this.commitEditing();
    const value = String(text);
    const own = html !== undefined && this.copied !== undefined && this.copyId !== '' && String(html).indexOf(`data-bym-copy="${this.copyId}"`) >= 0;
    // No text is nothing to paste - but for a copy of this view whose cells are empty and have a note or a rule
    // (a list to choose from, say), which shows no text at all.
    if (value === '' && !(own && this.copied && (this.copied.rules !== undefined || this.copied.notes !== undefined))) return;
    const state = this.state();
    if (state.sheet.kind !== 'worksheet') return;
    const result = pasteAt(this.workbook, this.sheetIndex, state.range, value, own ? this.copied : undefined, { locale: this.locale });
    const after = { active: { row: result.range.top, col: result.range.left }, range: result.range };
    if (!this.applyOps(result.ops, after)) this.restoreSelection(after);
    // The cells were pasted; notes and rules that could not go with them are said, not refused.
    const said: string[] = [];
    if (result.skippedNotes) said.push(this.strings.notesNotPasted(result.skippedNotes));
    if (result.skippedRules) said.push(this.strings.rulesNotPasted(result.skippedRules));
    if (said.length > 0) this.notifyLater(said.join('; '));
  }

  insertRows(at: number, count: number): void {
    this.structural('insert', 'row', at, count);
  }

  deleteRows(at: number, count: number): void {
    this.structural('delete', 'row', at, count);
  }

  insertColumns(at: number, count: number): void {
    this.structural('insert', 'col', at, count);
  }

  deleteColumns(at: number, count: number): void {
    this.structural('delete', 'col', at, count);
  }

  private structural(kind: 'insert' | 'delete', axis: 'row' | 'col', at: number, count: number): void {
    if (this.destroyed || this.locked) return;
    this.commitEditing();
    const state = this.state();
    // The new rows or columns end up selected; after a delete the selection stays where it is.
    let after: { active: CellPos; range: CellRange } | undefined;
    if (kind === 'insert') {
      after =
        axis === 'row'
          ? { active: { row: at, col: state.active.col }, range: { top: at, bottom: at + count - 1, left: 0, right: MAX_COLS - 1 } }
          : { active: { row: state.active.row, col: at }, range: { top: 0, bottom: MAX_ROWS - 1, left: at, right: at + count - 1 } };
    }
    // Deleting leaves out rows the filter hides, and takes the criterion of a deleted column with it.
    const ops: Op[] =
      kind === 'insert'
        ? [{ type: kind, sheet: this.sheetIndex, axis, at, count }]
        : axis === 'row'
          ? deleteRowOps(this.workbook, this.sheetIndex, at, count)
          : deleteColumnOps(this.workbook, this.sheetIndex, at, count, { locale: this.locale });
    this.applyOps(ops, after);
  }

  /** Ctrl with plus or minus on whole rows or columns. Returns false when the selection is neither. */
  private structuralKey(kind: 'insert' | 'delete'): boolean {
    if (this.locked) return false;
    const range = this.state().range;
    const rows = fullWidth(range) && !fullHeight(range);
    const cols = fullHeight(range) && !fullWidth(range);
    if (!rows && !cols) return false;
    const at = rows ? range.top : range.left;
    const count = (rows ? range.bottom : range.right) - at + 1;
    this.attempt(() => this.structural(kind, rows ? 'row' : 'col', at, count));
    return true;
  }

  private mark(sheets: number[], structural: boolean, selection?: { active: CellPos; range: CellRange }, cells?: HistoryMark['cells']): HistoryMark {
    const from = selection || this.state();
    return { sheet: this.sheetIndex, active: { ...from.active }, range: { ...from.range }, sheets, structural, cells };
  }

  /**
   * Recalculates the formulas that depend on what a step changed. Without a
   * list of cells (rows or columns moved) every formula is calculated again.
   * A failure here must never get in the way of editing: the cells then keep
   * the values they had.
   */
  private recalculate(cells: HistoryMark['cells'], rows?: number[]): number[] {
    if (this.options.calculate === false) return [];
    try {
      if (!this.engine) this.engine = new CalcEngine(this.workbook, { budget: { milliseconds: CALC_SLICE_MS } });
      this.marksDirty = true;
      if (!cells) this.engine.invalidate();
      else if (rows) {
        for (let i = 0; i < rows.length; i++) this.engine.rowsChanged(rows[i]);
        // A filter that is switched on or off defines or removes the name of its range.
        this.engine.namesChanged();
      }
      // After rows or columns moved every result may have changed for a reason the engine cannot
      // see in the cells (a deleted row makes a sum smaller), so the values of the file give way.
      const changed = cells ? this.engine.recalculate(cells) : this.engine.recalculate(undefined, { overwrite: true });
      if (this.engine.incomplete) this.continueCalculation();
      return changed.map((entry) => entry.sheet);
    } catch (error) {
      this.engine = undefined;
      return [];
    }
  }

  /**
   * A calculation that ran out of its time leaves the rest for later, so that the
   * page keeps answering; the rest is calculated in further turns, each drawn as it ends.
   */
  private continueCalculation(): void {
    if (this.calcTimer !== undefined) return;
    this.calcTimer = setTimeout(() => {
      this.calcTimer = undefined;
      const engine = this.engine;
      if (this.destroyed || !engine || !engine.incomplete) return;
      let sheets: number[] = [];
      this.marksDirty = true;
      try {
        sheets = engine.recalculate([]).map((entry) => entry.sheet);
        if (engine.incomplete) this.continueCalculation();
      } catch (error) {
        this.engine = undefined;
      }
      this.conditional.forEach((formats) => formats.invalidate());
      this.schedule();
      if (sheets.length > 0) this.printEpoch++;
      if (sheets.length > 0 && this.options.onChange) this.options.onChange({ sheets });
    }, 0);
  }

  /**
   * Goes on with a calculation that ran out of its time, here and now, for a few turns: for
   * what cannot wait for the next turn of the page. What is still left then goes on as before.
   */
  private finishCalculation(engine: CalcEngine): void {
    const sheets: number[] = [];
    try {
      for (let i = 0; engine.incomplete && i < MAX_FINISH_SLICES; i++) this.addSheets(sheets, engine.recalculate([]).map((entry) => entry.sheet));
    } catch (error) {
      this.engine = undefined;
    }
    this.marksDirty = true;
    this.conditional.forEach((formats) => formats.invalidate());
    this.schedule();
    if (sheets.length > 0) this.printEpoch++;
    if (sheets.length > 0 && this.options.onChange) this.options.onChange({ sheets });
  }

  /** Adds the sheets a recalculation wrote to the ones a step reports as changed. */
  private addSheets(sheets: number[], more: number[]): void {
    for (let i = 0; i < more.length; i++) if (sheets.indexOf(more[i]) < 0) sheets.push(more[i]);
  }

  /**
   * Applies operations to the shown sheet as one undo step and brings the view
   * up to date. Returns false when there was nothing to apply, and when a
   * dialog of rules with changes is open (see settleDialog()). An operation
   * that throws leaves the workbook and the view as they were.
   */
  private applyOps(ops: Op[], after?: { active: CellPos; range: CellRange }): boolean {
    if (ops.length === 0) return false;
    // A dialog of rules that is open shows the workbook as it was when it was opened (the range it is for, the place
    // of its rule): a change from outside it closes it when nothing was changed in it, and is refused otherwise.
    if (this.rulesDialog && !this.applyingRules && !this.settleDialog()) return false;
    // A note that is being typed is a step of its own, before this one.
    this.finishNote(true, true);
    const sheets: number[] = [];
    // Anything but a change of cells can change sizes, hidden runs or merged cells of its sheet.
    let structural = false;
    let moves = false;
    let rows: number[] | undefined;
    for (let i = 0; i < ops.length; i++) {
      const op = ops[i];
      if (sheets.indexOf(op.sheet) < 0) sheets.push(op.sheet);
      // A note or a rule changes no size and no position.
      if (op.type !== 'setCells' && op.type !== 'setComments' && op.type !== 'setDataValidations' && op.type !== 'setConditionalFormats') structural = true;
      if (op.type === 'insert' || op.type === 'delete') moves = true;
      if ((op.type === 'setRows' || op.type === 'setAutoFilter') && (!rows || rows.indexOf(op.sheet) < 0)) (rows || (rows = [])).push(op.sheet);
    }
    // Rows and columns that move take the formulas of every other sheet with them.
    if (moves) for (let i = 0; i < this.workbook.sheets.length; i++) if (sheets.indexOf(i) < 0) sheets.push(i);
    const before = this.mark(sheets, structural);
    const done = this.mark(sheets, structural, after);
    const inverse = this.history.apply(ops, before, done);
    // The cells written, so that what depends on them can be calculated again after the step and after
    // undoing it. They are read from the inverse: it also names the cells an operation changed on its
    // own, such as the dependants of a shared formula whose master was overwritten.
    let cells: HistoryMark['cells'] = moves ? undefined : [];
    for (let i = 0; cells && i < inverse.length; i++) {
      const op = inverse[i];
      if (op.type !== 'setCells') {
        if (op.type === 'restore') cells = undefined;
        continue;
      }
      if (cells.length + op.cells.length > MAX_RECALC_CELLS) cells = undefined;
      else for (let k = 0; k < op.cells.length; k++) cells.push({ sheet: op.sheet, row: op.cells[k].row, col: op.cells[k].col });
    }
    before.cells = cells;
    done.cells = cells;
    before.rows = rows;
    done.rows = rows;
    this.addSheets(sheets, this.recalculate(cells, rows));
    this.afterChange(sheets, structural, after);
    return true;
  }

  private travel(redo: boolean): boolean {
    if (this.destroyed || this.locked) return false;
    // A dialog of rules is modal: while it is open no step is taken back or made again (a button, a call from outside).
    if (this.rulesDialog) return false;
    // A tab that is being dragged is let go where it was: the step may change the list it was dragged in.
    this.endTabDrag();
    if (this.edit) {
      this.cancelEditing();
      return true;
    }
    // A name being typed for a sheet, or a note, is not dropped (the keys of its field undo within
    // it; this is a button or a call from outside): it is applied, as leaving the field would, and
    // that is all this call does - the next one undoes it. What cannot be applied stays in its field.
    if (this.renaming || this.noteEdit) {
      this.finishRename(true);
      this.finishNote(true);
      return !this.renaming && !this.noteEdit;
    }
    this.closeMenu();
    const listBefore = this.workbook.sheets;
    const namesBefore = this.workbook.definedNames;
    const step = redo ? this.history.redo() : this.history.undo();
    if (!step) return false;
    const meta = step.meta;
    if (meta.sheetOp) {
      // A sheet with readers goes when its deletion is done again and comes back when that is undone.
      const kind = meta.sheetOp !== 'delete' ? 'names' : redo ? 'gone' : 'back';
      this.recalculateSheets(kind, meta.sheetCells, kind === 'gone' ? listBefore : this.workbook.sheets, kind !== 'names' && namesBefore !== this.workbook.definedNames);
      this.syncSheets(meta.sheet, meta.sheetOp);
      // Each sheet has kept its selection and its scroll position; nothing is brought into view.
      this.afterChange(allSheets(this.workbook.sheets.length), false);
      return true;
    }
    if (meta.sheet !== this.sheetIndex && meta.sheet >= 0 && meta.sheet < this.workbook.sheets.length) this.setActiveSheet(meta.sheet);
    const sheets = meta.sheets.slice();
    this.addSheets(sheets, this.recalculate(meta.cells, meta.rows));
    this.afterChange(sheets, meta.structural, meta);
    return true;
  }

  /** Everything that has to follow a change of the model. */
  private afterChange(sheets: number[], structural: boolean, selection?: { active: CellPos; range: CellRange }): void {
    this.printEpoch++;
    // A rule may read any cell of any sheet.
    this.conditional.forEach((formats) => formats.invalidate());
    this.cellInfo = undefined;
    if (structural) {
      this.lastTap = undefined;
      // The cell under the pointer is another one now.
      this.hoverComment = undefined;
      this.commentDismissed = undefined;
      // Row heights, column widths and hidden runs moved; the merged ranges are indexed again on their own.
      for (let i = 0; i < sheets.length; i++) {
        const sheet = this.workbook.sheets[sheets[i]];
        const state = sheet ? this.states.get(sheet) : undefined;
        if (state) state.metrics = createSheetMetrics(state.sheet);
      }
    }
    if (selection) this.restoreSelection(selection);
    const state = this.state();
    this.growExtent(state.scrollX, state.scrollY);
    this.commitScroll();
    this.announce();
    this.schedule();
    if (this.options.onChange) this.options.onChange({ sheets: sheets.slice() });
  }

  private restoreSelection(selection: { active: CellPos; range: CellRange }): void {
    const range = selection.range;
    const active = { row: clamp(selection.active.row, range.top, range.bottom), col: clamp(selection.active.col, range.left, range.right) };
    const focus = { row: active.row === range.top ? range.bottom : range.top, col: active.col === range.left ? range.right : range.left };
    this.setSelection(active, range, focus, active);
  }

  /** Runs a command from a key, the clipboard or the menu; a failure is announced instead of thrown. */
  private attempt(run: () => void): void {
    try {
      run();
    } catch (error) {
      this.notify(this.localizeError(error), 'error');
    }
  }

  private beginEdit(idle?: string, keys = false): void {
    const state = this.state();
    const { row, col } = state.active;
    const style = resolveCellStyle(this.workbook.styles, styleAt(state.sheet, row, col));
    this.point = undefined;
    this.edit = {
      home: state.sheet,
      row,
      col,
      keys,
      pointed: false,
      // As the input will hold it: a text field turns every kind of line break into "\n".
      original: editTextAt(this.workbook, state.sheet, row, col, { locale: this.locale }).replace(/\r\n?/g, '\n'),
      idle,
      font: cssFont(style.font),
      wrap: style.alignment.wrapText === true,
    };
    this.input.setAttribute('inputmode', 'text');
    this.reveal(state.active);
    this.placeInput();
  }

  private endEdit(): void {
    const edit = this.edit;
    const pointed = edit !== undefined && edit.pointed;
    this.edit = undefined;
    this.point = undefined;
    this.composing = false;
    if (this.drag && this.drag.point) this.endDrag();
    // Back to the sheet of the cell, when another one was shown for the formula to point at.
    if (edit) this.returnHome(edit);
    this.noteActiveSheet();
    // Pointing may have scrolled the cell away.
    if (pointed) {
      this.reveal(this.state().active);
      this.schedule();
    }
    this.input.setAttribute('inputmode', 'none');
    if (this.keyboardShown) {
      this.keyboardShown = false;
      // The keyboard only follows inputmode when the focus arrives.
      if (this.hasFocus()) this.refocus();
    }
    this.syncInput();
    this.emitEditing();
  }

  /** On a touch device the on-screen keyboard has to be asked for. */
  private showKeyboard(): void {
    if (!this.touchInput) return;
    this.keyboardShown = true;
    this.refocus();
  }

  private refocus(): void {
    this.refocusing = true;
    this.quietFocus = true;
    this.input.blur();
    this.input.focus({ preventScroll: true });
    this.quietFocus = false;
    this.refocusing = false;
  }

  private hasFocus(): boolean {
    const root = this.input.getRootNode() as Document | ShadowRoot;
    return root.activeElement === this.input;
  }

  /** Gives the input the focus without the focus ring that keyboard users get. */
  private focusInput(quiet: boolean): void {
    if (quiet) this.root.classList.remove('bym-focus-ring');
    if (this.hasFocus()) return;
    this.quietFocus = quiet;
    this.input.focus({ preventScroll: true });
    this.quietFocus = false;
  }

  private selectIdleText(): void {
    if (this.edit) return;
    const input = this.input;
    if (input.selectionStart === 0 && input.selectionEnd === input.value.length) return;
    try {
      input.setSelectionRange(0, input.value.length);
    } catch {
      // Not rendered; nothing to select in.
    }
  }

  /**
   * While no cell is edited the input holds the active cell's text, selected:
   * typing replaces it, an input method sees it, and browsers that only fire
   * copy and cut for a selection have one. An empty cell gets a space.
   */
  private syncInput(): void {
    if (this.edit) return;
    const state = this.state();
    const text = state.sheet.kind === 'worksheet' ? editTextAt(this.workbook, state.sheet, state.active.row, state.active.col, { locale: this.locale }) : '';
    const value = text === '' ? ' ' : text;
    if (this.input.value !== value) this.input.value = value;
    if (this.hasFocus()) this.selectIdleText();
    this.placeInput();
  }

  /** Puts the input where its cell is: a hidden pixel while idle, the editor box while editing. */
  private placeInput(): void {
    const style = this.input.style;
    const state = this.state();
    const layout = this.layout(state);
    const edit = this.edit;
    // The editor of a cell of another sheet waits out of sight with its text while this sheet is pointed at.
    const here = edit !== undefined && edit.home === state.sheet;
    const cell = edit && here ? edit : state.active;
    const merge = mergeIn(this.merges(), cell.row, cell.col);
    const rect = rangeRect(layout, merge || { top: cell.row, left: cell.col, bottom: cell.row, right: cell.col });
    const paneLeft = layout.metrics.headerWidth + (cell.col < layout.frozenCols ? 0 : layout.frozenWidth);
    const paneTop = layout.metrics.headerHeight + (cell.row < layout.frozenRows ? 0 : layout.frozenHeight);
    const inView = rect.width > 0 && rect.height > 0 && rect.x >= paneLeft && rect.y >= paneTop && rect.x < this.width && rect.y < this.height;
    if (!edit || !here || !inView) {
      // Out of sight but still inside the view, so the browser never scrolls anything to show it. An editor
      // whose cell is scrolled away waits here with its text.
      style.left = `${clamp(rect.x, 0, Math.max(this.width - 1, 0))}px`;
      style.top = `${clamp(rect.y, 0, Math.max(this.height - 1, 0))}px`;
      style.width = '1px';
      style.height = '1px';
      style.opacity = '0';
      style.pointerEvents = 'none';
      style.border = '0';
      style.padding = '0';
      return;
    }
    // One pixel up and left: the box covers the grid lines around the cell, like the selection border does.
    const left = Math.max(rect.x - 1, 0);
    const top = Math.max(rect.y - 1, 0);
    const maxWidth = Math.max(this.width - left, 1);
    const maxHeight = Math.max(this.height - top, 1);
    const minWidth = Math.min(rect.width + 1, maxWidth);
    const minHeight = Math.min(rect.height + 1, maxHeight);
    style.opacity = '1';
    style.pointerEvents = 'auto';
    style.border = `2px solid ${DEFAULT_GRID_THEME.selectionBorder}`;
    style.padding = '0 2px';
    style.font = edit.font;
    style.lineHeight = '1.2';
    style.whiteSpace = edit.wrap ? 'pre-wrap' : 'pre';
    style.left = `${left}px`;
    style.top = `${top}px`;
    style.width = `${minWidth}px`;
    style.height = `${minHeight}px`;
    // Grow with the text: to the right until the edge of the view, then down.
    const border = 4;
    if (!edit.wrap) {
      const wanted = this.input.scrollWidth + border + 2;
      if (wanted > minWidth) style.width = `${Math.min(wanted, maxWidth)}px`;
    }
    const wantedHeight = this.input.scrollHeight + border;
    if (wantedHeight > minHeight) style.height = `${Math.min(wantedHeight, maxHeight)}px`;
  }

  private onInput = (event: Event): void => {
    if (this.destroyed) return;
    if (!this.edit) {
      // Typed over the idle text: that starts an edit which already holds what was typed. Only input that
      // adds something does; a deletion of the idle text (a key combination the view does not know) does not.
      const type = (event as InputEvent).inputType;
      const adds = type ? type.indexOf('insert') === 0 || (event as InputEvent).isComposing === true : this.input.value !== '';
      if (!adds || this.locked || this.state().sheet.kind !== 'worksheet') {
        this.syncInput();
        return;
      }
      this.closeMenu();
      // Begun by typing: the arrow keys point at cells where the formula takes a reference.
      this.beginEdit(undefined, true);
    }
    // The view itself wrote a reference: writePoint() places the editor and reports the text, once.
    if (this.pointWriting) return;
    // Typing goes on after the reference that was pointed at: the next one is added, not put in its place.
    this.typed();
    this.placeInput();
    this.emitEditing();
  };

  private onCompositionStart = (): void => {
    if (this.destroyed || this.edit || this.locked || this.state().sheet.kind !== 'worksheet') return;
    // An input method composes in place; the editor has to be there from its first key.
    this.selectIdleText();
    this.beginEdit(this.input.value, true);
    this.emitEditing();
  };

  /** An input method began or ended composing, in the editor or in a field of the shell: no pointing meanwhile. */
  private onComposition = (event: Event): void => {
    this.composing = event.type === 'compositionstart';
    if (!this.destroyed) this.syncSheetLock();
  };

  /** A press inside the editor places the caret: from here on the arrow keys move it. */
  private onInputMouseDown = (): void => {
    if (this.edit) this.edit.keys = false;
  };

  private onFocus = (): void => {
    if (this.destroyed) return;
    if (!this.quietFocus) this.root.classList.add('bym-focus-ring');
    this.selectIdleText();
  };

  private onBlur = (event: FocusEvent): void => {
    if (this.destroyed || this.refocusing) return;
    // A composition does not outlive the focus, whether its end is reported or not.
    this.composing = false;
    // The scrolling element can take the focus (a click on its scroll bar in some browsers); it hands it straight back.
    if (event.relatedTarget === this.scroller) return;
    this.root.classList.remove('bym-focus-ring');
  };

  /**
   * The focus left an element of the view or of the focus scope. If it went
   * outside both (or nowhere), an open edit is written: leaving the editor any
   * other way than Escape keeps what was typed. Inside them - a formula bar, a
   * toolbar button, a sheet tab - the edit goes on.
   */
  private onFocusOut = (event: FocusEvent): void => {
    if (this.destroyed || this.refocusing) return;
    // A composition does not outlive the focus, whether its end is reported or not.
    this.composing = false;
    if (!this.edit) return;
    const to = event.relatedTarget as Node | null;
    const scope = this.options.focusScope;
    if (to && (this.root.contains(to) || (scope !== undefined && scope.contains(to)))) return;
    this.tryCommit();
  };

  private onScrollerFocus = (): void => {
    if (!this.destroyed) this.focusInput(true);
  };

  /** A press on an empty part of the view (the tab strip beside the tabs, say) must not take the focus from the input. */
  private onRootMouseDown = (event: MouseEvent): void => {
    const target = event.target as Element | null;
    if (this.destroyed || !target || target === this.input || this.scroller.contains(target)) return;
    if (this.formulaEdit() && (this.tabs.contains(target) || target === this.addButton)) {
      // A formula is being typed: a press on a tab shows its sheet and leaves the focus in the field of the formula.
      event.preventDefault();
      return;
    }
    if (typeof target.closest === 'function' && target.closest('button, input, .bym-menu, .bym-note-edit')) return;
    event.preventDefault();
    this.focusInput(true);
  };

  private onEditKey(event: KeyboardEvent): void {
    // Enter that confirms an input method's candidate is not Enter for the cell.
    if (event.isComposing || event.keyCode === 229) return;
    const shift = event.shiftKey;
    const ctrl = event.ctrlKey || event.metaKey;
    const edit = this.edit as Edit;
    switch (event.key) {
      case 'Enter':
        event.preventDefault();
        if (event.altKey) {
          const input = this.input;
          input.setRangeText('\n', input.selectionStart, input.selectionEnd, 'end');
          this.typed();
          this.placeInput();
          this.emitEditing();
        } else {
          if (this.tryCommit() && !event.ctrlKey && !event.metaKey) this.moveBy(shift ? -1 : 1, 0, false, false);
        }
        break;
      case 'Tab':
        event.preventDefault();
        if (this.tryCommit()) this.moveBy(0, shift ? -1 : 1, false, false);
        break;
      case 'Escape':
        event.preventDefault();
        // The host page must not also act on it (close a dialog, say).
        event.stopPropagation();
        this.cancelEditing();
        break;
      case 'F2':
        // From pointing at cells with the arrow keys to moving the caret with them, and back.
        if (!ctrl && !event.altKey && !shift) {
          event.preventDefault();
          edit.keys = !edit.keys;
        }
        break;
      case 'F4':
        // Where there is no reference to change the key is the browser's.
        if (!ctrl && !event.altKey && !shift && this.cycleReference()) event.preventDefault();
        break;
      case 'PageUp':
      case 'PageDown': {
        const step = event.key === 'PageDown' ? 1 : -1;
        if (ctrl || event.altKey) {
          // The keys that change sheets: while a formula is typed the next sheet is shown for pointing at its cells.
          if (ctrl !== event.altKey && !shift && this.formulaEdit()) {
            event.preventDefault();
            this.stepSheet(step);
          }
        } else if (edit.keys && this.pointBy(() => this.page(step, shift), true)) event.preventDefault();
        break;
      }
      case 'ArrowUp':
      case 'ArrowDown':
      case 'ArrowLeft':
      case 'ArrowRight': {
        if (!edit.keys || event.altKey) break;
        const dRow = event.key === 'ArrowUp' ? -1 : event.key === 'ArrowDown' ? 1 : 0;
        const dCol = event.key === 'ArrowLeft' ? -1 : event.key === 'ArrowRight' ? 1 : 0;
        // Where the formula takes a reference the keys point at cells; anywhere else they move the caret.
        if (this.pointBy(() => this.moveBy(dRow, dCol, ctrl, shift), true)) event.preventDefault();
        break;
      }
      default:
      // Every other key is the editor's own.
    }
  }

  // ------------------------------------------------------------ point mode

  /** Whether a formula is being typed: sheets are then shown without ending the edit, and the list of sheets stays as it is. */
  private formulaEdit(): boolean {
    return this.edit !== undefined && !this.composing && isFormulaText(this.input.value);
  }

  /** Whether the sheet shown is another one than the sheet of the cell being edited. */
  private editAway(): boolean {
    return this.edit !== undefined && this.state().sheet !== this.edit.home;
  }

  /** Shows the sheet of the cell being edited again. */
  private returnHome(edit: Edit): void {
    const home = this.workbook.sheets.indexOf(edit.home);
    if (home >= 0 && home !== this.sheetIndex) this.show(home);
  }

  /** The button that adds a sheet says that it is not available while a formula is typed. */
  private syncSheetLock(): void {
    const locked = this.formulaEdit();
    if (locked === this.addButton.hasAttribute('aria-disabled')) return;
    if (locked) this.addButton.setAttribute('aria-disabled', 'true');
    else this.addButton.removeAttribute('aria-disabled');
  }

  /**
   * The field the edit is typed in: the editor of the grid, or a field of the shell around the
   * view (a formula bar) that has the focus, says that it edits the cell and holds the same text.
   */
  private editField(): HTMLTextAreaElement {
    const scope = this.options.focusScope;
    if (!scope) return this.input;
    const active = (this.input.getRootNode() as Document | ShadowRoot).activeElement as HTMLTextAreaElement | null;
    if (!active || active === this.input || active.tagName !== 'TEXTAREA' || !scope.contains(active)) return this.input;
    return active.hasAttribute(EDIT_FIELD_ATTRIBUTE) && active.value === this.input.value ? active : this.input;
  }

  /**
   * Where a reference that is pointed at now goes: in the place of the one pointed at last while
   * the caret still stands right behind it, else at the caret (in the place of the selected
   * text) when the formula takes a reference there. Undefined when pointing does nothing: no
   * formula is being typed, an input method is composing, the place is not one for a reference.
   */
  private pointTarget(): { field: HTMLTextAreaElement; start: number; end: number; live: Point | undefined } | undefined {
    if (!this.edit || this.destroyed || this.locked || this.composing || this.committing) return undefined;
    if (this.state().sheet.kind !== 'worksheet') return undefined;
    const field = this.editField();
    const text = field.value;
    const start = field.selectionStart;
    const end = field.selectionEnd;
    const point = this.point;
    if (point && point.field === field && pointSpanLive(text, start, end, point)) return { field, start: point.start, end: point.end, live: point };
    // The reference pointed at last is selected, with or without the name of its sheet (F4 leaves it so):
    // the next one takes its place all the same.
    if (point && point.field === field && start >= point.start && start < end && end === point.end && text.substring(point.start, end) === point.text) {
      return { field, start: point.start, end, live: point };
    }
    return canPointAt(text, start, end, { decimal: this.locale.decimal }) ? { field, start, end, live: undefined } : undefined;
  }

  /**
   * F4: the references of the formula being typed that the caret stands at, or that the
   * selection touches, go on to their next state ($A$1, A$1, $A1, A1; see cycleAbsolute()).
   * A caret ends up right behind the reference, a selection stays a selection. Works on the field the formula is typed in: the editor of the grid, or
   * the field of a formula bar. Returns false, and changes nothing, when no formula is being
   * typed or no reference is there.
   */
  cycleReference(): boolean {
    if (!this.edit || this.destroyed || this.locked || this.composing || this.committing) return false;
    const field = this.editField();
    const before = field.value;
    const next = cycleAbsolute(before, field.selectionStart, field.selectionEnd);
    if (!next) return false;
    const after = next.text;
    // Only what differs is written, as typing it would be: the field's own undo takes it back.
    let from = 0;
    while (from < before.length && from < after.length && before.charCodeAt(from) === after.charCodeAt(from)) from++;
    let tail = 0;
    while (tail < before.length - from && tail < after.length - from && before.charCodeAt(before.length - 1 - tail) === after.charCodeAt(after.length - 1 - tail)) tail++;
    const point = this.point;
    this.pointWriting = true;
    try {
      replaceFieldText(field, from, before.length - tail, after.substring(from, after.length - tail));
    } finally {
      this.pointWriting = false;
    }
    try {
      field.setSelectionRange(next.start, next.end);
    } catch {
      // Only fails on an element that is not rendered; the selection is a convenience.
    }
    if (field !== this.input) this.input.value = field.value;
    // The reference pointed at last stays the one the next press replaces when it is the one that was
    // changed, and is still where it was; any other trace would name a place that has moved.
    const kept =
      point !== undefined && point.field === field && before.substring(point.start, point.end) === point.text && from >= point.start && before.length - tail <= point.end;
    if (point && kept) {
      point.end += after.length - before.length;
      point.text = after.substring(point.start, point.end);
    } else this.clearPoint();
    this.placeInput();
    this.schedule();
    // What was changed, as it is written now; it reads the same in every language.
    const changedTo = after.length - tail;
    const changed = formulaReferences(after).filter((ref) => ref.bodyStart <= changedTo && ref.end >= from);
    if (changed.length > 0) this.live.textContent = after.substring(changed[0].bodyStart, changed[changed.length - 1].end);
    this.emitEditing();
    return true;
  }

  private clearPoint(): void {
    if (!this.point) return;
    this.point = undefined;
    this.schedule();
  }

  /**
   * Something was typed into the formula: the reference pointed at last is part of the text now, and a
   * drag that was pointing at a range is over (it would go on from the edited cell otherwise).
   */
  private typed(): void {
    this.clearPoint();
    if (this.drag && this.drag.point) this.endDrag();
  }

  /**
   * Whether the text being edited is a formula, an input method composing or not. Borders between
   * headers do not resize meanwhile: what is typed could not always be written first, and the press
   * would end the edit where it can.
   */
  private typingFormula(): boolean {
    return this.edit !== undefined && isFormulaText(this.input.value);
  }

  /**
   * Runs a command that moves the selection (a press, a drag, an arrow key) as pointing: the
   * command works on the cells pointed at last, or on the active cell of the shown sheet, and
   * where it ends is written into the formula as a reference. The selection itself stays as
   * it is. Returns false when the formula takes no reference now; the command is not run then.
   * With `keys` (an arrow key) a command that moved nothing writes nothing while no reference
   * was pointed at yet: at the edge of the sheet the formula would refer to its own cell.
   */
  private pointBy(run: () => void, keys = false): boolean {
    const target = this.pointTarget();
    const edit = this.edit;
    if (!target || !edit) return false;
    const state = this.state();
    const saved = { active: state.active, focus: state.focus, range: state.range };
    const live = target.live && target.live.sheet === state.sheet ? target.live : undefined;
    if (live) {
      state.active = live.active;
      state.focus = live.focus;
      state.range = live.range;
    } else {
      // From the active cell alone, whatever else is selected with it.
      const merge = mergeIn(this.merges(), saved.active.row, saved.active.col);
      state.focus = saved.active;
      state.range = merge ? { ...merge } : { top: saved.active.row, left: saved.active.col, bottom: saved.active.row, right: saved.active.col };
    }
    const from = { active: state.active, range: state.range };
    edit.pointed = true;
    let next: { active: CellPos; focus: CellPos; range: CellRange };
    this.pointing = true;
    try {
      run();
    } finally {
      this.pointing = false;
      next = { active: state.active, focus: state.focus, range: state.range };
      state.active = saved.active;
      state.focus = saved.focus;
      state.range = saved.range;
    }
    if (keys && !live && sameCell(next.active, from.active) && sameRange(next.range, from.range)) return true;
    this.writePoint(target, state.sheet, next);
    return true;
  }

  /** Writes the reference to cells of the shown sheet into the formula and remembers it as the one pointed at. */
  private writePoint(
    target: { field: HTMLTextAreaElement; start: number; end: number; live?: Point },
    sheet: Sheet,
    at: { active: CellPos; focus: CellPos; range: CellRange },
  ): void {
    const edit = this.edit;
    if (!edit) return;
    const merge = mergeIn(this.merges(), at.active.row, at.active.col);
    // The name of a sheet comes from the file: it only ever becomes text of the field.
    const text = pointReference(at.range, { single: merge !== undefined && sameRange(merge, at.range), sheet: sheet === edit.home ? undefined : sheet.name });
    const field = target.field;
    const live = target.live;
    if (live && live.sheet === sheet && live.text === text) {
      // The same reference again (the pointer moved inside the cell): nothing is written, said or reported,
      // and the undo of the field gets no step.
      live.active = { ...at.active };
      live.focus = { ...at.focus };
      live.range = { ...at.range };
      return;
    }
    // The focus was somewhere else in the view (on a tab chosen with the keyboard): typing goes on in the editor.
    if (field === this.input && !this.hasFocus()) this.focusInput(true);
    this.pointWriting = true;
    try {
      replaceFieldText(field, target.start, target.end, text);
    } finally {
      this.pointWriting = false;
    }
    if (field !== this.input) this.input.value = field.value;
    this.point = {
      sheet,
      active: { ...at.active },
      focus: { ...at.focus },
      range: { ...at.range },
      field,
      start: target.start,
      end: target.start + text.length,
      text,
    };
    this.placeInput();
    this.schedule();
    this.live.textContent = this.strings.referenceAdded(text);
    this.emitEditing();
  }

  /**
   * Writes what is being typed before a press is handled as a press on the sheet. When another
   * sheet was shown for the formula to point at, the press was on that sheet: it is shown again.
   * Returns false when the text could not be written; the editor stays open then.
   */
  private commitForPress(): boolean {
    if (!this.edit) return true;
    const shown = this.state().sheet;
    if (!this.tryCommit()) return false;
    const index = this.workbook.sheets.indexOf(shown);
    if (index >= 0 && index !== this.sheetIndex) this.setActiveSheet(index);
    return true;
  }

  private onCopy = (event: ClipboardEvent): void => {
    this.onClipboard(event, false);
  };

  private onCut = (event: ClipboardEvent): void => {
    this.onClipboard(event, true);
  };

  private onClipboard(event: ClipboardEvent, cut: boolean): void {
    // While a cell is edited the clipboard works on the editor's text, as in any text field.
    const data = event.clipboardData;
    if (this.destroyed || this.edit || !data) return;
    event.preventDefault();
    this.clipboardHandled = true;
    this.attempt(() => {
      const copied = cut ? this.cut() : this.copy();
      data.setData('text/plain', copied.text);
      data.setData('text/html', copied.html);
    });
  }

  private onPaste = (event: ClipboardEvent): void => {
    if (this.destroyed || this.edit) return;
    event.preventDefault();
    const data = event.clipboardData;
    if (!data) return;
    const text = data.getData('text/plain');
    const html = data.getData('text/html');
    this.attempt(() => this.paste(text, html));
  };

  // ------------------------------------------------------------------ menu

  private closeMenu(): void {
    if (this.menu) this.menu.close();
    if (this.panel) this.panel.close();
    if (this.tabPopover) this.tabPopover.close('owner');
    // A dialog that something else takes the place of is given up, as Escape gives it up.
    if (this.rulesDialog) this.rulesDialog.close();
  }

  /**
   * Opens the insert and delete menu for the row or column header at a canvas
   * point. Returns false when the point is not on a header.
   */
  private openHeaderMenu(x: number, y: number): boolean {
    // No menu on a sheet that is only shown for a formula to point at.
    if (this.state().sheet.kind !== 'worksheet' || this.editAway()) return false;
    const hit = hitTest(this.layout(), x, y);
    if (hit.area === 'cell') return this.openCellMenu(x, y, hit.row, hit.col);
    if (this.locked || (hit.area !== 'rowHeader' && hit.area !== 'colHeader')) return false;
    this.closeMenu();
    if (!this.tryCommit()) return false;
    const rows = hit.area === 'rowHeader';
    const index = hit.area === 'rowHeader' ? hit.row : hit.col;
    // A header outside the selected rows or columns is selected first, as a click would.
    let range = this.state().range;
    const whole = rows ? fullWidth(range) && !fullHeight(range) : fullHeight(range) && !fullWidth(range);
    const inside = whole && (rows ? index >= range.top && index <= range.bottom : index >= range.left && index <= range.right);
    if (!inside) {
      this.press(x, y, false);
      range = this.state().range;
    }
    const first = rows ? range.top : range.left;
    const count = (rows ? range.bottom : range.right) - first + 1;
    const axis = rows ? 'row' : 'col';
    const strings = this.strings;
    const items: MenuItem[] = [
      { label: rows ? strings.insertRowsAbove(count) : strings.insertColumnsLeft(count), action: () => this.attempt(() => this.structural('insert', axis, first, count)) },
      { label: rows ? strings.insertRowsBelow(count) : strings.insertColumnsRight(count), action: () => this.attempt(() => this.structural('insert', axis, first + count, count)) },
      { label: rows ? strings.deleteRows(count) : strings.deleteColumns(count), action: () => this.attempt(() => this.structural('delete', axis, first, count)) },
    ];
    this.showMenu(items, x, y);
    return true;
  }

  private showMenu(items: MenuItem[], x: number, y: number): void {
    this.menu = new ContextMenu(
      this.root,
      items,
      x,
      y,
      () => {
        this.menu = undefined;
        // Back to the cells, unless the view is going away.
        if (!this.destroyed) this.focusInput(true);
      },
      // Opened under a finger that is still down.
      this.touch !== undefined,
    );
  }

  /** The menu of a cell: clipboard, clearing, and rows and columns of the selection. */
  private openCellMenu(x: number, y: number, row: number, col: number): boolean {
    this.closeMenu();
    if (!this.tryCommit()) return false;
    let range = this.state().range;
    // A cell outside the selection is selected first, as a click would.
    if (row < range.top || row > range.bottom || col < range.left || col > range.right) {
      this.collapseTo({ row, col }, false);
      range = this.state().range;
    }
    const strings = this.strings;
    const canCopy = typeof document.execCommand === 'function' || (navigator.clipboard !== undefined && typeof navigator.clipboard.writeText === 'function');
    const items: MenuItem[] = [];
    if (!this.locked) items.push({ label: strings.cut, disabled: !canCopy, action: () => this.menuCopy(true) });
    items.push({ label: strings.copy, disabled: !canCopy, action: () => this.menuCopy(false) });
    if (!this.locked) {
      items.push({ label: strings.paste, action: () => this.menuPaste() });
      items.push({ label: strings.clearContents, action: () => this.attempt(() => this.clearSelection()) });
      // Whole columns have no rows to speak of, and whole rows no columns.
      const rows = fullHeight(range) ? 0 : range.bottom - range.top + 1;
      const cols = fullWidth(range) ? 0 : range.right - range.left + 1;
      if (rows > 0 || cols > 0) items.push({ label: '', separator: true });
      if (rows > 0) items.push({ label: strings.insertRowsAbove(rows), action: () => this.attempt(() => this.structural('insert', 'row', range.top, rows)) });
      if (cols > 0) items.push({ label: strings.insertColumnsLeft(cols), action: () => this.attempt(() => this.structural('insert', 'col', range.left, cols)) });
      if (rows > 0) items.push({ label: strings.deleteRows(rows), action: () => this.attempt(() => this.structural('delete', 'row', range.top, rows)) });
      if (cols > 0) items.push({ label: strings.deleteColumns(cols), action: () => this.attempt(() => this.structural('delete', 'col', range.left, cols)) });
      this.noteItems(items);
      if (this.state().sheet.kind === 'worksheet') {
        items.push({ label: '', separator: true });
        items.push({ label: strings.rules.validationCommand, action: () => void this.openValidationDialog() });
      }
    }
    this.showMenu(items, x, y);
    return true;
  }

  /** The note commands of the active cell for its menu: a note is added, or the one it has is edited or deleted. */
  private noteItems(items: MenuItem[]): void {
    const state = this.state();
    if (state.sheet.kind !== 'worksheet') return;
    const { row, col } = state.active;
    const strings = this.strings;
    let problem: string | undefined;
    try {
      problem = noteProblem(this.workbook, this.sheetIndex, row, col);
    } catch (error) {
      return;
    }
    // What cannot be done is shown all the same, and says why.
    const reason = problem === undefined ? undefined : this.localizeError(problem);
    const explain = reason === undefined ? undefined : () => this.notify(reason, 'error');
    const edit = () => this.attempt(() => void this.editComment(row, col));
    items.push({ label: '', separator: true });
    if (commentAtCell(state.sheet, row, col)) {
      items.push({ label: strings.editNote, disabled: reason !== undefined, title: reason, explain, action: edit });
      items.push({ label: strings.deleteNote, disabled: reason !== undefined, title: reason, explain, action: () => this.attempt(() => this.setComment(row, col, undefined)) });
    } else items.push({ label: strings.addNote, disabled: reason !== undefined, title: reason, explain, action: edit });
  }

  /**
   * Copy or cut from the menu. The clipboard can only be written from a copy
   * or cut event, so one is asked for: with the input focused it arrives at the
   * same handlers Ctrl+C and Ctrl+X reach. Where that is not possible the text
   * alone is written through the asynchronous clipboard.
   */
  private menuCopy(cut: boolean): void {
    this.focusInput(true);
    this.selectIdleText();
    this.clipboardHandled = false;
    try {
      document.execCommand(cut ? 'cut' : 'copy');
    } catch {
      // Not available; the fallback below takes over.
    }
    if (this.clipboardHandled) return;
    this.attempt(() => {
      const copied = cut ? this.cut() : this.copy();
      const clipboard = navigator.clipboard;
      if (clipboard && typeof clipboard.writeText === 'function') {
        clipboard.writeText(copied.text).then(undefined, () => {
          if (!this.destroyed) this.notify(this.strings.actionFailed, 'error');
        });
      } else this.notify(this.strings.actionFailed, 'error');
    });
  }

  /**
   * Paste from the menu: only the asynchronous clipboard can be read without a paste event.
   * Where it hands out the kinds of content it holds (read()), the HTML is taken with the text,
   * so that a copy of this view is known by its mark and comes with its formats, notes and
   * rules; elsewhere the text alone is read, and pasted as if typed.
   */
  private menuPaste(): void {
    const unavailable = () => {
      if (!this.destroyed) this.notify(this.strings.pasteWithKeyboard, 'info');
    };
    const clipboard = navigator.clipboard;
    if (!clipboard || (typeof clipboard.readText !== 'function' && typeof clipboard.read !== 'function')) {
      unavailable();
      return;
    }
    const paste = (text: string, html?: string) => {
      if (!this.destroyed) this.attempt(() => this.paste(text, html));
    };
    const asText = () => {
      if (typeof clipboard.readText !== 'function') unavailable();
      else clipboard.readText().then((text) => paste(text), unavailable);
    };
    try {
      if (typeof clipboard.read !== 'function') {
        asText();
        return;
      }
      const textOf = (item: ClipboardItem, type: string): Promise<string | undefined> =>
        item.types.indexOf(type) < 0 ? Promise.resolve(undefined) : item.getType(type).then((blob) => (typeof blob.text === 'function' ? blob.text() : new Response(blob).text()));
      clipboard
        .read()
        .then((items) => {
          const item = items[0];
          if (!item) return paste('');
          return Promise.all([textOf(item, 'text/plain'), textOf(item, 'text/html')]).then(([text, html]) => paste(text === undefined ? '' : text, html));
        })
        .then(undefined, (error: unknown) => {
          // Refused is refused: asking a second time would show the browser's question twice.
          if (error && (error as { name?: string }).name === 'NotAllowedError') unavailable();
          else asText();
        });
    } catch {
      unavailable();
    }
  }

  // ---------------------------------------------------------------- resize

  /** The column or row whose border in the headers a canvas point is on, if any. */
  private resizeAt(x: number, y: number, reach: number): Resize | undefined {
    if (this.locked) return undefined;
    const state = this.state();
    if (state.sheet.kind !== 'worksheet') return undefined;
    const layout = this.layout(state);
    const { headerWidth, headerHeight, rows, cols } = layout.metrics;
    const overColumns = y >= 0 && y < headerHeight && x >= headerWidth;
    const overRows = x >= 0 && x < headerWidth && y >= headerHeight;
    if (overColumns === overRows) return undefined;
    const hit = hitTest(layout, x, y);
    if (hit.area !== 'colHeader' && hit.area !== 'rowHeader') return undefined;
    const axis = hit.area === 'colHeader' ? 'col' : 'row';
    const metrics = axis === 'col' ? cols : rows;
    const index = hit.area === 'colHeader' ? hit.col : hit.row;
    const at = axis === 'col' ? x : y;
    const start = axis === 'col' ? colX(layout, index) : rowY(layout, index);
    const size = metrics.size(index);
    // The far edge is this column's own border; the near edge belongs to the one before, which may be hidden:
    // dragging from there is how a hidden column or row is shown again.
    // Never more than a third of the header from either side: the middle stays a place to press the header itself.
    const near = Math.min(reach, Math.floor(size / 3));
    if (at >= start + size - near) return { axis, index, origin: start };
    if (at <= start + near && index > 0) return { axis, index: index - 1, origin: start - metrics.size(index - 1) };
    return undefined;
  }

  private showGuide(resize: Resize, position: number): void {
    let guide = this.guide;
    if (!guide) {
      guide = document.createElement('div');
      guide.className = 'bym-resize-guide';
      guide.style.cssText = 'position:absolute;z-index:2;pointer-events:none';
      this.root.appendChild(guide);
      this.guide = guide;
    }
    const at = Math.max(position, resize.origin);
    const style = guide.style;
    if (resize.axis === 'col') {
      style.left = `${at}px`;
      style.top = '0';
      style.width = '1px';
      style.height = `${this.height}px`;
    } else {
      style.left = '0';
      style.top = `${at}px`;
      style.width = `${this.width}px`;
      style.height = '1px';
    }
  }

  private hideGuide(): void {
    if (this.guide && this.guide.parentNode) this.guide.parentNode.removeChild(this.guide);
    this.guide = undefined;
  }

  /** The columns or rows a resize of one of them applies to: all selected whole ones when it is among them. */
  private resizeSpan(resize: Resize): { first: number; last: number } {
    const range = this.state().range;
    if (resize.axis === 'col' && fullHeight(range) && !fullWidth(range) && resize.index >= range.left && resize.index <= range.right) {
      return { first: range.left, last: range.right };
    }
    if (resize.axis === 'row' && fullWidth(range) && !fullHeight(range) && resize.index >= range.top && resize.index <= range.bottom) {
      return { first: range.top, last: range.bottom };
    }
    return { first: resize.index, last: resize.index };
  }

  /** Applies the size a drag ended at, in pixels; 0 hides. */
  private applyResize(resize: Resize, pixels: number): void {
    const span = this.resizeSpan(resize);
    this.attempt(() => {
      if (resize.axis === 'col') this.setColumnWidth(span.first, span.last, pixelsToColumnWidth(Math.min(pixels, MAX_COLUMN_PIXELS)));
      else this.setRowHeight(span.first, span.last, (Math.min(pixels, MAX_ROW_PIXELS) * 72) / 96);
    });
  }

  /** A double click on a border: the column takes the width of its content, the row its default height. */
  private autoFit(resize: Resize): void {
    this.attempt(() => {
      if (resize.axis === 'row') {
        const span = this.resizeSpan(resize);
        this.setRowHeight(span.first, span.last, undefined);
        return;
      }
      const pixels = this.contentWidth(resize.index);
      this.setColumnWidth(resize.index, resize.index, pixels === undefined ? undefined : pixelsToColumnWidth(pixels));
    });
  }

  /** Width in pixels that shows the widest text of a column, from a sample of its rows; undefined for an empty column. */
  private contentWidth(col: number): number | undefined {
    if (this.measureCanvas === undefined) this.measureCanvas = document.createElement('canvas').getContext('2d');
    const context = this.measureCanvas;
    if (!context) return undefined;
    const state = this.state();
    const merges = this.merges();
    const fonts = new Map<number, string>();
    let widest = 0;
    let seen = 0;
    let current = '';
    const rows = state.sheet.rows;
    const iterator = rows.entries();
    for (let entry = iterator.next(); !entry.done && seen < AUTO_FIT_ROWS; entry = iterator.next()) {
      const row = entry.value[1];
      const cell = row.cells[col];
      if (!cell || row.hidden) continue;
      seen++;
      // A cell merged across columns does not ask its first column to hold all of it.
      const merge = mergeIn(merges, entry.value[0], col);
      if (merge && merge.right > merge.left) continue;
      const text = cellDisplayText(this.workbook, cell, this.locale);
      if (text === '') continue;
      let font = fonts.get(cell.style);
      if (font === undefined) {
        font = cssFont(resolveCellStyle(this.workbook.styles, cell.style).font);
        fonts.set(cell.style, font);
      }
      if (font !== current) {
        context.font = font;
        current = font;
      }
      const lines = text.split('\n');
      for (let i = 0; i < lines.length; i++) widest = Math.max(widest, context.measureText(lines[i]).width);
    }
    return widest > 0 ? Math.min(Math.ceil(widest) + CELL_TEXT_MARGIN, MAX_COLUMN_PIXELS) : undefined;
  }

  private onTouchStart = (event: TouchEvent): void => {
    if (this.destroyed || event.touches.length !== 1) return;
    const touch = event.touches[0];
    const rect = this.canvas.getBoundingClientRect();
    const x = (touch.clientX - rect.left) * (rect.width > 0 ? this.width / rect.width : 1);
    const y = (touch.clientY - rect.top) * (rect.height > 0 ? this.height / rect.height : 1);
    // Keeps the browser from scrolling with a finger that is resizing (none does while a formula is typed).
    if (!this.typingFormula() && this.resizeAt(x, y, RESIZE_REACH_TOUCH) && event.cancelable) event.preventDefault();
  };

  private onContextMenu = (event: MouseEvent): void => {
    if (this.destroyed) return;
    const point = this.canvasPoint(event);
    if (this.editAway()) {
      // The sheet is only shown for a formula to point at: neither the view's menu nor the browser's.
      event.preventDefault();
      return;
    }
    if (this.openHeaderMenu(point.x, point.y)) {
      event.preventDefault();
      // On a touch screen this is the long press; lifting the finger must not count as a tap.
      if (this.touch) this.touch.consumed = true;
    }
  };

  private onLongPress = (): void => {
    this.longPressTimer = 0;
    const touch = this.touch;
    if (this.destroyed || !touch || touch.moved || touch.consumed || this.menu) return;
    this.readScroll();
    if (touch.domX !== this.domX || touch.domY !== this.domY) return;
    if (this.openHeaderMenu(touch.x, touch.y)) touch.consumed = true;
  };

  private clearLongPress(): void {
    if (this.longPressTimer) window.clearTimeout(this.longPressTimer);
    this.longPressTimer = 0;
  }

  private endTouch(): void {
    this.clearLongPress();
    this.touch = undefined;
  }

  private onDoubleClick = (event: MouseEvent): void => {
    if (this.destroyed) return;
    const point = this.canvasPoint(event);
    const resize = this.resizeAt(point.x, point.y, RESIZE_REACH_MOUSE);
    if (resize) {
      // Not while a formula is typed, as a drag of the border.
      if (!this.typingFormula()) this.autoFit(resize);
    }
    // A second click on a button inside a cell belongs to the button.
    else if (this.filterButtonAt(point.x, point.y, 0) >= 0 || this.dropdownAt(point.x, point.y, 0)) return;
    else if (hitTest(this.layout(), point.x, point.y).area === 'cell') this.startEditing();
  };

  // --------------------------------------------------------------- pointer

  private canvasPoint(event: MouseEvent): { x: number; y: number } {
    // The rectangle is in screen terms; a CSS transform on an ancestor makes it differ from the layout size.
    const rect = this.canvas.getBoundingClientRect();
    const scaleX = rect.width > 0 ? this.width / rect.width : 1;
    const scaleY = rect.height > 0 ? this.height / rect.height : 1;
    return { x: (event.clientX - rect.left) * scaleX, y: (event.clientY - rect.top) * scaleY };
  }

  /** The cell under a point; points over the headers or outside the canvas give the nearest cell. */
  private cellAt(x: number, y: number): CellPos {
    const layout = this.layout();
    const hit = hitTest(layout, Math.max(x, layout.metrics.headerWidth), Math.max(y, layout.metrics.headerHeight));
    return hit.area === 'cell' ? { row: hit.row, col: hit.col } : { row: 0, col: 0 };
  }

  /** Applies a press at a canvas point and says what a drag from there selects. */
  private press(x: number, y: number, shift: boolean): Drag['mode'] {
    const state = this.state();
    const layout = this.layout(state);
    const { rows, cols } = state.metrics;
    const hit = hitTest(layout, x, y);
    switch (hit.area) {
      case 'corner':
        this.selectAll();
        return 'all';
      case 'rowHeader': {
        const anchor = shift ? state.active : { row: hit.row, col: layout.frozenCols > 0 ? cols.next(0, 1) : cols.indexAt(layout.scrollX) };
        this.setSelection(
          anchor,
          { top: Math.min(anchor.row, hit.row), bottom: Math.max(anchor.row, hit.row), left: 0, right: MAX_COLS - 1 },
          { row: hit.row, col: MAX_COLS - 1 },
        );
        return 'row';
      }
      case 'colHeader': {
        const anchor = shift ? state.active : { row: layout.frozenRows > 0 ? rows.next(0, 1) : rows.indexAt(layout.scrollY), col: hit.col };
        this.setSelection(
          anchor,
          { top: 0, bottom: MAX_ROWS - 1, left: Math.min(anchor.col, hit.col), right: Math.max(anchor.col, hit.col) },
          { row: MAX_ROWS - 1, col: hit.col },
        );
        return 'col';
      }
      default:
        if (shift) this.extendTo({ row: hit.row, col: hit.col }, false);
        else this.collapseTo({ row: hit.row, col: hit.col }, false);
        return 'cell';
    }
  }

  /** Stretches the selection of a drag to the cell, row or column under the pointer. */
  private dragTo(drag: Drag): void {
    if (drag.point && !this.pointing) {
      // The drag points at cells for the formula; it ends when the formula takes no reference any more.
      if (!this.pointBy(() => this.dragTo(drag))) this.endDrag();
      return;
    }
    const state = this.state();
    const cell = this.cellAt(drag.x, drag.y);
    const anchor = state.active;
    if (drag.mode === 'cell') this.extendTo(cell, false);
    else if (drag.mode === 'row') {
      this.setSelection(
        anchor,
        { top: Math.min(anchor.row, cell.row), bottom: Math.max(anchor.row, cell.row), left: 0, right: MAX_COLS - 1 },
        { row: cell.row, col: MAX_COLS - 1 },
      );
    } else if (drag.mode === 'col') {
      this.setSelection(
        anchor,
        { top: 0, bottom: MAX_ROWS - 1, left: Math.min(anchor.col, cell.col), right: Math.max(anchor.col, cell.col) },
        { row: MAX_ROWS - 1, col: cell.col },
      );
    }
  }

  private onMouseDown = (event: MouseEvent): void => {
    event.preventDefault();
  };

  private onPointerDown = (event: MouseEvent): void => {
    if (this.destroyed) return;
    const pointer = event as PointerEvent;
    const pointerId = pointer.pointerId === undefined ? 0 : pointer.pointerId;
    const point = this.canvasPoint(event);

    if (pointer.pointerType === 'touch') {
      // A touch that moves is a native scroll; only a tap selects (see onPointerUp).
      this.readScroll();
      this.endTouch();
      this.touchInput = true;
      // No resizing while a formula is typed: the finger on a border is one that taps or scrolls.
      const resize = this.typingFormula() ? undefined : this.resizeAt(point.x, point.y, RESIZE_REACH_TOUCH);
      this.touch = { pointerId, x: point.x, y: point.y, domX: this.domX, domY: this.domY, moved: false, consumed: false, resize };
      this.longPressTimer = window.setTimeout(this.onLongPress, LONG_PRESS_MS);
      return;
    }
    if (event.button !== 0) return;
    // Keeps the browser from starting a text selection or an image drag.
    event.preventDefault();
    this.touchInput = false;
    if (this.typingFormula()) {
      this.readScroll();
      // No resizing while a formula is typed: a press on a border between headers does nothing at all,
      // the edit goes on and nothing is written.
      if (this.resizeAt(point.x, point.y, RESIZE_REACH_MOUSE)) return;
    }
    if (this.pointTarget()) {
      this.readScroll();
      // A formula is being typed and takes a reference here: the press points at the cell, row or column, and
      // a drag at a range. The focus stays in the field of the formula. A press on the corner does nothing.
      this.endDrag();
      if (hitTest(this.layout(), point.x, point.y).area === 'corner') return;
      let mode: Drag['mode'] = 'cell';
      const pointed = this.pointBy(() => {
        mode = this.press(point.x, point.y, event.shiftKey);
      });
      if (pointed) {
        this.drag = { pointerId, mode, x: point.x, y: point.y, originX: point.x, originY: point.y, moved: false, point: true };
        this.capturePointer(pointerId);
      }
      return;
    }
    if (!this.commitForPress()) return;
    this.focusInput(true);
    this.endDrag();
    this.readScroll();
    // A press on the border between two headers resizes; it neither selects nor starts a selection drag.
    const resize = this.resizeAt(point.x, point.y, RESIZE_REACH_MOUSE);
    if (resize) {
      this.beginResize(pointerId, point, resize);
      return;
    }
    if (this.pressCellButton(point.x, point.y, false)) return;
    const mode = this.press(point.x, point.y, event.shiftKey);
    this.drag = { pointerId, mode, x: point.x, y: point.y, originX: point.x, originY: point.y, moved: false };
    this.capturePointer(pointerId);
  };

  private beginResize(pointerId: number, point: { x: number; y: number }, resize: Resize): void {
    this.drag = { pointerId, mode: 'resize', resize, x: point.x, y: point.y, originX: point.x, originY: point.y, moved: false };
    this.capturePointer(pointerId);
  }

  private capturePointer(pointerId: number): void {
    if (this.pointerEvents) {
      try {
        this.canvas.setPointerCapture(pointerId);
      } catch {
        // The pointer is already gone; the drag ends with the next event.
      }
    } else {
      window.addEventListener('mousemove', this.onPointerMove);
      window.addEventListener('mouseup', this.onPointerUp);
    }
  }

  private onPointerMove = (event: MouseEvent): void => {
    const pointer = event as PointerEvent;
    const pointerId = pointer.pointerId === undefined ? 0 : pointer.pointerId;
    const point = this.canvasPoint(event);
    const touch = this.touch;
    if (touch && touch.pointerId === pointerId) {
      const resize = this.typingFormula() ? undefined : touch.resize;
      if (resize && !touch.consumed && (Math.abs(point.x - touch.x) > DRAG_SLOP || Math.abs(point.y - touch.y) > DRAG_SLOP)) {
        // The finger on a border has started to move: from here on it resizes.
        this.endTouch();
        // What is being typed is written first, on the sheet of its cell; the finger is on the sheet that is shown.
        if (!this.commitForPress()) return;
        this.endDrag();
        // Looked for again: the border found when the finger came down was one of the sheet shown then.
        const border = this.resizeAt(touch.x, touch.y, RESIZE_REACH_TOUCH);
        if (!border) return;
        this.beginResize(pointerId, { x: touch.x, y: touch.y }, border);
        const drag = this.drag as Drag | undefined;
        if (drag) {
          drag.x = point.x;
          drag.y = point.y;
          drag.moved = true;
          this.showGuide(border, border.axis === 'col' ? point.x : point.y);
        }
        return;
      }
      if (Math.abs(point.x - touch.x) > TAP_SLOP || Math.abs(point.y - touch.y) > TAP_SLOP) {
        touch.moved = true;
        this.clearLongPress();
      }
      return;
    }
    const drag = this.drag;
    if (!drag) {
      // Hovering: the cursor says where a press would resize.
      if (pointer.pointerType !== 'touch') {
        const over = this.resizeAt(point.x, point.y, RESIZE_REACH_MOUSE);
        const cursor = over ? (over.axis === 'col' ? 'col-resize' : 'row-resize') : '';
        if (this.canvas.style.cursor !== cursor) this.canvas.style.cursor = cursor;
        this.hoverAt(point.x, point.y);
      }
      return;
    }
    if (drag.pointerId !== pointerId) return;
    drag.x = point.x;
    drag.y = point.y;
    if (!drag.moved) {
      if (Math.abs(point.x - drag.originX) <= DRAG_SLOP && Math.abs(point.y - drag.originY) <= DRAG_SLOP) return;
      drag.moved = true;
    }
    if (drag.resize) {
      this.showGuide(drag.resize, drag.resize.axis === 'col' ? point.x : point.y);
      return;
    }
    this.readScroll();
    this.dragTo(drag);
    if (!this.autoScrollFrame) this.autoScrollFrame = requestAnimationFrame(this.autoScroll);
  };

  private onPointerUp = (event: MouseEvent): void => {
    const pointer = event as PointerEvent;
    const pointerId = pointer.pointerId === undefined ? 0 : pointer.pointerId;
    const touch = this.touch;
    if (touch && touch.pointerId === pointerId) {
      this.endTouch();
      this.readScroll();
      const point = this.canvasPoint(event);
      const still = Math.abs(point.x - touch.x) <= TAP_SLOP && Math.abs(point.y - touch.y) <= TAP_SLOP;
      if (touch.consumed || touch.moved || !still || touch.domX !== this.domX || touch.domY !== this.domY) return;
      if (this.pointTarget()) {
        // A formula is being typed and takes a reference here: the tap points at the cell, row or column.
        this.lastTap = undefined;
        if (hitTest(this.layout(), point.x, point.y).area !== 'corner') {
          this.pointBy(() => {
            this.press(point.x, point.y, false);
          });
        }
        return;
      }
      if (!this.commitForPress()) return;
      this.focusInput(true);
      if (this.pressCellButton(point.x, point.y, true)) {
        this.lastTap = undefined;
        return;
      }
      const mode = this.press(point.x, point.y, false);
      // A second tap that selects the same cell (or the same merged cell) edits it; two quick taps on neighbours do not.
      const last = this.lastTap;
      const active = this.state().active;
      this.lastTap = mode === 'cell' ? { time: event.timeStamp, sheet: this.state().sheet, row: active.row, col: active.col } : undefined;
      if (
        mode === 'cell' &&
        last &&
        event.timeStamp - last.time < DOUBLE_TAP_MS &&
        last.sheet === this.state().sheet &&
        last.row === active.row &&
        last.col === active.col
      ) {
        this.lastTap = undefined;
        this.startEditing();
      }
      return;
    }
    const drag = this.drag;
    if (!drag || drag.pointerId !== pointerId) return;
    this.endDrag();
    if (drag.resize) {
      if (drag.moved) {
        const point = this.canvasPoint(event);
        this.applyResize(drag.resize, Math.max(Math.round((drag.resize.axis === 'col' ? point.x : point.y) - drag.resize.origin), 0));
      }
      return;
    }
    if (drag.point) return;
    // A click on a cell that is cut off by the edge brings it into view; with Shift that is the clicked cell, not the active one.
    if (!drag.moved && drag.mode === 'cell') this.reveal(this.state().focus, this.state().range);
  };

  private onPointerCancel = (event: Event): void => {
    const pointerId = (event as PointerEvent).pointerId;
    if (this.touch && this.touch.pointerId === pointerId) this.endTouch();
    if (this.drag && this.drag.pointerId === pointerId) this.endDrag();
  };

  private endDrag(): void {
    const drag = this.drag;
    this.drag = undefined;
    this.hideGuide();
    if (this.autoScrollFrame) cancelAnimationFrame(this.autoScrollFrame);
    this.autoScrollFrame = 0;
    window.removeEventListener('mousemove', this.onPointerMove);
    window.removeEventListener('mouseup', this.onPointerUp);
    if (drag && this.pointerEvents) {
      try {
        this.canvas.releasePointerCapture(drag.pointerId);
      } catch {
        // Already released.
      }
    }
  }

  /**
   * While a drag holds the pointer outside the area the cells scroll in (over
   * the headers or frozen cells, or beyond the canvas), scrolls that way once
   * per frame, faster the further out the pointer is.
   */
  private autoScroll = (): void => {
    this.autoScrollFrame = 0;
    const drag = this.drag;
    if (this.destroyed || !drag || !drag.moved) return;
    const state = this.state();
    const layout = this.layout(state);
    const speed = (position: number, start: number, end: number) => {
      const out = position < start ? position - start : position >= end ? position - end + 1 : 0;
      if (out === 0) return 0;
      return Math.sign(out) * Math.min(AUTO_SCROLL_MIN + Math.abs(out) / 2, AUTO_SCROLL_MAX);
    };
    const dx = drag.mode === 'cell' || drag.mode === 'col' ? speed(drag.x, layout.metrics.headerWidth + layout.frozenWidth, this.width) : 0;
    const dy = drag.mode === 'cell' || drag.mode === 'row' ? speed(drag.y, layout.metrics.headerHeight + layout.frozenHeight, this.height) : 0;
    if (dx !== 0 || dy !== 0) {
      this.readScroll();
      if (this.scrollToOffsets(Math.round(state.scrollX + dx), Math.round(state.scrollY + dy))) this.dragTo(drag);
    }
    this.autoScrollFrame = requestAnimationFrame(this.autoScroll);
  };
}
