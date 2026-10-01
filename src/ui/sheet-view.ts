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
import { MAX_COLS, MAX_ROWS, cellName, getCell, parseCellRef, rangeName, type CellRange, type Sheet, type Workbook } from '../model/sheet/types';
import { cellDisplayText } from '../format/cell-format';
import { getLocale, normalizeLocale, type FormatLocale } from '../format/locale';
import { getUiStrings, type UiStrings } from './strings';
import { resolveCellStyle, resolveColor, type ResolvedCellStyle } from '../model/sheet/color';
import { clearRange, copyRange, editTextAt, pasteAt, setCellInput, styleAt, type CopiedRange } from '../ops/commands';
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
import { History, type Op } from '../ops/ops';
import { DEFAULT_GRID_THEME, renderGrid, type GridSelection } from '../render/grid';
import { cssFont } from '../render/text';
import { ContextMenu, type MenuItem } from './context-menu';

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
}

/** Longest a recalculation holds the page before it leaves the rest for the next turn. */
const CALC_SLICE_MS = 250;

/** Above this many written cells a step recalculates every formula instead of tracking what depends on them. */
const MAX_RECALC_CELLS = 100000;

/** The cell being edited. */
interface Edit {
  sheet: number;
  row: number;
  col: number;
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

/** How long a touch has to rest on a header to open its menu, and how close in time and place two taps make a double tap. */
const LONG_PRESS_MS = 550;
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

/** Language of the view's own words: the tag given as `locale`, else the browser's. */
function resolveStrings(locale: string | FormatLocale | undefined): UiStrings {
  return getUiStrings(typeof locale === 'string' ? locale : typeof navigator !== 'undefined' ? navigator.language : undefined);
}

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
  '.bym-menu-item:disabled{color:#9aa0a6;background:transparent;cursor:default}' +
  '.bym-menu-separator{height:1px;margin:4px 0;background:#dadce0}' +
  '.bym-resize-guide{background:#1a7f45}' +
  '.bym-tabs{background:#f1f3f4;border-top:1px solid #dadce0}' +
  '.bym-tab{-webkit-appearance:none;appearance:none;display:block;float:none;flex:none;box-sizing:border-box;width:auto;height:auto;min-width:0;' +
  'max-width:none;max-height:none;margin:0;border:0;border-right:1px solid #dadce0;border-radius:0;' +
  'background:transparent;color:#444746;font:inherit;line-height:1.2;padding:0 14px;min-height:32px;white-space:nowrap;cursor:pointer}' +
  // A finger needs more to hit than a mouse pointer.
  '@media (pointer:coarse){.bym-tab{min-height:40px}}' +
  '.bym-tab:hover{background:#e8eaed}' +
  '.bym-tab[aria-selected="true"]{background:#fff;color:#1a7f45;font-weight:600}' +
  '.bym-note{color:#5f6368;font-size:14px}';

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
  private strings: UiStrings;
  private locked: boolean;

  private readonly container: HTMLElement;
  private readonly options: WorkbookViewOptions;
  private readonly root: HTMLDivElement;
  private readonly scroller: HTMLDivElement;
  private readonly canvas: HTMLCanvasElement;
  private readonly spacer: HTMLDivElement;
  private readonly tabs: HTMLDivElement;
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

  private readonly states: (SheetState | undefined)[] = [];
  private sheetIndex: number;

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
  private menu: ContextMenu | undefined;
  /** What copy() last put on the clipboard; pasting the same text back pastes these cells. */
  private copied: CopiedRange | undefined;
  private longPressTimer = 0;
  /** The cell the last tap selected, for telling a double tap: the second tap has to select the same cell. */
  private lastTap: { time: number; sheet: number; row: number; col: number } | undefined;
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
  private dprQuery: MediaQueryList | undefined;
  private destroyed = false;

  constructor(container: HTMLElement, workbook: Workbook, options: WorkbookViewOptions = {}) {
    if (workbook.sheets.length === 0) throw new Error('The workbook has no sheets');
    this.container = container;
    this.workbook = workbook;
    this.options = options;
    this.locale = resolveLocale(options.locale);
    this.strings = resolveStrings(options.locale);
    this.locked = options.readOnly === true;
    this.history = new History<HistoryMark>(workbook);
    this.sheetIndex = this.initialSheet();
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
    this.note = note;
    this.live = live;

    this.buildTabs();
    scroller.appendChild(canvas);
    scroller.appendChild(spacer);
    root.appendChild(scroller);
    root.appendChild(input);
    root.appendChild(note);
    root.appendChild(tabs);
    root.appendChild(live);

    this.styleHost = acquireStyle(container);
    container.appendChild(root);

    scroller.addEventListener('scroll', this.onScroll, { passive: true });
    input.addEventListener('keydown', this.onKeyDown);
    input.addEventListener('input', this.onInput);
    input.addEventListener('compositionstart', this.onCompositionStart);
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
    root.addEventListener('mousedown', this.onRootMouseDown);
    root.addEventListener('focusout', this.onFocusOut);
    if (options.focusScope) options.focusScope.addEventListener('focusout', this.onFocusOut);
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
    } else {
      // Browsers without pointer events: a tap arrives as the mouse events the browser makes up for it.
      canvas.addEventListener('mousedown', this.onPointerDown);
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
  }

  get activeSheet(): number {
    return this.sheetIndex;
  }

  /** Shows a sheet. Hidden sheets have no tab but can be shown this way. */
  setActiveSheet(index: number): void {
    if (!Number.isInteger(index) || index < 0 || index >= this.workbook.sheets.length) {
      throw new RangeError(`No sheet with index ${index}`);
    }
    if (this.destroyed || index === this.sheetIndex) return;
    // What is being typed belongs to the sheet that is being left.
    if (this.edit) this.commitEditing();
    this.closeMenu();
    this.endDrag();
    this.endTouch();
    this.lastTap = undefined;
    this.sheetIndex = index;
    this.showSheet();
    this.schedule();
    if (this.options.onSheetChange) this.options.onSheetChange(index);
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

  /** Changes the language of number and date formats, and of the view's own labels, and redraws. */
  setLocale(locale: string | FormatLocale): void {
    this.locale = resolveLocale(locale);
    // The view's own words follow a language tag; a locale object only changes formats.
    if (typeof locale === 'string') this.strings = resolveStrings(locale);
    this.root.setAttribute('aria-label', this.strings.spreadsheet);
    this.input.setAttribute('aria-label', this.strings.cellEditor);
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
    if (this.calcTimer !== undefined) clearTimeout(this.calcTimer);
    this.endDrag();
    this.endTouch();
    this.edit = undefined;
    this.closeMenu();
    const input = this.input;
    input.removeEventListener('keydown', this.onKeyDown);
    input.removeEventListener('input', this.onInput);
    input.removeEventListener('compositionstart', this.onCompositionStart);
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
    this.root.removeEventListener('mousedown', this.onRootMouseDown);
    this.root.removeEventListener('focusout', this.onFocusOut);
    if (this.options.focusScope) this.options.focusScope.removeEventListener('focusout', this.onFocusOut);
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
    if (this.root.parentNode) this.root.parentNode.removeChild(this.root);
    releaseStyle(this.styleHost);
    // Give the backing store back right away; some browsers keep it until collection otherwise.
    canvas.width = 0;
    canvas.height = 0;
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
    let state = this.states[this.sheetIndex];
    if (!state) {
      state = this.createState(this.workbook.sheets[this.sheetIndex]);
      this.states[this.sheetIndex] = state;
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
        this.attempt(() => this.setActiveSheet(i));
        return;
      }
    }
  }

  private onTabClick = (event: MouseEvent): void => {
    let node = event.target as HTMLElement | null;
    while (node && node !== this.tabs && !node.hasAttribute('data-sheet')) node = node.parentElement;
    if (!node || node === this.tabs) return;
    const index = Number(node.getAttribute('data-sheet'));
    // A pointer click hands the keyboard back to the cells; a tab chosen with the keyboard keeps the focus.
    if (event.detail > 0) this.focusInput(true);
    this.attempt(() => this.setActiveSheet(index));
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
    if (this.destroyed || !this.measure()) return;
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
    renderGrid(context, {
      workbook: this.workbook,
      sheet: state.sheet,
      layout: this.layout(state),
      selection: { active: state.active, range: state.range },
      locale: this.locale,
      devicePixelRatio: this.dpr,
    });
    // The input follows its cell through every scroll and resize.
    this.placeInput();
  }

  // ------------------------------------------------------------- selection

  private expandRange(range: CellRange): CellRange {
    return expandByMerges(this.merges(), range);
  }

  private setSelection(active: CellPos, range: CellRange, focus: CellPos, revealCell?: CellPos): void {
    const state = this.state();
    const changed = !sameCell(active, state.active) || !sameRange(range, state.range);
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
    if (this.edit) {
      this.onEditKey(event);
      return;
    }
    // 229: the key went to an input method, also where isComposing is not reported.
    if (event.isComposing || event.keyCode === 229) return;
    // Whatever is typed next replaces the idle text, and copy needs a selection in some browsers.
    this.selectIdleText();
    const ctrl = event.ctrlKey || event.metaKey;
    const shift = event.shiftKey;
    if (event.altKey) {
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
        this.startEditing();
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
    return strings.actionFailed;
  }

  /** Says something to the user: in the live region, and to the host page if it listens. */
  private notify(message: string, kind: 'error' | 'info'): void {
    this.live.textContent = message;
    if (this.options.onNotice) this.options.onNotice({ message, kind });
  }

  private emitEditing(): void {
    if (this.options.onEditingChange) this.options.onEditingChange({ editing: this.edit !== undefined, text: this.editText });
  }

  /** Applies a style patch to the selection as one undo step. No-op when read-only. */
  format(patch: StylePatch): void {
    if (this.destroyed || this.locked) return;
    this.commitEditing();
    this.applyOps(formatRange(this.workbook, this.sheetIndex, this.state().range, patch));
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
    if (!unchanged) {
      // The editor stays open until the cell is written: if that fails, what was typed is still there.
      this.committing = true;
      try {
        this.applyOps(setCellInput(this.workbook, edit.sheet, edit.row, edit.col, text, { locale: this.locale }));
      } finally {
        this.committing = false;
      }
    }
    this.endEdit();
    return true;
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
      this.commitEditing();
      this.closeMenu();
    }
    this.locked = next;
    this.input.readOnly = next;
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
   * steps before it are out of reach until the edit is over.
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
    return copied;
  }

  /** Copies the selection and empties it, as one undo step. Read-only views only copy. */
  cut(): CopiedRange {
    if (!this.destroyed && !this.locked) this.commitEditing();
    const copied = this.copy();
    if (!this.destroyed && !this.locked) {
      this.applyOps(clearRange(this.workbook, this.sheetIndex, this.state().range));
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
    if (value === '') return;
    const state = this.state();
    if (state.sheet.kind !== 'worksheet') return;
    const own = html !== undefined && this.copied !== undefined && this.copyId !== '' && String(html).indexOf(`data-bym-copy="${this.copyId}"`) >= 0;
    const result = pasteAt(this.workbook, this.sheetIndex, state.range, value, own ? this.copied : undefined, { locale: this.locale });
    const after = { active: { row: result.range.top, col: result.range.left }, range: result.range };
    if (!this.applyOps(result.ops, after)) this.restoreSelection(after);
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
    this.applyOps([{ type: kind, sheet: this.sheetIndex, axis, at, count }], after);
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
  private recalculate(cells: HistoryMark['cells']): number[] {
    if (this.options.calculate === false) return [];
    try {
      if (!this.engine) this.engine = new CalcEngine(this.workbook, { budget: { milliseconds: CALC_SLICE_MS } });
      if (!cells) this.engine.invalidate();
      const changed = cells ? this.engine.recalculate(cells) : this.engine.recalculate();
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
      try {
        sheets = engine.recalculate([]).map((entry) => entry.sheet);
        if (engine.incomplete) this.continueCalculation();
      } catch (error) {
        this.engine = undefined;
      }
      this.schedule();
      if (sheets.length > 0 && this.options.onChange) this.options.onChange({ sheets });
    }, 0);
  }

  /** Adds the sheets a recalculation wrote to the ones a step reports as changed. */
  private addSheets(sheets: number[], more: number[]): void {
    for (let i = 0; i < more.length; i++) if (sheets.indexOf(more[i]) < 0) sheets.push(more[i]);
  }

  /**
   * Applies operations to the shown sheet as one undo step and brings the view
   * up to date. Returns false when there was nothing to apply. An operation
   * that throws leaves the workbook and the view as they were.
   */
  private applyOps(ops: Op[], after?: { active: CellPos; range: CellRange }): boolean {
    if (ops.length === 0) return false;
    const sheets: number[] = [];
    // Anything but a change of cells can change sizes, hidden runs or merged cells of its sheet.
    let structural = false;
    let moves = false;
    for (let i = 0; i < ops.length; i++) {
      const op = ops[i];
      if (sheets.indexOf(op.sheet) < 0) sheets.push(op.sheet);
      if (op.type !== 'setCells') structural = true;
      if (op.type === 'insert' || op.type === 'delete') moves = true;
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
    this.addSheets(sheets, this.recalculate(cells));
    this.afterChange(sheets, structural, after);
    return true;
  }

  private travel(redo: boolean): boolean {
    if (this.destroyed || this.locked) return false;
    if (this.edit) {
      this.cancelEditing();
      return true;
    }
    this.closeMenu();
    const step = redo ? this.history.redo() : this.history.undo();
    if (!step) return false;
    const meta = step.meta;
    if (meta.sheet !== this.sheetIndex && meta.sheet >= 0 && meta.sheet < this.workbook.sheets.length) this.setActiveSheet(meta.sheet);
    const sheets = meta.sheets.slice();
    this.addSheets(sheets, this.recalculate(meta.cells));
    this.afterChange(sheets, meta.structural, meta);
    return true;
  }

  /** Everything that has to follow a change of the model. */
  private afterChange(sheets: number[], structural: boolean, selection?: { active: CellPos; range: CellRange }): void {
    if (structural) {
      this.lastTap = undefined;
      // Row heights, column widths and hidden runs moved; the merged ranges are indexed again on their own.
      for (let i = 0; i < sheets.length; i++) {
        const state = this.states[sheets[i]];
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

  private beginEdit(idle?: string): void {
    const state = this.state();
    const { row, col } = state.active;
    const style = resolveCellStyle(this.workbook.styles, styleAt(state.sheet, row, col));
    this.edit = {
      sheet: this.sheetIndex,
      row,
      col,
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
    this.edit = undefined;
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
    const cell = edit || state.active;
    const merge = mergeIn(this.merges(), cell.row, cell.col);
    const rect = rangeRect(layout, merge || { top: cell.row, left: cell.col, bottom: cell.row, right: cell.col });
    const paneLeft = layout.metrics.headerWidth + (cell.col < layout.frozenCols ? 0 : layout.frozenWidth);
    const paneTop = layout.metrics.headerHeight + (cell.row < layout.frozenRows ? 0 : layout.frozenHeight);
    const inView = rect.width > 0 && rect.height > 0 && rect.x >= paneLeft && rect.y >= paneTop && rect.x < this.width && rect.y < this.height;
    if (!edit || !inView) {
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
      this.beginEdit();
    }
    this.placeInput();
    this.emitEditing();
  };

  private onCompositionStart = (): void => {
    if (this.destroyed || this.edit || this.locked || this.state().sheet.kind !== 'worksheet') return;
    // An input method composes in place; the editor has to be there from its first key.
    this.selectIdleText();
    this.beginEdit(this.input.value);
    this.emitEditing();
  };

  private onFocus = (): void => {
    if (this.destroyed) return;
    if (!this.quietFocus) this.root.classList.add('bym-focus-ring');
    this.selectIdleText();
  };

  private onBlur = (event: FocusEvent): void => {
    if (this.destroyed || this.refocusing) return;
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
    if (this.destroyed || this.refocusing || !this.edit) return;
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
    if (typeof target.closest === 'function' && target.closest('button, .bym-menu')) return;
    event.preventDefault();
    this.focusInput(true);
  };

  private onEditKey(event: KeyboardEvent): void {
    // Enter that confirms an input method's candidate is not Enter for the cell.
    if (event.isComposing || event.keyCode === 229) return;
    const shift = event.shiftKey;
    switch (event.key) {
      case 'Enter':
        event.preventDefault();
        if (event.altKey) {
          const input = this.input;
          input.setRangeText('\n', input.selectionStart, input.selectionEnd, 'end');
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
      default:
      // Every other key is the editor's own: arrows move the caret, not the selection.
    }
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
  }

  /**
   * Opens the insert and delete menu for the row or column header at a canvas
   * point. Returns false when the point is not on a header.
   */
  private openHeaderMenu(x: number, y: number): boolean {
    if (this.state().sheet.kind !== 'worksheet') return false;
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
    }
    this.showMenu(items, x, y);
    return true;
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

  /** Paste from the menu: only the asynchronous clipboard can be read without a paste event, and only as text. */
  private menuPaste(): void {
    const unavailable = () => {
      if (!this.destroyed) this.notify(this.strings.pasteWithKeyboard, 'info');
    };
    const clipboard = navigator.clipboard;
    if (!clipboard || typeof clipboard.readText !== 'function') {
      unavailable();
      return;
    }
    try {
      clipboard.readText().then((text) => {
        if (!this.destroyed) this.attempt(() => this.paste(text));
      }, unavailable);
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
    // Keeps the browser from scrolling with a finger that is resizing.
    if (this.resizeAt(x, y, RESIZE_REACH_TOUCH) && event.cancelable) event.preventDefault();
  };

  private onContextMenu = (event: MouseEvent): void => {
    if (this.destroyed) return;
    const point = this.canvasPoint(event);
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
    if (resize) this.autoFit(resize);
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
      const resize = this.resizeAt(point.x, point.y, RESIZE_REACH_TOUCH);
      this.touch = { pointerId, x: point.x, y: point.y, domX: this.domX, domY: this.domY, moved: false, consumed: false, resize };
      this.longPressTimer = window.setTimeout(this.onLongPress, LONG_PRESS_MS);
      return;
    }
    if (event.button !== 0) return;
    // Keeps the browser from starting a text selection or an image drag.
    event.preventDefault();
    this.touchInput = false;
    if (!this.tryCommit()) return;
    this.focusInput(true);
    this.endDrag();
    this.readScroll();
    // A press on the border between two headers resizes; it neither selects nor starts a selection drag.
    const resize = this.resizeAt(point.x, point.y, RESIZE_REACH_MOUSE);
    if (resize) {
      this.beginResize(pointerId, point, resize);
      return;
    }
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
      const resize = touch.resize;
      if (resize && !touch.consumed && (Math.abs(point.x - touch.x) > DRAG_SLOP || Math.abs(point.y - touch.y) > DRAG_SLOP)) {
        // The finger on a border has started to move: from here on it resizes.
        this.endTouch();
        if (!this.tryCommit()) return;
        this.endDrag();
        this.beginResize(pointerId, { x: touch.x, y: touch.y }, resize);
        const drag = this.drag as Drag | undefined;
        if (drag) {
          drag.x = point.x;
          drag.y = point.y;
          drag.moved = true;
          this.showGuide(resize, resize.axis === 'col' ? point.x : point.y);
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
      if (!this.tryCommit()) return;
      this.focusInput(true);
      const mode = this.press(point.x, point.y, false);
      // A second tap that selects the same cell (or the same merged cell) edits it; two quick taps on neighbours do not.
      const last = this.lastTap;
      const active = this.state().active;
      this.lastTap = mode === 'cell' ? { time: event.timeStamp, sheet: this.sheetIndex, row: active.row, col: active.col } : undefined;
      if (
        mode === 'cell' &&
        last &&
        event.timeStamp - last.time < DOUBLE_TAP_MS &&
        last.sheet === this.sheetIndex &&
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
