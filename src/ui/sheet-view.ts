// Workbook view: shows one sheet of a workbook inside a container element and
// handles scrolling, selection, keyboard, touch and the sheet tabs.
//
// Scrolling is native: a scroll container holds the canvas (kept in place with
// sticky positioning) and a spacer that gives the container its scrollable
// size. The canvas only ever covers the visible area; drawing is done by
// renderGrid(). Nothing in this module touches the DOM at import time, so it
// can be imported where there is no document.

import { createLayout, hitTest, maxScroll, rangeRect, scrollAreaSize, scrollToReveal, type GridLayout } from '../layout/grid-geometry';
import { createSheetMetrics, type AxisMetrics, type SheetMetrics } from '../layout/sheet-metrics';
import { MAX_COLS, MAX_ROWS, cellName, getCell, parseCellRef, rangeName, type CellRange, type Sheet, type Workbook } from '../model/sheet/types';
import { cellDisplayText } from '../format/cell-format';
import { getLocale, normalizeLocale, type FormatLocale } from '../format/locale';
import { resolveCellStyle, resolveColor } from '../model/sheet/color';
import { clearRange, copyRange, editTextAt, pasteAt, setCellInput, styleAt, type CopiedRange } from '../ops/commands';
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
  /** Called after every change to the workbook, including undo and redo. */
  onChange?(change: { sheets: number[] }): void;
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
  /** Rows or columns were inserted or deleted, so sizes and positions changed too. */
  structural: boolean;
}

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

interface Drag {
  pointerId: number;
  mode: 'cell' | 'row' | 'col' | 'all';
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
const DOUBLE_TAP_MS = 500;

const LABEL = 'Spreadsheet';
const INPUT_LABEL = 'Cell editor';
const ACTION_FAILED = 'The action could not be completed';
const SHEET_LABEL = 'Sheet';
const CHART_SHEET_NOTE = 'This sheet is a chart sheet. Charts are not shown yet.';
const OTHER_SHEET_NOTE = 'This kind of sheet is not shown yet.';

// Cosmetic rules only; everything the layout depends on is set inline so the
// view still works where a style element is not allowed.
const STYLE_TEXT =
  '.bym-view{font:12px/1.2 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;color:#202124;background:#fff}' +
  '.bym-focus-ring .bym-scroller,.bym-tab:focus{outline:2px solid #1a73e8;outline-offset:-2px}' +
  '.bym-tab:focus:not(:focus-visible){outline:none}' +
  '.bym-input{color:#000;background:#fff}' +
  '.bym-menu{min-width:160px;padding:4px 0;background:#fff;border:1px solid #dadce0;border-radius:4px;box-shadow:0 2px 8px rgba(0,0,0,.25)}' +
  '.bym-menu-item{-webkit-appearance:none;appearance:none;margin:0;border:0;border-radius:0;background:transparent;color:inherit;font:inherit;' +
  'text-align:left;padding:8px 16px;white-space:nowrap;cursor:pointer}' +
  '.bym-menu-item:hover,.bym-menu-item:focus{background:#e8eaed;outline:none}' +
  '.bym-tabs{background:#f1f3f4;border-top:1px solid #dadce0}' +
  '.bym-tab{-webkit-appearance:none;appearance:none;flex:none;margin:0;border:0;border-right:1px solid #dadce0;border-radius:0;' +
  'background:transparent;color:#444746;font:inherit;padding:0 14px;min-height:32px;white-space:nowrap;cursor:pointer}' +
  '.bym-tab:hover{background:#e8eaed}' +
  '.bym-tab[aria-selected="true"]{background:#fff;color:#0b57d0;font-weight:600}' +
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

export class WorkbookView {
  readonly workbook: Workbook;
  /** Undo and redo. Each step carries the sheet and selection to go back to. */
  readonly history: History<HistoryMark>;
  private locale: FormatLocale;
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
    this.locked = options.readOnly === true;
    this.history = new History<HistoryMark>(workbook);
    this.sheetIndex = this.initialSheet();
    this.pointerEvents = 'PointerEvent' in window;

    const root = document.createElement('div');
    root.className = 'bym-view';
    root.setAttribute('role', 'group');
    root.setAttribute('aria-label', LABEL);
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
    canvas.style.cssText = 'display:block;position:-webkit-sticky;top:0;left:0';
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
    input.setAttribute('aria-label', INPUT_LABEL);
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
      'position:absolute;z-index:2;left:0;top:0;width:1px;height:1px;margin:0;padding:0;border:0;outline:none;resize:none;overflow:hidden;' +
      'box-sizing:border-box;opacity:0;pointer-events:none;white-space:pre;line-height:1.2;border-radius:0';

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

  /** Changes the language of number and date formats and redraws. */
  setLocale(locale: string | FormatLocale): void {
    this.locale = resolveLocale(locale);
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

    this.scroller.setAttribute('aria-label', `${SHEET_LABEL} ${state.sheet.name}`);
    const kind = state.sheet.kind;
    this.note.textContent = kind === 'worksheet' ? '' : kind === 'chartsheet' ? CHART_SHEET_NOTE : OTHER_SHEET_NOTE;
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
    if (!single && !(merge && sameRange(merge, range))) text += `, ${rangeName(range)} selected`;
    this.live.textContent = text;
    this.syncInput();
  }

  // -------------------------------------------------------------- keyboard

  private onKeyDown = (event: KeyboardEvent): void => {
    if (this.destroyed || event.defaultPrevented) return;
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
      this.live.textContent = error instanceof Error && error.message ? error.message : ACTION_FAILED;
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
    this.applyOps([{ type: kind, sheet: this.sheetIndex, axis, at, count }], after, true);
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

  private mark(sheets: number[], structural: boolean, selection?: { active: CellPos; range: CellRange }): HistoryMark {
    const from = selection || this.state();
    return { sheet: this.sheetIndex, active: { ...from.active }, range: { ...from.range }, sheets, structural };
  }

  /**
   * Applies operations to the shown sheet as one undo step and brings the view
   * up to date. Returns false when there was nothing to apply. An operation
   * that throws leaves the workbook and the view as they were.
   */
  private applyOps(ops: Op[], after?: { active: CellPos; range: CellRange }, structural = false): boolean {
    if (ops.length === 0) return false;
    const sheets: number[] = [];
    for (let i = 0; i < ops.length; i++) if (sheets.indexOf(ops[i].sheet) < 0) sheets.push(ops[i].sheet);
    // Rows and columns that move take the formulas of every other sheet with them.
    if (structural) for (let i = 0; i < this.workbook.sheets.length; i++) if (sheets.indexOf(i) < 0) sheets.push(i);
    this.history.apply(ops, this.mark(sheets, structural), this.mark(sheets, structural, after));
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
    this.afterChange(meta.sheets, meta.structural, meta);
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
      this.live.textContent = error instanceof Error && error.message ? error.message : ACTION_FAILED;
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
  };

  private onCompositionStart = (): void => {
    if (this.destroyed || this.edit || this.locked || this.state().sheet.kind !== 'worksheet') return;
    // An input method composes in place; the editor has to be there from its first key.
    this.selectIdleText();
    this.beginEdit(this.input.value);
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
    // Leaving the editor any other way than Escape keeps what was typed.
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
    if (this.locked || this.state().sheet.kind !== 'worksheet') return false;
    const hit = hitTest(this.layout(), x, y);
    if (hit.area !== 'rowHeader' && hit.area !== 'colHeader') return false;
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
    const noun = rows ? (count === 1 ? 'row' : `${count} rows`) : count === 1 ? 'column' : `${count} columns`;
    const items: MenuItem[] = [
      { label: `Insert ${noun} ${rows ? 'above' : 'left'}`, action: () => this.attempt(() => this.structural('insert', axis, first, count)) },
      { label: `Insert ${noun} ${rows ? 'below' : 'right'}`, action: () => this.attempt(() => this.structural('insert', axis, first + count, count)) },
      { label: `Delete ${noun}`, action: () => this.attempt(() => this.structural('delete', axis, first, count)) },
    ];
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
    return true;
  }

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
    if (hitTest(this.layout(), point.x, point.y).area === 'cell') this.startEditing();
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
      this.touch = { pointerId, x: point.x, y: point.y, domX: this.domX, domY: this.domY, moved: false, consumed: false };
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
    const mode = this.press(point.x, point.y, event.shiftKey);
    this.drag = { pointerId, mode, x: point.x, y: point.y, originX: point.x, originY: point.y, moved: false };
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
  };

  private onPointerMove = (event: MouseEvent): void => {
    const pointer = event as PointerEvent;
    const pointerId = pointer.pointerId === undefined ? 0 : pointer.pointerId;
    const point = this.canvasPoint(event);
    const touch = this.touch;
    if (touch && touch.pointerId === pointerId) {
      if (Math.abs(point.x - touch.x) > TAP_SLOP || Math.abs(point.y - touch.y) > TAP_SLOP) {
        touch.moved = true;
        this.clearLongPress();
      }
      return;
    }
    const drag = this.drag;
    if (!drag || drag.pointerId !== pointerId) return;
    drag.x = point.x;
    drag.y = point.y;
    if (!drag.moved) {
      if (Math.abs(point.x - drag.originX) <= DRAG_SLOP && Math.abs(point.y - drag.originY) <= DRAG_SLOP) return;
      drag.moved = true;
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
