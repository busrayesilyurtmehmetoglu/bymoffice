// Workbook view: shows one sheet of a workbook inside a container element and
// handles scrolling, selection, keyboard, touch and the sheet tabs.
//
// Scrolling is native: a scroll container holds the canvas (kept in place with
// sticky positioning) and a spacer that gives the container its scrollable
// size. The canvas only ever covers the visible area; drawing is done by
// renderGrid(). Nothing in this module touches the DOM at import time, so it
// can be imported where there is no document.

import { createLayout, hitTest, maxScroll, scrollAreaSize, scrollToReveal, type GridLayout } from '../layout/grid-geometry';
import { createSheetMetrics, type AxisMetrics, type SheetMetrics } from '../layout/sheet-metrics';
import { resolveColor } from '../model/sheet/color';
import { MAX_COLS, MAX_ROWS, cellName, getCell, parseCellRef, rangeName, type CellRange, type Sheet, type Workbook } from '../model/sheet/types';
import { cellText } from '../render/cell-text';
import { renderGrid, type GridSelection } from '../render/grid';

export interface WorkbookViewOptions {
  /** Sheet shown first; defaults to workbook.activeSheet, or the first visible sheet when that one is hidden. */
  sheet?: number;
  onSelectionChange?(selection: GridSelection, sheetIndex: number): void;
  onSheetChange?(sheetIndex: number): void;
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

const LABEL = 'Spreadsheet';
const SHEET_LABEL = 'Sheet';
const CHART_SHEET_NOTE = 'This sheet is a chart sheet. Charts are not shown yet.';
const OTHER_SHEET_NOTE = 'This kind of sheet is not shown yet.';

// Cosmetic rules only; everything the layout depends on is set inline so the
// view still works where a style element is not allowed.
const STYLE_TEXT =
  '.bym-view{font:12px/1.2 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;color:#202124;background:#fff}' +
  '.bym-scroller:focus,.bym-tab:focus{outline:2px solid #1a73e8;outline-offset:-2px}' +
  '.bym-scroller:focus:not(:focus-visible),.bym-tab:focus:not(:focus-visible){outline:none}' +
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

export class WorkbookView {
  readonly workbook: Workbook;

  private readonly container: HTMLElement;
  private readonly options: WorkbookViewOptions;
  private readonly root: HTMLDivElement;
  private readonly scroller: HTMLDivElement;
  private readonly canvas: HTMLCanvasElement;
  private readonly spacer: HTMLDivElement;
  private readonly tabs: HTMLDivElement;
  private readonly note: HTMLDivElement;
  private readonly live: HTMLDivElement;
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
  private resizeObserver: ResizeObserver | undefined;
  private dprQuery: MediaQueryList | undefined;
  private destroyed = false;

  constructor(container: HTMLElement, workbook: Workbook, options: WorkbookViewOptions = {}) {
    if (workbook.sheets.length === 0) throw new Error('The workbook has no sheets');
    this.container = container;
    this.workbook = workbook;
    this.options = options;
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
    // The cells are the grid and take the keyboard focus; Tab leaves them like any other control.
    scroller.tabIndex = 0;
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

    this.root = root;
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
    root.appendChild(note);
    root.appendChild(tabs);
    root.appendChild(live);

    this.styleHost = acquireStyle(container);
    container.appendChild(root);

    scroller.addEventListener('scroll', this.onScroll, { passive: true });
    scroller.addEventListener('keydown', this.onKeyDown);
    tabs.addEventListener('click', this.onTabClick);
    if (this.pointerEvents) {
      canvas.addEventListener('pointerdown', this.onPointerDown);
      canvas.addEventListener('pointermove', this.onPointerMove);
      canvas.addEventListener('pointerup', this.onPointerUp);
      canvas.addEventListener('pointercancel', this.onPointerCancel);
      canvas.addEventListener('lostpointercapture', this.onPointerCancel);
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
    this.endDrag();
    this.touch = undefined;
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
    this.touch = undefined;
    if (this.frame) cancelAnimationFrame(this.frame);
    this.frame = 0;
    if (this.resizeObserver) this.resizeObserver.disconnect();
    this.resizeObserver = undefined;
    this.unwatchPixelRatio();
    window.removeEventListener('resize', this.onResize);
    this.scroller.removeEventListener('scroll', this.onScroll);
    this.scroller.removeEventListener('keydown', this.onKeyDown);
    this.tabs.removeEventListener('click', this.onTabClick);
    const canvas = this.canvas;
    canvas.removeEventListener('pointerdown', this.onPointerDown);
    canvas.removeEventListener('pointermove', this.onPointerMove);
    canvas.removeEventListener('pointerup', this.onPointerUp);
    canvas.removeEventListener('pointercancel', this.onPointerCancel);
    canvas.removeEventListener('lostpointercapture', this.onPointerCancel);
    canvas.removeEventListener('mousedown', this.onPointerDown);
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
        this.setActiveSheet(i);
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
    if (event.detail > 0) this.scroller.focus({ preventScroll: true });
    this.setActiveSheet(index);
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
    return changed;
  }

  private onScroll = (): void => {
    if (this.destroyed || !this.readScroll()) return;
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
      devicePixelRatio: this.dpr,
    });
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
    const value = cellText(getCell(state.sheet, active.row, active.col));
    if (value !== '') text += ` ${value}`;
    const merge = mergeIn(this.merges(), active.row, active.col);
    const single = range.top === range.bottom && range.left === range.right;
    if (!single && !(merge && sameRange(merge, range))) text += `, ${rangeName(range)} selected`;
    this.live.textContent = text;
  }

  // -------------------------------------------------------------- keyboard

  private onKeyDown = (event: KeyboardEvent): void => {
    if (this.destroyed || event.defaultPrevented || event.isComposing) return;
    if (event.target !== this.scroller) return;
    const ctrl = event.ctrlKey || event.metaKey;
    const shift = event.shiftKey;
    if (event.altKey) {
      // Browsers keep Ctrl+PageUp/PageDown for their own tabs; Alt does the same here.
      if (ctrl || shift || (event.key !== 'PageUp' && event.key !== 'PageDown')) return;
      this.stepSheet(event.key === 'PageDown' ? 1 : -1);
      event.preventDefault();
      return;
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
      return cell !== undefined && cell.value !== null && cell.value !== '';
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

  private onPointerDown = (event: MouseEvent): void => {
    if (this.destroyed) return;
    const pointer = event as PointerEvent;
    const pointerId = pointer.pointerId === undefined ? 0 : pointer.pointerId;
    const point = this.canvasPoint(event);

    if (pointer.pointerType === 'touch') {
      // A touch that moves is a native scroll; only a tap selects (see onPointerUp).
      this.readScroll();
      this.touch = { pointerId, x: point.x, y: point.y, domX: this.domX, domY: this.domY, moved: false };
      return;
    }
    if (event.button !== 0) return;
    // Keeps the browser from starting a text selection or an image drag.
    event.preventDefault();
    this.scroller.focus({ preventScroll: true });
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
      if (Math.abs(point.x - touch.x) > TAP_SLOP || Math.abs(point.y - touch.y) > TAP_SLOP) touch.moved = true;
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
      this.touch = undefined;
      this.readScroll();
      const point = this.canvasPoint(event);
      const still = Math.abs(point.x - touch.x) <= TAP_SLOP && Math.abs(point.y - touch.y) <= TAP_SLOP;
      if (touch.moved || !still || touch.domX !== this.domX || touch.domY !== this.domY) return;
      this.scroller.focus({ preventScroll: true });
      this.press(point.x, point.y, false);
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
    if (this.touch && this.touch.pointerId === pointerId) this.touch = undefined;
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
