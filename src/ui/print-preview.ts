// The print preview of a workbook: a layer over the editor that shows the
// pages a print job has, as they will be printed. It does not lay out or
// paint anything itself: the view hands it the sizes of the pages (from the
// plan printing makes) and a function that paints one page on a canvas (the
// painting printing and PDF use), so what is previewed is what is printed.
//
// Only the pages that are in view (and half a screen around) have an element
// and a canvas; a canvas is given back when its page leaves, so a preview of
// 2,000 pages holds no more than a screenful (see preview-layout.ts). Pages are
// painted one after the other, the current one first, in slices that leave
// the page responsive.
//
// It is a modal dialog: the editor behind it is inert, the focus stays in it
// and goes back where it was when it closes. Keys: PageUp/PageDown the page
// before/after, Home/End the first/last, + and - the zoom, Escape closes,
// Ctrl+P prints. Two fingers zoom, as Ctrl+wheel (a pinch of a trackpad) does.
// Every text reaches the DOM through textContent. Nothing here touches the DOM at import time.

import {
  CanvasPool,
  MAX_PREVIEW_ZOOM,
  MIN_PREVIEW_ZOOM,
  PREVIEW_GAP,
  SettleClock,
  clampZoom,
  maxPreviewZoom,
  previewWindow,
  fitZoom,
  layoutPreview,
  pageAtScroll,
  previewBudget,
  previewPixelWidths,
  scrollForPage,
  previewTransition,
  stepZoom,
  type PreviewBudget,
  type PreviewLayout,
  type PreviewPageSize,
  type PreviewState,
} from './preview-layout';

export interface PrintPreviewStrings {
  title: string;
  /** The label of the area the pages are in. */
  pages: string;
  /** The label of one page: "Page 3 of 12". */
  page(number: number, total: number): string;
  /** What the counter shows: "3 / 12". */
  counter(number: number, total: number): string;
  previous: string;
  next: string;
  zoomIn: string;
  zoomOut: string;
  /** The zoom that shows a whole page. */
  fit: string;
  /** 100 %. */
  actualSize: string;
  /** The label of the zoom: "Zoom 75%". */
  zoom(percent: number): string;
  /** A percentage as the language writes it: "75%", "%75". */
  percent(value: number): string;
  print: string;
  pdf: string;
  pageSetup: string;
  close: string;
  /** The list of what is printed, and its entries. */
  which: string;
  active: string;
  all: string;
  planning: string;
}

export interface PrintPreviewOptions {
  /** The element the preview covers and is put into: the root element of the editor. */
  root: HTMLElement;
  strings: PrintPreviewStrings;
  sheets: 'active' | 'all';
  /** Whether "all sheets" may be chosen: there is more than one sheet to print. */
  canChooseSheets: boolean;
  /** Whether the page setup can be changed from here. */
  canPageSetup: boolean;
  coarse: boolean;
  onSheets(sheets: 'active' | 'all'): void;
  onPrint(): void;
  onPdf(): void;
  /** The page setup of the sheet of a page (the current one). */
  onPageSetup(page: number): void;
  /** Called once, when the preview has closed (by its button, Escape, or close()). */
  onClose(): void;
  /**
   * Paints a page on a canvas that is `pixelWidth` pixels across (its height
   * follows from the paper). Returns false when what it painted is still to
   * come (pictures, formats being worked out): the page is painted again a
   * little later.
   */
  paint(page: number, context: CanvasRenderingContext2D, pixelWidth: number): boolean;
  /** Called once before each turn of painting (several pages): the pictures of one turn are one pass of the store. */
  beginPass?(): void;
  /** Whether the page setup can be opened for a page (not for that of a chart sheet); all pages when missing. */
  pageSetupFor?(page: number): boolean;
}

/** What the preview shows now: for tests and for those who drive it. */
export interface PrintPreviewInfo {
  state: PreviewState;
  sheets: 'active' | 'all';
  pages: number;
  /** The current page, from 0. */
  page: number;
  zoom: number;
  fit: boolean;
  /** The pages that have an element now. */
  shown: number[];
  /** The pages whose canvas holds a painting of the pages as they are now. */
  painted: number[];
  canvases: number;
  /** Pixels all canvases hold. */
  pixels: number;
  message: string;
}

const CONTAINED = ['keypress', 'keyup', 'mousedown', 'mouseup', 'click', 'dblclick', 'pointerdown', 'pointerup', 'touchstart', 'touchend', 'contextmenu', 'copy', 'cut', 'paste', 'input', 'beforeinput', 'compositionstart', 'compositionend'];
const stop = (event: Event): void => event.stopPropagation();

/** How long a page that is not settled is painted again for, and how often. */
const SETTLE_MS = 30000;
const SETTLE_EVERY_MS = 120;
/** Time painting may take in one turn before the page gets a say. */
const SLICE_MS = 12;

interface Slot {
  element: HTMLDivElement;
  canvas?: HTMLCanvasElement;
  /** What it was painted for: `${generation}:${pixelWidth}`; '' when not yet. */
  key: string;
  /** How long it is painted again while what it shows is still to come. */
  settle: SettleClock;
  retry?: ReturnType<typeof setTimeout>;
}

let nextId = 1;

export class PrintPreview {
  private readonly options: PrintPreviewOptions;
  private strings: PrintPreviewStrings;
  private readonly backdrop: HTMLDivElement;
  private readonly scroller: HTMLDivElement;
  private readonly content: HTMLDivElement;
  private readonly message: HTMLDivElement;
  private readonly counter: HTMLSpanElement;
  private readonly zoomLabel: HTMLButtonElement;
  private readonly sheetsSelect: HTMLSelectElement;
  private readonly buttons: { [key: string]: HTMLButtonElement } = {};
  private readonly silenced: Element[] = [];
  private readonly before: HTMLElement | null;
  private unplaced: string | undefined;
  private readonly budget: PreviewBudget;
  private readonly pool: CanvasPool<HTMLCanvasElement>;
  private readonly slots = new Map<number, Slot>();

  private status: PreviewState = 'planning';
  private sheets: 'active' | 'all';
  private sizes: PreviewPageSize[] = [];
  private layout: PreviewLayout | undefined;
  /** The size of the view the layout was made for. */
  private laidWidth = -1;
  private laidHeight = -1;
  private zoom = 1;
  /** The largest zoom these pages may have (see maxPreviewZoom). */
  private maxZoom = MAX_PREVIEW_ZOOM;
  private fit = true;
  private current = 0;
  /** Counts the times the pages changed: a painting of an older one is out of date. */
  private generation = 0;
  /** Scroll the next frame brings the current page to (after a jump or a change of zoom); -1 for none. */
  private anchor = -1;
  /** Where in its page the anchor is, as a share of its height. */
  private anchorShare = 0;
  /** The scroll position the preview put the pages at itself: until the user moves them, the current page is the one it chose. */
  private placedTop = -1;
  private frame: number | undefined;
  private pumpTimer: ReturnType<typeof setTimeout> | undefined;
  private queue: number[] = [];
  private widths = new Map<number, number>();
  /** Two fingers on the pages: how far apart they started, and the zoom then. */
  private pinch = { distance: 0, zoom: 1 };
  private closed = false;
  private observer: ResizeObserver | undefined;

  constructor(options: PrintPreviewOptions) {
    this.options = options;
    this.strings = options.strings;
    this.sheets = options.sheets;
    this.budget = previewBudget(options.coarse);
    this.pool = new CanvasPool(() => {
      const canvas = document.createElement('canvas');
      canvas.className = 'bym-preview-canvas';
      canvas.style.cssText = 'display:block;width:100%;height:100%';
      canvas.setAttribute('aria-hidden', 'true');
      return canvas;
    });
    const id = nextId++;
    const coarse = options.coarse;
    const doc = options.root.ownerDocument || document;
    const rootNode = options.root.getRootNode() as Document | ShadowRoot;
    this.before = ((rootNode && rootNode.activeElement) || doc.activeElement) as HTMLElement | null;

    const backdrop = doc.createElement('div');
    backdrop.className = 'bym-preview';
    backdrop.setAttribute('role', 'dialog');
    backdrop.setAttribute('aria-modal', 'true');
    backdrop.style.cssText =
      'position:absolute;left:0;top:0;right:0;bottom:0;z-index:35;display:flex;flex-direction:column;box-sizing:border-box;overflow:hidden;' +
      'background:#e8eaed;color:#202124;font:13px/1.35 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;letter-spacing:normal;text-transform:none;text-align:left;direction:ltr';
    const title = doc.createElement('span');
    title.id = `bym-preview-title-${id}`;
    title.className = 'bym-preview-title';
    title.style.cssText = 'font-weight:600;margin-right:4px;white-space:nowrap';
    backdrop.setAttribute('aria-labelledby', title.id);

    const bar = doc.createElement('div');
    bar.className = 'bym-preview-bar';
    bar.style.cssText = 'display:flex;flex-wrap:wrap;align-items:center;gap:4px 6px;padding:6px 8px;background:#fff;border-bottom:1px solid #dadce0;flex:none';
    const size = coarse ? 40 : 30;
    const buttonCss = `-webkit-appearance:none;appearance:none;box-sizing:border-box;min-height:${size}px;min-width:${size}px;padding:0 10px;font:inherit;border:1px solid #dadce0;border-radius:6px;background:#fff;color:#202124;cursor:pointer;white-space:nowrap`;
    const button = (key: string): HTMLButtonElement => {
      const element = doc.createElement('button');
      element.type = 'button';
      element.className = `bym-preview-btn bym-preview-${key}`;
      element.setAttribute('data-bym-preview', key);
      element.style.cssText = buttonCss;
      this.buttons[key] = element;
      return element;
    };
    const group = (): HTMLSpanElement => {
      const element = doc.createElement('span');
      element.style.cssText = 'display:inline-flex;align-items:center;gap:4px';
      return element;
    };

    bar.appendChild(title);
    const actions = group();
    actions.appendChild(button('print'));
    actions.appendChild(button('pdf'));
    actions.appendChild(button('pageSetup'));
    bar.appendChild(actions);

    const select = doc.createElement('select');
    select.className = 'bym-preview-sheets';
    select.style.cssText = `box-sizing:border-box;height:${size}px;font:inherit;${coarse ? 'font-size:16px;' : ''}border:1px solid #dadce0;border-radius:4px;padding:0 6px;background:#fff;color:#202124`;
    this.sheetsSelect = select;
    bar.appendChild(select);

    const paging = group();
    paging.appendChild(button('previous'));
    const counter = doc.createElement('span');
    counter.className = 'bym-preview-counter';
    counter.setAttribute('aria-live', 'polite');
    counter.style.cssText = 'min-width:64px;text-align:center;font-variant-numeric:tabular-nums';
    this.counter = counter;
    paging.appendChild(counter);
    paging.appendChild(button('next'));
    bar.appendChild(paging);

    const zooming = group();
    zooming.appendChild(button('zoomOut'));
    this.zoomLabel = button('actualSize');
    zooming.appendChild(this.zoomLabel);
    zooming.appendChild(button('zoomIn'));
    zooming.appendChild(button('fit'));
    bar.appendChild(zooming);

    const spacer = doc.createElement('span');
    spacer.style.cssText = 'flex:1 1 0;min-width:0';
    bar.appendChild(spacer);
    bar.appendChild(button('close'));
    backdrop.appendChild(bar);

    const scroller = doc.createElement('div');
    scroller.className = 'bym-preview-pages';
    scroller.tabIndex = 0;
    scroller.setAttribute('role', 'region');
    // Fingers scroll here; two of them zoom the pages, not the page around the editor.
    scroller.style.cssText = 'position:relative;flex:1 1 0;min-height:0;overflow:auto;overscroll-behavior:contain;touch-action:pan-x pan-y;-webkit-overflow-scrolling:touch;outline:none';
    const content = doc.createElement('div');
    content.className = 'bym-preview-content';
    content.style.cssText = 'position:relative;margin:0;padding:0';
    scroller.appendChild(content);
    const message = doc.createElement('div');
    message.className = 'bym-preview-message';
    message.setAttribute('role', 'status');
    message.style.cssText = 'position:absolute;left:0;right:0;top:40%;text-align:center;color:#3c4043;font-size:14px;padding:0 16px;pointer-events:none';
    scroller.appendChild(message);
    backdrop.appendChild(scroller);
    this.scroller = scroller;
    this.content = content;
    this.message = message;
    this.backdrop = backdrop;

    for (const key of Object.keys(this.buttons)) this.buttons[key].addEventListener('click', this.onButton);
    select.addEventListener('change', this.onSelect);
    backdrop.addEventListener('keydown', this.onKeyDown);
    backdrop.addEventListener('mousedown', this.onMouseDown);
    backdrop.addEventListener('focusout', this.onFocusOut);
    for (const name of CONTAINED) backdrop.addEventListener(name, stop);
    scroller.addEventListener('scroll', this.onScroll, { passive: true });
    scroller.addEventListener('wheel', this.onWheel, { passive: false });
    // Touch events, not pointer events: a finger that scrolls has its pointer cancelled, and the second one would then be alone.
    scroller.addEventListener('touchstart', this.onTouchStart, { passive: true });
    scroller.addEventListener('touchmove', this.onTouchMove, { passive: false });
    scroller.addEventListener('touchend', this.onTouchEnd, { passive: true });
    scroller.addEventListener('touchcancel', this.onTouchEnd, { passive: true });
    const win = doc.defaultView;
    if (win && typeof (win as unknown as { ResizeObserver?: unknown }).ResizeObserver === 'function') {
      this.observer = new (win as unknown as { ResizeObserver: typeof ResizeObserver }).ResizeObserver(() => this.schedule());
      this.observer.observe(scroller);
    }

    this.applyStrings();

    // The editor behind the preview takes no input while it is open.
    const root = options.root;
    for (let i = 0; i < root.children.length; i++) {
      const child = root.children[i];
      if (child.hasAttribute('inert')) continue;
      child.setAttribute('inert', '');
      this.silenced.push(child);
    }
    if (win && typeof win.getComputedStyle === 'function' && win.getComputedStyle(root).position === 'static') {
      this.unplaced = root.style.position;
      root.style.position = 'relative';
    }
    root.appendChild(backdrop);
    this.showMessage(this.strings.planning);
    this.updateControls();
    scroller.focus({ preventScroll: true });
  }

  get isOpen(): boolean {
    return !this.closed;
  }

  get element(): HTMLElement {
    return this.backdrop;
  }

  /** The element the preview covers: dialogs opened from it go there too, so that they are not inside what it made inert. */
  get root(): HTMLElement {
    return this.options.root;
  }

  /** The words of the preview, in another language. */
  setStrings(strings: PrintPreviewStrings): void {
    this.strings = strings;
    this.applyStrings();
    this.updateControls();
    this.slots.forEach((slot, page) => slot.element.setAttribute('aria-label', strings.page(page + 1, this.sizes.length)));
  }

  /** The pages are being laid out (anew): what is shown stays until they come. */
  setPlanning(): void {
    if (this.closed) return;
    this.status = previewTransition(this.status, 'changed');
    if (this.sizes.length === 0) this.showMessage(this.strings.planning);
    this.updateControls();
  }

  /**
   * The pages to show, in px of their paper. The current page stays the same
   * page number (or the last one) unless `page` says another; every page is
   * painted anew.
   */
  setPages(sizes: readonly PreviewPageSize[], page?: number): void {
    if (this.closed) return;
    const first = this.sizes.length === 0;
    // The page stays where it was on the screen (worked out on the layout it is in now).
    if (!first && page === undefined) this.keepAnchor();
    this.sizes = sizes.slice();
    this.maxZoom = maxPreviewZoom(this.sizes);
    if (this.zoom > this.maxZoom) this.zoom = this.maxZoom;
    // A layout is the answer to a change, whether or not the change was said first.
    this.status = previewTransition(previewTransition(this.status, 'changed'), sizes.length > 0 ? 'planned' : 'nothing');
    this.generation++;
    this.layout = undefined;
    this.laidWidth = -1;
    const wanted = page !== undefined ? page : this.current;
    this.current = Math.max(0, Math.min(wanted, sizes.length - 1));
    if (first || page !== undefined) this.setAnchor(this.current, 0);
    else this.anchor = Math.min(this.anchor, this.current);
    // The elements stay, with what they show, until they are painted anew: no blank flash.
    this.showMessage('');
    this.updateControls();
    this.schedule();
  }

  /** Nothing to show: the reason is said in the middle. */
  setProblem(state: 'empty' | 'failed', text: string): void {
    if (this.closed) return;
    this.status = previewTransition(previewTransition(this.status, 'changed'), state === 'empty' ? 'nothing' : 'error');
    this.sizes = [];
    this.layout = undefined;
    this.laidWidth = -1;
    this.current = 0;
    this.generation++;
    this.clearSlots();
    this.content.style.width = '0px';
    this.content.style.height = '0px';
    this.showMessage(text);
    this.updateControls();
  }

  /** The pages look different now (pictures came, the language changed): every page is painted anew, where it is. */
  repaint(): void {
    if (this.closed) return;
    this.generation++;
    this.schedule();
  }

  /** What is shown, as chosen from outside: the list follows. */
  setSheets(sheets: 'active' | 'all'): void {
    if (this.closed) return;
    this.sheets = sheets;
    this.updateControls();
  }

  /** Whether the page setup button may be used. */
  setCanPageSetup(can: boolean): void {
    this.options.canPageSetup = can;
    this.updateControls();
  }

  /** Brings a page (from 0) into view and makes it the current one. */
  goTo(page: number): void {
    if (this.closed || this.sizes.length === 0) return;
    this.current = Math.max(0, Math.min(Math.floor(page), this.sizes.length - 1));
    this.setAnchor(this.current, 0);
    this.updateControls();
    this.schedule();
  }

  /** A zoom (1 = 100 %), or 'fit' for a whole page in view. */
  setZoom(zoom: number | 'fit'): void {
    if (this.closed) return;
    this.keepAnchor();
    if (zoom === 'fit') this.fit = true;
    else {
      this.fit = false;
      this.zoom = Math.min(clampZoom(zoom), this.maxZoom);
    }
    this.layout = undefined;
    this.updateControls();
    this.schedule();
  }

  info(): PrintPreviewInfo {
    const painted: number[] = [];
    // While a new layout waits for its frame (after a zoom, a resize, new pages) nothing is painted as it will be.
    if (this.layout) {
      this.slots.forEach((slot, page) => {
        if (slot.key === this.keyOf(page)) painted.push(page);
      });
    }
    painted.sort((a, b) => a - b);
    const shown = Array.from(this.slots.keys()).sort((a, b) => a - b);
    return {
      state: this.closed ? 'closed' : this.status,
      sheets: this.sheets,
      pages: this.sizes.length,
      page: this.current,
      zoom: this.layout ? this.layout.zoom : this.zoom,
      fit: this.fit,
      shown,
      painted,
      canvases: this.pool.size,
      pixels: this.pool.pixels,
      message: this.message.textContent || '',
    };
  }

  close(): void {
    if (this.closed) return;
    this.closed = true;
    this.status = previewTransition(this.status, 'close');
    if (this.frame !== undefined) cancelFrame(this.frame);
    this.frame = undefined;
    if (this.pumpTimer !== undefined) clearTimeout(this.pumpTimer);
    this.pumpTimer = undefined;
    if (this.observer) this.observer.disconnect();
    this.clearSlots();
    this.pool.clear();
    for (const key of Object.keys(this.buttons)) this.buttons[key].removeEventListener('click', this.onButton);
    this.sheetsSelect.removeEventListener('change', this.onSelect);
    const backdrop = this.backdrop;
    backdrop.removeEventListener('keydown', this.onKeyDown);
    backdrop.removeEventListener('mousedown', this.onMouseDown);
    backdrop.removeEventListener('focusout', this.onFocusOut);
    for (const name of CONTAINED) backdrop.removeEventListener(name, stop);
    this.scroller.removeEventListener('scroll', this.onScroll);
    this.scroller.removeEventListener('wheel', this.onWheel);
    this.scroller.removeEventListener('touchstart', this.onTouchStart);
    this.scroller.removeEventListener('touchmove', this.onTouchMove);
    this.scroller.removeEventListener('touchend', this.onTouchEnd);
    this.scroller.removeEventListener('touchcancel', this.onTouchEnd);
    for (const element of this.silenced) element.removeAttribute('inert');
    this.silenced.length = 0;
    if (backdrop.parentNode) backdrop.parentNode.removeChild(backdrop);
    if (this.unplaced !== undefined) this.options.root.style.position = this.unplaced;
    const before = this.before;
    this.options.onClose();
    // Back where the focus was before, unless the closing put it somewhere on purpose.
    if (before && before.isConnected && typeof before.focus === 'function') {
      const active = (this.options.root.getRootNode() as Document | ShadowRoot).activeElement;
      if (!active || active === before || !this.options.root.contains(active) || active === this.options.root) before.focus({ preventScroll: true });
    }
  }

  // ------------------------------------------------------------- building

  private applyStrings(): void {
    const s = this.strings;
    const title = this.backdrop.querySelector('.bym-preview-title');
    if (title) title.textContent = s.title;
    this.scroller.setAttribute('aria-label', s.pages);
    const label = (key: string, text: string, shown: string = text) => {
      const element = this.buttons[key];
      element.textContent = shown;
      element.setAttribute('aria-label', text);
      element.title = text;
    };
    label('print', s.print);
    label('pdf', s.pdf);
    label('pageSetup', s.pageSetup);
    label('previous', s.previous, '‹');
    label('next', s.next, '›');
    label('zoomOut', s.zoomOut, '−');
    label('zoomIn', s.zoomIn, '+');
    label('fit', s.fit);
    label('close', s.close, '×');
    const select = this.sheetsSelect;
    while (select.firstChild) select.removeChild(select.firstChild);
    const doc = select.ownerDocument || document;
    for (const [value, text] of [
      ['active', s.active],
      ['all', s.all],
    ]) {
      const option = doc.createElement('option');
      option.value = value;
      option.textContent = text;
      select.appendChild(option);
    }
    select.value = this.sheets;
    select.setAttribute('aria-label', s.which);
    select.title = s.which;
  }

  private updateControls(): void {
    if (this.closed) return;
    const s = this.strings;
    const total = this.sizes.length;
    const has = total > 0;
    // Written only when it changes: it is a live region, and every frame of a scroll comes here.
    const count = has ? s.counter(this.current + 1, total) : '';
    if (this.counter.textContent !== count) this.counter.textContent = count;
    this.setDisabled('previous', !has || this.current <= 0);
    this.setDisabled('next', !has || this.current >= total - 1);
    this.setDisabled('print', !has);
    this.setDisabled('pdf', !has);
    this.setDisabled('pageSetup', !this.options.canPageSetup || !has || (this.options.pageSetupFor !== undefined && !this.options.pageSetupFor(this.current)));
    const zoom = this.layout ? this.layout.zoom : this.zoom;
    this.setDisabled('zoomOut', !has || zoom <= MIN_PREVIEW_ZOOM + 1e-6);
    this.setDisabled('zoomIn', !has || zoom >= this.maxZoom - 1e-6);
    this.setDisabled('fit', !has || this.fit);
    this.setDisabled('actualSize', !has);
    const percent = Math.round(zoom * 100);
    const label = this.zoomLabel;
    // The zoom shown, and a press of it is 100 %.
    const shown = s.percent(percent);
    if (label.textContent !== shown) {
      label.textContent = shown;
      label.setAttribute('aria-label', `${s.zoom(percent)}. ${s.actualSize}`);
      label.title = s.actualSize;
    }
    this.sheetsSelect.disabled = !this.options.canChooseSheets;
    if (this.sheetsSelect.value !== this.sheets) this.sheetsSelect.value = this.sheets;
  }

  private setDisabled(key: string, disabled: boolean): void {
    const button = this.buttons[key];
    if (button && button.disabled !== disabled) {
      // A button that goes off while it has the focus hands it to the pages.
      const active = (this.backdrop.getRootNode() as Document | ShadowRoot).activeElement;
      button.disabled = disabled;
      if (disabled && active === button) this.scroller.focus({ preventScroll: true });
    }
  }

  private showMessage(text: string): void {
    if (this.message.textContent !== text) this.message.textContent = text;
    this.message.style.display = text ? 'block' : 'none';
  }

  // --------------------------------------------------------------- frames

  private schedule(): void {
    if (this.closed || this.frame !== undefined) return;
    this.frame = requestFrame(() => {
      this.frame = undefined;
      this.update();
    });
  }

  /** The current page and where in it the top of the view is: kept across a change of zoom or of width. */
  private keepAnchor(): void {
    const layout = this.layout;
    if (!layout || this.sizes.length === 0) {
      this.setAnchor(this.current, 0);
      return;
    }
    const slot = layout.slots[this.current];
    if (!slot) return this.setAnchor(this.current, 0);
    const share = slot.height > 0 ? (this.scroller.scrollTop - (slot.y - PREVIEW_GAP)) / slot.height : 0;
    this.setAnchor(this.current, Math.max(0, Math.min(1, share)));
  }

  private setAnchor(page: number, share: number): void {
    this.anchor = page;
    this.anchorShare = share;
  }

  private update(): void {
    if (this.closed) return;
    const scroller = this.scroller;
    const width = scroller.clientWidth;
    const height = scroller.clientHeight;
    if (this.sizes.length === 0 || width <= 0 || height <= 0) return;
    // A whole page in view depends on the height as well.
    if (!this.layout || width !== this.laidWidth || (this.fit && height !== this.laidHeight)) {
      if (this.layout && this.anchor < 0) this.keepAnchor();
      const zoom = Math.min(this.fit ? fitZoom(this.sizes[this.current], width, height) : this.zoom, this.maxZoom);
      this.zoom = zoom;
      this.layout = layoutPreview(this.sizes, zoom, width);
      this.laidWidth = width;
      this.laidHeight = height;
      // Every page is painted anew at its new size, and has its own time again to wait for its pictures.
      this.generation++;
      this.content.style.width = `${this.layout.width}px`;
      this.content.style.height = `${this.layout.height}px`;
      // The pages moved: each element goes to its new place, and is painted at its new size.
      this.slots.forEach((slot, page) => this.place(slot, page));
      this.updateControls();
    }
    const layout = this.layout;
    if (this.anchor >= 0) {
      const slot = layout.slots[Math.min(this.anchor, layout.slots.length - 1)];
      const top = this.anchorShare > 0 && slot ? slot.y - PREVIEW_GAP + this.anchorShare * slot.height : scrollForPage(layout, this.anchor);
      this.anchor = -1;
      scroller.scrollTop = top;
      if (slot) {
        // Across too: the page (or the middle of it) in view.
        const left = slot.x + slot.width / 2 - width / 2;
        scroller.scrollLeft = Math.max(0, left);
      }
      // As the browser has it (rounded, or held at the end).
      this.placedTop = scroller.scrollTop;
    }
    const top = scroller.scrollTop;
    // Moved by hand: the current page is the one in view.
    if (top !== this.placedTop) {
      this.placedTop = -1;
      this.current = pageAtScroll(layout, top, height, this.current);
    }
    // Paper for every page in view and around it; canvases for those in view, and for those around while the limit allows.
    const window = previewWindow(layout, top, height, this.current, this.budget.canvases);
    if (window.last < window.first) return;
    const painted = new Set(window.canvases);
    const inWindow = (page: number) => page >= window.first && page <= window.last;
    // The pages that left the window, or have no canvas in it now, give theirs back.
    this.pool.keepOnly((page) => painted.has(page));
    this.slots.forEach((slot, page) => {
      if (inWindow(page) && painted.has(page)) return;
      if (slot.retry !== undefined) clearTimeout(slot.retry);
      slot.retry = undefined;
      if (slot.canvas) {
        if (slot.canvas.parentNode === slot.element) slot.element.removeChild(slot.canvas);
        slot.canvas = undefined;
        slot.key = '';
      }
      if (inWindow(page)) return;
      if (slot.element.parentNode) slot.element.parentNode.removeChild(slot.element);
      this.slots.delete(page);
    });
    const ratio = (scroller.ownerDocument && scroller.ownerDocument.defaultView && scroller.ownerDocument.defaultView.devicePixelRatio) || 1;
    const widths = previewPixelWidths(layout.slots, this.sizes, window.canvases, ratio, this.budget);
    this.widths.clear();
    for (let i = 0; i < window.canvases.length; i++) this.widths.set(window.canvases[i], widths[i]);
    const total = this.sizes.length;
    for (let page = window.first; page <= window.last; page++) {
      if (this.slots.has(page)) continue;
      const element = (scroller.ownerDocument || document).createElement('div');
      element.className = 'bym-preview-page';
      element.setAttribute('role', 'img');
      element.setAttribute('data-page', String(page));
      element.setAttribute('aria-label', this.strings.page(page + 1, total));
      // The paper: white, with its edge.
      element.style.cssText = 'position:absolute;box-sizing:border-box;background:#fff;box-shadow:0 1px 3px rgba(60,64,67,.3),0 0 0 1px rgba(60,64,67,.15);overflow:hidden';
      const slot: Slot = { element, key: '', settle: new SettleClock(SETTLE_MS) };
      this.slots.set(page, slot);
      this.place(slot, page);
      this.content.appendChild(element);
    }
    this.updateControls();
    this.queuePaint(window.canvases);
  }

  private place(slot: Slot, page: number): void {
    const box = this.layout ? this.layout.slots[page] : undefined;
    if (!box) return;
    const label = this.strings.page(page + 1, this.sizes.length);
    if (slot.element.getAttribute('aria-label') !== label) slot.element.setAttribute('aria-label', label);
    const style = slot.element.style;
    style.left = `${box.x}px`;
    style.top = `${box.y}px`;
    style.width = `${box.width}px`;
    style.height = `${box.height}px`;
  }

  /** The pages that are out of date, the current one first and then by their distance from it. */
  private queuePaint(indexes: number[]): void {
    const wanted = new Set<number>();
    const add = (page: number) => {
      const slot = this.slots.get(page);
      if (slot && slot.key !== this.keyOf(page) && slot.retry === undefined) wanted.add(page);
    };
    for (let i = 0; i < this.queue.length; i++) add(this.queue[i]);
    for (let i = 0; i < indexes.length; i++) add(indexes[i]);
    const queue = Array.from(wanted);
    const current = this.current;
    queue.sort((a, b) => Math.abs(a - current) - Math.abs(b - current) || a - b);
    this.queue = queue;
    if (queue.length > 0 && this.pumpTimer === undefined) this.pumpTimer = setTimeout(this.pump, 0);
  }

  private keyOf(page: number): string {
    return `${this.generation}:${this.widths.get(page) || 0}`;
  }

  private pump = (): void => {
    this.pumpTimer = undefined;
    if (this.closed) return;
    const started = now();
    // One turn of painting is one pass of the pictures: those of the pages of this turn are kept together.
    if (this.options.beginPass) this.options.beginPass();
    while (this.queue.length > 0) {
      const page = this.queue.shift() as number;
      this.paintPage(page);
      if (now() - started > SLICE_MS) break;
    }
    if (this.queue.length > 0) this.pumpTimer = setTimeout(this.pump, 0);
  };

  private paintPage(page: number): void {
    const slot = this.slots.get(page);
    const size = this.sizes[page];
    const pixelWidth = this.widths.get(page);
    if (!slot || !size || !pixelWidth) return;
    const key = this.keyOf(page);
    if (slot.key === key) return;
    const canvas = this.pool.take(page);
    const pixelHeight = Math.max(1, Math.round((pixelWidth * size.height) / size.width));
    if (canvas.width !== pixelWidth) canvas.width = pixelWidth;
    if (canvas.height !== pixelHeight) canvas.height = pixelHeight;
    if (canvas.parentNode !== slot.element) slot.element.appendChild(canvas);
    slot.canvas = canvas;
    const context = canvas.getContext('2d');
    if (!context) return;
    let settled = true;
    try {
      settled = this.options.paint(page, context, pixelWidth);
    } catch {
      // A page that cannot be painted stays blank paper; the others go on.
      settled = true;
    }
    slot.key = key;
    if (settled) return;
    // Pictures or formats still to come: painted again in a while, for 30 s of this layout at most - however often.
    if (!slot.settle.again(this.generation, now())) return;
    slot.retry = setTimeout(() => {
      slot.retry = undefined;
      if (this.closed || this.slots.get(page) !== slot) return;
      slot.key = '';
      this.queuePaint([page]);
    }, SETTLE_EVERY_MS);
  }

  private clearSlots(): void {
    this.slots.forEach((slot) => {
      if (slot.retry !== undefined) clearTimeout(slot.retry);
      if (slot.element.parentNode) slot.element.parentNode.removeChild(slot.element);
    });
    this.slots.clear();
    this.pool.keepOnly(() => false);
    this.queue = [];
  }

  // --------------------------------------------------------------- events

  private onButton = (event: Event): void => {
    const button = event.currentTarget as HTMLButtonElement;
    if (this.closed || button.disabled) return;
    // Safari does not focus a button that is clicked: the focus would be left on the page, outside the dialog,
    // and a dialog opened from here (the page setup) would give it back there.
    const active = (this.backdrop.getRootNode() as Document | ShadowRoot).activeElement;
    if (active !== button) button.focus({ preventScroll: true });
    switch (button.getAttribute('data-bym-preview')) {
      case 'print':
        this.options.onPrint();
        break;
      case 'pdf':
        this.options.onPdf();
        break;
      case 'pageSetup':
        this.options.onPageSetup(this.current);
        break;
      case 'previous':
        this.goTo(this.current - 1);
        break;
      case 'next':
        this.goTo(this.current + 1);
        break;
      case 'zoomIn':
        this.setZoom(stepZoom(this.layout ? this.layout.zoom : this.zoom, 1));
        break;
      case 'zoomOut':
        this.setZoom(stepZoom(this.layout ? this.layout.zoom : this.zoom, -1));
        break;
      case 'actualSize':
        this.setZoom(1);
        break;
      case 'fit':
        this.setZoom('fit');
        break;
      case 'close':
        this.close();
        break;
    }
  };

  /**
   * A press on the preview outside its controls (the bar, the title, the counter, the grey around the pages)
   * keeps the focus where it is: without this it would go to the page, outside the dialog, and the keys of the
   * preview (Escape, Tab, Ctrl+P) would no longer reach it. The list opens on a press, and the pages take the
   * focus (and their scroll bars the press).
   */
  private onMouseDown = (event: MouseEvent): void => {
    const target = event.target as Element | null;
    if (!target || typeof target.closest !== 'function') return;
    if (target.closest('select, option')) return;
    if (this.scroller.contains(target)) {
      if (this.activeElement() !== this.scroller) this.scroller.focus({ preventScroll: true });
      return;
    }
    // A button takes the focus on its click (see onButton), in every browser.
    event.preventDefault();
  };

  /** The focus that leaves for nowhere (the page itself) comes back to the pages; one that goes to a dialog over the preview is let go. */
  private onFocusOut = (event: FocusEvent): void => {
    if (this.closed || event.relatedTarget) return;
    setTimeout(() => {
      if (this.closed || this.backdrop.hasAttribute('inert')) return;
      const doc = this.backdrop.ownerDocument;
      if (doc && typeof doc.hasFocus === 'function' && !doc.hasFocus()) return;
      const active = this.activeElement();
      const body = doc ? doc.body : null;
      if (!active || active === body || (doc && active === doc.documentElement)) this.scroller.focus({ preventScroll: true });
    }, 0);
  };

  private activeElement(): Element | null {
    return (this.backdrop.getRootNode() as Document | ShadowRoot).activeElement;
  }

  private onSelect = (): void => {
    const value = this.sheetsSelect.value === 'all' ? 'all' : 'active';
    if (value === this.sheets) return;
    this.sheets = value;
    this.options.onSheets(value);
  };

  /** The controls the Tab key goes through, in order. */
  private focusables(): HTMLElement[] {
    const out: HTMLElement[] = [];
    const all = this.backdrop.querySelectorAll('button, select');
    for (let i = 0; i < all.length; i++) {
      const element = all[i] as HTMLButtonElement | HTMLSelectElement;
      if (!element.disabled) out.push(element);
    }
    out.push(this.scroller);
    return out;
  }

  private onKeyDown = (event: KeyboardEvent): void => {
    event.stopPropagation();
    if (this.closed || event.isComposing) return;
    const key = event.key;
    const target = event.target as HTMLElement | null;
    const inList = target === this.sheetsSelect;
    if (key === 'Escape' || key === 'Esc') {
      event.preventDefault();
      this.close();
      return;
    }
    if (key === 'Tab') {
      // The focus stays in the preview.
      const list = this.focusables();
      const at = target ? list.indexOf(target) : -1;
      const next = event.shiftKey ? (at <= 0 ? list.length - 1 : at - 1) : at < 0 || at >= list.length - 1 ? 0 : at + 1;
      event.preventDefault();
      list[next].focus({ preventScroll: true });
      return;
    }
    if ((event.ctrlKey || event.metaKey) && !event.altKey && !event.shiftKey && (key === 'p' || key === 'P')) {
      // The browser would print the host page.
      event.preventDefault();
      if (this.sizes.length > 0) this.options.onPrint();
      return;
    }
    if (event.ctrlKey || event.metaKey || event.altKey || inList) return;
    switch (key) {
      case 'PageDown':
        this.goTo(this.current + 1);
        break;
      case 'PageUp':
        this.goTo(this.current - 1);
        break;
      case 'Home':
        this.goTo(0);
        break;
      case 'End':
        this.goTo(this.sizes.length - 1);
        break;
      case '+':
      case '=':
        this.setZoom(stepZoom(this.layout ? this.layout.zoom : this.zoom, 1));
        break;
      case '-':
      case '_':
        this.setZoom(stepZoom(this.layout ? this.layout.zoom : this.zoom, -1));
        break;
      default:
        return;
    }
    event.preventDefault();
  };

  private onScroll = (): void => {
    this.schedule();
  };

  /** Ctrl+wheel (and the pinch of a trackpad, which browsers send as one) zooms the pages, not the page around them. */
  private onWheel = (event: WheelEvent): void => {
    event.stopPropagation();
    if (!event.ctrlKey || this.sizes.length === 0) return;
    event.preventDefault();
    const zoom = this.layout ? this.layout.zoom : this.zoom;
    const delta = Math.max(-100, Math.min(100, event.deltaMode === 1 ? event.deltaY * 16 : event.deltaY));
    this.setZoom(zoom * Math.exp(-delta * 0.01));
  };

  private onTouchStart = (event: TouchEvent): void => {
    if (event.touches.length !== 2) return;
    this.pinch.distance = touchDistance(event.touches);
    this.pinch.zoom = this.layout ? this.layout.zoom : this.zoom;
  };

  private onTouchMove = (event: TouchEvent): void => {
    if (event.touches.length !== 2 || this.pinch.distance <= 0 || this.sizes.length === 0) return;
    // Two fingers zoom the pages, not the page around the editor.
    if (event.cancelable) event.preventDefault();
    const zoom = clampZoom((this.pinch.zoom * touchDistance(event.touches)) / this.pinch.distance);
    const now = this.layout ? this.layout.zoom : this.zoom;
    // Small changes wait: each zoom lays the pages out anew.
    if (Math.abs(zoom - now) / now > 0.04) this.setZoom(zoom);
  };

  private onTouchEnd = (event: TouchEvent): void => {
    if (event.touches.length < 2) this.pinch.distance = 0;
  };
}

function touchDistance(touches: TouchList): number {
  const a = touches[0];
  const b = touches[1];
  return a && b ? Math.hypot(a.clientX - b.clientX, a.clientY - b.clientY) : 0;
}

const now = (): number => (typeof performance !== 'undefined' && typeof performance.now === 'function' ? performance.now() : Date.now());

function requestFrame(callback: () => void): number {
  return typeof requestAnimationFrame === 'function' ? requestAnimationFrame(callback) : (setTimeout(callback, 16) as unknown as number);
}

function cancelFrame(handle: number): void {
  if (typeof cancelAnimationFrame === 'function') cancelAnimationFrame(handle);
  clearTimeout(handle as unknown as ReturnType<typeof setTimeout>);
}
