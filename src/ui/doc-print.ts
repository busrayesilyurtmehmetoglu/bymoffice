// Printing a document: every page is painted on a canvas at the resolution of
// a printer, turned into a PNG and put, as an image, into an element that
// only shows in print; rules for the print medium hide the rest of the host
// page and set the paper to the size of the document's pages. The browser's
// own print dialog does the rest (and saves a PDF where it offers that).
//
// Nothing of the document becomes HTML here either: a page is a picture, and
// the only words in the element are the view's own ("Page 3").
//
// This module holds what does not depend on the view: the plan (resolution
// and size of every page within the limits on memory), the style sheet, the
// element with the pages, and the wait for the dialog. Nothing in it touches
// the DOM at import time; the document and the window are handed in.

/** Dots per inch the pages are painted at when nothing else is asked for. */
export const DEFAULT_PRINT_DPI = 200;
export const MIN_PRINT_DPI = 72;
export const MAX_PRINT_DPI = 600;
/** Px of a page per inch: the unit of the layout. */
const PAGE_DPI = 96;
/** Pixels of one page; a larger page is painted at a lower resolution. Phones allow canvases of 16.7 million pixels at most. */
export const MAX_PRINT_PAGE_PIXELS = 16000000;
export const MAX_PRINT_PAGE_PIXELS_COARSE = 8000000;
/** Pixels of all pages together; a longer document is painted at a lower resolution, down to MIN_PRINT_DPI. */
export const MAX_PRINT_TOTAL_PIXELS = 600000000;
export const MAX_PRINT_TOTAL_PIXELS_COARSE = 200000000;
/** Never more pages than this, however small they are. */
export const MAX_PRINT_PAGES = 2000;
/** Bytes of the PNGs of all pages together. */
export const MAX_PRINT_BYTES = 768 * 1024 * 1024;
export const MAX_PRINT_BYTES_COARSE = 192 * 1024 * 1024;
/** Paper sizes with a page rule of their own; pages of further sizes are fitted into the first size. */
export const MAX_PRINT_SIZES = 32;
/** How long the job waits for the end of a dialog that never says it was closed. */
export const PRINT_DIALOG_MAX_MS = 5 * 60 * 1000;

export interface DocumentPrintOptions {
  /**
   * Resolution the pages are painted at, in dots per inch. Default 200,
   * between 72 and 600. Long documents and very large pages are painted at
   * less, so that the pages stay within the limits on memory.
   */
  dpi?: number;
}

export interface DocumentPrintResult {
  /** Pages handed to the browser. */
  pages: number;
  /** Resolution the pages were painted at (single pages larger than the limit have less). */
  dpi: number;
}

export type DocumentPrintErrorCode = 'unavailable' | 'too-long' | 'empty' | 'cancelled' | 'failed';

/**
 * Why a document was not printed: 'unavailable' where the window has no
 * print(), 'too-long' for a document beyond the limits, 'empty' for one
 * without a page, 'cancelled' when the view was destroyed while the pages
 * were prepared, 'failed' for everything else.
 */
export class DocumentPrintError extends Error {
  readonly code: DocumentPrintErrorCode;
  constructor(code: DocumentPrintErrorCode, message: string) {
    super(message);
    this.name = 'DocumentPrintError';
    this.code = code;
  }
}

export interface PrintLimits {
  maxPagePixels: number;
  maxTotalPixels: number;
  maxBytes: number;
}

/** The limits for a device: those with a coarse pointer (phones, tablets) get less. */
export function printLimits(coarse: boolean): PrintLimits {
  return coarse
    ? { maxPagePixels: MAX_PRINT_PAGE_PIXELS_COARSE, maxTotalPixels: MAX_PRINT_TOTAL_PIXELS_COARSE, maxBytes: MAX_PRINT_BYTES_COARSE }
    : { maxPagePixels: MAX_PRINT_PAGE_PIXELS, maxTotalPixels: MAX_PRINT_TOTAL_PIXELS, maxBytes: MAX_PRINT_BYTES };
}

export interface PrintPagePlan {
  /** Size of the page in its own px (96 to the inch). */
  width: number;
  height: number;
  /** Size of its picture in pixels. */
  pixelWidth: number;
  pixelHeight: number;
  /** Which of the plan's paper sizes it has; -1 for a page that is fitted into the first size. */
  size: number;
}

export interface PrintPlan {
  /** Resolution of the pages after the limits. */
  dpi: number;
  pages: PrintPagePlan[];
  /** The distinct paper sizes in the order they first appear, in pt, rounded to hundredths. */
  sizes: { width: number; height: number }[];
}

const toPt = (px: number): number => Math.round(px * 0.75 * 100) / 100;

/**
 * Decides the resolution and the size of the picture of every page. Throws a
 * RangeError for a `dpi` that is no positive number and a DocumentPrintError
 * for a document without pages or beyond the limits.
 */
export function planPrint(pageSizes: readonly { width: number; height: number }[], options: DocumentPrintOptions = {}, limits: PrintLimits = printLimits(false)): PrintPlan {
  let dpi = DEFAULT_PRINT_DPI;
  if (options.dpi !== undefined) {
    if (typeof options.dpi !== 'number' || !(options.dpi > 0) || options.dpi === Infinity) throw new RangeError('The print resolution must be a positive number of dots per inch');
    dpi = Math.min(MAX_PRINT_DPI, Math.max(MIN_PRINT_DPI, options.dpi));
  }
  const count = pageSizes.length;
  if (count === 0) throw new DocumentPrintError('empty', 'The document has no page to print');
  if (count > MAX_PRINT_PAGES) throw new DocumentPrintError('too-long', `The document has ${count} pages; at most ${MAX_PRINT_PAGES} can be printed at once`);
  // Square inches of all pages: what the pixels are shared out over.
  let area = 0;
  for (let i = 0; i < count; i++) {
    const { width, height } = pageSizes[i];
    if (!(width >= 1) || !(height >= 1) || width === Infinity || height === Infinity) throw new DocumentPrintError('failed', `Page ${i + 1} has no usable size`);
    area += (width / PAGE_DPI) * (height / PAGE_DPI);
  }
  if (area * dpi * dpi > limits.maxTotalPixels) dpi = Math.sqrt(limits.maxTotalPixels / area);
  if (dpi < MIN_PRINT_DPI) throw new DocumentPrintError('too-long', `The document has too many pages to print (${count})`);

  const sizes: PrintPlan['sizes'] = [];
  const keys = new Map<string, number>();
  const pages: PrintPagePlan[] = [];
  for (let i = 0; i < count; i++) {
    const { width, height } = pageSizes[i];
    let scale = dpi / PAGE_DPI;
    if (width * height * scale * scale > limits.maxPagePixels) scale = Math.sqrt(limits.maxPagePixels / (width * height));
    const paper = { width: toPt(width), height: toPt(height) };
    const key = `${paper.width}x${paper.height}`;
    let size = keys.get(key);
    if (size === undefined) {
      size = sizes.length < MAX_PRINT_SIZES ? sizes.length : -1;
      if (size >= 0) {
        sizes.push(paper);
        keys.set(key, size);
      }
    }
    pages.push({ width, height, pixelWidth: Math.max(1, Math.round(width * scale)), pixelHeight: Math.max(1, Math.round(height * scale)), size });
  }
  return { dpi, pages, sizes };
}

/** Class of the element that holds the pages, a direct child of the body. */
export const PRINT_CLASS = 'bym-print';

/**
 * The style sheet of a print job. On screen the pages are not shown; in
 * print nothing but the pages is, each on paper of its own size:
 * the first size is that of the unnamed page rule, and with `namedPages`
 * (the browser knows the `page` property) every size has a named rule.
 */
export function printCss(plan: Pick<PrintPlan, 'sizes'>, namedPages: boolean): string {
  const sizes = plan.sizes;
  const first = sizes[0];
  const rules: string[] = [`@page{size:${first.width}pt ${first.height}pt;margin:0}`];
  if (namedPages && sizes.length > 1) {
    for (let i = 0; i < sizes.length; i++) {
      rules.push(`@page ${PRINT_CLASS}-${i}{size:${sizes[i].width}pt ${sizes[i].height}pt;margin:0}`);
      rules.push(`.${PRINT_CLASS}-s${i}{page:${PRINT_CLASS}-${i}}`);
    }
  }
  return (
    `@media screen{.${PRINT_CLASS}{display:none!important}}` +
    '@media print{' +
    rules.join('') +
    // Whatever the host page does with its root elements (a fixed height, hidden overflow, a flex layout, a dark background) must not reach the paper.
    'html,body{display:block!important;position:static!important;float:none!important;width:auto!important;height:auto!important;min-width:0!important;' +
    'min-height:0!important;max-width:none!important;max-height:none!important;margin:0!important;padding:0!important;border:0!important;' +
    'overflow:visible!important;background:#fff!important;transform:none!important;filter:none!important;columns:auto!important}' +
    `body>:not(.${PRINT_CLASS}){display:none!important}` +
    `.${PRINT_CLASS}{display:block!important;position:static!important;float:none!important;visibility:visible!important;opacity:1!important;width:auto!important;height:auto!important;` +
    'margin:0!important;padding:0!important;border:0!important;overflow:visible!important;background:#fff!important;transform:none!important;filter:none!important}' +
    `.${PRINT_CLASS}-sheet{display:block!important;position:relative!important;float:none!important;visibility:visible!important;opacity:1!important;box-sizing:border-box!important;` +
    'margin:0!important;padding:0!important;border:0!important;overflow:hidden!important;transform:none!important;' +
    'break-inside:avoid;page-break-inside:avoid;break-after:page;page-break-after:always}' +
    `.${PRINT_CLASS}-sheet:last-child{break-after:auto;page-break-after:auto}` +
    `.${PRINT_CLASS}-img{display:block!important;position:absolute!important;visibility:visible!important;opacity:1!important;box-sizing:border-box!important;` +
    'min-width:0!important;min-height:0!important;max-width:none!important;max-height:none!important;margin:0!important;padding:0!important;border:0!important;' +
    'transform:none!important;filter:none!important;object-fit:fill!important}' +
    '}'
  );
}

/** Where a page goes on its sheet, in pt: the paper and the picture on it. */
export function sheetBox(plan: PrintPlan, index: number, namedPages: boolean): { width: number; height: number; left: number; top: number; imageWidth: number; imageHeight: number } {
  const page = plan.pages[index];
  const own = { width: toPt(page.width), height: toPt(page.height) };
  // The page has paper of its own size: the first size always, the others where pages can be named.
  if (page.size === 0 || (page.size > 0 && namedPages)) return { width: own.width, height: own.height, left: 0, top: 0, imageWidth: own.width, imageHeight: own.height };
  // Otherwise it is printed on paper of the first size: whole, in the middle, smaller when it does not fit.
  const paper = plan.sizes[0];
  const scale = Math.min(1, paper.width / own.width, paper.height / own.height);
  const imageWidth = Math.round(own.width * scale * 100) / 100;
  const imageHeight = Math.round(own.height * scale * 100) / 100;
  return {
    width: paper.width,
    height: paper.height,
    left: Math.round(((paper.width - imageWidth) / 2) * 100) / 100,
    top: Math.round(((paper.height - imageHeight) / 2) * 100) / 100,
    imageWidth,
    imageHeight,
  };
}

/** Whether the browser gives pages of different sizes paper of their own (the `page` property). */
export function supportsNamedPages(win: { CSS?: { supports?(property: string, value: string): boolean } }): boolean {
  try {
    return !!win.CSS && typeof win.CSS.supports === 'function' && win.CSS.supports('page', `${PRINT_CLASS}-0`);
  } catch {
    return false;
  }
}

/** The pages of one print job in the document: a style element and an element with one image per page. */
export class PrintFrame {
  readonly element: HTMLElement;
  private readonly style: HTMLElement;
  private readonly urls: string[] = [];
  /** Ends the wait for the picture that is being loaded. */
  private waiting: ((error: Error) => void) | undefined;
  private disposed = false;

  constructor(
    private readonly doc: Document,
    private readonly plan: PrintPlan,
    private readonly namedPages: boolean,
    private readonly revoke: (url: string) => void,
  ) {
    const style = doc.createElement('style');
    style.setAttribute('data-bymoffice-print', '');
    style.textContent = printCss(plan, namedPages);
    const element = doc.createElement('div');
    element.className = PRINT_CLASS;
    // Hidden on screen even where the style element is not applied.
    element.style.cssText = 'display:none';
    this.style = style;
    this.element = element;
    // At the level of the document, whatever shadow root the view is in: only there can the rest of the page be hidden around it.
    (doc.head || doc.documentElement).appendChild(style);
    (doc.body || doc.documentElement).appendChild(element);
  }

  /** Number of pages added so far. */
  get count(): number {
    return this.urls.length;
  }

  /**
   * Adds the next page: its picture, at a blob: or data: address, and its
   * label. Resolves once the browser has the picture; rejects when it cannot load it.
   */
  add(url: string, label: string): Promise<void> {
    const index = this.urls.length;
    if (this.disposed || index >= this.plan.pages.length) {
      this.revoke(url);
      return Promise.reject(new DocumentPrintError('cancelled', 'Printing was cancelled'));
    }
    this.urls.push(url);
    const page = this.plan.pages[index];
    const box = sheetBox(this.plan, index, this.namedPages);
    const sheet = this.doc.createElement('div');
    sheet.className = `${PRINT_CLASS}-sheet` + (page.size >= 0 ? ` ${PRINT_CLASS}-s${page.size}` : '');
    sheet.style.cssText = `width:${box.width}pt;height:${box.height}pt`;
    const image = this.doc.createElement('img');
    image.className = `${PRINT_CLASS}-img`;
    image.setAttribute('alt', label);
    image.setAttribute('width', String(page.pixelWidth));
    image.setAttribute('height', String(page.pixelHeight));
    image.style.cssText = `left:${box.left}pt;top:${box.top}pt;width:${box.imageWidth}pt;height:${box.imageHeight}pt`;
    const loaded = new Promise<void>((resolve, reject) => {
      this.waiting = reject;
      image.addEventListener('load', () => resolve());
      image.addEventListener('error', () => reject(new DocumentPrintError('failed', `The picture of page ${index + 1} could not be loaded`)));
    });
    image.setAttribute('src', url);
    sheet.appendChild(image);
    this.element.appendChild(sheet);
    return loaded;
  }

  /** Takes the pages and the style out of the document and gives the pictures back. Safe to call twice. */
  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    // A picture taken away while it loads may never say how that ended.
    if (this.waiting) this.waiting(new DocumentPrintError('cancelled', 'Printing was cancelled'));
    this.waiting = undefined;
    if (this.element.parentNode) this.element.parentNode.removeChild(this.element);
    if (this.style.parentNode) this.style.parentNode.removeChild(this.style);
    for (let i = 0; i < this.urls.length; i++) this.revoke(this.urls[i]);
    this.urls.length = 0;
  }
}

/** What the wait for the dialog needs of a window. */
export interface PrintWindow {
  print?: unknown;
  addEventListener(type: string, listener: () => void, capture?: boolean): void;
  removeEventListener(type: string, listener: () => void, capture?: boolean): void;
  setTimeout(callback: () => void, ms: number): unknown;
  clearTimeout(handle: never): void;
}

export interface PrintDialog {
  /** Resolves when the dialog was closed; rejects when print() threw. Printing and cancelling look the same. */
  closed: Promise<void>;
  /** Ends the wait now: the promise resolves. */
  finish(): void;
}

/**
 * Calls print() of the window and waits until its dialog is closed: until
 * the `afterprint` event. Where print() returns and no such event has come
 * (a window whose print() was replaced by the application, say), the dialog
 * counts as closed with the first key or pointer press in the window (none
 * gets there while a dialog is open) or after `maxMs`.
 */
export function openPrintDialog(win: PrintWindow, maxMs: number = PRINT_DIALOG_MAX_MS): PrintDialog {
  let over = false;
  let timer: unknown;
  let settle: () => void = () => undefined;
  const finish = () => {
    if (over) return;
    over = true;
    win.removeEventListener('afterprint', finish);
    win.removeEventListener('pointerdown', finish, true);
    win.removeEventListener('mousedown', finish, true);
    win.removeEventListener('touchstart', finish, true);
    win.removeEventListener('keydown', finish, true);
    if (timer !== undefined) win.clearTimeout(timer as never);
    settle();
  };
  const closed = new Promise<void>((resolve, reject) => {
    settle = resolve;
    if (typeof win.print !== 'function') {
      over = true;
      reject(new DocumentPrintError('unavailable', 'Printing is not available here: the window has no print()'));
      return;
    }
    win.addEventListener('afterprint', finish);
    try {
      (win as { print(): void }).print();
    } catch (error) {
      over = true;
      win.removeEventListener('afterprint', finish);
      reject(new DocumentPrintError('failed', `The print dialog could not be opened: ${error instanceof Error ? error.message : String(error)}`));
      return;
    }
    if (over) return;
    win.addEventListener('pointerdown', finish, true);
    win.addEventListener('mousedown', finish, true);
    win.addEventListener('touchstart', finish, true);
    win.addEventListener('keydown', finish, true);
    timer = win.setTimeout(finish, maxMs);
  });
  return { closed, finish };
}

/** The address of a picture that was made of a canvas, and its size in bytes. */
export interface CanvasPicture {
  url: string;
  bytes: number;
}

/**
 * The content of a canvas as a PNG at a blob: address; at a data: address
 * where canvases have no toBlob(). `frame` (requestAnimationFrame) is kept
 * going until the picture is there: some browsers encode in idle time, and a
 * page on which nothing moves has none for seconds.
 */
export function canvasPicture(canvas: HTMLCanvasElement, urls: { createObjectURL(blob: Blob): string }, frame?: (callback: () => void) => unknown): Promise<CanvasPicture> {
  return new Promise<CanvasPicture>((resolve, reject) => {
    const failed = () => reject(new DocumentPrintError('failed', 'A page could not be turned into a picture'));
    try {
      if (typeof canvas.toBlob === 'function') {
        let waiting = true;
        const tick = () => {
          if (waiting && frame) frame(tick);
        };
        canvas.toBlob((blob) => {
          waiting = false;
          if (!blob) return failed();
          try {
            resolve({ url: urls.createObjectURL(blob), bytes: blob.size });
          } catch {
            failed();
          }
        }, 'image/png');
        tick();
        return;
      }
      const url = canvas.toDataURL('image/png');
      // Not in the same turn: the pages are prepared one per turn, so that the page stays responsive.
      setTimeout(() => resolve({ url, bytes: url.length }), 0);
    } catch {
      failed();
    }
  });
}

/** Gives a picture made by canvasPicture() back. */
export function revokePicture(url: string, urls: { revokeObjectURL(url: string): void }): void {
  if (url.slice(0, 5) !== 'blob:') return;
  try {
    urls.revokeObjectURL(url);
  } catch {
    // Nothing depends on it.
  }
}
