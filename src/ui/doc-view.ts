// The document view: shows a Word document read-only, as a strip of pages
// that scrolls vertically. Pages are laid out in slices so that the first
// ones appear at once, and only the pages in sight (and one more on each
// side) have a canvas, so a document of thousands of pages costs no more
// memory on screen than one of three.
//
// Nothing of the document becomes HTML: text is drawn on canvases and reaches
// the DOM only as textContent or as an attribute value. Pictures are decoded
// from the package's own parts; nothing is ever fetched from the network.
//
// Nothing in this module touches the DOM at import time.

import { DocLayout } from '../layout/doc/layout';
import type { DocLayoutOptions, FontMetrics, ImageItem, LayoutFont, LayoutPage, LayoutProgress, TextItem, TextMeasurer } from '../layout/doc/types';
import { findWrapper, inFieldCode, paragraphText, safeHyperlinkTarget, type Block, type DocModel, type Paragraph, type RunItem, type Table } from '../model/doc/types';
import { renderPage } from '../render/doc';
import { IMAGE_BUDGET_BYTES, IMAGE_BUDGET_BYTES_COARSE, ImageStore } from '../render/images';
import { CanvasTextMeasurer } from '../render/doc-measure';
import { getUiStrings, type UiStrings } from './strings';

// The picture decoder lives with the renderers; its helpers are still reachable from here.
export { declaredImagePixels, declaredImageSize, decodedImageSize, sniffImageType } from '../render/images';

export type DocumentZoom = number | 'fit-width' | 'fit-page';

export interface DocumentViewOptions {
  /** The package the document was read from: where its pictures are. Without it pictures are shown as placeholders. */
  package?: { readPart(part: string): Promise<Uint8Array> };
  /** Language of the view's own words, as a language tag ("tr", "en-US"). Default: the browser's. */
  locale?: string | object;
  /**
   * 1 is 100%. 'fit-width' and 'fit-page' follow the size of the container.
   * Default: 100%, or less when the page is wider than the container.
   */
  zoom?: DocumentZoom;
  /** Show the bar with the page indicator and the zoom buttons. Default true. */
  toolbar?: boolean;
  /** Called when the page in view or the number of pages changes. `page` counts from 1. */
  onPageChange?(page: number, pageCount: number): void;
  /** Called with what the view tells the user; the same text goes to its live region. */
  onNotice?(notice: { message: string; kind: 'error' | 'info' }): void;
  /**
   * Called when a hyperlink is clicked or activated with Enter. Unless it
   * returns false the target is opened in a new window. Only http, https,
   * mailto and tel targets ever get here.
   */
  onLinkClick?(target: string): boolean | void;
  /** Measures text instead of the canvas measurer (tests). `cssFont` of the canvas measurer is still used for painting. */
  measurer?: TextMeasurer;
  /** Passed to the layout engine. */
  layout?: DocLayoutOptions;
}

export interface DocumentViewState {
  zoom: number;
  zoomMode: 'auto' | 'fit-width' | 'fit-page' | 'manual';
  page: number;
  pageCount: number;
  layoutDone: boolean;
  /** Indexes (from 0) of the pages that have a canvas. */
  mounted: number[];
  scrollX: number;
  scrollY: number;
  width: number;
  height: number;
  devicePixelRatio: number;
  /** Bytes of decoded pictures the view holds right now, and the most it ever holds. */
  imageBytes: number;
  imageBudget: number;
}

const MIN_ZOOM = 0.1;
const MAX_ZOOM = 5;
const ZOOM_STEPS = [0.1, 0.25, 0.33, 0.5, 0.67, 0.75, 0.8, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4, 5];
/** Space between pages and around the strip, in screen px. */
const GAP = 12;
const PAD = 12;
/** Pages kept ready beyond the ones in sight, on each side. */
const BUFFER_PAGES = 1;
/** Never more canvases than this, however small the pages are on screen. */
const MAX_MOUNTED = 24;
/** Browsers limit the size of an element; a taller strip is scrolled proportionally. */
const MAX_EXTENT = 8000000;
/** Size of a page the layout gives no usable size for (US Letter). */
const DEFAULT_PAGE = { width: 816, height: 1056 };
const MAX_PAGE_SIDE = 20000;
/** Backing store of one page; beyond it the page is drawn at a lower resolution. */
const MAX_CANVAS_PIXELS = 12000000;
const MAX_CANVAS_PIXELS_COARSE = 6000000;
const FIRST_SLICE_MS = 30;
/** How long the constructor may go on when the first slice ended without a page. */
const FIRST_PAGE_MS = 300;
const SLICE_MS = 12;
const LINE_SCROLL = 40;
/** Layouts started over because fonts arrived; more than this and the view stays as it is. */
const MAX_FONT_RELAYOUTS = 12;
const MAX_LINKS_PER_PAGE = 2000;
const MAX_TEXT_DEPTH = 64;
const SELECTION_FILL = 'rgba(26, 115, 232, 0.28)';

// Cosmetic rules only; what the layout depends on is set inline.
const STYLE_TEXT =
  '.bym-doc{font:13px/1.2 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;color:#202124;background:#e8eaed;' +
  'letter-spacing:normal;text-transform:none;text-align:left;-webkit-text-size-adjust:100%;text-size-adjust:100%}' +
  '.bym-doc-bar{background:#f8f9fa;border-bottom:1px solid #dadce0;color:#202124}' +
  // Sizes are stated in full: a page rule for every button must not reach into the view.
  '.bym-doc-btn{-webkit-appearance:none;appearance:none;display:inline-flex;align-items:center;justify-content:center;float:none;flex:none;' +
  'box-sizing:border-box;width:auto;min-width:32px;height:32px;min-height:0;max-width:none;max-height:none;margin:0 1px;padding:0 6px;border:0;' +
  'border-radius:6px;background:transparent;color:inherit;font:inherit;line-height:1.2;letter-spacing:normal;text-transform:none;text-indent:0;' +
  'text-shadow:none;text-decoration:none;opacity:1;cursor:pointer;-webkit-tap-highlight-color:transparent;touch-action:manipulation}' +
  '@media (pointer:coarse){.bym-doc-btn{min-width:40px;height:40px}}' +
  '@media (hover:hover){.bym-doc-btn:hover{background:rgba(60,64,67,.09)}}' +
  '.bym-doc-btn:active{background:rgba(60,64,67,.17)}' +
  '.bym-doc-btn[aria-pressed="true"]{background:#d2e3d8;color:#0f5c31}' +
  '.bym-doc-btn:focus{outline:2px solid #1a7f45;outline-offset:-2px}' +
  '.bym-doc-btn:focus:not(:focus-visible){outline:none}' +
  '.bym-doc-icon{display:block;flex:none;width:20px;height:20px;min-width:0;min-height:0;max-width:none;max-height:none;margin:0;padding:0;border:0;pointer-events:none}' +
  '.bym-doc-pages{color:#5f6368;white-space:nowrap}' +
  '.bym-doc-focus .bym-doc-scroller{outline:2px solid #1a7f45;outline-offset:-2px}' +
  '.bym-doc-page{background:#fff;box-shadow:0 1px 3px rgba(60,64,67,.3),0 1px 2px rgba(60,64,67,.15)}' +
  '.bym-doc-link{cursor:pointer;outline:none}' +
  '.bym-doc-link:focus{outline:2px solid #1a7f45;outline-offset:1px}' +
  '.bym-doc-link:focus:not(:focus-visible){outline:none}' +
  '.bym-doc-note{color:#5f6368;font-size:14px}';

/** One style element per document or shadow root, shared by the views in it. */
const styleUsers = new WeakMap<Node, { element: HTMLStyleElement; count: number }>();

function acquireStyle(container: HTMLElement): Node {
  // Older browsers have no getRootNode (and no shadow roots either): the document of the container is the root there.
  const root: Node = typeof container.getRootNode === 'function' ? container.getRootNode() : container.ownerDocument || document;
  const host: Node = root.nodeType === 9 ? (root as Document).head || (root as Document).documentElement : root;
  const entry = styleUsers.get(host);
  if (entry) {
    entry.count++;
    return host;
  }
  const element = document.createElement('style');
  element.setAttribute('data-bymoffice-doc', '');
  element.textContent = STYLE_TEXT;
  host.appendChild(element);
  styleUsers.set(host, { element, count: 1 });
  return host;
}

function releaseStyle(host: Node): void {
  const entry = styleUsers.get(host);
  if (!entry || --entry.count > 0) return;
  if (entry.element.parentNode) entry.element.parentNode.removeChild(entry.element);
  styleUsers.delete(host);
}

const HIDDEN = 'position:absolute;width:1px;height:1px;margin:-1px;padding:0;overflow:hidden;clip:rect(0 0 0 0);border:0';

const clamp = (value: number, low: number, high: number): number => (value < low ? low : value > high ? high : value);

const coarsePointer = (): boolean => typeof window.matchMedia === 'function' && window.matchMedia('(pointer:coarse)').matches;

function resolveStrings(locale: string | object | undefined): UiStrings {
  return getUiStrings(typeof locale === 'string' ? locale : typeof navigator !== 'undefined' ? navigator.language : undefined);
}

const SVG_NS = 'http://www.w3.org/2000/svg';

/** A line drawing on a 24 x 24 grid, built element by element. */
function icon(paths: string[]): SVGElement {
  const svg = document.createElementNS(SVG_NS, 'svg');
  svg.setAttribute('class', 'bym-doc-icon');
  svg.setAttribute('viewBox', '0 0 24 24');
  svg.setAttribute('width', '20');
  svg.setAttribute('height', '20');
  svg.setAttribute('fill', 'none');
  svg.setAttribute('stroke', 'currentColor');
  svg.setAttribute('stroke-width', '1.75');
  svg.setAttribute('stroke-linecap', 'round');
  svg.setAttribute('stroke-linejoin', 'round');
  svg.setAttribute('aria-hidden', 'true');
  svg.setAttribute('focusable', 'false');
  for (let i = 0; i < paths.length; i++) {
    const path = document.createElementNS(SVG_NS, 'path');
    path.setAttribute('d', paths[i]);
    svg.appendChild(path);
  }
  return svg;
}

// ---------------------------------------------------------------------- text

function blocksText(blocks: Block[], out: string[], depth: number): void {
  if (depth > MAX_TEXT_DEPTH) return;
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    if (block.type === 'paragraph') out.push(paragraphText(block));
    else if (block.type === 'table') {
      for (let r = 0; r < block.rows.length; r++) {
        const cells: string[] = [];
        const row = block.rows[r];
        for (let c = 0; c < row.cells.length; c++) {
          const inner: string[] = [];
          blocksText(row.cells[c].blocks, inner, depth + 1);
          cells.push(inner.join('\n'));
        }
        out.push(cells.join('\t'));
      }
    }
  }
}

/** The text of the body of a document: paragraphs on lines of their own, the cells of a table row separated by tabs. */
export function documentText(doc: DocModel): string {
  const out: string[] = [];
  blocksText(doc.body.blocks, out, 0);
  return out.join('\n');
}

/**
 * A place in the text of the document: before character `offset` (as in
 * paragraphText) of the paragraph at `path` of a story. The same as
 * TextItem.source, so a place does not depend on how the pages are painted.
 */
interface TextPosition {
  story: string;
  path: number[];
  offset: number;
}

/**
 * Length of the part of a path that says which flow of text a paragraph is
 * in: nothing for the story itself, the path up to and with the text box for
 * a paragraph in a text box (see TextItem.source). A selection stays in the
 * flow it began in.
 */
function flowPrefix(path: number[]): number {
  const marker = path.lastIndexOf(-1);
  return marker < 0 ? 0 : Math.min(marker + 3, path.length);
}

function sameFlow(a: { story: string; path: number[] }, b: { story: string; path: number[] }): boolean {
  if (a.story !== b.story) return false;
  const prefix = flowPrefix(a.path);
  if (flowPrefix(b.path) !== prefix) return false;
  for (let i = 0; i < prefix; i++) if (a.path[i] !== b.path[i]) return false;
  return true;
}

/** Document order of two paths of one flow. */
function comparePaths(a: number[], b: number[]): number {
  const shared = Math.min(a.length, b.length);
  for (let i = 0; i < shared; i++) if (a[i] !== b[i]) return a[i] - b[i];
  return a.length - b.length;
}

const comparePositions = (a: TextPosition, b: TextPosition): number => comparePaths(a.path, b.path) || a.offset - b.offset;

/** A selection put in order. */
interface TextRange {
  start: TextPosition;
  end: TextPosition;
  /** Length of the flow prefix of the paths. */
  prefix: number;
  /**
   * The two ends are in different cells of one table: whole cells are
   * selected, those of the rows and of the cells (by their number in the row)
   * between the two. `table` is the path of the table.
   */
  cells?: { table: number[]; row: number; lastRow: number; cell: number; lastCell: number };
}

function orderedRange(anchor: TextPosition, focus: TextPosition): TextRange | undefined {
  if (!sameFlow(anchor, focus)) return undefined;
  const order = comparePositions(anchor, focus);
  if (order === 0) return undefined;
  const start = order < 0 ? anchor : focus;
  const end = order < 0 ? focus : anchor;
  const range: TextRange = { start, end, prefix: flowPrefix(start.path) };
  const a = start.path;
  const b = end.path;
  // Down the tables both ends are in, as long as they are in the same cell.
  for (let i = range.prefix; i + 3 < a.length && i + 3 < b.length && a[i] === b[i]; i += 3) {
    if (a[i + 1] === b[i + 1] && a[i + 2] === b[i + 2]) continue;
    range.cells = {
      table: a.slice(0, i + 1),
      row: Math.min(a[i + 1], b[i + 1]),
      lastRow: Math.max(a[i + 1], b[i + 1]),
      cell: Math.min(a[i + 2], b[i + 2]),
      lastCell: Math.max(a[i + 2], b[i + 2]),
    };
    break;
  }
  return range;
}

/** The part of a text item that is in the range, as offsets into its text; undefined when none is. */
function selectedPart(range: TextRange, item: TextItem): [number, number] | undefined {
  const source = item.source;
  // Text the engine made up and repeated header rows are not part of the story.
  if (!source || item.repeat || !sameFlow(source, range.start)) return undefined;
  const path = source.path;
  const length = item.text.length;
  const cells = range.cells;
  if (cells) {
    const at = cells.table.length;
    if (path.length < at + 3) return undefined;
    for (let i = 0; i < at; i++) if (path[i] !== cells.table[i]) return undefined;
    return path[at] >= cells.row && path[at] <= cells.lastRow && path[at + 1] >= cells.cell && path[at + 1] <= cells.lastCell ? [0, length] : undefined;
  }
  const afterStart = comparePaths(path, range.start.path);
  const beforeEnd = comparePaths(path, range.end.path);
  if (afterStart < 0 || beforeEnd > 0) return undefined;
  const from = afterStart === 0 ? clamp(range.start.offset - source.offset, 0, length) : 0;
  const to = beforeEnd === 0 ? clamp(range.end.offset - source.offset, 0, length) : length;
  return to > from ? [from, to] : undefined;
}

/** The text boxes of the n-th drawing of a paragraph, counted as the layout engine counts them. */
function drawingBoxes(paragraph: Paragraph, wanted: number): Block[][] | undefined {
  let count = 0;
  const search = (items: RunItem[], code: (index: number) => boolean, depth: number): Block[][] | null | undefined => {
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      if (item.type === 'fieldChar' || code(i)) continue;
      if (item.type === 'alternate') {
        if (depth >= MAX_TEXT_DEPTH) continue;
        const inner = search(item.items, () => false, depth + 1);
        if (inner !== undefined) return inner;
      } else if (item.type === 'drawing' || item.type === 'pict' || item.type === 'object') {
        if (count++ === wanted) return item.textBoxes || null;
      }
    }
    return undefined;
  };
  for (const inline of paragraph.children) {
    if (inline.type !== 'run') continue;
    if (inline.wrap && (findWrapper(inline, 'del') || findWrapper(inline, 'moveFrom'))) continue;
    const found = search(inline.content, (index) => inFieldCode(inline, index), 0);
    if (found !== undefined) return found || undefined;
  }
  return undefined;
}

/** The blocks of a story of the document ("body", "header:<rId>", "footnote:<id>", ...). */
function storyBlocks(doc: DocModel, story: string): Block[] | undefined {
  if (story === 'body') return doc.body.blocks;
  const colon = story.indexOf(':');
  if (colon < 0) return undefined;
  const kind = story.slice(0, colon);
  const id = story.slice(colon + 1);
  if (kind === 'header' || kind === 'footer') {
    const part = (kind === 'header' ? doc.headers : doc.footers).get(id);
    return part ? part.blocks : undefined;
  }
  if (kind === 'footnote' || kind === 'endnote') {
    const notes = kind === 'footnote' ? doc.footnotes : doc.endnotes;
    const note = notes && /^-?\d{1,10}$/.test(id) ? notes.items.get(Number(id)) : undefined;
    return note ? note.blocks : undefined;
  }
  return undefined;
}

/** The table at a path (block index, then row, cell and block index for every table it is in) under `blocks`. */
function tableAt(blocks: Block[], path: number[]): Table | undefined {
  let list = blocks;
  for (let i = 0; i < path.length; i += 3) {
    const block = list[path[i]];
    if (!block || block.type !== 'table') return undefined;
    if (i + 1 >= path.length) return block;
    const row = block.rows[path[i + 1]];
    const cell = row ? row.cells[path[i + 2]] : undefined;
    if (!cell) return undefined;
    list = cell.blocks;
  }
  return undefined;
}

/** The blocks of the flow a path prefix names: those of the story, or of a text box in it. */
function flowBlocks(doc: DocModel, story: string, prefix: number[]): Block[] | undefined {
  let blocks: Block[] | undefined = storyBlocks(doc, story);
  let i = 0;
  // Every round follows one text box: the paragraph that anchors it (through the tables it is in), then the box.
  while (blocks && i < prefix.length) {
    let list: Block[] = blocks;
    let block: Block | undefined = list[prefix[i++]];
    while (block && block.type === 'table') {
      const table: Table = block;
      const row = table.rows[prefix[i]];
      const inner: Block[] | undefined = row && row.cells[prefix[i + 1]] ? row.cells[prefix[i + 1]].blocks : undefined;
      if (!inner) return undefined;
      list = inner;
      block = list[prefix[i + 2]];
      i += 3;
    }
    if (!block || block.type !== 'paragraph' || prefix[i] !== -1) return undefined;
    const boxes = drawingBoxes(block, prefix[i + 1]);
    blocks = boxes ? boxes[prefix[i + 2]] : undefined;
    i += 3;
  }
  return blocks;
}

/**
 * The text of `blocks` from the place (`from`, `fromOffset`) to (`to`,
 * `toOffset`), both paths relative to the blocks; a missing path stands for
 * the start or the end of the blocks. Paragraphs on lines of their own, the
 * cells of a table row separated by tabs: the rules of documentText().
 */
function blocksRangeText(blocks: Block[], from: number[] | undefined, fromOffset: number, to: number[] | undefined, toOffset: number, out: string[], depth: number): void {
  if (depth > MAX_TEXT_DEPTH) return;
  const first = from && from.length ? from[0] : 0;
  const last = to && to.length ? Math.min(to[0], blocks.length - 1) : blocks.length - 1;
  for (let i = Math.max(0, first); i <= last; i++) {
    const block = blocks[i];
    const starts = from !== undefined && from.length > 0 && i === from[0];
    const ends = to !== undefined && to.length > 0 && i === to[0];
    if (block.type === 'paragraph') {
      const text = paragraphText(block);
      const a = starts && from.length === 1 ? clamp(fromOffset, 0, text.length) : 0;
      const b = ends && to.length === 1 ? clamp(toOffset, 0, text.length) : text.length;
      out.push(text.slice(a, Math.max(a, b)));
    } else if (block.type === 'table') {
      const deepStart = starts && from.length >= 4;
      const deepEnd = ends && to.length >= 4;
      const firstRow = deepStart ? Math.max(0, from[1]) : 0;
      const lastRow = deepEnd ? Math.min(to[1], block.rows.length - 1) : block.rows.length - 1;
      for (let r = firstRow; r <= lastRow; r++) {
        const row = block.rows[r];
        const firstCell = deepStart && r === from[1] ? Math.max(0, from[2]) : 0;
        const lastCell = deepEnd && r === to[1] ? Math.min(to[2], row.cells.length - 1) : row.cells.length - 1;
        const cells: string[] = [];
        for (let c = firstCell; c <= lastCell; c++) {
          const inner: string[] = [];
          blocksRangeText(
            row.cells[c].blocks,
            deepStart && r === from[1] && c === from[2] ? from.slice(3) : undefined,
            fromOffset,
            deepEnd && r === to[1] && c === to[2] ? to.slice(3) : undefined,
            toOffset,
            inner,
            depth + 1,
          );
          cells.push(inner.join('\n'));
        }
        out.push(cells.join('\t'));
      }
    }
  }
}

/** The text of a selection, from the model. */
function rangeTextOf(doc: DocModel, range: TextRange): string {
  const prefix = range.prefix;
  const blocks = flowBlocks(doc, range.start.story, range.start.path.slice(0, prefix));
  if (!blocks) return '';
  const out: string[] = [];
  const cells = range.cells;
  if (cells) {
    const table = tableAt(blocks, cells.table.slice(prefix));
    if (!table) return '';
    for (let r = cells.row; r <= cells.lastRow && r < table.rows.length; r++) {
      const row = table.rows[r];
      const line: string[] = [];
      for (let c = cells.cell; c <= cells.lastCell && c < row.cells.length; c++) {
        const inner: string[] = [];
        blocksText(row.cells[c].blocks, inner, 1);
        line.push(inner.join('\n'));
      }
      out.push(line.join('\t'));
    }
  } else blocksRangeText(blocks, range.start.path.slice(prefix), range.start.offset, range.end.path.slice(prefix), range.end.offset, out, 0);
  return out.join('\n');
}

function sameSource(a: TextItem, b: TextItem): boolean {
  const x = a.source;
  const y = b.source;
  if (!x || !y || x.story !== y.story || x.path.length !== y.path.length) return false;
  for (let i = 0; i < x.path.length; i++) if (x.path[i] !== y.path[i]) return false;
  return true;
}

interface Slot {
  index: number;
  element: HTMLDivElement;
  canvas: HTMLCanvasElement;
  context: CanvasRenderingContext2D | null;
  text: HTMLDivElement;
  links: HTMLElement[];
  dirty: boolean;
  left: number;
  top: number;
  width: number;
  height: number;
}

export class DocumentView {
  readonly document: DocModel;
  private readonly options: DocumentViewOptions;
  private readonly container: HTMLElement;
  private readonly root: HTMLDivElement;
  private readonly scroller: HTMLDivElement;
  private readonly strip: HTMLDivElement;
  private readonly input: HTMLTextAreaElement;
  private readonly live: HTMLDivElement;
  private readonly note: HTMLDivElement;
  private readonly bar: HTMLDivElement | undefined;
  private readonly zoomOutButton: HTMLButtonElement | undefined;
  private readonly zoomInButton: HTMLButtonElement | undefined;
  private readonly zoomButton: HTMLButtonElement | undefined;
  private readonly fitButton: HTMLButtonElement | undefined;
  private readonly pagesLabel: HTMLSpanElement | undefined;
  private readonly styleHost: Node;
  private readonly painter: CanvasTextMeasurer;
  private readonly measurer: TextMeasurer;
  private readonly images: ImageStore;
  private readonly pointerEvents: boolean;
  private readonly maxCanvasPixels: number;
  private strings: UiStrings;
  private engine: DocLayout;
  private pages: LayoutPage[];
  /** Sum of the heights (px of the page) of the pages before each page; one entry more than there are pages. */
  private offsets: number[] = [0];
  private done = false;
  private failed = false;
  private truncationSaid = false;
  private paintFailed = false;

  private mode: 'auto' | 'fit-width' | 'fit-page' | 'manual' = 'auto';
  private scale = 1;
  private width = 0;
  private height = 0;
  private dpr = 1;
  private slots = new Map<number, Slot>();
  private linkTargets = new WeakMap<Element, string>();

  private frame = 0;
  private layoutTimer: ReturnType<typeof setTimeout> | undefined;
  private fontTimer: ReturnType<typeof setTimeout> | undefined;
  private fontRelayouts = 0;
  private resizeObserver: ResizeObserver | undefined;
  private dprQuery: MediaQueryList | undefined;
  private destroyed = false;

  private reportedPage = 0;
  private reportedCount = 0;
  /** A page asked for before it was laid out (from 0). */
  private pendingPage: number | undefined;
  private quietFocus = false;

  private selection: { anchor: TextPosition; focus: TextPosition } | 'all' | undefined;
  private dragging = false;
  private dragMoved = false;
  private pinch: { distance: number; zoom: number } | undefined;
  private gestureZoom = 0;

  constructor(container: HTMLElement, doc: DocModel, options: DocumentViewOptions = {}) {
    this.container = container;
    this.document = doc;
    this.options = options;
    this.strings = resolveStrings(options.locale);
    this.pointerEvents = 'PointerEvent' in window;
    const coarse = coarsePointer();
    this.maxCanvasPixels = coarse ? MAX_CANVAS_PIXELS_COARSE : MAX_CANVAS_PIXELS;

    this.painter = new CanvasTextMeasurer({ fonts: doc.fonts });
    this.measurer = options.measurer || this.painter;
    this.images = new ImageStore(options.package, coarse ? IMAGE_BUDGET_BYTES_COARSE : IMAGE_BUDGET_BYTES, this.onImageLoad);
    // A document that cannot be laid out at all throws here, before anything is added to the container.
    this.engine = new DocLayout(doc, this.measurer, options.layout);
    this.pages = this.engine.pages;

    const root = document.createElement('div');
    root.className = 'bym-doc';
    root.setAttribute('role', 'group');
    root.setAttribute('aria-label', this.strings.document);
    root.style.cssText =
      'position:relative;display:flex;flex-direction:column;width:100%;height:100%;box-sizing:border-box;overflow:hidden;direction:ltr;' +
      '-webkit-user-select:none;user-select:none';

    if (options.toolbar !== false) {
      const bar = document.createElement('div');
      bar.className = 'bym-doc-bar';
      bar.setAttribute('role', 'toolbar');
      bar.style.cssText = 'position:relative;display:flex;flex:none;align-items:center;box-sizing:border-box;padding:4px 8px;overflow:hidden';
      this.zoomOutButton = this.button(bar, 'bym-doc-zoom-out', icon(['M6 12h12']));
      this.zoomButton = this.button(bar, 'bym-doc-zoom-level');
      this.zoomButton.style.minWidth = '52px';
      this.zoomInButton = this.button(bar, 'bym-doc-zoom-in', icon(['M6 12h12', 'M12 6v12']));
      this.fitButton = this.button(bar, 'bym-doc-fit', icon(['M4 5v14', 'M20 5v14', 'M8 12h8', 'm10 9-2.5 3 2.5 3', 'm14 9 2.5 3-2.5 3']));
      const pagesLabel = document.createElement('span');
      pagesLabel.className = 'bym-doc-pages';
      pagesLabel.style.cssText = 'flex:1 1 0;min-width:0;margin:0 4px 0 12px;overflow:hidden;text-overflow:ellipsis;text-align:right';
      bar.appendChild(pagesLabel);
      this.pagesLabel = pagesLabel;
      this.bar = bar;
      root.appendChild(bar);
      bar.addEventListener('click', this.onBarClick);
    }

    const scroller = document.createElement('div');
    scroller.className = 'bym-doc-scroller';
    // The focus lives in the input element; some browsers would make a scrolling element a Tab stop of its own.
    scroller.tabIndex = -1;
    // The vertical bar is always there, so that fitting the width never changes whether there is one.
    scroller.style.cssText =
      'position:relative;flex:1 1 0;min-height:0;min-width:0;overflow-x:auto;overflow-y:scroll;overscroll-behavior:contain;' +
      '-webkit-overflow-scrolling:touch;touch-action:pan-x pan-y';

    const strip = document.createElement('div');
    strip.className = 'bym-doc-strip';
    strip.style.cssText = 'position:relative;box-sizing:content-box;margin:0;padding:0;border:0;overflow:hidden;width:1px;height:1px';

    const note = document.createElement('div');
    note.className = 'bym-doc-note';
    note.style.cssText = 'position:absolute;left:0;right:0;top:45%;padding:0 16px;text-align:center;pointer-events:none;display:none';

    const input = document.createElement('textarea');
    input.className = 'bym-doc-input';
    input.setAttribute('aria-label', this.strings.document);
    // Read-only and without an on-screen keyboard: it only holds the focus and takes the keys.
    input.readOnly = true;
    input.setAttribute('inputmode', 'none');
    input.setAttribute('autocomplete', 'off');
    input.setAttribute('autocorrect', 'off');
    input.setAttribute('autocapitalize', 'off');
    input.spellcheck = false;
    input.rows = 1;
    // Something to select: some browsers send no copy event when nothing is selected.
    input.value = ' ';
    input.style.cssText =
      'position:absolute;z-index:2;left:0;top:0;display:block;float:none;width:1px;height:1px;min-width:0;min-height:0;max-width:none;max-height:none;' +
      'margin:0;padding:0;border:0;outline:none;resize:none;overflow:hidden;-webkit-appearance:none;appearance:none;' +
      'box-sizing:border-box;opacity:0;pointer-events:none;white-space:pre;font:12px/1.2 sans-serif;border-radius:0';

    const live = document.createElement('div');
    live.className = 'bym-doc-live';
    live.setAttribute('aria-live', 'polite');
    live.setAttribute('aria-atomic', 'true');
    live.style.cssText = HIDDEN + ';white-space:nowrap';

    this.root = root;
    this.scroller = scroller;
    this.strip = strip;
    this.note = note;
    this.input = input;
    this.live = live;

    scroller.appendChild(strip);
    root.appendChild(input);
    root.appendChild(scroller);
    root.appendChild(note);
    root.appendChild(live);

    this.styleHost = acquireStyle(container);
    container.appendChild(root);

    try {
      this.readZoomOption(options.zoom);
      // The bar gets its words first: its height decides the height of the pages' area.
      this.updateBar();
      this.measure();
      const started = Date.now();
      let progress = this.engine.layout({ milliseconds: FIRST_SLICE_MS });
      while (!progress.done && this.pages.length === 0 && Date.now() - started < FIRST_PAGE_MS) progress = this.engine.layout({ milliseconds: FIRST_SLICE_MS });
      this.listen();
      this.afterLayout(progress);
      this.update();
    } catch (error) {
      // Nothing of a failed view stays behind.
      this.destroy();
      throw error;
    }
  }

  // ---------------------------------------------------------------- public

  /** Pages laid out so far; the final number once `layoutDone` is true. */
  get pageCount(): number {
    return this.pages.length;
  }

  /** Whether the whole document has been laid out. */
  get layoutDone(): boolean {
    return this.done;
  }

  /** The page in view, counted from 1: the one that shows most of itself. 0 when there is no page yet. */
  get currentPage(): number {
    return this.pages.length === 0 ? 0 : this.mostVisible() + 1;
  }

  /** The zoom in effect, 1 = 100%. */
  get zoom(): number {
    return this.scale;
  }

  get zoomMode(): 'auto' | 'fit-width' | 'fit-page' | 'manual' {
    return this.mode;
  }

  /** Sets the zoom: a factor (1 = 100%), or a mode that follows the size of the container. */
  setZoom(zoom: DocumentZoom): void {
    if (this.destroyed) return;
    if (zoom === 'fit-width' || zoom === 'fit-page') {
      this.mode = zoom;
      this.applyZoom(this.fitZoom());
    } else if (typeof zoom === 'number' && zoom > 0 && zoom < Infinity) {
      this.mode = 'manual';
      this.applyZoom(zoom);
    } else throw new RangeError('The zoom must be a positive number, "fit-width" or "fit-page"');
  }

  zoomIn(): void {
    for (let i = 0; i < ZOOM_STEPS.length; i++) {
      if (ZOOM_STEPS[i] > this.scale * 1.001) return this.setZoom(ZOOM_STEPS[i]);
    }
  }

  zoomOut(): void {
    for (let i = ZOOM_STEPS.length - 1; i >= 0; i--) {
      if (ZOOM_STEPS[i] < this.scale * 0.999) return this.setZoom(ZOOM_STEPS[i]);
    }
  }

  /**
   * Scrolls to a page, counted from 1. A page that is not laid out yet is
   * shown as soon as it is; a number beyond the document shows the last page.
   */
  goToPage(page: number): void {
    if (this.destroyed) return;
    if (typeof page !== 'number' || !(page >= 1)) throw new RangeError('The page number must be 1 or more');
    const index = Math.floor(page) - 1;
    if (index >= this.pages.length && !this.done) {
      this.pendingPage = index;
      this.scrollToPage(this.pages.length - 1);
      return;
    }
    this.pendingPage = undefined;
    this.scrollToPage(Math.min(index, this.pages.length - 1));
  }

  /** Selects the whole document; copying then yields its text. */
  selectAll(): void {
    if (this.destroyed) return;
    this.selection = 'all';
    this.repaintAll();
  }

  clearSelection(): void {
    if (this.destroyed || this.selection === undefined) return;
    this.selection = undefined;
    this.repaintAll();
  }

  /** The selected text: what a copy puts on the clipboard. Empty when nothing is selected. */
  getSelectedText(): string {
    const selection = this.selection;
    if (selection === undefined) return '';
    if (selection === 'all') return documentText(this.document);
    const range = orderedRange(selection.anchor, selection.focus);
    return range ? rangeTextOf(this.document, range) : '';
  }

  /** Changes the language of the view's own words. */
  setLocale(locale: string | object): void {
    if (this.destroyed) return;
    this.strings = resolveStrings(locale);
    this.root.setAttribute('aria-label', this.strings.document);
    this.input.setAttribute('aria-label', this.strings.document);
    this.slots.forEach((slot) => {
      slot.element.setAttribute('aria-label', this.strings.page(slot.index + 1));
      // Placeholders carry words.
      slot.dirty = true;
    });
    this.reportedPage = -1;
    this.update();
  }

  /** Moves the keyboard focus into the view. */
  focus(): void {
    if (!this.destroyed) this.focusInput(false);
  }

  /** Lays the document out again from the start: after fonts were added to the page, say. */
  relayout(): void {
    if (this.destroyed) return;
    const page = Math.max(0, this.mostVisible());
    if (this.layoutTimer !== undefined) clearTimeout(this.layoutTimer);
    this.layoutTimer = undefined;
    this.painter.clear();
    this.selection = undefined;
    this.unmountAll();
    this.engine = new DocLayout(this.document, this.measurer, this.options.layout);
    this.pages = this.engine.pages;
    this.offsets = [0];
    this.done = false;
    this.failed = false;
    this.pendingPage = page > 0 ? page : undefined;
    this.runLayout(FIRST_SLICE_MS);
    // At once, so that the pages are not gone for a frame.
    this.update();
  }

  /** What the view shows right now; for tests and for host pages that mirror it. */
  getViewState(): DocumentViewState {
    const mounted: number[] = [];
    this.slots.forEach((slot) => mounted.push(slot.index));
    mounted.sort((a, b) => a - b);
    return {
      zoom: this.scale,
      zoomMode: this.mode,
      page: this.currentPage,
      pageCount: this.pages.length,
      layoutDone: this.done,
      mounted,
      scrollX: this.scroller.scrollLeft,
      scrollY: this.virtualY(),
      width: this.width,
      height: this.height,
      devicePixelRatio: this.dpr,
      imageBytes: this.images.bytes,
      imageBudget: this.images.budget,
    };
  }

  /** Removes the view from its container with every listener, timer, canvas and decoded picture. Safe to call twice. */
  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    if (this.frame) cancelAnimationFrame(this.frame);
    this.frame = 0;
    if (this.layoutTimer !== undefined) clearTimeout(this.layoutTimer);
    if (this.fontTimer !== undefined) clearTimeout(this.fontTimer);
    this.layoutTimer = this.fontTimer = undefined;
    this.endDrag();
    this.unlisten();
    this.unmountAll();
    this.images.destroy();
    if (this.root.parentNode) this.root.parentNode.removeChild(this.root);
    releaseStyle(this.styleHost);
  }

  // ---------------------------------------------------------------- set-up

  private button(bar: HTMLElement, className: string, content?: SVGElement): HTMLButtonElement {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'bym-doc-btn ' + className;
    if (content) button.appendChild(content);
    bar.appendChild(button);
    return button;
  }

  private listen(): void {
    const { root, scroller, strip, input } = this;
    scroller.addEventListener('scroll', this.onScroll, { passive: true });
    scroller.addEventListener('focus', this.onScrollerFocus);
    scroller.addEventListener('wheel', this.onWheel, { passive: false });
    scroller.addEventListener('touchstart', this.onTouchStart, { passive: false });
    scroller.addEventListener('touchmove', this.onTouchMove, { passive: false });
    scroller.addEventListener('touchend', this.onTouchEnd);
    scroller.addEventListener('touchcancel', this.onTouchEnd);
    if (!('ontouchstart' in window)) {
      // The trackpad pinch of desktop Safari.
      scroller.addEventListener('gesturestart', this.onGestureStart);
      scroller.addEventListener('gesturechange', this.onGestureChange);
    }
    root.addEventListener('keydown', this.onKeyDown);
    root.addEventListener('copy', this.onCopy);
    root.addEventListener('mousedown', this.onRootMouseDown);
    input.addEventListener('focus', this.onFocus);
    input.addEventListener('blur', this.onBlur);
    strip.addEventListener('click', this.onStripClick);
    strip.addEventListener('dblclick', this.onDoubleClick);
    strip.addEventListener(this.pointerEvents ? 'pointerdown' : 'mousedown', this.onPointerDown as EventListener);
    window.addEventListener('resize', this.onResize);
    if (typeof ResizeObserver === 'function') {
      this.resizeObserver = new ResizeObserver(this.onResize);
      this.resizeObserver.observe(scroller);
    }
    this.watchPixelRatio();
    const fonts = (document as { fonts?: { addEventListener?(type: string, listener: () => void): void } }).fonts;
    if (fonts && typeof fonts.addEventListener === 'function') fonts.addEventListener('loadingdone', this.onFontsLoaded);
  }

  private unlisten(): void {
    const { root, scroller, strip, input } = this;
    scroller.removeEventListener('scroll', this.onScroll);
    scroller.removeEventListener('focus', this.onScrollerFocus);
    scroller.removeEventListener('wheel', this.onWheel);
    scroller.removeEventListener('touchstart', this.onTouchStart);
    scroller.removeEventListener('touchmove', this.onTouchMove);
    scroller.removeEventListener('touchend', this.onTouchEnd);
    scroller.removeEventListener('touchcancel', this.onTouchEnd);
    scroller.removeEventListener('gesturestart', this.onGestureStart);
    scroller.removeEventListener('gesturechange', this.onGestureChange);
    root.removeEventListener('keydown', this.onKeyDown);
    root.removeEventListener('copy', this.onCopy);
    root.removeEventListener('mousedown', this.onRootMouseDown);
    input.removeEventListener('focus', this.onFocus);
    input.removeEventListener('blur', this.onBlur);
    strip.removeEventListener('click', this.onStripClick);
    strip.removeEventListener('dblclick', this.onDoubleClick);
    strip.removeEventListener('pointerdown', this.onPointerDown as EventListener);
    strip.removeEventListener('mousedown', this.onPointerDown as EventListener);
    if (this.bar) this.bar.removeEventListener('click', this.onBarClick);
    window.removeEventListener('resize', this.onResize);
    if (this.resizeObserver) this.resizeObserver.disconnect();
    this.resizeObserver = undefined;
    this.unwatchPixelRatio();
    const fonts = (document as { fonts?: { removeEventListener?(type: string, listener: () => void): void } }).fonts;
    if (fonts && typeof fonts.removeEventListener === 'function') fonts.removeEventListener('loadingdone', this.onFontsLoaded);
  }

  private readZoomOption(zoom: DocumentZoom | undefined): void {
    if (zoom === 'fit-width' || zoom === 'fit-page') this.mode = zoom;
    else if (typeof zoom === 'number' && zoom > 0 && zoom < Infinity) {
      this.mode = 'manual';
      this.scale = clamp(zoom, MIN_ZOOM, MAX_ZOOM);
    } else if (zoom !== undefined) throw new RangeError('The zoom must be a positive number, "fit-width" or "fit-page"');
  }

  // ---------------------------------------------------------------- layout

  private runLayout = (milliseconds: number = SLICE_MS): void => {
    this.layoutTimer = undefined;
    if (this.destroyed || this.done || this.failed) return;
    let progress: LayoutProgress;
    try {
      progress = this.engine.layout({ milliseconds });
    } catch {
      // The pages laid out so far stay; the rest of the document is not shown.
      this.failed = true;
      this.notify(this.strings.documentLayoutFailed, 'error');
      this.update();
      return;
    }
    this.afterLayout(progress);
    this.schedule();
  };

  /** Takes over what a slice of layout produced and asks for the next one. */
  private afterLayout(progress: LayoutProgress): void {
    const pages = this.pages;
    const before = this.offsets.length - 1;
    for (let i = before; i < pages.length; i++) this.offsets.push(this.offsets[i] + this.pageHeight(i));
    for (let i = 0; i < progress.updated.length; i++) {
      const slot = this.slots.get(progress.updated[i]);
      if (slot) {
        slot.dirty = true;
        this.describe(slot);
      }
    }
    this.done = progress.done;
    if (this.engine.truncated && !this.truncationSaid) {
      this.truncationSaid = true;
      this.notify(this.strings.documentTooLong, 'info');
    }
    // The first page decides the zoom of the modes that fit.
    if (before === 0 && pages.length > 0 && this.mode !== 'manual') this.scale = this.fitZoom();
    this.sizeStrip();
    if (this.pendingPage !== undefined && (this.pendingPage < pages.length || this.done)) {
      const target = Math.min(this.pendingPage, pages.length - 1);
      this.pendingPage = undefined;
      this.scrollToPage(target);
    }
    if (!this.done) this.layoutTimer = setTimeout(this.runLayout, 0);
  }

  private onFontsLoaded = (): void => {
    if (this.destroyed || this.fontTimer !== undefined || this.fontRelayouts >= MAX_FONT_RELAYOUTS) return;
    // Fonts tend to arrive in groups.
    this.fontTimer = setTimeout(() => {
      this.fontTimer = undefined;
      this.fontRelayouts++;
      this.relayout();
    }, 50);
  };

  // -------------------------------------------------------------- geometry

  private pageWidth(index: number): number {
    const width = this.pages[index] ? this.pages[index].width : DEFAULT_PAGE.width;
    return width >= 1 && width < Infinity ? Math.min(width, MAX_PAGE_SIDE) : DEFAULT_PAGE.width;
  }

  private pageHeight(index: number): number {
    const height = this.pages[index] ? this.pages[index].height : DEFAULT_PAGE.height;
    return height >= 1 && height < Infinity ? Math.min(height, MAX_PAGE_SIDE) : DEFAULT_PAGE.height;
  }

  /** Top of a page in the strip as it would be without a limit on its height, screen px. */
  private pageTop(index: number): number {
    return PAD + this.offsets[index] * this.scale + index * GAP;
  }

  private totalHeight(): number {
    const count = this.pages.length;
    return PAD * 2 + this.offsets[count] * this.scale + Math.max(0, count - 1) * GAP;
  }

  private extent(): number {
    return Math.min(this.totalHeight(), MAX_EXTENT);
  }

  /** Scroll offset in the strip without limit; the same as scrollTop unless the document is taller than an element can be. */
  private virtualY(): number {
    const top = this.scroller.scrollTop;
    const total = this.totalHeight();
    if (total <= MAX_EXTENT) return top;
    const range = MAX_EXTENT - this.height;
    return range > 0 ? (top / range) * (total - this.height) : 0;
  }

  private setVirtualY(y: number): void {
    const total = this.totalHeight();
    const max = Math.max(0, total - this.height);
    const target = clamp(y, 0, max);
    if (total <= MAX_EXTENT) this.scroller.scrollTop = target;
    else this.scroller.scrollTop = max > 0 ? (target / max) * (MAX_EXTENT - this.height) : 0;
  }

  /** The page whose top is the last at or above `y` (screen px in the unlimited strip). */
  private pageAt(y: number): number {
    let low = 0;
    let high = this.pages.length - 1;
    while (low < high) {
      const mid = (low + high + 1) >> 1;
      if (this.pageTop(mid) <= y) low = mid;
      else high = mid - 1;
    }
    return Math.max(0, low);
  }

  private mostVisible(): number {
    const count = this.pages.length;
    if (count === 0) return -1;
    const top = this.virtualY();
    const bottom = top + Math.max(this.height, 1);
    let best = this.pageAt(top);
    let bestShare = -1;
    for (let i = best; i < count && this.pageTop(i) < bottom; i++) {
      const pageTop = this.pageTop(i);
      const shown = Math.min(bottom, pageTop + this.pageHeight(i) * this.scale) - Math.max(top, pageTop);
      if (shown > bestShare + 0.5) {
        bestShare = shown;
        best = i;
      }
    }
    return best;
  }

  private stripWidth(): number {
    let widest = 0;
    // The widest of the pages near the view would make the strip jump while scrolling: the first page and the mounted ones decide.
    if (this.pages.length > 0) widest = this.pageWidth(0);
    this.slots.forEach((slot) => {
      widest = Math.max(widest, this.pageWidth(slot.index));
    });
    return Math.max(this.width, widest * this.scale + PAD * 2);
  }

  private sizeStrip(): void {
    this.strip.style.height = `${Math.max(1, this.extent())}px`;
    this.strip.style.width = `${Math.max(1, this.stripWidth())}px`;
  }

  private fitZoom(): number {
    if (this.mode === 'manual') return this.scale;
    const width = this.pageWidth(0);
    const byWidth = (this.width - PAD * 2) / width;
    let zoom = byWidth;
    if (this.mode === 'auto') zoom = Math.min(1, byWidth);
    else if (this.mode === 'fit-page') zoom = Math.min(byWidth, (this.height - PAD * 2) / this.pageHeight(0));
    // A container without a size yet: the zoom is fitted when it gets one.
    return zoom > 0 ? clamp(zoom, MIN_ZOOM, MAX_ZOOM) : 1;
  }

  /** Reads the size of the visible area and the pixel ratio; true when one of them changed. */
  private measure(): boolean {
    const width = this.scroller.clientWidth;
    const height = this.scroller.clientHeight;
    const dpr = window.devicePixelRatio > 0 ? window.devicePixelRatio : 1;
    if (width === this.width && height === this.height && dpr === this.dpr) return false;
    this.width = width;
    this.height = height;
    this.dpr = dpr;
    return true;
  }

  /**
   * Changes the zoom and keeps the point of the document that is at (ax, ay)
   * of the visible area where it is; by default the top of the visible area.
   */
  private applyZoom(zoom: number, ax?: number, ay?: number): void {
    const next = clamp(zoom, MIN_ZOOM, MAX_ZOOM);
    if (next === this.scale) {
      this.updateBar();
      return;
    }
    const anchorX = ax === undefined ? this.width / 2 : ax;
    const anchorY = ay === undefined ? 0 : ay;
    const y = this.virtualY() + anchorY;
    const count = this.pages.length;
    let page = 0;
    let share = 0;
    // Distance in screen px from the page, for a point in the space around it: that space does not grow with the zoom.
    let outside = 0;
    if (count > 0) {
      page = this.pageAt(y);
      const height = this.pageHeight(page) * this.scale;
      const inPage = y - this.pageTop(page);
      share = clamp(inPage / height, 0, 1);
      outside = inPage < 0 ? inPage : inPage > height ? inPage - height : 0;
    }
    const oldWidth = this.stripWidth();
    const across = oldWidth > 0 ? (this.scroller.scrollLeft + anchorX) / oldWidth : 0.5;
    this.scale = next;
    this.sizeStrip();
    if (count > 0) this.setVirtualY(this.pageTop(page) + share * this.pageHeight(page) * next + outside - anchorY);
    this.scroller.scrollLeft = Math.max(0, across * this.stripWidth() - anchorX);
    this.slots.forEach((slot) => {
      slot.dirty = true;
    });
    this.update();
  }

  private scrollToPage(index: number): void {
    if (index < 0 || this.pages.length === 0) return;
    this.sizeStrip();
    // The gap above the page stays in sight, so that the page is seen to begin there.
    this.setVirtualY(this.pageTop(index) - (index === 0 ? PAD : GAP / 2));
    this.update();
  }

  // ---------------------------------------------------------------- update

  /** Brings the screen up to date in the next frame; any number of changes before it cost one update. */
  private schedule(): void {
    if (this.destroyed || this.frame) return;
    this.frame = requestAnimationFrame(() => {
      this.frame = 0;
      this.update();
    });
  }

  private update(): void {
    if (this.destroyed) return;
    if (this.frame) {
      cancelAnimationFrame(this.frame);
      this.frame = 0;
    }
    const count = this.pages.length;
    this.images.pass++;
    this.note.style.display = count === 0 && !this.done && !this.failed ? 'block' : 'none';
    if (count === 0) this.note.textContent = this.strings.preparingPages;

    let first = 0;
    let last = -1;
    if (count > 0 && this.height > 0) {
      const top = this.virtualY();
      first = Math.max(0, this.pageAt(top) - BUFFER_PAGES);
      last = Math.min(count - 1, this.pageAt(top + this.height) + BUFFER_PAGES);
      if (last - first + 1 > MAX_MOUNTED) last = first + MAX_MOUNTED - 1;
    }
    const gone: Slot[] = [];
    this.slots.forEach((slot) => {
      if (slot.index < first || slot.index > last) gone.push(slot);
    });
    for (let i = 0; i < gone.length; i++) this.unmount(gone[i]);
    for (let i = first; i <= last; i++) if (!this.slots.has(i)) this.mount(i);

    this.sizeStrip();
    const stripWidth = this.stripWidth();
    const shift = this.scroller.scrollTop - this.virtualY();
    this.slots.forEach((slot) => {
      const width = this.pageWidth(slot.index) * this.scale;
      const height = this.pageHeight(slot.index) * this.scale;
      const left = Math.max(PAD, (stripWidth - width) / 2);
      const top = this.pageTop(slot.index) + shift;
      const style = slot.element.style;
      if (slot.left !== left) style.left = `${(slot.left = left)}px`;
      if (slot.top !== top) style.top = `${(slot.top = top)}px`;
      if (slot.width !== width) style.width = `${(slot.width = width)}px`;
      if (slot.height !== height) style.height = `${(slot.height = height)}px`;
      if (slot.dirty) this.paint(slot);
    });

    this.updateBar();
    const page = this.currentPage;
    if (page !== this.reportedPage || count !== this.reportedCount) {
      const turned = page !== this.reportedPage;
      this.reportedPage = page;
      this.reportedCount = count;
      // The live region follows the page in view, not the count that grows while the layout runs.
      if (turned && page > 0) this.live.textContent = this.strings.pageOf(page, count);
      if (this.options.onPageChange) this.options.onPageChange(page, count);
    }
  }

  private updateBar(): void {
    const strings = this.strings;
    if (this.zoomOutButton && this.zoomInButton && this.zoomButton && this.fitButton) {
      const percent = Math.round(this.scale * 100);
      const label = (button: HTMLButtonElement, text: string) => {
        if (button.getAttribute('aria-label') !== text) {
          button.setAttribute('aria-label', text);
          button.title = text;
        }
      };
      label(this.zoomOutButton, strings.zoomOut);
      label(this.zoomInButton, strings.zoomIn);
      label(this.fitButton, strings.fitWidth);
      label(this.zoomButton, strings.zoomLevel(percent));
      const text = `${percent}%`;
      if (this.zoomButton.textContent !== text) this.zoomButton.textContent = text;
      this.zoomOutButton.disabled = this.scale <= MIN_ZOOM;
      this.zoomInButton.disabled = this.scale >= MAX_ZOOM;
      this.fitButton.setAttribute('aria-pressed', this.mode === 'fit-width' ? 'true' : 'false');
      if (this.bar) this.bar.setAttribute('aria-label', strings.documentTools);
    }
    if (this.pagesLabel) {
      const count = this.pages.length;
      // "+": more pages are on their way.
      const text = count === 0 ? (this.done || this.failed ? '' : strings.preparingPages) : strings.pageOf(this.currentPage, count) + (this.done || this.failed ? '' : '+');
      if (this.pagesLabel.textContent !== text) this.pagesLabel.textContent = text;
    }
  }

  // ----------------------------------------------------------------- pages

  private mount(index: number): void {
    const element = document.createElement('div');
    element.className = 'bym-doc-page';
    element.setAttribute('role', 'region');
    element.setAttribute('aria-label', this.strings.page(index + 1));
    element.setAttribute('data-page', String(index + 1));
    element.style.cssText = 'position:absolute;box-sizing:content-box;margin:0;padding:0;border:0;overflow:hidden';

    const canvas = document.createElement('canvas');
    canvas.className = 'bym-doc-canvas';
    canvas.setAttribute('aria-hidden', 'true');
    canvas.style.cssText =
      'display:block;position:absolute;left:0;top:0;width:100%;height:100%;float:none;box-sizing:content-box;margin:0;padding:0;border:0;' +
      'min-width:0;min-height:0;max-width:none;max-height:none';

    const text = document.createElement('div');
    text.className = 'bym-doc-text';
    // For assistive technology: the page as text. Never markup.
    text.style.cssText = HIDDEN + ';white-space:pre-wrap';

    element.appendChild(canvas);
    element.appendChild(text);
    const slot: Slot = { index, element, canvas, context: null, text, links: [], dirty: true, left: NaN, top: NaN, width: NaN, height: NaN };
    this.describe(slot);
    this.slots.set(index, slot);
    // In page order, whatever order the pages came into view in: the order of the Tab key and of a screen reader.
    let next: Element | null = null;
    let nextIndex = Infinity;
    this.slots.forEach((other) => {
      if (other.index > index && other.index < nextIndex) {
        nextIndex = other.index;
        next = other.element;
      }
    });
    this.strip.insertBefore(element, next);
  }

  private unmount(slot: Slot): void {
    this.slots.delete(slot.index);
    if (slot.element.parentNode) slot.element.parentNode.removeChild(slot.element);
    // Give the backing store back right away; some browsers keep it until collection otherwise.
    slot.canvas.width = 0;
    slot.canvas.height = 0;
    slot.context = null;
    slot.links.length = 0;
  }

  private unmountAll(): void {
    const all: Slot[] = [];
    this.slots.forEach((slot) => all.push(slot));
    for (let i = 0; i < all.length; i++) this.unmount(all[i]);
  }

  /** Fills in what a page offers besides its picture: its text for assistive technology and its hyperlinks. */
  private describe(slot: Slot): void {
    const page = this.pages[slot.index];
    for (let i = 0; i < slot.links.length; i++) {
      const link = slot.links[i];
      if (link.parentNode) link.parentNode.removeChild(link);
    }
    slot.links.length = 0;
    if (!page) return;
    const width = this.pageWidth(slot.index);
    const height = this.pageHeight(slot.index);
    const lines: string[] = [];
    let line = '';
    let lineY = NaN;
    let previous: TextItem | undefined;
    let previousLink: { element: HTMLElement; item: TextItem; right: number } | undefined;
    const items = page.items;
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      if (item.type !== 'text') continue;
      if (previous && Math.abs(item.y - lineY) > 0.5 && !(previous.source && sameSource(previous, item) && line.length > 0 && /\s$/.test(line))) {
        // A new line on the page is a new line of text, except inside a paragraph that just wrapped.
        lines.push(line);
        line = '';
      }
      line += item.text;
      lineY = item.y;
      previous = item;

      // A link in vertical text is painted but has no element: its box would be in the wrong place.
      if (item.link === undefined || slot.links.length >= MAX_LINKS_PER_PAGE || !(item.width > 0) || item.rotation !== undefined) {
        if (item.link === undefined) previousLink = undefined;
        continue;
      }
      const target = safeHyperlinkTarget(item.link);
      if (target === undefined) continue;
      const box = this.measurer.metrics(item.font);
      if (previousLink && previousLink.item.link === item.link && Math.abs(previousLink.item.y - item.y) <= 0.5 && Math.abs(previousLink.right - item.x) < 1) {
        // The next piece of the same link on the same line: one element for both.
        const element = previousLink.element;
        const left = parseFloat(element.style.left) / 100;
        element.style.width = `${((item.x + item.width) / width - left) * 100}%`;
        element.setAttribute('aria-label', (element.getAttribute('aria-label') || '') + item.text);
        previousLink.right = item.x + item.width;
        previousLink.item = item;
        continue;
      }
      const element = document.createElement('span');
      element.className = 'bym-doc-link';
      element.setAttribute('role', 'link');
      element.setAttribute('aria-label', item.text);
      // A link that goes on from the line before is one stop for the Tab key, not two.
      element.tabIndex = previousLink && previousLink.item.link === item.link ? -1 : 0;
      // In shares of the page: the element keeps its place at every zoom.
      element.style.cssText =
        'position:absolute;display:block;box-sizing:border-box;margin:0;padding:0;border:0;background:transparent;' +
        `left:${(item.x / width) * 100}%;top:${((item.y - box.ascent) / height) * 100}%;` +
        `width:${(item.width / width) * 100}%;height:${((box.ascent + box.descent) / height) * 100}%`;
      this.linkTargets.set(element, target);
      slot.element.appendChild(element);
      slot.links.push(element);
      previousLink = { element, item, right: item.x + item.width };
    }
    if (line.length > 0) lines.push(line);
    slot.text.textContent = lines.join('\n');
  }

  private paint(slot: Slot): void {
    const page = this.pages[slot.index];
    if (!page) return;
    slot.dirty = false;
    const pageWidth = this.pageWidth(slot.index);
    const pageHeight = this.pageHeight(slot.index);
    const cssWidth = pageWidth * this.scale;
    const cssHeight = pageHeight * this.scale;
    let ratio = this.dpr;
    if (cssWidth * cssHeight * ratio * ratio > this.maxCanvasPixels) ratio = Math.sqrt(this.maxCanvasPixels / (cssWidth * cssHeight));
    const width = Math.max(1, Math.round(cssWidth * ratio));
    const height = Math.max(1, Math.round(cssHeight * ratio));
    const canvas = slot.canvas;
    if (canvas.width !== width || canvas.height !== height) {
      canvas.width = width;
      canvas.height = height;
    }
    const context = slot.context || (slot.context = canvas.getContext('2d'));
    if (!context) return;
    try {
      context.setTransform(1, 0, 0, 1, 0, 0);
      context.fillStyle = '#ffffff';
      context.fillRect(0, 0, width, height);
      const scale = width / pageWidth;
      renderPage(context, page, {
        scale,
        images: this.images,
        cssFont: this.painter.cssFont,
        metrics: this.metricsOf,
        placeholderLabel: this.placeholderLabel,
      });
      this.paintSelection(context, slot.index, scale);
    } catch {
      // One page that cannot be painted must not take the view down; it stays blank.
      if (!this.paintFailed) this.notify(this.strings.actionFailed, 'error');
      this.paintFailed = true;
    }
  }

  private metricsOf = (font: LayoutFont): FontMetrics => this.measurer.metrics(font);

  private placeholderLabel = (item: ImageItem): string | undefined => {
    if (item.alt !== undefined && item.alt !== '') return item.alt;
    return this.strings.placeholder(item.placeholder === undefined ? 'picture' : item.placeholder);
  };

  private onImageLoad = (part: string): void => {
    if (this.destroyed) return;
    this.slots.forEach((slot) => {
      if (slot.dirty) return;
      const page = this.pages[slot.index];
      const items = page ? page.items : [];
      for (let i = 0; i < items.length; i++) {
        const item = items[i];
        if (item.type === 'image' && item.part === part) {
          slot.dirty = true;
          break;
        }
      }
    });
    this.schedule();
  };

  private repaintAll(): void {
    this.slots.forEach((slot) => {
      slot.dirty = true;
    });
    this.schedule();
  }

  // ------------------------------------------------------------- selection

  /** Advance from the left edge of a text item to its character `offset`. */
  private advance(item: TextItem, offset: number): number {
    if (offset <= 0) return 0;
    if (offset >= item.text.length) return item.width;
    const width = this.measurer.width(item.font, item.text.slice(0, offset)) * (item.scale || 1) + (item.letterSpacing || 0) * offset;
    return clamp(width, 0, item.width);
  }

  /**
   * The text item nearest to a point of a page (px of the page) and the
   * character boundary in it nearest to the point. Only text of the document
   * counts: not what the engine made up, not repeated header rows. With
   * `flow`, only text of the flow that place is in.
   */
  private hitAt(pageIndex: number, x: number, y: number, flow?: TextPosition): { item: TextItem; offset: number } | undefined {
    const page = this.pages[pageIndex];
    if (!page) return undefined;
    const items = page.items;
    let best: TextItem | undefined;
    let bestDistance = Infinity;
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      // Vertical text is painted and copied, but a point is not looked up in it.
      if (item.type !== 'text' || item.text.length === 0 || !item.source || item.repeat || item.rotation !== undefined) continue;
      if (flow && !sameFlow(item.source, flow)) continue;
      const box = this.measurer.metrics(item.font);
      const dy = y < item.y - box.ascent ? item.y - box.ascent - y : y > item.y + box.descent ? y - item.y - box.descent : 0;
      const dx = x < item.x ? item.x - x : x > item.x + item.width ? x - item.x - item.width : 0;
      // The line counts for more than the place in it.
      const distance = dy * 1000 + dx;
      if (distance < bestDistance) {
        bestDistance = distance;
        best = item;
      }
    }
    if (!best) return undefined;
    const item = best;
    const target = x - item.x;
    let low = 0;
    let high = item.text.length;
    // The last boundary at or before the point, then the nearer of it and the next.
    while (low < high) {
      const mid = (low + high + 1) >> 1;
      if (this.advance(item, mid) <= target) low = mid;
      else high = mid - 1;
    }
    if (low < item.text.length && target - this.advance(item, low) > this.advance(item, low + 1) - target) low++;
    return { item, offset: low };
  }

  private static positionOf(item: TextItem, offset: number): TextPosition {
    const source = item.source as NonNullable<TextItem['source']>;
    return { story: source.story, path: source.path, offset: source.offset + offset };
  }

  /** The page under a point of the window and the point in px of that page. */
  private pagePoint(clientX: number, clientY: number): { page: number; x: number; y: number } | undefined {
    let found: { page: number; x: number; y: number } | undefined;
    let nearest = Infinity;
    this.slots.forEach((slot) => {
      // The rectangle on screen, whatever transforms lie on the view.
      const rect = slot.element.getBoundingClientRect();
      if (!(rect.width > 0) || !(rect.height > 0)) return;
      const distance = clientY < rect.top ? rect.top - clientY : clientY > rect.bottom ? clientY - rect.bottom : 0;
      if (distance >= nearest) return;
      nearest = distance;
      found = {
        page: slot.index,
        x: ((clientX - rect.left) / rect.width) * this.pageWidth(slot.index),
        y: ((clientY - rect.top) / rect.height) * this.pageHeight(slot.index),
      };
    });
    return found;
  }

  private hitFromEvent(event: { clientX: number; clientY: number }, flow?: TextPosition): { item: TextItem; offset: number } | undefined {
    const point = this.pagePoint(event.clientX, event.clientY);
    return point ? this.hitAt(point.page, point.x, point.y, flow) : undefined;
  }

  private positionFromEvent(event: { clientX: number; clientY: number }, flow?: TextPosition): TextPosition | undefined {
    const hit = this.hitFromEvent(event, flow);
    return hit ? DocumentView.positionOf(hit.item, hit.offset) : undefined;
  }

  /** Paints the selection over a page: only over the text that is selected, wherever on the page it is painted. */
  private paintSelection(context: CanvasRenderingContext2D, pageIndex: number, scale: number): void {
    const selection = this.selection;
    if (selection === undefined) return;
    const range = selection === 'all' ? undefined : orderedRange(selection.anchor, selection.focus);
    if (selection !== 'all' && !range) return;
    const items = this.pages[pageIndex].items;
    context.save();
    context.scale(scale, scale);
    context.fillStyle = SELECTION_FILL;
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      if (item.type !== 'text' || !(item.width > 0) || item.rotation !== undefined) continue;
      let a = 0;
      let b = item.width;
      if (range) {
        const part = selectedPart(range, item);
        if (!part) continue;
        a = this.advance(item, part[0]);
        b = this.advance(item, part[1]);
      }
      if (!(b > a)) continue;
      const box = this.measurer.metrics(item.font);
      context.fillRect(item.x + a, item.y - box.ascent, b - a, box.ascent + box.descent);
    }
    context.restore();
  }

  private setFocusPosition(position: TextPosition): void {
    const selection = this.selection;
    if (!selection || selection === 'all') return;
    if (comparePositions(selection.focus, position) === 0) return;
    selection.focus = position;
    // The selected text may be on any page that is mounted: there are only a few.
    this.repaintAll();
  }

  private onPointerDown = (event: PointerEvent | MouseEvent): void => {
    if (this.destroyed || event.button !== 0) return;
    this.dragMoved = false;
    // A finger scrolls; only a mouse or a pen selects.
    if ('pointerType' in event && event.pointerType === 'touch') return;
    const selection = this.selection;
    const extend = event.shiftKey && selection !== undefined && selection !== 'all';
    // A selection stays in the flow of text it began in: the body, one header, one note, one text box.
    const position = this.positionFromEvent(event, extend ? (selection as { anchor: TextPosition }).anchor : undefined);
    if (!position) {
      if (!extend) this.clearSelection();
      return;
    }
    if (extend) {
      this.setFocusPosition(position);
    } else {
      const had = selection !== undefined;
      this.selection = { anchor: position, focus: position };
      if (had) this.repaintAll();
    }
    this.dragging = true;
    window.addEventListener(this.pointerEvents ? 'pointermove' : 'mousemove', this.onPointerMove as EventListener);
    window.addEventListener(this.pointerEvents ? 'pointerup' : 'mouseup', this.onPointerUp);
    if (this.pointerEvents) window.addEventListener('pointercancel', this.onPointerUp);
  };

  private onPointerMove = (event: PointerEvent | MouseEvent): void => {
    if (!this.dragging || this.destroyed) return;
    const selection = this.selection;
    if (!selection || selection === 'all') return;
    const position = this.positionFromEvent(event, selection.anchor);
    if (!position) return;
    if (comparePositions(selection.anchor, position) !== 0) this.dragMoved = true;
    this.setFocusPosition(position);
  };

  private onPointerUp = (): void => {
    this.endDrag();
  };

  private endDrag(): void {
    if (!this.dragging) return;
    this.dragging = false;
    window.removeEventListener('pointermove', this.onPointerMove as EventListener);
    window.removeEventListener('mousemove', this.onPointerMove as EventListener);
    window.removeEventListener('pointerup', this.onPointerUp);
    window.removeEventListener('mouseup', this.onPointerUp);
    window.removeEventListener('pointercancel', this.onPointerUp);
  }

  private onDoubleClick = (event: MouseEvent): void => {
    if (this.destroyed) return;
    const hit = this.hitFromEvent(event);
    if (!hit) return;
    const text = hit.item.text;
    let start = Math.min(hit.offset, text.length - 1);
    if (start < 0 || /\s/.test(text.charAt(start))) return;
    let end = start;
    while (start > 0 && !/\s/.test(text.charAt(start - 1))) start--;
    while (end < text.length && !/\s/.test(text.charAt(end))) end++;
    this.selection = { anchor: DocumentView.positionOf(hit.item, start), focus: DocumentView.positionOf(hit.item, end) };
    this.repaintAll();
  };

  private onCopy = (event: ClipboardEvent): void => {
    const data = event.clipboardData;
    if (this.destroyed || !data) return;
    const text = this.getSelectedText();
    if (text === '') {
      // The blank the input holds for the sake of this event must not replace what is on the clipboard.
      if (event.target === this.input) event.preventDefault();
      return;
    }
    event.preventDefault();
    data.setData('text/plain', text);
  };

  // ----------------------------------------------------------------- links

  private linkOf(target: EventTarget | null): string | undefined {
    let node = target as Element | null;
    // The link element itself or something inside it; never beyond the strip.
    for (let depth = 0; node && node !== this.strip && depth < 8; depth++) {
      const link = this.linkTargets.get(node);
      if (link !== undefined) return link;
      node = node.parentElement;
    }
    return undefined;
  }

  private activateLink(target: string): void {
    const handler = this.options.onLinkClick;
    if (handler && handler(target) === false) return;
    // Checked again where it is used: nothing but http, https, mailto and tel is ever opened.
    const safe = safeHyperlinkTarget(target);
    if (safe === undefined) return;
    window.open(safe, '_blank', 'noopener,noreferrer');
  }

  private onStripClick = (event: MouseEvent): void => {
    if (this.destroyed || this.dragMoved) return;
    const target = this.linkOf(event.target);
    if (target === undefined) return;
    event.preventDefault();
    this.activateLink(target);
  };

  // -------------------------------------------------------------- keyboard

  private onKeyDown = (event: KeyboardEvent): void => {
    if (this.destroyed || event.defaultPrevented) return;
    const target = event.target as Element | null;
    // Keys on the buttons of the bar are theirs.
    if (this.bar && target && this.bar.contains(target)) return;
    const command = event.ctrlKey || event.metaKey;
    const key = event.key;
    if (!command && !event.altKey && (key === 'Enter' || key === ' ')) {
      const link = this.linkOf(target);
      if (link !== undefined) {
        event.preventDefault();
        this.activateLink(link);
        return;
      }
    }
    if (event.altKey) return;
    if (command) {
      if (key === 'a' || key === 'A') this.selectAll();
      else if (key === '+' || key === '=') this.zoomIn();
      else if (key === '-' || key === '_') this.zoomOut();
      else if (key === '0') this.setZoom(1);
      else if (key === 'Home') this.setVirtualY(0);
      else if (key === 'End') this.setVirtualY(this.totalHeight());
      else return;
      event.preventDefault();
      this.schedule();
      return;
    }
    const page = Math.max(LINE_SCROLL, this.height - LINE_SCROLL);
    const y = this.virtualY();
    switch (key) {
      case 'PageDown':
        this.setVirtualY(y + page);
        break;
      case 'PageUp':
        this.setVirtualY(y - page);
        break;
      case ' ':
        this.setVirtualY(y + (event.shiftKey ? -page : page));
        break;
      case 'ArrowDown':
        this.setVirtualY(y + LINE_SCROLL);
        break;
      case 'ArrowUp':
        this.setVirtualY(y - LINE_SCROLL);
        break;
      case 'ArrowRight':
        this.scroller.scrollLeft += LINE_SCROLL;
        break;
      case 'ArrowLeft':
        this.scroller.scrollLeft -= LINE_SCROLL;
        break;
      case 'Home':
        this.setVirtualY(0);
        break;
      case 'End':
        this.setVirtualY(this.totalHeight());
        break;
      case 'Escape':
        if (this.selection === undefined) return;
        this.clearSelection();
        break;
      default:
        // Tab among them: the focus moves on, out of the view.
        return;
    }
    event.preventDefault();
    this.update();
  };

  private focusInput(quiet: boolean): void {
    this.quietFocus = quiet;
    try {
      this.input.focus({ preventScroll: true });
    } catch {
      this.input.focus();
    }
    this.quietFocus = false;
  }

  private onFocus = (): void => {
    if (this.destroyed) return;
    if (!this.quietFocus) this.root.classList.add('bym-doc-focus');
    try {
      this.input.select();
    } catch {
      // Nothing depends on it but the copy event of some browsers.
    }
  };

  private onBlur = (): void => {
    this.root.classList.remove('bym-doc-focus');
  };

  private onScrollerFocus = (): void => {
    if (!this.destroyed) this.focusInput(true);
  };

  /** A press inside the view puts the focus in the input, so that the keys work; buttons and links keep their own behaviour. */
  private onRootMouseDown = (event: MouseEvent): void => {
    const target = event.target as Element | null;
    if (this.destroyed || !target || target === this.input) return;
    if (this.bar && this.bar.contains(target)) {
      if (typeof target.closest === 'function' && target.closest('button')) return;
      event.preventDefault();
      this.focusInput(true);
      return;
    }
    // On the scrolling element itself the press is on a scroll bar: left to the browser, the focus is handed on from there.
    if (target === this.scroller) return;
    event.preventDefault();
    this.focusInput(true);
  };

  private onBarClick = (event: MouseEvent): void => {
    if (this.destroyed) return;
    const target = event.target as Element | null;
    const button = target && typeof target.closest === 'function' ? target.closest('button') : null;
    if (!button) return;
    if (button === this.zoomOutButton) this.zoomOut();
    else if (button === this.zoomInButton) this.zoomIn();
    else if (button === this.zoomButton) this.setZoom(1);
    else if (button === this.fitButton) this.setZoom('fit-width');
    // After a press with a pointer the keys belong to the document again; a key press leaves the focus on the button.
    if (event.detail > 0) this.focusInput(true);
  };

  // ------------------------------------------------------ scroll and zoom

  private onScroll = (): void => {
    this.schedule();
  };

  private onResize = (): void => {
    const ratio = this.dpr;
    if (this.destroyed || !this.measure()) return;
    if (this.mode !== 'manual') {
      const zoom = this.fitZoom();
      if (zoom !== this.scale) {
        this.applyZoom(zoom);
        return;
      }
    }
    // At another pixel ratio the pages are drawn again.
    if (this.dpr !== ratio) {
      this.slots.forEach((slot) => {
        slot.dirty = true;
      });
    }
    this.update();
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
    // The query matched the old ratio; watch the new one.
    this.watchPixelRatio();
    this.onResize();
  };

  /** A point of the window in px of the visible area, whatever transforms lie on the view. */
  private viewportPoint(clientX: number, clientY: number): { x: number; y: number } {
    const rect = this.scroller.getBoundingClientRect();
    const sx = rect.width > 0 && this.scroller.offsetWidth > 0 ? this.scroller.offsetWidth / rect.width : 1;
    const sy = rect.height > 0 && this.scroller.offsetHeight > 0 ? this.scroller.offsetHeight / rect.height : 1;
    return { x: (clientX - rect.left) * sx, y: (clientY - rect.top) * sy };
  }

  private zoomAround(zoom: number, clientX: number, clientY: number): void {
    const point = this.viewportPoint(clientX, clientY);
    this.mode = 'manual';
    this.applyZoom(zoom, point.x, point.y);
  }

  private onWheel = (event: WheelEvent): void => {
    if (this.destroyed || !(event.ctrlKey || event.metaKey)) return;
    event.preventDefault();
    // Lines and pages count for more than pixels; a pinch on a trackpad arrives here too, in small steps.
    const delta = event.deltaY * (event.deltaMode === 1 ? 16 : event.deltaMode === 2 ? 100 : 1);
    if (!delta) return;
    this.zoomAround(this.scale * Math.exp(-clamp(delta, -200, 200) * 0.0025), event.clientX, event.clientY);
  };

  private onTouchStart = (event: TouchEvent): void => {
    if (this.destroyed || event.touches.length !== 2) {
      this.pinch = undefined;
      return;
    }
    const a = event.touches[0];
    const b = event.touches[1];
    const distance = Math.sqrt((a.clientX - b.clientX) * (a.clientX - b.clientX) + (a.clientY - b.clientY) * (a.clientY - b.clientY));
    if (distance > 0) this.pinch = { distance, zoom: this.scale };
  };

  private onTouchMove = (event: TouchEvent): void => {
    const pinch = this.pinch;
    if (this.destroyed || !pinch || event.touches.length !== 2) return;
    if (event.cancelable) event.preventDefault();
    const a = event.touches[0];
    const b = event.touches[1];
    const distance = Math.sqrt((a.clientX - b.clientX) * (a.clientX - b.clientX) + (a.clientY - b.clientY) * (a.clientY - b.clientY));
    if (!(distance > 0)) return;
    this.zoomAround(pinch.zoom * (distance / pinch.distance), (a.clientX + b.clientX) / 2, (a.clientY + b.clientY) / 2);
  };

  private onTouchEnd = (event: TouchEvent): void => {
    if (event.touches.length < 2) this.pinch = undefined;
  };

  private onGestureStart = (event: Event): void => {
    if (this.destroyed) return;
    event.preventDefault();
    this.gestureZoom = this.scale;
  };

  private onGestureChange = (event: Event): void => {
    const gesture = event as Event & { scale?: number; clientX?: number; clientY?: number };
    if (this.destroyed || !(this.gestureZoom > 0) || !(typeof gesture.scale === 'number' && gesture.scale > 0)) return;
    event.preventDefault();
    const rect = this.scroller.getBoundingClientRect();
    this.zoomAround(
      this.gestureZoom * gesture.scale,
      typeof gesture.clientX === 'number' ? gesture.clientX : rect.left + rect.width / 2,
      typeof gesture.clientY === 'number' ? gesture.clientY : rect.top + rect.height / 2,
    );
  };

  private notify(message: string, kind: 'error' | 'info'): void {
    this.live.textContent = message;
    if (this.options.onNotice) this.options.onNotice({ message, kind });
  }
}
