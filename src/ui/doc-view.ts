// The document view: shows a Word document as a strip of pages that scrolls
// vertically, read-only or, with `readOnly: false`, as an editor: a caret, a
// selection, typing (see doc-input.ts), the clipboard, undo and redo. Every
// change goes through the commands of ops/doc and its history; the layout
// follows an edit from the page it is on (DocLayout.edited). Pages are laid out in slices so that the first
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
import type { DocLayoutOptions, FontMetrics, ImageItem, LayoutFont, LayoutPage, LayoutProgress, LineMark, TextItem, TextMeasurer } from '../layout/doc/types';
import * as ops from '../ops/doc/index';
import type { DocEdit, DocPosition, DocRange, ParaFormatPatch, RangeFormat, RunFormatPatch, ToggleKey } from '../ops/doc/index';
import { enforcedProtection, findWrapper, inFieldCode, paragraphText, safeHyperlinkTarget, type Block, type DocModel, type Paragraph, type RunItem, type Table } from '../model/doc/types';
import { renderPage } from '../render/doc';
import { IMAGE_BUDGET_BYTES, IMAGE_BUDGET_BYTES_COARSE, ImageStore, declaredImageSize, sniffImageType } from '../render/images';
import { CanvasTextMeasurer } from '../render/doc-measure';
import { onFontsChanged } from '../render/fonts';
import { onEmbeddedFontNotice } from '../render/embedded-fonts';
import { DocLines, lineEnd, type Affinity, type CaretBox, type PageLine } from './doc-caret';
import { cutLoses, fragmentOfRange, setCopyMedia, sliceHtml, sliceOf, storeCopy, storedCopy, type DocSlice, type StoredCopy } from './doc-clipboard';
import { DocInput, type InputChange, type InputContext } from './doc-input';
import {
  DocumentPrintError,
  PrintFrame,
  canvasPicture,
  openPrintDialog,
  planPrint,
  printLimits,
  MAX_PRINT_PAGES,
  revokePicture,
  supportsNamedPages,
  type DocumentPrintOptions,
  type DocumentPrintResult,
  type PrintDialog,
  type PrintLimits,
  type PrintPlan,
} from './doc-print';
import { resolveUiStrings, type UiStrings } from './strings';

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
  /**
   * False makes the view an editor: a caret, typing, the clipboard, undo and
   * redo. Default true: the document is shown read-only, with a selection
   * that can be copied.
   */
  readOnly?: boolean;
  /** Called after every change of the document made in the view, undo and redo included. */
  onChange?(): void;
  /** Called when the caret or the selection moves; undefined when nothing is selected. */
  onSelectionChange?(selection: DocRange | undefined): void;
  /** Called for Ctrl+K (Cmd+K): the shell around the view asks for the address of a link. */
  onLinkShortcut?(): void;
  /**
   * Whether the package the document was read from has a part of this name
   * ("/word/media/image3.png"): a picture put in gets a name of its own.
   * Without it, `package.hasPart` is asked when the package has one.
   */
  partTaken?(part: string): boolean;
  /**
   * Tracks the changes made in the view although the document does not ask
   * for it. A document that asks for it (its "track changes" setting, or a
   * protection that allows only tracked changes) is always edited with
   * tracked changes; false does not turn that off.
   */
  trackChanges?: boolean;
  /** Who is editing: the author tracked changes name. Default "Author". */
  author?: string;
  /** The clock tracked changes take their date from (tests). Default: the system's. */
  now?(): Date;
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
/** Text of the paragraph the field for typing mirrors before and after the caret (see doc-input.ts). */
const CONTEXT_BEFORE = 80;
const CONTEXT_AFTER = 40;
/** Longest text taken from the clipboard, in characters, and largest picture, in bytes. */
const MAX_PASTE_TEXT = 5000000;
const MAX_PICTURE_BYTES = 20000000;
/** How long an edit may lay out pages before the screen is brought up to date; the rest follows in slices. */
const EDIT_LAYOUT_MS = 120;
/** Two presses this close in time and place are a double click. */
const CLICK_MS = 500;
const CLICK_DISTANCE = 5;
const LONG_PRESS_MS = 500;
/** A drag this close to the top or bottom of the pages scrolls them. */
const DRAG_SCROLL_EDGE = 24;
const CARET_MARGIN = 8;
const EMU_PER_PX = 9525;
/** A selection over more blocks of the body than this is too large to look at as a whole for the toolbar or to read out. */
const MAX_LOOKED_AT_BLOCKS = 400;
/** How often the pages are prepared for printing anew because the document changed meanwhile. */
const MAX_PRINT_RESTARTS = 3;
/** How long one page waits for its pictures before it is printed with placeholders. */
const PRINT_PICTURES_MS = 30000;

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
  '.bym-doc-note{color:#5f6368;font-size:14px}' +
  // The caret: two animations with the same frames, so that changing the name starts the blinking anew.
  '@keyframes bym-doc-blink-a{0%{opacity:1}50%{opacity:0}}@keyframes bym-doc-blink-b{0%{opacity:1}50%{opacity:0}}' +
  '.bym-doc-caret{background:#000;animation:bym-doc-blink-a 1.06s step-end infinite}' +
  '.bym-doc-caret.bym-doc-caret-b{animation-name:bym-doc-blink-b}' +
  '@media (prefers-reduced-motion:reduce){.bym-doc-caret,.bym-doc-caret.bym-doc-caret-b{animation:none}}' +
  '.bym-doc-compose{background:#1a73e8}' +
  '.bym-doc-handle{background:#1a73e8;border-radius:0 50% 50% 50%;touch-action:none}' +
  '.bym-doc-handle-start{border-radius:50% 0 50% 50%}' +
  '.bym-doc-bubble{background:#fff;border:1px solid #dadce0;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,.2);color:#202124}';

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

const isApple = (): boolean => typeof navigator !== 'undefined' && /Mac|iPhone|iPad|iPod/.test(navigator.platform || '');

const isLowSurrogate = (code: number): boolean => code >= 0xdc00 && code <= 0xdfff;

const copyPosition = (position: DocPosition): DocPosition => {
  const copy: DocPosition = { story: position.story, path: position.path.slice(), offset: position.offset };
  if (position.before === true) copy.before = true;
  return copy;
};

/** "header" of "header:rId7". */
const storyKind = (story: string): string => (story.indexOf(':') < 0 ? story : story.slice(0, story.indexOf(':')));

/** Lines of the text itself: not those in a text box, whose text cannot be edited. */
const outsideBoxes = (mark: LineMark): boolean => mark.source.path.indexOf(-1) < 0;

const coarsePointer = (): boolean => typeof window.matchMedia === 'function' && window.matchMedia('(pointer:coarse)').matches;

const resolveStrings = resolveUiStrings;

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
  /** Before what shows without width at the offset (a picture): see DocPosition.before. */
  before?: true;
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

const comparePositions = (a: TextPosition, b: TextPosition): number =>
  comparePaths(a.path, b.path) || a.offset - b.offset || (a.before === true ? 0 : 1) - (b.before === true ? 0 : 1);

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

/** One call of print(): what it holds while it runs. */
interface PrintJob {
  promise: Promise<DocumentPrintResult>;
  frame?: PrintFrame;
  store?: ImageStore;
  /** Set once the pages are ready and the dialog was asked for. */
  dialog?: PrintDialog;
  /** Pages ready so far, while they are prepared. */
  progress?: { done: number; total: number };
}

const sleep = (ms: number): Promise<void> => new Promise<void>((resolve) => setTimeout(resolve, ms));

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
  private readonly printButton: HTMLButtonElement | undefined;
  private readonly pagesLabel: HTMLSpanElement | undefined;
  private readonly styleHost: Node;
  private readonly painter: CanvasTextMeasurer;
  private readonly measurer: TextMeasurer;
  private readonly images: ImageStore;
  private readonly pictures: { readPart(part: string): Promise<Uint8Array> };
  private printJob: PrintJob | undefined;
  /** Counts the times the pages were laid out anew: pages prepared for printing before are out of date. */
  private layoutEpoch = 0;
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
  private unwatchFonts: (() => void) | undefined;
  private unwatchEmbeddedFont: (() => void) | undefined;
  /** Fonts arrived while a composition or a drag was under way: the layout is due when it ends. */
  private fontsDue = false;
  private resizeObserver: ResizeObserver | undefined;
  private dprQuery: MediaQueryList | undefined;
  private destroyed = false;

  private reportedPage = 0;
  private reportedCount = 0;
  /** A page asked for before it was laid out (from 0). */
  private pendingPage: number | undefined;
  /** With `pendingPage`: where in that page the view began (page px from its top) before a new layout, to begin there again; undefined for the top of the page. */
  private pendingWithin: number | undefined;
  private quietFocus = false;

  private selection: { anchor: TextPosition; focus: TextPosition } | 'all' | undefined;
  private dragging = false;
  private dragMoved = false;
  private pinch: { distance: number; zoom: number } | undefined;
  private gestureZoom = 0;

  /** Undo and redo of what was changed in the view. */
  readonly history: ops.DocHistory<DocRange | undefined>;
  private editable = false;
  private readonly lines: DocLines;
  private readonly typing: DocInput;
  private readonly caret: HTMLDivElement;
  private readonly underline: HTMLDivElement;
  private handles: HTMLDivElement[] = [];
  private bubble: HTMLDivElement | undefined;
  private modified = false;
  private focused = false;
  /** Which line a caret on the border between two lines is shown on. */
  private affinity: Affinity = 'start';
  /** The x the caret keeps while it moves up and down, px of the page. */
  private targetX: number | undefined;
  /** The story the caret is in: the body, or a header, footer or note entered with a double click. */
  private story = 'body';
  /** Formatting chosen at a caret, for what is typed next. */
  private pendingFormat: RunFormatPatch | undefined;
  /** Page the caret was last seen on: where it is looked for first. */
  private caretPage = 0;
  private blink = false;
  private hadRange = false;
  /** Characters before the caret that an input method is still composing. */
  private composed = 0;
  /** Escape was pressed: the next Tab leaves the editor. */
  private escaped = false;
  private lastDown = { time: 0, x: 0, y: 0, count: 0 };
  private dragPoint: { clientX: number; clientY: number } | undefined;
  private dragTimer: ReturnType<typeof setInterval> | undefined;
  private changes = 0;
  /** The selection as it is shown, for the selection it was worked out for. */
  private expanded: { anchor: TextPosition; focus: TextPosition; range: DocRange } | undefined;
  /** Height the strip keeps while pages that were there before an edit are laid out again, screen px. */
  private hold = 0;
  private steps = 0;
  /** A finger is down: where, and the timer of a long press. */
  private press: { x: number; y: number; timer: ReturnType<typeof setTimeout> | undefined; long: boolean; moved: boolean } | undefined;
  private lastTouch = 0;
  /** The selection was made by touch: it has handles and a bubble. */
  private touchSelection = false;
  private handleFixed: DocPosition | undefined;
  private inputPlace = '';
  private viewport: { addEventListener(type: string, listener: () => void): void; removeEventListener(type: string, listener: () => void): void } | undefined;
  /** The protection of the document that keeps the view read-only, whatever is asked of it. */
  private readonly lock: 'readOnly' | 'comments' | 'forms' | undefined;
  /** The user was told why the document cannot be edited. */
  private lockTold = false;

  constructor(container: HTMLElement, doc: DocModel, options: DocumentViewOptions = {}) {
    this.container = container;
    this.document = doc;
    this.options = options;
    this.strings = resolveStrings(options.locale);
    this.pointerEvents = 'PointerEvent' in window;
    const coarse = coarsePointer();
    this.maxCanvasPixels = coarse ? MAX_CANVAS_PIXELS_COARSE : MAX_CANVAS_PIXELS;

    // A protection that allows comments or form fields only is one that allows no editing here: the view has neither.
    const protection = enforcedProtection(doc);
    this.lock = protection !== undefined && protection !== 'trackedChanges' ? protection : undefined;
    this.editable = options.readOnly === false && this.lock === undefined;
    ops.setTracking(doc, { force: options.trackChanges === true, author: typeof options.author === 'string' ? options.author : undefined, now: typeof options.now === 'function' ? options.now : undefined });
    this.history = new ops.DocHistory<DocRange | undefined>(doc);
    this.painter = new CanvasTextMeasurer({ fonts: doc.fonts });
    this.measurer = options.measurer || this.painter;
    // Pictures put in while editing are in the model (a view may be made editable later); the others where the document was read from.
    const source = options.package;
    const pictures = {
      readPart: (part: string): Promise<Uint8Array> => {
        const own = doc.media && doc.media.get(part);
        if (own) return Promise.resolve(own.bytes);
        return source ? source.readPart(part) : Promise.reject(new Error('No such picture'));
      },
    };
    this.pictures = pictures;
    this.images = new ImageStore(pictures, coarse ? IMAGE_BUDGET_BYTES_COARSE : IMAGE_BUDGET_BYTES, this.onImageLoad);
    // A document that cannot be laid out at all throws here, before anything is added to the container.
    this.engine = new DocLayout(doc, this.measurer, this.layoutOptions());
    this.pages = this.engine.pages;
    this.lines = new DocLines(() => this.pages, this.measurer);

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
      this.printButton = this.button(bar, 'bym-doc-print', icon(['M7 8V4h10v4', 'M7 17H5a1 1 0 0 1-1-1v-6a2 2 0 0 1 2-2h12a2 2 0 0 1 2 2v6a1 1 0 0 1-1 1h-2', 'M7 14h10v6H7z']));
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

    const caret = document.createElement('div');
    caret.className = 'bym-doc-caret';
    caret.setAttribute('aria-hidden', 'true');
    caret.style.cssText = 'position:absolute;z-index:1;width:2px;min-width:1px;margin:0 0 0 -1px;padding:0;border:0;pointer-events:none';
    const underline = document.createElement('div');
    underline.className = 'bym-doc-compose';
    underline.setAttribute('aria-hidden', 'true');
    underline.style.cssText = 'position:absolute;z-index:1;height:1px;margin:0;padding:0;border:0;pointer-events:none';

    this.root = root;
    this.scroller = scroller;
    this.strip = strip;
    this.note = note;
    this.input = input;
    this.live = live;
    this.caret = caret;
    this.underline = underline;
    this.typing = new DocInput(input, {
      context: () => this.inputContext(),
      change: (change, composition) => this.inputChange(change, composition),
      type: (text) => {
        this.typeText(text);
      },
      remove: (direction, unit) => this.removeText(direction, unit),
      paragraph: () => this.enter('paragraph'),
      undo: () => {
        this.undo();
      },
      redo: () => {
        this.redo();
      },
      compositionStart: () => {
        this.history.seal();
        this.composed = 0;
      },
      compositionEnd: () => {
        this.history.seal();
        this.composed = 0;
        this.schedule();
        this.fontsSettled();
      },
    });
    this.applyInputMode();

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
    // Editing was asked for and the protection of the document does not allow it.
    if (options.readOnly === false) this.tellLock();
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
    this.pendingWithin = undefined;
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
    if (this.editable) {
      // A range of the story the caret is in: what is typed next takes its place.
      const start = this.storyEdge(-1);
      const end = this.storyEdge(1);
      if (!start || !end) return;
      this.typing.commit();
      this.selection = { anchor: start, focus: end };
      this.affinity = 'start';
      this.selectionMoved(false);
      return;
    }
    this.selection = 'all';
    this.repaintAll();
  }

  clearSelection(): void {
    if (this.destroyed || this.selection === undefined) return;
    this.selection = undefined;
    this.hideTouchTools();
    this.repaintAll();
    if (this.editable) {
      this.typing.sync();
      if (this.options.onSelectionChange) this.options.onSelectionChange(undefined);
    }
  }

  /** The selected text: what a copy puts on the clipboard. Empty when nothing is selected. */
  getSelectedText(): string {
    const selection = this.selection;
    if (selection === undefined) return '';
    if (selection === 'all') return documentText(this.document);
    const shown = this.shownRange();
    const range = shown ? orderedRange(shown.anchor, shown.focus) : undefined;
    return range ? rangeTextOf(this.document, range) : '';
  }

  /**
   * The selection as the pages show it. In an editor a selection that goes
   * into or out of a table covers whole cells: what a deletion would remove
   * (see expandRangeForTables of ops/doc).
   */
  private shownRange(): DocRange | undefined {
    const selection = this.selection;
    if (selection === undefined || selection === 'all') return undefined;
    if (!this.editable || ops.samePosition(selection.anchor, selection.focus)) return selection;
    const cached = this.expanded;
    if (cached && cached.anchor === selection.anchor && cached.focus === selection.focus) return cached.range;
    let range: DocRange = selection;
    try {
      range = ops.expandRangeForTables(this.document, selection);
    } catch {
      range = selection;
    }
    this.expanded = { anchor: selection.anchor, focus: selection.focus, range };
    return range;
  }

  /** Changes the language of the view's own words. */
  setLocale(locale: string | object): void {
    if (this.destroyed) return;
    this.strings = resolveStrings(locale);
    this.root.setAttribute('aria-label', this.strings.document);
    this.applyInputMode();
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
    // Where the reader is: the page most in view, and how far into it the view begins. The new layout shows
    // the same place again (fonts that arrive must not take a reader back to the top of the page), and the
    // strip keeps its height until the pages are there, so that the browser does not scroll in between.
    const within = this.pages.length > 0 && this.scale > 0 ? (this.virtualY() - this.pageTop(page)) / this.scale : undefined;
    const height = this.pages.length > 0 ? this.totalHeight() : 0;
    if (this.layoutTimer !== undefined) clearTimeout(this.layoutTimer);
    this.layoutTimer = undefined;
    this.painter.clear();
    this.layoutEpoch++;
    // Places in the text are places of the model: an editor keeps its caret.
    if (!this.editable) this.selection = undefined;
    this.hold = height;
    this.unmountAll();
    this.engine = new DocLayout(this.document, this.measurer, this.layoutOptions());
    this.pages = this.engine.pages;
    this.offsets = [0];
    this.done = false;
    this.failed = false;
    this.pendingPage = within !== undefined ? page : undefined;
    this.pendingWithin = within;
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
    this.endPress();
    this.typing.destroy();
    this.unlisten();
    this.unmountAll();
    this.images.destroy();
    // A print job under way ends here: its pages leave the document at once.
    if (this.printJob) this.endPrint(this.printJob);
    if (this.root.parentNode) this.root.parentNode.removeChild(this.root);
    releaseStyle(this.styleHost);
  }

  // -------------------------------------------------------------- printing

  /** Whether print() is at work: its pages are being prepared or its dialog is open. */
  get printing(): boolean {
    return this.printJob !== undefined;
  }

  /**
   * Prints the document: lays out what is not laid out yet, paints every
   * page at the resolution of a printer, and opens the print dialog of the
   * window the view is in, with paper of the size of the pages. The pages are
   * prepared one per turn, so the page stays responsive, and the document
   * can be read on meanwhile; when it is changed, they are prepared anew.
   *
   * Resolves when the dialog was closed (whether the user printed or
   * cancelled cannot be told). A call while the pages of an earlier one are
   * being prepared gets the promise of that call and opens no second dialog.
   * Rejects with a DocumentPrintError: where the window cannot print, for a
   * document beyond the limits, and when the view is destroyed meanwhile;
   * with a RangeError for a `dpi` that is no positive number.
   */
  print(options: DocumentPrintOptions = {}): Promise<DocumentPrintResult> {
    if (this.destroyed) return Promise.reject(new DocumentPrintError('cancelled', 'The view was destroyed'));
    // A wrong option is refused before any work is done, whatever else is under way.
    try {
      planPrint([DEFAULT_PAGE], options);
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
    this.updateBar();
    return job.promise;
  }

  private async runPrint(job: PrintJob, options: DocumentPrintOptions): Promise<DocumentPrintResult> {
    const cancelled = () => new DocumentPrintError('cancelled', 'Printing was cancelled: the view was destroyed');
    if (this.destroyed) throw cancelled();
    // The window of the document the view is in: inside a frame, the frame is what is printed.
    const doc = this.container.ownerDocument || document;
    const win = doc.defaultView;
    if (!win || typeof win.print !== 'function') throw new DocumentPrintError('unavailable', 'Printing is not available here: the window has no print()');
    this.notify(this.strings.printPreparing, 'info');
    const limits = printLimits(coarsePointer());
    const urls = (win as unknown as { URL?: typeof URL }).URL || URL;
    let plan: PrintPlan;
    for (let attempt = 0; ; attempt++) {
      // The rest of the document is laid out by the view's own slices.
      while (!this.done && !this.failed) {
        // A document beyond the limit is refused as soon as it is known to be, not after all of it was laid out.
        if (this.pages.length > MAX_PRINT_PAGES) break;
        await sleep(30);
        if (this.destroyed) throw cancelled();
      }
      // Pages up to where the layout gave up are not the document: nothing is printed rather than a part of it.
      if (this.failed) throw new DocumentPrintError('failed', 'The document could not be laid out to its end, so it is not printed');
      const sizes: { width: number; height: number }[] = [];
      for (let i = 0; i < this.pages.length; i++) sizes.push({ width: this.pageWidth(i), height: this.pageHeight(i) });
      plan = planPrint(sizes, options, limits);
      const frame = new PrintFrame(doc, plan, supportsNamedPages(win as unknown as { CSS?: { supports?(property: string, value: string): boolean } }), (url) => revokePicture(url, urls));
      const store = new ImageStore(this.pictures, this.images.budget, () => undefined);
      job.frame = frame;
      job.store = store;
      job.progress = { done: 0, total: plan.pages.length };
      this.updateBar();
      const complete = await this.paintForPrint(job, plan, frame, store, limits, doc, urls);
      store.destroy();
      job.store = undefined;
      if (this.destroyed) throw cancelled();
      if (complete) break;
      // The document changed under the pages: they are prepared again from the start.
      frame.dispose();
      job.frame = undefined;
      if (attempt >= MAX_PRINT_RESTARTS) throw new DocumentPrintError('failed', 'The document kept changing while it was prepared for printing');
    }
    job.progress = undefined;
    const dialog = openPrintDialog(win as unknown as Parameters<typeof openPrintDialog>[0]);
    job.dialog = dialog;
    this.updateBar();
    await dialog.closed;
    // Takes the place of "preparing": whether anything was printed cannot be told.
    if (!this.destroyed) this.notify(this.strings.printClosed, 'info');
    return { pages: plan.pages.length, dpi: plan.dpi };
  }

  /** Paints the pages of a plan and adds them to the frame; false when the pages were laid out anew meanwhile. */
  private async paintForPrint(
    job: PrintJob,
    plan: PrintPlan,
    frame: PrintFrame,
    store: ImageStore,
    limits: PrintLimits,
    doc: Document,
    urls: { createObjectURL(blob: Blob): string; revokeObjectURL(url: string): void },
  ): Promise<boolean> {
    const epoch = this.layoutEpoch;
    const stale = () => this.destroyed || this.layoutEpoch !== epoch || !(this.done || this.failed);
    // One canvas for all pages: a page is held as pixels only while it is painted.
    const canvas = doc.createElement('canvas');
    const win = doc.defaultView;
    const nextFrame = win && typeof win.requestAnimationFrame === 'function' ? (callback: () => void) => win.requestAnimationFrame(callback) : undefined;
    let bytes = 0;
    try {
      for (let i = 0; i < plan.pages.length; i++) {
        if (stale()) return false;
        const page = this.pages[i];
        const planned = plan.pages[i];
        if (!page) return false;
        if (canvas.width !== planned.pixelWidth) canvas.width = planned.pixelWidth;
        if (canvas.height !== planned.pixelHeight) canvas.height = planned.pixelHeight;
        const context = canvas.getContext('2d');
        if (!context) throw new DocumentPrintError('failed', 'No canvas to paint the pages on');
        const paint = () => {
          context.setTransform(1, 0, 0, 1, 0, 0);
          context.fillStyle = '#ffffff';
          context.fillRect(0, 0, planned.pixelWidth, planned.pixelHeight);
          renderPage(context, page, {
            scale: planned.pixelWidth / planned.width,
            images: store,
            cssFont: this.painter.cssFont,
            metrics: this.metricsOf,
            placeholderLabel: this.placeholderLabel,
          });
        };
        // Pictures of the pages before this one may make room for those of this one.
        store.pass++;
        try {
          paint();
          if (this.picturesPending(page, store)) {
            // Painting asked for the pictures; once they are decoded the page is painted with them.
            const deadline = Date.now() + PRINT_PICTURES_MS;
            while (store.busy && Date.now() < deadline) {
              await sleep(20);
              if (stale()) return false;
            }
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
          throw new DocumentPrintError('too-long', 'The pages of the document take too much memory to print at once');
        }
        await frame.add(picture.url, this.strings.page(i + 1));
        job.progress = { done: i + 1, total: plan.pages.length };
        this.updateBar();
      }
      return !stale();
    } finally {
      // Give the backing store back right away.
      canvas.width = 0;
      canvas.height = 0;
    }
  }

  private picturesPending(page: LayoutPage, store: ImageStore): boolean {
    const items = page.items;
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      if (item.type === 'image' && item.part !== undefined && item.placeholder === undefined && store.status(item.part) === 'pending') return true;
    }
    return false;
  }

  /** Takes what a print job holds out of the document. Safe to call twice. */
  private endPrint(job: PrintJob): void {
    if (job.frame) job.frame.dispose();
    if (job.store) job.store.destroy();
    if (job.dialog) job.dialog.finish();
    job.frame = undefined;
    job.store = undefined;
    job.progress = undefined;
    if (this.printJob !== job) return;
    this.printJob = undefined;
    if (!this.destroyed) this.updateBar();
  }

  private printFailed(error: unknown): void {
    // A wrong option is the caller's to hear about, not the user's; so is the end of a view.
    if (this.destroyed || error instanceof RangeError) return;
    const code = error instanceof DocumentPrintError ? error.code : 'failed';
    if (code === 'cancelled') return;
    const strings = this.strings;
    this.notify(code === 'unavailable' ? strings.printNotAvailable : code === 'too-long' ? strings.printTooLong : strings.printFailed, 'error');
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
    root.addEventListener('cut', this.onCut);
    root.addEventListener('paste', this.onPaste);
    root.addEventListener('mousedown', this.onRootMouseDown);
    strip.addEventListener('contextmenu', this.onContextMenu);
    const viewport = (window as { visualViewport?: DocumentView['viewport'] }).visualViewport;
    if (viewport && typeof viewport.addEventListener === 'function') {
      this.viewport = viewport;
      viewport.addEventListener('resize', this.onViewportChange);
      viewport.addEventListener('scroll', this.onViewportChange);
    }
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
    this.unwatchFonts = onFontsChanged(this.onFontsRegistered);
    // The font of the package was needed and could not be loaded: said once, by one view of the page.
    this.unwatchEmbeddedFont = onEmbeddedFontNotice((notice) => {
      if (!this.destroyed) this.notify(this.strings.embeddedFontNotLoaded(notice.family), 'info');
    });
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
    root.removeEventListener('cut', this.onCut);
    root.removeEventListener('paste', this.onPaste);
    root.removeEventListener('mousedown', this.onRootMouseDown);
    strip.removeEventListener('contextmenu', this.onContextMenu);
    if (this.viewport) {
      this.viewport.removeEventListener('resize', this.onViewportChange);
      this.viewport.removeEventListener('scroll', this.onViewportChange);
      this.viewport = undefined;
    }
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
    if (this.unwatchFonts) this.unwatchFonts();
    this.unwatchFonts = undefined;
    if (this.unwatchEmbeddedFont) this.unwatchEmbeddedFont();
    this.unwatchEmbeddedFont = undefined;
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
      // Nothing more comes: the strip is as high as its pages.
      this.hold = 0;
      this.notify(this.strings.documentLayoutFailed, 'error');
      this.update();
      return;
    }
    // An edit whose blocks were still being laid out turned out to be one the engine cannot follow.
    if (this.engine.broken) this.rebuild();
    else this.afterLayout(progress);
    this.schedule();
  };

  /** Takes over what a slice of layout produced and asks for the next one. */
  private afterLayout(progress: LayoutProgress): void {
    const pages = this.pages;
    const before = this.offsets.length - 1;
    for (let i = before; i < pages.length; i++) this.offsets.push(this.offsets[i] + this.pageHeight(i));
    // Pages that were laid out again after an edit may still have a canvas from before.
    if (this.slots.size > 0) {
      this.slots.forEach((slot) => {
        if (slot.index >= before && slot.index < pages.length) {
          slot.dirty = true;
          this.describe(slot);
        }
      });
    }
    for (let i = 0; i < progress.updated.length; i++) {
      const slot = this.slots.get(progress.updated[i]);
      if (slot) {
        slot.dirty = true;
        this.describe(slot);
      }
    }
    this.done = progress.done;
    if (this.done) this.hold = 0;
    if (this.engine.truncated && !this.truncationSaid) {
      this.truncationSaid = true;
      this.notify(this.strings.documentTooLong, 'info');
    }
    // The first page decides the zoom of the modes that fit.
    if (before === 0 && pages.length > 0 && this.mode !== 'manual') this.scale = this.fitZoom();
    this.sizeStrip();
    if (this.pendingPage !== undefined && (this.pendingPage < pages.length || this.done)) {
      const target = Math.min(this.pendingPage, pages.length - 1);
      const within = target === this.pendingPage ? this.pendingWithin : undefined;
      this.pendingPage = undefined;
      this.pendingWithin = undefined;
      if (within === undefined || target < 0) this.scrollToPage(target);
      else {
        // The place in the page the view was at before the layout was made anew.
        this.setVirtualY(this.pageTop(target) + within * this.scale);
        this.update();
      }
    }
    if (this.layoutTimer !== undefined) clearTimeout(this.layoutTimer);
    this.layoutTimer = this.done ? undefined : setTimeout(this.runLayout, 0);
  }

  private layoutOptions(): DocLayoutOptions | undefined {
    // The marks of the lines are what a caret is placed by; a view that only shows the document does without them.
    return this.editable ? { ...this.options.layout, lineMarks: true } : this.options.layout;
  }

  /** A new layout of the whole document that keeps what is on screen until its pages arrive: after an edit the engine cannot follow. */
  private rebuild(): void {
    if (this.layoutTimer !== undefined) clearTimeout(this.layoutTimer);
    this.layoutTimer = undefined;
    this.layoutEpoch++;
    let progress: LayoutProgress;
    try {
      this.engine = new DocLayout(this.document, this.measurer, this.layoutOptions());
      this.pages = this.engine.pages;
      this.offsets = [0];
      this.done = false;
      this.failed = false;
      progress = this.engine.layout({ milliseconds: FIRST_SLICE_MS });
    } catch {
      this.layoutBroke();
      return;
    }
    this.afterLayout(progress);
  }

  private layoutBroke(): void {
    this.failed = true;
    this.notify(this.strings.documentLayoutFailed, 'error');
  }

  /** Brings the pages up to date after the document was changed, at least as far as the caret. */
  private follow(): void {
    this.layoutEpoch++;
    const height = this.totalHeight();
    let progress: (LayoutProgress & { from: number }) | undefined;
    const started = Date.now();
    try {
      // A change of very many blocks (the whole document made bold) is laid out in slices like a document that is opened.
      progress = this.failed ? undefined : this.engine.edited({ milliseconds: EDIT_LAYOUT_MS });
    } catch {
      progress = undefined;
    }
    // Pages after the edit come back in slices: the strip keeps its height until then, so that nothing jumps.
    if (height > this.hold) this.hold = height;
    if (!progress) this.rebuild();
    else {
      if (progress.from < this.offsets.length - 1) this.offsets.length = progress.from + 1;
      this.afterLayout(progress);
    }
    const range = this.range();
    while (!this.done && !this.failed && range && !this.lines.lineOf(range.focus, 'start', this.caretPage) && Date.now() - started < EDIT_LAYOUT_MS) {
      let next: LayoutProgress;
      try {
        next = this.engine.layout({ pages: 2 });
      } catch {
        this.layoutBroke();
        break;
      }
      if (this.engine.broken) this.rebuild();
      else this.afterLayout(next);
    }
  }

  private onFontsLoaded = (): void => {
    if (this.destroyed || this.fontTimer !== undefined || this.fontRelayouts >= MAX_FONT_RELAYOUTS) return;
    // Fonts tend to arrive in groups.
    this.fontTimer = setTimeout(() => {
      this.fontTimer = undefined;
      this.fontRelayouts++;
      this.relayoutForFonts();
    }, 50);
  };

  /**
   * The application registered a font or took one out (registerFont()): what
   * was measured without it is measured anew. Not counted as the fonts of the
   * page are: the application decides how often this happens.
   */
  private onFontsRegistered = (): void => {
    if (this.destroyed || this.fontTimer !== undefined) return;
    // Several fonts registered together are one layout.
    this.fontTimer = setTimeout(() => {
      this.fontTimer = undefined;
      this.relayoutForFonts();
    }, 50);
  };

  /**
   * Lays the document out again because fonts arrived, but not under the
   * hands of the user: while an input method composes or a selection is
   * dragged the layout waits, and fontsSettled() runs it when that is over.
   */
  private relayoutForFonts(): void {
    if (this.destroyed) return;
    if (this.typing.composing || this.dragging) {
      this.fontsDue = true;
      return;
    }
    this.fontsDue = false;
    this.relayout();
  }

  /** A composition or a drag ended: the layout that waited for it runs. */
  private fontsSettled(): void {
    if (!this.fontsDue || this.destroyed || this.fontTimer !== undefined) return;
    // After the event that ended it has been dealt with to its end.
    this.fontTimer = setTimeout(() => {
      this.fontTimer = undefined;
      this.relayoutForFonts();
    }, 0);
  }

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
    const height = PAD * 2 + this.offsets[count] * this.scale + Math.max(0, count - 1) * GAP;
    return !this.done && this.hold > height ? this.hold : height;
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
    this.hold = 0;
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

    this.placeCaret();
    this.updateBar();
    const page = this.currentPage;
    if (page !== this.reportedPage || count !== this.reportedCount) {
      const turned = page !== this.reportedPage;
      this.reportedPage = page;
      this.reportedCount = count;
      // The live region follows the page in view, not the count that grows while the layout runs.
      // An editor says where the caret is instead: every page turn while typing would be too much.
      if (turned && page > 0 && !(this.editable && this.focused)) this.live.textContent = this.strings.pageOf(page, count);
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
    const job = this.printJob;
    if (this.printButton) {
      const text = strings.print;
      if (this.printButton.getAttribute('aria-label') !== text) {
        this.printButton.setAttribute('aria-label', text);
        this.printButton.title = text;
      }
      const busy = job && !job.dialog ? 'true' : 'false';
      if (this.printButton.getAttribute('aria-busy') !== busy) this.printButton.setAttribute('aria-busy', busy);
    }
    if (this.pagesLabel) {
      const count = this.pages.length;
      const preparing = job && !job.dialog ? job.progress : undefined;
      // "+": more pages are on their way.
      const text = preparing
        ? strings.printProgress(preparing.done, preparing.total)
        : count === 0 ? (this.done || this.failed ? '' : strings.preparingPages) : strings.pageOf(this.currentPage, count) + (this.done || this.failed ? '' : '+');
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
      // Text a tracked change deleted is on the page, struck through, but it is not text of the document.
      if (item.type !== 'text' || item.deleted) continue;
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
    const shown = this.shownRange();
    const range = selection === 'all' || !shown ? undefined : orderedRange(shown.anchor, shown.focus);
    if (selection !== 'all' && !range) return;
    const items = this.pages[pageIndex].items;
    context.save();
    context.scale(scale, scale);
    context.fillStyle = SELECTION_FILL;
    if (this.editable && range && !range.cells) {
      // Pictures in the line take no offset of their own: one is selected when the selection runs from before it to after it.
      const lines = this.lines.linesOf(pageIndex);
      for (let i = 0; i < lines.length; i++) {
        const mark = lines[i].mark;
        if (!mark.boxes || !sameFlow(mark.source, range.start)) continue;
        for (const box of mark.boxes) {
          const before: TextPosition = { story: mark.source.story, path: mark.source.path, offset: box.offset, before: true };
          const after: TextPosition = { story: mark.source.story, path: mark.source.path, offset: box.offset };
          if (comparePositions(range.start, before) <= 0 && comparePositions(after, range.end) <= 0) context.fillRect(mark.x + box.dx, mark.y, box.width, mark.height);
        }
      }
    }
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
    // The mouse events a browser makes up after a tap are not presses of a mouse.
    if (Date.now() - this.lastTouch < 800) return;
    if (this.editable) {
      this.editPointerDown(event);
      return;
    }
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
    // A second press without the end of the first one (a pointer that was lost) must not leave a timer behind.
    if (this.dragTimer !== undefined) clearInterval(this.dragTimer);
    this.dragging = true;
    window.addEventListener(this.pointerEvents ? 'pointermove' : 'mousemove', this.onPointerMove as EventListener);
    window.addEventListener(this.pointerEvents ? 'pointerup' : 'mouseup', this.onPointerUp);
    if (this.pointerEvents) window.addEventListener('pointercancel', this.onPointerUp);
  };

  private onPointerMove = (event: PointerEvent | MouseEvent): void => {
    if (!this.dragging || this.destroyed) return;
    const selection = this.selection;
    if (!selection || selection === 'all') return;
    if (this.editable) {
      this.dragPoint = { clientX: event.clientX, clientY: event.clientY };
      this.dragTo(this.dragPoint);
      return;
    }
    const position = this.positionFromEvent(event, selection.anchor);
    if (!position) return;
    if (comparePositions(selection.anchor, position) !== 0) this.dragMoved = true;
    this.setFocusPosition(position);
  };

  private onPointerUp = (): void => {
    this.endDrag();
  };

  private endDrag(): void {
    if (this.dragTimer !== undefined) clearInterval(this.dragTimer);
    this.dragTimer = undefined;
    this.dragPoint = undefined;
    if (!this.dragging) return;
    this.dragging = false;
    this.fontsSettled();
    window.removeEventListener('pointermove', this.onPointerMove as EventListener);
    window.removeEventListener('mousemove', this.onPointerMove as EventListener);
    window.removeEventListener('pointerup', this.onPointerUp);
    window.removeEventListener('mouseup', this.onPointerUp);
    window.removeEventListener('pointercancel', this.onPointerUp);
  }

  private onDoubleClick = (event: MouseEvent): void => {
    if (this.destroyed) return;
    if (this.editable) {
      // A mouse was counted when it was pressed; this is the second tap of a finger.
      if (Date.now() - this.lastDown.time > CLICK_MS * 2) this.selectWordAt(event, false);
      return;
    }
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
    this.copyTo(event, this.selectionSlice());
  };

  /**
   * Puts the selection on the clipboard: its text, and HTML made of that text
   * with the mark of the copy the page keeps. Says what the clipboard holds
   * afterwards: the mark of the copy ("rich"), the text alone ("plain"), or
   * nothing of the selection ("none": there was nothing to put there).
   */
  private copyTo(event: ClipboardEvent, slice: DocSlice | undefined, cut = false): 'rich' | 'plain' | 'none' {
    const data = event.clipboardData;
    if (this.destroyed || !data) return 'none';
    const text = this.getSelectedText();
    // A picture alone is a copy without text.
    const fragment = slice ? slice.fragment : undefined;
    if (text === '' && !(fragment && ops.fragmentHasContent(fragment))) {
      // The blank the input holds for the sake of this event must not replace what is on the clipboard.
      if (event.target === this.input) event.preventDefault();
      return 'none';
    }
    event.preventDefault();
    data.setData('text/plain', text);
    if (!slice) return 'plain';
    // The same as simple HTML, made here from escaped text, with a mark by which a paste into an editor of this page knows the copy.
    const copy = storeCopy(text, slice, cut);
    try {
      data.setData('text/html', sliceHtml(slice, copy.mark));
    } catch {
      // Plain text is on the clipboard then: no paste will ask for the copy.
      return 'plain';
    }
    this.readCopyMedia(copy);
    return 'rich';
  }

  /** Reads the pictures of a copy that the model does not hold: another document needs their bytes to take the copy. */
  private readCopyMedia(copy: StoredCopy): void {
    const fragment = copy.slice.fragment;
    if (!fragment) return;
    const reads: Promise<void>[] = [];
    fragment.media.forEach((media, name) => {
      if (media.bytes) return;
      reads.push(
        this.pictures.readPart(name).then(
          (bytes) => {
            if (bytes instanceof Uint8Array && bytes.length <= MAX_PICTURE_BYTES) setCopyMedia(copy, name, bytes, sniffImageType(bytes));
          },
          () => undefined,
        ),
      );
    });
    if (reads.length) copy.ready = Promise.all(reads).then(() => undefined);
  }

  private onCut = (event: ClipboardEvent): void => {
    if (this.destroyed) return;
    const range = this.range();
    let slice: DocSlice | undefined;
    let computed = false;
    if (this.editable && range && !ops.isCollapsed(range)) {
      // What is removed must be on the clipboard: a cut that cannot put it there removes nothing.
      let refused: string | undefined;
      computed = true;
      if (!event.clipboardData) refused = this.strings.cannotWriteClipboard;
      else {
        try {
          slice = this.selectionSlice();
          // A tracked deletion takes nothing out of the document but text the same author typed, which the
          // clipboard holds as text: there is nothing more the clipboard has to hold for it.
          if (!this.trackingChanges && cutLoses(this.document, range, slice)) refused = this.strings.cannotCutObjects;
        } catch (error) {
          event.preventDefault();
          this.refuse(error);
          return;
        }
      }
      if (refused !== undefined) {
        event.preventDefault();
        this.notify(refused, 'info');
        return;
      }
    }
    const deleting = this.editable && range !== undefined && !ops.isCollapsed(range);
    const written = this.copyTo(event, computed ? slice : this.selectionSlice(), deleting);
    // Nothing goes out of the field itself either way.
    event.preventDefault();
    if (!deleting || !range) return;
    if (written !== 'rich' && slice && !this.trackingChanges) {
      // The mark of the copy did not reach the clipboard: what a paste gets is the text, so the cut is judged as one of text.
      let loses = true;
      try {
        loses = cutLoses(this.document, range, undefined);
      } catch {
        loses = true;
      }
      if (loses) {
        this.notify(this.strings.cannotWriteClipboard, 'info');
        return;
      }
    }
    this.perform((selection) => ops.deleteRange(this.document, selection));
  };

  private onPaste = (event: ClipboardEvent): void => {
    if (this.destroyed) return;
    event.preventDefault();
    const data = event.clipboardData;
    if (!this.editable || !data || !this.range()) return;
    this.typing.commit();
    let text = '';
    let html = '';
    try {
      text = data.getData('text/plain') || '';
      html = data.getData('text/html') || '';
    } catch {
      // Nothing to read.
    }
    this.pasteClipboard(text, html, text === '' ? this.clipboardPicture(data) : undefined);
  };

  /** Pastes what the clipboard holds: a copy made on this page as it was, a picture, or text. HTML from elsewhere is never parsed: its text is what is pasted. */
  private pasteClipboard(text: string, html: string, picture: Blob | undefined): void {
    const copy = storedCopy(html, text);
    if (copy) this.pasteCopy(copy, false);
    else if (picture) this.readPicture(picture);
    else this.pasteText(text);
  }

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
    if (this.editable) {
      // A tap of a finger puts the caret there; a link is followed with Ctrl (Cmd) held, as in Word.
      if (Date.now() - this.lastTouch < 800) {
        if (this.press && this.press.long) return;
        this.tapAt(event);
        return;
      }
      if (!(event.ctrlKey || event.metaKey)) return;
    }
    const target = this.linkOf(event.target);
    if (target === undefined) return;
    event.preventDefault();
    this.activateLink(target);
  };

  // -------------------------------------------------------------- keyboard

  private onKeyDown = (event: KeyboardEvent): void => {
    if (this.destroyed || event.defaultPrevented) return;
    const target = event.target as Element | null;
    const command = event.ctrlKey || event.metaKey;
    const key = event.key;
    // Ctrl+P, on the buttons of the bar too: the browser would print the host page with a part of the document in it.
    if (command && !event.altKey && !event.shiftKey && (key === 'p' || key === 'P')) {
      event.preventDefault();
      this.print().catch(() => undefined);
      return;
    }
    // The other keys on the buttons of the bar are theirs.
    if (this.bar && target && this.bar.contains(target)) return;
    if (!command && !event.altKey && (key === 'Enter' || key === ' ')) {
      const link = this.linkOf(target);
      if (link !== undefined) {
        event.preventDefault();
        this.activateLink(link);
        return;
      }
    }
    if (this.editable && this.editKey(event)) return;
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
      case 'Backspace':
        // Some browsers go back a page for it when the focus is not in a field that can be edited.
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
    this.focused = true;
    if (this.editable) {
      // The caret is where it was left; a view that never had one gets it at the start of the text.
      if (this.selection === undefined) {
        const start = this.storyEdge(-1);
        if (start) this.selection = { anchor: start, focus: copyPosition(start) };
      }
      this.typing.sync();
      this.schedule();
      return;
    }
    try {
      this.input.select();
    } catch {
      // Nothing depends on it but the copy event of some browsers.
    }
  };

  private onBlur = (): void => {
    this.root.classList.remove('bym-doc-focus');
    this.focused = false;
    if (!this.editable || this.destroyed) return;
    // What an input method was composing stays as text; the caret is hidden until the focus comes back.
    this.typing.commit();
    this.history.seal();
    this.escaped = false;
    this.schedule();
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
    else if (button === this.printButton) this.print().catch(() => undefined);
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
    this.lastTouch = Date.now();
    this.endPress();
    if (!this.destroyed && this.editable && event.touches.length === 1) {
      const touch = event.touches[0];
      const press: NonNullable<DocumentView['press']> = { x: touch.clientX, y: touch.clientY, timer: undefined, long: false, moved: false };
      // A finger that stays where it is selects the word under it.
      press.timer = setTimeout(() => {
        press.timer = undefined;
        if (this.destroyed || this.press !== press || press.moved) return;
        press.long = true;
        this.selectWordAt({ clientX: press.x, clientY: press.y }, true);
      }, LONG_PRESS_MS);
      this.press = press;
    }
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
    this.lastTouch = Date.now();
    const press = this.press;
    if (press && event.touches.length > 0) {
      const touch = event.touches[0];
      if (Math.abs(touch.clientX - press.x) > 10 || Math.abs(touch.clientY - press.y) > 10) {
        press.moved = true;
        if (press.timer !== undefined) clearTimeout(press.timer);
        press.timer = undefined;
      }
    }
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
    this.lastTouch = Date.now();
    const press = this.press;
    if (press && press.timer !== undefined) {
      clearTimeout(press.timer);
      press.timer = undefined;
    }
    if (event.touches.length < 2) this.pinch = undefined;
  };

  private endPress(): void {
    const press = this.press;
    if (press && press.timer !== undefined) clearTimeout(press.timer);
    this.press = undefined;
  }

  /** The menu a browser opens for a long press would cover the selection the press made. */
  private onContextMenu = (event: Event): void => {
    if (this.editable && Date.now() - this.lastTouch < 1500) event.preventDefault();
  };

  private onViewportChange = (): void => {
    // The keyboard of a phone came or went: the caret stays in what is left of the view.
    if (this.destroyed || !this.editable || !this.focused) return;
    this.reveal();
    this.schedule();
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

  // --------------------------------------------------------------- editing

  /** Whether the document can only be read. */
  get readOnly(): boolean {
    return !this.editable;
  }

  /**
   * What the protection of the document enforces: "readOnly", "comments" or
   * "forms" (the view is read-only and stays so), "trackedChanges" (every
   * change is tracked); undefined for a document that enforces none.
   */
  get protection(): 'readOnly' | 'comments' | 'forms' | 'trackedChanges' | undefined {
    return enforcedProtection(this.document);
  }

  /** Whether the changes made in the view are tracked: the document asks for it, or the options of the view do. */
  get trackingChanges(): boolean {
    return ops.isTracking(this.document);
  }

  /** Says, once, why the document cannot be edited. */
  private tellLock(): void {
    if (this.lockTold || this.lock === undefined) return;
    this.lockTold = true;
    this.notify(this.strings.protectedDocument(this.lock), 'info');
  }

  /**
   * Turns editing on or off. The pages are laid out again when it is turned
   * on for the first time. A document whose protection allows no editing
   * stays read-only (see `protection`).
   */
  setReadOnly(readOnly: boolean): void {
    if (this.destroyed) return;
    if (!readOnly && this.lock !== undefined) {
      this.tellLock();
      return;
    }
    if (this.editable === !readOnly) return;
    this.typing.commit();
    this.history.seal();
    this.hideTouchTools();
    this.editable = !readOnly;
    this.pendingFormat = undefined;
    this.applyInputMode();
    if (this.editable) {
      // A caret needs the marks of the lines, which a view that only shows the document does not have.
      this.selection = undefined;
      this.relayout();
    } else this.schedule();
    this.typing.sync();
  }

  /** Whether the document was changed in the view since it was made or markSaved() was called. */
  get isModified(): boolean {
    return this.modified;
  }

  /** Says that the document as it is now has been saved. */
  markSaved(): void {
    this.modified = false;
  }

  /** Counts the changes made in the view, undo and redo included: the same number means the same document. */
  get changeCount(): number {
    return this.changes;
  }

  get canUndo(): boolean {
    return this.editable && this.history.canUndo;
  }

  get canRedo(): boolean {
    return this.editable && this.history.canRedo;
  }

  /** Undoes the last change; the caret goes back to where it was before it. */
  undo(): boolean {
    if (!this.editable || this.destroyed) return false;
    this.typing.commit();
    const step = this.history.undo();
    if (!step) return false;
    this.afterHistory(step.meta);
    return true;
  }

  redo(): boolean {
    if (!this.editable || this.destroyed) return false;
    this.typing.commit();
    const step = this.history.redo();
    if (!step) return false;
    this.afterHistory(step.meta);
    return true;
  }

  /** The caret (a range whose ends are the same place) or the selection; undefined when there is none or everything is selected in a read-only view. */
  getSelection(): DocRange | undefined {
    const range = this.range();
    return range ? { anchor: copyPosition(range.anchor), focus: copyPosition(range.focus) } : undefined;
  }

  /** Puts the caret at a position or selects a range. Throws RangeError for a place the document does not have. */
  setSelection(selection: DocRange | DocPosition | undefined): void {
    if (this.destroyed) return;
    if (selection === undefined) {
      this.clearSelection();
      return;
    }
    const range = ops.asRange(selection);
    ops.resolvePosition(this.document, range.anchor);
    ops.resolvePosition(this.document, range.focus);
    if (!ops.samePosition(range.anchor, range.focus)) ops.comparePositions(range.anchor, range.focus);
    this.typing.commit();
    this.history.seal();
    this.selection = { anchor: copyPosition(range.anchor), focus: copyPosition(range.focus) };
    this.story = range.focus.story;
    this.affinity = 'start';
    this.targetX = undefined;
    this.pendingFormat = undefined;
    this.selectionMoved(true);
  }

  /** Types text at the caret, or over the selection: "\n" ends a paragraph, "\t" is a tab. False when nothing was changed. */
  insertText(text: string): boolean {
    if (typeof text !== 'string' || text === '') return false;
    return this.perform((range) => ops.insertText(this.document, range, text, this.takeFormat()), true);
  }

  /** Character formatting for the selection, or at a caret for what is typed next. */
  format(patch: RunFormatPatch): boolean {
    const range = this.range();
    if (!this.editable || !range) return false;
    if (ops.isCollapsed(range)) {
      // An empty paragraph takes the formatting on its mark, which is where its text will take it from.
      let empty = false;
      try {
        empty = ops.resolvePosition(this.document, range.focus).length === 0;
      } catch {
        empty = false;
      }
      // Where changes are tracked the mark is left alone (its formatting cannot be tracked yet): the text typed next takes the formatting.
      if (empty && !this.trackingChanges && this.perform((selection) => ops.setRunFormat(this.document, selection, patch), true)) return true;
      // Checked now, so that a patch that is not valid is refused here and not at the next key.
      try {
        ops.patchRunProps({}, patch);
      } catch (error) {
        this.refuse(error);
        return false;
      }
      this.pendingFormat = { ...this.pendingFormat, ...patch };
      this.announceSelection();
      return true;
    }
    return this.perform((selection) => ops.setRunFormat(this.document, selection, patch), true);
  }

  /** Turns bold, italic, underline or strikethrough on, or off when all of the selection has it. */
  toggleFormat(key: ToggleKey): boolean {
    const range = this.range();
    if (!this.editable || !range) return false;
    if (ops.isCollapsed(range)) {
      const now = this.selectionFormat;
      return this.format({ [key]: !(now && now[key] === true) });
    }
    return this.perform((selection) => ops.toggleRunFormat(this.document, selection, key), true);
  }

  /** Paragraph formatting (alignment, style, spacing ...) for every paragraph the selection touches. */
  formatParagraph(patch: ParaFormatPatch): boolean {
    return this.perform((range) => ops.setParaFormat(this.document, range, patch), true);
  }

  /** Makes the paragraphs of the selection items of a list, or takes them out of it. */
  toggleList(kind: 'bullet' | 'number'): boolean {
    return this.perform((range) => ops.toggleList(this.document, range, kind), true);
  }

  /** One level deeper (1) or less deep (-1): list items change their level, other paragraphs their indent. */
  indent(delta: number): boolean {
    return this.perform((range) => ops.indentList(this.document, range, delta), true);
  }

  insertTable(rows: number, columns: number): boolean {
    return this.perform((range) => ops.insertTable(this.document, this.caretOf(range), rows, columns), true);
  }

  insertRow(where: 'above' | 'below'): boolean {
    return this.perform((range) => ops.insertRow(this.document, range.focus, where), true);
  }

  deleteRow(): boolean {
    return this.perform((range) => ops.deleteRow(this.document, range.focus), true);
  }

  insertColumn(side: 'left' | 'right'): boolean {
    return this.perform((range) => ops.insertColumn(this.document, range.focus, side), true);
  }

  deleteColumn(): boolean {
    return this.perform((range) => ops.deleteColumn(this.document, range.focus), true);
  }

  insertPageBreak(): boolean {
    return this.perform((range) => ops.insertPageBreak(this.document, this.caretOf(range)), true);
  }

  /**
   * Makes the selected text a link to `target`; at a caret the address itself
   * becomes the text. Only http, https, mailto and tel addresses and "#name"
   * (a bookmark) are taken.
   */
  insertLink(target: string): boolean {
    const safe = safeHyperlinkTarget(target);
    if (safe === undefined || safe === '' || safe === '#') {
      this.notify(this.strings.linkNotAllowed, 'error');
      return false;
    }
    return this.perform((range) => ops.insertHyperlink(this.document, range, safe), true);
  }

  /**
   * Inserts a picture (PNG, JPEG, GIF or BMP bytes) at the caret, in line with
   * the text, no wider than the text is. What the bytes are is read from the
   * bytes themselves.
   */
  insertPicture(bytes: Uint8Array, description?: string): boolean {
    const range = this.range();
    if (!this.editable || !range) return false;
    const type = bytes instanceof Uint8Array && bytes.length <= MAX_PICTURE_BYTES ? sniffImageType(bytes) : undefined;
    const size = type !== undefined && /^image\/(png|jpeg|gif|bmp|webp)$/.test(type) ? declaredImageSize(bytes, type) : undefined;
    if (!type || !size || !(size.width > 0) || !(size.height > 0) || size.width > 100000 || size.height > 100000) {
      this.notify(this.strings.pictureNotSupported, 'error');
      return false;
    }
    let width = size.width;
    let height = size.height;
    const room = this.textRoom(range.focus);
    if (width > room.width) {
      height = (height * room.width) / width;
      width = room.width;
    }
    if (height > room.height) {
      width = (width * room.height) / height;
      height = room.height;
    }
    const image: ops.NewImage = { bytes, type, width: Math.max(1, Math.round(width * EMU_PER_PX)), height: Math.max(1, Math.round(height * EMU_PER_PX)) };
    if (description !== undefined) image.description = description;
    // The name of the new part must not be one the package has already, though the model does not know it.
    const source = this.options.package as { hasPart?(part: string): boolean } | undefined;
    const taken = this.options.partTaken || (source && typeof source.hasPart === 'function' ? (part: string) => (source.hasPart as (part: string) => boolean)(part) : undefined);
    return this.perform((selection) => ops.insertImage(this.document, this.caretOf(selection), image, taken ? { taken } : undefined), true);
  }

  /** The formatting in effect for the selection, with what was chosen at a caret for the text typed next; undefined without a selection. */
  get selectionFormat(): RangeFormat | undefined {
    const range = this.range();
    if (!range) return undefined;
    let format: RangeFormat;
    try {
      // Of a very large selection the formatting where it starts: going through all of it at every move would hold the page up.
      format = ops.rangeFormat(this.document, this.isLarge(range) ? ops.orderRange(range).start : range);
    } catch {
      return undefined;
    }
    const pending = this.pendingFormat;
    if (pending && ops.isCollapsed(range)) {
      const out = format as unknown as { [key: string]: unknown };
      const keys: (keyof RunFormatPatch)[] = ['bold', 'italic', 'strike', 'font', 'size', 'highlight', 'vertAlign'];
      for (const key of keys) if (pending[key] !== undefined && pending[key] !== null) out[key] = pending[key];
      if (pending.underline !== undefined && pending.underline !== null) out.underline = pending.underline !== false;
      if (typeof pending.color === 'string') out.color = pending.color;
    }
    return format;
  }

  /** Whether the caret is in a table cell. */
  get inTable(): boolean {
    const range = this.range();
    return range !== undefined && range.focus.path.length >= 4 && range.focus.path.indexOf(-1) < 0;
  }

  /** Ends what an input method is composing: the text it has typed so far stays in the document. Call before saving. */
  commitEditing(): void {
    if (this.destroyed) return;
    this.typing.commit();
    this.history.seal();
  }

  /** The selection as a range of the model; undefined for "everything" in a read-only view and without a selection. */
  private range(): DocRange | undefined {
    const selection = this.selection;
    return selection === undefined || selection === 'all' ? undefined : selection;
  }

  /** Where something is put in: the caret, or the start of the selection. */
  private caretOf(range: DocRange): DocPosition {
    return ops.isCollapsed(range) ? range.focus : ops.orderRange(range).start;
  }

  private takeFormat(): RunFormatPatch | undefined {
    const format = this.pendingFormat;
    this.pendingFormat = undefined;
    return format;
  }

  private applyInputMode(): void {
    const input = this.input;
    input.readOnly = !this.editable;
    input.setAttribute('aria-label', this.editable ? this.strings.documentEditor : this.strings.document);
    if (this.editable) {
      input.setAttribute('role', 'textbox');
      input.setAttribute('aria-multiline', 'true');
      input.setAttribute('inputmode', 'text');
      // Text smaller than 16 px makes iOS zoom the page when the field takes the focus.
      input.style.fontSize = '16px';
    } else {
      input.removeAttribute('role');
      input.removeAttribute('aria-multiline');
      input.setAttribute('inputmode', 'none');
      input.style.fontSize = '';
    }
  }

  /** A place that cannot be edited is not an error of the program: said, and nothing happens. */
  private refuse(error: unknown): void {
    if (!(error instanceof RangeError)) throw error;
    this.notify(error.message === ops.TRACK_UNSUPPORTED ? this.strings.cannotTrackChange : this.strings.cannotEditHere, 'info');
  }

  /** First (-1) or last (1) place of the story the caret is in. */
  private storyEdge(direction: 1 | -1): DocPosition | undefined {
    try {
      return (direction === 1 ? ops.storyEnd(this.document, this.story) : ops.storyStart(this.document, this.story)) || undefined;
    } catch {
      this.story = 'body';
      return (direction === 1 ? ops.storyEnd(this.document, 'body') : ops.storyStart(this.document, 'body')) || undefined;
    }
  }

  /**
   * Runs a command on the selection and applies what it returns as one step
   * of the history. `seal`: the step is not joined with typing before or
   * after it.
   */
  private perform(make: (range: DocRange) => DocEdit, seal = false): boolean {
    if (!this.editable || this.destroyed) return false;
    this.typing.commit();
    const range = this.range();
    if (!range) return false;
    if (seal) this.history.seal();
    let edit: DocEdit;
    try {
      edit = make(range);
      if (edit.ops.length > 0) this.history.applyEdit(edit, { anchor: copyPosition(range.anchor), focus: copyPosition(range.focus) });
    } catch (error) {
      this.refuse(error);
      return false;
    }
    if (seal) this.history.seal();
    if (edit.ops.length === 0) {
      // Nothing to change, but the command may have moved the caret.
      this.adopt(edit.selection);
      this.selectionMoved(true);
      return false;
    }
    this.changed(edit.selection);
    return true;
  }

  private adopt(selection: DocRange): void {
    this.selection = { anchor: copyPosition(selection.anchor), focus: copyPosition(selection.focus) };
    this.story = selection.focus.story;
  }

  /** The document was changed: the pages follow, the caret goes where the change left it. */
  private changed(selection: DocRange | undefined): void {
    this.modified = true;
    this.changes++;
    if (selection) this.adopt(selection);
    this.affinity = 'start';
    this.targetX = undefined;
    this.follow();
    this.selectionMoved(true);
    if (this.options.onChange) this.options.onChange();
  }

  private afterHistory(meta: DocRange | undefined): void {
    this.pendingFormat = undefined;
    let selection: DocRange | undefined;
    if (meta) {
      try {
        ops.resolvePosition(this.document, meta.anchor);
        ops.resolvePosition(this.document, meta.focus);
        selection = meta;
      } catch {
        selection = undefined;
      }
    }
    if (!selection) {
      const start = this.storyEdge(-1);
      if (start) selection = ops.collapsed(start);
    }
    this.changed(selection);
  }

  /** The caret or the selection moved: the screen, the field for typing and the host are told. */
  private selectionMoved(reveal: boolean): void {
    const range = this.range();
    const wide = range !== undefined && !ops.samePosition(range.anchor, range.focus);
    // Only a selection is painted on the pages; a caret is an element of its own.
    if (wide || this.hadRange) this.repaintAll();
    this.hadRange = wide;
    if (!wide) this.hideTouchTools();
    this.escaped = false;
    this.blink = !this.blink;
    this.typing.sync();
    if (reveal) this.reveal();
    this.schedule();
    if (this.options.onSelectionChange) this.options.onSelectionChange(this.getSelection());
  }

  // ------------------------------------------------------------------ typing

  /** What the field for typing mirrors: the text of the paragraph around the caret. */
  private inputContext(): InputContext | undefined {
    const range = this.range();
    if (!this.editable || !range) return undefined;
    try {
      let start = range.anchor;
      let end = range.focus;
      const wide = !ops.samePosition(start, end);
      if (wide) {
        const ordered = ops.orderRange(range);
        start = ordered.start;
        end = ordered.end;
      }
      const first = paragraphText(ops.resolvePosition(this.document, start).paragraph);
      const last = wide ? paragraphText(ops.resolvePosition(this.document, end).paragraph) : first;
      let from = Math.max(0, start.offset - CONTEXT_BEFORE);
      if (from > 0 && isLowSurrogate(first.charCodeAt(from))) from++;
      let to = Math.min(last.length, end.offset + CONTEXT_AFTER);
      if (to < last.length && isLowSurrogate(last.charCodeAt(to))) to--;
      // The selected text itself is not mirrored: one character stands for all of it.
      return { before: first.slice(from, start.offset), after: last.slice(end.offset, Math.max(to, end.offset)), selected: wide ? ' ' : '', atStart: from === 0, atEnd: to === last.length };
    } catch {
      return undefined;
    }
  }

  /** A change the browser made in the field (an input method, a correction) is made in the document. */
  private inputChange(change: InputChange, composition: string | undefined): void {
    const range = this.range();
    if (!this.editable || !range) return;
    const doc = this.document;
    const before = { anchor: copyPosition(range.anchor), focus: copyPosition(range.focus) };
    let selection: DocRange | undefined;
    let any = false;
    const apply = (edit: DocEdit): void => {
      if (edit.ops.length > 0) {
        // The steps of one composition are one step of the history.
        this.history.apply(edit.ops, before, edit.selection, composition !== undefined ? composition : edit.mergeKey);
        if (composition === undefined && edit.seal) this.history.seal();
        any = true;
      }
      selection = edit.selection;
    };
    try {
      if (!ops.isCollapsed(range)) {
        if (change.replaceSelection) apply(change.text === '' ? ops.deleteRange(doc, range) : ops.insertText(doc, range, change.text, this.takeFormat()));
      } else {
        const at = range.focus;
        const length = ops.resolvePosition(doc, at).length;
        let from = at.offset - change.removeBefore;
        let to = at.offset + change.removeAfter;
        // One character more than the paragraph has stands for its boundary.
        const joinBefore = from < 0 && change.text === '';
        const joinAfter = to > length && change.text === '';
        from = Math.max(0, from);
        to = Math.min(length, to);
        const paragraph = ops.resolvePosition(doc, at).paragraph;
        const place = (offset: number): DocPosition => {
          const position: DocPosition = { story: at.story, path: at.path.slice(), offset };
          // Typed at a caret before a picture, the text goes before it: also the text that replaces what was
          // composed there, whose start is where the picture comes to stand once the composed text is gone.
          if (at.before === true && (offset === at.offset || !ops.hasAtomAt(paragraph, offset))) position.before = true;
          return position;
        };
        // One character taken away next to a picture: the picture goes, as with the key, not the letter beyond it.
        const single = change.text === '' && change.removeBefore + change.removeAfter === 1 && ops.hasAtomAt(paragraph, at.offset);
        let picture = false;
        if (single && change.removeBefore === 1 && at.before !== true) {
          apply(ops.deleteBackward(doc, at));
          picture = true;
        } else if (single && change.removeAfter === 1 && at.before === true) {
          apply(ops.deleteForward(doc, at));
          picture = true;
        } else if (to > from) {
          const span = { anchor: place(from), focus: place(to) };
          apply(change.text === '' ? ops.deleteRange(doc, span) : ops.insertText(doc, span, change.text, this.takeFormat()));
        } else if (change.text !== '') apply(ops.insertText(doc, place(from), change.text, this.takeFormat()));
        // The picture was the one thing to go: the paragraph is not joined with its neighbour as well.
        if (picture) {
          // Nothing more.
        } else if (joinBefore) apply(ops.deleteBackward(doc, place(0)));
        else if (joinAfter) apply(ops.deleteForward(doc, selection ? (selection as DocRange).focus : place(length)));
      }
    } catch (error) {
      this.refuse(error);
    }
    let caret = selection as DocRange | undefined;
    // An input method may leave its caret inside what it is composing.
    if (caret && change.caretBack > 0 && change.text.indexOf('\n') < 0 && caret.focus.offset >= change.caretBack) {
      caret = ops.collapsed({ story: caret.focus.story, path: caret.focus.path.slice(), offset: caret.focus.offset - change.caretBack });
    }
    this.composed = composition !== undefined ? Math.max(0, change.composing - change.caretBack) : 0;
    if (any) this.changed(caret);
    else if (caret) {
      this.adopt(caret);
      this.selectionMoved(true);
    }
  }

  private typeText(text: string): boolean {
    // Typing is joined into one step of the history by the command itself.
    return this.perform((range) => ops.insertText(this.document, range, text, this.takeFormat()));
  }

  private removeText(direction: 1 | -1, unit: 'char' | 'word' | 'line'): void {
    this.perform((range) => {
      const doc = this.document;
      if (!ops.isCollapsed(range)) return ops.deleteRange(doc, range);
      if (unit === 'line') {
        // To the start or the end of the line the caret is on.
        const at = this.lines.lineOf(range.focus, this.affinity, this.caretPage);
        const edge = at ? (direction === -1 ? at.line.mark.source.offset : lineEnd(at.line)) : range.focus.offset;
        if (edge === range.focus.offset) return direction === -1 ? ops.deleteBackward(doc, range.focus) : ops.deleteForward(doc, range.focus);
        return ops.deleteRange(doc, { anchor: range.focus, focus: { story: range.focus.story, path: range.focus.path.slice(), offset: edge } });
      }
      return direction === -1 ? ops.deleteBackward(doc, range.focus, unit) : ops.deleteForward(doc, range.focus, unit);
    });
  }

  /** Enter (a new paragraph), Shift+Enter (a line break) and Ctrl+Enter (a page break). */
  private enter(kind: 'paragraph' | 'line' | 'page'): void {
    this.perform((range) => {
      const doc = this.document;
      if (kind === 'line') return ops.insertText(doc, range, '\v', this.takeFormat());
      if (kind === 'page') return ops.insertPageBreak(doc, this.caretOf(range));
      return ops.isCollapsed(range) ? ops.splitParagraph(doc, range.focus) : ops.insertText(doc, range, '\n');
    }, true);
  }

  /** Tab: the next cell in a table (a new row after the last one), a deeper level in a list, a tab character elsewhere. */
  private tab(back: boolean): void {
    const range = this.range();
    if (!range) return;
    const doc = this.document;
    const path = range.focus.path;
    const story = range.focus.story;
    if (this.inTable) {
      try {
        const tablePath = path.slice(0, -3);
        const table = ops.containerBlocks(doc, story, tablePath.slice(0, -1))[tablePath[tablePath.length - 1]];
        if (!table || table.type !== 'table') return;
        let row = path[path.length - 3];
        let cell = path[path.length - 2] + (back ? -1 : 1);
        if (cell < 0) {
          row--;
          if (row < 0) return;
          cell = table.rows[row].cells.length - 1;
        } else if (cell >= table.rows[row].cells.length) {
          row++;
          cell = 0;
          if (row >= table.rows.length) {
            // Past the last cell: the table gets a row, as in Word, and the caret goes into its first cell.
            if (!this.perform((selection) => ops.insertRow(doc, selection.focus, 'below'), true)) return;
            const grown = ops.containerBlocks(doc, story, tablePath.slice(0, -1))[tablePath[tablePath.length - 1]];
            const cell = grown && grown.type === 'table' && grown.rows[row] ? grown.rows[row].cells[0] : undefined;
            const first = cell ? ops.paragraphIn(cell.blocks, tablePath.concat([row, 0]), 0, 1) : undefined;
            if (first) this.setSelection({ story, path: first, offset: 0 });
            return;
          }
        }
        const blocks = table.rows[row].cells[cell].blocks;
        const base = tablePath.concat([row, cell]);
        const first = ops.paragraphIn(blocks, base, 0, 1);
        const last = ops.paragraphIn(blocks, base, blocks.length - 1, -1);
        if (!first || !last) return;
        // The content of the cell is selected: typing replaces it.
        const end = ops.resolvePosition(doc, { story, path: last, offset: 0 }).length;
        this.setSelection({ anchor: { story, path: first, offset: 0 }, focus: { story, path: last, offset: end } });
      } catch (error) {
        this.refuse(error);
      }
      return;
    }
    const format = this.selectionFormat;
    if (format && format.list !== undefined && format.list !== 'none') this.indent(back ? -1 : 1);
    else if (!back) this.typeText('\t');
  }

  // --------------------------------------------------------------- the caret

  /** Shows the caret where it is, when the view has the focus, and keeps the field for typing beside it. */
  private placeCaret(): void {
    const range = this.range();
    const caret = this.caret;
    const underline = this.underline;
    let box: CaretBox | undefined;
    if (this.editable && range) box = this.lines.caret(range.focus, this.affinity, this.caretPage);
    if (box) this.caretPage = box.page;
    const slot = box ? this.slots.get(box.page) : undefined;
    const shown = box !== undefined && slot !== undefined && this.focused && range !== undefined && ops.samePosition(range.anchor, range.focus);
    let composing = false;
    if (shown && box && slot) {
      const width = this.pageWidth(box.page);
      const height = this.pageHeight(box.page);
      caret.style.left = `${(box.x / width) * 100}%`;
      caret.style.top = `${(box.top / height) * 100}%`;
      caret.style.height = `${(box.height / height) * 100}%`;
      // Another animation with the same frames: the caret shows at once after every move.
      const name = this.blink ? 'bym-doc-caret bym-doc-caret-b' : 'bym-doc-caret';
      if (caret.className !== name) caret.className = name;
      if (caret.parentNode !== slot.element) slot.element.appendChild(caret);
      if (this.composed > 0 && range && range.focus.offset >= this.composed) {
        // What an input method is composing is underlined, as far as it is on the line of the caret.
        const start = this.lines.caret({ story: range.focus.story, path: range.focus.path, offset: range.focus.offset - this.composed }, 'start', box.page);
        if (start && start.page === box.page && Math.abs(start.top - box.top) < 1 && box.x > start.x) {
          underline.style.left = `${(start.x / width) * 100}%`;
          underline.style.width = `${((box.x - start.x) / width) * 100}%`;
          underline.style.top = `${((box.top + box.height) / height) * 100}%`;
          if (underline.parentNode !== slot.element) slot.element.appendChild(underline);
          composing = true;
        }
      }
    } else if (caret.parentNode) caret.parentNode.removeChild(caret);
    if (!composing && underline.parentNode) underline.parentNode.removeChild(underline);
    if (box && slot && this.editable) this.placeInput(box, slot);
    this.placeTouchTools();
  }

  /** A point of a page in px of the view's root, whatever transforms lie on the view. */
  private rootPoint(slot: Slot, x: number, y: number): { x: number; y: number } {
    const root = this.root.getBoundingClientRect();
    const page = slot.element.getBoundingClientRect();
    const sx = root.width > 0 && this.root.offsetWidth > 0 ? this.root.offsetWidth / root.width : 1;
    const sy = root.height > 0 && this.root.offsetHeight > 0 ? this.root.offsetHeight / root.height : 1;
    return {
      x: (page.left - root.left + (x / this.pageWidth(slot.index)) * page.width) * sx,
      y: (page.top - root.top + (y / this.pageHeight(slot.index)) * page.height) * sy,
    };
  }

  /** The field for typing sits at the caret: an input method opens its window there, and a phone scrolls there for its keyboard. */
  private placeInput(box: CaretBox, slot: Slot): void {
    const key = `${box.page}:${box.x}:${box.top}:${slot.left}:${slot.top}:${this.scale}:${this.scroller.scrollLeft}:${this.scroller.scrollTop}`;
    if (key === this.inputPlace) return;
    this.inputPlace = key;
    const point = this.rootPoint(slot, box.x, box.top);
    const style = this.input.style;
    style.left = `${clamp(point.x, 0, Math.max(0, this.root.clientWidth - 2))}px`;
    style.top = `${clamp(point.y, 0, Math.max(0, this.root.clientHeight - 2))}px`;
  }

  /** Part of the visible area a keyboard on screen covers, px of the view. */
  private keyboardOverlap(): number {
    const viewport = this.viewport as { offsetTop?: number; height?: number } | undefined;
    if (!viewport || typeof viewport.height !== 'number') return 0;
    const rect = this.scroller.getBoundingClientRect();
    const hidden = rect.bottom - ((viewport.offsetTop || 0) + viewport.height);
    if (!(hidden > 0) || !(rect.height > 0)) return 0;
    return Math.min(Math.max(0, this.height - 40), hidden * (this.scroller.offsetHeight / rect.height));
  }

  /** Scrolls the caret into what is seen of the pages. */
  private reveal(): void {
    const range = this.range();
    if (!range) return;
    const box = this.lines.caret(range.focus, this.affinity, this.caretPage);
    if (!box) return;
    this.caretPage = box.page;
    this.sizeStrip();
    const top = this.pageTop(box.page) + box.top * this.scale;
    const bottom = top + box.height * this.scale;
    const visible = Math.max(40, this.height - this.keyboardOverlap());
    const y = this.virtualY();
    if (top - CARET_MARGIN < y) this.setVirtualY(top - CARET_MARGIN);
    else if (bottom + CARET_MARGIN > y + visible) this.setVirtualY(bottom + CARET_MARGIN - visible);
    const width = this.pageWidth(box.page) * this.scale;
    const x = Math.max(PAD, (this.stripWidth() - width) / 2) + box.x * this.scale;
    const left = this.scroller.scrollLeft;
    if (x - CARET_MARGIN < left) this.scroller.scrollLeft = Math.max(0, x - CARET_MARGIN);
    else if (x + CARET_MARGIN > left + this.width) this.scroller.scrollLeft = x + CARET_MARGIN - this.width;
  }

  /** A place found on a page as a position of the model: inside its paragraph, not inside a character, and before a line break that ends its line. */
  private settle(place: DocPosition, affinity: Affinity): { position: DocPosition; affinity: Affinity } | undefined {
    try {
      const at = ops.resolvePosition(this.document, { story: place.story, path: place.path, offset: 0 });
      const text = paragraphText(at.paragraph);
      let offset = ops.snapOffset(text, clamp(place.offset, 0, at.length));
      let side = affinity;
      if (side === 'end' && offset > 0 && text.charCodeAt(offset - 1) === 0x0a) {
        offset--;
        side = 'start';
      }
      const position: DocPosition = { story: place.story, path: place.path.slice(), offset };
      if (place.before === true && offset === place.offset) position.before = true;
      return { position, affinity: side };
    } catch {
      return undefined;
    }
  }

  /** The place of a story nearest to a point of the window. */
  private placeAt(point: { clientX: number; clientY: number }, story: string): { position: DocPosition; affinity: Affinity; page: number; box: boolean } | undefined {
    const at = this.pagePoint(point.clientX, point.clientY);
    if (!at) return undefined;
    const hit = this.lines.hit(at.page, at.x, at.y, (mark) => mark.source.story === story && outsideBoxes(mark));
    const settled = hit ? this.settle(hit.place, hit.affinity) : undefined;
    return settled && hit ? { position: settled.position, affinity: settled.affinity, page: at.page, box: hit.box === true } : undefined;
  }

  /**
   * The story a double click is in: the one whose line is under the pointer.
   * That is how the caret gets into a header, a footer or a note, and back
   * into the text.
   */
  private storyAt(point: { clientX: number; clientY: number }): string {
    const at = this.pagePoint(point.clientX, point.clientY);
    const line = at ? this.lines.lineAt(at.page, at.x, at.y, outsideBoxes) : undefined;
    if (!at || !line) return this.story;
    const mark = line.mark;
    return at.y >= mark.y - 4 && at.y <= mark.y + mark.height + 4 ? mark.source.story : this.story;
  }

  private enterStory(story: string): void {
    if (story === this.story) return;
    this.story = story;
    this.live.textContent = this.strings.editingStory(storyKind(story));
  }

  private editPointerDown(event: PointerEvent | MouseEvent): void {
    const now = Date.now();
    const down = this.lastDown;
    const again = now - down.time < CLICK_MS && Math.abs(event.clientX - down.x) < CLICK_DISTANCE && Math.abs(event.clientY - down.y) < CLICK_DISTANCE;
    // A fourth press in a row still selects the paragraph.
    const count = again ? Math.min(down.count + 1, 3) : 1;
    this.lastDown = { time: now, x: event.clientX, y: event.clientY, count };
    this.typing.commit();
    this.history.seal();
    this.pendingFormat = undefined;
    this.touchSelection = false;
    let story = count === 2 ? this.storyAt(event) : this.story;
    let hit = this.placeAt(event, story);
    if (!hit && story !== 'body') {
      // The page shows nothing of that story: back to the text.
      story = 'body';
      hit = this.placeAt(event, story);
    }
    if (!hit) return;
    this.enterStory(story);
    const range = this.range();
    this.caretPage = hit.page;
    this.targetX = undefined;
    if (count === 2) this.selectWord(hit.position);
    else if (count === 3) {
      const length = ops.resolvePosition(this.document, hit.position).length;
      this.selection = { anchor: { ...hit.position, offset: 0 }, focus: { ...hit.position, path: hit.position.path.slice(), offset: length } };
      this.affinity = 'start';
    } else if (event.shiftKey && range && range.anchor.story === story) {
      this.selection = { anchor: range.anchor, focus: hit.position };
      this.affinity = hit.affinity;
    } else if (hit.box) {
      // A press on a picture selects it: from before it to after it. Typing replaces it, Delete removes it.
      const after = copyPosition(hit.position);
      delete after.before;
      this.selection = { anchor: { ...after, path: after.path.slice(), before: true }, focus: after };
      this.affinity = 'start';
      this.selectionMoved(false);
      return;
    } else {
      this.selection = { anchor: hit.position, focus: copyPosition(hit.position) };
      this.affinity = hit.affinity;
    }
    this.selectionMoved(false);
    if (count !== 1) return;
    // A second press without the end of the first one (a pointer that was lost) must not leave a timer behind.
    if (this.dragTimer !== undefined) clearInterval(this.dragTimer);
    this.dragging = true;
    window.addEventListener(this.pointerEvents ? 'pointermove' : 'mousemove', this.onPointerMove as EventListener);
    window.addEventListener(this.pointerEvents ? 'pointerup' : 'mouseup', this.onPointerUp);
    if (this.pointerEvents) window.addEventListener('pointercancel', this.onPointerUp);
    // While the pointer is held near the top or the bottom, the pages scroll and the selection follows.
    this.dragTimer = setInterval(this.onDragTick, 50);
  }

  private selectWord(position: DocPosition): void {
    const text = paragraphText(ops.resolvePosition(this.document, position).paragraph);
    const [start, end] = ops.wordAt(text, position.offset);
    this.selection = { anchor: { ...position, offset: start }, focus: { story: position.story, path: position.path.slice(), offset: end } };
    this.affinity = 'start';
  }

  /** Selects the word at a point: a double tap, or (`tools`) a long press, which also brings the handles and the bubble. */
  private selectWordAt(point: { clientX: number; clientY: number }, tools: boolean): void {
    const story = this.storyAt(point);
    const hit = this.placeAt(point, story) || this.placeAt(point, 'body');
    if (!hit) return;
    this.typing.commit();
    this.history.seal();
    this.enterStory(hit.position.story);
    this.caretPage = hit.page;
    this.selectWord(hit.position);
    this.touchSelection = tools;
    this.selectionMoved(false);
  }

  /** A tap puts the caret there and, in the field that now has the focus, brings the keyboard. */
  private tapAt(point: { clientX: number; clientY: number }): void {
    const hit = this.placeAt(point, this.story) || this.placeAt(point, 'body');
    if (!hit) return;
    this.typing.commit();
    this.history.seal();
    this.pendingFormat = undefined;
    this.enterStory(hit.position.story);
    this.caretPage = hit.page;
    this.targetX = undefined;
    this.touchSelection = false;
    this.selection = { anchor: hit.position, focus: copyPosition(hit.position) };
    this.affinity = hit.affinity;
    this.focusInput(true);
    this.selectionMoved(true);
  }

  private dragTo(point: { clientX: number; clientY: number }): void {
    const range = this.range();
    if (!range) return;
    const hit = this.placeAt(point, range.anchor.story);
    if (!hit) return;
    if (!ops.samePosition(range.anchor, hit.position)) this.dragMoved = true;
    if (ops.samePosition(range.focus, hit.position) && this.affinity === hit.affinity) return;
    this.selection = { anchor: range.anchor, focus: hit.position };
    this.affinity = hit.affinity;
    this.caretPage = hit.page;
    this.selectionMoved(false);
  }

  private onDragTick = (): void => {
    const point = this.dragPoint;
    if (!this.dragging || !point || this.destroyed) return;
    const rect = this.scroller.getBoundingClientRect();
    const above = rect.top + DRAG_SCROLL_EDGE - point.clientY;
    const below = point.clientY - (rect.bottom - DRAG_SCROLL_EDGE);
    if (!(above > 0) && !(below > 0)) return;
    // Faster the further out the pointer is.
    const step = clamp((above > 0 ? above : below) * 0.6, 6, 80);
    this.setVirtualY(this.virtualY() + (above > 0 ? -step : step));
    this.update();
    this.dragTo(point);
  };

  // ------------------------------------------------------------ touch tools

  private hideTouchTools(): void {
    this.touchSelection = false;
    this.handleFixed = undefined;
  }

  /** The handles at the ends of a selection made by touch, and the bubble with what can be done with it. */
  private placeTouchTools(): void {
    const range = this.range();
    const shown = this.editable && this.touchSelection && range !== undefined && !ops.samePosition(range.anchor, range.focus);
    if (!shown || !range) {
      for (const handle of this.handles) if (handle.parentNode) handle.parentNode.removeChild(handle);
      if (this.bubble && this.bubble.parentNode) this.bubble.parentNode.removeChild(this.bubble);
      return;
    }
    if (this.handles.length === 0) this.buildTouchTools();
    const ordered = ops.orderRange(range);
    const ends = [ordered.start, ordered.end];
    let top: { x: number; y: number } | undefined;
    for (let i = 0; i < 2; i++) {
      const handle = this.handles[i];
      const box = this.lines.caret(ends[i], i === 0 ? 'start' : 'end', this.caretPage);
      const slot = box ? this.slots.get(box.page) : undefined;
      if (!box || !slot) {
        if (handle.parentNode) handle.parentNode.removeChild(handle);
        continue;
      }
      // In the strip, not in the page: an element that is moved to another parent while a finger is on it loses the finger.
      const x = slot.left + box.x * this.scale;
      const y = slot.top + (box.top + box.height) * this.scale;
      handle.style.left = `${x - (i === 0 ? 22 : 0)}px`;
      handle.style.top = `${y}px`;
      if (handle.parentNode !== this.strip) this.strip.appendChild(handle);
      if (i === 0) top = this.rootPoint(slot, box.x, box.top);
    }
    const bubble = this.bubble;
    if (!bubble) return;
    if (!top || this.handleFixed) {
      if (bubble.parentNode) bubble.parentNode.removeChild(bubble);
      return;
    }
    if (bubble.parentNode !== this.root) this.root.appendChild(bubble);
    const width = bubble.offsetWidth;
    const height = bubble.offsetHeight;
    bubble.style.left = `${clamp(top.x - width / 2, 4, Math.max(4, this.root.clientWidth - width - 4))}px`;
    bubble.style.top = `${clamp(top.y - height - 10, 4, Math.max(4, this.root.clientHeight - height - 4))}px`;
  }

  private buildTouchTools(): void {
    for (let i = 0; i < 2; i++) {
      const handle = document.createElement('div');
      handle.className = i === 0 ? 'bym-doc-handle bym-doc-handle-start' : 'bym-doc-handle';
      handle.setAttribute('aria-hidden', 'true');
      handle.style.cssText = 'position:absolute;z-index:3;width:22px;height:22px;margin:0;padding:0;border:0;box-sizing:border-box';
      const start = (event: Event) => {
        const range = this.range();
        if (!range) return;
        if (event.cancelable) event.preventDefault();
        event.stopPropagation();
        const ordered = ops.orderRange(range);
        // The other end stays where it is while this one is moved.
        this.handleFixed = copyPosition(i === 0 ? ordered.end : ordered.start);
      };
      const move = (event: TouchEvent) => {
        const fixed = this.handleFixed;
        const touch = event.touches[0];
        if (!fixed || !touch) return;
        if (event.cancelable) event.preventDefault();
        // The finger is below the line it points at.
        const hit = this.placeAt({ clientX: touch.clientX, clientY: touch.clientY - 22 }, fixed.story);
        if (!hit || ops.samePosition(hit.position, fixed)) return;
        this.selection = { anchor: fixed, focus: hit.position };
        this.caretPage = hit.page;
        this.touchSelection = true;
        this.selectionMoved(false);
        this.touchSelection = true;
      };
      const end = () => {
        this.handleFixed = undefined;
        this.schedule();
      };
      handle.addEventListener('touchstart', start, { passive: false });
      handle.addEventListener('touchmove', move as EventListener, { passive: false });
      handle.addEventListener('touchend', end);
      handle.addEventListener('touchcancel', end);
      // Nothing under the handle takes the press.
      handle.addEventListener('mousedown', (event) => {
        event.preventDefault();
        event.stopPropagation();
      });
      this.handles.push(handle);
    }
    const bubble = document.createElement('div');
    bubble.className = 'bym-doc-bubble';
    bubble.setAttribute('role', 'toolbar');
    bubble.setAttribute('aria-label', this.strings.selectionTools);
    bubble.style.cssText = 'position:absolute;z-index:4;display:flex;box-sizing:border-box;padding:2px;white-space:nowrap';
    const add = (name: string, label: string, action: () => void) => {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'bym-doc-btn bym-doc-bubble-' + name;
      button.textContent = label;
      button.addEventListener('mousedown', (event) => event.preventDefault());
      button.addEventListener('click', (event) => {
        event.stopPropagation();
        if (!this.destroyed) action();
      });
      bubble.appendChild(button);
    };
    const command = (name: 'copy' | 'cut') => () => {
      this.focusInput(true);
      this.typing.sync();
      let done = false;
      try {
        // In the press of a button: the copy and cut events of the view do the work.
        done = document.execCommand(name);
      } catch {
        done = false;
      }
      if (!done) this.notify(this.strings.pasteWithKeyboard, 'info');
      this.touchSelection = false;
      this.schedule();
    };
    add('cut', this.strings.cut, command('cut'));
    add('copy', this.strings.copy, command('copy'));
    add('paste', this.strings.paste, () => {
      interface Item {
        types: readonly string[];
        getType(type: string): Promise<Blob>;
      }
      const clipboard = (navigator as { clipboard?: { readText?(): Promise<string>; read?(): Promise<Item[]> } }).clipboard;
      const failed = () => {
        if (!this.destroyed) this.notify(this.strings.pasteWithKeyboard, 'info');
      };
      const plain = () => {
        if (!clipboard || typeof clipboard.readText !== 'function') return failed();
        clipboard.readText().then((text) => {
          if (this.destroyed) return;
          // Without text there may be a copy of a picture the page keeps: only the keys can tell.
          if (text === '') failed();
          else this.pasteText(text);
        }, failed);
      };
      if (!clipboard || typeof clipboard.read !== 'function' || typeof Blob === 'undefined' || typeof Blob.prototype.text !== 'function') return plain();
      // With its HTML, so that a copy made on this page is known by its mark.
      const part = (items: Item[], type: string): Promise<string> => {
        for (const item of items) if (item.types.indexOf(type) >= 0) return item.getType(type).then((blob) => blob.text());
        return Promise.resolve('');
      };
      clipboard
        .read()
        .then((items) => Promise.all([part(items, 'text/plain'), part(items, 'text/html')]))
        .then(([text, html]) => {
          if (this.destroyed) return;
          if (text === '' && !storedCopy(html, text)) failed();
          else this.pasteClipboard(text, html, undefined);
        }, plain);
    });
    add('all', this.strings.selectWholeText, () => {
      this.selectAll();
      this.touchSelection = true;
      this.schedule();
    });
    this.bubble = bubble;
  }

  // ------------------------------------------------------------ the clipboard

  /** The selection as the editor keeps a copy of it: its content, and its text as formatted paragraphs or as the cells of a table. */
  private selectionSlice(): DocSlice | undefined {
    const selection = this.selection;
    if (selection === undefined) return undefined;
    try {
      if (selection === 'all') {
        const start = ops.storyStart(this.document, 'body');
        const end = ops.storyEnd(this.document, 'body');
        return start && end ? sliceOf(this.document, start, end) : undefined;
      }
      // What a deletion would remove is what the copy holds: whole cells where the selection goes into or out of a table.
      const shown = this.shownRange() || selection;
      const range = orderedRange(shown.anchor, shown.focus);
      if (!range) return undefined;
      if (range.cells) {
        const blocks = flowBlocks(this.document, range.start.story, range.start.path.slice(0, range.prefix));
        const table = blocks ? tableAt(blocks, range.cells.table.slice(range.prefix)) : undefined;
        if (!table) return undefined;
        const rows: string[][] = [];
        for (let r = range.cells.row; r <= range.cells.lastRow && r < table.rows.length; r++) {
          const row: string[] = [];
          for (let c = range.cells.cell; c <= range.cells.lastCell && c < table.rows[r].cells.length; c++) {
            const inner: string[] = [];
            blocksText(table.rows[r].cells[c].blocks, inner, 1);
            row.push(inner.join('\n'));
          }
          rows.push(row);
        }
        const slice: DocSlice = { paragraphs: [], table: rows };
        const fragment = fragmentOfRange(this.document, range.start, range.end);
        if (fragment) slice.fragment = fragment;
        return slice;
      }
      return sliceOf(this.document, range.start, range.end);
    } catch {
      return undefined;
    }
  }

  private pasteText(text: string): void {
    if (typeof text !== 'string' || text === '') return;
    if (text.length > MAX_PASTE_TEXT) {
      this.notify(this.strings.documentPasteTooLarge, 'error');
      return;
    }
    this.perform((range) => ops.insertText(this.document, range, text), true);
  }

  /**
   * Pastes a copy made on this page as one step of the history: with its
   * pictures, links, fields and tables when it holds them, else its text with
   * the formatting it had. A copy whose pictures are still being read from
   * the document it came from is pasted once they are there.
   */
  private pasteCopy(copy: StoredCopy, waited: boolean): void {
    if (this.trackingChanges) {
      // Formatting, pictures and tables are not made tracked insertions yet: the text of the copy is, and the user is told.
      if (!this.editable || !this.range()) return;
      if (copy.text === '') {
        // A copy without text (a picture alone): nothing of it can be pasted as a tracked insertion, and that is said.
        this.notify(this.strings.cannotTrackChange, 'info');
        return;
      }
      this.pasteText(copy.text);
      this.notify(this.strings.pastedAsText, 'info');
      return;
    }
    const asked = this.range();
    const fragment = copy.slice.fragment;
    if (!fragment) {
      this.pasteSlice(copy.slice);
      return;
    }
    if (!this.editable || !this.range()) return;
    if (fragment.textLength > MAX_PASTE_TEXT) {
      this.notify(this.strings.documentPasteTooLarge, 'error');
      return;
    }
    let needs: string[];
    try {
      needs = ops.fragmentNeeds(this.document, fragment);
    } catch (error) {
      this.refuse(error);
      return;
    }
    if (needs.length) {
      if (!waited && copy.ready && asked) {
        // The paste is for this place in this state of the document: if either is another one when the pictures are
        // there, pasting would put the copy where nobody asked for it, or over a selection made since.
        const at = { anchor: copyPosition(asked.anchor), focus: copyPosition(asked.focus) };
        const changes = this.changes;
        this.notify(this.strings.pasteWaiting, 'info');
        copy.ready.then(() => {
          if (this.destroyed) return;
          const now = this.range();
          let same = this.editable && this.changes === changes && now !== undefined;
          try {
            same = same && now !== undefined && ops.samePosition(now.anchor, at.anchor) && ops.samePosition(now.focus, at.focus);
          } catch {
            same = false;
          }
          if (same) this.pasteCopy(copy, true);
          else this.notify(this.strings.pasteInterrupted, 'info');
        });
      } else this.notify(this.strings.cannotPasteHere, 'info');
      return;
    }
    // The name of a new part must not be one the package has already, though the model does not know it.
    const source = this.options.package as { hasPart?(part: string): boolean } | undefined;
    const taken = this.options.partTaken || (source && typeof source.hasPart === 'function' ? (part: string) => (source.hasPart as (part: string) => boolean)(part) : undefined);
    this.typing.commit();
    const range = this.range();
    if (!range) return;
    this.history.seal();
    let edit: DocEdit;
    try {
      const options: ops.InsertFragmentOptions = {};
      if (taken) options.taken = taken;
      // A cut that comes back to where it was gives the paragraph it ended in its formatting again.
      if (copy.cut) options.moved = true;
      edit = ops.insertFragment(this.document, range, fragment, options);
      if (edit.ops.length > 0) this.history.applyEdit(edit, { anchor: copyPosition(range.anchor), focus: copyPosition(range.focus) });
    } catch (error) {
      // A copy that does not fit here (another kind of document, tables too deep) is not an error of the program.
      if (!(error instanceof RangeError)) throw error;
      this.notify(this.strings.cannotPasteHere, 'info');
      return;
    }
    this.history.seal();
    if (edit.ops.length > 0) this.changed(edit.selection);
  }

  /** Pastes the text of a copy with the formatting it had, as one step of the history. */
  private pasteSlice(slice: DocSlice): void {
    const range = this.range();
    if (!this.editable || !range) return;
    const doc = this.document;
    const before = { anchor: copyPosition(range.anchor), focus: copyPosition(range.focus) };
    const key = 'paste:' + ++this.steps;
    this.history.seal();
    let at: DocRange = range;
    let any = false;
    const apply = (edit: DocEdit): void => {
      if (edit.ops.length > 0) {
        this.history.apply(edit.ops, before, edit.selection, key);
        any = true;
      }
      at = edit.selection;
    };
    try {
      if (!ops.isCollapsed(at)) apply(ops.deleteRange(doc, at));
      if (slice.table) {
        const rows = slice.table.length;
        let columns = 1;
        for (const row of slice.table) columns = Math.max(columns, row.length);
        apply(ops.insertTable(doc, at.focus, rows, columns));
        // The caret is in the first cell of the new table.
        const first = at.focus;
        const table = first.path.slice(0, -3);
        for (let r = 0; r < rows; r++) {
          for (let c = 0; c < slice.table[r].length; c++) {
            if (slice.table[r][c] !== '') apply(ops.insertText(doc, { story: first.story, path: table.concat([r, c, 0]), offset: 0 }, slice.table[r][c]));
          }
        }
      } else {
        const ended: { path: number[]; align?: string; style?: string }[] = [];
        for (let i = 0; i < slice.paragraphs.length; i++) {
          const paragraph = slice.paragraphs[i];
          for (const run of paragraph.runs) if (run.text !== '') apply(ops.insertText(doc, at.focus, run.text, run.format));
          if (i < slice.paragraphs.length - 1) {
            ended.push({ path: at.focus.path.slice(), align: paragraph.align, style: paragraph.style });
            apply(ops.insertText(doc, at.focus, '\n'));
          }
        }
        const caret = at;
        // The paragraphs whose end was copied get the alignment and the style they had.
        for (const entry of ended) {
          const patch: ParaFormatPatch = {};
          if (entry.align !== undefined) patch.align = entry.align;
          if (entry.style !== undefined && this.document.styles.styles.has(entry.style)) patch.style = entry.style;
          if (patch.align !== undefined || patch.style !== undefined) apply(ops.setParaFormat(doc, { story: caret.focus.story, path: entry.path, offset: 0 }, patch));
        }
        at = caret;
      }
    } catch (error) {
      this.refuse(error);
    }
    this.history.seal();
    if (any) this.changed(ops.collapsed(at.focus));
  }

  /** A picture among what is on the clipboard: PNG or JPEG. */
  private clipboardPicture(data: DataTransfer): Blob | undefined {
    const wanted = /^image\/(png|jpeg)$/;
    const items = data.items;
    if (items) {
      for (let i = 0; i < items.length; i++) {
        if (items[i].kind !== 'file' || !wanted.test(items[i].type)) continue;
        const file = items[i].getAsFile();
        if (file) return file;
      }
    }
    const files = data.files;
    if (files) for (let i = 0; i < files.length; i++) if (wanted.test(files[i].type)) return files[i];
    return undefined;
  }

  private readPicture(file: Blob): void {
    if (file.size > MAX_PICTURE_BYTES) {
      this.notify(this.strings.pictureNotSupported, 'error');
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      // What the bytes are is found out from the bytes, not from what the clipboard says.
      if (!this.destroyed && reader.result instanceof ArrayBuffer) this.insertPicture(new Uint8Array(reader.result));
    };
    reader.onerror = () => {
      if (!this.destroyed) this.notify(this.strings.pictureNotSupported, 'error');
    };
    reader.readAsArrayBuffer(file);
  }

  /** Room a picture has where a position is, px: the text area of the page, or a share of it inside a table. */
  private textRoom(position: DocPosition): { width: number; height: number } {
    const doc = this.document;
    let section = doc.body.section;
    if (position.story === 'body') {
      const blocks = doc.body.blocks;
      for (let i = position.path[0]; i >= 0 && i < blocks.length; i++) {
        const block = blocks[i];
        if (block.type === 'paragraph' && block.props.section) {
          section = block.props.section;
          break;
        }
      }
    }
    const page = (section && section.page) || {};
    const margins = (section && section.margins) || {};
    const px = (value: number | undefined, fallback: number) => (typeof value === 'number' && value > 0 && value < 500000 ? value : fallback) / 15;
    let width = px(page.width, 12240) - px(margins.left, 1440) - px(margins.right, 1440);
    const height = px(page.height, 15840) - px(margins.top, 1440) - px(margins.bottom, 1440);
    if (position.path.length >= 4) {
      try {
        const tablePath = position.path.slice(0, -3);
        const table = ops.containerBlocks(doc, position.story, tablePath.slice(0, -1))[tablePath[tablePath.length - 1]];
        if (table && table.type === 'table') width /= Math.max(1, table.rows[position.path[position.path.length - 3]].cells.length);
      } catch {
        // The whole width then.
      }
    }
    return { width: clamp(width, 24, 2000), height: clamp(height * 0.9, 24, 3000) };
  }

  // ---------------------------------------------------------- keys of an editor

  /** Moves the focus end of the selection to a place; without `extend` the caret is put there. */
  private moveTo(position: DocPosition, affinity: Affinity, extend: boolean, keepX: boolean): void {
    const range = this.range();
    this.history.seal();
    this.pendingFormat = undefined;
    this.touchSelection = false;
    this.selection = { anchor: extend && range ? range.anchor : copyPosition(position), focus: position };
    this.affinity = affinity;
    if (!keepX) this.targetX = undefined;
    this.selectionMoved(true);
  }

  /** One line up or down, or a screen (`pages`), keeping the x the caret had when the moving began. */
  private moveVertically(direction: 1 | -1, extend: boolean, screen: boolean): void {
    const range = this.range();
    if (!range) return;
    const box = this.lines.caret(range.focus, this.affinity, this.caretPage);
    if (!box) return;
    const x = this.targetX !== undefined ? this.targetX : box.x;
    const accept = (mark: LineMark) => mark.source.story === range.focus.story && outsideBoxes(mark);
    // A line that shows nothing but deleted text has no place of its own: a key never stops on one.
    let found: { place: DocPosition; affinity: Affinity } | undefined;
    if (screen) {
      // The line a screen away, as far as the pages laid out reach.
      const distance = Math.max(LINE_SCROLL, this.height - LINE_SCROLL) / this.scale;
      let y = this.offsets[box.page] + box.top + direction * distance;
      y = clamp(y, 0, Math.max(0, this.offsets[this.pages.length] - 1));
      let page = box.page;
      while (page > 0 && this.offsets[page] > y) page--;
      while (page < this.pages.length - 1 && this.offsets[page + 1] <= y) page++;
      const line = this.lines.lineAt(page, x, y - this.offsets[page], accept);
      if (line && line !== box.line) found = this.lines.ownPlace({ page, line }, x) || this.lines.verticalPlace({ page, line }, direction, x, accept);
    }
    if (!found) found = this.lines.verticalPlace(box, direction, x, accept);
    let target: { position: DocPosition; affinity: Affinity } | undefined;
    if (found) target = this.settle(found.place, found.affinity);
    else {
      // No line further on: the very start or end of the text.
      const edge = this.storyEdge(direction);
      if (edge) target = { position: edge, affinity: 'start' };
    }
    if (!target) return;
    this.targetX = x;
    this.moveTo(target.position, target.affinity, extend, true);
    this.announceCaret('line');
  }

  /** Start (-1) or end (1) of the line the caret is on. */
  private moveToLineEdge(direction: 1 | -1, extend: boolean): void {
    const range = this.range();
    if (!range) return;
    const at = this.lines.lineOf(range.focus, this.affinity, this.caretPage);
    if (!at) return;
    const source = at.line.mark.source;
    const offset = direction === -1 ? source.offset : lineEnd(at.line);
    // On a line a press named (it shows nothing but deleted text) both edges are the place the caret is at.
    const side: Affinity = typeof this.affinity === 'number' && at.line.mark.length === 0 ? this.affinity : direction === 1 && !at.line.mark.last ? 'end' : 'start';
    const settled = this.settle({ story: source.story, path: source.path, offset }, side);
    if (!settled) return;
    this.moveTo(settled.position, settled.affinity, extend, false);
    this.announceCaret('line');
  }

  private moveBy(unit: 'char' | 'word' | 'paragraph' | 'story', direction: 1 | -1, extend: boolean): void {
    const range = this.range();
    if (!range) return;
    let target: DocPosition;
    try {
      if (!extend && unit === 'char' && !ops.samePosition(range.anchor, range.focus)) {
        // Left and right leave a selection at its start and at its end.
        const ordered = ops.orderRange(range);
        target = copyPosition(direction === -1 ? ordered.start : ordered.end);
      } else target = ops.movePosition(this.document, range.focus, unit, direction);
    } catch {
      return;
    }
    this.moveTo(target, 'start', extend, false);
    this.announceCaret(unit === 'char' ? 'char' : unit === 'word' ? 'word' : 'line');
  }

  /** Whether a selection spans more of the body than is looked at as a whole. */
  private isLarge(range: DocRange): boolean {
    return Math.abs(range.anchor.path[0] - range.focus.path[0]) > MAX_LOOKED_AT_BLOCKS;
  }

  /** Says where the caret is now, briefly: the character or the word after it, or its line. */
  private announceCaret(unit: 'char' | 'word' | 'line'): void {
    const range = this.range();
    if (!range) return;
    try {
      if (!ops.samePosition(range.anchor, range.focus)) {
        if (this.isLarge(range)) return;
        const text = this.getSelectedText();
        this.live.textContent = this.strings.selected(text.length > 80 ? text.slice(0, 80) + '…' : text);
        return;
      }
      const at = ops.resolvePosition(this.document, range.focus);
      const text = paragraphText(at.paragraph);
      let said: string;
      if (unit === 'line') {
        const line = this.lines.lineOf(range.focus, this.affinity, this.caretPage);
        said = line ? text.slice(line.line.mark.source.offset, lineEnd(line.line)) : text;
      } else if (at.offset >= text.length) said = '';
      else if (unit === 'word') {
        const [start, end] = ops.wordAt(text, at.offset);
        said = text.slice(Math.max(start, at.offset), end);
      } else said = text.slice(at.offset, ops.nextCharOffset(text, at.offset));
      if (said === ' ' || said === ' ') said = this.strings.space;
      else if (said.replace(/\s+/g, '') === '') said = this.strings.blankLine;
      this.live.textContent = said.length > 200 ? said.slice(0, 200) + '…' : said;
    } catch {
      // Nothing to say.
    }
  }

  private announceSelection(): void {
    if (this.options.onSelectionChange) this.options.onSelectionChange(this.getSelection());
  }

  /** The keys of an editor. True when the key was one of them. */
  private editKey(event: KeyboardEvent): boolean {
    // 229: the key went to an input method, also where isComposing is not reported.
    if (event.isComposing || event.keyCode === 229) return true;
    const key = event.key;
    const apple = isApple();
    // AltGr (Ctrl+Alt) types characters; it is not a command.
    const command = (event.ctrlKey || event.metaKey) && !event.altKey;
    const word = apple ? event.altKey && !event.metaKey : event.ctrlKey && !event.altKey;
    const shift = event.shiftKey;
    const done = (): true => {
      event.preventDefault();
      return true;
    };
    if (key !== 'Tab' && key !== 'Escape' && key !== 'Shift' && key !== 'Control' && key !== 'Alt' && key !== 'Meta') this.escaped = false;
    if (!this.range()) {
      // No caret yet: the first key that needs one puts it at the start of the text.
      const start = this.storyEdge(-1);
      if (!start) return false;
      this.selection = { anchor: start, focus: copyPosition(start) };
    }
    switch (key) {
      case 'ArrowLeft':
      case 'ArrowRight': {
        const direction = key === 'ArrowLeft' ? -1 : 1;
        if (apple && event.metaKey) this.moveToLineEdge(direction, shift);
        else this.moveBy(word ? 'word' : 'char', direction, shift);
        return done();
      }
      case 'ArrowUp':
      case 'ArrowDown': {
        const direction = key === 'ArrowUp' ? -1 : 1;
        if (apple && event.metaKey) this.moveBy('story', direction, shift);
        else if (word) this.moveBy('paragraph', direction, shift);
        else this.moveVertically(direction, shift, false);
        return done();
      }
      case 'Home':
      case 'End': {
        const direction = key === 'Home' ? -1 : 1;
        if (event.ctrlKey || event.metaKey) this.moveBy('story', direction, shift);
        else this.moveToLineEdge(direction, shift);
        return done();
      }
      case 'PageUp':
      case 'PageDown':
        this.moveVertically(key === 'PageUp' ? -1 : 1, shift, true);
        return done();
      case 'Enter':
        this.enter(command ? 'page' : shift ? 'line' : 'paragraph');
        return done();
      case 'Backspace':
        this.removeText(-1, apple && event.metaKey ? 'line' : word ? 'word' : 'char');
        return done();
      case 'Delete':
        this.removeText(1, word ? 'word' : 'char');
        return done();
      case 'Tab':
        if (command || event.altKey) return false;
        // After Escape the key is the browser's again: the focus moves on, out of the editor.
        if (this.escaped) {
          this.escaped = false;
          return true;
        }
        this.tab(shift);
        return done();
      case 'Escape': {
        const range = this.range();
        if (range && !ops.samePosition(range.anchor, range.focus)) this.moveTo(copyPosition(range.focus), this.affinity, false, false);
        this.escaped = true;
        this.live.textContent = this.strings.tabLeaves;
        return done();
      }
      default:
    }
    if (!command) return key.length === 1 || key === 'Dead' || key === 'Process' || key === 'Unidentified';
    switch (key.toLowerCase()) {
      case 'a':
        this.selectAll();
        return done();
      case 'z':
        if (shift) this.redo();
        else this.undo();
        return done();
      case 'y':
        this.redo();
        return done();
      case 'b':
        this.toggleFormat('bold');
        return done();
      case 'i':
        this.toggleFormat('italic');
        return done();
      case 'u':
        this.toggleFormat('underline');
        return done();
      case 'e':
        this.formatParagraph({ align: 'center' });
        return done();
      case 'l':
        this.formatParagraph({ align: 'left' });
        return done();
      case 'r':
        this.formatParagraph({ align: 'right' });
        return done();
      case 'j':
        this.formatParagraph({ align: 'both' });
        return done();
      case 'k':
        if (!this.options.onLinkShortcut) return false;
        this.options.onLinkShortcut();
        return done();
      default:
        // Zoom, save, copy and paste: the view's other keys and the browser's.
        return false;
    }
  }

  private notify(message: string, kind: 'error' | 'info'): void {
    this.live.textContent = message;
    if (this.options.onNotice) this.options.onNotice({ message, kind });
  }
}
