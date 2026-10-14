// The pages a sheet is printed on: which cells go on which page, at what
// scale, on what paper. Works from the model alone (Sheet.pageSetup, the
// defined names of the print area and the print titles, row and column sizes)
// and touches no DOM, so the same code runs in Node.
//
// Lengths are in px, 96 to the inch, like the grid: "paper px" on the paper,
// "sheet px" in the cells, which are `scale` times as large on the paper.
//
// The rules are Excel's as far as they are known:
//  - Without a print area the used range is printed: from A1 to the last cell
//    with a value, a formula, a fill or a border, the last merged cell and
//    the last cell under a drawing.
//  - Hidden rows and columns take no room and never start a page.
//  - A page is filled with whole rows and columns. One that is larger than
//    the page is not cut in the middle (Excel does that): it gets a page of
//    its own and what does not fit is cut off at the margin.
//  - Manual breaks start a page; a sheet that is fitted to pages ignores them.
//  - "Fit to" never enlarges: the largest scale from 100% down to 10%, in
//    whole percents, at which the pages are enough.
//  - Print titles are repeated on every page that starts after them.

import { resolveCellStyle } from '../model/sheet/color';
import { MAX_COLS, MAX_ROWS, type CellRange, type PageBreak, type PageMargins, type PageSettings, type Sheet, type SheetStructuralEdit, type StyleSheet, type Workbook } from '../model/sheet/types';
import { movePageBreak } from '../model/sheet/page-breaks';
import { createSheetMetrics, headerWidthFor, type AxisMetrics, type SheetMetrics } from './sheet-metrics';

const PX_PER_INCH = 96;
/** Paper sizes by their code, in inches (width, height upright). Codes that are not here print on the default paper. */
const MM = 1 / 25.4;
const PAPER_SIZES = new Map<number, [number, number]>([
  [1, [8.5, 11]], // Letter
  [2, [8.5, 11]], // Letter small
  [3, [11, 17]], // Tabloid
  [4, [17, 11]], // Ledger
  [5, [8.5, 14]], // Legal
  [6, [5.5, 8.5]], // Statement
  [7, [7.25, 10.5]], // Executive
  [8, [297 * MM, 420 * MM]], // A3
  [9, [210 * MM, 297 * MM]], // A4
  [10, [210 * MM, 297 * MM]], // A4 small
  [11, [148 * MM, 210 * MM]], // A5
  [12, [257 * MM, 364 * MM]], // B4 (JIS)
  [13, [182 * MM, 257 * MM]], // B5 (JIS)
  [14, [8.5, 13]], // Folio
  [15, [215 * MM, 275 * MM]], // Quarto
  [16, [10, 14]],
  [17, [11, 17]],
  [18, [8.5, 11]], // Note
  [19, [3.875, 8.875]], // Envelope #9
  [20, [4.125, 9.5]], // Envelope #10
  [27, [110 * MM, 220 * MM]], // Envelope DL
  [28, [162 * MM, 229 * MM]], // Envelope C5
  [34, [176 * MM, 250 * MM]], // Envelope B5
  [66, [420 * MM, 594 * MM]], // A2
  [70, [105 * MM, 148 * MM]], // A6
]);
/** The size of a paper by its code in inches (width, height upright); undefined for a code that is not known here. */
export function paperInches(code: number): [number, number] | undefined {
  const size = PAPER_SIZES.get(code);
  return size ? [size[0], size[1]] : undefined;
}
export const PAPER_LETTER = 1;
export const PAPER_A4 = 9;
/** The sides of a custom paper: between half an inch and 100 inches. */
const MIN_PAPER_INCHES = 0.5;
const MAX_PAPER_INCHES = 100;
/** Excel's "Normal" margins, in inches. */
export const DEFAULT_PAGE_MARGINS: PageMargins = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 };
/** What margins have to leave of each side of the paper for the cells; margins that leave less are replaced by the defaults. */
const MIN_CONTENT_PX = 24;
const MIN_SCALE = 10;
const MAX_SCALE = 400;
/** Never more pages than this for one sheet, unless the caller sets another limit. */
export const MAX_SHEET_PAGES = 2000;
const MAX_PRINT_AREAS = 64;
/** Rows plus columns of all print areas together that fit-to-page still walks. */
const MAX_FIT_SPAN = 4000000;
/** Height of the row of column letters when headings are printed, in sheet px. */
const HEADING_HEIGHT = 20;
const EPSILON = 1e-6;

/** Rows or columns from `first` to `last`, both included. */
export interface PrintBand {
  first: number;
  last: number;
}

export interface SheetPrintPage {
  /** Which of the print areas the page belongs to. */
  area: number;
  /** The rows and columns of the page; hidden ones among them take no room. */
  rows: PrintBand;
  cols: PrintBand;
  /** Print titles repeated above and left of the cells of this page. */
  titleRows?: PrintBand;
  titleCols?: PrintBand;
  /** Where the printed block (headings, titles, cells) starts on the paper, in paper px. */
  x: number;
  y: number;
  /** Its size in paper px, cut at the margins. */
  width: number;
  height: number;
}

export interface SheetPrintLayout {
  /** The paper as it is fed: turned for landscape. Paper px. */
  paper: { width: number; height: number };
  landscape: boolean;
  /** Paper px. */
  margins: PageMargins;
  /** Paper px per sheet px: 1 is 100%. */
  scale: number;
  /** Row numbers and column letters are printed, and how much room they take in sheet px. */
  headings: boolean;
  headingWidth: number;
  headingHeight: number;
  gridLines: boolean;
  /** In the order they are printed and numbered. */
  pages: SheetPrintPage[];
  /** The number of the first page, when the sheet states one. */
  firstPageNumber?: number;
}

export interface SheetPagesOptions {
  /** Row and column sizes; made from the sheet when missing. */
  metrics?: SheetMetrics;
  /** For telling which cell formats show on paper; without it only cells with content count as used. */
  styles?: StyleSheet;
  /** The print areas in the order they are printed; without any the used range is printed. */
  areas?: readonly CellRange[];
  titleRows?: PrintBand;
  titleCols?: PrintBand;
  /** The boxes of the drawings of the sheet, in sheet px from the corner of A1: they are part of the used range. */
  drawings?: readonly { x: number; y: number; width: number; height: number }[];
  /** Paper size code for a sheet that states none. Default A4. */
  defaultPaper?: number;
  /** A sheet with more pages is refused. Default MAX_SHEET_PAGES. */
  maxPages?: number;
}

/** 'empty': the sheet has nothing to print. 'too-long': it has more pages than the limit. */
export class SheetPagesError extends Error {
  readonly code: 'empty' | 'too-long';
  constructor(code: 'empty' | 'too-long', message: string) {
    super(message);
    this.name = 'SheetPagesError';
    this.code = code;
  }
}

const finite = (value: unknown): value is number => typeof value === 'number' && value === value && value !== Infinity && value !== -Infinity;

/** A length with its unit ("210mm", "8.5in", "595pt") in inches. */
function measureInches(text: string | undefined): number | undefined {
  if (typeof text !== 'string') return undefined;
  const match = /^\s*([0-9]{1,6}(?:\.[0-9]{1,6})?)\s*(mm|cm|in|pt|pc|pi)\s*$/i.exec(text);
  if (!match) return undefined;
  const value = Number(match[1]);
  const unit = match[2].toLowerCase();
  const inches = unit === 'mm' ? value / 25.4 : unit === 'cm' ? value / 2.54 : unit === 'pt' ? value / 72 : unit === 'in' ? value : value / 6;
  return inches >= MIN_PAPER_INCHES && inches <= MAX_PAPER_INCHES ? inches : undefined;
}

/** The paper of a sheet in paper px, turned for landscape. */
export function paperSize(page: PageSettings | undefined, defaultPaper: number = PAPER_A4): { width: number; height: number; landscape: boolean } {
  let size: [number, number] | undefined;
  if (page) {
    const width = measureInches(page.paperWidth);
    const height = measureInches(page.paperHeight);
    if (width !== undefined && height !== undefined) size = [width, height];
    else if (page.paperSize !== undefined) size = PAPER_SIZES.get(page.paperSize);
  }
  if (!size) size = PAPER_SIZES.get(defaultPaper) || (PAPER_SIZES.get(PAPER_A4) as [number, number]);
  const landscape = !!page && page.orientation === 'landscape';
  const width = Math.round(size[0] * PX_PER_INCH * 100) / 100;
  const height = Math.round(size[1] * PX_PER_INCH * 100) / 100;
  return landscape ? { width: height, height: width, landscape } : { width, height, landscape };
}

/** The margins in paper px; margins that leave no room for cells are replaced by the defaults. */
export function marginsOf(stated: PageMargins | undefined, paper: { width: number; height: number }): PageMargins {
  const px = (source: PageMargins): PageMargins => ({
    left: source.left * PX_PER_INCH,
    right: source.right * PX_PER_INCH,
    top: source.top * PX_PER_INCH,
    bottom: source.bottom * PX_PER_INCH,
    header: source.header * PX_PER_INCH,
    footer: source.footer * PX_PER_INCH,
  });
  if (stated) {
    const side = (value: unknown, fallback: number): number => (finite(value) && value >= 0 ? value : fallback);
    const margins = px({
      left: side(stated.left, DEFAULT_PAGE_MARGINS.left),
      right: side(stated.right, DEFAULT_PAGE_MARGINS.right),
      top: side(stated.top, DEFAULT_PAGE_MARGINS.top),
      bottom: side(stated.bottom, DEFAULT_PAGE_MARGINS.bottom),
      header: side(stated.header, DEFAULT_PAGE_MARGINS.header),
      footer: side(stated.footer, DEFAULT_PAGE_MARGINS.footer),
    });
    // A header or footer further in than half the paper would be drawn off it: the default distance instead.
    if (!(margins.header <= paper.height / 2)) margins.header = Math.min(DEFAULT_PAGE_MARGINS.header * PX_PER_INCH, paper.height / 16);
    if (!(margins.footer <= paper.height / 2)) margins.footer = Math.min(DEFAULT_PAGE_MARGINS.footer * PX_PER_INCH, paper.height / 16);
    if (paper.width - margins.left - margins.right >= MIN_CONTENT_PX && paper.height - margins.top - margins.bottom >= MIN_CONTENT_PX) return margins;
  }
  const margins = px(DEFAULT_PAGE_MARGINS);
  if (paper.width - margins.left - margins.right >= MIN_CONTENT_PX && paper.height - margins.top - margins.bottom >= MIN_CONTENT_PX) return margins;
  // A paper smaller than the default margins: an eighth of each side.
  return { left: paper.width / 8, right: paper.width / 8, top: paper.height / 8, bottom: paper.height / 8, header: paper.height / 16, footer: paper.height / 16 };
}

// ---------------------------------------------------------------------------
// Print area and print titles

const PRINT_AREA = '_xlnm.print_area';
const PRINT_TITLES = '_xlnm.print_titles';

function columnNumber(letters: string): number {
  let n = 0;
  for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) & 0x5f) - 64;
  return n - 1;
}

/** The parts of a reference list, cut at the commas outside quoted sheet names. */
function splitReferences(value: string): string[] {
  const parts: string[] = [];
  let quoted = false;
  let start = 0;
  for (let i = 0; i < value.length && parts.length < MAX_PRINT_AREAS; i++) {
    const ch = value.charCodeAt(i);
    if (ch === 39) quoted = !quoted;
    // A bracket outside a sheet name: the value is a formula, not a list of references.
    else if (ch === 40 && !quoted) return [];
    else if (ch === 44 && !quoted) {
      parts.push(value.slice(start, i));
      start = i + 1;
    }
  }
  if (parts.length < MAX_PRINT_AREAS) parts.push(value.slice(start));
  return parts;
}

type Reference = { kind: 'cells'; range: CellRange } | { kind: 'rows'; band: PrintBand } | { kind: 'cols'; band: PrintBand };

/** One reference without its sheet name: "$A$1:$C$9", "$1:$3", "$A:$B" or "$B$2". Anything else (a formula, "#REF!") is undefined. */
function parseReference(part: string): Reference | undefined {
  let text = part.trim();
  // The sheet name ends at the last "!" (a quoted name may hold one).
  const bang = text.lastIndexOf('!');
  if (bang >= 0) text = text.slice(bang + 1);
  text = text.replace(/\$/g, '');
  let match = /^([A-Za-z]{1,3})([0-9]{1,7})(?::([A-Za-z]{1,3})([0-9]{1,7}))?$/.exec(text);
  if (match) {
    const c0 = columnNumber(match[1]);
    const r0 = Number(match[2]) - 1;
    const c1 = match[3] !== undefined ? columnNumber(match[3]) : c0;
    const r1 = match[4] !== undefined ? Number(match[4]) - 1 : r0;
    const range = { top: Math.min(r0, r1), left: Math.min(c0, c1), bottom: Math.max(r0, r1), right: Math.max(c0, c1) };
    if (range.top < 0 || range.left < 0 || range.bottom >= MAX_ROWS || range.right >= MAX_COLS) return undefined;
    return { kind: 'cells', range };
  }
  match = /^([0-9]{1,7}):([0-9]{1,7})$/.exec(text);
  if (match) {
    const a = Number(match[1]) - 1;
    const b = Number(match[2]) - 1;
    const band = { first: Math.min(a, b), last: Math.max(a, b) };
    return band.first >= 0 && band.last < MAX_ROWS ? { kind: 'rows', band } : undefined;
  }
  match = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(text);
  if (match) {
    const a = columnNumber(match[1]);
    const b = columnNumber(match[2]);
    const band = { first: Math.min(a, b), last: Math.max(a, b) };
    return band.first >= 0 && band.last < MAX_COLS ? { kind: 'cols', band } : undefined;
  }
  return undefined;
}

/**
 * The print area and the print titles of a sheet, from the defined names
 * "_xlnm.Print_Area" and "_xlnm.Print_Titles" that are local to it. A part
 * of a name that is no plain reference is left out.
 */
export function printRanges(workbook: Pick<Workbook, 'definedNames'>, sheetIndex: number): { areas: CellRange[]; titleRows?: PrintBand; titleCols?: PrintBand } {
  const out: { areas: CellRange[]; titleRows?: PrintBand; titleCols?: PrintBand } = { areas: [] };
  const names = workbook.definedNames || [];
  for (let i = 0; i < names.length; i++) {
    const name = names[i];
    if (!name || name.localSheet !== sheetIndex || typeof name.name !== 'string' || typeof name.value !== 'string') continue;
    const key = name.name.toLowerCase();
    if (key !== PRINT_AREA && key !== PRINT_TITLES) continue;
    if (name.value.length > 8192) continue;
    const parts = splitReferences(name.value);
    for (let p = 0; p < parts.length; p++) {
      const reference = parseReference(parts[p]);
      if (!reference) continue;
      if (key === PRINT_AREA) {
        if (reference.kind === 'cells') out.areas.push(reference.range);
        else if (reference.kind === 'rows') out.areas.push({ top: reference.band.first, bottom: reference.band.last, left: 0, right: MAX_COLS - 1 });
        else out.areas.push({ left: reference.band.first, right: reference.band.last, top: 0, bottom: MAX_ROWS - 1 });
      } else if (reference.kind === 'rows') out.titleRows = out.titleRows || reference.band;
      else if (reference.kind === 'cols') out.titleCols = out.titleCols || reference.band;
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Used range

/**
 * The range that is printed when the sheet has no print area: from A1 to
 * the last cell that shows something. Undefined for a sheet that shows nothing.
 */
export function usedRange(sheet: Sheet, options: Pick<SheetPagesOptions, 'styles' | 'drawings' | 'metrics'> = {}): CellRange | undefined {
  let bottom = -1;
  let right = -1;
  const styles = options.styles;
  // Whether a cell format shows on paper without any content: by its index, looked up once.
  const visible = new Map<number, boolean>();
  const shows = (index: number): boolean => {
    if (!styles) return false;
    let known = visible.get(index);
    if (known === undefined) {
      const style = resolveCellStyle(styles, index);
      const fill = style.fill;
      const border = style.border;
      const edge = (e: { style?: string } | undefined): boolean => !!e && !!e.style && e.style !== 'none';
      known = (!!fill && fill.type === 'pattern' && fill.pattern !== undefined && fill.pattern !== 'none') || (!!border && (edge(border.left) || edge(border.right) || edge(border.top) || edge(border.bottom)));
      visible.set(index, known);
    }
    return known;
  };
  sheet.rows.forEach((row, r) => {
    if (!row || !(r >= 0 && r < MAX_ROWS)) return;
    const cells = row.cells;
    if (!cells) return;
    // From the right: the first cell that shows something ends the search of the row.
    for (let c = Math.min(cells.length, MAX_COLS) - 1; c >= 0; c--) {
      const cell = cells[c];
      if (!cell) continue;
      const content = cell.value === null || cell.value === undefined ? cell.formula !== undefined && cell.formula.text !== '' : cell.value !== '';
      if (!content && !shows(cell.style)) continue;
      if (c > right) right = c;
      if (r > bottom) bottom = r;
      break;
    }
  });
  const merges = sheet.merges || [];
  for (let i = 0; i < merges.length; i++) {
    const merge = merges[i];
    if (!merge || !(merge.bottom >= merge.top && merge.right >= merge.left && merge.top >= 0 && merge.left >= 0)) continue;
    bottom = Math.max(bottom, Math.min(merge.bottom, MAX_ROWS - 1));
    right = Math.max(right, Math.min(merge.right, MAX_COLS - 1));
  }
  const drawings = options.drawings;
  if (drawings && drawings.length > 0) {
    const metrics = options.metrics || createSheetMetrics(sheet);
    for (let i = 0; i < drawings.length; i++) {
      const box = drawings[i];
      if (!box || !finite(box.x) || !finite(box.y) || !(box.width > 0) || !(box.height > 0)) continue;
      const lastX = box.x + box.width - 1;
      const lastY = box.y + box.height - 1;
      if (lastX < 0 || lastY < 0) continue;
      bottom = Math.max(bottom, metrics.rows.indexAt(Math.min(lastY, metrics.rows.total - 1)));
      right = Math.max(right, metrics.cols.indexAt(Math.min(lastX, metrics.cols.total - 1)));
    }
  }
  if (bottom < 0 || right < 0) return undefined;
  return { top: 0, left: 0, bottom, right };
}

// ---------------------------------------------------------------------------
// Pages along one axis

/** The rows and columns inserted or deleted after the first `since` of them: those that move breaks set since (SheetPageSetup.breaksSince). */
export function editsSince(edits: SheetStructuralEdit[] | undefined, since: number | undefined): SheetStructuralEdit[] | undefined {
  if (!edits || !(typeof since === 'number' && since > 0)) return edits;
  return edits.slice(Math.min(Math.floor(since), edits.length));
}

/**
 * The manual breaks of an axis as they are now (the first rows or columns of their pages, in
 * order): rows and columns inserted or deleted since the file was read, or since the breaks were
 * set, move them (see SheetPageSetup.breaksSince).
 */
export function breaksNow(sheet: Pick<Sheet, 'pageSetup' | 'structuralEdits'>, axis: 'row' | 'col'): number[] {
  const setup = sheet.pageSetup;
  if (!setup) return [];
  const moves = editsSince(sheet.structuralEdits, setup.breaksSince);
  return axis === 'row' ? currentBreaks(setup.rowBreaks, moves, 'row', MAX_ROWS) : currentBreaks(setup.colBreaks, moves, 'col', MAX_COLS);
}

/** The manual breaks of an axis as they are now: rows and columns inserted or deleted since the file was read move them. */
function currentBreaks(breaks: PageBreak[] | undefined, edits: SheetStructuralEdit[] | undefined, axis: 'row' | 'col', count: number): number[] {
  if (!breaks || breaks.length === 0) return [];
  const out: number[] = [];
  for (let i = 0; i < breaks.length && i < 4096; i++) {
    const entry = breaks[i];
    if (!entry || !finite(entry.at)) continue;
    const span = { at: Math.floor(entry.at), min: finite(entry.min) ? entry.min : 0, max: finite(entry.max) ? entry.max : count };
    const moved = movePageBreak(span, edits, axis);
    if (moved && moved.at > 0 && moved.at < count && out.indexOf(moved.at) < 0) out.push(moved.at);
  }
  out.sort((a, b) => a - b);
  return out;
}

/** The size of the visible items of a band. */
function bandSize(axis: AxisMetrics, band: PrintBand): number {
  return axis.offset(band.last + 1) - axis.offset(band.first);
}

/**
 * Cuts the items first..last into pages of `room` sheet px. A page that
 * starts after the titles has `titleSize` less. Returns undefined as soon as
 * there are more than `limit` pages, and an empty list when every item is hidden.
 */
function breakAxis(axis: AxisMetrics, first: number, last: number, room: number, title: PrintBand | undefined, titleSize: number, manual: number[], limit: number): PrintBand[] | undefined {
  const bands: PrintBand[] = [];
  let i = axis.skipHidden(first, 1);
  if (i > last || i < 0) return bands;
  const roomFor = (start: number): number => (title && start > title.last ? room - titleSize : room);
  let start = i;
  let previous = i;
  let used = 0;
  let available = roomFor(start);
  let next = 0;
  while (next < manual.length && manual[next] <= start) next++;
  while (i <= last) {
    const size = axis.size(i);
    if (i !== start && ((next < manual.length && manual[next] <= i) || used + size > available + EPSILON)) {
      bands.push({ first: start, last: previous });
      if (bands.length > limit) return undefined;
      start = i;
      used = 0;
      available = roomFor(start);
      while (next < manual.length && manual[next] <= start) next++;
    }
    used += size;
    previous = i;
    i = axis.skipHidden(i + 1, 1);
  }
  bands.push({ first: start, last: previous });
  return bands.length > limit ? undefined : bands;
}

/** A print title cut to the axis; undefined for one that is no range. */
function validBand(band: PrintBand | undefined, count: number): PrintBand | undefined {
  if (!band || !finite(band.first) || !finite(band.last)) return undefined;
  const first = Math.max(Math.floor(band.first), 0);
  const last = Math.min(Math.floor(band.last), count - 1);
  return first <= last ? { first, last } : undefined;
}

function validRange(range: CellRange | undefined): CellRange | undefined {
  if (!range || !finite(range.top) || !finite(range.left) || !finite(range.bottom) || !finite(range.right)) return undefined;
  const top = Math.max(Math.floor(range.top), 0);
  const left = Math.max(Math.floor(range.left), 0);
  const bottom = Math.min(Math.floor(range.bottom), MAX_ROWS - 1);
  const right = Math.min(Math.floor(range.right), MAX_COLS - 1);
  return top <= bottom && left <= right ? { top, left, bottom, right } : undefined;
}

/** A count of pages a sheet is fitted to: 0 is "as many as needed", anything that is no count is 1. */
function fitCount(value: unknown): number {
  if (!finite(value) || value < 0) return 1;
  return Math.min(Math.floor(value), 32767);
}

interface AreaPages {
  rows: PrintBand[];
  cols: PrintBand[];
}

/**
 * Lays a sheet out on pages. Throws a SheetPagesError with code 'empty' for
 * a sheet that has nothing to print (no cell in use, or every row or column
 * of what would be printed hidden) and 'too-long' for one beyond the limit.
 */
export function paginateSheet(sheet: Sheet, options: SheetPagesOptions = {}): SheetPrintLayout {
  const metrics = options.metrics || createSheetMetrics(sheet);
  const setup = sheet.pageSetup || {};
  const page = setup.page || {};
  const paperBox = paperSize(setup.page, options.defaultPaper);
  const paper = { width: paperBox.width, height: paperBox.height };
  const margins = marginsOf(setup.margins, paper);
  const contentWidth = paper.width - margins.left - margins.right;
  const contentHeight = paper.height - margins.top - margins.bottom;
  const limit = finite(options.maxPages) && options.maxPages >= 1 ? Math.floor(options.maxPages) : MAX_SHEET_PAGES;

  let areas: CellRange[] = [];
  if (options.areas) {
    for (let i = 0; i < options.areas.length && areas.length < MAX_PRINT_AREAS; i++) {
      const range = validRange(options.areas[i]);
      if (range) areas.push(range);
    }
  }
  if (areas.length === 0) {
    const used = usedRange(sheet, { styles: options.styles, drawings: options.drawings, metrics });
    if (used) areas = [used];
  } else {
    // A print area of whole rows or columns ends where the sheet does: nobody means a million empty rows.
    const used = usedRange(sheet, { styles: options.styles, drawings: options.drawings, metrics });
    areas = areas.map((range) => {
      const wholeRows = range.left === 0 && range.right === MAX_COLS - 1;
      const wholeCols = range.top === 0 && range.bottom === MAX_ROWS - 1;
      if (!wholeRows && !wholeCols) return range;
      return {
        top: range.top,
        left: range.left,
        bottom: wholeCols ? Math.max(used ? used.bottom : 0, range.top) : range.bottom,
        right: wholeRows ? Math.max(used ? used.right : 0, range.left) : range.right,
      };
    });
  }
  if (areas.length === 0) throw new SheetPagesError('empty', 'The sheet has nothing to print');

  const print = setup.print || {};
  const headings = print.headings === true;
  const gridLines = print.gridLines === true && print.gridLinesSet !== false;
  let lastRow = 0;
  for (let i = 0; i < areas.length; i++) lastRow = Math.max(lastRow, areas[i].bottom);
  const headingWidth = headings ? headerWidthFor(lastRow) : 0;
  const headingHeight = headings ? HEADING_HEIGHT : 0;

  const titleRows = validBand(options.titleRows, MAX_ROWS);
  const titleCols = validBand(options.titleCols, MAX_COLS);
  const titleHeight = titleRows ? bandSize(metrics.rows, titleRows) : 0;
  const titleWidth = titleCols ? bandSize(metrics.cols, titleCols) : 0;

  const fit = setup.fitToPage === true;
  // UNSURE: Excel is said to ignore manual breaks while a sheet is fitted to pages; that is what is done here.
  const moves = editsSince(sheet.structuralEdits, setup.breaksSince);
  const rowBreaks = fit ? [] : currentBreaks(setup.rowBreaks, moves, 'row', MAX_ROWS);
  const colBreaks = fit ? [] : currentBreaks(setup.colBreaks, moves, 'col', MAX_COLS);

  /** The pages of every area at a scale; undefined when they are more than the limit. */
  const pagesAt = (scale: number): AreaPages[] | undefined => {
    const roomX = contentWidth / scale - headingWidth;
    const roomY = contentHeight / scale - headingHeight;
    // Titles that leave no room for cells are not repeated.
    const repeatRows = titleRows && titleHeight > 0 && titleHeight < roomY - EPSILON ? titleRows : undefined;
    const repeatCols = titleCols && titleWidth > 0 && titleWidth < roomX - EPSILON ? titleCols : undefined;
    const out: AreaPages[] = [];
    let total = 0;
    for (let a = 0; a < areas.length; a++) {
      const area = areas[a];
      const cols = breakAxis(metrics.cols, area.left, area.right, roomX, repeatCols, titleWidth, colBreaks, limit);
      if (!cols) return undefined;
      const rows = breakAxis(metrics.rows, area.top, area.bottom, roomY, repeatRows, titleHeight, rowBreaks, limit);
      if (!rows) return undefined;
      total += rows.length * cols.length;
      if (total > limit) return undefined;
      out.push({ rows, cols });
    }
    return out;
  };

  let percent = 100;
  let found: AreaPages[] | undefined;
  if (fit) {
    const wide = fitCount(page.fitToWidth === undefined ? 1 : page.fitToWidth);
    const tall = fitCount(page.fitToHeight === undefined ? 1 : page.fitToHeight);
    // Where to start: the scale at which the cells would fit if pages could be cut anywhere.
    let estimate = 1;
    for (let a = 0; a < areas.length; a++) {
      const area = areas[a];
      const width = bandSize(metrics.cols, { first: area.left, last: area.right }) + headingWidth * Math.max(wide, 1);
      const height = bandSize(metrics.rows, { first: area.top, last: area.bottom }) + headingHeight * Math.max(tall, 1);
      if (wide > 0 && width > 0) estimate = Math.min(estimate, (contentWidth * wide) / width);
      if (tall > 0 && height > 0) estimate = Math.min(estimate, (contentHeight * tall) / height);
    }
    percent = Math.min(100, Math.max(MIN_SCALE, Math.floor(estimate * 100 + EPSILON)));
    // Every try walks the rows and columns of the areas, and a file can state the same large area
    // many times: beyond this many of them the sheet is refused rather than walked again and again.
    let span = 0;
    for (let a = 0; a < areas.length; a++) span += areas[a].bottom - areas[a].top + 1 + (areas[a].right - areas[a].left + 1);
    if (span > MAX_FIT_SPAN) throw new SheetPagesError('too-long', 'The print areas are too large to be fitted to pages');
    // Whole rows and columns fill a page less well than the estimate assumes, so scales below the
    // estimate are tried. A smaller scale never needs more pages: the largest that fits is found by halving.
    const fits = (scale: number): boolean => {
      const roomX = contentWidth / scale - headingWidth;
      const roomY = contentHeight / scale - headingHeight;
      const repeatRows = titleRows && titleHeight > 0 && titleHeight < roomY - EPSILON ? titleRows : undefined;
      const repeatCols = titleCols && titleWidth > 0 && titleWidth < roomX - EPSILON ? titleCols : undefined;
      // What cannot fit even if pages could be cut anywhere is told without walking anything.
      for (let a = 0; a < areas.length; a++) {
        const area = areas[a];
        if (wide > 0 && bandSize(metrics.cols, { first: area.left, last: area.right }) > roomX * wide + EPSILON) return false;
        if (tall > 0 && bandSize(metrics.rows, { first: area.top, last: area.bottom }) > roomY * tall + EPSILON) return false;
      }
      for (let a = 0; a < areas.length; a++) {
        const area = areas[a];
        if (wide > 0 && !breakAxis(metrics.cols, area.left, area.right, roomX, repeatCols, titleWidth, colBreaks, Math.min(wide, limit))) return false;
        if (tall > 0 && !breakAxis(metrics.rows, area.top, area.bottom, roomY, repeatRows, titleHeight, rowBreaks, Math.min(tall, limit))) return false;
      }
      return true;
    };
    if (percent > MIN_SCALE && !fits(percent / 100)) {
      // `low` fits or is the smallest scale there is; `high` does not fit.
      let low = MIN_SCALE;
      let high = percent;
      while (high - low > 1) {
        const middle = Math.floor((low + high) / 2);
        if (fits(middle / 100)) low = middle;
        else high = middle;
      }
      percent = low;
    }
    found = pagesAt(percent / 100);
  } else {
    const stated = page.scale;
    percent = finite(stated) && stated >= MIN_SCALE && stated <= MAX_SCALE ? stated : 100;
    found = pagesAt(percent / 100);
  }
  if (!found) throw new SheetPagesError('too-long', `The sheet has more than ${limit} pages`);

  const scale = percent / 100;
  const roomX = contentWidth / scale - headingWidth;
  const roomY = contentHeight / scale - headingHeight;
  const repeatRows = titleRows && titleHeight > 0 && titleHeight < roomY - EPSILON ? titleRows : undefined;
  const repeatCols = titleCols && titleWidth > 0 && titleWidth < roomX - EPSILON ? titleCols : undefined;
  const pages: SheetPrintPage[] = [];
  const add = (area: number, rows: PrintBand, cols: PrintBand) => {
    const entry: SheetPrintPage = { area, rows, cols, x: margins.left, y: margins.top, width: 0, height: 0 };
    let width = headingWidth + bandSize(metrics.cols, cols);
    let height = headingHeight + bandSize(metrics.rows, rows);
    if (repeatRows && rows.first > repeatRows.last) {
      entry.titleRows = repeatRows;
      height += titleHeight;
    }
    if (repeatCols && cols.first > repeatCols.last) {
      entry.titleCols = repeatCols;
      width += titleWidth;
    }
    entry.width = Math.min(width * scale, contentWidth);
    entry.height = Math.min(height * scale, contentHeight);
    if (print.horizontalCentered === true) entry.x += (contentWidth - entry.width) / 2;
    if (print.verticalCentered === true) entry.y += (contentHeight - entry.height) / 2;
    pages.push(entry);
  };
  const overThenDown = page.pageOrder === 'overThenDown';
  for (let a = 0; a < found.length; a++) {
    const { rows, cols } = found[a];
    if (overThenDown) {
      for (let r = 0; r < rows.length; r++) for (let c = 0; c < cols.length; c++) add(a, rows[r], cols[c]);
    } else {
      for (let c = 0; c < cols.length; c++) for (let r = 0; r < rows.length; r++) add(a, rows[r], cols[c]);
    }
  }
  if (pages.length === 0) throw new SheetPagesError('empty', 'Every row or column of what would be printed is hidden');

  const layout: SheetPrintLayout = { paper, landscape: paperBox.landscape, margins, scale, headings, headingWidth, headingHeight, gridLines, pages };
  if (page.useFirstPageNumber === true && finite(page.firstPageNumber) && Math.abs(page.firstPageNumber) <= 32767) layout.firstPageNumber = Math.floor(page.firstPageNumber);
  return layout;
}
