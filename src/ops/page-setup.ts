// How a worksheet is printed, as the editor asks to change it: the paper, its orientation, the
// margins, the scale or "fit to pages", what is printed beside the cells (grid lines, headings)
// and where (centred), the print area and the print titles, and the manual page breaks. Like the
// commands on cells these functions only build operations; the caller applies them through a
// History so that each call is one undo step.
//
// The page setup of a sheet is never changed in place: a change puts a new object into the
// sheet ("setPageSetup"), and members it leaves alone stay the objects they were, so that the
// writer of the file touches only the elements that changed. The print area and the print
// titles are defined names local to the sheet; they are set by replacing the list of names
// ("setDefinedNames"). Neither asks for anything to be calculated again.

import { quoteSheetName } from '../calc/refs';
import { movePageBreak } from '../model/sheet/page-breaks';
import { DEFAULT_PAGE_MARGINS, breaksNow, editsSince, paperInches, paperSize, printRanges, PAPER_A4, type PrintBand } from '../layout/sheet-pages';
import { MAX_COLS, MAX_ROWS, cellName, columnName, type CellRange, type DefinedName, type PageBreak, type PageMargins, type PageSettings, type PrintOptions, type SheetPageSetup, type Workbook } from '../model/sheet/types';
import type { Op } from './ops';

export const PRINT_AREA_NAME = '_xlnm.Print_Area';
export const PRINT_TITLES_NAME = '_xlnm.Print_Titles';

/** The paper codes the dialog offers (see PAPER_SIZES of the layout): Letter, Legal, Executive, A3, A4, A5, B4, B5, Tabloid. */
export const OFFERED_PAPERS: readonly number[] = [1, 5, 7, 8, 9, 11, 12, 13, 3];

/** Excel's dialog takes margins up to this many inches. */
const MAX_MARGIN_INCHES = 49;
const MIN_SCALE = 10;
const MAX_SCALE = 400;
const MAX_FIT_PAGES = 32767;
/** Most print areas one sheet can be given. */
const MAX_PRINT_AREAS = 64;
/** Most manual breaks along one axis (Excel's limit is 1,026). */
export const MAX_MANUAL_BREAKS = 1026;

/** What a sheet's page setup is now, as the dialog shows it: every value is the one that counts, stated or not. */
export interface PageSetupState {
  /** The paper code, or undefined for a paper the file gives as a measure (paperWidth/paperHeight) or by a code that is not known here. */
  paperSize: number | undefined;
  /** The file states the paper (by code or by measure); otherwise the paper is the default one. */
  paperStated: boolean;
  orientation: 'portrait' | 'landscape';
  /** Inches. */
  margins: PageMargins;
  fitToPage: boolean;
  /** Percent, 10..400. */
  scale: number;
  /** Pages across and down while fitted; 0 is "as many as needed". */
  fitToWidth: number;
  fitToHeight: number;
  horizontalCentered: boolean;
  verticalCentered: boolean;
  gridLines: boolean;
  headings: boolean;
  /** The print areas, in order; empty when the used range is printed. */
  printArea: CellRange[];
  titleRows?: PrintBand;
  titleCols?: PrintBand;
  /** The manual breaks as they are now: the first rows and columns of their pages. */
  rowBreaks: number[];
  colBreaks: number[];
}

/**
 * A change of the page setup of a sheet: only what is given changes. `scale` and `fitTo` exclude
 * each other (a scale ends "fit to pages", fitTo starts it). A print area or print titles of null
 * are cleared.
 */
export interface PageSetupPatch {
  /** A paper code (1 Letter, 5 Legal, 8 A3, 9 A4, 11 A5, ...: see OFFERED_PAPERS); it replaces a paper given as a measure. */
  paperSize?: number;
  orientation?: 'portrait' | 'landscape';
  /** Inches, each 0..49; those left out keep what counts now. */
  margins?: Partial<PageMargins>;
  /** Percent, 10..400. */
  scale?: number;
  /** Pages across and down, 0..32767 each (0: as many as needed). */
  fitTo?: { width: number; height: number };
  horizontalCentered?: boolean;
  verticalCentered?: boolean;
  gridLines?: boolean;
  headings?: boolean;
  printArea?: CellRange[] | null;
  titleRows?: PrintBand | null;
  titleCols?: PrintBand | null;
}

const bad = (message: string): never => {
  throw new RangeError(message);
};

const isWhole = (value: unknown, min: number, max: number): value is number => typeof value === 'number' && value >= min && value <= max && Math.floor(value) === value;
const isLength = (value: unknown): value is number => typeof value === 'number' && value >= 0 && value <= MAX_MARGIN_INCHES && value === value;

function worksheetAt(workbook: Workbook, index: number) {
  const sheet = workbook.sheets[index];
  if (!sheet) return bad(`No sheet at index ${index}`);
  if (sheet.kind !== 'worksheet') return bad('Only a worksheet has a page setup that can be changed');
  return sheet;
}

/** What the page setup of a sheet is now (see PageSetupState); `defaultPaper` is the paper of a sheet that states none (default A4). */
export function pageSetupState(workbook: Workbook, index: number, defaultPaper: number = PAPER_A4): PageSetupState {
  const sheet = workbook.sheets[index];
  if (!sheet) return bad(`No sheet at index ${index}`);
  const setup = sheet.pageSetup || {};
  const page = setup.page || {};
  const print = setup.print || {};
  const measured = typeof page.paperWidth === 'string' && typeof page.paperHeight === 'string';
  const stated = measured || page.paperSize !== undefined;
  let code: number | undefined = defaultPaper;
  if (measured) code = undefined;
  else if (page.paperSize !== undefined) code = paperInches(page.paperSize) ? page.paperSize : undefined;
  const stateMargins = setup.margins || DEFAULT_PAGE_MARGINS;
  const margin = (name: keyof PageMargins): number => (isLength(stateMargins[name]) ? stateMargins[name] : DEFAULT_PAGE_MARGINS[name]);
  const ranges = printRanges(workbook, index);
  const count = (value: unknown): number => (isWhole(value, 0, MAX_FIT_PAGES) ? value : 1);
  const state: PageSetupState = {
    paperSize: code,
    paperStated: stated,
    orientation: page.orientation === 'landscape' ? 'landscape' : 'portrait',
    margins: { left: margin('left'), right: margin('right'), top: margin('top'), bottom: margin('bottom'), header: margin('header'), footer: margin('footer') },
    fitToPage: setup.fitToPage === true,
    scale: isWhole(page.scale, MIN_SCALE, MAX_SCALE) ? page.scale : 100,
    fitToWidth: count(page.fitToWidth),
    fitToHeight: count(page.fitToHeight),
    horizontalCentered: print.horizontalCentered === true,
    verticalCentered: print.verticalCentered === true,
    gridLines: print.gridLines === true && print.gridLinesSet !== false,
    headings: print.headings === true,
    printArea: ranges.areas,
    rowBreaks: breaksNow(sheet, 'row'),
    colBreaks: breaksNow(sheet, 'col'),
  };
  if (ranges.titleRows) state.titleRows = ranges.titleRows;
  if (ranges.titleCols) state.titleCols = ranges.titleCols;
  return state;
}

/** A row or column band of print titles; anything else is refused. */
function checkBand(band: PrintBand, limit: number, what: string): PrintBand {
  if (!band || typeof band !== 'object' || !isWhole(band.first, 0, limit - 1) || !isWhole(band.last, 0, limit - 1) || band.first > band.last) return bad(`The ${what} are not a band of the sheet`);
  return { first: band.first, last: band.last };
}

function checkArea(range: CellRange): CellRange {
  if (!range || typeof range !== 'object' || !isWhole(range.top, 0, MAX_ROWS - 1) || !isWhole(range.bottom, 0, MAX_ROWS - 1) || !isWhole(range.left, 0, MAX_COLS - 1) || !isWhole(range.right, 0, MAX_COLS - 1)) {
    return bad('A print area is not a range of the sheet');
  }
  return { top: Math.min(range.top, range.bottom), bottom: Math.max(range.top, range.bottom), left: Math.min(range.left, range.right), right: Math.max(range.left, range.right) };
}

const absoluteRow = (row: number): string => `$${row + 1}`;
const absoluteCol = (col: number): string => `$${columnName(col)}`;
function absoluteCell(row: number, col: number): string {
  const name = cellName(row, col);
  const digits = name.search(/\d/);
  return `$${name.slice(0, digits)}$${name.slice(digits)}`;
}

/** A range as a print area names it: "$A$1:$C$9", whole rows "$1:$3", whole columns "$A:$B". */
function areaText(range: CellRange): string {
  if (range.left === 0 && range.right === MAX_COLS - 1) return `${absoluteRow(range.top)}:${absoluteRow(range.bottom)}`;
  if (range.top === 0 && range.bottom === MAX_ROWS - 1) return `${absoluteCol(range.left)}:${absoluteCol(range.right)}`;
  const first = absoluteCell(range.top, range.left);
  return range.top === range.bottom && range.left === range.right ? first : `${first}:${absoluteCell(range.bottom, range.right)}`;
}

/** The value of the print area name of a sheet: each range with the name of the sheet, separated by commas. */
export function printAreaValue(sheetName: string, areas: readonly CellRange[]): string {
  const prefix = `${quoteSheetName(sheetName)}!`;
  return areas.map((range) => prefix + areaText(range)).join(',');
}

/** The value of the print titles name of a sheet: the columns first, then the rows, as Excel writes them. */
export function printTitlesValue(sheetName: string, rows: PrintBand | undefined, cols: PrintBand | undefined): string {
  const prefix = `${quoteSheetName(sheetName)}!`;
  const parts: string[] = [];
  if (cols) parts.push(`${prefix}${absoluteCol(cols.first)}:${absoluteCol(cols.last)}`);
  if (rows) parts.push(`${prefix}${absoluteRow(rows.first)}:${absoluteRow(rows.last)}`);
  return parts.join(',');
}

const sameBand = (a: PrintBand | undefined, b: PrintBand | undefined): boolean => (!a && !b) || (!!a && !!b && a.first === b.first && a.last === b.last);
const sameAreas = (a: readonly CellRange[], b: readonly CellRange[]): boolean =>
  a.length === b.length && a.every((range, i) => range.top === b[i].top && range.left === b[i].left && range.bottom === b[i].bottom && range.right === b[i].right);

/**
 * The defined names with the name `name` local to the sheet set to `value` ('' removes it). The
 * list is a new one when anything changes; undefined when nothing does.
 */
function withSheetName(names: DefinedName[], index: number, name: string, value: string): DefinedName[] | undefined {
  const key = name.toLowerCase();
  const at = names.findIndex((entry) => entry.localSheet === index && typeof entry.name === 'string' && entry.name.toLowerCase() === key);
  if (at < 0) return value === '' ? undefined : names.concat([{ name, value, localSheet: index, hidden: false }]);
  if (value === '') return names.filter((_, i) => i !== at);
  if (names[at].value === value) return undefined;
  return names.map((entry, i) => (i === at ? { ...entry, value } : entry));
}

/** The keys of an object whose values differ from those of another (undefined and missing are the same). */
function sameFields(a: object | undefined, b: object | undefined): boolean {
  const left = (a || {}) as Record<string, unknown>;
  const right = (b || {}) as Record<string, unknown>;
  const keys = new Set(Object.keys(left).concat(Object.keys(right)));
  for (const key of keys) if (left[key] !== right[key]) return false;
  return true;
}

/**
 * The operations that change the page setup of a worksheet by `patch` (see PageSetupPatch): one
 * "setPageSetup" when the paper, the margins, the scale or what is printed changes, one
 * "setDefinedNames" when the print area or the print titles do. An empty list when nothing changes.
 * Throws RangeError for a patch that is not valid; nothing is built then.
 */
export function setPageSetup(workbook: Workbook, index: number, patch: PageSetupPatch): Op[] {
  const sheet = worksheetAt(workbook, index);
  if (!patch || typeof patch !== 'object') return bad('A change of the page setup is an object');
  if (patch.scale !== undefined && patch.fitTo !== undefined) bad('A sheet is printed either at a scale or fitted to pages');
  const before = sheet.pageSetup || {};
  const next: SheetPageSetup = { ...before };
  const page: PageSettings = { ...(before.page || {}) };
  const print: PrintOptions = { ...(before.print || {}) };

  if (patch.paperSize !== undefined) {
    if (!isWhole(patch.paperSize, 1, 118) || !paperInches(patch.paperSize)) bad('The paper size is not one that is known');
    page.paperSize = patch.paperSize;
    // A paper given as a measure counts before the code: it goes.
    delete page.paperWidth;
    delete page.paperHeight;
  }
  if (patch.orientation !== undefined) {
    if (patch.orientation !== 'portrait' && patch.orientation !== 'landscape') bad('The orientation is portrait or landscape');
    page.orientation = patch.orientation;
  }
  if (patch.scale !== undefined) {
    if (!isWhole(patch.scale, MIN_SCALE, MAX_SCALE)) bad(`The scale is a whole percent from ${MIN_SCALE} to ${MAX_SCALE}`);
    page.scale = patch.scale;
    if (next.fitToPage === true) delete next.fitToPage;
  }
  if (patch.fitTo !== undefined) {
    const fit = patch.fitTo;
    if (!fit || typeof fit !== 'object' || !isWhole(fit.width, 0, MAX_FIT_PAGES) || !isWhole(fit.height, 0, MAX_FIT_PAGES)) bad(`The pages to fit to are whole numbers from 0 to ${MAX_FIT_PAGES}`);
    next.fitToPage = true;
    // 1 is the default of both: it is stated only where the file stated it.
    if (fit.width !== 1 || page.fitToWidth !== undefined) page.fitToWidth = fit.width;
    if (fit.height !== 1 || page.fitToHeight !== undefined) page.fitToHeight = fit.height;
  }
  if (patch.margins !== undefined) {
    const given = patch.margins;
    if (!given || typeof given !== 'object') bad('The margins are an object');
    const now = before.margins;
    const margins: PageMargins = { ...DEFAULT_PAGE_MARGINS };
    for (const name of ['left', 'right', 'top', 'bottom', 'header', 'footer'] as const) {
      const value = given[name];
      if (value !== undefined) {
        if (!isLength(value)) bad(`A margin is a length from 0 to ${MAX_MARGIN_INCHES} inches`);
        margins[name] = value;
      } else if (now && isLength(now[name])) margins[name] = now[name];
    }
    // The margins have to leave room for the cells on the paper the sheet is printed on.
    const paper = paperSize(page, PAPER_A4);
    const inches = (px: number) => px / 96;
    if (margins.left + margins.right >= inches(paper.width) || margins.top + margins.bottom >= inches(paper.height)) bad('The margins leave no room on the paper');
    if (!now || !sameFields(now, margins)) next.margins = margins;
  }
  // What is printed beside the cells. Off is the default: a setting that is turned off is no longer stated.
  const flag = (name: 'horizontalCentered' | 'verticalCentered' | 'headings', value: boolean | undefined) => {
    if (value === undefined) return;
    if (typeof value !== 'boolean') bad('What is printed is either on or off');
    if (value) print[name] = true;
    else if (print[name] === true) delete print[name];
  };
  flag('horizontalCentered', patch.horizontalCentered);
  flag('verticalCentered', patch.verticalCentered);
  flag('headings', patch.headings);
  if (patch.gridLines !== undefined) {
    if (typeof patch.gridLines !== 'boolean') bad('What is printed is either on or off');
    if (patch.gridLines) {
      print.gridLines = true;
      // Grid lines count only together with gridLinesSet, which is on by default.
      if (print.gridLinesSet === false) delete print.gridLinesSet;
    } else if (print.gridLines === true) delete print.gridLines;
  }
  if (!sameFields(before.page, page)) next.page = page;
  if (!sameFields(before.print, print)) {
    if (Object.keys(print).length > 0 || before.print) next.print = print;
  }

  // The print area and the print titles.
  let names = workbook.definedNames;
  const ranges = printRanges(workbook, index);
  if (patch.printArea !== undefined) {
    const areas = patch.printArea === null ? [] : Array.isArray(patch.printArea) ? patch.printArea.map(checkArea) : bad('The print area is a list of ranges');
    if (areas.length > MAX_PRINT_AREAS) bad(`A sheet has at most ${MAX_PRINT_AREAS} print areas`);
    if (!sameAreas(areas, ranges.areas)) names = withSheetName(names, index, PRINT_AREA_NAME, areas.length === 0 ? '' : printAreaValue(sheet.name, areas)) || names;
  }
  if (patch.titleRows !== undefined || patch.titleCols !== undefined) {
    const rows = patch.titleRows === undefined ? ranges.titleRows : patch.titleRows === null ? undefined : checkBand(patch.titleRows, MAX_ROWS, 'title rows');
    const cols = patch.titleCols === undefined ? ranges.titleCols : patch.titleCols === null ? undefined : checkBand(patch.titleCols, MAX_COLS, 'title columns');
    if (!sameBand(rows, ranges.titleRows) || !sameBand(cols, ranges.titleCols)) names = withSheetName(names, index, PRINT_TITLES_NAME, printTitlesValue(sheet.name, rows, cols)) || names;
  }

  const ops: Op[] = [];
  if (!sameFields(before, next)) ops.push({ type: 'setPageSetup', sheet: index, setup: next });
  if (names !== workbook.definedNames) ops.push({ type: 'setDefinedNames', sheet: index, names });
  return ops;
}

/**
 * The breaks of one axis as they are now, as PageBreak objects: moved by the edits (their place
 * and the band of the other axis they span, see movePageBreak()), those that are gone left out;
 * what else a break of the file says stays.
 */
export function movedBreaks(list: PageBreak[] | undefined, edits: ReturnType<typeof editsSince>, axis: 'row' | 'col'): PageBreak[] {
  const limit = axis === 'row' ? MAX_ROWS : MAX_COLS;
  const out: PageBreak[] = [];
  const seen = new Set<number>();
  // Of two breaks that come to one place the one the file lists first stays (the writer keeps that one too).
  const ordered = (list || []).slice(0, 4096).filter((entry) => !!entry && typeof entry.at === 'number' && entry.at === entry.at);
  if (ordered.some((entry) => typeof entry.fileOrder === 'number')) ordered.sort((a, b) => (typeof a.fileOrder === 'number' ? a.fileOrder : Infinity) - (typeof b.fileOrder === 'number' ? b.fileOrder : Infinity));
  for (const entry of ordered) {
    let moved = movePageBreak(entry.at === Math.floor(entry.at) ? entry : { ...entry, at: Math.floor(entry.at) }, edits, axis);
    if (!moved || !(moved.at > 0 && moved.at < limit) || seen.has(moved.at)) continue;
    seen.add(moved.at);
    // The breaks are written in order from here on: their place in the file is that one.
    if (typeof moved.fileOrder === 'number') {
      moved = { ...moved };
      delete moved.fileOrder;
    }
    out.push(moved);
  }
  out.sort((a, b) => a.at - b.at);
  return out;
}

/** The page setup with the breaks of both axes changed by `change`, which gets them as they are now. */
function withBreaks(workbook: Workbook, index: number, change: (rows: PageBreak[], cols: PageBreak[]) => { rows: PageBreak[]; cols: PageBreak[] } | undefined): Op[] {
  const sheet = worksheetAt(workbook, index);
  const before = sheet.pageSetup || {};
  const edits = editsSince(sheet.structuralEdits, before.breaksSince);
  const rows = movedBreaks(before.rowBreaks, edits, 'row');
  const cols = movedBreaks(before.colBreaks, edits, 'col');
  const changed = change(rows, cols);
  if (!changed) return [];
  if (changed.rows.length > MAX_MANUAL_BREAKS || changed.cols.length > MAX_MANUAL_BREAKS) bad(`A sheet has at most ${MAX_MANUAL_BREAKS} page breaks in each direction`);
  const next: SheetPageSetup = { ...before };
  if (changed.rows.length > 0) next.rowBreaks = changed.rows;
  else delete next.rowBreaks;
  if (changed.cols.length > 0) next.colBreaks = changed.cols;
  else delete next.colBreaks;
  // The breaks are where they are now: the edits made so far are behind them.
  const count = sheet.structuralEdits ? sheet.structuralEdits.length : 0;
  if (count > 0) next.breaksSince = count;
  else delete next.breaksSince;
  return [{ type: 'setPageSetup', sheet: index, setup: next }];
}

/** Where a break goes along one axis: `at` is the first row or column of the new page; undefined leaves the axis alone. */
export interface BreakPlace {
  row?: number;
  col?: number;
}

function checkPlace(at: BreakPlace): void {
  if (!at || typeof at !== 'object') bad('A page break needs a place');
  if (at.row !== undefined && !isWhole(at.row, 0, MAX_ROWS - 1)) bad('The row of a page break is outside the sheet');
  if (at.col !== undefined && !isWhole(at.col, 0, MAX_COLS - 1)) bad('The column of a page break is outside the sheet');
}

/**
 * Manual page breaks above row `at.row` and left of column `at.col` (Excel's Insert Page Break at
 * a cell: row 0 and column 0 get none). An empty list when the breaks are there already.
 */
export function insertPageBreak(workbook: Workbook, index: number, at: BreakPlace): Op[] {
  checkPlace(at);
  return withBreaks(workbook, index, (rows, cols) => {
    let changed = false;
    const add = (list: PageBreak[], place: number | undefined, crossMax: number): PageBreak[] => {
      if (place === undefined || place === 0 || list.some((entry) => entry.at === place)) return list;
      changed = true;
      return list.concat([{ at: place, min: 0, max: crossMax, manual: true }]).sort((a, b) => a.at - b.at);
    };
    const nextRows = add(rows, at.row, MAX_COLS - 1);
    const nextCols = add(cols, at.col, MAX_ROWS - 1);
    return changed ? { rows: nextRows, cols: nextCols } : undefined;
  });
}

/** Takes away the breaks above row `at.row` and left of column `at.col`. An empty list when there are none. */
export function removePageBreak(workbook: Workbook, index: number, at: BreakPlace): Op[] {
  checkPlace(at);
  return withBreaks(workbook, index, (rows, cols) => {
    const nextRows = at.row === undefined ? rows : rows.filter((entry) => entry.at !== at.row);
    const nextCols = at.col === undefined ? cols : cols.filter((entry) => entry.at !== at.col);
    return nextRows.length !== rows.length || nextCols.length !== cols.length ? { rows: nextRows, cols: nextCols } : undefined;
  });
}

/** Takes away every manual page break of the sheet (Excel's Reset All Page Breaks); breaks of pivot tables stay. */
export function resetPageBreaks(workbook: Workbook, index: number): Op[] {
  return withBreaks(workbook, index, (rows, cols) => {
    const kept = (entry: PageBreak) => !entry.manual || entry.pivot === true;
    const nextRows = rows.filter(kept);
    const nextCols = cols.filter(kept);
    return nextRows.length !== rows.length || nextCols.length !== cols.length ? { rows: nextRows, cols: nextCols } : undefined;
  });
}

// ---------------------------------------------------------------------------
// The print area and the print titles as a dialog shows and takes them

/** The column number of letters ("A" 0, "AA" 26); -1 for what is no column of the sheet. */
function columnOf(letters: string): number {
  let n = 0;
  for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) & 0x5f) - 64;
  return n >= 1 && n <= MAX_COLS ? n - 1 : -1;
}

/** One reference as typed, with or without "$" and a sheet name: a range, a cell, whole rows or whole columns. */
function referenceOf(text: string): CellRange | undefined {
  let part = text.trim();
  const bang = part.lastIndexOf('!');
  if (bang >= 0) part = part.slice(bang + 1);
  part = part.replace(/\$/g, '').trim();
  let match = /^([A-Za-z]{1,3})(\d{1,7})(?::([A-Za-z]{1,3})(\d{1,7}))?$/.exec(part);
  if (match) {
    const c0 = columnOf(match[1]);
    const r0 = Number(match[2]) - 1;
    const c1 = match[3] !== undefined ? columnOf(match[3]) : c0;
    const r1 = match[4] !== undefined ? Number(match[4]) - 1 : r0;
    if (c0 < 0 || c1 < 0 || r0 < 0 || r1 < 0 || r0 >= MAX_ROWS || r1 >= MAX_ROWS) return undefined;
    return { top: Math.min(r0, r1), bottom: Math.max(r0, r1), left: Math.min(c0, c1), right: Math.max(c0, c1) };
  }
  match = /^(\d{1,7}):(\d{1,7})$/.exec(part);
  if (match) {
    const a = Number(match[1]) - 1;
    const b = Number(match[2]) - 1;
    if (a < 0 || b < 0 || a >= MAX_ROWS || b >= MAX_ROWS) return undefined;
    return { top: Math.min(a, b), bottom: Math.max(a, b), left: 0, right: MAX_COLS - 1 };
  }
  match = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(part);
  if (match) {
    const a = columnOf(match[1]);
    const b = columnOf(match[2]);
    if (a < 0 || b < 0) return undefined;
    return { top: 0, bottom: MAX_ROWS - 1, left: Math.min(a, b), right: Math.max(a, b) };
  }
  return undefined;
}

/**
 * A print area as typed: references separated by "," or ";" ("A1:D20", "$A$1:$D$20;F1:G5",
 * "Sheet1!A:C", "1:5"). Null for an empty text (no print area); undefined for one that is not
 * a list of references.
 */
export function parsePrintArea(text: string): CellRange[] | null | undefined {
  const trimmed = String(text).trim();
  if (trimmed === '') return null;
  if (trimmed.length > 4096) return undefined;
  // A quoted sheet name may hold a separator: the names are taken out first.
  const plain = trimmed.replace(/'(?:[^']|'')*'!/g, '');
  const parts = plain.split(/[;,]/);
  if (parts.length > MAX_PRINT_AREAS) return undefined;
  const out: CellRange[] = [];
  for (const part of parts) {
    const range = referenceOf(part);
    if (!range) return undefined;
    out.push(range);
  }
  return out;
}

/** Print titles as typed: rows ("1:2", "$1", "$1:$3") or columns ("A:B", "$A"). Null for an empty text; undefined for anything else. */
export function parseTitleBand(text: string, axis: 'row' | 'col'): PrintBand | null | undefined {
  let part = String(text).trim();
  if (part === '') return null;
  if (part.length > 512) return undefined;
  const bang = part.lastIndexOf('!');
  if (bang >= 0) part = part.slice(bang + 1);
  part = part.replace(/\$/g, '').trim();
  const match = axis === 'row' ? /^(\d{1,7})(?::(\d{1,7}))?$/.exec(part) : /^([A-Za-z]{1,3})(?::([A-Za-z]{1,3}))?$/.exec(part);
  if (!match) return undefined;
  const at = (value: string): number => (axis === 'row' ? Number(value) - 1 : columnOf(value));
  const a = at(match[1]);
  const b = match[2] !== undefined ? at(match[2]) : a;
  const limit = axis === 'row' ? MAX_ROWS : MAX_COLS;
  if (a < 0 || b < 0 || a >= limit || b >= limit) return undefined;
  return { first: Math.min(a, b), last: Math.max(a, b) };
}

/** A print area as the dialog shows it: absolute references, separated by `separator`. */
export function printAreaText(areas: readonly CellRange[], separator = ','): string {
  return areas.map(areaText).join(separator);
}

/** Print titles as the dialog shows them: "$1:$2", "$A:$B"; '' for none. */
export function titleBandText(band: PrintBand | undefined, axis: 'row' | 'col'): string {
  if (!band) return '';
  return axis === 'row' ? `${absoluteRow(band.first)}:${absoluteRow(band.last)}` : `${absoluteCol(band.first)}:${absoluteCol(band.last)}`;
}
