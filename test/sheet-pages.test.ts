import assert from 'node:assert/strict';
import { test } from 'node:test';
import { createSheetMetrics } from '../src/layout/sheet-metrics';
import { DEFAULT_PAGE_MARGINS, PAPER_LETTER, SheetPagesError, paginateSheet, paperSize, printRanges, usedRange, type SheetPrintLayout } from '../src/layout/sheet-pages';
import { MAX_COLS, MAX_ROWS, createSheet, emptyStyleSheet } from '../src/model/sheet/types';
import type { Cell, DefinedName, Row, Sheet, SheetPageSetup, StyleSheet } from '../src/model/sheet/types';

// Default geometry: columns 64 px, rows 20 px. A4 is 793.7 x 1122.52 px; with the default margins
// (0.7 in left and right, 0.75 in top and bottom) the cells get 659.3 x 978.52 px: 10 columns and 48 rows.
const A4 = { width: 793.7, height: 1122.52 };
const A4_COLS = 10;
const A4_ROWS = 48;

function row(cells: Cell[] = [], extra: Partial<Row> = {}): Row {
  return { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells, ...extra };
}

const num = (value: number, style = 0): Cell => ({ kind: 'number', value, style });

/** A sheet with a number in every cell of the first `rows` x `cols` cells. */
function grid(rows: number, cols: number, setup?: SheetPageSetup): Sheet {
  const sheet = createSheet({ name: 'Sheet1', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/xl/worksheets/sheet1.xml' });
  for (let r = 0; r < rows; r++) {
    const cells: Cell[] = [];
    for (let c = 0; c < cols; c++) cells[c] = num(r * cols + c);
    sheet.rows.set(r, row(cells));
  }
  sheet.rowCount = rows;
  sheet.colCount = cols;
  if (setup) sheet.pageSetup = setup;
  return sheet;
}

/** A sheet whose only cell is at (r, c). */
function single(r: number, c: number, setup?: SheetPageSetup): Sheet {
  const sheet = createSheet({ name: 'Sheet1', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/xl/worksheets/sheet1.xml' });
  const cells: Cell[] = [];
  cells[c] = num(1);
  sheet.rows.set(r, row(cells));
  if (setup) sheet.pageSetup = setup;
  return sheet;
}

const bands = (layout: SheetPrintLayout) => layout.pages.map((page) => `${page.rows.first}-${page.rows.last}/${page.cols.first}-${page.cols.last}`);

const errorCode = (run: () => unknown): string => {
  try {
    run();
  } catch (error) {
    if (error instanceof SheetPagesError) return error.code;
    throw error;
  }
  return 'none';
};

test('sheet pages: paper sizes by code, custom sizes, orientation and the default paper', () => {
  assert.deepEqual(paperSize(undefined), { width: 793.7, height: 1122.52, landscape: false });
  assert.deepEqual(paperSize(undefined, PAPER_LETTER), { width: 816, height: 1056, landscape: false });
  assert.deepEqual(paperSize({ paperSize: 1 }), { width: 816, height: 1056, landscape: false });
  assert.deepEqual(paperSize({ paperSize: 5 }), { width: 816, height: 1344, landscape: false });
  assert.deepEqual(paperSize({ paperSize: 8 }), { width: 1122.52, height: 1587.4, landscape: false });
  assert.deepEqual(paperSize({ paperSize: 11 }), { width: 559.37, height: 793.7, landscape: false });
  assert.deepEqual(paperSize({ paperSize: 9, orientation: 'landscape' }), { width: 1122.52, height: 793.7, landscape: true });
  // "default" leaves the choice to the application, which prints upright.
  assert.deepEqual(paperSize({ paperSize: 9, orientation: 'default' }), { width: 793.7, height: 1122.52, landscape: false });
  // A code that is not known, and one that is no code at all, print on the default paper.
  assert.deepEqual(paperSize({ paperSize: 4711 }), { width: 793.7, height: 1122.52, landscape: false });
  assert.deepEqual(paperSize({ paperSize: NaN }, PAPER_LETTER), { width: 816, height: 1056, landscape: false });
  // A custom size replaces the code; a size that cannot be read does not.
  assert.deepEqual(paperSize({ paperSize: 1, paperWidth: '100mm', paperHeight: '5in' }), { width: 377.95, height: 480, landscape: false });
  assert.deepEqual(paperSize({ paperSize: 1, paperWidth: '100furlongs', paperHeight: '5in' }), { width: 816, height: 1056, landscape: false });
  assert.deepEqual(paperSize({ paperSize: 1, paperWidth: '0.001mm', paperHeight: '99999in' }), { width: 816, height: 1056, landscape: false });
});

test('sheet pages: a sheet without page setup is printed on the default paper with Excel margins, whole rows and columns per page', () => {
  const layout = paginateSheet(grid(100, 25));
  assert.deepEqual(layout.paper, A4);
  assert.equal(layout.scale, 1);
  assert.equal(layout.headings, false);
  assert.equal(layout.gridLines, false);
  assert.equal(layout.firstPageNumber, undefined);
  assert.ok(Math.abs(layout.margins.left - DEFAULT_PAGE_MARGINS.left * 96) < 1e-9);
  assert.ok(Math.abs(layout.margins.top - DEFAULT_PAGE_MARGINS.top * 96) < 1e-9);
  // 100 rows are 3 pages down (48 + 48 + 4), 25 columns 3 pages across (10 + 10 + 5); down, then over.
  assert.deepEqual(bands(layout), [
    '0-47/0-9', '48-95/0-9', '96-99/0-9',
    '0-47/10-19', '48-95/10-19', '96-99/10-19',
    '0-47/20-24', '48-95/20-24', '96-99/20-24',
  ]);
  const first = layout.pages[0];
  assert.ok(Math.abs(first.x - 67.2) < 1e-9 && Math.abs(first.y - 72) < 1e-9);
  assert.equal(first.width, A4_COLS * 64);
  assert.equal(first.height, A4_ROWS * 20);
  assert.equal(layout.pages[8].width, 5 * 64);
  assert.equal(layout.pages[8].height, 4 * 20);
  assert.equal(first.titleRows, undefined);
});

test('sheet pages: the page order, over then down', () => {
  const layout = paginateSheet(grid(60, 15, { page: { pageOrder: 'overThenDown' } }));
  assert.deepEqual(bands(layout), ['0-47/0-9', '0-47/10-14', '48-59/0-9', '48-59/10-14']);
});

test('sheet pages: the used range ends at the last cell with content, a visible format, a merge or a drawing', () => {
  const styles: StyleSheet = emptyStyleSheet();
  styles.fills.push({ type: 'pattern', pattern: 'solid', fg: { rgb: 'FFFF0000' } });
  styles.borders.push({ diagonalUp: false, diagonalDown: false, bottom: { style: 'thin' } });
  styles.cellXfs.push({ numFmtId: 0, fontId: 0, fillId: styles.fills.length - 1, borderId: 0, quotePrefix: false });
  const filled = styles.cellXfs.length - 1;
  styles.cellXfs.push({ numFmtId: 0, fontId: 0, fillId: 0, borderId: styles.borders.length - 1, quotePrefix: false });
  const bordered = styles.cellXfs.length - 1;

  const sheet = single(2, 1);
  assert.deepEqual(usedRange(sheet), { top: 0, left: 0, bottom: 2, right: 1 });
  // An empty cell and a cell with an empty text are not content; a formula without a result is.
  (sheet.rows.get(2) as Row).cells[5] = { kind: 'blank', value: null, style: 0 };
  sheet.rows.set(9, row([{ kind: 'string', value: '', style: 0 }]));
  assert.deepEqual(usedRange(sheet, { styles }), { top: 0, left: 0, bottom: 2, right: 1 });
  sheet.rows.set(6, row([{ kind: 'blank', value: null, style: 0, formula: { text: 'A1+1' } } as Cell]));
  assert.deepEqual(usedRange(sheet, { styles }), { top: 0, left: 0, bottom: 6, right: 1 });
  // A blank cell with a fill or a border shows on paper, but only when the styles are known.
  const cells: Cell[] = [];
  cells[7] = { kind: 'blank', value: null, style: filled };
  sheet.rows.set(11, row(cells));
  const more: Cell[] = [];
  more[3] = { kind: 'blank', value: null, style: bordered };
  sheet.rows.set(14, row(more));
  assert.deepEqual(usedRange(sheet), { top: 0, left: 0, bottom: 6, right: 1 });
  assert.deepEqual(usedRange(sheet, { styles }), { top: 0, left: 0, bottom: 14, right: 7 });
  // Merged cells and drawings count.
  sheet.merges.push({ top: 20, left: 0, bottom: 21, right: 9 });
  assert.deepEqual(usedRange(sheet, { styles }), { top: 0, left: 0, bottom: 21, right: 9 });
  // A drawing of 100 x 50 px at (700, 500): its last pixel is in column 12 (768..831) and row 27 (540..559).
  assert.deepEqual(usedRange(sheet, { styles, drawings: [{ x: 700, y: 500, width: 100, height: 50 }] }), { top: 0, left: 0, bottom: 27, right: 12 });
  // A drawing without a size, or one that is no box, changes nothing.
  assert.deepEqual(usedRange(sheet, { styles, drawings: [{ x: 9000, y: 9000, width: 0, height: 10 }, { x: NaN, y: 0, width: 10, height: 10 }] }), { top: 0, left: 0, bottom: 21, right: 9 });
});

test('sheet pages: an empty sheet and a sheet whose rows are all hidden have nothing to print', () => {
  const empty = createSheet({ name: 'S', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/x' });
  assert.equal(usedRange(empty), undefined);
  assert.equal(errorCode(() => paginateSheet(empty)), 'empty');
  const hidden = grid(3, 3);
  hidden.rows.forEach((entry) => (entry.hidden = true));
  assert.equal(errorCode(() => paginateSheet(hidden)), 'empty');
  const columns = grid(3, 3);
  columns.columns = [{ min: 0, max: 2, hidden: true } as Sheet['columns'][number]];
  assert.equal(errorCode(() => paginateSheet(columns)), 'empty');
});

test('sheet pages: hidden rows and columns take no room and never start or end a page', () => {
  const sheet = grid(60, 12);
  // Rows 10..19 and 47..49 are hidden: the first page takes 48 visible rows, which ends at row 60.
  for (let r = 10; r < 20; r++) (sheet.rows.get(r) as Row).hidden = true;
  for (let r = 47; r < 50; r++) (sheet.rows.get(r) as Row).hidden = true;
  // Column B is hidden: the first page holds A, C..K (10 visible), the second L.
  sheet.columns = [{ min: 1, max: 1, hidden: true } as Sheet['columns'][number]];
  const layout = paginateSheet(sheet);
  assert.deepEqual(bands(layout), ['0-59/0-10', '0-59/11-11']);
  assert.equal(layout.pages[0].height, 47 * 20);
  assert.equal(layout.pages[0].width, 10 * 64);
  // Hidden rows at the start and at the end of the print area are not part of a page.
  const edges = grid(6, 2);
  (edges.rows.get(0) as Row).hidden = true;
  (edges.rows.get(5) as Row).hidden = true;
  assert.deepEqual(bands(paginateSheet(edges)), ['1-4/0-1']);
});

test('sheet pages: manual breaks start a page; breaks follow inserted and deleted rows', () => {
  const setup: SheetPageSetup = {
    rowBreaks: [
      { at: 5, min: 0, max: 16383, manual: true },
      { at: 20, min: 0, max: 16383, manual: true },
    ],
    colBreaks: [{ at: 2, min: 0, max: 1048575, manual: true }],
  };
  const sheet = grid(60, 4, setup);
  assert.deepEqual(bands(paginateSheet(sheet)), ['0-4/0-1', '5-19/0-1', '20-59/0-1', '0-4/2-3', '5-19/2-3', '20-59/2-3']);
  // A break at the first row, beyond the sheet or with a wrong position is no break.
  const odd = grid(10, 2, { rowBreaks: [{ at: 0, min: 0, max: 1, manual: true }, { at: NaN, min: 0, max: 1, manual: true }, { at: MAX_ROWS + 5, min: 0, max: 1, manual: true }] });
  assert.deepEqual(bands(paginateSheet(odd)), ['0-9/0-1']);
  // A break at a hidden row takes effect at the next visible one.
  const hidden = grid(10, 2, { rowBreaks: [{ at: 4, min: 0, max: 1, manual: true }] });
  (hidden.rows.get(4) as Row).hidden = true;
  (hidden.rows.get(5) as Row).hidden = true;
  assert.deepEqual(bands(paginateSheet(hidden)), ['0-3/0-1', '6-9/0-1']);
  // Three rows inserted above row 3 move both breaks down; two rows deleted at row 0 move them up again.
  sheet.structuralEdits = [{ kind: 'insert', axis: 'row', at: 2, count: 3 }];
  assert.deepEqual(bands(paginateSheet(sheet)).slice(0, 3), ['0-7/0-1', '8-22/0-1', '23-59/0-1']);
  sheet.structuralEdits.push({ kind: 'delete', axis: 'row', at: 0, count: 2 });
  assert.deepEqual(bands(paginateSheet(sheet)).slice(0, 3), ['0-5/0-1', '6-20/0-1', '21-59/0-1']);
  // An automatic break after 48 rows still happens between manual ones.
  const long = grid(120, 1, { rowBreaks: [{ at: 100, min: 0, max: 0, manual: true }] });
  assert.deepEqual(bands(paginateSheet(long)), ['0-47/0-0', '48-95/0-0', '96-99/0-0', '100-119/0-0']);
});

test('sheet pages: the scale of the page setup changes how much fits; a scale out of range is 100%', () => {
  // At 50% a page holds 20 columns and 97 rows (978.52 / 0.5 = 1957 px).
  const half = paginateSheet(grid(100, 25, { page: { scale: 50 } }));
  assert.equal(half.scale, 0.5);
  assert.deepEqual(bands(half), ['0-96/0-19', '97-99/0-19', '0-96/20-24', '97-99/20-24']);
  assert.equal(half.pages[0].width, 20 * 64 * 0.5);
  // At 200% it holds 5 columns and 24 rows.
  const twice = paginateSheet(grid(30, 6, { page: { scale: 200 } }));
  assert.deepEqual(bands(twice), ['0-23/0-4', '24-29/0-4', '0-23/5-5', '24-29/5-5']);
  for (const scale of [0, 5, 401, NaN, -100, Infinity]) assert.equal(paginateSheet(grid(2, 2, { page: { scale } })).scale, 1, String(scale));
  // The scale is not used while the sheet is fitted to pages.
  assert.equal(paginateSheet(grid(2, 2, { fitToPage: true, page: { scale: 300 } })).scale, 1);
});

test('sheet pages: fit to pages finds the largest whole percent at which the pages are enough, and never enlarges', () => {
  // 25 columns are 1600 px; one page across needs 659.3 / 1600 = 41%.
  const wide = paginateSheet(grid(100, 25, { fitToPage: true, page: { fitToWidth: 1, fitToHeight: 0 } }));
  assert.equal(wide.scale, 0.41);
  // 978.52 / 0.41 = 2386 px: 119 rows on a page, so all 100 on the first.
  assert.deepEqual(bands(wide), ['0-99/0-24']);
  // One page each way: 100 rows are 2000 px, 978.52 / 2000 = 48%, and the width needs 41%.
  const one = paginateSheet(grid(100, 25, { fitToPage: true }));
  assert.equal(one.scale, 0.41);
  assert.equal(one.pages.length, 1);
  // Two pages down, any number across: 2000 px on two pages of 978.52 px is 97%; whole rows need a percent less.
  const tall = paginateSheet(grid(100, 25, { fitToPage: true, page: { fitToWidth: 0, fitToHeight: 2 } }));
  assert.ok(tall.scale <= 0.97 && tall.scale >= 0.9, String(tall.scale));
  assert.equal(new Set(tall.pages.map((page) => page.rows.first)).size, 2);
  // The scale found is the largest: one percent more needs a page more.
  const more = paginateSheet(grid(100, 25, { page: { scale: Math.round(tall.scale * 100) + 1 } }));
  assert.equal(new Set(more.pages.map((page) => page.rows.first)).size, 3);
  // What fits at 100% stays at 100%.
  assert.equal(paginateSheet(grid(5, 5, { fitToPage: true })).scale, 1);
  // What does not fit at 10% is printed at 10% on as many pages as that takes.
  const huge = paginateSheet(grid(1, 400, { fitToPage: true }));
  assert.equal(huge.scale, 0.1);
  assert.equal(huge.pages.length, 4);
  // Manual breaks are ignored while the sheet is fitted.
  const fitted = paginateSheet(grid(10, 2, { fitToPage: true, rowBreaks: [{ at: 5, min: 0, max: 1, manual: true }] }));
  assert.equal(fitted.pages.length, 1);
  // Counts that are no counts read as one page.
  assert.equal(paginateSheet(grid(100, 25, { fitToPage: true, page: { fitToWidth: NaN, fitToHeight: -3 } })).pages.length, 1);
});

test('sheet pages: centring puts the cells in the middle of the room between the margins', () => {
  const layout = paginateSheet(grid(2, 2, { print: { horizontalCentered: true, verticalCentered: true } }));
  const page = layout.pages[0];
  assert.ok(Math.abs(page.x - (67.2 + (659.3 - 128) / 2)) < 1e-6);
  assert.ok(Math.abs(page.y - (72 + (978.52 - 40) / 2)) < 1e-6);
  const across = paginateSheet(grid(2, 2, { print: { horizontalCentered: true } })).pages[0];
  assert.ok(Math.abs(across.y - 72) < 1e-9);
});

test('sheet pages: margins of the file are used; margins that leave no room are replaced by the defaults', () => {
  const layout = paginateSheet(grid(2, 2, { margins: { left: 1, right: 0.5, top: 2, bottom: 0.25, header: 0.125, footer: 0.25 } }));
  assert.deepEqual(layout.margins, { left: 96, right: 48, top: 192, bottom: 24, header: 12, footer: 24 });
  assert.equal(layout.pages[0].x, 96);
  assert.equal(layout.pages[0].y, 192);
  for (const margins of [
    { left: 5, right: 5, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
    { left: 0.7, right: 0.7, top: 20, bottom: 0.75, header: 0.3, footer: 0.3 },
  ]) {
    const wrong = paginateSheet(grid(2, 2, { margins }));
    assert.ok(Math.abs(wrong.margins.left - 67.2) < 1e-9 && Math.abs(wrong.margins.top - 72) < 1e-9);
  }
  // A side that is no number takes its default; the others stay.
  const partly = paginateSheet(grid(2, 2, { margins: { left: NaN, right: -1, top: 1, bottom: 1, header: 0.3, footer: 0.3 } }));
  assert.ok(Math.abs(partly.margins.left - 67.2) < 1e-9 && Math.abs(partly.margins.right - 67.2) < 1e-9);
  assert.equal(partly.margins.top, 96);
});

test('sheet pages: print options, headings take room and the first page number', () => {
  const layout = paginateSheet(grid(100, 25, { print: { headings: true, gridLines: true }, page: { useFirstPageNumber: true, firstPageNumber: 7 } }));
  assert.equal(layout.headings, true);
  assert.equal(layout.gridLines, true);
  assert.equal(layout.firstPageNumber, 7);
  assert.equal(layout.headingWidth, 40);
  assert.equal(layout.headingHeight, 20);
  // 659.3 - 40 = 619.3 px: 9 columns; 978.52 - 20 = 958.52 px: 47 rows.
  assert.equal(bands(layout)[0], '0-46/0-8');
  assert.equal(layout.pages[0].width, 40 + 9 * 64);
  assert.equal(layout.pages[0].height, 20 + 47 * 20);
  // Grid lines are printed only while "gridLinesSet" is not off; the first page number only when it is to be used.
  assert.equal(paginateSheet(grid(2, 2, { print: { gridLines: true, gridLinesSet: false } })).gridLines, false);
  assert.equal(paginateSheet(grid(2, 2, { page: { firstPageNumber: 7 } })).firstPageNumber, undefined);
});

test('sheet pages: print areas replace the used range, each on pages of its own; whole rows and columns end at the used range', () => {
  const sheet = grid(100, 25);
  const one = paginateSheet(sheet, { areas: [{ top: 2, left: 1, bottom: 5, right: 3 }] });
  assert.deepEqual(bands(one), ['2-5/1-3']);
  assert.equal(one.pages[0].width, 3 * 64);
  const two = paginateSheet(sheet, { areas: [{ top: 0, left: 0, bottom: 1, right: 1 }, { top: 50, left: 12, bottom: 51, right: 13 }] });
  assert.deepEqual(bands(two), ['0-1/0-1', '50-51/12-13']);
  assert.deepEqual(two.pages.map((page) => page.area), [0, 1]);
  // "$A:$B": the columns down to the last row in use, not a million rows.
  const columns = paginateSheet(sheet, { areas: [{ top: 0, left: 0, bottom: MAX_ROWS - 1, right: 1 }] });
  assert.deepEqual(bands(columns), ['0-47/0-1', '48-95/0-1', '96-99/0-1']);
  const rows = paginateSheet(sheet, { areas: [{ top: 3, left: 0, bottom: 4, right: MAX_COLS - 1 }] });
  assert.deepEqual(bands(rows), ['3-4/0-9', '3-4/10-19', '3-4/20-24']);
  // An area that is no range is left out; without any left the used range is printed.
  const wrong = paginateSheet(grid(2, 2), { areas: [{ top: 5, left: 0, bottom: 1, right: 0 }, { top: NaN, left: 0, bottom: 1, right: 1 }] });
  assert.deepEqual(bands(wrong), ['0-1/0-1']);
  // An area outside the cells in use is printed all the same: it was asked for.
  assert.deepEqual(bands(paginateSheet(grid(2, 2), { areas: [{ top: 500, left: 30, bottom: 501, right: 31 }] })), ['500-501/30-31']);
});

test('sheet pages: print titles are repeated on the pages that start after them and take their room', () => {
  const sheet = grid(100, 25);
  const layout = paginateSheet(sheet, { titleRows: { first: 0, last: 1 }, titleCols: { first: 0, last: 0 } });
  // The first page has the titles as its own cells; the next ones have 46 rows (2 are taken) and 9 columns.
  assert.deepEqual(bands(layout), [
    '0-47/0-9', '48-93/0-9', '94-99/0-9',
    '0-47/10-18', '48-93/10-18', '94-99/10-18',
    '0-47/19-24', '48-93/19-24', '94-99/19-24',
  ]);
  assert.equal(layout.pages[0].titleRows, undefined);
  assert.equal(layout.pages[0].titleCols, undefined);
  assert.deepEqual(layout.pages[1].titleRows, { first: 0, last: 1 });
  assert.equal(layout.pages[1].titleCols, undefined);
  assert.equal(layout.pages[3].titleRows, undefined);
  assert.deepEqual(layout.pages[3].titleCols, { first: 0, last: 0 });
  assert.deepEqual(layout.pages[4].titleRows, { first: 0, last: 1 });
  assert.equal(layout.pages[4].height, (2 + 46) * 20);
  assert.equal(layout.pages[4].width, (1 + 9) * 64);
  // Titles that leave no room for cells are not repeated.
  const tall = paginateSheet(sheet, { titleRows: { first: 0, last: 60 } });
  assert.equal(tall.pages.every((page) => page.titleRows === undefined), true);
  assert.equal(bands(tall)[1], '48-95/0-9');
  // A title that is no range is none.
  assert.equal(paginateSheet(sheet, { titleRows: { first: 4, last: 2 } }).pages[1].titleRows, undefined);
});

test('sheet pages: a row or column larger than the page gets a page of its own and never loops', () => {
  const sheet = grid(3, 3);
  (sheet.rows.get(1) as Row).height = 409;
  sheet.columns = [{ min: 1, max: 1, width: 255 } as Sheet['columns'][number]];
  const metrics = createSheetMetrics(sheet);
  assert.ok(metrics.cols.size(1) > 659.3);
  const layout = paginateSheet(sheet, { metrics, areas: [{ top: 0, left: 0, bottom: 2, right: 2 }] });
  assert.deepEqual(bands(layout), ['0-2/0-0', '0-2/1-1', '0-2/2-2']);
  // What does not fit is cut off at the margin.
  assert.ok(Math.abs(layout.pages[1].width - 659.3) < 1e-6);
  // The same at a scale at which even the smallest row is larger than the page.
  const tiny = paginateSheet(grid(4, 1, { page: { paperSize: 9, scale: 400 }, margins: { left: 4, right: 4, top: 5.6, bottom: 5.6, header: 0, footer: 0 } }));
  assert.equal(tiny.pages.length, 4);
});

test('sheet pages: a used range of a million rows is refused as too long, quickly', () => {
  const sheet = single(MAX_ROWS - 1, 0);
  const start = Date.now();
  assert.equal(errorCode(() => paginateSheet(sheet)), 'too-long');
  assert.equal(errorCode(() => paginateSheet(sheet, { maxPages: 10 })), 'too-long');
  // Fitted to one page across and any number down it is still too long, and still quick.
  sheet.pageSetup = { fitToPage: true, page: { fitToWidth: 1, fitToHeight: 0 } };
  assert.equal(errorCode(() => paginateSheet(sheet)), 'too-long');
  // Fitted to one page it ends at 10%: 489 rows a page, 2145 pages.
  sheet.pageSetup = { fitToPage: true };
  assert.equal(errorCode(() => paginateSheet(sheet)), 'too-long');
  // A wide used range as well: 16384 columns by 2000 rows.
  const wide = single(1999, MAX_COLS - 1);
  assert.equal(errorCode(() => paginateSheet(wide)), 'too-long');
  assert.ok(Date.now() - start < 2000, `took ${Date.now() - start} ms`);
  // With a limit that is large enough the pages are counted: 1,048,576 rows are 21,846 pages of 48.
  const all = paginateSheet(single(MAX_ROWS - 1, 0), { maxPages: 30000 });
  assert.equal(all.pages.length, Math.ceil(MAX_ROWS / A4_ROWS));
});

test('sheet pages: the limit on pages counts the pages of all areas', () => {
  const sheet = grid(100, 25);
  assert.equal(paginateSheet(sheet, { maxPages: 9 }).pages.length, 9);
  assert.equal(errorCode(() => paginateSheet(sheet, { maxPages: 8 })), 'too-long');
  // A limit that is no number is the default.
  assert.equal(paginateSheet(sheet, { maxPages: NaN }).pages.length, 9);
});

test('print ranges: the print area and the print titles of a sheet from its defined names', () => {
  const names: DefinedName[] = [
    { name: '_xlnm.Print_Area', value: 'Sheet1!$B$2:$D$10', localSheet: 0, hidden: false },
    { name: '_xlnm.Print_Titles', value: 'Sheet1!$A:$B,Sheet1!$1:$3', localSheet: 0, hidden: false },
    { name: '_xlnm.Print_Area', value: "'It''s, a sheet!'!$A$1:$C$3,'It''s, a sheet!'!$E$5", localSheet: 1, hidden: false },
    { name: '_xlnm.Print_Titles', value: 'Other!$2:$2', localSheet: 1, hidden: false },
    { name: '_XLNM.PRINT_AREA', value: 'Third!$C$3:$A$1', localSheet: 2, hidden: false },
    // Not local to a sheet, a formula, a broken reference, and something that is no name at all.
    { name: '_xlnm.Print_Area', value: 'Sheet1!$A$1:$Z$99', hidden: false },
    { name: '_xlnm.Print_Area', value: 'OFFSET(Sheet4!$A$1,0,0,5,5)', localSheet: 3, hidden: false },
    { name: '_xlnm.Print_Area', value: '#REF!', localSheet: 3, hidden: false },
    { name: '_xlnm.Print_Titles', value: 'Sheet4!$XFE:$XFF,Sheet4!$0:$3', localSheet: 3, hidden: false },
    { name: 'Print_Area', value: 'Sheet5!$A$1:$B$2', localSheet: 4, hidden: false },
    { name: 42 as unknown as string, value: 'x', localSheet: 4, hidden: false },
  ];
  assert.deepEqual(printRanges({ definedNames: names }, 0), {
    areas: [{ top: 1, left: 1, bottom: 9, right: 3 }],
    titleCols: { first: 0, last: 1 },
    titleRows: { first: 0, last: 2 },
  });
  assert.deepEqual(printRanges({ definedNames: names }, 1), {
    areas: [{ top: 0, left: 0, bottom: 2, right: 2 }, { top: 4, left: 4, bottom: 4, right: 4 }],
    titleRows: { first: 1, last: 1 },
  });
  assert.deepEqual(printRanges({ definedNames: names }, 2), { areas: [{ top: 0, left: 0, bottom: 2, right: 2 }] });
  assert.deepEqual(printRanges({ definedNames: names }, 3), { areas: [] });
  assert.deepEqual(printRanges({ definedNames: names }, 4), { areas: [] });
  assert.deepEqual(printRanges({ definedNames: [] }, 0), { areas: [] });
  // Whole rows and whole columns as a print area.
  const whole: DefinedName[] = [{ name: '_xlnm.Print_Area', value: 'S!$2:$4,S!$C:$D', localSheet: 0, hidden: false }];
  assert.deepEqual(printRanges({ definedNames: whole }, 0).areas, [
    { top: 1, bottom: 3, left: 0, right: MAX_COLS - 1 },
    { left: 2, right: 3, top: 0, bottom: MAX_ROWS - 1 },
  ]);
  // A name with thousands of parts is read up to a limit.
  const many: DefinedName[] = [{ name: '_xlnm.Print_Area', value: new Array(500).fill('S!$A$1').join(','), localSheet: 0, hidden: false }];
  assert.equal(printRanges({ definedNames: many }, 0).areas.length, 64);
});

test('sheet pages: fitting many large print areas to pages is refused or done quickly, never walked for minutes', () => {
  // Rows of one pixel, then 40,000 tall ones: an area that fits nowhere, behind areas that take long to walk.
  const sheet = single(MAX_ROWS - 1, 0);
  sheet.pageSetup = { fitToPage: true, page: { fitToWidth: 1, fitToHeight: 32767 } };
  const areas = [];
  for (let i = 0; i < 63; i++) areas.push({ top: 40000, left: 0, bottom: MAX_ROWS - 2, right: 0 });
  areas.push({ top: 0, left: 0, bottom: 39999, right: 0 });
  const start = Date.now();
  assert.equal(errorCode(() => paginateSheet(sheet, { areas })), 'too-long');
  // A few areas stay below the limit of what is walked and are still quick.
  errorCode(() => paginateSheet(sheet, { areas: areas.slice(61) }));
  assert.ok(Date.now() - start < 3000, `took ${Date.now() - start} ms`);
});

test('sheet pages: a header or footer distance beyond the paper falls back to the default', () => {
  const layout = paginateSheet(grid(4, 2, { margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 1e9, footer: -1 } }));
  assert.ok(layout.margins.header > 0 && layout.margins.header < layout.paper.height / 2, String(layout.margins.header));
  assert.ok(layout.margins.footer > 0 && layout.margins.footer < layout.paper.height / 2, String(layout.margins.footer));
});
