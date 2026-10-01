import assert from 'node:assert/strict';
import { test } from 'node:test';
import { emptyStyleSheet, getCell, open, resolveCellStyle, resolveColor, type Workbook } from '../src/index';
import { setCellInput } from '../src/ops/commands';
import { ensureStyle, formatAt, formatRange, mergeRange, setColumnWidth, setColumnsHidden, setRowHeight, setRowsHidden, unmergeRange } from '../src/ops/format';
import { History, applyOp } from '../src/ops/ops';
import { sampleWorkbook } from './fixtures';

const load = async (): Promise<Workbook> => (await open(await sampleWorkbook())).workbook();
const range = (top: number, left: number, bottom: number, right: number) => ({ top, left, bottom, right });

/** Model state that formatting may touch, as text. */
function picture(wb: Workbook): string {
  return JSON.stringify(
    wb.sheets.map((sheet) => ({
      columns: sheet.columns,
      merges: sheet.merges,
      rows: [...sheet.rows.entries()].sort((a, b) => a[0] - b[0]).map(([i, row]) => [i, row.height, row.customHeight, row.hidden, row.style, row.cells.map((c, col) => [col, c.kind, c.value, c.style])]),
    })),
  );
}

test('format: font, fill, alignment and number format patches', async () => {
  const wb = await load();
  const styles = wb.styles;
  const counts = [styles.fonts.length, styles.fills.length, styles.cellXfs.length];

  const bold = ensureStyle(styles, 0, { bold: true });
  assert.equal(ensureStyle(styles, 0, { bold: true }), bold, 'the same patch finds the same format');
  assert.equal(ensureStyle(styles, bold, { bold: true }), bold, 'a patch that changes nothing keeps the format');
  assert.equal(ensureStyle(styles, bold, { bold: false }), 0, 'undoing a patch finds the original format');
  assert.equal(styles.fonts.length, counts[0] + 1);
  assert.equal(styles.modified, true);

  const all = ensureStyle(styles, 0, {
    italic: true,
    underline: 'single',
    strike: true,
    fontName: 'Arial',
    fontSize: 14,
    color: { rgb: 'FFFF0000' },
    fill: { rgb: 'FF00FF00' },
    horizontal: 'center',
    vertical: 'top',
    wrapText: true,
    numFmt: '0.00%',
  });
  const resolved = resolveCellStyle(styles, all);
  assert.deepEqual(
    [resolved.font.italic, resolved.font.underline, resolved.font.strike, resolved.font.name, resolved.font.size, resolved.font.scheme],
    [true, 'single', true, 'Arial', 14, undefined],
  );
  assert.equal(resolveColor(resolved.font.color, styles), '#FF0000');
  assert.equal(resolved.fill.type === 'pattern' && resolved.fill.pattern, 'solid');
  assert.equal(resolved.fill.type === 'pattern' && resolveColor(resolved.fill.fg, styles), '#00FF00');
  assert.deepEqual(resolved.alignment, { horizontal: 'center', vertical: 'top', wrapText: true });
  assert.equal(resolved.numFmt, '0.00%');

  // null removes: no colour, no fill, default alignment.
  const cleared = resolveCellStyle(styles, ensureStyle(styles, all, { color: null, fill: null, horizontal: null, vertical: null, wrapText: false }));
  assert.equal(cleared.font.color, undefined);
  assert.equal(cleared.fill.type === 'pattern' && cleared.fill.pattern, 'none');
  assert.deepEqual(cleared.alignment, {});
  // "No fill" is the reserved first fill, not a new record.
  assert.equal(styles.fills.filter((fill) => fill.type === 'pattern' && fill.pattern === 'none').length, 1);

  // A style sheet without the two reserved fills gets them before its first own fill.
  const bare = emptyStyleSheet();
  const filled = resolveCellStyle(bare, ensureStyle(bare, 0, { fill: { rgb: 'FF112233' } }));
  assert.deepEqual(bare.fills.slice(0, 2).map((fill) => fill.type === 'pattern' && fill.pattern), ['none', 'gray125']);
  assert.equal(bare.fills.length, 3);
  assert.equal(filled.fill.type === 'pattern' && resolveColor(filled.fill.fg, bare), '#112233');

  // The records other cells use are never changed in place.
  assert.equal(styles.fonts[0].bold, false);
  assert.equal(resolveCellStyle(styles, 1).font.bold, true);
});

test('format: a range is formatted cell by cell, empty cells included, and undone exactly', async () => {
  const wb = await load();
  const history = new History(wb);
  const before = picture(wb);
  const sheet = wb.sheets[0];
  history.apply(formatRange(wb, 0, range(1, 0, 5, 1), { bold: true, fill: { rgb: 'FFFFFF00' } }), undefined, undefined);

  // B2 had the currency format: it keeps it and gains the rest.
  const b2 = formatAt(wb, sheet, 1, 1);
  assert.deepEqual([b2.font.bold, b2.numFmt], [true, '#,##0.00\\ "₺"']);
  assert.equal(getCell(sheet, 1, 1)?.value, 12.5);
  // A6 was empty: it is now a blank cell that carries the format.
  assert.deepEqual([getCell(sheet, 5, 0)?.kind, formatAt(wb, sheet, 5, 0).font.bold], ['blank', true]);
  assert.equal(formatAt(wb, sheet, 6, 0).font.bold, false);

  history.undo();
  assert.equal(picture(wb), before);
  history.redo();
  assert.equal(formatAt(wb, sheet, 1, 0).font.bold, true);
});

test('format: borders follow the position of each cell in the range', async () => {
  const wb = await load();
  const sheet = wb.sheets[0];
  for (const op of formatRange(wb, 0, range(10, 1, 12, 3), { border: { preset: 'outline', edge: { style: 'medium', color: { rgb: 'FF0000FF' } } } })) applyOp(wb, op);
  const edges = (row: number, col: number) => {
    const border = formatAt(wb, sheet, row, col).border;
    return ['top', 'bottom', 'left', 'right'].filter((side) => border[side as 'top'] !== undefined).join(',');
  };
  assert.deepEqual([edges(10, 1), edges(10, 2), edges(11, 2), edges(12, 3), edges(11, 1)], ['top,left', 'top', '', 'bottom,right', 'left']);
  assert.equal(formatAt(wb, sheet, 10, 1).border.top?.style, 'medium');

  for (const op of formatRange(wb, 0, range(10, 1, 12, 3), { border: { preset: 'all' } })) applyOp(wb, op);
  assert.equal(edges(11, 2), 'top,bottom,left,right');
  assert.equal(formatAt(wb, sheet, 11, 2).border.top?.style, 'thin');
  for (const op of formatRange(wb, 0, range(10, 1, 12, 3), { border: { preset: 'bottom' } })) applyOp(wb, op);
  assert.equal(edges(11, 2), 'top,bottom,left,right');
  for (const op of formatRange(wb, 0, range(10, 1, 12, 3), { border: { preset: 'none' } })) applyOp(wb, op);
  assert.deepEqual([edges(10, 1), edges(12, 3)], ['', '']);
});

test('format: whole columns and rows are formatted as such', async () => {
  const wb = await load();
  const history = new History(wb);
  const before = picture(wb);
  const sheet = wb.sheets[0];

  history.apply(formatRange(wb, 0, range(0, 1, 1048575, 2), { italic: true }), undefined, undefined);
  // B and C: the sample has column settings for A and for C:D; they are split around the span.
  assert.deepEqual(sheet.columns.map((c) => [c.min, c.max, c.style !== undefined]), [[0, 0, false], [1, 1, true], [2, 2, true], [3, 3, false]]);
  assert.equal(sheet.columns[2].hidden, true, 'other column settings survive');
  assert.equal(sheet.rows.size < 20, true, 'no cells are created down the column');
  assert.equal(formatAt(wb, sheet, 500, 1).font.italic, true);
  assert.equal(formatAt(wb, sheet, 1, 1).font.italic, true, 'cells already in the column follow');
  assert.equal(formatAt(wb, sheet, 1, 1).numFmt, '#,##0.00\\ "₺"');
  assert.equal(formatAt(wb, sheet, 500, 0).font.italic, false);

  history.apply(formatRange(wb, 0, range(20, 0, 21, 16383), { fill: { rgb: 'FFCCCCCC' } }), undefined, undefined);
  assert.deepEqual([sheet.rows.get(20)?.style !== undefined, sheet.rows.get(20)?.cells.length], [true, 0]);
  const cornerFill = formatAt(wb, sheet, 20, 9000).fill;
  assert.equal(cornerFill.type === 'pattern' && cornerFill.pattern, 'solid');

  assert.throws(() => formatRange(wb, 0, range(0, 0, 100000, 100), { bold: true }), /Too many cells/);
  while (history.undo());
  assert.equal(picture(wb), before);
});

test('format: column widths, row heights and hiding', async () => {
  const wb = await load();
  const history = new History(wb);
  const before = picture(wb);
  const sheet = wb.sheets[0];

  history.apply(setColumnWidth(wb, 0, 5, 7, 20), undefined, undefined);
  assert.deepEqual(sheet.columns.filter((c) => c.min === 5).map((c) => [c.min, c.max, c.width, c.customWidth]), [[5, 7, 20, true]]);
  assert.equal(sheet.columnsModified, true);
  history.apply(setColumnWidth(wb, 0, 6, 6, undefined), undefined, undefined);
  assert.deepEqual(sheet.columns.filter((c) => c.min >= 5).map((c) => [c.min, c.max, c.width]), [[5, 5, 20], [6, 6, undefined], [7, 7, 20]]);
  history.apply(setColumnsHidden(wb, 0, 2, 3, false), undefined, undefined);
  assert.deepEqual(sheet.columns.filter((c) => c.min === 2).map((c) => [c.max, c.hidden, c.width]), [[3, false, 9]]);

  history.apply(setRowHeight(0, 1, 2, 30), undefined, undefined);
  // Row 3 of the sample is hidden; giving it a height shows it, as dragging its border would.
  assert.deepEqual([1, 2].map((r) => [sheet.rows.get(r)?.height, sheet.rows.get(r)?.customHeight, sheet.rows.get(r)?.hidden]), [[30, true, false], [30, true, false]]);
  assert.ok(sheet.changedRows?.has(1));
  // Zero size hides and keeps the size to come back with.
  history.apply(setColumnWidth(wb, 0, 5, 5, 0), undefined, undefined);
  assert.deepEqual(sheet.columns.filter((c) => c.min === 5).map((c) => [c.hidden, c.width]), [[true, 20]]);
  history.apply(setRowHeight(0, 1, 1, 0), undefined, undefined);
  assert.deepEqual([sheet.rows.get(1)?.hidden, sheet.rows.get(1)?.height], [true, 30]);
  history.apply(setRowsHidden(0, 40, 41, true), undefined, undefined);
  assert.equal(sheet.rows.get(40)?.hidden, true);
  history.apply(setRowsHidden(0, 40, 41, false), undefined, undefined);
  assert.equal(sheet.rows.has(40), false, 'a row with nothing of its own is not kept');
  history.apply(setRowHeight(0, 0, 0, undefined), undefined, undefined);
  assert.deepEqual([sheet.rows.get(0)?.height, sheet.rows.get(0)?.customHeight], [undefined, false]);

  while (history.undo());
  assert.equal(picture(wb), before);
});

test('format: merging keeps the top-left cell and replaces merges it overlaps', async () => {
  const wb = await load();
  const history = new History(wb);
  const before = picture(wb);
  const sheet = wb.sheets[0];
  history.apply(setCellInput(wb, 0, 11, 2, 'kaybolur'), undefined, undefined);
  history.apply(setCellInput(wb, 0, 10, 1, 'kalır'), undefined, undefined);

  history.apply(mergeRange(wb, 0, range(10, 1, 12, 3)), undefined, undefined);
  assert.deepEqual(sheet.merges[sheet.merges.length - 1], range(10, 1, 12, 3));
  assert.equal(sheet.mergesModified, true);
  assert.deepEqual([getCell(sheet, 10, 1)?.value, getCell(sheet, 11, 2)], ['kalır', undefined]);

  // The sample merges A5:C6; a merge that overlaps it takes its place.
  history.apply(mergeRange(wb, 0, range(4, 2, 7, 2)), undefined, undefined);
  assert.deepEqual(sheet.merges.map((m) => [m.top, m.left, m.bottom, m.right]), [[10, 1, 12, 3], [4, 2, 7, 2]]);
  assert.deepEqual(mergeRange(wb, 0, range(3, 3, 3, 3)), []);
  assert.throws(() => mergeRange(wb, 0, range(0, 0, 1048575, 1)), /cannot be merged/);

  history.apply(unmergeRange(wb, 0, range(11, 2, 11, 2)), undefined, undefined);
  assert.deepEqual(sheet.merges.map((m) => [m.top, m.left]), [[4, 2]]);
  assert.deepEqual(unmergeRange(wb, 0, range(30, 0, 31, 1)), []);

  while (history.undo());
  assert.equal(picture(wb), before);
});

test('format: borders around whole columns go on the first and the last column only', async () => {
  const wb = await load();
  const sheet = wb.sheets[1];
  for (const op of formatRange(wb, 1, range(0, 1, 1048575, 3), { border: { preset: 'outline' } })) applyOp(wb, op);
  const sides = (col: number) => {
    const border = formatAt(wb, sheet, 100, col).border;
    return ['left', 'right'].filter((side) => border[side as 'left'] !== undefined).join(',');
  };
  assert.deepEqual([sides(0), sides(1), sides(2), sides(3), sides(4)], ['', 'left', '', 'right', '']);
  assert.deepEqual(sheet.columns.map((c) => [c.min, c.max]), [[1, 1], [2, 2], [3, 3]]);

  // One column: both sides on it.
  for (const op of formatRange(wb, 1, range(0, 6, 1048575, 6), { border: { preset: 'outline' } })) applyOp(wb, op);
  assert.equal(sides(6), 'left,right');
});

test('format: a row with a format of its own follows a column format across every column', async () => {
  const wb = await load();
  const sheet = wb.sheets[1];
  for (const op of setCellInput(wb, 1, 2, 0, 'x')) applyOp(wb, op);
  for (const op of formatRange(wb, 1, range(2, 0, 2, 16383), { italic: true })) applyOp(wb, op);
  // Select everything and make it bold: the styled row must be bold at its far end too.
  for (const op of formatRange(wb, 1, range(0, 0, 1048575, 16383), { bold: true })) applyOp(wb, op);
  for (const col of [0, 255, 256, 5000, 16383]) {
    const font = formatAt(wb, sheet, 2, col).font;
    assert.deepEqual([font.bold, font.italic], [true, true], `column ${col}`);
  }
  assert.equal(formatAt(wb, sheet, 0, 16383).font.bold, true);
  // A few columns only: the crossing cells are formatted one by one, the row keeps its own format.
  for (const op of formatRange(wb, 1, range(0, 3, 1048575, 4), { strike: true })) applyOp(wb, op);
  assert.deepEqual([formatAt(wb, sheet, 2, 3).font.strike, formatAt(wb, sheet, 2, 5).font.strike, formatAt(wb, sheet, 2, 3).font.italic], [true, false, true]);
});
