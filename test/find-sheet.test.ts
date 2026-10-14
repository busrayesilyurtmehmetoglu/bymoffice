// Find and Replace in workbooks: the order of the search, the way round the end of the sheet
// and of the workbook, formulas and values, hidden rows and columns, what replacing writes
// (text stays text, rich text keeps its formatting, numbers and formulas are read anew), what
// it passes over, "Replace all" as one step, and a large sheet searched in slices.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { LOCALE_TR } from '../src/format/locale';
import { createSheet, emptyStyleSheet, getCell, type Cell, type Sheet, type Workbook } from '../src/model/sheet/types';
import { ensureStyle } from '../src/ops/format';
import { type CellReplacement, MAX_REPLACE_CELLS, ReplaceAll, SheetSearch, cellSearchText, compileSheetFind, findNextCell, replaceAllOps, replaceInCell, searchStart, type CellAddress, type SheetFindOptions } from '../src/ops/find';
import { History, applyOp, type Op } from '../src/ops/ops';

type Value = number | string | boolean | null | Cell;

function cellOf(value: Value): Cell | undefined {
  if (value === null) return undefined;
  if (typeof value === 'number') return { kind: 'number', value, style: 0 };
  if (typeof value === 'string') return value.charAt(0) === '=' ? { kind: 'blank', value: null, style: 0, formula: { text: value.slice(1), type: 'normal' } } : { kind: 'string', value, style: 0 };
  if (typeof value === 'boolean') return { kind: 'boolean', value, style: 0 };
  return value;
}

function sheetOf(grid: Value[][], name = 'Sheet1', id = 1): Sheet {
  const sheet = createSheet({ name, sheetId: id, state: 'visible', kind: 'worksheet', part: `/xl/worksheets/sheet${id}.xml` });
  grid.forEach((values, r) => {
    const cells: Cell[] = [];
    values.forEach((value, c) => {
      const cell = cellOf(value);
      if (cell) {
        cells[c] = cell;
        sheet.colCount = Math.max(sheet.colCount, c + 1);
      }
    });
    if (cells.length > 0) {
      sheet.rows.set(r, { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells });
      sheet.rowCount = r + 1;
    }
  });
  return sheet;
}

const bookOf = (...sheets: Sheet[]): Workbook => ({ sheets, styles: emptyStyleSheet(), sharedStrings: [], definedNames: [], date1904: false, activeSheet: 0 });
const at = (sheet: number, row: number, col: number): CellAddress => ({ sheet, row, col });
const name = (hit: CellAddress | undefined) => (hit ? `${hit.sheet}:${String.fromCharCode(65 + hit.col)}${hit.row + 1}` : 'none');

/** Every cell a search reaches from a place, in order, until it comes round to the first again. */
function walkAll(workbook: Workbook, text: string, from: CellAddress, options: SheetFindOptions = {}, backward = false): string[] {
  const out: string[] = [];
  let place = from;
  for (let i = 0; i < 100; i++) {
    const hit = findNextCell(workbook, text, place, options, backward);
    if (!hit) break;
    const label = name(hit);
    if (out.indexOf(label) >= 0) break;
    out.push(label);
    place = hit;
  }
  return out;
}

/** What replacing wrote into a cell; throws when it wrote nothing. */
function written(result: CellReplacement): { change: { row: number; col: number; cell: Cell | undefined }; count: number } {
  if (!result || !('change' in result)) throw new Error('nothing was written');
  return result;
}

const apply = (workbook: Workbook, ops: Op[]): Op[] => ops.map((op) => applyOp(workbook, op)).reverse();

test('find in cells: by rows from the active cell, round the end of the sheet, and backwards', () => {
  const sheet = sheetOf([
    ['x', 'a', 'x'],
    ['x', null, 'b'],
    [null, 'x', null],
  ]);
  const workbook = bookOf(sheet);
  assert.deepEqual(walkAll(workbook, 'x', at(0, 0, 0)), ['0:C1', '0:A2', '0:B3', '0:A1']);
  assert.deepEqual(walkAll(workbook, 'x', at(0, 0, 0), {}, true), ['0:B3', '0:A2', '0:C1', '0:A1']);
  // From a cell that has no match itself, and from the last cell.
  assert.equal(name(findNextCell(workbook, 'x', at(0, 1, 1))), '0:B3');
  assert.equal(name(findNextCell(workbook, 'x', at(0, 2, 1))), '0:A1');
  // The only match is found again from itself, all the way round.
  assert.equal(name(findNextCell(workbook, 'b', at(0, 1, 2))), '0:C2');
  assert.equal(findNextCell(workbook, 'nothing', at(0, 0, 0)), undefined);
  assert.equal(findNextCell(workbook, '', at(0, 0, 0)), undefined);
});

test('find in cells: the workbook, sheet after sheet and back to the first; other sheets only when asked', () => {
  const workbook = bookOf(sheetOf([['k', null], [null, 'k']], 'One', 1), sheetOf([['no'], ['k']], 'Two', 2), sheetOf([[null, 'k']], 'Three', 3));
  const chart = createSheet({ name: 'Chart', sheetId: 4, state: 'visible', kind: 'chartsheet', part: '/xl/chartsheets/sheet1.xml' });
  workbook.sheets.splice(2, 0, chart);
  assert.deepEqual(walkAll(workbook, 'k', at(1, 0, 0), { within: 'workbook' }), ['1:A2', '3:B1', '0:A1', '0:B2']);
  assert.deepEqual(walkAll(workbook, 'k', at(1, 0, 0), { within: 'workbook' }, true), ['0:B2', '0:A1', '3:B1', '1:A2']);
  assert.deepEqual(walkAll(workbook, 'k', at(1, 0, 0)), ['1:A2']);
  // A hidden sheet is searched only when the search starts on it.
  workbook.sheets[3].state = 'hidden';
  assert.deepEqual(walkAll(workbook, 'k', at(1, 0, 0), { within: 'workbook' }), ['1:A2', '0:A1', '0:B2']);
  assert.deepEqual(walkAll(workbook, 'k', at(3, 0, 0), { within: 'workbook' }), ['3:B1', '0:A1', '0:B2', '1:A2']);
});

test('find in cells: only within a range when one is given', () => {
  const workbook = bookOf(sheetOf([['q', 'q', 'q'], ['q', 'q', 'q'], ['q', 'q', 'q']]));
  assert.deepEqual(walkAll(workbook, 'q', at(0, 1, 1), { range: { top: 1, left: 1, bottom: 2, right: 2 } }), ['0:C2', '0:B3', '0:C3', '0:B2']);
});

test('find in cells: formulas by default, values when asked; numbers as typed, text without its apostrophe', () => {
  const workbook = bookOf(sheetOf([[1234.5, '=ROUND(SUM(A1:A3),2)', '123', true]]));
  const styles = workbook.styles;
  const sheet = workbook.sheets[0];
  const row = sheet.rows.get(0)!;
  row.cells[0] = { ...row.cells[0], style: ensureStyle(styles, 0, { numFmt: '#,##0.00' }) };
  row.cells[1] = { ...row.cells[1], kind: 'number', value: 1234.5 };
  assert.equal(cellSearchText(workbook, sheet, 0, 0, row.cells[0]), '1234.5');
  assert.equal(cellSearchText(workbook, sheet, 0, 0, row.cells[0], 'values'), '1,234.50');
  assert.equal(cellSearchText(workbook, sheet, 0, 1, row.cells[1]), '=ROUND(SUM(A1:A3),2)');
  assert.equal(cellSearchText(workbook, sheet, 0, 2, row.cells[2]), '123');
  assert.equal(cellSearchText(workbook, sheet, 0, 3, row.cells[3]), 'TRUE');
  // In Turkish: the decimal comma, the separator of arguments and the names of truth, as the formula bar shows them.
  assert.equal(cellSearchText(workbook, sheet, 0, 0, row.cells[0], 'formulas', LOCALE_TR), '1234,5');
  assert.equal(cellSearchText(workbook, sheet, 0, 1, row.cells[1], 'formulas', LOCALE_TR), '=ROUND(SUM(A1:A3);2)');
  assert.equal(cellSearchText(workbook, sheet, 0, 3, row.cells[3], 'formulas', LOCALE_TR), 'DOĞRU');

  assert.deepEqual(walkAll(workbook, '1,234', at(0, 0, 3)), []);
  assert.deepEqual(walkAll(workbook, '1,234', at(0, 0, 3), { lookIn: 'values' }), ['0:A1']);
  assert.deepEqual(walkAll(workbook, '1234.5', at(0, 0, 3), { lookIn: 'values' }), ['0:B1']);
  assert.deepEqual(walkAll(workbook, 'sum', at(0, 0, 3)), ['0:B1']);
  assert.deepEqual(walkAll(workbook, 'sum', at(0, 0, 3), { lookIn: 'values' }), []);
  assert.deepEqual(walkAll(workbook, '123', at(0, 0, 3), { entireCell: true }), ['0:C1']);
});

test('find in cells: a shared formula is searched as the formula it stands for', () => {
  const sheet = sheetOf([
    [1, { kind: 'number', value: 2, style: 0, formula: { text: 'A1*2', type: 'shared', si: 0, ref: 'B1:B2' } }],
    [3, { kind: 'number', value: 6, style: 0, formula: { text: '', type: 'shared', si: 0 } }],
  ]);
  const workbook = bookOf(sheet);
  assert.deepEqual(walkAll(workbook, 'A2*2', at(0, 0, 0)), ['0:B2']);
});

test('find in cells: hidden rows and columns are searched for formulas, passed over for values', () => {
  const sheet = sheetOf([['h', 'h'], ['h', 'v']]);
  sheet.rows.get(1)!.hidden = true;
  sheet.columns.push({ min: 2, max: 2, customWidth: false, hidden: true, outlineLevel: 0, collapsed: false });
  const workbook = bookOf(sheet);
  assert.deepEqual(walkAll(workbook, 'h', at(0, 0, 0)), ['0:B1', '0:A2', '0:A1']);
  assert.deepEqual(walkAll(workbook, 'h', at(0, 0, 0), { lookIn: 'values' }), ['0:A1']);
});

test('find in cells: Turkish case folding follows the locale', () => {
  const workbook = bookOf(sheetOf([['İzmir', 'IZMIR', 'ızmır']]));
  assert.deepEqual(walkAll(workbook, 'izmir', at(0, 0, 2), { locale: LOCALE_TR }), ['0:A1']);
  assert.deepEqual(walkAll(workbook, 'ızmır', at(0, 0, 2), { locale: LOCALE_TR }), ['0:B1', '0:C1']);
  assert.deepEqual(walkAll(workbook, 'izmir', at(0, 0, 2)), ['0:B1']);
  assert.deepEqual(walkAll(workbook, 'IZMIR', at(0, 0, 2), { matchCase: true }), ['0:B1']);
});

test('replace in a cell: text stays text, numbers and formulas are read anew', () => {
  const workbook = bookOf(sheetOf([['order 12', 1205, '=SUM(A1:A2)', '0012', 'abc']]));
  const find = (text: string, options: SheetFindOptions = {}) => compileSheetFind(text, options)!;
  // Text: "12" in text stays text, also when the result looks like a number.
  let change = written(replaceInCell(workbook, 0, 0, 0, find('order '), ''));
  assert.deepEqual(change.change.cell, { kind: 'string', value: '12', style: 0 });
  change = written(replaceInCell(workbook, 0, 0, 3, find('00'), '99'));
  assert.deepEqual(change.change.cell, { kind: 'string', value: '9912', style: 0 });
  // A number as it is typed.
  change = written(replaceInCell(workbook, 0, 0, 1, find('12'), '34'));
  assert.equal(change.count, 1);
  assert.deepEqual([change.change.cell!.kind, change.change.cell!.value], ['number', 3405]);
  // A formula, which is calculated anew (it has no value until then).
  change = written(replaceInCell(workbook, 0, 0, 2, find('sum'), 'AVERAGE'));
  assert.equal(change.change.cell!.formula!.text, 'AVERAGE(A1:A2)');
  // In Turkish the formula is read the way it is typed there: the Turkish name of a function is understood.
  change = written(replaceInCell(workbook, 0, 0, 2, find('sum', { locale: LOCALE_TR }), 'ORTALAMA', LOCALE_TR));
  assert.equal(change.change.cell!.formula!.text, 'AVERAGE(A1:A2)');
  // A formula that cannot be read once replaced is passed over, and counted.
  assert.deepEqual(replaceInCell(workbook, 0, 0, 2, find(')'), '(('), { skipped: true, count: 1 });
  // Text replaced by nothing at all empties the cell.
  change = written(replaceInCell(workbook, 0, 0, 4, find('abc'), ''));
  assert.equal(change.change.cell, undefined);
  // A cell without the text gives nothing.
  assert.equal(replaceInCell(workbook, 0, 0, 4, find('zzz'), 'y'), undefined);
  assert.equal(replaceInCell(workbook, 0, 5, 5, find('a'), 'y'), undefined);
});

test('replace in a cell: rich text keeps the formatting of what was not replaced', () => {
  const bold = { name: 'Calibri', size: 11, bold: true, italic: false, strike: false, underline: 'none' as const };
  const runs = [{ text: 'Hello ' }, { text: 'big', font: bold }, { text: ' world' }];
  const workbook = bookOf(sheetOf([[{ kind: 'string', value: 'Hello big world', style: 0, runs, si: 0 }]]));
  const find = (text: string) => compileSheetFind(text)!;
  let result = written(replaceInCell(workbook, 0, 0, 0, find('big'), 'small'));
  assert.deepEqual(result.change.cell, { kind: 'string', value: 'Hello small world', style: 0, runs: [{ text: 'Hello ' }, { text: 'small', font: bold }, { text: ' world' }] });
  // A match across runs: the replacement takes the formatting of the run it starts in, the rest keeps theirs.
  result = written(replaceInCell(workbook, 0, 0, 0, find('o b'), '-'));
  assert.deepEqual(result.change.cell!.runs, [{ text: 'Hell-' }, { text: 'ig', font: bold }, { text: ' world' }]);
  // A run that is replaced as a whole goes; the text is what the runs say.
  result = written(replaceInCell(workbook, 0, 0, 0, find('big'), ''));
  assert.deepEqual(result.change.cell!.runs, [{ text: 'Hello ' }, { text: ' world' }]);
  assert.equal(result.change.cell!.value, 'Hello  world');
});

test('replace in a cell: a part of an array formula and text that grows too long are passed over', () => {
  const sheet = sheetOf([
    [{ kind: 'number', value: 2, style: 0, formula: { text: 'A3:A4*2', type: 'array', ref: 'A1:A2' } }],
    [{ kind: 'number', value: 4, style: 0 }],
    ['x'.repeat(30000)],
  ]);
  const workbook = bookOf(sheet);
  assert.deepEqual(replaceInCell(workbook, 0, 0, 0, compileSheetFind('2')!, '3'), { skipped: true, count: 1 });
  assert.deepEqual(replaceInCell(workbook, 0, 1, 0, compileSheetFind('4')!, '5'), { skipped: true, count: 1 });
  assert.deepEqual(replaceInCell(workbook, 0, 2, 0, compileSheetFind('xx')!, 'xxxxx'), { skipped: true, count: 15000 });
});

test('replace all: one setCells per sheet, as one undo step; rows the filter hides are left alone', () => {
  const one = sheetOf([['cat', 'cat cat'], ['dog', 'cat'], ['cat', null]], 'One', 1);
  // Row 3 is hidden by the filter of the sheet; row 2 is hidden by hand.
  one.autoFilter = { range: { top: 0, left: 0, bottom: 2, right: 1 }, columns: [{ colId: 0, kind: 'values', values: ['cat', 'dog'] }] };
  one.rows.get(2)!.hidden = true;
  const two = sheetOf([['Cat']], 'Two', 2);
  const workbook = bookOf(one, two);
  const result = replaceAllOps(workbook, 'cat', 'bird', { within: 'workbook' });
  assert.equal(result.replaced, 5);
  assert.equal(result.cells, 4);
  assert.equal(result.skipped, 0);
  assert.deepEqual(result.ops.map((op) => (op.type === 'setCells' ? [op.sheet, op.cells.length] : op.type)), [[0, 3], [1, 1]]);
  const history = new History(workbook);
  history.apply(result.ops, undefined, undefined);
  assert.deepEqual(['A1', 'B1', 'B2', 'A3'].map((ref, i) => getCell(one, [0, 0, 1, 2][i], [0, 1, 1, 0][i])!.value), ['bird', 'bird bird', 'bird', 'cat']);
  assert.equal(getCell(two, 0, 0)!.value, 'bird');
  history.undo();
  assert.equal(getCell(one, 0, 1)!.value, 'cat cat');
  assert.equal(getCell(two, 0, 0)!.value, 'Cat');
  // Only the sheet the search starts on, unless asked.
  assert.equal(replaceAllOps(workbook, 'cat', 'x', {}, searchStart(1)).cells, 1);
});

test('replace all: whole cells only, and nothing when there is no match', () => {
  const workbook = bookOf(sheetOf([['a', 'a b', 'A']]));
  const result = replaceAllOps(workbook, 'a', 'z', { entireCell: true });
  apply(workbook, result.ops);
  assert.deepEqual([0, 1, 2].map((c) => getCell(workbook.sheets[0], 0, c)!.value), ['z', 'a b', 'z']);
  assert.deepEqual(replaceAllOps(workbook, 'q', 'z'), { ops: [], replaced: 0, cells: 0, skipped: 0 });
});

test('replace all: more cells than one step takes is refused as a whole', () => {
  const sheet = sheetOf([]);
  const cells: Cell[] = [];
  for (let c = 0; c < 1000; c++) cells[c] = { kind: 'string', value: 'm', style: 0 };
  for (let r = 0; r <= MAX_REPLACE_CELLS / 1000; r++) sheet.rows.set(r, { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: cells.slice() });
  const workbook = bookOf(sheet);
  const result = replaceAllOps(workbook, 'm', 'n');
  assert.equal(result.tooMany, true);
  assert.equal(result.ops.length, 0);
});

test('find in cells: 100,000 rows are searched in slices, none of which takes long', () => {
  const sheet = sheetOf([]);
  for (let r = 0; r < 100000; r++) sheet.rows.set(r, { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [{ kind: 'string', value: `row ${r}`, style: 0 }, { kind: 'number', value: r, style: 0 }] });
  sheet.rowCount = 100000;
  sheet.colCount = 2;
  const workbook = bookOf(sheet);
  const find = compileSheetFind('needle')!;
  const search = new SheetSearch(workbook, find, {}, searchStart(0));
  let slices = 0;
  let longest = 0;
  const started = Date.now();
  for (;;) {
    const t = Date.now();
    const hit = search.step(4096);
    longest = Math.max(longest, Date.now() - t);
    slices++;
    if (hit !== undefined) {
      assert.equal(hit, null);
      break;
    }
  }
  const total = Date.now() - started;
  // About 200,000 cells, 4,096 at a time.
  assert.ok(slices >= 45 && slices <= 60, String(slices));
  // Guards against a slice that does everything at once, not measures of speed (see the report of the task for those).
  assert.ok(longest < 1000, `longest slice ${longest} ms, all ${total} ms`);
  // The row found last is found from the first.
  assert.equal(name(findNextCell(workbook, 'row 99999', searchStart(0))), '0:A100000');
  // Replace all goes in slices too.
  const job = new ReplaceAll(workbook, compileSheetFind('row 9999')!, 'R', {}, searchStart(0));
  let steps = 0;
  while (!job.step(4096)) steps++;
  assert.ok(steps > 10);
  assert.equal(job.result().cells, 11);
});
