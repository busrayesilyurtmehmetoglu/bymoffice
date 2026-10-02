// The filter, second round: case folding by locale, what deleting rows and
// columns does to a filter in use, and the rows a filter hides staying out of
// what is done to a selection.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { LOCALE_EN, LOCALE_TR } from '../src/format/locale';
import { FILTER_NAME, createSheet, emptyStyleSheet, getCell, type AutoFilter, type Cell, type Row, type Sheet, type Workbook } from '../src/model/sheet/types';
import { clearRange, copyRange, pasteAt } from '../src/ops/commands';
import { clearFilter, filterSummary, filterValues, setAutoFilter, setFilterColumn, type FilterCriterion } from '../src/ops/filter';
import * as filterModule from '../src/ops/filter';
import { History, applyOp, type Op } from '../src/ops/ops';

// Added by this round; read through the module so that the tests of what was there before still run without them.
const { deleteColumns, deleteRows, rowsHiddenByFilter } = filterModule as unknown as {
  deleteColumns(workbook: Workbook, sheet: number, at: number, count: number, options?: { locale?: typeof LOCALE_TR }): Op[];
  deleteRows(workbook: Workbook, sheet: number, at: number, count: number): Op[];
  rowsHiddenByFilter(sheet: Sheet, top: number, bottom: number): number[];
};

type Value = number | string | boolean | null | { f: string };

function sheetOf(grid: Value[][], name = 'Data'): Sheet {
  const sheet = createSheet({ name, sheetId: 1, state: 'visible', kind: 'worksheet', part: '/xl/worksheets/sheet1.xml' });
  grid.forEach((values, r) => {
    const cells: Cell[] = [];
    values.forEach((value, c) => {
      if (value === null) return;
      cells[c] =
        typeof value === 'number'
          ? { kind: 'number', value, style: 0 }
          : typeof value === 'boolean'
            ? { kind: 'boolean', value, style: 0 }
            : typeof value === 'string'
              ? { kind: 'string', value, style: 0 }
              : { kind: 'blank', value: null, style: 0, formula: { text: value.f, type: 'normal' } };
      sheet.colCount = Math.max(sheet.colCount, c + 1);
    });
    if (cells.length > 0) sheet.rows.set(r, { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells });
    sheet.rowCount = r + 1;
  });
  return sheet;
}

const bookOf = (...sheets: Sheet[]): Workbook => ({ sheets, styles: emptyStyleSheet(), sharedStrings: [], definedNames: [], date1904: false, activeSheet: 0 });

/** 1-based numbers of the hidden rows. */
const hidden = (sheet: Sheet) =>
  Array.from(sheet.rows.entries())
    .filter(([, row]) => row.hidden)
    .map(([index]) => index + 1)
    .sort((a, b) => a - b);

const values = (list: string[], blank = false): FilterCriterion => ({ kind: 'values', values: list, ...(blank ? { blank: true } : {}) });

const column = (sheet: Sheet, col: number, top: number, bottom: number) => {
  const out: unknown[] = [];
  for (let r = top; r <= bottom; r++) {
    const cell = getCell(sheet, r, col);
    out.push(cell ? (cell.formula ? `=${cell.formula.text}` : cell.value) : null);
  }
  return out;
};

function run(history: History, ops: Op[]): void {
  history.apply(ops, undefined, undefined);
}

/** A list of seven rows under a header; the filter on column B ("x") hides rows 3, 4 and 6. */
function listed(): { workbook: Workbook; sheet: Sheet; history: History } {
  const sheet = sheetOf([
    ['Name', 'Kind', 'Amount', 'Twice'],
    ['a', 'x', 1, { f: 'C2*2' }],
    ['b', 'y', 2, { f: 'C3*2' }],
    ['c', 'y', 3, { f: 'C4*2' }],
    ['d', 'x', 4, { f: 'C5*2' }],
    ['e', 'z', 5, { f: 'C6*2' }],
    ['f', 'x', 6, { f: 'C7*2' }],
    ['g', 'x', 7, { f: 'C8*2' }],
  ]);
  const workbook = bookOf(sheet);
  const history = new History(workbook);
  run(history, setAutoFilter(workbook, 0, { top: 0, left: 0, bottom: 7, right: 3 }));
  run(history, setFilterColumn(workbook, 0, 1, values(['x'])));
  assert.deepEqual(hidden(sheet), [3, 4, 6]);
  return { workbook, sheet, history };
}

test('filter: Turkish text keeps "ı" and "i" apart, in the value list and in criteria', () => {
  const sheet = sheetOf([['Word'], ['kıl'], ['kil'], ['KIL'], ['KİL'], ['ISTANBUL'], ['istanbul'], ['İstanbul'], ['ıstanbul']]);
  const workbook = bookOf(sheet);
  const history = new History(workbook);
  run(history, setAutoFilter(workbook, 0, { top: 0, left: 0, bottom: 8, right: 0 }));
  const texts = (locale: typeof LOCALE_TR) => filterValues(workbook, 0, 0, { locale }).items.map((item) => item.text);
  // Turkish: "kıl" and "KIL" are one word, "kil" and "KİL" another.
  assert.deepEqual(texts(LOCALE_TR), ['ISTANBUL', 'istanbul', 'kıl', 'kil']);
  // English: "I" is the capital of "i"; "ı" and "İ" are letters of their own.
  assert.equal(texts(LOCALE_EN).length, 6);

  run(history, setFilterColumn(workbook, 0, 0, values(['kıl']), { locale: LOCALE_TR }));
  assert.deepEqual(hidden(sheet), [3, 5, 6, 7, 8, 9]);
  run(history, setFilterColumn(workbook, 0, 0, values(['KİL']), { locale: LOCALE_TR }));
  assert.deepEqual(hidden(sheet), [2, 4, 6, 7, 8, 9]);
  run(history, setFilterColumn(workbook, 0, 0, { kind: 'custom', conditions: [{ operator: 'equal', value: 'istanbul' }] }, { locale: LOCALE_TR }));
  assert.deepEqual(hidden(sheet), [2, 3, 4, 5, 6, 9]);
  run(history, setFilterColumn(workbook, 0, 0, { kind: 'custom', conditions: [{ operator: 'equal', value: 'I*' }] }, { locale: LOCALE_TR }));
  assert.deepEqual(hidden(sheet), [2, 3, 4, 5, 7, 8]);
  // The same criterion in English pairs "I" with "i".
  run(history, setFilterColumn(workbook, 0, 0, { kind: 'custom', conditions: [{ operator: 'equal', value: 'istanbul' }] }, { locale: LOCALE_EN }));
  assert.deepEqual(hidden(sheet), [2, 3, 4, 5, 8, 9]);
  run(history, setFilterColumn(workbook, 0, 0, values(['kil']), { locale: LOCALE_EN }));
  assert.deepEqual(hidden(sheet), [2, 5, 6, 7, 8, 9]);
  // The operation itself states the rows: applying, undoing and redoing it needs no locale.
  history.undo();
  assert.deepEqual(hidden(sheet), [2, 3, 4, 5, 8, 9]);
  history.redo();
  assert.deepEqual(hidden(sheet), [2, 5, 6, 7, 8, 9]);
});

test('filter: TRUE and FALSE match logical cells in the language of the locale and in English', () => {
  const sheet = sheetOf([['Flag'], [true], [false], ['TRUE'], ['doğru'], [1]]);
  const workbook = bookOf(sheet);
  const history = new History(workbook);
  run(history, setAutoFilter(workbook, 0, { top: 0, left: 0, bottom: 5, right: 0 }));
  const equal = (value: string): FilterCriterion => ({ kind: 'custom', conditions: [{ operator: 'equal', value }] });
  run(history, setFilterColumn(workbook, 0, 0, equal('TRUE'), { locale: LOCALE_TR }));
  assert.deepEqual(hidden(sheet), [3, 5, 6]);
  run(history, setFilterColumn(workbook, 0, 0, equal('DOĞRU'), { locale: LOCALE_TR }));
  assert.deepEqual(hidden(sheet), [3, 4, 6]);
  run(history, setFilterColumn(workbook, 0, 0, { kind: 'custom', conditions: [{ operator: 'notEqual', value: 'false' }] }, { locale: LOCALE_TR }));
  assert.deepEqual(hidden(sheet), [3]);
  // A list of values written in English by another program.
  run(history, setFilterColumn(workbook, 0, 0, values(['FALSE']), { locale: LOCALE_TR }));
  assert.deepEqual(hidden(sheet), [2, 4, 5, 6]);
  run(history, setFilterColumn(workbook, 0, 0, equal('TRUE'), { locale: LOCALE_EN }));
  assert.deepEqual(hidden(sheet), [3, 5, 6]);
});

test('filter: deleting the column of a criterion shows the rows only that criterion hid', () => {
  const { workbook, sheet, history } = listed();
  // A second criterion, on the amounts: rows 2 and 3 go too (row 3 is hidden by both).
  run(history, setFilterColumn(workbook, 0, 2, { kind: 'custom', conditions: [{ operator: 'greaterThan', value: '2' }] }));
  assert.deepEqual(hidden(sheet), [2, 3, 4, 6]);
  const before = sheet.autoFilter as AutoFilter;
  run(history, deleteColumns(workbook, 0, 1, 1));
  // The criterion went with its column; the other one moved one column to the left and still holds.
  assert.deepEqual(sheet.autoFilter!.range, { top: 0, left: 0, bottom: 7, right: 2 });
  assert.deepEqual(
    sheet.autoFilter!.columns.map((entry) => [entry.colId, entry.kind]),
    [[1, 'custom']],
  );
  assert.deepEqual(hidden(sheet), [2, 3]);
  assert.deepEqual(column(sheet, 1, 0, 7), ['Amount', 1, 2, 3, 4, 5, 6, 7]);
  assert.deepEqual(filterSummary(sheet), { shown: 5, total: 7 });
  // "Clear" works on what is left.
  run(history, clearFilter(workbook, 0));
  assert.deepEqual(hidden(sheet), []);
  history.undo();
  // One step back: the column, its criterion and the rows as they were.
  history.undo();
  assert.equal(sheet.autoFilter, before);
  assert.deepEqual(hidden(sheet), [2, 3, 4, 6]);
  assert.deepEqual(column(sheet, 1, 0, 2), ['Kind', 'x', 'y']);
  history.redo();
  assert.deepEqual(hidden(sheet), [2, 3]);
});

test('filter: deleting the only criterion column leaves a filter that hides nothing', () => {
  const { workbook, sheet, history } = listed();
  run(history, deleteColumns(workbook, 0, 1, 1));
  assert.deepEqual(sheet.autoFilter!.columns, []);
  assert.deepEqual(hidden(sheet), []);
  assert.equal(filterSummary(sheet), undefined);
  history.undo();
  assert.deepEqual(hidden(sheet), [3, 4, 6]);
  // A column without a criterion is deleted as before: one operation, the filter follows.
  const ops = deleteColumns(workbook, 0, 0, 1);
  assert.deepEqual(ops, [{ type: 'delete', sheet: 0, axis: 'col', at: 0, count: 1 }]);
  run(history, ops);
  assert.deepEqual(sheet.autoFilter!.columns.map((entry) => entry.colId), [0]);
  assert.deepEqual(hidden(sheet), [3, 4, 6]);
});

test('filter: the criterion of a deleted column is dropped by the operation itself, the others keep their columns', () => {
  const { workbook, sheet } = listed();
  // The bare operation, as another client would send it.
  applyOp(workbook, { type: 'setAutoFilter', sheet: 0, autoFilter: { ...(sheet.autoFilter as AutoFilter), columns: [{ colId: 1, kind: 'values', values: ['x'] }, { colId: 3, kind: 'values', values: ['2'] }] } });
  applyOp(workbook, { type: 'delete', sheet: 0, axis: 'col', at: 1, count: 1 });
  assert.deepEqual(sheet.autoFilter!.columns, [{ colId: 2, kind: 'values', values: ['2'] }]);
});

test('filter: deleting the header row removes the filter, shows every row and drops its name', () => {
  for (const viaCommand of [true, false]) {
    const { workbook, sheet, history } = listed();
    const filter = sheet.autoFilter;
    assert.equal(workbook.definedNames.length, 1);
    run(history, viaCommand ? deleteRows(workbook, 0, 0, 1) : [{ type: 'delete', sheet: 0, axis: 'row', at: 0, count: 1 }]);
    assert.equal(sheet.autoFilter, undefined);
    assert.deepEqual(hidden(sheet), []);
    assert.deepEqual(workbook.definedNames, []);
    assert.deepEqual(column(sheet, 0, 0, 6), ['a', 'b', 'c', 'd', 'e', 'f', 'g']);
    assert.equal(sheet.autoFilterModified, true);
    // One step back: the filter, its name and the hidden rows are what they were.
    history.undo();
    assert.equal(sheet.autoFilter, filter);
    assert.deepEqual(hidden(sheet), [3, 4, 6]);
    assert.deepEqual(workbook.definedNames, [{ name: FILTER_NAME, value: 'Data!$A$1:$D$8', localSheet: 0, hidden: true }]);
    assert.equal(history.canUndo, true);
    history.redo();
    assert.equal(sheet.autoFilter, undefined);
    assert.deepEqual(hidden(sheet), []);
  }
});

test('filter: a row above the header only moves the filter; a row of data only shortens it', () => {
  const { workbook, sheet, history } = listed();
  run(history, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 2 }]);
  assert.deepEqual(sheet.autoFilter!.range, { top: 2, left: 0, bottom: 9, right: 3 });
  run(history, [{ type: 'delete', sheet: 0, axis: 'row', at: 0, count: 2 }]);
  assert.deepEqual(sheet.autoFilter!.range, { top: 0, left: 0, bottom: 7, right: 3 });
  assert.deepEqual(hidden(sheet), [3, 4, 6]);
  run(history, [{ type: 'delete', sheet: 0, axis: 'row', at: 1, count: 1 }]);
  assert.deepEqual(sheet.autoFilter!.range, { top: 0, left: 0, bottom: 6, right: 3 });
  assert.deepEqual(hidden(sheet), [2, 3, 5]);
  assert.equal(workbook.definedNames[0].value, 'Data!$A$1:$D$7');
});

test('filter: deleting all of its columns, or all of its rows, removes the filter and its name', () => {
  let { workbook, sheet, history } = listed();
  // Rows hidden outside the filter are none of its business.
  sheet.rows.set(9, { customHeight: false, hidden: true, outlineLevel: 0, collapsed: false, cells: [{ kind: 'number', value: 1, style: 0 }] });
  run(history, deleteColumns(workbook, 0, 0, 4));
  assert.equal(sheet.autoFilter, undefined);
  assert.deepEqual(hidden(sheet), [10]);
  assert.deepEqual(workbook.definedNames, []);
  history.undo();
  assert.deepEqual(hidden(sheet), [3, 4, 6, 10]);
  assert.equal(workbook.definedNames.length, 1);
  assert.deepEqual(sheet.autoFilter!.range, { top: 0, left: 0, bottom: 7, right: 3 });

  ({ workbook, sheet, history } = listed());
  run(history, [{ type: 'delete', sheet: 0, axis: 'row', at: 0, count: 8 }]);
  assert.equal(sheet.autoFilter, undefined);
  assert.deepEqual(workbook.definedNames, []);
  history.undo();
  assert.deepEqual(hidden(sheet), [3, 4, 6]);
  assert.equal(workbook.definedNames.length, 1);
});

test('filter: the rows a filter hides are told from rows hidden by hand', () => {
  const { workbook, sheet, history } = listed();
  assert.deepEqual(rowsHiddenByFilter(sheet, 0, 100), [2, 3, 5]);
  assert.deepEqual(rowsHiddenByFilter(sheet, 3, 4), [3]);
  // Without a criterion the filter hides nothing: what is hidden was hidden by hand.
  run(history, clearFilter(workbook, 0));
  run(history, [{ type: 'setRows', sheet: 0, rows: [{ row: 2, props: { hidden: true } }] }]);
  assert.deepEqual(rowsHiddenByFilter(sheet, 0, 100), []);
  // A hidden row outside the data rows of the filter is hidden by hand as well.
  run(history, setFilterColumn(workbook, 0, 1, values(['x'])));
  run(history, [{ type: 'setRows', sheet: 0, rows: [{ row: 12, props: { hidden: true } }] }]);
  assert.deepEqual(rowsHiddenByFilter(sheet, 0, 100), [2, 3, 5]);
  const plain = sheetOf([[1], [2]]);
  (plain.rows.get(1) as Row).hidden = true;
  assert.deepEqual(rowsHiddenByFilter(plain, 0, 10), []);
});

test('filter: clearing a selection leaves the cells of filtered rows alone', () => {
  const { workbook, sheet, history } = listed();
  run(history, clearRange(workbook, 0, { top: 1, left: 2, bottom: 6, right: 2 }));
  assert.deepEqual(column(sheet, 2, 0, 7), ['Amount', null, 2, 3, null, 5, null, 7]);
  history.undo();
  assert.deepEqual(column(sheet, 2, 0, 7), ['Amount', 1, 2, 3, 4, 5, 6, 7]);
  // Rows hidden by hand are part of a selection, as in Excel.
  run(history, clearFilter(workbook, 0));
  run(history, [{ type: 'setRows', sheet: 0, rows: [{ row: 2, props: { hidden: true } }] }]);
  run(history, clearRange(workbook, 0, { top: 1, left: 2, bottom: 3, right: 2 }));
  assert.deepEqual(column(sheet, 2, 0, 7), ['Amount', null, null, null, 4, 5, 6, 7]);
});

test('filter: a copy takes the shown rows only, and pastes as one block with its formulas in step', () => {
  const { workbook, sheet } = listed();
  const copied = copyRange(workbook, 0, { top: 1, left: 2, bottom: 6, right: 3 });
  // A formula that has no value yet shows as its text.
  assert.equal(copied.text, '1\t=C2*2\n4\t=C5*2\n6\t=C7*2');
  assert.equal(copied.cells.length, 3);
  assert.deepEqual(copied.rows, [1, 4, 6]);
  assert.deepEqual(copied.range, { top: 1, left: 2, bottom: 6, right: 3 });
  // Pasted two columns to the right, from row 11 on: three rows, each formula moved from its own row.
  const result = pasteAt(workbook, 0, { top: 10, left: 4, bottom: 10, right: 4 }, copied.text, copied);
  for (const op of result.ops) applyOp(workbook, op);
  assert.deepEqual(result.range, { top: 10, left: 4, bottom: 12, right: 5 });
  assert.deepEqual(column(sheet, 4, 10, 13), [1, 4, 6, null]);
  assert.deepEqual(column(sheet, 5, 10, 13), ['=E11*2', '=E12*2', '=E13*2', null]);
  // A copy without filtered rows in it is what it always was.
  const plain = copyRange(workbook, 0, { top: 0, left: 0, bottom: 1, right: 1 });
  assert.equal(plain.rows, undefined);
  assert.equal(plain.text, 'Name\tKind\na\tx');
  // A selection of filtered rows only copies nothing but still has a shape.
  const none = copyRange(workbook, 0, { top: 2, left: 0, bottom: 3, right: 0 });
  assert.equal(none.text, '');
  assert.deepEqual(none.cells, []);
});

test('filter: pasting writes into filtered rows too, as before', () => {
  const { workbook, sheet } = listed();
  const result = pasteAt(workbook, 0, { top: 1, left: 0, bottom: 1, right: 0 }, 'p\nq\nr');
  for (const op of result.ops) applyOp(workbook, op);
  assert.deepEqual(column(sheet, 0, 1, 3), ['p', 'q', 'r']);
});

test('filter: deleting selected rows of a filtered list deletes the shown ones, in one undo step', () => {
  const { workbook, sheet, history } = listed();
  const rows = sheet.rows;
  // Rows 2 to 5 are selected; 3 and 4 are filtered out.
  const ops = deleteRows(workbook, 0, 1, 4);
  assert.deepEqual(ops, [
    { type: 'delete', sheet: 0, axis: 'row', at: 4, count: 1 },
    { type: 'delete', sheet: 0, axis: 'row', at: 1, count: 1 },
  ]);
  run(history, ops);
  assert.deepEqual(column(sheet, 0, 0, 6), ['Name', 'b', 'c', 'e', 'f', 'g', null]);
  assert.deepEqual(hidden(sheet), [2, 3, 4]);
  assert.deepEqual(sheet.autoFilter!.range, { top: 0, left: 0, bottom: 5, right: 3 });
  // Formulas followed their rows.
  assert.deepEqual(column(sheet, 3, 1, 5), ['=C2*2', '=C3*2', '=C4*2', '=C5*2', '=C6*2']);
  history.undo();
  assert.equal(sheet.rows, rows);
  assert.deepEqual(column(sheet, 0, 0, 7), ['Name', 'a', 'b', 'c', 'd', 'e', 'f', 'g']);
  assert.deepEqual(hidden(sheet), [3, 4, 6]);
  assert.equal(history.canUndo, true);
  history.redo();
  assert.deepEqual(column(sheet, 0, 0, 5), ['Name', 'b', 'c', 'e', 'f', 'g']);

  // Nothing but filtered rows selected: nothing to delete. No filtered row selected: one plain operation.
  assert.deepEqual(deleteRows(workbook, 0, 1, 2), []);
  assert.deepEqual(deleteRows(workbook, 0, 4, 2), [{ type: 'delete', sheet: 0, axis: 'row', at: 4, count: 2 }]);
  assert.deepEqual(deleteRows(workbook, 0, 40, 3), [{ type: 'delete', sheet: 0, axis: 'row', at: 40, count: 3 }]);
});

test('filter: rows hidden by hand are deleted with the selection', () => {
  const sheet = sheetOf([['a'], ['b'], ['c'], ['d']]);
  const workbook = bookOf(sheet);
  (sheet.rows.get(1) as Row).hidden = true;
  assert.deepEqual(deleteRows(workbook, 0, 0, 3), [{ type: 'delete', sheet: 0, axis: 'row', at: 0, count: 3 }]);
});

test('filter: the value list says when the criterion of its column cannot be shown as values', () => {
  const { workbook, sheet } = listed();
  assert.equal(filterValues(workbook, 0, 1).opaque, false);
  assert.equal(filterValues(workbook, 0, 0).opaque, false);
  applyOp(workbook, { type: 'setAutoFilter', sheet: 0, autoFilter: { ...(sheet.autoFilter as AutoFilter), columns: [{ colId: 0, kind: 'color', raw: '<filterColumn colId="0"><colorFilter dxfId="0"/></filterColumn>' }] } });
  assert.equal(filterValues(workbook, 0, 0).opaque, true);
  assert.equal(filterValues(workbook, 0, 1).opaque, false);
});
