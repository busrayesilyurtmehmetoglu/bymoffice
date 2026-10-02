// The filter of a sheet: which rows its criteria hide, the value list of a
// column, switching it on and off, and how it follows moved columns and undo.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { LOCALE_TR } from '../src/format/locale';
import { FILTER_NAME, createSheet, emptyStyleSheet, type AutoFilter, type Cell, type Row, type Sheet, type Workbook } from '../src/model/sheet/types';
import { MAX_FILTER_VALUES, clearFilter, filterSummary, filterValues, setAutoFilter, setFilterColumn, wildcardMatch, type FilterCriterion } from '../src/ops/filter';
import { History, applyOp, type Op } from '../src/ops/ops';

type Value = number | string | boolean | null;

function sheetOf(grid: Value[][], name = 'Data'): Sheet {
  const sheet = createSheet({ name, sheetId: 1, state: 'visible', kind: 'worksheet', part: '/xl/worksheets/sheet1.xml' });
  grid.forEach((values, r) => {
    const cells: Cell[] = [];
    values.forEach((value, c) => {
      if (value === null) return;
      cells[c] = typeof value === 'number' ? { kind: 'number', value, style: 0 } : typeof value === 'boolean' ? { kind: 'boolean', value, style: 0 } : { kind: 'string', value, style: 0 };
      sheet.colCount = Math.max(sheet.colCount, c + 1);
    });
    if (cells.length > 0) sheet.rows.set(r, { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells });
    sheet.rowCount = r + 1;
  });
  return sheet;
}

const bookOf = (...sheets: Sheet[]): Workbook => ({ sheets, styles: emptyStyleSheet(), sharedStrings: [], definedNames: [], date1904: false, activeSheet: 0 });

/** Rows 2.. of the fruit list: name, colour, amount. Row 6 has no colour, row 7 is empty. */
function fruit(): { workbook: Workbook; sheet: Sheet; history: History } {
  const sheet = sheetOf([
    ['Name', 'Colour', 'Amount'],
    ['Apple', 'red', 10],
    ['Banana', 'yellow', 25],
    ['Cherry', 'Red', 5],
    ['Lemon', 'yellow', 40],
    ['Mystery', null, 15],
    [null, null, null],
    ['Plum', 'purple', 30],
  ]);
  const workbook = bookOf(sheet);
  const history = new History(workbook);
  history.apply(setAutoFilter(workbook, 0, { top: 0, left: 0, bottom: 7, right: 2 }), undefined, undefined);
  return { workbook, sheet, history };
}

/** 1-based numbers of the hidden rows. */
const hidden = (sheet: Sheet) =>
  Array.from(sheet.rows.entries())
    .filter(([, row]) => row.hidden)
    .map(([index]) => index + 1)
    .sort((a, b) => a - b);

const values = (list: string[], blank = false): FilterCriterion => ({ kind: 'values', values: list, ...(blank ? { blank: true } : {}) });

function run(history: History, ops: Op[]): void {
  history.apply(ops, undefined, undefined);
}

test('filter: switching it on names its range for other applications, switching it off removes the name', () => {
  const sheet = sheetOf([['a', 'b'], [1, 2], [3, 4]], "It's data");
  const workbook = bookOf(sheet);
  const history = new History(workbook);
  run(history, setAutoFilter(workbook, 0, { top: 0, left: 0, bottom: 2, right: 1 }));
  assert.deepEqual(sheet.autoFilter, { range: { top: 0, left: 0, bottom: 2, right: 1 }, columns: [] });
  assert.equal(sheet.autoFilterModified, true);
  assert.equal(sheet.modified, true);
  assert.deepEqual(workbook.definedNames, [{ name: FILTER_NAME, value: "'It''s data'!$A$1:$B$3", localSheet: 0, hidden: true }]);
  history.undo();
  assert.equal(sheet.autoFilter, undefined);
  assert.deepEqual(workbook.definedNames, []);
  history.redo();
  assert.equal(workbook.definedNames.length, 1);
  run(history, setAutoFilter(workbook, 0, undefined));
  assert.equal(sheet.autoFilter, undefined);
  assert.deepEqual(workbook.definedNames, []);
  assert.deepEqual(setAutoFilter(workbook, 0, undefined), []);
});

test('filter: an existing name follows a filter that is replaced, and whole columns are cut to the used rows', () => {
  const sheet = sheetOf([['a'], [1], [2]], 'Sheet1');
  const workbook = bookOf(sheet);
  workbook.definedNames = [
    { name: 'Other', value: 'Sheet1!$A$1', hidden: false },
    { name: '_xlnm._FilterDatabase', value: 'Sheet1!$A$1:$A$2', localSheet: 0, hidden: true },
  ];
  sheet.autoFilter = { range: { top: 0, left: 0, bottom: 1, right: 0 }, columns: [] };
  applyOp(workbook, setAutoFilter(workbook, 0, { top: 0, left: 0, bottom: 1048575, right: 0 })[0]);
  assert.deepEqual(sheet.autoFilter.range, { top: 0, left: 0, bottom: 2, right: 0 });
  assert.equal(workbook.definedNames[1].value, 'Sheet1!$A$1:$A$3');
  assert.equal(workbook.definedNames.length, 2);
});

test('filter: a list of values hides the other rows; text is compared without regard to case', () => {
  const { workbook, sheet, history } = fruit();
  run(history, setFilterColumn(workbook, 0, 1, values(['red'])));
  assert.deepEqual(hidden(sheet), [3, 5, 6, 7, 8]);
  assert.deepEqual(filterSummary(sheet), { shown: 2, total: 7 });
  // With blanks: the row without a colour and the empty row are shown as well.
  run(history, setFilterColumn(workbook, 0, 1, values(['RED'], true)));
  assert.deepEqual(hidden(sheet), [3, 5, 8]);
  // The row that was created only to be hidden is gone again.
  assert.equal(sheet.rows.has(6), false);
  history.undo();
  assert.deepEqual(hidden(sheet), [3, 5, 6, 7, 8]);
  history.undo();
  assert.deepEqual(hidden(sheet), []);
  assert.deepEqual(sheet.autoFilter!.columns, []);
  assert.equal(filterSummary(sheet), undefined);
  history.redo();
  assert.deepEqual(hidden(sheet), [3, 5, 6, 7, 8]);
});

test('filter: criteria of several columns all have to hold, and removing one shows what the others allow', () => {
  const { workbook, sheet, history } = fruit();
  run(history, setFilterColumn(workbook, 0, 1, values(['yellow', 'red'])));
  run(history, setFilterColumn(workbook, 0, 2, { kind: 'custom', conditions: [{ operator: 'greaterThan', value: '9' }] }));
  assert.deepEqual(hidden(sheet), [4, 6, 7, 8]);
  assert.deepEqual(sheet.autoFilter!.columns.map((column) => column.colId), [1, 2]);
  run(history, setFilterColumn(workbook, 0, 1, undefined));
  assert.deepEqual(hidden(sheet), [4, 7]);
  run(history, clearFilter(workbook, 0));
  assert.deepEqual(hidden(sheet), []);
  assert.deepEqual(clearFilter(workbook, 0), []);
  assert.deepEqual(setFilterColumn(workbook, 0, 0, undefined), []);
});

test('filter: clearing shows rows hidden by hand inside the range, rows outside it stay hidden', () => {
  const { workbook, sheet, history } = fruit();
  sheet.rows.set(2, { ...(sheet.rows.get(2) as Row), hidden: true });
  sheet.rows.set(20, { customHeight: false, hidden: true, outlineLevel: 0, collapsed: false, cells: [] });
  run(history, setFilterColumn(workbook, 0, 0, values(['Apple', 'Banana'])));
  assert.deepEqual(hidden(sheet), [4, 5, 6, 7, 8, 21]);
  run(history, setFilterColumn(workbook, 0, 0, undefined));
  assert.deepEqual(hidden(sheet), [21]);
  // Switching the filter off shows every row under its header as well.
  run(history, setFilterColumn(workbook, 0, 0, values(['Plum'])));
  run(history, setAutoFilter(workbook, 0, undefined));
  assert.deepEqual(hidden(sheet), [21]);
  history.undo();
  assert.deepEqual(hidden(sheet), [2, 3, 4, 5, 6, 7, 21]);
});

test('filter: custom conditions - comparisons, wildcards, and / or, blanks', () => {
  const custom = (and: boolean, ...conditions: [string, string][]): FilterCriterion => ({ kind: 'custom', and, conditions: conditions.map(([operator, value]) => ({ operator, value })) });
  const cases: [number, FilterCriterion, number[]][] = [
    [2, custom(true, ['greaterThanOrEqual', '10'], ['lessThan', '30']), [4, 5, 7, 8]],
    [2, custom(false, ['lessThan', '10'], ['greaterThan', '30']), [2, 3, 6, 7, 8]],
    [2, custom(false, ['equal', '25']), [2, 4, 5, 6, 7, 8]],
    [2, custom(false, ['notEqual', '25']), [3]],
    [0, custom(false, ['equal', '*e*']), [3, 7, 8]],
    [0, custom(false, ['equal', '?e*']), [2, 3, 4, 6, 7, 8]],
    [0, custom(false, ['notEqual', 'b*']), [3]],
    [0, custom(true, ['greaterThanOrEqual', 'b'], ['lessThan', 'd']), [2, 5, 6, 7, 8]],
    [1, custom(false, ['notEqual', ' ']), [6, 7]],
    [1, custom(false, ['equal', '']), [2, 3, 4, 5, 8]],
    [1, custom(false, ['lessThanOrEqual', 'RED']), [3, 5, 6, 7]],
  ];
  for (const [colId, criterion, expected] of cases) {
    const { workbook, sheet, history } = fruit();
    run(history, setFilterColumn(workbook, 0, colId, criterion));
    assert.deepEqual(hidden(sheet), expected, JSON.stringify(criterion));
  }
});

test('filter: wildcards match as in Excel and stay fast on text that almost matches', () => {
  assert.equal(wildcardMatch('apple', 'a*e'), true);
  assert.equal(wildcardMatch('apple', 'a?e'), false);
  assert.equal(wildcardMatch('apple', '*'), true);
  assert.equal(wildcardMatch('', '*'), true);
  assert.equal(wildcardMatch('', '?'), false);
  assert.equal(wildcardMatch('what?', 'what~?'), true);
  assert.equal(wildcardMatch('whatx', 'what~?'), false);
  assert.equal(wildcardMatch('a*b', 'a~*b'), true);
  assert.equal(wildcardMatch('axb', 'a~*b'), false);
  assert.equal(wildcardMatch('tilde~', 'tilde~'), true);
  const long = 'a'.repeat(20000);
  const started = Date.now();
  assert.equal(wildcardMatch(long, '*a*a*a*a*a*b'), false);
  assert.ok(Date.now() - started < 2000);
});

test('filter: top and bottom items, percentages, above and below average', () => {
  const cases: [FilterCriterion, number[]][] = [
    [{ kind: 'top10', top: true, value: 2 }, [2, 3, 4, 6, 7]],
    [{ kind: 'top10', top: false, value: 2 }, [3, 5, 6, 7, 8]],
    [{ kind: 'top10', top: true, percent: true, value: 50 }, [2, 4, 6, 7]],
    [{ kind: 'top10', top: true, percent: true, value: 1 }, [2, 3, 4, 6, 7, 8]],
    [{ kind: 'dynamic', dynamicType: 'aboveAverage' }, [2, 4, 6, 7]],
    [{ kind: 'dynamic', dynamicType: 'belowAverage' }, [3, 5, 7, 8]],
  ];
  for (const [criterion, expected] of cases) {
    const { workbook, sheet, history } = fruit();
    run(history, setFilterColumn(workbook, 0, 2, criterion));
    assert.deepEqual(hidden(sheet), expected, JSON.stringify(criterion));
  }
});

test('filter: a criterion that cannot be applied keeps the rows it hid while another column changes', () => {
  const { workbook, sheet, history } = fruit();
  // As read from a file: a colour filter on the first column that hid rows 3 and 8.
  const filter: AutoFilter = { ...sheet.autoFilter!, columns: [{ colId: 0, kind: 'color', raw: '<filterColumn colId="0"><colorFilter dxfId="0"/></filterColumn>' }] };
  applyOp(workbook, { type: 'setAutoFilter', sheet: 0, autoFilter: filter, hide: [2, 7] });
  run(history, setFilterColumn(workbook, 0, 1, values(['red', 'yellow'])));
  assert.deepEqual(hidden(sheet), [3, 6, 7, 8]);
  assert.deepEqual(sheet.autoFilter!.held, [2, 7]);
  assert.equal(sheet.autoFilter!.columns[0].raw, '<filterColumn colId="0"><colorFilter dxfId="0"/></filterColumn>');
  run(history, setFilterColumn(workbook, 0, 1, undefined));
  assert.deepEqual(hidden(sheet), [3, 8]);
  // Dynamic filters this engine does not know are of the same kind.
  run(history, setFilterColumn(workbook, 0, 2, { kind: 'dynamic', dynamicType: 'thisMonth' }));
  assert.deepEqual(hidden(sheet), [3, 8]);
  // Replacing the colour filter itself shows what it hid.
  run(history, setFilterColumn(workbook, 0, 2, undefined));
  run(history, setFilterColumn(workbook, 0, 0, values(['Banana', 'Plum'])));
  assert.deepEqual(hidden(sheet), [2, 4, 5, 6, 7]);
  assert.equal(sheet.autoFilter!.held, undefined);
});

test('filter: dates chosen by year, month or day', () => {
  const sheet = sheetOf([['When'], [45292], [45323.5], [45324], [44927], ['text']]);
  const workbook = bookOf(sheet);
  const history = new History(workbook);
  run(history, setAutoFilter(workbook, 0, { top: 0, left: 0, bottom: 5, right: 0 }));
  // 45292 = 2024-01-01, 45323 = 2024-02-01, 45324 = 2024-02-02, 44927 = 2023-01-01.
  run(history, setFilterColumn(workbook, 0, 0, { kind: 'values', values: [], dateGroups: [{ grouping: 'year', year: 2024 }] }));
  assert.deepEqual(hidden(sheet), [5, 6]);
  run(history, setFilterColumn(workbook, 0, 0, { kind: 'values', values: [], dateGroups: [{ grouping: 'month', year: 2024, month: 2 }] }));
  assert.deepEqual(hidden(sheet), [2, 5, 6]);
  run(history, setFilterColumn(workbook, 0, 0, { kind: 'values', values: ['text'], dateGroups: [{ grouping: 'day', year: 2024, month: 2, day: 2 }, { grouping: 'day', year: 2023, month: 1, day: 1 }] }));
  assert.deepEqual(hidden(sheet), [2, 3]);
});

test('filter: numbers are matched as the cells show them, in the locale or as a file in English states them', () => {
  const sheet = sheetOf([['n'], [1.5], [2], [1234.5]]);
  const workbook = bookOf(sheet);
  const history = new History(workbook);
  run(history, setAutoFilter(workbook, 0, { top: 0, left: 0, bottom: 3, right: 0 }));
  run(history, setFilterColumn(workbook, 0, 0, values(['1,5']), { locale: LOCALE_TR }));
  assert.deepEqual(hidden(sheet), [3, 4]);
  run(history, setFilterColumn(workbook, 0, 0, values(['1.5', '1234.5']), { locale: LOCALE_TR }));
  assert.deepEqual(hidden(sheet), [3]);
  assert.deepEqual(
    filterValues(workbook, 0, 0, { locale: LOCALE_TR }).items,
    [{ text: '1,5', checked: true }, { text: '2', checked: false }, { text: '1234,5', checked: true }],
  );
});

test('filter: the value list of a column - distinct, sorted, checked, with blanks, within what other columns show', () => {
  const { workbook, sheet, history } = fruit();
  assert.deepEqual(filterValues(workbook, 0, 1), {
    items: [{ text: 'purple', checked: true }, { text: 'red', checked: true }, { text: 'yellow', checked: true }],
    blanks: true,
    blanksChecked: true,
    truncated: false,
  });
  run(history, setFilterColumn(workbook, 0, 1, values(['red'])));
  assert.deepEqual(filterValues(workbook, 0, 1).items, [{ text: 'purple', checked: false }, { text: 'red', checked: true }, { text: 'yellow', checked: false }]);
  assert.equal(filterValues(workbook, 0, 1).blanksChecked, false);
  // Another column only offers what the first one lets through.
  assert.deepEqual(filterValues(workbook, 0, 0), { items: [{ text: 'Apple', checked: true }, { text: 'Cherry', checked: true }], blanks: false, blanksChecked: false, truncated: false });
  assert.deepEqual(filterValues(workbook, 0, 2).items.map((item) => item.text), ['5', '10']);
  // A criterion of another kind: checked is what it shows.
  run(history, setFilterColumn(workbook, 0, 1, undefined));
  run(history, setFilterColumn(workbook, 0, 2, { kind: 'top10', top: true, value: 2 }));
  assert.deepEqual(filterValues(workbook, 0, 2).items, [
    { text: '5', checked: false },
    { text: '10', checked: false },
    { text: '15', checked: false },
    { text: '25', checked: false },
    { text: '30', checked: true },
    { text: '40', checked: true },
  ]);
  assert.deepEqual(filterValues(bookOf(sheetOf([[1]])), 0, 0).items, []);
});

test('filter: a column with more distinct values than the list holds is cut and says so', () => {
  const grid: Value[][] = [['id']];
  for (let i = 0; i < MAX_FILTER_VALUES + 500; i++) grid.push([`v${i}`]);
  const sheet = sheetOf(grid);
  const workbook = bookOf(sheet);
  applyOp(workbook, setAutoFilter(workbook, 0, { top: 0, left: 0, bottom: grid.length - 1, right: 0 })[0]);
  const list = filterValues(workbook, 0, 0);
  assert.equal(list.items.length, MAX_FILTER_VALUES);
  assert.equal(list.truncated, true);
});

test('filter: it follows inserted and deleted columns and rows, and undo brings it back', () => {
  const { workbook, sheet, history } = fruit();
  run(history, setFilterColumn(workbook, 0, 2, values(['10', '25'])));
  run(history, [{ type: 'insert', sheet: 0, axis: 'col', at: 1, count: 2 }]);
  assert.deepEqual(sheet.autoFilter!.range, { top: 0, left: 0, bottom: 7, right: 4 });
  assert.deepEqual(sheet.autoFilter!.columns.map((column) => column.colId), [4]);
  assert.equal(workbook.definedNames[0].value, 'Data!$A$1:$E$8');
  run(history, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 }]);
  assert.deepEqual(sheet.autoFilter!.range, { top: 1, left: 0, bottom: 8, right: 4 });
  // The column of the criterion is deleted: the criterion goes with it.
  run(history, [{ type: 'delete', sheet: 0, axis: 'col', at: 4, count: 1 }]);
  assert.deepEqual(sheet.autoFilter!.range, { top: 1, left: 0, bottom: 8, right: 3 });
  assert.deepEqual(sheet.autoFilter!.columns, []);
  history.undo();
  assert.deepEqual(sheet.autoFilter!.columns.map((column) => column.colId), [4]);
  history.undo();
  history.undo();
  assert.deepEqual(sheet.autoFilter!.range, { top: 0, left: 0, bottom: 7, right: 2 });
  assert.deepEqual(sheet.autoFilter!.columns.map((column) => column.colId), [2]);
  // All of it deleted: no filter is left.
  run(history, [{ type: 'delete', sheet: 0, axis: 'row', at: 0, count: 8 }]);
  assert.equal(sheet.autoFilter, undefined);
  history.undo();
  assert.deepEqual(sheet.autoFilter!.range, { top: 0, left: 0, bottom: 7, right: 2 });
});

test('filter: what cannot be done is refused', () => {
  const sheet = sheetOf([['a'], [1]]);
  const workbook = bookOf(sheet);
  assert.throws(() => setFilterColumn(workbook, 0, 0, values(['1'])), /The sheet has no filter/);
  applyOp(workbook, setAutoFilter(workbook, 0, { top: 0, left: 0, bottom: 1, right: 0 })[0]);
  assert.throws(() => setFilterColumn(workbook, 0, 1, values(['1'])), /outside the filter/);
  assert.throws(() => setFilterColumn(workbook, 0, -1, values(['1'])), /outside the filter/);
  assert.throws(() => setAutoFilter(workbook, 3, undefined), /No sheet/);
  assert.throws(() => applyOp(workbook, { type: 'setAutoFilter', sheet: 0, autoFilter: undefined, hide: [-1] }), /outside the sheet/);
  const tabled = sheetOf([['a'], [1]]);
  tabled.anchored = ['table'];
  const other = bookOf(tabled);
  assert.throws(() => setAutoFilter(other, 0, { top: 0, left: 0, bottom: 1, right: 0 }), /Cannot filter in a sheet with table/);
});

test('filter: 100,000 rows are filtered and listed quickly', () => {
  const grid: Value[][] = [['k', 'n']];
  for (let i = 0; i < 100000; i++) grid.push([`group ${i % 200}`, i]);
  const sheet = sheetOf(grid);
  const workbook = bookOf(sheet);
  const history = new History(workbook);
  run(history, setAutoFilter(workbook, 0, { top: 0, left: 0, bottom: 100000, right: 1 }));
  let started = Date.now();
  const list = filterValues(workbook, 0, 0);
  const listed = Date.now() - started;
  assert.equal(list.items.length, 200);
  started = Date.now();
  run(history, setFilterColumn(workbook, 0, 0, values(['group 7', 'group 19'])));
  const applied = Date.now() - started;
  assert.deepEqual(filterSummary(sheet), { shown: 1000, total: 100000 });
  started = Date.now();
  history.undo();
  const undone = Date.now() - started;
  assert.deepEqual(hidden(sheet), []);
  console.log(`# filter 100,000 rows: value list ${listed} ms, apply ${applied} ms, undo ${undone} ms`);
  assert.ok(listed < 3000 && applied < 3000, `listing took ${listed} ms, applying ${applied} ms`);
});
