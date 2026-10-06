// Sorting: the order of values, the rows that move and what moves with them,
// what is refused, and undo. Every workbook is built in memory.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { LOCALE_EN, LOCALE_TR } from '../src/format/locale';
import { createSheet, emptyStyleSheet, getCell, type Cell, type Row, type Sheet, type Workbook } from '../src/model/sheet/types';
import { History, applyOp, type Op } from '../src/ops/ops';
import { collationKey, currentRegion, sortRange } from '../src/ops/sort';

type Value = number | string | boolean | null | { f: string; v?: number | string } | { e: string } | { s: number; v: number | string };

function cellOf(value: Value): Cell | undefined {
  if (value === null) return undefined;
  if (typeof value === 'number') return { kind: 'number', value, style: 0 };
  if (typeof value === 'string') return { kind: 'string', value, style: 0 };
  if (typeof value === 'boolean') return { kind: 'boolean', value, style: 0 };
  if ('e' in value) return { kind: 'error', value: value.e, style: 0 };
  if ('s' in value) return { kind: typeof value.v === 'number' ? 'number' : 'string', value: value.v, style: value.s };
  const result = value.v === undefined ? null : value.v;
  return { kind: result === null ? 'blank' : typeof result === 'number' ? 'number' : 'string', value: result, style: 0, formula: { text: value.f, type: 'normal' } };
}

function sheetOf(grid: Value[][], name = 'Data'): Sheet {
  const sheet = createSheet({ name, sheetId: 1, state: 'visible', kind: 'worksheet', part: '/xl/worksheets/sheet1.xml' });
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

function bookOf(...sheets: Sheet[]): Workbook {
  return { sheets, styles: emptyStyleSheet(), sharedStrings: [], definedNames: [], date1904: false, activeSheet: 0 };
}

const column = (sheet: Sheet, col: number, top: number, bottom: number) => {
  const out: unknown[] = [];
  for (let r = top; r <= bottom; r++) {
    const cell = getCell(sheet, r, col);
    out.push(cell ? (cell.formula ? `=${cell.formula.text}` : cell.value) : null);
  }
  return out;
};

function apply(workbook: Workbook, ops: Op[]): Op[] {
  return ops.map((op) => applyOp(workbook, op)).reverse();
}

const full = (sheet: Sheet) => ({ top: 0, left: 0, bottom: sheet.rowCount - 1, right: sheet.colCount - 1 });

test('sort: numbers come before text, text before booleans, those before errors; empty cells are last', () => {
  const sheet = sheetOf([[{ e: '#N/A' }], [true], ['pear'], [null], [3], ['Apple'], [false], [-1], [{ e: '#DIV/0!' }], ['apple 2']]);
  const workbook = bookOf(sheet);
  apply(workbook, sortRange(workbook, 0, { top: 0, left: 0, bottom: 9, right: 0 }, [{ col: 0 }]));
  assert.deepEqual(column(sheet, 0, 0, 9), [-1, 3, 'Apple', 'apple 2', 'pear', false, true, '#N/A', '#DIV/0!', null]);
});

test('sort: descending turns the order round and still leaves empty cells last', () => {
  const sheet = sheetOf([[null], [2], ['b'], [true], [{ e: '#REF!' }], [10], ['a'], [null]]);
  sheet.rowCount = 8;
  const workbook = bookOf(sheet);
  apply(workbook, sortRange(workbook, 0, { top: 0, left: 0, bottom: 7, right: 0 }, [{ col: 0, descending: true }]));
  assert.deepEqual(column(sheet, 0, 0, 7), ['#REF!', true, 'b', 'a', 10, 2, null, null]);
});

test('sort: rows that compare equal keep their order, also in a long list', () => {
  const grid: Value[][] = [];
  for (let i = 0; i < 500; i++) grid.push([i % 3, i]);
  const sheet = sheetOf(grid);
  const workbook = bookOf(sheet);
  apply(workbook, sortRange(workbook, 0, full(sheet), [{ col: 0 }]));
  let last = -1;
  let group = 0;
  for (let r = 0; r < 500; r++) {
    const key = getCell(sheet, r, 0)!.value as number;
    const tag = getCell(sheet, r, 1)!.value as number;
    if (key !== group) {
      assert.equal(key, group + 1);
      group = key;
      last = -1;
    }
    assert.ok(tag > last, `row ${r}`);
    last = tag;
  }
  // Descending as well: equal keys still in their first order.
  apply(workbook, sortRange(workbook, 0, full(sheet), [{ col: 0, descending: true }]));
  assert.deepEqual(column(sheet, 1, 0, 2), [2, 5, 8]);
});

test('sort: text follows the Turkish alphabet in the Turkish locale and ignores case', () => {
  const words = ['zil', 'ıslak', 'çay', 'İnci', 'ugur', 'cam', 'şeker', 'süt', 'Ilık', 'ödev', 'oda', 'ğ', 'gül', 'üzüm', 'iz', 'CAN'];
  const sheet = sheetOf(words.map((word) => [word]));
  const workbook = bookOf(sheet);
  apply(workbook, sortRange(workbook, 0, full(sheet), [{ col: 0 }], { locale: LOCALE_TR }));
  assert.deepEqual(column(sheet, 0, 0, words.length - 1), ['cam', 'CAN', 'çay', 'gül', 'ğ', 'Ilık', 'ıslak', 'İnci', 'iz', 'oda', 'ödev', 'süt', 'şeker', 'ugur', 'üzüm', 'zil']);
  assert.ok(collationKey('I', LOCALE_TR) === collationKey('ı', LOCALE_TR));
  assert.ok(collationKey('İ', LOCALE_TR) === collationKey('i', LOCALE_TR));
  assert.ok(collationKey('ı', LOCALE_TR) < collationKey('i', LOCALE_TR));
  assert.ok(collationKey('9', LOCALE_TR) < collationKey('a', LOCALE_TR));
  assert.equal(collationKey('ABC', LOCALE_EN), 'abc');
});

test('sort: several keys, and a header row that stays', () => {
  const sheet = sheetOf([
    ['Name', 'Group', 'Score'],
    ['d', 2, 5],
    ['a', 1, 7],
    ['c', 2, 9],
    ['b', 1, 7],
  ]);
  const workbook = bookOf(sheet);
  apply(workbook, sortRange(workbook, 0, full(sheet), [{ col: 1 }, { col: 2, descending: true }, { col: 0, descending: true }], { header: true }));
  // Group 1 first; within it the scores tie, so the names decide, from the end of the alphabet.
  assert.deepEqual(column(sheet, 0, 0, 4), ['Name', 'b', 'a', 'c', 'd']);
  assert.deepEqual(column(sheet, 2, 0, 4), ['Score', 7, 7, 9, 5]);
});

test('sort: only the columns of the range move; cells beside it stay', () => {
  const sheet = sheetOf([
    [3, 'c', 'keep 1'],
    [1, 'a', 'keep 2'],
    [2, 'b', 'keep 3'],
  ]);
  const workbook = bookOf(sheet);
  apply(workbook, sortRange(workbook, 0, { top: 0, left: 0, bottom: 2, right: 1 }, [{ col: 0 }]));
  assert.deepEqual(column(sheet, 0, 0, 2), [1, 2, 3]);
  assert.deepEqual(column(sheet, 1, 0, 2), ['a', 'b', 'c']);
  assert.deepEqual(column(sheet, 2, 0, 2), ['keep 1', 'keep 2', 'keep 3']);
});

test('sort: formulas move with their rows; relative references follow, absolute ones stay', () => {
  const sheet = sheetOf([
    [3, { f: 'A1*2', v: 6 }, { f: '$A$1+A1', v: 6 }, { f: 'SUM(A$1:A1)', v: 3 }],
    [1, { f: 'A2*2', v: 2 }, { f: '$A$1+A2', v: 4 }, { f: 'SUM(A$1:A2)', v: 4 }],
    [2, { f: 'A3*2', v: 4 }, { f: '$A$1+A3', v: 5 }, { f: 'SUM(A$1:A3)', v: 6 }],
  ]);
  const workbook = bookOf(sheet);
  apply(workbook, sortRange(workbook, 0, full(sheet), [{ col: 0 }]));
  assert.deepEqual(column(sheet, 0, 0, 2), [1, 2, 3]);
  assert.deepEqual(column(sheet, 1, 0, 2), ['=A1*2', '=A2*2', '=A3*2']);
  assert.deepEqual(column(sheet, 2, 0, 2), ['=$A$1+A1', '=$A$1+A2', '=$A$1+A3']);
  assert.deepEqual(column(sheet, 3, 0, 2), ['=SUM(A$1:A1)', '=SUM(A$1:A2)', '=SUM(A$1:A3)']);
});

test('sort: a shared formula group that is moved is written out for every cell of the group', () => {
  const sheet = sheetOf([[3], [1], [2], [9]]);
  const shared = (row: number, text: string): Cell => ({ kind: 'number', value: 0, style: 0, formula: { text, type: 'shared', si: 0, ...(text ? { ref: 'B1:B4' } : {}) } });
  sheet.rows.get(0)!.cells[1] = shared(0, 'A1+1');
  for (let r = 1; r < 4; r++) sheet.rows.get(r)!.cells[1] = shared(r, '');
  sheet.colCount = 2;
  const workbook = bookOf(sheet);
  // Row 4 is outside the sorted range but belongs to the group.
  apply(workbook, sortRange(workbook, 0, { top: 0, left: 0, bottom: 2, right: 1 }, [{ col: 0 }]));
  assert.deepEqual(column(sheet, 1, 0, 3), ['=A1+1', '=A2+1', '=A3+1', '=A4+1']);
  for (let r = 0; r < 4; r++) assert.equal(getCell(sheet, r, 1)!.formula!.type, 'normal');
});

test('sort: cell formats move with the cells, row heights and hidden rows stay', () => {
  const sheet = sheetOf([[{ s: 0, v: 3 }], [{ s: 0, v: 1 }], [{ s: 0, v: 2 }], [{ s: 0, v: 0 }]]);
  const workbook = bookOf(sheet);
  workbook.styles.cellXfs.push({ ...workbook.styles.cellXfs[0], numFmtId: 2 });
  sheet.rows.set(0, { ...(sheet.rows.get(0) as Row), height: 30, customHeight: true, cells: [{ kind: 'number', value: 3, style: 1 }] });
  sheet.rows.set(3, { ...(sheet.rows.get(3) as Row), hidden: true });
  apply(workbook, sortRange(workbook, 0, full(sheet), [{ col: 0 }]));
  // The hidden row (value 0) is not part of the sort.
  assert.deepEqual(column(sheet, 0, 0, 3), [1, 2, 3, 0]);
  assert.equal(sheet.rows.get(0)!.height, 30);
  assert.equal(sheet.rows.get(2)!.height, undefined);
  assert.equal(getCell(sheet, 2, 0)!.style, 1);
  assert.equal(getCell(sheet, 0, 0)!.style, 0);
  assert.equal(sheet.rows.get(3)!.hidden, true);
});

test('sort: a link on one cell goes with its cell', () => {
  const sheet = sheetOf([['b'], ['a'], ['c']]);
  sheet.hyperlinks = [
    { range: { top: 0, left: 0, bottom: 0, right: 0 }, target: 'https://example.com/b' },
    { range: { top: 0, left: 0, bottom: 2, right: 0 }, target: 'https://example.com/all' },
  ];
  const workbook = bookOf(sheet);
  apply(workbook, sortRange(workbook, 0, full(sheet), [{ col: 0 }]));
  assert.deepEqual(sheet.hyperlinks[0].range, { top: 1, left: 0, bottom: 1, right: 0 });
  assert.deepEqual(sheet.hyperlinks[1].range, { top: 0, left: 0, bottom: 2, right: 0 });
  assert.equal(sheet.hyperlinksModified, true);
});

test('sort: merged cells, array formulas and anchored comments are refused with a message, and nothing changes', () => {
  const merged = sheetOf([[2], [1], [3]]);
  merged.merges = [{ top: 1, left: 0, bottom: 2, right: 0 }];
  const wb1 = bookOf(merged);
  assert.throws(() => apply(wb1, sortRange(wb1, 0, full(merged), [{ col: 0 }])), /Cannot sort a range that contains merged cells/);
  assert.deepEqual(column(merged, 0, 0, 2), [2, 1, 3]);

  const array = sheetOf([[2, 0], [1, 0], [3, 0]]);
  array.rows.get(0)!.cells[1] = { kind: 'number', value: 0, style: 0, formula: { text: 'A1:A2*2', type: 'array', ref: 'B1:B2' } };
  const wb2 = bookOf(array);
  assert.throws(() => apply(wb2, sortRange(wb2, 0, full(array), [{ col: 0 }])), /Cannot sort through the array formula in B1:B2/);

  const single = sheetOf([[2, 0], [1, 0]]);
  single.rows.get(0)!.cells[1] = { kind: 'number', value: 4, style: 0, formula: { text: 'SUM(A1*2)', type: 'array', ref: 'B1' } };
  const wb3 = bookOf(single);
  apply(wb3, sortRange(wb3, 0, full(single), [{ col: 0 }]));
  assert.deepEqual(getCell(single, 1, 1)!.formula, { text: 'SUM(A2*2)', type: 'array', ref: 'B2' });

  const noted = sheetOf([[2], [1]]);
  noted.anchored = ['comments', 'vmlDrawing'];
  // As the reader leaves a sheet with a note: the part is named and the note is in the model.
  noted.comments = [{ row: 0, col: 0, text: 'not' }];
  const wb4 = bookOf(noted);
  assert.throws(() => apply(wb4, sortRange(wb4, 0, full(noted), [{ col: 0 }])), /Cannot sort in a sheet with comments: not supported yet/);
});

test('sort: a key outside the range and an order that is no permutation are refused', () => {
  const sheet = sheetOf([[2], [1]]);
  const workbook = bookOf(sheet);
  assert.throws(() => sortRange(workbook, 0, full(sheet), [{ col: 3 }]), /outside the range/);
  assert.throws(() => sortRange(workbook, 0, full(sheet), []), /at least one key/);
  for (const order of [[0, 0], [0], [0, 2], [0.5, 1], [1, NaN]]) {
    assert.throws(() => applyOp(workbook, { type: 'sortRange', sheet: 0, range: { top: 0, left: 0, bottom: 1, right: 0 }, order }), /Invalid sort order/);
  }
  assert.deepEqual(column(sheet, 0, 0, 1), [2, 1]);
});

test('sort: whole columns are cut down to the used rows, and a sorted range gives no operation', () => {
  const sheet = sheetOf([[1], [2], [3]]);
  const workbook = bookOf(sheet);
  assert.deepEqual(sortRange(workbook, 0, { top: 0, left: 0, bottom: 1048575, right: 0 }, [{ col: 0 }]), []);
  const ops = sortRange(workbook, 0, { top: 0, left: 0, bottom: 1048575, right: 0 }, [{ col: 0, descending: true }]);
  assert.equal(ops.length, 1);
  assert.deepEqual(ops[0].type === 'sortRange' && ops[0].range, { top: 0, left: 0, bottom: 2, right: 0 });
});

test('sort: undo and redo bring back exactly the rows there were', () => {
  const sheet = sheetOf([
    ['h', 'x'],
    [3, { f: 'A2+1', v: 4 }],
    [1, { f: 'A3+1', v: 2 }],
    [2, { f: 'A4+1', v: 3 }],
  ]);
  const workbook = bookOf(sheet);
  const history = new History(workbook);
  const before = new Map(sheet.rows);
  history.apply(sortRange(workbook, 0, full(sheet), [{ col: 0 }], { header: true }), undefined, undefined);
  assert.deepEqual(column(sheet, 0, 0, 3), ['h', 1, 2, 3]);
  const sorted = new Map(sheet.rows);
  history.undo();
  assert.equal(sheet.rows.size, before.size);
  for (const [index, row] of before) assert.equal(sheet.rows.get(index), row, `row ${index} is the object it was`);
  history.redo();
  for (const [index, row] of sorted) assert.equal(sheet.rows.get(index), row);
});

test('sort: inside a filter the sort is noted in the filter, and only shown rows move', () => {
  const sheet = sheetOf([['Name'], ['c'], ['x'], ['a'], ['b']]);
  sheet.autoFilter = { range: { top: 0, left: 0, bottom: 4, right: 0 }, columns: [] };
  sheet.rows.set(2, { ...(sheet.rows.get(2) as Row), hidden: true });
  const workbook = bookOf(sheet);
  const history = new History(workbook);
  history.apply(sortRange(workbook, 0, full(sheet), [{ col: 0, descending: true }], { header: true }), undefined, undefined);
  assert.deepEqual(column(sheet, 0, 0, 4), ['Name', 'c', 'x', 'b', 'a']);
  assert.deepEqual(sheet.autoFilter!.sortState, {
    range: { top: 1, left: 0, bottom: 4, right: 0 },
    conditions: [{ range: { top: 1, left: 0, bottom: 4, right: 0 }, descending: true }],
  });
  assert.equal(sheet.autoFilterModified, true);
  history.undo();
  assert.equal(sheet.autoFilter!.sortState, undefined);
});

test('sort: the region around a cell ends at empty rows and columns', () => {
  const sheet = sheetOf([
    ['a', 'b', null, 'far'],
    [1, 2, null, null],
    [null, 3, null, null],
    [null, null, null, null],
    [9, null, null, null],
  ]);
  sheet.rowCount = 5;
  assert.deepEqual(currentRegion(sheet, 1, 0), { top: 0, left: 0, bottom: 2, right: 1 });
  assert.deepEqual(currentRegion(sheet, 2, 1), { top: 0, left: 0, bottom: 2, right: 1 });
  assert.deepEqual(currentRegion(sheet, 4, 0), { top: 4, left: 0, bottom: 4, right: 0 });
  assert.deepEqual(currentRegion(sheet, 0, 3), { top: 0, left: 3, bottom: 0, right: 3 });
  // An empty cell next to data belongs to the data; one with nothing around it is alone.
  assert.deepEqual(currentRegion(sheet, 3, 1), { top: 0, left: 0, bottom: 4, right: 1 });
  assert.deepEqual(currentRegion(sheet, 40, 40), { top: 40, left: 40, bottom: 40, right: 40 });
});

test('sort: 100,000 rows of 20 columns are sorted, undone and redone quickly, and one step holds the rows', () => {
  const sheet = createSheet({ name: 'Big', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/xl/worksheets/sheet1.xml' });
  const ROWS = 100000;
  const COLS = 20;
  let seed = 12345;
  const random = () => (seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff;
  for (let r = 0; r < ROWS; r++) {
    const cells: Cell[] = [];
    for (let c = 0; c < COLS; c++) cells.push(c === 1 ? { kind: 'string', value: `name ${Math.floor(random() * 50000)}`, style: 0 } : { kind: 'number', value: c === 0 ? random() : r * c, style: 0 });
    sheet.rows.set(r, { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells });
  }
  sheet.rowCount = ROWS;
  sheet.colCount = COLS;
  const workbook = bookOf(sheet);
  const history = new History(workbook);
  const started = Date.now();
  history.apply(sortRange(workbook, 0, full(sheet), [{ col: 0 }]), undefined, undefined);
  const numeric = Date.now() - started;
  let previous = -1;
  for (let r = 0; r < ROWS; r += 997) {
    const value = getCell(sheet, r, 0)!.value as number;
    assert.ok(value >= previous);
    previous = value;
  }
  const textStarted = Date.now();
  history.apply(sortRange(workbook, 0, full(sheet), [{ col: 1 }, { col: 0, descending: true }], { locale: LOCALE_TR }), undefined, undefined);
  const text = Date.now() - textStarted;
  const undoStarted = Date.now();
  history.undo();
  history.undo();
  const undo = Date.now() - undoStarted;
  assert.equal(getCell(sheet, 5, 2)!.value, 10);
  console.log(`# sort 100,000 x 20: by a number ${numeric} ms, by text and a number (Turkish) ${text} ms, two undos ${undo} ms`);
  // Generous: the figures above are what is reported; this only catches a change of order of magnitude.
  assert.ok(numeric < 5000 && text < 8000, `sorting took ${numeric} and ${text} ms`);
});
