// Making pivot tables (src/ops/pivot-create.ts): from a range or a defined name, on a new sheet or at
// a cell; what is refused; undo; what is saved, read back by the engine and by independent readers.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { ZipArchive } from '../src/index';
import { buildXlsx } from '../src/formats/xlsx/build';
import { encodeOp } from '../src/collab/wire';
import { CollabUnsupportedError } from '../src/collab/types';
import type { Op } from '../src/ops/ops';
import { createPivot, createPivotProblem, pivotSourceFields } from '../src/ops/pivot-create';
import { listPivots, refreshPivot } from '../src/ops/pivots';
import { renameSheetProblem, deleteSheetProblem } from '../src/ops/ops';
import { copySheetLosses } from '../src/ops/sheets';
import { tableState } from '../src/model/sheet/pivot';
import { bookWith, grid, hasPython, independentPivots, load, pivotBook, same, save, text, type Datum, type Loaded } from './pivot-fixtures';

const DATA: Datum[][] = [
  ['Region', 'Product', 'Sales'],
  ['East', 'A', 10],
  ['West', 'B', 20],
  ['East', 'B', 5],
  ['West', 'A', 7],
  ['East', 'A', 1],
];

const whole = (data: Datum[][]) => ({ top: 0, left: 0, bottom: data.length - 1, right: data[0].length - 1 });

function apply(doc: Loaded, ops: Op[]): void {
  doc.history.apply(ops, undefined, undefined);
}

test('a table made on a new sheet: Excel\'s first guess, the rows of the first column of text and the sum of the first of numbers', async () => {
  const doc = await bookWith(DATA);
  const made = createPivot(doc.wb, { source: { sheet: 0, range: whole(DATA) } });
  assert.equal(doc.wb.sheets.length, 1, 'nothing changes before the operations are applied');
  assert.equal(doc.wb.pivots, undefined);
  apply(doc, made.ops);
  assert.equal(doc.wb.sheets.length, 2);
  assert.equal(made.sheet, 1);
  assert.equal(made.index, 0);
  assert.deepEqual(made.range, { top: 2, left: 0, bottom: 5, right: 1 });
  assert.deepEqual(grid(doc.wb, 1, 2, 0, 5, 1), [
    ['Row Labels', 'Sum of Sales'],
    ['East', 16],
    ['West', 27],
    ['Grand Total', 43],
  ]);
  assert.deepEqual(listPivots(doc.wb), [{ index: 0, name: 'PivotTable1', sheet: 1, range: { top: 2, left: 0, bottom: 5, right: 1 } }]);
  // A refresh of it changes nothing.
  assert.deepEqual(refreshPivot(doc.wb, 0), []);
});

test('pivotSourceFields names the fields of a source and tells the numbers; one cell stands for the block around it', async () => {
  const doc = await bookWith(DATA);
  const fields = pivotSourceFields(doc.wb, { sheet: 0, range: { top: 2, left: 1, bottom: 2, right: 1 } });
  assert.deepEqual(fields.names, ['Region', 'Product', 'Sales']);
  assert.deepEqual(fields.numeric, [false, false, true]);
  assert.deepEqual(fields.range, whole(DATA));
});

test('saved, the table is read back as a table of the file, and independent readers find it whole', async () => {
  const doc = await bookWith(DATA);
  apply(doc, createPivot(doc.wb, { source: { sheet: 0, range: whole(DATA) }, rows: [0], cols: [1], values: [{ field: 2 }] }).ops);
  const bytes = await save(doc);
  const again = await load(bytes);
  const pivots = again.wb.pivots!;
  assert.equal(pivots.tables.length, 1);
  const table = pivots.tables[0];
  assert.equal(table.refusal, undefined);
  assert.equal(table.cache.refusal, undefined);
  assert.deepEqual(table.rowFields, [0]);
  assert.deepEqual(table.colFields, [1]);
  assert.deepEqual(table.cache.source, { kind: 'range', sheet: 'Sheet1', ref: 'A1:C6' });
  assert.equal(table.sheet, again.wb.sheets[1]);
  assert.deepEqual(grid(again.wb, 1, 2, 0, 6, 3), [
    ['Sum of Sales', 'Column Labels', undefined, undefined],
    ['Row Labels', 'A', 'B', 'Grand Total'],
    ['East', 11, 5, 16],
    ['West', 7, 20, 27],
    ['Grand Total', 18, 25, 43],
  ]);
  // Opened again, a refresh changes no cell.
  assert.deepEqual(refreshPivot(again.wb, 0).filter((op) => op.type === 'setCells'), []);
  const book = await text(bytes, 'xl/workbook.xml');
  assert.match(book, /<pivotCaches><pivotCache cacheId="0" r:id="rId\d+"\/><\/pivotCaches>/);
  const types = await text(bytes, '[Content_Types].xml');
  assert.match(types, /pivotTable\+xml/);
  assert.match(types, /pivotCacheDefinition\+xml/);
  assert.match(types, /pivotCacheRecords\+xml/);
  if (hasPython()) assert.deepEqual(independentPivots({ made: bytes }), []);
});

test('saved twice, the parts keep their names; undone and saved, the file is what it was byte for byte', async () => {
  const doc = await bookWith(DATA);
  const original = await save(doc);
  apply(doc, createPivot(doc.wb, { source: { sheet: 0, range: whole(DATA) } }).ops);
  const first = await save(doc);
  const second = await save(doc);
  assert.ok(same(first, second));
  doc.history.undo();
  const back = await save(doc);
  assert.ok(same(back, original), 'made, undone and saved: the bytes of the file');
  doc.history.redo();
  assert.ok(same(await save(doc), first));
});

test('made at a cell of a sheet without relationships, saved, undone and saved: the file is what it was; the parts made for it are gone', async () => {
  const doc = await bookWith(DATA);
  const original = await save(doc);
  assert.equal(ZipArchive.read(original).names().indexOf('xl/worksheets/_rels/sheet1.xml.rels'), -1);
  apply(doc, createPivot(doc.wb, { source: { sheet: 0, range: whole(DATA) }, target: { sheet: 0, row: 0, col: 5 } }).ops);
  const made = await save(doc);
  assert.match(await text(made, 'xl/worksheets/_rels/sheet1.xml.rels'), /pivotTable" Target="..\/pivotTables\/pivotTable1.xml"/);
  doc.history.undo();
  const back = await save(doc);
  assert.ok(same(back, original));
  doc.history.redo();
  assert.ok(same(await save(doc), made));
});

test('beside a pivot table of the file: the new cache takes a number the workbook does not use, and both tables are read back', async () => {
  const doc = await load(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] }));
  apply(doc, refreshPivot(doc.wb, 0));
  apply(doc, createPivot(doc.wb, { source: { sheet: 0, range: whole(DATA) }, rows: [1], values: [{ field: 2, subtotal: 'count' }] }).ops);
  const bytes = await save(doc);
  assert.match(await text(bytes, 'xl/workbook.xml'), /<pivotCache cacheId="7" r:id="rId3"\/><pivotCache cacheId="8" r:id="rId\d+"\/>/);
  const again = await load(bytes);
  assert.equal(again.wb.pivots!.tables.length, 2);
  assert.deepEqual(again.wb.pivots!.caches.map((cache) => cache.cacheId).sort(), [7, 8]);
  for (const table of again.wb.pivots!.tables) assert.equal(table.refusal, undefined);
  if (hasPython()) assert.deepEqual(independentPivots({ beside: bytes }), []);
});

test('at a cell of the sheet of the source: beside it, not over it, not over what holds something', async () => {
  const doc = await bookWith(DATA);
  const made = createPivot(doc.wb, { source: { sheet: 0, range: whole(DATA) }, target: { sheet: 0, row: 0, col: 5 } });
  assert.equal(made.sheet, 0);
  assert.deepEqual(made.range, { top: 0, left: 5, bottom: 3, right: 6 });
  assert.match(createPivotProblem(doc.wb, { source: { sheet: 0, range: whole(DATA) }, target: { sheet: 0, row: 3, col: 1 } }) || '', /holds something|over its source|not empty/);
  apply(doc, made.ops);
  assert.deepEqual(grid(doc.wb, 0, 0, 5, 3, 6), [
    ['Row Labels', 'Sum of Sales'],
    ['East', 16],
    ['West', 27],
    ['Grand Total', 43],
  ]);
  if (hasPython()) assert.deepEqual(independentPivots({ beside: await save(doc) }), []);
});

test('filters at creation stand above the body; the body moves down to make room', async () => {
  const doc = await bookWith(DATA);
  const made = createPivot(doc.wb, { source: { sheet: 0, range: whole(DATA) }, target: { sheet: 0, row: 0, col: 5 }, rows: [1], filters: [0], values: [{ field: 2, subtotal: 'count' }] });
  assert.deepEqual(made.range, { top: 2, left: 5, bottom: 5, right: 6 });
  apply(doc, made.ops);
  assert.deepEqual(grid(doc.wb, 0, 0, 5, 5, 6), [
    ['Region', '(All)'],
    [undefined, undefined],
    ['Row Labels', 'Count of Sales'],
    ['A', 3],
    ['B', 2],
    ['Grand Total', 5],
  ]);
  const bytes = await save(doc);
  const again = await load(bytes);
  assert.deepEqual(again.wb.pivots!.tables[0].pageFields, [{ fld: 0 }]);
  assert.deepEqual(again.wb.pivots!.tables[0].pages, { rows: 1, cols: 1 });
  if (hasPython()) assert.deepEqual(independentPivots({ filters: bytes }), []);
});

test('refused: an empty header, a header twice, nothing below the header, a layout not built here, a name taken', async () => {
  const empty = await bookWith([['Region', null, 'Sales'], ['East', 'A', 1]]);
  assert.match(createPivotProblem(empty.wb, { source: { sheet: 0, range: { top: 0, left: 0, bottom: 1, right: 2 } } }) || '', /header of column B1 of its source is empty/);
  const twice = await bookWith([['Sales', 'sales'], [1, 2]]);
  assert.match(createPivotProblem(twice.wb, { source: { sheet: 0, range: { top: 0, left: 0, bottom: 1, right: 1 } } }) || '', /two columns of its source have the header "sales"/);
  const doc = await bookWith(DATA);
  assert.match(createPivotProblem(doc.wb, { source: { sheet: 0, range: { top: 0, left: 0, bottom: 0, right: 2 } } }) || '', /no rows below its header/);
  assert.match(createPivotProblem(doc.wb, { source: { sheet: 0, range: whole(DATA) }, rows: [] }) || '', /needs a field in its rows/);
  assert.match(createPivotProblem(doc.wb, { source: { sheet: 0, range: whole(DATA) }, values: [] }) || '', /needs a value/);
  assert.match(createPivotProblem(doc.wb, { source: { sheet: 0, range: whole(DATA) }, rows: [0], cols: [0] }) || '', /stands in two places/);
  assert.match(createPivotProblem(doc.wb, { source: { sheet: 0, range: whole(DATA) }, cols: [1], values: [{ field: 2 }, { field: 2, subtotal: 'count' }] }) || '', /several values and fields in the columns/);
  assert.match(createPivotProblem(doc.wb, { source: { sheet: 0, range: whole(DATA) }, rows: [7] }) || '', /no field 7/);
  assert.match(createPivotProblem(doc.wb, { source: { sheet: 0, range: whole(DATA) }, values: [{ field: 2, subtotal: 'median' as never }] }) || '', /not known here/);
  apply(doc, createPivot(doc.wb, { source: { sheet: 0, range: whole(DATA) } }).ops);
  assert.match(createPivotProblem(doc.wb, { source: { sheet: 0, range: whole(DATA) }, name: 'pivottable1' }) || '', /called "pivottable1" already/);
  assert.equal(createPivot(doc.wb, { source: { sheet: 0, range: whole(DATA) } }).index, 1);
});

test('a workbook with pivot tables made here keeps its sheets\' names; a copy of the sheet says the table is not copied', async () => {
  const doc = await bookWith(DATA);
  assert.equal(renameSheetProblem(doc.wb, 0, 'Data'), undefined);
  apply(doc, createPivot(doc.wb, { source: { sheet: 0, range: whole(DATA) } }).ops);
  assert.match(renameSheetProblem(doc.wb, 0, 'Data') || '', /pivotTable/);
  assert.match(deleteSheetProblem(doc.wb, 1) || '', /pivotTable/);
  assert.deepEqual(copySheetLosses(doc.wb, 1), [{ what: 'pivotTable', count: 1 }]);
});

test('while the workbook is edited together a table is not made: its operation is not sent', async () => {
  const doc = await bookWith(DATA);
  const ops = createPivot(doc.wb, { source: { sheet: 0, range: whole(DATA) } }).ops;
  const add = ops.find((op) => op.type === 'addPivot') as Op;
  assert.throws(() => encodeOp(doc.wb, add), CollabUnsupportedError);
});

test('a defined name for its source; Turkish text, blanks, errors and dates in the fields', async () => {
  const data: Datum[][] = [
    ['Şehir', 'Tarih', 'Tutar'],
    ['İzmir', { date: 45292 }, 10],
    ['ığdır', { date: 45293 }, '#N/A'],
    ['Çorum', { date: 45292 }, 5],
    [null, { date: 45293 }, 2],
    ['İzmir', null, 3],
  ];
  const doc = await bookWith(data, '<definedName name="Veri">Sheet1!$A$1:$C$6</definedName>');
  const made = createPivot(doc.wb, { source: { name: 'Veri' }, rows: [0, 1], values: [{ field: 2 }] });
  apply(doc, made.ops);
  const ref = made.range;
  const body = grid(doc.wb, made.sheet, ref.top, ref.left, ref.bottom, ref.right);
  assert.deepEqual(body[0], ['Row Labels', 'Sum of Tutar']);
  assert.deepEqual(body.map((row) => row[0]), ['Row Labels', 'İzmir', 45292, '(blank)', 'ığdır', 45293, 'Çorum', 45292, '(blank)', 45293, 'Grand Total']);
  // The sum of a set with an error is the error.
  assert.equal(body[4][1], '#N/A');
  // A date label is shown as a date.
  const dateCell = doc.wb.sheets[made.sheet].rows.get(ref.top + 2)!.cells[0];
  assert.equal(doc.wb.styles.cellXfs[dateCell.style].numFmtId, 14);
  const bytes = await save(doc);
  const again = await load(bytes);
  assert.deepEqual(again.wb.pivots!.tables[0].cache.source, { kind: 'name', name: 'Veri' });
  // The field of dates has the format of its dates, as Excel gives it.
  assert.match(await text(bytes, 'xl/pivotCache/pivotCacheDefinition1.xml'), /<cacheField name="Tarih" numFmtId="14">/);
  assert.deepEqual(refreshPivot(again.wb, 0).filter((op) => op.type === 'setCells'), []);
  if (hasPython()) assert.deepEqual(independentPivots({ named: bytes }), []);
});

test('the parts are listed under names no part of the package has, and the workbook lists the cache under a free number', async () => {
  const doc = await bookWith(DATA);
  apply(doc, createPivot(doc.wb, { source: { sheet: 0, range: whole(DATA) } }).ops);
  apply(doc, createPivot(doc.wb, { source: { sheet: 0, range: whole(DATA) }, target: { sheet: 0, row: 0, col: 6 } }).ops);
  const bytes = await save(doc);
  const zip = ZipArchive.read(bytes);
  const names = zip.names().filter((name) => /pivot/i.test(name)).sort();
  assert.deepEqual(names, [
    'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels',
    'xl/pivotCache/_rels/pivotCacheDefinition2.xml.rels',
    'xl/pivotCache/pivotCacheDefinition1.xml',
    'xl/pivotCache/pivotCacheDefinition2.xml',
    'xl/pivotCache/pivotCacheRecords1.xml',
    'xl/pivotCache/pivotCacheRecords2.xml',
    'xl/pivotTables/_rels/pivotTable1.xml.rels',
    'xl/pivotTables/_rels/pivotTable2.xml.rels',
    'xl/pivotTables/pivotTable1.xml',
    'xl/pivotTables/pivotTable2.xml',
  ]);
  const book = await text(bytes, 'xl/workbook.xml');
  assert.match(book, /cacheId="0"[^>]*\/><pivotCache cacheId="1"/);
  const again = await load(bytes);
  assert.equal(again.wb.pivots!.tables.length, 2);
  // Read sheet by sheet: the table of the first sheet comes first.
  assert.deepEqual(tableState(again.wb.pivots!.tables[0]).location.ref, { top: 0, left: 6, bottom: 3, right: 7 });
  if (hasPython()) assert.deepEqual(independentPivots({ two: bytes }), []);
});

test('the writer of new files (buildXlsx) writes the cells of a table, not the table, and says so', async () => {
  const doc = await bookWith(DATA);
  apply(doc, createPivot(doc.wb, { source: { sheet: 0, range: whole(DATA) } }).ops);
  const report = { parts: [] as string[], warnings: [] as string[] };
  await buildXlsx(doc.wb, { report });
  assert.ok(report.warnings.indexOf('Pivot table: a pivot table is not written; the cells it shows are') >= 0, report.warnings.join('; '));
  assert.ok(!report.parts.some((part) => /pivot/i.test(part)));
});

test('rows inserted above a table made here and inside its source move them, as for a table of the file', async () => {
  const doc = await bookWith(DATA);
  apply(doc, createPivot(doc.wb, { source: { sheet: 0, range: whole(DATA) } }).ops);
  apply(doc, [{ type: 'insert', sheet: 1, axis: 'row', at: 0, count: 2 }]);
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 2, count: 1 }]);
  assert.deepEqual(tableState(doc.wb.pivots!.tables[0]).location.ref, { top: 4, left: 0, bottom: 7, right: 1 });
  assert.throws(() => doc.history.apply([{ type: 'delete', sheet: 0, axis: 'row', at: 0, count: 1 }], undefined, undefined), /its header row would go/);
  const bytes = await save(doc);
  const again = await load(bytes);
  assert.deepEqual(again.wb.pivots!.tables[0].cache.source, { kind: 'range', sheet: 'Sheet1', ref: 'A1:C7' });
  assert.deepEqual(tableState(again.wb.pivots!.tables[0]).location.ref, { top: 4, left: 0, bottom: 7, right: 1 });
  // The records stay as they were until a refresh (Excel keeps them too); refreshed, the file is whole again.
  again.history.apply(refreshPivot(again.wb, 0), undefined, undefined);
  if (hasPython()) assert.deepEqual(independentPivots({ moved: await save(again) }), []);
});
