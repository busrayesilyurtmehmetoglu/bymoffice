// Deleting pivot tables (src/ops/pivot-delete.ts): of the file and made here; the cells it clears;
// a cache two tables share; what is refused; the parts a save takes out and brings back; undo.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { ZipArchive } from '../src/index';
import { createPivot } from '../src/ops/pivot-create';
import { deletePivot, deletePivotProblem } from '../src/ops/pivot-delete';
import { refreshPivot } from '../src/ops/pivots';
import { deleteSheet, renameSheet } from '../src/ops/sheets';
import { applyOp, type Op } from '../src/ops/ops';
import { encodeOp } from '../src/collab/wire';
import { grid, hasPython, independentPivots, load, pivotBook, same, save, text, type Datum, type Loaded } from './pivot-fixtures';

const encoder = new TextEncoder();
const DATA: Datum[][] = [
  ['Region', 'Product', 'Sales'],
  ['East', 'A', 10],
  ['West', 'B', 20],
  ['East', 'B', 5],
  ['West', 'A', 7],
];
const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

const apply = (doc: Loaded, ops: Op[]) => doc.history.apply(ops, undefined, undefined);
const names = async (bytes: Uint8Array) => ZipArchive.read(bytes).names();

/** The parts of a package with what each holds, by name; the declared content types in their order of name. */
async function contents(bytes: Uint8Array): Promise<Record<string, string>> {
  const zip = ZipArchive.read(bytes);
  const out: Record<string, string> = {};
  for (const name of zip.names().slice().sort()) {
    let xml = new TextDecoder().decode(await zip.read(name));
    if (name === '[Content_Types].xml') xml = (xml.match(/<Override [^>]*\/>/g) || []).sort().join('');
    out[name] = xml;
  }
  return out;
}

async function edit(bytes: Uint8Array, change: (zip: ZipArchive) => void | Promise<void>): Promise<Uint8Array> {
  const zip = ZipArchive.read(bytes);
  await change(zip);
  return zip.write();
}

test('a table of the file: deleted, its cells are empty, its parts and its cache leave the package; undone and saved, the bytes of the file', async () => {
  const bytes = await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }], around: [[0, 5, 'note']] });
  const doc = await load(bytes);
  assert.equal(deletePivotProblem(doc.wb, 0), undefined);
  const ops = deletePivot(doc.wb, 0);
  assert.deepEqual(ops.map((op) => op.type), ['deletePivot', 'setCells']);
  apply(doc, ops);
  assert.equal(doc.wb.pivots!.tables.length, 0);
  assert.equal(doc.wb.pivots!.caches.length, 0);
  // The body (A3:B4) is empty, formats and all; what stands beside it stays.
  assert.deepEqual(grid(doc.wb, 1, 2, 0, 3, 1), [
    [undefined, undefined],
    [undefined, undefined],
  ]);
  assert.equal(doc.wb.sheets[1].rows.get(0)!.cells[5].value, 'note');
  const deleted = await save(doc);
  const list = await names(deleted);
  for (const part of ['xl/pivotTables/pivotTable1.xml', 'xl/pivotTables/_rels/pivotTable1.xml.rels', 'xl/pivotCache/pivotCacheDefinition1.xml', 'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels', 'xl/pivotCache/pivotCacheRecords1.xml']) {
    assert.ok(list.indexOf(part) < 0, part);
  }
  assert.doesNotMatch(await text(deleted, 'xl/workbook.xml'), /pivotCache/);
  assert.doesNotMatch(await text(deleted, 'xl/_rels/workbook.xml.rels'), /pivotCache/);
  assert.doesNotMatch(await text(deleted, 'xl/worksheets/_rels/sheet2.xml.rels'), /pivotTable/);
  assert.doesNotMatch(await text(deleted, '[Content_Types].xml'), /pivot/);
  // The rest of the workbook part is as it was (cells changed: a calculation is asked for on opening).
  assert.equal(
    await text(deleted, 'xl/workbook.xml'),
    (await text(bytes, 'xl/workbook.xml')).replace('<pivotCaches><pivotCache cacheId="7" r:id="rId3"/></pivotCaches>', '').replace('<calcPr calcId="191029"/>', '<calcPr calcId="191029" fullCalcOnLoad="1"/>'),
  );
  const again = await load(deleted);
  assert.equal(again.wb.pivots, undefined);
  assert.deepEqual(grid(again.wb, 1, 2, 0, 3, 1), [
    [undefined, undefined],
    [undefined, undefined],
  ]);
  if (hasPython()) assert.deepEqual(independentPivots({ deleted }), []);
  // Undone: the table is back, and so is every byte of the file.
  doc.history.undo();
  assert.equal(doc.wb.pivots!.tables.length, 1);
  assert.equal(doc.wb.pivots!.deleted, undefined);
  assert.ok(same(await save(doc), bytes), 'deleted, saved, undone and saved: the bytes of the file');
  // Done again: the file of the deletion.
  doc.history.redo();
  assert.ok(same(await save(doc), deleted));
});

test('deleted and undone before any save: the bytes of the file', async () => {
  const bytes = await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] });
  const doc = await load(bytes);
  apply(doc, deletePivot(doc.wb, 0));
  doc.history.undo();
  assert.ok(same(await save(doc), bytes));
});

test('a table refreshed and saved, then deleted and saved, then undone: the refreshed file; the refresh undone as well: the file', async () => {
  const bytes = await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] });
  const doc = await load(bytes);
  apply(doc, refreshPivot(doc.wb, 0));
  const refreshed = await save(doc);
  apply(doc, deletePivot(doc.wb, 0));
  const deleted = await save(doc);
  assert.ok((await names(deleted)).indexOf('xl/pivotTables/pivotTable1.xml') < 0);
  doc.history.undo();
  assert.ok(same(await save(doc), refreshed), 'the refreshed parts come back as they were written');
  doc.history.undo();
  assert.ok(same(await save(doc), bytes), 'and the parts of the file as the file had them');
  // The other way round: refresh undone first, while the table is deleted.
  const other = await load(bytes);
  apply(other, refreshPivot(other.wb, 0));
  await save(other);
  apply(other, deletePivot(other.wb, 0));
  await save(other);
  other.history.undo();
  other.history.undo();
  assert.ok(same(await save(other), bytes), 'the table comes back with the bytes of the file, not with those of the refresh');
});

test('two tables of one cache: the first goes alone, the cache stays and the other is still refreshed; the second takes the cache along', async () => {
  const bytes = await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }], twin: { row: 2, col: 5 } });
  const doc = await load(bytes);
  assert.equal(doc.wb.pivots!.tables.length, 2);
  apply(doc, deletePivot(doc.wb, 0));
  assert.equal(doc.wb.pivots!.tables.length, 1);
  assert.equal(doc.wb.pivots!.caches.length, 1);
  assert.deepEqual(doc.wb.pivots!.deleted!.map((entry) => entry.cache), [undefined]);
  // The part of the deleted table still names the cache in the package: it does not count any more.
  assert.equal(deletePivotProblem(doc.wb, 0), undefined);
  apply(doc, refreshPivot(doc.wb, 0));
  assert.deepEqual(grid(doc.wb, 1, 2, 5, 5, 6), [
    ['Row Labels', 'Sum of Sales'],
    ['East', 15],
    ['West', 27],
    ['Grand Total', 42],
  ]);
  const one = await save(doc);
  const list = await names(one);
  assert.ok(list.indexOf('xl/pivotTables/pivotTable1.xml') < 0);
  assert.ok(list.indexOf('xl/pivotTables/pivotTable2.xml') >= 0);
  assert.ok(list.indexOf('xl/pivotCache/pivotCacheDefinition1.xml') >= 0);
  assert.match(await text(one, 'xl/workbook.xml'), /<pivotCache cacheId="7" r:id="rId3"\/>/);
  assert.equal(await text(one, 'xl/worksheets/_rels/sheet2.xml.rels'), (await text(bytes, 'xl/worksheets/_rels/sheet2.xml.rels')).replace(/<Relationship Id="rId1"[^>]*\/>/, ''));
  const reread = await load(one);
  assert.equal(reread.wb.pivots!.tables.length, 1);
  assert.equal(reread.wb.pivots!.tables[0].name, 'PivotTable2');
  if (hasPython()) assert.deepEqual(independentPivots({ one }), []);
  // The second one: the cache goes with it.
  apply(doc, deletePivot(doc.wb, 0));
  assert.equal(doc.wb.pivots!.caches.length, 0);
  const none = await save(doc);
  assert.ok((await names(none)).every((name) => !/pivot/i.test(name)));
  assert.doesNotMatch(await text(none, 'xl/workbook.xml'), /pivotCache/);
  if (hasPython()) assert.deepEqual(independentPivots({ none }), []);
  doc.history.undo();
  doc.history.undo();
  doc.history.undo();
  assert.equal(doc.wb.pivots!.tables.length, 2);
  assert.deepEqual(doc.wb.pivots!.tables.map((table) => table.name), ['PivotTable1', 'PivotTable2']);
  assert.ok(same(await save(doc), bytes));
});

test('a cache a pivot chart or a slicer refers to: the deletion is refused, before anything changes', async () => {
  const chart = encoder.encode(`<c:chartSpace xmlns:c="${CHART_NS}"><c:pivotSource><c:name>Pivot!PivotTable1</c:name><c:fmtId val="0"/></c:pivotSource><c:chart/></c:chartSpace>`);
  const charted = await load(await edit(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] }), (zip) => zip.set('xl/charts/chart1.xml', chart)));
  assert.throws(() => deletePivot(charted.wb, 0), /Cannot delete the pivot table "PivotTable1": a pivot chart is made from a table of its cache/);
  assert.equal(charted.wb.pivots!.tables.length, 1);
  const sliced = await load(
    await edit(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] }), (zip) => {
      zip.set('xl/slicerCaches/slicerCache1.xml', encoder.encode('<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="S" sourceName="Region"/>'));
    }),
  );
  assert.match(deletePivotProblem(sliced.wb, 0) || '', /slicer or a timeline/);
  // Another part of the package names the cache.
  const named = await load(
    await edit(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] }), (zip) => {
      zip.set('xl/other/x.xml', encoder.encode('<x/>'));
      zip.set('xl/other/_rels/x.xml.rels', encoder.encode('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>'));
    }),
  );
  assert.match(deletePivotProblem(named.wb, 0) || '', /another part of the workbook refers to its cache/);
  // No table at the index.
  assert.throws(() => deletePivot(named.wb, 3), /No pivot table at index 3/);
  assert.throws(() => deletePivot(named.wb, 0.5), /No pivot table/);
});

test('a table the model does not hold all of: a table of the file is not deleted', async () => {
  // A second table part no sheet shows: the model holds no table for it.
  const bytes = await edit(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] }), async (zip) => {
    zip.set('xl/pivotTables/pivotTable9.xml', await zip.read('xl/pivotTables/pivotTable1.xml'));
  });
  const doc = await load(bytes);
  assert.match(deletePivotProblem(doc.wb, 0) || '', /the package has a pivot table no sheet shows/);
});

test('the sheets of a workbook whose last pivot table was deleted are renamed; undone, they are not', async () => {
  const doc = await load(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }], twin: { row: 2, col: 5 } }));
  assert.throws(() => renameSheet(doc.wb, 0, 'Source'), /pivot/);
  apply(doc, deletePivot(doc.wb, 0));
  assert.throws(() => renameSheet(doc.wb, 0, 'Source'), /pivot/, 'one table is left');
  apply(doc, deletePivot(doc.wb, 0));
  apply(doc, renameSheet(doc.wb, 0, 'Source'));
  assert.equal(doc.wb.sheets[0].name, 'Source');
  const saved = await save(doc);
  assert.match(await text(saved, 'xl/workbook.xml'), /name="Source"/);
  if (hasPython()) assert.deepEqual(independentPivots({ saved }), []);
  doc.history.undo();
  doc.history.undo();
  assert.throws(() => renameSheet(doc.wb, 0, 'Source'), /pivot/);
});

test('a table made here: deleted, its parts are not written; undone, saved, it is written as it was; two tables, the first deleted', async () => {
  const doc = await load(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] }));
  // The table of the file is brought up to date, for the independent check of what it shows.
  apply(doc, refreshPivot(doc.wb, 0));
  const first = createPivot(doc.wb, { source: { sheet: 0, range: { top: 0, left: 0, bottom: 4, right: 2 } }, target: { sheet: 0, row: 0, col: 6 } });
  apply(doc, first.ops);
  const second = createPivot(doc.wb, { source: { sheet: 0, range: { top: 0, left: 0, bottom: 4, right: 2 } }, target: { sheet: 0, row: 0, col: 10 } });
  apply(doc, second.ops);
  assert.equal(doc.wb.pivots!.tables.length, 3);
  const made = await save(doc);
  assert.ok((await names(made)).indexOf('xl/pivotTables/pivotTable3.xml') >= 0);
  // The first one made here goes: its cells, its parts, the entry of its cache.
  apply(doc, deletePivot(doc.wb, first.index));
  assert.equal(doc.wb.pivots!.deleted, undefined, 'a table made here is in no file');
  assert.deepEqual(grid(doc.wb, 0, 0, 6, 3, 7), [
    [undefined, undefined],
    [undefined, undefined],
    [undefined, undefined],
    [undefined, undefined],
  ]);
  const deleted = await save(doc);
  const book = await text(deleted, 'xl/workbook.xml');
  assert.equal((book.match(/<pivotCache /g) || []).length, 2);
  const reread = await load(deleted);
  assert.deepEqual(reread.wb.pivots!.tables.map((table) => table.name).sort(), ['PivotTable1', 'PivotTable3']);
  if (hasPython()) assert.deepEqual(independentPivots({ deleted }), []);
  doc.history.undo();
  // Undone: written as it was, the same parts under the same names, the same numbers. (Parts made here come back as new
  // entries of the archive, at its end: the order of the entries and of the content types is not that of the save before.)
  assert.deepEqual(await contents(await save(doc)), await contents(made));
  doc.history.redo();
  assert.deepEqual(await contents(await save(doc)), await contents(deleted));
});

test('deletePivot and restorePivot are each other\'s inverse, and check what they are given', async () => {
  const doc = await load(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }], twin: { row: 2, col: 5 } }));
  const tables = doc.wb.pivots!.tables.slice();
  const undo = applyOp(doc.wb, { type: 'deletePivot', sheet: 1, table: 1 });
  assert.deepEqual(undo, { type: 'restorePivot', sheet: 1, at: 1, table: tables[1] });
  assert.throws(() => applyOp(doc.wb, { ...undo, at: 5 } as Op), /cannot be put at 5/);
  const back = applyOp(doc.wb, undo);
  assert.deepEqual(back, { type: 'deletePivot', sheet: 1, table: 1 });
  assert.deepEqual(doc.wb.pivots!.tables, tables);
  assert.throws(() => applyOp(doc.wb, undo), /in the workbook already/);
  assert.throws(() => applyOp(doc.wb, { type: 'deletePivot', sheet: 1, table: -1 }), /No pivot table/);
});

test('a deletion is not shared with a room: the operation is refused before anything is sent', async () => {
  const doc = await load(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] }));
  assert.throws(() => encodeOp(doc.wb, { type: 'deletePivot', sheet: 1, table: 0 }), /cannot be deleted while the workbook is edited together/);
});

test('a table whose fields of the page stand above it: those cells are cleared as well', async () => {
  const bytes = await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }], pages: [{ fld: 1 }], at: { row: 3, col: 0 } });
  const doc = await load(bytes);
  assert.equal(doc.wb.sheets[1].rows.get(1)!.cells[0].value, 'Product');
  apply(doc, deletePivot(doc.wb, 0));
  assert.deepEqual(grid(doc.wb, 1, 0, 0, 5, 1), [
    [undefined, undefined],
    [undefined, undefined],
    [undefined, undefined],
    [undefined, undefined],
    [undefined, undefined],
    [undefined, undefined],
  ]);
  const saved = await save(doc);
  if (hasPython()) assert.deepEqual(independentPivots({ saved }), []);
  doc.history.undo();
  assert.ok(same(await save(doc), bytes));
});

test('the last table deleted, its sheet deleted and the other renamed: the file holds no pivot part; all undone, the bytes of the file', async () => {
  const bytes = await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] });
  const doc = await load(bytes);
  apply(doc, deletePivot(doc.wb, 0));
  const deleted = await save(doc);
  apply(doc, deleteSheet(doc.wb, 1));
  apply(doc, renameSheet(doc.wb, 0, 'Source'));
  const saved = await save(doc);
  assert.ok((await names(saved)).every((name) => !/pivot|sheet2/i.test(name)), (await names(saved)).join(', '));
  const again = await load(saved);
  assert.deepEqual(again.wb.sheets.map((sheet) => sheet.name), ['Source']);
  if (hasPython()) assert.deepEqual(independentPivots({ saved }), []);
  doc.history.undo();
  doc.history.undo();
  // The sheet is back with its parts. (The content type of a part that comes back is declared last: the order of the
  // declarations is that of the file only once the package declares what the file did.)
  assert.deepEqual(await contents(await save(doc)), await contents(deleted));
  doc.history.undo();
  assert.ok(same(await save(doc), bytes));
});

test('deleted, saved and undone, the table is refreshed and saved: its parts are written from what the file had', async () => {
  const bytes = await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] });
  const doc = await load(bytes);
  apply(doc, deletePivot(doc.wb, 0));
  await save(doc);
  doc.history.undo();
  // The package lacks the parts until the next save puts them back; the refresh reads them from what was kept.
  apply(doc, refreshPivot(doc.wb, 0));
  const refreshed = await save(doc);
  const again = await load(refreshed);
  assert.equal(again.wb.pivots!.tables.length, 1);
  assert.deepEqual(grid(again.wb, 1, 2, 0, 5, 1), [
    ['Row Labels', 'Sum of Sales'],
    ['East', 15],
    ['West', 27],
    ['Grand Total', 42],
  ]);
  if (hasPython()) assert.deepEqual(independentPivots({ refreshed }), []);
  // The same as refreshing the file that was never deleted.
  const plain = await load(bytes);
  apply(plain, refreshPivot(plain.wb, 0));
  assert.deepEqual(await contents(refreshed), await contents(await save(plain)));
});
