// Rows and columns inserted and deleted in a workbook with pivot tables: the range of the source
// of a cache and the place of a table follow, what would cut through a table or take what a source
// needs is refused, and the parts are written with only those ranges changed.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { ZipArchive } from '../src/index';
import { applyWire } from '../src/collab/wire';
import { contentVersion } from '../src/collab/protocol';
import { applyOp, type Op } from '../src/ops/ops';
import { refreshPivot } from '../src/ops/pivots';
import { pivotPlacesAfter, pivotSourceLostBy, pivotsFollow } from '../src/ops/pivot-places';
import { patchedAttribute } from '../src/formats/xlsx/pivot-write';
import { mergeProblem } from '../src/ops/format';
import { cacheSourceNow, fieldColumnsNow, namedSourceRange, sourceArea, tableState } from '../src/model/sheet/pivot';
import { parseXml, serializeXml, type XmlDocument } from '../src/xml';
import { WorkbookView } from '../src/ui/sheet-view';
import { STRINGS_EN, STRINGS_TR } from '../src/ui/strings';
import { act, assertSame, int, join, Net, pick, saveOf, seeded } from './collab-fixtures';
import { grid, hasPython, independentPivots, load, pivotBook, same, save, setData, text, type Datum, type Loaded, type PivotSpec } from './pivot-fixtures';

const DATA: Datum[][] = [
  ['Region', 'Product', 'Sales'],
  ['East', 'A', 10],
  ['West', 'B', 20],
  ['East', 'B', 5],
  ['West', 'A', 7],
  ['East', 'A', 1],
];
const SPEC: PivotSpec = { data: DATA, rows: [0], values: [{ fld: 2 }] };
const CACHE = 'xl/pivotCache/pivotCacheDefinition1.xml';
const RECORDS = 'xl/pivotCache/pivotCacheRecords1.xml';
const TABLE = 'xl/pivotTables/pivotTable1.xml';

const ins = (sheet: number, axis: 'row' | 'col', at: number, count = 1): Op => ({ type: 'insert', sheet, axis, at, count });
const del = (sheet: number, axis: 'row' | 'col', at: number, count = 1): Op => ({ type: 'delete', sheet, axis, at, count });
const apply = (doc: Loaded, op: Op) => doc.history.apply([op], undefined, undefined);
const sourceOf = async (bytes: Uint8Array) => /<worksheetSource [^>]*>/.exec(await text(bytes, CACHE))![0];
const locationOf = async (bytes: Uint8Array, part = TABLE) => /<location [^>]*>/.exec(await text(bytes, part))![0];

/** The text of a part with one stretch of it replaced: what a part changed in place should be. */
async function only(bytes: Uint8Array, original: Uint8Array, part: string, from: string, to: string): Promise<void> {
  const before = await text(original, part);
  assert.equal(before.split(from).length, 2, `${from} stands once in ${part}`);
  assert.equal(await text(bytes, part), before.replace(from, to));
}

test('a workbook whose pivot tables the model holds lets rows and columns move', async () => {
  const doc = await load(await pivotBook(SPEC));
  assert.ok(pivotsFollow(doc.wb));
  assert.deepEqual(doc.wb.sheets[0].anchored, ['pivotTable']);
});

test('rows inserted above the source move it; only the range of the source is written anew', async () => {
  const bytes = await pivotBook(SPEC);
  const doc = await load(bytes);
  apply(doc, ins(0, 'row', 0, 2));
  assert.equal(doc.wb.pivots!.caches[0].sourceRef, 'A3:C8');
  // The table stands on another sheet: it stays.
  assert.equal(doc.wb.pivots!.tables[0].placed, undefined);
  const saved = await save(doc);
  await only(saved, bytes, CACHE, 'ref="A1:C6"', 'ref="A3:C8"');
  // Neither the records nor the table are written.
  assert.equal(await text(saved, RECORDS), await text(bytes, RECORDS));
  assert.equal(await text(saved, TABLE), await text(bytes, TABLE));
  const reread = await load(saved);
  assert.deepEqual(reread.wb.pivots!.caches[0].source, { kind: 'range', sheet: 'Data', ref: 'A3:C8' });
  assert.deepEqual(await refreshedBody(saved), [['Row Labels', 'Sum of Sales'], ['East', 16], ['West', 27], ['Grand Total', 43]]);
});

/**
 * Opens a saved file, refreshes every table and gives the body of the first as values; the file
 * that refresh writes is checked by the independent reader (the stand-in tables of the fixtures
 * are out of date until they are refreshed).
 */
async function refreshedBody(saved: Uint8Array): Promise<unknown[][]> {
  const doc = await load(saved);
  for (let i = 0; i < doc.wb.pivots!.tables.length; i++) if (!doc.wb.pivots!.tables[i].now) doc.history.apply(refreshPivot(doc.wb, i), undefined, undefined);
  if (hasPython()) assert.deepEqual(independentPivots({ refreshed: await save(doc) }), []);
  const table = doc.wb.pivots!.tables[0];
  const ref = tableState(table).location.ref;
  return grid(doc.wb, doc.wb.sheets.indexOf(table.sheet), ref.top, ref.left, ref.bottom, ref.right);
}

test('rows inserted inside the source make it longer; at its header or right after it they do not', async () => {
  const cases: [number, string | undefined][] = [
    [0, 'A2:C7'], // at the header: the source moves down
    [1, 'A1:C7'], // under the header: it grows
    [5, 'A1:C7'], // before its last row: it grows
    [6, undefined], // right after its last row: nothing
  ];
  for (const [at, ref] of cases) {
    const doc = await load(await pivotBook(SPEC));
    apply(doc, ins(0, 'row', at));
    assert.equal(doc.wb.pivots!.caches[0].sourceRef, ref, `insert at row ${at + 1}`);
  }
});

test('columns inserted before or inside the source move it or make it wider', async () => {
  const doc = await load(await pivotBook(SPEC));
  apply(doc, ins(0, 'col', 0));
  assert.equal(doc.wb.pivots!.caches[0].sourceRef, 'B1:D6');
  apply(doc, ins(0, 'col', 2));
  assert.equal(doc.wb.pivots!.caches[0].sourceRef, 'B1:E6');
  // A column after it does nothing; deleting the one before it moves it back.
  apply(doc, ins(0, 'col', 5));
  assert.equal(doc.wb.pivots!.caches[0].sourceRef, 'B1:E6');
  doc.history.undo();
  doc.history.undo();
  apply(doc, del(0, 'col', 0));
  assert.equal(doc.wb.pivots!.caches[0].sourceRef, undefined);
});

test('a deletion that takes a column of the source, its header row or all of it is refused; rows under the header shorten it', async () => {
  const doc = await load(await pivotBook(SPEC));
  const rows = doc.wb.sheets[0].rows;
  for (const op of [del(0, 'col', 1), del(0, 'col', 0, 3), del(0, 'col', 2, 5), del(0, 'row', 0), del(0, 'row', 0, 6), del(0, 'row', 0, 10)]) {
    assert.throws(() => applyOp(doc.wb, op), /^RangeError: Cannot delete (row|col)s of the source of the pivot table "PivotTable1": (its header row|a column of it) would go$/);
    assert.equal(doc.wb.sheets[0].rows, rows, 'nothing changed');
  }
  assert.equal(doc.wb.pivots!.caches[0].sourceRef, undefined);
  apply(doc, del(0, 'row', 2, 2));
  assert.equal(doc.wb.pivots!.caches[0].sourceRef, 'A1:C4');
  // Its last rows as well: the source ends higher.
  apply(doc, del(0, 'row', 3));
  assert.equal(doc.wb.pivots!.caches[0].sourceRef, 'A1:C3');
  // A column after it is no column of it.
  apply(doc, del(0, 'col', 3));
  assert.equal(doc.wb.pivots!.caches[0].sourceRef, 'A1:C3');
});

test('rows and columns inserted above or to the left of a table move it, as a whole', async () => {
  const bytes = await pivotBook(SPEC);
  const doc = await load(bytes);
  const table = doc.wb.pivots!.tables[0];
  // The stand-in of the fixture: a header and a grand total.
  assert.deepEqual(table.location.ref, { top: 2, left: 0, bottom: 3, right: 1 });
  apply(doc, ins(1, 'row', 0, 3));
  assert.deepEqual(table.placed, { top: 5, left: 0, bottom: 6, right: 1 });
  apply(doc, ins(1, 'col', 0));
  assert.deepEqual(tableState(table).location.ref, { top: 5, left: 1, bottom: 6, right: 2 });
  // The source on another sheet stays.
  assert.equal(doc.wb.pivots!.caches[0].sourceRef, undefined);
  // The cells moved with it: its header is where the table now starts.
  assert.equal(doc.wb.sheets[1].rows.get(5)!.cells[1]!.value, 'Row Labels');
  const saved = await save(doc);
  await only(saved, bytes, TABLE, 'ref="A3:B4"', 'ref="B6:C7"');
  assert.equal(await text(saved, CACHE), await text(bytes, CACHE));
  // Below it and to its right nothing moves.
  apply(doc, ins(1, 'row', 7));
  apply(doc, ins(1, 'col', 3));
  assert.deepEqual(table.placed, { top: 5, left: 1, bottom: 6, right: 2 });
  // Rows deleted above it move it up.
  apply(doc, del(1, 'row', 0, 2));
  assert.deepEqual(table.placed, { top: 3, left: 1, bottom: 4, right: 2 });
  // Opened again and refreshed, the table is built where it now stands.
  const again = await save(doc);
  assert.deepEqual(await refreshedBody(again), [['Row Labels', 'Sum of Sales'], ['East', 16], ['West', 27], ['Grand Total', 43]]);
  const reread = await load(again);
  reread.history.apply(refreshPivot(reread.wb, 0), undefined, undefined);
  assert.deepEqual(tableState(reread.wb.pivots!.tables[0]).location.ref, { top: 3, left: 1, bottom: 6, right: 2 });
});

test('an insertion or a deletion through a table is refused, and nothing changes', async () => {
  const doc = await load(await pivotBook(SPEC));
  const rows = doc.wb.sheets[1].rows;
  const cases: [Op, RegExp][] = [
    [ins(1, 'row', 3), /^RangeError: Cannot insert rows through the pivot table "PivotTable1"$/],
    [ins(1, 'row', 3, 5), /^RangeError: Cannot insert rows through the pivot table "PivotTable1"$/],
    [ins(1, 'col', 1), /^RangeError: Cannot insert cols through the pivot table "PivotTable1"$/],
    [del(1, 'row', 2), /^RangeError: Cannot delete rows through the pivot table "PivotTable1"$/],
    [del(1, 'row', 0, 3), /^RangeError: Cannot delete rows through the pivot table "PivotTable1"$/],
    [del(1, 'row', 3, 4), /^RangeError: Cannot delete rows through the pivot table "PivotTable1"$/],
    [del(1, 'row', 0, 20), /^RangeError: Cannot delete rows through the pivot table "PivotTable1"$/],
    [del(1, 'col', 0), /^RangeError: Cannot delete cols through the pivot table "PivotTable1"$/],
    [del(1, 'col', 1, 3), /^RangeError: Cannot delete cols through the pivot table "PivotTable1"$/],
  ];
  for (const [op, error] of cases) {
    assert.throws(() => applyOp(doc.wb, op), error);
    assert.equal(doc.wb.sheets[1].rows, rows);
    assert.equal(doc.wb.pivots!.tables[0].placed, undefined);
  }
  assert.equal(doc.wb.pivotsModified, undefined);
});

test('the fields of the page and the blank row under them belong to the table', async () => {
  const doc = await load(await pivotBook({ data: DATA, rows: [1], values: [{ fld: 2 }], pages: [{ fld: 0, item: 'East' }], at: { row: 4, col: 0 } }));
  const table = doc.wb.pivots!.tables[0];
  assert.deepEqual(table.pages, { rows: 1, cols: 1 });
  // The page field stands in row 3, the blank row is row 4, the body starts in row 5.
  assert.throws(() => applyOp(doc.wb, ins(1, 'row', 3)), /through the pivot table/);
  assert.throws(() => applyOp(doc.wb, del(1, 'row', 3)), /through the pivot table/);
  assert.throws(() => applyOp(doc.wb, del(1, 'row', 2)), /through the pivot table/);
  apply(doc, ins(1, 'row', 2));
  assert.equal(table.placed!.top, 5);
  assert.equal(doc.wb.sheets[1].rows.get(3)!.cells[0]!.value, 'Region');
  // Its two columns (caption and item) as well.
  assert.throws(() => applyOp(doc.wb, ins(1, 'col', 1)), /through the pivot table/);
});

test('two tables of one cache: each moves for itself, the cache once', async () => {
  const bytes = await pivotBook({ ...SPEC, twin: { row: 2, col: 4 } });
  const doc = await load(bytes);
  const [one, two] = doc.wb.pivots!.tables;
  apply(doc, ins(1, 'col', 3, 2));
  assert.equal(one.placed, undefined);
  assert.deepEqual(two.placed, { top: 2, left: 6, bottom: 3, right: 7 });
  apply(doc, ins(1, 'row', 0));
  assert.deepEqual(one.placed, { top: 3, left: 0, bottom: 4, right: 1 });
  assert.deepEqual(two.placed, { top: 3, left: 6, bottom: 4, right: 7 });
  apply(doc, ins(0, 'row', 0));
  assert.equal(doc.wb.pivots!.caches[0].sourceRef, 'A2:C7');
  const saved = await save(doc);
  await only(saved, bytes, TABLE, 'ref="A3:B4"', 'ref="A4:B5"');
  await only(saved, bytes, 'xl/pivotTables/pivotTable2.xml', 'ref="E3:F4"', 'ref="G4:H5"');
  await only(saved, bytes, CACHE, 'ref="A1:C6"', 'ref="A2:C7"');
  assert.deepEqual(await refreshedBody(saved), [['Row Labels', 'Sum of Sales'], ['East', 16], ['West', 27], ['Grand Total', 43]]);
});

test('a source given by a defined name: the name follows, the cache part is not changed', async () => {
  const bytes = await pivotBook({ ...SPEC, sourceName: 'Sales' });
  const doc = await load(bytes);
  assert.ok(pivotsFollow(doc.wb));
  apply(doc, ins(0, 'row', 3));
  assert.equal(doc.wb.definedNames.find((name) => name.name === 'Sales')!.value, 'Data!$A$1:$C$7');
  assert.equal(doc.wb.pivots!.caches[0].sourceRef, undefined);
  assert.throws(() => applyOp(doc.wb, del(0, 'col', 1)), /of the source of the pivot table "PivotTable1": a column of it would go/);
  assert.throws(() => applyOp(doc.wb, del(0, 'row', 0)), /its header row would go/);
  const saved = await save(doc);
  assert.equal(await text(saved, CACHE), await text(bytes, CACHE));
  // Filled and refreshed, the table reads the longer range.
  setData(doc, 3, 0, 'North');
  setData(doc, 3, 1, 'C');
  setData(doc, 3, 2, 100);
  doc.history.apply(refreshPivot(doc.wb, 0), undefined, undefined);
  const ref = tableState(doc.wb.pivots!.tables[0]).location.ref;
  // A value new to the source comes after the items the cache had (the field is sorted by hand).
  assert.deepEqual(grid(doc.wb, 1, ref.top, ref.left, ref.bottom, ref.right), [
    ['Row Labels', 'Sum of Sales'],
    ['East', 16],
    ['West', 27],
    ['North', 100],
    ['Grand Total', 143],
  ]);
});

test('insert, undo, save gives the file back byte for byte; so does a save in between', async () => {
  const bytes = await pivotBook({ ...SPEC, twin: { row: 2, col: 4 } });
  const doc = await load(bytes);
  apply(doc, ins(0, 'row', 2));
  apply(doc, ins(1, 'row', 0));
  apply(doc, ins(1, 'col', 3));
  doc.history.undo();
  doc.history.undo();
  doc.history.undo();
  assert.ok(same(await save(doc), bytes));

  const again = await load(bytes);
  apply(again, ins(0, 'row', 2));
  apply(again, ins(1, 'row', 0));
  const between = await save(again);
  assert.ok(!same(between, bytes));
  again.history.undo();
  again.history.undo();
  assert.ok(same(await save(again), bytes));
  // And redone, the parts are those of the save in between.
  again.history.redo();
  again.history.redo();
  assert.ok(same(await save(again), between));
});

test('insert, save, open again, refresh: the table is built from the longer source', async () => {
  const doc = await load(await pivotBook(SPEC));
  apply(doc, ins(0, 'row', 3, 1));
  setData(doc, 3, 0, 'North');
  setData(doc, 3, 1, 'C');
  setData(doc, 3, 2, 100);
  apply(doc, ins(1, 'row', 0, 2));
  const saved = await save(doc);
  const reread = await load(saved);
  assert.equal(reread.wb.pivots!.caches[0].source!.kind === 'range' && reread.wb.pivots!.caches[0].source!.ref, 'A1:C7');
  assert.deepEqual(reread.wb.pivots!.tables[0].location.ref, { top: 4, left: 0, bottom: 5, right: 1 });
  reread.history.apply(refreshPivot(reread.wb, 0), undefined, undefined);
  const ref = tableState(reread.wb.pivots!.tables[0]).location.ref;
  assert.deepEqual(ref, { top: 4, left: 0, bottom: 8, right: 1 });
  assert.deepEqual(grid(reread.wb, 1, ref.top, ref.left, ref.bottom, ref.right), [
    ['Row Labels', 'Sum of Sales'],
    ['East', 16],
    ['West', 27],
    ['North', 100],
    ['Grand Total', 143],
  ]);
  const refreshed = await save(reread);
  if (hasPython()) assert.deepEqual(independentPivots({ refreshed }), []);
});

test('a refreshed table moves in its refreshed state; undone, each step puts back its own', async () => {
  const bytes = await pivotBook(SPEC);
  const doc = await load(bytes);
  doc.history.apply(refreshPivot(doc.wb, 0), undefined, undefined);
  const table = doc.wb.pivots!.tables[0];
  const refreshedAt = tableState(table).location.ref;
  apply(doc, ins(1, 'row', 0, 2));
  assert.deepEqual(table.now!.location.ref, { ...refreshedAt, top: refreshedAt.top + 2, bottom: refreshedAt.bottom + 2 });
  // Without the refresh it would stand two rows lower as well.
  assert.deepEqual(table.placed, { top: 4, left: 0, bottom: 5, right: 1 });
  const saved = await save(doc);
  assert.match(await locationOf(saved), /ref="A5:B8"/);
  if (hasPython()) assert.deepEqual(independentPivots({ saved }), []);
  doc.history.undo();
  assert.deepEqual(table.now!.location.ref, refreshedAt);
  assert.equal(table.placed, undefined);
  doc.history.undo();
  assert.equal(table.now, undefined);
  assert.ok(same(await save(doc), bytes));
});

test('cells of a moved table cannot be merged where it now stands', async () => {
  const doc = await load(await pivotBook(SPEC));
  apply(doc, ins(1, 'row', 0, 4));
  assert.equal(mergeProblem(doc.wb, 1, { top: 6, left: 0, bottom: 6, right: 1 }), 'Cannot merge cells of the pivot table "PivotTable1"');
  assert.equal(mergeProblem(doc.wb, 1, { top: 2, left: 0, bottom: 2, right: 1 }), undefined);
});

test('a pivot table of the package that the model does not hold keeps rows and columns where they are', async () => {
  const bytes = await pivotBook(SPEC);
  const zip = ZipArchive.read(bytes);
  // A table part no sheet shows.
  zip.set('xl/pivotTables/pivotTable9.xml', await zip.read('xl/pivotTables/pivotTable1.xml'));
  const doc = await load(await zip.write());
  assert.ok(!pivotsFollow(doc.wb));
  assert.equal(doc.wb.pivots!.fixed, 'the package has a pivot table no sheet shows');
  assert.throws(() => applyOp(doc.wb, ins(0, 'row', 0)), /^RangeError: Cannot insert rows in a sheet with pivotTable: not supported yet$/);

  // A source that is not one range of cells.
  const odd = ZipArchive.read(bytes);
  odd.set(CACHE, new TextEncoder().encode((await text(bytes, CACHE)).replace('ref="A1:C6"', 'ref="A1:C6 E1:E6"')));
  const other = await load(await odd.write());
  assert.match(other.wb.pivots!.caches[0].placesFixed || '', /is not one range of cells/);
  assert.throws(() => applyOp(other.wb, ins(1, 'row', 0)), /not supported yet/);
});

test('a source of whole columns does not move with rows; deleting its first row is refused', async () => {
  const bytes = await pivotBook(SPEC);
  const zip = ZipArchive.read(bytes);
  zip.set(CACHE, new TextEncoder().encode((await text(bytes, CACHE)).replace('ref="A1:C6"', 'ref="$A:$C"')));
  const doc = await load(await zip.write());
  assert.ok(pivotsFollow(doc.wb));
  apply(doc, ins(0, 'row', 3));
  assert.equal(doc.wb.pivots!.caches[0].sourceRef, undefined);
  apply(doc, ins(0, 'col', 0));
  assert.equal(doc.wb.pivots!.caches[0].sourceRef, '$B:$D');
  assert.throws(() => applyOp(doc.wb, del(0, 'row', 0)), /its header row would go/);
});

test('the range is changed in the text of the part: its other bytes, comments and line ends stay', () => {
  const xml = '<?xml version="1.0" encoding="UTF-8"?>\r\n<!-- <worksheetSource ref="Z1"/> -->\r\n<p:pivotCacheDefinition xmlns:p="urn:x"><p:cacheSource type="worksheet"><p:worksheetSource sheet="D" ref = \'A1:C6\' x:y="1" xmlns:x="urn:y"/></p:cacheSource><![CDATA[<worksheetSource ref="Q1"/>]]></p:pivotCacheDefinition>';
  const out = patchedAttribute(new TextEncoder().encode(xml), ['pivotCacheDefinition', 'cacheSource', 'worksheetSource'], 'ref', 'A2:C7');
  assert.ok(out instanceof Uint8Array);
  assert.equal(new TextDecoder().decode(out as Uint8Array), xml.replace("'A1:C6'", "'A2:C7'"));
  // A part in another encoding is written from its tree, with a declaration that says what it is.
  const utf16 = '﻿<?xml version="1.0" encoding="UTF-16"?><pivotCacheDefinition><cacheSource><worksheetSource ref="A1:C6" sheet="D"/></cacheSource></pivotCacheDefinition>';
  const bytes = new Uint8Array(utf16.length * 2);
  for (let i = 0; i < utf16.length; i++) {
    bytes[2 * i] = utf16.charCodeAt(i) & 255;
    bytes[2 * i + 1] = utf16.charCodeAt(i) >> 8;
  }
  const tree = patchedAttribute(bytes, ['pivotCacheDefinition', 'cacheSource', 'worksheetSource'], 'ref', 'A2:C7') as XmlDocument;
  assert.ok(!(tree instanceof Uint8Array));
  const written = serializeXml(tree);
  assert.match(written, /encoding="UTF-8"/);
  assert.match(written, /<worksheetSource ref="A2:C7" sheet="D"\/>/);
  // An attribute that stands twice along the path is changed in the first element, as the reader reads it.
  const twice = '<pivotCacheDefinition><cacheSource><worksheetSource ref="A1"/><worksheetSource ref="B1"/></cacheSource></pivotCacheDefinition>';
  const fixed = patchedAttribute(new TextEncoder().encode(twice), ['pivotCacheDefinition', 'cacheSource', 'worksheetSource'], 'ref', 'C1') as XmlDocument;
  assert.equal(serializeXml(fixed), '<pivotCacheDefinition><cacheSource><worksheetSource ref="C1"/><worksheetSource ref="B1"/></cacheSource></pivotCacheDefinition>');
  assert.throws(() => patchedAttribute(new TextEncoder().encode('<pivotCacheDefinition/>'), ['pivotCacheDefinition', 'cacheSource', 'worksheetSource'], 'ref', 'C1'), /has no pivotCacheDefinition\/cacheSource\/worksheetSource with ref/);
  assert.ok(parseXml(new TextEncoder().encode(xml)));
});

test('the view says what is refused in the language of the user', async () => {
  const doc = await load(await pivotBook(SPEC));
  const said = (op: Op, strings: typeof STRINGS_EN): string => {
    try {
      applyOp(doc.wb, op);
    } catch (error) {
      return WorkbookView.prototype.localizeError.call({ strings } as unknown as WorkbookView, error);
    }
    return assert.fail('the edit was made');
  };
  assert.equal(said(ins(1, 'row', 3), STRINGS_EN), 'Cannot insert rows through the pivot table "PivotTable1"');
  assert.equal(said(del(1, 'col', 1), STRINGS_TR), '"PivotTable1" özet tablosunun içinden sütun silinemez');
  assert.equal(said(del(0, 'col', 1), STRINGS_EN), 'Columns of the fields of the source of the pivot table "PivotTable1" cannot be deleted');
  assert.equal(said(del(0, 'row', 0), STRINGS_TR), '"PivotTable1" özet tablosunun kaynağındaki başlık satırı silinemez');
  const view = (strings: typeof STRINGS_EN, message: string) => WorkbookView.prototype.localizeError.call({ strings } as unknown as WorkbookView, new RangeError(message));
  assert.equal(view(STRINGS_TR, 'Cannot insert rows: the pivot table "T" would be pushed off the sheet'), 'Satır eklenemez: "T" özet tablosu sayfanın dışına itilirdi');
  assert.equal(view(STRINGS_EN, 'Cannot delete cols: the source of the pivot table "T" is not one range of cells'), 'Cannot delete columns: the source of the pivot table "T" is not one range of cells');
  assert.equal(view(STRINGS_TR, 'Cannot delete cols: the source of the pivot table "T" is not one range of cells'), 'Sütun silinemez: "T" özet tablosunun kaynağı tek bir hücre aralığı değil');
  assert.equal(view(STRINGS_TR, 'Cannot insert rows in a workbook with pivot tables: the package has a pivot table no sheet shows'), 'Bu çalışma kitabında özet tabloları yüzünden satır eklenemez: pakette hiçbir sayfanın göstermediği bir pivot tablo var');
  assert.equal(view(STRINGS_EN, 'Cannot insert rows in a workbook with pivot tables: a pivot table cannot be read'), 'Cannot insert rows in this workbook because of its pivot tables: a pivot table cannot be read');
});

test('nothing is worked out for an edit that touches no pivot table', async () => {
  const doc = await load(await pivotBook(SPEC));
  assert.equal(pivotPlacesAfter(doc.wb, doc.wb.sheets[0], { kind: 'insert', axis: 'row', at: 6, count: 1 }), undefined);
  assert.equal(pivotPlacesAfter(doc.wb, doc.wb.sheets[1], { kind: 'insert', axis: 'row', at: 6, count: 1 }), undefined);
});

test('the last rows of a source do not come back with rows inserted where they were: an undo of a shared step says so', async () => {
  const doc = await load(await pivotBook(SPEC));
  const sheet = doc.wb.sheets[0];
  assert.equal(pivotSourceLostBy(doc.wb, sheet, { kind: 'delete', axis: 'row', at: 4, count: 2 }), true);
  assert.equal(pivotSourceLostBy(doc.wb, sheet, { kind: 'delete', axis: 'row', at: 2, count: 2 }), false);
  assert.equal(pivotSourceLostBy(doc.wb, sheet, { kind: 'delete', axis: 'row', at: 7, count: 2 }), false);
  assert.equal(pivotSourceLostBy(doc.wb, doc.wb.sheets[1], { kind: 'delete', axis: 'row', at: 4, count: 2 }), false);
  // A shared deletion says its undo does not bring everything back.
  assert.equal(applyWire(doc.wb, { t: 'del', s: 0, x: 'row', at: 2, n: 1 }).lossy, false);
  assert.equal(applyWire(doc.wb, { t: 'del', s: 0, x: 'row', at: 3, n: 2 }).lossy, true);
  assert.equal(doc.wb.pivots!.caches[0].sourceRef, 'A1:C3');
});

test('editing together: every participant moves the pivot tables alike from the edit alone', async () => {
  const bytes = await pivotBook({ ...SPEC, twin: { row: 2, col: 4 } });
  const net = new Net();
  const ada = await join(net, 'Ada', bytes);
  const bob = await join(net, 'Bob', bytes);
  net.drain(() => 0);
  assert.ok(act(ada, () => [ins(0, 'row', 2, 2)]));
  assert.ok(act(bob, () => [ins(1, 'row', 0, 1)]));
  assert.ok(act(bob, () => [ins(1, 'col', 3, 1)]));
  // Through a table: refused before anything is sent.
  assert.equal(act(ada, () => [ins(1, 'row', 3)]), false);
  net.drain(() => 0.5);
  assertSame([ada, bob], 'after the edits');
  for (const member of [ada, bob]) {
    const pivots = member.wb.pivots!;
    assert.equal(pivots.caches[0].sourceRef, 'A1:C8', member.name);
    assert.deepEqual(pivots.tables[0].placed, { top: 3, left: 0, bottom: 4, right: 1 }, member.name);
    assert.deepEqual(pivots.tables[1].placed, { top: 3, left: 5, bottom: 4, right: 6 }, member.name);
  }
  const [a, b] = [await saveOf(ada), await saveOf(bob)];
  for (const part of [CACHE, TABLE, 'xl/pivotTables/pivotTable2.xml']) assert.equal(await text(a, part), await text(b, part), part);
  assert.match(await sourceOf(a), /ref="A1:C8"/);
});

/** Where the pivot tables of a workbook and the sources of their caches stand: what every participant has to agree on. */
const placesOf = (wb: Loaded['wb']) => JSON.stringify({ caches: wb.pivots!.caches.map((cache) => cache.sourceRef || null), tables: wb.pivots!.tables.map((table) => tableState(table).location.ref) });

test('editing together: edits that cross on their way, undone and redone, leave every participant with the same pivot tables', async () => {
  const bytes = await pivotBook({ ...SPEC, twin: { row: 2, col: 4 } });
  let moved = 0;
  for (let seed = 1; seed <= 12; seed++) {
    const random = seeded(seed);
    const net = new Net();
    const members = [await join(net, 'Ada', bytes), await join(net, 'Bob', bytes), await join(net, 'Cem', bytes)];
    net.drain(random);
    for (let step = 0; step < 30; step++) {
      const member = pick(random, members);
      const roll = random();
      if (roll < 0.15) member.session.undo();
      else if (roll < 0.2) member.session.redo();
      else {
        const sheet = int(random, 0, 1);
        const axis = random() < 0.6 ? 'row' : 'col';
        const op = random() < 0.6 ? ins(sheet, axis, int(random, 0, 9), int(random, 1, 2)) : del(sheet, axis, int(random, 0, 9), 1);
        act(member, () => [op]);
      }
      // Some messages go now, the others later: edits cross on their way.
      for (let k = int(random, 0, 4); k > 0; k--) net.step(random);
    }
    net.drain(random);
    assertSame(members, `seed ${seed}`);
    const places = members.map((member) => placesOf(member.wb));
    for (const one of places) assert.equal(one, places[0], `seed ${seed}`);
    if (places[0] !== placesOf((await load(bytes)).wb)) moved++;
    // What each saves names the same ranges.
    const saved = await Promise.all(members.map(saveOf));
    for (const part of [CACHE, TABLE, 'xl/pivotTables/pivotTable2.xml']) for (const one of saved) assert.equal(await text(one, part), await text(saved[0], part), `seed ${seed} ${part}`);
  }
  // The tables and sources did move in most runs: the test is about something.
  assert.ok(moved >= 8, `moved in ${moved} runs`);
});

// ---------------------------------------------------------------------------------------------
// Columns inserted inside a source are no fields of its cache

/** The columns of the source of the first cache that are fields now, as the header row says (fieldColumnsNow()). */
function fieldsOf(wb: Loaded['wb']): number[] | 'all' {
  const cache = wb.pivots!.caches[0];
  const source = cacheSourceNow(cache)!;
  const at = source.kind === 'range' ? { sheetName: source.sheet, ref: source.ref } : (namedSourceRange(wb, source) as { sheetName: string; ref: string });
  const sheet = wb.sheets.find((one) => one.name === at.sheetName)!;
  return fieldColumnsNow(cache, sheet, sourceArea(at.ref)!);
}

test('a column inserted inside the source is no field: it may be deleted again, a column of a field may not', async () => {
  const bytes = await pivotBook(SPEC);
  const doc = await load(bytes);
  const cache = doc.wb.pivots!.caches[0];
  apply(doc, ins(0, 'col', 1));
  assert.equal(cache.sourceRef, 'A1:D6');
  assert.deepEqual(fieldsOf(doc.wb), [0, 2, 3]);
  // The columns of Region, Product and Sales stay.
  for (const col of [0, 2, 3]) assert.throws(() => applyOp(doc.wb, del(0, 'col', col)), /a column of it would go/);
  assert.throws(() => applyOp(doc.wb, del(0, 'col', 1, 2)), /a column of it would go/);
  // The new one goes: the source is what it was.
  apply(doc, del(0, 'col', 1));
  assert.equal(cache.sourceRef, undefined);
  assert.deepEqual(fieldsOf(doc.wb), [0, 1, 2]);
  assert.ok(same(await save(doc), bytes));
  // Undone and redone step by step, each step puts back its own.
  doc.history.undo();
  assert.deepEqual(fieldsOf(doc.wb), [0, 2, 3]);
  apply(doc, ins(0, 'col', 3, 2));
  assert.equal(cache.sourceRef, 'A1:F6');
  assert.deepEqual(fieldsOf(doc.wb), [0, 2, 5]);
  apply(doc, ins(0, 'col', 0));
  assert.deepEqual(fieldsOf(doc.wb), [1, 3, 6]);
  assert.equal(cache.sourceRef, 'B1:G6');
  doc.history.undo();
  doc.history.undo();
  assert.deepEqual(fieldsOf(doc.wb), [0, 2, 3]);
  doc.history.undo();
  assert.deepEqual(fieldsOf(doc.wb), [0, 1, 2]);
  assert.ok(same(await save(doc), bytes));
  doc.history.redo();
  doc.history.redo();
  assert.deepEqual(fieldsOf(doc.wb), [0, 2, 5]);
});

test('a source given by a defined name knows its fields the same way', async () => {
  const bytes = await pivotBook({ ...SPEC, sourceName: 'Sales' });
  const doc = await load(bytes);
  const cache = doc.wb.pivots!.caches[0];
  apply(doc, ins(0, 'col', 2));
  assert.equal(doc.wb.definedNames.find((name) => name.name === 'Sales')!.value, 'Data!$A$1:$D$6');
  assert.equal(cache.sourceRef, undefined);
  assert.deepEqual(fieldsOf(doc.wb), [0, 1, 3]);
  assert.throws(() => applyOp(doc.wb, del(0, 'col', 3)), /a column of it would go/);
  apply(doc, del(0, 'col', 2));
  assert.deepEqual(fieldsOf(doc.wb), [0, 1, 2]);
  assert.equal(doc.wb.definedNames.find((name) => name.name === 'Sales')!.value, 'Data!$A$1:$C$6');
  assert.equal(await text(await save(doc), CACHE), await text(bytes, CACHE));
});

test('a source with a column that is no field is not refreshed, and says why', async () => {
  const doc = await load(await pivotBook(SPEC));
  apply(doc, ins(0, 'col', 1));
  assert.throws(() => refreshPivot(doc.wb, 0), /^PivotRefreshError: Cannot refresh the pivot table "PivotTable1": its source has 4 columns and its cache 3 fields: columns were inserted into it, and a refresh here adds no fields$/);
  try {
    refreshPivot(doc.wb, 0);
  } catch (error) {
    assert.match(STRINGS_TR.pivots.refused((error as Error).message), /kaynağa sütun eklenmiş, buradaki yenileme önbelleğe yeni alan eklemez/);
  }
});

test('saved and opened again, the fields of a source are found by its headers', async () => {
  for (const header of [undefined, 'New']) {
    const doc = await load(await pivotBook(SPEC));
    apply(doc, ins(0, 'col', 1));
    if (header) setData(doc, 0, 1, header);
    const reread = await load(await save(doc));
    const cache = reread.wb.pivots!.caches[0];
    assert.deepEqual(cache.source, { kind: 'range', sheet: 'Data', ref: 'A1:D6' });
    // Region, Product and Sales stand in A, C and D: those stay, the inserted column B may go.
    assert.deepEqual(fieldsOf(reread.wb), [0, 2, 3], String(header));
    for (const col of [0, 2, 3]) assert.throws(() => applyOp(reread.wb, del(0, 'col', col)), /a column of it would go/);
    apply(reread, del(0, 'col', 1));
    assert.equal(cache.sourceRef, 'A1:C6');
    assert.deepEqual(fieldsOf(reread.wb), [0, 1, 2]);
  }
});

test('a file whose headers do not match the fields one to one: every column of the source is a field', async () => {
  const bytes = await pivotBook(SPEC);
  const wider = async (sheet: (xml: string) => string): Promise<Loaded> => {
    const zip = ZipArchive.read(bytes);
    zip.set(CACHE, new TextEncoder().encode((await text(bytes, CACHE)).replace('ref="A1:C6"', 'ref="A1:D6"')));
    zip.set('xl/worksheets/sheet1.xml', new TextEncoder().encode(sheet(await text(bytes, 'xl/worksheets/sheet1.xml'))));
    return load(await zip.write());
  };
  // Headers that match: the column beyond the fields is none, and may go.
  const plain = await wider((xml) => xml);
  assert.deepEqual(fieldsOf(plain.wb), [0, 1, 2]);
  apply(plain, del(0, 'col', 3));
  assert.equal(plain.wb.pivots!.caches[0].sourceRef, 'A1:C6');
  // A header that is not a field's name, one that stands twice, or names that do not come in the order of the fields.
  const cases: ((xml: string) => string)[] = [
    (xml) => xml.replace('<t>Sales</t>', '<t>Amount</t>'),
    (xml) => xml.replace('<t>Region</t>', '<t>Product</t>'),
    (xml) => xml.replace('<t>Region</t>', '<t>@</t>').replace('<t>Product</t>', '<t>Region</t>').replace('<t>@</t>', '<t>Product</t>'),
  ];
  for (const change of cases) {
    const doc = await wider(change);
    assert.deepEqual(fieldsOf(doc.wb), 'all');
    for (let col = 0; col < 4; col++) assert.throws(() => applyOp(doc.wb, del(0, 'col', col)), /a column of it would go/);
    // Still every column after an insertion inside it, the inserted one too: nothing tells it from a field.
    apply(doc, ins(0, 'col', 2));
    assert.deepEqual(fieldsOf(doc.wb), 'all');
    assert.throws(() => applyOp(doc.wb, del(0, 'col', 2)), /a column of it would go/);
  }
  // A name that two headers have, even where either would keep the order of the fields.
  const twice = await load(bytes);
  apply(twice, ins(0, 'col', 1));
  setData(twice, 0, 1, 'product');
  const reread = await load(await save(twice));
  assert.deepEqual(fieldsOf(reread.wb), 'all');
  // A header beside the source, outside it, is none of its headers.
  const beside = await load(bytes);
  setData(beside, 0, 4, 'Region');
  assert.deepEqual(fieldsOf(beside.wb), [0, 1, 2]);
  // Headers matched without regard to case.
  const cased = await wider((xml) => xml.replace('<t>Sales</t>', '<t>SALES</t>'));
  assert.deepEqual(fieldsOf(cased.wb), [0, 1, 2]);
});

test('editing together: a column inserted inside a source is undone and redone like any other step', async () => {
  const bytes = await pivotBook(SPEC);
  const net = new Net();
  const random = seeded(1);
  const ada = await join(net, 'Ada', bytes);
  const bob = await join(net, 'Bob', bytes);
  net.drain(random);
  assert.ok(act(ada, () => [ins(0, 'col', 1)]));
  net.drain(random);
  for (const member of [ada, bob]) assert.deepEqual(fieldsOf(member.wb), [0, 2, 3], member.name);
  assert.ok(ada.session.undo());
  net.drain(random);
  assert.deepEqual(ada.notices, []);
  for (const member of [ada, bob]) {
    assert.equal(member.wb.pivots!.caches[0].sourceRef, undefined, member.name);
    assert.deepEqual(fieldsOf(member.wb), [0, 1, 2], member.name);
    assert.equal(member.wb.sheets[0].colCount, 3, member.name);
  }
  assert.ok(ada.session.redo());
  net.drain(random);
  assert.deepEqual(ada.notices, []);
  for (const member of [ada, bob]) {
    assert.equal(member.wb.pivots!.caches[0].sourceRef, 'A1:D6', member.name);
    assert.deepEqual(fieldsOf(member.wb), [0, 2, 3], member.name);
  }
  assertSame([ada, bob], 'after the redo');
});

test('an inserted column whose header is turned into the name of a field: the headers decide, and two of a name make every column a field', async () => {
  const doc = await load(await pivotBook(SPEC));
  // Region, Product, (new), Sales.
  apply(doc, ins(0, 'col', 2));
  assert.deepEqual(fieldsOf(doc.wb), [0, 1, 3]);
  setData(doc, 0, 2, 'sales');
  assert.equal(fieldsOf(doc.wb), 'all');
  for (let col = 0; col < 4; col++) assert.throws(() => applyOp(doc.wb, del(0, 'col', col)), /a column of it would go/);
  // The header of the old column of Sales renamed: the new column is the field now, the old one is none.
  setData(doc, 0, 3, 'Old sales');
  assert.deepEqual(fieldsOf(doc.wb), [0, 1, 2]);
  assert.throws(() => applyOp(doc.wb, del(0, 'col', 2)), /a column of it would go/);
  apply(doc, del(0, 'col', 3));
  assert.equal(doc.wb.pivots!.caches[0].sourceRef, undefined);
});

test('editing together: a header changed and saved, a participant who joins from that file, a deletion: every one decides alike', async () => {
  const bytes = await pivotBook(SPEC);
  const net = new Net();
  const random = seeded(3);
  const ada = await join(net, 'Ada', bytes);
  const bob = await join(net, 'Bob', bytes);
  net.drain(random);
  const header = (text: string): Op => ({ type: 'setCells', sheet: 0, cells: [{ row: 0, col: 2, cell: { kind: 'string', value: text, style: 0 } }] });
  assert.ok(act(ada, () => [ins(0, 'col', 1)]));
  // The header of Product becomes something else: the headers no longer match the fields one to one.
  assert.ok(act(ada, () => [header('Item')]));
  net.drain(random);
  const saved = await saveOf(ada);
  // The file is stored and the room told so, as mount() does: who joins from it now joins the room.
  ada.session.announceSaved([contentVersion(saved)], ada.session.savableRev);
  net.drain(random);
  const cem = await join(net, 'Cem', saved);
  net.drain(random);
  const members = [ada, bob, cem];
  for (const member of members) assert.equal(fieldsOf(member.wb), 'all', member.name);
  // The inserted column may not go, for any of them.
  for (const member of members) assert.equal(act(member, () => [del(0, 'col', 1)]), false, member.name);
  net.drain(random);
  assertSame(members, 'after the refused deletion');
  // The header back: the inserted column is none of the fields again, for every one.
  assert.ok(act(cem, () => [header('Product')]));
  net.drain(random);
  for (const member of members) assert.deepEqual(fieldsOf(member.wb), [0, 2, 3], member.name);
  assert.ok(act(bob, () => [del(0, 'col', 1)]));
  net.drain(random);
  assertSame(members, 'after the deletion');
  for (const member of members) {
    assert.equal((cacheSourceNow(member.wb.pivots!.caches[0]) as { ref: string }).ref, 'A1:C6', member.name);
    assert.equal(member.wb.sheets[0].colCount, 3, member.name);
    assert.deepEqual(member.notices, [], member.name);
  }
});

// ---------------------------------------------------------------------------------------------
// Ends that the tests above do not reach

test('rows inserted into the source and refreshed at once, without a save: the refresh reads the range as it now stands', async () => {
  const doc = await load(await pivotBook(SPEC));
  apply(doc, ins(0, 'row', 0, 2));
  apply(doc, ins(0, 'row', 4));
  setData(doc, 4, 0, 'North');
  setData(doc, 4, 1, 'A');
  setData(doc, 4, 2, 100);
  doc.history.apply(refreshPivot(doc.wb, 0), undefined, undefined);
  const ref = tableState(doc.wb.pivots!.tables[0]).location.ref;
  assert.deepEqual(grid(doc.wb, 1, ref.top, ref.left, ref.bottom, ref.right), [['Row Labels', 'Sum of Sales'], ['East', 16], ['West', 27], ['North', 100], ['Grand Total', 143]]);
  // The cache refreshed from its tree carries the moved range as well.
  const saved = await save(doc);
  assert.match(await sourceOf(saved), /ref="A3:C9"/);
  if (hasPython()) assert.deepEqual(independentPivots({ saved }), []);
});

test('a change in place that the tree would read otherwise is not kept', () => {
  // Only the second worksheetSource has a range: the text has it once, the reader takes the first element.
  const xml = '<pivotCacheDefinition><cacheSource><worksheetSource sheet="D"/></cacheSource><cacheSource><worksheetSource ref="A1:C6"/></cacheSource></pivotCacheDefinition>';
  assert.throws(() => patchedAttribute(new TextEncoder().encode(xml), ['pivotCacheDefinition', 'cacheSource', 'worksheetSource'], 'ref', 'A2:C7'), /has no pivotCacheDefinition\/cacheSource\/worksheetSource with ref/);
});

test('a cache part that no table is made from keeps rows and columns where they are', async () => {
  const bytes = await pivotBook(SPEC);
  const zip = ZipArchive.read(bytes);
  zip.set('xl/pivotCache/pivotCacheDefinition5.xml', await zip.read(CACHE));
  const doc = await load(await zip.write());
  assert.equal(doc.wb.pivots!.fixed, 'the package has a pivot cache no table is made from');
  assert.throws(() => applyOp(doc.wb, ins(0, 'row', 0)), /not supported yet/);
});

test('a workbook with more pivot tables than are read keeps rows and columns where they are', async () => {
  const bytes = await pivotBook(SPEC);
  const zip = ZipArchive.read(bytes);
  const table = await zip.read(TABLE);
  const tableRels = await zip.read('xl/pivotTables/_rels/pivotTable1.xml.rels');
  const rels: string[] = [];
  for (let i = 1; i <= 1001; i++) {
    if (i > 1) {
      zip.set(`xl/pivotTables/pivotTable${i}.xml`, table);
      zip.set(`xl/pivotTables/_rels/pivotTable${i}.xml.rels`, tableRels);
    }
    rels.push(`<Relationship Id="rId${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable${i}.xml"/>`);
  }
  zip.set('xl/worksheets/_rels/sheet2.xml.rels', new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`));
  const doc = await load(await zip.write());
  assert.equal(doc.wb.pivots!.tables.length, 1000);
  assert.equal(doc.wb.pivots!.fixed, 'the workbook has more pivot tables than are read');
});

test('fields of the page in more rows than there are fields: the rows the file states count', async () => {
  const bytes = await pivotBook({ data: DATA, rows: [1], values: [{ fld: 2 }], pages: [{ fld: 0, item: 'East' }], at: { row: 6, col: 0 } });
  const zip = ZipArchive.read(bytes);
  zip.set(TABLE, new TextEncoder().encode((await text(bytes, TABLE)).replace('rowPageCount="1"', 'rowPageCount="3"')));
  const doc = await load(await zip.write());
  assert.deepEqual(doc.wb.pivots!.tables[0].pages, { rows: 3, cols: 1 });
  // The body starts in row 7: the fields of the page take rows 3 to 5, row 6 is blank.
  assert.throws(() => applyOp(doc.wb, ins(1, 'row', 3)), /through the pivot table/);
  apply(doc, ins(1, 'row', 2));
});

test('a table that would end past the last row of the sheet is not moved there', async () => {
  const bytes = await pivotBook(SPEC);
  const zip = ZipArchive.read(bytes);
  // The file says the table reaches the last row (its cells stand higher up).
  zip.set(TABLE, new TextEncoder().encode((await text(bytes, TABLE)).replace('ref="A3:B4"', 'ref="A1048575:B1048576"')));
  const doc = await load(await zip.write());
  assert.throws(() => applyOp(doc.wb, ins(1, 'row', 0)), /^RangeError: Cannot insert rows: the pivot table "PivotTable1" would be pushed off the sheet$/);
  // One row higher it fits exactly.
  const fits = ZipArchive.read(bytes);
  fits.set(TABLE, new TextEncoder().encode((await text(bytes, TABLE)).replace('ref="A3:B4"', 'ref="A1048574:B1048575"')));
  const other = await load(await fits.write());
  apply(other, ins(1, 'row', 0));
  assert.deepEqual(other.wb.pivots!.tables[0].placed, { top: 1048574, left: 0, bottom: 1048575, right: 1 });
});
