// Pivot tables: reading them, refreshing them from their source, writing what a refresh made,
// and what is refused. The workbooks are made by test/pivot-fixtures.ts.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { ZipArchive } from '../src/index';
import { encodeOp } from '../src/collab/wire';
import { CollabUnsupportedError } from '../src/collab/types';
import { applyOp } from '../src/ops/ops';
import { listPivots, pivotAt, refreshAllPivots, refreshPivot, refreshPivotProblem } from '../src/ops/pivots';
import { tableState } from '../src/model/sheet/pivot';
import { join, Net } from './collab-fixtures';
import { grid, hasPython, independentPivots, load, pivotBook, same, save, setData, text, type Datum, type PivotSpec } from './pivot-fixtures';

const DATA: Datum[][] = [
  ['Region', 'Product', 'Sales'],
  ['East', 'A', 10],
  ['West', 'B', 20],
  ['East', 'B', 5],
  ['West', 'A', 7],
  ['East', 'A', 1],
];

/** Loads a workbook made from the spec, refreshes its first table and gives the body of the table as values. */
async function refreshed(spec: PivotSpec, change?: (doc: Awaited<ReturnType<typeof load>>) => void) {
  const doc = await load(await pivotBook(spec));
  if (change) change(doc);
  doc.history.apply(refreshPivot(doc.wb, 0), undefined, undefined);
  const ref = tableState(doc.wb.pivots!.tables[0]).location.ref;
  return { doc, ref, body: grid(doc.wb, 1, ref.top, ref.left, ref.bottom, ref.right) };
}

test('the reader takes in the table, its fields and its cache; nothing of the sheet changes', async () => {
  const doc = await load(await pivotBook({ data: DATA, rows: [0], cols: [1], values: [{ fld: 2 }] }));
  const pivots = doc.wb.pivots!;
  assert.equal(pivots.tables.length, 1);
  assert.equal(pivots.caches.length, 1);
  const table = pivots.tables[0];
  assert.equal(table.name, 'PivotTable1');
  assert.equal(table.sheet, doc.wb.sheets[1]);
  assert.deepEqual(table.rowFields, [0]);
  assert.deepEqual(table.colFields, [1]);
  assert.deepEqual(table.dataFields, [{ name: 'Sum of Sales', fld: 2, subtotal: 'sum' }]);
  assert.deepEqual(table.location.ref, { top: 2, left: 0, bottom: 4, right: 1 });
  assert.equal(table.refusal, undefined);
  const cache = pivots.caches[0];
  assert.equal(cache.cacheId, 7);
  assert.deepEqual(cache.source, { kind: 'range', sheet: 'Data', ref: 'A1:C6' });
  assert.deepEqual(cache.fields.map((f) => f.name), ['Region', 'Product', 'Sales']);
  assert.deepEqual(cache.fields[0].items, [{ t: 's', v: 'East' }, { t: 's', v: 'West' }]);
  assert.equal(cache.fields[2].items, undefined);
  assert.equal(cache.recordsPart, '/xl/pivotCache/pivotCacheRecords1.xml');
  assert.deepEqual(listPivots(doc.wb), [{ index: 0, name: 'PivotTable1', sheet: 1, range: { top: 2, left: 0, bottom: 4, right: 1 } }]);
  assert.equal(pivotAt(doc.wb, 1, 3, 1), 0);
  assert.equal(pivotAt(doc.wb, 1, 5, 1), undefined);
  assert.equal(pivotAt(doc.wb, 0, 3, 1), undefined);
});

test('open and save without a change gives the file back byte for byte', async () => {
  const bytes = await pivotBook({ data: DATA, rows: [0, 1], values: [{ fld: 2 }] });
  const doc = await load(bytes);
  assert.ok(same(await save(doc), bytes));
});

test('compact form, one field in rows', async () => {
  const { body } = await refreshed({ data: DATA, rows: [0], values: [{ fld: 2 }] });
  assert.deepEqual(body, [
    ['Row Labels', 'Sum of Sales'],
    ['East', 16],
    ['West', 27],
    ['Grand Total', 43],
  ]);
});

test('compact form, two fields in rows: the subtotal stands on the line of the item above', async () => {
  const { body, doc } = await refreshed({ data: DATA, rows: [0, 1], values: [{ fld: 2 }] });
  assert.deepEqual(body, [
    ['Row Labels', 'Sum of Sales'],
    ['East', 16],
    ['A', 11],
    ['B', 5],
    ['West', 27],
    ['A', 7],
    ['B', 20],
    ['Grand Total', 43],
  ]);
  const state = doc.wb.pivots!.tables[0].now!;
  assert.deepEqual(state.rowItems, [
    { r: 0, i: 0, x: [0] },
    { r: 1, i: 0, x: [0] },
    { r: 1, i: 0, x: [1] },
    { r: 0, i: 0, x: [1] },
    { r: 1, i: 0, x: [0] },
    { r: 1, i: 0, x: [1] },
    { t: 'grand', r: 0, i: 0, x: [0] },
  ]);
});

test('subtotals at the bottom, and none at all', async () => {
  const bottom = await refreshed({ data: DATA, rows: [0, 1], values: [{ fld: 2 }], subtotalTop: false });
  assert.deepEqual(bottom.body.slice(1, 5), [
    ['East', undefined],
    ['A', 11],
    ['B', 5],
    ['East Total', 16],
  ]);
  const none = await refreshed({ data: DATA, rows: [0, 1], values: [{ fld: 2 }], defaultSubtotal: false });
  assert.deepEqual(none.body.slice(1, 4), [
    ['East', undefined],
    ['A', 11],
    ['B', 5],
  ]);
  assert.equal(none.body.length, 8);
});

test('outline and tabular form', async () => {
  const outline = await refreshed({ data: DATA, rows: [0, 1], values: [{ fld: 2 }], form: 'outline' });
  assert.deepEqual(outline.body.slice(0, 4), [
    ['Region', 'Product', 'Sum of Sales'],
    ['East', undefined, 16],
    [undefined, 'A', 11],
    [undefined, 'B', 5],
  ]);
  const tabular = await refreshed({ data: DATA, rows: [0, 1], values: [{ fld: 2 }], form: 'tabular' });
  assert.deepEqual(tabular.body, [
    ['Region', 'Product', 'Sum of Sales'],
    ['East', 'A', 11],
    [undefined, 'B', 5],
    ['East Total', undefined, 16],
    ['West', 'A', 7],
    [undefined, 'B', 20],
    ['West Total', undefined, 27],
    ['Grand Total', undefined, 43],
  ]);
});

test('a field in columns, and several values', async () => {
  const cols = await refreshed({ data: DATA, rows: [0], cols: [1], values: [{ fld: 2 }] });
  assert.deepEqual(cols.body, [
    ['Sum of Sales', 'Column Labels', undefined, undefined],
    ['Row Labels', 'A', 'B', 'Grand Total'],
    ['East', 11, 5, 16],
    ['West', 7, 20, 27],
    ['Grand Total', 18, 25, 43],
  ]);
  const values = await refreshed({ data: DATA, rows: [0], values: [{ fld: 2 }, { fld: 2, subtotal: 'average' }, { fld: 1, subtotal: 'count' }] });
  assert.deepEqual(values.body, [
    ['Row Labels', 'Sum of Sales', 'Average of Sales', 'Count of Product'],
    ['East', 16, 16 / 3, 3],
    ['West', 27, 13.5, 2],
    ['Grand Total', 43, 43 / 5, 5],
  ]);
  assert.equal(values.doc.wb.pivots!.tables[0].now!.location.firstHeaderRow, 0);
});

test('two fields in columns: the subtotal of the outer one follows its items', async () => {
  const data: Datum[][] = [['R', 'C1', 'C2', 'V'], ['a', 'x', 'p', 1], ['a', 'x', 'q', 2], ['a', 'y', 'p', 4], ['b', 'y', 'q', 8]];
  const { body } = await refreshed({ data, rows: [0], cols: [1, 2], values: [{ fld: 3 }] });
  assert.deepEqual(body, [
    ['Sum of V', 'Column Labels', undefined, undefined, undefined, undefined, undefined, undefined],
    [undefined, 'x', undefined, 'x Total', 'y', undefined, 'y Total', 'Grand Total'],
    ['Row Labels', 'p', 'q', undefined, 'p', 'q', undefined, undefined],
    ['a', 1, 2, 3, 4, undefined, 4, 7],
    ['b', undefined, undefined, undefined, undefined, 8, 8, 8],
    ['Grand Total', 1, 2, 3, 4, 8, 12, 15],
  ]);
});

test('every summary function', async () => {
  const data: Datum[][] = [['K', 'V'], ['a', 2], ['a', 4], ['a', 'text'], ['a', null], ['b', 3]];
  const fns = ['sum', 'count', 'average', 'max', 'min', 'product', 'countNums', 'stdDev', 'stdDevp', 'var', 'varp'];
  const { body } = await refreshed({ data, rows: [0], values: fns.map((subtotal) => ({ fld: 1, subtotal })) });
  const a = body[1].slice(1) as number[];
  const b = body[2].slice(1);
  const expected = [6, 3, 3, 4, 2, 8, 2, Math.sqrt(2), 1, 2, 1];
  a.forEach((value, i) => assert.ok(Math.abs(value - expected[i]) < 1e-12, `${fns[i]}: ${value}`));
  assert.deepEqual(b, [3, 1, 3, 3, 3, 3, 1, '#DIV/0!', 0, '#DIV/0!', 0]);
});

test('an error in the values is the value of its total', async () => {
  const data: Datum[][] = [['K', 'V'], ['a', 2], ['a', '#N/A'], ['b', 3]];
  const { body, doc } = await refreshed({ data, rows: [0], values: [{ fld: 1 }] });
  assert.deepEqual(body.slice(1), [['a', '#N/A'], ['b', 3], ['Grand Total', '#N/A']]);
  assert.equal(doc.wb.sheets[1].rows.get(3)!.cells[1].kind, 'error');
});

test('a refresh takes in what the source has now: new items at the end, gone items out, hidden items stay hidden', async () => {
  const spec: PivotSpec = { data: DATA, rows: [0], values: [{ fld: 2 }], hidden: [{ fld: 0, values: ['West'] }] };
  const { body, doc } = await refreshed(spec, (doc) => {
    setData(doc, 1, 0, 'North');
    setData(doc, 3, 0, 'Central');
    setData(doc, 5, 0, 'Central');
  });
  // East is gone, West stays hidden, North and Central come in the order they come in the source.
  assert.deepEqual(body, [
    ['Row Labels', 'Sum of Sales'],
    ['North', 10],
    ['Central', 6],
    ['Grand Total', 16],
  ]);
  const items = doc.wb.pivots!.tables[0].now!.items[0]!;
  const values = doc.wb.pivots!.caches[0].now!.fields[0].items!;
  assert.deepEqual(
    items.map((item) => (item.t ? item.t : `${(values[item.x!] as { v: string }).v}${item.h ? '-' : ''}`)),
    ['West-', 'North', 'Central', 'default'],
  );
});

test('a field sorted ascending is sorted again', async () => {
  const { body } = await refreshed({ data: DATA, rows: [0], values: [{ fld: 2 }], sortType: [{ fld: 0, type: 'descending' }] }, (doc) => setData(doc, 2, 0, 'Zeta'));
  assert.deepEqual(body.map((row) => row[0]), ['Row Labels', 'Zeta', 'West', 'East', 'Grand Total']);
});

test('the page shows one item, or every item that is not hidden', async () => {
  const one = await refreshed({ data: DATA, rows: [1], values: [{ fld: 2 }], pages: [{ fld: 0, item: 'West' }], at: { row: 4, col: 0 } });
  assert.deepEqual(one.body.slice(1), [['A', 7], ['B', 20], ['Grand Total', 27]]);
  const hidden = await refreshed({ data: DATA, rows: [1], values: [{ fld: 2 }], pages: [{ fld: 0 }], hidden: [{ fld: 0, values: ['East'] }], at: { row: 4, col: 0 } });
  assert.deepEqual(hidden.body.slice(1), [['A', 7], ['B', 20], ['Grand Total', 27]]);
  // The page field shows its item at its new place.
  const moved = await load(await pivotBook({ data: DATA, rows: [1], values: [{ fld: 2 }], pages: [{ fld: 0, item: 'West' }], at: { row: 4, col: 0 } }));
  setData(moved, 1, 0, 'Aardvark');
  moved.history.apply(refreshPivot(moved.wb, 0), undefined, undefined);
  const state = moved.wb.pivots!.tables[0].now!;
  const items = state.items[0]!;
  const values = moved.wb.pivots!.caches[0].now!.fields[0].items!;
  assert.deepEqual(values[items[state.pageItems[0]!].x!], { t: 's', v: 'West' });
  // The item the page shows is gone from the source: refused.
  const gone = await load(await pivotBook({ data: DATA, rows: [1], values: [{ fld: 2 }], pages: [{ fld: 0, item: 'West' }], at: { row: 4, col: 0 } }));
  setData(gone, 2, 0, 'East');
  setData(gone, 4, 0, 'East');
  assert.throws(() => refreshPivot(gone.wb, 0), /not in its source any more/);
});

test('dates in rows: the cache holds them as dates, the labels are numbers with a date format', async () => {
  const data: Datum[][] = [['Day', 'V'], [{ date: 45292 }, 1], [{ date: 45293 }, 2], [{ date: 45292 }, 3]];
  const { body, doc } = await refreshed({ data, rows: [0], values: [{ fld: 1 }] });
  assert.deepEqual(body.slice(1), [[45292, 4], [45293, 2], ['Grand Total', 6]]);
  const label = doc.wb.sheets[1].rows.get(3)!.cells[0];
  assert.equal(doc.wb.styles.cellXfs[label.style].numFmtId, 14);
  const bytes = await save(doc);
  const cache = await text(bytes, 'xl/pivotCache/pivotCacheDefinition1.xml');
  assert.match(cache, /<cacheField name="Day" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsNonDate="0" containsDate="1" containsString="0" minDate="2024-01-01T00:00:00" maxDate="2024-01-02T00:00:00" count="2"><d v="2024-01-01T00:00:00"\/><d v="2024-01-02T00:00:00"\/><\/sharedItems>/);
});

test('the words of the file are kept: captions, the word of the grand total and of subtotals', async () => {
  const words = { rowHeader: 'Satır Etiketleri', grand: 'Genel Toplam' };
  const first = await refreshed({ data: DATA, rows: [0, 1], values: [{ fld: 2 }], form: 'tabular', words });
  assert.equal(first.body[first.body.length - 1][0], 'Genel Toplam');
  // A subtotal written as Excel in Turkish writes it: the next refresh says the others that way.
  const doc = first.doc;
  doc.history.apply([{ type: 'setCells', sheet: 1, cells: [{ row: 5, col: 0, cell: { kind: 'string', value: 'East Toplam', style: 0 } }] }], undefined, undefined);
  setData(doc, 2, 0, 'North');
  doc.history.apply(refreshPivot(doc.wb, 0), undefined, undefined);
  const ref = doc.wb.pivots!.tables[0].now!.location.ref;
  const labels = grid(doc.wb, 1, ref.top, ref.left, ref.bottom, ref.left).map((row) => row[0]);
  assert.ok(labels.indexOf('North Toplam') >= 0, JSON.stringify(labels));
  assert.ok(labels.indexOf('East Toplam') >= 0, JSON.stringify(labels));
  const compact = await refreshed({ data: DATA, rows: [0], values: [{ fld: 2 }], words });
  assert.equal(compact.body[0][0], 'Satır Etiketleri');
});

test('the styles of the cells of the table as it was are those of the table as it is', async () => {
  const doc = await load(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] }));
  // The stand-in's grand total in B4: style 1 (#,##0); its label: style 3 (bold).
  doc.history.apply(
    [
      {
        type: 'setCells',
        sheet: 1,
        cells: [
          { row: 3, col: 1, cell: { kind: 'number', value: 0, style: 1 } },
          { row: 3, col: 0, cell: { kind: 'string', value: 'Grand Total', style: 3 } },
        ],
      },
    ],
    undefined,
    undefined,
  );
  doc.history.apply(refreshPivot(doc.wb, 0), undefined, undefined);
  const rows = doc.wb.sheets[1].rows;
  assert.equal(rows.get(5)!.cells[0].style, 3);
  assert.equal(rows.get(5)!.cells[1].style, 1);
  // Values of the lines of items had no cell to go by: they take the style of a value of the same field.
  assert.equal(rows.get(3)!.cells[1].style, 1);
  assert.equal(rows.get(3)!.cells[0].style, 0);
});

test('refresh, undo, save: the file as it was, byte for byte', async () => {
  const bytes = await pivotBook({ data: DATA, rows: [0, 1], values: [{ fld: 2 }], cols: [] });
  const doc = await load(bytes);
  setData(doc, 1, 2, 99);
  doc.history.apply(refreshPivot(doc.wb, 0), undefined, undefined);
  await save(doc);
  doc.history.undo();
  doc.history.undo();
  const back = await save(doc);
  const a = ZipArchive.read(bytes);
  const b = ZipArchive.read(back);
  for (const name of a.names()) assert.ok(same(await b.read(name), await a.read(name)), name);
  assert.deepEqual(b.names().sort(), a.names().sort());
});

test('a refreshed file is read back as it was written, and a second refresh changes no cell', async () => {
  const { doc } = await refreshed({ data: DATA, rows: [0, 1], values: [{ fld: 2 }, { fld: 2, subtotal: 'max' }], form: 'tabular' });
  const again = await load(await save(doc));
  const table = again.wb.pivots!.tables[0];
  assert.deepEqual(table.rowItems, doc.wb.pivots!.tables[0].now!.rowItems);
  assert.deepEqual(table.location, doc.wb.pivots!.tables[0].now!.location);
  const ops = refreshPivot(again.wb, 0);
  assert.deepEqual(ops.map((op) => op.type), ['setPivot']);
  // A refresh that changes no cell still writes its parts; they come out as they were written the first time.
  const first = await save(doc);
  again.history.apply(ops, undefined, undefined);
  const second = await save(again);
  for (const name of ['xl/pivotTables/pivotTable1.xml', 'xl/pivotCache/pivotCacheDefinition1.xml', 'xl/pivotCache/pivotCacheRecords1.xml']) {
    assert.equal(await text(second, name), await text(first, name), name);
  }
  if (hasPython()) assert.deepEqual(independentPivots({ second }), []);
});

test('what the file has beside what a refresh changes stays: unknown elements and attributes, extensions of the records', async () => {
  const ext = '<extLst><ext uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:pivotTableDefinition hideValuesRow="1"/></ext></extLst>';
  const spec: PivotSpec = {
    data: DATA,
    rows: [0],
    values: [{ fld: 2 }],
    tableTail: ext,
    // An attribute of a namespace the part says may be passed over: kept, and no reason to refuse.
    tableAttrs: ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:z="urn:z" mc:Ignorable="z" z:futureThing="7"',
    cacheTail: '<extLst><ext uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:pivotCacheDefinition pivotCacheId="12"/></ext></extLst>',
    recordsTail: '<extLst><ext uri="{0000}"><keep/></ext></extLst>',
    fieldTail: { 0: '<extLst><ext uri="{1111}"><keep-field/></ext></extLst>' },
  };
  const { doc } = await refreshed(spec, (doc) => setData(doc, 1, 0, 'North'));
  const bytes = await save(doc);
  const table = await text(bytes, 'xl/pivotTables/pivotTable1.xml');
  assert.ok(table.indexOf(ext) >= 0);
  assert.match(table, /z:futureThing="7"/);
  const cache = await text(bytes, 'xl/pivotCache/pivotCacheDefinition1.xml');
  assert.match(cache, /pivotCacheId="12"/);
  assert.match(cache, /<keep-field\/>/);
  assert.match(cache, /refreshedDate="45000.5"/);
  const records = await text(bytes, 'xl/pivotCache/pivotCacheRecords1.xml');
  assert.match(records, /<extLst><ext uri="\{0000\}"><keep\/><\/ext><\/extLst><\/pivotCacheRecords>$/);
  assert.match(records, /count="5"/);
});

test('a table that would grow over cells that hold something is not refreshed; a smaller one clears what it left', async () => {
  const blocked = await load(await pivotBook({ data: DATA, rows: [0, 1], values: [{ fld: 2 }], around: [[7, 0, 'mine']] }));
  assert.throws(() => refreshPivot(blocked.wb, 0), /would grow over A8, which is not empty/);
  assert.match(refreshPivotProblem(blocked.wb, 0) || '', /A8/);
  const doc = (await refreshed({ data: DATA, rows: [0, 1], values: [{ fld: 2 }] })).doc;
  setData(doc, 2, 0, 'East');
  setData(doc, 4, 0, 'East');
  doc.history.apply(refreshPivot(doc.wb, 0), undefined, undefined);
  const ref = doc.wb.pivots!.tables[0].now!.location.ref;
  assert.deepEqual(ref, { top: 2, left: 0, bottom: 6, right: 1 });
  for (let r = 7; r <= 9; r++) assert.equal(doc.wb.sheets[1].rows.get(r), undefined);
});

test('merged cells and another table in the way', async () => {
  const doc = await load(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] }));
  doc.history.apply([{ type: 'setMerges', sheet: 1, merges: [{ top: 4, left: 0, bottom: 4, right: 1 }] }], undefined, undefined);
  assert.throws(() => refreshPivot(doc.wb, 0), /merged cells A5:B5/);
  const twins = await load(await pivotBook({ data: DATA, rows: [0, 1], values: [{ fld: 2 }], twin: { row: 6, col: 0 } }));
  assert.throws(() => refreshPivot(twins.wb, 0), /would grow over the pivot table "PivotTable2"/);
});

test('tables of one cache are refreshed together', async () => {
  const doc = await load(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }], twin: { row: 2, col: 4 } }));
  setData(doc, 1, 2, 100);
  doc.history.apply(refreshPivot(doc.wb, 1), undefined, undefined);
  for (const left of [0, 4]) assert.deepEqual(grid(doc.wb, 1, 3, left, 5, left + 1), [['East', 106], ['West', 27], ['Grand Total', 133]]);
  // Refreshed together: the other one has nothing left to change.
  assert.deepEqual(refreshPivot(doc.wb, 0), []);
});

test('a source by a defined name', async () => {
  const { body } = await refreshed({ data: DATA, rows: [0], values: [{ fld: 2 }], sourceName: 'SalesData' });
  assert.deepEqual(body[3], ['Grand Total', 43]);
});

test('a cache without records: only its definition is written', async () => {
  const { doc } = await refreshed({ data: DATA, rows: [0], values: [{ fld: 2 }], noRecords: true });
  const bytes = await save(doc);
  assert.deepEqual(ZipArchive.read(bytes).names().filter((name) => /pivotCacheRecords/.test(name)), []);
  // The count of records the definition states is brought up to date all the same.
  assert.match(await text(bytes, 'xl/pivotCache/pivotCacheDefinition1.xml'), /recordCount="5"/);
});

test('a source that does not fit its cache is refused', async () => {
  const doc = await load(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] }));
  setData(doc, 0, 1, 'Item');
  assert.throws(() => refreshPivot(doc.wb, 0), /"Item", the field of the cache is "Product"/);
  setData(doc, 0, 1, null);
  assert.throws(() => refreshPivot(doc.wb, 0), /header of column B1 of its source is empty/);
});

test('what is not built here is refused with the reason, and the table stays as the file has it', async () => {
  const cases: [Partial<PivotSpec>, RegExp][] = [
    [{ fieldTail: { 0: '<fieldGroup base="0"/>' } }, /groups items/],
    [{ tableTail: '<filters count="1"><filter fld="0" type="captionEqual" evalOrder="-1" id="1"/></filters>' }, /filters by labels or by values/],
    [{ values: [{ fld: 2, name: 'Share' }], tableTail: '' }, /^$/],
    [{ tableAttrs: ' dataOnRows="1"' }, /values stand in rows/],
    [{ tableTail: '<formats count="1"><format dxfId="0"><pivotArea><references count="1"><reference field="0" count="1"><x v="0"/></reference></references></pivotArea></format></formats>' }, /formats some of its items/],
    [{ tableTail: '<chartFormats count="1"><chartFormat chart="0" format="0" series="1"><pivotArea type="data" outline="0" fieldPosition="0"/></chartFormat></chartFormats>' }, /pivot chart/],
  ];
  for (const [extra, reason] of cases) {
    const spec = { data: DATA, rows: [0], values: [{ fld: 2 }], ...extra } as PivotSpec;
    const doc = await load(await pivotBook(spec));
    if (reason.source === '^$') continue;
    assert.throws(() => refreshPivot(doc.wb, 0), reason);
    assert.match(listPivots(doc.wb)[0].refusal || '', reason);
  }
  // A value shown as a share of something.
  const share = await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] });
  const zip = ZipArchive.read(share);
  const table = new TextDecoder().decode(await zip.read('xl/pivotTables/pivotTable1.xml')).replace('baseField="0"', 'showDataAs="percentOfTotal" baseField="0"');
  zip.set('xl/pivotTables/pivotTable1.xml', new TextEncoder().encode(table));
  const doc = await load(await zip.write());
  assert.throws(() => refreshPivot(doc.wb, 0), /shown as a calculation/);
  // A cache of an outside source.
  const outside = ZipArchive.read(share);
  const cache = new TextDecoder().decode(await outside.read('xl/pivotCache/pivotCacheDefinition1.xml')).replace('<cacheSource type="worksheet">', '<cacheSource type="external" connectionId="1">');
  outside.set('xl/pivotCache/pivotCacheDefinition1.xml', new TextEncoder().encode(cache));
  const ext = await load(await outside.write());
  assert.throws(() => refreshPivot(ext.wb, 0), /does not come from a range of the workbook/);
});

test('a part that cannot be read leaves its table out; the workbook opens and saves as it was', async () => {
  for (const broken of ['<!DOCTYPE x [<!ENTITY a "b">]><pivotTableDefinition/>', 'not xml', '<pivotTableDefinition xmlns="x"><location ref="nowhere"/></pivotTableDefinition>']) {
    const zip = ZipArchive.read(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] }));
    zip.set('xl/pivotTables/pivotTable1.xml', new TextEncoder().encode(broken));
    const bytes = await zip.write();
    const doc = await load(bytes);
    assert.equal(doc.wb.pivots, undefined);
    assert.ok(same(await save(doc), bytes));
  }
  // A table whose lines reach far: the numbers are bounded, nothing hangs.
  const zip = ZipArchive.read(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] }));
  const table = new TextDecoder().decode(await zip.read('xl/pivotTables/pivotTable1.xml')).replace('<i t="grand"><x/></i>', '<i r="99999999999" i="-4"><x v="5000000000"/></i>');
  zip.set('xl/pivotTables/pivotTable1.xml', new TextEncoder().encode(table));
  const doc = await load(await zip.write());
  assert.deepEqual(doc.wb.pivots!.tables[0].rowItems, [{ r: 0, i: 0, x: [0] }]);
  doc.history.apply(refreshPivot(doc.wb, 0), undefined, undefined);
  assert.deepEqual(grid(doc.wb, 1, 3, 0, 5, 1), [['East', 16], ['West', 27], ['Grand Total', 43]]);
});

test('refresh all: what can be refreshed is, the rest is said', async () => {
  const doc = await load(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }], around: [[5, 0, 'mine']] }));
  const all = refreshAllPivots(doc.wb);
  assert.deepEqual(all.refreshed, []);
  assert.equal(all.refused.length, 1);
  assert.match(all.refused[0].reason, /A6/);
  const ok = await load(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] }));
  const before = ok.wb.sheets[1].rows;
  const result = refreshAllPivots(ok.wb);
  assert.deepEqual(result.refreshed, [0]);
  assert.equal(ok.wb.sheets[1].rows, before);
  assert.equal(ok.wb.pivots!.tables[0].now, undefined);
  ok.history.apply(result.ops, undefined, undefined);
  assert.deepEqual(grid(ok.wb, 1, 5, 0, 5, 1), [['Grand Total', 43]]);
});

test('a refresh is one undo step; the operation takes back the state of the cache and the tables', async () => {
  const doc = await load(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] }));
  const ops = refreshPivot(doc.wb, 0);
  const inverse = ops.map((op) => applyOp(doc.wb, op)).reverse();
  assert.ok(doc.wb.pivots!.caches[0].now);
  for (const op of inverse) applyOp(doc.wb, op);
  assert.equal(doc.wb.pivots!.caches[0].now, undefined);
  assert.equal(doc.wb.pivots!.tables[0].now, undefined);
  assert.throws(() => applyOp(doc.wb, { type: 'setPivot', sheet: 1, cache: 5, state: undefined, tables: [] }), /No pivot cache/);
  assert.throws(() => applyOp(doc.wb, { type: 'setPivot', sheet: 1, cache: 0, state: undefined, tables: [{ table: 3, state: undefined }] }), /not made from its cache/);
});

test('editing together: a refresh is not shared, and refused before anything changes', async () => {
  const doc = await load(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] }));
  const ops = refreshPivot(doc.wb, 0);
  const pivot = ops.find((op) => op.type === 'setPivot')!;
  assert.throws(() => encodeOp(doc.wb, pivot), (error: unknown) => error instanceof CollabUnsupportedError && /cannot be refreshed while the workbook is edited together/.test((error as Error).message));
});

test('editing together: a session refuses the refresh, and the workbook is as it was', async () => {
  const bytes = await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] });
  const net = new Net();
  const member = await join(net, 'Ada', bytes);
  net.drain(() => 0);
  const rows = member.wb.sheets[1].rows;
  const ops = refreshPivot(member.wb, 0);
  assert.equal(ops[0].type, 'setPivot');
  assert.throws(() => member.session.submit(ops), (error: unknown) => error instanceof CollabUnsupportedError);
  assert.equal(member.wb.sheets[1].rows, rows);
  assert.equal(member.wb.pivots!.caches[0].now, undefined);
  assert.equal(member.wb.pivots!.tables[0].now, undefined);
});

test('an independent reader finds the refreshed files whole and their values right', { skip: !hasPython() }, async () => {
  const data: Datum[][] = [['Region', 'Product', 'Year', 'Sales', 'Units'], ['East', 'A', 2023, 10.5, 1], ['West', 'B', 2024, 20, 2], ['East', 'B', 2023, 5, null], ['West', 'A', 2024, 7, 4], ['East', 'A', 2024, 1, 5], ['North', 'C', 2023, -3, 6]];
  const specs: PivotSpec[] = [
    { data, rows: [0], values: [{ fld: 3 }] },
    { data, rows: [0, 1], values: [{ fld: 3 }] },
    { data, rows: [0, 1], values: [{ fld: 3 }], form: 'outline' },
    { data, rows: [0, 1, 2], values: [{ fld: 3 }], form: 'tabular' },
    { data, rows: [0], cols: [1], values: [{ fld: 3, subtotal: 'average' }] },
    { data, rows: [1], cols: [0, 2], values: [{ fld: 4, subtotal: 'count' }] },
    { data, rows: [0], values: [{ fld: 3 }, { fld: 4, subtotal: 'max' }, { fld: 3, subtotal: 'stdDev' }] },
    { data, rows: [0, 1], values: [{ fld: 3 }], hidden: [{ fld: 1, values: ['B'] }], subtotalTop: false },
    { data, rows: [1], values: [{ fld: 3 }], pages: [{ fld: 0, item: 'East' }], at: { row: 4, col: 0 } },
    { data, rows: [2, 0], values: [{ fld: 3, subtotal: 'product' }, { fld: 3, subtotal: 'varp' }], showAll: true },
  ];
  const files: { [name: string]: Uint8Array } = {};
  for (let i = 0; i < specs.length; i++) {
    const doc = await load(await pivotBook(specs[i]));
    setData(doc, 3, 0, 'South');
    setData(doc, 2, 3, 42);
    doc.history.apply(refreshPivot(doc.wb, 0), undefined, undefined);
    files[`pivot${i}`] = await save(doc);
  }
  assert.deepEqual(independentPivots(files), []);
});

test('the independent reader would see a wrong value', { skip: !hasPython() }, async () => {
  const { doc } = await refreshed({ data: DATA, rows: [0], values: [{ fld: 2 }] });
  doc.history.apply([{ type: 'setCells', sheet: 1, cells: [{ row: 3, col: 1, cell: { kind: 'number', value: 17, style: 0 } }] }], undefined, undefined);
  const problems = independentPivots({ wrong: await save(doc) });
  assert.equal(problems.length, 1);
  assert.match(problems[0], /holds 17\.0, 16\.0 was expected/);
});
