// Pivot tables after their independent review: what a refresh must refuse (parts of the cache
// the model does not hold, settings it does not apply, files built to exhaust it), the grand
// totals in their places, the order of Turkish text, and a refresh that reads what another one
// wrote. Each of these failed before it was fixed.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { ZipArchive } from '../src/index';
import { refreshAllPivots, refreshPivot } from '../src/ops/pivots';
import { tableState } from '../src/model/sheet/pivot';
import { grid, hasPython, independentPivots, load, pivotBook, save, setData, type Datum, type PivotSpec } from './pivot-fixtures';

const DATA: Datum[][] = [
  ['Region', 'Product', 'Sales'],
  ['East', 'A', 10],
  ['West', 'B', 20],
  ['East', 'B', 5],
  ['West', 'A', 7],
  ['East', 'A', 1],
];

const encoder = new TextEncoder();
const decoder = new TextDecoder();

async function patch(bytes: Uint8Array, name: string, change: (text: string) => string): Promise<Uint8Array> {
  const zip = ZipArchive.read(bytes);
  zip.set(name, encoder.encode(change(decoder.decode(await zip.read(name)))));
  return zip.write();
}

async function body(spec: PivotSpec, fix?: (text: string) => string, change?: (doc: Awaited<ReturnType<typeof load>>) => void) {
  let bytes = await pivotBook(spec);
  if (fix) bytes = await patch(bytes, 'xl/pivotTables/pivotTable1.xml', fix);
  const doc = await load(bytes);
  if (change) change(doc);
  doc.history.apply(refreshPivot(doc.wb, 0), undefined, undefined);
  const ref = tableState(doc.wb.pivots!.tables[0]).location.ref;
  return { doc, body: grid(doc.wb, 1, ref.top, ref.left, ref.bottom, ref.right) };
}

async function refusal(spec: PivotSpec, fix?: (text: string) => string): Promise<string> {
  let bytes = await pivotBook(spec);
  if (fix) bytes = await patch(bytes, 'xl/pivotTables/pivotTable1.xml', fix);
  const doc = await load(bytes);
  try {
    refreshPivot(doc.wb, 0);
  } catch (error) {
    assert.ok(error instanceof RangeError);
    return (error as Error).message;
  }
  return '';
}

// ---------------------------------------------------------------------------------------------
// 1. Parts of the cache the model does not hold

test('a cache that a table the model could not take in also shows is not refreshed', async () => {
  const twin = await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }], twin: { row: 2, col: 5 } });
  // Too large to be read into the model.
  const large = await patch(twin, 'xl/pivotTables/pivotTable2.xml', (s) => s.replace('</pivotTableDefinition>', `${' '.repeat(17 * 1024 * 1024)}</pivotTableDefinition>`));
  // Not a table at all.
  const broken = await patch(twin, 'xl/pivotTables/pivotTable2.xml', () => 'not xml');
  for (const bytes of [large, broken]) {
    const doc = await load(bytes);
    assert.equal(doc.wb.pivots!.tables.length, 1);
    assert.throws(() => refreshPivot(doc.wb, 0), /another part of the workbook refers to its cache/);
  }
});

test('a cache that a slicer or a timeline reads is not refreshed', async () => {
  const zip = ZipArchive.read(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] }));
  const types = decoder.decode(await zip.read('[Content_Types].xml')).replace('</Types>', '<Override PartName="/xl/slicerCaches/slicerCache1.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/></Types>');
  zip.set('[Content_Types].xml', encoder.encode(types));
  zip.set('xl/slicerCaches/slicerCache1.xml', encoder.encode('<?xml version="1.0"?><slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="Slicer_Region" sourceName="Region"><data><tabular pivotCacheId="1"><items count="1"><i x="0" s="1"/></items></tabular></data></slicerCacheDefinition>'));
  const doc = await load(await zip.write());
  assert.throws(() => refreshPivot(doc.wb, 0), /slicer or a timeline/);
  const listed = await patch(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] }), 'xl/workbook.xml', (s) =>
    s.replace('</workbook>', '<extLst><ext uri="{46BE6895-7355-4a93-B00E-2C351335B9C9}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:timelineCacheRefs><x15:timelineCacheRef r:id="rId9"/></x15:timelineCacheRefs></ext></extLst></workbook>'),
  );
  const timeline = await load(listed);
  assert.throws(() => refreshPivot(timeline.wb, 0), /slicer or a timeline/);
});

// ---------------------------------------------------------------------------------------------
// 2-4. Files built to exhaust a refresh

test('lines of rows that name more fields than the table has leave the table out', async () => {
  const lines: string[] = [];
  for (let i = 0; i < 3000; i++) lines.push(`<i r="${i}"><x/></i>`);
  const bytes = await patch(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] }), 'xl/pivotTables/pivotTable1.xml', (s) =>
    s.replace(/<rowItems[^]*?<\/rowItems>/, `<rowItems count="3000">${lines.join('')}</rowItems>`),
  );
  const doc = await load(bytes);
  assert.equal(doc.wb.pivots, undefined);
});

test('a field named twice in rows and columns leaves the table out', async () => {
  const bytes = await patch(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] }), 'xl/pivotTables/pivotTable1.xml', (s) =>
    s.replace(/<rowFields[^]*?<\/rowFields>/, `<rowFields count="3">${'<field x="0"/>'.repeat(3)}</rowFields>`),
  );
  assert.equal((await load(bytes)).wb.pivots, undefined);
  const twice = await patch(await pivotBook({ data: DATA, rows: [0], cols: [1], values: [{ fld: 2 }] }), 'xl/pivotTables/pivotTable1.xml', (s) =>
    s.replace(/<colFields[^]*?<\/colFields>/, '<colFields count="1"><field x="0"/></colFields>'),
  );
  assert.equal((await load(twice)).wb.pivots, undefined);
});

test('lines that do not fit where the table stands cost nothing to read', async () => {
  const N = 20000;
  const bytes = await patch(await pivotBook({ data: [['R', 'C', 'V'], ['a', 'x', 1], ['b', 'y', 2]], rows: [0], cols: [1], values: [{ fld: 2 }] }), 'xl/pivotTables/pivotTable1.xml', (s) =>
    s.replace(/<rowItems[^]*?<\/rowItems>/, `<rowItems count="${N}">${'<i/>'.repeat(N)}</rowItems>`).replace(/<colItems[^]*?<\/colItems>/, `<colItems count="${N}">${'<i/>'.repeat(N)}</colItems>`),
  );
  const doc = await load(bytes);
  const started = Date.now();
  try {
    refreshPivot(doc.wb, 0);
  } catch (error) {
    assert.ok(error instanceof RangeError);
  }
  assert.ok(Date.now() - started < 1000, `${Date.now() - started} ms`);
});

// ---------------------------------------------------------------------------------------------
// 5. Grand totals

test('grand totals: of the rows in the column at the right, of the columns in the row at the bottom', async () => {
  const data: Datum[][] = [['R', 'C', 'V'], ['a', 'x', 1], ['b', 'y', 2]];
  const spec = { data, rows: [0], cols: [1], values: [{ fld: 2 }] };
  const both = await body(spec);
  assert.deepEqual(both.body, [
    ['Sum of V', 'Column Labels', undefined, undefined],
    ['Row Labels', 'x', 'y', 'Grand Total'],
    ['a', 1, undefined, 1],
    ['b', undefined, 2, 2],
    ['Grand Total', 1, 2, 3],
  ]);
  const noRows = await body({ ...spec, rowGrandTotals: false });
  assert.deepEqual(noRows.body, [
    ['Sum of V', 'Column Labels', undefined],
    ['Row Labels', 'x', 'y'],
    ['a', 1, undefined],
    ['b', undefined, 2],
    ['Grand Total', 1, 2],
  ]);
  const noCols = await body({ ...spec, colGrandTotals: false });
  assert.deepEqual(noCols.body, [
    ['Sum of V', 'Column Labels', undefined, undefined],
    ['Row Labels', 'x', 'y', 'Grand Total'],
    ['a', 1, undefined, 1],
    ['b', undefined, 2, 2],
  ]);
  const none = await body({ ...spec, rowGrandTotals: false, colGrandTotals: false });
  assert.deepEqual(none.body, [
    ['Sum of V', 'Column Labels', undefined],
    ['Row Labels', 'x', 'y'],
    ['a', 1, undefined],
    ['b', undefined, 2],
  ]);
  // Without fields in columns there is no column of grand totals: the row at the bottom is that of the columns.
  const rowsOnly = await body({ data, rows: [0], values: [{ fld: 2 }], colGrandTotals: false });
  assert.deepEqual(rowsOnly.body, [['Row Labels', 'Sum of V'], ['a', 1], ['b', 2]]);
  const keep = await body({ data, rows: [0], values: [{ fld: 2 }], rowGrandTotals: false });
  assert.deepEqual(keep.body[keep.body.length - 1], ['Grand Total', 3]);
});

// ---------------------------------------------------------------------------------------------
// 6. Settings that change what a table shows

test('settings a refresh does not apply are refused', async () => {
  const spec: PivotSpec = { data: DATA, rows: [0, 1], values: [{ fld: 2 }] };
  const cases: [string, (s: string) => string, RegExp][] = [
    ['sd=0', (s) => s.replace('<item x="0"/>', '<item x="0" sd="0"/>'), /collapsed/],
    ['subtotalCaption', (s) => s.replace('<pivotField axis="axisRow"', '<pivotField axis="axisRow" subtotalCaption="Toplam!"'), /subtotalCaption/],
    ['fillDownLabelsDefault', (s) => s.replace('</pivotTableDefinition>', '<extLst><ext uri="{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:pivotTableDefinition fillDownLabelsDefault="1"/></ext></extLst></pivotTableDefinition>'), /fillDownLabelsDefault/],
    ['visualTotals', (s) => s.replace('<pivotTableDefinition ', '<pivotTableDefinition visualTotals="0" '), /visualTotals/],
    ['subtotalHiddenItems', (s) => s.replace('<pivotTableDefinition ', '<pivotTableDefinition subtotalHiddenItems="1" '), /subtotalHiddenItems/],
    ['showEmptyRow', (s) => s.replace('<pivotTableDefinition ', '<pivotTableDefinition showEmptyRow="1" '), /showEmptyRow/],
    ['autoShow', (s) => s.replace('<pivotField axis="axisRow"', '<pivotField axis="axisRow" autoShow="1"'), /autoShow/],
    ['unknown attribute', (s) => s.replace('<pivotTableDefinition ', '<pivotTableDefinition futureThing="7" '), /futureThing/],
    ['unknown element', (s) => s.replace('</pivotTableDefinition>', '<futureElement/></pivotTableDefinition>'), /futureElement/],
    ['item kind', (s) => s.replace('<item t="default"/>', '<item t="sum"/>'), /"sum"/],
    ['showDataAs in an extension', (s) => s.replace(/<dataField ([^>]*)\/>/, '<dataField $1><extLst><ext uri="{E15A36E0-9728-4e99-A89B-3F7291B0FE68}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dataField pivotShowAs="percentOfParentRow"/></ext></extLst></dataField>'), /pivotShowAs/],
  ];
  for (const [label, fix, reason] of cases) assert.match(await refusal(spec, fix), reason, label);
  // What the reader may pass over stays allowed: an attribute of a namespace the part says may be ignored.
  const ignorable = await refusal(spec, (s) => s.replace('<pivotTableDefinition ', '<pivotTableDefinition xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:z="urn:z" mc:Ignorable="z" z:note="1" '));
  assert.equal(ignorable, '');
});

test('captions of missing values and errors, and new items hidden, are applied', async () => {
  const missing = await body({ data: DATA, rows: [0], cols: [1], values: [{ fld: 2 }] }, (s) => s.replace('<pivotTableDefinition ', '<pivotTableDefinition missingCaption="-" '), (doc) => setData(doc, 2, 0, 'North'));
  // North has only B: its A is empty, shown with the caption.
  assert.deepEqual(missing.body[4], ['North', '-', 20, 20]);
  const off = await body({ data: DATA, rows: [0], cols: [1], values: [{ fld: 2 }] }, (s) => s.replace('<pivotTableDefinition ', '<pivotTableDefinition showMissing="0" missingCaption="-" '), (doc) => setData(doc, 2, 0, 'North'));
  assert.deepEqual(off.body[4], ['North', undefined, 20, 20]);
  const errors = await body({ data: DATA, rows: [0], values: [{ fld: 2 }] }, (s) => s.replace('<pivotTableDefinition ', '<pivotTableDefinition showError="1" errorCaption="HATA" '), (doc) =>
    doc.history.apply([{ type: 'setCells', sheet: 0, cells: [{ row: 1, col: 2, cell: { kind: 'error', value: '#N/A', style: 0 } }] }], undefined, undefined),
  );
  assert.deepEqual(errors.body.slice(1), [['East', 'HATA'], ['West', 27], ['Grand Total', 'HATA']]);
  const hidden = await body({ data: DATA, rows: [0], values: [{ fld: 2 }] }, (s) => s.replace('<pivotField axis="axisRow"', '<pivotField axis="axisRow" hideNewItems="1"'), (doc) => setData(doc, 1, 0, 'North'));
  assert.deepEqual(hidden.body.slice(1), [['East', 6], ['West', 27], ['Grand Total', 33]]);
});

// ---------------------------------------------------------------------------------------------
// 7. The order of text

test('items sorted ascending are in the order of the Turkish alphabet, numbers first', async () => {
  const names: Datum[] = ['Zonguldak', 'Çorum', 'Ankara', 'İzmir', 'Iğdır', 'Ödemiş', 'Uşak', 'Ümraniye', 'Évian', 'Edirne', 'Eskişehir', 'Fatsa', 'Bursa', 'Şile', 'Sinop', 'Ordu', 'Ğ', 'Gebze', 'Hatay', 12, 7, 'ılgaz', 'irmak'];
  const data: Datum[][] = [['City', 'V'], ...names.map((name, i) => [name, i] as Datum[])];
  const { body: rows } = await body({ data, rows: [0], values: [{ fld: 1 }], sortType: [{ fld: 0, type: 'ascending' }] });
  // Written out by hand: A B C Ç D E É F G Ğ H I(ı) İ(i) ... O Ö ... S Ş ... U Ü ... Z.
  assert.deepEqual(
    rows.slice(1, -1).map((row) => row[0]),
    [7, 12, 'Ankara', 'Bursa', 'Çorum', 'Edirne', 'Eskişehir', 'Évian', 'Fatsa', 'Gebze', 'Ğ', 'Hatay', 'Iğdır', 'ılgaz', 'irmak', 'İzmir', 'Ordu', 'Ödemiş', 'Sinop', 'Şile', 'Uşak', 'Ümraniye', 'Zonguldak'],
  );
});

// ---------------------------------------------------------------------------------------------
// 8. A refresh that reads what another one writes

/** The workbook of `spec` with a sheet "Calc" whose B2 reads "Pivot"!`read` by a formula, and a second cache and table over Calc!A1:B2. */
async function withSecondCache(spec: PivotSpec, read: string): Promise<Uint8Array> {
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const P = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml';
  const zip = ZipArchive.read(await pivotBook(spec));
  const edit = async (name: string, change: (s: string) => string) => zip.set(name, encoder.encode(change(decoder.decode(await zip.read(name)))));
  await edit('[Content_Types].xml', (s) =>
    s.replace(
      '</Types>',
      `<Override PartName="/xl/worksheets/sheet3.xml" ContentType="${CT}.worksheet+xml"/><Override PartName="/xl/pivotCache/pivotCacheDefinition2.xml" ContentType="${CT}.pivotCacheDefinition+xml"/><Override PartName="/xl/pivotCache/pivotCacheRecords2.xml" ContentType="${CT}.pivotCacheRecords+xml"/><Override PartName="/xl/pivotTables/pivotTable9.xml" ContentType="${CT}.pivotTable+xml"/></Types>`,
    ),
  );
  await edit('xl/workbook.xml', (s) => s.replace('</sheets>', '<sheet name="Calc" sheetId="3" r:id="rId9"/></sheets>').replace('</pivotCaches>', '<pivotCache cacheId="8" r:id="rId10"/></pivotCaches>'));
  await edit('xl/_rels/workbook.xml.rels', (s) =>
    s.replace('</Relationships>', `<Relationship Id="rId9" Type="${R}/worksheet" Target="worksheets/sheet3.xml"/><Relationship Id="rId10" Type="${R}/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition2.xml"/></Relationships>`),
  );
  zip.set(
    'xl/worksheets/sheet3.xml',
    encoder.encode(
      `<?xml version="1.0"?><worksheet xmlns="${NS}" xmlns:r="${R}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>K</t></is></c><c r="B1" t="inlineStr"><is><t>V</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>x</t></is></c><c r="B2"><f>${read}</f><v>0</v></c></row>` +
        '<row r="4"><c r="A4" t="inlineStr"><is><t>Row Labels</t></is></c><c r="B4" t="inlineStr"><is><t>Sum of V</t></is></c></row><row r="5"><c r="A5" t="inlineStr"><is><t>Grand Total</t></is></c><c r="B5"><v>0</v></c></row></sheetData></worksheet>',
    ),
  );
  zip.set('xl/worksheets/_rels/sheet3.xml.rels', encoder.encode(`<?xml version="1.0"?><Relationships xmlns="${P}"><Relationship Id="rId1" Type="${R}/pivotTable" Target="../pivotTables/pivotTable9.xml"/></Relationships>`));
  zip.set(
    'xl/pivotCache/pivotCacheDefinition2.xml',
    encoder.encode(
      `<?xml version="1.0"?><pivotCacheDefinition xmlns="${NS}" xmlns:r="${R}" r:id="rId1" recordCount="1"><cacheSource type="worksheet"><worksheetSource ref="A1:B2" sheet="Calc"/></cacheSource><cacheFields count="2"><cacheField name="K" numFmtId="0"><sharedItems count="1"><s v="x"/></sharedItems></cacheField><cacheField name="V" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="0" maxValue="0"/></cacheField></cacheFields></pivotCacheDefinition>`,
    ),
  );
  zip.set('xl/pivotCache/_rels/pivotCacheDefinition2.xml.rels', encoder.encode(`<?xml version="1.0"?><Relationships xmlns="${P}"><Relationship Id="rId1" Type="${R}/pivotCacheRecords" Target="pivotCacheRecords2.xml"/></Relationships>`));
  zip.set('xl/pivotCache/pivotCacheRecords2.xml', encoder.encode(`<?xml version="1.0"?><pivotCacheRecords xmlns="${NS}" count="1"><r><x/><n v="0"/></r></pivotCacheRecords>`));
  zip.set(
    'xl/pivotTables/pivotTable9.xml',
    encoder.encode(
      `<?xml version="1.0"?><pivotTableDefinition xmlns="${NS}" name="Second" cacheId="8" dataCaption="Values"><location ref="A4:B5" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="2"><pivotField axis="axisRow" showAll="0"><items count="2"><item x="0"/><item t="default"/></items></pivotField><pivotField dataField="1" showAll="0"/></pivotFields><rowFields count="1"><field x="0"/></rowFields><rowItems count="1"><i t="grand"><x/></i></rowItems><colItems count="1"><i/></colItems><dataFields count="1"><dataField name="Sum of V" fld="1" baseField="0" baseItem="0"/></dataFields></pivotTableDefinition>`,
    ),
  );
  zip.set('xl/pivotTables/_rels/pivotTable9.xml.rels', encoder.encode(`<?xml version="1.0"?><Relationships xmlns="${P}"><Relationship Id="rId1" Type="${R}/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition2.xml"/></Relationships>`));
  return zip.write();
}

test('refresh all: a source with formulas after a refresh that changed cells waits for its own refresh', async () => {
  const doc = await load(await withSecondCache({ data: DATA, rows: [0], values: [{ fld: 2 }] }, 'Pivot!B6'));
  assert.equal(doc.wb.pivots!.tables.length, 2);
  const all = refreshAllPivots(doc.wb);
  assert.deepEqual(all.refreshed, [0]);
  assert.equal(all.refused.length, 1);
  assert.equal(all.refused[0].name, 'Second');
  assert.match(all.refused[0].reason, /formulas/);
  // Alone it is refreshed (whoever calls it calculates the formulas first).
  const alone = await load(await withSecondCache({ data: DATA, rows: [0], values: [{ fld: 2 }] }, 'Pivot!B6'));
  assert.doesNotThrow(() => refreshPivot(alone.wb, 1));
});

// ---------------------------------------------------------------------------------------------
// Small ones

test('a refresh that would change nothing is no step', async () => {
  const { doc } = await body({ data: DATA, rows: [0], values: [{ fld: 2 }] });
  assert.deepEqual(refreshPivot(doc.wb, 0), []);
  setData(doc, 1, 2, 11);
  assert.notDeepEqual(refreshPivot(doc.wb, 0), []);
});

test('a total past the largest number is #NUM!, in the model as in the file', async () => {
  const { doc, body: rows } = await body({ data: [['K', 'V'], ['a', 1e308], ['a', 1e308], ['b', 1e200]], rows: [0], values: [{ fld: 1 }] });
  assert.deepEqual(rows.slice(1), [['a', '#NUM!'], ['b', 1e200], ['Grand Total', '#NUM!']]);
  assert.equal(doc.wb.sheets[1].rows.get(3)!.cells[1].kind, 'error');
  if (hasPython()) assert.deepEqual(independentPivots({ inf: await save(doc) }), []);
});

test('two refreshes and one undo: the state of the first refresh, not that of the file', async () => {
  const { doc } = await body({ data: DATA, rows: [0], values: [{ fld: 2 }] });
  const first = doc.wb.pivots!.caches[0].now;
  const firstTable = doc.wb.pivots!.tables[0].now;
  setData(doc, 1, 0, 'North');
  doc.history.apply(refreshPivot(doc.wb, 0), undefined, undefined);
  assert.notEqual(doc.wb.pivots!.caches[0].now, first);
  doc.history.undo();
  assert.equal(doc.wb.pivots!.caches[0].now, first);
  assert.equal(doc.wb.pivots!.tables[0].now, firstTable);
});
