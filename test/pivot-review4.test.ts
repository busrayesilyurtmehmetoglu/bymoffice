// Pivot tables after their fourth review: a chart too deep to be looked through is a pivot chart
// of no known table; one interaction looks through the workbook once; names of months and days in
// capitals; letters made of a Turkish letter and other marks; the word of subtotals.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { ZipArchive } from '../src/index';
import { addChart } from '../src/ops/charts';
import { PIVOT_STATS, compareText, lookThroughPivots, refreshAllPivots, refreshPivot } from '../src/ops/pivots';
import { PIVOT_SCAN_LIMITS, tableState } from '../src/model/sheet/pivot';
import { grid, load, pivotBook, save, setData, type Datum } from './pivot-fixtures';

const encoder = new TextEncoder();
const decoder = new TextDecoder();
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml';
const SMALL: Datum[][] = [['R', 'V'], ['a', 1], ['b', 2]];

const reason = (run: () => unknown): string => {
  try {
    run();
  } catch (error) {
    assert.ok(error instanceof RangeError, String(error));
    return (error as Error).message;
  }
  return '';
};

async function edit(bytes: Uint8Array, change: (zip: ZipArchive, read: (name: string) => Promise<string>) => void | Promise<void>): Promise<Uint8Array> {
  const zip = ZipArchive.read(bytes);
  await change(zip, async (name) => decoder.decode(await zip.read(name)));
  return zip.write();
}

/** A workbook with `count` caches, each with a table of its own on a sheet of its own. */
async function manyCaches(count: number): Promise<Uint8Array> {
  return edit(await pivotBook({ data: SMALL, rows: [0], values: [{ fld: 1 }] }), async (zip, read) => {
    const def = await read('xl/pivotCache/pivotCacheDefinition1.xml');
    const rec = await read('xl/pivotCache/pivotCacheRecords1.xml');
    const defRels = await read('xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels');
    const table = await read('xl/pivotTables/pivotTable1.xml');
    const tableRels = await read('xl/pivotTables/_rels/pivotTable1.xml.rels');
    const sheet = await read('xl/worksheets/sheet2.xml');
    const sheetRels = await read('xl/worksheets/_rels/sheet2.xml.rels');
    let sheets = '';
    let caches = '';
    let rels = '';
    let types = '';
    for (let i = 2; i <= count; i++) {
      zip.set(`xl/pivotCache/pivotCacheDefinition${i}.xml`, encoder.encode(def));
      zip.set(`xl/pivotCache/pivotCacheRecords${i}.xml`, encoder.encode(rec));
      zip.set(`xl/pivotCache/_rels/pivotCacheDefinition${i}.xml.rels`, encoder.encode(defRels.replace('pivotCacheRecords1', `pivotCacheRecords${i}`)));
      zip.set(`xl/pivotTables/pivotTable${i}.xml`, encoder.encode(table.replace('cacheId="7"', `cacheId="${100 + i}"`).replace('PivotTable1', `PivotTable${i}`)));
      zip.set(`xl/pivotTables/_rels/pivotTable${i}.xml.rels`, encoder.encode(tableRels.replace('pivotCacheDefinition1', `pivotCacheDefinition${i}`)));
      zip.set(`xl/worksheets/sheet${i + 1}.xml`, encoder.encode(sheet));
      zip.set(`xl/worksheets/_rels/sheet${i + 1}.xml.rels`, encoder.encode(sheetRels.replace('pivotTable1', `pivotTable${i}`)));
      sheets += `<sheet name="P${i}" sheetId="${i + 1}" r:id="rS${i}"/>`;
      caches += `<pivotCache cacheId="${100 + i}" r:id="rC${i}"/>`;
      rels += `<Relationship Id="rS${i}" Type="${OFFICE_REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/><Relationship Id="rC${i}" Type="${OFFICE_REL}/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition${i}.xml"/>`;
      types +=
        `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="${CT}.worksheet+xml"/><Override PartName="/xl/pivotCache/pivotCacheDefinition${i}.xml" ContentType="${CT}.pivotCacheDefinition+xml"/>` +
        `<Override PartName="/xl/pivotCache/pivotCacheRecords${i}.xml" ContentType="${CT}.pivotCacheRecords+xml"/><Override PartName="/xl/pivotTables/pivotTable${i}.xml" ContentType="${CT}.pivotTable+xml"/>`;
    }
    zip.set('xl/workbook.xml', encoder.encode((await read('xl/workbook.xml')).replace('</sheets>', `${sheets}</sheets>`).replace('</pivotCaches>', `${caches}</pivotCaches>`)));
    zip.set('xl/_rels/workbook.xml.rels', encoder.encode((await read('xl/_rels/workbook.xml.rels')).replace('</Relationships>', `${rels}</Relationships>`)));
    zip.set('[Content_Types].xml', encoder.encode((await read('[Content_Types].xml')).replace('</Types>', `${types}</Types>`)));
  });
}

// ---------------------------------------------------------------------------------------------
// 1. A chart too deep to be looked through

test('a pivot chart whose part is larger than is looked through is one of no known table', async () => {
  const doc0 = await load(await pivotBook({ data: SMALL, rows: [0], values: [{ fld: 1 }] }));
  doc0.history.apply(addChart(doc0.wb, 1, { type: 'bar', range: { top: 0, left: 0, bottom: 2, right: 1 } } as never), undefined, undefined);
  const bytes = await edit(await save(doc0), async (zip, read) => {
    const name = zip.names().find((entry) => /charts\/chart/.test(entry)) as string;
    zip.set(
      name,
      encoder.encode(
        (await read(name))
          .replace(/<c:chart>/, '<c:pivotSource><c:name>[x.xlsx]Pivot!PivotTable1</c:name><c:fmtId val="0"/></c:pivotSource><c:chart>')
          .replace(/<\/c:chartSpace>/, `<c:extLst><c:ext uri="{X}">${'<a/>'.repeat(200)}</c:ext></c:extLst></c:chartSpace>`),
      ),
    );
  });
  const limit = PIVOT_SCAN_LIMITS.chartNodes;
  PIVOT_SCAN_LIMITS.chartNodes = 50;
  try {
    const doc = await load(bytes);
    assert.equal(doc.wb.sheets[1].drawings![0].object.type === 'chart' && (doc.wb.sheets[1].drawings![0].object as { chart: { pivotSource?: string } }).chart.pivotSource, '');
    setData(doc, 1, 0, 'zz');
    assert.match(reason(() => refreshPivot(doc.wb, 0)), /pivot chart is made from a table that is not known here/);
  } finally {
    PIVOT_SCAN_LIMITS.chartNodes = limit;
  }
});

// ---------------------------------------------------------------------------------------------
// 2. One look for one interaction

test('refreshing every table of many caches looks through the workbook once; a menu and the refresh after it too', async () => {
  const doc = await load(await manyCaches(4));
  assert.equal(doc.wb.pivots!.caches.length, 4);
  setData(doc, 1, 0, 'zz');
  const before = PIVOT_STATS.scans;
  const all = refreshAllPivots(doc.wb);
  assert.equal(all.refreshed.length, 4);
  assert.equal(PIVOT_STATS.scans - before, 1);
  // The menu, then the refresh: one look, as long as nothing was changed in between.
  const other = await load(await manyCaches(2));
  setData(other, 1, 0, 'zz');
  const start = PIVOT_STATS.scans;
  lookThroughPivots(other.wb);
  lookThroughPivots(other.wb);
  refreshPivot(other.wb, 0);
  assert.equal(PIVOT_STATS.scans - start, 1);
  // An edit since: looked through again.
  setData(other, 2, 0, 'yy');
  lookThroughPivots(other.wb);
  assert.equal(PIVOT_STATS.scans - start, 2);
});

test('relationships and charts together are looked through within one budget', async () => {
  const bytes = await edit(await pivotBook({ data: SMALL, rows: [0], values: [{ fld: 1 }] }), (zip) =>
    zip.set('xl/charts/chart1.xml', encoder.encode('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart/></c:chartSpace>')),
  );
  const zip = ZipArchive.read(bytes);
  let rels = 0;
  // The relationships of parts (those of the package itself name no part of the workbook and are not looked at).
  for (const name of zip.names()) if (/\/_rels\/[^/]+\.rels$/.test(name)) rels += (await zip.read(name)).length;
  const limits = { ...PIVOT_SCAN_LIMITS };
  try {
    PIVOT_SCAN_LIMITS.totalBytes = rels + 5;
    const doc = await load(bytes);
    assert.match(reason(() => refreshPivot(doc.wb, 0)), /cannot tell what refers to its cache/);
    // Without the chart the same budget is enough.
    const plain = await load(await pivotBook({ data: SMALL, rows: [0], values: [{ fld: 1 }] }));
    assert.equal(reason(() => refreshPivot(plain.wb, 0)), '');
  } finally {
    Object.assign(PIVOT_SCAN_LIMITS, limits);
  }
});

// ---------------------------------------------------------------------------------------------
// 3, 4. Names in capitals; letters made of more than one mark

test('names of months and days written in capitals are names all the same', async () => {
  const run = async (names: Datum[]) => {
    const doc = await load(await pivotBook({ data: [['M', 'V'], ...names.map((name, i) => [name, i] as Datum[])], rows: [0], values: [{ fld: 1 }], sortType: [{ fld: 0, type: 'ascending' }] }));
    return reason(() => refreshPivot(doc.wb, 0));
  };
  assert.match(await run(['APRIL', 'Zeta', 'Beta']), /names of months or days/);
  assert.match(await run(['FRI', 'SAT', 'SUN']), /names of months or days/);
  assert.match(await run(['FRIDAY', 'x']), /names of months or days/);
  assert.match(await run(['PAZARTESİ', 'x']), /names of months or days/);
});

test('a letter made of a Turkish letter and another mark is that Turkish letter with an accent', () => {
  assert.deepEqual(['ǖa', 'üb', 'ub', 'uz'].sort(compareText), ['ub', 'uz', 'ǖa', 'üb']);
  assert.deepEqual(['ḉa', 'çb', 'cz'].sort(compareText), ['cz', 'ḉa', 'çb']);
  assert.deepEqual(['ȫa', 'oz', 'öb'].sort(compareText), ['oz', 'ȫa', 'öb']);
});

// ---------------------------------------------------------------------------------------------
// 5. Records referred to by their own type

test('another part that refers to the records by the type of records refuses the refresh', async () => {
  const doc = await load(
    await edit(await pivotBook({ data: SMALL, rows: [0], values: [{ fld: 1 }] }), (zip) => {
      zip.set('xl/other/x.xml', encoder.encode('<x/>'));
      zip.set('xl/other/_rels/x.xml.rels', encoder.encode(`<?xml version="1.0"?><Relationships xmlns="${REL_NS}"><Relationship Id="r1" Type="${OFFICE_REL}/pivotCacheRecords" Target="../pivotCache/pivotCacheRecords1.xml"/></Relationships>`));
    }),
  );
  assert.match(reason(() => refreshPivot(doc.wb, 0)), /refers to its records/);
});

// ---------------------------------------------------------------------------------------------
// 6. The word of subtotals

test('subtotals are named in the words of the table: "Toplam" beside "Genel Toplam"', async () => {
  const doc = await load(await pivotBook({ data: [['R', 'P', 'V'], ['East', 'A', 1], ['East', 'B', 2], ['West', 'A', 3]], rows: [0, 1], values: [{ fld: 2 }], form: 'tabular', words: { grand: 'Genel Toplam' } }));
  doc.history.apply(refreshPivot(doc.wb, 0), undefined, undefined);
  const ref = tableState(doc.wb.pivots!.tables[0]).location.ref;
  const labels = grid(doc.wb, 1, ref.top, ref.left, ref.bottom, ref.left).map((row) => row[0]);
  assert.ok(labels.indexOf('East Toplam') >= 0, JSON.stringify(labels));
  assert.equal(labels[labels.length - 1], 'Genel Toplam');
  const english = await load(await pivotBook({ data: [['R', 'P', 'V'], ['East', 'A', 1], ['East', 'B', 2]], rows: [0, 1], values: [{ fld: 2 }], form: 'tabular' }));
  english.history.apply(refreshPivot(english.wb, 0), undefined, undefined);
  const refE = tableState(english.wb.pivots!.tables[0]).location.ref;
  assert.ok(grid(english.wb, 1, refE.top, refE.left, refE.bottom, refE.left).some((row) => row[0] === 'East Total'));
});
