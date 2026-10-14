// Pivot tables after their third review: what refers to a cache is judged on the workbook as it
// is when a refresh is made (the charts of the model with the parts of the package), whatever was
// looked through before; part names with percent signs on both sides; names of months and days;
// the order of text by keys made once per item, with marks of every script as accents.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { ZipArchive } from '../src/index';
import { addChart, removeChart } from '../src/ops/charts';
import { PIVOT_STATS, compareText, listPivots, lookThroughPivots, refreshPivot } from '../src/ops/pivots';
import { PIVOT_SCAN_LIMITS } from '../src/model/sheet/pivot';
import { load, pivotBook, save, setData, type Datum } from './pivot-fixtures';

const encoder = new TextEncoder();
const decoder = new TextDecoder();
const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
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

async function edit(bytes: Uint8Array, change: (zip: ZipArchive) => void | Promise<void>): Promise<Uint8Array> {
  const zip = ZipArchive.read(bytes);
  await change(zip);
  return zip.write();
}

const text = async (zip: ZipArchive, name: string) => decoder.decode(await zip.read(name));

/** A workbook whose table has a pivot chart: a chart added on "Pivot", whose part then names the table as its source. */
async function withPivotChart(): Promise<Uint8Array> {
  const doc = await load(await pivotBook({ data: SMALL, rows: [0], values: [{ fld: 1 }] }));
  doc.history.apply(addChart(doc.wb, 1, { type: 'bar', range: { top: 0, left: 0, bottom: 2, right: 1 } } as never), undefined, undefined);
  return edit(await save(doc), async (zip) => {
    const name = zip.names().find((entry) => /charts\/chart/.test(entry)) as string;
    zip.set(name, encoder.encode((await text(zip, name)).replace(/<c:chart>/, '<c:pivotSource><c:name>[x.xlsx]Pivot!PivotTable1</c:name><c:fmtId val="0"/></c:pivotSource><c:chart>')));
  });
}

// ---------------------------------------------------------------------------------------------
// 1. Judged on the workbook as it is

test('a pivot chart that comes back by undo after what was looked through without it still refuses the refresh', async () => {
  const bytes = await withPivotChart();
  const first = await load(bytes);
  setData(first, 1, 0, 'zz');
  assert.match(reason(() => refreshPivot(first.wb, 0)), /pivot chart/);
  // Nothing was looked through before the chart went.
  const doc = await load(bytes);
  setData(doc, 1, 0, 'zz');
  doc.history.apply(removeChart(doc.wb, 1, 0), undefined, undefined);
  await save(doc);
  // The menu looks through the package as it is now: without the chart.
  lookThroughPivots(doc.wb);
  doc.history.undo();
  assert.equal(doc.wb.sheets[1].drawings!.length, 1);
  assert.match(reason(() => refreshPivot(doc.wb, 0)), /pivot chart/);
  assert.match(listPivots(doc.wb)[0].refusal || '', /pivot chart/, 'and the list says so once it was looked through again');
});

// ---------------------------------------------------------------------------------------------
// 2. Percent signs on both sides

test('a cache whose own name has percent signs is found by a part that names it so', async () => {
  for (const odd of ['pivot%20Cache1.xml', 'pivot%2541.xml']) {
    const bytes = await edit(await pivotBook({ data: SMALL, rows: [0], values: [{ fld: 1 }] }), async (zip) => {
      const def = await zip.read('xl/pivotCache/pivotCacheDefinition1.xml');
      const rels = await zip.read('xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels');
      zip.delete('xl/pivotCache/pivotCacheDefinition1.xml');
      zip.delete('xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels');
      zip.set(`xl/pivotCache/${odd}`, def);
      zip.set(`xl/pivotCache/_rels/${odd}.rels`, rels);
      for (const name of ['xl/_rels/workbook.xml.rels', '[Content_Types].xml', 'xl/pivotTables/_rels/pivotTable1.xml.rels']) zip.set(name, encoder.encode((await text(zip, name)).split('pivotCacheDefinition1.xml').join(odd)));
      zip.set('xl/other/x.xml', encoder.encode('<x/>'));
      zip.set('xl/other/_rels/x.xml.rels', encoder.encode(`<?xml version="1.0"?><Relationships xmlns="${REL_NS}"><Relationship Id="r1" Type="${OFFICE_REL}/pivotCacheDefinition" Target="../pivotCache/${odd}"/></Relationships>`));
    });
    const doc = await load(bytes);
    assert.equal(doc.wb.pivots!.tables.length, 1, odd);
    assert.match(reason(() => refreshPivot(doc.wb, 0)), /another part of the workbook refers to its cache/, odd);
  }
});

// ---------------------------------------------------------------------------------------------
// 3. Names of months and days

test('a sorted field with any full name of a month or a day is refused; short names only when every item is one', async () => {
  const run = async (names: Datum[]) => {
    const doc = await load(await pivotBook({ data: [['M', 'V'], ...names.map((name, i) => [name, i] as Datum[])], rows: [0], values: [{ fld: 1 }], sortType: [{ fld: 0, type: 'ascending' }] }));
    return reason(() => refreshPivot(doc.wb, 0));
  };
  assert.match(await run(['Ocak', 'Ankara']), /names of months or days/);
  assert.match(await run(['January', 12]), /names of months or days/);
  assert.match(await run(['ara', 'cum']), /names of months or days/);
  assert.equal(await run(['ara', 'cum', 'Ankara']), '');
});

// ---------------------------------------------------------------------------------------------
// 4, 5. The order of text

test('items are sorted by keys made once for each of them', async () => {
  const names: Datum[] = [];
  for (let i = 0; i < 2000; i++) names.push(`İtem ${(i * 7919) % 2000}`);
  const doc = await load(await pivotBook({ data: [['K', 'V'], ...names.map((name, i) => [name, i] as Datum[])], rows: [0], values: [{ fld: 1 }], sortType: [{ fld: 0, type: 'ascending' }] }));
  const before = PIVOT_STATS.textKeys;
  refreshPivot(doc.wb, 0);
  assert.ok(PIVOT_STATS.textKeys - before <= 2000, `${PIVOT_STATS.textKeys - before} keys`);
});

test('accents weigh before case, and marks of every script are accents', () => {
  assert.deepEqual(['é', 'E'].sort(compareText), ['E', 'é']);
  // Arabic: a vowel sign is no letter (َ), nor is the grapheme joiner (͏).
  assert.deepEqual(['بَس', 'بت'].sort(compareText), ['بت', 'بَس']);
  assert.deepEqual(['بَ', 'ب'].sort(compareText), ['ب', 'بَ']);
  assert.deepEqual(['a͏c', 'ab'].sort(compareText), ['ab', 'a͏c']);
  // Hebrew points as well.
  assert.deepEqual(['בָג', 'בא'].sort(compareText), ['בא', 'בָג']);
});

// ---------------------------------------------------------------------------------------------
// 6. A pivot source anywhere in a chart

test('a pivot source after the chart, in alternate content, with its name after other children, or written with a reference', async () => {
  const bodies = [
    '<c:chart/><c:pivotSource><c:name>Pivot!PivotTable1</c:name></c:pivotSource>',
    '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c"><c:pivotSource><c:name>Pivot!PivotTable1</c:name></c:pivotSource></mc:Choice></mc:AlternateContent><c:chart/>',
    '<c:pivotSource><c:fmtId val="0"/><c:name>Pivot!PivotTable1</c:name></c:pivotSource><c:chart/>',
    '<c:pivotSource><c:name>Pivot!Pivot&#84;able1</c:name></c:pivotSource><c:chart/>',
  ];
  for (const body of bodies) {
    const doc = await load(await edit(await pivotBook({ data: SMALL, rows: [0], values: [{ fld: 1 }] }), (zip) => zip.set('xl/charts/chart1.xml', encoder.encode(`<c:chartSpace xmlns:c="${CHART_NS}">${body}</c:chartSpace>`))));
    assert.match(reason(() => refreshPivot(doc.wb, 0)), /pivot chart is made from a table of its cache/, body);
  }
});

// ---------------------------------------------------------------------------------------------
// 7. Limits, each on its own

test('a chart part larger than is read, made of good XML, is not read and refuses the refresh', async () => {
  const filler = `<c:extLst>${'<c:ext uri="x"/>'.repeat(600000)}</c:extLst>`;
  const doc = await load(await edit(await pivotBook({ data: SMALL, rows: [0], values: [{ fld: 1 }] }), (zip) => zip.set('xl/charts/chart1.xml', encoder.encode(`<c:chartSpace xmlns:c="${CHART_NS}"><c:chart/>${filler}</c:chartSpace>`))));
  assert.match(reason(() => refreshPivot(doc.wb, 0)), /cannot tell what refers to its cache/);
});

test('the bytes of relationships and of charts looked through have budgets of their own', async () => {
  const chart = encoder.encode(`<c:chartSpace xmlns:c="${CHART_NS}"><c:chart/></c:chartSpace>`);
  const withChart = await edit(await pivotBook({ data: SMALL, rows: [0], values: [{ fld: 1 }] }), (zip) => zip.set('xl/charts/chart1.xml', chart));
  const limits = { ...PIVOT_SCAN_LIMITS };
  const rels = await load(withChart);
  PIVOT_SCAN_LIMITS.relsBytes = 100;
  try {
    assert.match(reason(() => refreshPivot(rels.wb, 0)), /cannot tell what refers to its cache/);
  } finally {
    Object.assign(PIVOT_SCAN_LIMITS, limits);
  }
  const charts = await load(withChart);
  PIVOT_SCAN_LIMITS.chartBytes = 10;
  try {
    assert.match(reason(() => refreshPivot(charts.wb, 0)), /cannot tell what refers to its cache/);
  } finally {
    Object.assign(PIVOT_SCAN_LIMITS, limits);
  }
  const plain = await load(withChart);
  assert.equal(reason(() => refreshPivot(plain.wb, 0)), '');
});

test('a target no part name can have makes what refers to the cache unknown; a link to a place does not', async () => {
  const rels = (target: string, type: string) => encoder.encode(`<?xml version="1.0"?><Relationships xmlns="${REL_NS}"><Relationship Id="r1" Type="${type}" Target="${target}"/></Relationships>`);
  for (const target of ['../pivotCache/pivotCacheDefinition1.xml#a', '../pivotCache/pivotCacheDefinition1.xml?x', '../pivotCache\\pivotCacheDefinition1.xml', '../pivotCache/pivotCacheDefinition1.xml ']) {
    const doc = await load(
      await edit(await pivotBook({ data: SMALL, rows: [0], values: [{ fld: 1 }] }), (zip) => {
        zip.set('xl/other/x.xml', encoder.encode('<x/>'));
        zip.set('xl/other/_rels/x.xml.rels', rels(target, 'x/y'));
      }),
    );
    assert.match(reason(() => refreshPivot(doc.wb, 0)), /cannot tell what refers to its cache/, target);
  }
  const link = await load(
    await edit(await pivotBook({ data: SMALL, rows: [0], values: [{ fld: 1 }] }), (zip) => {
      zip.set('xl/other/x.xml', encoder.encode('<x/>'));
      zip.set('xl/other/_rels/x.xml.rels', rels('#Data!A1', `${OFFICE_REL}/hyperlink`));
    }),
  );
  assert.equal(reason(() => refreshPivot(link.wb, 0)), '');
});
