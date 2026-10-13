// Charts of a Word document whose part caches no values: the workbook Word
// embeds beside the part (word/embeddings/*.xlsx) is opened, once and within
// limits, and the series are read from its cells. A series that has a cache
// keeps it; a workbook that cannot be read leaves the series empty, with a
// note; and nothing of the package is ever changed.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { MAX_DOC_CHART_BOOKS, MAX_DOC_CHART_BOOK_BYTES, MAX_DOC_CHART_BOOKS_BYTES } from '../src/formats/docx/charts';
import { readDocx, type DocxReadOptions } from '../src/formats/docx/document';
import { open as openFile } from '../src/index';
import { chartValues } from '../src/layout/chart';
import { chartCacheList, type Chart } from '../src/model/sheet/chart';
import type { DocModel } from '../src/model/doc/types';
import { DocHistory, insertText } from '../src/ops/doc/index';
import { OpcPackage } from '../src/opc';
import { writeDocx } from '../src/formats/docx/write';
import { ZipArchive } from '../src/zip';
import { ser, sheetData, solid, spPr } from './chart-fixtures';
import { p } from './doc-layout-fixtures';
import { chartParts, columnsChart, inlineChart, type ChartPartSpec } from './docx-chart-fixtures';
import { buildDocx as buildSpec, type DocxSpec } from './docx-fixtures';
import { buildCfb } from './doc-binary-fixtures';
import { buildXlsx, worksheetXml } from './fixtures';
import { number as xlsNumber, workbookStream, xlsFile } from './xls-fixtures';

const CHART1 = '/word/charts/chart1.xml';
const CHART2 = '/word/charts/chart2.xml';

/** The workbook of a chart as Word embeds it: categories in A2:A4, the values of two series in B2:B4 and C2:C4, their names in the first row. */
function book(values: (number | string | null)[] = [50, 100, 25], more: Partial<Parameters<typeof buildXlsx>[0]> = {}): Promise<Uint8Array> {
  const rows: (number | string | null)[][] = [[null, 'Sales', 'Costs']];
  values.forEach((value, i) => rows.push([`cat${i + 1}`, value, i + 1]));
  return buildXlsx({ sheets: [{ name: 'Sheet1', xml: worksheetXml(sheetData(rows)) }], ...more });
}

async function spec(charts: ChartPartSpec[]): Promise<DocxSpec> {
  return { body: charts.map((chart) => p(inlineChart(chart.rel, 300, 200))).join(''), ...(await chartParts(charts)) };
}

async function read(charts: ChartPartSpec[], options?: DocxReadOptions): Promise<DocModel> {
  return readDocx(await OpcPackage.open(await buildSpec(await spec(charts))), options);
}

const chartOf = (doc: DocModel, part = CHART1): Chart => {
  const chart = doc.charts ? doc.charts.get(part) : undefined;
  assert.ok(chart, `the chart ${part} was read`);
  return chart;
};
const valuesOf = (doc: DocModel, part = CHART1, series = 0) => chartCacheList(chartOf(doc, part).groups[0].series[series].values);
const bookNotes = (doc: DocModel, part = CHART1) => chartOf(doc, part).notes.filter((note) => /embedded workbook/.test(note));

test('book: a series without cached values is read from the cells of the embedded workbook', async () => {
  const doc = await read([{ rel: 'c1', xml: columnsChart({ uncached: true, nameRef: 'Sheet1!$B$1' }), workbook: await book() }]);
  const chart = chartOf(doc);
  const series = chart.groups[0].series[0];
  assert.deepEqual(valuesOf(doc), [50, 100, 25]);
  assert.equal(series.values && series.values.count, 3);
  assert.deepEqual(chartCacheList(series.categories), ['cat1', 'cat2', 'cat3']);
  assert.deepEqual(series.name, { ref: 'Sheet1!$B$1', text: 'Sales' });
  // What is drawn: the numbers and the labels of the cells.
  const drawn = chartValues(chart);
  assert.deepEqual(drawn.series[0].values.toArray(), [50, 100, 25]);
  assert.equal(drawn.series[0].name, 'Sales');
  assert.deepEqual([0, 1, 2].map((i) => drawn.label(i)), ['cat1', 'cat2', 'cat3']);
  assert.deepEqual(bookNotes(doc), ['Read from the embedded workbook of the chart, where the chart part caches nothing: values of series, categories, names of series']);
  assert.deepEqual(doc.warnings, []);
});

test('book: a chart whose series are all cached never opens its workbook', async () => {
  // The cells say something else than the cache: the cache is what Word showed.
  const cached = await read([{ rel: 'c1', xml: columnsChart({ values: [1, 2, 3] }), workbook: await book([7, 8, 9]) }]);
  assert.deepEqual(valuesOf(cached), [1, 2, 3]);
  assert.deepEqual(bookNotes(cached), []);
  // A workbook that cannot be read is not even looked at.
  const broken = await read([{ rel: 'c1', xml: columnsChart({ values: [1, 2, 3] }), workbook: new Uint8Array([1, 2, 3, 4]) }]);
  assert.deepEqual(valuesOf(broken), [1, 2, 3]);
  assert.deepEqual(bookNotes(broken), []);
  assert.deepEqual(broken.warnings, []);
});

test('book: of two series only the one without a cache is read from the cells', async () => {
  const second = ser(1, { cat: '<c:strRef><c:f>Sheet1!$A$2:$A$4</c:f></c:strRef>', val: '<c:numRef><c:f>Sheet1!$C$2:$C$4</c:f></c:numRef>', props: spPr(solid('FF0000')) });
  const xml = columnsChart({ values: [9, 9, 9], more: second });
  const doc = await read([{ rel: 'c1', xml, workbook: await book([50, 100, 25]) }]);
  assert.deepEqual(valuesOf(doc, CHART1, 0), [9, 9, 9], 'the cache of the first series stays, whatever the cells hold');
  assert.deepEqual(valuesOf(doc, CHART1, 1), [1, 2, 3], 'the second series has the cells');
});

test('book: a cell that holds no number is a point without a value; texts are no numbers', async () => {
  const doc = await read([{ rel: 'c1', xml: columnsChart({ uncached: true }), workbook: await book([5, 'text', null]) }]);
  const data = chartOf(doc).groups[0].series[0].values;
  assert.deepEqual(chartCacheList(data), [5]);
  assert.equal(data && data.count, 3, 'the series has three points');
});

test('book: a macro-enabled workbook gives its cells; nothing of it is run', async () => {
  const bytes = await book();
  const zip = ZipArchive.read(bytes);
  const types = new TextDecoder().decode(await zip.read('[Content_Types].xml'));
  zip.set('[Content_Types].xml', new TextEncoder().encode(types.replace('spreadsheetml.sheet.main+xml', 'XX').replace('application/vnd.openxmlformats-officedocument.XX', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml')));
  zip.set('xl/vbaProject.bin', new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3]));
  const doc = await read([{ rel: 'c1', xml: columnsChart({ uncached: true }), workbook: await zip.write(), embedding: 'Book1.xlsm' }]);
  assert.deepEqual(valuesOf(doc), [50, 100, 25]);
});

test('book: an embedded workbook of the old binary form (xls) gives its cells', async () => {
  const stream = workbookStream({ sheets: [{ name: 'Sheet1', records: [...xlsNumber(1, 1, 50), ...xlsNumber(2, 1, 100), ...xlsNumber(3, 1, 25)] }] });
  const doc = await read([{ rel: 'c1', xml: columnsChart({ uncached: true }), workbook: xlsFile(stream), embedding: 'Book1.xls' }]);
  assert.deepEqual(valuesOf(doc), [50, 100, 25]);
});

const xlsb = async (): Promise<Uint8Array> => {
  const zip = ZipArchive.read(await book());
  const types = new TextDecoder().decode(await zip.read('[Content_Types].xml'));
  zip.set('[Content_Types].xml', new TextEncoder().encode(types.replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', 'application/vnd.ms-excel.sheet.binary.macroEnabled.main')));
  return zip.write();
};

const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const unreadable: [name: string, chart: () => Promise<Partial<ChartPartSpec>>, note: RegExp][] = [
  ['no zip and no compound file', async () => ({ workbook: new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]) }), /not a workbook/],
  ['a zip that is no package', async () => ({}), /could not be read/],
  ['a package that is a document', async () => ({ workbook: await buildSpec({ body: p('text') }) }), /not a workbook/],
  ['a binary workbook (xlsb)', async () => ({ workbook: await xlsb(), embedding: 'Book1.xlsb' }), /binary workbook/],
  ['an encrypted workbook', async () => ({ workbook: buildCfb([['EncryptionInfo', new Uint8Array(64)], ['EncryptedPackage', new Uint8Array(5000)]]) }), /encrypted/],
  ['a compound file that is no workbook', async () => ({ workbook: buildCfb([['WordDocument', new Uint8Array(5000)]]) }), /not a workbook/],
  ['a compound file whose workbook cannot be read', async () => ({ workbook: buildCfb([['Workbook', new Uint8Array(5000).fill(0xff)]]), embedding: 'Book1.xls' }), /could not be read/],
  ['a workbook cut short', async () => ({ workbook: (await book()).subarray(0, 300) }), /could not be read/],
  ['not related to the chart', async () => ({ bookRel: '' }), /names no embedded workbook/],
  ['related to no part', async () => ({ bookRel: `<Relationship Id="rId1" Type="${NS_REL}/package" Target="../embeddings/gone.xlsx"/>` }), /names no embedded workbook|could not be read/],
  ['linked, not embedded', async () => ({ bookRel: `<Relationship Id="rId1" Type="${NS_REL}/oleObject" Target="file:///C:/book.xlsx" TargetMode="External"/>` }), /linked/],
  ['a workbook without the sheet of the series', async () => ({ workbook: await buildXlsx({ sheets: [{ name: 'Other', xml: worksheetXml(sheetData([[1]])) }] }) }), /cannot be followed/],
];
for (const [name, make, note] of unreadable) {
  test(`book: ${name} leaves the series empty, with a note; the document opens`, async () => {
    const doc = await read([{ rel: 'c1', xml: columnsChart({ uncached: true }), ...(await make()) }]);
    assert.deepEqual(valuesOf(doc), []);
    const notes = bookNotes(doc);
    assert.equal(notes.length, 1, JSON.stringify(chartOf(doc).notes));
    assert.match(notes[0], /^Shown empty, without cached values: values of series, categories \(/);
    assert.match(notes[0], note);
    assert.deepEqual(doc.warnings, []);
  });
}

test('book: the limits are stated', () => {
  assert.equal(MAX_DOC_CHART_BOOK_BYTES, 16 * 1024 * 1024);
  assert.equal(MAX_DOC_CHART_BOOKS_BYTES, 64 * 1024 * 1024);
  assert.equal(MAX_DOC_CHART_BOOKS, 64);
});

test('book: a workbook larger than the limit is not opened', async () => {
  const bytes = await book();
  const zip = ZipArchive.read(bytes);
  let size = 0;
  for (const name of zip.names()) size += (zip.entry(name) as { uncompressedSize: number }).uncompressedSize;
  const charts = [{ rel: 'c1', xml: columnsChart({ uncached: true }), workbook: bytes }];
  const fits = await read(charts, { maxChartBookBytes: size });
  assert.deepEqual(valuesOf(fits), [50, 100, 25]);
  const over = await read(charts, { maxChartBookBytes: size - 1 });
  assert.deepEqual(valuesOf(over), []);
  assert.match(bookNotes(over)[0], /too large/);
  // A limit that is asked for cannot be wider than that of the engine.
  const wide = await read(charts, { maxChartBookBytes: MAX_DOC_CHART_BOOK_BYTES * 100 });
  assert.deepEqual(valuesOf(wide), [50, 100, 25]);
});

test('book: the workbooks of a document together are bounded, in bytes and in number', async () => {
  const bytes = await book();
  const zip = ZipArchive.read(bytes);
  let size = 0;
  for (const name of zip.names()) size += (zip.entry(name) as { uncompressedSize: number }).uncompressedSize;
  const charts = [
    { rel: 'c1', xml: columnsChart({ uncached: true }), workbook: bytes, embedding: 'a.xlsx' },
    { rel: 'c2', xml: columnsChart({ uncached: true }), workbook: bytes, embedding: 'b.xlsx' },
  ];
  const both = await read(charts, { maxChartBooksBytes: size * 2 });
  assert.deepEqual([valuesOf(both, CHART1), valuesOf(both, CHART2)], [[50, 100, 25], [50, 100, 25]]);
  const one = await read(charts, { maxChartBooksBytes: size * 2 - 1 });
  assert.deepEqual([valuesOf(one, CHART1), valuesOf(one, CHART2)], [[50, 100, 25], []]);
  assert.match(bookNotes(one, CHART2)[0], /too many bytes/);
  const counted = await read(charts, { maxChartBooks: 1 });
  assert.deepEqual([valuesOf(counted, CHART1), valuesOf(counted, CHART2)], [[50, 100, 25], []]);
  assert.match(bookNotes(counted, CHART2)[0], /too many embedded workbooks/);
  const none = await read(charts, { maxChartBooks: 0 });
  assert.deepEqual([valuesOf(none, CHART1), valuesOf(none, CHART2)], [[], []]);
});

test('book: a workbook two charts share is opened once', async () => {
  const bytes = await book();
  const charts = [
    { rel: 'c1', xml: columnsChart({ uncached: true }), workbook: bytes, embedding: 'shared.xlsx' },
    { rel: 'c2', xml: columnsChart({ uncached: true }), workbook: bytes, embedding: 'shared.xlsx' },
  ];
  // One workbook is all the document may open: both charts get their cells.
  const doc = await read(charts, { maxChartBooks: 1 });
  assert.deepEqual([valuesOf(doc, CHART1), valuesOf(doc, CHART2)], [[50, 100, 25], [50, 100, 25]]);
  // Counted in a package that says how often a part is read.
  const pkg = await OpcPackage.open(await buildSpec(await spec(charts)));
  let reads = 0;
  const readPart = pkg.readPart.bind(pkg);
  pkg.readPart = (part: string) => {
    if (/shared\.xlsx$/.test(part)) reads++;
    return readPart(part);
  };
  await readDocx(pkg);
  assert.equal(reads, 1);
});

test('book: the values read from cells count against the cached values of the document', async () => {
  const charts = [{ rel: 'c1', xml: columnsChart({ uncached: true }), workbook: await book() }];
  // Three values and three categories.
  const fits = await read(charts, { maxChartPoints: 6 });
  assert.deepEqual(valuesOf(fits), [50, 100, 25]);
  const over = await read(charts, { maxChartPoints: 2 });
  assert.deepEqual(valuesOf(over), []);
  assert.match(bookNotes(over)[0], /too many/);
});

test('book: dates of the workbook count as its date system says', async () => {
  const doc = await read([{ rel: 'c1', xml: columnsChart({ uncached: true }), workbook: await book([1, 2, 3], { workbookHead: '<workbookPr date1904="1"/>' }) }]);
  assert.equal(chartOf(doc).date1904, true);
});

test('book: a series with literal values and a chart without a reference open nothing', async () => {
  const literal = columnsChart({ values: [4, 5, 6] }).replace(/<c:numRef>.*?<\/c:numRef>/, '<c:numLit><c:ptCount val="0"/></c:numLit>');
  const doc = await read([{ rel: 'c1', xml: literal, workbook: new Uint8Array([9, 9]) }]);
  assert.deepEqual(valuesOf(doc), []);
  assert.deepEqual(bookNotes(doc), [], 'literal data that is empty is no missing cache of a reference');
});

// ------------------------------------------ a cache that is there, and empty

const VALUES = '<c:numRef><c:f>Sheet1!$B$2:$B$4</c:f></c:numRef>';
const CATEGORIES = '<c:strRef><c:f>Sheet1!$A$2:$A$4</c:f></c:strRef>';
/** The chart whose series names its cells and caches nothing, with caches that are there and hold no point. */
const emptyCaches = (): string => {
  const xml = columnsChart({ uncached: true });
  assert.ok(xml.indexOf(VALUES) >= 0 && xml.indexOf(CATEGORIES) >= 0, 'the fixture is as this test takes it to be');
  return xml
    .replace(VALUES, VALUES.replace('</c:numRef>', '<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="0"/></c:numCache></c:numRef>'))
    .replace(CATEGORIES, CATEGORIES.replace('</c:strRef>', '<c:strCache><c:ptCount val="0"/></c:strCache></c:strRef>'));
};
/** A chart whose values are cached (9, 9, 9) and whose categories are not. */
const cachedValuesOnly = (spec: Parameters<typeof columnsChart>[0] = {}): string => {
  const xml = columnsChart({ values: [9, 9, 9], ...spec });
  const cache = /<c:cat><c:strRef><c:f>[^<]*<\/c:f>(<c:strCache>.*?<\/c:strCache>)/.exec(xml);
  assert.ok(cache, 'the fixture caches its categories');
  return xml.replace(cache[1], '');
};

test('book: a cache that is there and holds no point is what is drawn: the workbook is not asked', async () => {
  // The cells hold 50, 100 and 25; the part says, of the same cells, that there is nothing.
  const pkg = await OpcPackage.open(await buildSpec(await spec([{ rel: 'c1', xml: emptyCaches(), workbook: await book() }])));
  let reads = 0;
  const readPart = pkg.readPart.bind(pkg);
  pkg.readPart = (part: string) => {
    if (/embeddings/.test(part)) reads++;
    return readPart(part);
  };
  const doc = await readDocx(pkg);
  assert.deepEqual(valuesOf(doc), []);
  assert.deepEqual(chartCacheList(chartOf(doc).groups[0].series[0].categories), []);
  assert.equal(reads, 0, 'the embedded workbook was not read');
  assert.deepEqual(chartOf(doc).notes.filter((note) => /cached|embedded/.test(note)), []);
  // Without the cache elements the same cells are read.
  const without = await read([{ rel: 'c1', xml: columnsChart({ uncached: true }), workbook: await book() }]);
  assert.deepEqual(valuesOf(without), [50, 100, 25]);
});

test('book: the name of a series and a title whose caches are there and empty are left as they are', async () => {
  const tx = '<c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="0"/></c:strCache></c:strRef>';
  const head = '<c:title><c:tx><c:strRef><c:f>Sheet1!$C$1</c:f><c:strCache><c:ptCount val="0"/></c:strCache></c:strRef></c:tx></c:title>';
  const xml = columnsChart({ uncached: true, head }).replace('<c:order val="0"/>', `<c:order val="0"/><c:tx>${tx}</c:tx>`);
  const doc = await read([{ rel: 'c1', xml, workbook: await book() }]);
  const chart = chartOf(doc);
  assert.deepEqual(valuesOf(doc), [50, 100, 25]);
  assert.deepEqual(chart.groups[0].series[0].name, { ref: 'Sheet1!$B$1' });
  assert.equal(chart.title && chart.title.text, undefined);
  assert.deepEqual(bookNotes(doc), ['Read from the embedded workbook of the chart, where the chart part caches nothing: values of series, categories']);
});

// ----------------------------------------------- what came from where is said

test('book: categories without a cache and a workbook that cannot be read are said, though the values are cached', async () => {
  const doc = await read([{ rel: 'c1', xml: cachedValuesOnly(), workbook: new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]) }]);
  assert.deepEqual(valuesOf(doc), [9, 9, 9]);
  assert.deepEqual(chartCacheList(chartOf(doc).groups[0].series[0].categories), []);
  assert.deepEqual(bookNotes(doc), ['Shown empty, without cached values: categories (what stands for the embedded workbook of the chart is not a workbook)']);
  assert.deepEqual(doc.warnings, []);
});

test('book: a name of a series without a cache and a workbook that cannot be read are said', async () => {
  const xml = columnsChart({ values: [9, 9, 9] }).replace('<c:order val="0"/>', '<c:order val="0"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>');
  const doc = await read([{ rel: 'c1', xml, bookRel: '' }]);
  assert.deepEqual(chartOf(doc).groups[0].series[0].name, { ref: 'Sheet1!$B$1' });
  assert.deepEqual(bookNotes(doc), ['Shown empty, without cached values: names of series (the chart names no embedded workbook)']);
});

test('book: what the note says was read from the workbook is what was read from it', async () => {
  const xml = cachedValuesOnly().replace('<c:order val="0"/>', '<c:order val="0"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>');
  const doc = await read([{ rel: 'c1', xml, workbook: await book() }]);
  const series = chartOf(doc).groups[0].series[0];
  assert.deepEqual(valuesOf(doc), [9, 9, 9], 'the values are those of the cache');
  assert.deepEqual(chartCacheList(series.categories), ['cat1', 'cat2', 'cat3']);
  assert.deepEqual(series.name, { ref: 'Sheet1!$B$1', text: 'Sales' });
  assert.deepEqual(bookNotes(doc), ['Read from the embedded workbook of the chart, where the chart part caches nothing: categories, names of series']);
});

test('book: a title without a cache is read from its cell; the note names it', async () => {
  const head = '<c:title><c:tx><c:strRef><c:f>Sheet1!$C$1</c:f></c:strRef></c:tx></c:title>';
  const doc = await read([{ rel: 'c1', xml: columnsChart({ values: [9, 9, 9], head }), workbook: await book() }]);
  assert.deepEqual(chartOf(doc).title && [chartOf(doc).title!.ref, chartOf(doc).title!.text], ['Sheet1!$C$1', 'Costs']);
  assert.deepEqual(bookNotes(doc), ['Read from the embedded workbook of the chart, where the chart part caches nothing: titles']);
});

test('book: a name and a title the part caches are not replaced by what the cells hold', async () => {
  const cached = (text: string) => `<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${text}</c:v></c:pt></c:strCache>`;
  const head = `<c:title><c:tx><c:strRef><c:f>Sheet1!$C$1</c:f>${cached('Cached title')}</c:strRef></c:tx></c:title>`;
  const xml = columnsChart({ uncached: true, head }).replace('<c:order val="0"/>', `<c:order val="0"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f>${cached('Cached name')}</c:strRef></c:tx>`);
  const doc = await read([{ rel: 'c1', xml, workbook: await book() }]);
  const chart = chartOf(doc);
  assert.deepEqual(valuesOf(doc), [50, 100, 25], 'the workbook was read for the values');
  assert.equal(chart.groups[0].series[0].name!.text, 'Cached name');
  assert.equal(chart.title!.text, 'Cached title');
});

// ---------------------------------------------------------- more of the limits

test('book: an embedded workbook of the old binary form counts by its bytes', async () => {
  const stream = workbookStream({ sheets: [{ name: 'Sheet1', records: [...xlsNumber(1, 1, 50), ...xlsNumber(2, 1, 100), ...xlsNumber(3, 1, 25)] }] });
  const bytes = xlsFile(stream);
  const charts = [{ rel: 'c1', xml: columnsChart({ uncached: true }), workbook: bytes, embedding: 'Book1.xls' }];
  assert.deepEqual(valuesOf(await read(charts, { maxChartBookBytes: bytes.length, maxChartBooksBytes: bytes.length })), [50, 100, 25]);
  const large = await read(charts, { maxChartBookBytes: bytes.length - 1 });
  assert.deepEqual(valuesOf(large), []);
  assert.match(bookNotes(large)[0], /too large/);
  const many = await read(charts, { maxChartBooksBytes: bytes.length - 1 });
  assert.deepEqual(valuesOf(many), []);
  assert.match(bookNotes(many)[0], /too many bytes/);
});

test('book: values that fit what is left of the points of the document are read, the categories after them are not', async () => {
  const charts = [{ rel: 'c1', xml: columnsChart({ uncached: true }), workbook: await book() }];
  // Three values fit five points; the three categories after them do not fit the two that are left.
  const doc = await read(charts, { maxChartPoints: 5 });
  assert.deepEqual(valuesOf(doc), [50, 100, 25]);
  assert.deepEqual(chartCacheList(chartOf(doc).groups[0].series[0].categories), []);
  assert.deepEqual(bookNotes(doc), [
    'Read from the embedded workbook of the chart, where the chart part caches nothing: values of series',
    'Shown empty, without cached values: categories (the charts of the document hold too many values; the embedded workbook of the chart is not read for it)',
  ]);
  // Exactly what is needed is enough.
  assert.deepEqual(chartCacheList(chartOf(await read(charts, { maxChartPoints: 6 })).groups[0].series[0].categories), ['cat1', 'cat2', 'cat3']);
});

test('book: categories of several levels are left to their cache', async () => {
  // UNSURE: which column of the block is which level is not certain, so the block is not read from the cells.
  const xml = cachedValuesOnly().replace(/<c:cat><c:strRef>(<c:f>[^<]*<\/c:f>)<\/c:strRef><\/c:cat>/, '<c:cat><c:multiLvlStrRef>$1</c:multiLvlStrRef></c:cat>');
  assert.ok(xml.indexOf('multiLvlStrRef') >= 0, 'the fixture has categories of several levels');
  const doc = await read([{ rel: 'c1', xml, workbook: await book() }]);
  const categories = chartOf(doc).groups[0].series[0].categories;
  assert.deepEqual([categories && categories.multiLevel, chartCacheList(categories)], [true, []]);
  assert.deepEqual(bookNotes(doc), []);
});

test('book: numbers read from cells are shown in the number format of their cells', async () => {
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const styles =
    `<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="${NS}"><numFmts count="1"><numFmt numFmtId="164" formatCode="0.00&quot; kg&quot;"/></numFmts>` +
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>';
  const rows = [50, 100, 25].map((value, i) => `<row r="${i + 2}"><c r="B${i + 2}" s="1"><v>${value}</v></c></row>`).join('');
  const workbook = await buildXlsx({ sheets: [{ name: 'Sheet1', xml: worksheetXml(`<sheetData>${rows}</sheetData>`) }], styles });
  const doc = await read([{ rel: 'c1', xml: columnsChart({ uncached: true }), workbook }]);
  const data = chartOf(doc).groups[0].series[0].values;
  assert.deepEqual([chartCacheList(data), data && data.formatCode], [[50, 100, 25], '0.00" kg"']);
});

// --------------------------------------------------------------- lossless

test('lossless: a document whose chart was read from its workbook is saved as the file it was', async () => {
  const bytes = await buildSpec(await spec([{ rel: 'c1', xml: columnsChart({ uncached: true }), workbook: await book() }]));
  const file = await openFile(bytes);
  const doc = await file.document();
  assert.deepEqual(valuesOf(doc), [50, 100, 25]);
  assert.deepEqual(await file.save(), bytes);
});

test('lossless: after an edit the chart part and the embedded workbook are the entries of the file, byte for byte', async () => {
  const workbook = await book();
  const bytes = await buildSpec({ body: p('text') + p(inlineChart('c1', 300, 200)), ...(await chartParts([{ rel: 'c1', xml: columnsChart({ uncached: true }), workbook }])) });
  const pkg = await OpcPackage.open(bytes);
  const doc = await readDocx(pkg);
  assert.deepEqual(valuesOf(doc), [50, 100, 25]);
  const history = new DocHistory(doc);
  history.applyEdit(insertText(doc, { story: 'body', path: [0], offset: 4 }, ' added'), undefined);
  await writeDocx(pkg, doc);
  const saved = await pkg.save();
  const before = ZipArchive.read(bytes);
  const after = ZipArchive.read(saved);
  for (const name of ['word/charts/chart1.xml', 'word/charts/_rels/chart1.xml.rels', 'word/embeddings/Microsoft_Excel_Sheet1.xlsx']) {
    assert.deepEqual((after.entry(name) as { raw: Uint8Array }).raw, (before.entry(name) as { raw: Uint8Array }).raw, name);
  }
  assert.deepEqual(await after.read('word/embeddings/Microsoft_Excel_Sheet1.xlsx'), workbook, 'the workbook is the one that was embedded');
  // The chart part still caches nothing: what was read from the cells is not written into it.
  assert.ok(new TextDecoder().decode(await after.read('word/charts/chart1.xml')).indexOf('numCache') < 0, 'no cache was written into the chart part');
});
