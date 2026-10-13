// Charts of an xls workbook that are not what the reader expects: records cut
// short, blocks that do not close, series that name nothing, too much of
// everything, and thousands of damaged substreams. Whatever a chart holds, the
// reader does not throw and does not hang, the chart stays a placeholder with
// a warning, and the rest of the workbook is read as before.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readXlsChart, XLS_MAX_CHART_DEPTH, XLS_MAX_CHART_RECORDS, type XlsChartContext } from '../src/formats/xls/chart';
import { recordStream } from '../src/formats/xls/records';
import { readXlsStream } from '../src/formats/xls/workbook';
import { chartValues, layoutChart } from '../src/layout/chart';
import { MAX_CHART_SERIES, type Chart } from '../src/model/sheet/chart';
import { getCell, type Sheet, type StyleSheet } from '../src/model/sheet/types';
import { fakeMeasure } from './chart-fixtures';
import { drawingRecords, shape } from './xls-art-fixtures';
import {
  ai,
  areaFormat,
  attachedLabel,
  axisParentRec,
  axisRec,
  bar,
  begin,
  block,
  cachedLabel,
  cachedNumber,
  catSerRange,
  chartFormatRec,
  chartRec,
  chartRecords,
  chartSubstream,
  dataFormat,
  defaultAxes,
  end,
  legendRec,
  line,
  lineFormat,
  markerFormat,
  objectLink,
  pie,
  REC,
  scatter,
  seriesRec,
  serToCrt,
  siIndex,
  textBlock,
  textRec,
  valueRange,
  type XlsChartSpec,
  type XlsSeriesSpec,
} from './xls-chart-fixtures';
import { bof, cat, eof, externSheet, number, ptg, rec, supBookSelf, u16, workbookStream, type Bytes } from './xls-fixtures';

const tail = (): Uint8Array => cat(supBookSelf(1), externSheet([0, 0, 0]));
const column = (col: number): number[] => ptg.area3d(0, 1, 3, col, col);
const sales: XlsSeriesSpec = { name: 'Sales', values: column(1), categories: column(0) };
const cells = (): Uint8Array => cat(...[50, 100, 25].map((value, i) => number(i + 1, 1, value)));

/** A sheet with cells before and after a chart whose substream is `substream`. */
function sheetWith(substream: Bytes[]): Sheet {
  const drawing = drawingRecords([shape({ kind: 201, ot: 5, id: 1, extra: substream.map((part) => (part instanceof Uint8Array ? part : new Uint8Array(part))) })]);
  return readXlsStream(workbookStream({ tail: tail(), sheets: [{ name: 'Data', records: cat(cells(), drawing, number(9, 0, 9)) }] })).workbook.sheets[0];
}

/** The warning of a sheet whose chart was not read; the cells around the chart are read all the same. */
function problem(substream: Bytes[]): string {
  const sheet = sheetWith(substream);
  assert.deepEqual(sheet.drawings && sheet.drawings.map((drawing) => drawing.object.type), ['placeholder']);
  assert.equal(sheet.warnings && sheet.warnings.length, 1, JSON.stringify(sheet.warnings));
  assert.equal(getCell(sheet, 1, 1)?.value, 50, 'the cells before the chart are read');
  assert.equal(getCell(sheet, 9, 0)?.value, 9, 'the cells after the chart are read');
  return (sheet.warnings as string[])[0];
}

const chartBof = (): Uint8Array => bof(0x0020);
/** A chart substream around the records that stand inside the Chart block. */
const around = (...inside: Bytes[]): Bytes[] => [chartBof(), block(chartRec(), ...inside), eof()];
const oneSeries = (...children: Bytes[]): Uint8Array => block(seriesRec(3, 3), ai(1, 2, column(1)), ...children);
const axesWith = (...groups: Bytes[]): Uint8Array => block(axisParentRec(0), ...defaultAxes(), ...groups);
const barGroup = (order = 0): Uint8Array => block(chartFormatRec(order), bar());

function chartIn(substream: Bytes[]): Chart {
  const sheet = sheetWith(substream);
  const object = sheet.drawings && sheet.drawings[0] ? sheet.drawings[0].object : undefined;
  assert.ok(object && object.type === 'chart', `the chart was read: ${JSON.stringify(sheet.warnings)}`);
  return object.chart;
}

// -------------------------------------------------- what is not what it should be

test('xls chart: the smallest chart that is read: one series, one chart group', () => {
  const chart = chartIn(around(oneSeries(serToCrt(0)), axesWith(barGroup())));
  assert.equal(chart.groups[0].series[0].values!.ref, 'Data!$B$2:$B$4');
  // Without a SerToCrt record the one chart group there is has the series.
  assert.equal(chartIn(around(oneSeries(), axesWith(barGroup()))).groups[0].series.length, 1);
});

const unexpected: [name: string, substream: () => Bytes[], warning: RegExp][] = [
  ['no records after its object', () => [], /no records of a chart follow/],
  ['a Chart record without a block', () => [chartBof(), chartRec(), eof()], /Chart record holds nothing/],
  ['two Chart records', () => [chartBof(), block(chartRec(), oneSeries(), axesWith(barGroup())), block(chartRec()), eof()], /two Chart records/],
  ['a block that is not closed', () => [chartBof(), chartRec(), begin(), oneSeries(), axesWith(barGroup()), eof()], /not closed/],
  ['an End too many', () => [chartBof(), block(chartRec(), oneSeries(), axesWith(barGroup())), end(), eof()], /belongs to no record/],
  ['a Begin after a Begin', () => around(oneSeries(), begin(), begin(), end(), end(), axesWith(barGroup())), /belongs to no record/],
  ['blocks nested too deeply', () => around(oneSeries(), axesWith(barGroup()), ...new Array(XLS_MAX_CHART_DEPTH + 1).fill(cat(rec(0x10f0), begin())), ...new Array(XLS_MAX_CHART_DEPTH + 1).fill(end())), /nested too deeply/],
  ['a Chart record cut short', () => [chartBof(), block(rec(REC.chart, [0, 0, 0, 0]), oneSeries(), axesWith(barGroup())), eof()], /Chart record is too short/],
  ['no chart group', () => around(oneSeries(), block(axisParentRec(0), ...defaultAxes())), /no chart group/],
  ['no axis group', () => around(oneSeries()), /no chart group/],
  ['no series', () => around(axesWith(barGroup())), /no series/],
  ['only a trend line for a series', () => around(oneSeries(rec(REC.serParent, u16(1))), axesWith(barGroup())), /no series/],
  ['a chart group without a chart type', () => around(oneSeries(), axesWith(block(chartFormatRec(0)))), /no chart type/],
  ['a chart group with two chart types', () => around(oneSeries(), axesWith(block(chartFormatRec(0), bar(), line()))), /two chart types/],
  ['a series that names a chart group that is not there', () => around(oneSeries(serToCrt(3)), axesWith(barGroup())), /names a chart group/],
  ['a series that names no chart group where there are two', () => around(oneSeries(), axesWith(barGroup(0), block(chartFormatRec(1), line()))), /names a chart group/],
  ['a Series record cut short', () => around(block(rec(REC.series, u16(3), u16(1)), ai(1, 2, column(1))), axesWith(barGroup())), /Series record is too short/],
  ['a SerToCrt record cut short', () => around(oneSeries(rec(REC.serToCrt, [0])), axesWith(barGroup())), /SerToCrt record is too short/],
  ['a ChartFormat record cut short', () => around(oneSeries(), axesWith(block(rec(REC.chartFormat, new Array(8).fill(0)), bar()))), /ChartFormat record is too short/],
  ['a Bar record cut short', () => around(oneSeries(), axesWith(block(chartFormatRec(0), rec(REC.bar, u16(0))))), /Bar record is too short/],
  ['a Pie record cut short', () => around(oneSeries(), axesWith(block(chartFormatRec(0), rec(REC.pie, u16(0))))), /Pie record is too short/],
  ['an AxisParent record cut short', () => around(oneSeries(), block(rec(REC.axisParent, [0]), barGroup())), /AxisParent record is too short/],
  ['a third axis group', () => around(oneSeries(), axesWith(barGroup()), block(axisParentRec(1), block(chartFormatRec(1), line())), block(axisParentRec(2), block(chartFormatRec(2), line()))), /axis groups/],
  ['two primary axis groups', () => around(oneSeries(), axesWith(barGroup()), axesWith(barGroup(1))), /axis groups/],
  ['two value axes in one axis group', () => around(oneSeries(), block(axisParentRec(0), block(axisRec(1), valueRange()), block(axisRec(1), valueRange()), barGroup())), /two axes of one kind/],
  ['an Axis record cut short', () => around(oneSeries(), block(axisParentRec(0), block(rec(REC.axis, [0])), barGroup())), /Axis record is too short/],
  ['a ValueRange record cut short', () => around(oneSeries(), block(axisParentRec(0), block(axisRec(1), rec(REC.valueRange, new Array(20).fill(0))), barGroup())), /ValueRange record is too short/],
  ['a CatSerRange record cut short', () => around(oneSeries(), block(axisParentRec(0), block(axisRec(0), rec(REC.catSerRange, u16(1))), barGroup())), /CatSerRange record is too short/],
  ['a LineFormat record cut short', () => around(oneSeries(block(dataFormat(), rec(REC.lineFormat, [0, 0, 0, 0]))), axesWith(barGroup())), /LineFormat record is too short/],
  ['an AreaFormat record cut short', () => around(oneSeries(block(dataFormat(), rec(REC.areaFormat, [0, 0, 0, 0]))), axesWith(barGroup())), /AreaFormat record is too short/],
  ['a MarkerFormat record cut short', () => around(oneSeries(block(dataFormat(), rec(REC.markerFormat, [0, 0, 0, 0]))), axesWith(barGroup())), /MarkerFormat record is too short/],
  ['a DataFormat record cut short', () => around(oneSeries(block(rec(REC.dataFormat, u16(0xffff)))), axesWith(barGroup())), /DataFormat record is too short/],
  ['an AI record cut short', () => around(block(seriesRec(3, 3), rec(REC.ai, [1, 2, 0, 0])), axesWith(barGroup())), /AI record is too short/],
  ['an AI record whose formula is longer than the record', () => around(block(seriesRec(3, 3), rec(REC.ai, [1, 2], u16(0), u16(0), u16(200), column(1))), axesWith(barGroup())), /runs past its record/],
  ['a Legend record cut short', () => around(oneSeries(), axesWith(block(chartFormatRec(0), bar(), rec(REC.legend, [0, 0, 0, 0])))), /Legend record is too short/],
  ['a title cut short', () => around(oneSeries(), axesWith(barGroup()), block(rec(REC.text, [2, 2, 1, 0]), objectLink(1))), /Text record is too short/],
  ['a cached number cut short', () => [chartBof(), block(chartRec(), oneSeries(), axesWith(barGroup())), siIndex(1), rec(0x0203, u16(0), u16(0), u16(0), [1, 2]), eof()], /cached number is too short/],
];
for (const [name, substream, warning] of unexpected) {
  test(`xls chart: ${name} leaves a placeholder and a warning; the sheet is read`, () => {
    assert.match(problem(substream()), warning);
  });
}

test('xls chart: a reference that is no reference of cells leaves the series to its cache, with a note', () => {
  // A formula: 1+2. And tokens that are no formula at all.
  for (const tokens of [[...ptg.int(1), ...ptg.int(2), ...ptg.add], [0xff, 0xff, 0xff]]) {
    const chart = chartIn([chartBof(), block(chartRec(), block(seriesRec(3, 3), ai(1, 2, tokens)), axesWith(barGroup())), siIndex(1), cachedNumber(0, 0, 4), eof()]);
    const data = chart.groups[0].series[0].values!;
    assert.deepEqual(data.cache, { index: [0], values: [4] });
    if (data.ref === undefined) assert.ok(chart.notes.some((note) => /reference of the chart cannot be read/.test(note)), JSON.stringify(chart.notes));
    // What is no reference of cells is not followed: the cache is drawn.
    const workbook = readXlsStream(workbookStream({ tail: tail(), sheets: [{ name: 'Data', records: cells() }] })).workbook;
    assert.deepEqual(chartValues(chart, { workbook }).series[0].values.toArray(), [4, null, null]);
  }
});

test('xls chart: a reference to a sheet that was deleted is not followed', () => {
  const chart = chartIn([chartBof(), block(chartRec(), block(seriesRec(3, 3), ai(1, 2, ptg.refErr3d(0))), axesWith(barGroup())), eof()]);
  const workbook = readXlsStream(workbookStream({ tail: tail(), sheets: [{ name: 'Data', records: cells() }] })).workbook;
  assert.deepEqual(chartValues(chart, { workbook }).series[0].values.toArray(), [null, null, null]);
});

test('xls chart: records that are shorter than Excel writes them but long enough for what is read are read', () => {
  // A Tick record of three bytes, a Text of 26 bytes, a Scatter record without its fields.
  const chart = chartIn(
    around(oneSeries(), block(axisParentRec(0), block(axisRec(0), valueRange(), rec(REC.tick, [0, 0, 1])), block(axisRec(1), valueRange()), block(chartFormatRec(0), rec(REC.scatter))), block(rec(REC.text, new Array(26).fill(0)), ai(0, 1), rec(REC.seriesText, u16(0), [2, 0, 72, 105]), objectLink(1))),
  );
  assert.equal(chart.groups[0].kind, 'scatter');
  assert.equal(chart.axes[0].tickLabels, 'low');
  assert.deepEqual(chart.title, { text: 'Hi', style: { color: '#000000' } });
});

test('xls chart: the words of a title that run past their record end where the record does; wide characters are read', () => {
  const title = (text: Bytes): Chart => chartIn(around(oneSeries(), axesWith(barGroup()), block(textRec(), ai(0, 1), rec(REC.seriesText, u16(0), text), objectLink(1))));
  assert.equal(title([200, 0, 65, 66]).title!.text, 'AB');
  assert.equal(title([2, 1, 0x1e, 0x01, 0x5f, 0x01]).title!.text, 'Ğş');
  assert.equal(title([0]).title, undefined, 'a title without words is none');
});

// ------------------------------------------------------------------- budgets

function context(counts: { charts: number; points: number }, styles?: StyleSheet): XlsChartContext {
  // The same object goes on being counted down: the tests look at it afterwards.
  const budget = counts as XlsChartContext['budget'];
  if (budget.bytes === undefined) budget.bytes = 1 << 30;
  return {
    styles: styles || readXlsStream(workbookStream({ sheets: [{ name: 'Data' }] })).workbook.styles,
    fontOf: () => undefined,
    formula: { biff8: true, codepage: 1252, sheetNames: ['Data'], names: [], xti: [{ book: 0, first: 0, last: 0 }], supBooks: [{ kind: 'self', sheetNames: ['Data'], externNames: [], ordinal: 0 }], failures: new Map() },
    date1904: false,
    budget,
  };
}

test('xls chart: the charts of a workbook together are bounded, in number and in cached values', () => {
  const cache = [siIndex(1), cachedNumber(0, 0, 1), cachedNumber(1, 0, 2), cachedNumber(2, 0, 3)];
  const stream = recordStream(chartSubstream({ series: [sales], cache }));
  const budget = { charts: 2, points: 5 };
  const first = readXlsChart(stream, 0, context(budget));
  assert.ok(first.chart, first.problem);
  assert.deepEqual([budget.charts, budget.points], [1, 2]);
  // Three values do not fit the two that are left: the chart is not half a chart.
  assert.deepEqual(readXlsChart(stream, 0, context(budget)), { problem: 'the charts of the workbook hold too many cached values' });
  assert.equal(budget.charts, 0);
  assert.deepEqual(readXlsChart(stream, 0, context(budget)), { problem: 'the workbook holds too many charts' });
});

test('xls chart: what is no chart substream at the offset is not read', () => {
  const stream = recordStream(cat(bof(0x0010), eof(), chartSubstream({ series: [sales] })));
  for (const start of [0, 3, -1, 1e9, stream.bytes.length - 2]) assert.deepEqual(readXlsChart(stream, start, context({ charts: 9, points: 9 })), { problem: 'it does not start as a chart does' });
});

test('xls chart: series beyond the limit of a chart are left out, with a note', () => {
  const series: XlsSeriesSpec[] = [];
  for (let i = 0; i < MAX_CHART_SERIES + 5; i++) series.push({ name: `S${i}`, values: column(1) });
  const read = readXlsChart(recordStream(chartSubstream({ series })), 0, context({ charts: 1, points: 0 }));
  assert.ok(read.chart, read.problem);
  assert.equal(read.chart.groups[0].series.length, MAX_CHART_SERIES);
  assert.ok(read.chart.notes.some((note) => new RegExp(`first ${MAX_CHART_SERIES} series`).test(note)), JSON.stringify(read.chart.notes));
});

test('xls chart: a chart of more records than the limit is not read; many cached values are read quickly', () => {
  const filler = new Uint8Array(4 * (XLS_MAX_CHART_RECORDS + 1));
  for (let p = 0; p < filler.length; p += 4) filler.set([0x22, 0x10, 0, 0], p);
  const much = recordStream(cat(chartBof(), block(chartRec(), oneSeries(), axesWith(barGroup()), filler), eof()));
  assert.deepEqual(readXlsChart(much, 0, context({ charts: 1, points: 0 })), { problem: 'it holds too many records' });

  const values: Uint8Array[] = [siIndex(1)];
  for (let point = 0; point < 60000; point++) values.push(cachedNumber(point, 0, point));
  const started = Date.now();
  const budget = { charts: 1, points: 100000 };
  const read = readXlsChart(recordStream(cat(chartBof(), block(chartRec(), block(seriesRec(0, 60000), ai(1, 1)), axesWith(barGroup())), ...values, eof())), 0, context(budget));
  assert.ok(read.chart, read.problem);
  assert.equal(read.chart.groups[0].series[0].values!.cache.index.length, 60000);
  assert.equal(budget.points, 40000);
  assert.ok(Date.now() - started < 5000, `sixty thousand cached values in ${Date.now() - started} ms`);
});

test('xls chart: cached values of points and series beyond the limits, of no kind, and blanks are left out', () => {
  const read = readXlsChart(
    recordStream(
      cat(
        chartBof(),
        // Before the Chart record nothing is a cache.
        siIndex(1),
        cachedNumber(0, 0, 99),
        block(chartRec(), oneSeries(), axesWith(barGroup())),
        // No SIIndex yet: not known what these are of.
        cachedNumber(0, 0, 98),
        siIndex(7),
        cachedNumber(0, 0, 97),
        siIndex(1),
        cachedNumber(2, 0, 3),
        cachedNumber(0, 0, 1),
        cachedNumber(0, 0, 2),
        cachedNumber(0xffff, 0, 5),
        cachedNumber(0, 300, 5),
        rec(0x0201, u16(1), u16(0), u16(0)),
        rec(0x0205, u16(1), u16(0), u16(0), [1, 0]),
        rec(0x0203, u16(1), u16(0), u16(0), [0, 0, 0, 0, 0, 0, 0xf0, 0x7f]),
        siIndex(2),
        cachedLabel(1, 0, 'b'),
        cachedNumber(0, 0, 2020),
        eof(),
      ),
    ),
    0,
    context({ charts: 1, points: 100 }),
  );
  assert.ok(read.chart, read.problem);
  const series = read.chart.groups[0].series[0];
  // The value stated last for a point is its value; the points come out in their order.
  assert.deepEqual(series.values!.cache, { index: [0, 2, 65535], values: [2, 3, 5] });
  // Texts among the categories make them texts.
  assert.deepEqual([series.categories!.kind, series.categories!.cache], ['str', { index: [0, 1], values: ['2020', 'b'] }]);
});

// ----------------------------------------------------------- nothing changes

test('xls chart: a workbook without charts is read as before', () => {
  // A substream inside a sheet that is no chart is passed over, without a word.
  const sheet = readXlsStream(workbookStream({ sheets: [{ name: 'S', records: cat(number(0, 0, 1), bof(0x0010), number(5, 5, 5), eof(), number(1, 0, 2)) }] })).workbook.sheets[0];
  assert.deepEqual([getCell(sheet, 0, 0)?.value, getCell(sheet, 1, 0)?.value, getCell(sheet, 5, 5)], [1, 2, undefined]);
  assert.deepEqual([sheet.drawings, sheet.warnings], [undefined, undefined]);
  // A chart substream that follows no chart object is no chart of the sheet.
  const stray = readXlsStream(workbookStream({ tail: tail(), sheets: [{ name: 'Data', records: cat(cells(), chartSubstream({ series: [sales] })) }] })).workbook.sheets[0];
  assert.deepEqual([stray.drawings, stray.warnings], [undefined, undefined]);
});

test('xls chart: a chart sheet of the older format (BIFF5) stays empty, as before', () => {
  const { workbook } = readXlsStream(workbookStream({ version: 0x0500, sheets: [{ name: 'Data', records: number(0, 0, 1) }, { name: 'Chart1', type: 2, records: chartRecords({ series: [sales] }) }] }), 'Book');
  assert.equal(workbook.sheets[1].kind, 'chartsheet');
  assert.deepEqual([workbook.sheets[1].drawings, workbook.sheets[1].warnings], [undefined, undefined]);
});

test('xls chart: of many charts that are not read twenty are said, and the others counted', () => {
  const shapes = [];
  for (let i = 0; i < 25; i++) shapes.push(shape({ kind: 201, ot: 5, id: i + 1, extra: [chartBof(), eof()] }));
  const sheet = readXlsStream(workbookStream({ sheets: [{ name: 'Data', records: drawingRecords(shapes) }] })).workbook.sheets[0];
  assert.equal(sheet.drawings && sheet.drawings.length, 25);
  assert.equal(sheet.warnings && sheet.warnings.length, 21);
  assert.equal((sheet.warnings as string[])[20], 'More charts that are not shown: 5');
});

test('xls chart: two charts of one sheet are each the chart of their own object', () => {
  const drawing = drawingRecords([
    shape({ kind: 201, ot: 5, id: 1, extra: [chartSubstream({ series: [sales], groups: [{ type: line() }] })] }),
    shape({ kind: 1, ot: 2, id: 2 }),
    shape({ kind: 201, ot: 5, id: 3, extra: [chartSubstream({ series: [sales], axes: [], groups: [{ type: pie() }] })] }),
  ]);
  const sheet = readXlsStream(workbookStream({ tail: tail(), sheets: [{ name: 'Data', records: cat(cells(), drawing) }] })).workbook.sheets[0];
  assert.deepEqual(
    (sheet.drawings || []).map((item) => (item.object.type === 'chart' ? item.object.chart.groups[0].kind : item.object.type)),
    ['line', 'shape', 'pie'],
  );
});

// ---------------------------------------------------------------------- fuzz

/** A generator of numbers that gives the same run every time. */
function random(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const rich = (): XlsChartSpec[] => [
  {
    series: [
      { ...sales, format: [lineFormat({ color: 'FF0000' }), areaFormat({ color: '00FF00' }), markerFormat({ type: 3 }), attachedLabel(1)], points: [[1, [areaFormat({ color: '0000FF' })]]] },
      { name: ptg.ref3d(0, 0, 2), values: column(2), categories: column(0), group: 1 },
    ],
    groups: [{ type: bar(2, 50, -100), children: [block(legendRec(3), block(textRec()))] }, { type: line(1) }],
    tail: [textBlock(objectLink(1), 'Title', { color: '112233' }), textBlock(objectLink(4, 0, 0xffff), undefined, { flags: 4 })],
    cache: [siIndex(1), cachedNumber(0, 0, 1), cachedNumber(1, 0, 2), cachedNumber(0, 1, 3), siIndex(2), cachedLabel(0, 0, 'a'), cachedLabel(1, 0, 'b')],
  },
  { series: [{ ...sales, format: [areaFormat({ color: '00FF00' })] }], axes: [], groups: [{ type: pie(45, 30), varied: true }] },
  {
    series: [{ ...sales, numericCategories: true, sizes: column(3) }],
    axes: [block(axisRec(0), valueRange({ min: 0, max: 10, log: true }), catSerRange()), block(axisRec(1), valueRange({ min: -5, max: 5, major: 1, reversed: true }))],
    groups: [{ type: scatter(1) }],
  },
];

test('xls chart: thousands of damaged chart substreams neither throw nor hang; what is read can be drawn', () => {
  const next = random(20261013);
  const bases = rich().map((spec) => chartSubstream(spec));
  const book = readXlsStream(workbookStream({ tail: tail(), sheets: [{ name: 'Data', records: cells() }] })).workbook;
  const pick = (count: number): number => Math.floor(next() * count);
  let read = 0;
  let refused = 0;
  const started = Date.now();
  for (let round = 0; round < 6000; round++) {
    const base = bases[round % bases.length];
    let bytes = new Uint8Array(base);
    const changes = 1 + pick(4);
    for (let change = 0; change < changes; change++) {
      const kind = pick(6);
      // The header of the substream stays: what follows it is what is damaged.
      const at = 20 + pick(bytes.length - 20);
      if (kind === 0) bytes[at] = pick(256);
      else if (kind === 1) bytes[at] ^= 1 << pick(8);
      else if (kind === 2) bytes = bytes.slice(0, at);
      else if (kind === 3) bytes = cat(bytes.subarray(0, at), bytes.subarray(Math.min(at + 1 + pick(40), bytes.length)));
      else if (kind === 4) bytes = cat(bytes.subarray(0, at), bytes.subarray(at, Math.min(at + 1 + pick(60), bytes.length)), bytes.subarray(at));
      else bytes.fill(pick(2) ? 0xff : 0, at, Math.min(at + 1 + pick(8), bytes.length));
      if (bytes.length <= 20) break;
    }
    let result;
    try {
      result = readXlsChart(recordStream(bytes), 0, context({ charts: 1, points: 1000 }, book.styles));
    } catch (error) {
      assert.fail(`round ${round}: the reader threw ${error instanceof Error ? error.stack : String(error)}`);
    }
    if (!result.chart) {
      assert.equal(typeof result.problem, 'string', `round ${round}: a chart that is not read says why`);
      refused++;
      continue;
    }
    read++;
    try {
      const values = chartValues(result.chart, { workbook: book });
      layoutChart(result.chart, values, 320, 240, fakeMeasure);
    } catch (error) {
      assert.fail(`round ${round}: a chart that was read cannot be drawn: ${error instanceof Error ? error.stack : String(error)}`);
    }
  }
  assert.ok(read > 100 && refused > 100, `both happen: ${read} read, ${refused} refused`);
  assert.ok(Date.now() - started < 120000, `six thousand charts in ${Date.now() - started} ms`);
});

test('xls chart: a workbook whose chart is damaged opens, or is refused as a workbook that is not valid', () => {
  const next = random(77);
  const drawing = drawingRecords([shape({ kind: 201, ot: 5, id: 1, extra: [chartSubstream(rich()[0])] })]);
  const stream = workbookStream({ tail: tail(), sheets: [{ name: 'Data', records: cat(cells(), drawing, number(9, 0, 9)) }, { name: 'Chart1', type: 2, records: chartRecords(rich()[1]) }] });
  // Where the charts are: after the cells of the first sheet.
  const from = stream.length - drawing.length - chartRecords(rich()[1]).length - 60;
  let opened = 0;
  for (let round = 0; round < 1500; round++) {
    const bytes = new Uint8Array(stream);
    for (let change = 0; change < 1 + Math.floor(next() * 3); change++) bytes[from + Math.floor(next() * (bytes.length - from))] = Math.floor(next() * 256);
    try {
      const { workbook } = readXlsStream(bytes);
      assert.equal(getCell(workbook.sheets[0], 1, 1)?.value, 50, `round ${round}: the cells are read`);
      opened++;
    } catch (error) {
      assert.ok(error instanceof Error && /^Invalid xls|xls/.test(error.message), `round ${round}: ${error instanceof Error ? error.stack : String(error)}`);
    }
  }
  assert.ok(opened > 500, `most workbooks open: ${opened}`);
});
