// What an independent review of the xls chart reader found, and the rules of
// the reader no test held yet: an axis that is switched off, substreams that
// share their records, the budgets at their edges, the flags of data labels,
// markers and axes, and how a series finds its chart group.
//
// Tests named UNSURE hold a rule that was taken from the specification of the
// format as remembered and never compared with a file of Excel: they say what
// the reader does today, not that it is right.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readXlsChart, XLS_MAX_CHART_CACHE_BYTES, XLS_MAX_CHART_DEPTH, XLS_MAX_CHART_RECORDS, type XlsChartContext } from '../src/formats/xls/chart';
import { recordStream } from '../src/formats/xls/records';
import { readXlsStream } from '../src/formats/xls/workbook';
import { buildXlsx } from '../src/formats/xlsx/build';
import { MAX_CHART_CACHE } from '../src/formats/xlsx/chart';
import { chartCacheList, MAX_SHEET_CHARTS, type Chart, type ChartAxis, type ChartDataLabels } from '../src/model/sheet/chart';
import { getCell, type DrawingChart, type Sheet, type Workbook } from '../src/model/sheet/types';
import { chartProblem, moveChart, removeChart, updateChart } from '../src/ops/charts';
import { drawingRecords, shape } from './xls-art-fixtures';
import {
  ai,
  areaFormat,
  axcExt,
  axisLine,
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
  frame,
  line,
  lineFormat,
  markerFormat,
  objectLink,
  plotArea,
  REC,
  rgb,
  seriesRec,
  serToCrt,
  siIndex,
  textBlock,
  tick,
  valueRange,
  type XlsChartSpec,
  type XlsSeriesSpec,
} from './xls-chart-fixtures';
import { bof, cat, defaultStyles, eof, externSheet, f64, number, ptg, rec, str8, supBookSelf, u16, u32, workbookStream, type Bytes } from './xls-fixtures';

const tail = (): Uint8Array => cat(supBookSelf(1), externSheet([0, 0, 0]));
const column = (col: number): number[] => ptg.area3d(0, 1, 3, col, col);
const sales: XlsSeriesSpec = { name: 'Sales', values: column(1), categories: column(0) };
const costs: XlsSeriesSpec = { name: 'Costs', values: column(2), categories: column(0) };
const cells = (): Uint8Array => cat(...[50, 100, 25].map((value, i) => number(i + 1, 1, value)));

function bookWith(substream: Bytes[], head?: Bytes): Workbook {
  const drawing = drawingRecords([shape({ kind: 201, ot: 5, id: 1, extra: substream.map((part) => (part instanceof Uint8Array ? part : new Uint8Array(part))) })]);
  return readXlsStream(workbookStream({ head, tail: tail(), sheets: [{ name: 'Data', records: cat(cells(), drawing, number(9, 0, 9)) }] })).workbook;
}

function chartOf(spec: XlsChartSpec): Chart {
  const sheet = bookWith([chartSubstream(spec)]).sheets[0];
  const object = sheet.drawings && sheet.drawings[0] ? sheet.drawings[0].object : undefined;
  assert.ok(object && object.type === 'chart', `the chart was read: ${JSON.stringify(sheet.warnings)}`);
  return (object as DrawingChart).chart;
}

function problemOf(substream: Bytes[]): string {
  const sheet = bookWith(substream).sheets[0];
  assert.deepEqual(sheet.drawings && sheet.drawings.map((drawing) => drawing.object.type), ['placeholder']);
  assert.equal(sheet.warnings && sheet.warnings.length, 1, JSON.stringify(sheet.warnings));
  return (sheet.warnings as string[])[0];
}

const chartBof = (): Uint8Array => bof(0x0020);
const around = (...inside: Bytes[]): Bytes[] => [chartBof(), block(chartRec(), ...inside), eof()];
const oneSeries = (...children: Bytes[]): Uint8Array => block(seriesRec(3, 3), ai(1, 2, column(1)), ...children);
const axesWith = (...groups: Bytes[]): Uint8Array => block(axisParentRec(0), ...defaultAxes(), ...groups);
const barGroup = (order = 0): Uint8Array => block(chartFormatRec(order), bar());

function context(counts: Partial<XlsChartContext['budget']> = {}): XlsChartContext {
  // The same object goes on being counted down: the tests look at it afterwards.
  const budget = counts as XlsChartContext['budget'];
  if (budget.charts === undefined) budget.charts = 10;
  if (budget.points === undefined) budget.points = 10000000;
  if (budget.bytes === undefined) budget.bytes = XLS_MAX_CHART_CACHE_BYTES;
  return {
    styles: readXlsStream(workbookStream({ sheets: [{ name: 'Data' }] })).workbook.styles,
    fontOf: () => undefined,
    formula: { biff8: true, codepage: 1252, sheetNames: ['Data'], names: [], xti: [{ book: 0, first: 0, last: 0 }], supBooks: [{ kind: 'self', sheetNames: ['Data'], externNames: [], ordinal: 0 }], failures: new Map() },
    date1904: false,
    budget,
  };
}

const valueAxis = (chart: Chart): ChartAxis => chart.axes.filter((axis) => axis.kind === 'val')[0];
const categoryAxis = (chart: Chart): ChartAxis => chart.axes.filter((axis) => axis.kind === 'cat')[0];
const withAxes = (category: Bytes[], value: Bytes[]): Chart => chartOf({ series: [sales], axes: [block(axisRec(0), ...category), block(axisRec(1), ...value)] });

// ------------------------------------------------------ B1: an axis that is off

test('xls chart: an axis whose line says that the axis is off is not drawn', () => {
  // AxisLine 0 is the axis itself: bit 2 of the LineFormat after it (fAxisOn) says whether the axis is shown.
  const off = withAxes([catSerRange(), tick(), axisLine(0), lineFormat({ auto: true })], [valueRange(), tick(), axisLine(0), lineFormat({ auto: true, axisOn: true })]);
  assert.equal(categoryAxis(off).deleted, true);
  assert.equal(valueAxis(off).deleted, undefined);
  // The line of an axis that is on is read as before; one that states "no line" hides the line, not the axis.
  const lineless = withAxes([catSerRange(), axisLine(0), lineFormat({ style: 5, axisOn: true })], [valueRange()]);
  assert.deepEqual([categoryAxis(lineless).deleted, categoryAxis(lineless).line], [undefined, { color: null }]);
  // UNSURE: an axis without an AxisLine record of its own is shown; the flag of grid lines says nothing about the axis.
  const bare = withAxes([catSerRange()], [valueRange(), axisLine(1), lineFormat({ color: 'C0C0C0' })]);
  assert.deepEqual([categoryAxis(bare).deleted, valueAxis(bare).deleted], [undefined, undefined]);
  // UNSURE: an AxisLine 0 that no LineFormat follows says nothing.
  const alone = withAxes([catSerRange(), axisLine(0)], [valueRange()]);
  assert.equal(categoryAxis(alone).deleted, undefined);
});

test('xls chart: the hidden category axis of a second axis group is not drawn', () => {
  const chart = chartOf({
    series: [{ ...sales, group: 0 }, { ...costs, group: 1 }],
    parents: [{ groups: [{ type: bar() }] }, { axes: [block(axisRec(0), catSerRange(), tick({ labels: 0 }), axisLine(0), lineFormat({ auto: true })), block(axisRec(1), valueRange(), tick(), axisLine(0), lineFormat({ auto: true, axisOn: true }))], groups: [{ type: line() }] }],
  });
  assert.deepEqual(chart.axes.map((axis) => [axis.id, axis.deleted === true]), [[1, false], [2, false], [4, true], [5, false]]);
});

// ------------------------------------------- B2: substreams that share records

/**
 * A workbook whose sheets start where `starts` says inside one run of bytes: [name, type, offset into `body`].
 * The offsets of the sheet list are worked out here; nothing checks that the substreams are apart.
 */
function bookOf(sheets: [name: string, type: number, at: number][], body: Uint8Array): Uint8Array {
  const head = cat(bof(0x0005), defaultStyles());
  const end = cat(supBookSelf(sheets.length), externSheet([0, 0, 0]), eof());
  const bound = (offset: number, name: string, type: number) => rec(0x0085, u32(offset), [0, type], str8(name));
  const size = sheets.reduce((sum, [name, type]) => sum + bound(0, name, type).length, 0);
  const base = head.length + size + end.length;
  return cat(head, ...sheets.map(([name, type, at]) => bound(base + at, name, type)), end, body);
}

test('xls chart: of two chart sheets that share their records the second is not read; the workbook opens', () => {
  const data = cat(bof(0x0010), cells(), eof());
  const records = chartRecords({ series: [sales] });
  // The outer substream starts with a record that swallows the BOF record of the inner one: both run into the same chart.
  const outer = cat(chartBof(), rec(0x10f0).subarray(0, 2), u16(20));
  const body = cat(data, outer, chartBof(), records, eof());
  const inner = data.length + outer.length;
  for (const order of ['inner first', 'outer first']) {
    const sheets: [string, number, number][] = order === 'inner first' ? [['Data', 0, 0], ['Inner', 2, inner], ['Outer', 2, data.length]] : [['Data', 0, 0], ['Outer', 2, data.length], ['Inner', 2, inner]];
    const { workbook } = readXlsStream(bookOf(sheets, body));
    assert.equal(getCell(workbook.sheets[0], 1, 1)?.value, 50, order);
    const [first, second] = [workbook.sheets[1], workbook.sheets[2]];
    assert.equal(first.drawings && first.drawings[0].object.type, 'chart', `${order}: the first chart sheet shows its chart`);
    assert.equal(second.drawings, undefined, `${order}: the second shows none`);
    assert.equal(second.warnings && second.warnings.length, 1, order);
    assert.match((second.warnings as string[])[0], /The chart of the sheet is not shown: its records are those of another sheet/, order);
  }
});

test('xls chart: of two worksheets that share their records the second stays empty, with a warning', () => {
  const payload = cat(...[1, 2, 3].map((value, i) => number(i, 0, value)), eof());
  const outer = cat(bof(0x0010), rec(0x10f0).subarray(0, 2), u16(20));
  const body = cat(outer, bof(0x0010), payload);
  for (const order of ['inner first', 'outer first']) {
    const sheets: [string, number, number][] = order === 'inner first' ? [['Inner', 0, outer.length], ['Outer', 0, 0]] : [['Outer', 0, 0], ['Inner', 0, outer.length]];
    const { workbook } = readXlsStream(bookOf(sheets, body));
    assert.equal(getCell(workbook.sheets[0], 2, 0)?.value, 3, `${order}: the first sheet has its cells`);
    assert.equal(workbook.sheets[1].rows.size, 0, `${order}: the second has none`);
    assert.deepEqual(workbook.sheets[1].warnings, ['The sheet is not read: its records are those of another sheet'], order);
    assert.equal(workbook.sheets[1].name, sheets[1][0]);
  }
});

test('xls chart: sheets that follow each other closely are read, each to its own end', () => {
  const body = cat(bof(0x0010), number(0, 0, 1), eof(), chartBof(), chartRecords({ series: [sales] }), eof(), bof(0x0010), number(0, 0, 3), eof());
  const second = 20 + number(0, 0, 1).length + 4;
  const third = second + 20 + chartRecords({ series: [sales] }).length + 4;
  // Listed back to front: what is read first lies behind what is read after it.
  const { workbook } = readXlsStream(bookOf([['C', 0, third], ['B', 2, second], ['A', 0, 0]], body));
  assert.deepEqual([getCell(workbook.sheets[0], 0, 0)?.value, getCell(workbook.sheets[2], 0, 0)?.value], [3, 1]);
  assert.equal(workbook.sheets[1].drawings && workbook.sheets[1].drawings[0].object.type, 'chart');
  assert.deepEqual(workbook.sheets.map((sheet) => sheet.warnings), [undefined, undefined, undefined]);
});

test('xls chart: the cached values of the charts of a workbook are bounded in bytes, their texts included', () => {
  assert.equal(XLS_MAX_CHART_CACHE_BYTES, 64 * 1024 * 1024);
  const long = 'x'.repeat(1000);
  const stream = recordStream(chartSubstream({ series: [sales], cache: [siIndex(2), cachedLabel(0, 0, long), cachedLabel(1, 0, long), siIndex(1), cachedNumber(0, 0, 1)] }));
  // A text costs two bytes a character and sixteen for its place, a number sixteen.
  const cost = 2 * (16 + 2000) + 16;
  const budget = { bytes: cost };
  const fits = readXlsChart(stream, 0, context(budget));
  assert.ok(fits.chart, fits.problem);
  assert.equal(budget.bytes, 0);
  const tight = { bytes: cost - 1 };
  assert.deepEqual(readXlsChart(stream, 0, context(tight)).problem, 'the charts of the workbook hold too many bytes of cached values');
});

// --------------------------------------------- B5: a chart that is only shown

test('xls chart: a chart that is only shown is neither changed, moved nor deleted; all three say the same', () => {
  const workbook = bookWith([chartSubstream({ series: [sales] })]);
  const expected = /Cannot change this chart: it was read from an xls file and is only shown/;
  assert.throws(() => updateChart(workbook, 0, 0, { title: 'New' }), expected);
  assert.throws(() => moveChart(workbook, 0, 0, { top: 0, left: 0, bottom: 5, right: 5 }), expected);
  assert.throws(() => removeChart(workbook, 0, 0), expected);
  assert.match(chartProblem(workbook, 0, 0) || '', expected);
});

test('xls chart: saving as xlsx says that the chart of an xls file is left out, not that its type is not written', async () => {
  const report = { parts: [] as string[], warnings: [] as string[] };
  await buildXlsx(bookWith([chartSubstream({ series: [sales] })]), { report });
  assert.deepEqual(report.warnings.filter((warning) => /chart/.test(warning)), ['Drawing: a chart read from an xls file is only shown and is left out']);
});

// ------------------------------------------------------------- B7: data labels

const labelsWith = (flags: number, position = 0): ChartDataLabels | undefined => chartOf({ series: [sales], tail: [textBlock(objectLink(4, 0, 0xffff), undefined, { flags: flags | 0x10, position })] }).groups[0].series[0].dataLabels;

test('xls chart: each flag of the Text record of a data label says one thing', () => {
  assert.deepEqual(labelsWith(0), {});
  assert.deepEqual(labelsWith(0x0004), { showValue: true });
  assert.deepEqual(labelsWith(0x1000), { showPercent: true });
  assert.deepEqual(labelsWith(0x0800), { showPercent: true, showCategory: true });
  assert.deepEqual(labelsWith(0x4000), { showCategory: true });
  assert.deepEqual(labelsWith(0x2000), { showBubbleSize: true });
  assert.deepEqual(labelsWith(0x0002), { showLegendKey: true });
  assert.deepEqual(labelsWith(0x0040 | 0x0004), { deleted: true });
  // Bits that say something else (fAutoMode, the ones not used) show nothing.
  assert.deepEqual(labelsWith(0x0080 | 0x0008 | 0x0700 | 0x8000), {});
});

test('xls chart: where a data label stands', () => {
  const places = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 15].map((position) => (labelsWith(0x4, position) as ChartDataLabels).position);
  assert.deepEqual(places, [undefined, 'outEnd', 'inEnd', 'ctr', 'inBase', 't', 'b', 'l', 'r', 'bestFit', undefined, undefined]);
  // Only the low four bits are the place.
  assert.equal((labelsWith(0x4, 0x33) as ChartDataLabels).position, 'ctr');
});

test('xls chart: a marker without its inside, and one without its outline', () => {
  const marker = (spec: Parameters<typeof markerFormat>[0]) => chartOf({ series: [{ ...sales, format: [markerFormat(spec)] }], groups: [{ type: line() }] }).groups[0].series[0].marker;
  assert.deepEqual(marker({ type: 8, fill: '00FF00', line: 'FF0000', noFill: true }), { symbol: 'circle', fill: null, line: { color: '#FF0000' }, size: 5 });
  assert.deepEqual(marker({ type: 8, fill: '00FF00', line: 'FF0000', noLine: true }), { symbol: 'circle', fill: '#00FF00', line: { color: null }, size: 5 });
  assert.deepEqual(marker({ type: 8, fill: '00FF00', line: 'FF0000' }), { symbol: 'circle', fill: '#00FF00', line: { color: '#FF0000' }, size: 5 });
});

// -------------------------------------------------- B7: a series and its group

test('UNSURE xls chart: a series names its chart group by the number the group states (icrt), not by its place', () => {
  // The specification says "index into the collection of ChartFormat records"; the reader goes by icrt, which in the files of Excel is the same.
  const chart = chartOf({ series: [{ ...sales, group: 0 }, { ...costs, group: 1 }], groups: [{ type: bar(), order: 1 }, { type: line(), order: 0 }] });
  assert.deepEqual(chart.groups.map((group) => [group.kind, group.series.map((series) => (series.name as { text: string }).text)]), [['bar', ['Costs']], ['line', ['Sales']]]);
  // Numbers that are no places at all still find their groups.
  const far = chartOf({ series: [{ ...sales, group: 7 }, { ...costs, group: 3 }], groups: [{ type: bar(), order: 3 }, { type: line(), order: 7 }] });
  assert.deepEqual(far.groups.map((group) => (group.series[0].name as { text: string }).text), ['Costs', 'Sales']);
});

test('UNSURE xls chart: where two chart groups state the same number, a series goes by the place of its group', () => {
  const chart = chartOf({ series: [{ ...sales, group: 1 }, { ...costs, group: 0 }], groups: [{ type: bar(), order: 5 }, { type: line(), order: 5 }] });
  assert.deepEqual(chart.groups.map((group) => [group.kind, (group.series[0].name as { text: string }).text]), [['bar', 'Costs'], ['line', 'Sales']]);
  // The number itself then names nothing.
  assert.match(problemOf([chartSubstream({ series: [{ ...sales, group: 5 }], groups: [{ type: bar(), order: 5 }, { type: line(), order: 5 }] })]), /names a chart group/);
});

test('xls chart: the Series record says how many points there are and what the categories are', () => {
  const counted = chartOf({ series: [{ ...sales, counts: [5, 7] }] }).groups[0].series[0];
  assert.deepEqual([counted.categories && counted.categories.count, counted.values && counted.values.count], [5, 7]);
  // UNSURE: sdtX 3 says that the categories are texts, 1 that they are numbers; what the cache holds says more.
  assert.equal(chartOf({ series: [{ ...sales }] }).groups[0].series[0].categories!.kind, 'str');
  assert.equal(chartOf({ series: [{ ...sales, numericCategories: true }] }).groups[0].series[0].categories!.kind, 'num');
  assert.equal(chartOf({ series: [{ ...sales, numericCategories: true }], cache: [siIndex(2), cachedLabel(0, 0, 'a')] }).groups[0].series[0].categories!.kind, 'str');
  assert.equal(chartOf({ series: [{ ...sales }], cache: [siIndex(2), cachedNumber(0, 0, 4)] }).groups[0].series[0].categories!.kind, 'num');
  // The count of the record is at least what the cache reaches to.
  const cached = chartOf({ series: [{ ...sales, counts: [1, 1] }], cache: [siIndex(1), cachedNumber(8, 0, 4)] }).groups[0].series[0];
  assert.equal(cached.values!.count, 9);
});

// ---------------------------------------------------------------- B7: axes

test('xls chart: each limit of a value axis is automatic by its own flag', () => {
  const low = valueAxis(withAxes([catSerRange()], [valueRange({ min: 5 })]));
  assert.deepEqual([low.min, low.max, low.majorUnit], [5, undefined, undefined]);
  const high = valueAxis(withAxes([catSerRange()], [valueRange({ max: 50 })]));
  assert.deepEqual([high.min, high.max, high.majorUnit], [undefined, 50, undefined]);
  const unit = valueAxis(withAxes([catSerRange()], [valueRange({ major: 10 })]));
  assert.deepEqual([unit.min, unit.max, unit.majorUnit], [undefined, undefined, 10]);
  // UNSURE: the unit of a logarithmic axis is not read (what it counts in is not certain).
  const log = valueAxis(withAxes([catSerRange()], [valueRange({ min: 0, max: 3, major: 1, log: true })]));
  assert.deepEqual([log.logBase, log.majorUnit], [10, undefined]);
  // A unit that is no step is none.
  const zero = valueAxis(withAxes([catSerRange()], [rec(REC.valueRange, f64(0), f64(0), f64(-2), f64(0), f64(0), u16(0x1b))]));
  assert.equal(zero.majorUnit, undefined);
});

test('xls chart: the category the value axis crosses at, the unit of a date axis and labels that are turned by themselves', () => {
  assert.equal(valueAxis(withAxes([catSerRange(1, 1)], [valueRange()])).crossesAt, undefined, 'the first category is where the axis stands by itself');
  assert.equal(valueAxis(withAxes([catSerRange(1, 2)], [valueRange()])).crossesAt, 2);
  // At the far end the place is not read: "the last category" says it.
  const far = valueAxis(withAxes([catSerRange(3, 4)], [valueRange()]));
  assert.deepEqual([far.crosses, far.crossesAt], ['max', undefined]);
  // AxcExt: a date axis whose unit is automatic (bit 5) states none; one that is no date axis states nothing.
  const automatic = categoryAxis(withAxes([catSerRange(), axcExt(0x30, 2)], [valueRange()]));
  assert.deepEqual([automatic.date, automatic.baseTimeUnit], [true, undefined]);
  const stated = categoryAxis(withAxes([catSerRange(), axcExt(0x10, 2)], [valueRange()]));
  assert.deepEqual([stated.date, stated.baseTimeUnit], [true, 'years']);
  const plain = categoryAxis(withAxes([catSerRange(), axcExt(0x20, 2)], [valueRange()]));
  assert.deepEqual([plain.date, plain.baseTimeUnit], [undefined, undefined]);
  // Tick: a turn that is automatic (bit 5) is none, whatever the turn says.
  const tickWith = (flags: number) => rec(REC.tick, [2, 0, 3, 1], rgb('000000'), new Array(16).fill(0), u16(flags), u16(0x4d), u16(0));
  assert.equal(categoryAxis(withAxes([catSerRange(), tickWith(1 | (2 << 2) | 0x20)], [valueRange()])).labelRotation, undefined);
  // UNSURE: rot 2 is a quarter turn counter-clockwise, 3 clockwise; 1 (stacked letters) is not applied.
  assert.deepEqual([1, 2, 3].map((turn) => categoryAxis(withAxes([catSerRange(), tickWith(1 | (turn << 2))], [valueRange()])).labelRotation), [undefined, -90, 90]);
});

test('xls chart: the frame of the plot area is that of the first axis group', () => {
  const chart = chartOf({
    series: [{ ...sales, group: 0 }, { ...costs, group: 1 }],
    parents: [
      { axes: [block(axisRec(0), catSerRange()), block(axisRec(1), valueRange()), plotArea(), frame({ color: '111111' }, { color: 'AAAAAA' })], groups: [{ type: bar() }] },
      { axes: [block(axisRec(0), catSerRange()), block(axisRec(1), valueRange()), plotArea(), frame({ color: '222222' }, { color: 'BBBBBB' })], groups: [{ type: line() }] },
    ],
  });
  assert.deepEqual(chart.plotArea, { fill: '#AAAAAA', line: { color: '#111111', width: 1 } });
  // A frame that does not follow a PlotArea record is no plot area.
  const none = chartOf({ series: [sales], axes: [block(axisRec(0), catSerRange()), block(axisRec(1), valueRange()), frame({ color: '111111' }, { color: 'AAAAAA' })] });
  assert.equal(none.plotArea, undefined);
});

// -------------------------------------------------------------- B7: limits

test('xls chart: blocks are followed to the depth the reader states, and no deeper', () => {
  // The Chart block is the first level; `count` blocks inside each other reach count + 1.
  const nested = (count: number) => around(oneSeries(), axesWith(barGroup()), ...new Array(count).fill(cat(rec(0x10f0), begin())), ...new Array(count).fill(end()));
  const sheet = bookWith(nested(XLS_MAX_CHART_DEPTH - 1)).sheets[0];
  assert.equal(sheet.drawings && sheet.drawings[0].object.type, 'chart', JSON.stringify(sheet.warnings));
  assert.match(problemOf(nested(XLS_MAX_CHART_DEPTH)), /nested too deeply/);
});

test('xls chart: a chart holds as many records as the reader states, and not one more', () => {
  const inside = cat(oneSeries(), axesWith(barGroup()));
  // The records of the block, Begin and End not counted.
  let own = 0;
  for (let p = 0; p < inside.length; p += 4 + (inside[p + 2] | (inside[p + 3] << 8))) {
    const id = inside[p] | (inside[p + 1] << 8);
    if (id !== REC.begin && id !== REC.end) own++;
  }
  const filled = (count: number) => {
    const filler = new Uint8Array(4 * count);
    for (let p = 0; p < filler.length; p += 4) filler.set([0x22, 0x10, 0, 0], p);
    return recordStream(cat(chartBof(), block(chartRec(), inside, filler), eof()));
  };
  const most = readXlsChart(filled(XLS_MAX_CHART_RECORDS - own), 0, context());
  assert.ok(most.chart, most.problem);
  assert.deepEqual(readXlsChart(filled(XLS_MAX_CHART_RECORDS - own + 1), 0, context()).problem, 'it holds too many records');
});

test('xls chart: the cached values of one chart beyond the limit are left out, with a note', () => {
  // More than a million values: sixteen series of 65,536 points each hold 1,048,576.
  const perSeries = 65536;
  const series = 16;
  const record = 18;
  const values = new Uint8Array(series * perSeries * record);
  const view = new DataView(values.buffer);
  let p = 0;
  for (let s = 0; s < series; s++) {
    for (let point = 0; point < perSeries; point++, p += record) {
      view.setUint16(p, 0x0203, true);
      view.setUint16(p + 2, 14, true);
      view.setUint16(p + 4, point, true);
      view.setUint16(p + 6, s, true);
      view.setFloat64(p + 10, 1, true);
    }
  }
  const head = cat(chartBof(), block(chartRec(), block(seriesRec(0, 0), ai(1, 1)), axesWith(barGroup())), siIndex(1));
  const exact = MAX_CHART_CACHE * record;
  const fits = readXlsChart(recordStream(cat(head, values.subarray(0, exact), eof())), 0, context());
  assert.ok(fits.chart, fits.problem);
  assert.deepEqual(fits.chart.notes, []);
  const budget = { points: 10000000 };
  const over = readXlsChart(recordStream(cat(head, values.subarray(0, exact + record), eof())), 0, context(budget));
  assert.ok(over.chart, over.problem);
  assert.deepEqual(over.chart.notes, [`Only the first ${MAX_CHART_CACHE} cached values of a chart are read`]);
  assert.equal(budget.points, 10000000 - MAX_CHART_CACHE, 'what is left out costs nothing');
});

test('xls chart: a sheet shows as many charts as the limit of a sheet, the next one is a placeholder', () => {
  const small = cat(chartBof(), block(chartRec(), block(seriesRec(0, 1), ai(1, 1), serToCrt(0)), block(axisParentRec(0), barGroup())), eof());
  const shapes = [];
  for (let i = 0; i < MAX_SHEET_CHARTS + 1; i++) shapes.push(shape({ kind: 201, ot: 5, id: i + 1, extra: [small] }));
  const sheet = readXlsStream(workbookStream({ sheets: [{ name: 'Data', records: drawingRecords(shapes) }] })).workbook.sheets[0];
  const kinds = (sheet.drawings as NonNullable<Sheet['drawings']>).map((drawing) => drawing.object.type);
  assert.equal(kinds.filter((kind) => kind === 'chart').length, MAX_SHEET_CHARTS);
  assert.equal(kinds[MAX_SHEET_CHARTS], 'placeholder');
  assert.deepEqual(sheet.warnings, ['Chart is not shown: the sheet holds too many charts']);
});

test('xls chart: sixteen chart groups are read, a seventeenth is too many', () => {
  const groups = (count: number): XlsChartSpec => ({ series: [sales], groups: Array.from({ length: count }, (_value, i) => ({ type: i === 0 ? bar() : line() })) });
  assert.equal(chartOf(groups(16)).groups.length, 1, 'groups without a series are not kept');
  assert.match(problemOf([chartSubstream(groups(17))]), /too many chart groups/);
});

test('xls chart: a record that runs past the end of the stream leaves the chart unread, wherever it stands', () => {
  const whole = chartSubstream({ series: [sales] });
  // Inside the blocks, and after them.
  const inBlock = whole.slice(0, whole.length - 30);
  assert.match(readXlsChart(recordStream(inBlock), 0, context()).problem || '', /runs past the end of the stream|not closed/);
  const long = cat(whole.subarray(0, whole.length - 4), rec(0x10f0).subarray(0, 2), u16(500));
  assert.deepEqual(readXlsChart(recordStream(long), 0, context()).problem, 'a record runs past the end of the stream');
  const inside = cat(chartBof(), chartRec(), begin(), rec(REC.series).subarray(0, 2), u16(900), [1, 2, 3]);
  assert.deepEqual(readXlsChart(recordStream(inside), 0, context()).problem, 'a record runs past the end of the stream');
  const headless = cat(chartBof(), chartRec(), begin(), [0x03]);
  assert.deepEqual(readXlsChart(recordStream(headless), 0, context()).problem, 'a block of its records is not closed');
});

// ------------------------------------------------------------ B7: structure

test('xls chart: the series of error bars is no series of the chart', () => {
  // SerAuxErrBar: sertm, ebsrc, fTeeTop, reserved, numValue, cnum.
  const bars = rec(0x105b, [1, 2, 1, 0], f64(5), u16(3));
  const chart = chartOf({ series: [sales, { name: 'Error', tail: [bars] }, costs] });
  assert.deepEqual(chart.groups[0].series.map((series) => (series.name as { text: string }).text), ['Sales', 'Costs']);
  assert.ok(chart.notes.some((note) => /error bars/.test(note)), JSON.stringify(chart.notes));
});

test('xls chart: of two formats of the whole series the first one is the format', () => {
  const chart = chartOf({ series: [{ ...sales, number: 3, format: [areaFormat({ color: 'FF0000' })], tail: [block(dataFormat(0xffff, 0, 9), areaFormat({ color: '00FF00' }))] }] });
  const series = chart.groups[0].series[0];
  assert.deepEqual([series.index, series.shape], [3, { fill: '#FF0000' }]);
});

test('xls chart: another substream that starts inside a block leaves the chart unread', () => {
  assert.match(problemOf(around(oneSeries(), bof(0x0010), eof(), axesWith(barGroup()))), /not closed/);
  // One between the Chart block and the cache is passed over.
  const sheet = bookWith([chartBof(), block(chartRec(), oneSeries(), axesWith(barGroup())), bof(0x0010), number(0, 0, 1), eof(), siIndex(1), cachedNumber(0, 0, 7), eof()]).sheets[0];
  assert.equal(sheet.drawings && sheet.drawings[0].object.type, 'chart', JSON.stringify(sheet.warnings));
  assert.deepEqual(chartCacheList(((sheet.drawings as NonNullable<Sheet['drawings']>)[0].object as DrawingChart).chart.groups[0].series[0].values), [7]);
});

test('xls chart: the dates of a chart count as the workbook says', () => {
  const substream = [chartSubstream({ series: [sales] })];
  const from1904 = bookWith(substream, cat(rec(0x0022, u16(1)), defaultStyles()));
  assert.equal(from1904.date1904, true);
  assert.equal(((from1904.sheets[0].drawings as NonNullable<Sheet['drawings']>)[0].object as DrawingChart).chart.date1904, true);
  assert.equal(((bookWith(substream).sheets[0].drawings as NonNullable<Sheet['drawings']>)[0].object as DrawingChart).chart.date1904, undefined);
});

// ------------------------------------------------------------------ UNSURE

test('UNSURE xls chart: the overlap of bars is stated with the other sign than a chart part states it', () => {
  const overlap = (stated: number) => chartOf({ series: [sales], groups: [{ type: bar(0, 150, stated) }] }).groups[0].overlap;
  assert.deepEqual([overlap(-100), overlap(-30), overlap(0), overlap(30), overlap(100)], [100, 30, 0, -30, -100]);
});

test('UNSURE xls chart: the place of the legend is the seventeenth byte of its record', () => {
  const at = (bytes: number[]) => chartOf({ series: [sales], groups: [{ type: bar(), children: [block(rec(REC.legend, bytes))] }] }).legend;
  const record = (type: number) => [...new Array(16).fill(0xee), type, 1, 0x1f, 0];
  assert.deepEqual([0, 1, 2, 3, 4].map((type) => (at(record(type)) || { position: '' }).position), ['b', 'tr', 't', 'r', 'l']);
});

test('UNSURE xls chart: series without a colour take the palette from 24 on where they are filled, from 32 on where they are lines; six of each', () => {
  // Excel 2003 may go on through all eight of a row and into the next; here the six the model holds are taken, and tones of them after that.
  assert.deepEqual(chartOf({ series: [sales] }).palette, ['#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080']);
  assert.deepEqual(chartOf({ series: [sales], groups: [{ type: line() }] }).palette, ['#000080', '#FF00FF', '#FFFF00', '#00FFFF', '#800080', '#800000']);
  // A chart of bars and lines takes the row of its first group for both.
  assert.deepEqual(chartOf({ series: [{ ...sales, group: 0 }, { ...costs, group: 1 }], groups: [{ type: bar() }, { type: line() }] }).palette[0], '#9999FF');
});

test('UNSURE xls chart: a cached number that no SIIndex record precedes is of nothing', () => {
  const chart = chartOf({ series: [sales], cache: [cachedNumber(0, 0, 9), siIndex(1), cachedNumber(1, 0, 4)] });
  assert.deepEqual(chartCacheList(chart.groups[0].series[0].values), [null, 4]);
});

test('UNSURE xls chart: a logarithmic axis states the exponents of its limits and of where it is crossed', () => {
  const chart = withAxes([catSerRange()], [valueRange({ min: -1, max: 2, cross: 1, log: true })]);
  assert.deepEqual([valueAxis(chart).min, valueAxis(chart).max, categoryAxis(chart).crossesAt], [0.1, 100, 10]);
});

test('UNSURE xls chart: a series of lines that states no marker has the markers the chart gives', () => {
  const group = chartOf({ series: [sales], groups: [{ type: line() }] }).groups[0];
  assert.deepEqual([group.markers, group.series[0].marker], [true, undefined]);
});
