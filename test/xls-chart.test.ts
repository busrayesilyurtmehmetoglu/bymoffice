// Charts of an xls workbook: the substream of an embedded chart and of a chart
// sheet is read into the model the charts of an xlsx file are read into, the
// series are drawn from the cells their references name, and what cannot be
// read for certain stays a placeholder. Every stream is built here, record by
// record (see xls-chart-fixtures.ts: none of it was compared with a file of
// Excel).

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readXlsStream } from '../src/formats/xls/workbook';
import { buildXlsx } from '../src/formats/xlsx/build';
import { open } from '../src/index';
import { chartValues, layoutChart } from '../src/layout/chart';
import { chartCacheList, chartRefAreas, chartWritable, type Chart, type ChartAxis, type ChartGroup } from '../src/model/sheet/chart';
import type { DrawingChart, Sheet, Workbook } from '../src/model/sheet/types';
import { ZipArchive } from '../src/zip';
import { fakeMeasure } from './chart-fixtures';
import { clientAnchor, drawingRecords, shape } from './xls-art-fixtures';
import {
  area,
  areaFormat,
  attachedLabel,
  axcExt,
  axisLine,
  axisRec,
  bar,
  block,
  cachedLabel,
  cachedNumber,
  catSerRange,
  chart3d,
  chartRecords,
  chartSubstream,
  crtLine,
  defaultAxes,
  defaultText,
  dropBar,
  fontX,
  frame,
  ifmt,
  legendRec,
  line,
  lineFormat,
  markerFormat,
  objectLink,
  pie,
  pieFormat,
  plotArea,
  radar,
  radarArea,
  scatter,
  serAuxTrend,
  serFmt,
  serParent,
  shtProps,
  siIndex,
  textBlock,
  textRec,
  tick,
  valueRange,
  type XlsChartSpec,
  type XlsSeriesSpec,
} from './xls-chart-fixtures';
import { buildCfb } from './doc-binary-fixtures';
import { cat, defaultStyles, externSheet, font, format, number, ptg, rec, str16, supBookSelf, u16, workbookStream, type Bytes } from './xls-fixtures';

const label = (row: number, col: number, text: string): Uint8Array => rec(0x0204, u16(row), u16(col), u16(15), str16(text));

/** The cells of the sheet "Data": categories a, b, c in A2:A4, the series "Sales" (50, 100, 25) in B and "Costs" (1, 2, 3) in C, sizes in D. */
const cells = (): Uint8Array =>
  cat(
    label(0, 1, 'Sales'),
    label(0, 2, 'Costs'),
    ...['a', 'b', 'c'].map((text, i) => label(i + 1, 0, text)),
    ...[50, 100, 25].map((value, i) => number(i + 1, 1, value)),
    ...[1, 2, 3].map((value, i) => number(i + 1, 2, value)),
    ...[4, 9, 16].map((value, i) => number(i + 1, 3, value)),
  );

/** The sheet "Data" is the first of the workbook: the one the references name. */
const tail = (): Uint8Array => cat(supBookSelf(2), externSheet([0, 0, 0]));
const column = (col: number): number[] => ptg.area3d(0, 1, 3, col, col);
const sales: XlsSeriesSpec = { name: ptg.ref3d(0, 0, 1), values: column(1), categories: column(0) };
const costs: XlsSeriesSpec = { name: ptg.ref3d(0, 0, 2), values: column(2), categories: column(0) };

function embeddedBook(spec: XlsChartSpec, head?: Bytes): Workbook {
  const drawing = drawingRecords([shape({ kind: 201, ot: 5, id: 1, extra: [chartSubstream(spec)], props: [[0x0380, 'Chart 1']], anchor: clientAnchor([4, 0, 1, 0], [9, 0, 16, 0]) })]);
  return readXlsStream(workbookStream({ head, tail: tail(), sheets: [{ name: 'Data', records: cat(cells(), drawing) }] })).workbook;
}

function chartOf(spec: XlsChartSpec, head?: Bytes): Chart {
  const sheet = embeddedBook(spec, head).sheets[0];
  const object = sheet.drawings && sheet.drawings[0] ? sheet.drawings[0].object : undefined;
  assert.ok(object && object.type === 'chart', `the chart was read: ${JSON.stringify(sheet.warnings)}`);
  assert.equal(sheet.warnings, undefined);
  return object.chart;
}

/** The warning of the sheet whose chart was not read. */
function problemOf(spec: XlsChartSpec): string {
  const sheet = embeddedBook(spec).sheets[0];
  assert.deepEqual(sheet.drawings && sheet.drawings.map((drawing) => drawing.object.type), ['placeholder']);
  assert.equal(sheet.warnings && sheet.warnings.length, 1, JSON.stringify(sheet.warnings));
  return (sheet.warnings as string[])[0];
}

const groupOf = (spec: XlsChartSpec): ChartGroup => chartOf(spec).groups[0];

// -------------------------------------------------------------------- reader

test('xls chart: an embedded chart is read with its series, their references and its place on the sheet', () => {
  const workbook = embeddedBook({ series: [sales, costs], groups: [{ type: bar(0, 150, 0) }] });
  const sheet = workbook.sheets[0];
  assert.equal(sheet.drawings && sheet.drawings.length, 1);
  const drawing = (sheet.drawings as NonNullable<Sheet['drawings']>)[0];
  assert.deepEqual(drawing.anchor, { kind: 'twoCell', from: { col: 4, colOff: 0, colShare: 0, row: 1, rowOff: 0, rowShare: 0 }, to: { col: 9, colOff: 0, colShare: 0, row: 16, rowOff: 0, rowShare: 0 } });
  const object = drawing.object as DrawingChart;
  assert.equal(object.type, 'chart');
  assert.equal(object.name, 'Chart 1');
  assert.equal(object.part, undefined);
  const chart = object.chart;
  assert.equal(chart.groups.length, 1);
  const group = chart.groups[0];
  assert.deepEqual([group.kind, group.barDir, group.grouping, group.gapWidth, group.overlap], ['bar', 'col', 'standard', 150, 0]);
  assert.deepEqual(group.series.map((series) => [series.index, series.order]), [[0, 0], [1, 1]]);
  assert.deepEqual(group.series.map((series) => series.name), [{ ref: 'Data!$B$1' }, { ref: 'Data!$C$1' }]);
  assert.deepEqual(group.series.map((series) => series.values && series.values.ref), ['Data!$B$2:$B$4', 'Data!$C$2:$C$4']);
  assert.deepEqual(group.series.map((series) => series.categories && series.categories.ref), ['Data!$A$2:$A$4', 'Data!$A$2:$A$4']);
  assert.deepEqual(group.series.map((series) => series.values && [series.values.kind, series.values.count]), [['num', 3], ['num', 3]]);
  assert.equal(chart.shownOnly, true);
  assert.equal(chart.autoTitleDeleted, true, 'a chart of this format makes up no title');
  assert.equal(sheet.warnings, undefined);
  // The cells of the sheet are read as before.
  assert.equal(sheet.rows.get(1)?.cells[1]?.value, 50);
});

test('xls chart: the series are drawn from the cells as they are now', () => {
  const workbook = embeddedBook({ series: [sales, costs] });
  const chart = ((workbook.sheets[0].drawings as NonNullable<Sheet['drawings']>)[0].object as DrawingChart).chart;
  const areas = chartRefAreas(workbook, 'Data!$B$2:$B$4');
  assert.ok(areas && areas.length === 1, 'the reference is followed to the sheet');
  let values = chartValues(chart, { workbook });
  assert.deepEqual(values.series.map((series) => [series.name, series.live, series.values.toArray()]), [['Sales', true, [50, 100, 25]], ['Costs', true, [1, 2, 3]]]);
  assert.deepEqual([0, 1, 2].map((i) => values.label(i)), ['a', 'b', 'c']);
  // A cell changes: the chart shows it.
  const cell = workbook.sheets[0].rows.get(2)?.cells[1];
  assert.ok(cell, 'B3 is a cell');
  workbook.sheets[0].rows.get(2)!.cells[1] = { ...cell, value: 70 };
  values = chartValues(chart, { workbook });
  assert.deepEqual(values.series[0].values.toArray(), [50, 70, 25]);
  // And the chart can be laid out.
  const layout = layoutChart(chart, values, 400, 300, fakeMeasure);
  assert.ok(layout.items.some((item) => item.kind === 'rect'), 'columns are drawn');
});

test('xls chart: the columns of a series are as high as the cells say, in the colour the series states', () => {
  // What the browser test looks for in the pixels: three columns of one colour over an axis from 0 to 100.
  const workbook = embeddedBook({
    series: [{ ...sales, format: [lineFormat({ style: 5 }), areaFormat({ color: '00B050' })] }],
    axes: [block(axisRec(0), catSerRange(), tick()), block(axisRec(1), valueRange({ min: 0, max: 100, major: 20 }), tick())],
    groups: [{ type: bar(0, 100) }],
  });
  const chart = ((workbook.sheets[0].drawings as NonNullable<Sheet['drawings']>)[0].object as DrawingChart).chart;
  const layout = layoutChart(chart, chartValues(chart, { workbook }), 256, 260, fakeMeasure);
  const columns = layout.items.filter((item) => item.kind === 'rect' && item.fill === '#00B050') as { x: number; height: number }[];
  assert.equal(columns.length, 3);
  columns.sort((a, b) => a.x - b.x);
  // 50, 100 and 25.
  assert.ok(Math.abs(columns[0].height / columns[1].height - 0.5) < 0.02, `the first column is half the second: ${columns[0].height} / ${columns[1].height}`);
  assert.ok(Math.abs(columns[2].height / columns[1].height - 0.25) < 0.02, `the third column is a quarter of the second: ${columns[2].height} / ${columns[1].height}`);
  assert.ok(columns[1].height > 100, `the column of 100 fills most of the plot area: ${columns[1].height}`);
});

test('xls chart: the values the file cached are read, and drawn where a series names no cells', () => {
  const cache = [
    siIndex(1),
    cachedNumber(0, 0, 7),
    cachedNumber(1, 0, 8),
    cachedNumber(2, 0, 9),
    cachedNumber(0, 1, 5),
    cachedNumber(2, 1, 6),
    siIndex(2),
    cachedLabel(0, 0, 'x'),
    cachedLabel(1, 0, 'y'),
    cachedLabel(2, 0, 'z'),
  ];
  const workbook = embeddedBook({ series: [{ name: 'Literal', counts: [3, 3] }, { ...costs }], cache });
  const chart = ((workbook.sheets[0].drawings as NonNullable<Sheet['drawings']>)[0].object as DrawingChart).chart;
  const [literal, linked] = chart.groups[0].series;
  assert.deepEqual(literal.name, { text: 'Literal' });
  assert.equal(literal.values && literal.values.ref, undefined);
  assert.deepEqual(chartCacheList(literal.values), [7, 8, 9]);
  assert.deepEqual(literal.categories && [literal.categories.kind, chartCacheList(literal.categories)], ['str', ['x', 'y', 'z']]);
  assert.deepEqual(chartCacheList(linked.values), [5, null, 6], 'a point the file cached nothing for is empty');
  const values = chartValues(chart, { workbook });
  assert.deepEqual(values.series.map((series) => [series.live, series.values.toArray()]), [[false, [7, 8, 9]], [true, [1, 2, 3]]]);
  assert.deepEqual([0, 1, 2].map((i) => values.label(i)), ['x', 'y', 'z']);
});

test('xls chart: a chart sheet shows its chart; the other sheets are read as before', async () => {
  const stream = workbookStream({ tail: tail(), sheets: [{ name: 'Data', records: cells() }, { name: 'Chart1', type: 2, records: chartRecords({ series: [sales], size: [600, 400] }) }] });
  const { workbook } = readXlsStream(stream);
  const sheet = workbook.sheets[1];
  assert.equal(sheet.kind, 'chartsheet');
  assert.equal(sheet.drawings && sheet.drawings.length, 1);
  const drawing = (sheet.drawings as NonNullable<Sheet['drawings']>)[0];
  assert.deepEqual(drawing.anchor, { kind: 'absolute', x: 0, y: 0, width: 600 * 12700, height: 400 * 12700 });
  assert.equal(drawing.object.type, 'chart');
  const chart = (drawing.object as DrawingChart).chart;
  assert.deepEqual(chartValues(chart, { workbook }).series[0].values.toArray(), [50, 100, 25]);
  assert.equal(sheet.warnings, undefined);
  assert.equal(workbook.sheets[0].drawings, undefined);
});

test('xls chart: a reference of several areas is read as one', () => {
  const union = [...ptg.memFunc(22), ...ptg.area3d(0, 1, 1, 1, 1), ...ptg.area3d(0, 3, 3, 1, 1), ...ptg.union];
  const workbook = embeddedBook({ series: [{ ...sales, values: union, counts: [3, 2] }] });
  const chart = ((workbook.sheets[0].drawings as NonNullable<Sheet['drawings']>)[0].object as DrawingChart).chart;
  const ref = chart.groups[0].series[0].values!.ref as string;
  assert.equal(ref, '(Data!$B$2:$B$2,Data!$B$4:$B$4)');
  assert.deepEqual(chartValues(chart, { workbook }).series[0].values.toArray(), [50, 25]);
});

// --------------------------------------------------------------------- types

test('xls chart: bars that lie, stacked bars and bars stacked to a hundred per cent', () => {
  assert.deepEqual([groupOf({ series: [sales], groups: [{ type: bar(1) }] }).barDir, groupOf({ series: [sales], groups: [{ type: bar(0) }] }).barDir], ['bar', 'col']);
  const stacked = groupOf({ series: [sales], groups: [{ type: bar(2, 80, -100) }] });
  assert.deepEqual([stacked.grouping, stacked.gapWidth, stacked.overlap], ['stacked', 80, 100]);
  assert.equal(groupOf({ series: [sales], groups: [{ type: bar(6, 150, -100) }] }).grouping, 'percentStacked');
  // UNSURE: the record states the overlap with the other sign than a chart part does.
  assert.equal(groupOf({ series: [sales], groups: [{ type: bar(0, 150, 30) }] }).overlap, -30);
  // The categories of bars that lie stand at the left, the values below.
  const lying = chartOf({ series: [sales], groups: [{ type: bar(1) }] });
  assert.deepEqual(lying.axes.map((axis) => [axis.kind, axis.position]), [['cat', 'l'], ['val', 'b']]);
});

test('xls chart: lines and areas, stacked or not', () => {
  const plain = groupOf({ series: [sales], groups: [{ type: line(0) }] });
  assert.deepEqual([plain.kind, plain.grouping, plain.markers], ['line', 'standard', true]);
  assert.equal(groupOf({ series: [sales], groups: [{ type: line(1) }] }).grouping, 'stacked');
  assert.equal(groupOf({ series: [sales], groups: [{ type: line(3) }] }).grouping, 'percentStacked');
  const filled = groupOf({ series: [sales], groups: [{ type: area(1) }] });
  assert.deepEqual([filled.kind, filled.grouping], ['area', 'stacked']);
});

test('xls chart: a pie, the angle of its first slice, and a doughnut', () => {
  const round = groupOf({ series: [sales], axes: [], groups: [{ type: pie(90, 0), varied: true }] });
  assert.deepEqual([round.kind, round.firstSliceAngle, round.varyColors, round.holeSize], ['pie', 90, true, undefined]);
  const ring = groupOf({ series: [sales], axes: [], groups: [{ type: pie(0, 50), varied: true }] });
  assert.deepEqual([ring.kind, ring.holeSize], ['doughnut', 50]);
  assert.deepEqual(chartOf({ series: [sales], axes: [], groups: [{ type: pie() }] }).axes, []);
});

test('xls chart: a scatter chart has numbers along both axes; a bubble chart has sizes', () => {
  const xy = { ...sales, categories: column(2), numericCategories: true };
  const axes = [block(axisRec(0), valueRange(), tick()), block(axisRec(1), valueRange(), tick())];
  const workbook = embeddedBook({ series: [xy], axes, groups: [{ type: scatter(0) }] });
  const chart = ((workbook.sheets[0].drawings as NonNullable<Sheet['drawings']>)[0].object as DrawingChart).chart;
  assert.deepEqual([chart.groups[0].kind, chart.groups[0].scatterStyle], ['scatter', 'lineMarker']);
  assert.deepEqual(chart.axes.map((axis) => [axis.kind, axis.position]), [['val', 'b'], ['val', 'l']]);
  assert.equal(chart.groups[0].series[0].categories!.kind, 'num');
  const values = chartValues(chart, { workbook });
  assert.deepEqual(values.series[0].xs && values.series[0].xs.toArray(), [1, 2, 3]);

  const bubbles = embeddedBook({ series: [{ ...xy, sizes: column(3) }], axes, groups: [{ type: scatter(3, 120, 2) }] });
  const bubble = ((bubbles.sheets[0].drawings as NonNullable<Sheet['drawings']>)[0].object as DrawingChart).chart;
  const group = bubble.groups[0];
  assert.deepEqual([group.kind, group.bubbleScale, group.sizeRepresents, group.showNegBubbles], ['bubble', 120, 'w', true]);
  assert.equal(group.series[0].bubbleSizes!.ref, 'Data!$D$2:$D$4');
  const sizes = chartValues(bubble, { workbook: bubbles }).series[0].sizes;
  assert.deepEqual(sizes && sizes.toArray(), [4, 9, 16]);
});

test('xls chart: a radar chart, with lines or filled', () => {
  assert.deepEqual([groupOf({ series: [sales], groups: [{ type: radar() }] }).kind, groupOf({ series: [sales], groups: [{ type: radar() }] }).radarStyle], ['radar', 'marker']);
  assert.equal(groupOf({ series: [sales], groups: [{ type: radarArea() }] }).radarStyle, 'filled');
});

test('xls chart: a 3-D chart is read as its flat form, with a note', () => {
  const chart = chartOf({ series: [sales], groups: [{ type: bar(), children: [chart3d()] }] });
  assert.equal(chart.groups[0].threeD, true);
  assert.ok(chart.notes.some((note) => /3-D/.test(note)), JSON.stringify(chart.notes));
});

test('xls chart: high-low lines, drop lines and up-down bars of a line chart', () => {
  const group = groupOf({
    series: [sales, costs],
    groups: [{ type: line(), children: [crtLine(1), lineFormat({ color: 'FF0000' }), crtLine(0), lineFormat({ auto: true }), block(dropBar(80), lineFormat({ color: '000000' }), areaFormat({ color: 'FFFFFF' })), block(dropBar(80), lineFormat({ color: '000000' }), areaFormat({ color: '000000' }))] }],
  });
  assert.deepEqual(group.hiLowLines, { color: '#FF0000', width: 1 });
  assert.deepEqual(group.dropLines, {});
  assert.deepEqual(group.upDownBars, { gapWidth: 80, up: { fill: '#FFFFFF', line: { color: '#000000', width: 1 } }, down: { fill: '#000000', line: { color: '#000000', width: 1 } } });
});

test('xls chart: two chart groups on one set of axes are a combination chart', () => {
  const chart = chartOf({ series: [{ ...sales, group: 0 }, { ...costs, group: 1 }], groups: [{ type: bar() }, { type: line() }] });
  assert.deepEqual(chart.groups.map((group) => [group.kind, group.series.length, group.axisIds]), [['bar', 1, [1, 2]], ['line', 1, [1, 2]]]);
  assert.deepEqual(chart.groups[1].series[0].name, { ref: 'Data!$C$1' });
  // The series keep their order among all of them.
  assert.deepEqual(chart.groups.map((group) => group.series[0].order), [0, 1]);
});

test('xls chart: a second axis group gives its series axes of their own', () => {
  const chart = chartOf({
    series: [{ ...sales, group: 0 }, { ...costs, group: 1 }],
    parents: [{ groups: [{ type: bar() }] }, { axes: [block(axisRec(0), catSerRange(), tick({ labels: 0 })), block(axisRec(1), valueRange({ max: 10 }), tick())], groups: [{ type: line() }] }],
  });
  assert.deepEqual(chart.axes.map((axis) => [axis.id, axis.kind, axis.position, axis.crossAxis]), [[1, 'cat', 'b', 2], [2, 'val', 'l', 1], [4, 'cat', 't', 5], [5, 'val', 'r', 4]]);
  assert.deepEqual(chart.groups.map((group) => [group.kind, group.axisIds]), [['bar', [1, 2]], ['line', [4, 5]]]);
  assert.equal(chart.axes[3].max, 10);
  assert.equal(chart.axes[2].tickLabels, 'none');
});

// ---------------------------------------------------------------------- axes

const valueAxis = (chart: Chart): ChartAxis => chart.axes.filter((axis) => axis.kind === 'val')[0];
const categoryAxis = (chart: Chart): ChartAxis => chart.axes.filter((axis) => axis.kind === 'cat')[0];
const withAxes = (category: Bytes[], value: Bytes[]): Chart => chartOf({ series: [sales], axes: [block(axisRec(0), ...category), block(axisRec(1), ...value)] });

test('xls chart: the range, the unit and the direction of a value axis', () => {
  const automatic = valueAxis(withAxes([catSerRange()], [valueRange()]));
  assert.deepEqual([automatic.min, automatic.max, automatic.majorUnit, automatic.reversed, automatic.logBase], [undefined, undefined, undefined, undefined, undefined]);
  const stated = valueAxis(withAxes([catSerRange()], [valueRange({ min: -10, max: 120, major: 20, reversed: true })]));
  assert.deepEqual([stated.min, stated.max, stated.majorUnit, stated.reversed], [-10, 120, 20, true]);
  // UNSURE: a logarithmic axis states the exponents of its limits.
  const log = valueAxis(withAxes([catSerRange()], [valueRange({ min: 0, max: 3, log: true })]));
  assert.deepEqual([log.logBase, log.min, log.max], [10, 1, 1000]);
});

test('xls chart: where the axes cross, and whether the points stand between the tick marks', () => {
  const plain = withAxes([catSerRange(1)], [valueRange()]);
  assert.deepEqual([valueAxis(plain).crossBetween, valueAxis(plain).crosses, categoryAxis(plain).crosses, categoryAxis(plain).crossesAt], ['between', undefined, undefined, undefined]);
  assert.equal(valueAxis(withAxes([catSerRange(0)], [valueRange()])).crossBetween, 'midCat');
  // The category axis says where the value axis crosses it, the value axis where the category axis does.
  assert.equal(valueAxis(withAxes([catSerRange(3)], [valueRange()])).crosses, 'max');
  assert.equal(valueAxis(withAxes([catSerRange(1, 3)], [valueRange()])).crossesAt, 3);
  assert.equal(categoryAxis(withAxes([catSerRange()], [valueRange({ maxCross: true })])).crosses, 'max');
  assert.equal(categoryAxis(withAxes([catSerRange()], [valueRange({ cross: 40 })])).crossesAt, 40);
  assert.equal(categoryAxis(withAxes([catSerRange(5)], [valueRange()])).reversed, true);
});

test('xls chart: the labels, the lines and the number format of an axis', () => {
  const head = cat(defaultStyles(), format(164, '0.0%'));
  const chart = chartOf(
    {
      series: [sales],
      axes: [
        block(axisRec(0), catSerRange(), axcExt(0x10, 1), tick({ labels: 1, color: '112233', rotation: 2 }), fontX(5), axisLine(0), lineFormat({ style: 5, axisOn: true })),
        block(axisRec(1), valueRange(), ifmt(164), tick({ labels: 0 }), axisLine(0), lineFormat({ color: '445566', weight: 1, style: 1, axisOn: true }), axisLine(1), lineFormat({ color: 'C0C0C0', weight: -1 }), axisLine(2), lineFormat({ auto: true })),
      ],
    },
    cat(head, font({ name: 'Tahoma', size: 9, bold: true })),
  );
  const category = categoryAxis(chart);
  const value = valueAxis(chart);
  assert.deepEqual([category.tickLabels, category.labelRotation, category.line], ['low', -90, { color: null }]);
  assert.equal(category.labelStyle && category.labelStyle.color, '#112233');
  assert.deepEqual([category.date, category.baseTimeUnit], [true, 'months']);
  assert.equal(value.tickLabels, 'none');
  assert.deepEqual([category.deleted, value.deleted], [undefined, undefined], 'both axes are on');
  assert.deepEqual([value.formatCode, value.sourceLinked], ['0.0%', false]);
  assert.deepEqual(value.line, { color: '#445566', width: 2, dash: true });
  assert.deepEqual(value.majorGridlines, { color: '#C0C0C0', width: 0.5 });
  assert.deepEqual(value.minorGridlines, {});
  assert.equal(category.majorGridlines, undefined);
});

test('xls chart: a series axis is not drawn and is noted', () => {
  const chart = chartOf({ series: [sales], axes: [...defaultAxes(), block(axisRec(2), catSerRange(), tick())] });
  assert.equal(chart.axes.length, 2);
  assert.ok(chart.notes.some((note) => /series axis/.test(note)), JSON.stringify(chart.notes));
});

// ------------------------------------------------- titles, legend and frames

test('xls chart: the title of the chart and the titles of its axes, with their fonts', () => {
  const fonts = cat(defaultStyles(), font({ name: 'Tahoma', size: 14, bold: true, italic: true }));
  const chart = chartOf(
    {
      series: [sales],
      axes: [...defaultAxes(), textBlock(objectLink(3), 'Months'), textBlock(objectLink(2), 'Amount', { color: '0000FF' })],
      tail: [textBlock(objectLink(1), 'Sales of the year', { font: 5, color: 'FF0000' })],
    },
    fonts,
  );
  assert.deepEqual(chart.title, { text: 'Sales of the year', style: { size: 14, bold: true, italic: true, font: 'Tahoma', color: '#FF0000' } });
  assert.equal(chart.autoTitleDeleted, undefined);
  assert.deepEqual(categoryAxis(chart).title, { text: 'Months' });
  assert.deepEqual(valueAxis(chart).title, { text: 'Amount', style: { color: '#0000FF' } });
});

test('xls chart: a title that is read from a cell, and one that was taken away', () => {
  const linked = chartOf({ series: [sales], tail: [textBlock(objectLink(1), ptg.ref3d(0, 0, 1))] });
  assert.deepEqual(linked.title, { ref: 'Data!$B$1' });
  const workbook = embeddedBook({ series: [sales], tail: [textBlock(objectLink(1), ptg.ref3d(0, 0, 1))] });
  const chart = ((workbook.sheets[0].drawings as NonNullable<Sheet['drawings']>)[0].object as DrawingChart).chart;
  assert.equal(chartValues(chart, { workbook }).title, 'Sales');
  const deleted = chartOf({ series: [sales], tail: [textBlock(objectLink(1), 'Gone', { flags: 0x40 })] });
  assert.deepEqual([deleted.title, deleted.autoTitleDeleted], [undefined, true]);
});

test('xls chart: the legend stands where the file puts it', () => {
  const at = (type: number) => chartOf({ series: [sales], groups: [{ type: bar(), children: [block(legendRec(type), block(textRec(), fontX(0)))] }] }).legend;
  assert.deepEqual([0, 1, 2, 3, 4].map((type) => (at(type) || { position: '' }).position), ['b', 'tr', 't', 'r', 'l']);
  assert.equal(chartOf({ series: [sales] }).legend, undefined);
  const moved = chartOf({ series: [sales], groups: [{ type: bar(), children: [block(legendRec(7))] }] });
  assert.equal(moved.legend && moved.legend.position, 'r');
  assert.ok(moved.notes.some((note) => /legend/.test(note)), JSON.stringify(moved.notes));
});

test('xls chart: the frame of the chart and of the plot area', () => {
  const chart = chartOf({ series: [sales], head: [frame({ color: '112233', weight: 2 }, { color: 'FFEEDD' })], axes: [block(axisRec(0), catSerRange()), block(axisRec(1), valueRange()), plotArea(), frame({ style: 5 }, { pattern: 0 })] });
  assert.deepEqual(chart.area, { fill: '#FFEEDD', line: { color: '#112233', width: 3 } });
  assert.deepEqual(chart.plotArea, { fill: null, line: { color: null } });
  // Frames that leave their look to the chart state nothing.
  const automatic = chartOf({ series: [sales], head: [frame()], axes: [block(axisRec(0), catSerRange()), block(axisRec(1), valueRange()), plotArea(), frame()] });
  assert.deepEqual([automatic.area, automatic.plotArea], [undefined, undefined]);
});

test('xls chart: the font of all text of the chart, and what stands for empty cells', () => {
  const fonts = cat(defaultStyles(), font({ name: 'Tahoma', size: 8 }));
  const chart = chartOf({ series: [sales], head: [frame(), defaultText(2), block(textRec(), fontX(5))], sheet: shtProps(false, 1) }, fonts);
  assert.deepEqual(chart.textStyle, { size: 8, font: 'Tahoma' });
  assert.deepEqual([chart.plotVisOnly, chart.blanks], [false, 'zero']);
  const other = chartOf({ series: [sales], sheet: shtProps(true, 2) });
  assert.deepEqual([other.plotVisOnly, other.blanks], [true, 'span']);
});

// ------------------------------------------------------- the look of series

test('xls chart: the colours of a series are the ones it states; automatic ones are left to the chart', () => {
  const stated = groupOf({ series: [{ ...sales, format: [lineFormat({ color: '102030', weight: 1 }), areaFormat({ color: 'A0B0C0', invert: true })] }] }).series[0];
  assert.deepEqual(stated.shape, { fill: '#A0B0C0', line: { color: '#102030', width: 2 } });
  assert.equal(stated.invertIfNegative, true);
  const automatic = groupOf({ series: [{ ...sales, format: [lineFormat({ auto: true, color: '102030' }), areaFormat({ auto: true, color: 'A0B0C0' })] }] }).series[0];
  assert.equal(automatic.shape, undefined);
  const hollow = groupOf({ series: [{ ...sales, format: [lineFormat({ style: 5 }), areaFormat({ pattern: 0 })] }] }).series[0];
  assert.deepEqual(hollow.shape, { fill: null, line: { color: null } });
  // The colours a series without any gets are those of the palette of the workbook: 24 and up for what is filled, 32 and up for lines.
  assert.deepEqual(chartOf({ series: [sales] }).palette, ['#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080']);
  assert.deepEqual(chartOf({ series: [sales], groups: [{ type: line() }] }).palette, ['#000080', '#FF00FF', '#FFFF00', '#00FFFF', '#800080', '#800000']);
});

test('xls chart: the line and the markers of a line series', () => {
  const series = groupOf({
    series: [{ ...sales, format: [lineFormat({ color: 'FF0000', style: 2, weight: 2 }), areaFormat({ color: '00FF00' }), markerFormat({ type: 2, line: '0000FF', fill: 'FFFF00', size: 140 }), serFmt(1)] }, { ...costs, format: [markerFormat({ type: 0 })] }, { ...costs, format: [markerFormat({ auto: true, type: 3 })] }],
    groups: [{ type: line() }],
  }).series;
  assert.deepEqual(series[0].shape, { line: { color: '#FF0000', width: 3, dash: true } }, 'the fill of a line series is not a fill of its line');
  assert.deepEqual(series[0].marker, { symbol: 'diamond', fill: '#FFFF00', line: { color: '#0000FF' }, size: 7 });
  assert.equal(series[0].smooth, true);
  assert.deepEqual(series[1].marker, { symbol: 'none' });
  assert.equal(series[2].marker, undefined, 'an automatic marker is left to the chart');
  const symbols = [1, 2, 3, 4, 5, 6, 7, 8, 9, 40].map((type) => groupOf({ series: [{ ...sales, format: [markerFormat({ type, noFill: true, noLine: true })] }], groups: [{ type: line() }] }).series[0].marker);
  assert.deepEqual(symbols.map((marker) => marker && marker.symbol), ['square', 'diamond', 'triangle', 'x', 'star', 'dot', 'dash', 'circle', 'plus', 'square']);
  assert.deepEqual(symbols[0], { symbol: 'square', fill: null, line: { color: null }, size: 5 });
});

test('xls chart: single points of a series: the slices of a pie', () => {
  const series = groupOf({
    series: [{ ...sales, format: [pieFormat(10)], points: [[2, [areaFormat({ color: '0000FF' }), pieFormat(25)]], [0, [areaFormat({ color: 'FF0000' })]], [1, [areaFormat({ auto: true })]]] }],
    axes: [],
    groups: [{ type: pie(), varied: true }],
  }).series[0];
  assert.equal(series.explosion, 10);
  assert.deepEqual(series.points, [{ index: 0, shape: { fill: '#FF0000' } }, { index: 2, shape: { fill: '#0000FF' }, explosion: 25 }]);
});

test('xls chart: the number of a series among the colours is the one its format states', () => {
  const series = groupOf({ series: [{ ...sales, number: 4 }, { ...costs, number: 2 }] }).series;
  assert.deepEqual(series.map((item) => [item.index, item.order]), [[4, 0], [2, 1]]);
});

test('xls chart: data labels of a series and of single points', () => {
  const attached = groupOf({ series: [{ ...sales, format: [attachedLabel(0x01)] }, { ...costs, format: [attachedLabel(0x04)] }, { ...costs, format: [attachedLabel(0x70)] }, { ...costs, format: [attachedLabel(0)] }] }).series;
  assert.deepEqual(attached.map((series) => series.dataLabels), [{ showValue: true }, { showPercent: true, showCategory: true }, { showCategory: true, showBubbleSize: true, showSeries: true }, undefined]);
  const linked = groupOf({
    series: [sales, costs],
    tail: [
      textBlock(objectLink(4, 0, 0xffff), undefined, { flags: 0x4 | 0x10, position: 1, color: '00FF00' }),
      textBlock(objectLink(4, 1, 2), 'Peak', { flags: 0 }),
      textBlock(objectLink(4, 1, 0), undefined, { flags: 0x40 | 0x10 }),
      textBlock(objectLink(4, 9, 0), 'No such series'),
    ],
  }).series;
  assert.deepEqual(linked[0].dataLabels, { showValue: true, position: 'outEnd', style: { color: '#00FF00' } });
  assert.deepEqual(linked[1].pointLabels, [{ index: 0, deleted: true }, { index: 2, text: 'Peak' }]);
});

test('xls chart: a trend line is no series of the chart; it is noted', () => {
  const chart = chartOf({
    series: [sales, { name: 'Linear (Sales)', tail: [serParent(1), serAuxTrend()] }, costs],
    cache: [siIndex(1), cachedNumber(0, 2, 77)],
  });
  assert.deepEqual(chart.groups[0].series.map((series) => series.name), [{ ref: 'Data!$B$1' }, { ref: 'Data!$C$1' }]);
  assert.ok(chart.notes.some((note) => /[Tt]rend/.test(note)), JSON.stringify(chart.notes));
  // The cache counts the series as the file does, the trend line included.
  assert.deepEqual(chartCacheList(chart.groups[0].series[1].values), [77]);
});

test('xls chart: records nothing is known of are named in a note; those of later versions in one line', () => {
  const chart = chartOf({ series: [sales], head: [frame(), rec(0x10f0, [1, 2]), rec(0x10f0, [3]), rec(0x10f1), rec(0x0850, [0, 0]), rec(0x089d, [0])] });
  assert.ok(chart.notes.some((note) => /not read: 0x10F0 \(2\), 0x10F1/.test(note)), JSON.stringify(chart.notes));
  assert.ok(chart.notes.some((note) => /later versions/.test(note)), JSON.stringify(chart.notes));
  assert.deepEqual(chartOf({ series: [sales] }).notes, []);
});

// ------------------------------------------------------------ what is not read

test('xls chart: a surface chart and a pie with a second plot stay placeholders, with a warning', () => {
  assert.match(problemOf({ series: [sales], groups: [{ type: rec(0x103f, u16(0)) }] }), /Chart "Chart 1" is not shown: a surface chart is not drawn/);
  assert.match(problemOf({ series: [sales], groups: [{ type: rec(0x1061, new Array(22).fill(0)) }] }), /second plot/);
});

test('xls chart: of an unreadable chart on a chart sheet the sheet says why and shows its note', () => {
  const stream = workbookStream({ tail: tail(), sheets: [{ name: 'Data', records: cells() }, { name: 'Chart1', type: 2, records: chartRecords({ series: [sales], groups: [{ type: rec(0x103f, u16(0)) }] }) }] });
  const sheet = readXlsStream(stream).workbook.sheets[1];
  assert.equal(sheet.kind, 'chartsheet');
  assert.equal(sheet.drawings, undefined);
  assert.deepEqual(sheet.warnings, ['The chart of the sheet is not shown: a surface chart is not drawn']);
});

// ------------------------------------------------------------------ writing

test('xls chart: saving as xlsx leaves the chart out and says so', async () => {
  const stream = workbookStream({ tail: tail(), sheets: [{ name: 'Data', records: cat(cells(), drawingRecords([shape({ kind: 201, ot: 5, id: 1, extra: [chartSubstream({ series: [sales] })] })])) }] });
  const workbook = readXlsStream(stream).workbook;
  const object = (workbook.sheets[0].drawings as NonNullable<Sheet['drawings']>)[0].object;
  assert.equal(object.type, 'chart');
  assert.equal(chartWritable((object as DrawingChart).chart), false);
  const report = { parts: [] as string[], warnings: [] as string[] };
  const bytes = await buildXlsx(workbook, { report });
  const names = ZipArchive.read(bytes).names();
  assert.deepEqual(names.filter((name) => /chart/i.test(name)), [], 'no chart part is written');
  assert.ok(report.warnings.some((warning) => /chart/.test(warning) && /left out/.test(warning)), JSON.stringify(report.warnings));
});

test('xls chart: a chart sheet is saved as an empty worksheet, its chart left out', async () => {
  const stream = workbookStream({ tail: tail(), sheets: [{ name: 'Data', records: cells() }, { name: 'Chart1', type: 2, records: chartRecords({ series: [sales] }) }] });
  const workbook = readXlsStream(stream).workbook;
  assert.equal(workbook.sheets[1].drawings && workbook.sheets[1].drawings[0].object.type, 'chart');
  const report = { parts: [] as string[], warnings: [] as string[] };
  const bytes = await buildXlsx(workbook, { report });
  assert.deepEqual(ZipArchive.read(bytes).names().filter((name) => /chart|drawing/i.test(name)), [], 'neither a chart nor a drawing is written');
  assert.ok(report.warnings.some((warning) => /chart sheet/.test(warning)), JSON.stringify(report.warnings));
  const again = await (await open(bytes)).workbook();
  assert.deepEqual(again.sheets.map((sheet) => [sheet.name, sheet.kind]), [['Data', 'worksheet'], ['Chart1', 'worksheet']]);
});

test('xls chart: an xls file with a chart is saved as the bytes it was', async () => {
  const stream = workbookStream({ tail: tail(), sheets: [{ name: 'Data', records: cat(cells(), drawingRecords([shape({ kind: 201, ot: 5, id: 1, extra: [chartSubstream({ series: [sales] })] })])) }] });
  const bytes = buildCfb([['Workbook', stream]]);
  const file = await open(bytes);
  const workbook = await file.workbook();
  assert.equal(workbook.sheets[0].drawings && workbook.sheets[0].drawings[0].object.type, 'chart');
  assert.deepEqual(await file.save(), bytes);
});
