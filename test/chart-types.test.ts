// Radar, bubble and stock charts, the pie of a pie with a second plot, and
// what lines carry besides their points: trend lines, error bars, drop lines,
// high-low lines, up-down bars and smoothing. Every chart part is generated
// by code; text is measured by a stand-in.
//
// Excel is not at hand: what follows its known behaviour without having been
// compared with it is marked UNSURE.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { MAX_DRAWN_ITEMS, chartValues, layoutChart, type ChartItem, type ChartLayout, type ChartScale } from '../src/layout/chart';
import type { Chart } from '../src/model/sheet/chart';
import { catAx, chartXml, fakeMeasure, group, legend, numLit, readChart, ser, solid, spPr, strLit, valAx, xyAxes, xySer } from './chart-fixtures';

type Of<K extends ChartItem['kind']> = Extract<ChartItem, { kind: K }>;

function model(xml: string): Chart {
  const chart = readChart(xml);
  assert.ok(chart, 'the chart is read');
  return chart as Chart;
}

const lay = (xml: string, width = 480, height = 320): ChartLayout => {
  const chart = model(xml);
  return layoutChart(chart, chartValues(chart), width, height, fakeMeasure);
};
const items = <K extends ChartItem['kind']>(layout: ChartLayout, kind: K, role?: string): Of<K>[] => layout.items.filter((item) => item.kind === kind && (role === undefined || item.role === role)) as Of<K>[];
const texts = (layout: ChartLayout, role: string): string[] => items(layout, 'text', role).map((item) => item.text);
const near = (actual: number, expected: number, message: string, tolerance = 1e-6) => assert.ok(Math.abs(actual - expected) <= tolerance, `${message}: ${actual} is not ${expected}`);
const bottom = (layout: ChartLayout) => layout.plot.y + layout.plot.height;
const valueY = (layout: ChartLayout, value: number): number => {
  const scale = layout.scale as ChartScale;
  return bottom(layout) - ((value - scale.min) / (scale.max - scale.min)) * layout.plot.height;
};

const CATS = strLit(['Kuzey', 'Doğu', 'Güney', 'Batı']);
const FIXED = (max: number, min = 0) => valAx('<c:majorGridlines/>', `<c:orientation val="minMax"/><c:max val="${max}"/><c:min val="${min}"/>`);

// --------------------------------------------------------------------- radar

const radar = (style: string, series: string, axes = catAx() + FIXED(40)) => chartXml({ plot: group('radarChart', `<c:radarStyle val="${style}"/>`, series) + axes });
const quarter = (values: (number | null)[], index = 0, props = '') => ser(index, { tx: `<c:v>S${index}</c:v>`, cat: CATS, val: numLit(values), props });

test('a radar chart: the categories around a centre from the top clockwise, the values outwards from it', () => {
  const layout = lay(radar('standard', quarter([10, 20, 30, 40])));
  const { x, y, width, height } = layout.plot;
  near(width, height, 'the plot area is the square around the circle');
  const cx = x + width / 2;
  const cy = y + height / 2;
  const radius = width / 2;
  const path = items(layout, 'path', 'line')[0];
  assert.ok(path.closed === true && path.fill === undefined, 'a line that comes back to its start, not filled');
  assert.equal(path.stroke, '#4472C4', 'in the colour of the series');
  const expected = [cx, cy - radius * 0.25, cx + radius * 0.5, cy, cx, cy + radius * 0.75, cx - radius, cy];
  path.points.forEach((value, i) => near(value, expected[i], `number ${i} of the outline`));
  assert.deepEqual([layout.scale?.min, layout.scale?.max], [0, 40]);
  // A ring for every tick above the centre, a spoke for every category.
  const rings = items(layout, 'path', 'grid');
  assert.equal(rings.length, (layout.scale as ChartScale).ticks.length - 1);
  assert.ok(rings.every((ring) => ring.closed === true && ring.points.length === 8), 'the rings are polygons through the spokes');
  near(rings[rings.length - 1].points[1], cy - radius, 'the outermost ring reaches the rim');
  const spokes = items(layout, 'path', 'axis');
  assert.equal(spokes.length, 4);
  for (const spoke of spokes) near(Math.hypot(spoke.points[2] - cx, spoke.points[3] - cy), radius, 'a spoke runs from the centre to the rim');
  // The labels of the categories stand outside the rim, those of the values along the first spoke.
  const around = items(layout, 'text', 'category');
  assert.deepEqual(around.map((item) => item.text), ['Kuzey', 'Doğu', 'Güney', 'Batı']);
  assert.ok(around[0].y < cy - radius && around[1].x > cx + radius && around[2].y > cy + radius && around[3].x < cx - radius, 'each beyond the end of its spoke');
  assert.deepEqual([around[0].align, around[1].align, around[3].align], ['center', 'left', 'right']);
  const ticks = items(layout, 'text', 'tick');
  assert.ok(ticks.length >= 3 && ticks.every((item) => item.x < cx && item.align === 'right'), 'value labels beside the spoke that points up');
  assert.ok(ticks.some((item) => item.text === '40'), 'the largest value of the scale is labelled');
  for (const item of layout.items) if (item.kind === 'text') assert.ok(item.x >= 0 && item.x <= layout.width && item.y >= 0 && item.y <= layout.height, `"${item.text}" lies in the frame`);
});

test('a radar chart: filled, with markers, with a point that is missing, with several series and a legend', () => {
  const filled = lay(radar('filled', quarter([10, 20, 30, 40])));
  const area = items(filled, 'path', 'area')[0];
  assert.ok(area && area.closed === true && area.fill === '#4472C4', 'filled in the colour of the series');
  const marked = lay(radar('marker', quarter([10, 20, 30, 40])));
  assert.equal(items(marked, 'markers', 'marker')[0].points.length, 8, 'a marker at every point');
  assert.equal(items(lay(radar('standard', quarter([10, 20, 30, 40]))), 'markers', 'marker').length, 0, 'none by default');
  const gap = lay(radar('standard', quarter([10, null, 30, 40])));
  const open = items(gap, 'path', 'line');
  assert.ok(open.length === 1 && open[0].closed !== true && open[0].points.length === 6, 'a point without a number breaks the outline: Güney, Batı and back to Kuzey');
  const two = lay(chartXml({ plot: group('radarChart', '<c:radarStyle val="standard"/>', quarter([10, 20, 30, 40]) + quarter([40, 30, 20, 10], 1)) + catAx() + FIXED(40), tail: legend('r') }));
  assert.equal(items(two, 'path', 'line').length, 2);
  assert.deepEqual(texts(two, 'legend'), ['S0', 'S1']);
  assert.equal(items(two, 'path', 'swatch').length, 2, 'the keys of the legend are lines');
  const labelled = lay(radar('standard', quarter([10, 20, 30, 40], 0, '<c:dLbls><c:showVal val="1"/></c:dLbls>')));
  assert.deepEqual(texts(labelled, 'label'), ['10', '20', '30', '40']);
});

test('a radar chart of very many categories, of none, in a small frame and beside another type', () => {
  const many = Array.from({ length: 5000 }, (_value, i) => i % 17);
  const big = lay(chartXml({ plot: group('radarChart', '', ser(0, { val: numLit(many) })) + catAx() + valAx() }));
  const path = items(big, 'path').filter((item) => item.role === 'line' || item.role === 'area');
  assert.ok(path.length >= 1 && path.every((item) => item.points.length <= 2 * 720 + 2), 'sampled down to what a circle can show');
  assert.ok(items(big, 'path', 'axis').length <= 360 && items(big, 'text', 'category').length <= 90, 'spokes and labels are thinned');
  assert.doesNotThrow(() => lay(chartXml({ plot: group('radarChart', '', ser(0, { val: numLit([]) })) + catAx() + valAx() })));
  const tiny = lay(radar('standard', quarter([10, 20, 30, 40])), 60, 50);
  assert.equal(items(tiny, 'text').length, 0, 'no words in a tiny frame');
  assert.equal(items(tiny, 'path', 'line').length, 1, 'the series alone');
  const mixed = lay(chartXml({ plot: group('radarChart', '', quarter([10, 20, 30, 40])) + group('barChart', '<c:barDir val="col"/>', quarter([1, 2, 3, 4], 1)) + catAx() + valAx() }));
  assert.equal(items(mixed, 'rect', 'bar').length, 0, 'a radar beside another type: the radar alone');
  const noted = model(chartXml({ plot: group('radarChart', '', quarter([10, 20, 30, 40])) + group('barChart', '<c:barDir val="col"/>', quarter([1, 2, 3, 4], 1)) + catAx() + valAx() })).notes;
  assert.ok(noted.some((note) => /radar combined with another chart type/.test(note)), 'which is said in the notes of the chart');
});

// -------------------------------------------------------------------- bubble

const bubbles = (sizes: string, tail = '', props = '') =>
  chartXml({
    plot:
      group('bubbleChart', '<c:varyColors val="0"/>', `<c:ser><c:idx val="0"/><c:order val="0"/>${props}<c:xVal>${numLit([1, 2, 3])}</c:xVal><c:yVal>${numLit([10, 20, 30])}</c:yVal>${sizes}</c:ser>`, tail) +
      '<c:valAx><c:axId val="111"/><c:scaling><c:max val="4"/><c:min val="0"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222"/></c:valAx>' +
      '<c:valAx><c:axId val="222"/><c:scaling><c:max val="40"/><c:min val="0"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111"/></c:valAx>',
  });
const sized = (values: (number | null)[]) => `<c:bubbleSize>${numLit(values)}</c:bubbleSize>`;

test('a bubble chart: circles at their x and y, as large as their sizes say', () => {
  const layout = lay(bubbles(sized([1, 4, 9])));
  const circles = items(layout, 'circles', 'bubble');
  assert.equal(circles.length, 1);
  const [x1, y1, r1, x2, y2, r2, x3, y3, r3] = circles[0].points;
  near(x1, layout.plot.x + layout.plot.width / 4, 'x = 1 of 0..4');
  near(y1, valueY(layout, 10), 'y = 10');
  near(x3, layout.plot.x + (layout.plot.width * 3) / 4, 'x = 3');
  near(y2, valueY(layout, 20), 'y = 20');
  near(x2, layout.plot.x + layout.plot.width / 2, 'x = 2');
  near(y3, valueY(layout, 30), 'y = 30');
  // UNSURE: the largest bubble is a quarter of the smaller side of the plot area across.
  near(r3 * 2, Math.min(layout.plot.width, layout.plot.height) * 0.25, 'the largest bubble');
  near(r1 / r3, 1 / 3, 'sizes are areas: a ninth of the size is a third of the radius');
  near(r2 / r3, 2 / 3, 'four ninths: two thirds');
  assert.equal(circles[0].fill, '#4472C4');
  assert.deepEqual([layout.xScale?.max, layout.scale?.max], [4, 40]);
});

test('bubbles: sizes as widths, the scale of the bubbles, sizes that are missing, nothing or below zero', () => {
  const widths = items(lay(bubbles(sized([1, 4, 9]), '<c:sizeRepresents val="w"/>')), 'circles')[0].points;
  near(widths[2] / widths[8], 1 / 9, 'sizes are widths: a ninth of the size is a ninth of the radius');
  const half = lay(bubbles(sized([1, 4, 9]), '<c:bubbleScale val="50"/>'));
  near(items(half, 'circles')[0].points[8] * 2, Math.min(half.plot.width, half.plot.height) * 0.125, 'bubbleScale 50: half as large');
  const none = items(lay(bubbles('')), 'circles')[0].points;
  assert.ok(none.length === 9 && none[2] === none[5] && none[5] === none[8] && none[2] > 0, 'without sizes every bubble is the same');
  const partly = items(lay(bubbles(sized([4, null, 0]))), 'circles')[0].points;
  assert.equal(partly.length, 3, 'a bubble without a size, or of the size nothing, is not drawn');
  const negative = lay(bubbles(sized([4, -9, 1])));
  assert.equal(items(negative, 'circles')[0].points.length, 6, 'a bubble of a negative size is not drawn');
  near(items(negative, 'circles')[0].points[2] * 2, Math.min(negative.plot.width, negative.plot.height) * 0.25, 'and does not count for the largest');
  const shown = lay(bubbles(sized([4, -9, 1]), '<c:showNegBubbles val="1"/>'));
  const all = items(shown, 'circles');
  assert.equal(all.reduce((sum, item) => sum + item.points.length / 3, 0), 3, 'unless the chart shows them');
  const hollow = all.filter((item) => item.fill === '#FFFFFF')[0];
  assert.ok(hollow && hollow.stroke === '#4472C4' && hollow.points.length === 3, 'UNSURE: by the size without its sign, white with the outline of the series');
  const labelled = lay(bubbles(sized([1, 4, 9]), '', '<c:dLbls><c:showBubbleSize val="1"/></c:dLbls>'));
  assert.deepEqual(texts(labelled, 'label'), ['1', '4', '9']);
});

test('bubbles are bounded: of very many only what a frame is worth, and every number is finite', () => {
  const count = 60000;
  const numbers = Array.from({ length: count }, (_value, i) => i);
  const xml = chartXml({ plot: group('bubbleChart', '', `<c:ser><c:idx val="0"/><c:order val="0"/><c:xVal>${numLit(numbers)}</c:xVal><c:yVal>${numLit(numbers)}</c:yVal><c:bubbleSize>${numLit(numbers)}</c:bubbleSize></c:ser>`) + xyAxes() });
  const layout = lay(xml);
  const points = items(layout, 'circles').reduce((sum, item) => sum + item.points.length / 3, 0);
  assert.ok(points > 100 && points <= MAX_DRAWN_ITEMS / 4, `${points} bubbles are drawn`);
  for (const item of items(layout, 'circles')) for (const value of item.points) assert.ok(Number.isFinite(value), 'a finite number');
});

// --------------------------------------------------------------------- stock

const hidden = spPr('<a:ln w="28575"><a:noFill/></a:ln>') + '<c:marker><c:symbol val="none"/></c:marker>';
const stockSeries = (rows: number[][], props = hidden) => rows.map((values, i) => ser(i, { cat: strLit(['Pzt', 'Sal', 'Çar']), val: numLit(values), props })).join('');
const OHLC = [
  [10, 30, 20],
  [35, 32, 28],
  [5, 12, 15],
  [30, 15, 20],
];

test('a stock chart: a line from the highest to the lowest value of each day, a box from the open to the close', () => {
  const layout = lay(chartXml({ plot: group('stockChart', '', stockSeries(OHLC), '<c:hiLowLines/><c:upDownBars><c:gapWidth val="100"/><c:upBars/><c:downBars/></c:upDownBars>') + catAx() + FIXED(40) }));
  const slot = layout.plot.width / 3;
  const centre = (i: number) => layout.plot.x + (i + 0.5) * slot;
  const lines = items(layout, 'path', 'hilow');
  assert.equal(lines.length, 1);
  assert.equal(lines[0].segments, true, 'one item of separate lines');
  const expected = [centre(0), valueY(layout, 35), centre(0), valueY(layout, 5), centre(1), valueY(layout, 32), centre(1), valueY(layout, 12), centre(2), valueY(layout, 28), centre(2), valueY(layout, 15)];
  assert.equal(lines[0].points.length, expected.length);
  lines[0].points.forEach((value, i) => near(value, expected[i], `number ${i} of the high-low lines`, 0.51));
  const boxes = items(layout, 'rect', 'updown');
  assert.equal(boxes.length, 3);
  near(boxes[0].width, slot / 2, 'gap width 100: half of the room of a category');
  near(boxes[0].x + boxes[0].width / 2, centre(0), 'centred on the category');
  near(boxes[0].y, valueY(layout, 30), 'the first day rose from 10 to 30: from the close down');
  near(boxes[0].height, valueY(layout, 10) - valueY(layout, 30), 'to the open');
  near(boxes[1].y, valueY(layout, 30), 'the second fell from 30 to 15: from the open down');
  // UNSURE: a day that rose is white, a day that fell is black, both outlined, where the file states no look.
  assert.deepEqual([boxes[0].fill, boxes[1].fill, boxes[0].stroke], ['#FFFFFF', '#000000', '#000000']);
  near(boxes[2].height, 0, 'open and close alike: a box without height');
  assert.ok(layout.items.indexOf(lines[0]) < layout.items.indexOf(boxes[0]), 'the boxes are drawn over the high-low lines, not under them');
  assert.equal(items(layout, 'path', 'line').length, 0, 'the series themselves have no line where the file takes it away');
  assert.equal(items(layout, 'markers', 'marker').length, 0);
});

test('a stock chart of three series (high, low, close): markers as the file states them, colours of the boxes from the file', () => {
  const close = spPr('<a:ln><a:noFill/></a:ln>') + `<c:marker><c:symbol val="dash"/><c:size val="7"/>${spPr(solid('FF0000'))}</c:marker>`;
  const series = stockSeries(OHLC.slice(1, 3)) + ser(2, { cat: strLit(['Pzt', 'Sal', 'Çar']), val: numLit(OHLC[3]), props: close });
  const layout = lay(chartXml({ plot: group('stockChart', '', series, `<c:hiLowLines>${spPr(`<a:ln w="19050">${solid('00AA00')}</a:ln>`)}</c:hiLowLines>`) + catAx() + FIXED(40) }));
  const line = items(layout, 'path', 'hilow')[0];
  assert.deepEqual([line.stroke, line.lineWidth], ['#00AA00', 2]);
  const marks = items(layout, 'markers', 'marker');
  assert.equal(marks.length, 1);
  assert.deepEqual([marks[0].symbol, marks[0].points.length], ['dash', 6]);
  near(marks[0].points[1], valueY(layout, 30), 'the close of the first day');
  const coloured = lay(chartXml({ plot: group('stockChart', '', stockSeries(OHLC), `<c:upDownBars><c:upBars>${spPr(solid('00B050'))}</c:upBars><c:downBars>${spPr(solid('FF0000'))}</c:downBars></c:upDownBars>`) + catAx() + FIXED(40) }));
  assert.deepEqual(items(coloured, 'rect', 'updown').slice(0, 2).map((item) => item.fill), ['#00B050', '#FF0000']);
  const bare = lay(chartXml({ plot: group('stockChart', '', stockSeries(OHLC, '')) + catAx() + FIXED(40) }));
  assert.equal(items(bare, 'path', 'line').length, 0, 'UNSURE: a series of a stock chart has no line unless the file gives it one');
});

// ---------------------------------------------------------- lines on a line

const lineChart = (series: string, tail = '<c:marker val="1"/>', axes = catAx() + FIXED(40)) => chartXml({ plot: group('lineChart', '<c:grouping val="standard"/>', series, tail) + axes });
const three = (values: (number | null)[], index = 0, props = '', tail = '') => ser(index, { cat: strLit(['a', 'b', 'c']), val: numLit(values), props, tail });

test('drop lines, high-low lines and up-down bars of a line chart', () => {
  const layout = lay(lineChart(three([10, 30, 20]) + three([20, 10, 20], 1), `<c:dropLines>${spPr(`<a:ln>${solid('AA0000')}</a:ln>`)}</c:dropLines><c:hiLowLines/><c:upDownBars/><c:marker val="1"/>`));
  const slot = layout.plot.width / 3;
  const drop = items(layout, 'path', 'drop')[0];
  assert.equal(drop.stroke, '#AA0000');
  assert.equal(drop.points.length, 12);
  near(drop.points[1], valueY(layout, 20), 'from the highest point of the first category');
  near(drop.points[3], bottom(layout), 'down to the category axis', 0.51);
  near(drop.points[5], valueY(layout, 30), 'the highest of the second');
  const hiLow = items(layout, 'path', 'hilow')[0];
  near(hiLow.points[5], valueY(layout, 30), 'high of the second');
  near(hiLow.points[7], valueY(layout, 10), 'low of the second');
  const boxes = items(layout, 'rect', 'updown');
  assert.equal(boxes.length, 3);
  // Gap width 150 by default: a box takes 1 / 2.5 of the room of its category.
  near(boxes[0].width, slot / 2.5, 'the width of a box');
  assert.deepEqual([boxes[0].fill, boxes[1].fill], ['#FFFFFF', '#000000'], 'the last series above the first: up; below: down');
  const area = lay(chartXml({ plot: group('areaChart', '<c:grouping val="standard"/>', three([10, 30, 20]), '<c:dropLines/>') + catAx() + FIXED(40) }));
  assert.equal(items(area, 'path', 'drop')[0].points.length, 12, 'an area chart has drop lines too');
  const gappy = lay(lineChart(three([10, null, 20]) + three([null, null, 20], 1), '<c:dropLines/><c:hiLowLines/><c:upDownBars/>'));
  assert.equal(items(gappy, 'path', 'drop')[0].points.length, 8, 'no drop line where no series has a point');
  assert.equal(items(gappy, 'rect', 'updown').length, 1, 'a box needs both ends');
});

test('a smoothed line is a curve through its points', () => {
  const layout = lay(lineChart(three([10, 30, 20], 0, '', '<c:smooth val="1"/>') + three([5, 5, 5], 1)));
  const [curve, straight] = items(layout, 'path', 'line');
  assert.equal(curve.curve, true);
  assert.equal(curve.points.length, 2 + 2 * 6, 'the first point and three more numbers pairs for each further one');
  const slot = layout.plot.width / 3;
  near(curve.points[0], layout.plot.x + slot / 2, 'starting at the first point');
  near(curve.points[1], valueY(layout, 10), 'at its value');
  near(curve.points[6], layout.plot.x + slot * 1.5, 'through the second point');
  near(curve.points[7], valueY(layout, 30), 'at its value');
  near(curve.points[13], valueY(layout, 20), 'ending at the third');
  assert.ok(straight.curve !== true && straight.points.length === 6, 'the other series stays straight');
  for (let i = 0; i < curve.points.length; i += 2) {
    assert.ok(curve.points[i] >= 0 && curve.points[i] <= layout.width && curve.points[i + 1] >= 0 && curve.points[i + 1] <= layout.height, 'every point that shapes the curve lies in the frame');
  }
  const scatter = lay(chartXml({ plot: group('scatterChart', '<c:scatterStyle val="smoothMarker"/>', xySer(0, { cat: numLit([1, 2, 3]), val: numLit([1, 9, 2]), tail: '<c:smooth val="1"/>' })) + xyAxes() }));
  assert.equal(items(scatter, 'path', 'line')[0].curve, true, 'a scatter series is smoothed where it says so');
  const two = lay(lineChart(three([10, 30, null], 0, '', '<c:smooth val="1"/>')));
  assert.ok(items(two, 'path', 'line')[0].curve !== true, 'two points are joined by a straight line');
});

// --------------------------------------------------------------- trend lines

const five = (trend: string, values = [2, 4, 5, 4, 5]) => ser(0, { cat: strLit(['a', 'b', 'c', 'd', 'e']), val: numLit(values), props: trend });
const trendOf = (inner: string) => `<c:trendline>${inner}</c:trendline>`;
const columns = (series: string, axes = catAx() + FIXED(10)) => chartXml({ plot: group('barChart', '<c:barDir val="col"/><c:grouping val="clustered"/>', series) + axes });

test('a linear trend line runs through the categories at the values of least squares', () => {
  const layout = lay(columns(five(trendOf('<c:trendlineType val="linear"/>'))));
  const slot = layout.plot.width / 5;
  const trend = items(layout, 'path', 'trend');
  assert.equal(trend.length, 1);
  const points = trend[0].points;
  // Slope 0.6, intercept 2.2 (see chart-fit.test.ts): 2.8 at the first category, 5.2 at the fifth.
  near(points[0], layout.plot.x + slot / 2, 'starting at the first category');
  near(points[1], valueY(layout, 2.8), 'at 2.8');
  near(points[points.length - 2], layout.plot.x + slot * 4.5, 'ending at the fifth');
  near(points[points.length - 1], valueY(layout, 5.2), 'at 5.2');
  // UNSURE: a trend line the file gives no look is dashed, two pixels wide, in the colour of its series.
  assert.deepEqual([trend[0].stroke, trend[0].dash, trend[0].lineWidth], ['#4472C4', true, 2]);
  const styled = lay(columns(five(trendOf(`${spPr(`<a:ln w="28575">${solid('FF0000')}</a:ln>`)}<c:trendlineType val="linear"/>`))));
  assert.deepEqual([items(styled, 'path', 'trend')[0].stroke, items(styled, 'path', 'trend')[0].dash, items(styled, 'path', 'trend')[0].lineWidth], ['#FF0000', undefined, 3]);
});

test('a trend line goes on forward and backward as far as the plot area reaches', () => {
  const layout = lay(columns(five(trendOf('<c:trendlineType val="linear"/><c:forward val="2"/><c:backward val="0.25"/>'))));
  const slot = layout.plot.width / 5;
  const points = items(layout, 'path', 'trend')[0].points;
  near(points[0], layout.plot.x + slot * 0.25, 'a quarter of a category before the first');
  near(points[1], valueY(layout, 2.2 + 0.6 * 0.75), 'at the value of the line there');
  // UNSURE: the categories are not made more for a trend line that goes beyond them: it ends at the edge.
  near(points[points.length - 2], layout.plot.x + layout.plot.width, 'forward: to the edge of the plot area');
  near(points[points.length - 1], valueY(layout, 2.2 + 0.6 * 5.5), 'at the value of the line there');
});

test('curved trend lines, a moving average, and a trend line that leaves the scale', () => {
  const squares = [1, 4, 9, 16, 25];
  const poly = lay(columns(five(trendOf('<c:trendlineType val="poly"/><c:order val="2"/>'), squares), catAx() + FIXED(30)));
  const curve = items(poly, 'path', 'trend')[0].points;
  assert.ok(curve.length >= 2 * 20, 'a curve is drawn through many points');
  const slot = poly.plot.width / 5;
  // Halfway between the second and the third category x is 2.5: 6.25.
  let closest = 0;
  for (let i = 0; i < curve.length; i += 2) if (Math.abs(curve[i] - (poly.plot.x + slot * 2)) < Math.abs(curve[closest] - (poly.plot.x + slot * 2))) closest = i;
  near(curve[closest + 1], valueY(poly, Math.pow((curve[closest] - poly.plot.x) / slot + 0.5, 2)), 'the parabola through the squares', 0.01);
  const moving = lay(columns(five(trendOf('<c:trendlineType val="movingAvg"/><c:period val="2"/>'), squares), catAx() + FIXED(30)));
  const average = items(moving, 'path', 'trend')[0].points;
  assert.equal(average.length, 8, 'from the second point on');
  near(average[0], moving.plot.x + slot * 1.5, 'the first average at the second category');
  near(average[1], valueY(moving, 2.5), 'the mean of 1 and 4');
  near(average[7], valueY(moving, 20.5), 'the mean of 16 and 25');
  // Cut at the top of the scale: 25 does not fit an axis that ends at 10.
  const cut = lay(columns(five(trendOf('<c:trendlineType val="poly"/><c:order val="2"/>'), squares), catAx() + FIXED(10)));
  for (const item of items(cut, 'path', 'trend')) {
    for (let i = 1; i < item.points.length; i += 2) assert.ok(item.points[i] >= cut.plot.y - 1e-6 && item.points[i] <= bottom(cut) + 1e-6, 'no point of the trend line outside the plot area');
  }
  assert.equal(items(lay(columns(five(trendOf('<c:trendlineType val="exp"/>'), [1, 2, 0, 4, 5]))), 'path', 'trend').length, 0, 'no exponential trend of a series with a zero');
  assert.equal(items(lay(columns(five(trendOf('<c:trendlineType val="poly"/><c:order val="6"/>'), [1, 2, 3, 4, 5]))), 'path', 'trend').length, 0, 'no polynomial of a degree the points do not give');
});

test('trend lines of a scatter series follow its x values; stacked series and pies have none; their number is bounded', () => {
  const scatter = lay(
    chartXml({
      plot:
        group('scatterChart', '<c:scatterStyle val="marker"/>', xySer(0, { cat: numLit([10, 20, 30]), val: numLit([1, 2, 3]), props: trendOf('<c:trendlineType val="linear"/>') })) +
        '<c:valAx><c:axId val="111"/><c:scaling><c:max val="40"/><c:min val="0"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222"/></c:valAx>' +
        '<c:valAx><c:axId val="222"/><c:scaling><c:max val="4"/><c:min val="0"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111"/></c:valAx>',
    }),
  );
  const points = items(scatter, 'path', 'trend')[0].points;
  near(points[0], scatter.plot.x + scatter.plot.width / 4, 'from x = 10');
  near(points[1], valueY(scatter, 1), 'y = 1');
  near(points[points.length - 2], scatter.plot.x + (scatter.plot.width * 3) / 4, 'to x = 30');
  near(points[points.length - 1], valueY(scatter, 3), 'y = 3');
  const stacked = lay(chartXml({ plot: group('barChart', '<c:barDir val="col"/><c:grouping val="stacked"/>', five(trendOf('<c:trendlineType val="linear"/>'))) + catAx() + valAx() }));
  assert.equal(items(stacked, 'path', 'trend').length, 0, 'none on what is stacked');
  const pie = lay(chartXml({ plot: group('pieChart', '', five(trendOf('<c:trendlineType val="linear"/>')), '', false) }));
  assert.equal(items(pie, 'path', 'trend').length, 0);
  const many = Array.from({ length: 40 }, (_value, i) => ser(i, { val: numLit([1, 2 + i, 3, 5]), props: trendOf('<c:trendlineType val="linear"/>').repeat(8) })).join('');
  const crowded = lay(chartXml({ plot: group('lineChart', '<c:grouping val="standard"/>', many) + catAx() + valAx() }));
  assert.ok(items(crowded, 'path', 'trend').length <= 64, `${items(crowded, 'path', 'trend').length} trend lines of 320`);
});

// ---------------------------------------------------------------- error bars

const err = (inner: string) => `<c:errBars>${inner}</c:errBars>`;
const errChart = (bars: string, values = [10, 20, 30]) => lay(lineChart(three(values, 0, bars)));
/** The vertical lines of the error bars of an upright chart: [x, from, to] of each. */
const uprights = (layout: ChartLayout): number[][] => {
  const out: number[][] = [];
  for (const item of items(layout, 'path', 'error')) {
    assert.equal(item.segments, true, 'error bars are separate lines');
    for (let i = 0; i + 3 < item.points.length; i += 4) if (Math.abs(item.points[i] - item.points[i + 2]) < 1e-9) out.push([item.points[i], item.points[i + 1], item.points[i + 3]]);
  }
  return out;
};

test('error bars of a fixed amount: up, down or both, with caps unless the file takes them away', () => {
  const both = errChart(err('<c:errBarType val="both"/><c:errValType val="fixedVal"/><c:val val="5"/>'));
  const lines = uprights(both);
  assert.equal(lines.length, 3);
  const slot = both.plot.width / 3;
  near(lines[1][0], both.plot.x + slot * 1.5, 'at the point');
  near(Math.min(lines[1][1], lines[1][2]), valueY(both, 25), 'up to 25');
  near(Math.max(lines[1][1], lines[1][2]), valueY(both, 15), 'down to 15');
  const item = items(both, 'path', 'error')[0];
  assert.equal(item.points.length, 3 * 12, 'a line and two caps for each point');
  const caps = item.points.length / 4 - lines.length;
  assert.equal(caps, 6);
  const bare = errChart(err('<c:errBarType val="both"/><c:errValType val="fixedVal"/><c:noEndCap val="1"/><c:val val="5"/>'));
  assert.equal(items(bare, 'path', 'error')[0].points.length, 3 * 4, 'no caps');
  const plus = uprights(errChart(err('<c:errBarType val="plus"/><c:errValType val="fixedVal"/><c:noEndCap val="1"/><c:val val="5"/>')));
  near(Math.max(plus[0][1], plus[0][2]), valueY(both, 10), 'plus: from the point');
  near(Math.min(plus[0][1], plus[0][2]), valueY(both, 15), 'up');
  const minus = uprights(errChart(err('<c:errBarType val="minus"/><c:errValType val="fixedVal"/><c:noEndCap val="1"/><c:val val="5"/>')));
  near(Math.max(minus[0][1], minus[0][2]), valueY(both, 5), 'minus: down');
  const styled = errChart(err(`<c:errBarType val="both"/><c:errValType val="fixedVal"/><c:val val="5"/>${spPr(`<a:ln w="19050">${solid('AA00AA')}</a:ln>`)}`));
  assert.deepEqual([items(styled, 'path', 'error')[0].stroke, items(styled, 'path', 'error')[0].lineWidth], ['#AA00AA', 2]);
  assert.equal(items(both, 'path', 'error')[0].stroke, '#595959', 'grey where the file states no look');
});

test('error bars in per cent, of the standard deviation, of the standard error, and with amounts of their own', () => {
  const percent = uprights(errChart(err('<c:errBarType val="both"/><c:errValType val="percentage"/><c:noEndCap val="1"/><c:val val="10"/>')));
  const scale = errChart('');
  near(Math.min(percent[2][1], percent[2][2]), valueY(scale, 33), 'ten per cent of 30 above it');
  near(Math.max(percent[0][1], percent[0][2]), valueY(scale, 9), 'ten per cent of 10 below it');
  // 10, 20, 30: mean 20, deviation of the sample 10.
  const deviation = uprights(errChart(err('<c:errBarType val="both"/><c:errValType val="stdDev"/><c:noEndCap val="1"/><c:val val="1"/>')));
  for (const line of deviation) {
    // UNSURE: bars of the standard deviation stand around the mean of the series at every point, not around the point.
    near(Math.min(line[1], line[2]), valueY(scale, 30), 'the mean and one deviation');
    near(Math.max(line[1], line[2]), valueY(scale, 10), 'the mean less one deviation');
  }
  const half = uprights(errChart(err('<c:errBarType val="both"/><c:errValType val="stdDev"/><c:noEndCap val="1"/><c:val val="0.5"/>')));
  near(Math.min(half[0][1], half[0][2]), valueY(scale, 25), 'half a deviation');
  // The standard error: 10 / √3 = 5.7735...
  const error = uprights(errChart(err('<c:errBarType val="both"/><c:errValType val="stdErr"/><c:noEndCap val="1"/>')));
  near(Math.min(error[0][1], error[0][2]), valueY(scale, 10 + 10 / Math.sqrt(3)), 'UNSURE: the deviation over the root of the number of points, around each point');
  near(Math.max(error[2][1], error[2][2]), valueY(scale, 30 - 5.773502691896258), 'the same below the last point');
  const custom = uprights(errChart(err(`<c:errBarType val="both"/><c:errValType val="cust"/><c:noEndCap val="1"/><c:plus>${numLit([1, 2, null])}</c:plus><c:minus>${numLit([4, 5, 6])}</c:minus>`)));
  near(Math.min(custom[0][1], custom[0][2]), valueY(scale, 11), 'its own amount up');
  near(Math.max(custom[1][1], custom[1][2]), valueY(scale, 15), 'its own amount down');
  near(Math.min(custom[2][1], custom[2][2]), valueY(scale, 30), 'a point without an amount up has no bar up');
  near(Math.max(custom[2][1], custom[2][2]), valueY(scale, 24), 'but the one down');
});

test('error bars of columns, of bars that lie, along x of a scatter series, and cut at the ends of the axis', () => {
  const fixed = err('<c:errBarType val="both"/><c:errValType val="fixedVal"/><c:noEndCap val="1"/><c:val val="5"/>');
  const cols = lay(columns(ser(0, { cat: strLit(['a', 'b', 'c']), val: numLit([2, 5, 8]), props: fixed }), catAx() + FIXED(10)));
  const lines = uprights(cols);
  const bars = items(cols, 'rect', 'bar');
  near(lines[0][0], bars[0].x + bars[0].width / 2, 'in the middle of the column');
  near(Math.max(lines[0][1], lines[0][2]), bottom(cols), '2 - 5 is below the axis: cut at its end');
  near(Math.min(lines[2][1], lines[2][2]), cols.plot.y, '8 + 5 is above it: cut');
  const lying = lay(chartXml({ plot: group('barChart', '<c:barDir val="bar"/><c:grouping val="clustered"/>', ser(0, { cat: strLit(['a', 'b', 'c']), val: numLit([10, 5, 8]), props: fixed })) + catAx() + FIXED(20) }));
  const flat = items(lying, 'path', 'error')[0].points;
  near(flat[1], flat[3], 'bars that lie have error bars that lie');
  near(Math.abs(flat[2] - flat[0]), lying.plot.width * 0.5, 'ten of twenty across');
  const scatter = lay(
    chartXml({
      plot:
        group('scatterChart', '<c:scatterStyle val="marker"/>', xySer(0, { cat: numLit([10, 20, 30]), val: numLit([1, 2, 3]), props: err('<c:errDir val="x"/><c:errBarType val="both"/><c:errValType val="fixedVal"/><c:noEndCap val="1"/><c:val val="5"/>') + err('<c:errDir val="y"/><c:errBarType val="plus"/><c:errValType val="fixedVal"/><c:noEndCap val="1"/><c:val val="1"/>') })) +
        '<c:valAx><c:axId val="111"/><c:scaling><c:max val="40"/><c:min val="0"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222"/></c:valAx>' +
        '<c:valAx><c:axId val="222"/><c:scaling><c:max val="4"/><c:min val="0"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111"/></c:valAx>',
    }),
  );
  const sets = items(scatter, 'path', 'error');
  assert.equal(sets.length, 2, 'one set of bars along each axis');
  const horizontal = sets.filter((item) => Math.abs(item.points[1] - item.points[3]) < 1e-9)[0];
  near(Math.abs(horizontal.points[2] - horizontal.points[0]), scatter.plot.width * 0.25, 'ten of forty along x');
  near(horizontal.points[1], valueY(scatter, 1), 'at the height of the point');
  assert.equal(uprights(scatter).length, 3, 'and three bars along y');
});

// ------------------------------------------------------- a pie of two plots

test('a pie with a second plot: the points of the second plot are one slice of the pie', () => {
  const of = (inner: string, values = [10, 20, 30, 40]) =>
    lay(chartXml({ plot: group('ofPieChart', '<c:ofPieType val="pie"/><c:varyColors val="1"/>', ser(0, { cat: strLit(['a', 'b', 'c', 'd']), val: numLit(values) }), inner, false) }));
  const shares = (layout: ChartLayout) => items(layout, 'slice').map((slice) => Math.round(((slice.end - slice.start) / (Math.PI * 2)) * 1000) / 1000);
  assert.deepEqual(shares(of('<c:splitType val="pos"/><c:splitPos val="2"/>')), [0.1, 0.2, 0.7], 'the last two together');
  assert.deepEqual(shares(of('<c:splitType val="val"/><c:splitPos val="25"/>')), [0.3, 0.4, 0.3], 'the values below 25 together, after the others');
  assert.deepEqual(shares(of('<c:splitType val="percent"/><c:splitPos val="15"/>')), [0.2, 0.3, 0.4, 0.1], 'the shares below 15 %');
  assert.deepEqual(shares(of('<c:splitType val="cust"/><c:custSplit><c:secondPiePt val="0"/><c:secondPiePt val="2"/></c:custSplit>')), [0.2, 0.4, 0.4], 'the points the file lists');
  // UNSURE: without a rule of its own the last third of the points, rounded up, is the second plot.
  assert.deepEqual(shares(of('')), [0.1, 0.2, 0.7]);
  assert.deepEqual(shares(of('<c:splitType val="pos"/><c:splitPos val="0"/>')), [0.1, 0.2, 0.3, 0.4], 'no point in the second plot: a plain pie');
  assert.deepEqual(shares(of('<c:splitType val="pos"/><c:splitPos val="99"/>')), [1], 'every point in it: one slice');
  const slices = items(of('<c:splitType val="pos"/><c:splitPos val="2"/>'), 'slice');
  assert.deepEqual(slices.map((slice) => slice.fill), ['#4472C4', '#ED7D31', '#5B9BD5'], 'UNSURE: the slice of the second plot has the colour after the last point');
});
