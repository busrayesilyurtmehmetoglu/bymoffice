// The scale of a value axis and the layout of a chart: where the bars, lines,
// slices, labels and the legend are put in a frame, measured with a stand-in
// for text measurement (every character half the font size wide).
//
// Excel is not at hand: what follows Excel's known behaviour without having
// been compared with it is marked UNSURE.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { LOCALE_TR } from '../src/format/locale';
import { MAX_AXIS_TICKS, MAX_DRAWN_POINTS, MAX_LEGEND_ENTRIES, MAX_PIE_SLICES, chartAutoColor, chartScale, chartValues, layoutChart, type ChartItem, type ChartLayout, type ChartLayoutOptions } from '../src/layout/chart';
import type { Chart } from '../src/model/sheet/chart';
import { catAx, chartXml, fakeMeasure, group, legend, numLit, numRef, readChart, ser, solid, spPr, strLit, title, valAx, xyAxes, xySer } from './chart-fixtures';

type Of<K extends ChartItem['kind']> = Extract<ChartItem, { kind: K }>;

function model(xml: string): Chart {
  const chart = readChart(xml);
  assert.ok(chart, 'the chart is read');
  return chart as Chart;
}

function lay(xml: string, width = 480, height = 320, options: ChartLayoutOptions = {}): ChartLayout {
  const chart = model(xml);
  return layoutChart(chart, chartValues(chart, { locale: options.locale }), width, height, fakeMeasure, options);
}

const items = <K extends ChartItem['kind']>(layout: ChartLayout, kind: K, role?: string): Of<K>[] => layout.items.filter((item) => item.kind === kind && (role === undefined || item.role === role)) as Of<K>[];
const texts = (layout: ChartLayout, role: string): string[] => items(layout, 'text', role).map((item) => item.text);
const near = (actual: number, expected: number, message: string, tolerance = 1e-6) => assert.ok(Math.abs(actual - expected) <= tolerance, `${message}: ${actual} is not ${expected}`);

const CATS = strLit(['Ocak', 'Şubat', 'Mart']);
const col = (series: string, head = '<c:grouping val="clustered"/>', tail = '', axes = catAx() + valAx(), rest: { head?: string; tail?: string; after?: string; before?: string } = {}) =>
  chartXml({ plot: group('barChart', `<c:barDir val="col"/>${head}`, series, tail) + axes, ...rest });
const one = (values: (number | null)[], props = '') => ser(0, { tx: '<c:v>Bir</c:v>', cat: CATS, val: numLit(values), props });
const two = (values: (number | null)[], props = '') => ser(1, { tx: '<c:v>İki</c:v>', cat: CATS, val: numLit(values), props });

// ----------------------------------------------------------------------- scale

test('scale: data of one sign starts at zero and leaves room at the open end', () => {
  // UNSURE: the rules are close to what Excel does; the exact ends were not compared with it.
  assert.deepEqual(chartScale(0, 30), { min: 0, max: 35, step: 5, ticks: [0, 5, 10, 15, 20, 25, 30, 35] });
  assert.deepEqual(chartScale(10, 30), { min: 0, max: 35, step: 5, ticks: [0, 5, 10, 15, 20, 25, 30, 35] });
  assert.deepEqual(chartScale(-30, 0), { min: -35, max: 0, step: 5, ticks: [-35, -30, -25, -20, -15, -10, -5, 0] });
  assert.deepEqual(chartScale(-5, 12), { min: -10, max: 15, step: 5, ticks: [-10, -5, 0, 5, 10, 15] });
  assert.deepEqual(chartScale(12, -5), chartScale(-5, 12), 'the ends in any order');
});

test('scale: all zero, a single value, nothing that is a number', () => {
  assert.deepEqual(chartScale(0, 0), { min: 0, max: 1, step: 0.2, ticks: [0, 0.2, 0.4, 0.6, 0.8, 1] });
  const single = chartScale(7, 7);
  assert.equal(single.min, 0);
  assert.ok(single.max > 7, 'a single value lies inside the axis');
  const negative = chartScale(-7, -7);
  assert.equal(negative.max, 0);
  assert.ok(negative.min < -7, 'a single negative value too');
  for (const [low, high] of [[NaN, NaN], [Infinity, -Infinity], [NaN, 5], [-Infinity, Infinity]] as [number, number][]) {
    const scale = chartScale(low, high);
    assert.ok(Number.isFinite(scale.min) && Number.isFinite(scale.max) && scale.max > scale.min && scale.ticks.length >= 2, `a scale for ${low}..${high}`);
  }
});

test('scale: very large and very small numbers get steps of 1, 2 or 5 times a power of ten', () => {
  for (const high of [1.23e12, 4.2e-7, 987654321, 0.55, 3, 1e-300, 1e300]) {
    const scale = chartScale(0, high);
    const mantissa = scale.step / Math.pow(10, Math.floor(Math.log10(scale.step) + 1e-9));
    assert.ok([1, 2, 5, 10].some((nice) => Math.abs(mantissa - nice) < 1e-6), `${high}: the step ${scale.step} is a round one`);
    assert.ok(scale.max >= high && scale.min === 0, `${high}: the data lies inside`);
    assert.ok(scale.ticks.length >= 2 && scale.ticks.length <= 12, `${high}: ${scale.ticks.length} ticks`);
  }
  assert.deepEqual(chartScale(0, 0.55).ticks, [0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6], 'no floating point noise in the ticks');
});

test('scale: data far from zero and close together does not start at zero', () => {
  // UNSURE: Excel's rule is "the difference is less than a sixth of the largest value".
  const scale = chartScale(98, 100);
  assert.ok(scale.min > 90 && scale.min <= 98, `starts near the data (${scale.min})`);
  assert.ok(scale.max >= 100, 'and holds the largest value');
  const wide = chartScale(50, 100);
  assert.equal(wide.min, 0, 'half the largest value apart: from zero');
});

test('scale: what the axis states is taken as it is', () => {
  const stated = chartScale(0, 30, { min: 10, max: 50 });
  assert.deepEqual([stated.min, stated.max], [10, 50]);
  assert.equal(stated.ticks[0], 10);
  assert.equal(stated.ticks[stated.ticks.length - 1], 50);
  assert.deepEqual(chartScale(0, 30, { step: 10 }).ticks, [0, 10, 20, 30, 40]);
  assert.deepEqual(chartScale(0, 30, { min: 3, step: 10 }).ticks, [3, 13, 23, 33], 'ticks count from a stated minimum');
  assert.equal(chartScale(0, 30, { max: 100 }).max, 100);
  // A minimum that is not below the maximum: neither is taken, the axis works its ends out by itself.
  assert.deepEqual(chartScale(0, 30, { min: 50, max: 10 }), chartScale(0, 30), 'ends the wrong way round are none');
  assert.deepEqual(chartScale(0, 30, { min: 7, max: 7 }), chartScale(0, 30), 'ends that are the same are none');
  // The data lies below a stated minimum.
  const above = chartScale(0, 3, { min: 10 });
  assert.ok(above.max > above.min, 'the axis has a length');
  // A step far too small for the range.
  const fine = chartScale(0, 1000, { step: 0.001 });
  assert.ok(fine.ticks.length <= MAX_AXIS_TICKS + 1, `${fine.ticks.length} ticks at most`);
  for (const bad of [0, -5, NaN, Infinity]) assert.deepEqual(chartScale(0, 30, { step: bad }), chartScale(0, 30), `a step of ${bad} is none`);
});

test('scale: fewer steps where there is less room', () => {
  assert.ok(chartScale(0, 30, { maxSteps: 2 }).ticks.length <= 3, 'room for two steps: three ticks at most');
  assert.ok(chartScale(0, 30, { maxSteps: 10 }).ticks.length >= 7, 'room for ten steps: more ticks');
  assert.ok(chartScale(0, 30, { maxSteps: 0 }).ticks.length >= 2, 'no room is still an axis');
});

// ------------------------------------------------------------------------ bars

test('columns: as high as their values, side by side in the order of the categories, standing on zero', () => {
  const layout = lay(col(one([10, 20, 30])));
  const bars = items(layout, 'rect', 'bar');
  assert.equal(bars.length, 3);
  const scale = layout.scale!;
  assert.deepEqual([scale.min, scale.max], [0, 35]);
  const unit = layout.plot.height / 35;
  bars.forEach((bar, i) => {
    near(bar.height, (i + 1) * 10 * unit, `column ${i} is as high as its value`);
    near(bar.y + bar.height, layout.plot.y + layout.plot.height, `column ${i} stands on the axis`);
    near(bar.width, bars[0].width, 'all columns are as wide');
    assert.equal(bar.fill, '#4472C4', 'the first accent colour');
    assert.deepEqual([bar.series, bar.point], [0, i]);
  });
  assert.ok(bars[0].x < bars[1].x && bars[1].x < bars[2].x, 'from left to right');
  // A gap of 150 %: a column is 1 / 2.5 of its slot, in the middle of it.
  const slot = layout.plot.width / 3;
  near(bars[0].width, slot / 2.5, 'the width of a column');
  near(bars[0].x + bars[0].width / 2, layout.plot.x + slot / 2, 'in the middle of its slot');
});

test('columns of two series: gap width and overlap decide their width and place', () => {
  const layout = lay(col(one([1, 2, 3]) + two([3, 2, 1]), '<c:grouping val="clustered"/>', '<c:gapWidth val="100"/><c:overlap val="0"/>'));
  const bars = items(layout, 'rect', 'bar');
  assert.equal(bars.length, 6);
  const slot = layout.plot.width / 3;
  const a = bars.filter((bar) => bar.series === 0);
  const b = bars.filter((bar) => bar.series === 1);
  near(a[0].width, slot / 3, 'two columns and a gap of one column');
  near(b[0].x, a[0].x + a[0].width, 'the second series stands right of the first');
  assert.equal(b[0].fill, '#ED7D31', 'the second accent colour');

  const over = items(lay(col(one([1, 2, 3]) + two([3, 2, 1]), '', '<c:gapWidth val="0"/><c:overlap val="50"/>')), 'rect', 'bar');
  near(over[3].x - over[0].x, over[0].width / 2, 'half an overlap: half a column apart');
  const apart = items(lay(col(one([1, 2, 3]) + two([3, 2, 1]), '', '<c:gapWidth val="0"/><c:overlap val="-100"/>')), 'rect', 'bar');
  near(apart[3].x - apart[0].x, apart[0].width * 2, 'a negative overlap: a column of room between them');
});

test('stacked columns stand on each other; negative ones hang below zero', () => {
  const layout = lay(col(one([10, 20, -5]) + two([5, 10, -10]), '<c:grouping val="stacked"/>', '<c:overlap val="100"/>'));
  const bars = items(layout, 'rect', 'bar');
  const a = bars.filter((bar) => bar.series === 0);
  const b = bars.filter((bar) => bar.series === 1);
  const zero = layout.plot.y + layout.plot.height - ((0 - layout.scale!.min) / (layout.scale!.max - layout.scale!.min)) * layout.plot.height;
  near(a[0].y + a[0].height, zero, 'the first series stands on zero');
  near(b[0].y + b[0].height, a[0].y, 'the second on the first');
  near(b[0].x, a[0].x, 'in the same place');
  near(a[2].y, zero, 'a negative column hangs from zero');
  near(b[2].y, a[2].y + a[2].height, 'the next negative one below it');
  assert.ok(layout.scale!.max >= 30 && layout.scale!.min <= -15, 'the axis holds the sums');
});

test('100 % stacked columns fill the height; the axis counts in per cent', () => {
  const layout = lay(col(one([1, 2, 30]) + two([3, 2, 10]), '<c:grouping val="percentStacked"/>'));
  assert.deepEqual([layout.scale!.min, layout.scale!.max], [0, 1]);
  const bars = items(layout, 'rect', 'bar');
  for (let i = 0; i < 3; i++) near(bars[i].height + bars[i + 3].height, layout.plot.height, `category ${i} is filled`);
  near(bars[0].height, layout.plot.height * 0.25, '1 of 4 is a quarter');
  const ticks = texts(layout, 'tick');
  assert.equal(ticks[0], '0%');
  assert.equal(ticks[ticks.length - 1], '100%');
  // With negative values the axis runs from -100 % to 100 %.
  const mixed = lay(col(one([1, -1, 3]) + two([3, 1, -1]), '<c:grouping val="percentStacked"/>'));
  assert.deepEqual([mixed.scale!.min, mixed.scale!.max], [-1, 1]);
});

test('bars that lie: as long as their values, categories from the bottom up, values along the bottom', () => {
  const layout = lay(chartXml({ plot: group('barChart', '<c:barDir val="bar"/><c:grouping val="clustered"/>', one([10, 20, 30])) + catAx() + valAx() }));
  const bars = items(layout, 'rect', 'bar');
  const unit = layout.plot.width / layout.scale!.max;
  bars.forEach((bar, i) => {
    near(bar.width, (i + 1) * 10 * unit, `bar ${i} is as long as its value`);
    near(bar.x, layout.plot.x, `bar ${i} starts at the axis`);
  });
  assert.ok(bars[0].y > bars[1].y && bars[1].y > bars[2].y, 'the first category is the lowest');
  const categories = items(layout, 'text', 'category');
  assert.deepEqual(categories.map((item) => item.text), ['Ocak', 'Şubat', 'Mart']);
  assert.ok(categories.every((item) => item.align === 'right' && item.x < layout.plot.x), 'category labels at the left');
  assert.ok(items(layout, 'text', 'tick').every((item) => item.y > layout.plot.y + layout.plot.height), 'value labels below');
});

test('an axis that runs the other way turns the columns and the order of the categories', () => {
  const layout = lay(col(one([10, 20, 30]), '', '', catAx() + valAx('', '<c:orientation val="maxMin"/>')));
  const bars = items(layout, 'rect', 'bar');
  bars.forEach((bar) => near(bar.y, layout.plot.y, 'columns hang from the top'));
  const turned = lay(col(one([10, 20, 30]), '', '', '<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="maxMin"/></c:scaling><c:axPos val="b"/><c:crossAx val="222"/></c:catAx>' + valAx()));
  const reversed = items(turned, 'rect', 'bar');
  assert.ok(reversed[0].x > reversed[1].x && reversed[1].x > reversed[2].x, 'the first category at the right');
});

test('a stated minimum and maximum cut the columns to the axis', () => {
  const layout = lay(col(one([10, 20, 300]), '', '', catAx() + valAx('', '<c:orientation val="minMax"/><c:max val="100"/><c:min val="0"/>')));
  assert.deepEqual([layout.scale!.min, layout.scale!.max], [0, 100]);
  const bars = items(layout, 'rect', 'bar');
  near(bars[2].height, layout.plot.height, 'a value beyond the axis ends at its end');
  near(bars[2].y, layout.plot.y, 'and does not leave the plot area');
});

test('a point without a number: no column; "invert if negative" shows a negative column white', () => {
  const layout = lay(col(one([10, null, -30], '<c:invertIfNegative val="1"/>')));
  const bars = items(layout, 'rect', 'bar');
  assert.deepEqual(bars.map((bar) => bar.point), [0, 2]);
  assert.equal(bars[0].fill, '#4472C4');
  // UNSURE: without a second colour Excel shows such a column white.
  assert.deepEqual([bars[1].fill, bars[1].stroke], ['#FFFFFF', '#4472C4']);
});

test('colours: what the file states, a colour per point, and the accent colours in turn with their lighter and darker forms', () => {
  const chart = model(col(one([1, 2, 3])));
  assert.deepEqual([0, 1, 5].map((i) => chartAutoColor(chart, i)), ['#4472C4', '#ED7D31', '#70AD47']);
  // UNSURE: the order of the forms.
  assert.equal(chartAutoColor(chart, 6), '#294476', 'the seventh series: the first colour, darker');
  assert.equal(chartAutoColor(chart, 12), '#698ED0', 'the thirteenth: lighter');
  assert.equal(chartAutoColor(chart, 6 * 9), '#4472C4', 'and round again');
  for (const index of [-1, NaN, Infinity, 1e12]) assert.match(chartAutoColor(chart, index), /^#[0-9A-F]{6}$/, `a colour for ${index}`);
  assert.match(chartAutoColor({ palette: [] }, 0), /^#[0-9A-F]{6}$/, 'no palette');

  const stated = items(lay(col(one([1, 2, 3], spPr(solid('00B050') + `<a:ln w="19050">${solid('000000')}</a:ln>`) + `<c:dPt><c:idx val="1"/>${spPr(solid('FF0000'))}</c:dPt>`))), 'rect', 'bar');
  assert.deepEqual(stated.map((bar) => bar.fill), ['#00B050', '#FF0000', '#00B050']);
  assert.deepEqual([stated[0].stroke, stated[0].lineWidth], ['#000000', 2]);
  const varied = items(lay(col(one([1, 2, 3]), '<c:varyColors val="1"/>')), 'rect', 'bar');
  assert.deepEqual(varied.map((bar) => bar.fill), ['#4472C4', '#ED7D31', '#A5A5A5'], 'a single series whose points vary');
  const none = items(lay(col(one([1, 2, 3], spPr('<a:noFill/>')))), 'rect', 'bar');
  assert.deepEqual(none.map((bar) => bar.fill), [undefined, undefined, undefined], 'no fill is no fill');
});

// ------------------------------------------------------------- lines and areas

const line = (series: string, head = '<c:grouping val="standard"/>', tail = '<c:marker val="1"/>', blanks = 'gap', axes = catAx() + valAx()) =>
  chartXml({ plot: group('lineChart', head, series, tail) + axes, tail: `<c:plotVisOnly val="1"/><c:dispBlanksAs val="${blanks}"/>` });

test('a line goes through its points, in the middle of their categories; markers where the chart asks for them', () => {
  const layout = lay(line(one([10, 20, 30])));
  const lines = items(layout, 'path', 'line');
  assert.equal(lines.length, 1);
  const points = lines[0].points;
  assert.equal(points.length, 6);
  const slot = layout.plot.width / 3;
  for (let i = 0; i < 3; i++) {
    near(points[i * 2], layout.plot.x + (i + 0.5) * slot, `point ${i} across`);
    near(points[i * 2 + 1], layout.plot.y + layout.plot.height * (1 - ((i + 1) * 10) / layout.scale!.max), `point ${i} up`);
  }
  assert.deepEqual([lines[0].stroke, lines[0].lineWidth], ['#4472C4', 3]);
  const markers = items(layout, 'markers', 'marker');
  assert.equal(markers.length, 1);
  assert.deepEqual(markers[0].points, points);
  // UNSURE: the marker a chart gives by itself is a circle.
  assert.equal(markers[0].symbol, 'circle');
  assert.equal(items(lay(line(one([10, 20, 30]), undefined, '')), 'markers', 'marker').length, 0, 'no markers unless asked for');
  assert.equal(items(lay(line(one([10, 20, 30], '<c:marker><c:symbol val="none"/></c:marker>'))), 'markers', 'marker').length, 0, 'a series that states none has none');
  const own = items(lay(line(one([10, 20, 30], '<c:marker><c:symbol val="square"/><c:size val="9"/></c:marker>'), undefined, '')), 'markers', 'marker');
  assert.deepEqual([own[0].symbol, own[0].size], ['square', 12], 'a stated marker: 9 points are 12 pixels');
});

test('points without a number: a gap breaks the line, "span" joins across it, "zero" goes down to zero', () => {
  const gap = lay(line(one([10, null, 30, 40, null, 60])));
  const runs = items(gap, 'path', 'line');
  assert.deepEqual(runs.map((run) => run.points.length / 2), [2], 'only the two points next to each other make a line');
  const dots = items(lay(line(one([10, null, 30, 40, null, 60]), undefined, '')), 'markers', 'marker');
  assert.equal(dots.length, 1, 'points with a gap on both sides are shown as dots');
  assert.equal(dots[0].points.length / 2, 2);

  const span = items(lay(line(one([10, null, 30, 40, null, 60]), undefined, undefined, 'span')), 'path', 'line');
  assert.deepEqual(span.map((run) => run.points.length / 2), [4], 'one line through the four points');
  const zero = lay(line(one([10, null, 30]), undefined, undefined, 'zero'));
  const through = items(zero, 'path', 'line')[0].points;
  assert.equal(through.length, 6);
  near(through[3], zero.plot.y + zero.plot.height, 'the empty point is at zero');
});

test('stacked lines add up; an area is closed along what it stands on', () => {
  const stacked = lay(line(one([10, 20, 30]) + two([5, 5, 5]), '<c:grouping val="stacked"/>'));
  const [a, b] = items(stacked, 'path', 'line');
  const unit = stacked.plot.height / (stacked.scale!.max - stacked.scale!.min);
  near(a.points[1] - b.points[1], 5 * unit, 'the second line lies five above the first');
  assert.ok(stacked.scale!.max >= 35, 'the axis holds the sums');

  const area = lay(chartXml({ plot: group('areaChart', '<c:grouping val="stacked"/>', one([10, 20, 30]) + two([5, 5, 5])) + catAx() + valAx('<c:crossBetween val="midCat"/>') }));
  const shapes = items(area, 'path', 'area');
  assert.equal(shapes.length, 2);
  assert.ok(shapes.every((shape) => shape.closed === true && shape.points.length === 12), 'three points along the top and three back');
  assert.deepEqual(shapes.map((shape) => shape.fill), ['#4472C4', '#ED7D31']);
  near(shapes[0].points[0], area.plot.x, 'points on the tick marks: the first at the left edge');
  near(shapes[0].points[4], area.plot.x + area.plot.width, 'the last at the right edge');
  near(shapes[0].points[11], area.plot.y + area.plot.height, 'the first area stands on the axis');
  near(shapes[1].points[11], shapes[0].points[1], 'the second on the first');
});

// ---------------------------------------------------------------- pie, scatter

const pie = (values: (number | null)[], name = 'pieChart', head = '', tail = '', props = '') =>
  chartXml({ plot: group(name, head, ser(0, { cat: strLit(['a', 'b', 'c', 'd']), val: numLit(values), props }), tail, false), tail: legend('r') });

test('a pie: slices as large as their shares, clockwise from the top, together a full turn', () => {
  const layout = lay(pie([10, 20, 30, 40]));
  const slices = items(layout, 'slice');
  assert.equal(slices.length, 4);
  let total = 0;
  slices.forEach((slice, i) => {
    near(slice.end - slice.start, ((i + 1) / 10) * Math.PI * 2, `slice ${i}`);
    if (i > 0) near(slice.start, slices[i - 1].end, `slice ${i} starts where the one before ends`);
    total += slice.end - slice.start;
    assert.equal(slice.inner, 0);
  });
  near(total, Math.PI * 2, 'a full turn');
  near(slices[0].start, -Math.PI / 2, 'from the top');
  assert.deepEqual(slices.map((slice) => slice.fill), ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000'], 'a colour per slice');
  near(slices[0].cx, layout.plot.x + layout.plot.width / 2, 'in the middle of its area');
  assert.ok(slices[0].radius * 2 <= Math.min(layout.width, layout.height), 'inside the frame');
  assert.equal(layout.scale, undefined, 'a pie has no axis');
  assert.equal(items(layout, 'text', 'tick').length, 0);
  assert.deepEqual(texts(layout, 'legend'), ['a', 'b', 'c', 'd'], 'the legend lists the slices');
});

test('a pie: the first slice angle, slices with a colour of their own, values that are nothing or negative', () => {
  const turned = items(lay(pie([1, 1, 1, 1], 'pieChart', '', '<c:firstSliceAng val="90"/>')), 'slice');
  near(turned[0].start, 0, 'a quarter turn: from the right');
  const own = items(lay(pie([1, 1, 1, 1], 'pieChart', '', '', `<c:dPt><c:idx val="2"/>${spPr(solid('112233') + `<a:ln w="19050">${solid('FFFFFF')}</a:ln>`)}</c:dPt>`)), 'slice');
  assert.deepEqual([own[2].fill, own[2].stroke, own[2].lineWidth], ['#112233', '#FFFFFF', 2]);
  assert.equal(own[1].stroke, undefined);
  const same = items(lay(pie([1, 1, 1, 1], 'pieChart', '<c:varyColors val="0"/>')), 'slice');
  assert.deepEqual(same.map((slice) => slice.fill), ['#4472C4', '#4472C4', '#4472C4', '#4472C4'], 'colours that do not vary');
  // UNSURE: a negative value takes the share of its size.
  const mixed = items(lay(pie([10, null, -10, 0])), 'slice');
  assert.deepEqual(mixed.map((slice) => slice.point), [0, 2]);
  near(mixed[1].end - mixed[1].start, Math.PI, 'half the pie');
  assert.equal(items(lay(pie([0, 0, 0, 0])), 'slice').length, 0, 'nothing to share out: no slices');
  assert.equal(items(lay(pie([null, null, null, null])), 'slice').length, 0);
});

test('a doughnut: a hole, and a ring per series', () => {
  const xml = chartXml({ plot: group('doughnutChart', '', ser(0, { val: numLit([1, 1]) }) + ser(1, { val: numLit([1, 3]) }), '<c:holeSize val="50"/>', false) });
  const slices = items(lay(xml), 'slice');
  assert.equal(slices.length, 4);
  const inner = slices.filter((slice) => slice.series === 0);
  const outer = slices.filter((slice) => slice.series === 1);
  // UNSURE: the first series is the innermost ring.
  near(inner[0].inner, outer[0].radius / 2, 'the hole is half the diameter');
  near(inner[0].radius, outer[0].inner, 'the rings touch');
  near(outer[1].end - outer[1].start, Math.PI * 1.5, 'three quarters of the outer ring');
});

test('a pie of very many slices is drawn up to a limit; what is left is one slice', () => {
  const values: number[] = [];
  for (let i = 0; i < MAX_PIE_SLICES + 500; i++) values.push(1);
  const slices = items(lay(chartXml({ plot: group('pieChart', '', ser(0, { val: numLit(values) }), '', false) })), 'slice');
  assert.equal(slices.length, MAX_PIE_SLICES + 1);
  near(slices[slices.length - 1].end, -Math.PI / 2 + Math.PI * 2, 'the last slice closes the pie');
});

test('a pie beside another chart type: the pie alone is drawn', () => {
  const xml = chartXml({ plot: group('pieChart', '', ser(0, { val: numLit([1, 2]) }), '', false) + group('barChart', '<c:barDir val="col"/>', ser(1, { val: numLit([5, 6]) })) + catAx() + valAx() });
  const layout = lay(xml);
  assert.equal(items(layout, 'slice').length, 2);
  assert.equal(items(layout, 'rect', 'bar').length, 0);
});

const scatter = (series: string, style = 'lineMarker') => chartXml({ plot: group('scatterChart', `<c:scatterStyle val="${style}"/>`, series) + xyAxes() });

test('a scatter chart: points placed by both their values, each axis with a scale of its own', () => {
  const layout = lay(scatter(xySer(0, { cat: numLit([0, 50, 100]), val: numLit([10, 30, 20]) })));
  assert.deepEqual([layout.xScale!.min, layout.xScale!.max], [0, 120]);
  const markers = items(layout, 'markers', 'marker');
  assert.equal(markers.length, 1);
  const points = markers[0].points;
  near(points[0], layout.plot.x, 'x = 0 at the left edge');
  near(points[2] - points[0], (50 / 120) * layout.plot.width, 'x = 50');
  near(points[1], layout.plot.y + layout.plot.height * (1 - 10 / layout.scale!.max), 'y = 10');
  assert.equal(items(layout, 'path', 'line').length, 1, 'line and markers');
  assert.equal(items(lay(scatter(xySer(0, { cat: numLit([0, 50, 100]), val: numLit([10, 30, 20]) }), 'marker')), 'path', 'line').length, 0, 'markers alone');
  assert.equal(items(lay(scatter(xySer(0, { cat: numLit([0, 50, 100]), val: numLit([10, 30, 20]), props: spPr('<a:ln w="19050"><a:noFill/></a:ln>') }))), 'path', 'line').length, 0, 'a line that is stated as none');
  assert.equal(items(lay(scatter(xySer(0, { cat: numLit([0, 50, 100]), val: numLit([10, 30, 20]) }), 'line')), 'markers', 'marker').length, 0, 'a line alone');
  const labels = texts(layout, 'category');
  assert.ok(labels.indexOf('0') >= 0 && labels.indexOf('120') >= 0, 'numbers along the bottom');
});

test('several chart types on the same axes: areas under bars under lines', () => {
  const xml = chartXml({
    plot:
      group('lineChart', '<c:grouping val="standard"/>', two([30, 10, 20])) +
      group('barChart', '<c:barDir val="col"/>', one([10, 20, 30])) +
      group('areaChart', '', ser(2, { val: numLit([5, 5, 5]) })) +
      catAx() +
      valAx(),
  });
  const layout = lay(xml);
  const order = layout.items.filter((item) => item.role === 'bar' || item.role === 'line' || item.role === 'area').map((item) => item.role);
  assert.deepEqual(order, ['area', 'bar', 'bar', 'bar', 'line']);
  const bars = items(layout, 'rect', 'bar');
  const points = items(layout, 'path', 'line')[0].points;
  near(points[0], bars[0].x + bars[0].width / 2, 'the line goes over the middle of the columns');
  near(points[1], bars[2].y, 'on the same axis: 30 is 30');
});

// -------------------------------------------------------------- title, legend

test('the title: above everything, centred, cut to the frame; no title element, no title', () => {
  const layout = lay(col(one([1, 2, 3]), '', '', undefined, { head: title('Yıllık satışlar', 'sz="2000" b="1"') }));
  const [heading] = items(layout, 'text', 'title');
  assert.equal(heading.text, 'Yıllık satışlar');
  assert.deepEqual([heading.align, heading.x], ['center', 240]);
  assert.ok(heading.y < layout.plot.y, 'above the plot area');
  assert.deepEqual([heading.font.bold, Math.round(heading.font.size * 100) / 100], [true, 26.67], '20 points');
  const long = items(lay(col(one([1, 2, 3]), '', '', undefined, { head: title('çok uzun bir başlık '.repeat(20)) })), 'text', 'title')[0];
  assert.ok(long.text.endsWith('…') && fakeMeasure(long.text, long.font) <= 480, 'a long title is cut with an ellipsis');
  assert.equal(items(lay(col(one([1, 2, 3]))), 'text', 'title').length, 0, 'no title element: none');
  // UNSURE: a title element without a text is the name of the series of a chart with one series, and nothing for several.
  assert.deepEqual(texts(lay(col(one([1, 2, 3]), '', '', undefined, { head: '<c:title><c:overlay val="0"/></c:title>' })), 'title'), ['Bir']);
  assert.deepEqual(texts(lay(col(one([1, 2, 3]) + two([1, 2, 3]), '', '', undefined, { head: '<c:title><c:overlay val="0"/></c:title>' })), 'title'), []);
});

test('the legend stands where the file puts it, and the plot area makes room for it', () => {
  const series = one([1, 2, 3]) + two([3, 2, 1]);
  const without = lay(col(series));
  assert.equal(items(without, 'text', 'legend').length, 0);
  for (const position of ['r', 'l', 't', 'b', 'tr'] as const) {
    const layout = lay(col(series, '', '', undefined, { tail: legend(position) }));
    const entries = items(layout, 'text', 'legend');
    assert.deepEqual(entries.map((entry) => entry.text), ['Bir', 'İki'], position);
    assert.equal(items(layout, 'rect', 'swatch').length, 2, `${position}: a key per entry`);
    const { plot } = layout;
    for (const entry of entries) {
      if (position === 'r' || position === 'tr') assert.ok(entry.x > plot.x + plot.width, `${position}: right of the plot area`);
      if (position === 'l') assert.ok(entry.x < plot.x, 'l: left of the plot area');
      if (position === 't') assert.ok(entry.y < plot.y, 't: above the plot area');
      if (position === 'b') assert.ok(entry.y > plot.y + plot.height, 'b: below the plot area');
      assert.ok(entry.x >= 0 && entry.x + fakeMeasure(entry.text, entry.font) <= layout.width && entry.y > 0 && entry.y <= layout.height, `${position}: inside the frame`);
    }
    if (position === 'r' || position === 'l' || position === 'tr') assert.ok(plot.width < without.plot.width, `${position}: the plot area is narrower`);
    else assert.ok(plot.height < without.plot.height, `${position}: the plot area is lower`);
    if (position === 'tr') assert.ok(entries[0].y < layout.height / 3, 'tr: at the top');
    if (position === 'r') assert.ok(Math.abs((entries[0].y + entries[1].y) / 2 - layout.height / 2) < 20, 'r: in the middle of the height');
  }
});

test('legend: entries the file leaves out, keys of lines, long names, more entries than there is room for', () => {
  const left = lay(col(one([1, 2, 3]) + two([3, 2, 1]), '', '', undefined, { tail: legend('r', '<c:legendEntry><c:idx val="0"/><c:delete val="1"/></c:legendEntry>') }));
  assert.deepEqual(texts(left, 'legend'), ['İki']);
  const lined = lay(chartXml({ plot: group('lineChart', '', one([1, 2, 3]), '<c:marker val="1"/>') + catAx() + valAx(), tail: legend('b') }));
  assert.equal(items(lined, 'path', 'swatch').length, 1, 'a line is the key of a line');
  assert.equal(items(lined, 'markers', 'swatch').length, 1, 'with its marker');
  const long = lay(col(ser(0, { tx: `<c:v>${'uzun ad '.repeat(40)}</c:v>`, val: numLit([1, 2, 3]) }), '', '', undefined, { tail: legend('r') }));
  const entry = items(long, 'text', 'legend')[0];
  assert.ok(entry.text.endsWith('…'), 'a long name is cut');
  assert.ok(entry.x + fakeMeasure(entry.text, entry.font) <= long.width, 'and stays in the frame');
  assert.ok(long.plot.width > long.width * 0.4, 'the legend takes no more than its share');
  let many = '';
  for (let i = 0; i < 150; i++) many += ser(i, { tx: `<c:v>S${i}</c:v>`, val: numLit([i]) });
  const crowded = lay(col(many, '', '', undefined, { tail: legend('r') }));
  const shown = items(crowded, 'text', 'legend');
  assert.ok(shown.length > 0 && shown.length < MAX_LEGEND_ENTRIES, `${shown.length} entries fit`);
  assert.ok(shown.every((item) => item.y <= crowded.height), 'none below the frame');
});

// ------------------------------------------------------------- axes and labels

test('the value axis: a label per tick in the format of the axis, grid lines where the file has them', () => {
  const layout = lay(col(one([1000, 2000, 3000]), '', '', catAx() + valAx('<c:majorGridlines/><c:numFmt formatCode="#,##0.0" sourceLinked="0"/>')));
  const ticks = items(layout, 'text', 'tick');
  assert.deepEqual(ticks.map((tick) => tick.text), layout.scale!.ticks.map((value) => value.toLocaleString('en-US', { minimumFractionDigits: 1 })));
  assert.ok(ticks.every((tick) => tick.align === 'right' && tick.x < layout.plot.x), 'left of the plot area');
  assert.equal(items(layout, 'path', 'grid').length, layout.scale!.ticks.length, 'a grid line per tick');
  assert.equal(items(lay(col(one([1, 2, 3]), '', '', catAx() + valAx(''))), 'path', 'grid').length, 0, 'no grid lines unless the file has them');
  const turkish = lay(col(one([1000, 2000, 3000]), '', '', catAx() + valAx('<c:numFmt formatCode="#,##0.0" sourceLinked="0"/>')), 480, 320, { locale: LOCALE_TR });
  assert.equal(texts(turkish, 'tick')[1], '500,0', 'in the language of the view');
  // Linked to the source: the format of the values of the first series.
  const linked = lay(col(ser(0, { val: numRef('X!$A$1:$A$3', [0.1, 0.2, 0.3], '0%') }), '', '', catAx() + valAx('<c:numFmt formatCode="General" sourceLinked="1"/>')));
  assert.ok(texts(linked, 'tick').every((text) => /%$/.test(text)), 'per cent, as the cells are');
  const minor = lay(col(one([1, 2, 3]), '', '', catAx() + valAx('<c:majorGridlines/><c:minorGridlines/>')));
  assert.equal(items(minor, 'path', 'grid').length, minor.scale!.ticks.length + (minor.scale!.ticks.length - 1) * 4, 'four minor lines between two major ones');
});

test('axes that are deleted or show no labels take no room; titles of the axes are placed beside them', () => {
  const plain = lay(col(one([1, 2, 3])));
  const bare = lay(col(one([1, 2, 3]), '', '', '<c:catAx><c:axId val="111"/><c:delete val="1"/><c:axPos val="b"/></c:catAx><c:valAx><c:axId val="222"/><c:delete val="0"/><c:axPos val="l"/><c:tickLblPos val="none"/></c:valAx>'));
  assert.equal(items(bare, 'text', 'category').length, 0);
  assert.equal(items(bare, 'text', 'tick').length, 0);
  assert.ok(bare.plot.width > plain.plot.width && bare.plot.height > plain.plot.height, 'the plot area takes the room');
  assert.equal(items(bare, 'path', 'axis').length, 1, 'a deleted axis has no line');

  const titled = lay(col(one([1, 2, 3]), '', '', catAx(title('Aylar')) + valAx(title('Tutar'))));
  const axisTitles = items(titled, 'text', 'axisTitle');
  assert.deepEqual(axisTitles.map((item) => item.text).sort(), ['Aylar', 'Tutar']);
  const below = axisTitles.filter((item) => item.text === 'Aylar')[0];
  const beside = axisTitles.filter((item) => item.text === 'Tutar')[0];
  assert.ok(below.y > titled.plot.y + titled.plot.height && !below.rotated, 'the title of the categories below them');
  assert.ok(beside.x < titled.plot.x && beside.rotated === true, 'the title of the values at the left, turned');
});

test('long category labels are cut and many are thinned out, so that none runs into the next', () => {
  const names: string[] = [];
  for (let i = 0; i < 40; i++) names.push(`Kategori numara ${i} uzun`);
  const values = names.map((_name, i) => i);
  const layout = lay(col(ser(0, { cat: strLit(names), val: numLit(values) })));
  const labels = items(layout, 'text', 'category');
  assert.ok(labels.length > 1 && labels.length < 40, `${labels.length} of 40 labels are shown`);
  for (let i = 1; i < labels.length; i++) {
    const room = labels[i].x - labels[i - 1].x;
    assert.ok(fakeMeasure(labels[i - 1].text, labels[i - 1].font) <= room + 1e-6, `label ${i - 1} ends before label ${i} starts`);
  }
  assert.equal(labels[0].text.indexOf('Kategori'), 0, 'the first one is the first category');
  const few = items(lay(col(ser(0, { cat: strLit(['a', 'x'.repeat(500), 'c']), val: numLit([1, 2, 3]) }))), 'text', 'category');
  assert.ok(few.some((label) => label.text.endsWith('…') && label.text.length < 100), 'a label wider than its slot is cut');
});

// ----------------------------------------------------------- small and strange

test('a small frame shows less and less: no words, then the series alone; nothing leaves the frame', () => {
  const xml = col(one([10, 20, 30]) + two([30, 20, 10]), '', '', catAx(title('Aylar')) + valAx(), { head: title('Başlık'), tail: legend('r') });
  const full = lay(xml, 480, 320);
  assert.equal(full.detail, 'full');
  assert.ok(items(full, 'text').length > 5);
  const compact = lay(xml, 150, 100);
  assert.equal(compact.detail, 'compact');
  assert.equal(items(compact, 'text').length, 0, 'no words');
  assert.equal(items(compact, 'rect', 'bar').length, 6);
  assert.ok(items(compact, 'path', 'axis').length > 0, 'axes still');
  for (const [width, height] of [[40, 30], [12, 200], [200, 12], [3, 3], [1, 1], [0.4, 0.4]]) {
    const bare = lay(xml, width, height);
    assert.equal(bare.detail, 'bare', `${width} x ${height}`);
    assert.equal(items(bare, 'text').length, 0);
    assert.equal(items(bare, 'path').length, 0, 'the series alone');
    const bars = items(bare, 'rect', 'bar');
    assert.equal(bars.length, 6);
    for (const bar of bars) {
      assert.ok(bar.x >= -1e-6 && bar.y >= -1e-6 && bar.x + bar.width <= width + 1e-6 && bar.y + bar.height <= height + 1e-6, `${width} x ${height}: a bar inside the frame`);
      assert.ok(Number.isFinite(bar.x) && Number.isFinite(bar.y) && Number.isFinite(bar.width) && Number.isFinite(bar.height), 'with numbers that are numbers');
    }
  }
});

test('sizes that are no sizes give an empty layout, and nothing is thrown', () => {
  for (const [width, height] of [[0, 100], [100, 0], [-5, 5], [NaN, 100], [100, Infinity], [Infinity, Infinity]] as [number, number][]) {
    const layout = lay(col(one([1, 2, 3])), width, height);
    assert.ok(layout.items.every((item) => item.kind !== 'text'), `${width} x ${height}`);
    for (const item of layout.items) if (item.kind === 'rect') assert.ok(Number.isFinite(item.width) && Number.isFinite(item.height), 'finite');
  }
  assert.deepEqual(lay(col(one([1, 2, 3])), 0, 0).items, []);
});

test('series without values, without points and with numbers that are none lay out to something that can be drawn', () => {
  const cases = [
    col(ser(0, {})),
    col(ser(0, { val: numLit([]) })),
    col(ser(0, { val: numLit([null, null]) })),
    col(ser(0, { val: numRef('X!$A$1', ['1e999', 'NaN', '-1e999']) })),
    line(ser(0, { val: numLit([5]) })),
    chartXml({ plot: group('areaChart', '', ser(0, { val: numLit([5]) })) + catAx() + valAx('<c:crossBetween val="midCat"/>') }),
    scatter(xySer(0, {})),
    scatter(xySer(0, { cat: numLit([1, 1, 1]), val: numLit([2, 2, 2]) })),
    chartXml({ plot: group('doughnutChart', '', ser(0, {}), '', false) }),
    col(one([1e308, -1e308, 1e308])),
  ];
  for (const xml of cases) {
    for (const [width, height] of [[480, 320], [150, 100], [30, 20]]) {
      const layout = lay(xml, width, height);
      for (const item of layout.items) {
        const numbers: number[] = item.kind === 'rect' ? [item.x, item.y, item.width, item.height] : item.kind === 'path' || item.kind === 'markers' ? item.points : item.kind === 'slice' ? [item.cx, item.cy, item.radius, item.start, item.end] : [item.x, item.y];
        assert.ok(numbers.every((value) => Number.isFinite(value)), `every coordinate is a number (${item.kind} ${item.role}, ${width} x ${height})`);
      }
    }
  }
});

test('very many points are sampled down to what the frame can show, quickly', () => {
  const count = 100000;
  const values: number[] = [];
  for (let i = 0; i < count; i++) values.push(Math.sin(i / 500) * 100 + (i % 1000 === 0 ? 500 : 0));
  const chartLine = model(line(ser(0, { val: numLit(values) }), undefined, ''));
  const started = Date.now();
  const data = chartValues(chartLine);
  const layout = layoutChart(chartLine, data, 480, 320, fakeMeasure);
  const drawn = items(layout, 'path', 'line').reduce((sum, path) => sum + path.points.length / 2, 0);
  assert.ok(drawn > 100 && drawn <= MAX_DRAWN_POINTS * 2, `${drawn} points are drawn for ${count}`);
  let peak = -Infinity;
  for (const value of values) if (value > peak) peak = value;
  assert.ok(peak > 500 && layout.scale!.max >= peak, 'the axis holds the peaks');
  let highest = Infinity;
  for (const path of items(layout, 'path', 'line')) for (let i = 1; i < path.points.length; i += 2) highest = Math.min(highest, path.points[i]);
  near(highest, layout.plot.y + layout.plot.height * (1 - (peak - layout.scale!.min) / (layout.scale!.max - layout.scale!.min)), 'a peak of one point in a thousand is on the line', 1e-6);

  const bars = items(layoutChart(model(col(ser(0, { val: numLit(values) }))), data, 480, 320, fakeMeasure), 'rect', 'bar');
  assert.ok(bars.length > 100 && bars.length <= 480, `${bars.length} columns are drawn for ${count}`);
  const wide = layoutChart(chartLine, data, 50000, 320, fakeMeasure);
  const most = items(wide, 'path', 'line').reduce((sum, path) => sum + path.points.length / 2, 0);
  assert.ok(most <= MAX_DRAWN_POINTS * 2, `never more than ${MAX_DRAWN_POINTS * 2} points, however wide the frame (${most})`);

  const cloud = model(scatter(xySer(0, { cat: numLit(values), val: numLit(values) }), 'marker'));
  const dots = items(layoutChart(cloud, chartValues(cloud), 480, 320, fakeMeasure), 'markers', 'marker');
  assert.ok(dots[0].points.length / 2 <= MAX_DRAWN_POINTS * 2, 'a scatter chart of very many points shows every n-th');
  // A guard against hanging, not a measurement.
  assert.ok(Date.now() - started < 10000, 'laid out in well under ten seconds');
});

test('the frame: white with a grey outline, or what the file states; rounded corners; the fill of the plot area', () => {
  const plain = lay(col(one([1, 2, 3])));
  const frame = items(plain, 'rect', 'frame')[0];
  assert.deepEqual([frame.x, frame.y, frame.width, frame.height, frame.fill, frame.stroke], [0, 0, 480, 320, '#FFFFFF', '#D9D9D9']);
  assert.equal(plain.radius, 0);
  assert.equal(plain.items[0], frame, 'the frame is drawn first');
  const styled = lay(col(one([1, 2, 3]), '', '', catAx() + valAx() + spPr(solid('EEEEEE')), { before: '<c:roundedCorners val="1"/>', after: spPr('<a:noFill/><a:ln><a:noFill/></a:ln>') }));
  const none = items(styled, 'rect', 'frame')[0];
  assert.deepEqual([none.fill, none.stroke], [undefined, undefined], 'no fill and no outline');
  assert.ok(styled.radius > 0, 'rounded corners');
  const plot = items(styled, 'rect', 'plot')[0];
  assert.deepEqual([plot.fill, plot.x, plot.y, plot.width, plot.height], ['#EEEEEE', styled.plot.x, styled.plot.y, styled.plot.width, styled.plot.height]);
});

test('texts from the file reach the layout as plain lines: no control characters, a bounded length', () => {
  const nasty = 'a\u0000b\u0007c\td\u001be';
  const xml = col(ser(0, { tx: `<c:v>${'x'.repeat(3000)}</c:v>`, cat: strLit(['<img src=x onerror=alert(1)>', 'iki\nsatır', 'x'.repeat(3000)]), val: numLit([1, 2, 3]) }), '', '', undefined, { tail: legend('b') });
  const chart = model(xml);
  chart.title = { text: nasty };
  const layout = layoutChart(chart, chartValues(chart), 480, 320, fakeMeasure);
  for (const item of items(layout, 'text')) {
    assert.ok(!/[\u0000-\u001f\u007f]/.test(item.text), `no control characters in ${JSON.stringify(item.text.slice(0, 20))}`);
    assert.ok(item.text.length <= 301, 'bounded');
  }
  assert.equal(texts(layout, 'title')[0], 'a b c d e');
  assert.ok(texts(layout, 'category').some((text) => text.indexOf('<img') === 0), 'markup is text like any other');
});
