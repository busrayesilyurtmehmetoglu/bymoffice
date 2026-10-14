// The layout of the newer chart types and of the surface chart: where the bars, boxes, pieces
// and rings are put in a frame, measured with a stand-in for text measurement (every
// character half the font size wide), and what a chart of very many points costs.
//
// Excel is not at hand: what follows Excel's known behaviour without having been compared
// with it is marked UNSURE in the layout and in the user guide.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { chartAutoColor, chartLayoutWeight, chartValues, layoutChart, MAX_DRAWN_ITEMS, MAX_LAYOUT_WEIGHT, type ChartItem, type ChartLayout } from '../src/layout/chart';
import type { Chart } from '../src/model/sheet/chart';
import { catAx, chartXml, fakeMeasure, group, numLit, readChart, ser, solid, spPr, strLit, valAx } from './chart-fixtures';
import { catAxis, chartExXml, CATS, data, numDim, readChartEx, series, seriesName, strDim, treemapXml, valAxis, waterfallXml } from './chartex-fixtures';

type Of<K extends ChartItem['kind']> = Extract<ChartItem, { kind: K }>;
const items = <K extends ChartItem['kind']>(layout: ChartLayout, kind: K, role?: string): Of<K>[] => layout.items.filter((item) => item.kind === kind && (role === undefined || (item as { role?: string }).role === role)) as Of<K>[];
const texts = (layout: ChartLayout, role: string): string[] => items(layout, 'text', role).map((item) => item.text);
const near = (actual: number, expected: number, message: string, tolerance = 1e-6) => assert.ok(Math.abs(actual - expected) <= tolerance, `${message}: ${actual} is not ${expected}`);

function model(xml: string): Chart {
  const chart = readChartEx(xml) || readChart(xml);
  assert.ok(chart, 'the chart is read');
  return chart as Chart;
}

function lay(chart: Chart, width = 480, height = 320, words?: (word: 'increase' | 'decrease' | 'total') => string): ChartLayout {
  return layoutChart(chart, chartValues(chart, { chartWord: words }), width, height, fakeMeasure);
}

const funnelXml = (values: number[]) =>
  chartExXml({
    data: data(0, strDim(undefined, [values.map((_v, i) => `Adım ${i + 1}`)]) + numDim(undefined, values)),
    series: series({ layout: 'funnel', dataId: 0, inner: seriesName('Huni'), axes: [0] }),
    axes: catAxis(0),
  });

const histogramXml = (values: number[], layoutPr: string, more = '') =>
  chartExXml({
    data: data(0, numDim(undefined, values)),
    series: series({ layout: 'clusteredColumn', dataId: 0, layoutPr, axes: [0, 1] }) + more,
    axes: catAxis(0, '<cx:tickLabels/>', '0') + valAxis(1),
  });

// ------------------------------------------------------------------ waterfall

test('waterfall: each bar from the running total, totals on the axis, colours by what the bar is, connector lines', () => {
  const chart = model(waterfallXml());
  const layout = lay(chart);
  const bars = items(layout, 'rect', 'bar');
  assert.equal(bars.length, 5);
  const scale = layout.scale;
  assert.ok(scale);
  const y = (value: number) => layout.plot.y + layout.plot.height - ((value - scale.min) / (scale.max - scale.min)) * layout.plot.height;
  // 100 (total), +20, -50, +30, 100 (total).
  const spans: [number, number][] = [
    [0, 100],
    [100, 120],
    [70, 120],
    [70, 100],
    [0, 100],
  ];
  bars.forEach((bar, i) => {
    near(bar.y, y(spans[i][1]), `top of bar ${i}`);
    near(bar.y + bar.height, y(spans[i][0]), `bottom of bar ${i}`);
  });
  // UNSURE: increase, decrease and total take the first three colours of the chart.
  assert.deepEqual(bars.map((bar) => bar.fill), [chartAutoColor(chart, 2), chartAutoColor(chart, 0), chartAutoColor(chart, 1), chartAutoColor(chart, 0), chartAutoColor(chart, 2)]);
  assert.ok(bars.every((bar, i) => i === 0 || bar.x > bars[i - 1].x), 'left to right');
  const lines = items(layout, 'path', 'line');
  assert.equal(lines.length, 1);
  assert.equal(lines[0].points.length, 16, 'four lines between five bars');
  near(lines[0].points[1], y(100), 'the first line at the total after the first bar');
  assert.deepEqual(texts(layout, 'legend'), ['Increase', 'Decrease', 'Total']);
  const words: Record<string, string> = { increase: 'Artış', decrease: 'Azalış', total: 'Toplam' };
  assert.deepEqual(texts(lay(chart, 480, 320, (word) => words[word]), 'legend'), ['Artış', 'Azalış', 'Toplam']);
  assert.deepEqual(texts(layout, 'category'), CATS);
  assert.deepEqual(texts(layout, 'title'), ['Nakit']);
  // No connector lines where the file says so; a colour of its own for a point.
  const plainer = model(waterfallXml([1, 2, 3, 4, 5], `<cx:dataPt idx="1"><cx:spPr>${solid('123456')}</cx:spPr></cx:dataPt>`).replace('</cx:subtotals>', '</cx:subtotals><cx:visibility connectorLines="0"/>'));
  const other = lay(plainer);
  assert.equal(items(other, 'path', 'line').length, 0);
  assert.equal(items(other, 'rect', 'bar')[1].fill, '#123456');
});

test('waterfall: data labels show the value of each bar beside its end', () => {
  const chart = model(waterfallXml([100, 20, -50, 30, 100], '<cx:dataLabels pos="outEnd"><cx:visibility seriesName="0" categoryName="0" value="1"/></cx:dataLabels>'));
  const layout = lay(chart);
  assert.deepEqual(texts(layout, 'label'), ['100', '20', '-50', '30', '100']);
  const bars = items(layout, 'rect', 'bar');
  const labels = items(layout, 'text', 'label');
  assert.ok(labels[0].y < bars[0].y, 'above a bar that goes up');
  assert.ok(labels[2].y > bars[2].y + bars[2].height, 'below a bar that goes down');
});

// ------------------------------------------------------------------ funnel

test('funnel: bars as wide as their share of the widest, centred, one below the other, names on the left', () => {
  const layout = lay(model(funnelXml([100, 50, 25, 0])));
  const bars = items(layout, 'rect', 'bar');
  assert.equal(bars.length, 3, 'a value of zero has no bar');
  near(bars[1].width, bars[0].width / 2, 'half as wide');
  near(bars[2].width, bars[0].width / 4, 'a quarter');
  const middle = bars[0].x + bars[0].width / 2;
  for (const bar of bars) near(bar.x + bar.width / 2, middle, 'centred');
  near(bars[0].width, layout.plot.width, 'the widest fills the plot area');
  assert.ok(bars[1].y > bars[0].y && bars[2].y > bars[1].y, 'top to bottom');
  assert.deepEqual(texts(layout, 'category'), ['Adım 1', 'Adım 2', 'Adım 3', 'Adım 4']);
  assert.ok(items(layout, 'text', 'category').every((text) => text.x <= layout.plot.x), 'names on the left of the bars');
});

// ------------------------------------------------------------------ histogram

test('histogram: a bar per bin as high as its count, labelled by the edges of its bin', () => {
  const layout = lay(model(histogramXml([1, 2, 3, 4, 5, 6, 7, 8, 9, 10], '<cx:binning intervalClosed="r"><cx:binSize val="3"/></cx:binning>')));
  const bars = items(layout, 'rect', 'bar');
  assert.equal(bars.length, 3);
  near(bars[0].height / bars[1].height, 4 / 3, 'four values and three');
  assert.deepEqual(texts(layout, 'category'), ['[1, 4]', '(4, 7]', '(7, 10]']);
  // Bins of their own below and above.
  const ends = lay(model(histogramXml([-5, 0, 1, 2, 3, 4, 5, 6, 50], '<cx:binning intervalClosed="r" underflow="1" overflow="5"><cx:binSize val="2"/></cx:binning>')));
  assert.deepEqual(texts(ends, 'category'), ['≤1', '(1, 3]', '(3, 5]', '>5']);
  const left = lay(model(histogramXml([1, 2, 3, 4], '<cx:binning intervalClosed="l"><cx:binCount val="2"/></cx:binning>')));
  assert.deepEqual(texts(left, 'category'), ['[1, 2.5)', '[2.5, 4]']);
});

test('pareto: the bars from the largest down and the running share on a second axis that ends at 100%', () => {
  const chart = model(
    chartExXml({
      data: data(0, strDim(undefined, [['a', 'b', 'a', 'c', 'b', 'a']]) + numDim(undefined, [1, 1, 1, 1, 1, 1])),
      series: series({ layout: 'clusteredColumn', dataId: 0, layoutPr: '<cx:aggregation/>', axes: [0, 1] }) + series({ layout: 'paretoLine', attrs: 'ownerIdx="0"', axes: [2] }),
      axes: catAxis(0) + valAxis(1) + valAxis(2, '<cx:tickLabels/>', 'max="1" min="0"'),
    }),
  );
  const layout = lay(chart);
  assert.deepEqual(texts(layout, 'category'), ['a', 'b', 'c'], 'categories summed, largest first');
  const bars = items(layout, 'rect', 'bar');
  near(bars[0].height / bars[2].height, 3, 'three a and one c');
  const [line] = items(layout, 'path', 'line');
  assert.ok(line);
  assert.equal(line.points.length, 6);
  near(line.points[5], layout.plot.y, 'the share ends at the top: 100%');
  near(line.points[1], layout.plot.y + layout.plot.height * 0.5, 'a is half of all');
  assert.ok(texts(layout, 'tick').indexOf('100%') >= 0, 'the second axis is labelled in per cent');
  assert.equal(line.stroke, chartAutoColor(chart, 1), 'the line takes the colour after that of the bars');
});

// ------------------------------------------------------------------ box and whisker

test('box and whisker: the box from the first to the third quartile, whiskers, outliers and the mean', () => {
  const values = [2, 4, 4, 5, 7, 9, 10, 12, 30];
  const chart = model(
    chartExXml({
      data: data(0, numDim(undefined, values)),
      series: series({ layout: 'boxWhisker', dataId: 0, inner: seriesName('Ölçüm'), layoutPr: '<cx:statistics quartileMethod="inclusive"/><cx:visibility meanLine="0" meanMarker="1" nonoutliers="0" outliers="1"/>', axes: [0, 1] }),
      axes: catAxis(0) + valAxis(1),
    }),
  );
  const layout = lay(chart);
  const scale = layout.scale;
  assert.ok(scale);
  const y = (value: number) => layout.plot.y + layout.plot.height - ((value - scale.min) / (scale.max - scale.min)) * layout.plot.height;
  const [box] = items(layout, 'rect', 'bar');
  near(box.y, y(10), 'the top of the box is the third quartile');
  near(box.y + box.height, y(4), 'its bottom the first');
  const markers = items(layout, 'markers', 'marker');
  const outliers = markers.find((item) => item.symbol === 'circle');
  assert.ok(outliers);
  assert.equal(outliers.points.length, 2, 'one outlier: 30');
  near(outliers.points[1], y(30), 'where 30 is');
  const mean = markers.find((item) => item.symbol === 'x');
  assert.ok(mean);
  near(mean.points[1], y(83 / 9), 'the mean');
  const [lines] = items(layout, 'path', 'line');
  // The median, two whiskers and their two ends: five lines.
  assert.equal(lines.points.length, 20);
  near(lines.points[1], y(7), 'the median');
  const ends = [lines.points[13], lines.points[17]].sort((a, b) => a - b);
  near(ends[0], y(12), 'the upper whisker ends at the highest value within reach');
  near(ends[1], y(2), 'the lower at the lowest');
});

test('box and whisker: boxes of several series side by side in each category, in the order the categories first come', () => {
  const chart = model(
    chartExXml({
      data: data(0, strDim(undefined, [['K', 'L', 'K', 'L']]) + numDim(undefined, [1, 5, 2, 6])) + data(1, strDim(undefined, [['K', 'L', 'K', 'L']]) + numDim(undefined, [3, 7, 4, 8])),
      series: series({ layout: 'boxWhisker', dataId: 0, inner: seriesName('Bir'), formatIdx: 0 }) + series({ layout: 'boxWhisker', dataId: 1, inner: seriesName('İki'), formatIdx: 1 }),
      axes: catAxis(0) + valAxis(1),
      legend: '<cx:legend pos="b"/>',
    }),
  );
  const layout = lay(chart);
  const boxes = items(layout, 'rect', 'bar');
  assert.equal(boxes.length, 4);
  assert.deepEqual(texts(layout, 'category'), ['K', 'L']);
  assert.deepEqual(texts(layout, 'legend'), ['Bir', 'İki']);
  const [k1, l1, k2, l2] = boxes;
  assert.ok(k1.x < k2.x && k2.x < l1.x && l1.x < l2.x, 'in each category the first series, then the second');
  assert.equal(k1.fill, chartAutoColor(chart, 0));
  assert.equal(k2.fill, chartAutoColor(chart, 1));
});

// ------------------------------------------------------------------ treemap

test('treemap: leaves as large as their values fill the plot area, coloured by their branch, names on top', () => {
  const chart = model(treemapXml());
  const layout = lay(chart, 480, 320);
  const rects = items(layout, 'rect', 'bar');
  // Two branches, three parents, six leaves.
  assert.equal(rects.length, 11);
  const leaves = rects.filter((rect) => rect.point !== undefined && ['a1', 'a2', 'b1', 'b2', 'c1', 'c2'].length > 0 && rect.width * rect.height > 0).slice(-6);
  const plot = layout.plot.width * layout.plot.height;
  const total = 100;
  const sizes = [10, 20, 30, 5, 15, 20];
  for (const leaf of leaves) near((leaf.width * leaf.height) / plot, sizes[leaf.point as number] / total, `leaf ${leaf.point}`, 1e-6);
  // "Kök" (65) is the larger branch: the first colour; "Diğer" (35) the second.
  const colorOf = (point: number) => (leaves.find((leaf) => leaf.point === point) as Of<'rect'>).fill;
  assert.equal(colorOf(0), chartAutoColor(chart, 0));
  assert.equal(colorOf(2), chartAutoColor(chart, 0));
  assert.equal(colorOf(4), chartAutoColor(chart, 1));
  const labels = texts(layout, 'label');
  for (const name of ['a1', 'a2', 'b1', 'b2', 'c1', 'c2', 'Kök', 'Diğer', 'A', 'B', 'C']) assert.ok(labels.indexOf(name) >= 0, `${name} is shown`);
  assert.ok(labels.indexOf('Kök') > labels.indexOf('c1'), 'the names of parents lie over those of the leaves');
  // A banner takes room at the top of each parent.
  const banner = lay(model(treemapXml('treemap', '<cx:parentLabelLayout val="banner"/>')), 480, 320);
  const top = items(banner, 'rect', 'bar')[0];
  const child = items(banner, 'rect', 'bar').find((rect) => rect !== top && rect.x >= top.x && rect.x < top.x + top.width && rect.y >= top.y && rect.y < top.y + top.height);
  assert.ok(child && child.y > top.y, 'the children start below the band');
  const none = lay(model(treemapXml('treemap', '<cx:parentLabelLayout val="none"/>')), 480, 320);
  assert.ok(texts(none, 'label').indexOf('Kök') < 0);
});

// ------------------------------------------------------------------ sunburst

test('sunburst: rings from the outermost level inwards, each piece as wide as its share of its parent', () => {
  const chart = model(treemapXml('sunburst', ''));
  const layout = lay(chart, 400, 400);
  const slices = items(layout, 'slice');
  assert.equal(slices.length, 11);
  const turn = Math.PI * 2;
  const inner = slices.filter((slice) => slice.inner === Math.min(...slices.map((other) => other.inner)));
  assert.deepEqual(inner.length, 2, 'two branches in the innermost ring');
  near(inner[0].end - inner[0].start, (65 / 100) * turn, 'Kök: 65 of 100');
  near(inner[0].start, -Math.PI / 2, 'from the top');
  near(inner[1].start, inner[0].end, 'clockwise, one after the other');
  const rings = new Set(slices.map((slice) => Math.round(slice.inner * 1000)));
  assert.equal(rings.size, 3, 'three levels, three rings');
  const leaf = slices.find((slice) => slice.point === 2);
  assert.ok(leaf);
  near(leaf.end - leaf.start, (30 / 100) * turn, 'b1: 30 of 100');
  assert.ok(leaf.radius > inner[0].radius);
  assert.equal(leaf.fill, chartAutoColor(chart, 0), 'the colour of its branch');
  // Every slice starts where the pieces of its ring before it end: the rings close.
  const outer = slices.filter((slice) => slice.inner === Math.max(...slices.map((other) => other.inner)));
  near(outer.reduce((sum, slice) => sum + slice.end - slice.start, 0), turn, 'the outer ring is whole');
});

// ------------------------------------------------------------------ surface

const surfaceXml = (name = 'surfaceChart', head = '', rows = [[0, 5, 10], [5, 10, 15], [10, 15, 20]]) =>
  chartXml({
    plot:
      group(
        name,
        head,
        rows.map((values, i) => ser(i, { tx: `<c:v>R${i + 1}</c:v>`, cat: strLit(['x', 'y', 'z']), val: numLit(values) })).join(''),
        `<c:bandFmts><c:bandFmt><c:idx val="0"/>${spPr(solid('0000FF'))}</c:bandFmt></c:bandFmts><c:axId val="111"/><c:axId val="222"/><c:axId val="333"/>`,
        false,
      ) +
      catAx() +
      valAx('<c:majorGridlines/><c:majorUnit val="10"/>', '<c:orientation val="minMax"/><c:max val="20"/><c:min val="0"/>') +
      '<c:serAx><c:axId val="333"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:tickLblPos val="nextTo"/><c:crossAx val="222"/></c:serAx>',
  }).replace('</c:plotArea>', '</c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>');

test('surface: seen from above, every point of the grid coloured by the band its value lies in; the legend names the bands', () => {
  const chart = model(surfaceXml());
  const layout = lay(chart);
  const areas = items(layout, 'path', 'area');
  assert.ok(areas.length > 0, 'pieces are drawn');
  const fills = new Set(areas.map((area) => area.fill));
  assert.deepEqual(Array.from(fills).sort(), ['#0000FF', chartAutoColor(chart, 1)].sort(), 'the band below 10 in its own colour, the band above in the second colour');
  assert.deepEqual(texts(layout, 'legend'), ['0-10', '10-20']);
  // The corner of the lowest value (bottom left) lies in the first band, that of the highest (top right) in the second.
  const plot = layout.plot;
  const inside = (area: Of<'path'>, x: number, y: number) => {
    let hit = false;
    const p = area.points;
    for (let i = 0, j = p.length - 2; i < p.length; j = i, i += 2) {
      if (p[i + 1] > y !== p[j + 1] > y && x < ((p[j] - p[i]) * (y - p[i + 1])) / (p[j + 1] - p[i + 1]) + p[i]) hit = !hit;
    }
    return hit;
  };
  const at = (x: number, y: number) => areas.find((area) => inside(area, x, y));
  assert.equal(at(plot.x + 2, plot.y + plot.height - 2)?.fill, '#0000FF');
  assert.equal(at(plot.x + plot.width - 2, plot.y + 2)?.fill, chartAutoColor(chart, 1));
  // The categories below, the names of the series beside their rows.
  assert.deepEqual(texts(layout, 'category'), ['x', 'y', 'z', 'R1', 'R2', 'R3']);
  // A wireframe draws the lines between the bands alone.
  const wire = lay(model(surfaceXml('surfaceChart', '<c:wireframe val="1"/>')));
  assert.equal(items(wire, 'path', 'area').length, 0);
  const lines = items(wire, 'path', 'line');
  assert.equal(lines.length, 1, 'one line: where the values are 10');
  assert.equal(lines[0].stroke, '#0000FF');
  assert.ok(lines[0].segments);
  // The 3-D form is drawn the same way.
  assert.ok(items(lay(model(surfaceXml('surface3DChart'))), 'path', 'area').length > 0);
});

// ------------------------------------------------------------------ cost

test('what a chart of very many points costs to lay out is bounded, whatever its type', () => {
  const many = 100000;
  const names = Array.from({ length: many }, (_v, i) => `n${i}`);
  const parents = Array.from({ length: many }, (_v, i) => `p${i % 500}`);
  const values = Array.from({ length: many }, (_v, i) => (i % 97) + 1);
  const big = (layout: string, layoutPr = '') =>
    model(chartExXml({ data: data(0, strDim(undefined, layout === 'treemap' || layout === 'sunburst' ? [names, parents] : [names]) + numDim(undefined, values, layout === 'treemap' || layout === 'sunburst' ? 'size' : 'val')), series: series({ layout, dataId: 0, layoutPr }) }));
  for (const [layout, layoutPr] of [
    ['treemap', ''],
    ['sunburst', ''],
    ['waterfall', ''],
    ['funnel', ''],
    ['clusteredColumn', '<cx:binning><cx:binCount val="100000"/></cx:binning>'],
    ['clusteredColumn', '<cx:aggregation/>'],
    ['boxWhisker', ''],
  ]) {
    const chart = big(layout, layoutPr);
    const started = Date.now();
    const result = lay(chart, 800, 600);
    const took = Date.now() - started;
    assert.ok(chartLayoutWeight(result) <= MAX_LAYOUT_WEIGHT, `${layout}: ${chartLayoutWeight(result)}`);
    assert.ok(result.items.length <= MAX_DRAWN_ITEMS * 3, `${layout}: ${result.items.length} items`);
    // A guard against hanging, not a measure of speed.
    assert.ok(took < 20000, `${layout}: ${took} ms`);
  }
  // A surface of many rows and columns.
  const rows = Array.from({ length: 60 }, (_v, r) => Array.from({ length: 3000 }, (_w, c) => Math.sin(r / 5) * Math.cos(c / 50) * 100));
  const surface = model(
    chartXml({
      plot: group('surfaceChart', '', rows.map((row, i) => ser(i, { val: numLit(row) })).join('')) + catAx() + valAx(),
    }),
  );
  const layout = lay(surface, 800, 600);
  assert.ok(chartLayoutWeight(layout) <= MAX_LAYOUT_WEIGHT, `surface: ${chartLayoutWeight(layout)}`);
  assert.ok(items(layout, 'path', 'area').length <= MAX_DRAWN_ITEMS + 16);
});

test('a fill the series states is the colour of its bars and boxes', () => {
  const fill = `<cx:spPr>${solid('00B050')}</cx:spPr>`;
  const cases: [string, string][] = [
    ['funnel', chartExXml({ data: data(0, strDim(undefined, [['a', 'b']]) + numDim(undefined, [2, 1])), series: series({ layout: 'funnel', dataId: 0, inner: fill }) })],
    ['histogram', chartExXml({ data: data(0, numDim(undefined, [1, 2, 3])), series: series({ layout: 'clusteredColumn', dataId: 0, inner: fill }) })],
    ['box', chartExXml({ data: data(0, numDim(undefined, [1, 2, 3, 4])), series: series({ layout: 'boxWhisker', dataId: 0, inner: fill }) })],
  ];
  for (const [label, xml] of cases) {
    const bars = items(lay(model(xml)), 'rect', 'bar');
    assert.ok(bars.length > 0, label);
    for (const bar of bars) assert.equal(bar.fill, '#00B050', label);
  }
  // The legend shows it too.
  const legend = lay(model(cases[0][1].replace('</cx:chart>', '<cx:legend pos="b"/></cx:chart>')));
  assert.equal(items(legend, 'rect', 'swatch')[0]?.fill, '#00B050');
});

test('frames too small for words show the shapes alone; a frame of no size shows nothing', () => {
  for (const xml of [waterfallXml(), treemapXml(), treemapXml('sunburst', ''), funnelXml([3, 2, 1])]) {
    const chart = model(xml);
    const small = lay(chart, 70, 50);
    assert.equal(small.detail, 'bare');
    assert.equal(items(small, 'text').length, 0);
    assert.ok(small.items.length > 1, 'something is drawn');
    assert.deepEqual(lay(chart, 0, 100).items, []);
  }
  void ser;
});
