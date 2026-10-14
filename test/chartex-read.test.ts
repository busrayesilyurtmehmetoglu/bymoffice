// The reader of the parts of the newer chart types (chartex) and of surface charts: what of
// <cx:chartSpace> reaches the model, which parts are no chart of the model (a region map, an
// unknown layout, another namespace), and what a damaged or hostile part costs.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { ChartBudgetError, MAX_CHART_CACHE, parseChart } from '../src/formats/xlsx/chart';
import { chartKnown } from '../src/formats/xlsx/chart-write';
import { isChartExRoot, parseChartEx } from '../src/formats/xlsx/chartex';
import { chartCacheList, chartWritable, type Chart } from '../src/model/sheet/chart';
import { parseXml } from '../src/xml';
import { catAx, chartXml, group, numRef, OFFICE_THEME, readChart, ser, solid, spPr, strRef, valAx } from './chart-fixtures';
import { catAxis, chartExXml, CATS, cxTitle, data, NS_CX1, numDim, readChartEx, series, seriesName, strDim, treemapXml, valAxis, waterfallXml } from './chartex-fixtures';

function must(xml: string): Chart {
  const chart = readChartEx(xml);
  assert.ok(chart, 'the chart is read');
  return chart as Chart;
}

const plain = (layout: string, layoutPr = '', more = '') =>
  chartExXml({
    data: data(0, strDim('Sayfa1!$A$2:$A$6', [CATS]) + numDim('Sayfa1!$B$2:$B$6', [1, 2, 3, 4, 5])),
    series: series({ layout, dataId: 0, layoutPr, axes: [0, 1] }) + more,
    axes: catAxis() + valAxis(),
  });

test('chartex: every layout that is drawn is read as its kind; the chart is only shown and never written', () => {
  const kinds: [string, string][] = [
    ['clusteredColumn', 'histogram'],
    ['waterfall', 'waterfall'],
    ['funnel', 'funnel'],
    ['boxWhisker', 'boxWhisker'],
    ['treemap', 'treemap'],
    ['sunburst', 'sunburst'],
  ];
  for (const [layout, kind] of kinds) {
    const chart = must(plain(layout));
    assert.equal(chart.groups.length, 1, layout);
    assert.equal(chart.groups[0].kind, kind, layout);
    assert.equal(chart.extended, true);
    assert.equal(chart.shownOnly, true);
    assert.equal(chartWritable(chart), false, `${layout}: not written from the model`);
  }
});

test('chartex: the data of a series by its number, with the formulas, the cache and the levels', () => {
  const chart = must(waterfallXml());
  const group = chart.groups[0];
  const [only] = group.series;
  assert.deepEqual(only.name, { ref: 'Sayfa1!$B$1', text: 'Akış' });
  assert.deepEqual(only.categories, { kind: 'str', ref: 'Sayfa1!$A$2:$A$6', cache: { index: [0, 1, 2, 3, 4], values: CATS }, count: 5 });
  assert.deepEqual(only.values, { kind: 'num', ref: 'Sayfa1!$B$2:$B$6', formatCode: 'General', cache: { index: [0, 1, 2, 3, 4], values: [100, 20, -50, 30, 100] }, count: 5 });
  assert.deepEqual(only.ex, { subtotals: [0, 4] });
  assert.equal(group.gapWidth, 50, 'the gap of the axis of the categories (0.5) as a percentage');
  assert.deepEqual(chart.title, { text: 'Nakit' });
  assert.deepEqual(chart.legend, { position: 't' });
  assert.deepEqual(
    chart.axes.map((axis) => [axis.id, axis.kind, axis.tickLabels, axis.majorGridlines !== undefined]),
    [
      [0, 'cat', 'nextTo', false],
      [1, 'val', 'nextTo', true],
    ],
  );
  const tree = must(treemapXml());
  const cats = tree.groups[0].series[0].categories;
  assert.equal(cats?.multiLevel, true);
  assert.deepEqual(chartCacheList(cats), ['a1', 'a2', 'b1', 'b2', 'c1', 'c2'], 'the first level is that of the leaves');
  assert.deepEqual(cats?.levels?.map((level) => level.values), [
    ['A', 'A', 'B', 'B', 'C', 'C'],
    ['Kök', 'Kök', 'Kök', 'Kök', 'Diğer', 'Diğer'],
  ]);
  assert.deepEqual(chartCacheList(tree.groups[0].series[0].values), [10, 20, 30, 5, 15, 20], 'the sizes are the values');
  assert.deepEqual(tree.groups[0].series[0].ex, { parentLabels: 'overlapping' });
  assert.deepEqual(tree.groups[0].series[0].dataLabels, { position: 'inEnd', showValue: false, showCategory: true, showSeries: false });
  // Levels by rows.
  const rows = must(chartExXml({ data: data(0, strDim('S!$A$1:$C$2', [['x', 'y', 'z'], ['X', 'X', 'Y']], 'cat', 'row') + numDim(undefined, [1, 2, 3], 'size')), series: series({ layout: 'sunburst', dataId: 0 }) }));
  assert.equal(rows.groups[0].series[0].categories?.byRow, true);
});

test('chartex: binning, aggregation, statistics and what is shown of a box', () => {
  const bins = must(plain('clusteredColumn', '<cx:binning intervalClosed="l" underflow="2" overflow="auto"><cx:binSize val="2.5"/></cx:binning>'));
  assert.deepEqual(bins.groups[0].series[0].ex, { binning: { closed: 'l', underflow: 2, overflow: 'auto', size: 2.5 } });
  const counted = must(plain('clusteredColumn', '<cx:binning intervalClosed="r"><cx:binCount val="7"/></cx:binning>'));
  assert.deepEqual(counted.groups[0].series[0].ex, { binning: { closed: 'r', count: 7 } });
  const huge = must(plain('clusteredColumn', '<cx:binning><cx:binCount val="999999999"/></cx:binning>'));
  assert.equal(huge.groups[0].series[0].ex?.binning?.count, 10000, 'a number of bins is held to a limit');
  assert.deepEqual(must(plain('clusteredColumn', '<cx:aggregation/>')).groups[0].series[0].ex, { aggregate: true });
  const box = must(plain('boxWhisker', '<cx:statistics quartileMethod="inclusive"/><cx:visibility meanLine="1" meanMarker="0" nonoutliers="1" outliers="0"/>'));
  assert.deepEqual(box.groups[0].series[0].ex, { quartiles: 'inclusive', meanLine: true, meanMarker: false, nonoutliers: true, outliers: false });
  const water = must(plain('waterfall', '<cx:subtotals><cx:idx val="4"/><cx:idx val="-1"/><cx:idx val="4"/><cx:idx val="2"/></cx:subtotals><cx:visibility connectorLines="0"/>'));
  assert.deepEqual(water.groups[0].series[0].ex, { subtotals: [2, 4], connectorLines: false });
  assert.deepEqual(must(plain('treemap', '<cx:parentLabelLayout val="banner"/>')).groups[0].series[0].ex, { parentLabels: 'banner' });
});

test('chartex: a pareto line belongs to the histogram it names', () => {
  const chart = must(
    plain(
      'clusteredColumn',
      '<cx:aggregation/>',
      series({ layout: 'paretoLine', attrs: 'ownerIdx="0"', inner: '<cx:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></cx:spPr>', axes: [2] }),
    ),
  );
  assert.equal(chart.groups.length, 1, 'the line is no chart type of its own');
  assert.deepEqual(chart.groups[0].series[0].ex, { aggregate: true, pareto: { line: { color: '#FF0000', width: 3 }, axisId: 2 } });
});

test('chartex: series of one layout make one chart type; colours of the series, points and labels', () => {
  const two = must(
    chartExXml({
      data: data(0, strDim(undefined, [['a', 'a', 'b']]) + numDim(undefined, [1, 2, 3])) + data(1, strDim(undefined, [['a', 'b', 'b']]) + numDim(undefined, [4, 5, 6])),
      series:
        series({ layout: 'boxWhisker', dataId: 0, formatIdx: 0, inner: seriesName('Bir') + `<cx:spPr>${solid('112233')}</cx:spPr>` }) +
        series({ layout: 'boxWhisker', dataId: 1, formatIdx: 1, inner: seriesName('İki') + `<cx:dataPt idx="2"><cx:spPr>${solid('445566')}</cx:spPr></cx:dataPt>` }),
    }),
  );
  assert.equal(two.groups.length, 1);
  assert.deepEqual(two.groups[0].series.map((item) => [item.index, item.name?.text, item.shape?.fill]), [
    [0, 'Bir', '#112233'],
    [1, 'İki', undefined],
  ]);
  assert.deepEqual(two.groups[0].series[1].points, [{ index: 2, shape: { fill: '#445566' } }]);
  const labelled = must(
    plain('waterfall', '').replace(
      '<cx:dataId',
      '<cx:dataLabels pos="outEnd"><cx:numFmt formatCode="#,##0" sourceLinked="0"/><cx:visibility seriesName="0" categoryName="0" value="1"/><cx:dataLabelHidden idx="1"/><cx:dataLabel idx="2"><cx:visibility value="0" categoryName="1"/></cx:dataLabel></cx:dataLabels><cx:dataId',
    ),
  );
  const only = labelled.groups[0].series[0];
  assert.deepEqual(only.dataLabels, { showValue: true, showCategory: false, showSeries: false, position: 'outEnd', formatCode: '#,##0', sourceLinked: false });
  assert.deepEqual(only.pointLabels, [
    { index: 1, deleted: true },
    { index: 2, showValue: false, showCategory: true },
  ]);
});

test('chartex: what is no chart of the model is not read, and nothing is thrown', () => {
  const cases: [string, string][] = [
    ['a region map', plain('regionMap')],
    ['a layout of a later version', plain('somethingNew')],
    ['a region map beside a funnel', plain('funnel', '', series({ layout: 'regionMap', dataId: 0 }))],
    ['no series', chartExXml({ data: data(0, numDim(undefined, [1])), series: '' })],
    ['a series without data', chartExXml({ data: '', series: series({ layout: 'funnel', dataId: 4 }) })],
    ['another namespace', plain('funnel').replace(/xmlns:cx="[^"]*"/, 'xmlns:cx="urn:other"')],
    ['the older chart part', chartXml({ plot: group('barChart', '<c:barDir val="col"/>', ser(0, { val: numRef('S!$B$2:$B$4', [1, 2, 3]) })) + catAx() + valAx() })],
    ['another root', '<cx:chartData xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"/>'],
  ];
  for (const [name, xml] of cases) assert.equal(readChartEx(xml), undefined, name);
  assert.equal(parseChartEx(undefined as never, OFFICE_THEME), undefined, 'no element at all');
  assert.ok(isChartExRoot(parseXml(plain('funnel').replace(/2014\/chartex/g, '2015/9/8/chartex')).root), 'a later namespace of the same kind');
  // A hidden series is left out; the older reader takes nothing of a newer part.
  assert.equal(must(plain('funnel', '', series({ layout: 'funnel', dataId: 0, attrs: 'hidden="1"' }))).groups[0].series.length, 1);
  assert.equal(parseChart(parseXml(plain('funnel')).root, OFFICE_THEME), undefined);
});

test('chartex: cached values count against the budget of the workbook like those of other charts', () => {
  const root = parseXml(waterfallXml()).root;
  const budget = { points: 100 };
  assert.ok(parseChartEx(root, OFFICE_THEME, budget));
  assert.equal(budget.points, 90, 'five categories and five values');
  assert.throws(() => parseChartEx(root, OFFICE_THEME, { points: 9 }), ChartBudgetError);
  // A part of very many points is cut where a chart is; it costs no more than it holds.
  const many = Array.from({ length: 2000 }, (_v, i) => i);
  const big = chartExXml({ data: data(0, numDim(undefined, many)), series: series({ layout: 'clusteredColumn', dataId: 0 }) });
  const read = must(big);
  assert.equal(read.groups[0].series[0].values?.cache.index.length, 2000);
  assert.ok(MAX_CHART_CACHE >= 2000);
  // Points far beyond the limit and indexes out of order cost what they hold.
  const sparse = chartExXml({ data: data(0, `<cx:numDim type="val"><cx:lvl ptCount="99999999"><cx:pt idx="5">5</cx:pt><cx:pt idx="1">1</cx:pt><cx:pt idx="900000000">9</cx:pt></cx:lvl></cx:numDim>`), series: series({ layout: 'funnel', dataId: 0 }) });
  const values = must(sparse).groups[0].series[0].values;
  assert.deepEqual(values?.cache, { index: [1, 5], values: [1, 5] });
  assert.ok((values?.count || 0) <= 100000);
  assert.ok(must(sparse).notes.some((note) => /first 100000 points/.test(note)));
});

test('chartex: deep levels are held to a limit; a title, a later namespace on the Choice, text styles', () => {
  const levels = Array.from({ length: 20 }, (_v, l) => ['n' + l]);
  const deep = must(chartExXml({ data: data(0, strDim(undefined, levels) + numDim(undefined, [1], 'size')), series: series({ layout: 'treemap', dataId: 0 }) }));
  assert.equal(deep.groups[0].series[0].categories?.levels?.length, 7, 'eight levels: the leaves and seven above');
  assert.ok(deep.notes.some((note) => /levels of categories/.test(note)));
  const titled = must(chartExXml({ data: data(0, numDim(undefined, [1])), series: series({ layout: 'funnel', dataId: 0 }), title: cxTitle('Huni').replace('pos="t"', 'pos="b"') }));
  assert.equal(titled.title?.text, 'Huni');
  assert.ok(titled.notes.some((note) => /top/.test(note)));
  assert.ok(NS_CX1.endsWith('/chartex'));
});

// ------------------------------------------------------------------ surface

test('surface: a surface chart is read with its wireframe and bands; the 3-D form is drawn from above', () => {
  const surface = (name: string, head: string, tail = '') =>
    chartXml({
      plot:
        group(name, head, ser(0, { tx: '<c:v>Bir</c:v>', cat: strRef('S!$A$2:$A$4', ['a', 'b', 'c']), val: numRef('S!$B$2:$B$4', [1, 2, 3]) }) + ser(1, { tx: '<c:v>İki</c:v>', cat: strRef('S!$A$2:$A$4', ['a', 'b', 'c']), val: numRef('S!$C$2:$C$4', [3, 2, 1]) }), tail) +
        catAx() +
        valAx() +
        '<c:serAx><c:axId val="333"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:tickLblPos val="nextTo"/><c:crossAx val="222"/></c:serAx>',
    });
  const flat = readChart(surface('surfaceChart', '<c:wireframe val="1"/>', `<c:bandFmts><c:bandFmt><c:idx val="1"/>${spPr(solid('ABCDEF'))}</c:bandFmt><c:bandFmt><c:idx val="0"/>${spPr(solid('123456'))}</c:bandFmt></c:bandFmts>`));
  assert.ok(flat);
  assert.equal(flat.groups[0].kind, 'surface');
  assert.equal(flat.groups[0].wireframe, true);
  assert.deepEqual(flat.groups[0].bands, [
    { index: 0, shape: { fill: '#123456' } },
    { index: 1, shape: { fill: '#ABCDEF' } },
  ]);
  assert.equal(flat.groups[0].series.length, 2);
  assert.ok(!flat.notes.some((note) => /series axis/.test(note)), 'the series axis of a surface is drawn');
  const deep = readChart(surface('surface3DChart', '<c:wireframe val="0"/>'));
  assert.ok(deep);
  assert.deepEqual([deep.groups[0].kind, deep.groups[0].threeD, deep.groups[0].wireframe], ['surface', true, undefined]);
  assert.ok(deep.notes.some((note) => /drawn from above/.test(note)));
  assert.equal(chartWritable(deep), false, 'a surface chart is not written from the model');
  // Its part is not changed in place either: no edit knows its series.
  assert.equal(chartKnown(parseXml(surface('surfaceChart', '')).root, flat, false), undefined);
});
