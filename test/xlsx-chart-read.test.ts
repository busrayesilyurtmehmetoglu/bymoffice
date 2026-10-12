// The reader of chart parts: what of <c:chartSpace> reaches the model, which
// charts are not charts of the model (and stay placeholders), and what a
// damaged or hostile part costs.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { MAX_CHART_CACHE, chartCachedPoints, parseChart } from '../src/formats/xlsx/chart';
import { MAX_DRAWING_CHARTS } from '../src/formats/xlsx/drawings';
import { open, type Chart, type DrawingChart } from '../src/index';
import { MAX_CHART_POINTS as POINTS, MAX_CHART_SERIES as SERIES, chartCacheList, chartCached } from '../src/model/sheet/chart';
import { parseXml } from '../src/xml';
import {
  NS_C,
  OFFICE_THEME,
  catAx,
  chartFrameOf,
  chartWorkbook,
  chartXml,
  group,
  legend,
  numLit,
  numRef,
  readChart,
  ser,
  sheetData,
  solid,
  spPr,
  strLit,
  strRef,
  title,
  valAx,
  xyAxes,
  xySer,
} from './chart-fixtures';
import { chartFrame, twoCell } from './sheet-drawing-fixtures';

const simple = (index = 0, props = '') => ser(index, { cat: strRef('S!$A$2:$A$4', ['a', 'b', 'c']), val: numRef('S!$B$2:$B$4', [1, 2, 3]), props });

function must(xml: string, theme?: string[]): Chart {
  const chart = readChart(xml, theme);
  assert.ok(chart, 'the chart is read');
  return chart as Chart;
}

test('a bar chart: direction, grouping, gap and overlap, the series with references and cached values', () => {
  const chart = must(
    chartXml({
      plot:
        group(
          'barChart',
          '<c:barDir val="bar"/><c:grouping val="percentStacked"/><c:varyColors val="0"/>',
          ser(0, { tx: strRef('S!$B$1', ['Adet']), cat: strRef("'S a'!$A$2:$A$4", ['a', 'b', 'c']), val: numRef('S!$B$2:$B$4', [1, null, 3], '0.0') }) + ser(1, { tx: '<c:v>Elle</c:v>', val: numLit([4, 5]) }),
          '<c:gapWidth val="80"/><c:overlap val="100"/>',
        ) +
        catAx() +
        valAx(),
    }),
  );
  assert.equal(chart.groups.length, 1);
  const bar = chart.groups[0];
  assert.deepEqual([bar.kind, bar.barDir, bar.grouping, bar.gapWidth, bar.overlap, bar.varyColors], ['bar', 'bar', 'percentStacked', 80, 100, false]);
  assert.deepEqual(bar.axisIds, [111, 222]);
  const [first, second] = bar.series;
  assert.deepEqual(first.name, { ref: 'S!$B$1', text: 'Adet' });
  assert.deepEqual(first.categories, { kind: 'str', cache: { index: [0, 1, 2], values: ['a', 'b', 'c'] }, count: 3, ref: "'S a'!$A$2:$A$4" });
  assert.deepEqual(first.values, { kind: 'num', cache: { index: [0, 2], values: [1, 3] }, count: 3, ref: 'S!$B$2:$B$4', formatCode: '0.0' }, 'one entry per point the file has');
  assert.deepEqual(chartCacheList(first.values), [1, null, 3], 'as a list: the place of the missing point is empty');
  assert.deepEqual(second.name, { text: 'Elle' });
  assert.deepEqual(second.values, { kind: 'num', cache: { index: [0, 1], values: [4, 5] }, count: 2, formatCode: 'General' });
  assert.equal(second.values?.ref, undefined, 'literal data has no reference');
});

test('every chart type that is drawn is read as its kind; the 3-D forms as their flat ones, with a note', () => {
  const cases: [string, string, string][] = [
    ['barChart', 'bar', '<c:barDir val="col"/><c:grouping val="clustered"/>'],
    ['bar3DChart', 'bar', '<c:barDir val="col"/><c:grouping val="stacked"/>'],
    ['lineChart', 'line', '<c:grouping val="standard"/>'],
    ['line3DChart', 'line', '<c:grouping val="standard"/>'],
    ['pieChart', 'pie', ''],
    ['pie3DChart', 'pie', ''],
    ['doughnutChart', 'doughnut', ''],
    ['areaChart', 'area', '<c:grouping val="stacked"/>'],
    ['area3DChart', 'area', '<c:grouping val="standard"/>'],
  ];
  for (const [name, kind, head] of cases) {
    const chart = must(chartXml({ plot: group(name, head, simple()) }));
    assert.equal(chart.groups[0].kind, kind, name);
    assert.equal(chart.groups[0].threeD === true, /3D/.test(name), `${name}: 3-D or not`);
    assert.equal(chart.notes.some((note) => /3-D chart is drawn flat/.test(note)), /3D/.test(name), `${name}: the note about 3-D`);
  }
  const scatter = must(chartXml({ plot: group('scatterChart', '<c:scatterStyle val="lineMarker"/>', xySer(0, { cat: numRef('S!$A$2:$A$4', [1, 2, 3]), val: numRef('S!$B$2:$B$4', [4, 5, 6]) })) + xyAxes() }));
  assert.equal(scatter.groups[0].kind, 'scatter');
  assert.equal(scatter.groups[0].scatterStyle, 'lineMarker');
  assert.deepEqual(chartCacheList(scatter.groups[0].series[0].categories), [1, 2, 3], 'the x values are the categories of a scatter series');
  assert.deepEqual(chartCacheList(scatter.groups[0].series[0].values), [4, 5, 6]);
});

test('a chart with a type that is not drawn is no chart of the model, whatever else it holds', () => {
  for (const name of ['surfaceChart', 'surface3DChart', 'contourChart']) {
    assert.equal(readChart(chartXml({ plot: group(name, '', simple()) + catAx() + valAx() })), undefined, name);
    assert.equal(readChart(chartXml({ plot: group('barChart', '<c:barDir val="col"/>', simple()) + group(name, '', simple(1)) + catAx() + valAx() })), undefined, `${name} beside a bar chart`);
  }
});

test('what is no chart at all is not read, and nothing is thrown', () => {
  const cases: [string, string][] = [
    ['another root', `<c:userShapes xmlns:c="${NS_C}"/>`],
    ['no chart element', `<c:chartSpace xmlns:c="${NS_C}"/>`],
    ['no plot area', `<c:chartSpace xmlns:c="${NS_C}"><c:chart/></c:chartSpace>`],
    ['an empty plot area', `<c:chartSpace xmlns:c="${NS_C}"><c:chart><c:plotArea/></c:chart></c:chartSpace>`],
    ['a chart type without series', chartXml({ plot: group('barChart', '<c:barDir val="col"/>', '') + catAx() + valAx() })],
    ['an extended chart', '<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart><cx:plotArea/></cx:chart></cx:chartSpace>'],
  ];
  for (const [name, xml] of cases) assert.equal(readChart(xml), undefined, name);
  assert.equal(parseChart(undefined as never, OFFICE_THEME), undefined, 'no element at all');
  assert.equal(parseChart(parseXml(chartXml({ plot: group('pieChart', '', simple()) })).root, undefined as never)?.palette.length, 6, 'no theme colours: the Office accents');
});

test('series without data, with data of an unknown form or with broken numbers are read as far as they go', () => {
  const chart = must(
    chartXml({
      plot:
        group(
          'lineChart',
          '',
          ser(0, {}) +
            ser(1, { val: '<c:mystery/>' }) +
            ser(2, { val: numRef('S!$A$1:$A$4', ['1e2', '#N/A', 'abc', '0x10']) }) +
            ser(3, { val: `<c:numRef><c:f>S!$A$1</c:f><c:numCache><c:pt idx="-1"><c:v>1</c:v></c:pt><c:pt idx="x"><c:v>2</c:v></c:pt><c:pt><c:v>3</c:v></c:pt><c:pt idx="2"><c:v>9</c:v></c:pt></c:numCache></c:numRef>` }),
        ) +
        catAx() +
        valAx(),
    }),
  );
  const series = chart.groups[0].series;
  assert.equal(series.length, 4);
  assert.equal(series[0].values, undefined);
  assert.equal(series[1].values, undefined);
  assert.deepEqual(chartCacheList(series[2].values), [100, null, null, null], 'only numbers of the schema are numbers');
  assert.deepEqual(chartCacheList(series[3].values), [null, null, 9], 'points without a usable index are left out; the places before a point are empty');
});

test('colours: sRGB, theme colours with lumMod, lumOff, tint, shade and alpha, no fill, gradients and patterns', () => {
  const scheme = (name: string, inner: string) => spPr(`<a:solidFill><a:schemeClr val="${name}">${inner}</a:schemeClr></a:solidFill>`);
  const chart = must(
    chartXml({
      plot:
        group(
          'barChart',
          '<c:barDir val="col"/>',
          simple(0, spPr(solid('ff0000') + `<a:ln w="19050">${solid('00FF00')}<a:prstDash val="dash"/></a:ln>`)) +
            simple(1, scheme('accent1', '')) +
            simple(2, scheme('accent1', '<a:lumMod val="50000"/>')) +
            simple(3, scheme('tx1', '<a:lumMod val="65000"/><a:lumOff val="35000"/>')) +
            simple(4, scheme('accent2', '<a:alpha val="50000"/>')) +
            simple(5, spPr('<a:noFill/><a:ln><a:noFill/></a:ln>')) +
            simple(6, spPr(`<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="112233"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst></a:gradFill>`)) +
            simple(7, spPr('<a:pattFill prst="pct5"><a:fgClr><a:srgbClr val="445566"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>')) +
            simple(8, scheme('accent6', '<a:shade val="50000"/>')) +
            simple(9, scheme('bg1', '<a:alpha val="0"/>')) +
            simple(10, spPr('<a:solidFill><a:srgbClr val="nothex"/></a:solidFill>')) +
            simple(11),
        ) +
        catAx() +
        valAx(),
    }),
  );
  const shapes = chart.groups[0].series.map((series) => series.shape);
  assert.deepEqual(shapes[0], { fill: '#FF0000', line: { color: '#00FF00', width: 2, dash: true } });
  assert.equal(shapes[1]?.fill, '#4472C4', 'accent1 of the theme');
  assert.equal(shapes[2]?.fill, '#223962', 'lumMod halves the channels');
  assert.equal(shapes[3]?.fill, '#595959', 'the grey of the text of a new chart');
  assert.equal(shapes[4]?.fill, '#F6BE98', 'half transparent: the colour as it shows on white');
  assert.deepEqual(shapes[5], { fill: null, line: { color: null } }, 'no fill and no line are stated as such');
  assert.equal(shapes[6]?.fill, '#112233', 'a gradient: its first colour');
  assert.equal(shapes[7]?.fill, '#445566', 'a pattern: its foreground colour');
  assert.equal(shapes[8]?.fill, '#385724', 'shade darkens');
  assert.equal(shapes[9]?.fill, null, 'fully transparent is no fill');
  assert.equal(shapes[10], undefined, 'a colour that is none: the chart gives its own');
  assert.equal(shapes[11], undefined, 'nothing stated: the chart gives its own');
  for (const wanted of [/gradient/, /pattern/, /Transparent/]) assert.ok(chart.notes.some((note) => wanted.test(note)), `a note about ${wanted}`);
});

test('the palette is the six accent colours of the theme, whatever the theme is', () => {
  const theme = OFFICE_THEME.slice();
  theme[4] = 'aa0000';
  theme[9] = 'not a colour';
  const chart = must(chartXml({ plot: group('pieChart', '', simple()) }), theme);
  assert.deepEqual(chart.palette, ['#AA0000', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']);
  assert.deepEqual(must(chartXml({ plot: group('pieChart', '', simple()) }), []).palette[0], '#4472C4', 'no theme: the Office colours');
});

test('markers, points with a look of their own, invert if negative, smooth, data labels', () => {
  const chart = must(
    chartXml({
      plot:
        group(
          'lineChart',
          '<c:grouping val="standard"/>',
          simple(0, `<c:marker><c:symbol val="diamond"/><c:size val="9"/>${spPr(solid('010203'))}</c:marker><c:dPt><c:idx val="1"/>${spPr(solid('AABBCC'))}</c:dPt><c:dPt><c:idx val="-3"/></c:dPt><c:dLbls><c:showVal val="1"/></c:dLbls>`) +
            ser(1, { val: numRef('S!$B$2:$B$4', [1, 2, 3]), props: '<c:marker><c:symbol val="none"/></c:marker><c:invertIfNegative val="1"/><c:dLbls><c:delete val="1"/></c:dLbls>', tail: '<c:smooth val="1"/>' }) +
            ser(2, { val: numRef('S!$B$2:$B$4', [1, 2, 3]), props: '<c:marker><c:symbol val="heart"/><c:size val="900"/></c:marker>', tail: '<c:smooth val="0"/>' }),
          '<c:dLbls><c:showPercent val="1"/></c:dLbls><c:marker val="1"/>',
        ) +
        catAx() +
        valAx(),
    }),
  );
  const line = chart.groups[0];
  assert.equal(line.markers, true);
  assert.deepEqual(line.dataLabels, { showPercent: true });
  const [a, b, c] = line.series;
  assert.deepEqual(a.marker, { symbol: 'diamond', size: 9, fill: '#010203' });
  assert.deepEqual(a.points, [{ index: 1, shape: { fill: '#AABBCC' } }]);
  assert.deepEqual(a.dataLabels, { showValue: true });
  assert.deepEqual(b.marker, { symbol: 'none' });
  assert.equal(b.invertIfNegative, true);
  assert.equal(b.smooth, true);
  assert.deepEqual(b.dataLabels, { deleted: true }, 'deleted labels say so');
  assert.deepEqual(c.marker, { symbol: 'auto', size: 72 }, 'an unknown symbol is the chart\'s own; a size is kept within bounds');
  assert.equal(c.smooth, undefined);
  assert.ok(!chart.notes.some((note) => /Data labels|Smoothed/.test(note)), 'labels and smoothed lines are drawn: nothing to note');
});

test('pie and doughnut: first slice angle, hole size, colours that vary by default', () => {
  const pie = must(chartXml({ plot: group('pieChart', '', simple(), '<c:firstSliceAng val="90"/>', false) })).groups[0];
  assert.deepEqual([pie.kind, pie.varyColors, pie.firstSliceAngle, pie.holeSize], ['pie', true, 90, undefined]);
  const doughnut = must(chartXml({ plot: group('doughnutChart', '<c:varyColors val="0"/>', simple(), '<c:firstSliceAng val="400"/><c:holeSize val="75"/>', false) })).groups[0];
  assert.deepEqual([doughnut.kind, doughnut.varyColors, doughnut.firstSliceAngle, doughnut.holeSize], ['doughnut', false, 360, 75]);
  assert.equal(must(chartXml({ plot: group('doughnutChart', '', simple(), '', false) })).groups[0].holeSize, 50, 'a doughnut without a stated hole');
});

test('the title: plain text of the rich text with the look of its first run; the automatic title; a title from a cell', () => {
  const rich =
    '<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1800" b="0"/></a:pPr>' +
    `<a:r><a:rPr sz="2000" b="1" i="1">${solid('C00000')}<a:latin typeface="Arial"/></a:rPr><a:t>Yıllık </a:t></a:r><a:r><a:rPr sz="900"/><a:t>&lt;satış&gt;</a:t></a:r></a:p><a:p><a:r><a:t>ikinci satır</a:t></a:r></a:p></c:rich></c:tx></c:title>`;
  const chart = must(chartXml({ head: rich + '<c:autoTitleDeleted val="0"/>', plot: group('pieChart', '', simple()) }));
  assert.deepEqual(chart.title, { text: 'Yıllık <satış>\nikinci satır', style: { size: 20, bold: true, italic: true, color: '#C00000', font: 'Arial' } });
  assert.equal(chart.autoTitleDeleted, undefined);

  const none = must(chartXml({ head: '<c:autoTitleDeleted val="1"/>', plot: group('pieChart', '', simple()) }));
  assert.equal(none.title, undefined);
  assert.equal(none.autoTitleDeleted, true);

  const empty = must(chartXml({ head: '<c:title><c:overlay val="0"/></c:title>', plot: group('pieChart', '', simple()) }));
  assert.deepEqual(empty.title, {}, 'a title element without a text: the chart makes one up');

  const linked = must(chartXml({ head: '<c:title><c:tx><c:strRef><c:f>S!$A$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Hücreden</c:v></c:pt></c:strCache></c:strRef></c:tx></c:title>', plot: group('pieChart', '', simple()) }));
  assert.deepEqual(linked.title, { ref: 'S!$A$1', text: 'Hücreden' });
});

test('axes: kind, side, scaling, unit, number format, grid lines, labels, title, crossing', () => {
  const chart = must(
    chartXml({
      plot:
        group('barChart', '<c:barDir val="col"/>', simple()) +
        catAx('<c:majorGridlines/><c:tickLblPos val="low"/><c:crosses val="max"/>', 'dateAx') +
        valAx(
          `<c:majorGridlines>${spPr(`<a:ln w="9525">${solid('D9D9D9')}</a:ln>`)}</c:majorGridlines><c:minorGridlines/>` +
            `${title('Tutar', 'sz="1000" b="1"')}<c:numFmt formatCode="#,##0.00" sourceLinked="0"/><c:tickLblPos val="none"/>${spPr('<a:ln><a:noFill/></a:ln>')}` +
            '<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="800" b="1"/></a:pPr></a:p></c:txPr><c:crossesAt val="5"/><c:crossBetween val="midCat"/><c:majorUnit val="25"/>',
          '<c:logBase val="10"/><c:orientation val="maxMin"/><c:max val="100"/><c:min val="-50"/>',
        ),
    }),
  );
  const [cat, value] = chart.axes;
  assert.deepEqual(cat, { id: 111, kind: 'cat', position: 'b', tickLabels: 'low', date: true, majorGridlines: {}, crosses: 'max', crossAxis: 222 });
  assert.deepEqual(value, {
    id: 222,
    kind: 'val',
    position: 'l',
    tickLabels: 'none',
    min: -50,
    max: 100,
    reversed: true,
    logBase: 10,
    majorUnit: 25,
    formatCode: '#,##0.00',
    sourceLinked: false,
    majorGridlines: { color: '#D9D9D9', width: 1 },
    minorGridlines: {},
    title: { text: 'Tutar', style: { size: 10, bold: true } },
    crossesAt: 5,
    crossAxis: 111,
    crossBetween: 'midCat',
    line: { color: null },
    labelStyle: { size: 8, bold: true },
  });
  assert.ok(chart.notes.some((note) => /logarithmic axis/.test(note)), 'a logarithmic axis is noted');

  const deleted = must(chartXml({ plot: group('barChart', '', simple()) + '<c:catAx><c:axId val="111"/><c:delete val="1"/><c:axPos val="t"/><c:numFmt formatCode="General" sourceLinked="1"/></c:catAx><c:valAx><c:delete val="0"/></c:valAx>' }));
  assert.equal(deleted.axes.length, 1, 'an axis without an id is left out');
  assert.deepEqual([deleted.axes[0].deleted, deleted.axes[0].position, deleted.axes[0].sourceLinked], [true, 't', true]);
});

test('the legend: position, entries left out; plotVisOnly and what stands for empty cells; rounded corners; fills of the areas', () => {
  const chart = must(
    chartXml({
      before: '<c:date1904 val="1"/><c:roundedCorners val="1"/><c:style val="18"/>',
      plot: group('barChart', '<c:barDir val="col"/>', simple()) + catAx() + valAx() + spPr(solid('EEEEEE')),
      tail: legend('tr', '<c:legendEntry><c:idx val="1"/><c:delete val="1"/></c:legendEntry><c:legendEntry><c:idx val="2"/><c:delete val="0"/></c:legendEntry>') + '<c:plotVisOnly val="0"/><c:dispBlanksAs val="span"/>',
      after: spPr(solid('FAFAFA') + `<a:ln w="9525">${solid('123456')}</a:ln>`) + '<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1200"/></a:pPr></a:p></c:txPr>',
    }),
  );
  assert.deepEqual(chart.legend, { position: 'tr', deleted: [1] });
  assert.equal(chart.plotVisOnly, false);
  assert.equal(chart.blanks, 'span');
  assert.equal(chart.roundedCorners, true);
  assert.equal(chart.date1904, true);
  assert.deepEqual(chart.area, { fill: '#FAFAFA', line: { color: '#123456', width: 1 } });
  assert.deepEqual(chart.plotArea, { fill: '#EEEEEE' });
  assert.deepEqual(chart.textStyle, { size: 12 });
  assert.ok(chart.notes.some((note) => /style number/.test(note)), 'the style number is noted as not applied');

  for (const position of ['r', 'l', 't', 'b', 'tr']) assert.equal(must(chartXml({ plot: group('pieChart', '', simple()), tail: legend(position) })).legend?.position, position);
  assert.equal(must(chartXml({ plot: group('pieChart', '', simple()), tail: legend('nowhere') })).legend?.position, 'r', 'an unknown position is the right side');

  const defaults = must(chartXml({ plot: group('pieChart', '', simple()), tail: '' }));
  assert.equal(defaults.plotVisOnly, true, 'hidden cells are left out unless the file says otherwise');
  // UNSURE: what a chart without the element does with empty cells.
  assert.equal(defaults.blanks, 'gap');
  assert.equal(must(chartXml({ plot: group('pieChart', '', simple()), tail: '<c:dispBlanksAs/>' })).blanks, 'zero', 'the element without a value: the default of the schema');
});

test('several chart types in one plot area are all read, with the axes of a second scale and a place set by hand', () => {
  const second = '<c:lineChart><c:grouping val="standard"/>' + simple(1) + '<c:axId val="333"/><c:axId val="444"/></c:lineChart>';
  const chart = must(
    chartXml({
      plot:
        group('barChart', '<c:barDir val="col"/>', simple(0)) +
        second +
        catAx() +
        valAx() +
        '<c:catAx><c:axId val="333"/><c:delete val="1"/><c:axPos val="b"/><c:crossAx val="444"/></c:catAx><c:valAx><c:axId val="444"/><c:axPos val="r"/><c:crosses val="max"/><c:crossAx val="333"/></c:valAx>',
      tail: legend('r', '<c:layout><c:manualLayout><c:x val="0.1"/></c:manualLayout></c:layout>'),
    }),
  );
  assert.deepEqual(chart.groups.map((item) => item.kind), ['bar', 'line']);
  assert.equal(chart.axes.length, 4);
  assert.deepEqual(chart.groups[1].axisIds, [333, 444], 'the second chart type names its own axes');
  assert.deepEqual(chart.legend?.layout, { x: 0.1 }, 'the place of the legend as the file states it');
  assert.ok(!chart.notes.some((note) => /secondary axis|set by hand/.test(note)), 'neither is noted as not applied');
});

test('series are put in their order; an index that is missing is the place in the file', () => {
  const chart = must(chartXml({ plot: group('barChart', '', ser(5, { order: 2, val: numLit([1]) }) + ser(3, { order: 0, val: numLit([2]) }) + '<c:ser><c:val>' + numLit([3]) + '</c:val></c:ser>') + catAx() + valAx() }));
  assert.deepEqual(chart.groups[0].series.map((series) => [series.index, series.order]), [[3, 0], [5, 2], [2, 2]]);
});

test('limits: series beyond the limit are left out, points beyond the limit too, and a sparse cache cannot take the memory', () => {
  assert.equal(SERIES, 255);
  assert.equal(POINTS, 100000);
  let many = '';
  for (let i = 0; i < SERIES + 20; i++) many += ser(i, { val: numLit([i]) });
  const wide = must(chartXml({ plot: group('barChart', '', many) + catAx() + valAx() }));
  assert.equal(wide.groups[0].series.length, SERIES);
  assert.ok(wide.notes.some((note) => /first 255 series/.test(note)), 'the cut is noted');

  // A point far beyond the limit, and a count that claims a billion points.
  const far = must(chartXml({ plot: group('lineChart', '', ser(0, { val: `<c:numLit><c:ptCount val="1000000000"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="${POINTS}"><c:v>2</c:v></c:pt><c:pt idx="2147483647"><c:v>3</c:v></c:pt></c:numLit>` })) + catAx() + valAx() }));
  const values = far.groups[0].series[0].values;
  assert.deepEqual(values?.cache, { index: [0], values: [1] }, 'only what the file has a value for takes room');
  assert.equal(values?.count, POINTS);
  assert.ok(far.notes.some((note) => /first 100000 points/.test(note)), 'the cut is noted');

  // Every series claims its last point: as places they would be 25 million; as what the file holds they are 255.
  let sparse = '';
  for (let i = 0; i < SERIES; i++) sparse += ser(i, { val: `<c:numLit><c:pt idx="${POINTS - 1}"><c:v>${i}</c:v></c:pt></c:numLit>` });
  const heavy = must(chartXml({ plot: group('lineChart', '', sparse) + catAx() + valAx() }));
  assert.equal(chartCachedPoints(heavy), SERIES, 'one entry per point of the file');
  heavy.groups[0].series.forEach((series, i) => {
    assert.deepEqual(series.values?.cache, { index: [POINTS - 1], values: [i] }, `series ${i}: its one point`);
    assert.equal(series.values?.count, POINTS, `series ${i}: as many points as its last index says`);
  });
  assert.ok(MAX_CHART_CACHE >= POINTS, 'a full series fits the limit of a chart');
});

test('texts of a chart are cut to a length that can be drawn', () => {
  const long = 'x'.repeat(20000);
  const chart = must(chartXml({ head: title(long), plot: group('pieChart', '', ser(0, { tx: `<c:v>${long}</c:v>`, cat: strLit([long]), val: numLit([1]) })) }));
  assert.equal(chart.title?.text?.length, 4096);
  assert.equal(chart.groups[0].series[0].name?.text?.length, 4096);
  assert.equal((chartCached(chart.groups[0].series[0].categories, 0) as string).length, 4096);
});

// ----------------------------------------------------------------- in a workbook

const plain = chartXml({ plot: group('barChart', '<c:barDir val="col"/>', ser(0, { cat: strRef('Sayfa1!$A$1:$A$3', ['a', 'b', 'c']), val: numRef('Sayfa1!$B$1:$B$3', [1, 2, 3]) })) + catAx() + valAx() });
const cells = sheetData([['a', 1], ['b', 2], ['c', 3]]);

async function drawingsOf(bytes: Uint8Array) {
  const workbook = await (await open(bytes)).workbook();
  return workbook.sheets.map((sheet) => sheet.drawings || []);
}

test('the frame of a chart in a drawing is a chart of the model, with the anchor of any other object', async () => {
  const [first] = await drawingsOf(await chartWorkbook({ sheet: cells, drawing: twoCell([3, 1], [9, 15], chartFrameOf('rIdC1', 'Satış grafiği')), charts: [plain] }));
  assert.equal(first.length, 1);
  const object = first[0].object as DrawingChart;
  assert.equal(object.type, 'chart');
  assert.equal(object.name, 'Satış grafiği');
  assert.equal(object.part, '/xl/charts/chart1.xml');
  assert.equal(object.chart.groups[0].kind, 'bar');
  assert.equal(first[0].anchor.kind, 'twoCell');
  assert.deepEqual([first[0].anchor.from?.col, first[0].anchor.from?.row, first[0].anchor.to?.col, first[0].anchor.to?.row], [3, 1, 9, 15]);
});

test('a chart that cannot be read is a placeholder; the other objects and the sheet are still there', async () => {
  const radar = chartXml({ plot: group('surfaceChart', '', simple()) + catAx() + valAx() });
  const dtd = `<?xml version="1.0"?><!DOCTYPE c:chartSpace [<!ENTITY a "aaaa">]>${plain.slice(plain.indexOf('<c:chartSpace'))}`;
  const cut = plain.slice(0, plain.length - 40);
  const junk = new Uint8Array([0xff, 0xfe, 0x00, 0x01, 0x02]);
  const frames = [1, 2, 3, 4, 5, 6].map((n, i) => twoCell([3, i * 5], [6, i * 5 + 4], chartFrameOf(`rIdC${n}`))).join('') + twoCell([7, 0], [9, 4], chartFrameOf('rIdMissing')) + twoCell([7, 5], [9, 9], chartFrame());
  const [first] = await drawingsOf(await chartWorkbook({ sheet: cells, drawing: frames, charts: [plain, radar, dtd, cut, junk, plain] }));
  assert.deepEqual(
    first.map((drawing) => drawing.object.type),
    ['chart', 'placeholder', 'placeholder', 'placeholder', 'placeholder', 'chart', 'placeholder', 'placeholder'],
    'a chart, a radar chart, a part with a DTD, a part that is cut off, bytes that are no XML, a chart, a frame without its part, a frame without a reference',
  );
  for (const drawing of first) if (drawing.object.type === 'placeholder') assert.equal(drawing.object.what, 'chart');
});

test('a frame whose relationship is not that of a chart (an extended chart) stays a placeholder', async () => {
  const [first] = await drawingsOf(await chartWorkbook({ sheet: cells, drawing: twoCell([3, 1], [9, 15], chartFrameOf('rIdC1')), charts: [plain], relType: 'http://schemas.microsoft.com/office/2014/relationships/chartEx' }));
  assert.equal(first[0].object.type, 'placeholder');
});

test('a chart sheet has the drawing that holds its chart; one whose chart is not drawn has a placeholder', async () => {
  const sheets = await drawingsOf(await chartWorkbook({ sheet: cells, chartSheet: plain }));
  assert.equal(sheets.length, 2);
  assert.equal(sheets[1].length, 1);
  assert.equal(sheets[1][0].object.type, 'chart');
  assert.equal(sheets[1][0].anchor.kind, 'absolute');

  const radar = chartXml({ plot: group('surfaceChart', '', simple()) + catAx() + valAx() });
  const other = await drawingsOf(await chartWorkbook({ sheet: cells, chartSheet: radar }));
  assert.equal(other[1][0].object.type, 'placeholder');

  const workbook = await (await open(await chartWorkbook({ sheet: cells, chartSheet: new Uint8Array([1, 2, 3]) }))).workbook();
  assert.equal(workbook.sheets[1].kind, 'chartsheet', 'the sheet is listed whatever its chart part holds');
  assert.equal(workbook.sheets[1].drawings?.[0].object.type, 'placeholder');
});

test('the number of chart parts read for one sheet is limited', async () => {
  assert.equal(MAX_DRAWING_CHARTS, 1000);
  const count = MAX_DRAWING_CHARTS + 3;
  const tiny = chartXml({ plot: group('pieChart', '', ser(0, { val: numLit([1]) })) });
  let frames = '';
  const charts: string[] = [];
  for (let i = 0; i < count; i++) {
    frames += twoCell([0, i], [1, i + 1], chartFrameOf(`rIdC${i + 1}`));
    charts.push(tiny);
  }
  const [first] = await drawingsOf(await chartWorkbook({ drawing: frames, charts }));
  assert.equal(first.length, count);
  assert.equal(first.filter((drawing) => drawing.object.type === 'chart').length, MAX_DRAWING_CHARTS);
  assert.equal(first.filter((drawing) => drawing.object.type === 'placeholder').length, 3);
});
