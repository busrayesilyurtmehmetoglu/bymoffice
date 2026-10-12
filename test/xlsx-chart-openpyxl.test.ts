// Cross-check of the chart reader against openpyxl, an independent
// implementation: openpyxl writes workbooks with radar, bubble and stock
// charts, a second scale, data labels, a trend line, error bars, a logarithmic
// axis, a date axis, turned labels, places set by hand and a pie of two plots,
// and says what it wrote; what the reader makes of the files has to agree.
// Skipped where python3 or openpyxl is not installed.
//
// What the cross-check found is pinned down without openpyxl too, with chart
// parts written the way openpyxl writes them (the second test).

import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdirSync, readFileSync, rmSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { open, type Chart, type DrawingChart, type Workbook } from '../src/index';
import { chartValues, layoutChart, type ChartItem, type ChartLayout } from '../src/layout/chart';
import { chartXml, fakeMeasure, readChart } from './chart-fixtures';

// Writes one workbook with a chart per sheet and prints what it set, as JSON.
const PYTHON = String.raw`
import json, sys, datetime
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, RadarChart, BubbleChart, StockChart, ScatterChart, PieChart, ProjectedPieChart, Reference, Series
from openpyxl.chart.axis import DateAxis
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.trendline import Trendline
from openpyxl.chart.error_bar import ErrorBars
from openpyxl.chart.data_source import NumDataSource, NumRef, AxDataSource, MultiLevelStrRef
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.text import RichText
from openpyxl.chart.updown_bars import UpDownBars
from openpyxl.chart.axis import ChartLines
from openpyxl.drawing.text import RichTextProperties, Paragraph, ParagraphProperties, CharacterProperties

wb = Workbook()
ws = wb.active
ws.title = "Veri"
rows = [
    ["Ay", "Açılış", "Yüksek", "Düşük", "Kapanış", "Adet", "Tarih", "Yıl", "Ay"],
    ["Oca", 10, 35, 5, 30, 1200, datetime.date(2024, 1, 15), "2024", "Oca"],
    ["Şub", 30, 32, 12, 15, 900, datetime.date(2024, 2, 20), None, "Şub"],
    ["Mar", 20, 28, 15, 20, 1500, datetime.date(2024, 6, 5), "2025", "Mar"],
    ["Nis", 22, 40, 18, 38, 100, datetime.date(2024, 7, 1), None, "Nis"],
]
for row in rows:
    ws.append(row)
cats = Reference(ws, min_col=1, min_row=2, max_row=5)
spec = {}

def place(chart, name):
    sheet = wb.create_sheet(name)
    sheet.add_chart(chart, "A1")

# 1. A filled radar of two series.
radar = RadarChart()
radar.type = "filled"
radar.add_data(Reference(ws, min_col=2, max_col=3, min_row=1, max_row=5), titles_from_data=True)
radar.set_categories(cats)
place(radar, "Radar")
spec["Radar"] = {"style": "filled", "series": [[10, 30, 20, 22], [35, 32, 28, 40]], "names": ["Açılış", "Yüksek"], "cats": ["Oca", "Şub", "Mar", "Nis"]}

# 2. Bubbles: x, y and sizes.
bubble = BubbleChart()
bubble.bubbleScale = 50
bubble.showNegBubbles = True
bubble.sizeRepresents = "w"
series = Series(values=Reference(ws, min_col=3, min_row=2, max_row=5), xvalues=Reference(ws, min_col=2, min_row=2, max_row=5), zvalues=Reference(ws, min_col=6, min_row=2, max_row=5), title="Balon")
bubble.series.append(series)
place(bubble, "Balon")
spec["Balon"] = {"x": [10, 30, 20, 22], "y": [35, 32, 28, 40], "size": [1200, 900, 1500, 100], "scale": 50}

# 3. Open, high, low, close with high-low lines and up-down bars.
stock = StockChart()
stock.add_data(Reference(ws, min_col=2, max_col=5, min_row=1, max_row=5), titles_from_data=True)
stock.set_categories(cats)
for s in stock.series:
    s.graphicalProperties.line.noFill = True
    s.marker.symbol = "none"
stock.hiLowLines = ChartLines()
stock.upDownBars = UpDownBars(gapWidth=120)
place(stock, "Borsa")
spec["Borsa"] = {"series": 4, "gap": 120, "close": [30, 15, 20, 38]}

# 4. Columns with a line on a second scale at the right.
bar = BarChart()
bar.add_data(Reference(ws, min_col=5, min_row=1, max_row=5), titles_from_data=True)
bar.set_categories(cats)
line = LineChart()
line.add_data(Reference(ws, min_col=6, min_row=1, max_row=5), titles_from_data=True)
line.y_axis.axId = 200
line.y_axis.title = "Adet"
line.y_axis.crosses = "max"
bar.y_axis.delete = False
bar.x_axis.delete = False
line.y_axis.delete = False
bar += line
place(bar, "İki")
spec["İki"] = {"first": [30, 15, 20, 38], "second": [1200, 900, 1500, 100], "axis": 200}

# 5. Data labels: values and categories outside the end, in a format of their own.
labelled = BarChart()
labelled.add_data(Reference(ws, min_col=2, min_row=1, max_row=5), titles_from_data=True)
labelled.set_categories(cats)
labels = DataLabelList()
labels.showVal = True
labels.showCatName = True
labels.showSerName = False
labels.showLegendKey = False
labels.position = "outEnd"
labels.numFmt = "0.0"
labels.separator = " | "
labelled.series[0].dLbls = labels
labelled.x_axis.delete = False
labelled.y_axis.delete = False
labelled.x_axis.txPr = RichText(bodyPr=RichTextProperties(rot=-2700000), p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties()), endParaRPr=CharacterProperties())])
place(labelled, "Etiket")
spec["Etiket"] = {"texts": ["Oca | 10.0", "Şub | 30.0", "Mar | 20.0", "Nis | 22.0"], "rotation": -45}

# 6. A scatter series with a trend line and error bars along both axes.
scatter = ScatterChart()
scatter.style = 13
points = Series(Reference(ws, min_col=3, min_row=2, max_row=5), Reference(ws, min_col=2, min_row=2, max_row=5), title="Ölçüm")
points.trendline = Trendline(trendlineType="poly", order=3, forward=2, backward=1, dispRSqr=True, dispEq=False)
points.errBars = ErrorBars(errDir="y", errBarType="plus", errValType="fixedVal", noEndCap=True, val=2)
scatter.series.append(points)
scatter.x_axis.delete = False
scatter.y_axis.delete = False
place(scatter, "Eğilim")
spec["Eğilim"] = {"trend": {"type": "poly", "order": 3, "forward": 2, "backward": 1}, "bars": {"direction": "y", "type": "plus", "valueType": "fixedVal", "value": 2, "noEndCap": True}}

# 7. A logarithmic value axis.
log = LineChart()
log.add_data(Reference(ws, min_col=6, min_row=1, max_row=5), titles_from_data=True)
log.set_categories(cats)
log.y_axis.scaling.logBase = 10
log.y_axis.scaling.min = 10
log.y_axis.scaling.max = 10000
log.x_axis.delete = False
log.y_axis.delete = False
place(log, "Log")
spec["Log"] = {"base": 10, "min": 10, "max": 10000, "values": [1200, 900, 1500, 100]}

# 8. A date axis that counts in months.
dated = LineChart()
dated.add_data(Reference(ws, min_col=5, min_row=1, max_row=5), titles_from_data=True)
dated.set_categories(Reference(ws, min_col=7, min_row=2, max_row=5))
dated.x_axis = DateAxis(crossAx=100)
dated.x_axis.number_format = "mmm"
dated.x_axis.majorTimeUnit = "months"
dated.x_axis.baseTimeUnit = "months"
dated.x_axis.majorUnit = 1
dated.x_axis.delete = False
dated.y_axis.delete = False
dated.y_axis.crossAx = 500
place(dated, "Tarih")
spec["Tarih"] = {"unit": "months", "labels": ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul"], "months": [0, 1, 5, 6]}

# 9. A pie with its slices pulled out, in a plot area set by hand.
pie = PieChart()
pie.add_data(Reference(ws, min_col=2, min_row=1, max_row=5), titles_from_data=True)
pie.set_categories(cats)
pie.series[0].explosion = 20
pie.layout = Layout(manualLayout=ManualLayout(x=0.25, y=0.25, w=0.5, h=0.5, xMode="edge", yMode="edge", layoutTarget="inner"))
place(pie, "Pasta")
spec["Pasta"] = {"explosion": 20, "layout": {"target": "inner", "xMode": "edge", "yMode": "edge", "x": 0.25, "y": 0.25, "w": 0.5, "h": 0.5}}

# 10. A pie whose last two points are a second plot.
of = ProjectedPieChart()
of.type = "bar"
of.splitType = "pos"
of.splitPos = 2
of.add_data(Reference(ws, min_col=2, min_row=1, max_row=5), titles_from_data=True)
of.set_categories(cats)
place(of, "İkili")
spec["İkili"] = {"split": "pos", "position": 2, "shares": [10 / 82, 30 / 82, 42 / 82]}

# 11. Two levels of categories, read from the cells.
levels = BarChart()
levels.add_data(Reference(ws, min_col=2, min_row=1, max_row=5), titles_from_data=True)
levels.series[0].cat = AxDataSource(multiLvlStrRef=MultiLevelStrRef(f="'Veri'!$H$2:$I$5"))
levels.x_axis.delete = False
levels.y_axis.delete = False
place(levels, "Düzey")
spec["Düzey"] = {"inner": ["Oca", "Şub", "Mar", "Nis"], "outer": ["2024", "", "2025", ""]}

wb.save(sys.argv[1])
print(json.dumps(spec))
`;

type Of<K extends ChartItem['kind']> = Extract<ChartItem, { kind: K }>;
const items = <K extends ChartItem['kind']>(layout: ChartLayout, kind: K, role?: string): Of<K>[] => layout.items.filter((item) => item.kind === kind && (role === undefined || item.role === role)) as Of<K>[];
const near = (actual: number, expected: number, message: string, tolerance = 1e-6) => assert.ok(Math.abs(actual - expected) <= tolerance, `${message}: ${actual} is not ${expected}`);

test('charts: what openpyxl writes of the new chart types and options is read as openpyxl wrote it', async (t) => {
  const probe = spawnSync('python3', ['-c', 'import openpyxl; print(openpyxl.__version__)'], { encoding: 'utf8' });
  if (probe.status !== 0) {
    t.skip('python3 with openpyxl is not available');
    return;
  }
  // A scratch folder inside the repository (ignored by git), removed afterwards.
  const folder = join(dirname(fileURLToPath(import.meta.url)), '..', `tmp-chart-openpyxl-${process.pid}`);
  mkdirSync(folder, { recursive: true });
  let bytes: Uint8Array;
  let spec: Record<string, any>;
  try {
    const file = join(folder, 'charts.xlsx');
    const run = spawnSync('python3', ['-c', PYTHON, file], { encoding: 'utf8' });
    assert.equal(run.status, 0, run.stderr);
    spec = JSON.parse(run.stdout);
    bytes = new Uint8Array(readFileSync(file));
  } finally {
    rmSync(folder, { recursive: true, force: true });
  }
  const workbook = (await (await open(bytes)).workbook()) as Workbook;
  const chartOf = (name: string): Chart => {
    const sheet = workbook.sheets.filter((candidate) => candidate.name === name)[0];
    assert.ok(sheet && sheet.drawings && sheet.drawings.length === 1, `the sheet "${name}" has one drawing`);
    const object = (sheet.drawings as NonNullable<typeof sheet.drawings>)[0].object as DrawingChart;
    assert.equal(object.type, 'chart', `"${name}" holds a chart`);
    assert.ok(object.chart, `the chart of "${name}" is read, not a placeholder`);
    return object.chart as Chart;
  };
  // openpyxl writes no cached values: everything is read from the cells.
  const numbers = (chart: Chart) => chartValues(chart, { workbook });
  const lay = (chart: Chart, width = 640, height = 400) => layoutChart(chart, numbers(chart), width, height, fakeMeasure, { date1904: workbook.date1904 === true });

  // 1. Radar.
  const radar = chartOf('Radar');
  assert.deepEqual([radar.groups[0].kind, radar.groups[0].radarStyle], ['radar', spec.Radar.style]);
  const radarValues = numbers(radar);
  assert.deepEqual(radarValues.series.map((entry) => entry.values.toArray()), spec.Radar.series);
  assert.deepEqual(radarValues.series.map((entry) => entry.name), spec.Radar.names);
  assert.deepEqual([0, 1, 2, 3].map((i) => radarValues.label(i)), spec.Radar.cats);
  assert.equal(items(lay(radar), 'path', 'area').length, 2, 'two filled outlines');

  // 2. Bubbles.
  const bubble = chartOf('Balon');
  const group = bubble.groups[0];
  assert.deepEqual([group.kind, group.bubbleScale, group.showNegBubbles, group.sizeRepresents], ['bubble', spec.Balon.scale, true, 'w']);
  const entry = numbers(bubble).series[0];
  assert.deepEqual([entry.xs?.toArray(), entry.values.toArray(), entry.sizes?.toArray()], [spec.Balon.x, spec.Balon.y, spec.Balon.size]);
  const circles = items(lay(bubble), 'circles', 'bubble')[0].points;
  near(circles[2] / circles[8], 1200 / 1500, 'sizes as widths: the radii are as the sizes');

  // 3. Stock.
  const stock = chartOf('Borsa');
  assert.deepEqual([stock.groups[0].kind, stock.groups[0].series.length], ['stock', spec.Borsa.series]);
  assert.ok(stock.groups[0].hiLowLines, 'the high-low lines');
  assert.equal(stock.groups[0].upDownBars?.gapWidth, spec.Borsa.gap);
  const stockLayout = lay(stock);
  assert.equal(items(stockLayout, 'path', 'hilow')[0].points.length, 16, 'a line for each of the four days');
  assert.deepEqual(items(stockLayout, 'rect', 'updown').map((box) => box.fill), ['#FFFFFF', '#000000', '#FFFFFF', '#FFFFFF'], 'up, down, level, up');
  assert.equal(items(stockLayout, 'path', 'line').length, 0, 'the series have no lines: openpyxl was told to take them away');
  assert.deepEqual(numbers(stock).series[3].values.toArray(), spec.Borsa.close);

  // 4. A second scale.
  const two = chartOf('İki');
  assert.deepEqual(two.groups.map((item) => item.kind), ['bar', 'line']);
  assert.ok(two.groups[1].axisIds.indexOf(spec.İki.axis) >= 0, 'the line names the axis openpyxl gave it');
  const secondAxis = two.axes.filter((axis) => axis.id === spec.İki.axis)[0];
  assert.deepEqual([secondAxis.kind, secondAxis.crosses, secondAxis.title?.text], ['val', 'max', 'Adet']);
  const twoLayout = lay(two);
  assert.ok(twoLayout.scale && twoLayout.scale2, 'two scales');
  assert.ok((twoLayout.scale?.max as number) < 100 && (twoLayout.scale2?.max as number) >= 1500, `each for its own numbers (${twoLayout.scale?.max}, ${twoLayout.scale2?.max})`);
  const right = twoLayout.plot.x + twoLayout.plot.width;
  assert.ok(items(twoLayout, 'text', 'tick').some((item) => item.x > right), 'labels of the second scale at the right');
  assert.ok(items(twoLayout, 'text', 'axisTitle').some((item) => item.text === 'Adet' && item.x > right), 'with its title');
  assert.deepEqual(numbers(two).series.map((item) => item.values.toArray()), [spec.İki.first, spec.İki.second]);

  // 5. Data labels and turned category labels.
  const labelled = chartOf('Etiket');
  // openpyxl states the separator as an attribute and leaves out whether the format is that of the cells.
  assert.deepEqual(labelled.groups[0].series[0].dataLabels, { showValue: true, showCategory: true, showSeries: false, showLegendKey: false, separator: ' | ', formatCode: '0.0', position: 'outEnd' });
  assert.equal(labelled.axes.filter((axis) => axis.kind === 'cat')[0].labelRotation, spec.Etiket.rotation);
  const labelLayout = lay(labelled);
  assert.deepEqual(items(labelLayout, 'text', 'label').map((item) => item.text), spec.Etiket.texts);
  assert.ok(items(labelLayout, 'text', 'category').every((item) => item.angle === -45), 'the labels of the categories are turned as the file says');

  // 6. A trend line and error bars.
  const trend = chartOf('Eğilim');
  const measured = trend.groups[0].series[0];
  assert.deepEqual(measured.trendlines, [spec.Eğilim.trend]);
  assert.deepEqual(measured.errorBars, [spec.Eğilim.bars]);
  assert.ok(trend.notes.some((note) => /R²/.test(note)), 'R² is noted as not shown');
  const trendLayout = lay(trend);
  assert.ok(items(trendLayout, 'path', 'trend').length >= 1, 'the trend line is drawn');
  assert.equal(items(trendLayout, 'path', 'error')[0].points.length, 16, 'four bars without caps');

  // 7. A logarithmic axis.
  const log = chartOf('Log');
  const logAxis = log.axes.filter((axis) => axis.kind === 'val')[0];
  assert.deepEqual([logAxis.logBase, logAxis.min, logAxis.max], [spec.Log.base, spec.Log.min, spec.Log.max]);
  const logLayout = lay(log);
  assert.deepEqual(logLayout.scale?.ticks, [10, 100, 1000, 10000]);
  const logPoints = items(logLayout, 'path', 'line')[0].points;
  near(logPoints[7], logLayout.plot.y + logLayout.plot.height * (2 / 3), '100 stands a third of the way up from 10 to 10000');

  // 8. A date axis.
  const dated = chartOf('Tarih');
  const dateAxis = dated.axes.filter((axis) => axis.date)[0];
  assert.deepEqual([dateAxis.baseTimeUnit, dateAxis.majorTimeUnit, dateAxis.majorUnit, dateAxis.formatCode], [spec.Tarih.unit, 'months', 1, 'mmm']);
  const dateLayout = lay(dated);
  assert.deepEqual(items(dateLayout, 'text', 'category').map((item) => item.text), spec.Tarih.labels);
  const datePoints = items(dateLayout, 'path', 'line')[0].points;
  const month = dateLayout.plot.width / 7;
  spec.Tarih.months.forEach((at: number, i: number) => near(datePoints[i * 2], dateLayout.plot.x + (at + 0.5) * month, `point ${i} stands in its month`));

  // 9. A pie pulled apart in a place of its own.
  const pie = chartOf('Pasta');
  assert.equal(pie.groups[0].series[0].explosion, spec.Pasta.explosion);
  assert.deepEqual(pie.plotLayout, spec.Pasta.layout);
  const slices = items(lay(pie, 640, 400), 'slice');
  near((slices[0].cx + slices[2].cx) / 2, 320, 'the pie stands in the middle of its place', 40);
  assert.ok(slices.every((slice) => slice.radius <= 100 / 1.2 + 1e-6), `no larger than the place leaves a pie that is pulled apart (${slices[0].radius})`);

  // 10. A pie of two plots.
  const of = chartOf('İkili');
  assert.deepEqual(of.groups[0].ofPie, { split: spec.İkili.split, position: spec.İkili.position });
  const shares = items(lay(of), 'slice').map((slice) => (slice.end - slice.start) / (Math.PI * 2));
  shares.forEach((share, i) => near(share, spec.İkili.shares[i], `share ${i}`));

  // 11. Two levels of categories from the cells (openpyxl writes the reference alone).
  const levels = chartOf('Düzey');
  assert.equal(levels.groups[0].series[0].categories?.multiLevel, true);
  const levelValues = numbers(levels);
  assert.deepEqual([0, 1, 2, 3].map((i) => levelValues.label(i)), spec.Düzey.inner, 'the last column of the block is the level next to the axis');
  assert.deepEqual([0, 1, 2, 3].map((i) => (levelValues.outerLabel ? levelValues.outerLabel(i) : '?')), spec.Düzey.outer, 'the column before it the level outside');
  const levelTexts = items(lay(levels), 'text', 'category').map((item) => item.text);
  assert.ok(levelTexts.indexOf('2024') >= 0 && levelTexts.indexOf('2025') >= 0 && levelTexts.indexOf('Nis') >= 0, 'both levels are shown');
});

// ------------------------------------------- the same, without openpyxl at hand

/** A chart part the way openpyxl 3.1 writes one: no caches, every element in the order of the schema. */
const openpyxlChart = (plot: string) => chartXml({ plot, tail: '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>' });

test('charts: parts written the way openpyxl writes them are read without it', () => {
  // A stock chart: series without lines or markers, high-low lines and up-down bars as empty elements.
  const hidden = '<c:spPr><a:ln><a:noFill/><a:prstDash val="solid"/></a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker>';
  const series = [0, 1, 2, 3].map((i) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:strRef><c:f>'Veri'!${'BCDE'[i]}1</c:f></c:strRef></c:tx>${hidden}<c:cat><c:numRef><c:f>'Veri'!$A$2:$A$5</c:f></c:numRef></c:cat><c:val><c:numRef><c:f>'Veri'!$${'BCDE'[i]}$2:$${'BCDE'[i]}$5</c:f></c:numRef></c:val></c:ser>`).join('');
  const stock = readChart(openpyxlChart(`<c:stockChart>${series}<c:hiLowLines/><c:upDownBars><c:gapWidth val="120"/></c:upDownBars><c:axId val="10"/><c:axId val="100"/></c:stockChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:crossAx val="100"/><c:lblOffset val="100"/></c:catAx><c:valAx><c:axId val="100"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:majorGridlines/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:crossAx val="10"/></c:valAx>`)) as Chart;
  assert.ok(stock, 'the stock chart is read');
  assert.deepEqual([stock.groups[0].kind, stock.groups[0].series.length, stock.groups[0].upDownBars?.gapWidth], ['stock', 4, 120]);
  assert.deepEqual(stock.groups[0].hiLowLines, {});
  assert.deepEqual(stock.groups[0].series[0].shape, { line: { color: null } }, 'a line that is taken away');
  assert.equal(stock.groups[0].series[0].values?.ref, "'Veri'!$B$2:$B$5", 'a reference without a cache');
  assert.equal(stock.groups[0].series[0].values?.count, 0);
  // Both axes say "l", as openpyxl writes every axis: the categories still run along the bottom.
  const layout = layoutChart(stock, chartValues(stock), 480, 320, fakeMeasure);
  assert.ok(Number.isFinite(layout.plot.width) && layout.plot.width > 100, 'a chart without cached values lays out to an empty plot area');

  // A second scale: the second value axis has an id of its own and crosses at the maximum; openpyxl writes c:crosses before c:crossAx... after it.
  const two = readChart(
    openpyxlChart(
      '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>\'Veri\'!$E$2:$E$5</c:f></c:numRef></c:val></c:ser><c:gapWidth val="150"/><c:axId val="10"/><c:axId val="100"/></c:barChart>' +
        '<c:lineChart><c:grouping val="standard"/><c:ser><c:idx val="1"/><c:order val="1"/><c:val><c:numRef><c:f>\'Veri\'!$F$2:$F$5</c:f></c:numRef></c:val></c:ser><c:axId val="10"/><c:axId val="200"/></c:lineChart>' +
        '<c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="100"/></c:catAx>' +
        '<c:valAx><c:axId val="100"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="10"/></c:valAx>' +
        '<c:valAx><c:axId val="200"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Adet</a:t></a:r></a:p></c:rich></c:tx></c:title><c:crossAx val="10"/><c:crosses val="max"/></c:valAx>',
    ),
  ) as Chart;
  assert.deepEqual(two.groups.map((item) => item.axisIds), [[10, 100], [10, 200]], 'the two chart types share the category axis and differ in the value axis');
  // Numbers of its own, so that the scales can be told apart without cells.
  two.groups[0].series[0].values = { kind: 'num', cache: { index: [0, 1, 2], values: [30, 15, 20] }, count: 3 };
  two.groups[1].series[0].values = { kind: 'num', cache: { index: [0, 1, 2], values: [1200, 900, 1500] }, count: 3 };
  const twoLayout = layoutChart(two, chartValues(two), 640, 400, fakeMeasure);
  assert.ok((twoLayout.scale?.max as number) <= 40 && (twoLayout.scale2?.max as number) >= 1500, 'a scale for each value axis');
  const right = twoLayout.plot.x + twoLayout.plot.width;
  const title = items(twoLayout, 'text', 'axisTitle').filter((item) => item.text === 'Adet')[0];
  assert.ok(title && title.x > right, 'the axis that crosses at the maximum stands at the right, though it says "l"');

  // A bubble series: c:bubbleSize after c:yVal, c:bubble3D after it; the settings of the chart type after the series.
  const bubble = readChart(
    openpyxlChart(
      "<c:bubbleChart><c:varyColors val=\"0\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/><c:xVal><c:numRef><c:f>'Veri'!$B$2:$B$5</c:f></c:numRef></c:xVal><c:yVal><c:numRef><c:f>'Veri'!$C$2:$C$5</c:f></c:numRef></c:yVal><c:bubbleSize><c:numRef><c:f>'Veri'!$F$2:$F$5</c:f></c:numRef></c:bubbleSize><c:bubble3D val=\"0\"/></c:ser><c:bubbleScale val=\"50\"/><c:showNegBubbles val=\"1\"/><c:sizeRepresents val=\"w\"/><c:axId val=\"10\"/><c:axId val=\"100\"/></c:bubbleChart>" +
        '<c:valAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="100"/></c:valAx><c:valAx><c:axId val="100"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="10"/></c:valAx>',
    ),
  ) as Chart;
  assert.deepEqual([bubble.groups[0].kind, bubble.groups[0].bubbleScale, bubble.groups[0].showNegBubbles, bubble.groups[0].sizeRepresents], ['bubble', 50, true, 'w']);
  assert.equal(bubble.groups[0].series[0].bubbleSizes?.ref, "'Veri'!$F$2:$F$5");
  assert.ok(!bubble.notes.some((note) => /3-D/.test(note)), 'bubble3D="0" is no 3-D chart');

  // A trend line and error bars: the type before the settings of it, the amount as c:val.
  const fitted = readChart(
    openpyxlChart(
      "<c:scatterChart><c:scatterStyle val=\"lineMarker\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/><c:trendline><c:trendlineType val=\"poly\"/><c:order val=\"3\"/><c:forward val=\"2\"/><c:backward val=\"1\"/><c:dispRSqr val=\"1\"/><c:dispEq val=\"0\"/></c:trendline><c:errBars><c:errDir val=\"y\"/><c:errBarType val=\"plus\"/><c:errValType val=\"fixedVal\"/><c:noEndCap val=\"1\"/><c:val val=\"2\"/></c:errBars><c:xVal><c:numRef><c:f>'Veri'!$B$2:$B$5</c:f></c:numRef></c:xVal><c:yVal><c:numRef><c:f>'Veri'!$C$2:$C$5</c:f></c:numRef></c:yVal></c:ser><c:axId val=\"10\"/><c:axId val=\"100\"/></c:scatterChart>" +
        '<c:valAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="100"/></c:valAx><c:valAx><c:axId val="100"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="10"/></c:valAx>',
    ),
  ) as Chart;
  assert.deepEqual(fitted.groups[0].series[0].trendlines, [{ type: 'poly', order: 3, forward: 2, backward: 1 }]);
  assert.deepEqual(fitted.groups[0].series[0].errorBars, [{ direction: 'y', type: 'plus', valueType: 'fixedVal', value: 2, noEndCap: true }]);

  // Data labels: what stands between the parts as an attribute (the schema says: as text), a format without "sourceLinked".
  const labelled = readChart(
    openpyxlChart(
      '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:dLbls><c:numFmt formatCode="0.0"/><c:dLblPos val="outEnd"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="1"/><c:showSerName val="0"/><c:separator val=" | "/></c:dLbls>' +
        '<c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>Oca</c:v></c:pt><c:pt idx="1"><c:v>Şub</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>30</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val="10"/><c:axId val="100"/></c:barChart>' +
        '<c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="100"/></c:catAx><c:valAx><c:axId val="100"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="10"/></c:valAx>',
    ),
  ) as Chart;
  assert.deepEqual(labelled.groups[0].series[0].dataLabels, { showValue: true, showCategory: true, showSeries: false, showLegendKey: false, separator: ' | ', formatCode: '0.0', position: 'outEnd' });
  assert.deepEqual(items(layoutChart(labelled, chartValues(labelled), 640, 400, fakeMeasure), 'text', 'label').map((item) => item.text), ['Oca | 10.0', 'Şub | 30.0'], 'a format that does not say it is that of the cells is the format of the labels');

  // A place set by hand: openpyxl writes what the width and the height mean ("factor": sizes). As edges they are where the place ends.
  const placed = (modes: string, w: number, h: number) =>
    readChart(
      chartXml({ plot: '<c:pieChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser></c:pieChart>' }).replace(
        '<c:plotArea><c:layout/>',
        `<c:plotArea><c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/>${modes}<c:x val="0.25"/><c:y val="0.25"/><c:w val="${w}"/><c:h val="${h}"/></c:manualLayout></c:layout>`,
      ),
    ) as Chart;
  const sized = placed('<c:wMode val="factor"/><c:hMode val="factor"/>', 0.5, 0.5);
  assert.deepEqual(sized.plotLayout, { target: 'inner', xMode: 'edge', yMode: 'edge', x: 0.25, y: 0.25, w: 0.5, h: 0.5 });
  const edged = placed('<c:wMode val="edge"/><c:hMode val="edge"/>', 0.75, 0.75);
  assert.deepEqual(edged.plotLayout, { target: 'inner', xMode: 'edge', yMode: 'edge', wMode: 'edge', hMode: 'edge', x: 0.25, y: 0.25, w: 0.75, h: 0.75 });
  const centreOf = (chart: Chart) => {
    const slice = items(layoutChart(chart, chartValues(chart), 400, 400, fakeMeasure), 'slice')[0];
    return [slice.cx, slice.cy, slice.radius];
  };
  assert.deepEqual(centreOf(edged), centreOf(sized), 'UNSURE: from a quarter to three quarters is the same place as a quarter in and half the frame wide');
  assert.deepEqual(centreOf(sized).slice(0, 2), [200, 200]);
});
