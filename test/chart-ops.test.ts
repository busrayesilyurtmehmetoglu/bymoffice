// Charts as operations: a chart added from a range, moved, changed and deleted, each one
// step that undo takes back exactly; what a range gives a chart; what rows, columns and
// sheets do to the references of charts; and what is refused.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import type { Chart, DrawingChart, Sheet, SheetDrawing, Workbook } from '../src/index';
import { createSheetMetrics } from '../src/layout/sheet-metrics';
import { chartRefs } from '../src/model/sheet/chart';
import { addChart, chartProblem, drawingAnchorNow, moveChart, removeChart, seriesOfRange, sheetCharts, updateChart } from '../src/ops/charts';
import { History, applyOp, type Op } from '../src/ops/ops';
import { SheetDrawings } from '../src/render/sheet-drawings';
import { bookOf } from './chart-edit-fixtures';

type Value = number | string | null;

const SALES: Value[][] = [
  ['Month', 'North', 'South'],
  ['Jan', 10, 30],
  ['Feb', 20, 20],
  ['Mar', 30, 10],
];

const chartOf = (sheet: Sheet, id = 0): Chart => ((sheet.drawings as SheetDrawing[])[id].object as DrawingChart).chart;
const refsOf = (sheet: Sheet, id = 0): string[] => chartRefs(chartOf(sheet, id));
const all = { top: 0, left: 0, bottom: 3, right: 2 };

test('charts: a range with a row of names and a column of names gives one series per column', () => {
  const wb = bookOf(SALES);
  const made = seriesOfRange(wb, 0, all);
  assert.equal(made.seriesIn, 'columns');
  assert.equal(made.headerRow, true, 'the first row holds no number');
  assert.equal(made.headerColumn, true, 'the first column holds no number');
  assert.deepEqual(made.series, [
    { values: 'Data!$B$2:$B$4', nameRef: 'Data!$B$1', categories: 'Data!$A$2:$A$4' },
    { values: 'Data!$C$2:$C$4', nameRef: 'Data!$C$1', categories: 'Data!$A$2:$A$4' },
  ]);
  // In rows: the names come from the first column, the categories from the first row.
  assert.deepEqual(seriesOfRange(wb, 0, all, 'rows').series[0], { values: 'Data!$B$2:$C$2', nameRef: 'Data!$A$2', categories: 'Data!$B$1:$C$1' });
});

test('charts: a range of numbers alone has no names; more columns than rows of values puts the series in rows (UNSURE)', () => {
  const wb = bookOf([
    [1, 2, 3, 4],
    [5, 6, 7, 8],
  ]);
  const made = seriesOfRange(wb, 0, { top: 0, left: 0, bottom: 1, right: 3 });
  assert.equal(made.headerRow, false);
  assert.equal(made.headerColumn, false);
  assert.equal(made.seriesIn, 'rows');
  assert.deepEqual(made.series, [{ values: 'Data!$A$1:$D$1' }, { values: 'Data!$A$2:$D$2' }]);
  // A scatter chart takes the first column as its horizontal values whatever it holds.
  const scatter = seriesOfRange(wb, 0, { top: 0, left: 0, bottom: 1, right: 3 }, 'columns', 'scatter');
  assert.deepEqual(scatter.series[0], { values: 'Data!$B$1:$B$2', categories: 'Data!$A$1:$A$2' });
  // One column under a name: the name is that of the series.
  assert.deepEqual(seriesOfRange(bookOf([['Units'], [4], [5]]), 0, { top: 0, left: 0, bottom: 2, right: 0 }).series, [{ values: 'Data!$A$2:$A$3', nameRef: 'Data!$A$1' }]);
});

test('charts: a sheet name that needs quotes is quoted in the references', () => {
  const wb = bookOf(SALES);
  wb.sheets[0].name = "My 'data'";
  assert.equal(seriesOfRange(wb, 0, all).series[0].values, "'My ''data'''!$B$2:$B$4");
});

test('charts: a chart is added, moved, changed and deleted, each as one step that undo takes back exactly', () => {
  const wb = bookOf(SALES);
  const sheet = wb.sheets[0];
  const history = new History(wb);
  history.apply(addChart(wb, 0, { type: 'column', range: all, title: 'Sales' }), undefined, undefined);
  assert.equal(sheet.drawings?.length, 1);
  assert.equal(sheet.drawingsModified, true);
  assert.equal(sheet.modified, undefined, 'no cell changed');
  const added = sheet.drawings;
  const chart = chartOf(sheet);
  assert.equal(chart.live, true);
  assert.equal(chart.groups[0].kind, 'bar');
  assert.equal(chart.groups[0].barDir, 'col');
  assert.equal(chart.title?.text, 'Sales');
  assert.equal(chart.legend?.position, 'r');
  assert.deepEqual(refsOf(sheet), ['Data!$B$1', 'Data!$A$2:$A$4', 'Data!$B$2:$B$4', 'Data!$C$1', 'Data!$A$2:$A$4', 'Data!$C$2:$C$4']);
  assert.equal(chart.groups[0].series[0].categories?.kind, 'str', 'texts are written as texts');
  // Beside the range, in the default size.
  const rect = new SheetDrawings(sheet).rects(createSheetMetrics(sheet))[0];
  assert.deepEqual([Math.round(rect.width), Math.round(rect.height)], [480, 288]);
  assert.ok(rect.x >= 3 * 64, `the chart starts right of the range: ${rect.x}`);

  history.apply(moveChart(wb, 0, 0, { top: 10, left: 1, bottom: 19, right: 5 }), undefined, undefined);
  assert.deepEqual(sheetCharts(wb, 0)[0].anchor, { kind: 'twoCell', from: { col: 1, colOff: 0, row: 10, rowOff: 0 }, to: { col: 6, colOff: 0, row: 20, rowOff: 0 } });
  assert.deepEqual(moveChart(wb, 0, 0, { top: 10, left: 1, bottom: 19, right: 5 }), [], 'a chart that is there already is not moved');

  history.apply(updateChart(wb, 0, 0, { type: 'pie', title: 'Share', legend: 'b' }), undefined, undefined);
  const pie = chartOf(sheet);
  assert.equal(pie.groups[0].kind, 'pie');
  assert.deepEqual(pie.axes, [], 'a pie has no axes');
  assert.equal(pie.title?.text, 'Share');
  assert.equal(pie.legend?.position, 'b');
  assert.deepEqual(refsOf(sheet).length, 6, 'the series stay');

  history.apply(removeChart(wb, 0, 0), undefined, undefined);
  assert.equal(sheet.drawings, undefined);

  history.undo();
  assert.equal(chartOf(sheet), pie);
  history.undo();
  assert.equal(chartOf(sheet), chart);
  history.undo();
  assert.equal(sheet.drawings, added, 'undo puts the very list back');
  history.undo();
  assert.equal(sheet.drawings, undefined);
  assert.equal(history.canUndo, false);
  for (let i = 0; i < 4; i++) history.redo();
  assert.equal(sheet.drawings, undefined);
  history.undo();
  assert.equal(chartOf(sheet), pie);
});

test('charts: every type is made with the axes it needs', () => {
  const wb = bookOf(SALES);
  const kinds: [Parameters<typeof addChart>[2]['type'], string, number][] = [
    ['column', 'bar', 2],
    ['bar', 'bar', 2],
    ['line', 'line', 2],
    ['area', 'area', 2],
    ['pie', 'pie', 0],
    ['doughnut', 'doughnut', 0],
    ['scatter', 'scatter', 2],
  ];
  for (const [type, kind, axes] of kinds) {
    const [op] = addChart(wb, 0, { type, range: all, grouping: 'stacked' });
    const chart = ((op as Extract<Op, { type: 'setDrawings' }>).drawings as SheetDrawing[])[0].object as DrawingChart;
    assert.equal(chart.chart.groups[0].kind, kind, type);
    assert.equal(chart.chart.axes.length, axes, type);
    assert.equal(chart.part, undefined, 'a chart that is in no file has no part');
    assert.equal(sheetChartsOf(wb, op)[0].type, type, 'the type is read back from the model');
  }
  const [bar] = addChart(wb, 0, { type: 'bar', range: all });
  const model = (((bar as Extract<Op, { type: 'setDrawings' }>).drawings as SheetDrawing[])[0].object as DrawingChart).chart;
  assert.equal(model.axes[0].position, 'l', 'the categories of lying bars stand on the left');
  assert.equal(model.groups[0].barDir, 'bar');
  const [scatter] = addChart(wb, 0, { type: 'scatter', range: { top: 1, left: 1, bottom: 3, right: 2 } });
  const points = (((scatter as Extract<Op, { type: 'setDrawings' }>).drawings as SheetDrawing[])[0].object as DrawingChart).chart;
  assert.deepEqual(points.axes.map((axis) => axis.kind), ['val', 'val'], 'a scatter chart has two value axes');
  assert.equal(points.groups[0].series[0].categories?.kind, 'num');
});

function sheetChartsOf(wb: Workbook, op: Op) {
  const undo = applyOp(wb, op);
  const list = sheetCharts(wb, 0);
  applyOp(wb, undo);
  return list;
}

test('charts: what cannot be made is refused with a reason', () => {
  const wb = bookOf(SALES);
  assert.throws(() => addChart(wb, 0, { type: 'radar' as 'line', range: all }), /A chart is one of/);
  assert.throws(() => addChart(wb, 0, { type: 'line' }), /needs a range/);
  assert.throws(() => addChart(wb, 0, { type: 'line', series: [{ values: 'Nowhere!A1:A3' }] }), /not a range of cells of this workbook/);
  assert.throws(() => addChart(wb, 0, { type: 'line', range: all, title: 'x'.repeat(256) }), /longer than 255/);
  assert.throws(() => addChart(wb, 0, { type: 'line', range: all, legend: 'middle' as 'r' }), /legend/);
  assert.throws(() => addChart(wb, 5, { type: 'line', range: all }), /No sheet/);
  assert.throws(() => removeChart(wb, 0, 0), /No chart at index 0/);
  assert.throws(() => moveChart(wb, 0, 3, all), /No chart/);
  assert.throws(() => updateChart(wb, 0, 0, { title: 'x' }), /No chart/);
  applyOp(wb, addChart(wb, 0, { type: 'line', range: all })[0]);
  assert.throws(() => moveChart(wb, 0, 0, { kind: 'twoCell' } as never), /place of a chart/);
  assert.throws(() => updateChart(wb, 0, 0, { series: [] }), /at least one series/);
  assert.throws(() => updateChart(wb, 0, 0, { series: [{ from: 0 }, { from: 0 }] }), /named twice/);
  assert.throws(() => updateChart(wb, 0, 0, { series: [{ from: 0, color: 'red' }] }), /#RRGGBB/);
  assert.throws(() => updateChart(wb, 0, 0, { axisBounds: { min: 5, max: 1 } }), /larger than its smallest/);
  assert.deepEqual(updateChart(wb, 0, 0, { title: null, legend: 'r', type: 'line' }), [], 'what the chart has already is no change');
});

test('charts: the series of a chart are renamed, recoloured, reordered, added and taken away', () => {
  const wb = bookOf(SALES);
  const sheet = wb.sheets[0];
  applyOp(wb, addChart(wb, 0, { type: 'column', range: all })[0]);
  applyOp(wb, updateChart(wb, 0, 0, { series: [{ from: 1, color: '#ff0000', name: 'Güney' }, { from: 0 }, { values: 'B2:B4', name: 'Again' }] })[0]);
  const series = chartOf(sheet).groups[0].series;
  assert.deepEqual(
    series.map((item) => [item.index, item.order, item.name, item.values?.ref, item.shape]),
    [
      [1, 0, { text: 'Güney' }, 'Data!$C$2:$C$4', { fill: '#FF0000' }],
      [0, 1, { ref: 'Data!$B$1' }, 'Data!$B$2:$B$4', undefined],
      [2, 2, { text: 'Again' }, 'Data!B2:B4', undefined],
    ],
  );
  assert.equal(series[2].categories?.ref, 'Data!$A$2:$A$4', 'a new series takes the categories of the first');
  const info = sheetCharts(wb, 0)[0];
  assert.deepEqual(info.series[0], { name: 'Güney', categories: 'Data!$A$2:$A$4', values: 'Data!$C$2:$C$4', color: '#FF0000' });
  assert.equal(info.range, undefined, 'the series are no longer one block of cells');

  applyOp(wb, updateChart(wb, 0, 0, { range: { top: 0, left: 0, bottom: 3, right: 1 } })[0]);
  const again = sheetCharts(wb, 0)[0];
  assert.deepEqual(again.range, { top: 0, left: 0, bottom: 3, right: 1 });
  assert.equal(again.seriesIn, 'columns');
  assert.equal(again.series.length, 1);
  assert.equal(again.series[0].color, '#FF0000', 'the series that was first keeps its look');

  applyOp(wb, updateChart(wb, 0, 0, { dataLabels: true, axisTitles: { category: 'Month', value: 'Units' }, axisBounds: { min: 0, max: 50 }, grouping: 'stacked' })[0]);
  const now = sheetCharts(wb, 0)[0];
  assert.equal(now.dataLabels, true);
  assert.deepEqual(now.axisTitles, { category: 'Month', value: 'Units' });
  assert.deepEqual(now.axisBounds, { min: 0, max: 50 });
  assert.equal(now.grouping, 'stacked');
  assert.equal(chartOf(sheet).groups[0].overlap, 100, 'stacked bars overlap');
  applyOp(wb, updateChart(wb, 0, 0, { dataLabels: false, axisBounds: { min: null }, axisTitles: { value: null } })[0]);
  const last = sheetCharts(wb, 0)[0];
  assert.equal(last.dataLabels, false);
  assert.deepEqual(last.axisBounds, { max: 50 });
  assert.deepEqual(last.axisTitles, { category: 'Month' });
});

test('charts: rows and columns that move take the references of the charts along; cells that are gone leave #REF!', () => {
  const wb = bookOf(SALES, [{ name: 'Other', rows: [[1]] }]);
  const sheet = wb.sheets[0];
  const history = new History(wb);
  history.apply(addChart(wb, 0, { type: 'line', range: all, anchor: { top: 6, left: 0, bottom: 15, right: 4 } }), undefined, undefined);
  const before = sheet.drawings;
  // A row inside the values: the ranges grow.
  history.apply([{ type: 'insert', sheet: 0, axis: 'row', at: 2, count: 1 }], undefined, undefined);
  assert.deepEqual(refsOf(sheet), ['Data!$B$1', 'Data!$A$2:$A$5', 'Data!$B$2:$B$5', 'Data!$C$1', 'Data!$A$2:$A$5', 'Data!$C$2:$C$5']);
  assert.notEqual(sheet.drawings, before, 'the list is replaced, not altered');
  assert.deepEqual(drawingAnchorNow((sheet.drawings as SheetDrawing[])[0], sheet.structuralEdits).from, { col: 0, colOff: 0, row: 7, rowOff: 0 }, 'the chart moves down with its cells');
  // A column before them: they move.
  history.apply([{ type: 'insert', sheet: 0, axis: 'col', at: 0, count: 2 }], undefined, undefined);
  assert.deepEqual(refsOf(sheet).slice(0, 3), ['Data!$D$1', 'Data!$C$2:$C$5', 'Data!$D$2:$D$5']);
  // The column of the second series goes: its references are errors.
  history.apply([{ type: 'delete', sheet: 0, axis: 'col', at: 4, count: 1 }], undefined, undefined);
  assert.deepEqual(refsOf(sheet).slice(3), ['Data!#REF!', 'Data!$C$2:$C$5', 'Data!#REF!']);
  // Rows of another sheet do nothing to them.
  const now = sheet.drawings;
  history.apply([{ type: 'insert', sheet: 1, axis: 'row', at: 0, count: 3 }], undefined, undefined);
  assert.equal(sheet.drawings, now);
  history.undo();
  history.undo();
  history.undo();
  history.undo();
  assert.equal(sheet.drawings, before, 'undo puts the very list back');
  assert.deepEqual(refsOf(sheet), ['Data!$B$1', 'Data!$A$2:$A$4', 'Data!$B$2:$B$4', 'Data!$C$1', 'Data!$A$2:$A$4', 'Data!$C$2:$C$4']);
});

test('charts: a sheet that is renamed is renamed in the references of every chart; one that is deleted leaves #REF!', () => {
  const wb = bookOf(SALES, [{ name: 'Other', rows: [[1]] }]);
  const history = new History(wb);
  history.apply(addChart(wb, 1, { type: 'line', series: [{ values: 'Data!$B$2:$B$4', nameRef: 'Data!$B$1' }] }), undefined, undefined);
  history.apply([{ type: 'renameSheet', sheet: 0, name: 'Sales 2024' }], undefined, undefined);
  assert.deepEqual(refsOf(wb.sheets[1]), ["'Sales 2024'!$B$1", "'Sales 2024'!$B$2:$B$4"]);
  history.apply([{ type: 'deleteSheet', sheet: 0 }], undefined, undefined);
  assert.deepEqual(refsOf(wb.sheets[0]), ['#REF!', '#REF!']);
  history.undo();
  history.undo();
  assert.deepEqual(refsOf(wb.sheets[1]), ['Data!$B$1', 'Data!$B$2:$B$4']);
});

test('charts: a workbook whose file has charts keeps its rows, columns and sheet names unless its charts follow', () => {
  const wb = bookOf(SALES, [{ name: 'Other', rows: [[1]] }]);
  // As the reader leaves a workbook with a chart part it cannot keep up to date.
  for (const sheet of wb.sheets) sheet.anchored = ['chart'];
  assert.throws(() => applyOp(wb, { type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 }), /sheet with chart/);
  assert.throws(() => applyOp(wb, { type: 'renameSheet', sheet: 0, name: 'X' }), /workbook with chart/);
  assert.throws(() => applyOp(wb, { type: 'deleteSheet', sheet: 1 }), /workbook with chart/);
  wb.chartsFollow = true;
  applyOp(wb, { type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 });
  applyOp(wb, { type: 'renameSheet', sheet: 0, name: 'X' });
  applyOp(wb, { type: 'deleteSheet', sheet: 1 });
  assert.equal(wb.sheets.length, 1);
});

test('charts: a chart of the file is changed by what its part can take, and only moved and deleted when its references are not known', () => {
  const wb = bookOf(SALES);
  const sheet = wb.sheets[0];
  applyOp(wb, addChart(wb, 0, { type: 'column', range: all })[0]);
  // As a reader leaves a chart of a file: no count of edits, a part, and what is known of the part.
  const made = (sheet.drawings as SheetDrawing[])[0];
  const read = (known?: 'all' | 'refs'): SheetDrawing => {
    const chart = { ...(made.object as DrawingChart).chart };
    delete chart.live;
    const object: DrawingChart = { type: 'chart', chart, part: '/xl/charts/chart1.xml' };
    if (known) object.known = known;
    return { anchor: made.anchor, object };
  };
  const fromFile = read('refs');
  sheet.drawings = [fromFile];
  assert.equal(sheetCharts(wb, 0)[0].editable, 'parts');
  assert.equal(sheetCharts(wb, 0)[0].ofFile, true);
  assert.match(chartProblem(wb, 0, 0, { type: 'line' }) || '', /Cannot change the type/);
  assert.throws(() => updateChart(wb, 0, 0, { type: 'line' }), /Cannot change the type/);
  assert.deepEqual(updateChart(wb, 0, 0, { type: 'column' }), [], 'the type it has is no change');
  applyOp(wb, updateChart(wb, 0, 0, { title: 'T', series: [{ from: 1 }, { values: 'B2:B4' }] })[0]);
  const changed = (sheet.drawings as SheetDrawing[])[0];
  assert.equal(changed.from, fromFile, 'the drawing knows the drawing of the file it was');
  const object = changed.object as DrawingChart;
  assert.equal(object.origin?.chart, (fromFile.object as DrawingChart).chart, 'and the chart as it was read');
  assert.deepEqual(object.origin?.series, [[1, -1]], 'and which series of the file each of its series is');
  applyOp(wb, updateChart(wb, 0, 0, { series: [{ from: 1 }, { from: 0 }] })[0]);
  assert.deepEqual(((sheet.drawings as SheetDrawing[])[0].object as DrawingChart).origin?.series, [[-1, 1]]);
  applyOp(wb, moveChart(wb, 0, 0, { top: 0, left: 5, bottom: 5, right: 9 })[0]);
  assert.equal((sheet.drawings as SheetDrawing[])[0].from, fromFile, 'a move keeps what the drawing came from');

  const unknown = read();
  sheet.drawings = [unknown];
  assert.equal(sheetCharts(wb, 0)[0].editable, 'none');
  assert.throws(() => updateChart(wb, 0, 0, { title: 'T' }), /only moved and deleted/);
  applyOp(wb, moveChart(wb, 0, 0, { top: 20, left: 5, bottom: 25, right: 9 })[0]);
  assert.equal((sheet.drawings as SheetDrawing[])[0].from, unknown);
  applyOp(wb, removeChart(wb, 0, 0)[0]);
  assert.equal(sheet.drawings, undefined);

  sheet.drawings = [read('all')];
  assert.equal(sheetCharts(wb, 0)[0].editable, 'all');
  applyOp(wb, updateChart(wb, 0, 0, { type: 'line' })[0]);
  assert.equal(chartOf(sheet).groups[0].kind, 'line');
  assert.equal(((sheet.drawings as SheetDrawing[])[0].object as DrawingChart).origin, undefined, 'a chart the engine wrote is written anew: nothing of the part is kept');
});
