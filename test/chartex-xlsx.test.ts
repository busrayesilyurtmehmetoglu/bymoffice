// Charts of the newer types (chartex) and surface charts in a workbook: the frame of such a
// chart in the alternate content Excel writes is shown as the chart where its part can be
// drawn, as the fallback beside it where it cannot, never both; the series read the cells of
// the sheet; nothing of the parts is ever written anew, and no command changes such a chart.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { chartValues, layoutChart } from '../src/layout/chart';
import { MAX_WORKBOOK_CHART_POINTS, type Chart } from '../src/model/sheet/chart';
import { chartCopyProblem, chartProblem, moveChart, removeChart, updateChart } from '../src/ops/charts';
import { applyOp } from '../src/ops/ops';
import { catAx, chartWorkbook, chartXml, fakeMeasure, group, numRef, ser, sheetData, strRef, valAx } from './chart-fixtures';
import { chartsOf, load, names, packageProblems, part, reopen, same, save } from './chart-edit-fixtures';
import { addChartEx, anchor, chartExFrame, chartExXml, CATS, data, numDim, series, strDim, treemapXml, waterfallXml } from './chartex-fixtures';

const CELLS = sheetData([
  ['Ad', 'Değer', 'Orta', 'Yaprak', 'Boyut'],
  ['Start', 200, 'A', 'a1', 1],
  ['Q1', 10, 'A', 'a2', 2],
  ['Q2', -20, 'B', 'b1', 3],
  ['Q3', 30, 'B', 'b2', 4],
  ['End', 220, 'C', 'c1', 5],
  [null, null, 'C', 'c2', 6],
]);

/** A workbook whose first sheet shows the given parts of newer chart types, each in its own anchor, and the older chart parts given. */
async function book(parts: string[], older: string[] = [], frames?: string): Promise<Uint8Array> {
  const drawing =
    frames ??
    parts.map((_xml, i) => anchor(chartExFrame(`rIdX${i + 1}`, `Chart ${i + 1}`), [6, i * 16], [14, i * 16 + 15])).join('') +
      older.map((_xml, i) => anchor(chartFrame(i), [16, i * 16], [24, i * 16 + 15])).join('');
  return addChartEx(await chartWorkbook({ sheet: CELLS, drawing, charts: older }), parts);
}

const chartFrame = (i: number) =>
  `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${20 + i}" name="Old ${i}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` +
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdC${i + 1}"/></a:graphicData></a:graphic></xdr:graphicFrame>`;

const regionMap = () =>
  chartExXml({ data: data(0, strDim('Sayfa1!$A$2:$A$3', [['Türkiye', 'Fransa']]) + numDim('Sayfa1!$B$2:$B$3', [1, 2])), series: series({ layout: 'regionMap', dataId: 0 }) });

test('a chart of a newer type in its alternate content is the chart, read from the cells of the sheet', async () => {
  const doc = await load(await book([waterfallXml()]));
  const sheet = doc.wb.sheets[0];
  const [object] = chartsOf(sheet);
  assert.ok(object, 'a chart, not a placeholder or the fallback');
  assert.equal(object.part, '/xl/charts/chartEx1.xml');
  assert.equal(object.name, 'Chart 1');
  assert.equal(object.chart.extended, true);
  assert.equal(object.chart.groups[0].kind, 'waterfall');
  assert.equal(sheet.drawings?.length, 1, 'the fallback is not shown beside it');
  // The cells, not the cache: the sheet says 200, 10, -20, 30, 220.
  const values = chartValues(object.chart, { workbook: doc.wb });
  assert.deepEqual(values.series[0].values.toArray(), [200, 10, -20, 30, 220]);
  assert.equal(values.series[0].live, true);
  assert.deepEqual([0, 1, 2, 3, 4].map((i) => values.label(i)), CATS);
  const layout = layoutChart(object.chart, values, 480, 300, fakeMeasure);
  assert.equal(layout.items.filter((item) => item.kind === 'rect' && item.role === 'bar').length, 5);
  assert.deepEqual(doc.wb.sheets[0].warnings, undefined);
});

test('the levels of a treemap are read from the columns of its reference: the last is the leaves', async () => {
  // The part caches other names: the cells win.
  const xml = treemapXml().replace("Sayfa1!$A$2:$C$7", 'Sayfa1!$C$2:$D$7').replace('Sayfa1!$D$2:$D$7', 'Sayfa1!$E$2:$E$7');
  const doc = await load(await book([xml]));
  const [object] = chartsOf(doc.wb.sheets[0]);
  assert.equal(object.chart.groups[0].kind, 'treemap');
  const values = chartValues(object.chart, { workbook: doc.wb });
  assert.ok(values.levels);
  assert.equal(values.levels.depth, 2);
  assert.deepEqual([0, 1, 2, 3, 4, 5].map((i) => values.levels?.name(0, i)), ['a1', 'a2', 'b1', 'b2', 'c1', 'c2']);
  assert.deepEqual([0, 1, 2, 3, 4, 5].map((i) => values.levels?.name(1, i)), ['A', 'A', 'B', 'B', 'C', 'C']);
  assert.deepEqual(values.series[0].values.toArray(), [1, 2, 3, 4, 5, 6]);
  const layout = layoutChart(object.chart, values, 480, 300, fakeMeasure);
  // Three parents and six leaves.
  assert.equal(layout.items.filter((item) => item.kind === 'rect' && item.role === 'bar').length, 9);
});

test('a chart of a newer type that is not drawn shows the fallback of its alternate content, not a placeholder', async () => {
  const doc = await load(await book([regionMap(), waterfallXml()]));
  const drawings = doc.wb.sheets[0].drawings || [];
  assert.equal(drawings.length, 2);
  assert.equal(drawings[0].object.type, 'shape', 'the text box Excel puts beside a region map');
  const shape = drawings[0].object as Extract<(typeof drawings)[number]['object'], { type: 'shape' }>;
  assert.ok(JSON.stringify(shape.text).indexOf('available in your version') >= 0);
  assert.equal(drawings[1].object.type, 'chart');
  // A part of a newer type that cannot be read at all: the fallback too.
  const broken = await load(await book(['<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart>']));
  assert.equal(broken.wb.sheets[0].drawings?.[0].object.type, 'shape');
  // Without alternate content around it, a frame whose chart is not drawn is a placeholder.
  const bare = `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="X"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdX1"/></a:graphicData></a:graphic></xdr:graphicFrame>`;
  const placeholder = await reopen(await book([regionMap()], [], anchor(bare)));
  assert.deepEqual(placeholder.sheets[0].drawings?.map((drawing) => drawing.object.type), ['placeholder']);
  const shown = await reopen(await book([waterfallXml()], [], anchor(bare)));
  assert.deepEqual(shown.sheets[0].drawings?.map((drawing) => drawing.object.type), ['chart']);
});

test('a workbook with charts of newer types is saved with their parts as they were', async () => {
  const bytes = await book([waterfallXml(), treemapXml(), regionMap()]);
  const doc = await load(bytes);
  // A cell is changed: the sheet is written anew, the charts and their drawing are not.
  applyOp(doc.wb, { type: 'setCells', sheet: 0, cells: [{ row: 1, col: 1, cell: { kind: 'number', value: 999, style: 0 } }] });
  const saved = await save(doc);
  assert.deepEqual(await packageProblems(saved), []);
  for (const name of names(bytes).filter((entry) => /^xl\/(charts|drawings)\//.test(entry))) same(await part(saved, name), await part(bytes, name), name);
  assert.deepEqual(names(saved), names(bytes));
  // Read again, the chart follows the cell.
  const again = await reopen(saved);
  const [object] = chartsOf(again.sheets[0]);
  assert.equal(chartValues(object.chart, { workbook: again }).series[0].values.at(0), 999);
});

test('no command changes, moves, copies or deletes a chart of a newer type; rows stay where they are', async () => {
  const doc = await load(await book([waterfallXml()]));
  const expected = /a chart of this type is only shown/;
  assert.throws(() => updateChart(doc.wb, 0, 0, { title: 'x' }), expected);
  assert.throws(() => moveChart(doc.wb, 0, 0, { top: 1, left: 1, bottom: 5, right: 5 }), expected);
  assert.throws(() => removeChart(doc.wb, 0, 0), expected);
  assert.match(chartProblem(doc.wb, 0, 0) || '', expected);
  assert.match(chartCopyProblem(doc.wb, 0, 0) || '', /a chart of this type is only shown/);
  assert.equal(doc.wb.chartsFollow, undefined);
  assert.throws(() => applyOp(doc.wb, { type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 }), /sheet with chart/);
});

test('a surface chart is a chart of the sheet now; it is moved and deleted like any chart whose part the model does not know', async () => {
  const surface = chartXml({
    plot:
      group('surfaceChart', '', ser(0, { cat: strRef('Sayfa1!$A$2:$A$4', ['a', 'b', 'c']), val: numRef('Sayfa1!$B$2:$B$4', [1, 2, 3]) }) + ser(1, { cat: strRef('Sayfa1!$A$2:$A$4', ['a', 'b', 'c']), val: numRef('Sayfa1!$E$2:$E$4', [3, 2, 1]) })) +
      catAx() +
      valAx(),
  });
  const doc = await load(await chartWorkbook({ sheet: CELLS, drawing: anchor(chartFrame(0)), charts: [surface] }));
  const [object] = chartsOf(doc.wb.sheets[0]);
  assert.ok(object);
  assert.equal(object.chart.groups[0].kind, 'surface');
  assert.equal(object.known, undefined, 'its part is not changed in place');
  assert.throws(() => updateChart(doc.wb, 0, 0, { title: 'x' }), /only moved and deleted/);
  assert.equal(doc.wb.chartsFollow, undefined);
  // Moved, saved, read again: still the chart, at its new place.
  doc.history.apply(moveChart(doc.wb, 0, 0, { top: 20, left: 2, bottom: 30, right: 8 }), undefined, undefined);
  const saved = await save(doc);
  assert.deepEqual(await packageProblems(saved), []);
  const again = await reopen(saved);
  const [moved] = chartsOf(again.sheets[0]);
  assert.equal(moved.chart.groups[0].kind, 'surface');
  same(await part(saved, 'xl/charts/chart1.xml'), await part(await chartWorkbook({ sheet: CELLS, drawing: anchor(chartFrame(0)), charts: [surface] }), 'xl/charts/chart1.xml'), 'the chart part');
});

test('charts of newer types share the limits of the workbook with the other charts', async () => {
  // Points: each part holds ten cached values; a workbook may hold MAX_WORKBOOK_CHART_POINTS of them.
  assert.ok(MAX_WORKBOOK_CHART_POINTS >= 10);
  const many = Array.from({ length: 30 }, () => waterfallXml());
  const wb = await reopen(await book(many));
  const charts = chartsOf(wb.sheets[0]);
  assert.equal(charts.length, 30);
  // One part named by many frames is read once.
  const frames = Array.from({ length: 5 }, (_v, i) => anchor(chartExFrame('rIdX1', `C${i}`), [0, i * 5], [4, i * 5 + 4])).join('');
  const shared = await reopen(await book([waterfallXml()], [], frames));
  const list = chartsOf(shared.sheets[0]);
  assert.equal(list.length, 5);
  assert.ok(list.every((item) => item.chart === list[0].chart), 'one model for one part');
  void ([] as Chart[]);
});
