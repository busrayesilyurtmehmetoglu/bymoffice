// Charts that are copied, cut, pasted and duplicated: the operations, what a copy of a chart of
// the file takes along of its part, and what the writer makes of it.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { ZipArchive, parseXml, type Sheet } from '../src/index';
import { chartLoose, chartKnown } from '../src/formats/xlsx/chart-write';
import { parseChart } from '../src/formats/xlsx/chart';
import { chartRefs } from '../src/model/sheet/chart';
import { CHART_COPY_OFFSET, addChart, chartCopyOutdated, chartCopyProblem, copyChart, cutChart, duplicateChart, moveChart, pasteChart, removeChart, sheetCharts, updateChart } from '../src/ops/charts';
import { History } from '../src/ops/ops';
import { deleteSheet, insertSheet, moveSheet, renameSheet } from '../src/ops/sheets';
import { anchorRect, drawingAnchorNow } from '../src/render/sheet-drawings';
import { createSheetMetrics } from '../src/layout/sheet-metrics';
import { COLORS, LATIN_TITLE, RICH, SHAPED, STYLE, UNKNOWN, latinChart, richBook } from './chart-copy-fixtures';
import { pythonHere, readByPython } from './chart-copy-python';
import { SALES, bare, bookOf, chartsOf, encodingProblem, load, names, packageProblems, part, relationProblems, reopen, same, save, text, type Loaded } from './chart-edit-fixtures';
import { OFFICE_THEME } from './chart-fixtures';

const ALL = { top: 0, left: 0, bottom: 3, right: 2 };
const whole = async (bytes: Uint8Array, label: string) => assert.deepEqual(await packageProblems(bytes), [], label);
const apply = (doc: Loaded, ops: Parameters<History['apply']>[0]) => doc.history.apply(ops, undefined, undefined);
const rectOf = (sheet: Sheet, id: number) => anchorRect(drawingAnchorNow((sheet.drawings || [])[id], sheet.structuralEdits), createSheetMetrics(sheet));
const chartParts = (bytes: Uint8Array) => names(bytes).filter((name) => /^xl\/(charts|media)\//.test(name));

test('chart copy: a chart of the model is pasted beside itself, at a cell and on another sheet, each as one step', () => {
  const wb = bookOf(SALES, [{ name: 'Other', rows: [['x', 1]] }]);
  const history = new History(wb);
  history.apply(addChart(wb, 0, { type: 'column', range: ALL, title: 'Sales' }), undefined, undefined);
  const copy = copyChart(wb, 0, 0);
  assert.deepEqual(copy.dropped, []);
  const before = rectOf(wb.sheets[0], 0);

  const first = pasteChart(wb, 0, copy);
  assert.equal(first.id, 1);
  assert.deepEqual([first.dropped, first.unbound], [[], 0]);
  history.apply(first.ops, undefined, undefined);
  const one = rectOf(wb.sheets[0], 1);
  assert.deepEqual([Math.round(one.x - before.x), Math.round(one.y - before.y)], [CHART_COPY_OFFSET, CHART_COPY_OFFSET], 'beside the chart it was copied from');
  assert.deepEqual([Math.round(one.width), Math.round(one.height)], [Math.round(before.width), Math.round(before.height)], 'in its size');
  // A second paste of the same copy does not cover the first.
  history.apply(pasteChart(wb, 0, copy).ops, undefined, undefined);
  const two = rectOf(wb.sheets[0], 2);
  assert.equal(Math.round(two.x - before.x), 2 * CHART_COPY_OFFSET);
  const charts = chartsOf(wb.sheets[0]);
  assert.deepEqual(charts.map((chart) => chart.name), ['Chart 1', 'Chart 2', 'Chart 3']);
  assert.deepEqual(bare(charts[1].chart), bare(charts[0].chart), 'the copy says what the chart says');
  assert.notEqual(charts[1].chart, charts[0].chart, 'and is a chart of its own');
  assert.notEqual(charts[2].chart, charts[1].chart);
  assert.equal(sheetCharts(wb, 0)[1].editable, 'all');

  // At a cell: its corner is the corner of the cell.
  history.apply(pasteChart(wb, 0, copy, { row: 20, col: 3 }).ops, undefined, undefined);
  const placed = drawingAnchorNow((wb.sheets[0].drawings || [])[3], undefined);
  assert.deepEqual([placed.from?.row, placed.from?.col, placed.from?.rowOff, placed.from?.colOff], [20, 3, 0, 0]);
  assert.throws(() => pasteChart(wb, 0, copy, { row: -1, col: 0 }), /The place of a chart is a cell/);

  // On another sheet: where the chart stood, and still about the cells of the sheet it came from.
  const other = pasteChart(wb, 1, copy);
  assert.equal(other.id, 0);
  history.apply(other.ops, undefined, undefined);
  const there = rectOf(wb.sheets[1], 0);
  assert.deepEqual([Math.round(there.x), Math.round(there.y)], [Math.round(before.x), Math.round(before.y)]);
  assert.deepEqual(chartRefs(chartsOf(wb.sheets[1])[0].chart), chartRefs(charts[0].chart));
  assert.ok(chartRefs(charts[0].chart).every((ref) => /^Data!/.test(ref)));

  // Every paste is one step.
  for (let n = 0; n < 4; n++) assert.ok(history.undo());
  assert.equal((wb.sheets[0].drawings || []).length, 1);
  assert.equal(wb.sheets[1].drawings, undefined);
  for (let n = 0; n < 4; n++) assert.ok(history.redo());
  assert.equal((wb.sheets[0].drawings || []).length, 4);
  assert.equal((wb.sheets[1].drawings || []).length, 1);
});

test('chart copy: a chart is duplicated, and cut and pasted elsewhere, each as one step', () => {
  const wb = bookOf(SALES, [{ name: 'Other', rows: [['x', 1]] }]);
  const history = new History(wb);
  history.apply(addChart(wb, 0, { type: 'line', range: ALL }), undefined, undefined);
  const made = duplicateChart(wb, 0, 0);
  history.apply(made.ops, undefined, undefined);
  assert.equal(made.id, 1);
  assert.equal(chartsOf(wb.sheets[0]).length, 2);
  history.apply(updateChart(wb, 0, 1, { title: 'Only the copy' }), undefined, undefined);
  assert.deepEqual(sheetCharts(wb, 0).map((chart) => chart.title), [undefined, 'Only the copy'], 'a change of the copy is not one of the chart');
  history.undo();
  history.undo();
  assert.equal(chartsOf(wb.sheets[0]).length, 1, 'one step takes the duplicate away');

  const where = rectOf(wb.sheets[0], 0);
  const cut = cutChart(wb, 0, 0);
  history.apply(cut.ops, undefined, undefined);
  assert.equal(wb.sheets[0].drawings, undefined);
  // Where it was cut from nothing stands: pasted back, it stands where it stood.
  const back = pasteChart(wb, 0, cut.copy);
  history.apply(back.ops, undefined, undefined);
  assert.deepEqual([Math.round(rectOf(wb.sheets[0], 0).x), Math.round(rectOf(wb.sheets[0], 0).y)], [Math.round(where.x), Math.round(where.y)]);
  history.undo();
  history.apply(pasteChart(wb, 1, cut.copy).ops, undefined, undefined);
  assert.equal(chartsOf(wb.sheets[1]).length, 1);
  history.undo();
  history.undo();
  assert.equal(chartsOf(wb.sheets[0]).length, 1, 'one step brings the chart that was cut back');
});

test('chart copy: a copy is not pasted into another workbook, nor once rows, columns or sheets changed', () => {
  const wb = bookOf(SALES, [{ name: 'Other', rows: [['x', 1]] }]);
  const history = new History(wb);
  history.apply(addChart(wb, 0, { type: 'column', range: ALL }), undefined, undefined);
  const copy = copyChart(wb, 0, 0);
  const elsewhere = bookOf(SALES);
  assert.throws(() => pasteChart(elsewhere, 0, copy), /pasted into the workbook it was copied from/);
  assert.throws(() => pasteChart(wb, 0, {} as never), /not a chart that was copied/);

  history.apply([{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 }], undefined, undefined);
  assert.equal(chartCopyOutdated(wb, copy), true);
  assert.throws(() => pasteChart(wb, 1, copy), /rows, columns or sheets changed since it was copied/);
  history.undo();
  assert.equal(chartCopyOutdated(wb, copy), false, 'a step that was undone leaves the copy as good as it was');
  history.apply(pasteChart(wb, 1, copy).ops, undefined, undefined);

  history.apply(renameSheet(wb, 0, 'Veri'), undefined, undefined);
  assert.throws(() => pasteChart(wb, 1, copy), /changed since it was copied/);
  // The chart that was pasted before follows the name.
  assert.ok(chartRefs(chartsOf(wb.sheets[1])[0].chart).every((ref) => /^Veri!/.test(ref)));
  // Cells that change are nothing to a copy; neither are sheets that are added or put in another order.
  const again = copyChart(wb, 0, 0);
  history.apply([{ type: 'setCells', sheet: 0, cells: [{ row: 1, col: 1, cell: { kind: 'number', value: 99, style: 0 } }] }], undefined, undefined);
  assert.equal(chartCopyOutdated(wb, again), false);
  history.apply(insertSheet(wb, 0, { name: 'Yeni' }), undefined, undefined);
  history.apply(moveSheet(wb, 2, 0), undefined, undefined);
  assert.equal(chartCopyOutdated(wb, again), false);
  const onNew = pasteChart(wb, wb.sheets.findIndex((sheet) => sheet.name === 'Yeni'), again);
  history.apply(onNew.ops, undefined, undefined);
  assert.ok(chartRefs(chartsOf(wb.sheets.find((sheet) => sheet.name === 'Yeni') as Sheet)[0].chart).every((ref) => /^Veri!/.test(ref)));
  // A sheet the chart does not show goes: nothing to the copy. The sheet it shows goes: the copy is put down nowhere.
  history.apply(deleteSheet(wb, wb.sheets.findIndex((sheet) => sheet.name === 'Other')), undefined, undefined);
  assert.equal(chartCopyOutdated(wb, again), false);
  history.apply(deleteSheet(wb, wb.sheets.findIndex((sheet) => sheet.name === 'Veri')), undefined, undefined);
  assert.equal(chartCopyOutdated(wb, again), true);

  assert.match(chartCopyProblem(wb, 7, 0) || '', /No sheet/);
  assert.match(chartCopyProblem(wb, 0, 9) || '', /No chart at index 9/);
  assert.throws(() => copyChart(wb, 0, 9), RangeError);
  assert.throws(() => cutChart(wb, 0, 9), RangeError);
  assert.throws(() => duplicateChart(wb, 0, 9), RangeError);
});

test('chart copy: the copy of a chart of the file has its part byte for byte, and parts of its own for what the part leads to', async () => {
  const bytes = await richBook();
  await whole(bytes, 'the fixture is whole');
  const doc = await load(bytes);
  assert.deepEqual(chartsOf(doc.wb.sheets[0]).map((chart) => chart.known), ['refs', 'refs']);
  const copy = copyChart(doc.wb, 0, 0);
  const pasted = pasteChart(doc.wb, 1, copy);
  assert.deepEqual([pasted.dropped, pasted.unbound], [[], 0]);
  apply(doc, pasted.ops);
  const made = chartsOf(doc.wb.sheets[1])[0];
  assert.equal(made.part, undefined);
  assert.deepEqual(made.source, { part: '/xl/charts/chart1.xml' });
  assert.equal(made.known, 'refs');
  assert.equal(sheetCharts(doc.wb, 1)[0].editable, 'parts');
  assert.equal(sheetCharts(doc.wb, 1)[0].ofFile, false);

  const out = await save(doc);
  await whole(out, 'the package is whole');
  assert.deepEqual(
    chartParts(out),
    ['xl/charts/_rels/chart1.xml.rels', 'xl/charts/_rels/chart3.xml.rels', 'xl/charts/chart1.xml', 'xl/charts/chart2.xml', 'xl/charts/chart3.xml', 'xl/charts/colors1.xml', 'xl/charts/colors2.xml', 'xl/charts/style1.xml', 'xl/charts/style2.xml', 'xl/media/image1.png', 'xl/media/image2.png'],
  );
  same(await part(out, 'xl/charts/chart3.xml'), await part(bytes, 'xl/charts/chart1.xml'), 'the part of the copy is the part of the chart');
  same(await part(out, 'xl/charts/chart1.xml'), await part(bytes, 'xl/charts/chart1.xml'), 'the chart keeps its part');
  assert.equal(await text(out, 'xl/charts/style2.xml'), STYLE);
  assert.equal(await text(out, 'xl/charts/colors2.xml'), COLORS);
  same(await part(out, 'xl/media/image2.png'), await part(bytes, 'xl/media/image1.png'), 'the picture');
  const rels = await text(out, 'xl/charts/_rels/chart3.xml.rels');
  assert.match(rels, /Id="rId1"[^>]*Target="style2\.xml"/);
  assert.match(rels, /Id="rId2"[^>]*Target="colors2\.xml"/);
  assert.match(rels, /Id="rId3"[^>]*Target="\.\.\/media\/image2\.png"/);
  assert.deepEqual((await relationProblems(out, 'xl/charts/chart3.xml')).filter((line) => /names the relationship/.test(line)), []);
  const types = await text(out, '[Content_Types].xml');
  assert.match(types, /PartName="\/xl\/charts\/chart3\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.drawingml\.chart\+xml"/);
  assert.match(types, /PartName="\/xl\/charts\/style2\.xml" ContentType="application\/vnd\.ms-office\.chartstyle\+xml"/);
  assert.match(types, /PartName="\/xl\/charts\/colors2\.xml" ContentType="application\/vnd\.ms-office\.chartcolorstyle\+xml"/);
  // The sheet got its first drawing.
  assert.match(await text(out, 'xl/worksheets/sheet2.xml'), /<drawing r:id="rId\d+"\/>/);

  const again = await reopen(out);
  const read = chartsOf(again.sheets[1]);
  assert.equal(read.length, 1);
  assert.equal(read[0].known, 'refs');
  assert.deepEqual(bare(read[0].chart), bare(made.chart), 'read back as it is in the model');
  assert.deepEqual(bare(read[0].chart), bare(chartsOf(again.sheets[0])[0].chart));

  // The same save again gives the same file; the paste undone gives the file that was opened.
  same(await save(doc), out, 'a second save');
  doc.history.undo();
  same(await save(doc), bytes, 'copy, paste, undo, save');
  doc.history.redo();
  same(await save(doc), out, 'and done again');
});

test('chart copy: a reference that names no sheet gets the name of the sheet the chart came from, on another sheet only', async () => {
  const bytes = await richBook();
  const doc = await load(bytes);
  const copy = copyChart(doc.wb, 0, 1);
  apply(doc, pasteChart(doc.wb, 0, copy).ops);
  apply(doc, pasteChart(doc.wb, 1, copy).ops);
  assert.deepEqual(chartRefs(chartsOf(doc.wb.sheets[0])[2].chart), ['$B$1', '$A$2:$A$4', '$B$2:$B$4'], 'on its own sheet the copy is written as the chart is');
  assert.deepEqual(chartRefs(chartsOf(doc.wb.sheets[1])[0].chart), ['Sayfa1!$B$1', 'Sayfa1!$A$2:$A$4', 'Sayfa1!$B$2:$B$4']);
  const out = await save(doc);
  await whole(out, 'the package is whole');
  same(await part(out, 'xl/charts/chart3.xml'), await part(bytes, 'xl/charts/chart2.xml'), 'the copy on the same sheet');
  const moved = await text(out, 'xl/charts/chart4.xml');
  assert.match(moved, /<c:f>Sayfa1!\$B\$2:\$B\$4<\/c:f>/);
  assert.match(moved, /<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="112233"\/>/, 'what no model holds is in the copy');
  assert.match(moved, /<c:gapWidth val="77"\/>/);
  assert.equal(names(out).indexOf('xl/charts/_rels/chart4.xml.rels'), -1, 'a part without relationships gets none');
  const again = await reopen(out);
  assert.deepEqual(chartRefs(chartsOf(again.sheets[1])[0].chart), ['Sayfa1!$B$1', 'Sayfa1!$A$2:$A$4', 'Sayfa1!$B$2:$B$4']);

  // A name is no cells: it stays, and is counted.
  const named = await load(await richBook({ first: RICH.replace('<c:f>Sayfa1!$B$2:$B$4</c:f>', '<c:f>Kuzey</c:f>') }));
  const taken = pasteChart(named.wb, 1, copyChart(named.wb, 0, 0));
  assert.equal(taken.unbound, 1);
  assert.equal(pasteChart(named.wb, 0, copyChart(named.wb, 0, 0)).unbound, 0, 'on its own sheet nothing is to be bound');
});

test('chart copy: a copy follows rows, columns and sheets, is changed like the chart it was copied from, and is copied again', async () => {
  const bytes = await richBook();
  const doc = await load(bytes);
  assert.equal(doc.wb.chartsFollow, true);
  apply(doc, pasteChart(doc.wb, 1, copyChart(doc.wb, 0, 0)).ops);
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 2 }]);
  const refs = chartRefs(chartsOf(doc.wb.sheets[1])[0].chart);
  assert.deepEqual(refs, ['Sayfa1!$B$3', 'Sayfa1!$A$4:$A$6', 'Sayfa1!$B$4:$B$6', 'Sayfa1!$C$3', 'Sayfa1!$A$4:$A$6', 'Sayfa1!$C$4:$C$6']);
  apply(doc, renameSheet(doc.wb, 0, 'Satış'));
  assert.ok(chartRefs(chartsOf(doc.wb.sheets[1])[0].chart).every((ref) => /^'Satış'!/.test(ref)));

  // What a chart of the file takes, its copy takes; what it refuses, the copy refuses.
  apply(doc, updateChart(doc.wb, 1, 0, { title: 'Kopya ₺', legend: 'b', series: [{ from: 1 }, { from: 0, color: '#AA0000' }] }));
  assert.throws(() => updateChart(doc.wb, 1, 0, { type: 'pie' }), /Cannot change the type of this chart/);
  // A copy of the copy is made from the same part.
  const second = pasteChart(doc.wb, 1, copyChart(doc.wb, 1, 0));
  apply(doc, second.ops);
  assert.deepEqual(chartsOf(doc.wb.sheets[1])[1].source, { part: '/xl/charts/chart1.xml' });
  apply(doc, updateChart(doc.wb, 1, 1, { title: 'İkinci' }));

  const out = await save(doc);
  await whole(out, 'the package is whole');
  const first = await text(out, 'xl/charts/chart3.xml');
  assert.match(first, /<a:t>Kopya ₺<\/a:t>/);
  assert.match(first, /<a:rPr lang="tr-TR" sz="1400" b="1"\/>/, 'the title keeps its look');
  assert.match(first, /<c:f>'Satış'!\$C\$4:\$C\$6<\/c:f>/);
  assert.match(first, /c14:pivotOptions/);
  assert.match(first, /<mc:AlternateContent/);
  assert.match(await text(out, 'xl/charts/chart4.xml'), /<a:t>İkinci<\/a:t>/);
  assert.match(await text(out, 'xl/charts/chart1.xml'), /<a:t>Old<\/a:t>/, 'the chart keeps its title');
  assert.match(await text(out, 'xl/charts/chart1.xml'), /<c:f>'Satış'!\$B\$4:\$B\$6<\/c:f>/, 'and follows the rows and the name as well');
  assert.deepEqual(
    chartParts(out).filter((name) => /style|colors|image/.test(name)),
    ['xl/charts/colors1.xml', 'xl/charts/colors2.xml', 'xl/charts/colors3.xml', 'xl/charts/style1.xml', 'xl/charts/style2.xml', 'xl/charts/style3.xml', 'xl/media/image1.png', 'xl/media/image2.png', 'xl/media/image3.png'],
  );
  const again = await reopen(out);
  const read = chartsOf(again.sheets[1]);
  assert.deepEqual(read.map((chart) => bare(chart.chart)), chartsOf(doc.wb.sheets[1]).map((chart) => bare(chart.chart)));

  // The sheet the chart shows goes: the copy says that the cells are gone.
  apply(doc, deleteSheet(doc.wb, 0));
  assert.ok(chartRefs(chartsOf(doc.wb.sheets[0])[0].chart).every((ref) => ref.indexOf('#REF!') >= 0));
  const left = await save(doc);
  await whole(left, 'with the sheet gone the package is whole');
  assert.deepEqual(chartParts(left).filter((name) => /chart\d\.xml$/.test(name)), ['xl/charts/chart3.xml', 'xl/charts/chart4.xml'], 'the parts of the copies stay when those of the chart go');
  assert.equal(chartsOf((await reopen(left)).sheets[0]).length, 2);
  while (doc.history.undo());
  same(await save(doc), bytes, 'everything undone');
});

test('chart copy: a chart of the file that is cut and pasted is written from its part, which goes', async () => {
  const bytes = await richBook();
  for (const target of [0, 1]) {
    const doc = await load(bytes);
    const cut = cutChart(doc.wb, 0, 0);
    apply(doc, cut.ops);
    const gone = await save(doc);
    await whole(gone, 'cut: the package is whole');
    assert.equal(chartParts(gone).some((name) => /chart1|style|colors|image/.test(name)), false, 'the parts of the chart went with it');
    apply(doc, pasteChart(doc.wb, target, cut.copy).ops);
    const out = await save(doc);
    await whole(out, `pasted on sheet ${target}: the package is whole`);
    const made = chartParts(out).filter((name) => /charts\/chart\d+\.xml$/.test(name) && name !== 'xl/charts/chart2.xml');
    assert.equal(made.length, 1);
    same(await part(out, made[0]), await part(bytes, 'xl/charts/chart1.xml'), 'the part it had');
    assert.equal(chartParts(out).filter((name) => /style\d|colors\d|image\d/.test(name)).length, 3);
    const again = await reopen(out);
    assert.equal(chartsOf(again.sheets[target]).length, target === 0 ? 2 : 1);
    same(await save(doc), out, 'a second save');
    doc.history.undo();
    doc.history.undo();
    same(await save(doc), bytes, 'cut and paste undone');
  }
});

test('chart copy: shapes drawn over a chart do not go with its copy, and that is said', async () => {
  const bytes = await richBook({ shapes: true });
  await whole(bytes, 'the fixture is whole');
  const doc = await load(bytes);
  const chart = chartsOf(doc.wb.sheets[0])[0];
  assert.equal(chart.known, undefined);
  assert.deepEqual(chart.loose, ['userShapes']);
  assert.equal(sheetCharts(doc.wb, 0)[0].editable, 'none', 'the chart itself is moved and deleted, not changed');
  assert.equal(chartCopyProblem(doc.wb, 0, 0), undefined);
  const copy = copyChart(doc.wb, 0, 0);
  assert.deepEqual(copy.dropped, ['userShapes']);
  const pasted = pasteChart(doc.wb, 1, copy);
  assert.deepEqual(pasted.dropped, ['userShapes']);
  apply(doc, pasted.ops);
  assert.deepEqual(chartsOf(doc.wb.sheets[1])[0].source, { part: '/xl/charts/chart1.xml', drops: ['userShapes'] });
  // A copy of the copy leaves nothing more out.
  assert.deepEqual(copyChart(doc.wb, 1, 0).dropped, []);
  const out = await save(doc);
  await whole(out, 'the package is whole');
  same(await part(out, 'xl/charts/chart1.xml'), await part(bytes, 'xl/charts/chart1.xml'), 'the chart keeps its part');
  const made = await text(out, 'xl/charts/chart3.xml');
  assert.equal(made, SHAPED.replace('<c:userShapes r:id="rId4"/>', ''), 'the copy is the part without the element');
  const rels = await text(out, 'xl/charts/_rels/chart3.xml.rels');
  assert.equal(/rId4/.test(rels), false);
  assert.deepEqual((await relationProblems(out, 'xl/charts/chart3.xml')).filter((line) => /names the relationship/.test(line)), []);
  assert.deepEqual(names(out).filter((name) => /drawing7/.test(name)), ['xl/drawings/drawing7.xml'], 'the shapes are not copied');
  const again = await reopen(out);
  assert.equal(chartsOf(again.sheets[1])[0].known, 'refs');
  doc.history.undo();
  same(await save(doc), bytes, 'undone');
});

test('chart copy: a chart whose part holds references the model does not is not copied', async () => {
  const doc = await load(await richBook({ first: UNKNOWN }));
  assert.equal(chartsOf(doc.wb.sheets[0])[0].known, undefined);
  assert.equal(chartsOf(doc.wb.sheets[0])[0].loose, undefined);
  assert.match(chartCopyProblem(doc.wb, 0, 0) || '', /Cannot copy this chart: its part holds references the model does not/);
  assert.throws(() => copyChart(doc.wb, 0, 0), /Cannot copy this chart/);
  assert.throws(() => cutChart(doc.wb, 0, 0), /Cannot copy this chart/);
  assert.throws(() => duplicateChart(doc.wb, 0, 0), /Cannot copy this chart/);
  assert.equal(chartsOf(doc.wb.sheets[0]).length, 2, 'nothing was taken');
  // The other chart of the sheet is copied all the same.
  apply(doc, duplicateChart(doc.wb, 0, 1).ops);
  await whole(await save(doc), 'the package is whole');
});

test('chart copy: what a part says of itself decides what a copy leaves out', () => {
  const facts = (xml: string) => {
    const root = parseXml(xml).root;
    const chart = parseChart(root, OFFICE_THEME);
    assert.ok(chart);
    return { known: chartKnown(root, chart, true), loose: chartLoose(root, chart) };
  };
  assert.deepEqual(facts(RICH), { known: { known: 'refs' }, loose: undefined });
  assert.deepEqual(facts(SHAPED), { known: undefined, loose: { loose: ['userShapes'] } });
  assert.deepEqual(facts(SHAPED.replace('<c:userShapes r:id="rId4"/>', '<c:externalData r:id="rId5"><c:autoUpdate val="0"/></c:externalData><c:userShapes r:id="rId4"/>')).loose, { loose: ['externalData', 'userShapes'] });
  assert.deepEqual(facts(UNKNOWN), { known: undefined, loose: undefined });
  // Shapes anywhere but at the root, and a pivot chart, are not something a copy can do without.
  assert.equal(facts(SHAPED.replace('<c:userShapes r:id="rId4"/>', '<c:pivotSource><c:name>[a.xlsx]S!P</c:name><c:fmtId val="0"/></c:pivotSource><c:userShapes r:id="rId4"/>')).loose, undefined);
  assert.equal(facts(UNKNOWN.replace('<c:extLst><c:ext uri="{781A3756', '<c:userShapes r:id="rId4"/><c:extLst><c:ext uri="{781A3756')).loose, undefined);
});

test('chart copy: the copy of a part that declares another encoding is written as UTF-8 that says so', async () => {
  for (const encoding of ['ISO-8859-1', 'UTF-16LE']) {
    const title = encoding === 'UTF-16LE' ? 'Yıllık Şube İğü' : LATIN_TITLE;
    const xml = RICH.replace('>Old<', `>${title}<`).replace(/^<\?xml[^?]*\?>/, '<?xml version="1.0" encoding="UTF-16"?>');
    let first: Uint8Array;
    if (encoding === 'UTF-16LE') {
      first = new Uint8Array(2 + xml.length * 2);
      first[0] = 0xff;
      first[1] = 0xfe;
      for (let i = 0; i < xml.length; i++) {
        first[2 + i * 2] = xml.charCodeAt(i) & 0xff;
        first[3 + i * 2] = xml.charCodeAt(i) >> 8;
      }
    } else first = latinChart();
    const bytes = await richBook({ first });
    const doc = await load(bytes);
    assert.equal(chartsOf(doc.wb.sheets[0])[0].known, 'refs', encoding);
    assert.equal(sheetCharts(doc.wb, 0)[0].title, title, `${encoding}: the letters are read`);
    apply(doc, pasteChart(doc.wb, 1, copyChart(doc.wb, 0, 0)).ops);
    // A copy that is left as it is has the letters of the chart.
    apply(doc, duplicateChart(doc.wb, 1, 0).ops);
    apply(doc, updateChart(doc.wb, 1, 0, { title: 'Şırnak ₺ İki' }));
    const out = await save(doc);
    await whole(out, `${encoding}: the package is whole`);
    same(await part(out, 'xl/charts/chart1.xml'), first, `${encoding}: the chart keeps its part`);
    const made = await part(out, 'xl/charts/chart3.xml');
    assert.equal(encodingProblem(made), undefined, encoding);
    const left = new TextDecoder('utf-8', { fatal: true }).decode(await part(out, 'xl/charts/chart4.xml'));
    assert.match(left, /^<\?xml version="1\.0" encoding="UTF-8"/, `${encoding}: the copy that was left as it is`);
    assert.ok(left.indexOf(`<a:t>${title}</a:t>`) > 0, `${encoding}: its title has every letter`);
    assert.equal(sheetCharts(await reopen(out), 1)[1].title, title, encoding);
    const said = new TextDecoder().decode(made);
    assert.match(said, /^<\?xml version="1\.0" encoding="UTF-8"/, `${encoding}: the declaration says what the bytes are`);
    assert.match(said, /<a:t>Şırnak ₺ İki<\/a:t>/);
    assert.match(said, /c14:pivotOptions/);
    const again = await reopen(out);
    assert.deepEqual(bare(chartsOf(again.sheets[1])[0].chart), bare(chartsOf(doc.wb.sheets[1])[0].chart), encoding);
  }
});

test('chart copy: charts the engine wrote and charts that were added are copied from their models', async () => {
  const start = await load(await richBook());
  apply(start, addChart(start.wb, 1, { type: 'pie', range: { top: 0, left: 0, bottom: 3, right: 1 }, title: 'Own' }));
  const own = await save(start);
  const doc = await load(own);
  assert.equal(chartsOf(doc.wb.sheets[1])[0].known, 'all');
  apply(doc, pasteChart(doc.wb, 0, copyChart(doc.wb, 1, 0)).ops);
  const made = chartsOf(doc.wb.sheets[0])[2];
  assert.equal(made.source, undefined);
  assert.equal(made.known, undefined);
  assert.equal(sheetCharts(doc.wb, 0)[2].editable, 'all');
  apply(doc, updateChart(doc.wb, 0, 2, { type: 'column' }));
  apply(doc, duplicateChart(doc.wb, 0, 2).ops);
  apply(doc, moveChart(doc.wb, 0, 3, { top: 40, left: 1, bottom: 50, right: 6 }));
  const out = await save(doc);
  await whole(out, 'the package is whole');
  const again = await reopen(out);
  assert.deepEqual(chartsOf(again.sheets[0]).map((chart) => chart.known), ['refs', 'refs', 'all', 'all']);
  assert.deepEqual(bare(chartsOf(again.sheets[0])[2].chart), bare(chartsOf(doc.wb.sheets[0])[2].chart));
  // Deleted again, the copies leave nothing behind.
  apply(doc, removeChart(doc.wb, 0, 3));
  apply(doc, removeChart(doc.wb, 0, 2));
  same(await save(doc), own, 'copies deleted');
});

test('chart copy: what lies beyond what a copy takes along is named by it, and stays while the copy names it', async () => {
  const base = await richBook();
  const zip = ZipArchive.read(base);
  let rels = (await text(base, 'xl/charts/_rels/chart1.xml.rels')).replace('</Relationships>', '');
  for (let n = 10; n < 90; n++) {
    zip.set(`xl/media/extra${n}.png`, new Uint8Array([0x89, 0x50, 0x4e, 0x47, n]));
    rels += `<Relationship Id="rId${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/extra${n}.png"/>`;
  }
  zip.set('xl/charts/_rels/chart1.xml.rels', new TextEncoder().encode(`${rels}</Relationships>`));
  const bytes = await zip.write();
  await whole(bytes, 'the fixture is whole');
  const doc = await load(bytes);
  apply(doc, pasteChart(doc.wb, 1, copyChart(doc.wb, 0, 0)).ops);
  const out = await save(doc);
  await whole(out, 'the package is whole');
  const extras = (file: Uint8Array) => names(file).filter((name) => /media\/extra/.test(name)).length;
  // The chart part and 63 parts beside it are the copy's own; the other twenty it shares with the chart.
  assert.equal(extras(out), 80 + 60);
  assert.equal((await text(out, 'xl/charts/_rels/chart3.xml.rels')).match(/<Relationship /g)?.length, 83);
  // The chart goes: what the copy names stays, and what only the chart led to goes.
  apply(doc, removeChart(doc.wb, 0, 0));
  const left = await save(doc);
  await whole(left, 'with the chart gone the package is whole');
  assert.equal(extras(left), 60 + 20);
  assert.equal(chartsOf((await reopen(left)).sheets[1]).length, 1);
  while (doc.history.undo());
  same(await save(doc), bytes, 'everything undone');
});

test('chart copy: relationships that cannot be read are written for the copy as they are', async () => {
  const zip = ZipArchive.read(await richBook());
  const broken = new TextEncoder().encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1"');
  zip.set('xl/charts/_rels/chart1.xml.rels', broken);
  const doc = await load(await zip.write());
  assert.equal(chartsOf(doc.wb.sheets[0])[0].known, 'refs');
  apply(doc, pasteChart(doc.wb, 1, copyChart(doc.wb, 0, 0)).ops);
  const out = await save(doc);
  same(await part(out, 'xl/charts/_rels/chart3.xml.rels'), broken, 'the relationships of the copy');
  assert.equal(chartsOf((await reopen(out)).sheets[1]).length, 1);
});

test('chart copy: a sheet whose drawing takes no change takes no copy, and a workbook takes no more charts than it reads', async () => {
  const doc = await load(await richBook());
  const copy = copyChart(doc.wb, 0, 0);
  doc.wb.sheets[1].drawingsFixed = 'it is nested too deeply';
  assert.throws(() => pasteChart(doc.wb, 1, copy), /Cannot change the charts of this sheet/);
  delete doc.wb.sheets[1].drawingsFixed;
  const wb = bookOf(SALES);
  const history = new History(wb);
  history.apply(addChart(wb, 0, { type: 'column', range: ALL }), undefined, undefined);
  const one = (wb.sheets[0].drawings || [])[0];
  wb.sheets[0].drawings = new Array(1000).fill(one);
  assert.throws(() => duplicateChart(wb, 0, 0), /A sheet has at most 1000 charts/);
});

test('chart copy: files with copies are read by a reader that is not this engine', async (t) => {
  if (!pythonHere()) {
    t.skip('python3 is not available');
    return;
  }
  const files: { name: string; bytes: Uint8Array; charts: number[] }[] = [];
  const bytes = await richBook();
  files.push({ name: 'the fixture', bytes, charts: [2, 0] });
  {
    const doc = await load(bytes);
    apply(doc, pasteChart(doc.wb, 1, copyChart(doc.wb, 0, 0)).ops);
    files.push({ name: 'pasted on another sheet', bytes: await save(doc), charts: [2, 1] });
    apply(doc, pasteChart(doc.wb, 1, copyChart(doc.wb, 0, 1)).ops);
    apply(doc, duplicateChart(doc.wb, 0, 0).ops);
    apply(doc, duplicateChart(doc.wb, 1, 0).ops);
    files.push({ name: 'more copies', bytes: await save(doc), charts: [3, 3] });
    apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 1, count: 1 }]);
    apply(doc, renameSheet(doc.wb, 0, "It's & <1>"));
    apply(doc, updateChart(doc.wb, 1, 0, { title: 'Başlık', series: [{ from: 1 }, { from: 0 }, { values: 'Other!$B$2:$B$4' }] }));
    files.push({ name: 'rows, a name and a change', bytes: await save(doc), charts: [3, 3] });
    apply(doc, removeChart(doc.wb, 0, 0));
    files.push({ name: 'the chart deleted', bytes: await save(doc), charts: [2, 3] });
    apply(doc, deleteSheet(doc.wb, 0));
    files.push({ name: 'its sheet deleted', bytes: await save(doc), charts: [3] });
  }
  {
    const doc = await load(bytes);
    const cut = cutChart(doc.wb, 0, 0);
    apply(doc, cut.ops);
    files.push({ name: 'cut', bytes: await save(doc), charts: [1, 0] });
    apply(doc, pasteChart(doc.wb, 1, cut.copy, { row: 5, col: 5 }).ops);
    files.push({ name: 'cut and pasted', bytes: await save(doc), charts: [1, 1] });
  }
  {
    const doc = await load(await richBook({ shapes: true }));
    apply(doc, pasteChart(doc.wb, 1, copyChart(doc.wb, 0, 0)).ops);
    apply(doc, duplicateChart(doc.wb, 0, 0).ops);
    files.push({ name: 'shapes left out', bytes: await save(doc), charts: [3, 1] });
  }
  {
    const doc = await load(await richBook({ first: latinChart() }));
    apply(doc, pasteChart(doc.wb, 1, copyChart(doc.wb, 0, 0)).ops);
    apply(doc, duplicateChart(doc.wb, 0, 0).ops);
    apply(doc, updateChart(doc.wb, 1, 0, { title: 'Şırnak ₺' }));
    files.push({ name: 'another encoding', bytes: await save(doc), charts: [3, 1] });
  }
  const found = readByPython(files, 'unit');
  assert.deepEqual(found.problems, []);
  assert.equal(found.files, files.length);
  t.diagnostic(`${found.files} files read by Python${found.openpyxl ? ' and openpyxl' : ' (openpyxl is not installed: charts were not counted by it)'}`);
});
