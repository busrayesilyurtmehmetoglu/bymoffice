// Charts of the file that are changed: what the writer rewrites in the part of a chart of
// another application and what it leaves alone, what is refused, the references that follow
// rows, columns and sheets, and the workbooks whose charts cannot follow.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { type SheetDrawing } from '../src/index';
import { chartRefs } from '../src/model/sheet/chart';
import { addChart, removeChart, sheetCharts, updateChart } from '../src/ops/charts';
import { applyOp } from '../src/ops/ops';
import { FOREIGN, bare, chartsOf, load, names, packageProblems, part, reopen, salesBook, same, save, text } from './chart-edit-fixtures';
import { catAx, chartFrameOf, chartXml, group, numRef, ser, strRef, valAx } from './chart-fixtures';
import { twoCell } from './sheet-drawing-fixtures';

const frame = (rel: string, name: string) => twoCell([4, 1], [9, 12], chartFrameOf(rel, name));
const foreignBook = (more: Parameters<typeof salesBook>[0] = {}) => salesBook({ drawing: frame('rIdC1', 'Grafik 1'), charts: [FOREIGN], ...more });
const CHART = 'xl/charts/chart1.xml';

/** What of FOREIGN no model holds: it has to be in the part whatever is changed. */
const KEPT = [
  '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart" Requires="c14"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent>',
  '<c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}"><c16:uniqueId xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart" val="{00000000-AAAA}"/></c:ext></c:extLst>',
  '<c:gapWidth val="219"/>',
  '<c:txPr><a:bodyPr rot="-2700000"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"/></a:pPr><a:endParaRPr lang="tr-TR"/></a:p></c:txPr>',
  '<c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings>',
  '<c:extLst><c:ext uri="{781A3756-C4B2-4CAC-9D66-4F8BD8637D16}"><c14:pivotOptions xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"/></c:ext></c:extLst>',
  '<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln><a:effectLst/>',
];

function keeps(xml: string, label: string, except: string[] = []): void {
  for (const piece of KEPT) if (except.indexOf(piece) < 0) assert.ok(xml.indexOf(piece) >= 0, `${label}: the part still holds ${piece.slice(0, 60)}...`);
}

test('chart file: a chart nobody touched is written back byte for byte, whatever its cells do', async () => {
  const bytes = await foreignBook();
  const doc = await load(bytes);
  doc.history.apply([{ type: 'setCells', sheet: 0, cells: [{ row: 1, col: 1, cell: { kind: 'number', value: 77, style: 0 } }] }], undefined, undefined);
  const saved = await save(doc);
  same(await part(saved, CHART), await part(bytes, CHART), 'the chart part');
  same(await part(saved, 'xl/drawings/drawing1.xml'), await part(bytes, 'xl/drawings/drawing1.xml'), 'the drawing part');
});

test('chart file: a title that is changed keeps the look of its first run, and everything the model does not hold stays', async () => {
  const bytes = await foreignBook();
  const doc = await load(bytes);
  doc.history.apply(updateChart(doc.wb, 0, 0, { title: 'New' }), undefined, undefined);
  const saved = await save(doc);
  assert.deepEqual(await packageProblems(saved), []);
  const xml = await text(saved, CHART);
  keeps(xml, 'title');
  assert.ok(
    xml.indexOf('<c:title><c:tx><c:rich><a:bodyPr rot="0"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="1"/></a:pPr><a:r><a:rPr lang="tr-TR" sz="1400" b="1"/><a:t>New</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/><c:txPr>') >= 0,
    'the text is put into the run the title had',
  );
  // A part that is written again caches the cells as they are.
  assert.ok(xml.indexOf('<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>30</c:v></c:pt></c:numCache>') >= 0, 'the cache holds the cells');
  assert.ok(xml.indexOf('<c:v>9</c:v>') < 0, 'and nothing of what the file cached');
  const again = await reopen(saved);
  const read = chartsOf(again.sheets[0])[0];
  assert.equal(read.chart.title?.text, 'New');
  assert.equal(read.known, 'refs');
  assert.deepEqual(bare(read.chart), bare(chartsOf(doc.wb.sheets[0])[0].chart), 'the chart that is read is the model that was written');
  // Two lines: a paragraph each, made like the first. No title: the element goes, and the chart says that none is made up.
  doc.history.apply(updateChart(doc.wb, 0, 0, { title: 'One\nTwo' }), undefined, undefined);
  const lines = await text(await save(doc), CHART);
  assert.equal(lines.split('<a:rPr lang="tr-TR" sz="1400" b="1"/>').length, 3, 'both lines have the look of the run');
  doc.history.apply(updateChart(doc.wb, 0, 0, { title: null }), undefined, undefined);
  const none = await text(await save(doc), CHART);
  assert.ok(none.indexOf('<c:title>') < 0 && none.indexOf('<c:autoTitleDeleted val="1"/>') >= 0, 'no title');
  keeps(none, 'no title');
  doc.history.undo();
  doc.history.undo();
  doc.history.undo();
  same(await save(doc), bytes, 'everything undone');
});

test('chart file: the series of a chart of the file are recoloured, renamed, reordered, taken away and added in the part', async () => {
  const bytes = await foreignBook();
  const doc = await load(bytes);
  doc.history.apply(updateChart(doc.wb, 0, 0, { series: [{ from: 1, color: '#AA0000' }, { from: 0, name: 'Kuzey' }, { values: 'Sayfa1!$B$2:$B$4', name: 'Third' }] }), undefined, undefined);
  const saved = await save(doc);
  const xml = await text(saved, CHART);
  keeps(xml, 'series');
  // The gradient gives way to the colour; the outline and the effects of the series stay.
  assert.ok(xml.indexOf('<c:spPr><a:solidFill><a:srgbClr val="AA0000"/></a:solidFill><a:ln w="12700">') >= 0, 'the fill of the series');
  assert.ok(xml.indexOf('gradFill') < 0, 'the gradient is gone');
  assert.ok(xml.indexOf('<c:idx val="0"/><c:order val="1"/><c:tx><c:v>Kuzey</c:v></c:tx>') >= 0, 'the first series of the file is the second now, under its new name');
  assert.ok(xml.indexOf('<c:idx val="1"/><c:order val="0"/>') >= 0, 'the second is the first');
  assert.ok(xml.indexOf('<c:idx val="2"/><c:order val="2"/><c:tx><c:v>Third</c:v></c:tx>') >= 0, 'the new series');
  const again = chartsOf((await reopen(saved)).sheets[0])[0];
  assert.deepEqual(again.chart.groups[0].series.map((item) => [item.index, item.name?.text, item.values?.ref]), [
    [1, 'South', 'Sayfa1!$C$2:$C$4'],
    [0, 'Kuzey', 'Sayfa1!$B$2:$B$4'],
    [2, 'Third', 'Sayfa1!$B$2:$B$4'],
  ]);
  assert.equal(again.chart.groups[0].series[0].shape?.fill, '#AA0000');
  assert.equal(again.known, 'refs');
  // One series is taken away: its element goes, the other stays as it is.
  doc.history.apply(updateChart(doc.wb, 0, 0, { series: [{ from: 1 }] }), undefined, undefined);
  const one = await text(await save(doc), CHART);
  assert.equal(one.split('<c:ser>').length, 2, 'one series is left');
  assert.ok(one.indexOf(KEPT[1]) >= 0, 'the one with its extension');
  doc.history.undo();
  doc.history.undo();
  same(await save(doc), bytes, 'everything undone');
});

test('chart file: legend, axes, labels and grouping are changed in place', async () => {
  const bytes = await foreignBook();
  const doc = await load(bytes);
  doc.history.apply(updateChart(doc.wb, 0, 0, { legend: 'b', axisTitles: { category: 'Ay', value: 'Adet' }, axisBounds: { min: 0, max: 40 }, dataLabels: true, grouping: 'stacked' }), undefined, undefined);
  const saved = await save(doc);
  const xml = await text(saved, CHART);
  keeps(xml, 'legend and axes');
  assert.ok(xml.indexOf('<c:legend><c:legendPos val="b"/><c:overlay val="0"/><c:txPr>') >= 0, 'the legend keeps its text properties');
  assert.ok(xml.indexOf('<c:grouping val="stacked"/>') >= 0 && xml.indexOf('<c:gapWidth val="219"/><c:overlap val="100"/><c:axId') >= 0, 'stacked bars overlap, stated where the schema has it');
  assert.ok(xml.indexOf('<c:scaling><c:orientation val="minMax"/><c:max val="40"/><c:min val="0"/></c:scaling>') >= 0, 'the bounds of the value axis');
  assert.ok(/<c:axPos val="b"\/><c:title><c:tx><c:rich>.*?<a:t>Ay<\/a:t>.*?<\/c:title><c:txPr>/.test(xml), 'the title of the category axis stands before its text properties');
  assert.ok(/<c:axPos val="l"\/><c:majorGridlines\/><c:title>.*?<a:t>Adet<\/a:t>/.test(xml), 'the title of the value axis stands behind its gridlines');
  assert.ok(xml.indexOf('<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/>') >= 0, 'the labels of the values');
  const read = sheetCharts(await reopen(saved), 0)[0];
  assert.deepEqual([read.legend, read.axisTitles, read.axisBounds, read.dataLabels, read.grouping], ['b', { category: 'Ay', value: 'Adet' }, { min: 0, max: 40 }, true, 'stacked']);
  doc.history.apply(updateChart(doc.wb, 0, 0, { legend: 'none', axisTitles: { category: null }, axisBounds: { max: null }, dataLabels: false }), undefined, undefined);
  const less = await text(await save(doc), CHART);
  assert.ok(less.indexOf('<c:legend>') < 0 && less.indexOf('<c:dLbls>') < 0 && less.indexOf('<a:t>Ay</a:t>') < 0 && less.indexOf('<c:max') < 0, 'what was taken away is gone');
  assert.ok(less.indexOf('<c:min val="0"/>') >= 0 && less.indexOf('<a:t>Adet</a:t>') >= 0, 'what was not stays');
  doc.history.undo();
  doc.history.undo();
  same(await save(doc), bytes, 'everything undone');
});

test('chart file: another type is refused for a chart whose part holds more than the model, and taken by one the engine wrote', async () => {
  const doc = await load(await foreignBook());
  assert.equal(sheetCharts(doc.wb, 0)[0].editable, 'parts');
  assert.throws(() => updateChart(doc.wb, 0, 0, { type: 'line' }), /Cannot change the type of this chart/);
  // A chart the engine wrote: saved, read again, changed in every way.
  const own = await load(await salesBook());
  own.history.apply(addChart(own.wb, 0, { type: 'column', range: { top: 0, left: 0, bottom: 3, right: 2 }, title: 'T' }), undefined, undefined);
  const again = await load(await save(own));
  assert.equal(sheetCharts(again.wb, 0)[0].editable, 'all');
  again.history.apply(updateChart(again.wb, 0, 0, { type: 'pie', title: 'Pie' }), undefined, undefined);
  const saved = await save(again);
  assert.deepEqual(await packageProblems(saved), []);
  const read = chartsOf((await reopen(saved)).sheets[0])[0];
  assert.equal(read.chart.groups[0].kind, 'pie');
  assert.equal(read.chart.title?.text, 'Pie');
  assert.equal(read.known, 'all');
  assert.deepEqual(bare(read.chart), bare(chartsOf(again.wb.sheets[0])[0].chart));
  again.history.undo();
  same(await save(again), again.bytes, 'undone: the file as it was opened');
});

test('chart file: rows and columns that move take the references of the charts of the file along, in the part too', async () => {
  const bytes = await foreignBook({ sheets: [{ name: 'Other', xml: '<sheetData/>' }] });
  const doc = await load(bytes);
  assert.equal(doc.wb.chartsFollow, true);
  doc.history.apply([{ type: 'insert', sheet: 0, axis: 'row', at: 2, count: 1 }], undefined, undefined);
  doc.history.apply([{ type: 'setCells', sheet: 0, cells: [{ row: 2, col: 1, cell: { kind: 'number', value: 15, style: 0 } }] }], undefined, undefined);
  assert.deepEqual(chartRefs(chartsOf(doc.wb.sheets[0])[0].chart), ['Sayfa1!$B$1', 'Sayfa1!$A$2:$A$5', 'Sayfa1!$B$2:$B$5', 'Sayfa1!$C$1', 'Sayfa1!$A$2:$A$5', 'Sayfa1!$C$2:$C$5']);
  const saved = await save(doc);
  assert.deepEqual(await packageProblems(saved), []);
  const xml = await text(saved, CHART);
  keeps(xml, 'a row inserted');
  assert.ok(xml.indexOf('<c:f>Sayfa1!$B$2:$B$5</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="4"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>15</c:v></c:pt><c:pt idx="2"><c:v>20</c:v></c:pt><c:pt idx="3"><c:v>30</c:v></c:pt></c:numCache>') >= 0, 'the range grew and the cache holds the new cell');
  // Nothing but references and caches differs from the part of the file.
  const strip = (s: string) =>
    s
      .replace(/\?>\s*</, '?><')
      .replace(/<c:f>[^<]*<\/c:f>/g, '<c:f/>')
      .replace(/<c:(num|str)Cache>.*?<\/c:\1Cache>/g, '<c:$1Cache/>');
  assert.ok(strip(xml) === strip(await text(bytes, CHART)), 'everything else of the part is as the file has it');
  // The frame moved down by a row, as the writer moves every drawing.
  const again = await reopen(saved);
  assert.deepEqual((again.sheets[0].drawings as SheetDrawing[])[0].anchor.to, { col: 9, colOff: 0, row: 13, rowOff: 0 });
  assert.deepEqual(chartRefs(chartsOf(again.sheets[0])[0].chart), chartRefs(chartsOf(doc.wb.sheets[0])[0].chart));

  // The column of the second series goes: "#REF!", as Excel writes it.
  doc.history.apply([{ type: 'delete', sheet: 0, axis: 'col', at: 2, count: 1 }], undefined, undefined);
  const gone = await text(await save(doc), CHART);
  assert.ok(gone.indexOf('<c:f>Sayfa1!#REF!</c:f>') >= 0, 'the reference is an error');
  // Rows of another sheet change no chart.
  const before = doc.wb.sheets[0].drawings;
  doc.history.apply([{ type: 'insert', sheet: 1, axis: 'row', at: 0, count: 2 }], undefined, undefined);
  assert.equal(doc.wb.sheets[0].drawings, before);
  for (let i = 0; i < 4; i++) doc.history.undo();
  const undone = await save(doc);
  same(await part(undone, CHART), await part(bytes, CHART), 'everything undone: the chart part');
  same(await part(undone, 'xl/drawings/drawing1.xml'), await part(bytes, 'xl/drawings/drawing1.xml'), 'everything undone: the drawing part');
});

test('chart file: a sheet that is renamed or deleted is renamed, or an error, in the parts of the charts', async () => {
  const bytes = await foreignBook({ sheets: [{ name: 'Other', xml: '<sheetData/>' }] });
  const doc = await load(bytes);
  doc.history.apply(addChart(doc.wb, 1, { type: 'line', series: [{ values: 'Sayfa1!$B$2:$B$4' }] }), undefined, undefined);
  doc.history.apply([{ type: 'renameSheet', sheet: 0, name: 'Satış 2024' }], undefined, undefined);
  const saved = await save(doc);
  assert.deepEqual(await packageProblems(saved), []);
  const xml = await text(saved, CHART);
  keeps(xml, 'renamed');
  assert.ok(xml.indexOf("<c:f>'Satış 2024'!$B$2:$B$4</c:f>") >= 0, 'the chart of the file names the sheet as it is called now');
  assert.ok((await text(saved, 'xl/charts/chart2.xml')).indexOf("<c:f>'Satış 2024'!$B$2:$B$4</c:f>") >= 0, 'and so does the chart that was added');
  const again = await reopen(saved);
  assert.deepEqual(chartsOf(again.sheets[1])[0].chart.groups[0].series[0].values?.cache.values, [10, 20, 30], 'the reference is followed in the file that was written');
  doc.history.apply([{ type: 'deleteSheet', sheet: 0 }], undefined, undefined);
  const without = await save(doc);
  assert.deepEqual(await packageProblems(without), []);
  assert.deepEqual(names(without).filter((name) => /charts\//.test(name)), ['xl/charts/chart2.xml'], 'the chart of the deleted sheet went with it');
  assert.ok((await text(without, 'xl/charts/chart2.xml')).indexOf('<c:f>#REF!</c:f>') >= 0, 'the chart of the other sheet shows an error');
  doc.history.undo();
  doc.history.undo();
  doc.history.undo();
  same(await save(doc), bytes, 'everything undone');
});

test('chart file: a workbook with a chart whose references the model does not hold keeps its rows, columns and sheets', async () => {
  // A reference in an extension (the range of data labels): no model holds it.
  const hidden = FOREIGN.replace('<c16:uniqueId', '<c15:datalabelsRange xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"><c15:f>Sayfa1!$C$2:$C$4</c15:f></c15:datalabelsRange><c16:uniqueId');
  const doc = await load(await foreignBook({ charts: [hidden], sheets: [{ name: 'Other', xml: '<sheetData/>' }] }));
  assert.equal(chartsOf(doc.wb.sheets[0])[0].known, undefined);
  assert.equal(doc.wb.chartsFollow, undefined);
  assert.equal(sheetCharts(doc.wb, 0)[0].editable, 'none');
  assert.throws(() => applyOp(doc.wb, { type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 }), /sheet with chart/);
  assert.throws(() => applyOp(doc.wb, { type: 'insert', sheet: 1, axis: 'col', at: 0, count: 1 }), /sheet with chart/);
  assert.throws(() => applyOp(doc.wb, { type: 'renameSheet', sheet: 0, name: 'X' }), /workbook with chart/);
  assert.throws(() => updateChart(doc.wb, 0, 0, { title: 'x' }), /only moved and deleted/);
  // Deleted it can be: with it gone and the file saved, the workbook is one without charts.
  doc.history.apply(removeChart(doc.wb, 0, 0), undefined, undefined);
  const saved = await save(doc);
  assert.deepEqual(await packageProblems(saved), []);
  const again = await reopen(saved);
  applyOp(again, { type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 });

  // A chart on a chart sheet, a pivot chart, a chart of a type that is not read: none follows.
  const line = chartXml({ plot: group('lineChart', '<c:grouping val="standard"/>', ser(0, { cat: strRef('Sayfa1!$A$2:$A$4', ['a']), val: numRef('Sayfa1!$B$2:$B$4', [1]) })) + catAx() + valAx() });
  for (const [label, spec] of [
    ['a chart sheet', { chartSheet: line }],
    ['a pivot chart', { drawing: frame('rIdC1', 'P'), charts: [line.replace('<c:chart>', '<c:pivotSource><c:name>[b.xlsx]S!Pivot</c:name><c:fmtId val="0"/></c:pivotSource><c:chart>')] }],
    ['a surface chart', { drawing: frame('rIdC1', 'S'), charts: [line.replace(/lineChart/g, 'surfaceChart')] }],
    ['two frames for one chart', { drawing: frame('rIdC1', 'A') + frame('rIdC1', 'B'), charts: [line] }],
  ] as const) {
    const wb = await reopen(await salesBook(spec));
    assert.equal(wb.chartsFollow, undefined, label);
    assert.throws(() => applyOp(wb, { type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 }), /sheet with chart/, label);
  }
  // The plain line chart on a worksheet does follow.
  assert.equal((await reopen(await salesBook({ drawing: frame('rIdC1', 'L'), charts: [line] }))).chartsFollow, true);
});
