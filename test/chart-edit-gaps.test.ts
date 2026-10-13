// The gates and rules of chart editing, one by one: what keeps the charts of a workbook from
// following its rows, columns and sheets; what the writer leaves and what it refuses; where a
// chart stands that was added or moved after rows were inserted; the limits; and what a chart
// that follows edits draws.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { ZipArchive, type Chart, type DrawingChart, type SheetDrawing, type Workbook } from '../src/index';
import { chartValues } from '../src/layout/chart';
import { chartRefs, MAX_CHART_SERIES, MAX_SHEET_CHARTS, MAX_WORKBOOK_CHARTS } from '../src/model/sheet/chart';
import { addChart, chartProblem, moveChart, removeChart, seriesOfRange, sheetCharts, updateChart, MAX_SHEET_DRAWINGS } from '../src/ops/charts';
import { applyOp, History } from '../src/ops/ops';
import { FOREIGN, bare, bookOf, chartsOf, load, packageProblems, reopen, salesBook, save, text } from './chart-edit-fixtures';
import { catAx, chartFrameOf, chartXml, group, numRef, ser, sheetData, strRef, valAx, xyAxes, xySer } from './chart-fixtures';
import { twoCell } from './sheet-drawing-fixtures';

const ALL = { top: 0, left: 0, bottom: 3, right: 2 };
const CHART = 'xl/charts/chart1.xml';
const frame = (rel = 'rIdC1', name = 'Grafik 1', row = 1, editAs?: string) => twoCell([4, row], [9, row + 11], chartFrameOf(rel, name), editAs);
const BAR = '<c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>';
const S = (extra: Partial<Parameters<typeof ser>[1]> = {}) => ser(0, { tx: strRef('Sayfa1!$B$1', ['N']), cat: strRef('Sayfa1!$A$2:$A$4', ['a', 'b', 'c']), val: numRef('Sayfa1!$B$2:$B$4', [1, 2, 3]), ...extra });
const plain = (series = S(), more: { before?: string; after?: string; head?: string } = {}) => chartXml({ ...more, plot: group('barChart', BAR, series) + catAx() + valAx() });
const OTHER = [{ name: 'Other', xml: sheetData([['x', 1], ['y', 2]]) }];
const encoder = new TextEncoder();
const drawingOf = (op: ReturnType<typeof addChart>[number]): SheetDrawing => ((op as { drawings: SheetDrawing[] }).drawings.slice(-1)[0]);
const modelOf = (op: ReturnType<typeof addChart>[number]): Chart => (drawingOf(op).object as DrawingChart).chart;

async function facts(chart: string, spec: Parameters<typeof salesBook>[0] = {}, change?: (zip: ZipArchive, bytes: Uint8Array) => Promise<void>): Promise<{ wb: Workbook; known: string | undefined; follow: boolean | undefined }> {
  let bytes = await salesBook({ drawing: frame(), charts: [chart], sheets: OTHER, ...spec });
  if (change) {
    const zip = ZipArchive.read(bytes);
    await change(zip, bytes);
    bytes = await zip.write();
  }
  const wb = await reopen(bytes);
  const chartOf = chartsOf(wb.sheets[0])[0];
  return { wb, known: chartOf ? chartOf.known : undefined, follow: wb.chartsFollow };
}

function stays(wb: Workbook, label: string): void {
  assert.throws(() => applyOp(wb, { type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 }), /sheet with chart/, `${label}: rows of the sheet of the chart`);
  assert.throws(() => applyOp(wb, { type: 'delete', sheet: 1, axis: 'col', at: 0, count: 1 }), /sheet with chart/, `${label}: columns of another sheet`);
  assert.throws(() => applyOp(wb, { type: 'renameSheet', sheet: 1, name: 'X' }), /workbook with chart/, `${label}: the name of a sheet`);
  assert.throws(() => applyOp(wb, { type: 'deleteSheet', sheet: 1 }), /workbook with chart/, `${label}: a sheet`);
}

test('chart gates: a plain chart of another application follows; each thing no model holds a reference for keeps the workbook as it is', async () => {
  const ok = await facts(plain());
  assert.deepEqual([ok.known, ok.follow], ['refs', true]);
  applyOp(ok.wb, { type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 });
  const gates: [string, string][] = [
    ['amounts of error bars read from cells', plain(S({ props: `<c:errBars><c:errBarType val="both"/><c:errValType val="cust"/><c:plus>${numRef('Sayfa1!$C$2:$C$4', [1, 1, 1])}</c:plus><c:minus>${numRef('Sayfa1!$C$2:$C$4', [1, 1, 1])}</c:minus></c:errBars>` }))],
    ['data kept outside the workbook', plain(S(), { after: '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>' })],
    ['shapes drawn over the chart', plain(S(), { after: '<c:userShapes r:id="rId1"/>' })],
    ['the source of a pivot chart', plain(S(), { before: '<c:pivotSource><c:name>[b.xlsx]S!P</c:name><c:fmtId val="0"/></c:pivotSource>' })],
    ['a range of an extension (sqref)', plain(S({ tail: '<c:extLst><c:ext uri="{x}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"><c15:fullRef><c15:sqref>Sayfa1!$B$2:$B$5</c15:sqref></c15:fullRef></c:ext></c:extLst>' }))],
    ['a formula of an extension (formulaRef)', plain(S({ tail: '<c:extLst><c:ext uri="{x}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"><c15:formulaRef>Sayfa1!$B$2:$B$5</c15:formulaRef></c:ext></c:extLst>' }))],
    ['a formula of an extension (f)', plain(S({ tail: '<c:extLst><c:ext uri="{x}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"><c15:datalabelsRange><c15:f>Sayfa1!$C$2:$C$4</c15:f></c15:datalabelsRange></c:ext></c:extLst>' }))],
    ['a reference whose sheet name is never closed', plain(S({ val: numRef("'Sayfa1!$B$2:$B$4", [1, 2, 3]) }))],
  ];
  for (const [label, xml] of gates) {
    const found = await facts(xml);
    assert.deepEqual([found.known, found.follow], [undefined, undefined], label);
    assert.equal(sheetCharts(found.wb, 0)[0].editable, 'none', label);
    stays(found.wb, label);
  }
});

test('chart gates: two frames for one part, a chart of a newer kind, a chart sheet and a chart inside a group each keep the workbook as it is', async () => {
  const two = await facts(plain(), { drawing: frame() + frame('rIdC1', 'Grafik 2', 20) });
  assert.deepEqual([two.known, two.follow, chartsOf(two.wb.sheets[0]).length], [undefined, undefined, 2], 'two frames');
  stays(two.wb, 'two frames');
  const newer = await facts(plain(), {}, async (zip, bytes) => {
    zip.set('xl/charts/chartEx1.xml', encoder.encode('<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"/>'));
    zip.set('[Content_Types].xml', encoder.encode((await text(bytes, '[Content_Types].xml')).replace('</Types>', '<Override PartName="/xl/charts/chartEx1.xml" ContentType="application/vnd.ms-office.chartex+xml"/></Types>')));
  });
  assert.deepEqual([newer.known, newer.follow], ['refs', undefined], 'a chart of a newer kind beside it');
  stays(newer.wb, 'chartex');
  const sheet = await facts(plain(), { chartSheet: plain() });
  assert.deepEqual([sheet.known, sheet.follow], ['refs', undefined], 'a chart sheet');
  assert.throws(() => applyOp(sheet.wb, { type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 }), /sheet with chart/);
  const grouped = await facts(plain(), {
    drawing: twoCell([4, 1], [9, 12], `<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="5" name="g"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></xdr:grpSpPr>${chartFrameOf('rIdC1')}</xdr:grpSp>`),
  });
  assert.equal(grouped.follow, undefined, 'a chart inside a group');
  stays(grouped.wb, 'a chart inside a group');
});

test('chart gates: a part the engine wrote that has relationships is not one it writes anew', async () => {
  const doc = await load(await salesBook());
  doc.history.apply(addChart(doc.wb, 0, { type: 'column', range: ALL }), undefined, undefined);
  const own = await save(doc);
  assert.equal(chartsOf((await reopen(own)).sheets[0])[0].known, 'all');
  const zip = ZipArchive.read(own);
  zip.set('xl/charts/style1.xml', encoder.encode('<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="201"/>'));
  zip.set('xl/charts/_rels/chart1.xml.rels', encoder.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/></Relationships>'));
  zip.set('[Content_Types].xml', encoder.encode((await text(own, '[Content_Types].xml')).replace('</Types>', '<Override PartName="/xl/charts/style1.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/></Types>')));
  const related = await load(await zip.write());
  assert.equal(chartsOf(related.wb.sheets[0])[0].known, 'refs', 'written anew, it would lose its relationships');
  assert.throws(() => updateChart(related.wb, 0, 0, { type: 'line' }), /Cannot change the type/);
  related.history.apply(updateChart(related.wb, 0, 0, { title: 'T' }), undefined, undefined);
  const out = await save(related);
  assert.deepEqual(await packageProblems(out), [], 'the style part is still led to');
  // The same part spelled with the strict namespaces is not what the engine writes either.
  const strict = ZipArchive.read(own);
  strict.set(CHART, encoder.encode((await text(own, CHART)).replace(/http:\/\/schemas\.openxmlformats\.org\/drawingml\/2006\/(chart|main)/g, 'http://purl.oclc.org/ooxml/drawingml/$1')));
  assert.equal(chartsOf((await reopen(await strict.write())).sheets[0])[0].known, 'refs');
});

test('chart writer: a drawing part two sheets show is not changed for one of them', async () => {
  const bytes = await salesBook({ drawing: frame(), charts: [plain()], sheets: [{ name: 'Other', xml: `${sheetData([['x', 1]])}<drawing r:id="rId9"/>` }] });
  const zip = ZipArchive.read(bytes);
  zip.set('xl/worksheets/_rels/sheet2.xml.rels', encoder.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>'));
  const doc = await load(await zip.write());
  doc.history.apply(moveChart(doc.wb, 0, 0, { top: 20, left: 0, bottom: 30, right: 4 }), undefined, undefined);
  await assert.rejects(save(doc), /is shared with sheet "Other"/);
});

test('chart writer: a frame that is moved keeps how it follows its cells while it has two corners, and loses it with another kind of anchor', async () => {
  const bytes = await salesBook({ drawing: frame('rIdC1', 'Grafik 1', 1, 'oneCell'), charts: [plain()] });
  const doc = await load(bytes);
  doc.history.apply(moveChart(doc.wb, 0, 0, { top: 20, left: 0, bottom: 30, right: 4 }), undefined, undefined);
  let xml = await text(await save(doc), 'xl/drawings/drawing1.xml');
  assert.ok(xml.indexOf('<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>20</xdr:row>') >= 0, 'two corners: the attribute stays');
  doc.history.apply(moveChart(doc.wb, 0, 0, { kind: 'oneCell', from: { col: 1, colOff: 0, row: 2, rowOff: 0 }, width: 3000000, height: 2000000 }), undefined, undefined);
  xml = await text(await save(doc), 'xl/drawings/drawing1.xml');
  assert.ok(xml.indexOf('<xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="3000000" cy="2000000"/><xdr:graphicFrame') >= 0, 'one corner and a size: no attribute, the frame as it was');
  assert.deepEqual((await reopen(await save(doc))).sheets[0].drawings?.[0].anchor, { kind: 'oneCell', from: { col: 1, colOff: 0, row: 2, rowOff: 0 }, width: 3000000, height: 2000000 });
  doc.history.apply(moveChart(doc.wb, 0, 0, { kind: 'absolute', x: 10, y: 20, width: 3000000, height: 2000000 }), undefined, undefined);
  xml = await text(await save(doc), 'xl/drawings/drawing1.xml');
  assert.ok(xml.indexOf('<xdr:absoluteAnchor><xdr:pos x="10" y="20"/><xdr:ext cx="3000000" cy="2000000"/>') >= 0, 'a place of the sheet');
  assert.throws(() => moveChart(doc.wb, 0, 0, { kind: 'twoCell', from: { col: 5, colOff: 0, row: 5, rowOff: 0 }, to: { col: 2, colOff: 0, row: 9, rowOff: 0 } }), /place of a chart/, 'a corner left of the other');
  assert.throws(() => moveChart(doc.wb, 0, 0, { kind: 'twoCell', from: { col: 5, colOff: 0, row: 5, rowOff: 0 }, to: { col: 6, colOff: 0, row: 2, rowOff: 0 } }), /place of a chart/, 'a corner above the other');
  assert.throws(() => moveChart(doc.wb, 0, 0, { kind: 'oneCell', from: { col: 1.5, colOff: 0, row: 2, rowOff: 0 }, width: 1, height: 1 }), /place of a chart/);
});

test('chart writer: what is put into a part stands where the schema has it; a series that is added stands behind the others, a title of two lines has two paragraphs', async () => {
  const doc = await load(await salesBook({ drawing: frame(), charts: [plain(S({ val: `${numRef('Sayfa1!$B$2:$B$4', [1, 2, 3]).replace('</c:numCache>', '<c:extLst><c:ext uri="{KEPT}"/></c:extLst></c:numCache>')}` }))] }));
  doc.history.apply(updateChart(doc.wb, 0, 0, { title: 'One\nTwo', series: [{ from: 0 }, { values: 'Sayfa1!$C$2:$C$4', name: 'Second' }] }), undefined, undefined);
  const xml = await text(await save(doc), CHART);
  assert.ok(xml.indexOf('<c:chart><c:title>') > 0, 'the title is the first child of the chart');
  assert.equal(xml.split('<a:p>').length, 3, 'a paragraph for each line of the title');
  assert.ok(/<c:idx val="0"\/>.*<c:idx val="1"\/><c:order val="1"\/><c:tx><c:v>Second<\/c:v><\/c:tx>.*<c:gapWidth|<c:idx val="1"\/>.*<c:axId/.test(xml) && xml.indexOf('<c:idx val="0"/>') < xml.indexOf('<c:idx val="1"/>'), 'the new series stands behind the one of the file');
  assert.ok(xml.indexOf('</c:ser><c:axId') > 0, 'and before the axes of the chart type');
  // The cache whose values are others is written anew; what it held beside its points stays.
  assert.ok(xml.indexOf('<c:pt idx="2"><c:v>30</c:v></c:pt><c:extLst><c:ext uri="{KEPT}"/></c:extLst></c:numCache>') > 0, 'the extension list of the cache is kept');
  // A chart that is in no file: two paragraphs too, and the title is read back with its line break.
  const fresh = await load(await salesBook());
  fresh.history.apply(addChart(fresh.wb, 0, { type: 'line', range: ALL, title: 'One\nTwo' }), undefined, undefined);
  const out = await save(fresh);
  assert.equal((await text(out, CHART)).split('<a:p>').length, 3);
  assert.equal(chartsOf((await reopen(out)).sheets[0])[0].chart.title?.text, 'One\nTwo');
});

test('chart ops: a chart that is added or moved after rows were inserted stands where it was put, and moves with the rows inserted after', async () => {
  const doc = await load(await salesBook());
  const row = (wb: Workbook, id = 0) => [sheetCharts(wb, 0)[id].anchor.from?.row, sheetCharts(wb, 0)[id].anchor.to?.row];
  doc.history.apply([{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 2 }], undefined, undefined);
  doc.history.apply(addChart(doc.wb, 0, { type: 'column', range: { top: 2, left: 0, bottom: 5, right: 2 }, anchor: { top: 8, left: 0, bottom: 17, right: 4 } }), undefined, undefined);
  assert.deepEqual(row(doc.wb), [8, 18], 'the rows inserted before it was added do not move it');
  assert.deepEqual(row(await reopen(await save(doc))), [8, 18], 'nor in the file');
  doc.history.apply([{ type: 'insert', sheet: 0, axis: 'row', at: 1, count: 3 }], undefined, undefined);
  assert.deepEqual(row(doc.wb), [11, 21], 'rows inserted afterwards do');
  assert.deepEqual(row(await reopen(await save(doc))), [11, 21]);
  doc.history.apply(moveChart(doc.wb, 0, 0, { top: 30, left: 0, bottom: 39, right: 4 }), undefined, undefined);
  assert.deepEqual(row(doc.wb), [30, 40], 'moved, it stands where it was moved to');
  assert.deepEqual(row(await reopen(await save(doc))), [30, 40]);
  doc.history.apply([{ type: 'delete', sheet: 0, axis: 'row', at: 0, count: 4 }], undefined, undefined);
  assert.deepEqual(row(doc.wb), [26, 36]);
  doc.history.apply(updateChart(doc.wb, 0, 0, { title: 'T' }), undefined, undefined);
  assert.deepEqual(row(doc.wb), [26, 36], 'a change of the chart leaves it where it is');
  assert.deepEqual(row(await reopen(await save(doc))), [26, 36]);
  // The same for a chart of the file that is moved after rows were inserted.
  const file = await load(await salesBook({ drawing: frame(), charts: [plain()] }));
  file.history.apply([{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 5 }], undefined, undefined);
  assert.deepEqual(row(file.wb), [6, 17], 'the frame of the file moved down with its rows');
  file.history.apply(moveChart(file.wb, 0, 0, { top: 2, left: 0, bottom: 9, right: 4 }), undefined, undefined);
  assert.deepEqual(row(file.wb), [2, 10]);
  assert.deepEqual(row(await reopen(await save(file))), [2, 10], 'in the file it is where it was moved to, not five rows further down');
  file.history.apply(moveChart(file.wb, 0, 0, { top: 2, left: 0, bottom: 9, right: 4 }), undefined, undefined);
  assert.equal(file.history.canUndo && file.wb.sheets[0].drawings?.[0].since, 1, 'moving it to where it is is no step');
});

test('chart ops: references without a sheet are those of the sheet of the chart; bubble sizes and the cells titles are read from follow too', async () => {
  const unqualified = plain(S({ tx: strRef('$B$1', ['N']), cat: strRef('$A$2:$A$4', ['a', 'b', 'c']), val: numRef('$B$2:$B$4', [1, 2, 3]) }));
  const doc = await load(await salesBook({ drawing: frame(), charts: [unqualified], sheets: OTHER }));
  assert.equal(doc.wb.chartsFollow, true);
  doc.history.apply([{ type: 'insert', sheet: 1, axis: 'row', at: 0, count: 1 }], undefined, undefined);
  assert.deepEqual(chartRefs(chartsOf(doc.wb.sheets[0])[0].chart), ['$B$1', '$A$2:$A$4', '$B$2:$B$4'], 'rows of another sheet are not theirs');
  doc.history.apply([{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 }], undefined, undefined);
  assert.deepEqual(chartRefs(chartsOf(doc.wb.sheets[0])[0].chart), ['$B$2', '$A$3:$A$5', '$B$3:$B$5'], 'rows of their own sheet are');
  const tRef = (f: string) => `<c:title><c:tx><c:strRef><c:f>${f}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>t</c:v></c:pt></c:strCache></c:strRef></c:tx><c:overlay val="0"/></c:title>`;
  const bubble = chartXml({
    head: tRef('Sayfa1!$A$1'),
    plot:
      group('bubbleChart', '<c:varyColors val="0"/>', xySer(0, { cat: numRef('Sayfa1!$B$2:$B$4', [1, 2, 3]), val: numRef('Sayfa1!$C$2:$C$4', [7, 8, 9]), tail: `<c:bubbleSize>${numRef('Sayfa1!$B$2:$B$3', [8, 9])}</c:bubbleSize>` })) +
      xyAxes().replace('<c:crossAx val="222"/>', `${tRef('Sayfa1!$B$1')}<c:crossAx val="222"/>`),
  });
  const sized = await load(await salesBook({ drawing: frame(), charts: [bubble] }));
  assert.equal(chartsOf(sized.wb.sheets[0])[0].known, 'refs');
  sized.history.apply([{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 2 }], undefined, undefined);
  const moved = chartsOf(sized.wb.sheets[0])[0].chart;
  assert.deepEqual(chartRefs(moved), ['Sayfa1!$A$3', 'Sayfa1!$B$3', 'Sayfa1!$B$4:$B$6', 'Sayfa1!$C$4:$C$6', 'Sayfa1!$B$4:$B$5']);
  assert.equal(moved.groups[0].series[0].bubbleSizes?.ref, 'Sayfa1!$B$4:$B$5');
  const out = await save(sized);
  assert.deepEqual(chartRefs(chartsOf((await reopen(out)).sheets[0])[0].chart), chartRefs(moved), 'in the part as in the model');
});

test('chart ops: a sheet that is deleted takes its own charts along untouched; a reference that cannot be read refuses the rename', () => {
  const wb = bookOf([['a', 1], ['b', 2]], [{ name: 'Other', rows: [[1]] }]);
  const history = new History(wb);
  history.apply(addChart(wb, 0, { type: 'line', series: [{ values: 'Data!$B$1:$B$2' }] }), undefined, undefined);
  history.apply(addChart(wb, 1, { type: 'line', series: [{ values: 'Data!$B$1:$B$2' }] }), undefined, undefined);
  const own = wb.sheets[0].drawings;
  const sheet = wb.sheets[0];
  history.apply([{ type: 'deleteSheet', sheet: 0 }], undefined, undefined);
  assert.equal(sheet.drawings, own, 'the charts of the sheet that went are as they were');
  assert.deepEqual(chartRefs(chartsOf(wb.sheets[0])[0].chart), ['#REF!']);
  history.undo();
  assert.deepEqual(chartRefs(chartsOf(wb.sheets[1])[0].chart), ['Data!$B$1:$B$2']);
  // As a reader could never leave it: a reference with a name that is never closed.
  const broken = chartsOf(wb.sheets[1])[0];
  broken.chart.groups[0].series[0].values!.ref = "'Data!$B$1:$B$2";
  assert.throws(() => applyOp(wb, { type: 'renameSheet', sheet: 0, name: 'X' }), /a reference of a chart cannot be read/);
  assert.equal(wb.sheets[0].name, 'Data', 'nothing was changed');
});

test('chart ops: what a new chart is made with, and what is refused', () => {
  const wb = bookOf([
    ['Year', 'A', 'B'],
    [2021, 1, 4],
    [2022, 2, 5],
    [2023, 3, 6],
  ]);
  const none = modelOf(addChart(wb, 0, { type: 'column', range: ALL, legend: 'none' })[0]);
  assert.equal(none.legend, undefined, 'no legend when none is asked for');
  assert.equal(none.autoTitleDeleted, true, 'a chart without a title says that none is made up');
  assert.equal(none.date1904, undefined);
  const titled = modelOf(addChart(wb, 0, { type: 'column', range: ALL, title: 'T' })[0]);
  assert.equal(titled.autoTitleDeleted, undefined);
  assert.deepEqual(titled.legend, { position: 'r' });
  // The years are numbers: the first column is a series, not the categories.
  assert.equal(titled.groups[0].series.length, 3);
  const years = modelOf(addChart(wb, 0, { type: 'line', series: [{ values: 'B2:B4', categories: 'A2:A4' }, { values: 'C2:C4', categories: 'A1:A4' }] })[0]);
  assert.equal(years.groups[0].series[0].categories?.kind, 'num', 'categories that are numbers are written as numbers');
  assert.equal(years.groups[0].series[1].categories?.kind, 'str', 'a text among them makes them texts');
  assert.equal(years.groups[0].series[0].values?.ref, 'Data!B2:B4', 'a reference without a sheet is of the sheet of the chart');
  wb.date1904 = true;
  assert.equal(modelOf(addChart(wb, 0, { type: 'column', range: ALL })[0]).date1904, true, 'the calendar of the workbook');
  // More columns of values than rows: the series are the rows.
  assert.equal(seriesOfRange(bookOf([[1, 2, 3, 4], [5, 6, 7, 8]]), 0, { top: 0, left: 0, bottom: 1, right: 3 }).seriesIn, 'rows');
  assert.equal(seriesOfRange(bookOf([[1, 2], [5, 6], [7, 8]]), 0, { top: 0, left: 0, bottom: 2, right: 1 }).seriesIn, 'columns');
  // Beside the range: two columns right of it.
  assert.equal(drawingOf(addChart(wb, 0, { type: 'column', range: { top: 0, left: 1, bottom: 3, right: 2 } })[0]).anchor.from?.col, 4);

  assert.throws(() => addChart(wb, 0, { type: 'line', range: ALL, title: 'a\u0001b' }), /holds a character a file cannot hold/);
  assert.throws(() => addChart(wb, 0, { type: 'line', series: new Array(MAX_CHART_SERIES + 1).fill({ values: 'B2:B4' }) }), /at most 255 series/);
  assert.throws(() => seriesOfRange(wb, 0, { top: 0, left: 0, bottom: 300, right: 300 }, 'columns'), /at most 255 series/);
  applyOp(wb, addChart(wb, 0, { type: 'line', range: ALL })[0]);
  assert.throws(() => updateChart(wb, 0, 0, { axisBounds: { min: 5, max: 5 } }), /larger than its smallest/);
  assert.throws(() => updateChart(wb, 0, 0, { axisBounds: { min: Number.NaN } }), /are numbers/);
  assert.throws(() => updateChart(wb, 0, 0, { grouping: 'sideways' as 'stacked' }), /stand side by side/);
  applyOp(wb, updateChart(wb, 0, 0, { type: 'pie' })[0]);
  assert.throws(() => updateChart(wb, 0, 0, { grouping: 'stacked' }), /Cannot change how the series of this chart are grouped/);
  assert.match(chartProblem(wb, 0, 0, { grouping: 'stacked' }) || '', /grouped/);
  assert.throws(() => updateChart(wb, 0, 0, { type: 'scatter', grouping: 'percentStacked' }), /grouped/);
  assert.throws(() => updateChart(wb, 0, 0, { axisTitles: { value: 'V' } }), /no such axis/);
  // Taken away, the chart at the place asked for goes, not another.
  applyOp(wb, addChart(wb, 0, { type: 'area', range: ALL })[0]);
  applyOp(wb, removeChart(wb, 0, 1)[0]);
  assert.deepEqual(sheetCharts(wb, 0).map((info) => info.type), ['pie']);
});

test('chart ops: no more charts are added than a file gives back when it is read', () => {
  const wb = bookOf([['a', 1], ['b', 2]], [{ name: 'Other', rows: [[1]] }, { name: 'Third', rows: [[1]] }]);
  const one = drawingOf(addChart(wb, 0, { type: 'line', series: [{ values: 'Data!$B$1:$B$2' }] })[0]);
  wb.sheets[0].drawings = new Array(MAX_SHEET_CHARTS - 1).fill(one);
  applyOp(wb, addChart(wb, 0, { type: 'line', series: [{ values: 'Data!$B$1:$B$2' }] })[0]);
  assert.equal(wb.sheets[0].drawings.length, MAX_SHEET_CHARTS);
  assert.throws(() => addChart(wb, 0, { type: 'line', series: [{ values: 'Data!$B$1:$B$2' }] }), new RegExp(`A sheet has at most ${MAX_SHEET_CHARTS} charts`));
  wb.sheets[1].drawings = new Array(MAX_WORKBOOK_CHARTS - MAX_SHEET_CHARTS).fill(one);
  assert.throws(() => addChart(wb, 2, { type: 'line', series: [{ values: 'Data!$B$1:$B$2' }] }), new RegExp(`A workbook has at most ${MAX_WORKBOOK_CHARTS} charts`));
  const picture: SheetDrawing = { anchor: one.anchor, object: { type: 'placeholder', what: 'object' } };
  wb.sheets[2].drawings = new Array(MAX_SHEET_DRAWINGS).fill(picture);
  wb.sheets[1].drawings = undefined;
  assert.throws(() => addChart(wb, 2, { type: 'line', series: [{ values: 'Data!$B$1:$B$2' }] }), /too many drawings/);
});

test('chart ops: a title that is added and taken away again leaves the chart saying what its part says', async () => {
  // The part states nothing about a title that is made up.
  const doc = await load(await salesBook({ drawing: frame(), charts: [plain()] }));
  const flag = async () => [chartsOf(doc.wb.sheets[0])[0].chart.autoTitleDeleted, chartsOf((await reopen(await save(doc))).sheets[0])[0].chart.autoTitleDeleted];
  doc.history.apply(updateChart(doc.wb, 0, 0, { title: 'T' }), undefined, undefined);
  assert.deepEqual(await flag(), [undefined, undefined], 'with a title');
  doc.history.apply(updateChart(doc.wb, 0, 0, { title: null }), undefined, undefined);
  assert.deepEqual(await flag(), [true, true], 'without one again: the model and the part say that none is made up');
  assert.ok((await text(await save(doc), CHART)).indexOf('<c:chart><c:autoTitleDeleted val="1"/><c:plotArea>') > 0, 'stated where the schema has it');
  doc.history.apply(updateChart(doc.wb, 0, 0, { title: 'Again' }), undefined, undefined);
  assert.deepEqual(await flag(), [undefined, undefined], 'and with a title once more');
});

test('chart layout: a series whose cells are gone shows nothing of what the file cached; the name of a series is read from the sheet called so now', () => {
  const wb = bookOf(
    [
      ['N', 1],
      ['M', 2],
    ],
    [{ name: 'Other', rows: [['other name']] }],
  );
  const history = new History(wb);
  history.apply(addChart(wb, 0, { type: 'line', series: [{ values: 'Data!$B$1:$B$2', nameRef: 'Data!$A$1' }] }), undefined, undefined);
  const chart = () => chartsOf(wb.sheets[0])[0].chart;
  // As a part of a file: with what it cached.
  chart().groups[0].series[0].values!.cache = { index: [0, 1], values: [7, 8] };
  chart().groups[0].series[0].values!.count = 2;
  assert.equal(chartValues(chart(), { workbook: wb }).series[0].values.length, 2);
  history.apply([{ type: 'delete', sheet: 0, axis: 'col', at: 1, count: 1 }], undefined, undefined);
  assert.equal(chart().groups[0].series[0].values?.ref, 'Data!#REF!');
  assert.equal(chartValues(chart(), { workbook: wb }).series[0].values.length, 0, 'nothing is drawn for cells that are gone');
  history.undo();
  // The sheet gets another name; a sheet that is added takes the old one. The chart reads the sheet it named.
  history.apply([{ type: 'renameSheet', sheet: 0, name: 'Renamed' }], undefined, undefined);
  history.apply([{ type: 'renameSheet', sheet: 1, name: 'Data' }], undefined, undefined);
  assert.equal(chart().groups[0].series[0].name?.ref, 'Renamed!$A$1');
  assert.equal(chartValues(chart(), { workbook: wb }).series[0].name, 'N', 'the cell of the sheet as it is called now');
  // A chart that names the sheet "Data" now means the sheet that is called so now.
  history.apply(addChart(wb, 0, { type: 'line', series: [{ values: 'Renamed!$B$1:$B$2', nameRef: 'Data!$A$1' }] }), undefined, undefined);
  assert.equal(chartValues(chartsOf(wb.sheets[0])[1].chart, { workbook: wb }).series[0].name, 'other name');
});

test('chart ops: what is taken away from a chart of the file and put back is in the model as the part will hold it', async () => {
  const doc = await load(await salesBook({ drawing: frame(), charts: [FOREIGN] }));
  const read = chartsOf(doc.wb.sheets[0])[0].chart;
  assert.deepEqual(read.legend, { position: 'r', style: { size: 8 } }, 'the legend of the file has a look of its own');
  const same = async (label: string) => assert.deepEqual(bare(chartsOf((await reopen(await save(doc))).sheets[0])[0].chart), bare(chartsOf(doc.wb.sheets[0])[0].chart), label);
  doc.history.apply(updateChart(doc.wb, 0, 0, { legend: 'none', title: null }), undefined, undefined);
  await same('without legend and title');
  // Both come back: the part still holds how its legend looked, and a title that is new has no look of its own.
  doc.history.apply(updateChart(doc.wb, 0, 0, { legend: 'b', title: 'Back' }), undefined, undefined);
  assert.deepEqual(chartsOf(doc.wb.sheets[0])[0].chart.legend, { position: 'b', style: { size: 8 } }, 'the legend has the look the part gives it');
  await same('with legend and title again');
  doc.history.apply(updateChart(doc.wb, 0, 0, { dataLabels: true }), undefined, undefined);
  doc.history.apply(updateChart(doc.wb, 0, 0, { dataLabels: false }), undefined, undefined);
  doc.history.apply(updateChart(doc.wb, 0, 0, { dataLabels: true, axisTitles: { value: 'V' } }), undefined, undefined);
  await same('labels on, off and on');
});

test('chart gates: a chart part is told by its content type and by the relationship that leads to it, wherever it stands', async () => {
  const moved = async (type: string, relType = 'chart') =>
    facts(plain(), {}, async (zip, bytes) => {
      zip.set('xl/graphs/c1.xml', await zip.read(CHART));
      zip.delete(CHART);
      zip.set('xl/drawings/_rels/drawing1.xml.rels', encoder.encode((await text(bytes, 'xl/drawings/_rels/drawing1.xml.rels')).replace('../charts/chart1.xml', '../graphs/c1.xml').replace('relationships/chart"', `relationships/${relType}"`)));
      zip.set('[Content_Types].xml', encoder.encode((await text(bytes, '[Content_Types].xml')).replace('<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>', `<Override PartName="/xl/graphs/c1.xml" ContentType="${type}"/>`)));
    });
  // Outside the usual folder and with a content type that says nothing: the relationship of the drawing says it is a chart.
  const byRel = await moved('application/xml');
  assert.deepEqual([byRel.known, byRel.follow], ['refs', true], 'known by its relationship: it follows');
  applyOp(byRel.wb, { type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 });
  assert.deepEqual(chartRefs(chartsOf(byRel.wb.sheets[0])[0].chart), ['Sayfa1!$B$2', 'Sayfa1!$A$3:$A$5', 'Sayfa1!$B$3:$B$5']);
  // A chart of a newer kind, told by nothing but its relationship or by nothing but its content type.
  const newerRel = await moved('application/xml', 'chartEx');
  assert.equal(newerRel.follow, undefined, 'a relationship of the newer kind');
  stays(newerRel.wb, 'a relationship of the newer kind');
  // The same beside a chart that would follow: the workbook stays as it is, the sheet without the newer chart too.
  const beside = await facts(plain(), { drawing: frame() + frame('rIdX', 'Yeni', 20) }, async (zip, bytes) => {
    zip.set('xl/graphs/x.xml', encoder.encode('<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"/>'));
    zip.set('xl/drawings/_rels/drawing1.xml.rels', encoder.encode((await text(bytes, 'xl/drawings/_rels/drawing1.xml.rels')).replace('</Relationships>', '<Relationship Id="rIdX" Type="http://schemas.microsoft.com/office/2014/relationships/chartEx" Target="../graphs/x.xml"/></Relationships>')));
    zip.set('[Content_Types].xml', encoder.encode((await text(bytes, '[Content_Types].xml')).replace('</Types>', '<Override PartName="/xl/graphs/x.xml" ContentType="application/xml"/></Types>')));
  });
  assert.deepEqual([beside.known, beside.follow], ['refs', undefined], 'a chart of the newer kind beside one that would follow');
  stays(beside.wb, 'a chart of the newer kind beside one that would follow');
  const newerType = await facts(plain(), {}, async (zip, bytes) => {
    zip.set('xl/extra/x1.xml', encoder.encode('<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"/>'));
    zip.set('[Content_Types].xml', encoder.encode((await text(bytes, '[Content_Types].xml')).replace('</Types>', '<Override PartName="/xl/extra/x1.xml" ContentType="application/vnd.ms-office.chartex+xml"/></Types>')));
  });
  assert.equal(newerType.follow, undefined, 'a content type of the newer kind');
  stays(newerType.wb, 'a content type of the newer kind');
});

test('chart writer: a namespace the part has no prefix for is declared where it is needed; a frame that was not moved is not written', async () => {
  // The part declares no prefix for the namespace of drawings: it has no text and no colour of its own.
  const bare_ = plain().replace(' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"', '');
  const bytes = await salesBook({ drawing: frame().replace('<xdr:from>', '<xdr:from >').replace('<xdr:col>', '\n <xdr:col>'), charts: [bare_] });
  const doc = await load(bytes);
  doc.history.apply(updateChart(doc.wb, 0, 0, { title: 'T', axisTitles: { value: 'V' }, series: [{ from: 0, color: '#102030' }] }), undefined, undefined);
  const out = await save(doc);
  assert.deepEqual(await packageProblems(out), [], 'every prefix that was written is bound');
  const xml = await text(out, CHART);
  assert.ok(xml.indexOf('<c:title xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">') > 0, 'declared on the element that was put in');
  assert.equal(chartsOf((await reopen(out)).sheets[0])[0].chart.groups[0].series[0].shape?.fill, '#102030');
  // Only the chart changed: its frame in the drawing part is the bytes of the file.
  assert.deepEqual(Array.from(await ZipArchive.read(out).read('xl/drawings/drawing1.xml')), Array.from(await ZipArchive.read(bytes).read('xl/drawings/drawing1.xml')), 'the drawing part');
});
