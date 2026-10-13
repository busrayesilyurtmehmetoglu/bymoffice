// Charts written into the package: the first chart of a sheet (a chart part, a drawing part,
// their relationships and content types, <drawing> in the sheet), a chart added beside
// other drawings, charts moved and deleted, undo back to the bytes of the file, a sheet that
// is not in the file, and what the reader makes of what was written.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { buildXlsx } from '../src/formats/xlsx/build';
import { parseXml, type SheetDrawing, type XmlElement } from '../src/index';
import { setCellInput } from '../src/ops/commands';
import { ZipArchive } from '../src/zip';
import { addChart, moveChart, removeChart, sheetCharts, updateChart, CHART_TYPES } from '../src/ops/charts';
import { FOREIGN, bare, chartsOf, load, names, packageProblems, part, relationProblems, reopen, salesBook, same, save, text } from './chart-edit-fixtures';
import { chartFrameOf } from './chart-fixtures';
import { sp, twoCell } from './sheet-drawing-fixtures';

const ALL = { top: 0, left: 0, bottom: 3, right: 2 };
const CHART_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const DRAWING_TYPE = 'application/vnd.openxmlformats-officedocument.drawing+xml';

const whole = async (bytes: Uint8Array, label: string) => assert.deepEqual(await packageProblems(bytes), [], label);

/** The local names of the children of an element, in order. */
const kids = (el: XmlElement): string[] => el.children.filter((node): node is XmlElement => node.type === 'element').map((node) => node.name.replace(/^.*:/, ''));
const find = (el: XmlElement, path: string[]): XmlElement => {
  let at = el;
  for (const name of path) {
    const next = at.children.find((node): node is XmlElement => node.type === 'element' && node.name.replace(/^.*:/, '') === name);
    if (!next) throw new Error(`no <${name}> in <${at.name}>`);
    at = next;
  }
  return at;
};

test('chart write: the first chart of a sheet gets a chart part, a drawing part, relationships, content types and <drawing> in its place', async () => {
  const doc = await load(await salesBook());
  doc.history.apply(addChart(doc.wb, 0, { type: 'column', range: ALL, title: 'Sales' }), undefined, undefined);
  const saved = await save(doc);
  await whole(saved, 'the package is whole');
  const added = names(saved).filter((name) => names(doc.bytes).indexOf(name) < 0);
  assert.deepEqual(added, ['xl/charts/chart1.xml', 'xl/drawings/_rels/drawing1.xml.rels', 'xl/drawings/drawing1.xml', 'xl/worksheets/_rels/sheet1.xml.rels']);
  const types = await text(saved, '[Content_Types].xml');
  assert.ok(types.indexOf(`<Override PartName="/xl/charts/chart1.xml" ContentType="${CHART_TYPE}"/>`) >= 0, 'the chart part has its content type');
  assert.ok(types.indexOf(`<Override PartName="/xl/drawings/drawing1.xml" ContentType="${DRAWING_TYPE}"/>`) >= 0, 'the drawing part has its content type');
  // <drawing> stands behind the margins of the page, where the schema has it.
  const sheet = parseXml(await part(saved, 'xl/worksheets/sheet1.xml')).root;
  assert.deepEqual(kids(sheet), ['sheetData', 'pageMargins', 'drawing']);
  assert.deepEqual(await relationProblems(saved, 'xl/worksheets/sheet1.xml'), []);
  assert.deepEqual(await relationProblems(saved, 'xl/drawings/drawing1.xml'), []);
  const rels = await text(saved, 'xl/worksheets/_rels/sheet1.xml.rels');
  assert.match(rels, /Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/drawing" Target="\.\.\/drawings\/drawing1\.xml"/);
  assert.match(await text(saved, 'xl/drawings/_rels/drawing1.xml.rels'), /relationships\/chart" Target="\.\.\/charts\/chart1\.xml"/);
  // The frame: two corners, a graphic frame that names the chart.
  const drawing = parseXml(await part(saved, 'xl/drawings/drawing1.xml')).root;
  assert.deepEqual(kids(drawing), ['twoCellAnchor']);
  assert.deepEqual(kids(find(drawing, ['twoCellAnchor'])), ['from', 'to', 'graphicFrame', 'clientData']);
  // The sheet part changed in nothing but the element.
  assert.equal((await text(saved, 'xl/worksheets/sheet1.xml')).replace(/<drawing [^>]*\/>/, ''), await text(doc.bytes, 'xl/worksheets/sheet1.xml'));
});

test('chart write: the chart part states its elements in the order of the schema, with the values of the cells cached', async () => {
  const doc = await load(await salesBook());
  doc.history.apply(addChart(doc.wb, 0, { type: 'column', range: ALL, title: 'Sales' }), undefined, undefined);
  const saved = await save(doc);
  const root = parseXml(await part(saved, 'xl/charts/chart1.xml')).root;
  assert.equal(root.name, 'c:chartSpace');
  assert.deepEqual(kids(root), ['date1904', 'roundedCorners', 'chart']);
  assert.deepEqual(kids(find(root, ['chart'])), ['title', 'autoTitleDeleted', 'plotArea', 'legend', 'plotVisOnly', 'dispBlanksAs']);
  assert.deepEqual(kids(find(root, ['chart', 'plotArea'])), ['layout', 'barChart', 'catAx', 'valAx']);
  assert.deepEqual(kids(find(root, ['chart', 'plotArea', 'barChart'])), ['barDir', 'grouping', 'varyColors', 'ser', 'ser', 'gapWidth', 'axId', 'axId']);
  assert.deepEqual(kids(find(root, ['chart', 'plotArea', 'barChart', 'ser'])), ['idx', 'order', 'tx', 'invertIfNegative', 'cat', 'val']);
  assert.deepEqual(kids(find(root, ['chart', 'plotArea', 'catAx'])), ['axId', 'scaling', 'delete', 'axPos', 'numFmt', 'majorTickMark', 'minorTickMark', 'tickLblPos', 'crossAx', 'crosses', 'auto', 'lblAlgn', 'lblOffset', 'noMultiLvlLbl']);
  assert.deepEqual(kids(find(root, ['chart', 'plotArea', 'valAx'])), ['axId', 'scaling', 'delete', 'axPos', 'majorGridlines', 'numFmt', 'majorTickMark', 'minorTickMark', 'tickLblPos', 'crossAx', 'crosses', 'crossBetween']);
  const xml = await text(saved, 'xl/charts/chart1.xml');
  assert.ok(xml.indexOf('<c:tx><c:strRef><c:f>Sayfa1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>North</c:v></c:pt></c:strCache></c:strRef></c:tx>') > 0, 'the cached values are those of the cells');
  assert.ok(xml.indexOf('<c:cat><c:strRef><c:f>Sayfa1!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>Jan</c:v></c:pt><c:pt idx="1"><c:v>Feb</c:v></c:pt><c:pt idx="2"><c:v>Mar</c:v></c:pt></c:strCache></c:strRef></c:cat>') > 0, 'the cached values are those of the cells');
  assert.ok(xml.indexOf('<c:val><c:numRef><c:f>Sayfa1!$C$2:$C$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>30</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>10</c:v></c:pt></c:numCache></c:numRef></c:val>') > 0, 'the cached values are those of the cells');
});

test('chart write: what is written is read back as the model, for every type, and as a chart the engine wrote', async () => {
  for (const type of CHART_TYPES) {
    for (const grouping of ['standard', 'stacked', 'percentStacked'] as const) {
      const doc = await load(await salesBook());
      doc.history.apply(addChart(doc.wb, 0, { type, grouping, range: ALL, title: type === 'pie' ? undefined : 'A title\nof two lines', legend: type === 'line' ? 'none' : 'b' }), undefined, undefined);
      if (type !== 'pie' && type !== 'doughnut') doc.history.apply(updateChart(doc.wb, 0, 0, { axisTitles: { category: 'X', value: 'Y' }, axisBounds: { min: 0, max: 100 }, dataLabels: true, series: [{ from: 1, color: '#112233' }, { from: 0 }] }), undefined, undefined);
      const model = chartsOf(doc.wb.sheets[0])[0];
      const saved = await save(doc);
      await whole(saved, `${type} ${grouping}: the package is whole`);
      const again = await reopen(saved);
      const read = chartsOf(again.sheets[0])[0];
      assert.deepEqual(bare(read.chart), bare(model.chart), `${type} ${grouping}: the chart that is read is the chart that was written`);
      assert.equal(read.known, 'all', `${type} ${grouping}: the part is known as one the engine wrote`);
      assert.equal(read.part, '/xl/charts/chart1.xml');
      assert.equal(read.name, 'Chart 1');
      assert.equal(again.chartsFollow, true, 'the charts of the file follow its rows and columns');
      assert.deepEqual((again.sheets[0].drawings as SheetDrawing[])[0].anchor, sheetCharts(doc.wb, 0)[0].anchor, 'the chart is where it was put');
      // The values the file cached are the cells.
      assert.deepEqual(read.chart.groups[0].series[type === 'pie' || type === 'doughnut' ? 0 : 1].values?.cache, { index: [0, 1, 2], values: [10, 20, 30] });
    }
  }
});

test('chart write: a chart that was added and taken back leaves the file as it was, before and after it was saved', async () => {
  const bytes = await salesBook();
  const doc = await load(bytes);
  doc.history.apply(addChart(doc.wb, 0, { type: 'line', range: ALL }), undefined, undefined);
  doc.history.undo();
  same(await save(doc), bytes, 'added, undone, saved');
  doc.history.redo();
  const saved = await save(doc);
  assert.ok(names(saved).indexOf('xl/charts/chart1.xml') >= 0, 'the chart is written');
  doc.history.undo();
  same(await save(doc), bytes, 'added, saved, undone, saved');
  doc.history.redo();
  same(await save(doc), saved, 'and written again the same way under the same names');
});

const shape = twoCell([8, 1], [10, 6], sp({ name: 'Box' }));
const frame = (rel: string, name: string, from: [number, number], to: [number, number]) => twoCell(from, to, chartFrameOf(rel, name));

/** A workbook with a shape, a chart of another application (with what the model does not hold) and nothing else on its sheet. */
const foreignBook = () => salesBook({ drawing: shape + frame('rIdC1', 'Grafik 1', [4, 1], [9, 12]), charts: [FOREIGN] });

test('chart write: a chart added beside the drawings of the file leaves them as they are', async () => {
  const bytes = await foreignBook();
  const doc = await load(bytes);
  assert.equal(chartsOf(doc.wb.sheets[0])[0].known, 'refs', 'the chart of the file holds more than the model, and every reference of it is in the model');
  assert.equal(doc.wb.chartsFollow, true);
  doc.history.apply(addChart(doc.wb, 0, { type: 'pie', range: { top: 0, left: 0, bottom: 3, right: 1 } }), undefined, undefined);
  const saved = await save(doc);
  await whole(saved, 'the package is whole');
  same(await part(saved, 'xl/charts/chart1.xml'), await part(bytes, 'xl/charts/chart1.xml'), 'the chart of the file');
  same(await part(saved, 'xl/worksheets/sheet1.xml'), await part(bytes, 'xl/worksheets/sheet1.xml'), 'the sheet: it has its <drawing> already');
  assert.ok(names(saved).indexOf('xl/charts/chart2.xml') >= 0, 'the new chart has a part of its own');
  const drawing = await text(saved, 'xl/drawings/drawing1.xml');
  assert.ok(drawing.indexOf(shape) > 0, 'the shape is as it was');
  assert.ok(drawing.indexOf(frame('rIdC1', 'Grafik 1', [4, 1], [9, 12])) > 0, 'the frame of the chart of the file is as it was');
  assert.deepEqual(await relationProblems(saved, 'xl/drawings/drawing1.xml'), []);
  const again = await reopen(saved);
  assert.deepEqual(
    (again.sheets[0].drawings as SheetDrawing[]).map((item) => item.object.type),
    ['shape', 'chart', 'chart'],
  );
  assert.equal(chartsOf(again.sheets[0])[1].chart.groups[0].kind, 'pie');
  assert.equal(chartsOf(again.sheets[0])[1].name, 'Chart 2');
  doc.history.undo();
  same(await save(doc), bytes, 'undone and saved');
});

test('chart write: a chart of the file that is moved gets another anchor and keeps its part', async () => {
  const bytes = await foreignBook();
  const doc = await load(bytes);
  doc.history.apply(moveChart(doc.wb, 0, 1, { top: 20, left: 2, bottom: 29, right: 7 }), undefined, undefined);
  const saved = await save(doc);
  await whole(saved, 'the package is whole');
  same(await part(saved, 'xl/charts/chart1.xml'), await part(bytes, 'xl/charts/chart1.xml'), 'the chart part');
  same(await part(saved, 'xl/worksheets/sheet1.xml'), await part(bytes, 'xl/worksheets/sheet1.xml'), 'the sheet part');
  const drawing = await text(saved, 'xl/drawings/drawing1.xml');
  assert.ok(drawing.indexOf(shape) > 0, 'the shape is as it was');
  assert.ok(drawing.indexOf(frame('rIdC1', 'Grafik 1', [2, 20], [8, 30])) > 0, 'the frame stands where the chart was put');
  const again = await reopen(saved);
  assert.deepEqual((again.sheets[0].drawings as SheetDrawing[])[1].anchor, { kind: 'twoCell', from: { col: 2, colOff: 0, row: 20, rowOff: 0 }, to: { col: 8, colOff: 0, row: 30, rowOff: 0 } });
  doc.history.undo();
  same(await save(doc), bytes, 'undone and saved');
});

test('chart write: a chart of the file that is deleted takes its part, its relationship and its content type along; the last drawing takes the drawing part', async () => {
  const bytes = await foreignBook();
  const doc = await load(bytes);
  doc.history.apply(removeChart(doc.wb, 0, 1), undefined, undefined);
  const saved = await save(doc);
  await whole(saved, 'the package is whole');
  assert.ok(names(saved).indexOf('xl/charts/chart1.xml') < 0, 'the chart part is gone');
  assert.ok((await text(saved, '[Content_Types].xml')).indexOf('chart1.xml') < 0, 'and its content type');
  assert.deepEqual(await relationProblems(saved, 'xl/drawings/drawing1.xml'), []);
  assert.ok((await text(saved, 'xl/drawings/drawing1.xml')).indexOf(shape) > 0, 'the shape is as it was');
  assert.deepEqual((await reopen(saved)).sheets[0].drawings?.map((item) => item.object.type), ['shape']);
  doc.history.undo();
  same(await save(doc), bytes, 'undone and saved');

  // A sheet whose only drawing is a chart: the drawing part goes with it, and <drawing> with that.
  const lone = await salesBook({ drawing: frame('rIdC1', 'Grafik 1', [4, 1], [9, 12]), charts: [FOREIGN] });
  const other = await load(lone);
  other.history.apply(removeChart(other.wb, 0, 0), undefined, undefined);
  const emptied = await save(other);
  await whole(emptied, 'the package is whole');
  assert.deepEqual(
    names(emptied).filter((name) => /chart|drawing/.test(name)),
    [],
    'nothing of the chart or the drawing is left',
  );
  // The relationships of the sheet are a part of the file: it stays, and holds none.
  assert.ok(!/<Relationship /.test(await text(emptied, 'xl/worksheets/_rels/sheet1.xml.rels')), 'the sheet has no relationship left');
  assert.ok((await text(emptied, 'xl/worksheets/sheet1.xml')).indexOf('<drawing') < 0, 'the sheet names no drawing');
  assert.ok((await text(emptied, '[Content_Types].xml')).indexOf('drawing') < 0, 'no content type is left behind');
  assert.equal((await reopen(emptied)).sheets[0].drawings, undefined);
  other.history.undo();
  same(await save(other), lone, 'undone and saved');
});

test('chart write: a chart on a sheet that is not in the file is written with the sheet', async () => {
  const bytes = await salesBook();
  const doc = await load(bytes);
  doc.history.apply([{ type: 'insertSheet', sheet: 1, name: 'New' }], undefined, undefined);
  doc.history.apply(addChart(doc.wb, 1, { type: 'line', series: [{ values: 'Sayfa1!$B$2:$B$4', categories: 'Sayfa1!$A$2:$A$4', nameRef: 'Sayfa1!$B$1' }] }), undefined, undefined);
  const saved = await save(doc);
  await whole(saved, 'the package is whole');
  const again = await reopen(saved);
  assert.equal(again.sheets[1].name, 'New');
  assert.equal(chartsOf(again.sheets[1]).length, 1);
  assert.equal(chartsOf(again.sheets[1])[0].known, 'all');
  assert.deepEqual(kids(parseXml(await part(saved, 'xl/worksheets/sheet2.xml')).root).slice(-1), ['drawing']);
  doc.history.undo();
  const without = await save(doc);
  await whole(without, 'the package is whole without the chart');
  assert.deepEqual(names(without).filter((name) => /chart|drawing/.test(name)), []);
  doc.history.undo();
  same(await save(doc), bytes, 'everything undone');
});

test('chart write: a file built anew from the model holds its charts, written from their models', async () => {
  const doc = await load(await foreignBook());
  doc.history.apply(addChart(doc.wb, 0, { type: 'line', range: ALL, title: 'New' }), undefined, undefined);
  const report: { parts: string[]; warnings: string[] } = { parts: [], warnings: [] };
  const built = await buildXlsx(doc.wb, { report });
  await whole(built, 'the package is whole');
  assert.deepEqual(names(built).filter((name) => /charts\//.test(name)), ['xl/charts/chart1.xml', 'xl/charts/chart2.xml']);
  assert.deepEqual(await relationProblems(built, 'xl/drawings/drawing1.xml'), []);
  assert.ok(report.warnings.some((line) => /a chart is written from the model; formatting its part held beside it is not kept/.test(line)), `what is not kept is said: ${report.warnings.join(' | ')}`);
  const again = await reopen(built);
  const read = chartsOf(again.sheets[0]);
  assert.deepEqual(read.map((chart) => [chart.chart.groups[0].kind, chart.known, chart.chart.title?.text]), [
    ['bar', 'all', 'Old'],
    ['line', 'all', 'New'],
  ]);
  assert.deepEqual(read[0].chart.groups[0].series[0].values?.cache.values, [10, 20, 30], 'the cells are cached as they are');
});

test('chart write: a title names no language, since the workbook says none', async () => {
  const doc = await load(await salesBook());
  doc.history.apply(addChart(doc.wb, 0, { type: 'column', range: ALL, title: 'Satışlar: ğüşıöç\nİkinci satır' }), undefined, undefined);
  doc.history.apply(updateChart(doc.wb, 0, sheetCharts(doc.wb, 0)[0].id, { axisTitles: { category: 'Aylar', value: 'Tutar (₺)' } }), undefined, undefined);
  const saved = await save(doc);
  await whole(saved, 'the package is whole');
  const xml = await text(saved, 'xl/charts/chart1.xml');
  assert.equal((xml.match(/<c:title>/g) || []).length, 3, 'the chart and both axes have a title');
  // "en-US" on Turkish text makes a spelling checker underline all of it; without the attribute the application decides.
  assert.equal(xml.indexOf(' lang='), -1, 'no language is claimed');
  assert.match(xml, /<c:rich><a:bodyPr\/><a:lstStyle\/><a:p><a:r><a:t>Satışlar: ğüşıöç<\/a:t><\/a:r><\/a:p><a:p><a:r><a:t>İkinci satır<\/a:t><\/a:r><\/a:p><\/c:rich>/);
  const read = chartsOf((await reopen(saved)).sheets[0]);
  assert.equal(read[0].chart.title?.text, 'Satışlar: ğüşıöç\nİkinci satır');
  assert.deepEqual(read[0].chart.axes.map((axis) => axis.title?.text), ['Aylar', 'Tutar (₺)']);
});

test('chart write: a chart this engine wrote when its titles still named a language is known as its own, and loses the language when it is changed', async () => {
  const doc = await load(await salesBook());
  doc.history.apply(addChart(doc.wb, 0, { type: 'column', range: ALL, title: 'Sales\nby month' }), undefined, undefined);
  doc.history.apply(updateChart(doc.wb, 0, sheetCharts(doc.wb, 0)[0].id, { axisTitles: { value: 'Amount' } }), undefined, undefined);
  const saved = await save(doc);
  // The part as it was written before: every run of a title said lang="en-US".
  const zip = ZipArchive.read(saved);
  const now = await text(saved, 'xl/charts/chart1.xml');
  const then = now.split('<a:r><a:t>').join('<a:r><a:rPr lang="en-US"/><a:t>');
  assert.equal(then.split('lang="en-US"').length - 1, 3);
  zip.set('xl/charts/chart1.xml', new TextEncoder().encode(then));
  const old = await zip.write();
  const again = await load(old);
  const chart = chartsOf(again.wb.sheets[0])[0];
  assert.equal(chart.known, 'all');
  assert.equal(chart.chart.title?.text, 'Sales\nby month');
  // Saved without a change of the chart: the part is the one of the file.
  again.history.apply(setCellInput(again.wb, 0, 9, 9, 'x'), undefined, undefined);
  assert.equal(await text(await save(again), 'xl/charts/chart1.xml'), then);
  // Changed in what only a chart of its own allows (the type): written anew, as it is written now.
  again.history.apply(updateChart(again.wb, 0, sheetCharts(again.wb, 0)[0].id, { type: 'line' }), undefined, undefined);
  const rewritten = await save(again);
  await whole(rewritten, 'the package is whole');
  assert.equal((await text(rewritten, 'xl/charts/chart1.xml')).indexOf(' lang='), -1);
  // Another language, or more than the language, is not what the engine wrote.
  zip.set('xl/charts/chart1.xml', new TextEncoder().encode(now.split('<a:r><a:t>').join('<a:r><a:rPr lang="tr-TR" b="1"/><a:t>')));
  assert.equal(chartsOf((await load(await zip.write())).wb.sheets[0])[0].known, 'refs');
});
