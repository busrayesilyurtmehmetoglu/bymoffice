// What a review of copied charts found, and what it asked to be held down by tests: the budget
// of a workbook, the frame of a copy, when a copy is outdated, relationships that lead astray,
// and what goes with a copy and what does not.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { ZipArchive, parseXml } from '../src/index';
import { parseChart } from '../src/formats/xlsx/chart';
import { chartLoose, stripChartPart } from '../src/formats/xlsx/chart-write';
import { chartRefs } from '../src/model/sheet/chart';
import { chartCopyOutdated, chartCopyProblem, copyChart, cutChart, duplicateChart, pasteChart, removeChart, sheetCharts, updateChart } from '../src/ops/charts';
import { deleteSheet, insertSheet, renameSheet } from '../src/ops/sheets';
import { drawingAnchorNow } from '../src/render/sheet-drawings';
import { RICH, SHAPED, frameAt, richBook } from './chart-copy-fixtures';
import { pythonHere, readByPython } from './chart-copy-python';
import { FOREIGN, chartsOf, load, names, packageProblems, part, reopen, salesBook, same, save, text, type Loaded } from './chart-edit-fixtures';
import { OFFICE_THEME, chartFrameOf } from './chart-fixtures';
import { twoCell } from './sheet-drawing-fixtures';

const encoder = new TextEncoder();
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const RELS = (inner: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;
const STD =
  '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/>' +
  `<Relationship Id="rId3" Type="${REL}/image" Target="../media/image1.png"/>`;
const whole = async (bytes: Uint8Array, label: string) => assert.deepEqual(await packageProblems(bytes), [], label);
const apply = (doc: Loaded, ops: Parameters<Loaded['history']['apply']>[0]) => doc.history.apply(ops, undefined, undefined);

/** A file with parts put in, changed (a function of the text they have) or taken out (null). */
async function mod(bytes: Uint8Array, change: { [name: string]: string | Uint8Array | null | ((had: string) => string) }): Promise<Uint8Array> {
  const zip = ZipArchive.read(bytes);
  for (const name of Object.keys(change)) {
    const value = change[name];
    if (value === null) zip.delete(name);
    else if (typeof value === 'function') zip.set(name, encoder.encode(value(new TextDecoder().decode(await zip.read(name)))));
    else zip.set(name, typeof value === 'string' ? encoder.encode(value) : value);
  }
  return zip.write();
}

/** The anchors of a drawing part as texts. */
const anchorsOf = (xml: string): string[] => xml.match(/<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[\s\S]*?<\/xdr:\1>/g) || [];

// ------------------------------------------------------------------ the budget of a workbook

test('chart copy review: copies that would hold more than a workbook reads back are refused when they are made', async () => {
  const pad = `<!--${'x'.repeat(22 * 1024 * 1024)}-->`;
  const doc = await load(await richBook({ first: encoder.encode(RICH.replace('<c:chart>', `${pad}<c:chart>`)) }));
  assert.equal(chartsOf(doc.wb.sheets[0])[0].known, 'refs');
  apply(doc, duplicateChart(doc.wb, 0, 0).ops);
  // A third time the part would be 66 MB of chart parts: more than a workbook is read with (64 MB).
  assert.throws(() => duplicateChart(doc.wb, 0, 0), /Cannot add the chart: the charts of the workbook would hold more than is read back from a file/);
  assert.throws(() => pasteChart(doc.wb, 1, copyChart(doc.wb, 0, 0)), /would hold more than is read back/);
  // A small chart still fits.
  apply(doc, duplicateChart(doc.wb, 0, 1).ops);
  const again = await reopen(await save(doc));
  assert.deepEqual((again.sheets[0].drawings || []).map((drawing) => drawing.object.type), ['chart', 'chart', 'chart', 'chart'], 'every chart is read back as a chart');
});

// ------------------------------------------------------------------ the frame

const FRAMED = async (chart = RICH): Promise<Uint8Array> =>
  mod(await richBook({ first: chart }), {
    'xl/drawings/drawing1.xml': (had) =>
      had
        .replace('<xdr:twoCellAnchor>', '<xdr:twoCellAnchor editAs="oneCell">')
        .replace('<xdr:graphicFrame macro="">', '<xdr:graphicFrame macro="[0]!Makro1" fPublished="1">')
        .replace('name="Grafik 1"', 'name="Grafik 1" descr="alt &amp; text" hidden="1" title="T"')
        .replace('<xdr:clientData/>', '<xdr:clientData fLocksWithSheet="0" fPrintsWithSheet="0"/>'),
  });

test('chart copy review: the frame of a copy is the frame of the chart, but for its id, its name, its place and its macro', async () => {
  const bytes = await FRAMED();
  for (const target of [0, 1]) {
    const doc = await load(bytes);
    const source = chartsOf(doc.wb.sheets[0])[0];
    assert.deepEqual([source.description, source.hidden, source.macro], ['alt & text', true, '[0]!Makro1']);
    const copy = copyChart(doc.wb, 0, 0);
    assert.deepEqual(copy.dropped, ['macro'], 'the macro does not go with a copy, and that is said');
    assert.equal(copy.editAs, 'oneCell');
    const pasted = pasteChart(doc.wb, target, copy);
    assert.deepEqual(pasted.dropped, ['macro']);
    apply(doc, pasted.ops);
    const made = chartsOf(doc.wb.sheets[target])[pasted.id];
    assert.deepEqual([made.description, made.hidden, made.macro], ['alt & text', true, undefined]);
    assert.equal(drawingAnchorNow((doc.wb.sheets[target].drawings || [])[pasted.id], undefined).editAs, 'oneCell');
    // A copy of the copy has no macro to leave behind.
    assert.deepEqual(copyChart(doc.wb, target, pasted.id).dropped, []);
    const out = await save(doc);
    await whole(out, `sheet ${target}: the package is whole`);
    const drawing = await text(out, target === 0 ? 'xl/drawings/drawing1.xml' : 'xl/drawings/drawing2.xml');
    const anchors = anchorsOf(drawing);
    const frame = anchors[anchors.length - 1];
    assert.match(frame, /^<xdr:twoCellAnchor[^>]* editAs="oneCell"/, 'how it moves with the cells');
    assert.match(frame, /<xdr:graphicFrame[^>]* macro=""/, 'no macro');
    assert.equal(/Makro1/.test(frame), false);
    assert.match(frame, /<xdr:graphicFrame[^>]* fPublished="1"/);
    assert.match(frame, /<xdr:cNvPr [^>]*descr="alt &amp; text"/);
    assert.match(frame, /<xdr:cNvPr [^>]*hidden="1"/);
    assert.match(frame, /<xdr:cNvPr [^>]*title="T"/);
    assert.match(frame, /<xdr:clientData fLocksWithSheet="0" fPrintsWithSheet="0"\/>/);
    assert.match(frame, new RegExp(`<xdr:cNvPr [^>]*name="Chart ${target === 0 ? 3 : 1}"`));
    // Its id is its own in the part, and its relationship leads to its own chart part.
    const ids = (drawing.match(/<xdr:cNvPr id="(\d+)"/g) || []).slice();
    assert.equal(new Set(ids).size, ids.length, 'no id twice');
    const rel = (/r:id="([^"]+)"/.exec(frame) as RegExpExecArray)[1];
    assert.match(await text(out, `xl/drawings/_rels/drawing${target + 1}.xml.rels`), new RegExp(`Id="${rel}"[^>]*Target="\\.\\./charts/chart3\\.xml"`));
    if (target === 0) assert.match(anchors[0], /macro="\[0\]!Makro1" fPublished="1"/, 'the frame of the chart is as it was');
    const again = await reopen(out);
    const read = chartsOf(again.sheets[target])[pasted.id];
    assert.deepEqual([read.description, read.hidden, read.macro], ['alt & text', true, undefined]);
    doc.history.undo();
    same(await save(doc), bytes, 'undone');
  }
});

test('chart copy review: a chart the engine wrote keeps its frame in a copy as well', async () => {
  const start = await load(await richBook());
  apply(start, (await import('../src/ops/charts')).addChart(start.wb, 1, { type: 'pie', range: { top: 0, left: 0, bottom: 3, right: 1 }, title: 'Own' }));
  const own = await mod(await save(start), { 'xl/drawings/drawing2.xml': (had) => had.replace('<xdr:graphicFrame macro="">', '<xdr:graphicFrame macro="" fPublished="1">') });
  const doc = await load(own);
  assert.equal(chartsOf(doc.wb.sheets[1])[0].known, 'all');
  apply(doc, duplicateChart(doc.wb, 1, 0).ops);
  const made = chartsOf(doc.wb.sheets[1])[1];
  assert.equal(made.source, undefined, 'written from its model');
  assert.equal(made.frame, '/xl/charts/chart3.xml');
  const out = await save(doc);
  await whole(out, 'the package is whole');
  const anchors = anchorsOf(await text(out, 'xl/drawings/drawing2.xml'));
  assert.equal(anchors.length, 2);
  assert.match(anchors[1], /<xdr:graphicFrame macro="" fPublished="1">/);
  assert.equal(chartsOf((await reopen(out)).sheets[1]).length, 2);
});

// ------------------------------------------------------------------ when a copy is outdated

test('chart copy review: a copy is outdated only by what changes the cells it shows', async () => {
  const doc = await load(await richBook());
  const copy = copyChart(doc.wb, 0, 0);
  const outdated = () => chartCopyOutdated(doc.wb, copy);
  const undo = () => assert.ok(doc.history.undo());
  assert.equal(outdated(), false);
  // Sheets the chart does not name are nothing to it.
  apply(doc, [{ type: 'insert', sheet: 1, axis: 'row', at: 0, count: 1 }]);
  assert.equal(outdated(), false, 'a row on a sheet the chart does not show');
  apply(doc, renameSheet(doc.wb, 1, 'Zzz'));
  assert.equal(outdated(), false, 'the name of a sheet the chart does not show');
  apply(doc, insertSheet(doc.wb, 0, { name: 'New' }));
  assert.equal(outdated(), false, 'a sheet that is added');
  apply(doc, deleteSheet(doc.wb, 0));
  assert.equal(outdated(), false, 'and deleted again');
  apply(doc, deleteSheet(doc.wb, 1));
  assert.equal(outdated(), false, 'a sheet the chart does not show is deleted');
  // Rows of the sheet it shows that leave its cells where they are.
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 50, count: 1 }]);
  assert.equal(outdated(), false, 'a row far below the cells');
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 2 }]);
  assert.equal(outdated(), true, 'rows above the cells');
  apply(doc, [{ type: 'delete', sheet: 0, axis: 'row', at: 0, count: 2 }]);
  assert.equal(outdated(), false, 'and taken out again: the cells are where they were');
  assert.equal(pasteChart(doc.wb, 0, copy).id, 2);
  assert.deepEqual(chartRefs(copy.object.chart), chartRefs(chartsOf(doc.wb.sheets[0])[0].chart), 'the copy says what the chart says now');
  // A row taken out of the middle of the cells and another put in: the chart shows the rows it showed, and so does the copy.
  apply(doc, [{ type: 'delete', sheet: 0, axis: 'row', at: 2, count: 1 }]);
  assert.equal(outdated(), true, 'a row of the cells deleted');
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 2, count: 1 }]);
  assert.equal(outdated(), false, 'and a row put in where it was');
  assert.deepEqual(chartRefs(copy.object.chart), chartRefs(chartsOf(doc.wb.sheets[0])[0].chart));
  undo();
  undo();
  // The cells themselves deleted: rows put in their place are not the cells the chart showed.
  apply(doc, [{ type: 'delete', sheet: 0, axis: 'row', at: 1, count: 3 }]);
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 1, count: 3 }]);
  assert.equal(outdated(), true, 'the cells deleted and rows put in');
  assert.notDeepEqual(chartRefs(copy.object.chart), chartRefs(chartsOf(doc.wb.sheets[0])[0].chart));
  undo();
  undo();
  assert.equal(outdated(), false);
  // As many edits as there were, but another one.
  const later = copyChart(doc.wb, 0, 0);
  undo();
  undo();
  undo();
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'col', at: 0, count: 1 }]);
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 70, count: 1 }]);
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 80, count: 1 }]);
  assert.equal((doc.wb.sheets[0].structuralEdits || []).length, later.state.filter((was) => was.sheet === doc.wb.sheets[0])[0].edits.length, 'as many edits');
  assert.equal(chartCopyOutdated(doc.wb, later), true, 'the edits are others');
  assert.throws(() => pasteChart(doc.wb, 0, later), /changed since it was copied/);
  // The sheet it shows is renamed, or gone.
  const named = await load(await richBook());
  const held = copyChart(named.wb, 0, 0);
  apply(named, renameSheet(named.wb, 0, 'Veri'));
  assert.equal(chartCopyOutdated(named.wb, held), true);
  ok(named.history.undo());
  assert.equal(chartCopyOutdated(named.wb, held), false);
  apply(named, deleteSheet(named.wb, 0));
  assert.equal(chartCopyOutdated(named.wb, held), true);
});

function ok(value: unknown): void {
  assert.ok(value);
}

test('chart copy review: a chart whose references name no sheet is outdated by the rows of its own sheet only', async () => {
  const doc = await load(await richBook());
  const copy = copyChart(doc.wb, 0, 1);
  apply(doc, [{ type: 'insert', sheet: 1, axis: 'row', at: 0, count: 1 }]);
  assert.equal(chartCopyOutdated(doc.wb, copy), false);
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 }]);
  assert.equal(chartCopyOutdated(doc.wb, copy), true);
});

// ------------------------------------------------------------------ relationships

const WEIRD =
  `<Relationship Id="rId7" Type="${REL}/x" Target="chart1.xml"/><Relationship Id="rId8" Type="${REL}/x" Target="/xl/workbook.xml"/><Relationship Id="rId9" Type="${REL}/x" Target="../worksheets/sheet1.xml"/>` +
  `<Relationship Id="rId14" Type="${REL}/x" Target="_rels/chart1.xml.rels"/><Relationship Id="rId15" Type="${REL}/hyperlink" Target="http://example.com/a b" TargetMode="External" Extra="kept"/>`;

test('chart copy review: relationships to itself, to the workbook, to a sheet and out of the file are written as they are meant; one to relationships is left out, and said', async () => {
  const bytes = await mod(await richBook(), { 'xl/charts/_rels/chart1.xml.rels': RELS(STD + WEIRD) });
  const doc = await load(bytes);
  assert.equal(chartsOf(doc.wb.sheets[0])[0].known, 'refs');
  const copy = copyChart(doc.wb, 0, 0);
  assert.deepEqual(copy.dropped, ['relationships']);
  apply(doc, pasteChart(doc.wb, 1, copy).ops);
  apply(doc, removeChart(doc.wb, 0, 0));
  const out = await save(doc);
  await whole(out, 'with the chart gone the package is whole');
  const rels = await text(out, 'xl/charts/_rels/chart3.xml.rels');
  assert.match(rels, /Id="rId7"[^>]*Target="chart3\.xml"/, 'what named the chart names the copy');
  assert.match(rels, /Id="rId8"[^>]*Target="\/xl\/workbook\.xml"/);
  assert.match(rels, /Id="rId9"[^>]*Target="\/xl\/worksheets\/sheet1\.xml"/);
  assert.equal(/rId14/.test(rels), false, 'the relationship to relationships is not in the copy');
  assert.match(rels, /<Relationship Id="rId15" Type="[^"]*\/hyperlink" Target="http:\/\/example\.com\/a b" TargetMode="External" Extra="kept"\/>/, 'what leads out of the file is written as it is');
  assert.equal(chartsOf((await reopen(out)).sheets[1]).length, 1);
});

test('chart copy review: parts that lead to each other in a circle are copied once each', async () => {
  const bytes = await mod(await richBook(), {
    'xl/charts/_rels/style1.xml.rels': RELS(`<Relationship Id="rId1" Type="${REL}/x" Target="colors1.xml"/><Relationship Id="rId2" Type="${REL}/x" Target="chart1.xml"/>`),
    'xl/charts/_rels/colors1.xml.rels': RELS(`<Relationship Id="rId1" Type="${REL}/x" Target="style1.xml"/>`),
  });
  const doc = await load(bytes);
  apply(doc, pasteChart(doc.wb, 1, copyChart(doc.wb, 0, 0)).ops);
  apply(doc, duplicateChart(doc.wb, 0, 0).ops);
  const out = await save(doc);
  await whole(out, 'the package is whole');
  assert.deepEqual(
    names(out).filter((name) => /style|colors/.test(name)),
    ['style1.xml', 'style2.xml', 'style3.xml', 'colors1.xml', 'colors2.xml', 'colors3.xml'].flatMap((name) => [`xl/charts/_rels/${name}.rels`, `xl/charts/${name}`]).sort(),
  );
  assert.match(await text(out, 'xl/charts/_rels/style2.xml.rels'), /Target="colors2\.xml"[\s\S]*Target="chart3\.xml"/);
  assert.match(await text(out, 'xl/charts/_rels/colors3.xml.rels'), /Target="style3\.xml"/);
  apply(doc, removeChart(doc.wb, 0, 0));
  await whole(await save(doc), 'with the chart gone the package is whole');
});

test('chart copy review: what a copy leaves out takes along only the relationships nothing else of the part uses', async () => {
  // The shapes are named by the relationship the picture of the fill is named by.
  const shared = RICH.replace('<c:extLst><c:ext uri="{781A3756', '<c:externalData r:id="rId5"><c:autoUpdate val="0"/></c:externalData><c:userShapes r:id="rId3"/><c:extLst><c:ext uri="{781A3756');
  const stripped = stripChartPart(encoder.encode(shared), ['userShapes', 'externalData']);
  assert.deepEqual(stripped.unrelated, ['rId5'], 'the relationship the picture uses stays');
  assert.equal(new TextDecoder().decode(stripped.bytes), RICH);
  assert.deepEqual(stripChartPart(encoder.encode(RICH), ['userShapes']).unrelated, []);
  const bytes = await mod(await richBook({ first: shared }), { 'xl/charts/_rels/chart1.xml.rels': RELS(`${STD}<Relationship Id="rId5" Type="${REL}/oleObject" Target="file:///C:/x/book.xlsx" TargetMode="External"/>`) });
  const doc = await load(bytes);
  assert.deepEqual(chartsOf(doc.wb.sheets[0])[0].loose, ['externalData', 'userShapes']);
  const pasted = pasteChart(doc.wb, 1, copyChart(doc.wb, 0, 0));
  assert.deepEqual(pasted.dropped, ['externalData', 'userShapes']);
  apply(doc, pasted.ops);
  const out = await save(doc);
  await whole(out, 'the package is whole');
  assert.equal(await text(out, 'xl/charts/chart3.xml'), RICH);
  const rels = await text(out, 'xl/charts/_rels/chart3.xml.rels');
  assert.match(rels, /Id="rId3"[^>]*Target="\.\.\/media\/image2\.png"/);
  assert.equal(/rId5/.test(rels), false);
});

test('chart copy review: a copy whose chart part is gone from the file is said, not written wrongly', async () => {
  const lost = await load(await richBook());
  apply(lost, pasteChart(lost.wb, 1, copyChart(lost.wb, 0, 0)).ops);
  // Somebody takes the part out of what the workbook was read from.
  lost.file.package!.deletePart('/xl/charts/chart1.xml');
  await assert.rejects(() => save(lost), /the chart part \/xl\/charts\/chart1\.xml a chart was copied from is missing/);
  // The chart that was deleted by a save is another matter: the file as it was opened is kept.
  const doc = await load(await richBook());
  apply(doc, pasteChart(doc.wb, 1, copyChart(doc.wb, 0, 0)).ops);
  apply(doc, removeChart(doc.wb, 0, 0));
  await save(doc);
  apply(doc, duplicateChart(doc.wb, 1, 0).ops);
  await whole(await save(doc), 'a copy of the copy, with the chart long gone');
});

// ------------------------------------------------------------------ what goes with a copy

test('chart copy review: that the series of a chart are bound to their order goes with every copy of it', async () => {
  // An entry of the legend that belongs to a trend line: the series are changed, not reordered.
  const fixed = RICH.replace('<c:legend><c:legendPos val="r"/>', '<c:legend><c:legendPos val="r"/><c:legendEntry><c:idx val="5"/><c:delete val="1"/></c:legendEntry>');
  const root = parseXml(SHAPED.replace('<c:legend><c:legendPos val="r"/>', '<c:legend><c:legendPos val="r"/><c:legendEntry><c:idx val="5"/><c:delete val="1"/></c:legendEntry>')).root;
  assert.deepEqual(chartLoose(root, parseChart(root, OFFICE_THEME) as NonNullable<ReturnType<typeof parseChart>>), { loose: ['userShapes'], seriesFixed: true });
  for (const shapes of [false, true]) {
    const doc = await load(await richBook({ shapes, first: shapes ? fixed.replace('<c:extLst><c:ext uri="{781A3756', '<c:userShapes r:id="rId4"/><c:extLst><c:ext uri="{781A3756') : fixed }));
    const source = chartsOf(doc.wb.sheets[0])[0];
    assert.equal(source.seriesFixed, true);
    const copy = copyChart(doc.wb, 0, 0);
    assert.equal(copy.object.seriesFixed, true);
    apply(doc, pasteChart(doc.wb, 1, copy).ops);
    assert.equal(chartsOf(doc.wb.sheets[1])[0].seriesFixed, true);
    assert.equal(sheetCharts(doc.wb, 1)[0].canReorderSeries, false);
    assert.throws(() => updateChart(doc.wb, 1, 0, { series: [{ from: 1 }, { from: 0 }] }), /bound to their order/);
    // What does not touch the order is written, and the entry of the legend stays.
    apply(doc, updateChart(doc.wb, 1, 0, { title: 'Sıra', series: [{ from: 0, color: '#AA0000' }, { from: 1 }] }));
    apply(doc, duplicateChart(doc.wb, 1, 0).ops);
    assert.equal(chartsOf(doc.wb.sheets[1])[1].seriesFixed, true, 'a copy of the copy');
    const out = await save(doc);
    await whole(out, 'the package is whole');
    const made = await text(out, 'xl/charts/chart3.xml');
    assert.match(made, /<c:legendEntry><c:idx val="5"\/><c:delete val="1"\/><\/c:legendEntry>/);
    assert.match(made, /<a:t>Sıra<\/a:t>/);
    assert.equal(chartsOf((await reopen(out)).sheets[1])[0].seriesFixed, true);
  }
});

test('chart copy review: the chart of a chart sheet is not copied, and a part two frames show leaves nothing to a copy', async () => {
  const withSheet = await load(await salesBook({ drawing: frameAt('rIdC1', 'Grafik 1', 1, 2), charts: [FOREIGN], chartSheet: FOREIGN, sheets: [{ name: 'Other', xml: '<sheetData/>' }] }));
  const at = withSheet.wb.sheets.findIndex((sheet) => sheet.kind === 'chartsheet');
  assert.ok(at > 0);
  assert.equal(chartsOf(withSheet.wb.sheets[at]).length, 1);
  assert.match(chartCopyProblem(withSheet.wb, at, 0) || '', /Cannot copy the chart of a chartsheet/);
  assert.throws(() => copyChart(withSheet.wb, at, 0), /Cannot copy the chart of a chartsheet/);
  assert.throws(() => cutChart(withSheet.wb, at, 0), RangeError);
  assert.throws(() => pasteChart(withSheet.wb, at, copyChart(withSheet.wb, 0, 0)), /Cannot change the charts of a chartsheet/);

  // One part with shapes drawn over it, shown by two frames: neither is known, and neither says what a copy could leave out.
  const two = await mod(await richBook({ shapes: true }), {
    'xl/drawings/_rels/drawing1.xml.rels': (had) => had.replace('../charts/chart2.xml', '../charts/chart1.xml'),
    'xl/charts/chart2.xml': null,
    '[Content_Types].xml': (had) => had.replace('<Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>', ''),
  });
  const doc = await load(two);
  const charts = chartsOf(doc.wb.sheets[0]);
  assert.equal(charts.length, 2);
  for (const chart of charts) assert.deepEqual([chart.known, chart.loose], [undefined, undefined]);
  assert.match(chartCopyProblem(doc.wb, 0, 0) || '', /Cannot copy this chart/);
  assert.match(chartCopyProblem(doc.wb, 0, 1) || '', /Cannot copy this chart/);
});

test('chart copy review: the files of these copies are read by a reader that is not this engine', async (t) => {
  if (!pythonHere()) {
    t.skip('python3 is not available');
    return;
  }
  const files: { name: string; bytes: Uint8Array; charts?: number[] }[] = [];
  {
    const doc = await load(await FRAMED());
    apply(doc, duplicateChart(doc.wb, 0, 0).ops);
    apply(doc, pasteChart(doc.wb, 1, copyChart(doc.wb, 0, 0)).ops);
    files.push({ name: 'frames', bytes: await save(doc), charts: [3, 1] });
    apply(doc, removeChart(doc.wb, 0, 0));
    files.push({ name: 'frames, the chart gone', bytes: await save(doc), charts: [2, 1] });
  }
  {
    const doc = await load(await mod(await richBook(), { 'xl/charts/_rels/chart1.xml.rels': RELS(STD + WEIRD) }));
    apply(doc, pasteChart(doc.wb, 1, copyChart(doc.wb, 0, 0)).ops);
    apply(doc, removeChart(doc.wb, 0, 0));
    files.push({ name: 'weird relationships', bytes: await save(doc), charts: [1, 1] });
  }
  {
    const first = Uint8Array.from(RICH.replace('>Old<', '>Café üö<').replace(/^<\?xml[^?]*\?>/, '<?xml version="1.0" encoding="ISO-8859-1"?>'), (ch) => ch.charCodeAt(0));
    const doc = await load(await richBook({ first }));
    apply(doc, pasteChart(doc.wb, 1, copyChart(doc.wb, 0, 0)).ops);
    files.push({ name: 'latin-1', bytes: await save(doc), charts: [2, 1] });
  }
  const found = readByPython(files, 'review');
  assert.deepEqual(found.problems, []);
  assert.equal(found.files, files.length);
});

// The frame of a chart that stands alone in a wrapper of alternatives: the copy has the anchor, not the wrapper.
test('chart copy review: the copy of a chart in a wrapper of alternatives has an anchor of its own', async () => {
  const frame = twoCell([4, 1], [9, 12], chartFrameOf('rIdC1', 'Grafik 1').replace('macro=""', 'macro="" fPublished="1"'));
  const wrapper = `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x14" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">${frame}</mc:Choice><mc:Fallback/></mc:AlternateContent>`;
  const doc = await load(await salesBook({ drawing: wrapper, charts: [FOREIGN], sheets: [{ name: 'Other', xml: '<sheetData/>' }] }));
  apply(doc, duplicateChart(doc.wb, 0, 0).ops);
  apply(doc, pasteChart(doc.wb, 1, copyChart(doc.wb, 0, 0)).ops);
  const out = await save(doc);
  await whole(out, 'the package is whole');
  for (const name of ['xl/drawings/drawing1.xml', 'xl/drawings/drawing2.xml']) {
    const anchors = anchorsOf(await text(out, name));
    assert.match(anchors[anchors.length - 1], /fPublished="1"/, name);
  }
  assert.deepEqual((await reopen(out)).sheets.map((sheet) => chartsOf(sheet).length), [2, 1]);
  same(await part(out, 'xl/charts/chart1.xml'), encoder.encode(FOREIGN), 'the chart keeps its part');
});
