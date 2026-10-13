// What a second review of chart editing found, each pinned down: parts that declare another
// encoding, parts nested too deeply to change, a chart and a sheet that go in one save, legend
// entries of trend lines, anchors that share a wrapper, and what of a part stays byte for byte.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { ZipArchive } from '../src/index';
import { addChart, moveChart, removeChart, sheetCharts, updateChart } from '../src/ops/charts';
import { deleteSheet } from '../src/ops/sheets';
import { FOREIGN, bare, chartsOf, encodingProblem, load, names, packageProblems, reopen, salesBook, same, save, text } from './chart-edit-fixtures';
import { catAx, chartFrameOf, chartXml, group, numRef, ser, sheetData, strRef, valAx, xySer } from './chart-fixtures';
import { sp, twoCell } from './sheet-drawing-fixtures';

const CHART = 'xl/charts/chart1.xml';
const DRAWING = 'xl/drawings/drawing1.xml';
const ALL = { top: 0, left: 0, bottom: 3, right: 2 };
const encoder = new TextEncoder();
const frame = (rel = 'rIdC1', name = 'Grafik 1', row = 1, id = 2) => twoCell([4, row], [9, row + 11], chartFrameOf(rel, name).replace('id="2"', `id="${id}"`));
const whole = async (bytes: Uint8Array, label: string) => assert.deepEqual(await packageProblems(bytes), [], label);
const S = (i = 0, col = 'B', extra: Partial<Parameters<typeof ser>[1]> = {}) => ser(i, { tx: strRef(`Sayfa1!$${col}$1`, ['N']), cat: strRef('Sayfa1!$A$2:$A$4', ['a', 'b', 'c']), val: numRef(`Sayfa1!$${col}$2:$${col}$4`, col === 'B' ? [10, 20, 30] : [30, 20, 10]), ...extra });
const BAR = '<c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>';
const PLAIN = (series = S(), tail?: string) => chartXml({ plot: group('barChart', BAR, series) + catAx() + valAx(), tail });
const deep = (n: number) => `<z:d xmlns:z="urn:z">${'<z:d>'.repeat(n)}${'</z:d>'.repeat(n)}</z:d>`;

test('chart review 2: a part that declares another encoding than UTF-8 is written as UTF-8 and says so', async () => {
  for (const [label, declaration, kept] of [
    ['Latin-1', '<?xml version="1.0" encoding="ISO-8859-1" standalone="yes"?>', false],
    ['ASCII', "<?xml version='1.0' encoding='us-ascii'?>", false],
    ['UTF-8 in small letters', "<?xml version='1.0' encoding='utf-8' standalone='no'?>", true],
  ] as [string, string, boolean][]) {
    const xml = FOREIGN.replace(/^<\?xml[^?]*\?>/, declaration);
    assert.ok(xml.indexOf(declaration) === 0, `${label}: the fixture declares it`);
    const bytes = await salesBook({ drawing: frame(), charts: [xml] });
    const doc = await load(bytes);
    doc.history.apply(updateChart(doc.wb, 0, 0, { title: 'Şırnak ₺ İki' }), undefined, undefined);
    const out = await save(doc);
    await whole(out, `${label}: the part can be read in the encoding it declares`);
    assert.equal(chartsOf((await reopen(out)).sheets[0])[0].chart.title?.text, 'Şırnak ₺ İki', label);
    const written = await text(out, CHART);
    assert.equal(written.indexOf(declaration) === 0, kept, `${label}: the declaration ${kept ? 'stays as it is' : 'names UTF-8 now'}`);
    if (!kept) assert.match(written.slice(0, 80), /encoding=["']UTF-8["']/, label);
  }
  // The drawing part is written the same way.
  const base = await salesBook({ drawing: frame(), charts: [FOREIGN] });
  const zip = ZipArchive.read(base);
  const drawing = (await text(base, DRAWING)).replace(/^<\?xml[^?]*\?>\s*/, '');
  zip.set(DRAWING, encoder.encode(`<?xml version="1.0" encoding="us-ascii"?>${drawing}`));
  const doc = await load(await zip.write());
  doc.history.apply(addChart(doc.wb, 0, { type: 'column', range: ALL, title: 'Şırnak' }), undefined, undefined);
  doc.history.apply(moveChart(doc.wb, 0, 0, { top: 30, left: 0, bottom: 40, right: 5 }), undefined, undefined);
  const out = await save(doc);
  await whole(out, 'a drawing part that declared ASCII');
  assert.match((await text(out, DRAWING)).slice(0, 60), /^<\?xml version="1.0" encoding="UTF-8"\?>/);
  assert.equal(chartsOf((await reopen(out)).sheets[0]).length, 2);
  assert.equal(encodingProblem(encoder.encode('<?xml version="1.0" encoding="us-ascii"?><a>Ş</a>')), 'declares the encoding us-ascii and holds bytes that are not ASCII', 'the check itself');
  assert.equal(encodingProblem(encoder.encode('<?xml version="1.0" encoding="utf-8"?><a>Ş</a>')), undefined);
});

test('chart review 2: a chart part nested too deeply to be changed is moved and deleted only; a drawing part nested so takes no change of its charts', async () => {
  for (const depth of [300, 2100]) {
    const xml = FOREIGN.replace('</c:chartSpace>', `<c:extLst><c:ext uri="{deep}">${deep(depth)}</c:ext></c:extLst></c:chartSpace>`);
    const doc = await load(await salesBook({ drawing: frame(), charts: [xml] }));
    assert.equal(chartsOf(doc.wb.sheets[0])[0].known, undefined, `depth ${depth}: the chart is not one that is changed`);
    assert.equal(doc.wb.chartsFollow, undefined);
    assert.throws(() => updateChart(doc.wb, 0, 0, { title: 'T' }), /only moved and deleted/);
    doc.history.apply(moveChart(doc.wb, 0, 0, { top: 30, left: 0, bottom: 40, right: 5 }), undefined, undefined);
    await whole(await save(doc), `depth ${depth}: moved and saved`);
  }
  // Not so deep: changed like any other.
  const fine = await load(await salesBook({ drawing: frame(), charts: [FOREIGN.replace('</c:chartSpace>', `<c:extLst><c:ext uri="{deep}">${deep(40)}</c:ext></c:extLst></c:chartSpace>`)] }));
  fine.history.apply(updateChart(fine.wb, 0, 0, { title: 'T' }), undefined, undefined);
  await whole(await save(fine), 'forty levels are changed');
  // A shape whose properties are nested deeply stands in the drawing part beside the chart.
  const shape = twoCell([0, 30], [2, 33], sp({ name: 'S' }).replace('<xdr:cNvPr id="3" name="S"/>', `<xdr:cNvPr id="3" name="S"><a:extLst><a:ext uri="{deep}">${deep(2100)}</a:ext></a:extLst></xdr:cNvPr>`));
  assert.ok(shape.indexOf('urn:z') > 0, 'the fixture holds the deep shape');
  const doc = await load(await salesBook({ drawing: frame() + shape, charts: [FOREIGN] }));
  assert.equal(doc.wb.sheets[0].drawings?.length, 2);
  assert.throws(() => moveChart(doc.wb, 0, 0, { top: 40, left: 1, bottom: 50, right: 6 }), /Cannot change the charts of this sheet/);
  assert.throws(() => removeChart(doc.wb, 0, 0), /Cannot change the charts of this sheet/);
  assert.throws(() => addChart(doc.wb, 0, { type: 'column', range: ALL }), /Cannot change the charts of this sheet/);
  assert.equal(sheetCharts(doc.wb, 0)[0].editable, 'parts', 'the chart itself is changed in its own part');
  // A drawing part that cannot be read: nothing is shown of it, and no chart is added to it.
  const zip = ZipArchive.read(await salesBook({ drawing: frame(), charts: [FOREIGN] }));
  zip.set(DRAWING, encoder.encode('<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><broken'));
  const unread = await load(await zip.write());
  assert.equal(unread.wb.sheets[0].drawings, undefined);
  assert.throws(() => addChart(unread.wb, 0, { type: 'column', range: ALL }), /Cannot change the charts of this sheet: its drawing cannot be read/);
  doc.history.apply(updateChart(doc.wb, 0, 0, { title: 'T' }), undefined, undefined);
  await whole(await save(doc), 'the chart part is written, the drawing part is as it was');
});

test('chart review 2: a drawing part two sheets show takes no change of its charts', async () => {
  const bytes = await salesBook({ drawing: frame(), charts: [PLAIN()], sheets: [{ name: 'Other', xml: `${sheetData([['x', 1]])}<drawing r:id="rId9"/>` }] });
  const zip = ZipArchive.read(bytes);
  zip.set('xl/worksheets/_rels/sheet2.xml.rels', encoder.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>'));
  const doc = await load(await zip.write());
  for (const sheet of [0, 1]) {
    assert.throws(() => moveChart(doc.wb, sheet, 0, { top: 20, left: 0, bottom: 30, right: 4 }), /Cannot change the charts of this sheet: its drawing is shared with another sheet/);
    assert.throws(() => addChart(doc.wb, sheet, { type: 'line', series: [{ values: 'Sayfa1!$B$2:$B$4' }] }), /Cannot change the charts of this sheet/);
    assert.throws(() => removeChart(doc.wb, sheet, 0), /Cannot change the charts of this sheet/);
  }
});

const RELS = (inner: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;

/** Two sheets with a chart each; both charts lead to one style part and one colours part. */
async function twoSheets(): Promise<Uint8Array> {
  const base = await salesBook({ drawing: frame(), charts: [PLAIN()], sheets: [{ name: 'Other', xml: `${sheetData([['x', 1]])}<drawing r:id="rId7"/>` }] });
  const zip = ZipArchive.read(base);
  zip.set('xl/worksheets/_rels/sheet2.xml.rels', encoder.encode(RELS('<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing2.xml"/>')));
  zip.set('xl/drawings/drawing2.xml', encoder.encode((await text(base, DRAWING)).replace('rIdC1', 'rId1')));
  zip.set('xl/drawings/_rels/drawing2.xml.rels', encoder.encode(RELS('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="/xl/charts/chart2.xml"/>')));
  zip.set('xl/charts/chart2.xml', encoder.encode(PLAIN()));
  zip.set('xl/charts/style1.xml', encoder.encode('<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="201"/>'));
  zip.set('xl/charts/colors1.xml', encoder.encode('<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="10"/>'));
  for (const n of [1, 2]) {
    zip.set(`xl/charts/_rels/chart${n}.xml.rels`, encoder.encode(RELS('<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/>')));
  }
  zip.set(
    '[Content_Types].xml',
    encoder.encode(
      (await text(base, '[Content_Types].xml')).replace(
        '</Types>',
        '<Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' +
          '<Override PartName="/xl/charts/style1.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/><Override PartName="/xl/charts/colors1.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/></Types>',
      ),
    ),
  );
  return zip.write();
}

test('chart review 2: a chart that goes and a sheet that goes in one save take the parts only they led to', async () => {
  const bytes = await twoSheets();
  await whole(bytes, 'the fixture is whole');
  const parts = (out: Uint8Array) => names(out).filter((name) => /chart|drawing/.test(name));
  for (const order of ['chart, sheet, save', 'sheet, chart, save', 'chart, save, sheet, save', 'sheet, save, chart, save']) {
    const doc = await load(bytes);
    assert.equal(doc.wb.chartsFollow, true);
    let out: Uint8Array | undefined;
    for (const step of order.split(', ')) {
      // The chart of the first sheet, and the other sheet with its chart.
      if (step === 'chart') doc.history.apply(removeChart(doc.wb, 0, 0), undefined, undefined);
      else if (step === 'sheet') doc.history.apply(deleteSheet(doc.wb, 1), undefined, undefined);
      else out = await save(doc);
    }
    await whole(out as Uint8Array, order);
    assert.deepEqual(parts(out as Uint8Array), [], `${order}: nothing of the charts is left`);
    doc.history.undo();
    const back = await save(doc);
    await whole(back, `${order}, one step back`);
    assert.ok(parts(back).indexOf('xl/charts/style1.xml') >= 0 && parts(back).indexOf('xl/charts/colors1.xml') >= 0, `${order}: with one of the two back, the parts they share are back`);
    doc.history.undo();
    same(await save(doc), bytes, `${order}: everything undone`);
  }
  // Only the sheet goes: the chart that stays still leads to the shared parts.
  const doc = await load(bytes);
  doc.history.apply(deleteSheet(doc.wb, 1), undefined, undefined);
  const out = await save(doc);
  await whole(out, 'the sheet alone');
  assert.deepEqual(parts(out), ['xl/charts/_rels/chart1.xml.rels', 'xl/charts/chart1.xml', 'xl/charts/colors1.xml', 'xl/charts/style1.xml', 'xl/drawings/_rels/drawing1.xml.rels', 'xl/drawings/drawing1.xml']);
});

test('chart review 2: a part that a frame which stays still names is never taken out', async () => {
  // Two frames name one chart by the same relationship; one of them goes.
  const bytes = await salesBook({ drawing: frame('rIdC1', 'A') + frame('rIdC1', 'B', 20, 3), charts: [FOREIGN] });
  const doc = await load(bytes);
  doc.history.apply(removeChart(doc.wb, 0, 0), undefined, undefined);
  const out = await save(doc);
  await whole(out, 'one frame gone');
  assert.ok(names(out).indexOf(CHART) >= 0, 'the chart part stays');
  assert.match(await text(out, 'xl/drawings/_rels/drawing1.xml.rels'), /Id="rIdC1"/, 'and the relationship the other frame names');
  assert.equal(chartsOf((await reopen(out)).sheets[0]).length, 1);
  doc.history.undo();
  same(await save(doc), bytes, 'undone');
});

test('chart review 2: a change that leaves the series in their order leaves the entries of the legend alone, those of trend lines too', async () => {
  const trend = '<c:trendline><c:trendlineType val="linear"/></c:trendline>';
  const xml = PLAIN(S(0, 'B', { props: trend }) + S(1, 'C'), '<c:legend><c:legendPos val="r"/><c:legendEntry><c:idx val="2"/><c:delete val="1"/></c:legendEntry><c:legendEntry><c:idx val="1"/><c:delete val="1"/></c:legendEntry><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>');
  const doc = await load(await salesBook({ drawing: frame(), charts: [xml] }));
  const object = chartsOf(doc.wb.sheets[0])[0];
  assert.deepEqual([object.known, object.seriesFixed, object.chart.legend?.deleted], ['refs', true, [2, 1]], 'the entry behind the series is that of the trend line');
  doc.history.apply(updateChart(doc.wb, 0, 0, { series: [{ from: 0, color: '#112233' }, { from: 1, name: 'Renamed' }] }), undefined, undefined);
  assert.deepEqual(chartsOf(doc.wb.sheets[0])[0].chart.legend?.deleted, [2, 1], 'the model keeps every entry');
  const out = await save(doc);
  assert.deepEqual(bare(chartsOf((await reopen(out)).sheets[0])[0].chart), bare(chartsOf(doc.wb.sheets[0])[0].chart), 'and so does the part');
  assert.throws(() => updateChart(doc.wb, 0, 0, { series: [{ from: 1 }, { from: 0 }] }), /Cannot add, remove or reorder the series of this chart/);
  // What binds a part to the order of its series, one by one.
  const bound = async (chart: string) => chartsOf((await reopen(await salesBook({ drawing: frame(), charts: [chart] }))).sheets[0])[0].seriesFixed;
  assert.equal(await bound(PLAIN(S() + S(1, 'C'))), undefined, 'nothing');
  assert.equal(await bound(PLAIN(S() + S(1, 'C')).replace('<c:plotArea>', '<c:pivotFmts><c:pivotFmt><c:idx val="0"/></c:pivotFmt></c:pivotFmts><c:plotArea>')), true, 'formats of a pivot chart');
  assert.equal(await bound(PLAIN(S(0, 'B', { props: trend }) + S(1, 'C'), '<c:legend><c:legendPos val="r"/><c:legendEntry><c:idx val="0"/><c:delete val="1"/></c:legendEntry></c:legend><c:plotVisOnly val="1"/>')), true, 'a trend line and an entry of the legend, whichever it is');
  assert.equal(await bound(PLAIN(S(0, 'B', { props: trend }) + S(1, 'C'), '<c:legend><c:legendPos val="r"/></c:legend><c:plotVisOnly val="1"/>')), undefined, 'a trend line without entries');
  assert.equal(await bound(PLAIN(S() + S(1, 'C'), '<c:legend><c:legendPos val="r"/><c:legendEntry><c:idx val="2"/><c:delete val="1"/></c:legendEntry></c:legend><c:plotVisOnly val="1"/>')), true, 'an entry behind the series');
});

test('chart review 2: a chart that shares its wrapper with another drawing is not taken out; one alone in its wrapper goes with it', async () => {
  const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
  const shape = twoCell([0, 30], [2, 33], sp({ name: 'Beside' }));
  const picture = twoCell([4, 1], [9, 12], sp({ name: 'Fallback' }));
  const together = `<mc:AlternateContent ${MC}><mc:Choice Requires="x" xmlns:x="urn:x">${frame()}${shape}</mc:Choice><mc:Fallback>${picture}</mc:Fallback></mc:AlternateContent>`;
  const doc = await load(await salesBook({ drawing: together, charts: [FOREIGN] }));
  assert.deepEqual(doc.wb.sheets[0].drawings?.map((drawing) => drawing.object.type), ['chart', 'shape']);
  assert.throws(() => removeChart(doc.wb, 0, 0), /Cannot delete this chart: it shares its place in the drawing with another object/);
  doc.history.apply(moveChart(doc.wb, 0, 0, { top: 40, left: 0, bottom: 50, right: 5 }), undefined, undefined);
  const moved = await save(doc);
  await whole(moved, 'moved');
  assert.deepEqual((await reopen(moved)).sheets[0].drawings?.map((drawing) => drawing.object.type), ['chart', 'shape'], 'both are there');
  // The writer holds as well, for a model that does not say so: the other drawing is never taken along.
  const unsaid = await load(await salesBook({ drawing: together, charts: [FOREIGN] }));
  for (const drawing of unsaid.wb.sheets[0].drawings || []) delete drawing.wrapped;
  unsaid.history.apply(removeChart(unsaid.wb, 0, 0), undefined, undefined);
  await assert.rejects(save(unsaid), /shares its wrapper with one that stays/);
  // Alone in its wrapper: the wrapper goes, fallback and all; the shape beside the wrapper stays.
  const alone = `<mc:AlternateContent ${MC}><mc:Choice Requires="x" xmlns:x="urn:x">${frame()}</mc:Choice><mc:Fallback>${picture}</mc:Fallback></mc:AlternateContent>${shape}`;
  const other = await load(await salesBook({ drawing: alone, charts: [FOREIGN] }));
  other.history.apply(removeChart(other.wb, 0, 0), undefined, undefined);
  const out = await save(other);
  await whole(out, 'removed');
  const xml = await text(out, DRAWING);
  assert.ok(xml.indexOf('AlternateContent') < 0 && xml.indexOf('name="Fallback"') < 0 && xml.indexOf(shape) > 0, 'the wrapper is gone with what stood in for the chart; the shape is as it was');
});

test('chart review 2: what is not changed in a part keeps its bytes: what follows the root, start tags, CDATA; a prefix bound anew inside the part is the one that counts there', async () => {
  const odd = FOREIGN.replace('<c:barChart>', '<c:barChart\t >')
    .replace('<a:t>Old</a:t>', '<a:t><![CDATA[Old <c:ser> & ]]></a:t>')
    .replace('<c:gapWidth', '<![CDATA[<c:ser><c:idx val="7"/></c:ser>]]><c:gapWidth')
    .concat('<!-- trail <c:ser/> --><?end x?>\n');
  const bytes = await salesBook({ drawing: frame(), charts: [odd] });
  const doc = await load(bytes);
  assert.equal(chartsOf(doc.wb.sheets[0])[0].chart.title?.text, 'Old <c:ser> & ', 'the title is read out of its CDATA');
  doc.history.apply(updateChart(doc.wb, 0, 0, { legend: 'b', series: [{ from: 1 }] }), undefined, undefined);
  const out = await save(doc);
  await whole(out, 'saved');
  const xml = await text(out, CHART);
  assert.ok(xml.endsWith('</c:chartSpace><!-- trail <c:ser/> --><?end x?>\n'), 'what follows the root is as it was');
  assert.ok(xml.indexOf('<c:barChart\t >') > 0, 'the start tag of the chart type whose series changed is as it was');
  assert.ok(xml.indexOf('<a:t><![CDATA[Old <c:ser> & ]]></a:t>') > 0 && xml.indexOf('<![CDATA[<c:ser><c:idx val="7"/></c:ser>]]><c:gapWidth') > 0, 'CDATA is as it was');
  assert.equal(chartsOf((await reopen(out)).sheets[0])[0].chart.groups[0].series.length, 1);
  // The plot area binds "c" to something else and writes charts with "x": what is put in there is written with "x".
  const rebound = FOREIGN.replace(/<c:plotArea>([\s\S]*)<\/c:plotArea>/, (_m, inner: string) => `<x:plotArea xmlns:x="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:c="urn:other">${inner.replace(/<(\/?)c:/g, '<$1x:')}</x:plotArea>`);
  const again = await load(await salesBook({ drawing: frame(), charts: [rebound] }));
  assert.equal(chartsOf(again.wb.sheets[0])[0].known, 'refs');
  again.history.apply(updateChart(again.wb, 0, 0, { dataLabels: true, series: [{ from: 0, color: '#010203' }, { from: 1 }, { values: 'Sayfa1!$C$2:$C$4', name: 'More' }] }), undefined, undefined);
  const written = await save(again);
  await whole(written, 'every element stands in the namespace of charts');
  assert.ok((await text(written, CHART)).indexOf('<x:ser><x:idx val="2"/>') > 0, 'the series that was added is written with the prefix in scope');
  assert.deepEqual(bare(chartsOf((await reopen(written)).sheets[0])[0].chart), bare(chartsOf(again.wb.sheets[0])[0].chart));
  // A series binds "a" to something else and drawings to "d": the nearest declaration counts, so the colour it is given is written with "d".
  const inner = FOREIGN.replace(/<c:ser>/g, '<c:ser xmlns:a="urn:other" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main">').replace(/(<c:ser [^>]*>)([\s\S]*?)(<\/c:ser>)/g, (_m, open: string, body: string, close: string) => open + body.replace(/<(\/?)a:/g, '<$1d:') + close);
  assert.ok(inner.indexOf('xmlns:a="urn:other"') > 0, 'the fixture binds the prefix anew');
  const third = await load(await salesBook({ drawing: frame(), charts: [inner] }));
  third.history.apply(updateChart(third.wb, 0, 0, { series: [{ from: 0, color: '#010203' }, { from: 1, color: '#040506' }] }), undefined, undefined);
  const coloured = await save(third);
  await whole(coloured, 'a prefix bound anew in a series');
  const colouredXml = await text(coloured, CHART);
  assert.ok(colouredXml.indexOf('<d:srgbClr val="010203"/>') > 0 && colouredXml.indexOf('<a:srgbClr val="010203"/>') < 0, 'the colour stands in the namespace of drawings');
  assert.deepEqual(bare(chartsOf((await reopen(coloured)).sheets[0])[0].chart), bare(chartsOf(third.wb.sheets[0])[0].chart));
  // A ">" inside an attribute does not end its tag: the element keeps its bytes, quotes and all.
  const quoted = FOREIGN.replace(/<c:idx val="0"\/>/, `<c:idx val='0' note='a>b"c/>'/>`).replace('<c:barDir val="col"/>', '<c:barDir\n   val = "col"\n/>');
  assert.ok(quoted.indexOf(`note='a>b"c/>'`) > 0, 'the fixture holds the attribute');
  const fourth = await load(await salesBook({ drawing: frame(), charts: [quoted] }));
  fourth.history.apply(updateChart(fourth.wb, 0, 0, { title: 'T', legend: 'b' }), undefined, undefined);
  const kept = await text(await save(fourth), CHART);
  assert.ok(kept.indexOf(`<c:idx val='0' note='a>b"c/>'/>`) > 0 && kept.indexOf('<c:barDir\n   val = "col"\n/>') > 0, 'what was not changed is spelled as it was');
});

test('chart review 2: a title or a legend inside a wrapper makes the chart one that is not changed; the axes of a scatter chart are told apart by the chart type', async () => {
  const wrap = (xml: string, what: string) =>
    xml.replace(new RegExp(`(<c:${what}>[\\s\\S]*?</c:${what}>)`), '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">$1</mc:Choice><mc:Fallback>$1</mc:Fallback></mc:AlternateContent>');
  for (const what of ['title', 'legend']) {
    const doc = await load(await salesBook({ drawing: frame(), charts: [wrap(FOREIGN, what)] }));
    assert.equal(chartsOf(doc.wb.sheets[0])[0].known, undefined, `a ${what} in a wrapper`);
    assert.throws(() => updateChart(doc.wb, 0, 0, { title: 'T' }), /only moved and deleted/);
  }
  // As openpyxl writes a scatter chart: both axes say they stand on the left; the first the chart type names is the horizontal one.
  const axis = (id: number, cross: number) => `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="${cross}"/></c:valAx>`;
  const scatter = chartXml({ plot: `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${xySer(0, { cat: numRef('Sayfa1!$B$2:$B$4', [1, 2, 3]), val: numRef('Sayfa1!$C$2:$C$4', [7, 8, 9]) })}<c:axId val="10"/><c:axId val="100"/></c:scatterChart>${axis(10, 100)}${axis(100, 10)}` });
  const doc = await load(await salesBook({ drawing: frame(), charts: [scatter] }));
  doc.history.apply(updateChart(doc.wb, 0, 0, { axisTitles: { category: 'X', value: 'Y' }, axisBounds: { min: 0, max: 50 } }), undefined, undefined);
  const axes = chartsOf(doc.wb.sheets[0])[0].chart.axes;
  assert.deepEqual(axes.map((item) => [item.id, item.title?.text, item.max]), [[10, 'X', undefined], [100, 'Y', 50]]);
  const out = await save(doc);
  await whole(out, 'saved');
  assert.deepEqual(sheetCharts(await reopen(out), 0)[0].axisTitles, { category: 'X', value: 'Y' });
});
