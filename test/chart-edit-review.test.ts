// What a review of chart editing found, each pinned down: a note and a chart on one sheet,
// charts of parts written with other prefixes and namespaces, changes a part cannot take,
// what depends on the order of the series, parts shared between sheets, and the cases that
// keep the rows, columns and sheets of a workbook where they are.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { ZipArchive } from '../src/index';
import { addChart, chartProblem, removeChart, sheetCharts, updateChart, type ChartPatch } from '../src/ops/charts';
import { setComment } from '../src/ops/notes';
import { FOREIGN, bare, chartsOf, load, names, packageProblems, reopen, salesBook, same, save, text, type Loaded } from './chart-edit-fixtures';
import { applyOp } from '../src/ops/ops';
import { catAx, chartFrameOf, chartXml, group, numRef, ser, sheetData, strRef, valAx, xyAxes, xySer } from './chart-fixtures';
import { twoCell } from './sheet-drawing-fixtures';

const ALL = { top: 0, left: 0, bottom: 3, right: 2 };
const whole = async (bytes: Uint8Array, label: string) => assert.deepEqual(await packageProblems(bytes), [], label);
const note = (doc: Loaded) => doc.history.apply(setComment(doc.wb, 0, 1, 1, 'a note'), undefined, undefined);
const noNote = (doc: Loaded) => doc.history.apply(setComment(doc.wb, 0, 1, 1, undefined), undefined, undefined);
const chart = (doc: Loaded) => doc.history.apply(addChart(doc.wb, 0, { type: 'column', range: ALL }), undefined, undefined);

test('chart review: a note and a chart added to one sheet get relationships with ids of their own, in whatever order and however many saves', async () => {
  const bytes = await salesBook();
  for (const order of ['note, chart, save', 'chart, note, save', 'note, save, chart, save', 'chart, save, note, save']) {
    const doc = await load(bytes);
    let out: Uint8Array | undefined;
    for (const step of order.split(', ')) {
      if (step === 'note') note(doc);
      else if (step === 'chart') chart(doc);
      else out = await save(doc);
    }
    await whole(out as Uint8Array, order);
    const again = await reopen(out as Uint8Array);
    assert.equal(chartsOf(again.sheets[0]).length, 1, `${order}: the chart is read back`);
    assert.equal((again.sheets[0].comments || []).length, 1, `${order}: the note is read back`);
    // One of the two goes again: nothing of it is left behind, and the other stays.
    noNote(doc);
    const without = await save(doc);
    await whole(without, `${order}, then without the note`);
    assert.deepEqual(names(without).filter((name) => /comments|vml/.test(name)), [], `${order}: no part of the note is left`);
    assert.equal(chartsOf((await reopen(without)).sheets[0]).length, 1);
    while (doc.history.undo());
    same(await save(doc), bytes, `${order}: everything undone`);
  }
});

test('chart review: a sheet that has relationships already gives the note and the chart the next free ids', async () => {
  const zip = ZipArchive.read(await salesBook());
  const encoder = new TextEncoder();
  zip.set('xl/worksheets/sheet1.xml', encoder.encode((await text(await zip.write(), 'xl/worksheets/sheet1.xml')).replace('<pageMargins', '<hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks><pageMargins')));
  zip.set(
    'xl/worksheets/_rels/sheet1.xml.rels',
    encoder.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://example.com/" TargetMode="External"/></Relationships>'),
  );
  const doc = await load(await zip.write());
  note(doc);
  chart(doc);
  const out = await save(doc);
  await whole(out, 'a link, a note and a chart');
  assert.equal(((await text(out, 'xl/worksheets/_rels/sheet1.xml.rels')).match(/<Relationship /g) || []).length, 4);
});

const frame = (rel = 'rIdC1', name = 'Grafik 1', row = 1) => twoCell([4, row], [9, row + 11], chartFrameOf(rel, name));
const CHART = 'xl/charts/chart1.xml';
const STRICT: [RegExp, string][] = [
  [/http:\/\/schemas\.openxmlformats\.org\/drawingml\/2006\/chart/g, 'http://purl.oclc.org/ooxml/drawingml/chart'],
  [/http:\/\/schemas\.openxmlformats\.org\/drawingml\/2006\/main/g, 'http://purl.oclc.org/ooxml/drawingml/main'],
  [/http:\/\/schemas\.openxmlformats\.org\/drawingml\/2006\/spreadsheetDrawing/g, 'http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing'],
];
const strict = (xml: string) => STRICT.reduce((out, [from, to]) => out.replace(from, to), xml);

/** A workbook whose chart part and drawing part are rewritten by `chart` and `drawing` before it is opened. */
async function rewritten(spec: Parameters<typeof salesBook>[0], chart: (xml: string) => string, drawing: (xml: string) => string = (xml) => xml): Promise<Uint8Array> {
  const bytes = await salesBook(spec);
  const zip = ZipArchive.read(bytes);
  const encoder = new TextEncoder();
  zip.set(CHART, encoder.encode(chart(await text(bytes, CHART))));
  zip.set('xl/drawings/drawing1.xml', encoder.encode(drawing(await text(bytes, 'xl/drawings/drawing1.xml'))));
  return zip.write();
}

const SPELLINGS: [string, (xml: string) => string, (xml: string) => string][] = [
  // As openpyxl writes every chart: the namespace of charts is the default one, no prefix "c" is declared.
  ['the default namespace', (xml) => xml.replace(/<(\/?)c:/g, '<$1').replace('xmlns:c=', 'xmlns='), (xml) => xml],
  ['other prefixes', (xml) => xml.replace(/<(\/?)c:/g, '<$1cc:').replace('xmlns:c=', 'xmlns:cc=').replace(/<(\/?)a:/g, '<$1dm:').replace('xmlns:a=', 'xmlns:dm='), (xml) => xml.replace(/<(\/?)xdr:/g, '<$1x:').replace('xmlns:xdr=', 'xmlns:x=').replace(/<(\/?)a:/g, '<$1dm:').replace('xmlns:a=', 'xmlns:dm=')],
  ['the prefix "a" bound to something else', (xml) => xml.replace(/<(\/?)a:/g, '<$1dm:').replace('xmlns:a=', 'xmlns:dm=').replace('<c:chartSpace ', '<c:chartSpace xmlns:a="urn:example:other" '), (xml) => xml],
  ['the strict namespaces', strict, strict],
];

test('chart review: a chart part written with other prefixes or namespaces is changed in its own: every prefix is bound, one namespace of charts', async () => {
  for (const [label, chartOf, drawingOf] of SPELLINGS) {
    const bytes = await rewritten({ drawing: frame(), charts: [FOREIGN] }, chartOf, drawingOf);
    assert.deepEqual(await packageProblems(bytes), [], `${label}: the fixture is whole`);
    const steps: [string, (doc: Loaded) => void][] = [
      ['a row inserted above the data', (doc) => doc.history.apply([{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 }], undefined, undefined)],
      ['the legend moved', (doc) => doc.history.apply(updateChart(doc.wb, 0, 0, { legend: 'b' }), undefined, undefined)],
      [
        'everything changed',
        (doc) =>
          doc.history.apply(
            updateChart(doc.wb, 0, 0, { title: 'New\nlines', legend: 'none', axisTitles: { category: 'C', value: 'V' }, axisBounds: { min: 0, max: 90 }, dataLabels: true, grouping: 'stacked', series: [{ from: 1, color: '#112233' }, { from: 0, name: 'N' }, { values: 'Sayfa1!$C$2:$C$4', color: '#445566' }] }),
            undefined,
            undefined,
          ),
      ],
      ['a chart added beside it', (doc) => doc.history.apply(addChart(doc.wb, 0, { type: 'line', range: ALL }), undefined, undefined)],
    ];
    for (const [what, act] of steps) {
      const doc = await load(bytes);
      assert.equal(chartsOf(doc.wb.sheets[0])[0].known, 'refs', `${label}: the chart is one that is changed in its part`);
      act(doc);
      const out = await save(doc);
      await whole(out, `${label}, ${what}`);
      const again = await reopen(out);
      assert.deepEqual(bare(chartsOf(again.sheets[0])[0].chart), bare(chartsOf(doc.wb.sheets[0])[0].chart), `${label}, ${what}: what is read back is the model`);
      assert.equal(chartsOf(again.sheets[0]).length, chartsOf(doc.wb.sheets[0]).length, `${label}, ${what}: every chart is read back`);
      doc.history.undo();
      same(await save(doc), bytes, `${label}, ${what}: undone`);
    }
  }
});

const S = (i: number, col: string) => ser(i, { tx: strRef(`Sayfa1!$${col}$1`, [col]), cat: strRef('Sayfa1!$A$2:$A$4', ['x', 'y', 'z']), val: numRef(`Sayfa1!$${col}$2:$${col}$4`, [5, 6, 7]) });
const OTHER_KINDS: [string, string][] = [
  ['radar', chartXml({ plot: group('radarChart', '<c:radarStyle val="marker"/><c:varyColors val="0"/>', S(0, 'B')) + catAx() + valAx() })],
  ['bubble', chartXml({ plot: group('bubbleChart', '<c:varyColors val="0"/>', xySer(0, { cat: numRef('Sayfa1!$B$2:$B$4', [1, 2, 3]), val: numRef('Sayfa1!$C$2:$C$4', [7, 8, 9]), tail: `<c:bubbleSize>${numRef('Sayfa1!$B$2:$B$4', [8, 9, 10])}</c:bubbleSize>` })) + xyAxes() })],
  ['stock', chartXml({ plot: group('stockChart', '', S(0, 'B') + S(1, 'C') + S(2, 'B'), '<c:hiLowLines/>') + catAx() + valAx() })],
  ['bar3D', chartXml({ plot: group('bar3DChart', '<c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>', S(0, 'B'), '<c:gapWidth val="150"/><c:shape val="box"/>') + catAx() + valAx() })],
  ['ofPie', chartXml({ plot: group('ofPieChart', '<c:ofPieType val="pie"/><c:varyColors val="1"/>', S(0, 'B'), '<c:gapWidth val="100"/><c:secondPieSize val="75"/>', false) })],
  ['pie', chartXml({ plot: group('pieChart', '<c:varyColors val="1"/>', S(0, 'B'), '<c:firstSliceAng val="0"/>', false) })],
  ['scatter', chartXml({ plot: group('scatterChart', '<c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>', xySer(0, { cat: numRef('Sayfa1!$B$2:$B$4', [1, 2, 3]), val: numRef('Sayfa1!$C$2:$C$4', [7, 8, 9]) })) + xyAxes() })],
];
const PATCHES: [string, ChartPatch][] = [
  ['a series added', { series: [{ from: 0 }, { values: 'Sayfa1!$C$2:$C$4', name: 'More' }] }],
  ['a range', { range: ALL }],
  ['a smaller range', { range: { top: 0, left: 0, bottom: 3, right: 1 } }],
  ['data labels', { dataLabels: true }],
  ['a grouping', { grouping: 'stacked' }],
  ['a colour', { series: [{ from: 0, color: '#123456' }] }],
  ['a title', { title: 'T', legend: 'b' }],
];

test('chart review: no change the writer cannot write passes as an operation: what is accepted is saved and read back as the model', async () => {
  const refused: string[] = [];
  for (const [kind, xml] of OTHER_KINDS) {
    const bytes = await salesBook({ drawing: frame(), charts: [xml] });
    for (const [what, patch] of PATCHES) {
      const doc = await load(bytes);
      const problem = chartProblem(doc.wb, 0, 0, patch);
      let ops;
      try {
        ops = updateChart(doc.wb, 0, 0, patch);
      } catch (error) {
        assert.ok(error instanceof RangeError, `${kind}, ${what}: what is refused is refused with a RangeError: ${error}`);
        assert.equal(problem, (error as Error).message, `${kind}, ${what}: chartProblem() says beforehand what updateChart() refuses`);
        refused.push(`${kind}: ${what}`);
        continue;
      }
      assert.equal(problem, undefined, `${kind}, ${what}: chartProblem() finds nothing wrong with what updateChart() takes`);
      if (ops.length === 0) continue;
      doc.history.apply(ops, undefined, undefined);
      const out = await save(doc);
      await whole(out, `${kind}, ${what}`);
      assert.deepEqual(bare(chartsOf((await reopen(out)).sheets[0])[0].chart), bare(chartsOf(doc.wb.sheets[0])[0].chart), `${kind}, ${what}: what is read back is the model`);
    }
  }
  assert.deepEqual(refused, [
    'radar: a series added',
    'radar: a range',
    'radar: a grouping',
    'bubble: a series added',
    'bubble: a range',
    'bubble: a grouping',
    'stock: a series added',
    'stock: a range',
    'stock: a smaller range',
    'stock: a grouping',
    'stock: a colour',
    'ofPie: a grouping',
    'pie: a grouping',
    'scatter: a grouping',
  ]);
});

const RICH_VALUES = (f: string, values: string[]) =>
  `<c:numRef><c:f>${f}</c:f><c:numCache><c:formatCode>0.00</c:formatCode><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}" formatCode="0.0%"><c:v>${v}</c:v></c:pt>`).join('')}<c:extLst><c:ext uri="{CACHE}"/></c:extLst></c:numCache><c:extLst><c:ext uri="{REF}"/></c:extLst></c:numRef>`;
const richSer = (i: number, col: string, values: string[], extra = '') =>
  ser(i, {
    tx: strRef(`Sayfa1!$${col}$1`, [col === 'B' ? 'North' : 'South']),
    cat: strRef('Sayfa1!$A$2:$A$4', ['Jan', 'Feb', 'Mar']),
    val: RICH_VALUES(`Sayfa1!$${col}$2:$${col}$4`, values),
    props: `<c:spPr><a:solidFill><a:schemeClr val="accent${i + 1}"><a:lumMod val="75000"/></a:schemeClr></a:solidFill><a:effectLst><a:glow rad="1"/></a:effectLst></c:spPr><c:invertIfNegative val="0"/><c:dPt><c:idx val="1"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="FF000${i}"/></a:solidFill></c:spPr></c:dPt>${extra}`,
    tail: `<c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}"><c16:uniqueId xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart" val="{0000000${i}-AAAA}"/></c:ext></c:extLst>`,
  });
/** A chart with what a part may hold beside the model, written with odd spacing; its cached values are what the cells hold, spelled their own way. */
const RICH = chartXml({
  head: '<c:title><c:tx><c:rich><a:bodyPr rot="0"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400"/></a:pPr><a:r><a:rPr lang="tr-TR" b="1"/><a:t>Old </a:t></a:r><a:r><a:rPr lang="tr-TR" i="1"/><a:t>title</a:t></a:r></a:p></c:rich></c:tx><c:layout><c:manualLayout><c:x val="0.1"/></c:manualLayout></c:layout><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>',
  plot:
    group('barChart', '<c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>', richSer(0, 'B', ['10.0', '20.00', '3E1']) + richSer(1, 'C', ['30', '20', '10']) + richSer(2, 'B', ['10', '20', '30']), '<c:gapWidth val="219"/><c:overlap val="-27"/>') +
    catAx('<c:numFmt formatCode="General" sourceLinked="1"/><c:txPr><a:bodyPr rot="-2700000"/><a:p><a:endParaRPr/></a:p></c:txPr>') +
    valAx('<c:majorGridlines><c:spPr><a:ln w="9525"/></c:spPr></c:majorGridlines><c:numFmt formatCode="#,##0" sourceLinked="0"/>'),
  tail: '<c:legend><c:legendPos val="r"/><c:legendEntry><c:idx val="2"/><c:delete val="1"/></c:legendEntry><c:legendEntry><c:idx val="0"/><c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr b="1"/></a:pPr><a:endParaRPr/></a:p></c:txPr></c:legendEntry><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>',
  after: '<c:printSettings><c:headerFooter/></c:printSettings><!-- a comment -->',
  before: '<c:date1904 val="0"/><c:lang  val="tr-TR"/>',
}).replace('<c:plotArea>', '\r\n  <c:plotArea >');
const richBook = () => salesBook({ drawing: frame(), charts: [RICH] });

test('chart review: only the bytes of what was changed differ in the part of a chart of the file', async () => {
  const bytes = await richBook();
  const before = await text(bytes, CHART);
  const cases: [string, (doc: Loaded) => void, (xml: string) => string][] = [
    ['the legend', (doc) => doc.history.apply(updateChart(doc.wb, 0, 0, { legend: 'b' }), undefined, undefined), (xml) => xml.replace('<c:legendPos val="r"/>', '<c:legendPos val="b"/>')],
    [
      'the title',
      (doc) => doc.history.apply(updateChart(doc.wb, 0, 0, { title: 'New & <b>' }), undefined, undefined),
      (xml) => xml.replace('<a:r><a:rPr lang="tr-TR" b="1"/><a:t>Old </a:t></a:r><a:r><a:rPr lang="tr-TR" i="1"/><a:t>title</a:t></a:r>', '<a:r><a:rPr lang="tr-TR" b="1"/><a:t>New &amp; &lt;b&gt;</a:t></a:r>'),
    ],
    [
      // The cached values say what the cells say (10, 20, 30 written "10.0", "20.00", "3E1"): no cache is written anew, no format changes.
      'a row above the data',
      (doc) => doc.history.apply([{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 }], undefined, undefined),
      (xml) => xml.replace(/\$([A-C])\$(\d)/g, (_m, col: string, row: string) => `$${col}$${Number(row) + 1}`),
    ],
    [
      // One value is another: that cache is written anew, in the format of the cell; the caches beside it stay.
      'a cell of the second series',
      (doc) => {
        doc.history.apply([{ type: 'setCells', sheet: 0, cells: [{ row: 1, col: 2, cell: { kind: 'number', value: 31, style: 0 } }] }], undefined, undefined);
        doc.history.apply(updateChart(doc.wb, 0, 0, { legend: 'b' }), undefined, undefined);
      },
      (xml) =>
        xml
          .replace('<c:legendPos val="r"/>', '<c:legendPos val="b"/>')
          .replace(
            '<c:numCache><c:formatCode>0.00</c:formatCode><c:ptCount val="3"/><c:pt idx="0" formatCode="0.0%"><c:v>30</c:v></c:pt><c:pt idx="1" formatCode="0.0%"><c:v>20</c:v></c:pt><c:pt idx="2" formatCode="0.0%"><c:v>10</c:v></c:pt><c:extLst><c:ext uri="{CACHE}"/></c:extLst></c:numCache>',
            '<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>31</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt><c:pt idx="2"><c:v>10</c:v></c:pt><c:extLst><c:ext uri="{CACHE}"/></c:extLst></c:numCache>',
          ),
    ],
  ];
  for (const [what, act, expected] of cases) {
    const doc = await load(bytes);
    act(doc);
    const after = await text(await save(doc), CHART);
    assert.ok(after === expected(before), `${what}: the part differs from the file in nothing else (lengths ${after.length} and ${expected(before).length}; first difference at ${[...after].findIndex((c, i) => c !== expected(before)[i])}: ...${after.slice(Math.max(0, [...after].findIndex((c, i) => c !== expected(before)[i]) - 80), [...after].findIndex((c, i) => c !== expected(before)[i]) + 80)}...)`);
  }
});

test('chart review: legend entries follow their series when series are taken away or put in another order; one of a series that goes, goes', async () => {
  const bytes = await richBook();
  const cases: [string, ChartPatch, string, number[] | undefined][] = [
    ['the first series taken away', { series: [{ from: 1 }, { from: 2 }] }, '<c:legendEntry><c:idx val="1"/><c:delete val="1"/></c:legendEntry><c:overlay', [1]],
    ['the last series taken away', { series: [{ from: 0 }, { from: 1 }] }, '<c:legendPos val="r"/><c:legendEntry><c:idx val="0"/><c:txPr>', undefined],
    ['the first two swapped', { series: [{ from: 1 }, { from: 0 }, { from: 2 }] }, '<c:legendEntry><c:idx val="2"/><c:delete val="1"/></c:legendEntry><c:legendEntry><c:idx val="1"/><c:txPr>', [2]],
    ['the last one first', { series: [{ from: 2 }, { from: 0 }, { from: 1 }] }, '<c:legendEntry><c:idx val="0"/><c:delete val="1"/></c:legendEntry><c:legendEntry><c:idx val="1"/><c:txPr>', [0]],
  ];
  for (const [what, patch, expected, deleted] of cases) {
    const doc = await load(bytes);
    doc.history.apply(updateChart(doc.wb, 0, 0, patch), undefined, undefined);
    assert.deepEqual(chartsOf(doc.wb.sheets[0])[0].chart.legend?.deleted, deleted, `${what}: the model`);
    const out = await save(doc);
    const xml = await text(out, CHART);
    assert.ok(xml.indexOf(expected) >= 0, `${what}: the entries of the legend in the part: ${xml.slice(xml.indexOf('<c:legend>'), xml.indexOf('</c:legend>'))}`);
    assert.deepEqual(bare(chartsOf((await reopen(out)).sheets[0])[0].chart), bare(chartsOf(doc.wb.sheets[0])[0].chart), `${what}: what is read back is the model`);
  }
  // An entry of the legend that belongs to no series (that of a trend line): the series keep their number and order.
  const trend = RICH.replace('<c:legendEntry><c:idx val="2"/>', '<c:legendEntry><c:idx val="3"/>');
  const fixed = await load(await salesBook({ drawing: frame(), charts: [trend] }));
  assert.equal(sheetCharts(fixed.wb, 0)[0].editable, 'parts');
  assert.throws(() => updateChart(fixed.wb, 0, 0, { series: [{ from: 1 }, { from: 0 }, { from: 2 }] }), /Cannot add, remove or reorder the series of this chart/);
  assert.throws(() => updateChart(fixed.wb, 0, 0, { series: [{ from: 0 }, { from: 1 }] }), /Cannot add, remove or reorder the series of this chart/);
  fixed.history.apply(updateChart(fixed.wb, 0, 0, { series: [{ from: 0, color: '#010203' }, { from: 1 }, { from: 2 }], title: 'T' }), undefined, undefined);
  await whole(await save(fixed), 'a colour and a title are changed all the same');
});

const PLAIN = (sheet = 'Sayfa1', extra = '') =>
  chartXml({ plot: group('barChart', '<c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>', ser(0, { tx: strRef(`${sheet}!$B$1`, ['N']), cat: strRef(`${sheet}!$A$2:$A$4`, ['a', 'b', 'c']), val: numRef(`${sheet}!$B$2:$B$4`, [1, 2, 3]), tail: extra })) + catAx() + valAx() });
const HIDDEN_REF = '<c:extLst><c:ext uri="{x}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"><c15:datalabelsRange><c15:f>Sayfa1!$C$2:$C$4</c15:f></c15:datalabelsRange></c:ext></c:extLst>';
const OTHER = [{ name: 'Other', xml: sheetData([['x', 1], ['y', 2]]) }];
const RELS = (inner: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;

test('chart review: a chart part is known by what it is, not by where it stands: one outside xl/charts that cannot follow keeps every sheet as it is', async () => {
  const bytes = await salesBook({ drawing: frame(), charts: [PLAIN('Sayfa1', HIDDEN_REF)], sheets: OTHER });
  const zip = ZipArchive.read(bytes);
  const encoder = new TextEncoder();
  zip.set('xl/graphs/c1.xml', await zip.read(CHART));
  zip.delete(CHART);
  zip.set('xl/drawings/_rels/drawing1.xml.rels', encoder.encode((await text(bytes, 'xl/drawings/_rels/drawing1.xml.rels')).replace('../charts/chart1.xml', '../graphs/c1.xml')));
  // Its content type says nothing of charts either: only the relationship of the drawing does.
  zip.set('[Content_Types].xml', encoder.encode((await text(bytes, '[Content_Types].xml')).replace('<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>', '<Override PartName="/xl/graphs/c1.xml" ContentType="application/xml"/>')));
  const doc = await load(await zip.write());
  assert.equal(chartsOf(doc.wb.sheets[0]).length, 1, 'the chart is read');
  assert.equal(chartsOf(doc.wb.sheets[0])[0].known, undefined);
  assert.equal(doc.wb.chartsFollow, undefined);
  for (const sheet of [0, 1]) {
    assert.throws(() => applyOp(doc.wb, { type: 'insert', sheet, axis: 'row', at: 0, count: 1 }), /sheet with chart/, `rows of sheet ${sheet}`);
    assert.throws(() => applyOp(doc.wb, { type: 'renameSheet', sheet, name: 'X' }), /workbook with chart/, `the name of sheet ${sheet}`);
  }
  // Whatever lets an edit through all the same (a model made by hand): the references of such a chart are not touched.
  doc.wb.chartsFollow = true;
  const before = doc.wb.sheets[0].drawings;
  applyOp(doc.wb, { type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 });
  assert.equal(doc.wb.sheets[0].drawings, before, 'the chart is as it was read');
  await save(doc);
});

test('chart review: a part two charts of two sheets share goes when both charts go, and stays while one of them is there', async () => {
  const STYLE = 'http://schemas.microsoft.com/office/2011/relationships/chartStyle';
  const base = await salesBook({ drawing: frame(), charts: [PLAIN()], sheets: [{ name: 'Other', xml: `${sheetData([['x', 1]])}<drawing r:id="rId7"/>` }] });
  const zip = ZipArchive.read(base);
  const encoder = new TextEncoder();
  zip.set('xl/worksheets/_rels/sheet2.xml.rels', encoder.encode(RELS('<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing2.xml"/>')));
  zip.set('xl/drawings/drawing2.xml', encoder.encode((await text(base, 'xl/drawings/drawing1.xml')).replace('rIdC1', 'rId1')));
  zip.set('xl/drawings/_rels/drawing2.xml.rels', encoder.encode(RELS('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="/xl/charts/chart2.xml"/>')));
  zip.set('xl/charts/chart2.xml', encoder.encode(PLAIN()));
  zip.set('xl/charts/style1.xml', encoder.encode('<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="201"/>'));
  for (const n of [1, 2]) zip.set(`xl/charts/_rels/chart${n}.xml.rels`, encoder.encode(RELS(`<Relationship Id="rId1" Type="${STYLE}" Target="style1.xml"/>`)));
  zip.set(
    '[Content_Types].xml',
    encoder.encode(
      (await text(base, '[Content_Types].xml')).replace(
        '</Types>',
        '<Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/><Override PartName="/xl/charts/style1.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/></Types>',
      ),
    ),
  );
  const bytes = await zip.write();
  assert.deepEqual(await packageProblems(bytes), [], 'the fixture is whole');
  const parts = (out: Uint8Array) => names(out).filter((name) => /chart|drawing/.test(name));
  // One of the two goes: the other still leads to the part they share.
  let doc = await load(bytes);
  assert.equal(chartsOf(doc.wb.sheets[0])[0].known, 'refs', 'a chart with relationships is not one the engine wrote');
  doc.history.apply(removeChart(doc.wb, 0, 0), undefined, undefined);
  let out = await save(doc);
  await whole(out, 'one chart gone');
  assert.deepEqual(parts(out), ['xl/charts/_rels/chart2.xml.rels', 'xl/charts/chart2.xml', 'xl/charts/style1.xml', 'xl/drawings/_rels/drawing2.xml.rels', 'xl/drawings/drawing2.xml']);
  // The other goes in a later save: nothing leads to the shared part any more.
  doc.history.apply(removeChart(doc.wb, 1, 0), undefined, undefined);
  out = await save(doc);
  await whole(out, 'both charts gone, a save each');
  assert.deepEqual(parts(out), []);
  while (doc.history.undo());
  same(await save(doc), bytes, 'both undone');
  // Both in one save.
  doc = await load(bytes);
  doc.history.apply(removeChart(doc.wb, 0, 0), undefined, undefined);
  doc.history.apply(removeChart(doc.wb, 1, 0), undefined, undefined);
  out = await save(doc);
  await whole(out, 'both charts gone in one save');
  assert.deepEqual(parts(out), []);
  doc.history.undo();
  out = await save(doc);
  await whole(out, 'one of them back');
  assert.deepEqual(parts(out), ['xl/charts/_rels/chart2.xml.rels', 'xl/charts/chart2.xml', 'xl/charts/style1.xml', 'xl/drawings/_rels/drawing2.xml.rels', 'xl/drawings/drawing2.xml']);
});
