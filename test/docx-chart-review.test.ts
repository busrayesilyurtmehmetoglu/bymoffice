// Charts of a document and of a workbook, after their review: a chart part is
// read once however its name is spelt, the bytes of all chart parts together
// are bounded, warnings are bounded, and the painter of the charts of a
// document hands fonts, locale and names on as it is told.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { MAX_DOC_CHART_BYTES, MAX_DOC_CHARTS, MAX_DOC_CHART_WARNINGS, readDocCharts } from '../src/formats/docx/charts';
import { MAX_CHART_FRAME_WARNINGS } from '../src/formats/docx/common';
import { readDocx, type DocxReadOptions } from '../src/formats/docx/document';
import { LOCALE_TR } from '../src/format/locale';
import { parseChart } from '../src/formats/xlsx/chart';
import { MAX_WORKBOOK_CHART_BYTES, MAX_WORKBOOK_DRAWING_BYTES, readSheetDrawings, type ChartReading } from '../src/formats/xlsx/drawings';
import type { LayoutFont } from '../src/layout/doc/types';
import type { Chart } from '../src/model/sheet/chart';
import type { DocModel, DrawingItem, Paragraph } from '../src/model/doc/types';
import { OpcPackage } from '../src/opc';
import { DocChartPainter } from '../src/render/doc-chart';
import { notifyFontsChanged } from '../src/render/fonts';
import { ZipArchive } from '../src/zip';
import { catAx, chartFrameOf, chartWorkbook, chartXml, group, legend, numLit, numRef, OFFICE_THEME, readChart, ser, strLit, strRef, title, valAx } from './chart-fixtures';
import { p } from './doc-layout-fixtures';
import { chartParts, columnsChart, inlineChart } from './docx-chart-fixtures';
import { buildDocx as buildSpec, type DocxSpec } from './docx-fixtures';
import { twoCell } from './sheet-drawing-fixtures';

const CHART1 = '/word/charts/chart1.xml';
const encoder = new TextEncoder();

async function read(spec: DocxSpec, options?: DocxReadOptions): Promise<DocModel> {
  return readDocx(await OpcPackage.open(await buildSpec(spec)), options);
}

/** The drawings of the paragraphs of the body, in order. */
function drawings(doc: DocModel): DrawingItem[] {
  const out: DrawingItem[] = [];
  for (const block of doc.body.blocks) {
    if (block.type !== 'paragraph') continue;
    for (const inline of (block as Paragraph).children) if (inline.type === 'run') for (const item of inline.content) if (item.type === 'drawing') out.push(item);
  }
  return out;
}

/** A package that counts how often each part is read. */
function counting(pkg: OpcPackage): { pkg: OpcPackage; reads: Map<string, number> } {
  const reads = new Map<string, number>();
  const readPart = pkg.readPart.bind(pkg);
  pkg.readPart = (part: string) => {
    const key = part.toLowerCase();
    reads.set(key, (reads.get(key) || 0) + 1);
    return readPart(part);
  };
  return { pkg, reads };
}

// ------------------------------------------------------- one part, one read

const SPELLINGS = ['charts/chart1.xml', 'charts/Chart1.xml', 'CHARTS/CHART1.XML', 'Charts/chart1.XML'];

test('docx: a chart part named in several spellings is read once and counted once', async () => {
  const bytes = await buildSpec({
    body: SPELLINGS.map((_name, i) => p(inlineChart(`c${i}`, 100, 80))).join(''),
    rels: SPELLINGS.map((name, i) => [`c${i}`, 'chart', name] as [string, string, string]),
    parts: [['word/charts/chart1.xml', columnsChart()]],
  });
  const { pkg, reads } = counting(await OpcPackage.open(bytes));
  // Six cached values and one chart are all the document may read.
  const doc = await readDocx(pkg, { maxCharts: 1, maxChartPoints: 6 });
  assert.equal(reads.get(CHART1), 1, 'the part is read once');
  assert.deepEqual(doc.warnings, []);
  assert.deepEqual(Array.from((doc.charts as Map<string, Chart>).keys()), [CHART1]);
  // Every frame names the part as the package spells it, and so finds the chart.
  assert.deepEqual(drawings(doc).map((drawing) => drawing.chartPart), SPELLINGS.map(() => CHART1));
});

test('docx: two entries of the package whose names differ in case only are two charts', async () => {
  const doc = await read({
    body: p(inlineChart('a', 100, 80)) + p(inlineChart('b', 100, 80)),
    rels: [
      ['a', 'chart', 'charts/chart1.xml'],
      ['b', 'chart', 'charts/CHART1.xml'],
    ],
    parts: [
      ['word/charts/chart1.xml', columnsChart()],
      ['word/charts/CHART1.xml', columnsChart({ values: [1, 2, 3] })],
    ],
  });
  assert.deepEqual(drawings(doc).map((drawing) => drawing.chartPart), [CHART1, '/word/charts/CHART1.xml']);
  assert.equal((doc.charts as Map<string, Chart>).size, 2);
});

test('docx: the bytes of all chart parts together are bounded; a chart beyond them stays a placeholder', async () => {
  assert.equal(MAX_DOC_CHART_BYTES, 64 * 1024 * 1024);
  const xml = columnsChart();
  const size = encoder.encode(xml).length;
  const charts = await chartParts([1, 2, 3].map((n) => ({ rel: `c${n}`, xml })));
  const doc = await read({ body: [1, 2, 3].map((n) => p(inlineChart(`c${n}`, 100, 80))).join(''), ...charts }, { maxChartBytes: size * 2 + 10 });
  assert.deepEqual(Array.from((doc.charts as Map<string, Chart>).keys()), [CHART1, '/word/charts/chart2.xml']);
  assert.ok(doc.warnings.some((line) => /chart3\.xml.*too many bytes/.test(line)), JSON.stringify(doc.warnings));
});

test('docx: a header part the document names in several spellings is read once', async () => {
  const bytes = await buildSpec({
    body: p('text') + '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:headerReference w:type="first" r:id="h2"/><w:headerReference w:type="even" r:id="h3"/></w:sectPr>',
    headers: { rIdH: p('head') },
    rels: [
      ['h2', 'header', 'Header1.xml'],
      ['h3', 'header', 'HEADER1.XML'],
    ],
  });
  const { pkg, reads } = counting(await OpcPackage.open(bytes));
  const doc = await readDocx(pkg);
  assert.equal(reads.get('/word/header1.xml'), 1, 'the header part is read once');
  assert.equal(doc.headers.size, 3);
  assert.equal(doc.headers.get('h2'), doc.headers.get('rIdH'), 'the ids share the story');
  assert.equal(doc.headers.get('h3'), doc.headers.get('rIdH'), 'the ids share the story');
  assert.deepEqual(doc.warnings, []);
});

/** A workbook whose drawing names chart1.xml through relationships of several spellings. */
async function spelledWorkbook(names: string[], charts: string[]): Promise<Uint8Array> {
  const drawing = names.map((_name, i) => twoCell([0, i * 10], [4, i * 10 + 8], chartFrameOf(`rIdX${i}`))).join('');
  const zip = ZipArchive.read(await chartWorkbook({ drawing, charts }));
  zip.set(
    'xl/drawings/_rels/drawing1.xml.rels',
    encoder.encode(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        names.map((name, i) => `<Relationship Id="rIdX${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="${name}"/>`).join('') +
        '</Relationships>',
    ),
  );
  return zip.write();
}

const plainColumns = () => chartXml({ plot: group('barChart', '<c:barDir val="col"/>', ser(0, { cat: strRef('Sayfa1!$A$1:$A$3', ['a', 'b', 'c']), val: numRef('Sayfa1!$B$1:$B$3', [1, 2, 3]) })) + catAx() + valAx() });

async function sheetDrawings(bytes: Uint8Array, reading: ChartReading): Promise<{ drawings: NonNullable<Awaited<ReturnType<typeof readSheetDrawings>>>; warnings: string[]; reads: Map<string, number> }> {
  const { pkg, reads } = counting(await OpcPackage.open(bytes));
  const warnings: string[] = [];
  const found = await readSheetDrawings(pkg, await pkg.relationships('/xl/worksheets/sheet1.xml'), OFFICE_THEME, new Set(), warnings, reading);
  assert.ok(found, 'the sheet has drawings');
  return { drawings: found, warnings, reads };
}

test('xlsx: a chart part named in several spellings is read once and counted once', async () => {
  const names = ['../charts/chart1.xml', '../charts/Chart1.xml', '../CHARTS/CHART1.XML', '../Charts/chart1.XML'];
  const reading: ChartReading = { read: parseChart, charts: 1, points: 6 };
  const { drawings: found, warnings, reads } = await sheetDrawings(await spelledWorkbook(names, [plainColumns()]), reading);
  assert.equal(reads.get('/xl/charts/chart1.xml'), 1, 'the part is read once');
  assert.deepEqual(warnings, []);
  assert.deepEqual(found.map((drawing) => drawing.object.type), ['chart', 'chart', 'chart', 'chart']);
  assert.equal(reading.charts, 0, 'one chart was taken from the budget');
});

test('xlsx: the chart a second drawing names again is not read again', async () => {
  const bytes = await spelledWorkbook(['../charts/chart1.xml'], [plainColumns()]);
  const { pkg, reads } = counting(await OpcPackage.open(bytes));
  const reading: ChartReading = { read: parseChart, charts: 1, points: 6 };
  const rels = await pkg.relationships('/xl/worksheets/sheet1.xml');
  for (let round = 0; round < 3; round++) {
    const warnings: string[] = [];
    const found = await readSheetDrawings(pkg, rels, OFFICE_THEME, new Set(), warnings, reading);
    assert.deepEqual(warnings, [], `round ${round}`);
    assert.equal(found && found[0].object.type, 'chart', `round ${round}`);
  }
  assert.equal(reads.get('/xl/charts/chart1.xml'), 1, 'the chart part is read once for the workbook');
});

test('xlsx: a chart part that could not be read is said for every sheet that shows it, and read once', async () => {
  const { pkg, reads } = counting(await OpcPackage.open(await spelledWorkbook(['../charts/chart1.xml'], ['<broken'])));
  const reading: ChartReading = { read: parseChart, charts: 10, points: 100 };
  const rels = await pkg.relationships('/xl/worksheets/sheet1.xml');
  for (let round = 0; round < 2; round++) {
    const warnings: string[] = [];
    const found = await readSheetDrawings(pkg, rels, OFFICE_THEME, new Set(), warnings, reading);
    assert.equal(found && found[0].object.type, 'placeholder', `round ${round}`);
    assert.equal(warnings.length, 1, `round ${round}: ${JSON.stringify(warnings)}`);
    assert.match(warnings[0], /^Chart could not be read: /);
  }
  assert.equal(reads.get('/xl/charts/chart1.xml'), 1);
  assert.equal(reading.charts, 9);
  // What is known of the parts of one package says nothing about those of another.
  const other = await OpcPackage.open(await spelledWorkbook(['../charts/chart1.xml'], [plainColumns()]));
  const found = await readSheetDrawings(other, await other.relationships('/xl/worksheets/sheet1.xml'), OFFICE_THEME, new Set(), [], reading);
  assert.equal(found && found[0].object.type, 'chart');
});

test('xlsx: the bytes of all chart parts of a workbook together are bounded', async () => {
  assert.equal(MAX_WORKBOOK_CHART_BYTES, 64 * 1024 * 1024);
  const xml = plainColumns();
  const size = encoder.encode(xml).length;
  const names = ['../charts/chart1.xml', '../charts/chart2.xml', '../charts/chart3.xml'];
  const reading: ChartReading = { read: parseChart, charts: 10, points: 1000, bytes: size * 2 + 10 };
  const { drawings: found, warnings } = await sheetDrawings(await spelledWorkbook(names, [xml, xml, xml]), reading);
  assert.deepEqual(found.map((drawing) => drawing.object.type), ['chart', 'chart', 'placeholder']);
  assert.ok(warnings.some((line) => /too many bytes/.test(line)), JSON.stringify(warnings));
});

test('xlsx: a drawing part a sheet names in several spellings is read once; the bytes of all drawings of a workbook are bounded', async () => {
  assert.equal(MAX_WORKBOOK_DRAWING_BYTES, 256 * 1024 * 1024);
  const zip = ZipArchive.read(await spelledWorkbook(['../charts/chart1.xml'], [plainColumns()]));
  const type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
  zip.set(
    'xl/worksheets/_rels/sheet1.xml.rels',
    encoder.encode(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        ['../drawings/drawing1.xml', '../drawings/Drawing1.xml', '../DRAWINGS/DRAWING1.XML'].map((name, i) => `<Relationship Id="rIdD${i || ''}" Type="${type}" Target="${name}"/>`).join('') +
        '</Relationships>',
    ),
  );
  const size = (await zip.read('xl/drawings/drawing1.xml')).length;
  const reading: ChartReading = { read: parseChart, charts: 10, points: 1000, drawingBytes: size * 2 + 10 };
  const { pkg, reads } = counting(await OpcPackage.open(await zip.write()));
  const rels = await pkg.relationships('/xl/worksheets/sheet1.xml');
  const read = async () => {
    const warnings: string[] = [];
    return { found: await readSheetDrawings(pkg, rels, OFFICE_THEME, new Set(), warnings, reading), warnings };
  };
  const first = await read();
  assert.equal(reads.get('/xl/drawings/drawing1.xml'), 1, 'the drawing part is read once for the sheet');
  assert.equal(first.found && first.found.length, 1, 'its one frame is there once');
  assert.deepEqual(first.warnings, []);
  // Another sheet that names the same drawing reads it again, within what the workbook may read.
  const second = await read();
  assert.equal(second.found && second.found.length, 1);
  const third = await read();
  assert.equal(third.found, undefined, 'the third reading is beyond the bytes of the workbook');
  assert.ok(third.warnings.some((line) => /too many bytes/.test(line)), JSON.stringify(third.warnings));
});

// ---------------------------------------------------------------- warnings

test('docx: of the frames that lead nowhere twenty are said, whatever their number', async () => {
  const count = MAX_CHART_FRAME_WARNINGS + 7;
  const doc = await read({
    body: Array.from({ length: count }, (_value, i) => p(inlineChart(`c${i}`, 50, 40))).join(''),
    rels: Array.from({ length: count }, (_value, i) => [`c${i}`, 'chart', i % 2 ? `charts/gone${i}.xml` : `https://example.invalid/${i}.xml`, i % 2 === 0] as [string, string, string, boolean]),
  });
  assert.equal(doc.warnings.length, MAX_CHART_FRAME_WARNINGS);
  assert.equal(doc.charts, undefined);
});

test('docx: of the chart parts that cannot be read twenty are said, and the others counted in one line', async () => {
  const count = MAX_DOC_CHART_WARNINGS + 5;
  const charts = await chartParts(Array.from({ length: count }, (_value, i) => ({ rel: `c${i}`, xml: i === 3 ? columnsChart() : '<broken' })));
  const doc = await read({ body: Array.from({ length: count }, (_value, i) => p(inlineChart(`c${i}`, 50, 40))).join(''), ...charts });
  assert.equal((doc.charts as Map<string, Chart>).size, 1);
  assert.equal(doc.warnings.filter((line) => /could not be read/.test(line)).length, MAX_DOC_CHART_WARNINGS);
  assert.equal(doc.warnings.length, MAX_DOC_CHART_WARNINGS + 1);
  assert.equal(doc.warnings[doc.warnings.length - 1], 'More charts that are not shown: 4');
});

test('docx: a limit that is asked for cannot be wider than that of the engine', async () => {
  const bytes = encoder.encode(columnsChart({ values: [1] }));
  let read = 0;
  const pkg = { readPart: async () => (read++, bytes) } as unknown as OpcPackage;
  const parts = Array.from({ length: MAX_DOC_CHARTS + 3 }, (_value, i) => `/word/charts/chart${i}.xml`);
  const warnings: string[] = [];
  const charts = await readDocCharts(pkg, parts, [], warnings, { maxCharts: MAX_DOC_CHARTS * 10 });
  assert.equal(read, MAX_DOC_CHARTS);
  assert.equal((charts as Map<string, Chart>).size, MAX_DOC_CHARTS);
  assert.deepEqual(warnings, [`Charts beyond the limit of a document (${MAX_DOC_CHARTS}) are not shown: 3`]);
});

test('docx: a series whose cache is there but holds no value is not said to be without a cache', async () => {
  const blank = await read({ body: p(inlineChart('c1', 100, 80)), ...(await chartParts([{ rel: 'c1', xml: columnsChart({ values: [null, null, null] }) }])) });
  const none = await read({ body: p(inlineChart('c1', 100, 80)), ...(await chartParts([{ rel: 'c1', xml: columnsChart({ uncached: true }) }])) });
  const noted = (doc: DocModel) => ((doc.charts as Map<string, Chart>).get(CHART1) as Chart).notes.filter((note) => /embedded workbook/.test(note)).length;
  assert.equal(noted(blank), 0, 'a cache of blanks is a cache');
  assert.equal(noted(none), 1, 'a series without a cache is noted');
});

// ------------------------------------------------------------- the painter

interface Recorded {
  ctx: CanvasRenderingContext2D;
  texts: { text: string; font: string }[];
  /** The fonts text was measured with. */
  measured: string[];
}

/** A context that records the words that are drawn and the fonts that are measured with. Each character is half the font size wide. */
function recorder(): Recorded {
  const out: Recorded = { ctx: undefined as unknown as CanvasRenderingContext2D, texts: [], measured: [] };
  let font = '10px sans-serif';
  const stack: string[] = [];
  const nothing = () => undefined;
  const ctx: { [name: string]: unknown } = {
    fillStyle: '#000000',
    strokeStyle: '#000000',
    lineWidth: 1,
    lineJoin: 'miter',
    lineCap: 'butt',
    globalAlpha: 1,
    textAlign: 'left',
    textBaseline: 'alphabetic',
    save: () => void stack.push(font),
    restore: () => void (font = stack.pop() as string),
    translate: nothing,
    scale: nothing,
    rotate: nothing,
    setLineDash: nothing,
    beginPath: nothing,
    closePath: nothing,
    moveTo: nothing,
    lineTo: nothing,
    quadraticCurveTo: nothing,
    bezierCurveTo: nothing,
    arc: nothing,
    rect: nothing,
    clip: nothing,
    fill: nothing,
    stroke: nothing,
    fillRect: nothing,
    fillText: (text: string) => void out.texts.push({ text, font }),
    measureText: (text: string) => {
      out.measured.push(font);
      const match = /(\d+(?:\.\d+)?)px/.exec(font);
      return { width: text.length * (match ? Number(match[1]) : 10) * 0.5 };
    },
  };
  Object.defineProperty(ctx, 'font', { get: () => font, set: (value: string) => void (font = value) });
  out.ctx = ctx as unknown as CanvasRenderingContext2D;
  return out;
}

const face = (font: LayoutFont) => `${font.italic ? 'italic ' : ''}${font.bold ? 'bold ' : ''}${font.size}px "${font.family}"`;
/** Columns of 0.5, 1.5 and 2.5 with a title, a legend and a series that has no name. */
const wordy = (head = title('Tea and coffee')) =>
  readChart(chartXml({ head, plot: group('barChart', '<c:barDir val="col"/>', ser(0, { cat: strLit(['a', 'b', 'c']), val: numLit([0.5, 1.5, 2.5]) })) + catAx() + valAx(), tail: legend('r') })) as Chart;

function paint(painter: DocChartPainter, chart: Chart): Recorded {
  const recorded = recorder();
  painter.paint(recorded.ctx, chart, 0, 0, 400, 300);
  return recorded;
}

test('painter: words that name no font are set in the font it was given for them, and in Calibri without one', () => {
  const themed = paint(new DocChartPainter({ cssFont: face, family: 'Theme Body' }), wordy());
  assert.ok(themed.texts.length > 3, 'the chart has words');
  for (const entry of themed.texts) assert.match(entry.font, /"Theme Body"$/, entry.text);
  const plain = paint(new DocChartPainter({ cssFont: face }), wordy());
  for (const entry of plain.texts) assert.match(entry.font, /"Calibri"$/, entry.text);
  // A font the chart names is the font of its words.
  const named = paint(new DocChartPainter({ cssFont: face, family: 'Theme Body' }), wordy(title('T').replace('/><a:t>', '><a:latin typeface="Named Face"/></a:rPr><a:t>')));
  assert.match((named.texts.find((entry) => entry.text === 'T') as { font: string }).font, /"Named Face"$/);
});

test('painter: every word is measured with the font it is painted with, that of the document', () => {
  const recorded = paint(new DocChartPainter({ cssFont: (font) => face(font) + ', doc-chain' }), wordy());
  assert.ok(recorded.measured.length > 0, 'text was measured');
  for (const font of recorded.measured) assert.match(font, /doc-chain$/);
  for (const entry of recorded.texts) assert.match(entry.font, /doc-chain$/, entry.text);
  const drawn = new Set(recorded.texts.map((entry) => entry.font));
  for (const font of drawn) assert.ok(recorded.measured.indexOf(font) >= 0, `a font words are painted with was never measured with: ${font}`);
});

test('painter: bold, italic and the size of the words of a chart reach the font of the document', () => {
  const recorded = paint(new DocChartPainter({ cssFont: face }), wordy(title('Loud', 'sz="2400" b="1" i="1"')));
  const loud = recorded.texts.find((entry) => entry.text === 'Loud');
  assert.ok(loud, 'the title is drawn');
  // 24 pt are 32 px.
  assert.equal(loud.font, 'italic bold 32px "Calibri"');
  const quiet = paint(new DocChartPainter({ cssFont: face }), wordy(title('Quiet', 'sz="1200" b="0"')));
  assert.equal((quiet.texts.find((entry) => entry.text === 'Quiet') as { font: string }).font, '16px "Calibri"');
});

test('painter: numbers are written in the locale and a series without a name gets the name it is given', () => {
  const english = paint(new DocChartPainter({ cssFont: face }), wordy()).texts.map((entry) => entry.text);
  assert.ok(english.indexOf('Series1') >= 0, JSON.stringify(english));
  assert.ok(english.some((text) => /^\d\.5$/.test(text)), `a tick with a decimal point: ${JSON.stringify(english)}`);
  const turkish = paint(new DocChartPainter({ cssFont: face, locale: () => LOCALE_TR, seriesName: (n) => `Seri${n}` }), wordy()).texts.map((entry) => entry.text);
  assert.ok(turkish.indexOf('Seri1') >= 0, JSON.stringify(turkish));
  assert.ok(turkish.some((text) => /^\d,5$/.test(text)), `a tick with a decimal comma: ${JSON.stringify(turkish)}`);
  assert.ok(!turkish.some((text) => /^\d\.5$/.test(text)), 'no tick with a decimal point');
});

test('painter: clear() lets go of the charts and of the font strings; fonts that change do too', () => {
  let version = 0;
  const painter = new DocChartPainter({ cssFont: (font) => `${face(font)} v${version}` });
  const chart = wordy();
  const fonts = (recorded: Recorded) => Array.from(new Set(recorded.texts.map((entry) => entry.font.replace(/^.* v/, 'v'))));
  assert.deepEqual(fonts(paint(painter, chart)), ['v0']);
  assert.deepEqual(painter.state().charts, 1);
  version = 1;
  // Nothing said that anything changed: what was worked out is kept.
  assert.deepEqual(fonts(paint(painter, chart)), ['v0']);
  assert.equal(painter.state().layouts, 1);
  painter.clear();
  assert.equal(painter.state().charts, 0);
  assert.equal(painter.state().weight, 0);
  const after = paint(painter, chart);
  assert.deepEqual(fonts(after), ['v1'], 'the font strings are asked for again after clear()');
  assert.ok(after.measured.every((font) => / v1$/.test(font)), 'and the chart is measured with them');
  assert.equal(painter.state().layouts, 2);
  // A font that was registered or arrived: the strings are asked for again without clear().
  version = 2;
  notifyFontsChanged();
  assert.deepEqual(fonts(paint(painter, chart)), ['v2']);
});
