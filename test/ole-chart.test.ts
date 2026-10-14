// Charts of objects embedded in doc and RTF files (Excel.Chart.8, Excel.Sheet.8,
// MSGraph.Chart.8): the object is a compound file with a BIFF workbook, read
// with the reader of xls files; the chart is drawn in the frame of the object
// in place of the picture the file stores for it, and a conversion to docx
// writes that picture and says so. Every object here is built in code: none was
// compared with an object written by Office (see src/formats/doc/ole-chart.ts).

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CfbFile } from '../src/cfb';
import { writeCfb } from '../src/cfb-write';
import { readDoc } from '../src/formats/doc/document';
import { compObjProgId, MAX_OLE_SMALL_STREAM, ole1Native, OleChartReader, oleChartKind, olePictureSize, placeableWmf } from '../src/formats/doc/ole-chart';
import { MAX_DOC_CHART_BOOKS_BYTES } from '../src/formats/docx/charts';
import { emptyWriteReport, buildDocx } from '../src/formats/docx/write';
import { DocLayout } from '../src/layout/doc/layout';
import type { ImageItem, PageItem } from '../src/layout/doc/types';
import { chartValues, layoutChart } from '../src/layout/chart';
import { chartCacheList, type Chart } from '../src/model/sheet/chart';
import type { DocModel, DrawingItem } from '../src/model/doc/types';
import { readDocx } from '../src/formats/docx/document';
import { OpcPackage } from '../src/opc';
import { fakeMeasure } from './chart-fixtures';
import { fakeMeasurer } from './doc-layout-fixtures';
import { DocBuilder, sprm, u32le } from './doc-binary-fixtures';
import { doc as rtfDoc, drawingsOf, fakePng, hexOf, paragraphs, read as readRtfText, texts as rtfTexts } from './rtf-fixtures';
import { areaFormat, bar, cachedNumber, chartRecords, lineFormat, siIndex } from './xls-chart-fixtures';
import { cat, defaultStyles, externSheet, number, ptg, rec, str16, supBookSelf, u16, workbookStream } from './xls-fixtures';

const TWIP = 635;
const GREEN = '#00B050';

// ------------------------------------------------------------------ fixtures

const label = (row: number, col: number, text: string): Uint8Array => rec(0x0204, u16(row), u16(col), u16(15), str16(text));
const column = (col: number): number[] => ptg.area3d(0, 1, 3, col, col);

/** The data sheet: categories a, b, c in A2:A4, "Sales" (50, 100, 25) in B. */
const cells = (): Uint8Array => cat(label(0, 1, 'Sales'), ...['a', 'b', 'c'].map((text, i) => label(i + 1, 0, text)), ...[50, 100, 25].map((value, i) => number(i + 1, 1, value)));

/** WINDOW1 with the active sheet. */
const window1 = (active: number): Uint8Array => rec(0x003d, u16(0), u16(0), u16(0), u16(0), u16(0x38), u16(active), u16(0), u16(1), u16(600));

interface BookSpec {
  /** The sheet the workbook shows: 0 the data sheet, 1 the chart sheet. */
  active?: number;
  /** Values the chart caches, behind SIIndex records. */
  cache?: Uint8Array[];
  /** No chart sheet: the workbook is only cells. */
  noChart?: boolean;
  /** The size the chart sheet states, points. */
  size?: [number, number];
}

/** A BIFF8 workbook as an embedded object holds it: a data sheet and a chart sheet with a column chart of it in green. */
function book(spec: BookSpec = {}): Uint8Array {
  const sheets: { name: string; type?: number; records: Uint8Array }[] = [{ name: 'Data', records: cells() }];
  if (!spec.noChart) {
    sheets.push({
      name: 'Chart1',
      type: 2,
      records: chartRecords({
        series: [{ name: ptg.ref3d(0, 0, 1), values: column(1), categories: column(0), format: [lineFormat({ style: 5 }), areaFormat({ color: '00B050' })] }],
        groups: [{ type: bar(0, 100) }],
        size: spec.size || [480, 288],
        cache: spec.cache,
      }),
    });
  }
  return workbookStream({ head: cat(window1(spec.active || 0), defaultStyles()), tail: cat(supBookSelf(2), externSheet([0, 0, 0])), sheets });
}

/** The stream "\1CompObj" of an object of the class. */
function compObj(progId: string, userType = 'Microsoft Excel Chart'): Uint8Array {
  const lp = (text: string) => cat(u32le(text.length + 1), Array.from(text, (c) => c.charCodeAt(0)), [0]);
  return cat(new Uint8Array(28), lp(userType), lp('Biff8'), lp(progId), new Uint8Array(12));
}

/** A PICF in "\3PIC": the size the object is shown at, twips, scaled in thousandths. */
function picf(width: number, height: number, scale = 1000): Uint8Array {
  const out = new Uint8Array(68);
  const view = new DataView(out.buffer);
  view.setUint32(0, 68, true);
  view.setUint16(4, 68, true);
  view.setUint16(28, width, true);
  view.setUint16(30, height, true);
  view.setUint16(32, scale, true);
  view.setUint16(34, scale, true);
  return out;
}

/** A Windows metafile that holds nothing: the header and the end record. */
const META = cat(u16(1), u16(9), u16(0x0300), u32le(12), u16(0), u32le(3), u16(0), u32le(3), u16(0));

interface ObjectSpec {
  progId?: string;
  workbook?: Uint8Array | null;
  streamName?: string;
  pic?: Uint8Array;
  meta?: Uint8Array;
}

/** The streams of an object under `prefix`. */
function objectStreams(prefix: string, spec: ObjectSpec): { path: string; data: Uint8Array }[] {
  const out = [{ path: prefix + '\u0001CompObj', data: compObj(spec.progId || 'Excel.Chart.8') }];
  if (spec.workbook !== null) out.push({ path: prefix + (spec.streamName || 'Workbook'), data: spec.workbook || book() });
  if (spec.pic) out.push({ path: prefix + '\u0003PIC', data: spec.pic });
  if (spec.meta) out.push({ path: prefix + '\u0003META', data: spec.meta });
  return out;
}

/** A doc file whose text is "a", the character of the object, "b": the character with sprmCFSpec, sprmCFOle2 and sprmCPicLocation = id. */
function docWithObjects(objects: [id: number, spec: ObjectSpec][], chars = objects.map(([id]) => id), flags: Uint8Array[] = [sprm(0x0855, 1), sprm(0x080a, 1)]): Uint8Array {
  const builder = new DocBuilder();
  const text = 'a' + '\u0001'.repeat(chars.length) + 'b\r';
  builder.text(text);
  builder.papxAll(text);
  chars.forEach((id, i) => builder.chpx(1 + i, 2 + i, ...flags, sprm(0x6a03, u32le(id))));
  const base = CfbFile.read(builder.build());
  const streams = ['WordDocument', '1Table', 'Data'].filter((name) => base.has(name)).map((name) => ({ path: name, data: base.stream(name) }));
  for (const [id, spec] of objects) streams.push(...objectStreams(`ObjectPool/_${id}/`, spec));
  return writeCfb(streams);
}

const readDocBytes = (bytes: Uint8Array): DocModel => readDoc(CfbFile.read(bytes));

/** The OLE1 object RTF writes in "\objdata" for an embedded object of OLE2: the class and the compound file. */
function ole1(className: string, data: Uint8Array): Uint8Array {
  return cat(u32le(0x0501), u32le(2), u32le(className.length + 1), Array.from(className, (c) => c.charCodeAt(0)), [0], u32le(0), u32le(0), u32le(data.length), data);
}

/** The compound file of an object as RTF holds it: the streams at its root. */
const objectFile = (spec: ObjectSpec): Uint8Array => writeCfb(objectStreams('', spec));

const PNG = fakePng(40, 20, 16);

function rtfObject(options: { className?: string; data?: Uint8Array; result?: boolean; size?: string } = {}): string {
  const className = options.className || 'Excel.Chart.8';
  const data = options.data || ole1(className, objectFile({ progId: className }));
  const result = options.result === false ? '' : String.raw`{\result {\pict\pngblip\picwgoal3000\pichgoal2000 ` + hexOf(PNG) + '}}';
  return String.raw`a{\object\objemb` + (options.size === undefined ? String.raw`\objw4000\objh2400` : options.size) + String.raw`{\*\objclass ` + className + String.raw`}{\*\objdata ` + hexOf(data) + '}' + result + String.raw`}b\par`;
}

function drawingsIn(model: DocModel): DrawingItem[] {
  return paragraphs(model).reduce<DrawingItem[]>((out, paragraph) => out.concat(drawingsOf(paragraph)), []);
}

function images(model: DocModel): ImageItem[] {
  const layout = new DocLayout(model, fakeMeasurer);
  const result = layout.layout();
  assert.ok(result.done, 'the layout finishes');
  const out: ImageItem[] = [];
  const walk = (items: PageItem[]) => {
    for (const item of items) {
      if (item.type === 'image') out.push(item);
      const inner = (item as { items?: PageItem[] }).items;
      if (Array.isArray(inner)) walk(inner);
    }
  };
  for (const page of layout.pages) walk(page.items);
  return out;
}

/** The green columns the chart is drawn with, from the values the model holds (no workbook: the doc painter has none). */
function greenColumns(chart: Chart): { x: number; height: number }[] {
  const layout = layoutChart(chart, chartValues(chart), 320, 192, fakeMeasure);
  return (layout.items.filter((item) => item.kind === 'rect' && item.fill === GREEN) as { x: number; height: number }[]).sort((a, b) => a.x - b.x);
}

function assertSalesColumns(chart: Chart): void {
  assert.deepEqual(chartValues(chart).series.map((series) => series.values.toArray()), [[50, 100, 25]]);
  const columns = greenColumns(chart);
  assert.equal(columns.length, 3, 'three columns are drawn');
  assert.ok(Math.abs(columns[0].height / columns[1].height - 0.5) < 0.02, `the first column is half the second: ${columns[0].height} / ${columns[1].height}`);
  assert.ok(Math.abs(columns[2].height / columns[1].height - 0.25) < 0.02, `the third column is a quarter of the second: ${columns[2].height} / ${columns[1].height}`);
}

// -------------------------------------------------------------------- pieces

test('ole chart: the class of an object is told by its ProgID; only Excel charts and sheets and Graph charts may hold a chart', () => {
  assert.equal(oleChartKind('Excel.Chart.8'), 'excelChart');
  assert.equal(oleChartKind(' excel.chart '), 'excelChart');
  assert.equal(oleChartKind('Excel.Sheet.8'), 'excelSheet');
  assert.equal(oleChartKind('MSGraph.Chart.8'), 'graph');
  for (const other of ['Equation.3', 'Package', 'Word.Document.8', 'Excel.Chart.8x', '', undefined]) assert.equal(oleChartKind(other), undefined, String(other));
  assert.equal(compObjProgId(compObj('MSGraph.Chart.8')), 'MSGraph.Chart.8');
  // A clipboard format given as a number.
  const numbered = cat(new Uint8Array(28), u32le(2), [0x41, 0], u32le(0xffffffff), u32le(3), u32le(14), Array.from('Excel.Sheet.8', (c) => c.charCodeAt(0)), [0]);
  assert.equal(compObjProgId(numbered), 'Excel.Sheet.8');
  // Cut short, or a length beyond the stream: no ProgID.
  assert.equal(compObjProgId(compObj('Excel.Chart.8').subarray(0, 40)), undefined);
  assert.equal(compObjProgId(cat(new Uint8Array(28), u32le(0x7fffffff))), undefined);
  assert.equal(compObjProgId(new Uint8Array(0)), undefined);
});

test('ole chart: the native data of an OLE1 object, and the size of a PICF', () => {
  const data = new Uint8Array([1, 2, 3]);
  assert.deepEqual(ole1Native(ole1('Excel.Chart.8', data)), { className: 'Excel.Chart.8', data });
  // A linked object (format 1), a size beyond the data, a cut header.
  const linked = ole1('Excel.Chart.8', data);
  linked[4] = 1;
  assert.equal(ole1Native(linked), undefined);
  const long = ole1('Excel.Chart.8', data);
  long.set(u32le(1000), long.length - 7);
  assert.equal(ole1Native(long), undefined);
  assert.equal(ole1Native(new Uint8Array([1, 5, 0, 0, 2, 0, 0, 0])), undefined);
  assert.deepEqual(olePictureSize(picf(4000, 2400)), { width: 4000 * TWIP, height: 2400 * TWIP });
  assert.deepEqual(olePictureSize(picf(4000, 2400, 500)), { width: 2000 * TWIP, height: 1200 * TWIP });
  assert.equal(olePictureSize(picf(0, 2400)), undefined);
  assert.equal(olePictureSize(new Uint8Array(20)), undefined);
  // No size beyond the 22 inches of a picture.
  assert.deepEqual(olePictureSize(picf(60000, 2400)), { width: 31680 * TWIP, height: 2400 * TWIP });
});

// ------------------------------------------------------------------------ doc

test('doc: the character of an embedded Excel chart is a chart drawn from the cells of its workbook, of the size the object states', () => {
  const model = readDocBytes(docWithObjects([[7, { pic: picf(4800, 2880), meta: META }]]));
  const [drawing] = drawingsIn(model);
  assert.ok(drawing, 'the object is a drawing');
  assert.deepEqual([drawing.kind, drawing.placement, drawing.chartPart, drawing.width, drawing.height], ['chart', 'inline', '/ObjectPool/_7', 4800 * TWIP, 2880 * TWIP]);
  // The picture of the object goes with it: a conversion writes it.
  assert.ok(drawing.image && drawing.image.target, 'the metafile of the object is its picture');
  assert.equal(model.media && model.media.get(drawing.image!.target!)!.type, 'image/x-wmf');
  const chart = model.charts && model.charts.get('/ObjectPool/_7');
  assert.ok(chart, 'the chart was read');
  assert.equal(chart.shownOnly, true);
  assertSalesColumns(chart);
  assert.deepEqual(chartValues(chart).series.map((series) => series.name), ['Sales']);
  assert.ok(chart.notes.some((note) => /Read from the embedded workbook/.test(note)), JSON.stringify(chart.notes));
  assert.deepEqual(model.warnings, []);
  // The layout gives the frame the chart, in place of the metafile.
  const [frame] = images(model);
  assert.equal(frame.chart, chart);
  assert.equal(frame.placeholder, 'chart');
  assert.equal(frame.part, undefined);
  assert.ok(Math.abs(frame.width - (4800 / 1440) * 96) < 0.5, `the frame is as wide as the object: ${frame.width}`);
});

test('doc: the values the chart cached are drawn as they are; a chart without a stated size takes that of its chart sheet', () => {
  const cache = [siIndex(1), cachedNumber(0, 0, 10), cachedNumber(1, 0, 20), cachedNumber(2, 0, 40)];
  const model = readDocBytes(docWithObjects([[3, { workbook: book({ cache }) }]]));
  const [drawing] = drawingsIn(model);
  assert.deepEqual([drawing.kind, drawing.width, drawing.height], ['chart', 480 * 12700, 288 * 12700]);
  assert.equal(drawing.image, undefined);
  const chart = model.charts!.get('/ObjectPool/_3') as Chart;
  assert.deepEqual(chartCacheList(chart.groups[0].series[0].values), [10, 20, 40]);
  assert.deepEqual(chartValues(chart).series[0].values.toArray(), [10, 20, 40]);
  // Without a picture the frame is the chart, not a placeholder.
  const [frame] = images(model);
  assert.equal(frame.chart, chart);
});

test('doc: an Excel workbook that shows its chart sheet is a chart; one that shows cells, and objects of other classes, show nothing as before', () => {
  const sheet = readDocBytes(docWithObjects([[1, { progId: 'Excel.Sheet.8', workbook: book({ active: 1 }) }]]));
  assert.equal(drawingsIn(sheet)[0].kind, 'chart');
  assertSalesColumns(sheet.charts!.get('/ObjectPool/_1') as Chart);

  const cellsShown = readDocBytes(docWithObjects([[1, { progId: 'Excel.Sheet.8', workbook: book({ active: 0 }) }]]));
  assert.deepEqual([drawingsIn(cellsShown).length, cellsShown.charts, cellsShown.warnings], [0, undefined, []]);

  const equation = readDocBytes(docWithObjects([[1, { progId: 'Equation.3', workbook: null }]]));
  assert.deepEqual([drawingsIn(equation).length, equation.charts, equation.warnings], [0, undefined, []]);

  // No storage of that name.
  const missing = readDocBytes(docWithObjects([[1, {}]], [2]));
  assert.deepEqual([drawingsIn(missing).length, missing.charts, missing.warnings], [0, undefined, []]);
});

test('doc: a chart of Microsoft Graph is read the same way from its workbook', () => {
  // UNSURE: that Graph 97 keeps a BIFF8 workbook in "Workbook" (see ole-chart.ts).
  const model = readDocBytes(docWithObjects([[9, { progId: 'MSGraph.Chart.8' }]]));
  assert.equal(drawingsIn(model)[0].kind, 'chart');
  assertSalesColumns(model.charts!.get('/ObjectPool/_9') as Chart);
});

test('doc: an object of a chart whose workbook is missing, cut short or of another kind shows nothing and says why; the document opens', () => {
  const stream = book();
  const cases: [string, ObjectSpec, RegExp][] = [
    ['without a workbook', { workbook: null }, /holds no workbook/],
    ['cut short', { workbook: stream.subarray(0, 300) }, /not shown/],
    ['of noise', { workbook: new Uint8Array(4096).fill(0x41) }, /not shown/],
    ['without a chart', { workbook: book({ noChart: true }) }, /holds no chart/],
  ];
  for (const [name, spec, why] of cases) {
    const model = readDocBytes(docWithObjects([[1, spec]]));
    assert.equal(drawingsIn(model).length, 0, name);
    assert.equal(model.charts, undefined, name);
    assert.equal(model.warnings.length, 1, `${name}: ${JSON.stringify(model.warnings)}`);
    assert.match(model.warnings[0], why, name);
    assert.match(model.warnings[0], /^The chart of an embedded object \(ObjectPool\/_1\) is not shown/, name);
  }
  // A package in the stream "Package" (an object of Excel 2007) is said to be one.
  const packaged = readDocBytes(docWithObjects([[1, { workbook: new Uint8Array([0x50, 0x4b, 3, 4]), streamName: 'Package' }]]));
  assert.match(packaged.warnings[0], /package \(xlsx\)/);
});

test('doc: an object that many characters show is read once; the objects of a document share the limits of charts', () => {
  const twice = readDocBytes(docWithObjects([[4, {}]], [4, 4]));
  const drawn = drawingsIn(twice);
  assert.deepEqual(drawn.map((drawing) => drawing.chartPart), ['/ObjectPool/_4', '/ObjectPool/_4']);
  assert.equal(twice.charts!.size, 1);

  // The limits, small: two workbooks at most, one chart, values and bytes of workbooks.
  const source = CfbFile.read(docWithObjects([1, 2, 3].map((id): [number, ObjectSpec] => [id, {}])));
  const charts: string[] = [];
  const reader = new OleChartReader(charts, { maxCharts: 1 });
  assert.ok(reader.readChart(source, 'ObjectPool/_1/', 'a', 'excelChart', 'one').chart);
  assert.match(reader.readChart(source, 'ObjectPool/_2/', 'b', 'excelChart', 'two').problem || '', /more than 1 charts/);
  const books: string[] = [];
  const fewBooks = new OleChartReader(books, { maxChartBooks: 1 });
  assert.ok(fewBooks.readChart(source, 'ObjectPool/_1/', 'a', 'excelChart', 'one').chart);
  assert.match(fewBooks.readChart(source, 'ObjectPool/_2/', 'b', 'excelChart', 'two').problem || '', /too many embedded workbooks/);
  const small = new OleChartReader([], { maxChartBookBytes: 100 });
  assert.match(small.readChart(source, 'ObjectPool/_1/', 'a', 'excelChart', 'one').problem || '', /too large/);
  const bytes = new OleChartReader([], { maxChartBooksBytes: book().length + 10 });
  assert.ok(bytes.readChart(source, 'ObjectPool/_1/', 'a', 'excelChart', 'one').chart);
  assert.match(bytes.readChart(source, 'ObjectPool/_2/', 'b', 'excelChart', 'two').problem || '', /too many bytes/);
  // Values taken from the cells count too: the values fit, the categories no longer do.
  const points = new OleChartReader([], { maxChartPoints: 4 });
  const partly = points.readChart(source, 'ObjectPool/_1/', 'a', 'excelChart', 'one').chart as Chart;
  assert.deepEqual(chartValues(partly).series[0].values.toArray(), [50, 100, 25]);
  assert.ok(partly.notes.some((note) => /^Shown empty, without cached values: categories \(the charts of the document hold too many values/.test(note)), JSON.stringify(partly.notes));
  // Cached values beyond what is left: the chart is not shown.
  const cached = CfbFile.read(docWithObjects([[1, { workbook: book({ cache: [siIndex(1), cachedNumber(0, 0, 1), cachedNumber(1, 0, 2), cachedNumber(2, 0, 3)] }) }]]));
  assert.match(new OleChartReader([], { maxChartPoints: 2 }).readChart(cached, 'ObjectPool/_1/', 'a', 'excelChart', 'one').problem || '', /too many values/);
});

test('doc: warnings about the charts of objects stop at a limit, and the rest are counted in one line', () => {
  const ids = Array.from({ length: 25 }, (_, i) => i + 1);
  const model = readDocBytes(docWithObjects(ids.map((id): [number, ObjectSpec] => [id, { workbook: null }])));
  assert.equal(model.warnings.length, 21, JSON.stringify(model.warnings));
  assert.equal(model.warnings[20], 'More embedded objects whose chart or picture is not shown: 5');
});

test('doc: converted to docx, the chart is written as the picture of the object and the report says so', async () => {
  const model = readDocBytes(docWithObjects([[7, { pic: picf(4800, 2880), meta: META }], [8, {}]]));
  assert.equal(model.charts!.size, 2);
  const report = emptyWriteReport();
  const bytes = await buildDocx(model, { report });
  assert.ok(report.warnings.some((warning) => /^1 chart\(s\) of embedded objects .* written as the picture the file stores for them/.test(warning)), JSON.stringify(report.warnings));
  assert.ok(report.warnings.some((warning) => /^1 chart\(s\) of embedded objects .* written as empty frames/.test(warning)), JSON.stringify(report.warnings));
  const pkg = await OpcPackage.open(bytes);
  const again = await readDocx(pkg);
  const drawn = drawingsIn(again);
  assert.deepEqual(drawn.map((drawing) => [drawing.kind, drawing.width, drawing.height]), [
    ['picture', 4800 * TWIP, 2880 * TWIP],
    ['picture', 480 * 12700, 288 * 12700],
  ]);
  assert.ok(drawn[0].image && drawn[0].image.target && /\.wmf$/.test(drawn[0].image.target), 'the metafile is written');
  assert.equal(again.charts, undefined, 'no chart part is written');
});

// ------------------------------------------------------------------------ RTF

test('rtf: an embedded Excel chart is drawn in place of the picture of its result, which stays with it', () => {
  const model = readRtfText(rtfDoc(rtfObject()));
  assert.deepEqual(paragraphs(model).map((paragraph) => drawingsOf(paragraph).length), [1]);
  const [drawing] = drawingsIn(model);
  assert.deepEqual([drawing.kind, drawing.chartPart, drawing.width, drawing.height], ['chart', '/rtf/object1', 3000 * TWIP, 2000 * TWIP]);
  assert.ok(drawing.image, 'the picture of the result stays');
  const chart = model.charts!.get('/rtf/object1') as Chart;
  assertSalesColumns(chart);
  assert.deepEqual(model.warnings, []);
  assert.equal(model.media!.size, 1, 'the data of the object is no picture of the document');
  const [frame] = images(model);
  assert.equal(frame.chart, chart);
  assert.equal(frame.part, undefined, 'the picture is not drawn');
});

test('rtf: a picture shape in the result is the frame of the chart; a result with a text box is left as it is', () => {
  const data = ole1('Excel.Chart.8', objectFile({}));
  const object = (result: string) => readRtfText(rtfDoc(String.raw`a{\object\objemb\objw4000\objh2400{\*\objclass Excel.Chart.8}{\*\objdata ` + hexOf(data) + String.raw`}{\result ` + result + String.raw`}}b\par`));
  const picture = object(String.raw`{\shp{\*\shpinst{\sp{\sn pib}{\sv {\pict\pngblip\picwgoal600\pichgoal300 ` + hexOf(PNG) + String.raw`}}}}}`);
  const drawn = drawingsIn(picture);
  assert.equal(drawn.length, 1, 'no second frame');
  assert.deepEqual([drawn[0].kind, drawn[0].chartPart], ['chart', '/rtf/object1']);
  const boxed = object(String.raw`{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 202}}{\shptxt box\par}}}`);
  assert.deepEqual([drawingsIn(boxed).length, drawingsIn(boxed)[0].kind, boxed.charts], [1, 'shape', undefined]);
});

test('rtf: an object of a chart without a result is a frame of the size it states, scaled', () => {
  const model = readRtfText(rtfDoc(rtfObject({ result: false, size: String.raw`\objw4000\objh2400\objscalex50\objscaley50` })));
  const [drawing] = drawingsIn(model);
  assert.deepEqual([drawing.kind, drawing.width, drawing.height, drawing.image], ['chart', 2000 * TWIP, 1200 * TWIP, undefined]);
  const unsized = drawingsIn(readRtfText(rtfDoc(rtfObject({ result: false, size: '' }))))[0];
  assert.deepEqual([unsized.width, unsized.height], [480 * 12700, 288 * 12700], 'the size of the chart sheet');
});

test('rtf: Graph and Excel sheets are read; other classes, broken and cut data keep the picture of the result', () => {
  const graph = readRtfText(rtfDoc(rtfObject({ className: 'MSGraph.Chart.8' })));
  assert.equal(drawingsIn(graph)[0].kind, 'chart');

  const shownCells = readRtfText(rtfDoc(rtfObject({ className: 'Excel.Sheet.8' })));
  assert.deepEqual([drawingsIn(shownCells)[0].kind, shownCells.charts, shownCells.warnings], ['picture', undefined, []]);

  const other = readRtfText(rtfDoc(rtfObject({ className: 'Package', data: ole1('Package', objectFile({ progId: 'Package' })) })));
  assert.deepEqual([drawingsIn(other)[0].kind, other.charts, other.warnings], ['picture', undefined, []]);

  const full = ole1('Excel.Chart.8', objectFile({}));
  const cases: [string, Uint8Array, RegExp][] = [
    ['cut short', full.subarray(0, 1200), /not that of an embedded object|cannot be read/],
    ['not a compound file', ole1('Excel.Chart.8', new Uint8Array(600).fill(7)), /not that of an embedded object/],
    ['no OLE1 header', new Uint8Array([1, 2, 3]), /not that of an embedded object/],
    ['without a workbook', ole1('Excel.Chart.8', writeCfb([{ path: '\u0001CompObj', data: compObj('Excel.Chart.8') }])), /holds no workbook/],
  ];
  for (const [name, data, why] of cases) {
    const model = readRtfText(rtfDoc(rtfObject({ data })));
    const [drawing] = drawingsIn(model);
    assert.equal(drawing.kind, 'picture', name);
    assert.equal(model.charts, undefined, name);
    assert.equal(model.warnings.length, 1, `${name}: ${JSON.stringify(model.warnings)}`);
    assert.match(model.warnings[0], why, name);
  }
});

test('rtf: data of an object beyond the limit is not kept, and the picture of the result stays', () => {
  const big = new Uint8Array(17 * 1024 * 1024);
  const model = readRtfText(rtfDoc(rtfObject({ data: ole1('Excel.Chart.8', big) })));
  assert.equal(drawingsIn(model)[0].kind, 'picture');
  assert.match(model.warnings.join('\n'), /its data is too large/);
});

test('rtf: converted to docx, the chart is written as the picture of the result', async () => {
  const model = readRtfText(rtfDoc(rtfObject()));
  const report = emptyWriteReport();
  const again = await readDocx(await OpcPackage.open(await buildDocx(model, { report })));
  assert.ok(report.warnings.some((warning) => /^1 chart\(s\) of embedded objects .* the picture the file stores for them/.test(warning)), JSON.stringify(report.warnings));
  const [drawing] = drawingsIn(again);
  assert.deepEqual([drawing.kind, drawing.width, drawing.height], ['picture', 3000 * TWIP, 2000 * TWIP]);
});

test('doc and rtf: objects whose workbook bytes are damaged anywhere never throw, and show a chart or nothing', () => {
  const stream = book();
  let seed = 12345;
  const random = (n: number) => {
    seed = (seed * 1103515245 + 12345) & 0x7fffffff;
    return seed % n;
  };
  for (let round = 0; round < 60; round++) {
    const damaged = stream.slice(0, round % 7 === 6 ? random(stream.length) : stream.length);
    for (let k = 0; k < 1 + random(6); k++) if (damaged.length) damaged[random(damaged.length)] = random(256);
    const model = readDocBytes(docWithObjects([[1, { workbook: damaged }]]));
    const drawn = drawingsIn(model);
    assert.ok(drawn.length === 0 || (drawn[0].kind === 'chart' && model.charts !== undefined && model.charts.has('/ObjectPool/_1')), `round ${round}`);
    if (drawn.length === 0) assert.equal(model.warnings.length, 1, `round ${round}: ${JSON.stringify(model.warnings)}`);
    const rtf = readRtfText(rtfDoc(rtfObject({ data: ole1('Excel.Chart.8', objectFile({ workbook: damaged })) })));
    assert.equal(drawingsIn(rtf).length, 1, `round ${round}: the picture or the chart stays`);
  }
});

// ------------------------------------------------------- hostile files (review)

type Lazy = { size: number; make: () => Uint8Array };

/**
 * The streams of a doc file with streams added that are made only when they are read: what the
 * entries state of them is all the reader may look at first. Counts the bytes read of the
 * storages of objects.
 */
function lazySource(bytes: Uint8Array, extra: Map<string, Lazy>) {
  const cfb = CfbFile.read(bytes);
  const out = { read: 0, calls: 0 };
  const source = {
    entries: cfb.entries.concat(Array.from(extra, ([path, stream]) => ({ path, name: path, type: 'stream' as const, size: stream.size }))),
    has: (path: string) => cfb.has(path) || extra.has(path),
    stream: (path: string): Uint8Array => {
      const made = extra.get(path);
      const data = made ? made.make() : cfb.stream(path);
      if (path.startsWith('ObjectPool/')) {
        out.read += data.length;
        out.calls++;
      }
      return data;
    },
  };
  return { source, out };
}

test('doc: an object that many characters show is looked at once, and a large \\1CompObj or \\3PIC is never read', () => {
  // 3,000 characters of one object of a class that is no chart, whose \1CompObj is 8 MB.
  const many = Array.from({ length: 3000 }, () => 5);
  const huge: Lazy = { size: 8 * 1024 * 1024, make: () => cat(compObj('Equation.3'), new Uint8Array(8 * 1024 * 1024)) };
  const big = lazySource(docWithObjects([], many), new Map([['ObjectPool/_5/\u0001CompObj', huge]]));
  const started = Date.now();
  const model = readDoc(big.source);
  assert.deepEqual([drawingsIn(model).length, model.warnings], [0, []]);
  assert.equal(big.out.calls, 0, 'nothing of the stream is read');
  assert.ok(Date.now() - started < 5000, 'quickly');

  // A small \1CompObj of an object that is no chart: read once for all characters.
  const small = compObj('Equation.3');
  const once = lazySource(docWithObjects([], many), new Map([['ObjectPool/_5/\u0001CompObj', { size: small.length, make: () => small }]]));
  readDoc(once.source);
  assert.deepEqual([once.out.calls, once.out.read], [1, small.length]);

  // Thousands of objects, each with a large \1CompObj: none is read.
  const ids = Array.from({ length: 2000 }, (_, i) => i + 1);
  const all = lazySource(docWithObjects([], ids), new Map(ids.map((id): [string, Lazy] => [`ObjectPool/_${id}/\u0001CompObj`, huge])));
  readDoc(all.source);
  assert.equal(all.out.calls, 0);

  // A \3PIC beyond the limit is not read either: the chart takes the size of its sheet.
  const sized = lazySource(docWithObjects([[1, {}]]), new Map([['ObjectPool/_1/\u0003PIC', { size: MAX_OLE_SMALL_STREAM + 1, make: () => picf(100, 100) }]]));
  const [drawing] = drawingsIn(readDoc(sized.source));
  assert.deepEqual([drawing.kind, drawing.width, drawing.height], ['chart', 480 * 12700, 288 * 12700]);
});

test('doc: the streams of objects (workbooks, metafiles) are read within one limit of the document; a stream beyond it is not read', () => {
  // 80 charts, each with a metafile the entries say is 2 MB: 160 MB, more than the limit of 64 MB.
  const ids = Array.from({ length: 80 }, (_, i) => i + 1);
  const meta: Lazy = { size: 2 * 1024 * 1024, make: () => cat(META, new Uint8Array(2 * 1024 * 1024 - META.length)) };
  const doc = lazySource(docWithObjects(ids.map((id): [number, ObjectSpec] => [id, {}])), new Map(ids.map((id): [string, Lazy] => [`ObjectPool/_${id}/\u0003META`, meta])));
  const model = readDoc(doc.source);
  assert.ok(doc.out.read <= MAX_DOC_CHART_BOOKS_BYTES, `read ${doc.out.read} bytes`);
  const drawn = drawingsIn(model);
  const pictured = drawn.filter((drawing) => drawing.image).length;
  assert.ok(pictured > 0 && pictured < 32, `the metafiles that fit are read: ${pictured}`);
  assert.ok(drawn.length > 0 && drawn.every((drawing) => drawing.kind === 'chart'));
});

test('doc: sprmCFData, or sprmCFObj without sprmCFOle2, is no object of the ObjectPool', () => {
  const withData = readDocBytes(docWithObjects([[1, {}]], [1], [sprm(0x0855, 1), sprm(0x080a, 1), sprm(0x0806, 1)]));
  assert.deepEqual([drawingsIn(withData).length, withData.charts], [0, undefined]);
  const objOnly = readDocBytes(docWithObjects([[1, {}]], [1], [sprm(0x0855, 1), sprm(0x0856, 1)]));
  assert.deepEqual([drawingsIn(objOnly).length, objOnly.charts], [0, undefined]);
  const both = readDocBytes(docWithObjects([[1, {}]], [1], [sprm(0x0855, 1), sprm(0x0856, 1), sprm(0x080a, 1)]));
  assert.equal(drawingsIn(both)[0].kind, 'chart');
});

test('doc: a chart sheet larger than 22 inches gives a frame of 22 inches', () => {
  const model = readDocBytes(docWithObjects([[1, { workbook: book({ size: [5000, 300] }) }]]));
  const [drawing] = drawingsIn(model);
  assert.deepEqual([drawing.width, drawing.height], [31680 * TWIP, 300 * 12700]);
});

test('doc: the metafile of an object is made placeable, its box the window it sets; a METAFILEPICT in front is passed over', () => {
  // SetWindowOrg (0, 0) and SetWindowExt (2000 wide, 1000 high) after the header, then the end record.
  const records = cat(u32le(5), u16(0x020b), u16(0), u16(0), u32le(5), u16(0x020c), u16(1000), u16(2000), u32le(3), u16(0));
  const wmf = cat(u16(1), u16(9), u16(0x0300), u32le(9 + records.length / 2), u16(0), u32le(5), u16(0), records);
  const placed = placeableWmf(wmf, 4000, 2000) as Uint8Array;
  const view = new DataView(placed.buffer, placed.byteOffset, placed.byteLength);
  assert.equal(view.getUint32(0, true), 0x9ac6cdd7);
  assert.deepEqual([view.getInt16(6, true), view.getInt16(8, true), view.getInt16(10, true), view.getInt16(12, true)], [0, 0, 2000, 1000]);
  assert.equal(view.getUint16(14, true), 720, 'the window fills 4000 twips: 720 units an inch');
  let sum = 0;
  for (let i = 0; i < 10; i++) sum ^= view.getUint16(i * 2, true);
  assert.equal(view.getUint16(20, true), sum);
  assert.deepEqual(Array.from(placed.subarray(22)), Array.from(wmf));
  // Behind a METAFILEPICT of 8 bytes: the same.
  assert.deepEqual(Array.from(placeableWmf(cat(u16(8), u16(4000), u16(2000), u16(0), wmf), 4000, 2000) as Uint8Array), Array.from(placed));
  // Without a window: the box is the frame in twips.
  const bare = placeableWmf(META, 3000, 1500) as Uint8Array;
  const bareView = new DataView(bare.buffer, bare.byteOffset, bare.byteLength);
  assert.deepEqual([bareView.getInt16(10, true), bareView.getInt16(12, true), bareView.getUint16(14, true)], [3000, 1500, 1440]);
  // Placeable already: kept; noise: none.
  assert.equal(placeableWmf(placed, 1, 1), placed);
  assert.equal(placeableWmf(new Uint8Array(64).fill(9), 100, 100), undefined);
  // In a document the picture of the object is the placeable metafile.
  const model = readDocBytes(docWithObjects([[7, { pic: picf(4800, 2880), meta: wmf }]]));
  const bytes = model.media!.get(drawingsIn(model)[0].image!.target!)!.bytes;
  assert.equal(new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint32(0, true), 0x9ac6cdd7);
});

test('rtf: a result of text, or of another object, keeps the result and adds no chart; a second \\objdata is passed over', () => {
  const data = ole1('Excel.Chart.8', objectFile({}));
  const object = (inner: string) => readRtfText(rtfDoc(String.raw`a{\object\objemb\objw4000\objh2400{\*\objclass Excel.Chart.8}{\*\objdata ` + hexOf(data) + '}' + inner + String.raw`}b\par`));
  const text = object(String.raw`{\result Chart fallback text}`);
  assert.deepEqual([drawingsIn(text).length, text.charts, text.warnings], [0, undefined, []]);
  assert.deepEqual(rtfTexts(text), ['aChart fallback textb']);
  const nested = object(String.raw`{\result {\object\objemb{\*\objclass Equation.3}{\result inner}}}`);
  assert.deepEqual([drawingsIn(nested).length, nested.charts], [0, undefined]);
  const twice = object(String.raw`{\*\objdata 00}{\result {\pict\pngblip\picwgoal3000\pichgoal2000 ` + hexOf(PNG) + '}}');
  assert.deepEqual([drawingsIn(twice)[0].kind, twice.warnings], ['chart', []]);
});

test('rtf: warnings about objects of charts stop at the limit like the others', () => {
  const model = readRtfText(rtfDoc(rtfObject({ data: new Uint8Array([1, 2, 3]) }).repeat(25)));
  assert.equal(model.warnings.length, 21, JSON.stringify(model.warnings));
  assert.equal(model.warnings[20], 'More embedded objects whose chart or picture is not shown: 5');
});

test('layout: a chart drawing with text boxes draws its picture and text, not the chart', () => {
  const model = readRtfText(rtfDoc(rtfObject()));
  const drawing = drawingsIn(model)[0];
  assert.equal(drawing.kind, 'chart');
  drawing.textBoxes = [[{ type: 'paragraph', props: {}, children: [] } as never]];
  const found = images(model);
  assert.ok(found.length > 0 && found.every((image) => image.chart === undefined), JSON.stringify(found.map((image) => image.placeholder)));
});

test('doc: a metafile of an object that is not kept is said, and so is why a workbook could not be read', () => {
  const large = lazySource(docWithObjects([[1, {}]]), new Map([['ObjectPool/_1/\u0003META', { size: 17 * 1024 * 1024, make: () => META }]]));
  const big = readDoc(large.source);
  assert.equal(drawingsIn(big)[0].image, undefined);
  assert.deepEqual(big.warnings, ['The picture of an embedded object (ObjectPool/_1) is not kept: it is larger than the reader keeps']);
  const noise = readDocBytes(docWithObjects([[1, { meta: new Uint8Array(64).fill(9) }]]));
  assert.deepEqual(noise.warnings, ['The picture of an embedded object (ObjectPool/_1) is not kept: it is no Windows metafile']);
  const broken = lazySource(docWithObjects([[1, { workbook: null }]]), new Map([['ObjectPool/_1/Workbook', { size: 100, make: () => { throw new Error('Corrupt compound file: sector chain loop'); } }]]));
  assert.deepEqual(readDoc(broken.source).warnings, ['The chart of an embedded object (ObjectPool/_1) is not shown: its workbook cannot be read (Corrupt compound file: sector chain loop)']);
});

test('doc: the window of a metafile is its last origin and extent, whichever comes first', () => {
  // SetWindowExt (2000 x 1000) before SetWindowOrg (100, 50).
  const records = cat(u32le(5), u16(0x020c), u16(1000), u16(2000), u32le(5), u16(0x020b), u16(50), u16(100), u32le(3), u16(0));
  const wmf = cat(u16(1), u16(9), u16(0x0300), u32le(9 + records.length / 2), u16(0), u32le(5), u16(0), records);
  const placed = placeableWmf(wmf, 4000, 2000) as Uint8Array;
  const view = new DataView(placed.buffer, placed.byteOffset, placed.byteLength);
  assert.deepEqual([view.getInt16(6, true), view.getInt16(8, true), view.getInt16(10, true), view.getInt16(12, true)], [100, 50, 2100, 1050]);
});
