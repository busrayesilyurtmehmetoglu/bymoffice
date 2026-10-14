// The painters of the engine on the PDF canvas: a page of a document, a page
// of a sheet and a chart sheet, painted by the code that paints them on the
// screen and in print, come out as text, drawings and nothing else - and a
// page with something the PDF canvas does not write comes out as a reason.
// What a painter asks of a canvas is listed here too: the list is what the
// PDF canvas has to answer for.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { open, type Workbook } from '../src/index';
import type { LayoutPage, PageItem } from '../src/layout/doc/types';
import { createSheetMetrics } from '../src/layout/sheet-metrics';
import { createSheet, emptyStyleSheet } from '../src/model/sheet/types';
import type { Cell, Sheet } from '../src/model/sheet/types';
import { PDF_CANVAS_MEMBERS, createPdfCanvas, type PdfCanvasMeasure } from '../src/pdf/canvas';
import { buildPdf, PdfDocument } from '../src/pdf/document';
import { PdfExportError, pdfSettings, writePdfPage, type PdfPageInfo } from '../src/pdf/export';
import { PdfFontSet } from '../src/pdf/font-set';
import { canvasMeasure, canvasPictures, drawPage, loadPdfFonts, pictureOfPixels, type PdfVectorTools } from '../src/pdf/vector';
import { joinParts } from '../src/pdf/writer';
import { renderPage } from '../src/render/doc';
import { clearTextMeasureCache } from '../src/render/text';
import { paintSheetPage, planSheetPages } from '../src/ui/sheet-print';
import { catAx, chartWorkbook, chartXml, group, numRef, ser, sheetData, solid, spPr, strRef, title, valAx } from './chart-fixtures';
import { checkPdf, poppler, popplerPixels, popplerTest, popplerWords } from './pdf-fixtures';
import { bonum, testFont } from './pdf-font-fixtures';

const DATE = new Date(Date.UTC(2026, 9, 13, 12, 0, 0));
const BONUM = { regular: bonum('regular'), bold: bonum('bold'), italic: bonum('italic'), bolditalic: bonum('bolditalic') };

/** The font of the package under its own name and, as an application would register it, in place of Calibri. */
function fonts(): PdfFontSet {
  const faces = (family: string) => [
    { family, weight: 400, style: 'normal' as const, bytes: BONUM.regular },
    { family, weight: 700, style: 'normal' as const, bytes: BONUM.bold },
    { family, weight: 400, style: 'italic' as const, bytes: BONUM.italic },
    { family, weight: 700, style: 'italic' as const, bytes: BONUM.bolditalic },
  ];
  return new PdfFontSet([...faces('TeX Gyre Bonum'), ...faces('Calibri')], (family) => family === 'Arial');
}

/** As wide as the font says: a canvas without kerning. */
function measure(set: PdfFontSet): PdfCanvasMeasure {
  return {
    width: (font, text) => {
      const resolved = set.resolve(font);
      if (!('face' in resolved)) return text.length * 6;
      let units = 0;
      for (const char of text) units += resolved.face.advance(resolved.face.glyph(char.codePointAt(0) as number));
      return (units * resolved.size) / resolved.face.unitsPerEm;
    },
    baseline: (font, baseline) => {
      const resolved = set.resolve(font);
      const size = 'face' in resolved ? resolved.size : 10;
      return baseline === 'middle' ? size * 0.3 : baseline === 'top' ? size * 0.8 : NaN;
    },
  };
}

function tools(): PdfVectorTools {
  const set = fonts();
  return { fonts: set, measure: measure(set) };
}

const near = (actual: number[], expected: number[], slack: number, what: string): void => {
  for (let i = 0; i < expected.length; i++) assert.ok(Math.abs(actual[i] - expected[i]) <= slack, `${what}: ${JSON.stringify(actual)} is not ${JSON.stringify(expected)} (±${slack})`);
};

// ---------------------------------------------------------------------------
// What the painters ask of a canvas

/** A canvas that does nothing and remembers the names of what was asked of it. */
function listening(): { ctx: CanvasRenderingContext2D; used: Set<string> } {
  const used = new Set<string>();
  const values: { [name: string]: unknown } = { font: '10px sans-serif', fillStyle: '#000000', strokeStyle: '#000000', lineWidth: 1, globalAlpha: 1, textAlign: 'left', textBaseline: 'alphabetic', lineCap: 'butt', lineJoin: 'miter', letterSpacing: '0px' };
  const ctx = new Proxy(
    {},
    {
      get(_target, key) {
        if (typeof key !== 'string' || key.slice(0, 2) === '__') return undefined;
        used.add(key);
        if (Object.prototype.hasOwnProperty.call(values, key)) return values[key];
        if (key === 'measureText') return (text: string) => ({ width: String(text).length * 6 });
        if (key === 'createLinearGradient') return () => ({ addColorStop: () => undefined });
        return () => undefined;
      },
      set(_target, key, value) {
        if (typeof key === 'string' && key.slice(0, 2) !== '__') {
          used.add(key);
          values[key] = value;
        }
        return true;
      },
      has: (_target, key) => key === 'letterSpacing',
    },
  );
  return { ctx: ctx as unknown as CanvasRenderingContext2D, used };
}

const cssOf = (font: { family: string; size: number; bold: boolean; italic: boolean }): string => `${font.italic ? 'italic ' : ''}${font.bold ? 'bold ' : ''}${font.size}px "${font.family}", serif`;

function docPage(items: PageItem[]): LayoutPage {
  return { width: 400, height: 300, section: 0, number: 1, items };
}

const BODY = { family: 'TeX Gyre Bonum', size: 16, bold: false, italic: false };

/** A page with everything a page of a document is painted from but pictures: text plain and decorated, shadings, borders, shapes. */
function fullDocPage(): LayoutPage {
  return docPage([
    { type: 'rect', x: 40, y: 20, width: 320, height: 30, fill: '#ddeeff' },
    { type: 'text', x: 50, y: 42, width: 0, text: 'Başlık: Çalışma', font: { ...BODY, bold: true }, color: '#003366' },
    { type: 'text', x: 50, y: 80, width: 150, text: 'İğne iplik, şöyle böyle.', font: BODY, color: '#000000', underline: { color: '#ff0000', style: 'single' } },
    { type: 'text', x: 50, y: 110, width: 100, text: 'üstü çizili', font: { ...BODY, italic: true }, color: '#000000', strike: true, highlight: '#ffff00' },
    { type: 'text', x: 50, y: 140, width: 120, text: 'aralıklı', font: BODY, color: '#000000', letterSpacing: 2 },
    { type: 'text', x: 220, y: 140, width: 120, text: 'geniş', font: BODY, color: '#000000', scale: 1.5 },
    { type: 'text', x: 370, y: 100, width: 60, text: 'dikey', font: BODY, color: '#000000', rotation: 90 },
    { type: 'text', x: 50, y: 170, width: 100, text: 'dalgalı', font: BODY, color: '#000000', underline: { color: '#0000ff', style: 'wave' } },
    { type: 'line', x1: 40, y1: 190, x2: 360, y2: 190, color: '#000000', width: 2, style: 'single' },
    { type: 'line', x1: 40, y1: 200, x2: 360, y2: 200, color: '#008000', width: 1, style: 'dashed' },
    { type: 'line', x1: 40, y1: 210, x2: 360, y2: 210, color: '#800000', width: 3, style: 'double' },
    { type: 'line', x1: 40, y1: 215, x2: 360, y2: 240, color: '#808080', width: 1, style: 'single' },
    { type: 'shape', x: 60, y: 245, width: 80, height: 40, geometry: 'ellipse', fill: '#ffcc00', stroke: { color: '#333333', width: 2 } },
    { type: 'shape', x: 180, y: 245, width: 80, height: 40, geometry: 'roundRect', radius: 8, fill: '#00ccff', stroke: { color: '#333333', width: 1, dash: 'dash' }, rotation: 15 },
    { type: 'shape', x: 290, y: 245, width: 60, height: 40, geometry: 'line', stroke: { color: '#ff00ff', width: 2 } },
  ] as PageItem[]);
}

function sheetOf(cells: (string | number)[][]): Sheet {
  const sheet = createSheet({ name: 'Veri', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/xl/worksheets/sheet1.xml' });
  cells.forEach((values, r) => {
    const row: Cell[] = values.map((value) => (typeof value === 'number' ? { kind: 'number', value, style: 0 } : { kind: 'string', value, style: 0 }) as Cell);
    sheet.rows.set(r, { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: row });
  });
  return sheet;
}

function bookOf(sheet: Sheet): Workbook {
  return { sheets: [sheet], styles: emptyStyleSheet(), sharedStrings: [], definedNames: [], date1904: false, activeSheet: 0 } as Workbook;
}

const CHART = chartXml({
  head: title('Satışlar'),
  plot:
    group('barChart', '<c:barDir val="col"/><c:grouping val="clustered"/>', ser(0, { cat: strRef('Sayfa1!$A$1:$A$3', ['x', 'y', 'z']), val: numRef('Sayfa1!$B$1:$B$3', [1, 1, 1]), props: spPr(solid('00B050')) })) +
    catAx() +
    valAx('<c:majorGridlines/>', '<c:orientation val="minMax"/><c:max val="40"/><c:min val="0"/>'),
});

async function chartBook(): Promise<Workbook> {
  const bytes = await chartWorkbook({ sheet: sheetData([['Çay', 10], ['Kahve', 20], ['Süt', 40]]), chartSheet: CHART });
  return (await (await open(bytes)).workbook()) as Workbook;
}

test('pdf painters: everything the painters of a document page, a sheet page and a chart ask of a canvas is something the PDF canvas has', async () => {
  const doc = listening();
  renderPage(doc.ctx, fullDocPage(), { cssFont: cssOf });
  renderPage(
    doc.ctx,
    docPage([{ type: 'image', x: 10, y: 10, width: 50, height: 50, part: '/word/media/a.png' } as PageItem]),
    { cssFont: cssOf, images: { get: () => ({ source: {} as CanvasImageSource, width: 10, height: 10 }) }, placeholderLabel: () => 'picture' },
  );
  renderPage(doc.ctx, docPage([{ type: 'image', x: 10, y: 10, width: 80, height: 50, placeholder: 'object' } as PageItem]), { cssFont: cssOf, placeholderLabel: () => 'picture' });
  const sheet = listening();
  const workbook = bookOf(sheetOf([['Ad', 'Tutar'], ['Çay', 12.5], ['Kahve', 40]]));
  workbook.sheets[0].pageSetup = { print: { headings: true, gridLines: true }, headerFooter: { oddHeader: '&CRapor &P / &N', oddFooter: '&L&D&R&F' } };
  const plan = planSheetPages(workbook, [0]);
  paintSheetPage(sheet.ctx, { workbook, job: plan.pages[0], total: 1, metrics: createSheetMetrics(workbook.sheets[0]), pixelWidth: 800, now: DATE, fileName: 'rapor.xlsx' });
  const chart = listening();
  const charts = await chartBook();
  const chartPlan = planSheetPages(charts, [1]);
  paintSheetPage(chart.ctx, { workbook: charts, job: chartPlan.pages[0], total: 1, metrics: createSheetMetrics(charts.sheets[1]), pixelWidth: 800, now: DATE });
  // The widths this canvas made up are not those of any font: the painters measure anew after it.
  clearTextMeasureCache();
  const all = new Set<string>([...doc.used, ...sheet.used, ...chart.used]);
  const missing = Array.from(all)
    .filter((member) => PDF_CANVAS_MEMBERS.indexOf(member) < 0)
    .sort();
  assert.deepEqual(missing, [], 'members the painters use and the PDF canvas does not have');
  // The list itself, so that a painter that starts to ask for something new shows up here.
  assert.deepEqual(Array.from(doc.used).sort(), ['beginPath', 'bezierCurveTo', 'closePath', 'drawImage', 'fill', 'fillRect', 'fillStyle', 'fillText', 'font', 'letterSpacing', 'lineTo', 'lineWidth', 'measureText', 'moveTo', 'quadraticCurveTo', 'restore', 'rotate', 'save', 'scale', 'setLineDash', 'stroke', 'strokeStyle', 'textAlign', 'textBaseline', 'translate']);
  assert.deepEqual(Array.from(sheet.used).sort(), ['beginPath', 'clip', 'fillRect', 'fillStyle', 'fillText', 'font', 'globalAlpha', 'lineTo', 'lineWidth', 'measureText', 'moveTo', 'rect', 'restore', 'save', 'setLineDash', 'setTransform', 'stroke', 'strokeStyle', 'textAlign', 'textBaseline']);
  for (const member of ['clip', 'fillRect', 'fillText', 'lineJoin', 'setTransform', 'stroke', 'translate']) assert.ok(chart.used.has(member), 'the painter of charts uses ' + member);
});

// ---------------------------------------------------------------------------
// A page of a document

popplerTest('pdf painters: a page of a document is written as its text, its shadings, borders and shapes', async () => {
  const page = fullDocPage();
  // Twice the pixels of the page, as a page painted at 192 dots per inch; 400 x 300 px are 300 x 225 points.
  const drawn = drawPage(tools(), { pixelWidth: 800, pixelHeight: 600, pageWidth: 300, pageHeight: 225 }, (ctx) => {
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, 800, 600);
    renderPage(ctx, page, { scale: 2, cssFont: cssOf });
  });
  assert.equal(drawn.reason, undefined);
  assert.ok(drawn.vector, 'the page is drawn');
  const file = await buildPdf([{ width: 300, height: 225, vector: drawn.vector }], { date: DATE });
  const checked = checkPdf(file);
  assert.equal(checked.pages[0].images.length, 0, 'no picture of the page');
  assert.deepEqual(
    checked.pages[0].texts.map((text) => text.text),
    ['Başlık: Çalışma', 'İğne iplik, şöyle böyle.', 'üstü çizili', 'aralıklı', 'geniş', 'dikey', 'dalgalı'],
  );
  assert.deepEqual(checked.fonts.map((font) => font.baseFont).sort(), ['TeXGyreBonum-Bold', 'TeXGyreBonum-Italic', 'TeXGyreBonum-Regular']);
  const spaced = checked.pages[0].texts[3];
  near([spaced.spacing], [2], 0.001, 'the spacing of letters of the document');
  near(checked.pages[0].texts[4].matrix, [1.5 * 0.75, 0, 0, 0.75], 0.001, 'the stretch of the text of the document');
  near(checked.pages[0].texts[5].matrix, [0, -0.75, 0.75, 0], 0.001, 'text that reads downwards');
  const report = poppler(file);
  assert.equal(report.complaints, '', 'Poppler has nothing to complain of');
  assert.ok(report.fonts.every((font) => font.embedded && font.unicode), 'every font is embedded and maps to Unicode');
  const words = popplerWords(file).words;
  const word = (text: string) => words.filter((entry) => entry.text === text)[0];
  // px of the page to points: three quarters.
  near([word('Başlık:').xMin], [50 * 0.75], 0.2, 'the heading starts where it was laid out');
  near([word('İğne').xMin], [50 * 0.75], 0.2, 'the line under it too');
  assert.ok(word('İğne').yMin < 80 * 0.75 && word('İğne').yMax > 80 * 0.75, 'on its baseline');
  near([word('geniş').xMin], [220 * 0.75], 0.2, 'the stretched word');
  const pixels = popplerPixels(file);
  near(pixels.at(250, 20), [0xdd, 0xee, 0xff], 2, 'the shading behind the heading');
  near(pixels.at(250, 50), [255, 255, 255], 2, 'the paper');
  near(pixels.at(150, 142), [0, 0, 0], 60, 'the border, two px thick around 190 px: 141.75 to 143.25 points');
  near(pixels.at(75, Math.round(265 * 0.75)), [0xff, 0xcc, 0x00], 2, 'the ellipse');
  // The highlight behind the struck text: from 37.5 to 112.5 points across, around the line at 82.5.
  let yellow = 0;
  for (let y = 73; y < 85; y++) for (let x = 40; x < 110; x++) if (pixels.at(x, y)[0] > 200 && pixels.at(x, y)[1] > 200 && pixels.at(x, y)[2] < 80) yellow++;
  assert.ok(yellow > 12 * 70 * 0.5, `the highlight behind the struck text (${yellow} yellow pixels)`);
  // The underline of the second line, in its own colour.
  let red = 0;
  for (let y = 60; y < 66; y++) if (pixels.at(100, y)[0] > 200 && pixels.at(100, y)[1] < 80) red++;
  assert.ok(red >= 1, 'the red underline under the second line');
});

test('pdf painters: a page of a document with text in a font of the device, or a character its font lacks, is not drawn', () => {
  const device = drawPage(tools(), { pixelWidth: 400, pixelHeight: 300, pageWidth: 300, pageHeight: 225 }, (ctx) =>
    renderPage(ctx, docPage([{ type: 'text', x: 50, y: 42, width: 0, text: 'Arial ile', font: { ...BODY, family: 'Arial' }, color: '#000000' } as PageItem]), { cssFont: cssOf }),
  );
  assert.equal(device.vector, undefined);
  assert.match(String(device.reason), /font of the device \("Arial"\)/);
  const lira = drawPage(tools(), { pixelWidth: 400, pixelHeight: 300, pageWidth: 300, pageHeight: 225 }, (ctx) =>
    renderPage(ctx, docPage([{ type: 'text', x: 50, y: 42, width: 0, text: '100 ₺', font: BODY, color: '#000000' } as PageItem]), { cssFont: cssOf }),
  );
  assert.match(String(lira.reason), /no glyph for U\+20BA/);
  // The words of a placeholder are in a font of the device.
  const placeholder = drawPage(tools(), { pixelWidth: 400, pixelHeight: 300, pageWidth: 300, pageHeight: 225 }, (ctx) =>
    renderPage(ctx, docPage([{ type: 'image', x: 10, y: 10, width: 200, height: 50, placeholder: 'object' } as PageItem]), { cssFont: cssOf, placeholderLabel: () => 'Embedded object' }),
  );
  assert.match(String(placeholder.reason), /font of the device/);
});

test('pdf painters: a picture of a document page is drawn as a picture of its own; a painter that fails leaves the page undrawn', async () => {
  const source = {};
  const picture = { width: 2, height: 1, image: { kind: 'samples' as const, width: 2, height: 1, colors: 3 as const, data: Uint8Array.from([255, 0, 0, 0, 0, 255]) } };
  const with_ = { ...tools(), images: { read: (from: unknown) => (from === source ? picture : undefined) } };
  const page = docPage([{ type: 'image', x: 100, y: 50, width: 200, height: 100, part: '/word/media/a.png', crop: { left: 0.5, top: 0, right: 0, bottom: 0 } } as PageItem]);
  const drawn = drawPage(with_, { pixelWidth: 400, pixelHeight: 300, pageWidth: 300, pageHeight: 225 }, (ctx) => renderPage(ctx, page, { cssFont: cssOf, images: { get: () => ({ source: source as CanvasImageSource, width: 2, height: 1 }) } }));
  assert.equal(drawn.reason, undefined);
  const checked = checkPdf(await buildPdf([{ width: 300, height: 225, vector: drawn.vector }], { date: DATE }));
  assert.equal(checked.pages[0].images.length, 1);
  // The right half of the picture fills the box: the whole picture is twice as wide and starts a box width to the left.
  near(checked.pages[0].images[0].matrix as number[], [300, 0, 0, 75, -75, 225 - 150 * 0.75], 0.01, 'the picture behind its crop');
  assert.equal(checked.pages[0].clips.length, 1, 'cut to the box');
  const failing = drawPage(tools(), { pixelWidth: 400, pixelHeight: 300, pageWidth: 300, pageHeight: 225 }, () => {
    throw new Error('the painter broke');
  });
  assert.match(String(failing.reason), /could not be painted: the painter broke/);
});

// ---------------------------------------------------------------------------
// A page of a sheet, and a chart

popplerTest('pdf painters: a page of a sheet is written as its cells, grid lines, headings, header and footer', async () => {
  const workbook = bookOf(sheetOf([['Ürün', 'Tutar'], ['Çay', 12.5], ['Kahve', 40], ['Ihlamur', 7]]));
  workbook.sheets[0].pageSetup = { print: { headings: true, gridLines: true }, headerFooter: { oddHeader: '&CRapor &P / &N', oddFooter: '&LVeri&R&F' } };
  const plan = planSheetPages(workbook, [0]);
  const paper = plan.pages[0].layout.paper;
  const pixelWidth = Math.round(paper.width * 2);
  const pixelHeight = Math.round(paper.height * 2);
  const drawn = drawPage(tools(), { pixelWidth, pixelHeight, pageWidth: paper.width * 0.75, pageHeight: paper.height * 0.75 }, (ctx) =>
    paintSheetPage(ctx, { workbook, job: plan.pages[0], total: 1, metrics: createSheetMetrics(workbook.sheets[0]), pixelWidth, now: DATE, fileName: 'rapor.xlsx' }),
  );
  assert.equal(drawn.reason, undefined);
  const file = await buildPdf([{ width: paper.width * 0.75, height: paper.height * 0.75, vector: drawn.vector }], { date: DATE });
  const checked = checkPdf(file);
  assert.equal(checked.pages[0].images.length, 0);
  const texts = checked.pages[0].texts.map((text) => text.text);
  for (const wanted of ['Ürün', 'Tutar', 'Çay', '12.5', 'Kahve', '40', 'Ihlamur', '7', 'A', 'B', '1', '4', 'Rapor 1 / 1', 'Veri', 'rapor.xlsx']) assert.ok(texts.indexOf(wanted) >= 0, `"${wanted}" is text of the page (${JSON.stringify(texts)})`);
  assert.ok(checked.pages[0].shapes.some((shape) => shape.operator === 'S'), 'the grid lines are lines');
  assert.ok(checked.pages[0].clips.length > 0, 'the cells are clipped to the page');
  const report = poppler(file);
  assert.equal(report.complaints, '');
  assert.deepEqual(
    report.fonts.map((font) => [font.type, font.embedded, font.unicode]),
    [['CID Type 0C', true, true]],
  );
  // Numbers stand right in their cells, text left: the two columns do not start at the same place.
  const words = popplerWords(file).words;
  const word = (text: string) => words.filter((entry) => entry.text === text)[0];
  assert.ok(word('Çay') && word('12.5') && word('Kahve'), 'Poppler finds the words of the cells');
  near([word('Çay').xMin], [word('Kahve').xMin], 0.1, 'the texts of a column start at the same place');
  // The painter of cells lays text on whole pixels of its canvas: half a pixel of two per px is 0.19 points.
  near([word('12.5').xMax], [word('40').xMax], 0.4, 'the numbers of a column end at the same place');
  assert.ok(word('Kahve').yMin > word('Çay').yMin + 5, 'the rows follow each other down the page');
});

test('pdf painters: a sheet whose cells are in a font no file is at hand for is not drawn', () => {
  const workbook = bookOf(sheetOf([['Ürün', 'Tutar']]));
  const plan = planSheetPages(workbook, [0]);
  const paper = plan.pages[0].layout.paper;
  // No font file stands for Calibri here.
  const set = new PdfFontSet([{ family: 'TeX Gyre Bonum', weight: 400, style: 'normal', bytes: BONUM.regular }], () => false);
  const drawn = drawPage({ fonts: set, measure: measure(set) }, { pixelWidth: 800, pixelHeight: 1131, pageWidth: paper.width * 0.75, pageHeight: paper.height * 0.75 }, (ctx) =>
    paintSheetPage(ctx, { workbook, job: plan.pages[0], total: 1, metrics: createSheetMetrics(workbook.sheets[0]), pixelWidth: 800, now: DATE }),
  );
  assert.equal(drawn.vector, undefined);
  assert.match(String(drawn.reason), /font of the device|no font of the text is known/);
});

popplerTest('pdf painters: a chart sheet is written as the drawing of its chart, with its labels as text', async () => {
  const workbook = await chartBook();
  const plan = planSheetPages(workbook, [1]);
  const paper = plan.pages[0].layout.paper;
  const pixelWidth = Math.round(paper.width * 1.5);
  const pixelHeight = Math.round(paper.height * 1.5);
  const drawn = drawPage(tools(), { pixelWidth, pixelHeight, pageWidth: paper.width * 0.75, pageHeight: paper.height * 0.75 }, (ctx) =>
    paintSheetPage(ctx, { workbook, job: plan.pages[0], total: 1, metrics: createSheetMetrics(workbook.sheets[1]), pixelWidth, now: DATE }),
  );
  assert.equal(drawn.reason, undefined);
  const file = await buildPdf([{ width: paper.width * 0.75, height: paper.height * 0.75, vector: drawn.vector }], { date: DATE });
  const checked = checkPdf(file);
  const texts = checked.pages[0].texts.map((text) => text.text);
  for (const wanted of ['Satışlar', 'Çay', 'Kahve', 'Süt', '40']) assert.ok(texts.indexOf(wanted) >= 0, `"${wanted}" is text of the chart (${JSON.stringify(texts)})`);
  // The three columns, in the colour of the series.
  const green = checked.pages[0].shapes.filter((shape) => shape.operator === 'f' && shape.colour.length === 3 && Math.abs(shape.colour[1] - 0xb0 / 255) < 0.01 && shape.colour[0] === 0);
  assert.equal(green.length, 3, 'three filled columns');
  const report = poppler(file);
  assert.equal(report.complaints, '');
  // Poppler paints them: the tallest column holds the colour of the series.
  const tallest = green.map((shape) => shape.path[0] as number[]).sort((a, b) => (b[4] as number) - (a[4] as number))[0];
  const pixels = popplerPixels(file);
  const x = Math.round((tallest[1] as number) + (tallest[3] as number) / 2);
  const y = Math.round(pixels.height - ((tallest[2] as number) + (tallest[4] as number) / 2));
  near(pixels.at(x, y), [0x00, 0xb0, 0x50], 3, 'the middle of the tallest column');
});

// ---------------------------------------------------------------------------
// The decision for a page

function pictureJob(number: number, paint: (ctx: CanvasRenderingContext2D) => void, asked: string[]) {
  return {
    number,
    width: 200,
    height: 100,
    pixelWidth: 200,
    pixelHeight: 100,
    paint,
    picture: async () => {
      asked.push('picture ' + number);
      return { image: { kind: 'samples' as const, width: 1, height: 2, colors: 1 as const, data: Uint8Array.from([0, 255]) }, text: { runs: [{ text: 'gizli', x: 10, y: 50, width: 40, size: 12 }], width: 200 } };
    },
  };
}

const good = (ctx: CanvasRenderingContext2D): void => {
  ctx.font = '16px "TeX Gyre Bonum"';
  ctx.fillText('çizilen sayfa', 10, 50);
};
const bad = (ctx: CanvasRenderingContext2D): void => {
  ctx.font = '16px Arial';
  ctx.fillText('cihaz fontu', 10, 50);
};

test('pdf export: every page is written as text and drawings when it can be, and as its picture when not', async () => {
  const pdf = new PdfDocument({ date: DATE });
  const asked: string[] = [];
  const told: PdfPageInfo[] = [];
  const settings = pdfSettings({});
  assert.equal(settings.text, 'auto', 'the default');
  assert.equal(await writePdfPage(pdf, pictureJob(1, good, asked), settings, tools(), (info) => told.push(info)), true);
  assert.equal(await writePdfPage(pdf, pictureJob(2, bad, asked), settings, tools(), (info) => told.push(info)), true);
  assert.equal(
    await writePdfPage(pdf, pictureJob(3, good, asked), settings, tools(), () => {
      throw new Error('the listener failed');
    }),
    true,
    'a listener that throws does not end the work',
  );
  await pdf.finish();
  assert.deepEqual(asked, ['picture 2'], 'only the page that could not be drawn was asked for its picture');
  assert.equal(told.length, 2);
  assert.deepEqual(told[0], { page: 1, mode: 'vector' });
  assert.equal(told[1].page, 2);
  assert.equal(told[1].mode, 'image');
  assert.match(String(told[1].reason), /font of the device \("Arial"\)/);
  const checked = checkPdf(joinParts(pdf.takeParts()));
  assert.deepEqual(
    checked.pages.map((page) => [page.images.length, page.texts.map((text) => text.mode).join(), page.text.trim()]),
    [
      [0, '0', 'çizilen sayfa'],
      [1, '3,3', 'gizli'],
      [0, '0', 'çizilen sayfa'],
    ],
  );
});

test('pdf export: the links of a page are written over it whether it is drawn or a picture, where its text is', async () => {
  const pdf = new PdfDocument({ date: DATE });
  const links = [{ rects: [{ x: 10, y: 38, width: 90, height: 16 }], uri: 'https://example.com/a' }];
  await writePdfPage(pdf, { ...pictureJob(1, good, []), links }, pdfSettings({}), tools());
  await writePdfPage(pdf, { ...pictureJob(2, bad, []), links }, pdfSettings({}), tools());
  await pdf.finish();
  const checked = checkPdf(joinParts(pdf.takeParts()));
  const root = checked.objects.get((checked.trailer.Root as { id: number }).id)!.value as { Pages: { id: number } };
  const kids = (checked.objects.get(root.Pages.id)!.value as { Kids: { id: number }[] }).Kids;
  for (let i = 0; i < 2; i++) {
    const page = checked.objects.get(kids[i].id)!.value as { Annots: { id: number }[] };
    assert.equal(page.Annots.length, 1, `page ${i + 1}`);
    const annot = checked.objects.get(page.Annots[0].id)!.value as { Rect: number[] };
    // From the top of a page of 100 points: 38 to 54 is 62 to 46 from the bottom.
    assert.deepEqual(annot.Rect, [10, 46, 100, 62]);
    // The text of the page, drawn or lying invisible over its picture, starts inside the link, on its baseline (50).
    const text = checked.pages[i].texts[0];
    assert.ok(text.x >= 10 && text.x < 100, `page ${i + 1}: text at ${text.x}`);
    assert.ok(Math.abs(100 - text.y - 50) < 0.5, `page ${i + 1}: baseline at ${text.y}`);
  }
  assert.equal(pdf.links.external, 2);
});

test('pdf export: "image" paints no page on the PDF canvas, "vector" refuses the file for a page that cannot be drawn', async () => {
  const asked: string[] = [];
  let painted = 0;
  const counting = (ctx: CanvasRenderingContext2D): void => {
    painted++;
    good(ctx);
  };
  const image = new PdfDocument({ date: DATE });
  const told: PdfPageInfo[] = [];
  await writePdfPage(image, pictureJob(1, counting, asked), pdfSettings({ text: 'image' }), tools(), (info) => told.push(info));
  assert.equal(painted, 0);
  assert.deepEqual(told, [{ page: 1, mode: 'image' }], 'no reason where nothing was tried');
  assert.deepEqual(asked, ['picture 1']);
  const vector = new PdfDocument({ date: DATE });
  await writePdfPage(vector, pictureJob(1, good, asked), pdfSettings({ text: 'vector' }), tools());
  await assert.rejects(
    writePdfPage(vector, pictureJob(2, bad, asked), pdfSettings({ text: 'vector' }), tools()),
    (error: unknown) => error instanceof PdfExportError && error.code === 'unsupported' && error.page === 2 && /font of the device/.test(String(error.reason)) && /^Page 2 cannot be written as text and drawings/.test(error.message),
  );
  assert.equal(vector.pageCount, 1, 'the page that was refused is not in the file');
  assert.deepEqual(asked, ['picture 1'], 'and its picture was not asked for');
  // Without anything to draw with (no canvas to measure on) every page is its picture, or the file is refused.
  const none = new PdfDocument({ date: DATE });
  const why: PdfPageInfo[] = [];
  await writePdfPage(none, pictureJob(1, good, asked), pdfSettings({}), undefined, (info) => why.push(info));
  assert.match(String(why[0].reason), /no canvas to measure/);
  await assert.rejects(writePdfPage(none, pictureJob(2, good, asked), pdfSettings({ text: 'vector' }), undefined), (error: unknown) => error instanceof PdfExportError && error.code === 'unsupported');
  // A job that was overtaken while its picture was made writes nothing, and neither does one whose page was drawn.
  const stale = new PdfDocument({ date: DATE });
  assert.equal(await writePdfPage(stale, { ...pictureJob(1, bad, asked), proceed: () => false }, pdfSettings({}), tools()), false);
  assert.equal(await writePdfPage(stale, { ...pictureJob(1, good, asked), proceed: () => false }, pdfSettings({}), tools()), false);
  assert.equal(stale.pageCount, 0);
  // What ends a job (the caller does not want the file any more) ends it for a drawn page too.
  await assert.rejects(
    writePdfPage(stale, { ...pictureJob(1, good, asked), proceed: () => { throw new PdfExportError('cancelled', 'Making the PDF was cancelled'); } }, pdfSettings({}), tools()),
    (error: unknown) => error instanceof PdfExportError && error.code === 'cancelled',
  );
});

test('pdf export: the options that choose how pages are written are checked before any work', () => {
  assert.equal(pdfSettings({ text: 'image' }).text, 'image');
  assert.equal(pdfSettings({ text: 'vector' }).text, 'vector');
  assert.throws(() => pdfSettings({ text: 'sharp' as never }), RangeError);
  assert.throws(() => pdfSettings({ onPage: 'no' as never }), TypeError);
});

// ---------------------------------------------------------------------------
// Where the font files, the widths and the pictures come from

test('pdf fonts: the font files of the page are read where they are bytes and fetched once where they are an address', async () => {
  const fetchedUrls: string[] = [];
  const fetcher = async (url: string) => {
    fetchedUrls.push(url);
    if (url.indexOf('missing') >= 0) return { ok: false, arrayBuffer: async () => new ArrayBuffer(0) };
    if (url.indexOf('broken') >= 0) throw new Error('network');
    const bytes = url.indexOf('bold') >= 0 ? BONUM.bold : BONUM.regular;
    return { ok: true, arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer };
  };
  const sources = [
    { family: 'Given', weight: 400, style: 'normal' as const, bytes: testFont() },
    { family: 'Fetched', weight: 400, style: 'normal' as const, url: 'https://fonts.example/pdf-test/regular.otf' },
    { family: 'Fetched', weight: 700, style: 'normal' as const, url: 'https://fonts.example/pdf-test/bold.otf' },
    { family: 'Gone', weight: 400, style: 'normal' as const, url: 'https://fonts.example/pdf-test/missing.otf' },
    { family: 'Offline', weight: 400, style: 'normal' as const, url: 'https://fonts.example/pdf-test/broken.otf' },
    { family: 'Nowhere', weight: 400, style: 'normal' as const },
  ];
  const set = await loadPdfFonts(sources, fetcher, () => false);
  const name = (font: string): string => {
    const resolved = set.resolve(font);
    return 'face' in resolved ? resolved.face.name : resolved.reason;
  };
  assert.equal(name('12px Given'), 'BymTest-Regular');
  assert.equal(name('12px Fetched'), 'TeXGyreBonum-Regular');
  assert.equal(name('bold 12px Fetched'), 'TeXGyreBonum-Bold');
  // A face whose file did not arrive still is the face the browser paints with: the next family is not taken for it.
  assert.match(name('12px Gone, Given'), /"Gone" cannot be embedded: its font file is not at hand/);
  assert.match(name('12px Offline, Given'), /not at hand/);
  assert.match(name('12px Nowhere, Given'), /not at hand/);
  // The same address again is not asked for again; one that failed is.
  const again = await loadPdfFonts(sources, fetcher, () => false);
  assert.equal('face' in again.resolve('12px Fetched'), true);
  assert.deepEqual(fetchedUrls.filter((url) => /regular|bold/.test(url)).length, 2, 'each file that arrived was fetched once');
  assert.equal(fetchedUrls.filter((url) => /missing/.test(url)).length, 2, 'the file that did not arrive is asked for anew');
  // The same file gives the same face: a file embeds it once.
  const first = set.resolve('12px Fetched');
  const second = again.resolve('12px Fetched');
  assert.ok('face' in first && 'face' in second && first.face === second.face, 'one face for one file');
});

test('pdf fonts: the real canvas says how wide text is and where its baselines lie', () => {
  const calls: string[] = [];
  const context = {
    font: '',
    textBaseline: 'alphabetic',
    measureText(text: string) {
      calls.push(`${this.font}|${this.textBaseline}|${text}`);
      // The ink of "x" rises 5 above the alphabetic baseline, and lies 3 below the middle line.
      return { width: text.length * 7, actualBoundingBoxAscent: this.textBaseline === 'alphabetic' ? 5 : this.textBaseline === 'middle' ? -3 : undefined };
    },
  };
  const doc = { createElement: () => ({ getContext: () => context }) } as unknown as Document;
  const measurer = canvasMeasure(doc);
  assert.ok(measurer, 'a document with a canvas has a measurer');
  assert.equal(measurer.width('12px A', 'abc'), 21);
  assert.equal(measurer.width('12px A', 'abc'), 21);
  assert.equal(calls.length, 1, 'a width is measured once');
  assert.equal((measurer.baseline as NonNullable<typeof measurer.baseline>)('12px A', 'middle'), 8, 'the alphabetic baseline lies 8 below the middle line');
  assert.ok(isNaN((measurer.baseline as NonNullable<typeof measurer.baseline>)('12px A', 'top')), 'a canvas that does not say is not guessed for');
  assert.equal(context.textBaseline, 'alphabetic', 'the canvas is left as it was');
  assert.equal(canvasMeasure({ createElement: () => ({ getContext: () => null }) } as unknown as Document), undefined);
  assert.equal(
    canvasMeasure({
      createElement: () => {
        throw new Error('no DOM');
      },
    } as unknown as Document),
    undefined,
  );
});

test('pdf pictures: what a painter draws is read once through a canvas: its pixels, its mask, a JPEG for a photograph', () => {
  assert.deepEqual(pictureOfPixels([10, 10, 10, 255, 20, 20, 20, 255], 2, 1), { kind: 'samples', data: Uint8Array.from([10, 20]), width: 2, height: 1, colors: 1, alpha: undefined });
  assert.deepEqual(pictureOfPixels([10, 20, 30, 255, 1, 2, 3, 0], 2, 1), { kind: 'samples', data: Uint8Array.from([10, 20, 30, 1, 2, 3]), width: 2, height: 1, colors: 3, alpha: Uint8Array.from([255, 0]) });
  let reads = 0;
  let jpegs = 0;
  // A canvas that gives back what its source says its pixels are.
  const doc = {
    createElement: () => {
      let drawnSource: { pixels?: number[]; tainted?: boolean } = {};
      const canvas = {
        width: 0,
        height: 0,
        getContext: () => ({
          clearRect: () => undefined,
          drawImage: (source: { pixels?: number[]; tainted?: boolean }) => {
            drawnSource = source;
          },
          getImageData: (_x: number, _y: number, width: number, height: number) => {
            reads++;
            if (drawnSource.tainted) throw new Error('The canvas has been tainted');
            return { data: drawnSource.pixels || new Array(width * height * 4).fill(255) };
          },
        }),
        toDataURL: () => {
          jpegs++;
          // The smallest thing that reads as a JPEG of 2 x 2 pixels with three components.
          const bytes = [0xff, 0xd8, 0xff, 0xc0, 0x00, 0x11, 0x08, 0x00, 0x02, 0x00, 0x02, 0x03, 0x01, 0x11, 0x00, 0x02, 0x11, 0x00, 0x03, 0x11, 0x00, 0xff, 0xd9];
          return 'data:image/jpeg;base64,' + Buffer.from(bytes).toString('base64');
        },
      };
      return canvas;
    },
  } as unknown as Document;
  const noisy = [255, 0, 0, 255, 0, 255, 0, 255, 0, 0, 255, 255, 9, 9, 9, 255];
  const photo = { width: 2, height: 2, pixels: noisy };
  const pictures = canvasPictures(doc);
  const first = pictures.read(photo);
  assert.ok(first && first.image && first.image.kind === 'jpeg', 'a picture whose pixels all differ is a photograph');
  assert.equal(pictures.read(photo), first, 'the same source is the same picture');
  assert.equal(reads, 1, 'and is read once');
  assert.equal(jpegs, 1);
  const flat = pictures.read({ width: 2, height: 2, pixels: [1, 2, 3, 255, 1, 2, 3, 255, 1, 2, 3, 255, 1, 2, 3, 255] });
  assert.ok(flat && flat.image && flat.image.kind === 'samples' && flat.image.colors === 3, 'a flat picture keeps its pixels');
  const glass = pictures.read({ width: 2, height: 2, pixels: noisy.map((value, index) => (index === 3 ? 0 : value)) });
  assert.ok(glass && glass.image && glass.image.kind === 'samples' && glass.image.alpha, 'a picture that lets through what is behind it is no JPEG');
  const element = pictures.read({ naturalWidth: 2, naturalHeight: 2, width: 100, height: 100, pixels: noisy });
  assert.ok(element && element.width === 2 && element.height === 2, 'a picture element has the size of its file, not of its box');
  assert.equal(canvasPictures(doc, 'lossless').read({ width: 2, height: 2, pixels: noisy })?.image?.kind, 'samples', 'lossless keeps every pixel');
  assert.equal(pictures.read({ width: 2, height: 2, tainted: true }), undefined, 'pixels that may not be read');
  assert.equal(pictures.read({ width: 100000, height: 100000 }), undefined, 'a picture that is too large');
  assert.equal(pictures.read({ width: 0, height: 5 }), undefined);
  assert.equal(pictures.read({ width: 1.5, height: 5 }), undefined);
  assert.equal(pictures.read('not a picture'), undefined);
  assert.equal(pictures.read(null), undefined);
});
