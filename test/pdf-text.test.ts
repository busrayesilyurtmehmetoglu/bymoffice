// The invisible text of a PDF page: which runs a page of a document and a
// painted canvas give, in which order they are written, and where and how
// wide each one lies in the content stream.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import type { LayoutPage, PageItem, TextItem } from '../src/layout/doc/types';
import { buildPdf } from '../src/pdf/document';
import { PdfTextFont } from '../src/pdf/font';
import { MAX_RUN_CHARS, captureText, docPageRuns, orderRuns, textLayer, type PdfTextRun } from '../src/pdf/text';
import { checkPdf, parseToUnicode, poppler, popplerTest } from './pdf-fixtures';

const run = (text: string, x: number, y: number, more: Partial<PdfTextRun> = {}): PdfTextRun => ({ text, x, y, width: text.length * 6, size: 12, ...more });

/** The runs of a layer as they are written: text, matrix, size and scaling of each. */
function written(layer: string, font: PdfTextFont): { text: string; matrix: number[]; size: number; scaling: number }[] {
  const map = parseToUnicode(font.toUnicode());
  const out: { text: string; matrix: number[]; size: number; scaling: number }[] = [];
  let size = 0;
  let scaling = 100;
  for (const line of layer.split('\n')) {
    const tf = /\/F1 ([\d.]+) Tf/.exec(line);
    if (tf) size = Number(tf[1]);
    const tz = /([\d.]+) Tz/.exec(line);
    if (tz) scaling = Number(tz[1]);
    const shown = /([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) Tm <([0-9a-f]*)> Tj$/.exec(line);
    if (!shown) continue;
    let text = '';
    for (let i = 0; i < shown[7].length; i += 4) text += map.get(parseInt(shown[7].substr(i, 4), 16));
    out.push({ text, matrix: shown.slice(1, 7).map(Number), size, scaling });
  }
  return out;
}

// -------------------------------------------------------------------- order

test('runs are read from the top of the page down, and within a line from the left', () => {
  const runs = [run('world', 60, 100), run('second', 10, 130), run('Hello ', 10, 100), run('line', 60, 130)];
  assert.deepEqual(orderRuns(runs).map((entry) => entry.text), ['Hello ', 'world', 'second', 'line']);
});

test('text a little above or below the baseline stays in its line', () => {
  // A footnote mark is smaller and higher; a subscript lower.
  const runs = [run('2', 40, 95, { size: 7 }), run('x', 34, 100), run(' + y', 46, 100), run('i', 70, 103, { size: 7 }), run('next', 34, 118)];
  assert.deepEqual(orderRuns(runs).map((entry) => entry.text), ['x', '2', ' + y', 'i', 'next']);
});

test('runs that share a place keep the order they came in', () => {
  const runs = [run('a', 10, 50), run('b', 10, 50), run('c', 10, 50)];
  assert.deepEqual(orderRuns(runs).map((entry) => entry.text), ['a', 'b', 'c']);
});

test('runs without a place are left out', () => {
  const runs = [run('ok', 1, 2), run('nan', NaN, 2), run('inf', 1, Infinity), { text: 'none', x: 1, y: 2, width: 1, size: 0 }, { text: 5 as unknown as string, x: 1, y: 2, width: 1, size: 9 }, run('', 3, 4)];
  assert.deepEqual(orderRuns(runs).map((entry) => entry.text), ['ok']);
});

// -------------------------------------------------------------------- layer

test('a run is written invisible, at its baseline, as wide as it was measured', () => {
  const font = new PdfTextFont();
  // A page 1000 px high, at 0.75 pt to the px.
  const layer = textLayer([run('Hello', 100, 200, { width: 40, size: 16 })], { height: 1000, scale: 0.75, font });
  assert.match(layer, /^BT\n3 Tr\n/, 'text rendering mode 3: neither filled nor stroked');
  assert.match(layer, /\nET\n$/);
  const [hello] = written(layer, font);
  assert.equal(hello.text, 'Hello');
  assert.deepEqual(hello.matrix, [1, 0, 0, 1, 75, 600]);
  assert.equal(hello.size, 12);
  // Five characters of one em each are 60 pt; the run is 30 pt wide.
  assert.equal(hello.scaling, 50);
});

test('the width of a run counts characters, not UTF-16 units', () => {
  const font = new PdfTextFont();
  const [pair] = written(textLayer([run('😀😀', 0, 10, { width: 20, size: 10 })], { height: 100, scale: 1, font }), font);
  assert.equal(pair.text, '😀😀');
  assert.equal(pair.scaling, 100, 'two characters of 10 take 20');
});

test('a line ends with a space of its own, unless its text does', () => {
  const font = new PdfTextFont();
  const layer = written(textLayer([run('first', 10, 20, { width: 30 }), run('second ', 10, 40), run('third', 10, 60, { width: 30 })], { height: 100, scale: 1, font }), font);
  assert.deepEqual(layer.map((entry) => entry.text), ['first', ' ', 'second ', 'third', ' ']);
  // The space starts where the line ends.
  assert.deepEqual(layer[1].matrix, [1, 0, 0, 1, 40, 80]);
  assert.equal(layer[1].size, 12);
  assert.equal(layer[1].scaling, 25);
});

test('runs with room between them are separated by a space that fills it', () => {
  const font = new PdfTextFont();
  // Three cells of a row; then a word in two fonts, whose halves touch; then runs that bring their own space.
  const layer = written(
    textLayer([run('101', 10, 20), run('102', 60, 20), run('103', 110, 20), run('bo', 10, 50), run('ld ', 22, 50), run('a ', 10, 80), run('b', 60, 80), run(' c ', 100, 80)], { height: 100, scale: 1, font }),
    font,
  );
  assert.deepEqual(layer.map((entry) => entry.text), ['101', ' ', '102', ' ', '103', ' ', 'bo', 'ld ', 'a ', 'b', ' c ']);
  // The space starts where the cell before it ends (10 + 18) and is as wide as the room up to the next one (32).
  assert.deepEqual(layer[1].matrix, [1, 0, 0, 1, 28, 80]);
  assert.ok(Math.abs((layer[1].scaling * layer[1].size) / 100 - 32) < 0.01, `the space is 32 wide (${layer[1].scaling})`);
});

test('the font and the scaling are only stated when they change', () => {
  const font = new PdfTextFont();
  const layer = textLayer([run('ab', 0, 10), run('cd', 12, 10), run('ef ', 24, 10, { size: 9 })], { height: 100, scale: 1, font });
  assert.equal((layer.match(/ Tf/g) || []).length, 2);
  assert.equal((layer.match(/ Tz/g) || []).length, 2);
});

test('turned text has the matrix of its turn', () => {
  const font = new PdfTextFont();
  // Reading downwards (90 degrees clockwise) and upwards (270).
  const layer = written(textLayer([run('down ', 50, 20, { angle: 90 }), run('up ', 80, 90, { angle: 270 }), run('tilted ', 10, 95, { angle: -45 })], { height: 100, scale: 1, font }), font);
  assert.deepEqual(layer.map((entry) => entry.text), ['down ', 'up ', 'tilted ']);
  assert.deepEqual(layer[0].matrix, [0, -1, 1, 0, 50, 80]);
  assert.deepEqual(layer[1].matrix, [0, 1, -1, 0, 80, 10]);
  assert.deepEqual(layer[2].matrix, [0.7071, 0.7071, -0.7071, 0.7071, 10, 5]);
});

test('a run without a width, or with one that cannot be, still has a scaling that can', () => {
  const font = new PdfTextFont();
  const layer = written(
    textLayer([run('abcd ', 0, 10, { width: 0 }), run('efgh ', 0, 30, { width: NaN }), run('i ', 0, 50, { width: 1e5 }), run('wide text ', 0, 70, { width: 1e-9 })], { height: 100, scale: 1, font }),
    font,
  );
  assert.deepEqual(layer.map((entry) => entry.scaling), [50, 50, 10000, 1]);
});

test('runs with nothing to write give no layer at all', () => {
  const font = new PdfTextFont();
  assert.equal(textLayer([], { height: 100, scale: 1, font }), '');
  assert.equal(textLayer([run('\0\u0007', 0, 10)], { height: 100, scale: 1, font }), '', 'a run of controls only');
  assert.equal(font.codes, 0);
});

test('a very long run is cut, and its width with it', () => {
  const font = new PdfTextFont();
  const long = 'x'.repeat(MAX_RUN_CHARS * 3);
  const [cut] = written(textLayer([run(long + ' ', 0, 10, { width: 300 * MAX_RUN_CHARS + 100, size: 10 })], { height: 100, scale: 1, font }), font);
  assert.equal(cut.text.length, MAX_RUN_CHARS);
  // A third of the characters over a third of the width: 100 units per character at size 10.
  assert.ok(Math.abs(cut.scaling - 1000) < 0.5, `scaling ${cut.scaling}`);
});

// ---------------------------------------------------------------- documents

const FONT = { family: 'Test', size: 16, bold: false, italic: false };
const item = (text: string, x: number, y: number, more: Partial<TextItem> = {}): TextItem => ({ type: 'text', x, y, width: text.length * 8, text, font: FONT, color: '#000000', ...more });
const pageOf = (items: PageItem[]): LayoutPage => ({ width: 800, height: 1000, section: 0, number: 1, items });

test('the text items of a page become its runs, whatever lies between them', () => {
  const runs = docPageRuns(
    pageOf([
      { type: 'rect', x: 0, y: 0, width: 800, height: 1000, fill: '#ffffff' },
      item('Heading', 96, 120, { font: { ...FONT, size: 24, bold: true }, width: 110 }),
      { type: 'line', x1: 0, y1: 130, x2: 800, y2: 130, color: '#000000', width: 1, style: 'single' },
      item(' ', 206, 120, { width: 5 }),
      item('down', 300, 400, { rotation: 90 }),
      { type: 'image', x: 10, y: 10, width: 50, height: 50, alt: 'not text of the page' },
      { type: 'mark', x: 0, y: 0, width: 1, height: 1, baseline: 1, ascent: 1, descent: 0, source: { story: 'body', path: [0], offset: 0 }, length: 0, last: true },
    ]),
  );
  assert.deepEqual(runs, [
    { text: 'Heading', x: 96, y: 120, width: 110, size: 24 },
    { text: ' ', x: 206, y: 120, width: 5, size: 16 },
    { text: 'down', x: 300, y: 400, width: 32, size: 16, angle: 90 },
  ]);
});

test('text that is shown is in the runs: struck out by a tracked change, and the repeated header row of a table', () => {
  const runs = docPageRuns(pageOf([item('kept ', 10, 20), item('deleted', 60, 20, { deleted: true, strike: true }), item('Header', 10, 60, { repeat: true })]));
  assert.deepEqual(runs.map((entry) => entry.text), ['kept ', 'deleted', 'Header']);
});

test('items that are broken give no run and do not stop the others', () => {
  const broken = [
    item('', 10, 20),
    item('nan', NaN, 20),
    { ...item('no font', 10, 20), font: undefined as unknown as TextItem['font'] },
    { ...item('no size', 10, 20), font: { ...FONT, size: -3 } },
    { ...item('x', 10, 20), text: 12 as unknown as string },
    null as unknown as PageItem,
    item('fine', 10, 40),
  ];
  assert.deepEqual(docPageRuns(pageOf(broken)).map((entry) => entry.text), ['fine']);
  assert.deepEqual(docPageRuns({ width: 1, height: 1, section: 0, number: 1, items: undefined as unknown as PageItem[] }), []);
});

// ------------------------------------------------------------------- canvas

interface Call {
  text: string;
  x: number;
  y: number;
}

/** What captureText() needs of a canvas, with a transform that can be set and glyphs half as wide as the font is high. */
function fakeCanvas() {
  const drawn: Call[] = [];
  const ctx = {
    font: '16px Test',
    textAlign: 'left',
    textBaseline: 'alphabetic',
    direction: 'ltr',
    matrix: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
    getTransform() {
      return this.matrix;
    },
    measureText(text: string) {
      return { width: (text.length * Number((/([\d.]+)px/.exec(this.font) || [0, 0])[1])) / 2 };
    },
    fillText(text: string, x: number, y: number, _maxWidth?: number) {
      drawn.push({ text, x, y });
    },
  };
  return { ctx, drawn };
}

test('text painted on a canvas is caught with its place, and still painted', () => {
  const { ctx, drawn } = fakeCanvas();
  const runs: PdfTextRun[] = [];
  const own = ctx.fillText;
  const release = captureText(ctx, 2, (entry) => runs.push(entry));
  // The canvas has two pixels to the px of the page, and the painter scales by as much.
  ctx.matrix = { a: 2, b: 0, c: 0, d: 2, e: 20, f: 40 };
  ctx.fillText('Cell', 30, 50);
  assert.deepEqual(drawn, [{ text: 'Cell', x: 30, y: 50 }], 'the text is painted as before');
  assert.deepEqual(runs, [{ text: 'Cell', x: 40, y: 70, width: 32, size: 16 }]);
  release();
  ctx.fillText('after', 0, 0);
  assert.equal(runs.length, 1, 'nothing is caught after the release');
  assert.equal(drawn.length, 2);
  assert.equal(ctx.fillText, own, 'the canvas has its own fillText again');
  // A method of the prototype, as a real canvas has it, is no property of the canvas afterwards.
  const proto = { fillText: () => undefined };
  const real = Object.create(proto) as typeof ctx;
  captureText(real, 1, () => undefined)();
  assert.ok(!Object.prototype.hasOwnProperty.call(real, 'fillText'), 'nothing is left on the canvas');
});

test('the alignment and the baseline of the canvas are taken into account', () => {
  const { ctx } = fakeCanvas();
  const runs: PdfTextRun[] = [];
  captureText(ctx, 1, (entry) => runs.push(entry));
  ctx.font = 'italic bold 20px "Some Font", sans-serif';
  // 'abcd' is 40 wide.
  ctx.textAlign = 'center';
  ctx.fillText('abcd', 100, 50);
  ctx.textAlign = 'right';
  ctx.fillText('abcd', 100, 50);
  ctx.textAlign = 'end';
  ctx.fillText('abcd', 100, 50);
  ctx.direction = 'rtl';
  ctx.fillText('abcd', 100, 50);
  ctx.textAlign = 'start';
  ctx.fillText('abcd', 100, 50);
  assert.deepEqual(runs.map((entry) => entry.x), [80, 60, 60, 100, 60]);
  runs.length = 0;
  ctx.textAlign = 'left';
  for (const baseline of ['alphabetic', 'middle', 'top', 'hanging', 'bottom', 'ideographic']) {
    ctx.textBaseline = baseline;
    ctx.fillText('abcd', 0, 100);
  }
  assert.deepEqual(runs.map((entry) => entry.y), [100, 106, 116, 116, 96, 96]);
  assert.deepEqual(runs.map((entry) => entry.size), [20, 20, 20, 20, 20, 20]);
});

test('text painted under a turn is a turned run', () => {
  const { ctx } = fakeCanvas();
  const runs: PdfTextRun[] = [];
  captureText(ctx, 1, (entry) => runs.push(entry));
  // Turned by 90 degrees clockwise around (200, 100), and twice as large.
  ctx.matrix = { a: 0, b: 2, c: -2, d: 0, e: 200, f: 100 };
  ctx.fillText('up', 10, 5);
  assert.deepEqual(runs, [{ text: 'up', x: 190, y: 120, width: 32, size: 32, angle: 90 }]);
});

test('a width the painter limits the text to limits the run', () => {
  const { ctx } = fakeCanvas();
  const runs: PdfTextRun[] = [];
  captureText(ctx, 1, (entry) => runs.push(entry));
  ctx.fillText('squeezed', 0, 10, 20);
  assert.equal(runs[0].width, 20);
});

test('a canvas that cannot say where it paints gives no runs, and paints as before', () => {
  const { ctx, drawn } = fakeCanvas();
  const runs: PdfTextRun[] = [];
  (ctx as { getTransform?: unknown }).getTransform = undefined;
  const release = captureText(ctx, 1, (entry) => runs.push(entry));
  ctx.fillText('text', 1, 2);
  ctx.font = 'no size';
  ctx.fillText('text', 1, 2);
  release();
  assert.deepEqual(runs, []);
  assert.equal(drawn.length, 2);
  // A transform or a measurement that throws does not stop the painting either.
  const second = fakeCanvas();
  second.ctx.getTransform = () => {
    throw new Error('no transform');
  };
  captureText(second.ctx, 1, (entry) => runs.push(entry));
  second.ctx.fillText('text', 1, 2);
  assert.deepEqual(second.drawn, [{ text: 'text', x: 1, y: 2 }]);
  assert.deepEqual(runs, []);
});

test('a very long text of a cell is cut when it is caught', () => {
  const { ctx } = fakeCanvas();
  const runs: PdfTextRun[] = [];
  captureText(ctx, 1, (entry) => runs.push(entry));
  ctx.fillText('y'.repeat(MAX_RUN_CHARS * 4), 0, 10);
  assert.equal(runs[0].text.length, MAX_RUN_CHARS);
  assert.equal(runs[0].width, MAX_RUN_CHARS * 8);
});

// ------------------------------------------------------------- whole files

test('the text of a laid out page is found where it is on the page', async () => {
  const page = pageOf([item('Sol üst', 96, 96), item('sağ', 600, 96), item('İkinci satır', 96, 120), item('Dikey', 400, 500, { rotation: 270 })]);
  const file = await buildPdf([{ width: 600, height: 750, text: { runs: docPageRuns(page), width: 800 } }], { date: new Date(0) });
  const texts = checkPdf(file).pages[0].texts.filter((entry) => entry.text !== ' ');
  assert.deepEqual(texts.map((entry) => entry.text), ['Sol üst', 'sağ', 'İkinci satır', 'Dikey']);
  assert.deepEqual([texts[0].x, texts[0].y], [72, 750 - 72]);
  assert.deepEqual([texts[1].x, texts[1].y], [450, 750 - 72]);
  assert.equal(texts[0].size, 12);
  assert.ok(Math.abs(texts[0].advance - 7 * 8 * 0.75) < 0.05, `"Sol üst" is 42 pt wide (${texts[0].advance})`);
  assert.deepEqual(texts[3].matrix, [0, 1, -1, 0, 300, 375]);
});

popplerTest('Poppler finds the text in reading order, each line on its own', async () => {
  const page = pageOf([
    item('sonu.', 230, 100),
    item('Birinci satırın ', 96, 100),
    item('İkinci satır, ', 96, 124),
    item('ğüşiöç ĞÜŞİÖÇ ıI', 212, 124),
    item('Üçüncü 😀 satır', 96, 148),
  ]);
  const file = await buildPdf([{ width: 600, height: 750, text: { runs: docPageRuns(page), width: 800 } }], { date: new Date(0) });
  const report = poppler(file);
  assert.equal(report.complaints, '');
  assert.deepEqual(report.text[0].trim().split(/\s*\n/), ['Birinci satırın sonu.', 'İkinci satır, ğüşiöç ĞÜŞİÖÇ ıI', 'Üçüncü 😀 satır']);
});

// ------------------------------------------------------------ right to left

test('a run that is written from right to left is put down in the order it is seen', () => {
  const font = new PdfTextFont();
  const layer = written(
    textLayer([run('שלום עולם ', 10, 20), run('שנת 2024, 3.5 ', 10, 40), run('مرحبا ', 10, 60), run('abc שלום ', 10, 80), run('2024 ', 10, 95)], { height: 100, scale: 1, font }),
    font,
  );
  // What a reader of the file turns around again: the characters from left to right, numbers as they are read; in a run of both directions the stretch from right to left alone.
  assert.deepEqual(layer.map((entry) => entry.text), [' םלוע םולש', ' 3.5 ,2024 תנש', ' ابحرم', 'abc םולש ', '2024 ']);
});

popplerTest('Poppler gives text from right to left back as it was written', async () => {
  const lines = ['שלום עולם', 'مرحبا بالعالم', 'שנת 2024'];
  const runs = lines.map((text, index) => run(text, 96, 100 + index * 24, { width: text.length * 8, size: 16 }));
  const report = poppler(await buildPdf([{ width: 600, height: 750, text: { runs, width: 800 } }], { date: new Date(0) }));
  assert.equal(report.complaints, '');
  // The marks that say which way a line runs are Poppler's own.
  // In the order of the content stream ("raw") Poppler turns every word around and leaves the words where they are; its own reading order has both.
  assert.deepEqual(report.reading[0].replace(/[‎‏‪-‮⁦-⁩]/g, '').trim().split(/\s*\n/), lines);
});
