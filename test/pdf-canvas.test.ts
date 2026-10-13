// The canvas that writes a page of a PDF: what a painter draws on it is in
// the file as drawings, text and pictures, at the place and in the colour it
// was drawn - read back by a reader that shares no code with the engine
// (checkPdf) and painted by Poppler - and whatever it cannot write exactly
// makes the page unsupported instead of wrong.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { PDF_CANVAS_MEMBERS, createPdfCanvas, parseCssColor, type PdfCanvas, type PdfCanvasMeasure, type PdfPicture } from '../src/pdf/canvas';
import { buildPdf } from '../src/pdf/document';
import { PdfFontSet, parseCssFont } from '../src/pdf/font-set';
import { PdfError } from '../src/pdf/writer';
import { checkPdf, poppler, popplerPixels, popplerTest, popplerWords, type CheckedPdf } from './pdf-fixtures';
import { bonum, testFont } from './pdf-font-fixtures';

const DATE = new Date(Date.UTC(2026, 9, 13, 12, 0, 0));
const TEST_FONT = testFont();
const BONUM = { regular: bonum('regular'), bold: bonum('bold'), italic: bonum('italic'), bolditalic: bonum('bolditalic') };

function fontSet(more: { family: string; bytes: Uint8Array; weight?: number; style?: 'normal' | 'italic' }[] = []): PdfFontSet {
  return new PdfFontSet(
    [
      { family: 'Test', weight: 400, style: 'normal', bytes: TEST_FONT },
      { family: 'TeX Gyre Bonum', weight: 400, style: 'normal', bytes: BONUM.regular },
      { family: 'TeX Gyre Bonum', weight: 700, style: 'normal', bytes: BONUM.bold },
      { family: 'TeX Gyre Bonum', weight: 400, style: 'italic', bytes: BONUM.italic },
      { family: 'TeX Gyre Bonum', weight: 700, style: 'italic', bytes: BONUM.bolditalic },
      ...more.map((font) => ({ family: font.family, weight: font.weight || 400, style: font.style || ('normal' as const), bytes: font.bytes })),
    ],
    // The device of the tests has Arial and nothing else.
    (family) => family === 'Arial',
  );
}

/** Measures text as the font says it is wide: what a canvas without kerning would say. `more` is added per character. */
function measureWith(fonts: PdfFontSet, more = 0, baseline?: (font: string, baseline: string) => number): PdfCanvasMeasure {
  return {
    width: (font, text) => {
      const resolved = fonts.resolve(font);
      if (!('face' in resolved)) return text.length * 5;
      let units = 0;
      // As a browser: a tab or a line end is a space, a soft hyphen nothing.
      for (const char of text.replace(/[\t\n\r]/g, ' ').replace(/\u00ad/g, '')) {
        const code = char.codePointAt(0) as number;
        units += resolved.face.advance(resolved.face.glyph(code === 0xa0 && resolved.face.glyph(code) === 0 ? 0x20 : code));
      }
      return (units * resolved.size) / resolved.face.unitsPerEm + more * Array.from(text).length;
    },
    baseline,
  };
}

interface Made {
  ctx: CanvasRenderingContext2D;
  canvas: PdfCanvas;
  file(): Promise<Uint8Array>;
  checked(): Promise<CheckedPdf>;
}

function make(options: { width?: number; height?: number; pageWidth?: number; pageHeight?: number; fonts?: PdfFontSet; measure?: PdfCanvasMeasure; pictures?: Map<unknown, PdfPicture> } = {}): Made {
  const fonts = options.fonts || fontSet();
  const pageWidth = options.pageWidth || 200;
  const pageHeight = options.pageHeight || 100;
  const pictures = options.pictures;
  const made = createPdfCanvas({
    width: options.width || pageWidth,
    height: options.height || pageHeight,
    pageWidth,
    pageHeight,
    fonts,
    measure: options.measure || measureWith(fonts),
    images: pictures ? { read: (source) => pictures.get(source) } : undefined,
  });
  const file = () => buildPdf([{ width: pageWidth, height: pageHeight, vector: made.canvas.finish() }], { date: DATE });
  return { ctx: made.context, canvas: made.canvas, file, checked: async () => checkPdf(await file()) };
}

const near = (actual: number[], expected: number[], slack: number, what: string): void => {
  assert.equal(actual.length, expected.length, what);
  for (let i = 0; i < actual.length; i++) assert.ok(Math.abs(actual[i] - expected[i]) <= slack, `${what}: ${JSON.stringify(actual)} is not ${JSON.stringify(expected)} (±${slack})`);
};

const WHITE = [255, 255, 255];

// ---------------------------------------------------------------------------
// Paths, fills, strokes

test('pdf canvas: a filled rectangle is a rectangle of the page, y axis turned upwards', async () => {
  const made = make();
  made.ctx.fillStyle = '#ff0000';
  made.ctx.fillRect(20, 10, 60, 30);
  const page = (await made.checked()).pages[0];
  assert.equal(page.shapes.length, 1);
  assert.deepEqual(page.shapes[0], { operator: 'f', path: [['re', 20, 60, 60, 30]], colour: [1, 0, 0], alpha: 1, lineWidth: 1, dash: [] });
  assert.equal(page.images.length, 0, 'no picture in a drawn page');
});

popplerTest('pdf canvas: Poppler paints the rectangle where the canvas has it, in its colour', async () => {
  const made = make();
  made.ctx.fillStyle = '#ff0000';
  made.ctx.fillRect(20, 10, 60, 30);
  made.ctx.fillStyle = 'rgb(0, 128, 255)';
  made.ctx.fillRect(120, 60, 40, 20);
  const file = await made.file();
  const pixels = popplerPixels(file);
  assert.equal(pixels.width, 200);
  assert.equal(pixels.height, 100);
  near(pixels.at(50, 25), [255, 0, 0], 2, 'inside the first rectangle');
  near(pixels.at(22, 12), [255, 0, 0], 2, 'its top left corner');
  near(pixels.at(78, 38), [255, 0, 0], 2, 'its bottom right corner');
  near(pixels.at(17, 25), WHITE, 2, 'left of it');
  near(pixels.at(50, 43), WHITE, 2, 'below it');
  near(pixels.at(140, 70), [0, 128, 255], 2, 'inside the second rectangle');
  assert.equal(poppler(file).complaints, '', 'Poppler has nothing to complain of');
});

popplerTest('pdf canvas: the transform of the canvas places what is drawn: scale, translate, rotate, setTransform', async () => {
  // Two pixels of the canvas per point of the page, as a page painted at 144 dots per inch.
  const made = make({ width: 400, height: 200 });
  const ctx = made.ctx;
  ctx.setTransform(2, 0, 0, 2, 0, 0);
  ctx.fillStyle = '#00ff00';
  ctx.fillRect(20, 10, 60, 30);
  ctx.save();
  ctx.translate(150, 50);
  ctx.rotate(Math.PI / 2);
  ctx.fillStyle = '#0000ff';
  // Turned a quarter: 30 along the x axis of the canvas is 30 down the page.
  ctx.fillRect(0, 0, 30, 10);
  ctx.restore();
  ctx.scale(0.5, 0.5);
  ctx.fillRect(20, 160, 40, 20);
  const pixels = popplerPixels(await made.file());
  near(pixels.at(50, 25), [0, 255, 0], 2, 'the scaled rectangle');
  near(pixels.at(85, 25), WHITE, 2, 'right of it');
  near(pixels.at(145, 65), [0, 0, 255], 2, 'the turned rectangle');
  near(pixels.at(155, 65), WHITE, 2, 'right of the turned rectangle');
  near(pixels.at(145, 45), WHITE, 2, 'above the turned rectangle');
  near(pixels.at(145, 85), WHITE, 2, 'below the turned rectangle');
  // The fill is that of before the save again, at half the scale.
  near(pixels.at(20, 85), [0, 255, 0], 2, 'the rectangle drawn after the state was given back');
  near(pixels.at(35, 85), WHITE, 2, 'right of it');
});

test('pdf canvas: a line is as wide on the page as on the canvas, with its dashes, caps and joins', async () => {
  const made = make({ width: 400, height: 200 });
  const ctx = made.ctx;
  ctx.setTransform(2, 0, 0, 2, 0, 0);
  ctx.strokeStyle = '#0000ff';
  ctx.lineWidth = 4;
  ctx.setLineDash([10, 5]);
  ctx.lineCap = 'round';
  ctx.lineJoin = 'bevel';
  ctx.beginPath();
  ctx.moveTo(20, 50);
  ctx.lineTo(180, 50);
  ctx.stroke();
  ctx.setLineDash([]);
  ctx.strokeRect(10, 10, 20, 20);
  const page = (await made.checked()).pages[0];
  assert.equal(page.shapes.length, 2);
  assert.deepEqual(page.shapes[0], { operator: 'S', path: [['m', 20, 50], ['l', 180, 50]], colour: [0, 0, 1], alpha: 1, lineWidth: 4, dash: [10, 5] });
  assert.deepEqual(page.shapes[1].path, [['re', 10, 70, 20, 20]]);
  assert.deepEqual(page.shapes[1].dash, [], 'the dashes are taken off again');
  assert.match(page.content, /\b1 J\b/, 'round caps');
  assert.match(page.content, /\b2 j\b/, 'bevelled joins');
});

popplerTest('pdf canvas: Poppler paints the dashes of a line where they are', async () => {
  const made = make();
  const ctx = made.ctx;
  ctx.strokeStyle = '#0000ff';
  ctx.lineWidth = 6;
  ctx.setLineDash([10, 6]);
  ctx.beginPath();
  ctx.moveTo(20, 50);
  ctx.lineTo(180, 50);
  ctx.stroke();
  const pixels = popplerPixels(await made.file());
  near(pixels.at(25, 50), [0, 0, 255], 2, 'in the first dash');
  near(pixels.at(33, 50), WHITE, 2, 'in the first gap');
  near(pixels.at(41, 50), [0, 0, 255], 2, 'in the second dash');
  near(pixels.at(25, 44), WHITE, 2, 'above the line');
  near(pixels.at(25, 56), WHITE, 2, 'below the line');
});

popplerTest('pdf canvas: a clip holds until the state it was made in is given back', async () => {
  const made = make();
  const ctx = made.ctx;
  ctx.save();
  ctx.beginPath();
  ctx.rect(50, 20, 40, 40);
  ctx.clip();
  ctx.fillStyle = '#008000';
  ctx.fillRect(0, 0, 200, 100);
  ctx.restore();
  ctx.fillStyle = '#ff0000';
  ctx.fillRect(0, 0, 10, 10);
  const file = await made.file();
  assert.equal(checkPdf(file).pages[0].clips.length, 1);
  const pixels = popplerPixels(file);
  near(pixels.at(70, 40), [0, 128, 0], 2, 'inside the clip');
  near(pixels.at(30, 40), WHITE, 2, 'left of the clip');
  near(pixels.at(70, 70), WHITE, 2, 'below the clip');
  near(pixels.at(5, 5), [255, 0, 0], 2, 'drawn after the clip was given back');
});

popplerTest('pdf canvas: colours that let the paper through, and a canvas that does', async () => {
  const made = make();
  const ctx = made.ctx;
  ctx.fillStyle = 'rgba(255, 0, 0, 0.5)';
  ctx.fillRect(10, 10, 40, 40);
  ctx.globalAlpha = 0.25;
  ctx.fillStyle = '#000000';
  ctx.fillRect(60, 10, 40, 40);
  ctx.strokeStyle = '#0000ff';
  ctx.lineWidth = 10;
  ctx.beginPath();
  ctx.moveTo(110, 30);
  ctx.lineTo(190, 30);
  ctx.stroke();
  ctx.globalAlpha = 1;
  ctx.fillRect(10, 60, 40, 30);
  const file = await made.file();
  const page = checkPdf(file).pages[0];
  assert.deepEqual(
    page.shapes.map((shape) => shape.alpha),
    [0.5, 0.25, 0.25, 1],
  );
  const pixels = popplerPixels(file);
  near(pixels.at(30, 30), [255, 128, 128], 3, 'half a red over white');
  near(pixels.at(80, 30), [191, 191, 191], 3, 'a quarter of black over white');
  near(pixels.at(150, 30), [191, 191, 255], 3, 'a quarter of a blue line over white');
  near(pixels.at(30, 75), [0, 0, 0], 3, 'opaque again');
});

test('pdf canvas: what cannot be seen is not written', async () => {
  const made = make();
  const ctx = made.ctx;
  ctx.fillStyle = 'rgba(0, 0, 0, 0)';
  ctx.fillRect(10, 10, 40, 40);
  ctx.fillStyle = 'transparent';
  ctx.fillRect(10, 10, 40, 40);
  ctx.globalAlpha = 0;
  ctx.fillStyle = '#000000';
  ctx.fillRect(10, 10, 40, 40);
  ctx.fillRect(10, 10, 0, 40);
  ctx.fillRect(NaN, 10, 10, 40);
  assert.equal(made.canvas.unsupported, undefined);
  assert.equal((await made.checked()).pages[0].shapes.length, 0);
});

test('pdf canvas: white on paper nothing was drawn on yet is not written: an empty page has no content', async () => {
  const empty = make();
  const ctx = empty.ctx;
  // What the painters of pages do first: the paper.
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, 200, 100);
  ctx.save();
  ctx.scale(2, 2);
  ctx.fillStyle = '#FFFFFF';
  ctx.fillRect(0, 0, 100, 50);
  ctx.beginPath();
  ctx.rect(0, 0, 10, 10);
  ctx.clip();
  ctx.restore();
  assert.equal(empty.canvas.finish().content, '', 'nothing at all is written');
  const page = (await empty.checked()).pages[0];
  assert.equal(page.content, '');
  assert.equal(page.shapes.length, 0);
  // Once something is on the paper, white covers it and is written.
  const covered = make();
  covered.ctx.fillStyle = '#ffffff';
  covered.ctx.fillRect(0, 0, 200, 100);
  covered.ctx.fillStyle = '#ff0000';
  covered.ctx.fillRect(10, 10, 50, 50);
  covered.ctx.fillStyle = '#ffffff';
  covered.ctx.fillRect(20, 20, 10, 10);
  assert.deepEqual(
    (await covered.checked()).pages[0].shapes.map((shape) => shape.colour),
    [[1, 0, 0], [1]],
    'the red box and the white one over it; not the paper before them',
  );
  // White that lets the paper through is no paper, and neither is a white line or white text: they are written.
  const glass = make();
  glass.ctx.fillStyle = 'rgba(255, 255, 255, 0.5)';
  glass.ctx.fillRect(0, 0, 10, 10);
  assert.equal((await glass.checked()).pages[0].shapes.length, 1);
});

test('pdf canvas: arcs and curves are the curves of the page', async () => {
  const made = make();
  const ctx = made.ctx;
  ctx.beginPath();
  // A quarter of a circle, clockwise on the canvas from the right to the bottom.
  ctx.arc(100, 50, 30, 0, Math.PI / 2);
  ctx.quadraticCurveTo(70, 80, 40, 50);
  ctx.bezierCurveTo(40, 30, 60, 20, 80, 20);
  ctx.closePath();
  ctx.fill();
  const path = (await made.checked()).pages[0].shapes[0].path;
  assert.deepEqual(path[0], ['m', 130, 50], 'the arc starts at its first point');
  const reach = (4 / 3) * Math.tan(Math.PI / 8) * 30;
  near(path[1].slice(1) as number[], [130, 50 - reach, 100 + reach, 20, 100, 20], 0.002, 'a quarter turn as one curve');
  // From (100, 80) through the control point (70, 80) to (40, 50), as a cubic curve.
  near(path[2].slice(1) as number[], [80, 20, 60, 30, 40, 50], 0.002, 'the quadratic curve');
  near(path[3].slice(1) as number[], [40, 70, 60, 80, 80, 80], 0.002, 'the cubic curve');
  assert.deepEqual(path[4], ['h']);
});

popplerTest('pdf canvas: a full circle, a ring by the even-odd rule, and an arc drawn the other way round', async () => {
  const made = make();
  const ctx = made.ctx;
  ctx.fillStyle = '#ff0000';
  ctx.beginPath();
  ctx.arc(50, 50, 30, 0, Math.PI * 2);
  ctx.fill();
  ctx.fillStyle = '#0000ff';
  ctx.beginPath();
  ctx.arc(150, 50, 30, 0, Math.PI * 2);
  ctx.moveTo(160, 50);
  ctx.arc(150, 50, 10, 0, Math.PI * 2);
  ctx.fill('evenodd');
  const file = await made.file();
  assert.deepEqual(
    checkPdf(file).pages[0].shapes.map((shape) => shape.operator),
    ['f', 'f*'],
  );
  const pixels = popplerPixels(file);
  near(pixels.at(50, 50), [255, 0, 0], 2, 'the middle of the circle');
  near(pixels.at(50, 24), [255, 0, 0], 2, 'near its top');
  near(pixels.at(26, 26), WHITE, 2, 'the corner of the box around it');
  near(pixels.at(150, 50), WHITE, 2, 'the hole of the ring');
  near(pixels.at(150, 30), [0, 0, 255], 2, 'the ring');
  // The upper half only, drawn counter-clockwise from the right to the left.
  const half = make();
  half.ctx.beginPath();
  half.ctx.arc(100, 50, 30, 0, Math.PI, true);
  half.ctx.fill();
  const halfPixels = popplerPixels(await half.file());
  near(halfPixels.at(100, 35), [0, 0, 0], 2, 'the upper half');
  near(halfPixels.at(100, 65), WHITE, 2, 'the lower half');
});

// ---------------------------------------------------------------------------
// Text

test('pdf canvas: text is text in the font it is painted with, at its baseline', async () => {
  const made = make();
  made.ctx.font = '20px Test';
  made.ctx.fillStyle = '#102030';
  made.ctx.fillText('AB', 30, 60);
  const checked = await made.checked();
  const page = checked.pages[0];
  assert.equal(page.texts.length, 1);
  const text = page.texts[0];
  assert.equal(text.text, 'AB');
  assert.deepEqual(text.matrix, [1, 0, 0, 1, 30, 40]);
  assert.equal(text.size, 20);
  assert.equal(text.mode, 0, 'the text is filled, not hidden');
  assert.equal(text.spacing, 0);
  near([text.advance], [24], 0.001, 'the advance of "A" and "B" at 20 px');
  near(text.fill, [16 / 255, 32 / 255, 48 / 255], 0.001, 'the colour of the text');
  assert.deepEqual(text.glyphs, [2, 3], 'the codes are the glyph numbers of the font');
  assert.match(text.font, /^[A-Z]{6}\+BymTest-Regular$/);
  assert.equal(checked.fonts.length, 1);
  assert.equal(checked.fonts[0].kind, 'truetype');
  assert.deepEqual(Array.from(checked.fonts[0].widths.entries()), [
    [2, 700],
    [3, 500],
  ]);
  assert.equal(checked.fonts[0].mapped, 2);
});

popplerTest('pdf canvas: Poppler paints the glyphs where the canvas has them and reads the text back, Turkish letters too', async () => {
  const made = make();
  made.ctx.font = '20px Test';
  // "A" is a box from the baseline 14 px up, 1 px to 13 px from where it starts; "B" a triangle.
  made.ctx.fillText('AB', 30, 60);
  made.ctx.fillText('şİğ C', 100, 60);
  const file = await made.file();
  const pixels = popplerPixels(file);
  near(pixels.at(37, 53), [0, 0, 0], 2, 'inside the box of "A"');
  near(pixels.at(37, 44), WHITE, 2, 'above it');
  near(pixels.at(37, 62), WHITE, 2, 'below the baseline');
  near(pixels.at(28, 53), WHITE, 2, 'left of it');
  near(pixels.at(49, 57), [0, 0, 0], 2, 'inside the triangle of "B"');
  near(pixels.at(46, 48), WHITE, 2, 'beside the tip of the triangle');
  const report = poppler(file);
  assert.equal(report.complaints, '');
  assert.equal(report.text[0].replace(/\s+/g, ' ').trim(), 'AB şİğ C');
  assert.equal(report.fonts.length, 1);
  assert.match(report.fonts[0].name, /^[A-Z]{6}\+BymTest-Regular$/);
  assert.deepEqual({ ...report.fonts[0], name: '' }, { name: '', type: 'CID TrueType', embedded: true, subset: true, unicode: true });
  // The words stand where they were painted.
  const words = popplerWords(file).words;
  assert.deepEqual(
    words.map((word) => word.text),
    ['AB', 'şİğ', 'C'],
  );
  near([words[0].xMin, words[0].xMax], [30, 54], 0.1, 'the box of the first word');
  near([words[1].xMin, words[1].xMax], [100, 100 + (450 + 350 + 550) * 0.02], 0.1, 'the box of the second word');
});

test('pdf canvas: every glyph of a run stands where the canvas has it: what kerning moves it by is in the text', async () => {
  const fonts = fontSet();
  const plain = measureWith(fonts);
  // A canvas that kerns: "A" before "B" stands 1.5 px closer, "B" before "A" 0.5 px further.
  const kerned: PdfCanvasMeasure = { width: (font, text) => plain.width(font, text) - 1.5 * (text.match(/AB/g) || []).length + 0.5 * (text.match(/BA/g) || []).length };
  const made = make({ fonts, measure: kerned });
  made.ctx.font = '20px Test';
  made.ctx.fillText('ABAB C', 30, 60);
  made.ctx.fillText('AC', 30, 90);
  const texts = (await made.checked()).pages[0].texts;
  assert.equal(texts[0].text, 'ABAB C');
  // 14, 10, 14, 10 and 6 px are the advances of the glyphs at 20 px.
  near(texts[0].starts, [0, 14 - 1.5, 24 - 1.5 + 0.5, 38 - 3 + 0.5, 48 - 3 + 0.5, 54 - 3 + 0.5], 0.003, 'where each glyph starts');
  near([texts[0].advance], [kerned.width('20px Test', 'ABAB C')], 0.003, 'the run is as wide as it was measured');
  assert.equal(texts[0].spacing, 0, 'no spacing is spread over the run');
  assert.match((await made.checked()).pages[0].content, /\] TJ/, 'the moves stand between the strings');
  // A run nothing is moved in is one string.
  assert.deepEqual(texts[1].starts, [0, 14]);
  assert.match((await made.checked()).pages[0].content, /> Tj/);
});

test('pdf canvas: a run of more glyphs than are placed as one piece is written in pieces, each where the canvas has it', async () => {
  const fonts = fontSet();
  const made = make({ fonts, measure: measureWith(fonts, -0.3), pageWidth: 6000 });
  made.ctx.font = '20px Test';
  // No space to cut at: the pieces are as long as a piece may be.
  const long = 'AB'.repeat(200);
  made.ctx.fillText(long, 0, 60);
  const texts = (await made.checked()).pages[0].texts;
  assert.deepEqual(texts.map((text) => text.glyphs.length), [256, 144]);
  assert.ok(texts.every((text) => text.spacing === 0), 'nothing is spread over the run');
  // Every glyph ends 0.3 px before where its font alone would end it. The first glyph of the second piece stands
  // where the text up to and with it ends, less its own advance: 257 glyphs, each 0.3 px short.
  near([texts[1].x], [128 * 24 - 257 * 0.3], 0.01, 'where the second piece starts');
  // With spaces the cut is after one.
  const spaced = make({ fonts, measure: measureWith(fonts), pageWidth: 6000 });
  spaced.ctx.font = '20px Test';
  spaced.ctx.fillText('AB AB '.repeat(60), 0, 60);
  const pieces = (await spaced.checked()).pages[0].texts;
  assert.deepEqual(pieces.map((text) => text.glyphs.length), [255, 105]);
  assert.equal(pieces[0].text.slice(-1), ' ', 'the first piece ends with the space it was cut after');
  // "AB " is 30 px at 20 px: 85 of them stand before the second piece.
  near([pieces[1].x], [85 * 30], 0.01, 'the second piece starts where the canvas has its first glyph');
  near([pieces[1].x + pieces[1].advance], [120 * 30], 0.01, 'and the run ends where the canvas ends it');
});

test('pdf canvas: a run that is far from as wide as its glyphs is shaped by the browser, and the page is not written', () => {
  const fonts = fontSet();
  for (const more of [-1.5, 1.5]) {
    const made = make({ fonts, measure: measureWith(fonts, more) });
    made.ctx.font = '20px Test';
    made.ctx.fillText('ABA', 30, 60);
    assert.match(String(made.canvas.unsupported), /not as wide as its glyphs/, `${more} px per glyph`);
  }
  // One glyph that stands far from where its font puts it: two letters the browser drew as something else.
  const plain = measureWith(fonts);
  const jump = make({ fonts, measure: { width: (font, text) => plain.width(font, text) + (text.length >= 2 ? 7 : 0) - (text.length >= 3 ? 7 : 0) } });
  jump.ctx.font = '20px Test';
  jump.ctx.fillText('ABAB', 30, 60);
  assert.match(String(jump.canvas.unsupported), /not as wide as its glyphs/, 'a run of the right width with a glyph out of place');
  const broken = make({ fonts, measure: { width: () => NaN } });
  broken.ctx.font = '20px Test';
  broken.ctx.fillText('A', 0, 0);
  assert.match(String(broken.canvas.unsupported), /could not be measured/);
  const throwing = make({
    fonts,
    measure: {
      width: () => {
        throw new Error('no canvas');
      },
    },
  });
  throwing.ctx.font = '20px Test';
  throwing.ctx.fillText('A', 0, 0);
  assert.match(String(throwing.canvas.unsupported), /could not be measured/);
});

test('pdf canvas: text is placed by its alignment, its baseline and the spacing of its letters', async () => {
  const fonts = fontSet();
  const made = make({ fonts, measure: measureWith(fonts, 0, (_font, baseline) => (baseline === 'middle' ? 5 : baseline === 'top' ? 16 : NaN)) });
  const ctx = made.ctx;
  ctx.font = '20px Test';
  ctx.textAlign = 'center';
  ctx.fillText('A', 100, 60);
  ctx.textAlign = 'right';
  ctx.fillText('A', 100, 60);
  ctx.textAlign = 'left';
  ctx.textBaseline = 'middle';
  ctx.fillText('A', 100, 60);
  ctx.textBaseline = 'top';
  ctx.fillText('A', 100, 60);
  ctx.textBaseline = 'alphabetic';
  (ctx as unknown as { letterSpacing: string }).letterSpacing = '2px';
  assert.equal(ctx.measureText('AB').width, 24 + 4, 'measuring counts the spacing');
  ctx.textAlign = 'end';
  ctx.fillText('AB', 100, 60);
  const texts = (await made.checked()).pages[0].texts;
  assert.deepEqual(
    texts.map((text) => [text.x, text.y, text.spacing]),
    [
      [93, 40, 0],
      [86, 40, 0],
      [100, 35, 0],
      [100, 24, 0],
      [72, 40, 2],
    ],
  );
  // A baseline that cannot be measured is none text can be placed by.
  const blind = make({ fonts });
  blind.ctx.font = '20px Test';
  blind.ctx.textBaseline = 'middle';
  blind.ctx.fillText('A', 100, 60);
  assert.match(String(blind.canvas.unsupported), /middle line/);
});

popplerTest('pdf canvas: turned and stretched text stands as the canvas turns and stretches it', async () => {
  const made = make();
  const ctx = made.ctx;
  ctx.font = '20px Test';
  ctx.save();
  ctx.translate(100, 20);
  ctx.rotate(Math.PI / 2);
  ctx.fillText('A', 0, 0);
  ctx.restore();
  ctx.save();
  ctx.translate(20, 80);
  ctx.scale(2, 1);
  ctx.fillText('A', 0, 0);
  ctx.restore();
  const file = await made.file();
  const texts = checkPdf(file).pages[0].texts;
  near(texts[0].matrix, [0, -1, 1, 0, 100, 80], 0.0001, 'a quarter turn clockwise');
  near(texts[1].matrix, [2, 0, 0, 1, 20, 20], 0.0001, 'twice as wide');
  const pixels = popplerPixels(file);
  // The box of "A" reads downwards: 14 px to the right of its baseline, from 1 px to 13 px down.
  near(pixels.at(107, 27), [0, 0, 0], 2, 'inside the turned glyph');
  near(pixels.at(97, 27), WHITE, 2, 'on the other side of its baseline');
  near(pixels.at(107, 36), WHITE, 2, 'after its end');
  // Stretched: 2 px to 26 px from its start, 14 px high.
  near(pixels.at(44, 73), [0, 0, 0], 2, 'inside the stretched glyph');
  near(pixels.at(49, 73), WHITE, 2, 'right of the stretched glyph');
});

test('pdf canvas: the font of a text is the first of its list that can be shown', () => {
  const fonts = fontSet();
  const size = (font: string): unknown => {
    const resolved = fonts.resolve(font);
    return 'face' in resolved ? [resolved.face.name, resolved.size] : resolved.reason;
  };
  assert.deepEqual(size('20px Test'), ['BymTest-Regular', 20]);
  assert.deepEqual(size('italic bold 13.5px "TeX Gyre Bonum", serif'), ['TeXGyreBonum-BoldItalic', 13.5]);
  assert.deepEqual(size('700 12px "tex gyre bonum"'), ['TeXGyreBonum-Bold', 12]);
  // A family the device does not have is passed over, as a browser passes over it.
  assert.deepEqual(size('16px "Bookman Old Style", "TeX Gyre Bonum", serif'), ['TeXGyreBonum-Regular', 16]);
  assert.match(String(size('16px Arial, "TeX Gyre Bonum"')), /font of the device \("Arial"\)/);
  assert.match(String(size('16px serif')), /font of the device/);
  assert.match(String(size('16px "Nothing Here"')), /no font of the text is known/);
  // A face the family does not have would be made up by the browser: slanted, or emboldened.
  assert.match(String(size('bold 16px Test')), /no bold face/);
  assert.match(String(size('italic 16px Test')), /no italic face/);
  assert.match(String(size('16pt Test')), /form that is not read/);
  assert.match(String(size('oblique 16px Test')), /form that is not read/);
  assert.match(String(size('16px')), /form that is not read/);
  assert.deepEqual(parseCssFont('italic 600 12px/normal "A, B", C D , \'E\''), { style: 'italic', weight: 600, size: 12, families: ['A, B', 'C D', 'E'] });
  assert.equal(parseCssFont('12px "unclosed'), undefined);
  assert.equal(parseCssFont('small-caps 12px X'), undefined);
});

test('pdf canvas: a font that may not be embedded, or cannot be read, leaves its text to the picture', () => {
  const restricted = fontSet([
    { family: 'Restricted', bytes: testFont({ fsType: 0x0002 }) },
    { family: 'Bitmaps', bytes: testFont({ fsType: 0x0200 }) },
    { family: 'Silent', bytes: testFont({ os2: false }) },
    { family: 'Broken', bytes: new Uint8Array(100) },
    { family: 'Woff', bytes: Uint8Array.from([0x77, 0x4f, 0x46, 0x32, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]) },
  ]);
  const why = (family: string): string => {
    const made = make({ fonts: restricted });
    made.ctx.font = `20px ${family}`;
    made.ctx.fillText('A', 10, 50);
    return String(made.canvas.unsupported);
  };
  assert.match(why('Restricted'), /cannot be embedded: the licence of the font does not allow embedding/);
  assert.match(why('Bitmaps'), /only its bitmaps/);
  assert.match(why('Silent'), /does not say whether it may be embedded/);
  assert.match(why('Broken'), /cannot be embedded: The file is no TrueType or OpenType font/);
  assert.match(why('Woff'), /WOFF2/);
});

test('pdf canvas: a font that may only be embedded whole is embedded whole, under its own name', async () => {
  const bytes = testFont({ fsType: 0x0100, name: 'Whole-Font' });
  const made = make({ fonts: fontSet([{ family: 'Whole', bytes }]) });
  made.ctx.font = '20px Whole';
  made.ctx.fillText('A', 10, 50);
  const checked = await made.checked();
  assert.equal(checked.fonts[0].baseFont, 'Whole-Font');
  assert.deepEqual(Array.from(checked.fonts[0].file), Array.from(bytes), 'the font file as it was given');
});

test('pdf canvas: characters the font cannot show as the browser does make the page unsupported', () => {
  const why = (text: string, font = '20px Test'): string | undefined => {
    const made = make();
    made.ctx.font = font;
    made.ctx.fillText(text, 10, 50);
    return made.canvas.unsupported;
  };
  assert.equal(why('AB C'), undefined);
  assert.match(String(why('AZ')), /BymTest-Regular has no glyph for U\+5A/);
  assert.match(String(why('₺', '20px "TeX Gyre Bonum"')), /no glyph for U\+20BA/);
  assert.match(String(why('😀', '20px "TeX Gyre Bonum"')), /no glyph for U\+1F600/);
  assert.match(String(why('سلام', '20px "TeX Gyre Bonum"')), /rules that are not followed/);
  assert.match(String(why('e\u0301', '20px "TeX Gyre Bonum"')), /rules that are not followed/);
  assert.match(String(why('A\u200dB')), /rules that are not followed/);
  assert.match(String(why('A\u0001B')), /control character/);
  assert.match(String(why('\ud800', '20px "TeX Gyre Bonum"')), /rules that are not followed/, 'half a pair of surrogates');
  // A tab and a line end are spaces; a soft hyphen is nothing.
  assert.equal(why('A\tB\nC\u00ad'), undefined);
  assert.equal(why('A\u00a0B'), undefined, 'a space that does not break, in a font without a glyph of its own for it');
});

test('pdf canvas: tabs are spaces and soft hyphens nothing in what is written', async () => {
  const fonts = fontSet();
  // The canvas the widths come from does the same: a tab as wide as a space, a soft hyphen as nothing.
  const measure: PdfCanvasMeasure = { width: (font, text) => measureWith(fonts).width(font, text.replace(/\t/g, ' ').replace(/\u00ad/g, '')) };
  const made = make({ fonts, measure });
  made.ctx.font = '20px Test';
  made.ctx.fillText('A\tB\u00ad', 10, 50);
  const text = (await made.checked()).pages[0].texts[0];
  assert.equal(text.text, 'A B');
  assert.deepEqual(text.glyphs, [2, 1, 3]);
});

test('pdf canvas: a glyph that stands for two characters is copied as the plainer of them, whichever was written first', async () => {
  // A font whose space is also its space that does not break, and whose "A" is also the Greek capital alpha.
  const bytes = testFont({ map: { 0x20: 1, 0xa0: 1, 0x41: 2, 0x391: 2, 0x42: 3 }, name: 'Twice' });
  const fonts = fontSet([{ family: 'Twice', bytes }]);
  const made = make({ fonts });
  made.ctx.font = '20px Twice';
  made.ctx.fillText('\u0391\u00a0B', 10, 30);
  made.ctx.fillText('A B', 10, 60);
  const page = (await made.checked()).pages[0];
  assert.deepEqual(page.texts.map((text) => text.glyphs), [[2, 1, 3], [2, 1, 3]], 'the same glyphs both times');
  assert.deepEqual(page.texts.map((text) => text.text), ['A B', 'A B'], 'read back as the letter and the space');
});

test('pdf canvas: text of a file cannot become syntax of the PDF', async () => {
  // A page wide enough for all of it: what lies beyond the paper is not written.
  const made = make({ pageWidth: 400 });
  made.ctx.font = '16px "TeX Gyre Bonum"';
  const hostile = ') Tj ET /X Do <</A 1>> (\\) % [x] BT';
  made.ctx.fillText(hostile, 5, 50);
  assert.equal(made.canvas.unsupported, undefined);
  const page = (await made.checked()).pages[0];
  assert.equal(page.text, hostile);
  assert.equal((page.content.match(/Tj/g) || []).length, 1, 'one string is shown');
  assert.ok(!/[()%\\]|Do|<</.test(page.content), 'none of the characters of the text stands in the content as itself');
});

popplerTest('pdf canvas: text in the font of the package: embedded whole with its CFF outlines, found and read back by Poppler', async () => {
  const made = make({ pageWidth: 300, pageHeight: 120 });
  const ctx = made.ctx;
  ctx.font = '24px "TeX Gyre Bonum"';
  const text = 'Çığ gölü şöyle';
  ctx.fillText(text, 20, 50);
  ctx.font = 'italic bold 12px "Bookman Old Style", "TeX Gyre Bonum", serif';
  ctx.fillText('İstanbul’da ağaç', 20, 90);
  const file = await made.file();
  const checked = checkPdf(file);
  assert.equal(checked.fonts.length, 2);
  assert.deepEqual(checked.fonts.map((font) => font.baseFont).sort(), ['TeXGyreBonum-BoldItalic', 'TeXGyreBonum-Regular'], 'the fonts keep their names: they are not cut down');
  assert.ok(
    checked.fonts.every((font) => font.kind === 'cff'),
    'CFF outlines',
  );
  const report = poppler(file);
  assert.equal(report.complaints, '', 'Poppler has nothing to complain of');
  assert.deepEqual(report.text[0].trim().split('\n'), [text, 'İstanbul’da ağaç']);
  assert.deepEqual(
    report.fonts.map((font) => [font.name, font.type, font.embedded, font.subset, font.unicode]).sort(),
    [
      ['TeXGyreBonum-BoldItalic', 'CID Type 0C', true, false, true],
      ['TeXGyreBonum-Regular', 'CID Type 0C', true, false, true],
    ],
  );
  const words = popplerWords(file).words;
  assert.deepEqual(
    words.map((word) => word.text),
    ['Çığ', 'gölü', 'şöyle', 'İstanbul’da', 'ağaç'],
  );
  const width = measureWith(fontSet()).width('24px "TeX Gyre Bonum"', text);
  near([words[0].xMin], [20], 0.1, 'the first word starts where the text was painted');
  near([words[2].xMax], [20 + width], 0.1, 'the last word of the line ends where the text ends');
  assert.ok(words[0].yMin < 50 && words[0].yMax > 50, 'the baseline lies inside the box of the word');
  // The glyphs are painted: ink between the top of the capitals and the descenders, none outside.
  const pixels = popplerPixels(file);
  let ink = 0;
  let stray = 0;
  for (let y = 0; y < pixels.height; y++) {
    for (let x = 0; x < pixels.width; x++) {
      if (pixels.at(x, y)[0] > 128) continue;
      const first = y >= 26 && y <= 58 && x >= 18 && x <= 22 + width;
      const second = y >= 76 && y <= 95 && x >= 18 && x <= 160;
      if (first || second) ink++;
      else stray++;
    }
  }
  assert.ok(ink > 800, `the letters are painted (${ink} dark pixels)`);
  assert.equal(stray, 0, 'and nothing else');
});

// ---------------------------------------------------------------------------
// Pictures

function picture(alpha = false): PdfPicture {
  // Two by two: red, green; blue, white.
  const data = Uint8Array.from([255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 255]);
  return { width: 2, height: 2, image: { kind: 'samples', width: 2, height: 2, colors: 3, data, alpha: alpha ? Uint8Array.from([255, 255, 0, 255]) : undefined } };
}

popplerTest('pdf canvas: a picture is a picture of the page, where it was drawn, its first row at the top', async () => {
  const source = {};
  const made = make({ pictures: new Map([[source, picture()]]) });
  made.ctx.drawImage(source as CanvasImageSource, 20, 10, 120, 80);
  const file = await made.file();
  const page = checkPdf(file).pages[0];
  assert.equal(page.images.length, 1);
  near(page.images[0].matrix as number[], [120, 0, 0, 80, 20, 10], 0.001, 'the unit square of the picture on the page');
  const pixels = popplerPixels(file);
  const strongest = (pixel: number[]): string => (pixel[0] > 200 && pixel[1] > 200 && pixel[2] > 200 ? 'white' : pixel[0] > pixel[1] && pixel[0] > pixel[2] ? 'red' : pixel[1] > pixel[2] ? 'green' : 'blue');
  assert.equal(strongest(pixels.at(40, 22)), 'red', 'top left');
  assert.equal(strongest(pixels.at(120, 22)), 'green', 'top right');
  assert.equal(strongest(pixels.at(40, 78)), 'blue', 'bottom left');
  assert.equal(strongest(pixels.at(120, 78)), 'white', 'bottom right');
  near(pixels.at(10, 50), WHITE, 2, 'left of the picture');
  near(pixels.at(150, 50), WHITE, 2, 'right of the picture');
  assert.equal(poppler(file).complaints, '');
});

popplerTest('pdf canvas: a part of a picture is the picture behind a clip', async () => {
  const source = {};
  const made = make({ pictures: new Map([[source, picture()]]) });
  // The left column only: red over blue.
  made.ctx.drawImage(source as CanvasImageSource, 0, 0, 1, 2, 20, 10, 60, 80);
  const file = await made.file();
  const page = checkPdf(file).pages[0];
  near(page.images[0].matrix as number[], [120, 0, 0, 80, 20, 10], 0.001, 'the whole picture, twice as wide as the part');
  assert.deepEqual(page.clips, [[['re', 20, 10, 60, 80]]]);
  const pixels = popplerPixels(file);
  assert.ok(pixels.at(50, 25)[0] > 200 && pixels.at(50, 25)[1] < 80, 'red at the top');
  assert.ok(pixels.at(50, 75)[2] > 200 && pixels.at(50, 75)[0] < 80, 'blue at the bottom');
  near(pixels.at(100, 25), WHITE, 2, 'the other column is cut off');
});

test('pdf canvas: a picture that is drawn twice, on one page or on two, is in the file once', async () => {
  const source = {};
  const pictures = new Map<unknown, PdfPicture>([[source, picture(true)]]);
  const fonts = fontSet();
  const pages = [0, 1].map(() => {
    const made = createPdfCanvas({ width: 200, height: 100, pageWidth: 200, pageHeight: 100, fonts, measure: measureWith(fonts), images: { read: (from) => pictures.get(from) } });
    made.context.drawImage(source as CanvasImageSource, 0, 0);
    made.context.globalAlpha = 0.5;
    made.context.drawImage(source as CanvasImageSource, 50, 50, 20, 20);
    return { width: 200, height: 100, vector: made.canvas.finish() };
  });
  const checked = checkPdf(await buildPdf(pages, { date: DATE }));
  assert.equal(checked.pages[0].images.length, 2);
  near(checked.pages[0].images[0].matrix as number[], [2, 0, 0, 2, 0, 98], 0.001, 'at its own size where no size is given');
  assert.ok(checked.pages[0].images[0].mask, 'the picture brings its mask');
  let streams = 0;
  checked.objects.forEach((object) => {
    const value = object.value as { Subtype?: { name?: string }; SMask?: unknown; ColorSpace?: { name?: string } };
    if (object.stream && value.Subtype && value.Subtype.name === 'Image') streams++;
  });
  assert.equal(streams, 2, 'one picture and its mask in the file for four drawings of it');
  assert.match(checked.pages[0].content, /\/G\d+ gs/, 'the second drawing lets the paper through');
});

test('pdf canvas: a picture whose pixels cannot be read, or that is drawn from beyond its edges, makes the page unsupported', () => {
  const source = {};
  const none = make();
  none.ctx.drawImage(source as CanvasImageSource, 0, 0);
  assert.match(String(none.canvas.unsupported), /pixels cannot be read/);
  const unknown = make({ pictures: new Map() });
  unknown.ctx.drawImage(source as CanvasImageSource, 0, 0);
  assert.match(String(unknown.canvas.unsupported), /pixels cannot be read/);
  const beyond = make({ pictures: new Map([[source, picture()]]) });
  beyond.ctx.drawImage(source as CanvasImageSource, 1, 1, 5, 5, 0, 0, 10, 10);
  assert.match(String(beyond.canvas.unsupported), /beyond its edges/);
});

// ---------------------------------------------------------------------------
// What is not written

test('pdf canvas: a gradient, a pattern or a colour that is not read is a reason once something is painted with it', () => {
  const gradient = make();
  const paint = gradient.ctx.createLinearGradient(0, 0, 10, 0);
  paint.addColorStop(0, '#ffffff');
  gradient.ctx.fillStyle = paint;
  assert.equal(gradient.canvas.unsupported, undefined, 'choosing it paints nothing');
  gradient.ctx.fillStyle = '#000000';
  gradient.ctx.fillRect(0, 0, 10, 10);
  assert.equal(gradient.canvas.unsupported, undefined, 'and another paint was chosen before anything was painted');
  gradient.ctx.fillStyle = paint;
  gradient.ctx.fillRect(0, 0, 10, 10);
  assert.match(String(gradient.canvas.unsupported), /gradient/);
  const stroked = make();
  stroked.ctx.strokeStyle = stroked.ctx.createLinearGradient(0, 0, 1, 1);
  stroked.ctx.strokeRect(0, 0, 10, 10);
  assert.match(String(stroked.canvas.unsupported), /gradient/);
  const colour = make();
  colour.ctx.fillStyle = 'hsl(120, 50%, 50%)';
  colour.ctx.fillRect(0, 0, 10, 10);
  assert.match(String(colour.canvas.unsupported), /colour that is not read/);
  assert.deepEqual(parseCssColor('#abc'), { r: 170, g: 187, b: 204, a: 1 });
  assert.deepEqual(parseCssColor('#11223380'), { r: 17, g: 34, b: 51, a: 128 / 255 });
  assert.deepEqual(parseCssColor('rgba(26, 127, 69, 0.12)'), { r: 26, g: 127, b: 69, a: 0.12 });
  assert.deepEqual(parseCssColor(' RGB(1,2,3) '), { r: 1, g: 2, b: 3, a: 1 });
  assert.equal(parseCssColor('rebeccapurple'), undefined);
  assert.equal(parseCssColor('#12'), undefined);
  assert.equal(parseCssColor(undefined), undefined);
});

test('pdf canvas: a member of the canvas interface this canvas does not have makes the page unsupported, whether it is read, set or called', () => {
  const read = make();
  assert.equal((read.ctx as unknown as { canvas: unknown }).canvas, undefined);
  assert.match(String(read.canvas.unsupported), /"canvas" of the canvas/);
  const set = make();
  (set.ctx as unknown as { textRendering: string }).textRendering = 'geometricPrecision';
  assert.match(String(set.canvas.unsupported), /sets "textRendering"/);
  const called = make();
  // A painter may catch what a call throws: the reason is kept all the same.
  try {
    called.ctx.clearRect(0, 0, 10, 10);
  } catch {
    // Drawn another way, as far as the painter knows.
  }
  called.ctx.fillRect(0, 0, 10, 10);
  assert.match(String(called.canvas.unsupported), /"clearRect"/);
  assert.throws(() => called.canvas.finish(), (error: unknown) => error instanceof PdfError && /clearRect/.test(error.message));
  for (const member of ['ellipse', 'roundRect', 'arcTo', 'strokeText', 'getImageData', 'putImageData', 'isPointInPath', 'drawFocusIfNeeded']) {
    const made = make();
    assert.equal(typeof (made.ctx as unknown as { [name: string]: unknown })[member], 'undefined', `${member} is not there`);
    assert.match(String(made.canvas.unsupported), new RegExp(`"${member}"`), `asking for ${member} is a reason`);
  }
  // What painters ask a canvas to find out what it can do.
  const asked = make();
  assert.equal('letterSpacing' in asked.ctx, true);
  assert.equal('roundRect' in asked.ctx, false);
  assert.equal(typeof asked.ctx.setLineDash, 'function');
  assert.equal(typeof asked.ctx.createLinearGradient, 'function');
  // And the notes painters leave on a canvas for each other.
  (asked.ctx as unknown as { __bymTextBlock?: unknown }).__bymTextBlock = undefined;
  assert.equal((asked.ctx as unknown as { __bymTextBlock?: unknown }).__bymTextBlock, undefined);
  assert.equal(asked.canvas.unsupported, undefined, 'none of that is a reason');
  assert.ok(PDF_CANVAS_MEMBERS.indexOf('fillText') >= 0 && PDF_CANVAS_MEMBERS.indexOf('clearRect') < 0, 'the list of what is drawn');
});

test('pdf canvas: settings that change how everything is drawn are refused unless they are the plain ones', () => {
  const plain = make();
  const ctx = plain.ctx as unknown as { [name: string]: unknown };
  ctx.globalCompositeOperation = 'source-over';
  ctx.filter = 'none';
  ctx.shadowBlur = 0;
  ctx.shadowOffsetX = 0;
  ctx.shadowOffsetY = 0;
  ctx.shadowColor = '#000000';
  ctx.direction = 'ltr';
  ctx.fontKerning = 'auto';
  ctx.imageSmoothingEnabled = false;
  ctx.lineDashOffset = 3;
  ctx.miterLimit = 4;
  assert.equal(plain.canvas.unsupported, undefined);
  const refused: [string, unknown, RegExp][] = [
    ['globalCompositeOperation', 'multiply', /combined/],
    ['filter', 'blur(2px)', /filter/],
    ['shadowBlur', 4, /shadow/],
    ['shadowOffsetX', 1, /shadow/],
    ['direction', 'rtl', /right to left/],
    ['fontKerning', 'none', /kerning/],
    ['wordSpacing', '2px', /spacing of words/],
    ['letterSpacing', '0.1em', /spacing of letters/],
  ];
  for (const [member, value, why] of refused) {
    const made = make();
    (made.ctx as unknown as { [name: string]: unknown })[member] = value;
    assert.match(String(made.canvas.unsupported), why, member);
  }
  const squeezed = make();
  squeezed.ctx.font = '20px Test';
  squeezed.ctx.fillText('A', 0, 50, 5);
  assert.match(String(squeezed.canvas.unsupported), /squeezed/);
  const own = make();
  (own.ctx as unknown as { fill(path: object): void }).fill({});
  assert.match(String(own.canvas.unsupported), /path of its own/);
});

test('pdf canvas: a line on a canvas that is stretched one way more than the other is not written; a page whose pixels are not quite square is', async () => {
  const stretched = make();
  stretched.ctx.scale(2, 1);
  stretched.ctx.beginPath();
  stretched.ctx.moveTo(0, 0);
  stretched.ctx.lineTo(10, 10);
  stretched.ctx.stroke();
  assert.match(String(stretched.canvas.unsupported), /stretched/);
  // 1240 x 1754 pixels for a page of 595.28 x 841.89 points: the two axes differ by less than a thousandth.
  const page = make({ width: 1240, height: 1754, pageWidth: 595.28, pageHeight: 841.89 });
  page.ctx.lineWidth = 2;
  page.ctx.strokeRect(100, 100, 200, 200);
  assert.equal(page.canvas.unsupported, undefined);
  const shape = (await page.checked()).pages[0].shapes[0];
  near([shape.lineWidth], [(2 * 595.28) / 1240], 0.001, 'the width of the line in points');
  // A mirrored canvas stretches nothing.
  const mirrored = make();
  mirrored.ctx.translate(100, 0);
  mirrored.ctx.scale(-1, 1);
  mirrored.ctx.strokeRect(10, 10, 20, 20);
  assert.equal(mirrored.canvas.unsupported, undefined);
});

test('pdf canvas: once the page is unsupported nothing more is written and nothing throws', () => {
  const made = make();
  const ctx = made.ctx;
  ctx.fillStyle = 'hsl(0, 0%, 0%)';
  ctx.fillRect(0, 0, 10, 10);
  const first = made.canvas.unsupported;
  ctx.fillStyle = '#000000';
  ctx.font = '20px Arial';
  ctx.fillText('A', 0, 0);
  ctx.save();
  ctx.beginPath();
  ctx.arc(0, 0, 5, 0, 7);
  ctx.fill();
  ctx.stroke();
  ctx.clip();
  ctx.restore();
  ctx.restore();
  assert.equal(made.canvas.unsupported, first, 'the first reason stays');
});

test('pdf canvas: states that are saved and not given back are closed, and saving without end is refused', async () => {
  const open = make();
  open.ctx.save();
  open.ctx.save();
  open.ctx.fillRect(0, 0, 10, 10);
  // More restores than saves do nothing.
  open.ctx.restore();
  open.ctx.restore();
  open.ctx.restore();
  open.ctx.save();
  open.ctx.fillRect(20, 0, 10, 10);
  assert.equal((await open.checked()).pages[0].shapes.length, 2, 'the page is sound: every q has its Q');
  const deep = make();
  for (let i = 0; i < 100; i++) deep.ctx.save();
  assert.match(String(deep.canvas.unsupported), /saves the state/);
});

test('pdf canvas: numbers that are no numbers draw nothing, numbers beyond every page make the page unsupported', async () => {
  const made = make();
  const ctx = made.ctx;
  ctx.beginPath();
  ctx.moveTo(NaN, 0);
  ctx.lineTo(Infinity, 0);
  ctx.rect(0, 0, NaN, 5);
  ctx.arc(0, 0, NaN, 0, 1);
  ctx.bezierCurveTo(0, 0, 0, 0, 0, NaN);
  ctx.quadraticCurveTo(0, NaN, 0, 0);
  ctx.fill();
  ctx.stroke();
  ctx.translate(NaN, 0);
  ctx.scale(Infinity, 1);
  ctx.rotate(NaN);
  ctx.setTransform(NaN, 0, 0, 1, 0, 0);
  ctx.lineWidth = -1;
  ctx.lineWidth = NaN;
  ctx.globalAlpha = 7;
  ctx.setLineDash([1, -1]);
  ctx.fillRect(10, 10, 10, 10);
  assert.equal(made.canvas.unsupported, undefined);
  const page = (await made.checked()).pages[0];
  assert.deepEqual(page.shapes.map((shape) => shape.path), [[['re', 10, 80, 10, 10]]], 'only what has numbers is drawn, under the transform as it was');
  const far = make();
  far.ctx.fillRect(1e12, 0, 10, 10);
  assert.match(String(far.canvas.unsupported), /no page has/);
});

test('pdf document: drawn pages and pictured pages stand in one file, each font once', async () => {
  const fonts = fontSet();
  const drawn = (text: string) => {
    const made = createPdfCanvas({ width: 200, height: 100, pageWidth: 200, pageHeight: 100, fonts, measure: measureWith(fonts) });
    made.context.font = '20px Test';
    made.context.fillText(text, 10, 50);
    return { width: 200, height: 100, vector: made.canvas.finish() };
  };
  const pictured = { width: 200, height: 100, image: { kind: 'samples' as const, width: 2, height: 1, colors: 3 as const, data: Uint8Array.from([255, 0, 0, 0, 0, 255]) }, text: { runs: [{ text: 'hidden', x: 10, y: 50, width: 60, size: 20 }], width: 200 } };
  const file = await buildPdf([drawn('A'), pictured, drawn('BC')], { date: DATE });
  const checked = checkPdf(file);
  assert.deepEqual(
    checked.pages.map((page) => [page.text, page.images.length, page.texts.map((text) => text.mode).join()]),
    [
      ['A', 0, '0'],
      ['hidden ', 1, '3,3'],
      ['BC', 0, '0'],
    ],
  );
  assert.equal(checked.fonts.length, 2, 'the font of the drawn pages and the font of the hidden text');
  const subset = checked.fonts.filter((font) => font.baseFont !== 'BYMGlyphless')[0];
  assert.deepEqual(Array.from(subset.widths.keys()), [2, 3, 4], 'the glyphs of both drawn pages are in the one font');
  // The same pages give the same file, and other text another identifier.
  assert.deepEqual(Array.from(await buildPdf([drawn('A'), pictured, drawn('BC')], { date: DATE })), Array.from(file));
  const other = checkPdf(await buildPdf([drawn('A'), pictured, drawn('BD')], { date: DATE }));
  assert.notDeepEqual(other.trailer.ID, checked.trailer.ID);
  await assert.rejects(buildPdf([{ width: 10, height: 10, vector: { content: 1 } as never }]), (error: unknown) => error instanceof PdfError);
});
