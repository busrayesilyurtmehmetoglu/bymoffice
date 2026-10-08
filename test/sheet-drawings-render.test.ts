// The layer of pictures, shapes and comment boxes above the cells of a sheet:
// what it draws on a canvas, and that it only touches what is in sight.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { createSheetMetrics, type SheetMetrics } from '../src/layout/sheet-metrics';
import { createSheet, type DrawingAnchor, type DrawingObject, type DrawingShape, type Sheet, type SheetDrawing } from '../src/model/sheet/types';
import type { DecodedImage } from '../src/render/images';
import { SheetDrawings, type SheetDrawingsOptions } from '../src/render/sheet-drawings';
import { registerFont, unregisterFont } from '../src/render/fonts';
import { FONT_FALLBACK, clearTextMeasureCache } from '../src/render/text';

const PX = 9525;
const CW = 64;
const RH = 20;

interface Box {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** A stand-in for CanvasRenderingContext2D that records what is drawn. Each character is half the font size wide. */
class FakeContext {
  fillStyle = '#000000';
  strokeStyle = '#000000';
  lineWidth = 1;
  textAlign = 'left';
  textBaseline = 'alphabetic';
  font = '10px sans-serif';
  calls: string[] = [];
  images: { image: unknown; args: number[]; clips: Box[]; transforms: string[] }[] = [];
  rects: (Box & { fill: string; clips: Box[] })[] = [];
  strokeRects: (Box & { style: string })[] = [];
  texts: { text: string; x: number; y: number; font: string; fill: string; clips: Box[]; transforms: string[] }[] = [];
  fills: { fill: string; path: string[] }[] = [];
  strokes: { style: string; width: number; dash: number[]; path: string[] }[] = [];
  saves = 0;
  restores = 0;
  /** Makes drawImage() fail, as a canvas does for a picture it refuses. */
  failImages = false;

  private clips: Box[] = [];
  private transforms: string[] = [];
  private dash: number[] = [];
  private path: string[] = [];
  private pathRect: Box | undefined;
  private stack: { clips: Box[]; transforms: string[]; fill: string; font: string }[] = [];

  save(): void {
    this.saves++;
    this.stack.push({ clips: this.clips.slice(), transforms: this.transforms.slice(), fill: this.fillStyle, font: this.font });
  }
  restore(): void {
    this.restores++;
    const state = this.stack.pop();
    assert.ok(state, 'restore() without save()');
    this.clips = state.clips;
    this.transforms = state.transforms;
    this.fillStyle = state.fill;
    this.font = state.font;
  }
  translate(x: number, y: number): void {
    this.transforms.push(`translate(${x},${y})`);
  }
  rotate(angle: number): void {
    this.transforms.push(`rotate(${Math.round((angle * 180) / Math.PI)})`);
  }
  scale(x: number, y: number): void {
    this.transforms.push(`scale(${x},${y})`);
  }
  beginPath(): void {
    this.path = [];
    this.pathRect = undefined;
  }
  closePath(): void {
    this.path.push('close');
  }
  rect(x: number, y: number, width: number, height: number): void {
    this.pathRect = { x, y, width, height };
    this.path.push(`rect(${x},${y},${width},${height})`);
  }
  moveTo(x: number, y: number): void {
    this.path.push(`M${x},${y}`);
  }
  lineTo(x: number, y: number): void {
    this.path.push(`L${x},${y}`);
  }
  bezierCurveTo(): void {
    this.path.push('C');
  }
  quadraticCurveTo(): void {
    this.path.push('Q');
  }
  clip(): void {
    assert.ok(this.pathRect, 'clip() without a rect');
    this.clips = this.clips.concat([this.pathRect]);
  }
  fill(): void {
    this.fills.push({ fill: this.fillStyle, path: this.path.slice() });
  }
  stroke(): void {
    this.strokes.push({ style: this.strokeStyle, width: this.lineWidth, dash: this.dash.slice(), path: this.path.slice() });
  }
  setLineDash(dash: number[]): void {
    this.dash = dash.slice();
  }
  fillRect(x: number, y: number, width: number, height: number): void {
    this.rects.push({ x, y, width, height, fill: this.fillStyle, clips: this.clips });
  }
  strokeRect(x: number, y: number, width: number, height: number): void {
    this.strokeRects.push({ x, y, width, height, style: this.strokeStyle });
  }
  fillText(text: string, x: number, y: number): void {
    this.texts.push({ text, x, y, font: this.font, fill: this.fillStyle, clips: this.clips, transforms: this.transforms.slice() });
  }
  measureText(text: string): { width: number } {
    const match = /(\d+(?:\.\d+)?)px/.exec(this.font);
    return { width: text.length * (match ? Number(match[1]) : 10) * 0.5 };
  }
  drawImage(image: unknown, ...args: number[]): void {
    if (this.failImages) throw new Error('refused');
    this.images.push({ image, args, clips: this.clips, transforms: this.transforms.slice() });
  }
}

const newSheet = (): Sheet => createSheet({ name: 'S', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/x' });

const cells = (col: number, row: number, col2: number, row2: number): DrawingAnchor => ({
  kind: 'twoCell',
  from: { col, colOff: 0, row, rowOff: 0 },
  to: { col: col2, colOff: 0, row: row2, rowOff: 0 },
});

const picture = (media: string | undefined, extra: Partial<Extract<DrawingObject, { type: 'picture' }>> = {}): DrawingObject => ({ type: 'picture', media, ...extra });
const shape = (extra: Partial<DrawingShape> = {}): DrawingShape => ({ type: 'shape', geometry: 'rect', ...extra });

/** Pictures that are all there: 100 x 50 of their own pixels. */
function readyImages(): SheetDrawingsOptions['images'] & { asked: [string, number | undefined, number | undefined][] } {
  const asked: [string, number | undefined, number | undefined][] = [];
  return {
    asked,
    get(part, width, height) {
      asked.push([part, width, height]);
      return { source: part as unknown as CanvasImageSource, width: 100, height: 50 } as DecodedImage;
    },
    status: () => 'ready',
  };
}

interface PaintOptions extends SheetDrawingsOptions {
  clip?: Box;
  originX?: number;
  originY?: number;
  dpr?: number;
  metrics?: SheetMetrics;
  layer?: SheetDrawings;
  ctx?: FakeContext;
}

function paint(sheet: Sheet, options: PaintOptions = {}): FakeContext {
  const ctx = options.ctx || new FakeContext();
  const layer = options.layer || new SheetDrawings(sheet, options);
  layer.paint(ctx as unknown as CanvasRenderingContext2D, {
    metrics: options.metrics || createSheetMetrics(sheet),
    clip: options.clip || { x: 0, y: 0, width: 800, height: 600 },
    originX: options.originX || 0,
    originY: options.originY || 0,
    devicePixelRatio: options.dpr === undefined ? 1 : options.dpr,
  });
  assert.equal(ctx.saves, ctx.restores, 'save() and restore() are balanced');
  return ctx;
}

const withDrawings = (...drawings: SheetDrawing[]): Sheet => {
  const sheet = newSheet();
  sheet.drawings = drawings;
  return sheet;
};

test('layer: a sheet without drawings and comments has nothing for the layer', () => {
  const sheet = newSheet();
  assert.equal(SheetDrawings.isEmpty(sheet), true);
  const ctx = paint(sheet);
  assert.equal(ctx.saves, 0);
  sheet.drawings = [];
  sheet.comments = [];
  assert.equal(SheetDrawings.isEmpty(sheet), true);
  sheet.comments = [{ row: 0, col: 0, text: 'x' }];
  assert.equal(SheetDrawings.isEmpty(sheet), false);
});

test('layer: a picture is drawn into the rectangle of its anchor, under the clip of the pane', () => {
  const images = readyImages();
  const sheet = withDrawings({ anchor: cells(1, 2, 3, 6), object: picture('/xl/media/a.png') });
  const clip = { x: 72, y: 22, width: 400, height: 300 };
  const ctx = paint(sheet, { images, clip, originX: 72, originY: 22, dpr: 2 });
  assert.equal(ctx.images.length, 1);
  assert.equal(ctx.images[0].image, '/xl/media/a.png');
  // The whole picture (100 x 50 of its own pixels) into 128 x 80 at B3.
  assert.deepEqual(ctx.images[0].args, [0, 0, 100, 50, 72 + CW, 22 + 2 * RH, 2 * CW, 4 * RH]);
  assert.deepEqual(ctx.images[0].clips, [clip]);
  // Asked for at the size it covers on the canvas: CSS pixels times the device pixel ratio.
  assert.deepEqual(images.asked, [['/xl/media/a.png', 256, 160]]);
});

test('layer: a cropped picture shows its part, and is asked for with the pixels the whole of it needs', () => {
  const images = readyImages();
  const sheet = withDrawings({ anchor: cells(0, 0, 2, 4), object: picture('p', { crop: { left: 0.5, top: 0.2, right: 0.25, bottom: 0 } }) });
  const ctx = paint(sheet, { images });
  // A quarter of the width and four fifths of the height of the 100 x 50 picture, from (50, 10).
  assert.deepEqual(ctx.images[0].args, [50, 10, 25, 40, 0, 0, 128, 80]);
  assert.deepEqual(images.asked, [['p', 512, 100]]);
});

test('layer: rotation and flips are transforms around the centre of the box', () => {
  const images = readyImages();
  const sheet = withDrawings(
    { anchor: cells(0, 0, 2, 4), object: picture('p', { rotation: 30 }) },
    { anchor: cells(0, 0, 2, 4), object: picture('q', { flipH: true }) },
    { anchor: cells(0, 0, 2, 4), object: picture('r', { flipV: true, rotation: 180 }) },
  );
  const ctx = paint(sheet, { images });
  assert.deepEqual(ctx.images[0].transforms, ['translate(64,40)', 'rotate(30)', 'translate(-64,-40)']);
  assert.deepEqual(ctx.images[1].transforms, ['translate(64,40)', 'scale(-1,1)', 'translate(-64,-40)']);
  assert.deepEqual(ctx.images[2].transforms, ['translate(64,40)', 'rotate(180)', 'translate(-64,-40)', 'translate(64,40)', 'scale(1,-1)', 'translate(-64,-40)']);
});

test('layer: a picture that is on its way is a light rectangle; one that will not come is a placeholder with its words', () => {
  const sheet = withDrawings(
    { anchor: cells(0, 0, 2, 4), object: picture('waiting') },
    { anchor: cells(2, 0, 4, 4), object: picture('failed', { name: 'Picture 3', description: 'Company logo' }) },
    { anchor: cells(4, 0, 6, 4), object: picture(undefined, { external: true, name: 'Linked' }) },
    { anchor: cells(6, 0, 8, 4), object: picture(undefined) },
  );
  const ctx = paint(sheet, {
    images: { get: () => undefined, status: (part) => (part === 'failed' ? 'failed' : 'pending') },
    labels: { picture: 'Görsel' },
  });
  assert.equal(ctx.images.length, 0);
  assert.deepEqual(
    ctx.rects.map((rect) => [rect.x, rect.fill]),
    [
      [0, '#F6F7F8'],
      [128, '#F1F3F4'],
      [256, '#F1F3F4'],
      [384, '#F1F3F4'],
    ],
  );
  // The description, else the name, else the word of the view; nothing for the one that is still expected.
  assert.deepEqual(
    ctx.texts.map((call) => call.text),
    ['Company logo', 'Linked', 'Görsel'],
  );
  assert.equal(ctx.strokeRects.length, 3);
});

test('layer: without a source of pictures every picture is a placeholder; a source that throws is one too', () => {
  const sheet = withDrawings({ anchor: cells(0, 0, 2, 4), object: picture('p') });
  assert.deepEqual(
    paint(sheet).rects.map((rect) => rect.fill),
    ['#F1F3F4'],
  );
  const throwing = paint(sheet, {
    images: {
      get() {
        throw new Error('no');
      },
      status: () => 'failed',
    },
  });
  assert.deepEqual(
    throwing.rects.map((rect) => rect.fill),
    ['#F1F3F4'],
  );
});

test('layer: a placeholder label that does not fit is shortened; a small placeholder has none; a mirrored one is not mirrored', () => {
  const sheet = withDrawings(
    { anchor: cells(0, 0, 1, 2), object: picture(undefined, { description: 'A description that is much too long for its box' }) },
    { anchor: { kind: 'absolute', x: 200 * PX, y: 0, width: 20 * PX, height: 10 * PX }, object: picture(undefined, { name: 'x' }) },
    { anchor: cells(4, 0, 8, 4), object: { type: 'placeholder', what: 'chart', flipH: true } },
    { anchor: cells(4, 5, 8, 9), object: { type: 'placeholder', what: 'diagram' } },
    { anchor: cells(4, 10, 8, 14), object: { type: 'placeholder', what: 'object' } },
  );
  const ctx = paint(sheet, { labels: { chart: 'Grafik' } });
  // 12 px font: 6 px a character, 56 px of room.
  assert.deepEqual(
    ctx.texts.map((call) => call.text),
    ['A descri\u2026', 'Grafik', 'Diagram', 'Object'],
  );
  assert.deepEqual(ctx.texts[1].transforms, []);
  assert.equal(ctx.texts[1].font, `12px ${FONT_FALLBACK}`);
});

test('layer: fill and outline of a shape; colours that are not "#RRGGBB" are never used', () => {
  const sheet = withDrawings(
    { anchor: cells(0, 0, 2, 4), object: shape({ fill: '#00B050', line: { color: '#112233', width: 2 } }) },
    { anchor: cells(2, 0, 4, 4), object: shape({ geometry: 'ellipse', fill: '#FF0000' }) },
    { anchor: cells(4, 0, 6, 4), object: shape({ geometry: 'roundRect', line: { color: '#0000FF', width: 1, dash: true } }) },
    { anchor: cells(6, 0, 8, 4), object: shape({ fill: 'red; evil', line: { color: 'url(x)', width: 1 } }) },
    { anchor: cells(8, 0, 10, 4), object: shape({ geometry: 'cloud', fill: '#ABCDEF' }) },
  );
  const ctx = paint(sheet);
  assert.deepEqual(
    ctx.fills.map((call) => [call.fill, call.path[0]]),
    [
      ['#00B050', 'rect(0,0,128,80)'],
      ['#FF0000', 'M128,40'],
      // A geometry that is not known is drawn as its box.
      ['#ABCDEF', 'rect(512,0,128,80)'],
    ],
  );
  assert.equal(ctx.fills[1].path.filter((step) => step === 'C').length, 4);
  assert.deepEqual(
    ctx.strokes.map((call) => [call.style, call.width, call.dash, call.path[0]]),
    [
      ['#112233', 2, [], 'rect(0,0,128,80)'],
      // A hairline is put on the pixel centres; the corner is a sixth of the shorter side.
      ['#0000FF', 1, [4, 3], `M${256.5 + 79 / 6},0.5`],
    ],
  );
});

test('layer: a connector is a line from corner to corner; flat lines are drawn too', () => {
  const sheet = withDrawings(
    { anchor: cells(0, 0, 2, 4), object: shape({ connector: true, line: { color: '#0000FF', width: 2 } }) },
    { anchor: cells(2, 0, 4, 4), object: shape({ connector: true, flipV: true, line: { color: '#0000FF', width: 2 } }) },
    { anchor: cells(0, 5, 2, 5), object: shape({ connector: true, line: { color: '#FF0000', width: 1 } }) },
    { anchor: cells(3, 5, 3, 8), object: shape({ connector: true, line: { color: '#FF0000', width: 3 } }) },
    // Without a colour there is nothing to draw; a point is nothing either.
    { anchor: cells(0, 9, 2, 12), object: shape({ connector: true }) },
    { anchor: cells(5, 5, 5, 5), object: shape({ connector: true, line: { color: '#FF0000', width: 1 } }) },
  );
  const ctx = paint(sheet);
  assert.deepEqual(
    ctx.strokes.map((call) => call.path),
    [
      ['M0,0', 'L128,80'],
      ['M128,0', 'L256,80'],
      ['M0,100.5', 'L128,100.5'],
      ['M192,100', 'L192,160'],
    ],
  );
  assert.equal(ctx.fills.length, 0);
});

test('layer: the text of a shape is broken at its width, aligned, anchored and cut at its box', () => {
  clearTextMeasureCache();
  // 16 px font: 8 px a character. The box is 128 wide with 4 px insets: 15 characters a line.
  const text = {
    insets: [4, 2, 4, 2] as [number, number, number, number],
    paragraphs: [{ runs: [{ text: 'one two three four five', size: 12, color: '#FF0000' }] }, { runs: [{ text: 'right', size: 12 }], align: 'right' as const }, { runs: [{ text: 'mid', size: 12, bold: true, font: 'Arial' }], align: 'center' as const }],
  };
  const sheet = withDrawings({ anchor: cells(1, 1, 3, 9), object: shape({ fill: '#FFFFFF', text }) });
  const ctx = paint(sheet);
  const font = `16px ${FONT_FALLBACK}`;
  assert.deepEqual(
    ctx.texts.map((call) => [call.text, call.x, call.y, call.fill]),
    [
      ['one', 68, 37, '#FF0000'],
      ['two', 100, 37, '#FF0000'],
      ['three', 132, 37, '#FF0000'],
      ['four', 68, 57, '#FF0000'],
      ['five', 108, 57, '#FF0000'],
      // Right: 5 characters of 8 px end at 64 + 128 - 4.
      ['right', 148, 76, '#000000'],
      // Centre: 3 characters in 120 px of room.
      ['mid', 116, 95, '#000000'],
    ],
  );
  assert.equal(ctx.texts[0].font, font);
  assert.equal(ctx.texts[6].font, `bold 16px "Arial", ${FONT_FALLBACK}`);
  assert.deepEqual(ctx.texts[0].clips, [{ x: 0, y: 0, width: 800, height: 600 }, { x: 64, y: 20, width: 128, height: 160 }]);

  // Anchored to the bottom: the same four lines (76.8 px) end 2 px above the bottom edge.
  const bottom = paint(withDrawings({ anchor: cells(1, 1, 3, 9), object: shape({ text: { ...text, anchor: 'bottom' } }) }));
  assert.equal(bottom.texts[0].y, Math.round(22 + (156 - 76.8) + 19.2 * 0.8));
  const middle = paint(withDrawings({ anchor: cells(1, 1, 3, 9), object: shape({ text: { ...text, anchor: 'middle' } }) }));
  assert.equal(middle.texts[0].y, Math.round(22 + (156 - 76.8) / 2 + 19.2 * 0.8));
  // Taller than its box: from the top, whatever it asks for.
  const tall = paint(withDrawings({ anchor: cells(1, 1, 3, 3), object: shape({ text: { ...text, anchor: 'bottom' } }) }));
  assert.equal(tall.texts[0].y, 37);
  // Lines below the box are not drawn at all.
  assert.deepEqual(
    tall.texts.map((call) => call.text),
    ['one', 'two', 'three', 'four', 'five'],
  );
});

test('layer: line breaks, text that is not wrapped, a word wider than the box, underline and strike', () => {
  clearTextMeasureCache();
  const runs = [{ text: 'ab\ncd\r\nef', size: 12 }, { text: ' gh', size: 12, underline: true, strike: true }];
  const none = paint(withDrawings({ anchor: cells(0, 0, 1, 6), object: shape({ text: { wrap: false, paragraphs: [{ runs }] } }) }));
  assert.deepEqual(
    none.texts.map((call) => [call.text, call.x, call.y]),
    [
      ['ab', 0, 15],
      ['cd', 0, 35],
      ['ef', 0, 54],
      ['gh', 24, 54],
    ],
  );
  // Underline and strike: two thin rectangles under and through " gh" and "gh" (the space carries them too).
  assert.equal(none.rects.length, 4);
  // 64 px of room: eight characters a line.
  const long = paint(withDrawings({ anchor: cells(0, 0, 1, 6), object: shape({ text: { paragraphs: [{ runs: [{ text: 'abcdefghijklmnopqrst uv', size: 12 }] }] } }) }));
  assert.deepEqual(
    long.texts.map((call) => call.text),
    ['abcdefgh', 'ijklmnop', 'qrst', 'uv'],
  );
});

test('layer: text is turned with its shape but never mirrored', () => {
  const sheet = withDrawings({ anchor: cells(0, 0, 2, 4), object: shape({ rotation: 30, flipH: true, fill: '#FFFFFF', text: { paragraphs: [{ runs: [{ text: 'x' }] }] } }) });
  const ctx = paint(sheet);
  assert.deepEqual(ctx.texts[0].transforms, ['translate(64,40)', 'rotate(30)', 'translate(-64,-40)']);
});

test('layer: font names and colours of the text are never trusted', () => {
  const sheet = withDrawings({
    anchor: cells(0, 0, 4, 4),
    object: shape({ text: { paragraphs: [{ runs: [{ text: 'x', font: 'Arial"; } body { color: red', color: 'expression(alert(1))', size: Number.NaN }] }] } }),
  });
  const ctx = paint(sheet);
  assert.equal(ctx.texts[0].fill, '#000000');
  assert.equal(ctx.texts[0].font, `14.67px "Arial body color red", ${FONT_FALLBACK}`);
});

test('layer: the members of a group are placed in the box of the group', () => {
  const images = readyImages();
  const group: DrawingObject = {
    type: 'group',
    children: [
      { ...shape({ fill: '#FF0000' }), box: { x: 0, y: 0, width: 0.5, height: 0.5 } },
      { ...picture('p'), box: { x: 0.5, y: 0.5, width: 0.5, height: 0.25 } },
      { type: 'group', box: { x: 0, y: 0.5, width: 0.5, height: 0.5 }, children: [{ ...shape({ fill: '#0000FF' }), box: { x: 0.5, y: 0, width: 0.5, height: 1 } }] },
      // A member without a place is not drawn.
      shape({ fill: '#00FF00' }),
    ],
  };
  const ctx = paint(withDrawings({ anchor: cells(1, 1, 3, 5), object: group }), { images });
  assert.deepEqual(
    ctx.fills.map((call) => [call.fill, call.path[0]]),
    [
      ['#FF0000', 'rect(64,20,64,40)'],
      ['#0000FF', 'rect(96,60,32,40)'],
    ],
  );
  assert.deepEqual(ctx.images[0].args.slice(4), [128, 60, 64, 20]);
});

test('layer: a group nested without end, or with members without number, is drawn as far as the limits go', () => {
  let deep: DrawingObject = { ...shape({ fill: '#FF0000' }), box: { x: 0, y: 0, width: 1, height: 1 } };
  for (let i = 0; i < 100; i++) deep = { type: 'group', children: [deep], box: { x: 0, y: 0, width: 1, height: 1 } };
  assert.equal(paint(withDrawings({ anchor: cells(0, 0, 2, 4), object: deep })).fills.length, 0);
  // A group that holds itself.
  const loop: DrawingObject = { type: 'group', children: [], box: { x: 0, y: 0, width: 1, height: 1 } };
  loop.children.push(loop);
  paint(withDrawings({ anchor: cells(0, 0, 2, 4), object: loop }));
  const wide: DrawingObject = { type: 'group', children: [] };
  for (let i = 0; i < 30000; i++) wide.children.push({ ...shape({ fill: '#FF0000' }), box: { x: 0, y: 0, width: 1, height: 1 } });
  assert.ok(paint(withDrawings({ anchor: cells(0, 0, 2, 4), object: wide })).fills.length < 20000);
});

test('layer: what the model states is not trusted: sizes that are not numbers, objects that are not objects', () => {
  const sheet = withDrawings(
    { anchor: { kind: 'absolute', x: Number.NaN, y: 0, width: 100 * PX, height: 100 * PX }, object: shape({ fill: '#FF0000' }) },
    { anchor: { kind: 'absolute', x: 0, y: 0, width: Number.POSITIVE_INFINITY, height: 100 * PX }, object: shape({ fill: '#FF0000' }) },
    { anchor: { kind: 'absolute', x: 0, y: 0, width: -5, height: 100 * PX }, object: shape({ fill: '#FF0000' }) },
    { anchor: { kind: 'twoCell' }, object: shape({ fill: '#FF0000' }) },
    { anchor: cells(0, 0, 2, 4), object: undefined as unknown as DrawingObject },
    undefined as unknown as SheetDrawing,
    { anchor: cells(0, 0, 2, 4), object: { ...shape({ fill: '#00FF00' }), rotation: Number.NaN, text: { paragraphs: [{ runs: [{ text: 5 as unknown as string }] }], insets: [Number.NaN, -1, 1e9, 2] } } },
    { anchor: cells(0, 0, 2, 4), object: { type: 'group', children: undefined as unknown as DrawingObject[] } },
    { anchor: cells(0, 0, 2, 4), object: { type: 'shape', geometry: 'rect', hidden: true, fill: '#0000FF' } },
  );
  const ctx = paint(sheet);
  // A position that is not a number reads as 0; a size that is not one leaves nothing to draw.
  assert.deepEqual(
    ctx.fills.map((call) => [call.fill, call.path[0]]),
    [
      ['#FF0000', 'rect(0,0,100,100)'],
      ['#00FF00', 'rect(0,0,128,80)'],
    ],
  );
  assert.equal(ctx.texts.length, 0);
});

test('layer: an object the canvas refuses does not take the others with it, and leaves the context in order', () => {
  const images = readyImages();
  const sheet = withDrawings({ anchor: cells(0, 0, 2, 4), object: picture('p', { rotation: 20, flipH: true }) }, { anchor: cells(2, 0, 4, 4), object: shape({ fill: '#00FF00' }) });
  const ctx = new FakeContext();
  ctx.failImages = true;
  paint(sheet, { images, ctx });
  assert.deepEqual(
    ctx.fills.map((call) => call.fill),
    ['#00FF00'],
  );
});

test('layer: only what reaches into the clip is drawn, in the order of the file', () => {
  const images = readyImages();
  const sheet = newSheet();
  sheet.drawings = [];
  // 10,000 pictures, one per row, in a column; every 500th one is tall and reaches far down.
  for (let i = 0; i < 10000; i++) sheet.drawings.push({ anchor: cells(1, i, 2, i % 500 === 0 ? i + 400 : i + 1), object: picture(`p${i}`) });
  const layer = new SheetDrawings(sheet, { images });
  const metrics = createSheetMetrics(sheet);
  // The window shows rows 5000..5029 (600 px), scrolled there.
  const ctx = paint(sheet, { layer, metrics, originY: -5000 * RH });
  const drawn = ctx.images.map((call) => call.image);
  // The thirty rows in sight and the two that touch the edges of the window, the earlier one first.
  assert.equal(drawn.length, 32);
  assert.equal(drawn[0], 'p4999');
  assert.equal(drawn[31], 'p5030');
  // A row further up: the tall picture of row 4500 still reaches in (to row 4900).
  const up = paint(sheet, { layer, metrics, originY: -4880 * RH });
  assert.deepEqual(
    up.images.map((call) => call.image).slice(0, 3),
    ['p4500', 'p4879', 'p4880'],
  );
  // Nothing to the right of column B.
  assert.equal(paint(sheet, { layer, metrics, originX: -3 * CW }).images.length, 0);

  // A frame costs what is in sight, not what the sheet holds.
  images.asked.length = 0;
  const started = Date.now();
  for (let i = 0; i < 200; i++) paint(sheet, { layer, metrics, originY: -(i * 37) * RH });
  const elapsed = Date.now() - started;
  assert.ok(images.asked.length <= 200 * 34, `${images.asked.length} pictures asked for in 200 frames`);
  assert.ok(elapsed < 20000, `200 frames took ${elapsed} ms`);
});

test('layer: later objects are drawn over earlier ones, whatever their place on the sheet', () => {
  const sheet = withDrawings(
    { anchor: cells(0, 5, 2, 9), object: shape({ fill: '#000001' }) },
    { anchor: cells(0, 0, 2, 9), object: shape({ fill: '#000002' }) },
    { anchor: cells(0, 7, 2, 9), object: shape({ fill: '#000003' }) },
    { anchor: cells(0, 1, 2, 9), object: shape({ fill: '#000004' }) },
  );
  assert.deepEqual(
    paint(sheet).fills.map((call) => call.fill),
    ['#000001', '#000002', '#000003', '#000004'],
  );
});

test('layer: a rotated object is found by the circle around its box', () => {
  // 20 wide and 400 high, turned by 80 degrees: it reaches 200 px to each side of its centre.
  const sheet = withDrawings({ anchor: { kind: 'absolute', x: 300 * PX, y: 300 * PX, width: 20 * PX, height: 400 * PX }, object: shape({ fill: '#FF0000', rotation: 20 }) });
  assert.equal(paint(sheet, { clip: { x: 0, y: 0, width: 150, height: 600 } }).fills.length, 1);
  assert.equal(paint(sheet, { clip: { x: 0, y: 0, width: 80, height: 600 } }).fills.length, 0);
});

test('layer: places follow the sizes of rows and columns, and structural edits', () => {
  const sheet = withDrawings({ anchor: cells(1, 1, 2, 2), object: shape({ fill: '#FF0000' }) });
  const layer = new SheetDrawings(sheet);
  assert.equal(paint(sheet, { layer }).fills[0].path[0], 'rect(64,20,64,20)');
  // A taller first row: new metrics, new place.
  sheet.rows.set(0, { height: 30, customHeight: true, hidden: false, outlineLevel: 0, collapsed: false, cells: [] });
  assert.equal(paint(sheet, { layer }).fills[0].path[0], 'rect(64,40,64,20)');
  sheet.structuralEdits = [{ kind: 'insert', axis: 'col', at: 0, count: 2 }];
  assert.equal(paint(sheet, { layer }).fills[0].path[0], 'rect(192,40,64,20)');
  sheet.structuralEdits = undefined;
  assert.equal(paint(sheet, { layer }).fills[0].path[0], 'rect(64,40,64,20)');
  assert.equal(layer.count, 1);
});

test('layer: a comment that is always shown is a box with its text; the others are only looked up', () => {
  clearTextMeasureCache();
  const sheet = newSheet();
  sheet.comments = [
    { row: 1, col: 1, text: 'Ayşe:\nnot', author: 'Ayşe', visible: true, runs: [{ text: 'Ayşe:', font: { bold: true, italic: false, strike: false, underline: 'none', size: 9 } }, { text: '\nnot' }] },
    { row: 3, col: 0, text: 'kutulu', visible: true, box: cells(2, 6, 4, 9) },
    { row: 5, col: 5, text: 'gizli' },
    { row: 7, col: 0, text: 'konuşma', author: 'Bora', threaded: true, visible: true, box: cells(6, 0, 9, 9), replies: [{ author: 'Can', text: 'yanıt' }, { text: 'adsız' }] },
  ];
  const layer = new SheetDrawings(sheet);
  const ctx = paint(sheet, { layer });
  assert.equal(layer.count, 3);
  assert.equal(layer.has(5, 5), true);
  assert.equal(layer.has(5, 4), false);
  assert.equal(layer.commentAt(3, 0)?.text, 'kutulu');
  // Without a box of its own: beside the cell, a little above it, in the size Excel gives a new note.
  assert.deepEqual(
    ctx.fills.map((call) => [call.fill, call.path[0]]),
    [
      ['#FFFFE1', 'rect(139,13,144,79)'],
      ['#FFFFE1', 'rect(128,120,128,60)'],
      ['#FFFFE1', 'rect(384,0,192,180)'],
    ],
  );
  assert.deepEqual(
    ctx.texts.map((call) => [call.text, /^bold /.test(call.font)]),
    [
      ['Ayşe:', true],
      ['not', false],
      ['kutulu', false],
      ['Bora:', true],
      ['konuşma', false],
      ['Can:', true],
      ['yanıt', false],
      ['adsız', false],
    ],
  );
});

test('layer: the text of a shape follows the fonts that are registered, on the same layer', async () => {
  // The browser's web fonts, as far as the registry uses them.
  const globals = globalThis as { FontFace?: unknown; document?: unknown };
  const before = { face: globals.FontFace, document: globals.document };
  globals.FontFace = class {
    load(): Promise<unknown> {
      return Promise.resolve(this);
    }
  };
  globals.document = { fonts: { add: () => undefined, delete: () => undefined } };
  try {
    const text = { paragraphs: [{ runs: [{ text: 'one two', size: 12, font: 'Bookman Old Style Tur' }] }] };
    const sheet = withDrawings({ anchor: cells(1, 1, 3, 9), object: shape({ fill: '#FFFFFF', text }) });
    const layer = new SheetDrawings(sheet);
    const fonts = () => Array.from(new Set(paint(sheet, { layer }).texts.map((call) => call.font)));
    const plain = `16px "Bookman Old Style Tur", "Bookman Old Style", ${FONT_FALLBACK}`;
    assert.deepEqual(fonts(), [plain]);
    await registerFont({ family: 'My Bonum', source: new Uint8Array([0, 1, 0, 0]), as: 'Bookman Old Style' });
    assert.deepEqual(fonts(), [`16px "Bookman Old Style Tur", "Bookman Old Style", "My Bonum", ${FONT_FALLBACK}`]);
    assert.equal(unregisterFont('My Bonum'), true);
    assert.deepEqual(fonts(), [plain]);
  } finally {
    unregisterFont('My Bonum');
    globals.FontFace = before.face;
    globals.document = before.document;
  }
});
