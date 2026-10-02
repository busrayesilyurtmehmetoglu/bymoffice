import assert from 'node:assert/strict';
import { test } from 'node:test';
import type { ImageItem, LayoutFont, LayoutPage, LineItem, PageItem, TextItem } from '../src/layout/doc/types';
import { renderPage, type PageImage, type RenderPageOptions } from '../src/render/doc';
import { CanvasTextMeasurer, docCssFont, type MeasureContext } from '../src/render/doc-measure';
import { declaredImagePixels, declaredImageSize, decodedImageSize, documentText, sniffImageType } from '../src/ui/doc-view';
import { STRINGS_EN, STRINGS_TR, getUiStrings } from '../src/ui/strings';
import { p, r, tbl, tc, tr } from './doc-layout-fixtures';
import { readSpec } from './docx-fixtures';

// A context that records what is drawn, in canvas pixels: every call is kept
// with the transform that was in effect, so a test sees where things land.

type Matrix = [number, number, number, number, number, number];

interface RectCall {
  x: number;
  y: number;
  width: number;
  height: number;
  fill: string;
}
interface TextCall {
  text: string;
  x: number;
  y: number;
  font: string;
  fill: string;
  spacing: string | undefined;
  /** Horizontal and vertical scale of the transform. */
  sx: number;
  sy: number;
  align: string;
}
interface StrokeCall {
  points: [string, ...number[]][];
  color: string;
  width: number;
  dash: number[];
}
interface ImageCall {
  source: unknown;
  args: number[];
  matrix: Matrix;
}

interface Recorder {
  ctx: CanvasRenderingContext2D;
  rects: RectCall[];
  texts: TextCall[];
  strokes: StrokeCall[];
  images: ImageCall[];
  measured: string[];
  depth(): number;
}

const round = (value: number) => Math.round(value * 1000) / 1000;

function recorder(features: { letterSpacing?: boolean; dash?: boolean; failImages?: boolean } = {}): Recorder {
  const rects: RectCall[] = [];
  const texts: TextCall[] = [];
  const strokes: StrokeCall[] = [];
  const images: ImageCall[] = [];
  const measured: string[] = [];
  interface State {
    m: Matrix;
    fillStyle: string;
    strokeStyle: string;
    lineWidth: number;
    font: string;
    dash: number[];
    textAlign: string;
    textBaseline: string;
    letterSpacing: string;
  }
  let state: State = { m: [1, 0, 0, 1, 0, 0], fillStyle: '#000000', strokeStyle: '#000000', lineWidth: 1, font: '10px sans-serif', dash: [], textAlign: 'start', textBaseline: 'alphabetic', letterSpacing: '0px' };
  const stack: State[] = [];
  let path: [string, ...number[]][] = [];
  const apply = (x: number, y: number): [number, number] => [round(state.m[0] * x + state.m[2] * y + state.m[4]), round(state.m[1] * x + state.m[3] * y + state.m[5])];
  const multiply = (n: Matrix) => {
    const m = state.m;
    state.m = [m[0] * n[0] + m[2] * n[1], m[1] * n[0] + m[3] * n[1], m[0] * n[2] + m[2] * n[3], m[1] * n[2] + m[3] * n[3], m[0] * n[4] + m[2] * n[5] + m[4], m[1] * n[4] + m[3] * n[5] + m[5]];
  };
  const size = (font: string) => {
    const match = /(\d+(?:\.\d+)?)px/.exec(font);
    return match ? Number(match[1]) : 10;
  };
  const ctx: { [name: string]: unknown } = {
    save: () => stack.push({ ...state, m: state.m.slice() as Matrix, dash: state.dash.slice() }),
    restore: () => {
      const top = stack.pop();
      if (top) state = top;
    },
    scale: (x: number, y: number) => multiply([x, 0, 0, y, 0, 0]),
    translate: (x: number, y: number) => multiply([1, 0, 0, 1, x, y]),
    rotate: (angle: number) => multiply([Math.cos(angle), Math.sin(angle), -Math.sin(angle), Math.cos(angle), 0, 0]),
    setTransform: (a: number, b: number, c: number, d: number, e: number, f: number) => {
      state.m = [a, b, c, d, e, f];
    },
    fillRect: (x: number, y: number, width: number, height: number) => {
      const [x1, y1] = apply(x, y);
      const [x2, y2] = apply(x + width, y + height);
      rects.push({ x: Math.min(x1, x2), y: Math.min(y1, y2), width: round(Math.abs(x2 - x1)), height: round(Math.abs(y2 - y1)), fill: state.fillStyle });
    },
    fillText: (text: string, x: number, y: number) => {
      const [dx, dy] = apply(x, y);
      texts.push({ text, x: dx, y: dy, font: state.font, fill: state.fillStyle, spacing: features.letterSpacing ? state.letterSpacing : undefined, sx: round(state.m[0]), sy: round(state.m[3]), align: state.textAlign });
    },
    // Every character is half the font size wide.
    measureText: (text: string) => {
      measured.push(text);
      return { width: text.length * size(state.font) * 0.5 };
    },
    beginPath: () => {
      path = [];
    },
    moveTo: (x: number, y: number) => path.push(['M', ...apply(x, y)]),
    lineTo: (x: number, y: number) => path.push(['L', ...apply(x, y)]),
    quadraticCurveTo: (cx: number, cy: number, x: number, y: number) => path.push(['Q', ...apply(cx, cy), ...apply(x, y)]),
    stroke: () => strokes.push({ points: path.slice(), color: state.strokeStyle, width: round(state.lineWidth * state.m[0]), dash: state.dash.map((value) => round(value * state.m[0])) }),
    drawImage: (source: unknown, ...args: number[]) => {
      if (features.failImages) throw new Error('InvalidStateError');
      images.push({ source, args, matrix: state.m.map(round) as Matrix });
    },
  };
  if (features.dash !== false) {
    ctx.setLineDash = (dash: number[]) => {
      state.dash = dash.slice();
    };
  }
  for (const name of ['fillStyle', 'strokeStyle', 'lineWidth', 'font', 'textAlign', 'textBaseline'] as const) {
    Object.defineProperty(ctx, name, { get: () => state[name], set: (value) => ((state as unknown as { [key: string]: unknown })[name] = value), enumerable: true });
  }
  if (features.letterSpacing) {
    Object.defineProperty(ctx, 'letterSpacing', { get: () => state.letterSpacing, set: (value: string) => (state.letterSpacing = value), enumerable: true });
  }
  return { ctx: ctx as unknown as CanvasRenderingContext2D, rects, texts, strokes, images, measured, depth: () => stack.length };
}

const FONT: LayoutFont = { family: 'Test', size: 20, bold: false, italic: false };
const page = (items: PageItem[], width = 400, height = 300): LayoutPage => ({ width, height, section: 0, number: 1, items });
const text = (extra: Partial<TextItem> = {}): TextItem => ({ type: 'text', x: 40, y: 100, width: 50, text: 'Hello', font: FONT, color: '#112233', ...extra });
const line = (extra: Partial<LineItem> = {}): LineItem => ({ type: 'line', x1: 10, y1: 50, x2: 110, y2: 50, color: '#ff0000', width: 1, style: 'single', ...extra });
const image = (extra: Partial<ImageItem> = {}): ImageItem => ({ type: 'image', x: 20, y: 30, width: 100, height: 50, part: '/word/media/image1.png', ...extra });

/** Draws without the paper, so that only what the items draw is recorded. */
function draw(items: PageItem[], options: RenderPageOptions = {}, features?: Parameters<typeof recorder>[0]): Recorder {
  const rec = recorder(features);
  renderPage(rec.ctx, page(items), { background: null, ...options });
  assert.equal(rec.depth(), 0, 'every save has its restore');
  return rec;
}

test('renderPage: the paper is white, covers the page and scales with it', () => {
  const rec = recorder();
  renderPage(rec.ctx, page([{ type: 'rect', x: 10, y: 20, width: 30, height: 40, fill: '#abcdef' }]), { scale: 2 });
  assert.deepEqual(rec.rects, [
    { x: 0, y: 0, width: 800, height: 600, fill: '#ffffff' },
    { x: 20, y: 40, width: 60, height: 80, fill: '#abcdef' },
  ]);
  const custom = recorder();
  renderPage(custom.ctx, page([]), { background: '#fafafa' });
  assert.equal(custom.rects[0].fill, '#fafafa');
  assert.deepEqual(draw([]).rects, []);
});

test('renderPage: rectangles are fitted to whole canvas pixels; what is not a rectangle is skipped', () => {
  const rec = draw(
    [
      { type: 'rect', x: 10.3, y: 20.3, width: 9.9, height: 10, fill: '#00ff00' },
      { type: 'rect', x: 20.2, y: 20.3, width: 9.9, height: 10, fill: 'red; }' },
      { type: 'rect', x: NaN, y: 0, width: 10, height: 10, fill: '#000000' },
      { type: 'rect', x: 0, y: 0, width: -5, height: 10, fill: '#000000' },
      { type: 'rect', x: 0, y: 0, width: Infinity, height: 10, fill: '#000000' },
    ],
    { scale: 1.5 },
  );
  // 15.45..30.3 and 30.3..45.15 in canvas pixels: both meet at 30, without a seam or an overlap.
  assert.deepEqual(rec.rects, [
    { x: 15, y: 30, width: 15, height: 15, fill: '#00ff00' },
    // A colour that is not "#RRGGBB" is never handed to the canvas.
    { x: 30, y: 30, width: 15, height: 15, fill: '#ffffff' },
  ]);
});

test('renderPage: text is drawn on its baseline with the font it was measured with', () => {
  const rec = draw([text(), text({ x: 100, text: 'Bold', font: { family: 'Calibri', size: 12, bold: true, italic: true }, color: 'javascript:1' })], { scale: 2 });
  assert.deepEqual(rec.texts, [
    { text: 'Hello', x: 80, y: 200, font: '20px "Test", sans-serif', fill: '#112233', spacing: undefined, sx: 2, sy: 2, align: 'left' },
    { text: 'Bold', x: 200, y: 200, font: 'italic bold 12px "Calibri", "Carlito", sans-serif', fill: '#000000', spacing: undefined, sx: 2, sy: 2, align: 'left' },
  ]);
  const custom = draw([text()], { cssFont: (font) => `${font.size}px mine` });
  assert.equal(custom.texts[0].font, '20px mine');
});

test('renderPage: blank text and broken items draw nothing and do not stop the page', () => {
  const rec = draw([
    text({ text: '   ' }),
    text({ text: '' }),
    text({ x: NaN }),
    text({ y: Infinity }),
    { type: 'text', x: 0, y: 0, width: 10, text: 'x', color: '#000000' } as unknown as TextItem,
    { type: 'mystery' } as unknown as PageItem,
    text({ text: 'after' }),
  ]);
  assert.deepEqual(rec.texts.map((call) => call.text), ['after']);
  assert.deepEqual(rec.rects, []);
});

test('renderPage: a highlight lies behind the text, as high as the line of its font', () => {
  const rec = draw([text({ highlight: '#ffff00' })], { metrics: () => ({ ascent: 16, descent: 4 }) });
  assert.deepEqual(rec.rects, [{ x: 40, y: 84, width: 50, height: 20, fill: '#ffff00' }]);
  // Without metrics the height is estimated from the size.
  const estimated = draw([text({ highlight: '#ffff00' })]);
  assert.deepEqual(estimated.rects, [{ x: 40, y: 82, width: 50, height: 23, fill: '#ffff00' }]);
  assert.equal(rec.texts.length, 1);
});

test('renderPage: underlines are sharp lines below the baseline, in their own colour', () => {
  const single = draw([text({ underline: { color: '#0000ff', style: 'single' } })]);
  // 20 px text: 1.2 px thick, drawn as one whole pixel; 2.4 px below the baseline.
  assert.deepEqual(single.rects, [{ x: 40, y: 103, width: 50, height: 1, fill: '#0000ff' }]);
  assert.equal(single.strokes.length, 0);

  const scaled = draw([text({ underline: { color: '#0000ff', style: 'single' } })], { scale: 3 });
  assert.deepEqual(scaled.rects, [{ x: 120, y: 307, width: 150, height: 4, fill: '#0000ff' }]);

  const words = draw([text({ underline: { color: '#0000ff', style: 'words' } })]);
  assert.deepEqual(words.rects, single.rects);

  const thick = draw([text({ underline: { color: '#0000ff', style: 'thick' } })]);
  assert.deepEqual(thick.rects, [{ x: 40, y: 103, width: 50, height: 2, fill: '#0000ff' }]);

  const double = draw([text({ underline: { color: '#0000ff', style: 'double' } })]);
  assert.equal(double.rects.length, 2);
  assert.ok(double.rects[1].y - (double.rects[0].y + double.rects[0].height) >= 1, 'a clear pixel between the two lines');

  const none = draw([text({ underline: { color: '#0000ff', style: 'none' } })]);
  assert.deepEqual(none.rects, []);
});

test('renderPage: dotted, dashed and wavy underlines', () => {
  const dotted = draw([text({ underline: { color: '#0000ff', style: 'dotted' } })]);
  assert.equal(dotted.strokes.length, 1);
  assert.deepEqual(dotted.strokes[0].dash, [1.8, 1.8]);
  assert.equal(dotted.strokes[0].color, '#0000ff');
  assert.deepEqual(dotted.strokes[0].points, [
    ['M', 40, 103.5],
    ['L', 90, 103.5],
  ]);
  assert.deepEqual(draw([text({ underline: { color: '#0000ff', style: 'dash' } })]).strokes[0].dash, [6, 3.6]);
  assert.equal(draw([text({ underline: { color: '#0000ff', style: 'dotDotDash' } })]).strokes[0].dash.length, 6);
  // Heavy styles are twice as thick, and their dashes with them.
  assert.deepEqual(draw([text({ underline: { color: '#0000ff', style: 'dashedHeavy' } })]).strokes[0].dash, [12, 7.2]);

  const wave = draw([text({ underline: { color: '#0000ff', style: 'wave' } })]);
  assert.equal(wave.strokes.length, 1);
  assert.equal(wave.strokes[0].points[0][0], 'M');
  assert.ok(wave.strokes[0].points.length > 5 && wave.strokes[0].points.slice(1).every((point) => point[0] === 'Q'));
  const last = wave.strokes[0].points[wave.strokes[0].points.length - 1];
  assert.equal(last[3], 90);
  assert.equal(draw([text({ underline: { color: '#0000ff', style: 'wavyDouble' } })]).strokes.length, 2);

  // A canvas without dashes still underlines, solid.
  const plain = draw([text({ underline: { color: '#0000ff', style: 'dotted' } })], {}, { dash: false });
  assert.equal(plain.strokes.length, 0);
  assert.equal(plain.rects.length, 1);
});

test('renderPage: a wave over an absurd length stays bounded', () => {
  const rec = draw([text({ width: 1e9, underline: { color: '#0000ff', style: 'wave' } })]);
  assert.ok(rec.strokes[0].points.length <= 4001);
});

test('renderPage: strike-through, single and double, in the colour of the text', () => {
  const single = draw([text({ strike: true })]);
  assert.deepEqual(single.rects, [{ x: 40, y: 94, width: 50, height: 1, fill: '#112233' }]);
  const double = draw([text({ strike: 'double' })]);
  assert.deepEqual(double.rects.map((rect) => [rect.y, rect.height]), [
    [93, 1],
    [95, 1],
  ]);
});

test('renderPage: letter spacing uses the canvas where it can', () => {
  const rec = draw([text({ letterSpacing: 2, width: 60 }), text({ y: 200, text: 'after' })], {}, { letterSpacing: true });
  assert.equal(rec.texts.length, 2);
  assert.equal(rec.texts[0].spacing, '2px');
  assert.deepEqual([rec.texts[0].x, rec.texts[0].y], [40, 100]);
  // The spacing does not leak into what is drawn next.
  assert.equal(rec.texts[1].spacing, '0px');
});

test('renderPage: letter spacing without canvas support is drawn character by character', () => {
  const rec = draw([text({ text: 'ab\u{1F600}c', letterSpacing: 3 })]);
  // 10 px a character (half of 20 px) plus 3 px after every UTF-16 unit, as the layout counts it.
  assert.deepEqual(
    rec.texts.map((call) => [call.text, call.x]),
    [
      ['a', 40],
      ['b', 53],
      ['\u{1F600}', 66],
      ['c', 92],
    ],
  );
  assert.ok(rec.texts.every((call) => call.spacing === undefined));
});

test('renderPage: horizontally scaled text is stretched around its left edge only', () => {
  const rec = draw([text({ scale: 1.5 }), text({ y: 200, text: 'after' })], { scale: 2 });
  assert.deepEqual([rec.texts[0].x, rec.texts[0].y, rec.texts[0].sx, rec.texts[0].sy], [80, 200, 3, 2]);
  assert.deepEqual([rec.texts[1].sx, rec.texts[1].font], [2, '20px "Test", sans-serif']);

  // With spacing: the canvas spaces in the stretched space, so the spacing is divided by the stretch.
  const spaced = draw([text({ scale: 2, letterSpacing: 3 })], {}, { letterSpacing: true });
  assert.equal(spaced.texts[0].spacing, '1.5px');

  // Without canvas support: every character stretched, advancing by its stretched width plus the spacing.
  const manual = draw([text({ text: 'ab', scale: 2, letterSpacing: 3 })]);
  assert.deepEqual(
    manual.texts.map((call) => [call.text, call.x, call.sx]),
    [
      ['a', 40, 2],
      ['b', 63, 2],
    ],
  );
});

test('renderPage: single borders are whole canvas pixels wide, horizontal and vertical', () => {
  const rec = draw([line(), line({ x1: 30, y1: 10, x2: 30, y2: 60, width: 0.5 }), line({ width: 2.6, y1: 80, y2: 80 })]);
  assert.deepEqual(rec.rects, [
    { x: 10, y: 50, width: 100, height: 1, fill: '#ff0000' },
    { x: 30, y: 10, width: 1, height: 50, fill: '#ff0000' },
    { x: 10, y: 79, width: 100, height: 3, fill: '#ff0000' },
  ]);
  assert.equal(rec.strokes.length, 0);
  // At 150% a 1 px border is 2 canvas pixels, exactly.
  assert.deepEqual(draw([line()], { scale: 1.5 }).rects, [{ x: 15, y: 74, width: 150, height: 2, fill: '#ff0000' }]);
});

test('renderPage: styles this painter does not know, and the 3D ones, are single lines; none draws nothing', () => {
  const single = draw([line()]).rects;
  for (const style of ['thick', 'threeDEmboss', 'outset', 'inset', 'apples', 'constructor', '__proto__', '']) {
    assert.deepEqual(draw([line({ style })]).rects, single, style);
  }
  for (const style of ['none', 'nil']) assert.deepEqual(draw([line({ style })]).rects, [], style);
  assert.deepEqual(draw([line({ x2: 10 })]).rects, [], 'no length');
  assert.deepEqual(draw([line({ x1: NaN })]).rects, [], 'not a number');
  // A width that is missing or absurd is replaced or capped.
  assert.equal(draw([line({ width: NaN })]).rects[0].height, 1);
  assert.equal(draw([line({ width: 1e6 })]).rects[0].height, 64);
});

test('renderPage: double and triple borders are parallel lines centred on the border', () => {
  const double = draw([line({ style: 'double', width: 2 })]);
  assert.deepEqual(double.rects, [
    { x: 10, y: 47, width: 100, height: 2, fill: '#ff0000' },
    { x: 10, y: 51, width: 100, height: 2, fill: '#ff0000' },
  ]);
  const triple = draw([line({ style: 'triple', width: 1, x1: 50, y1: 10, x2: 50, y2: 90 })]);
  assert.deepEqual(triple.rects.map((rect) => [rect.x, rect.width, rect.y, rect.height]), [
    [48, 1, 10, 80],
    [50, 1, 10, 80],
    [52, 1, 10, 80],
  ]);
  // The same border given from the bottom up has its lines in the same places.
  assert.deepEqual(draw([line({ style: 'thinThickSmallGap', width: 3, x1: 50, y1: 90, x2: 50, y2: 10 })]).rects, draw([line({ style: 'thinThickSmallGap', width: 3, x1: 50, y1: 10, x2: 50, y2: 90 })]).rects);
  const thinThick = draw([line({ style: 'thinThickSmallGap', width: 3 })]);
  assert.deepEqual(thinThick.rects.map((rect) => [rect.y, rect.height]), [
    [48, 1],
    [50, 3],
  ]);
  assert.equal(draw([line({ style: 'thinThickThinLargeGap', width: 3 })]).rects.length, 3);
});

test('renderPage: dashed and dotted borders', () => {
  const dash = (style: string, width = 1, features?: Parameters<typeof recorder>[0]) => draw([line({ style, width })], {}, features);
  const dashed = dash('dashed');
  assert.deepEqual(dashed.rects, []);
  assert.deepEqual(dashed.strokes, [
    {
      points: [
        ['M', 10, 50.5],
        ['L', 110, 50.5],
      ],
      color: '#ff0000',
      width: 1,
      dash: [4, 3],
    },
  ]);
  assert.deepEqual(dash('dotted', 2).strokes[0].dash, [2, 2]);
  assert.deepEqual(dash('dashSmallGap').strokes[0].dash, [4, 1.5]);
  assert.deepEqual(dash('dotDash').strokes[0].dash, [4, 2, 1, 2]);
  assert.deepEqual(dash('dotDotDash').strokes[0].dash, [4, 2, 1, 2, 1, 2]);
  // A hairline keeps dashes one canvas pixel long at least.
  assert.deepEqual(dash('dotted', 0.25).strokes[0].dash, [1, 1]);
  // The next line is solid again.
  const mixed = draw([line({ style: 'dashed' }), line({ y1: 70, y2: 70 }), line({ x1: 0, y1: 0, x2: 30, y2: 40 })]);
  assert.deepEqual(mixed.strokes[1].dash, []);
  // No dashes on this canvas: the border is still there.
  const plain = dash('dashed', 1, { dash: false });
  assert.deepEqual(plain.rects, [{ x: 10, y: 50, width: 100, height: 1, fill: '#ff0000' }]);
});

test('renderPage: wavy and slanted borders are stroked paths', () => {
  const wave = draw([line({ style: 'wave' })]);
  assert.equal(wave.strokes.length, 1);
  assert.ok(wave.strokes[0].points.some((point) => point[0] === 'Q'));
  assert.equal(draw([line({ style: 'doubleWave' })]).strokes.length, 2);
  const slanted = draw([line({ x1: 0, y1: 0, x2: 30, y2: 40, width: 2 })], { scale: 2 });
  assert.deepEqual(slanted.strokes, [
    {
      points: [
        ['M', 0, 0],
        ['L', 60, 80],
      ],
      color: '#ff0000',
      width: 4,
      dash: [],
    },
  ]);
  const slantedDouble = draw([line({ style: 'double', x1: 0, y1: 0, x2: 30, y2: 40 })]);
  assert.equal(slantedDouble.strokes.length, 2);
});

const loaded = (width = 200, height = 100): { store: { get(part: string): PageImage | undefined }; asked: string[]; source: object } => {
  const source = { picture: true };
  const asked: string[] = [];
  return {
    source,
    asked,
    store: {
      get(part) {
        asked.push(part);
        return part === '/word/media/image1.png' ? { source: source as unknown as CanvasImageSource, width, height } : undefined;
      },
    },
  };
};

test('renderPage: a loaded picture fills its box; a crop selects the part of the picture', () => {
  const pictures = loaded();
  const rec = draw([image(), image({ x: 200, crop: { left: 0.1, top: 0.2, right: 0.3, bottom: 0.4 } })], { images: pictures.store, scale: 2 });
  assert.equal(rec.images.length, 2);
  assert.equal(rec.images[0].source, pictures.source);
  assert.deepEqual(rec.images[0].args, [0, 0, 200, 100, 20, 30, 100, 50]);
  assert.deepEqual(rec.images[0].matrix, [2, 0, 0, 2, 0, 0]);
  assert.deepEqual(rec.images[1].args.map(round), [20, 20, 120, 40, 200, 30, 100, 50]);
  assert.deepEqual(rec.rects, []);
});

test('renderPage: crops from a file are kept inside the picture', () => {
  const pictures = loaded();
  const crops = [
    { left: -1, top: -0.5, right: NaN, bottom: -3 },
    { left: 0.6, top: 0, right: 0.6, bottom: 0 },
    { left: 5, top: 5, right: 5, bottom: 5 },
  ];
  for (const crop of crops) {
    const rec = draw([image({ crop })], { images: pictures.store });
    assert.deepEqual(rec.images[0].args.slice(0, 4), [0, 0, 200, 100], JSON.stringify(crop));
  }
});

test('renderPage: rotation and flips turn the picture around its centre', () => {
  const pictures = loaded();
  const rec = draw([image({ rotation: 90 }), image({ flipH: true }), image({ flipV: true, rotation: 360 + 180 }), text({ text: 'after' })], { images: pictures.store });
  // Centre (70, 55): a quarter turn clockwise maps the top left corner (20, 30) to (95, 5).
  const corner = (m: Matrix, x: number, y: number) => [round(m[0] * x + m[2] * y + m[4]), round(m[1] * x + m[3] * y + m[5])];
  assert.deepEqual(corner(rec.images[0].matrix, 20, 30), [95, 5]);
  assert.deepEqual(corner(rec.images[1].matrix, 20, 30), [120, 30]);
  // Half a turn and a vertical flip together mirror left and right.
  assert.deepEqual(corner(rec.images[2].matrix, 20, 30), [120, 30]);
  // What follows is drawn straight again.
  assert.deepEqual([rec.texts[0].x, rec.texts[0].y, rec.texts[0].sx], [40, 100, 1]);
});

test('renderPage: a picture that is not loaded is a grey box with a border and its label', () => {
  const pictures = loaded();
  const rec = draw([image({ part: '/word/media/other.png', alt: 'Logo' })], { images: pictures.store, placeholderLabel: (item) => item.alt });
  assert.deepEqual(pictures.asked, ['/word/media/other.png']);
  assert.equal(rec.images.length, 0);
  assert.deepEqual(rec.rects[0], { x: 20, y: 30, width: 100, height: 50, fill: '#f1f3f4' });
  assert.equal(rec.rects.length, 5);
  assert.ok(rec.rects.slice(1).every((rect) => rect.fill === '#bdc1c6' && (rect.width === 1 || rect.height === 1)));
  assert.deepEqual(rec.texts.map((call) => [call.text, call.x, call.y, call.align]), [['Logo', 70, 55, 'center']]);
});

test('renderPage: charts, shapes and linked pictures are placeholders and are never asked for', () => {
  const pictures = loaded();
  const rec = draw([image({ placeholder: 'chart' }), image({ part: undefined, y: 100 })], { images: pictures.store, placeholderLabel: (item) => item.placeholder });
  assert.deepEqual(pictures.asked, []);
  assert.equal(rec.images.length, 0);
  assert.deepEqual(rec.texts.map((call) => call.text), ['chart']);
  // Without a store, and without a label function, there is still a box.
  assert.equal(draw([image()]).rects.length, 5);
});

test('renderPage: a placeholder label is cut to the box; a small box has none', () => {
  const long = draw([image({ alt: 'A very long description of the picture that cannot fit' })], { placeholderLabel: (item) => item.alt });
  assert.equal(long.texts.length, 1);
  // 11 px label font: 5.5 px a character, 92 px of room.
  assert.ok(long.texts[0].text.length <= 16 && long.texts[0].text.endsWith('…'), long.texts[0].text);
  assert.equal(draw([image({ width: 30, alt: 'Logo' })], { placeholderLabel: (item) => item.alt }).texts.length, 0);
  assert.equal(draw([image({ height: 10, alt: 'Logo' })], { placeholderLabel: (item) => item.alt }).texts.length, 0);
  // A huge label is not measured character by character.
  const huge = draw([image({ alt: 'x'.repeat(100000) })], { placeholderLabel: (item) => item.alt });
  assert.ok(huge.measured.length < 20);
});

test('renderPage: a picture the canvas refuses becomes a placeholder; boxes without a size are skipped', () => {
  const pictures = loaded();
  const refused = draw([image(), text({ text: 'after' })], { images: pictures.store }, { failImages: true });
  assert.equal(refused.rects.length, 5);
  assert.equal(refused.texts[0].text, 'after');
  const empty = draw([image({ width: 0 }), image({ height: -1 }), image({ x: NaN })], { images: pictures.store });
  assert.equal(empty.images.length + empty.rects.length, 0);
  const noSize = draw([image()], { images: { get: () => ({ source: {} as CanvasImageSource, width: 0, height: 0 }) } });
  assert.equal(noSize.images.length, 0);
  assert.equal(noSize.rects.length, 5);
});

test('renderPage: the state of the context is restored even when drawing throws', () => {
  const rec = recorder();
  (rec.ctx as unknown as { fillText: unknown }).fillText = () => {
    throw new Error('boom');
  };
  assert.throws(() => renderPage(rec.ctx, page([text()])), /boom/);
  assert.equal(rec.depth(), 0);
});

test('renderPage: an unusable scale is 1', () => {
  for (const scale of [0, -2, NaN, Infinity]) {
    const rec = draw([text()], { scale });
    assert.deepEqual([rec.texts[0].x, rec.texts[0].sx], [40, 1], String(scale));
  }
});

// ------------------------------------------------------------------ measurer

function measureContext(known: string[] = [], metrics: 'font' | 'actual' | 'none' = 'font'): { context: MeasureContext; calls: string[] } {
  const calls: string[] = [];
  const context: MeasureContext = {
    font: '',
    measureText(textToMeasure: string) {
      calls.push(`${this.font}|${textToMeasure}`);
      const size = Number((/(\d+(?:\.\d+)?)px/.exec(this.font) || [])[1] || 10);
      // A known family measures differently from the generic family behind it.
      const first = (/"([^"]+)"/.exec(this.font) || [])[1];
      const factor = first && known.indexOf(first) >= 0 ? 0.6 : 0.5;
      const out: ReturnType<MeasureContext['measureText']> = { width: textToMeasure.length * size * factor };
      if (metrics === 'font') {
        out.fontBoundingBoxAscent = size * 0.9;
        out.fontBoundingBoxDescent = size * 0.3;
      } else if (metrics === 'actual') {
        out.actualBoundingBoxAscent = size * 0.95;
        out.actualBoundingBoxDescent = size * 0.1;
      }
      return out;
    },
  };
  return { context, calls };
}

const font = (family: string, size = 16, bold = false, italic = false): LayoutFont => ({ family, size, bold, italic });

test('doc measure: Office fonts are followed by fonts with the same metrics and a generic family', () => {
  const measurer = new CanvasTextMeasurer();
  assert.equal(measurer.cssFamilies('Calibri'), '"Calibri", "Carlito", sans-serif');
  assert.equal(measurer.cssFamilies('Cambria'), '"Cambria", "Caladea", serif');
  assert.equal(measurer.cssFamilies('Arial'), '"Arial", "Liberation Sans", "Arimo", "Helvetica", sans-serif');
  assert.equal(measurer.cssFamilies('Times New Roman'), '"Times New Roman", "Liberation Serif", "Tinos", "Times", serif');
  assert.equal(measurer.cssFamilies('COURIER NEW'), '"COURIER NEW", "Liberation Mono", "Cousine", "Courier", monospace');
  assert.equal(measurer.cssFamilies('Some Font'), '"Some Font", sans-serif');
  assert.equal(measurer.cssFont(font('Calibri', 14.666, true, true)), 'italic bold 14.67px "Calibri", "Carlito", sans-serif');
  assert.equal(docCssFont(font('Arial', 10)), '10px "Arial", "Liberation Sans", "Arimo", "Helvetica", sans-serif');
});

test('doc measure: the font table of the document picks the alternative and the generic family', () => {
  const fonts = new Map([
    ['Antiqua X', { family: 'roman', altName: 'Book Antiqua' }],
    ['Fixed X', { pitch: 'fixed' }],
    ['Typewriter', { family: 'modern' }],
    ['Hand', { family: 'script', altName: 'Hand' }],
    ['Plain', { family: 'swiss' }],
    ['constructor', {}],
  ]);
  const measurer = new CanvasTextMeasurer({ fonts });
  assert.equal(measurer.cssFamilies('Antiqua X'), '"Antiqua X", "Book Antiqua", serif');
  assert.equal(measurer.cssFamilies('Fixed X'), '"Fixed X", monospace');
  assert.equal(measurer.cssFamilies('Typewriter'), '"Typewriter", monospace');
  assert.equal(measurer.cssFamilies('Hand'), '"Hand", cursive');
  assert.equal(measurer.cssFamilies('Plain'), '"Plain", sans-serif');
  // Names that are also properties of every object are just names.
  assert.equal(measurer.cssFamilies('constructor'), '"constructor", sans-serif');
  assert.equal(measurer.cssFamilies('__proto__'), '"proto", sans-serif');
});

test('doc measure: a family name from a file cannot leave its quotes', () => {
  const measurer = new CanvasTextMeasurer({ fonts: new Map([['Evil', { altName: 'x", url(http://example.com/a), "y' }]]) });
  assert.equal(measurer.cssFamilies('a"; } body { background: url(//example.com) } "'), '"a body background urlexamplecom", sans-serif');
  assert.equal(measurer.cssFamilies('Evil'), '"Evil", "x urlhttpexamplecoma y", sans-serif');
  assert.equal(measurer.cssFamilies(''), 'sans-serif');
  assert.equal(measurer.cssFamilies('";;'), 'sans-serif');
  for (const size of [NaN, -5, 0, Infinity, 1e9]) {
    const css = measurer.cssFont(font('A', size));
    assert.match(css, /^(0\.5|4000)px "A", sans-serif$/, String(size));
  }
});

test('doc measure: widths come from the canvas, with the font set only when it changes, and are cached', () => {
  const { context, calls } = measureContext();
  const measurer = new CanvasTextMeasurer({ context });
  const regular = font('Test', 20);
  assert.equal(measurer.width(regular, 'abcd'), 40);
  assert.equal(measurer.width(regular, 'abcd'), 40);
  assert.equal(measurer.width(regular, ''), 0);
  assert.equal(measurer.width(font('Test', 20, true), 'ab'), 20);
  assert.deepEqual(calls, ['20px "Test", sans-serif|abcd', 'bold 20px "Test", sans-serif|ab']);
  measurer.clear();
  measurer.width(regular, 'abcd');
  assert.equal(calls.length, 3);
});

test('doc measure: a very long text is measured in pieces that never split a surrogate pair', () => {
  const { context, calls } = measureContext();
  const measurer = new CanvasTextMeasurer({ context });
  const long = 'a'.repeat(4095) + '\u{1F600}' + 'b'.repeat(6000);
  assert.equal(measurer.width(font('Test', 10), long), long.length * 5);
  const pieces = calls.map((call) => call.split('|')[1]);
  assert.equal(pieces.join(''), long);
  assert.ok(pieces.length >= 3 && pieces.every((piece) => piece.length <= 4096));
  assert.equal(pieces[0].length, 4095);
});

test('doc measure: line metrics are the font box of the canvas, or an estimate where it has none', () => {
  const regular = font('Test', 20);
  const withBox = measureContext();
  const measurer = new CanvasTextMeasurer({ context: withBox.context });
  assert.deepEqual(measurer.metrics(regular), { ascent: 18, descent: 6, lineGap: 0 });
  measurer.metrics(regular);
  assert.equal(withBox.calls.length, 1, 'cached');

  // Only the ink of the letters is known: it can raise the estimate, never lower it.
  assert.deepEqual(new CanvasTextMeasurer({ context: measureContext([], 'actual').context }).metrics(regular), { ascent: 19, descent: 20 * 0.24, lineGap: 0 });
  assert.deepEqual(new CanvasTextMeasurer({ context: measureContext([], 'none').context }).metrics(regular), { ascent: 20 * 0.91, descent: 20 * 0.24, lineGap: 0 });
});

test('doc measure: hasFont is true when the family or a font with its metrics is there', () => {
  const measurer = new CanvasTextMeasurer({ context: measureContext(['Carlito', 'Arial']).context });
  assert.equal(measurer.hasFont('Arial'), true);
  // Calibri itself is missing, its substitute is not.
  assert.equal(measurer.hasFont('Calibri'), true);
  assert.equal(measurer.hasFont('Cambria'), false);
  assert.equal(measurer.hasFont('No Such Font'), false);
  assert.equal(measurer.hasFont('";'), false);

  const asked: string[] = [];
  const custom = new CanvasTextMeasurer({
    hasFamily: (family) => {
      asked.push(family);
      return family === 'Caladea';
    },
  });
  assert.equal(custom.hasFont('Cambria'), true);
  assert.equal(custom.hasFont('Cambria'), true);
  assert.deepEqual(asked, ['Cambria', 'Caladea'], 'answers are remembered');
});

test('doc measure: without a canvas the error says so', () => {
  assert.throws(() => new CanvasTextMeasurer().width(font('Test'), 'x'), /no 2D canvas context/);
});

// ------------------------------------------------- helpers of the document view

const bytes = (...values: (number | string)[]): Uint8Array => {
  const out: number[] = [];
  for (const value of values) {
    if (typeof value === 'number') out.push(value);
    else for (let i = 0; i < value.length; i++) out.push(value.charCodeAt(i));
  }
  while (out.length < 32) out.push(0);
  return Uint8Array.from(out);
};
const be32 = (value: number) => [(value >>> 24) & 255, (value >>> 16) & 255, (value >>> 8) & 255, value & 255];

test('document view: only raster pictures are recognised, by their content', () => {
  assert.equal(sniffImageType(bytes(0x89, 'PNG', 13, 10, 26, 10)), 'image/png');
  assert.equal(sniffImageType(bytes(0xff, 0xd8, 0xff, 0xe0)), 'image/jpeg');
  assert.equal(sniffImageType(bytes('GIF89a')), 'image/gif');
  assert.equal(sniffImageType(bytes('BM')), 'image/bmp');
  assert.equal(sniffImageType(bytes('RIFF', 0, 0, 0, 0, 'WEBPVP8 ')), 'image/webp');
  assert.equal(sniffImageType(bytes('RIFF', 0, 0, 0, 0, 'WAVEfmt ')), undefined);
  // Vector formats and everything else are never handed to the browser.
  assert.equal(sniffImageType(bytes('<svg xmlns="http://www.w3.org/2000/svg">')), undefined);
  assert.equal(sniffImageType(bytes('<?xml version="1.0"?><svg>')), undefined);
  assert.equal(sniffImageType(bytes(1, 0, 0, 0, 0x6c, 0, 0, 0)), undefined, 'EMF');
  assert.equal(sniffImageType(bytes(0xd7, 0xcd, 0xc6, 0x9a)), undefined, 'WMF');
  assert.equal(sniffImageType(bytes('II*', 0)), undefined, 'TIFF');
  assert.equal(sniffImageType(new Uint8Array(0)), undefined);
  assert.equal(sniffImageType(Uint8Array.from([0x89, 0x50, 0x4e, 0x47])), undefined, 'too short to be a picture');
});

test('document view: the size a picture declares is read without decoding it', () => {
  assert.equal(declaredImagePixels(bytes(0x89, 'PNG', 13, 10, 26, 10, 0, 0, 0, 13, 'IHDR', ...be32(640), ...be32(480)), 'image/png'), 640 * 480);
  assert.equal(declaredImagePixels(bytes(0x89, 'PNG', 13, 10, 26, 10, 0, 0, 0, 13, 'IHDR', ...be32(0xffffffff), ...be32(0xffffffff)), 'image/png'), 0xffffffff * 0xffffffff);
  assert.equal(declaredImagePixels(bytes('GIF89a', 0x20, 0x03, 0x58, 0x02), 'image/gif'), 800 * 600);
  // BMP: a negative height means top-down rows.
  const bmp = bytes('BM', ...new Array(16).fill(0), 100, 0, 0, 0, 0x38, 0xff, 0xff, 0xff);
  assert.equal(declaredImagePixels(bmp, 'image/bmp'), 100 * 200);
  // JPEG: APP0, then a table, then the frame header with height 300 and width 400.
  const jpeg = bytes(0xff, 0xd8, 0xff, 0xe0, 0, 4, 1, 2, 0xff, 0xdb, 0, 3, 9, 0xff, 0xc0, 0, 11, 8, 1, 0x2c, 1, 0x90, 3);
  assert.equal(declaredImagePixels(jpeg, 'image/jpeg'), 300 * 400);
  // A JPEG whose segments lead nowhere says nothing, and the search ends.
  assert.equal(declaredImagePixels(bytes(0xff, 0xd8, 0xff, 0xe0, 0, 0), 'image/jpeg'), undefined);
  assert.equal(declaredImagePixels(bytes(0xff, 0xd8, 0xff, 0xe0, 0xff, 0xff), 'image/jpeg'), undefined);
  assert.equal(declaredImagePixels(bytes(0xff, 0xd8, 1, 2, 3), 'image/jpeg'), undefined);
  const padded = new Uint8Array(70000).fill(0xff);
  padded[1] = 0xd8;
  assert.equal(declaredImagePixels(padded, 'image/jpeg'), undefined);
  const vp8x = bytes('RIFF', 0, 0, 0, 0, 'WEBPVP8X', 0, 0, 0, 0, 0, 0, 0, 0, 0xff, 0x03, 0, 0xff, 0x01, 0);
  assert.equal(declaredImagePixels(vp8x, 'image/webp'), 1024 * 512);
  assert.equal(declaredImagePixels(bytes('RIFF', 0, 0, 0, 0, 'WEBPXXXX'), 'image/webp'), undefined);
});

test('document view: the text of a document is its paragraphs, with table cells separated by tabs', async () => {
  const doc = await readSpec({
    body: p('First') + p(r('Second ') + r('line', '<w:b/>')) + tbl(tr(tc('a') + tc(p('b1') + p('b2'))) + tr(tc('c') + tc(tbl(tr(tc('in') + tc('side')), [1000, 1000]))), [2000, 2000]) + p('') + p('Last'),
  });
  assert.equal(documentText(doc), 'First\nSecond line\na\tb1\nb2\nc\tin\tside\n\nLast');
});

test('document view: its words in Turkish and English', () => {
  assert.equal(getUiStrings('tr-TR').pageOf(3, 12), 'Sayfa 3 / 12');
  assert.equal(getUiStrings('en').pageOf(3, 12), 'Page 3 / 12');
  assert.equal(getUiStrings('de'), STRINGS_EN);
  assert.equal(STRINGS_TR.page(7), 'Sayfa 7');
  assert.equal(STRINGS_TR.zoomLevel(125), "Yakınlaştırma %125, %100'e döndür");
  assert.equal(STRINGS_EN.zoomLevel(125), 'Zoom 125%, reset to 100%');
  assert.deepEqual(['picture', 'chart', 'diagram', 'object', 'shape', 'vml', 'anything'].map(STRINGS_TR.placeholder), ['Görsel', 'Grafik', 'Diyagram', 'Gömülü nesne', 'Çizim', 'Çizim', 'Çizim']);
  assert.deepEqual(['picture', 'chart', 'diagram', 'object', 'shape'].map(STRINGS_EN.placeholder), ['Picture', 'Chart', 'Diagram', 'Embedded object', 'Drawing']);
  // Every string of one language exists in the other.
  assert.deepEqual(Object.keys(STRINGS_TR).sort(), Object.keys(STRINGS_EN).sort());
});

// ------------------------------------------------ fixes after the layout review

test('renderPage: text turned by a quarter is drawn along its own baseline, with its highlight and underline', () => {
  const rec = draw(
    [
      text({ rotation: 270, highlight: '#ffff00', underline: { color: '#0000ff', style: 'single' } }),
      text({ rotation: 90, x: 200, highlight: '#00ff00' }),
      text({ text: 'after', x: 300 }),
    ],
    { metrics: () => ({ ascent: 16, descent: 4 }) },
  );
  // Upwards from (40, 100): the line of the font (16 above, 4 below the baseline) lies left and right of x = 40.
  assert.deepEqual(rec.rects[0], { x: 24, y: 50, width: 20, height: 50, fill: '#ffff00' });
  // The underline: 2.4 px "below" the baseline, which is to the right of it.
  assert.deepEqual(rec.rects[1], { x: 43, y: 50, width: 1, height: 50, fill: '#0000ff' });
  // Downwards from (200, 100): the ascent is on the right.
  assert.deepEqual(rec.rects[2], { x: 196, y: 100, width: 20, height: 50, fill: '#00ff00' });
  assert.deepEqual(
    rec.texts.map((call) => [call.text, call.x, call.y, call.sx]),
    [
      ['Hello', 40, 100, 0],
      ['Hello', 200, 100, 0],
      // What follows is drawn straight again.
      ['after', 300, 100, 1],
    ],
  );
  // A turn that is no quarter is not a turn.
  const odd = draw([text({ rotation: 45 as unknown as 90 })]);
  assert.deepEqual([odd.texts[0].x, odd.texts[0].y, odd.texts[0].sx], [40, 100, 1]);
});

/** A canvas whose font box is rounded to whole pixels, as some browsers report it; `box` is [ascent, descent] per em. */
function roundingContext(box: [number, number] = [0.93, 0.27]): { context: MeasureContext; calls: string[] } {
  const calls: string[] = [];
  const context: MeasureContext = {
    font: '',
    measureText(textToMeasure: string) {
      calls.push(`${this.font}|${textToMeasure}`);
      const size = Number((/(\d+(?:\.\d+)?)px/.exec(this.font) || [])[1] || 10);
      return { width: textToMeasure.length * size * 0.5, fontBoundingBoxAscent: Math.round(size * box[0]), fontBoundingBoxDescent: Math.round(size * box[1]) };
    },
  };
  return { context, calls };
}

const near = (actual: number, expected: number, tolerance: number, what: string) =>
  assert.ok(Math.abs(actual - expected) <= tolerance, `${what}: ${actual}, expected ${expected} within ${tolerance}`);
const lineOf = (metrics: { ascent: number; descent: number; lineGap: number }) => metrics.ascent + metrics.descent + metrics.lineGap;

test('doc measure: common fonts take their line metrics from a table, whatever the canvas says', () => {
  const { context, calls } = roundingContext();
  const measurer = new CanvasTextMeasurer({ context });
  // Word sets Times New Roman 12 pt on lines of 13.8 pt and Arial 11 pt on lines of 12.65 pt.
  near(lineOf(measurer.metrics(font('Times New Roman', 16))), 18.4, 0.01, 'Times New Roman 12 pt');
  near(lineOf(measurer.metrics(font('Arial', 44 / 3))), 16.87, 0.01, 'Arial 11 pt');
  near(lineOf(measurer.metrics(font('Arial', 40 / 3))), 15.33, 0.01, 'Arial 10 pt');
  near(lineOf(measurer.metrics(font('Calibri', 44 / 3))), (44 / 3) * 1.2207, 0.01, 'Calibri 11 pt');
  near(lineOf(measurer.metrics(font('Courier New', 16))), 16 * 1.1328, 0.01, 'Courier New 12 pt');
  near(lineOf(measurer.metrics(font('Verdana', 16))), 16 * 1.2153, 0.01, 'Verdana 12 pt');
  near(lineOf(measurer.metrics(font('Tahoma', 16))), 16 * 1.207, 0.01, 'Tahoma 12 pt');
  near(lineOf(measurer.metrics(font('Georgia', 16))), 16 * 1.1362, 0.01, 'Georgia 12 pt');
  // The parts: ascent, descent and the leading Word puts between lines.
  const arial = measurer.metrics(font('Arial', 2048));
  assert.deepEqual([arial.ascent, arial.descent, arial.lineGap], [1854, 434, 67]);
  const times = measurer.metrics(font('Times New Roman', 2048));
  assert.deepEqual([times.ascent, times.descent, times.lineGap], [1825, 443, 87]);
  // The name is what the document asks for, in any case; bold and italic share the metrics; substitutes have their own entry.
  assert.deepEqual(measurer.metrics(font('ARIAL', 2048, true, true)), arial);
  assert.deepEqual(measurer.metrics(font(' arial ', 2048)), arial);
  assert.deepEqual(measurer.metrics(font('Liberation Sans', 2048)), arial);
  assert.deepEqual(measurer.metrics(font('Liberation Serif', 2048)), times);
  // Nothing was measured for any of this.
  assert.deepEqual(calls, []);
});

test('doc measure: other fonts are measured at a large size, so that rounding by the canvas does not show', () => {
  const { context, calls } = roundingContext();
  const measurer = new CanvasTextMeasurer({ context });
  // At 13 px the canvas would say 12 and 4; the real font box is 12.09 and 3.51.
  const metrics = measurer.metrics(font('Some Font', 13));
  near(metrics.ascent, 12.09, 0.001, 'ascent');
  near(metrics.descent, 3.51, 0.001, 'descent');
  assert.equal(metrics.lineGap, 0);
  assert.deepEqual(calls, ['1000px "Some Font", sans-serif|Hg']);
  // Other sizes of the family need no other measurement; bold is another face.
  near(measurer.metrics(font('Some Font', 26)).ascent, 24.18, 0.001, 'ascent at 26 px');
  assert.equal(calls.length, 1);
  measurer.metrics(font('Some Font', 26, true));
  assert.deepEqual(calls[1], 'bold 1000px "Some Font", sans-serif|Hg');
  // After clear() the canvas is asked again.
  measurer.clear();
  measurer.metrics(font('Some Font', 13));
  assert.equal(calls.length, 3);
});

test('renderPage: the placeholder of a mirrored picture is turned with it but its words are not mirrored', () => {
  const rec = draw([image({ part: '/word/media/other.png', alt: 'Logo', flipH: true, flipV: true }), image({ part: '/word/media/other.png', alt: 'Logo', y: 150, flipH: true, rotation: 180 })], {
    images: loaded().store,
    placeholderLabel: (item) => item.alt,
  });
  // Not flipped: the label reads from left to right, upright.
  assert.deepEqual(rec.texts.map((call) => [call.text, call.x, call.y, call.sx, call.sy]), [
    ['Logo', 70, 55, 1, 1],
    // Half a turn turns the words with the box; the flip still does not mirror them.
    ['Logo', 70, 175, -1, -1],
  ]);
  assert.deepEqual(rec.rects[0], { x: 20, y: 30, width: 100, height: 50, fill: '#f1f3f4' });
});

test('renderPage: the picture store is told how many canvas pixels the whole picture covers', () => {
  const asked: (number | undefined)[][] = [];
  const store = {
    get(part: string, width?: number, height?: number) {
      asked.push([width, height]);
      return part === '' ? undefined : undefined;
    },
  };
  // 100 x 50 px at twice the size; then a picture of which the crop shows half the width and a quarter of the height.
  draw([image(), image({ crop: { left: 0.25, top: 0.5, right: 0.25, bottom: 0.25 } }), image({ crop: { left: 0.6, top: 0, right: 0.6, bottom: 0 } })], { images: store, scale: 2 });
  assert.deepEqual(asked, [
    [200, 100],
    [400, 400],
    // A crop that leaves nothing is not applied.
    [200, 100],
  ]);
});

test('document view: the size a picture declares, and the size it is decoded to', () => {
  const png = bytes(0x89, 'PNG', 13, 10, 26, 10, 0, 0, 0, 13, 'IHDR', ...be32(6500), ...be32(100));
  assert.deepEqual(declaredImageSize(png, 'image/png'), { width: 6500, height: 100 });
  assert.equal(declaredImagePixels(png, 'image/png'), 650000);
  assert.equal(declaredImageSize(bytes(0x89, 'PNG'), 'image/x-unknown'), undefined);

  const limit = (96 * 1024 * 1024) / 8 / 4;
  // Shown at 200 px: decoded at the smallest halving of 6500 px that is not below it.
  assert.deepEqual(decodedImageSize(6500, 6500, 200, 200, limit), { width: 203, height: 203 });
  // Shown larger than the budget allows one picture to be: as many pixels as its share of the budget.
  const capped = decodedImageSize(6500, 6500, 6500, 6500, limit);
  assert.deepEqual(capped, { width: 1773, height: 1773 });
  assert.ok(capped.width * capped.height <= limit);
  // No size asked for: the picture as it is, within the same limit.
  assert.deepEqual(decodedImageSize(6500, 6500, 0, 0, limit), capped);
  assert.deepEqual(decodedImageSize(300, 200, 0, 0, limit), { width: 300, height: 200 });
  // Never larger than the picture is, and not below a small size that a little zooming would outgrow at once.
  assert.deepEqual(decodedImageSize(100, 50, 1000, 1000, limit), { width: 100, height: 50 });
  assert.deepEqual(decodedImageSize(1000, 1000, 10, 10, limit), { width: 125, height: 125 });
  // The side that needs more decides.
  assert.deepEqual(decodedImageSize(4000, 1000, 100, 500, limit), { width: 2000, height: 500 });
});
