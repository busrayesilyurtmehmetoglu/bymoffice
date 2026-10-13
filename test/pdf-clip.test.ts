// What is in the text of a PDF page is what can be seen on it: text that a
// clip cuts off (the next cell of a row, the margin of the page, the frame of
// a chart) is caught only as far as it shows, text of a cell is read cell by
// cell, runs that mix two directions are written word by word, and a run
// with numbers that cannot be does not bring the file down.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import type { LayoutFont, LayoutPage, PageItem, TextItem } from '../src/layout/doc/types';
import { buildPdf } from '../src/pdf/document';
import { PdfTextFont } from '../src/pdf/font';
import { captureText, clipRun, docPageRuns, textBlock, textLayer, visualOrder, type PdfTextRun } from '../src/pdf/text';
import { lay, p, r, tbl, tc, tr } from './doc-layout-fixtures';
import { checkPdf, parseToUnicode, poppler, popplerTest } from './pdf-fixtures';

interface Matrix {
  a: number;
  b: number;
  c: number;
  d: number;
  e: number;
  f: number;
}

/** A canvas as the painters use it: a transform that is saved and restored, a path, a clip. Every character is half the font size wide. */
function fakeCanvas() {
  const drawn: string[] = [];
  const calls: string[] = [];
  const stack: Matrix[] = [];
  const proto = {
    getTransform(this: any): Matrix {
      return this.matrix;
    },
    setTransform(this: any, a: number, b: number, c: number, d: number, e: number, f: number) {
      this.matrix = { a, b, c, d, e, f };
    },
    translate(this: any, x: number, y: number) {
      const m: Matrix = this.matrix;
      this.matrix = { ...m, e: m.a * x + m.c * y + m.e, f: m.b * x + m.d * y + m.f };
    },
    measureText(this: any, text: string) {
      return { width: (text.length * Number((/([\d.]+)px/.exec(this.font) || [0, 0])[1])) / 2 };
    },
    fillText(text: string) {
      drawn.push(text);
    },
    save(this: any) {
      calls.push('save');
      stack.push(this.matrix);
    },
    restore(this: any) {
      calls.push('restore');
      const m = stack.pop();
      if (m) this.matrix = m;
    },
    beginPath() {
      calls.push('beginPath');
    },
    rect() {
      calls.push('rect');
    },
    moveTo() {
      calls.push('moveTo');
    },
    lineTo() {
      calls.push('lineTo');
    },
    arc() {
      calls.push('arc');
    },
    closePath() {
      calls.push('closePath');
    },
    clip() {
      calls.push('clip');
    },
  };
  const ctx = Object.create(proto) as typeof proto & { font: string; textAlign: string; textBaseline: string; direction: string; matrix: Matrix };
  ctx.font = '16px Test';
  ctx.textAlign = 'left';
  ctx.textBaseline = 'alphabetic';
  ctx.direction = 'ltr';
  ctx.matrix = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
  return { ctx, drawn, calls, proto };
}

/** Catches the text of `paint` on a fake canvas. */
function caught(paint: (ctx: ReturnType<typeof fakeCanvas>['ctx']) => void, bounds?: { width: number; height: number }, scale = 1): PdfTextRun[] {
  const { ctx } = fakeCanvas();
  const runs: PdfTextRun[] = [];
  const release = captureText(ctx, scale, (run) => runs.push(run), bounds);
  paint(ctx);
  release();
  return runs;
}

const clipTo = (ctx: ReturnType<typeof fakeCanvas>['ctx'], x: number, y: number, width: number, height: number) => {
  ctx.beginPath();
  ctx.rect(x, y, width, height);
  ctx.clip();
};

// --------------------------------------------------------------- the canvas

test('text that a clip cuts off is caught as far as it can be seen, and no further', () => {
  // A cell 64 wide from x = 10; sixteen characters of 8 each, of which eight fit.
  const runs = caught((ctx) => {
    ctx.save();
    clipTo(ctx, 10, 0, 64, 30);
    ctx.fillText('ABCDEFGHIJKLMNOP', 10, 20);
    ctx.restore();
  });
  assert.deepEqual(runs, [{ text: 'ABCDEFGH', x: 10, y: 20, width: 64, size: 16, cutAfter: true }]);
});

test('a character that shows in part is caught, one that only touches the edge is not', () => {
  const part = caught((ctx) => {
    clipTo(ctx, 10, 0, 60, 30);
    ctx.fillText('ABCDEFGHIJKLMNOP', 10, 20);
  });
  // "H" lies from 66 to 74 and the clip ends at 70: the letter is in the text, and the run ends where the clip does.
  assert.deepEqual(part, [{ text: 'ABCDEFGH', x: 10, y: 20, width: 60, size: 16, cutAfter: true }]);
  const left = caught((ctx) => {
    clipTo(ctx, 10, 0, 64, 30);
    ctx.textAlign = 'right';
    ctx.fillText('ABCDEFGHIJKLMNOP', 74, 20);
  });
  // Right aligned: what is cut off is the start of the text, and the run starts where the clip does.
  assert.deepEqual(left, [{ text: 'IJKLMNOP', x: 10, y: 20, width: 64, size: 16, cutBefore: true }]);
  const middle = caught((ctx) => {
    clipTo(ctx, 30, 0, 20, 30);
    ctx.fillText('ABCDEFGHIJKLMNOP', 10, 20);
  });
  // Characters 2 ("C", 26 to 34, of which 30 to 34 shows) to 4 ("E", 42 to 50); "F" starts at 50, where the clip ends.
  assert.deepEqual(middle, [{ text: 'CDE', x: 30, y: 20, width: 20, size: 16, cutBefore: true, cutAfter: true }]);
});

test('text that a clip hides altogether is not caught', () => {
  const runs = caught((ctx) => {
    ctx.save();
    clipTo(ctx, 10, 0, 64, 30);
    ctx.fillText('to the right', 80, 20);
    ctx.fillText('to the left', -200, 20);
    ctx.fillText('below', 10, 60);
    ctx.fillText('above', 10, -5);
    ctx.fillText('seen', 10, 20);
    ctx.restore();
  });
  assert.deepEqual(runs.map((run) => run.text), ['seen']);
});

test('a line whose letters reach into the clip from below or above is caught', () => {
  const runs = caught((ctx) => {
    clipTo(ctx, 0, 100, 500, 20);
    // The baseline lies below the clip, the letters reach up into it.
    ctx.fillText('reaches up', 0, 125);
    // The baseline lies above it, and what hangs below the line reaches down into it.
    ctx.fillText('hangs', 0, 99);
    // A whole line above and a whole line below.
    ctx.fillText('gone above', 0, 80);
    ctx.fillText('gone below', 0, 140);
  });
  assert.deepEqual(runs.map((run) => run.text), ['reaches up', 'hangs']);
});

test('a clip ends with the state it was made in, and clips inside each other narrow', () => {
  const runs = caught((ctx) => {
    ctx.save();
    clipTo(ctx, 0, 0, 100, 30);
    ctx.save();
    clipTo(ctx, 50, 0, 100, 30);
    // Both clips: 50 to 100.
    ctx.fillText('ABCDEFGHIJKLMNOPQRST', 0, 20);
    ctx.restore();
    // The outer one only.
    ctx.fillText('ABCDEFGHIJKLMNOPQRST', 0, 20);
    ctx.restore();
    ctx.fillText('ABCDEFGHIJKLMNOPQRST', 0, 20);
    // A restore too many changes nothing.
    ctx.restore();
    ctx.fillText('free', 0, 20);
  });
  assert.deepEqual(runs.map((run) => [run.text, run.x, run.width]), [
    ['GHIJKLM', 50, 50],
    ['ABCDEFGHIJKLM', 0, 100],
    ['ABCDEFGHIJKLMNOPQRST', 0, 160],
    ['free', 0, 32],
  ]);
});

test('a clip is where the transform of its path puts it', () => {
  const runs = caught((ctx) => {
    // The page is painted twice as large and moved; the clip is stated in units of the page.
    ctx.setTransform(2, 0, 0, 2, 100, 0);
    clipTo(ctx, 10, 0, 32, 30);
    ctx.fillText('ABCDEFGH', 10, 20);
  }, undefined, 2);
  // In px of the page (half a canvas pixel): the text starts at 60, and four characters of 8 are seen.
  assert.deepEqual(runs, [{ text: 'ABCD', x: 60, y: 20, width: 32, size: 16, cutAfter: true }]);
});

test('a clip that is no rectangle counts as the box around it, and an empty one hides everything', () => {
  const triangle = caught((ctx) => {
    ctx.beginPath();
    ctx.moveTo(20, 0);
    ctx.lineTo(60, 0);
    ctx.lineTo(40, 40);
    ctx.closePath();
    ctx.clip();
    ctx.fillText('ABCDEFGHIJKL', 0, 20);
  });
  assert.deepEqual(triangle.map((run) => [run.text, run.x]), [['CDEFGH', 20]]);
  const circle = caught((ctx) => {
    ctx.beginPath();
    ctx.arc(40, 20, 20, 0, Math.PI * 2);
    ctx.clip();
    ctx.fillText('ABCDEFGHIJKL', 0, 20);
  });
  assert.deepEqual(circle.map((run) => [run.text, run.x]), [['CDEFGH', 20]]);
  const empty = caught((ctx) => {
    ctx.beginPath();
    ctx.clip();
    ctx.fillText('nothing', 0, 20);
  });
  assert.deepEqual(empty, []);
});

test('a clip the canvas is not told of as a path hides what is painted under it', () => {
  const runs = caught((ctx) => {
    ctx.save();
    // A path object: where it lies cannot be followed.
    (ctx.clip as (path: object) => void)({});
    ctx.fillText('under an unknown clip', 0, 20);
    ctx.restore();
    ctx.fillText('after it', 0, 20);
  });
  assert.deepEqual(runs.map((run) => run.text), ['after it']);
});

test('the edge of the canvas is a clip too', () => {
  const runs = caught(
    (ctx) => {
      ctx.fillText('ABCDEFGHIJ', -40, 20);
      ctx.fillText('ABCDEFGHIJ', 160, 20);
      ctx.fillText('below the page', 0, 150);
    },
    { width: 200, height: 100 },
  );
  assert.deepEqual(runs.map((run) => [run.text, run.x, run.width]), [
    ['FGHIJ', 0, 40],
    ['ABCDE', 160, 40],
  ]);
});

test('turned text is cut along its own line', () => {
  const runs = caught((ctx) => {
    clipTo(ctx, 0, 0, 100, 40);
    // Turned by 90 degrees clockwise around (50, 0): the text reads downwards.
    ctx.setTransform(0, 1, -1, 0, 50, 0);
    ctx.fillText('ABCDEFGHIJ', 0, 0);
  });
  // Five characters of 8 reach down to 40.
  assert.deepEqual(runs, [{ text: 'ABCDE', x: 50, y: 0, width: 40, size: 16, angle: 90, cutAfter: true }]);
});

test('text from right to left is cut at its own end', () => {
  const runs = caught((ctx) => {
    clipTo(ctx, 10, 0, 32, 30);
    // Eight letters, 64 wide from x = 10: the first one stands at the right. What shows is the left half, the last four letters.
    ctx.fillText('אבגדהוזח', 10, 20);
  });
  // Along the line the text goes on to the right of what is seen.
  assert.deepEqual(runs, [{ text: 'הוזח', x: 10, y: 20, width: 32, size: 16, cutAfter: true }]);
});

test('a run of two directions that is cut off is left out rather than guessed at', () => {
  const runs = caught((ctx) => {
    clipTo(ctx, 10, 0, 32, 30);
    ctx.fillText('abc אבגד efg', 10, 20);
    ctx.fillText('אב c', 10, 20);
  });
  // The second one fits (4 x 8), and is caught whole.
  assert.deepEqual(runs.map((run) => run.text), ['אב c']);
});

test('a width the text is squeezed into is the width it is cut in', () => {
  const runs = caught((ctx) => {
    clipTo(ctx, 0, 0, 40, 30);
    // 160 wide, squeezed into 80: every character is 4 wide, and ten of them are seen.
    (ctx.fillText as (text: string, x: number, y: number, max: number) => void)('ABCDEFGHIJKLMNOPQRST', 0, 20, 80);
  });
  assert.deepEqual(runs, [{ text: 'ABCDEFGHIJ', x: 0, y: 20, width: 40, size: 16, cutAfter: true }]);
});

test('a pair of surrogates is never cut in two', () => {
  const runs = caught((ctx) => {
    clipTo(ctx, 0, 0, 20, 30);
    // "a" (0 to 8), the pair (8 to 24, two units), "b": the clip ends inside the second unit of the pair.
    ctx.fillText('a😀b', 0, 20);
  });
  assert.deepEqual(runs.map((run) => run.text), ['a😀']);
  const half = caught((ctx) => {
    clipTo(ctx, 0, 0, 14, 30);
    ctx.fillText('a😀b', 0, 20);
  });
  // Only its first unit lies in the clip: whether the character shows cannot be told, so it is left out.
  assert.deepEqual(half.map((run) => run.text), ['a']);
});

test('painting is as it was while the clips are followed, and the canvas is as it was afterwards', () => {
  const { ctx, drawn, calls, proto } = fakeCanvas();
  const runs: PdfTextRun[] = [];
  const release = captureText(ctx, 1, (run) => runs.push(run));
  ctx.save();
  clipTo(ctx, 0, 0, 10, 10);
  ctx.fillText('text', 0, 5);
  ctx.restore();
  assert.deepEqual(calls, ['save', 'beginPath', 'rect', 'clip', 'restore'], 'every call reached the canvas');
  assert.deepEqual(drawn, ['text']);
  release();
  for (const name of Object.keys(proto)) assert.ok(!Object.prototype.hasOwnProperty.call(ctx, name), `${name} is the method of the canvas again`);
  assert.equal(Object.keys(ctx).filter((key) => key.indexOf('bym') >= 0).length, 0, 'nothing of the listener is left on the canvas');
  ctx.fillText('later', 0, 5);
  assert.equal(runs.length, 1);
});

test('the lines a painter says belong together are one block', () => {
  const { ctx } = fakeCanvas();
  const runs: PdfTextRun[] = [];
  // Without a listener the painter's words go nowhere.
  textBlock(ctx, true);
  textBlock(ctx, false);
  const release = captureText(ctx, 1, (run) => runs.push(run));
  ctx.fillText('alone', 0, 10);
  textBlock(ctx, true);
  ctx.fillText('first line', 0, 30);
  ctx.fillText('second line', 0, 50);
  textBlock(ctx, false);
  textBlock(ctx, true);
  ctx.fillText('another cell', 200, 30);
  textBlock(ctx, false);
  ctx.fillText('alone again', 0, 90);
  release();
  assert.deepEqual(runs.map((run) => run.block), [undefined, 1, 1, 2, undefined]);
});

// ------------------------------------------------------------ cutting a run

test('a run is cut to a box by where its characters end', () => {
  const run: PdfTextRun = { text: 'abcdefghij', x: 100, y: 50, width: 100, size: 10 };
  const even = (count: number) => count * 10;
  assert.equal(clipRun(run, { left: 0, top: 0, right: 500, bottom: 500 }, even), run, 'a run that is seen whole is the run');
  assert.deepEqual(clipRun(run, { left: 125, top: 0, right: 500, bottom: 500 }, even), { text: 'cdefghij', x: 125, y: 50, width: 75, size: 10, cutBefore: true });
  assert.deepEqual(clipRun(run, { left: 0, top: 0, right: 130, bottom: 500 }, even), { text: 'abc', x: 100, y: 50, width: 30, size: 10, cutAfter: true });
  assert.equal(clipRun(run, { left: 0, top: 0, right: 100, bottom: 500 }, even), undefined);
  assert.equal(clipRun(run, { left: 0, top: 60, right: 500, bottom: 500 }, even), undefined, 'the letters end above the box');
  // Without a way to ask where the characters are, a run that is cut is left out.
  assert.equal(clipRun(run, { left: 125, top: 0, right: 500, bottom: 500 }), undefined);
  assert.equal(clipRun(run, { left: 0, top: 0, right: 500, bottom: 500 }), run);
});

// ---------------------------------------------------------------- documents

const FONT: LayoutFont = { family: 'Test', size: 16, bold: false, italic: false };
const item = (text: string, x: number, y: number, more: Partial<TextItem> = {}): TextItem => ({ type: 'text', x, y, width: text.length * 8, text, font: FONT, color: '#000000', ...more });
const pageOf = (items: PageItem[]): LayoutPage => ({ width: 800, height: 1000, section: 0, number: 1, items });
const measure = (font: LayoutFont, text: string) => (text.length * font.size) / 2;

test('text of a document that lies beyond the paper is not in its runs', () => {
  const page = pageOf([
    item('inside', 100, 100),
    item('ABCDEFGHIJ', 760, 200),
    item('ABCDEFGHIJ', -24, 300),
    item('off the right', 900, 400),
    item('off the bottom', 100, 1100),
    item('down the edge', 400, 960, { rotation: 90 }),
  ]);
  assert.deepEqual(
    docPageRuns(page, measure).map((run) => [run.text, run.x, run.y, run.width]),
    [
      ['inside', 100, 100, 48],
      ['ABCDE', 760, 200, 40],
      ['DEFGHIJ', 0, 300, 56],
      // Reading downwards from y = 960: five characters reach the lower edge.
      ['down ', 400, 960, 40],
    ],
  );
  // Without a measure a run that is cut is left out, and no part of it guessed.
  assert.deepEqual(docPageRuns(page).map((run) => run.text), ['inside']);
});

test('hidden text of a document is not laid out, so it is in no run', async () => {
  const laid = await lay(p(r('shown ') + r('HIDDEN', '<w:vanish/>') + r('too')));
  const text = docPageRuns(laid.pages[0], measure).map((run) => run.text).join('');
  assert.equal(text.replace(/\s+/g, ' ').trim(), 'shown too');
});

test('the lines of a table cell of a document are one block, and so is every other cell', async () => {
  const laid = await lay(tbl(tr(tc(p('one two three four five six seven eight nine ten')) + tc(p('alpha beta gamma delta epsilon zeta eta theta'))), [1500, 1500]) + p('after the table'));
  const runs = docPageRuns(laid.pages[0], measure);
  const blocksOf = (words: RegExp) => Array.from(new Set(runs.filter((entry) => words.test(entry.text)).map((entry) => entry.block)));
  const first = blocksOf(/one|two|three|four|five|six|seven|eight|nine|ten/);
  const second = blocksOf(/alpha|beta|gamma|delta|epsilon|zeta|eta|theta/);
  assert.equal(first.length, 1, 'the lines of a cell are one block');
  assert.ok(first[0] !== undefined, 'text of a cell is in a block');
  assert.equal(second.length, 1);
  assert.notEqual(second[0], first[0], 'another cell is another block');
  assert.ok(new Set(runs.filter((entry) => entry.block === first[0]).map((entry) => entry.y)).size > 1, 'the cell has more than one line');
  assert.deepEqual(blocksOf(/after/), [undefined], 'a paragraph of the body is in none');
});

// -------------------------------------------------------------------- order

const run = (text: string, x: number, y: number, more: Partial<PdfTextRun> = {}): PdfTextRun => ({ text, x, y, width: text.length * 6, size: 12, ...more });

function written(layer: string, font: PdfTextFont): string[] {
  const map = parseToUnicode(font.toUnicode());
  const out: string[] = [];
  for (const line of layer.split('\n')) {
    const shown = /<([0-9a-f]*)> Tj$/.exec(line);
    if (!shown) continue;
    let text = '';
    for (let i = 0; i < shown[1].length; i += 4) text += map.get(parseInt(shown[1].substr(i, 4), 16));
    out.push(text);
  }
  return out;
}

test('the lines of a block are written one after the other, block by block along the row', () => {
  const font = new PdfTextFont();
  // Two cells of a row, each with two lines; below them a row of cells with one line each.
  const runs = [
    run('Geminin sahibi', 60, 100, { block: 1 }),
    run('Teslim tarihi', 320, 100, { block: 2 }),
    run('Ahmet Yılmaz', 60, 115, { block: 1 }),
    run('12 Ekim 2026', 320, 115, { block: 2 }),
    run('sol', 60, 140, { block: 3 }),
    run('sağ', 320, 140),
  ];
  assert.deepEqual(written(textLayer(runs, { height: 800, scale: 1, font }), font), ['Geminin sahibi', ' ', 'Ahmet Yılmaz', ' ', 'Teslim tarihi', ' ', '12 Ekim 2026', ' ', 'sol', ' ', 'sağ', ' ']);
  // Without blocks the same runs are read line by line across the page, as before.
  const plain = runs.map((entry) => ({ ...entry, block: undefined }));
  assert.deepEqual(written(textLayer(plain, { height: 800, scale: 1, font }), font).filter((text) => text !== ' '), ['Geminin sahibi', 'Teslim tarihi', 'Ahmet Yılmaz', '12 Ekim 2026', 'sol', 'sağ']);
});

test('text that was cut off is kept apart from what stands next to it, however close', () => {
  const font = new PdfTextFont();
  // The cell on the left is cut off where the next one starts: no room between the two.
  const touching = [run('GIZLI-UZU', 10, 20, { width: 54, cutAfter: true }), run('komsu', 64, 20)];
  assert.deepEqual(written(textLayer(touching, { height: 100, scale: 1, font }), font), ['GIZLI-UZU', ' ', 'komsu', ' ']);
  // Two halves of a word in two fonts touch as well, and stay a word.
  const word = [run('GIZLI-UZU', 10, 20, { width: 54 }), run('komsu', 64, 20)];
  assert.deepEqual(written(textLayer(word, { height: 100, scale: 1, font }), font), ['GIZLI-UZU', 'komsu', ' ']);
  const before = [run('abc', 10, 20, { width: 18 }), run('tail', 28, 20, { cutBefore: true })];
  assert.deepEqual(written(textLayer(before, { height: 100, scale: 1, font }), font), ['abc', ' ', 'tail', ' ']);
});

popplerTest('Poppler copies a table cell by cell', async () => {
  const runs = [run('Geminin sahibi', 60, 100, { block: 1 }), run('Teslim tarihi', 320, 100, { block: 2 }), run('Ahmet Yılmaz', 60, 115, { block: 1 }), run('12 Ekim 2026', 320, 115, { block: 2 })];
  const report = poppler(await buildPdf([{ width: 595.28, height: 841.89, text: { runs, width: 794 } }], { date: new Date(0) }));
  assert.equal(report.complaints, '');
  assert.deepEqual(report.text[0].trim().split(/\s*\n/), ['Geminin sahibi', 'Ahmet Yılmaz', 'Teslim tarihi', '12 Ekim 2026']);
});

// -------------------------------------------------------- two directions

test('in a run of two directions every stretch from right to left is turned on its own', () => {
  assert.equal(visualOrder('abc مرحبا def'), 'abc ابحرم def');
  assert.equal(visualOrder('abc مرحبا بالعالم def'), 'abc ملاعلاب ابحرم def');
  assert.equal(visualOrder('No: שנת 2024, ok'), 'No: 2024 תנש, ok');
  assert.equal(visualOrder('שלום (abc) עולם'), 'םולש (abc) םלוע');
  assert.equal(visualOrder('plain text 123'), 'plain text 123');
  assert.equal(visualOrder('שלום עולם'), 'םלוע םולש', 'a run of one direction is turned as a whole');
});

popplerTest('Poppler finds each word of a run of two directions', async () => {
  const lines = ['abc مرحبا def', 'Total שלום 2024'];
  const runs = lines.map((text, index) => run(text, 96, 100 + index * 24, { width: text.length * 8, size: 16 }));
  const report = poppler(await buildPdf([{ width: 600, height: 750, text: { runs, width: 800 } }], { date: new Date(0) }));
  assert.equal(report.complaints, '');
  const read = report.reading[0].replace(/[‎‏‪-‮⁦-⁩]/g, '');
  for (const word of ['abc', 'مرحبا', 'def', 'Total', 'שלום', '2024']) assert.ok(read.indexOf(word) >= 0, `"${word}" is found in ${JSON.stringify(read)}`);
});

// ------------------------------------------------------------------ numbers

test('a run with numbers that cannot be written is left out, and the others are written', async () => {
  const font = new PdfTextFont();
  const runs = [
    run('far right', 1e300, 20),
    run('far down', 10, -1e300),
    run('giant', 10, 40, { size: 1e300 }),
    run('endless', 10, 60, { width: 1e300 }),
    run('beyond a number', 1e16, 80),
    run('fine', 10, 90),
  ];
  assert.deepEqual(written(textLayer(runs, { height: 100, scale: 1, font }), font), ['fine', ' ']);
  const file = await buildPdf([{ width: 300, height: 300, text: { runs, width: 300 } }], { date: new Date(0) });
  assert.equal(checkPdf(file).pages[0].text.trim(), 'fine');
});

// ------------------------------------------------- what the review left over

test('a pair of surrogates that the start of a clip cuts is left out too', () => {
  const runs = caught((ctx) => {
    clipTo(ctx, 12, 0, 100, 30);
    // The pair lies from 0 to 16 (two units of 8): only its second unit reaches into the clip.
    ctx.fillText('😀ab', 0, 20);
  });
  assert.deepEqual(runs.map((entry) => [entry.text, entry.x]), [['ab', 16]]);
});

test('a clip inside another one never widens it', () => {
  const runs = caught((ctx) => {
    clipTo(ctx, 50, 0, 50, 30);
    // Wider than the clip it is made in, on every side.
    clipTo(ctx, 0, -10, 300, 100);
    ctx.fillText('ABCDEFGHIJKLMNOPQRST', 0, 20);
  });
  assert.deepEqual(runs.map((entry) => [entry.text, entry.x, entry.width]), [['GHIJKLM', 50, 50]]);
});

test('a run too far away or too large to be on a page is left out, though its numbers could be written', () => {
  const font = new PdfTextFont();
  const runs = [run('far right', 5e6, 20), run('far up', 10, -5e6), run('huge', 10, 40, { size: 5e5 }), run('fine', 10, 90)];
  assert.deepEqual(written(textLayer(runs, { height: 100, scale: 1, font }), font), ['fine', ' ']);
});

test('what the layout puts before the text of a cell (the number of a list item) is read with the cell', () => {
  const inCell = (text: string, x: number, y: number, cell: number): TextItem => item(text, x, y, { source: { story: 'body', path: [0, 0, cell, 0], offset: 0 } });
  const runs = docPageRuns(pageOf([item('1.', 100, 100), inCell('first cell', 130, 100, 0), item('a)', 400, 100), inCell('second cell', 430, 100, 1), item('alone', 100, 300)]), measure);
  const blockOf = (text: string) => (runs.find((entry) => entry.text === text) as PdfTextRun).block;
  assert.ok(blockOf('first cell') !== undefined, 'the text of a cell is in a block');
  assert.equal(blockOf('1.'), blockOf('first cell'));
  assert.equal(blockOf('a)'), blockOf('second cell'));
  assert.notEqual(blockOf('first cell'), blockOf('second cell'));
  assert.equal(blockOf('alone'), undefined);
});
