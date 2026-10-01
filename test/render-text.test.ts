import assert from 'node:assert/strict';
import { test } from 'node:test';
import type { Font } from '../src/model/sheet/types';
import { clearTextMeasureCache, cssFont, FONT_FALLBACK, fontPixels, sanitizeFontFamily, TextPainter, wrapLines } from '../src/render/text';

const font = (init: Partial<Font>): Font => ({ bold: false, italic: false, strike: false, underline: 'none', ...init });

/** Every character is 10 wide. */
const measure = (text: string) => text.length * 10;

/** The lines as text. */
const wrapped = (text: string, width: number) => {
  const offsets = wrapLines(text, width, measure);
  const lines: string[] = [];
  for (let i = 0; i < offsets.length; i += 2) lines.push(text.slice(offsets[i], offsets[i + 1]));
  return lines;
};

/** A context that counts font assignments and measureText calls. */
function fakeContext() {
  const log = { fonts: [] as string[], fills: [] as string[], measured: [] as string[] };
  let current = '10px sans-serif';
  const ctx = {
    get font() {
      return current;
    },
    set font(value: string) {
      current = value;
      log.fonts.push(value);
    },
    set fillStyle(value: string) {
      log.fills.push(value);
    },
    measureText(text: string) {
      log.measured.push(text);
      return { width: text.length * 5 };
    },
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, log };
}

test('text: font family names keep only letters, digits, spaces and hyphens', () => {
  assert.equal(sanitizeFontFamily('Calibri'), 'Calibri');
  assert.equal(sanitizeFontFamily('  Times  New   Roman '), 'Times New Roman');
  assert.equal(sanitizeFontFamily('Source Sans-3'), 'Source Sans-3');
  assert.equal(sanitizeFontFamily('ＭＳ Ｐゴシック'), 'ＭＳ Ｐゴシック');
  assert.equal(sanitizeFontFamily('x"; } body { color: red'), 'x body color red');
  assert.equal(sanitizeFontFamily("a'b\\c\"d;e,f(g)h/i:j\nk"), 'abcdefghijk');
  assert.equal(sanitizeFontFamily('";,()'), '');
  assert.equal(sanitizeFontFamily(undefined), '');
  assert.equal(sanitizeFontFamily(42), '');
  assert.ok(sanitizeFontFamily('a'.repeat(100000)).length <= 128);
  // A surrogate pair cut in half by the length limit does not leave a lone surrogate behind.
  const cut = sanitizeFontFamily('a'.repeat(127) + '\u{20000}');
  assert.equal(cut, 'a'.repeat(127));
});

test('text: CSS font strings', () => {
  assert.equal(fontPixels(font({ size: 11 })), 14.67);
  assert.equal(fontPixels(font({ size: 12 })), 16);
  assert.equal(fontPixels(undefined), 14.67);
  for (const size of [0, -5, Number.NaN, Number.POSITIVE_INFINITY, undefined]) assert.equal(fontPixels(font({ size })), 14.67);
  assert.equal(fontPixels(font({ size: 100000 })), Math.round(((409 * 96) / 72) * 100) / 100);

  assert.equal(cssFont(font({ name: 'Arial', size: 12 })), `16px "Arial", ${FONT_FALLBACK}`);
  assert.equal(cssFont(font({ name: 'Arial', size: 9, bold: true, italic: true })), `italic bold 12px "Arial", ${FONT_FALLBACK}`);
  assert.equal(cssFont(font({ size: 12 })), `16px ${FONT_FALLBACK}`);
  assert.equal(cssFont(undefined), `14.67px ${FONT_FALLBACK}`);
  assert.equal(cssFont(font({ name: 'x"; } body { display: none', size: 12 })), `16px "x body display none", ${FONT_FALLBACK}`);
  // Values of the wrong type in a damaged model do not get through either.
  const damaged = { name: { toString: () => 'evil' }, size: '12px; x', bold: 'yes', italic: 1 } as unknown as Font;
  assert.equal(cssFont(damaged), `14.67px ${FONT_FALLBACK}`);
});

test('text: lines break at spaces', () => {
  assert.deepEqual(wrapped('aaa bbb ccc', 70), ['aaa bbb', 'ccc']);
  assert.deepEqual(wrapped('aaa bbb ccc', 110), ['aaa bbb ccc']);
  assert.deepEqual(wrapped('aaa bbb ccc', 30), ['aaa', 'bbb', 'ccc']);
  assert.deepEqual(wrapped('', 50), ['']);
  assert.deepEqual(wrapped('abc', 50), ['abc']);
  // Spaces at a break and at the end of a paragraph are dropped, spaces inside a line are kept.
  assert.deepEqual(wrapped('aa  bb    cc   ', 60), ['aa  bb', 'cc']);
  // Spaces that start a paragraph are an indent.
  assert.deepEqual(wrapped('  ab cd', 40), ['  ab', 'cd']);
});

test('text: explicit line breaks', () => {
  assert.deepEqual(wrapped('one\ntwo', 1000), ['one', 'two']);
  assert.deepEqual(wrapped('one\r\ntwo\r\n', 1000), ['one', 'two', '']);
  assert.deepEqual(wrapped('a\n\nb', 1000), ['a', '', 'b']);
  assert.deepEqual(wrapped('aaa bbb\nccc ddd eee', 70), ['aaa bbb', 'ccc ddd', 'eee']);
});

test('text: a word wider than the line breaks between characters', () => {
  assert.deepEqual(wrapped('abcdefghij', 30), ['abc', 'def', 'ghi', 'j']);
  assert.deepEqual(wrapped('ab cdefghij kl', 40), ['ab', 'cdef', 'ghij', 'kl']);
  // The rest of a broken word shares its line with the next word when that fits.
  assert.deepEqual(wrapped('abcde f', 30), ['abc', 'de', 'f']);
  assert.deepEqual(wrapped('abcd f', 30), ['abc', 'd f']);
  // Surrogate pairs stay whole.
  const emoji = '\u{1F600}\u{1F600}\u{1F600}';
  assert.deepEqual(wrapped(emoji, 20), ['\u{1F600}', '\u{1F600}', '\u{1F600}']);
});

test('text: every line holds a character whatever the width is', () => {
  for (const width of [0, -10, Number.NaN, 5]) {
    assert.deepEqual(wrapped('abc', width), ['a', 'b', 'c'], `width ${width}`);
  }
  assert.deepEqual(wrapped('a b', 0), ['a', 'b']);
});

test('text: a very long text wraps in bounded time', () => {
  const text = 'x'.repeat(32767);
  const started = Date.now();
  const offsets = wrapLines(text, 100, measure);
  assert.equal(offsets.length / 2, Math.ceil(32767 / 10));
  assert.equal(offsets[offsets.length - 1], text.length);
  // A huge word is never measured in one piece.
  let longest = 0;
  wrapLines(text, 1e9, (piece) => {
    longest = Math.max(longest, piece.length);
    return piece.length * 10;
  });
  assert.equal(longest, 1);
  assert.deepEqual(wrapLines('ab ' + text, 1e9, measure), [0, 3 + text.length]);
  const words = 'word '.repeat(6000);
  assert.ok(wrapLines(words, 100, measure).length > 0);
  assert.ok(Date.now() - started < 2000);
});

test('text: the painter assigns font and fill only when they change', () => {
  const { ctx, log } = fakeContext();
  const painter = new TextPainter(ctx);
  painter.setFont('12px a');
  painter.setFont('12px a');
  painter.setFont('13px a');
  painter.setFill('#000000');
  painter.setFill('#000000');
  assert.deepEqual(log.fonts, ['12px a', '13px a']);
  assert.deepEqual(log.fills, ['#000000']);
  // After a restore() the context may hold anything.
  painter.invalidate();
  painter.setFont('13px a');
  painter.setFill('#000000');
  assert.deepEqual(log.fonts, ['12px a', '13px a', '13px a']);
  assert.deepEqual(log.fills, ['#000000', '#000000']);
});

test('text: widths are cached by font and text', () => {
  clearTextMeasureCache();
  const { ctx, log } = fakeContext();
  const painter = new TextPainter(ctx);
  assert.equal(painter.measure('12px a', 'hello'), 25);
  assert.equal(painter.measure('12px a', 'hello'), 25);
  assert.equal(painter.measure('12px a', ''), 0);
  assert.deepEqual(log.measured, ['hello']);
  // The font is set before measuring, and another font is another entry.
  painter.measure('20px b', 'hello');
  assert.deepEqual(log.measured, ['hello', 'hello']);
  assert.deepEqual(log.fonts, ['12px a', '20px b']);
  // The cache is shared between painters.
  const other = fakeContext();
  new TextPainter(other.ctx).measure('12px a', 'hello');
  assert.deepEqual(other.log.measured, []);
  clearTextMeasureCache();
  painter.measure('12px a', 'hello');
  assert.equal(log.measured.length, 3);
});

test('text: the cache is bounded and long strings stay out of it', () => {
  clearTextMeasureCache();
  const { ctx, log } = fakeContext();
  const painter = new TextPainter(ctx);
  painter.measure('12px a', 'first');
  for (let i = 0; i < 5000; i++) painter.measure('12px a', 'n' + i);
  // 5,001 entries do not fit: the cache was emptied on the way, so the first entry is measured again.
  const before = log.measured.length;
  painter.measure('12px a', 'first');
  assert.equal(log.measured.length, before + 1);

  const long = 'y'.repeat(201);
  painter.measure('12px a', long);
  painter.measure('12px a', long);
  assert.equal(log.measured.filter((text) => text === long).length, 2);
  clearTextMeasureCache();
});
