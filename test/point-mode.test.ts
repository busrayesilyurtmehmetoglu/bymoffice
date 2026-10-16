// Pointing at cells while a formula is typed, without a DOM: where a reference
// can be put, what it is written as, and how the last one is replaced.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { MAX_POINT_TEXT, canPointAt, isFormulaText, pointReference, pointSpanLive, putReference } from '../src/edit/point';
import { MAX_COLS, MAX_ROWS } from '../src/model/sheet/types';
import { STRINGS_EN, STRINGS_TR } from '../src/ui/strings';

/** The text with "|" for the caret, or two of them around the selected part. */
function at(marked: string): boolean {
  const start = marked.indexOf('|');
  const second = marked.indexOf('|', start + 1);
  const text = marked.replace(/\|/g, '');
  return second < 0 ? canPointAt(text, start) : canPointAt(text, start, second - 1);
}

test('isFormulaText: only a text that begins with "="', () => {
  assert.equal(isFormulaText('='), true);
  assert.equal(isFormulaText('=A1'), true);
  assert.equal(isFormulaText(''), false);
  assert.equal(isFormulaText('+A1'), false);
  assert.equal(isFormulaText('-A1'), false);
  assert.equal(isFormulaText(' =A1'), false);
  assert.equal(isFormulaText("'=A1"), false);
  assert.equal(isFormulaText(undefined as unknown as string), false);
});

test('canPointAt: after "=", an operator, a separator or an opening bracket', () => {
  const yes = ['=|', '=A1+|', '=A1-|', '=A1*|', '=A1/|', '=A1^|', '=A1&|', '=SUM(|', '=SUM(A1;|', '=SUM(A1,|', '=A1:|', '=-|', '=1<|', '=1>|', '=1<=|', '=1>=|', '=1<>|', '=(|', '=((|'];
  for (const marked of yes) assert.equal(at(marked), true, marked);
});

test('canPointAt: not after an operand, a closing bracket, a percent sign or "!"', () => {
  const no = ['=A1|', '=1|', '=SUM|', '=SUM(A1)|', '=50%|', "='Sayfa 1'!|", '=Sheet2!|', '=A1 |', '=A1  |', '=TRUE|', '="a"|', '={|', '={1;|', '={1,2}|'];
  for (const marked of no) assert.equal(at(marked), false, marked);
});

test('canPointAt: spaces after an operator do not count, a space after an operand is the intersection', () => {
  assert.equal(at('=A1+ |'), true);
  assert.equal(at('=SUM(A1; |'), true);
  assert.equal(at('=A1+\n|'), true);
  assert.equal(at('= |'), true);
  assert.equal(at('=A1 |'), false);
  assert.equal(at('=SUM(A1) |'), false);
});

test('canPointAt: never inside a text in quotes or a sheet name in apostrophes', () => {
  assert.equal(at('="a+|'), false);
  assert.equal(at('="a+|"'), false);
  assert.equal(at('="a""+|'), false);
  assert.equal(at("='My sheet+|"), false);
  assert.equal(at("='It''s+|'!A1"), false);
  // Behind a text that is closed, after an operator, it can.
  assert.equal(at('="a"&|'), true);
  assert.equal(at('="a""b"&|'), true);
  assert.equal(at("='My sheet'!A1+|"), true);
  // An apostrophe inside a text in quotes opens no name.
  assert.equal(at('="it\'s"&|'), true);
  // A selection that reaches into a text is no place for a reference.
  assert.equal(at('=A1+|"ab|c"'), false);
  assert.equal(at('=A1+"|ab"|'), false);
});

test('canPointAt: what follows the caret has to be the end, an operator, a separator or ")"', () => {
  assert.equal(at('=SUM(|)'), true);
  assert.equal(at('=A1+|+B1'), true);
  assert.equal(at('=SUM(|;B1)'), true);
  assert.equal(at('=IF(|,1,2)'), true);
  assert.equal(at('=A1+| '), true);
  assert.equal(at('=A1+| )'), true);
  assert.equal(at('=|A1'), false);
  assert.equal(at('=A1+|B1'), false);
  assert.equal(at('=A1+| B1'), false);
  assert.equal(at('=SUM(|"a")'), false);
  assert.equal(at('=A1+|(B1)'), false);
});

test('canPointAt: a selection is replaced when both of its ends are places for a reference', () => {
  assert.equal(at('=A1+|B1|'), true);
  assert.equal(at('=|A1+B1|'), true);
  assert.equal(at('=SUM(|A1:B2|)'), true);
  assert.equal(at('=A|1+B1|'), false);
  assert.equal(at('=A1+|B|1'), false);
});

test('canPointAt: not in a text that is no formula, and not with a caret that is none', () => {
  assert.equal(canPointAt('', 0), false);
  assert.equal(canPointAt('abc', 3), false);
  assert.equal(canPointAt('+A1+', 4), false);
  assert.equal(canPointAt('-A1+', 4), false);
  assert.equal(canPointAt("'=A1+", 5), false);
  // Before the "=" nothing is put.
  assert.equal(canPointAt('=', 0), false);
  assert.equal(canPointAt('=A1+', 0, 4), false);
  assert.equal(canPointAt('=', 2), false);
  assert.equal(canPointAt('=', -1), false);
  assert.equal(canPointAt('=', 1.5), false);
  assert.equal(canPointAt('=', Number.NaN), false);
  assert.equal(canPointAt('=A1+', 4, 3), false);
  assert.equal(canPointAt('=A1+', 4, 9), false);
  assert.equal(canPointAt(undefined as unknown as string, 1), false);
});

test('canPointAt: not behind the sign of an exponent', () => {
  for (const marked of ['=1E+|', '=1.5E-|', '=1e+|', '=.5E-|', '=12.E+|', '=A1+1E+|', '=SUM(2e-|', '=A1*3E-|)']) assert.equal(at(marked), false, marked);
  // A cell or a name that ends in "E" is no number.
  for (const marked of ['=B2E+|', '=RATE+|', '=E+|', '=A1E+|', '=x1E-|', '=1E +|', '=1+|', '=1E3+|', '=_1E+|']) assert.equal(at(marked), true, marked);
});

test('canPointAt: behind a long run of digits in linear time', () => {
  // The pattern for the exponent once shared the digits out between two loops in every way: 0.7 s for this.
  const text = `=${'1'.repeat(MAX_POINT_TEXT - 3)}+`;
  const start = Date.now();
  assert.equal(canPointAt(text, text.length), true);
  assert.equal(canPointAt(`=${'1'.repeat(MAX_POINT_TEXT - 4)}E+`, MAX_POINT_TEXT - 1), false);
  assert.ok(Date.now() - start < 100, `${Date.now() - start} ms`);
});

test('canPointAt: not inside the square brackets of a reference to a table', () => {
  for (const marked of ['=T[[#This Row],|', '=T[[A]:|', '=T[|', '=SUM(T[[A]:|', '=T[[#All],[A]:|']) assert.equal(at(marked), false, marked);
  // Behind the closed brackets, after an operator, it can.
  assert.equal(at('=T[A]+|'), true);
  assert.equal(at('=SUM(T[[A]:[B]];|'), true);
  // A bracket inside a text in quotes opens nothing.
  assert.equal(at('="["&|'), true);
  // An apostrophe inside the brackets is part of a column name, not the start of a sheet name.
  assert.equal(at("=T[It's]+|"), true);
});

test('canPointAt: not inside an error constant', () => {
  for (const marked of ['=#N/|', '=#DIV/|', '=A1+#N/|', '=IF(A1;#DIV/|', '=#SAYI/|', '=#n/|']) assert.equal(at(marked), false, marked);
  // A division after an error constant that is complete, or after anything else, takes a reference.
  for (const marked of ['=#REF!/|', '=#N/A/|', '=#DIV/0!/|', '=A1/|', '=#REF!+1/|', '=(#N/A)/|']) assert.equal(at(marked), true, marked);
});

test('canPointAt: where numbers have a decimal comma, a comma behind a digit is no separator', () => {
  const tr = { decimal: ',' };
  const point = (marked: string, context?: { decimal?: string }) => canPointAt(marked.replace('|', ''), marked.indexOf('|'), marked.indexOf('|'), context);
  assert.equal(point('=1,|', tr), false);
  assert.equal(point('=A1+12,|', tr), false);
  // Inside a list of arguments too: the comma is no separator in that locale at all.
  assert.equal(point('=TOPLA(1,|', tr), false);
  assert.equal(point('=TOPLA(A1,|', tr), false);
  // The semicolon always separates.
  assert.equal(point('=TOPLA(1;|', tr), true);
  assert.equal(point('=TOPLA(A1;|', tr), true);
  assert.equal(point('=1+|', tr), true);
  // Not behind a digit: left to the user, as everywhere.
  assert.equal(point('=TOPLA(A:A,|', tr), true);
  // With a decimal point, and without a locale, the comma separates.
  assert.equal(point('=SUM(1,|', { decimal: '.' }), true);
  assert.equal(point('=SUM(A1,|', { decimal: '.' }), true);
  assert.equal(point('=SUM(1,|'), true);
  assert.equal(point('=SUM(1,|', {}), true);
});

test('canPointAt: a text longer than any formula is never read', () => {
  const long = `=${'1+'.repeat(MAX_POINT_TEXT / 2)}`;
  assert.equal(long.length, MAX_POINT_TEXT + 1);
  assert.equal(canPointAt(long, long.length), false);
  const fits = long.substring(0, MAX_POINT_TEXT - 1);
  assert.equal(fits.charAt(fits.length - 1), '+');
  assert.equal(canPointAt(fits, fits.length), true);
});

test('pointReference: a cell and a range, the corners in any order', () => {
  assert.equal(pointReference({ top: 4, left: 1, bottom: 4, right: 1 }), 'B5');
  assert.equal(pointReference({ top: 4, left: 1, bottom: 8, right: 3 }), 'B5:D9');
  // Dragged up and to the left.
  assert.equal(pointReference({ top: 8, left: 3, bottom: 4, right: 1 }), 'B5:D9');
  assert.equal(pointReference({ top: 4, left: 3, bottom: 8, right: 1 }), 'B5:D9');
  assert.equal(pointReference({ top: 0, left: 0, bottom: 0, right: 0 }), 'A1');
  assert.equal(pointReference({ top: MAX_ROWS - 1, left: MAX_COLS - 1, bottom: MAX_ROWS - 1, right: MAX_COLS - 1 }), 'XFD1048576');
});

test('pointReference: a merged cell is its top-left cell', () => {
  assert.equal(pointReference({ top: 4, left: 3, bottom: 5, right: 4 }, { single: true }), 'D5');
  assert.equal(pointReference({ top: 4, left: 3, bottom: 5, right: 4 }, { single: false }), 'D5:E6');
});

test('pointReference: whole columns and whole rows', () => {
  assert.equal(pointReference({ top: 0, left: 2, bottom: MAX_ROWS - 1, right: 2 }), 'C:C');
  assert.equal(pointReference({ top: 0, left: 3, bottom: MAX_ROWS - 1, right: 1 }), 'B:D');
  assert.equal(pointReference({ top: 2, left: 0, bottom: 2, right: MAX_COLS - 1 }), '3:3');
  assert.equal(pointReference({ top: 6, left: 0, bottom: 2, right: MAX_COLS - 1 }), '3:7');
  assert.equal(pointReference({ top: 0, left: 0, bottom: MAX_ROWS - 1, right: MAX_COLS - 1 }), '1:1048576');
  // Whole columns stay whole columns for a merged cell flag too.
  assert.equal(pointReference({ top: 0, left: 2, bottom: MAX_ROWS - 1, right: 2 }, { single: true }), 'C:C');
});

test('pointReference: another sheet, in apostrophes when its name needs them', () => {
  const cell = { top: 4, left: 1, bottom: 4, right: 1 };
  assert.equal(pointReference(cell, { sheet: 'Sayfa2' }), 'Sayfa2!B5');
  assert.equal(pointReference(cell, { sheet: 'Satış 2026' }), "'Satış 2026'!B5");
  assert.equal(pointReference(cell, { sheet: "It's" }), "'It''s'!B5");
  assert.equal(pointReference(cell, { sheet: '2026' }), "'2026'!B5");
  // A name that reads as a cell is quoted.
  assert.equal(pointReference(cell, { sheet: 'A1' }), "'A1'!B5");
  assert.equal(pointReference({ top: 4, left: 1, bottom: 8, right: 3 }, { sheet: 'Data' }), 'Data!B5:D9');
  assert.equal(pointReference({ top: 0, left: 2, bottom: MAX_ROWS - 1, right: 2 }, { sheet: 'My data' }), "'My data'!C:C");
  assert.equal(pointReference({ top: 2, left: 0, bottom: 2, right: MAX_COLS - 1 }, { sheet: 'Data' }), 'Data!3:3');
});

test('pointReference: the text is the same in every locale, and never absolute', () => {
  const text = pointReference({ top: 1, left: 1, bottom: 3, right: 2 }, { sheet: 'Sayfa1' });
  assert.equal(text, 'Sayfa1!B2:C4');
  assert.equal(/[$;,]/.test(text), false);
});

test('pointReference: a range outside the sheet is refused', () => {
  assert.throws(() => pointReference({ top: -1, left: 0, bottom: 0, right: 0 }), RangeError);
  assert.throws(() => pointReference({ top: 0, left: 0, bottom: MAX_ROWS, right: 0 }), RangeError);
  assert.throws(() => pointReference({ top: 0, left: 0, bottom: 0, right: MAX_COLS }), RangeError);
  assert.throws(() => pointReference({ top: 0.5, left: 0, bottom: 0, right: 0 }), RangeError);
  assert.throws(() => pointReference({ top: Number.NaN, left: 0, bottom: 0, right: 0 }), RangeError);
});

test('putReference: at the caret, and in the place of a selection', () => {
  assert.deepEqual(putReference('=', 1, 1, 'B5'), { text: '=B5', span: { start: 1, end: 3, text: 'B5' } });
  assert.deepEqual(putReference('=SUM()', 5, 5, 'B2:C4'), { text: '=SUM(B2:C4)', span: { start: 5, end: 10, text: 'B2:C4' } });
  assert.deepEqual(putReference('=A1+B1', 4, 6, 'Data!C3'), { text: '=A1+Data!C3', span: { start: 4, end: 11, text: 'Data!C3' } });
  // Positions that are none end up inside the text.
  assert.deepEqual(putReference('=', 9, 2, 'A1'), { text: '=A1', span: { start: 1, end: 3, text: 'A1' } });
  assert.deepEqual(putReference('=', Number.NaN, Number.NaN, 'A1').text, '=A1');
});

test('the reference pointed at last is replaced by the next one while the caret stands behind it', () => {
  const first = putReference('=SUM(', 5, 5, 'B5');
  assert.equal(first.text, '=SUM(B5');
  assert.equal(pointSpanLive(first.text, first.span.end, first.span.end, first.span), true);
  // The next cell takes its place, a longer or a shorter reference alike.
  const second = putReference(first.text, first.span.start, first.span.end, 'B5:D9');
  assert.equal(second.text, '=SUM(B5:D9');
  assert.equal(pointSpanLive(second.text, second.span.end, second.span.end, second.span), true);
  const third = putReference(second.text, second.span.start, second.span.end, "'Satış 2026'!A1");
  assert.equal(third.text, "=SUM('Satış 2026'!A1");
  const fourth = putReference(third.text, third.span.start, third.span.end, 'C:C');
  assert.equal(fourth.text, '=SUM(C:C');
});

test('pointSpanLive: not once something was typed, the caret moved or the text changed', () => {
  const { text, span } = putReference('=', 1, 1, 'B5');
  assert.equal(pointSpanLive(text, 3, 3, span), true);
  // Typed on: the caret is further on.
  assert.equal(pointSpanLive(`${text}+`, 4, 4, span), false);
  // The caret went back into the reference, or part of it is selected.
  assert.equal(pointSpanLive(text, 2, 2, span), false);
  assert.equal(pointSpanLive(text, 1, 3, span), false);
  // The reference was typed over with something as long.
  assert.equal(pointSpanLive('=C7', 3, 3, span), false);
  // The text is no formula any more.
  assert.equal(pointSpanLive('xB5', 3, 3, span), false);
  // A span that does not fit its text.
  assert.equal(pointSpanLive(text, 3, 3, { start: 1, end: 3, text: 'B55' }), false);
  assert.equal(pointSpanLive('=', 3, 3, span), false);
  assert.equal(pointSpanLive(text, 3, 3, { start: 0, end: 3, text: '=B5' }), false);
});

test('after the reference comes what the user types: the next place for a reference is behind the separator', () => {
  let text = putReference('=TOPLA(', 7, 7, 'B2').text;
  assert.equal(canPointAt(text, text.length), false);
  text += ';';
  assert.equal(canPointAt(text, text.length), true);
  text = putReference(text, text.length, text.length, 'C3').text;
  assert.equal(text, '=TOPLA(B2;C3');
});

test('what is said when a reference was added, in both languages', () => {
  assert.equal(STRINGS_EN.referenceAdded('B5:D9'), 'Reference B5:D9');
  assert.equal(STRINGS_TR.referenceAdded('Sayfa2!B5'), 'Başvuru Sayfa2!B5');
  assert.equal(STRINGS_EN.editSheetGone, 'The sheet of the cell is gone; what was typed was not written.');
  assert.equal(STRINGS_TR.editSheetGone, 'Hücrenin sayfası artık yok; yazılan kaydedilmedi.');
});
