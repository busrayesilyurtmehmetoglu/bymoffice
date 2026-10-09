// Pointing at cells for a field of a dialog of rules: where the reference goes in a list of
// ranges and in a value or formula, how it is written, and that what is written is read back
// by the dialog's own reading of its fields.
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { fieldSpanLive, formulaFieldTarget, pointReference, putReference, rangeFieldTarget, type FieldTarget } from '../src/edit/point';
import { MAX_COLS, MAX_ROWS, type CellRange } from '../src/model/sheet/types';
import { LOCALE_EN, LOCALE_TR } from '../src/format/locale';
import { conditionalFields, conditionalFromFields, listSeparator, parseRangeList, validationFields, validationFromFields } from '../src/ui/rules-dialog';
import { STRINGS_EN, STRINGS_TR } from '../src/ui/strings';

const range = (top: number, left: number, bottom = top, right = left): CellRange => ({ top, left, bottom, right });

/** Writes a reference at a target the way the view does: the lead first, the reference behind it. */
function write(text: string, target: FieldTarget, reference: string): { text: string; start: number; end: number } {
  const put = putReference(text, target.start, target.end, target.lead + reference);
  return { text: put.text, start: put.span.start + target.lead.length, end: put.span.end };
}

test('an absolute reference has "$" before every column and row', () => {
  assert.equal(pointReference(range(0, 0), { absolute: true }), '$A$1', 'one cell');
  assert.equal(pointReference(range(0, 0, 8, 0), { absolute: true }), '$A$1:$A$9', 'a range');
  assert.equal(pointReference(range(8, 3, 0, 1), { absolute: true }), '$B$1:$D$9', 'corners in any order');
  assert.equal(pointReference(range(0, 1, MAX_ROWS - 1, 3), { absolute: true }), '$B:$D', 'whole columns');
  assert.equal(pointReference(range(2, 0, 6, MAX_COLS - 1), { absolute: true }), '$3:$7', 'whole rows');
  assert.equal(pointReference(range(4, 3, 5, 4), { absolute: true, single: true }), '$D$5', 'a merged cell is its first cell');
  assert.equal(pointReference(range(0, 0, 2, 0), { absolute: true, sheet: 'Satış 2026' }), "'Satış 2026'!$A$1:$A$3", 'the name of the sheet is quoted as ever');
});

test('a reference stays relative unless asked for otherwise', () => {
  assert.equal(pointReference(range(4, 1, 8, 3)), 'B5:D9', 'no option');
  assert.equal(pointReference(range(4, 1, 8, 3), { absolute: false }), 'B5:D9', 'absolute: false');
  assert.equal(pointReference(range(0, 1, MAX_ROWS - 1, 3)), 'B:D', 'whole columns');
  assert.equal(pointReference(range(0, 0), { sheet: 'Other' }), 'Other!A1', 'another sheet');
});

test('a range pointed at takes the place of the whole list of ranges', () => {
  assert.deepEqual(rangeFieldTarget('A1:B5; D1:D9', false, ';'), { start: 0, end: 12, lead: '' }, 'a list of two');
  assert.deepEqual(rangeFieldTarget('', false, ','), { start: 0, end: 0, lead: '' }, 'an empty field');
  assert.deepEqual(rangeFieldTarget('not a range', false, ','), { start: 0, end: 11, lead: '' }, 'whatever the field holds');
  const done = write('A1:B5; D1:D9', rangeFieldTarget('A1:B5; D1:D9', false, ';'), 'C2:C4');
  assert.deepEqual(done, { text: 'C2:C4', start: 0, end: 5 }, 'the reference is all there is');
});

test('with Ctrl or Cmd the range is added behind the list, after the separator of the locale', () => {
  const tr = listSeparator(LOCALE_TR);
  const en = listSeparator(LOCALE_EN);
  assert.equal(tr, ';', 'Turkish lists with ";"');
  assert.equal(en, ',', 'English lists with ","');
  assert.deepEqual(write('A1:B5', rangeFieldTarget('A1:B5', true, tr), 'D1:D9'), { text: 'A1:B5; D1:D9', start: 7, end: 12 }, 'Turkish');
  assert.deepEqual(write('A1:B5', rangeFieldTarget('A1:B5', true, en), 'D1:D9'), { text: 'A1:B5, D1:D9', start: 7, end: 12 }, 'English');
  assert.deepEqual(write('', rangeFieldTarget('', true, en), 'D1'), { text: 'D1', start: 0, end: 2 }, 'nothing to add to: the range alone');
  assert.deepEqual(write('  ', rangeFieldTarget('  ', true, en), 'D1'), { text: 'D1', start: 0, end: 2 }, 'only spaces: the range alone');
  assert.deepEqual(write('A1; ', rangeFieldTarget('A1; ', true, tr), 'D1'), { text: 'A1; D1', start: 4, end: 6 }, 'a separator the list ends in is not doubled');
  assert.deepEqual(write('A1 B2,', rangeFieldTarget('A1 B2,', true, en), '3:7'), { text: 'A1 B2, 3:7', start: 7, end: 10 }, 'whole rows after a list typed with spaces');
});

test('a list of ranges that was pointed at is read by the dialog', () => {
  for (const [name, locale] of [['tr', LOCALE_TR], ['en', LOCALE_EN]] as const) {
    const separator = listSeparator(locale);
    let text = write('Z9', rangeFieldTarget('Z9', false, separator), pointReference(range(0, 0, 4, 1))).text;
    text = write(text, rangeFieldTarget(text, true, separator), pointReference(range(0, 3, 8, 3))).text;
    text = write(text, rangeFieldTarget(text, true, separator), pointReference(range(0, 5, MAX_ROWS - 1, 6))).text;
    assert.deepEqual(
      parseRangeList(text, locale),
      [range(0, 0, 4, 1), range(0, 3, 8, 3), range(0, 5, MAX_ROWS - 1, 6)],
      `three ranges pointed at in ${name}: ${text}`,
    );
  }
});

test('an empty value field, or one whose whole text is selected, takes "=" and the reference', () => {
  assert.deepEqual(formulaFieldTarget('', 0, 0), { start: 0, end: 0, lead: '=' }, 'empty');
  assert.deepEqual(formulaFieldTarget('  ', 1, 1), { start: 0, end: 2, lead: '=' }, 'only spaces');
  assert.deepEqual(formulaFieldTarget('5', 0, 1), { start: 0, end: 1, lead: '=' }, 'a number that is selected');
  assert.deepEqual(formulaFieldTarget('a; b; c', 0, 7), { start: 0, end: 7, lead: '=' }, 'the items of a list, selected');
  assert.deepEqual(formulaFieldTarget('=A1+1', 0, 5), { start: 0, end: 5, lead: '=' }, 'a formula that is selected with its "="');
  assert.deepEqual(write('5', formulaFieldTarget('5', 0, 1) as FieldTarget, '$A$1:$A$9'), { text: '=$A$1:$A$9', start: 1, end: 10 }, 'the reference stands behind the "="');
});

test('in a formula the reference goes to the caret, where a formula takes one', () => {
  assert.deepEqual(formulaFieldTarget('=', 1, 1), { start: 1, end: 1, lead: '' }, 'right behind "="');
  assert.deepEqual(formulaFieldTarget('=A1+', 4, 4), { start: 4, end: 4, lead: '' }, 'behind an operator');
  assert.deepEqual(formulaFieldTarget('=SUM(A1;)', 8, 8, { decimal: ',' }), { start: 8, end: 8, lead: '' }, 'behind the separator of the locale');
  assert.deepEqual(formulaFieldTarget('=A1+B2', 4, 6), { start: 4, end: 6, lead: '' }, 'in the place of a selected reference');
  assert.equal(formulaFieldTarget('=A1', 3, 3), undefined, 'not behind an operand');
  assert.equal(formulaFieldTarget('="a', 3, 3), undefined, 'not inside a text in quotes');
  assert.equal(formulaFieldTarget('=1,', 3, 3, { decimal: ',' }), undefined, 'not behind the comma of a number');
  assert.equal(formulaFieldTarget('=A1', 0, 0), undefined, 'not before the "="');
});

test('in a number or in the items of a list nothing is pointed at', () => {
  assert.equal(formulaFieldTarget('5', 1, 1), undefined, 'a caret behind a number');
  assert.equal(formulaFieldTarget('a; b', 2, 2), undefined, 'a caret among items');
  assert.equal(formulaFieldTarget('12', 0, 1), undefined, 'a part of a number selected');
  assert.equal(formulaFieldTarget('=A1', 2, 1), undefined, 'a selection that ends before it starts');
  assert.equal(formulaFieldTarget('=A1', 1, 9), undefined, 'a selection beyond the text');
  assert.equal(formulaFieldTarget(`=${'1+'.repeat(20000)}`, 1, 1), undefined, 'a text longer than any formula');
});

test('the reference pointed at last is live while it is still there with the caret behind it', () => {
  const span = { start: 0, end: 5, text: 'A1:B5' };
  assert.equal(fieldSpanLive('A1:B5', 5, 5, span), true, 'a list of ranges has no "=": the span may start at 0');
  assert.equal(fieldSpanLive('A1:B5; D1', 5, 5, span), true, 'more text behind it');
  assert.equal(fieldSpanLive('A1:B5', 4, 4, span), false, 'the caret moved');
  assert.equal(fieldSpanLive('A1:B5', 0, 5, span), false, 'a selection');
  assert.equal(fieldSpanLive('A1:B6', 5, 5, span), false, 'the text was changed');
  assert.equal(fieldSpanLive('A1', 5, 5, span), false, 'the text is shorter than the span');
  assert.equal(fieldSpanLive('=A1:B5', 6, 6, { start: 1, end: 6, text: 'A1:B5' }), true, 'in a formula');
  assert.equal(fieldSpanLive('A1:B5', 5, 5, { start: 0, end: 4, text: 'A1:B5' }), false, 'a span that does not fit its text');
});

test('a source or a formula that was pointed at is read by the dialogs in both locales', () => {
  for (const [name, locale, strings] of [['tr', LOCALE_TR, STRINGS_TR], ['en', LOCALE_EN, STRINGS_EN]] as const) {
    const source = write('', formulaFieldTarget('', 0, 0) as FieldTarget, pointReference(range(0, 0, 8, 0), { absolute: true })).text;
    assert.equal(source, '=$A$1:$A$9', `the source of a list in ${name}`);
    const list = validationFromFields({ ...validationFields(undefined, locale, false), type: 'list', value1: source }, locale, strings.rules);
    assert.equal(list.formula1, '$A$1:$A$9', `the rule reads the cells in ${name}`);
    const away = write('a', formulaFieldTarget('a', 0, 1) as FieldTarget, pointReference(range(0, 0, 2, 0), { absolute: true, sheet: 'Satış 2026' })).text;
    const other = validationFromFields({ ...validationFields(undefined, locale, false), type: 'list', value1: away }, locale, strings.rules);
    assert.equal(other.formula1, "'Satış 2026'!$A$1:$A$3", `the cells of another sheet in ${name}`);
    const typed = name === 'tr' ? '=TOPLA(1;' : '=SUM(1,';
    const sum = write(typed, formulaFieldTarget(typed, typed.length, typed.length, { decimal: locale.decimal }) as FieldTarget, pointReference(range(1, 1, 3, 2))).text + ')>0';
    const rule = conditionalFromFields({ ...conditionalFields(undefined, locale, false, 'expression'), value1: sum }, locale, strings.rules);
    assert.deepEqual(rule.formulas, ['SUM(1,B2:C4)>0'], `a reference pointed into a formula typed in ${name}: ${sum}`);
  }
});

test('what the dialog says about another sheet is there in both languages', () => {
  assert.ok(STRINGS_EN.rules.pointOwnSheet.length > 0, 'English');
  assert.ok(STRINGS_TR.rules.pointOwnSheet.length > 0, 'Turkish');
  assert.notEqual(STRINGS_EN.rules.pointOwnSheet, STRINGS_TR.rules.pointOwnSheet, 'the two are not the same text');
});
