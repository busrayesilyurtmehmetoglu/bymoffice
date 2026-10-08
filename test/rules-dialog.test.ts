// What the dialogs of rules make of their fields: values typed the way the locale writes
// them become what a file stores, and the rule of a file is shown the way it would be typed.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { LOCALE_EN, LOCALE_TR } from '../src/format/locale';
import type { ConditionalRule, DataValidation, DifferentialFormat } from '../src/index';
import type { ConditionalRuleInfo } from '../src/ops/rules';
import { FieldError, validationChange, conditionalPatch, conditionalFields, conditionalFromFields, describeRule, parseRangeList, rangeListText, validationFields, validationFromFields, type ConditionalFields, type ValidationFields } from '../src/ui/rules-dialog';
import { STRINGS_EN, STRINGS_TR } from '../src/ui/strings';

const blank = (): ValidationFields => validationFields(undefined, LOCALE_EN, false);
const en = (fields: Partial<ValidationFields>) => validationFromFields({ ...blank(), ...fields }, LOCALE_EN, STRINGS_EN.rules);
const tr = (fields: Partial<ValidationFields>) => validationFromFields({ ...blank(), ...fields }, LOCALE_TR, STRINGS_TR.rules);
const failsOn = (run: () => unknown, field: string, pattern: RegExp) =>
  assert.throws(run, (error: unknown) => error instanceof FieldError && error.field === field && pattern.test(error.message));

test('dialog: a new rule starts as Excel starts it', () => {
  assert.deepEqual(blank(), {
    type: 'none', operator: 'between', value1: '', value2: '', allowBlank: true, inCellDropdown: true,
    showInputMessage: true, promptTitle: '', prompt: '', showErrorMessage: true, errorStyle: 'stop', errorTitle: '', error: '',
  });
  assert.deepEqual(en({}), { type: 'none', allowBlank: true, showInputMessage: true, showErrorMessage: true, errorStyle: 'stop', hideDropDown: false, promptTitle: '', prompt: '', errorTitle: '', error: '' });
});

test('dialog: numbers are typed as the locale writes them and stored as files write them', () => {
  assert.deepEqual([en({ type: 'decimal', value1: '1.5', value2: '2,000.25' }).formula1, en({ type: 'decimal', value1: '1.5', value2: '2,000.25' }).formula2], ['1.5', '2000.25']);
  assert.deepEqual([tr({ type: 'decimal', value1: '1,5', value2: '2.000,25' }).formula1, tr({ type: 'decimal', value1: '1,5', value2: '2.000,25' }).formula2], ['1.5', '2000.25']);
  assert.equal(tr({ type: 'whole', operator: 'greaterThan', value1: ' 12 ' }).formula1, '12');
  failsOn(() => en({ type: 'whole', operator: 'equal', value1: '1.5' }), 'value1', /whole number/);
  failsOn(() => en({ type: 'decimal', operator: 'equal', value1: 'abc' }), 'value1', /number/);
  failsOn(() => en({ type: 'decimal', value1: '1', value2: '' }), 'value2', /required/);
  failsOn(() => tr({ type: 'textLength', operator: 'lessThan', value1: '' }), 'value1', /gerekli/);
  // The second value of an operator that has one value is not asked for, whatever the field holds.
  assert.equal(en({ type: 'whole', operator: 'equal', value1: '3', value2: 'junk' }).formula2, undefined);
});

test('dialog: dates and times are typed as dates and times and stored as their serial numbers', () => {
  assert.equal(en({ type: 'date', operator: 'greaterThan', value1: '12/31/2024' }).formula1, '45657');
  assert.equal(tr({ type: 'date', operator: 'greaterThan', value1: '31.12.2024' }).formula1, '45657');
  assert.equal(en({ type: 'time', operator: 'lessThan', value1: '12:00' }).formula1, '0.5');
  failsOn(() => tr({ type: 'date', operator: 'equal', value1: '12/31/2024' }), 'value1', /tarih/);
  failsOn(() => en({ type: 'time', operator: 'equal', value1: 'noon' }), 'value1', /time/);
  // The 1904 date system counts from another day.
  assert.equal(validationFromFields({ ...blank(), type: 'date', operator: 'equal', value1: '12/31/2024' }, LOCALE_EN, STRINGS_EN.rules, { date1904: true }).formula1, '44195');
});

test('dialog: formulas are typed with "=" and the separators of the locale', () => {
  assert.equal(en({ type: 'whole', operator: 'equal', value1: '=SUM($A$1:$A$3, 1.5)' }).formula1, 'SUM($A$1:$A$3, 1.5)');
  assert.equal(tr({ type: 'whole', operator: 'equal', value1: '=TOPLA($A$1:$A$3; 1,5)' }).formula1, 'SUM($A$1:$A$3, 1.5)');
  assert.equal(tr({ type: 'custom', value1: 'EĞER(A1>1,5;DOĞRU;YANLIŞ)' }).formula1, 'IF(A1>1.5,TRUE,FALSE)');
  assert.equal(en({ type: 'custom', value1: '=A1>B1' }).formula1, 'A1>B1');
  // A comma that is no decimal comma: refused, as it is in a cell.
  failsOn(() => tr({ type: 'custom', value1: '=IF(A1>5,1,0)' }), 'value1', /formül okunamıyor/);
  failsOn(() => en({ type: 'custom', value1: '=SUM(A1' }), 'value1', /cannot be read/);
  failsOn(() => en({ type: 'custom', value1: '' }), 'value1', /required/);
});

test('dialog: a list is its items, separated the way the locale separates, or the cells that hold them', () => {
  assert.equal(en({ type: 'list', value1: 'yes, no ,maybe' }).formula1, '"yes,no,maybe"');
  assert.equal(tr({ type: 'list', value1: 'evet; hayır ;belki' }).formula1, '"evet,hayır,belki"');
  assert.equal(en({ type: 'list', value1: 'say "hi", no' }).formula1, '"say ""hi"",no"');
  assert.equal(en({ type: 'list', value1: '=$A$1:$A$5' }).formula1, '$A$1:$A$5');
  assert.equal(tr({ type: 'list', value1: "='Diğer Sayfa'!$A$1:$A$5" }).formula1, "'Diğer Sayfa'!$A$1:$A$5");
  assert.equal(en({ type: 'list', value1: 'a,b', inCellDropdown: false }).hideDropDown, true);
  // The file separates with commas whatever the locale: an item cannot hold one.
  failsOn(() => tr({ type: 'list', value1: '1,5; 2,5' }), 'value1', /virgül/);
  failsOn(() => en({ type: 'list', value1: ' , ' }), 'value1', /items/);
  failsOn(() => en({ type: 'list', value1: 'x,'.repeat(130) }), 'value1', /255/);
});

test('dialog: messages have the limits Excel has', () => {
  failsOn(() => en({ type: 'whole', operator: 'equal', value1: '1', promptTitle: 'x'.repeat(33) }), 'promptTitle', /32/);
  failsOn(() => en({ type: 'whole', operator: 'equal', value1: '1', prompt: 'x'.repeat(256) }), 'prompt', /255/);
  failsOn(() => en({ type: 'whole', operator: 'equal', value1: '1', errorTitle: 'x'.repeat(33) }), 'errorTitle', /32/);
  failsOn(() => en({ type: 'whole', operator: 'equal', value1: '1', error: 'x'.repeat(256) }), 'error', /255/);
  const rule = en({ type: 'whole', operator: 'equal', value1: '1', promptTitle: 'T', prompt: 'two\nlines', errorStyle: 'warning', errorTitle: 'E', error: 'no' });
  assert.deepEqual([rule.promptTitle, rule.prompt, rule.errorStyle, rule.errorTitle, rule.error], ['T', 'two\nlines', 'warning', 'E', 'no']);
});

test('dialog: the rule of a file is shown the way it would be typed, and typed back it is the same rule', () => {
  const base = { ranges: [], allowBlank: false, hideDropDown: true, showInputMessage: false, showErrorMessage: true, errorStyle: 'information' as const, operator: 'between' };
  const cases: [DataValidation, string, string, string, string][] = [
    [{ ...base, type: 'decimal', formula1: '1.5', formula2: '2000.25' }, '1.5', '2000.25', '1,5', '2000,25'],
    [{ ...base, type: 'date', operator: 'greaterThan', formula1: '45657' }, '12/31/2024', '', '31.12.2024', ''],
    [{ ...base, type: 'time', operator: 'lessThan', formula1: '0.5' }, '12:00:00', '', '12:00:00', ''],
    [{ ...base, type: 'whole', operator: 'equal', formula1: 'SUM($A$1:$A$3,1.5)' }, '=SUM($A$1:$A$3,1.5)', '', '=SUM($A$1:$A$3;1,5)', ''],
    [{ ...base, type: 'list', formula1: '"yes,no,say ""hi"""' }, 'yes, no, say "hi"', '', 'yes; no; say "hi"', ''],
    [{ ...base, type: 'list', formula1: '$A$1:$A$5' }, '=$A$1:$A$5', '', '=$A$1:$A$5', ''],
    [{ ...base, type: 'custom', formula1: 'IF(A1>1.5,TRUE,FALSE)' }, '=IF(A1>1.5,TRUE,FALSE)', '', '=IF(A1>1,5;TRUE;FALSE)', ''],
  ];
  for (const [rule, en1, en2, tr1, tr2] of cases) {
    for (const [locale, strings, v1, v2] of [[LOCALE_EN, STRINGS_EN, en1, en2], [LOCALE_TR, STRINGS_TR, tr1, tr2]] as const) {
      const fields = validationFields(rule, locale, false);
      assert.deepEqual([fields.value1, fields.value2], [v1, v2], `${rule.type} ${rule.formula1}`);
      assert.deepEqual([fields.allowBlank, fields.inCellDropdown, fields.showInputMessage, fields.showErrorMessage, fields.errorStyle], [false, false, false, true, 'information']);
      const back = validationFromFields(fields, locale, strings.rules);
      assert.deepEqual([back.type, back.operator ?? 'between', back.formula1, back.formula2], [rule.type, rule.operator, rule.formula1, rule.formula2], `${rule.type} ${rule.formula1} typed back`);
    }
  }
});

// ---------------------------------------------------------------------------
// Conditional formats

const blankRule = (type = 'cellIs'): ConditionalFields => conditionalFields(undefined, LOCALE_EN, false, type);
const cen = (fields: Partial<ConditionalFields>) => conditionalFromFields({ ...blankRule(fields.type), ...fields }, LOCALE_EN, STRINGS_EN.rules);
const ctr = (fields: Partial<ConditionalFields>) => conditionalFromFields({ ...conditionalFields(undefined, LOCALE_TR, false, fields.type), ...fields }, LOCALE_TR, STRINGS_TR.rules);
const info = (rule: Partial<ConditionalRule> & { type: string }, format?: DifferentialFormat): ConditionalRuleInfo => ({
  index: 0,
  ranges: [{ top: 0, left: 0, bottom: 4, right: 0 }],
  rule: { priority: 1, stopIfTrue: false, formulas: [], percent: false, bottom: false, aboveAverage: true, equalAverage: false, ...rule },
  format,
  editable: true,
  removable: true,
  movable: true,
});

test('conditional dialog: a new rule starts with the format Excel offers first', () => {
  const fields = blankRule();
  assert.deepEqual([fields.type, fields.operator, fields.preset, fields.rank, fields.scale], ['cellIs', 'greaterThan', 'lightRed', '10', '3']);
  assert.deepEqual(cen({ value1: '5' }), { type: 'cellIs', operator: 'greaterThan', formulas: ['5'], stopIfTrue: false, format: { fill: '#FFC7CE', color: '#9C0006' } });
});

test('conditional dialog: values are typed the way the locale writes them; text is text', () => {
  assert.deepEqual(cen({ operator: 'between', value1: '1.5', value2: '2,000' }).formulas, ['1.5', '2000']);
  assert.deepEqual(ctr({ operator: 'between', value1: '1,5', value2: '=TOPLA($A$1:$A$3; 0,5)' }).formulas, ['1.5', 'SUM($A$1:$A$3, 0.5)']);
  assert.deepEqual(cen({ operator: 'equal', value1: 'done' }).formulas, ['"done"']);
  assert.deepEqual(cen({ operator: 'equal', value1: 'say "hi"' }).formulas, ['"say ""hi"""']);
  assert.deepEqual(cen({ operator: 'greaterThan', value1: '12/31/2024' }).formulas, ['45657']);
  assert.deepEqual(ctr({ operator: 'greaterThan', value1: '31.12.2024' }).formulas, ['45657']);
  assert.deepEqual(cen({ operator: 'equal', value1: 'true' }).formulas, ['TRUE']);
  failsOn(() => cen({ operator: 'between', value1: '1', value2: '' }), 'value2', /required/);
  failsOn(() => cen({ operator: 'equal', value1: '' }), 'value1', /required/);
  failsOn(() => ctr({ operator: 'equal', value1: '=IF(A1>5,1,0)' }), 'value1', /formül okunamıyor/);
  assert.deepEqual(cen({ type: 'expression', value1: 'MOD(ROW(),2)=0' }).formulas, ['MOD(ROW(),2)=0']);
  assert.deepEqual(ctr({ type: 'expression', value1: '=EĞER(A1>1,5;DOĞRU;YANLIŞ)' }).formulas, ['IF(A1>1.5,TRUE,FALSE)']);
  failsOn(() => cen({ type: 'expression', value1: '=SUM(A1' }), 'value1', /cannot be read/);
});

test('conditional dialog: every kind of rule says what it needs and nothing else', () => {
  assert.deepEqual(cen({ type: 'containsText', text: 'abc' }), { type: 'containsText', text: 'abc', stopIfTrue: false, format: { fill: '#FFC7CE', color: '#9C0006' } });
  failsOn(() => cen({ type: 'beginsWith', text: '' }), 'text', /required/);
  assert.deepEqual(cen({ type: 'top10', rank: '5', percent: true, bottom: true, preset: 'green' }), { type: 'top10', rank: 5, percent: true, bottom: true, stopIfTrue: false, format: { fill: '#C6EFCE', color: '#006100' } });
  failsOn(() => cen({ type: 'top10', rank: 'many' }), 'rank', /whole number/);
  failsOn(() => cen({ type: 'top10', rank: '101', percent: true }), 'rank', /100/);
  assert.deepEqual(cen({ type: 'aboveAverage', average: 'below', equalAverage: true, preset: 'yellow', stopIfTrue: true }), { type: 'aboveAverage', aboveAverage: false, equalAverage: true, stopIfTrue: true, format: { fill: '#FFEB9C', color: '#9C5700' } });
  assert.deepEqual(cen({ type: 'duplicateValues', preset: 'custom', useColor: true, color: '#112233', useFill: false, bold: true, italic: false, underline: true, strike: false }).format, { color: '#112233', bold: true, underline: true });
  // "No format" is said as that; "the format it has" is not said at all.
  assert.equal(cen({ type: 'uniqueValues', preset: 'none' }).format, null);
  assert.equal(cen({ type: 'uniqueValues', preset: 'keep' }).format, undefined);
  assert.deepEqual(cen({ type: 'colorScale', scale: '2', color1: '#FF0000', color2: '#00FF00', color3: '#0000FF' }), { type: 'colorScale', colors: ['#FF0000', '#0000FF'], stopIfTrue: false });
  assert.deepEqual(cen({ type: 'colorScale', scale: '3', color1: '#FF0000', color2: '#00FF00', color3: '#0000FF' }).colors, ['#FF0000', '#00FF00', '#0000FF']);
  assert.deepEqual(cen({ type: 'dataBar', barColor: '#638EC6' }), { type: 'dataBar', colors: ['#638EC6'], stopIfTrue: false });
  for (const type of ['containsBlanks', 'notContainsBlanks', 'containsErrors', 'notContainsErrors']) assert.equal(cen({ type }).type, type);
});

test('conditional dialog: a rule is shown the way it would be typed, and typed back it is the same rule', () => {
  const red: DifferentialFormat = { font: { color: { rgb: 'FF9C0006' } }, fill: { type: 'pattern', pattern: 'solid', bg: { rgb: 'FFFFC7CE' } } };
  const shown = conditionalFields(info({ type: 'cellIs', operator: 'between', formulas: ['1.5', 'SUM($A$1:$A$3,0.5)'], stopIfTrue: true }, red), LOCALE_TR, false);
  assert.deepEqual([shown.type, shown.operator, shown.value1, shown.value2, shown.stopIfTrue, shown.preset], ['cellIs', 'between', '1,5', '=SUM($A$1:$A$3;0,5)', true, 'lightRed']);
  assert.deepEqual(conditionalFromFields(shown, LOCALE_TR, STRINGS_TR.rules), { type: 'cellIs', operator: 'between', formulas: ['1.5', 'SUM($A$1:$A$3,0.5)'], stopIfTrue: true, format: { fill: '#FFC7CE', color: '#9C0006' } });
  assert.equal(conditionalFields(info({ type: 'cellIs', operator: 'equal', formulas: ['"say ""hi"""'] }), LOCALE_EN, false).value1, 'say "hi"');
  // A format that is none of the three offered ones is "the format the rule has"; the fields of a custom one show what it states.
  const custom = conditionalFields(info({ type: 'duplicateValues' }, { font: { bold: true, underline: 'single', color: { rgb: 'FF112233' } }, fill: { type: 'pattern', pattern: 'solid', bg: { rgb: 'FF445566' } } }), LOCALE_EN, false);
  assert.deepEqual([custom.preset, custom.useColor, custom.color, custom.useFill, custom.fill, custom.bold, custom.italic, custom.underline, custom.strike], ['keep', true, '#112233', true, '#445566', true, false, true, false]);
  assert.equal(conditionalFields(info({ type: 'uniqueValues' }), LOCALE_EN, false).preset, 'none');
  const text = conditionalFields(info({ type: 'endsWith', text: 'x"y', operator: 'endsWith', formulas: ['RIGHT(A1,LEN("x""y"))="x""y"'] }), LOCALE_EN, false);
  assert.deepEqual([text.type, text.text], ['endsWith', 'x"y']);
  const top = conditionalFields(info({ type: 'top10', rank: 3, percent: true, bottom: true }), LOCALE_EN, false);
  assert.deepEqual([top.rank, top.percent, top.bottom], ['3', true, true]);
  const average = conditionalFields(info({ type: 'aboveAverage', aboveAverage: false, equalAverage: true }), LOCALE_EN, false);
  assert.deepEqual([average.average, average.equalAverage], ['below', true]);
  const scale = conditionalFields(info({ type: 'colorScale', colorScale: { values: [{ type: 'min', gte: true }, { type: 'max', gte: true }], colors: [{ rgb: 'FFF8696B' }, { rgb: 'FF63BE7B' }] } }), LOCALE_EN, false);
  assert.deepEqual([scale.scale, scale.color1, scale.color3], ['2', '#F8696B', '#63BE7B']);
});

test('conditional dialog: a rule is named for the list of rules, in both languages', () => {
  const name = (rule: Partial<ConditionalRule> & { type: string }, strings = STRINGS_EN, locale = LOCALE_EN) => describeRule(info(rule).rule, strings.rules, locale, false);
  assert.equal(name({ type: 'cellIs', operator: 'greaterThan', formulas: ['5'] }), 'Cell value greater than 5');
  assert.equal(name({ type: 'cellIs', operator: 'between', formulas: ['1.5', '$B$1'] }, STRINGS_TR, LOCALE_TR), 'Hücre değeri arasında 1,5 ve =$B$1');
  assert.equal(name({ type: 'containsText', text: 'abc' }), 'Text contains "abc"');
  assert.equal(name({ type: 'top10', rank: 5, percent: true, bottom: true }), 'Bottom 5%');
  assert.equal(name({ type: 'top10', rank: 10 }), 'Top 10');
  assert.equal(name({ type: 'aboveAverage', aboveAverage: false }), 'Below average');
  assert.equal(name({ type: 'expression', formulas: ['A1>3'] }), 'Formula: =A1>3');
  assert.equal(name({ type: 'duplicateValues' }), 'Duplicate values');
  assert.equal(name({ type: 'colorScale' }), 'Colour scale');
  assert.equal(name({ type: 'dataBar' }), 'Data bar');
  assert.equal(name({ type: 'iconSet' }), 'Icon set');
  assert.equal(name({ type: 'timePeriod', timePeriod: 'today' }), 'Dates occurring: today');
  assert.equal(name({ type: 'somethingNew' }), 'somethingNew');
  assert.equal(name({ type: 'duplicateValues' }, STRINGS_TR), 'Yinelenen değerler');
});

test('conditional dialog: the cells of a rule are typed as a list of ranges', () => {
  assert.deepEqual(parseRangeList('A1:B10; D1:D5', LOCALE_TR)!.map((r) => [r.top, r.left, r.bottom, r.right]), [[0, 0, 9, 1], [0, 3, 4, 3]]);
  assert.deepEqual(parseRangeList(' $a$1 , c3:b2 ', LOCALE_EN)!.map((r) => [r.top, r.left, r.bottom, r.right]), [[0, 0, 0, 0], [1, 1, 2, 2]]);
  assert.deepEqual(parseRangeList('A:A', LOCALE_EN)!.map((r) => [r.top, r.left, r.bottom, r.right]), [[0, 0, 1048575, 0]]);
  assert.equal(parseRangeList('A1:B', LOCALE_EN), undefined);
  assert.equal(parseRangeList('', LOCALE_EN), undefined);
  assert.equal(parseRangeList('Sheet2!A1', LOCALE_EN), undefined);
  assert.equal(rangeListText([{ top: 0, left: 0, bottom: 9, right: 1 }, { top: 0, left: 3, bottom: 0, right: 3 }], LOCALE_TR), 'A1:B10; D1');
  assert.equal(rangeListText([{ top: 0, left: 0, bottom: 9, right: 1 }, { top: 0, left: 3, bottom: 0, right: 3 }], LOCALE_EN), 'A1:B10, D1');
});

// ---------------------------------------------------------------------------
// Review: what was only looked at is not changed

test('validation dialog: OK without a change is no rule to set; a field that was not touched keeps what the file wrote', () => {
  const rule: DataValidation = { ranges: [], type: 'list', operator: 'between', formula1: '"a, b,c"', allowBlank: true, hideDropDown: false, showInputMessage: false, showErrorMessage: true, errorStyle: 'stop' };
  for (const [locale, strings] of [[LOCALE_EN, STRINGS_EN], [LOCALE_TR, STRINGS_TR]] as const) {
    const initial = validationFields(rule, locale, false);
    assert.equal(validationChange({ ...initial }, { fields: initial, rule }, locale, strings.rules), undefined);
    const titled = validationChange({ ...initial, errorTitle: 'No' }, { fields: initial, rule }, locale, strings.rules)!;
    assert.deepEqual([titled.formula1, titled.errorTitle], ['"a, b,c"', 'No'], 'the list as the file spells it');
    const other = validationChange({ ...initial, value1: 'x' }, { fields: initial, rule }, locale, strings.rules)!;
    assert.equal(other.formula1, '"x"');
  }
  const spaced: DataValidation = { ...rule, type: 'whole', formula1: 'SUM( $A$1:$A$3 )', formula2: '9' };
  const initial = validationFields(spaced, LOCALE_TR, false);
  const changed = validationChange({ ...initial, value2: '10' }, { fields: initial, rule: spaced }, LOCALE_TR, STRINGS_TR.rules)!;
  assert.deepEqual([changed.formula1, changed.formula2], ['SUM( $A$1:$A$3 )', '10']);
  // Another kind of rule reads the fields anew.
  const retyped = validationChange({ ...initial, type: 'decimal' }, { fields: initial, rule: spaced }, LOCALE_TR, STRINGS_TR.rules)!;
  assert.equal(retyped.type, 'decimal');
  // Without a rule to start from there is nothing to keep.
  assert.equal(validationChange(validationFields(undefined, LOCALE_EN, false), { fields: validationFields(undefined, LOCALE_EN, false), rule: undefined }, LOCALE_EN, STRINGS_EN.rules), undefined);
});

test('conditional dialog: OK without a change changes nothing; one field changed is one thing changed', () => {
  const red: DifferentialFormat = { font: { color: { rgb: 'FF9C0006' } }, fill: { type: 'pattern', pattern: 'solid', bg: { rgb: 'FFFFC7CE' } } };
  const rule = info({ type: 'cellIs', operator: 'between', formulas: ['1', 'SUM( $A$1:$A$3 )'] }, red);
  for (const [locale, strings] of [[LOCALE_EN, STRINGS_EN], [LOCALE_TR, STRINGS_TR]] as const) {
    const initial = { fields: conditionalFields(rule, locale, false), ranges: rangeListText(rule.ranges, locale) };
    assert.deepEqual(conditionalPatch({ ...initial.fields }, initial.ranges, initial, rule, locale, strings.rules), {});
    // The other value keeps the spaces the file wrote it with; the format is not mentioned.
    assert.deepEqual(conditionalPatch({ ...initial.fields, value1: '2' }, initial.ranges, initial, rule, locale, strings.rules), { rule: { type: 'cellIs', operator: 'between', formulas: ['2', 'SUM( $A$1:$A$3 )'] } });
    assert.deepEqual(conditionalPatch({ ...initial.fields, stopIfTrue: true }, initial.ranges, initial, rule, locale, strings.rules), { stopIfTrue: true });
    assert.deepEqual(conditionalPatch({ ...initial.fields, preset: 'green' }, initial.ranges, initial, rule, locale, strings.rules), { format: { fill: '#C6EFCE', color: '#006100' } });
    assert.deepEqual(conditionalPatch({ ...initial.fields, preset: 'none' }, initial.ranges, initial, rule, locale, strings.rules), { format: null });
    assert.deepEqual(conditionalPatch({ ...initial.fields }, 'B2:B3', initial, rule, locale, strings.rules), { ranges: [{ top: 1, left: 1, bottom: 2, right: 1 }] });
    failsOn(() => conditionalPatch({ ...initial.fields }, 'nowhere', initial, rule, locale, strings.rules), 'ranges', /./);
    failsOn(() => conditionalPatch({ ...initial.fields }, ' ', initial, rule, locale, strings.rules), 'ranges', /./);
  }
});

test('conditional dialog: a format it cannot show is "the format the rule has", and stays unless another is chosen', () => {
  const border: DifferentialFormat = { border: { left: { style: 'thin' } } };
  const rule = info({ type: 'cellIs', operator: 'equal', formulas: ['5'] }, border);
  const initial = { fields: conditionalFields(rule, LOCALE_EN, false), ranges: 'A1:A5' };
  assert.equal(initial.fields.preset, 'keep');
  // The value changes: nothing is said about the format.
  assert.deepEqual(conditionalPatch({ ...initial.fields, value1: '6' }, 'A1:A5', initial, rule, LOCALE_EN, STRINGS_EN.rules), { rule: { type: 'cellIs', operator: 'equal', formulas: ['6'] } });
  // Custom, with nothing ticked: what the dialog shows of the format, which is what the format has of it.
  assert.deepEqual(conditionalPatch({ ...initial.fields, preset: 'custom' }, 'A1:A5', initial, rule, LOCALE_EN, STRINGS_EN.rules), { format: {} });
  // A format with bold stated as off: shown unticked, and no change when left so.
  const off = info({ type: 'cellIs', operator: 'equal', formulas: ['5'] }, { font: { bold: false, italic: true }, fill: { type: 'pattern', pattern: 'solid', bg: { rgb: 'FFFFC7CE' } } });
  const shown = conditionalFields(off, LOCALE_EN, false);
  assert.deepEqual([shown.preset, shown.bold, shown.italic, shown.useFill, shown.fill, shown.useColor], ['keep', false, true, true, '#FFC7CE', false]);
  assert.deepEqual(conditionalPatch({ ...shown, preset: 'custom', fill: '#C6EFCE' }, 'A1:A5', { fields: shown, ranges: 'A1:A5' }, off, LOCALE_EN, STRINGS_EN.rules), { format: { fill: '#C6EFCE', italic: true } });
  // A new rule has no format to keep.
  assert.equal(conditionalFields(undefined, LOCALE_EN, false).preset, 'lightRed');
  assert.equal(conditionalFields(info({ type: 'uniqueValues' }), LOCALE_EN, false).preset, 'none');
});

test('conditional dialog: a text rule whose formula is not the one Excel writes is shown as the formula it is', () => {
  const standard = info({ type: 'containsText', operator: 'containsText', text: 'ab', formulas: ['NOT(ISERROR(SEARCH("ab",$A1)))'] });
  assert.deepEqual([conditionalFields(standard, LOCALE_EN, false).type, conditionalFields(standard, LOCALE_EN, false).text], ['containsText', 'ab']);
  const own = info({ type: 'beginsWith', operator: 'beginsWith', text: 'q', formulas: ['LEFT(A1,1)="Q"'] });
  const fields = conditionalFields(own, LOCALE_EN, false);
  assert.deepEqual([fields.type, fields.value1], ['expression', '=LEFT(A1,1)="Q"']);
  const initial = { fields, ranges: 'A1:A5' };
  assert.deepEqual(conditionalPatch({ ...fields }, 'A1:A5', initial, own, LOCALE_EN, STRINGS_EN.rules), {});
  assert.deepEqual(conditionalPatch({ ...fields, value1: '=LEFT(A1,1)="Z"' }, 'A1:A5', initial, own, LOCALE_EN, STRINGS_EN.rules), { rule: { type: 'expression', formulas: ['LEFT(A1,1)="Z"'] } });
});
