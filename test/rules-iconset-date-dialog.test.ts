// The dialog of a conditional rule for icon sets and "dates occurring" rules: what its
// fields stand for, what a rule of a file is shown as, what a dialog that was only looked
// at changes (nothing), and how the sets and periods are named in both languages.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { ICON_SETS } from '../src/format/conditional';
import { LOCALE_EN, LOCALE_TR } from '../src/format/locale';
import { parseRange } from '../src/index';
import type { ConditionalRule, DifferentialFormat } from '../src/index';
import { ICON_SET_NAMES, TIME_PERIODS, conditionalRules, updateConditionalRule, type ConditionalRuleInfo } from '../src/ops/rules';
import { FieldError, conditionalFields, conditionalFromFields, conditionalPatch, describeRule, hasFormat, iconsOfSet, offeredIconSets, rangeListText, type ConditionalFields } from '../src/ui/rules-dialog';
import { STRINGS_EN, STRINGS_TR } from '../src/ui/strings';
import { STYLES, cf } from './conditional-fixtures';
import { buildXlsx, worksheetXml } from './fixtures';
import { load, partText, save } from './rules-fixtures';

const blank = (type: string): ConditionalFields => conditionalFields(undefined, LOCALE_EN, false, type);
const cen = (fields: Partial<ConditionalFields> & { type: string }) => conditionalFromFields({ ...blank(fields.type), ...fields }, LOCALE_EN, STRINGS_EN.rules);
const info = (rule: Partial<ConditionalRule> & { type: string }, format?: DifferentialFormat): ConditionalRuleInfo => ({
  index: 0,
  ranges: [{ top: 0, left: 0, bottom: 4, right: 0 }],
  rule: { priority: 1, stopIfTrue: false, formulas: [], percent: false, bottom: false, aboveAverage: true, equalAverage: false, ...rule },
  format,
  editable: true,
  removable: true,
  movable: true,
});
const RED: DifferentialFormat = { font: { color: { rgb: 'FF9C0006' } }, fill: { type: 'pattern', pattern: 'solid', bg: { rgb: 'FFFFC7CE' } } };
const VALUES = [
  { type: 'percent', value: '0', gte: true },
  { type: 'num', value: '2', gte: false },
  { type: 'formula', value: '$H$1', gte: true },
];

test('icon set dialog: a new rule starts with the set Excel starts with, and says its set, its direction and whether the value shows - no format', () => {
  const fields = blank('iconSet');
  assert.deepEqual([fields.type, fields.iconSet, fields.reverse, fields.iconOnly], ['iconSet', '3TrafficLights1', false, false], 'the fields of a new rule');
  assert.equal(hasFormat('iconSet'), false, 'an icon set has no format to choose');
  assert.equal(hasFormat('timePeriod'), true, 'a period has');
  assert.deepEqual(cen({ type: 'iconSet' }), { type: 'iconSet', iconSet: '3TrafficLights1', reverse: false, showValue: true, stopIfTrue: false }, 'a new rule as it is');
  assert.deepEqual(cen({ type: 'iconSet', iconSet: '5Quarters', reverse: true, iconOnly: true, stopIfTrue: true, preset: 'green' }), { type: 'iconSet', iconSet: '5Quarters', reverse: true, showValue: false, stopIfTrue: true }, 'every field set');
});

test('period dialog: a new rule starts with today and the format Excel offers first; a period is required', () => {
  const fields = blank('timePeriod');
  assert.deepEqual([fields.type, fields.period, fields.preset], ['timePeriod', 'today', 'lightRed'], 'the fields of a new rule');
  assert.deepEqual(cen({ type: 'timePeriod' }), { type: 'timePeriod', timePeriod: 'today', stopIfTrue: false, format: { fill: '#FFC7CE', color: '#9C0006' } }, 'a new rule as it is');
  for (const period of TIME_PERIODS) assert.equal(cen({ type: 'timePeriod', period, preset: 'none' }).timePeriod, period, period);
  assert.equal(cen({ type: 'timePeriod', period: 'lastWeek', preset: 'none' }).format, null, 'no format');
  assert.throws(() => cen({ type: 'timePeriod', period: '' }), (error: unknown) => error instanceof FieldError && error.field === 'period' && /required/.test(error.message), 'no period');
});

test('icon set dialog: a rule of a file is shown as it is; looked at and confirmed it is not changed, and one field changed is one thing changed', () => {
  const rule = info({ type: 'iconSet', iconSet: { set: '3Flags', values: VALUES, showValue: false, reverse: true } });
  for (const [locale, strings] of [[LOCALE_EN, STRINGS_EN], [LOCALE_TR, STRINGS_TR]] as const) {
    const initial = { fields: conditionalFields(rule, locale, false), ranges: rangeListText(rule.ranges, locale) };
    assert.deepEqual([initial.fields.type, initial.fields.iconSet, initial.fields.reverse, initial.fields.iconOnly, initial.fields.preset], ['iconSet', '3Flags', true, true, 'none'], 'the fields');
    assert.deepEqual(conditionalPatch({ ...initial.fields }, initial.ranges, initial, rule, locale, strings.rules), {}, 'nothing touched');
    assert.deepEqual(conditionalPatch({ ...initial.fields, reverse: false }, initial.ranges, initial, rule, locale, strings.rules), { rule: { type: 'iconSet', iconSet: '3Flags', reverse: false, showValue: false } }, 'the direction');
    assert.deepEqual(conditionalPatch({ ...initial.fields, iconSet: '4Rating' }, initial.ranges, initial, rule, locale, strings.rules), { rule: { type: 'iconSet', iconSet: '4Rating', reverse: true, showValue: false } }, 'the set');
    assert.deepEqual(conditionalPatch({ ...initial.fields, stopIfTrue: true }, initial.ranges, initial, rule, locale, strings.rules), { stopIfTrue: true }, 'stop if true');
    // The fields of a format are not those of an icon set: whatever they hold, nothing is said about a format.
    assert.deepEqual(conditionalPatch({ ...initial.fields, preset: 'green' }, initial.ranges, initial, rule, locale, strings.rules), {}, 'a format');
  }
});

test('period dialog: a rule of a file is shown as it is; looked at and confirmed it is not changed, and one field changed is one thing changed', () => {
  const rule = info({ type: 'timePeriod', timePeriod: 'lastMonth', formulas: ['AND(MONTH($A1)=MONTH(TODAY())-1, TRUE)'], dxfId: 0 }, RED);
  for (const [locale, strings] of [[LOCALE_EN, STRINGS_EN], [LOCALE_TR, STRINGS_TR]] as const) {
    const initial = { fields: conditionalFields(rule, locale, false), ranges: rangeListText(rule.ranges, locale) };
    assert.deepEqual([initial.fields.type, initial.fields.period, initial.fields.preset], ['timePeriod', 'lastMonth', 'lightRed'], 'the fields');
    assert.deepEqual(conditionalPatch({ ...initial.fields }, initial.ranges, initial, rule, locale, strings.rules), {}, 'nothing touched');
    assert.deepEqual(conditionalPatch({ ...initial.fields, period: 'nextWeek' }, initial.ranges, initial, rule, locale, strings.rules), { rule: { type: 'timePeriod', timePeriod: 'nextWeek' } }, 'the period');
    assert.deepEqual(conditionalPatch({ ...initial.fields, preset: 'yellow' }, initial.ranges, initial, rule, locale, strings.rules), { format: { fill: '#FFEB9C', color: '#9C5700' } }, 'the format');
    assert.deepEqual(conditionalPatch({ ...initial.fields }, 'B2:B3', initial, rule, locale, strings.rules), { ranges: [{ top: 1, left: 1, bottom: 2, right: 1 }] }, 'the cells');
    // Another kind of rule reads its own fields.
    assert.deepEqual(conditionalPatch({ ...initial.fields, type: 'iconSet', iconSet: '3Signs' }, initial.ranges, initial, rule, locale, strings.rules), { rule: { type: 'iconSet', iconSet: '3Signs', reverse: false, showValue: true } }, 'the kind');
  }
});

test('a set or a period of a file that is not offered is shown under its own name and stays while it is not changed', () => {
  const icons = info({ type: 'iconSet', iconSet: { set: '3Unheard', values: VALUES, showValue: true, reverse: false } });
  const shown = conditionalFields(icons, LOCALE_EN, false);
  assert.equal(shown.iconSet, '3Unheard', 'the set of the file');
  assert.deepEqual(conditionalPatch({ ...shown, iconOnly: true }, 'A1:A5', { fields: shown, ranges: 'A1:A5' }, icons, LOCALE_EN, STRINGS_EN.rules), { rule: { type: 'iconSet', iconSet: '3Unheard', reverse: false, showValue: false } }, 'the name goes back as it came');
  const period = info({ type: 'timePeriod', timePeriod: 'someday', formulas: ['A1=TODAY()'] }, RED);
  assert.equal(conditionalFields(period, LOCALE_EN, false).period, 'someday', 'the period of the file');
  assert.equal(conditionalFields(info({ type: 'timePeriod', formulas: [] }), LOCALE_EN, false).period, '', 'a rule without a period');
});

test('the dialog of a rule and the operations agree: what the dialog changes of a rule of a file is all that changes in the file', async () => {
  const icons = '<cfRule type="iconSet" priority="1"><iconSet iconSet="3Flags"><cfvo type="percent" val="0"/><cfvo type="num" val="2" gte="0"/><cfvo type="formula" val="$H$1"/></iconSet></cfRule>';
  const period = '<cfRule type="timePeriod" dxfId="0" priority="2" timePeriod="lastWeek"><formula>WEEKDAY( $A1 )&gt;0</formula></cfRule>';
  const doc = await load(await buildXlsx({ sheets: [{ name: 'Data', xml: worksheetXml(`<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>${cf('A1:A5', icons)}${cf('B1:B5', period)}`) }], styles: STYLES }));
  const reading = { date1904: false };
  for (const index of [0, 1]) {
    const listed = conditionalRules(doc.wb, 0)[index];
    const initial = { fields: conditionalFields(listed, LOCALE_TR, false), ranges: rangeListText(listed.ranges, LOCALE_TR) };
    const untouched = conditionalPatch({ ...initial.fields }, initial.ranges, initial, listed, LOCALE_TR, STRINGS_TR.rules, reading);
    assert.deepEqual(untouched, {}, `rule ${index}: nothing touched is an empty patch`);
    // Stated again in full, the way the dialog states a rule, it is still the rule it was.
    assert.deepEqual(updateConditionalRule(doc.wb, 0, index, { rule: conditionalFromFields(initial.fields, LOCALE_TR, STRINGS_TR.rules, reading) }), [], `rule ${index}: stated again as it is`);
  }
  const first = conditionalRules(doc.wb, 0)[0];
  const fields = conditionalFields(first, LOCALE_EN, false);
  doc.history.apply(updateConditionalRule(doc.wb, 0, 0, conditionalPatch({ ...fields, iconOnly: true }, 'A1:A5', { fields, ranges: 'A1:A5' }, first, LOCALE_EN, STRINGS_EN.rules, reading)), undefined, undefined);
  const xml = await partText(await save(doc), 'xl/worksheets/sheet1.xml');
  assert.ok(xml.includes('<iconSet iconSet="3Flags" showValue="0"><cfvo type="percent" val="0"/><cfvo type="num" val="2" gte="0"/><cfvo type="formula" val="$H$1"/></iconSet>'), xml);
  assert.ok(xml.includes(cf('B1:B5', period)), 'the rule that was not touched keeps its source text');
  assert.deepEqual(conditionalRules(doc.wb, 0, parseRange('B1')!)[0].rule.formulas, ['WEEKDAY( $A1 )>0'], 'the formula of the period as the file spells it');
});

test('sets and periods are named for the list of rules and for the dialog, in both languages', () => {
  const name = (rule: Partial<ConditionalRule> & { type: string }, strings = STRINGS_EN, locale = LOCALE_EN) => describeRule(info(rule).rule, strings.rules, locale, false);
  const set = (id: string) => ({ set: id, values: VALUES, showValue: true, reverse: false });
  assert.equal(name({ type: 'iconSet', iconSet: set('3Arrows') }), 'Icon set: 3 arrows (coloured)');
  assert.equal(name({ type: 'iconSet', iconSet: set('5Quarters') }, STRINGS_TR, LOCALE_TR), 'Simge kümesi: 5 çeyrek');
  assert.equal(name({ type: 'iconSet', iconSet: set('3Stars') }), 'Icon set: 3 stars', 'a set of the 2010 extension has a name too');
  assert.equal(name({ type: 'iconSet', iconSet: set('3Unheard') }), 'Icon set: 3Unheard', 'a set nobody knows is named as the file names it');
  assert.equal(name({ type: 'iconSet' }), 'Icon set', 'a rule without its element');
  assert.equal(name({ type: 'timePeriod', timePeriod: 'last7Days' }), 'Dates occurring: in the last 7 days');
  assert.equal(name({ type: 'timePeriod', timePeriod: 'nextMonth' }, STRINGS_TR, LOCALE_TR), 'Tarih: gelecek ay');
  for (const strings of [STRINGS_EN, STRINGS_TR]) {
    for (const id of ICON_SETS.keys()) assert.ok(typeof strings.rules.iconSets[id] === 'string' && strings.rules.iconSets[id] !== '', `a name for the set ${id}`);
    for (const period of TIME_PERIODS) assert.ok(typeof strings.rules.periods[period] === 'string' && strings.rules.periods[period] !== '', `a name for the period ${period}`);
    for (const type of ['iconSet', 'timePeriod']) assert.ok(strings.rules.conditionalTypes[type], `a name for the kind ${type}`);
    for (const key of ['iconStyle', 'iconPreview', 'reverseIcons', 'iconOnly', 'period'] as const) assert.ok(strings.rules[key] !== '', `the text ${key}`);
    assert.equal(new Set(ICON_SET_NAMES.map((id) => strings.rules.iconSets[id])).size, ICON_SET_NAMES.length, 'no two sets that are offered share a name');
  }
});

test('the sets the dialog offers are those that are made, and the picture of a set shows its icons from the lowest values to the highest', () => {
  assert.deepEqual(offeredIconSets(), [...ICON_SET_NAMES], 'the sets offered');
  for (const id of ICON_SET_NAMES) {
    const icons = iconsOfSet(id, false);
    assert.equal(icons.length, Number(id.charAt(0)), `${id}: as many icons as its name says`);
    assert.deepEqual(iconsOfSet(id, true), icons.slice().reverse(), `${id}: the other way round`);
    assert.deepEqual(icons, ICON_SETS.get(id), `${id}: the icons the grid draws`);
  }
  // A set that is not known is shown as the grid shows it.
  assert.deepEqual(iconsOfSet('4Unheard', false), ICON_SETS.get('4TrafficLights'), 'four icons');
  assert.deepEqual(iconsOfSet('5Unheard', false), ICON_SETS.get('5Quarters'), 'five icons');
  assert.deepEqual(iconsOfSet('whatever', false), ICON_SETS.get('3TrafficLights1'), 'anything else');
});
