// Icon sets and "dates occurring" rules that are made and changed in the editor: what the
// operations build, what the writer and the builder of a new file make of it, which cells
// get which icon or format, what a change leaves alone, and how the formula of a period
// follows its cells through rows, columns and the clipboard.

import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { CalcEngine } from '../src/calc/engine';
import { ConditionalFormats, ICON_SETS } from '../src/format/conditional';
import { buildXlsx as buildFromModel } from '../src/formats/xlsx/build';
import { open, parseRange } from '../src/index';
import type { CellRange, Workbook } from '../src/index';
import { copyRange, cutRange, pasteAt } from '../src/ops/commands';
import type { Op } from '../src/ops/ops';
import { CONDITIONAL_TYPES, ICON_SET_NAMES, TIME_PERIODS, addConditionalRule, conditionalRules, removeConditionalRule, updateConditionalRule, type ConditionalSettings } from '../src/ops/rules';
import { DXFS, EXT_ONLY, STYLES, cf, conditionalBook, conditionalContent, conditionalNow as now } from './conditional-fixtures';
import { buildXlsx, worksheetXml } from './fixtures';
import { checkPackage, checkSheetOrder, load as loadBytes, partText, reopen, save, sameBytes, type Loaded } from './rules-fixtures';
import { canonical } from './workbook-content';

const range = (ref: string): CellRange => parseRange(ref)!;
const SHEET1 = 'xl/worksheets/sheet1.xml';
const STYLES_PART = 'xl/styles.xml';
const apply = (doc: Loaded, ops: Op[]) => doc.history.apply(ops, undefined, undefined);
const RED = { fill: '#FFC7CE', color: '#9C0006' };
const escaped = (text: string) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const sameContent = async (bytes: Uint8Array, doc: Loaded, label: string) => assert.equal(canonical(conditionalContent(await reopen(bytes))), canonical(conditionalContent(doc.wb)), label);

/**
 * The sheet and the style sheet are written as the file has them. (After rows, columns or cells
 * were changed and the change undone the workbook part asks for a calculation on load, with
 * or without rules: the archive as a whole is not compared then.)
 */
async function sameParts(doc: Loaded, label: string): Promise<void> {
  const saved = await save(doc);
  for (const part of [SHEET1, STYLES_PART]) assert.equal(await partText(saved, part), await partText(doc.bytes, part), `${label}: ${part}`);
}

/** A sheet "Data" with numbers in A1..A`count` (1, 2, 3, ...), whatever else `inner` says, and the three formats of the fixtures. */
const numbers = (count: number, inner = '', value = (i: number) => i + 1) => {
  let rows = '';
  for (let i = 0; i < count; i++) rows += `<row r="${i + 1}"><c r="A${i + 1}"><v>${value(i)}</v></c></row>`;
  return buildXlsx({ sheets: [{ name: 'Data', xml: worksheetXml(`<sheetData>${rows}</sheetData>${inner}`) }], styles: STYLES });
};

/** The thresholds a new set of `count` icons is written with. */
const THRESHOLDS: Record<number, number[]> = { 3: [0, 33, 67], 4: [0, 25, 50, 75], 5: [0, 20, 40, 60, 80] };
const cfvos = (count: number) => THRESHOLDS[count].map((value) => `<cfvo type="percent" val="${value}"/>`).join('');

/** The formula each period is written with, for the cell `c`. */
const FORMULAS: Record<string, (c: string) => string> = {
  today: (c) => `FLOOR(${c},1)=TODAY()`,
  yesterday: (c) => `FLOOR(${c},1)=TODAY()-1`,
  tomorrow: (c) => `FLOOR(${c},1)=TODAY()+1`,
  last7Days: (c) => `AND(TODAY()-FLOOR(${c},1)<=6,FLOOR(${c},1)<=TODAY())`,
  lastWeek: (c) => `AND(TODAY()-ROUNDDOWN(${c},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(${c},0)<(WEEKDAY(TODAY())+7))`,
  thisWeek: (c) => `AND(TODAY()-ROUNDDOWN(${c},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(${c},0)-TODAY()<=7-WEEKDAY(TODAY()))`,
  nextWeek: (c) => `AND(ROUNDDOWN(${c},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(${c},0)-TODAY()<(15-WEEKDAY(TODAY())))`,
  lastMonth: (c) => `AND(MONTH(${c})=MONTH(EDATE(TODAY(),0-1)),YEAR(${c})=YEAR(EDATE(TODAY(),0-1)))`,
  thisMonth: (c) => `AND(MONTH(${c})=MONTH(TODAY()),YEAR(${c})=YEAR(TODAY()))`,
  nextMonth: (c) => `AND(MONTH(${c})=MONTH(EDATE(TODAY(),0+1)),YEAR(${c})=YEAR(EDATE(TODAY(),0+1)))`,
};

// ---------------------------------------------------------------------------
// What is offered

test('icon sets and periods: the two kinds are made here; every set that is offered can be drawn, and none of the 2010 extension is offered', () => {
  assert.ok((CONDITIONAL_TYPES as readonly string[]).includes('iconSet'), 'iconSet is a kind that is made');
  assert.ok((CONDITIONAL_TYPES as readonly string[]).includes('timePeriod'), 'timePeriod is a kind that is made');
  assert.equal(ICON_SET_NAMES.length, 17, 'the sets of ST_IconSetType');
  for (const name of ICON_SET_NAMES) {
    const icons = ICON_SETS.get(name);
    assert.ok(icons, `${name} can be drawn`);
    assert.equal(icons!.length, Number(name.charAt(0)), `${name} has as many icons as its name says`);
  }
  for (const name of ['3Stars', '3Triangles', '5Boxes', 'NoIcons']) assert.ok(!(ICON_SET_NAMES as readonly string[]).includes(name), `${name} is not offered`);
  assert.deepEqual([...TIME_PERIODS].sort(), Object.keys(FORMULAS).sort(), 'the periods of ST_TimePeriod');
});

// ---------------------------------------------------------------------------
// Made, written, read back

test('icon sets: every set is written into the main part with the thresholds Excel starts with, read back as the model has it, and undone to the bytes of the file', async () => {
  for (const name of ICON_SET_NAMES) {
    const doc = await loadBytes(await conditionalBook({ bare: true }));
    const original = await partText(doc.bytes, SHEET1);
    const ops = addConditionalRule(doc.wb, 0, range('G1:G10'), { type: 'iconSet', iconSet: name });
    assert.equal(ops.length, 1, `${name}: one operation`);
    apply(doc, ops);
    const count = Number(name.charAt(0));
    const rule = conditionalRules(doc.wb, 0)[0];
    assert.deepEqual(rule.rule.iconSet, { set: name, values: THRESHOLDS[count].map((value) => ({ type: 'percent', value: String(value), gte: true })), showValue: true, reverse: false }, `${name}: the model`);
    assert.equal(rule.format, undefined, `${name}: an icon set gives no format`);
    assert.deepEqual([rule.editable, rule.removable, rule.movable], [true, true, true], `${name}: what can be done to it`);
    const saved = await save(doc);
    await checkPackage(saved);
    const xml = await partText(saved, SHEET1);
    checkSheetOrder(xml, name);
    const added = `<conditionalFormatting sqref="G1:G10"><cfRule type="iconSet" priority="1"><iconSet iconSet="${name}">${cfvos(count)}</iconSet></cfRule></conditionalFormatting>`;
    assert.equal(xml, original.replace('<hyperlinks>', `${added}<hyperlinks>`), `${name}: the sheet part`);
    assert.ok(!xml.includes('extLst'), `${name}: nothing goes into an extension list`);
    assert.equal(await partText(saved, STYLES_PART), await partText(doc.bytes, STYLES_PART), `${name}: the style sheet stays`);
    await sameContent(saved, doc, `${name}: read back`);
    doc.history.undo();
    sameBytes(await save(doc), doc.bytes, `${name}: after undo`);
    doc.history.redo();
    sameBytes(await save(doc), saved, `${name}: after redo`);
  }
});

test('icon sets: the other way round and without the value; what is not stated is the default', async () => {
  const doc = await loadBytes(await conditionalBook({ bare: true }));
  apply(doc, addConditionalRule(doc.wb, 0, range('G1:G3'), { type: 'iconSet', iconSet: '4Rating', reverse: true, showValue: false, stopIfTrue: true }));
  apply(doc, addConditionalRule(doc.wb, 0, range('F1:F3'), { type: 'iconSet' }));
  const saved = await save(doc);
  await checkPackage(saved);
  const xml = await partText(saved, SHEET1);
  assert.ok(xml.includes(`<conditionalFormatting sqref="G1:G3"><cfRule type="iconSet" priority="2" stopIfTrue="1"><iconSet iconSet="4Rating" showValue="0" reverse="1">${cfvos(4)}</iconSet></cfRule></conditionalFormatting>`), xml);
  assert.ok(xml.includes(`<conditionalFormatting sqref="F1:F3"><cfRule type="iconSet" priority="1"><iconSet iconSet="3TrafficLights1">${cfvos(3)}</iconSet></cfRule></conditionalFormatting>`), xml);
  await sameContent(saved, doc, 'read back');
  // A format is nothing an icon set has: one that is stated is not kept.
  apply(doc, addConditionalRule(doc.wb, 0, range('E1'), { type: 'iconSet', iconSet: '3Flags', format: RED }));
  assert.equal(conditionalRules(doc.wb, 0)[0].format, undefined, 'no format');
  assert.equal(conditionalRules(doc.wb, 0)[0].rule.dxf, undefined, 'no format stated in the rule');
});

test('time periods: every period is written with its attribute and the formula for its first cell, read back as the model has it, and undone to the bytes of the file', async () => {
  for (const period of TIME_PERIODS) {
    const doc = await loadBytes(await conditionalBook({ bare: true }));
    const original = await partText(doc.bytes, SHEET1);
    apply(doc, addConditionalRule(doc.wb, 0, range('G2:H9'), { type: 'timePeriod', timePeriod: period, format: RED }));
    const info = conditionalRules(doc.wb, 0)[0];
    assert.deepEqual([info.rule.type, info.rule.timePeriod, info.rule.formulas], ['timePeriod', period, [FORMULAS[period]('G2')]], `${period}: the model`);
    assert.deepEqual([info.editable, info.removable, info.movable], [true, true, true], `${period}: what can be done to it`);
    const saved = await save(doc);
    await checkPackage(saved);
    const xml = await partText(saved, SHEET1);
    checkSheetOrder(xml, period);
    const added = `<conditionalFormatting sqref="G2:H9"><cfRule type="timePeriod" dxfId="3" priority="1" timePeriod="${period}"><formula>${escaped(FORMULAS[period]('G2'))}</formula></cfRule></conditionalFormatting>`;
    assert.equal(xml, original.replace('<hyperlinks>', `${added}<hyperlinks>`), `${period}: the sheet part`);
    assert.ok((await partText(saved, STYLES_PART)).includes(`<dxfs count="4">${DXFS}<dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>`), `${period}: the format is added to the style sheet`);
    await sameContent(saved, doc, `${period}: read back`);
    doc.history.undo();
    sameBytes(await save(doc), doc.bytes, `${period}: after undo`);
    doc.history.redo();
    sameBytes(await save(doc), saved, `${period}: after redo`);
  }
});

test('what an icon set or a period cannot say is refused, and the sheet stays as it is', async () => {
  const doc = await loadBytes(await conditionalBook({ bare: true }));
  const add = (settings: ConditionalSettings) => addConditionalRule(doc.wb, 0, range('A1'), settings);
  for (const name of ['3Stars', '3Triangles', '5Boxes', 'NoIcons', '3arrows', '']) assert.throws(() => add({ type: 'iconSet', iconSet: name as '3Arrows' }), /icon set is not known/, `the set "${name}"`);
  assert.throws(() => add({ type: 'timePeriod' }), /needs a period/, 'no period');
  for (const period of ['someday', 'Today', '']) assert.throws(() => add({ type: 'timePeriod', timePeriod: period as 'today' }), /period is not known/, `the period "${period}"`);
  assert.equal(doc.wb.sheets[0].conditionalFormats, undefined, 'no rule was added');
});

// ---------------------------------------------------------------------------
// Which cell gets what

const iconsOf = (wb: Workbook, count: number) => {
  const formats = new ConditionalFormats(wb, wb.sheets[0]);
  const out: (number | undefined)[] = [];
  for (let row = 0; row < count; row++) {
    const found = formats.resolve(row, 0);
    out.push(found && found.icon ? found.icon.index : undefined);
  }
  return out;
};

test('icon sets: a new rule gives each cell the icon of its share of the range, as a rule of a file with the same thresholds does', async () => {
  // 1..9: the thirds end at 3.64 and 6.36, the quarters at 3, 5 and 7, the fifths at 2.6, 4.2, 5.8 and 7.4; a threshold belongs to the icon above it.
  const expected: Record<number, number[]> = { 3: [0, 0, 0, 1, 1, 1, 2, 2, 2], 4: [0, 0, 1, 1, 2, 2, 3, 3, 3], 5: [0, 0, 1, 1, 2, 3, 3, 4, 4] };
  for (const name of ICON_SET_NAMES) {
    const count = Number(name.charAt(0));
    const doc = await loadBytes(await numbers(9));
    apply(doc, addConditionalRule(doc.wb, 0, range('A1:A9'), { type: 'iconSet', iconSet: name }));
    assert.deepEqual(iconsOf(doc.wb, 10), [...expected[count], undefined], `${name}: the model`);
    const formats = new ConditionalFormats(doc.wb, doc.wb.sheets[0]);
    assert.deepEqual([formats.resolve(0, 0)!.icon!.set, formats.resolve(0, 0)!.hideValue === true], [name, false], `${name}: the set that is drawn, beside the value`);
    // The same rule as a file states it.
    const file = await reopen(await numbers(9, cf('A1:A9', `<cfRule type="iconSet" priority="1"><iconSet iconSet="${name}">${cfvos(count)}</iconSet></cfRule>`)));
    assert.deepEqual(iconsOf(file, 10), iconsOf(doc.wb, 10), `${name}: as a file with these thresholds`);
    assert.deepEqual(iconsOf(await reopen(await save(doc)), 10), iconsOf(doc.wb, 10), `${name}: saved and read back`);
    // The other way round, and the icon alone.
    apply(doc, updateConditionalRule(doc.wb, 0, 0, { rule: { type: 'iconSet', reverse: true, showValue: false } }));
    assert.deepEqual(iconsOf(doc.wb, 9), expected[count].map((index) => count - 1 - index), `${name}: reversed`);
    assert.equal(new ConditionalFormats(doc.wb, doc.wb.sheets[0]).resolve(4, 0)!.hideValue, true, `${name}: the value is not shown`);
  }
});

/** The rows of column A (dates around `today`) a rule of `period` fills, by the rule and by the formula it is written with. */
async function periodCells(today: number, period: string): Promise<{ byRule: number[]; byFormula: number[]; formula: string }> {
  const SPAN = 75;
  const doc = await loadBytes(await numbers(2 * SPAN + 1, '', (i) => today - SPAN + i + (i % 3) * 0.25));
  apply(doc, addConditionalRule(doc.wb, 0, range(`A1:A${2 * SPAN + 1}`), { type: 'timePeriod', timePeriod: period as 'today', format: RED }));
  const clock = () => today + 0.4;
  const formats = new ConditionalFormats(doc.wb, doc.wb.sheets[0], { now: clock });
  const engine = new CalcEngine(doc.wb, { now: clock });
  const formula = conditionalRules(doc.wb, 0)[0].rule.formulas[0];
  const byRule: number[] = [];
  const byFormula: number[] = [];
  for (let row = 0; row <= 2 * SPAN; row++) {
    const found = formats.resolve(row, 0);
    if (found && found.style && found.style.fill !== undefined) byRule.push(row - SPAN);
    if (engine.evaluateText(formula.replace(/A1/g, `A${row + 1}`), { sheet: 0, row, col: 0 }) === true) byFormula.push(row - SPAN);
  }
  return { byRule, byFormula, formula };
}

const span = (from: number, to: number) => Array.from({ length: to - from + 1 }, (_, i) => from + i);

test('time periods: a new rule fills the days of its period, counted from a given today, at the edges of weeks and months', async () => {
  // Serial 45000 is Wednesday 15 March 2023; weeks run from Sunday to Saturday.
  const wednesday: Record<string, number[]> = {
    today: [0],
    yesterday: [-1],
    tomorrow: [1],
    last7Days: span(-6, 0),
    thisWeek: span(-3, 3),
    lastWeek: span(-10, -4),
    nextWeek: span(4, 10),
    // March has 31 days, February 2023 28, April 30.
    thisMonth: span(-14, 16),
    lastMonth: span(-42, -15),
    nextMonth: span(17, 46),
  };
  for (const period of TIME_PERIODS) assert.deepEqual((await periodCells(45000, period)).byRule, wednesday[period], `${period} on a Wednesday in the middle of March`);
  // Sunday 12 March 2023 (the first day of its week) and Saturday 18 (the last).
  assert.deepEqual((await periodCells(44997, 'thisWeek')).byRule, span(0, 6), 'this week on its Sunday');
  assert.deepEqual((await periodCells(44997, 'lastWeek')).byRule, span(-7, -1), 'last week on a Sunday');
  assert.deepEqual((await periodCells(45003, 'thisWeek')).byRule, span(-6, 0), 'this week on its Saturday');
  assert.deepEqual((await periodCells(45003, 'nextWeek')).byRule, span(1, 7), 'next week on a Saturday');
  // 1 January 2024 (45292): last month is December of the year before; 31 December 2023 (45291): next month is January.
  assert.deepEqual((await periodCells(45292, 'lastMonth')).byRule, span(-31, -1), 'last month in January');
  assert.deepEqual((await periodCells(45292, 'thisMonth')).byRule, span(0, 30), 'this month on its first day');
  assert.deepEqual((await periodCells(45291, 'nextMonth')).byRule, span(1, 31), 'next month in December');
  // 29 February 2024 (45351).
  assert.deepEqual((await periodCells(45351, 'thisMonth')).byRule, span(-28, 0), 'this month on a leap day');
  assert.deepEqual((await periodCells(45351, 'nextMonth')).byRule, span(1, 31), 'next month from a leap day');
});

test('time periods: the formula a rule is written with says what the rule does, on every kind of day', async () => {
  // What a reader goes by that does not know the attribute: calculated with the formula engine, it names the same days.
  const days = [45000, 44997, 45003, 45291, 45292, 45322, 45351, 45352, 45016, 45017, 44985, 44986];
  for (const period of TIME_PERIODS) {
    for (const today of days) {
      const found = await periodCells(today, period);
      assert.equal(found.formula, FORMULAS[period]('A1'), `${period}: the formula`);
      assert.deepEqual(found.byFormula, found.byRule, `${period} with today ${today}: the formula and the rule name the same days`);
    }
  }
});

// ---------------------------------------------------------------------------
// A rule of a file is changed

const XR = ' xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"';
const FILE_ICONS = '<cfRule type="iconSet" priority="1" xr:uid="{I1}"><iconSet iconSet="3Flags" percent="0"><cfvo type="percent" val="0"/><cfvo type="num" val="2" gte="0"/><cfvo type="formula" val="$H$1"/></iconSet><!-- kept --></cfRule>';
const WEEK = 'AND(TODAY()-ROUNDDOWN($D1,0)&gt;=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN($D1,0)&lt;(WEEKDAY(TODAY())+7))';
const FILE_PERIOD = `<cfRule type="timePeriod" dxfId="2"   priority="2" timePeriod="lastWeek"><formula>${WEEK}</formula></cfRule>`;
const FILE_BLOCK = cf('D1:D5', `\n ${FILE_ICONS}\n ${FILE_PERIOD}\n`, XR);
const OTHER_BLOCK = cf('E1:E5', '<cfRule type="cellIs" dxfId="0" priority="3" operator="greaterThan"><formula>1</formula></cfRule>');
const fileBook = () => numbers(5, FILE_BLOCK + OTHER_BLOCK);
const KEPT_VALUES = [
  { type: 'percent', value: '0', gte: true },
  { type: 'num', value: '2', gte: false },
  { type: 'formula', value: '$H$1', gte: true },
];

test('a rule of a file: stated again the way it is, it is not changed at all', async () => {
  const doc = await loadBytes(await fileBook());
  const [icons, period] = conditionalRules(doc.wb, 0);
  assert.deepEqual([icons.rule.type, icons.editable, period.rule.type, period.editable], ['iconSet', true, 'timePeriod', true], 'both can be changed');
  assert.deepEqual(updateConditionalRule(doc.wb, 0, 0, { rule: { type: 'iconSet' } }), [], 'the kind alone');
  assert.deepEqual(updateConditionalRule(doc.wb, 0, 0, { rule: { type: 'iconSet', iconSet: '3Flags', reverse: false, showValue: true } }), [], 'an icon set as it is');
  assert.deepEqual(updateConditionalRule(doc.wb, 0, 0, { format: RED }), [], 'a format for an icon set');
  assert.deepEqual(updateConditionalRule(doc.wb, 0, 1, { rule: { type: 'timePeriod' } }), [], 'the kind alone');
  assert.deepEqual(updateConditionalRule(doc.wb, 0, 1, { rule: { type: 'timePeriod', timePeriod: 'lastWeek' }, stopIfTrue: false }), [], 'a period as it is');
  sameBytes(await save(doc), doc.bytes, 'saved without a change');
});

test('a rule of a file: one thing changed is that thing changed; thresholds, the spelling of the formula and what the element holds beside the model stay', async () => {
  const doc = await loadBytes(await fileBook());
  const original = await partText(doc.bytes, SHEET1);
  // The direction of the icons alone.
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { rule: { type: 'iconSet', reverse: true } }));
  assert.deepEqual(conditionalRules(doc.wb, 0)[0].rule.iconSet, { set: '3Flags', values: KEPT_VALUES, showValue: true, reverse: true }, 'the thresholds of the file');
  let saved = await save(doc);
  await checkPackage(saved);
  let xml = await partText(saved, SHEET1);
  const icons = (attrs: string, values = '<cfvo type="percent" val="0"/><cfvo type="num" val="2" gte="0"/><cfvo type="formula" val="$H$1"/>') =>
    `<conditionalFormatting${XR} sqref="D1:D5"><cfRule type="iconSet" priority="1" xr:uid="{I1}"><iconSet ${attrs}>${values}</iconSet></cfRule></conditionalFormatting>`;
  // The rule leaves its element (the other rule stays in it, as the file wrote it) and keeps its unknown attributes.
  assert.equal(xml, original.replace(FILE_ICONS, '').replace(OTHER_BLOCK, OTHER_BLOCK + icons('iconSet="3Flags" reverse="1" percent="0"')), 'the direction alone');
  assert.ok(xml.includes(FILE_PERIOD) && xml.includes(OTHER_BLOCK), 'the rules that were not touched keep their source text');
  await sameContent(saved, doc, 'the direction: read back');

  // Another set of three icons: where one icon ends and the next begins stays.
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { rule: { type: 'iconSet', iconSet: '3Arrows', showValue: false } }));
  assert.deepEqual(conditionalRules(doc.wb, 0)[0].rule.iconSet, { set: '3Arrows', values: KEPT_VALUES, showValue: false, reverse: true }, 'the thresholds of the file under another set');
  xml = await partText(await save(doc), SHEET1);
  assert.ok(xml.includes(icons('iconSet="3Arrows" showValue="0" reverse="1" percent="0"')), xml);
  // A set of five has thresholds of its own.
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { rule: { type: 'iconSet', iconSet: '5Quarters' } }));
  saved = await save(doc);
  await checkPackage(saved);
  xml = await partText(saved, SHEET1);
  assert.ok(xml.includes(icons('iconSet="5Quarters" showValue="0" reverse="1" percent="0"', cfvos(5))), xml);
  await sameContent(saved, doc, 'five icons: read back');

  // The period alone: the formula is written anew for the cell the old one named, the way it named it; the format stays.
  apply(doc, updateConditionalRule(doc.wb, 0, 1, { rule: { type: 'timePeriod', timePeriod: 'yesterday' } }));
  const period = conditionalRules(doc.wb, 0)[1];
  assert.deepEqual([period.rule.timePeriod, period.rule.formulas, period.rule.dxfId], ['yesterday', ['FLOOR($D1,1)=TODAY()-1'], 2], 'the period');
  saved = await save(doc);
  await checkPackage(saved);
  xml = await partText(saved, SHEET1);
  checkSheetOrder(xml);
  // What the element it came from declares goes with it.
  assert.ok(xml.includes(`<conditionalFormatting${XR} sqref="D1:D5"><cfRule type="timePeriod" dxfId="2" priority="2" timePeriod="yesterday"><formula>FLOOR($D1,1)=TODAY()-1</formula></cfRule></conditionalFormatting>`), xml);
  assert.ok(xml.includes(OTHER_BLOCK), 'the rule that was not touched keeps its source text');
  await sameContent(saved, doc, 'the period: read back');
  // The format alone: the period and its formula stay.
  apply(doc, updateConditionalRule(doc.wb, 0, 1, { format: { fill: '#0000FF', italic: true } }));
  assert.deepEqual([conditionalRules(doc.wb, 0)[1].rule.timePeriod, conditionalRules(doc.wb, 0)[1].rule.formulas[0]], ['yesterday', 'FLOOR($D1,1)=TODAY()-1'], 'the format alone');
  await sameContent(await save(doc), doc, 'the format: read back');

  for (let i = 0; i < 5; i++) doc.history.undo();
  sameBytes(await save(doc), doc.bytes, 'after undoing every step');
});

test('a rule of a file: the formula of a period that was only looked at stays as the file spells it when the cells or "stop if true" change', async () => {
  const doc = await loadBytes(await fileBook());
  apply(doc, updateConditionalRule(doc.wb, 0, 1, { stopIfTrue: true }));
  let xml = await partText(await save(doc), SHEET1);
  assert.ok(xml.includes(`<cfRule type="timePeriod" dxfId="2" priority="2" stopIfTrue="1" timePeriod="lastWeek"><formula>${WEEK}</formula></cfRule>`), xml);
  // Other cells: the formula names the new first cell, spelled with the "$" the file had (the rule looks at its own cell).
  apply(doc, updateConditionalRule(doc.wb, 0, 1, { ranges: range('F3:G4') }));
  xml = await partText(await save(doc), SHEET1);
  assert.ok(xml.includes(`<conditionalFormatting${XR} sqref="F3:G4"><cfRule type="timePeriod" dxfId="2" priority="2" stopIfTrue="1" timePeriod="lastWeek"><formula>${WEEK.replace(/\$D1/g, '$F3')}</formula></cfRule></conditionalFormatting>`), xml);
  await sameContent(await save(doc), doc, 'read back');
});

test('one kind becomes the other: an icon set that becomes a period gets a formula and a format, a period that becomes an icon set loses both', async () => {
  const doc = await loadBytes(await fileBook());
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { rule: { type: 'timePeriod', timePeriod: 'nextMonth', format: RED } }));
  apply(doc, updateConditionalRule(doc.wb, 0, 1, { rule: { type: 'iconSet', iconSet: '4Arrows' } }));
  const [first, second] = conditionalRules(doc.wb, 0);
  assert.deepEqual([first.rule.type, first.rule.timePeriod, first.rule.formulas, first.rule.iconSet], ['timePeriod', 'nextMonth', [FORMULAS.nextMonth('D1')], undefined], 'the icon set is a period');
  assert.deepEqual([second.rule.type, second.rule.timePeriod, second.rule.formulas, second.format, second.rule.iconSet!.set], ['iconSet', undefined, [], undefined, '4Arrows'], 'the period is an icon set');
  const saved = await save(doc);
  await checkPackage(saved);
  const xml = await partText(saved, SHEET1);
  checkSheetOrder(xml);
  // The place of each in the file and its unknown attributes stay; what belonged to the kind it was does not.
  assert.ok(xml.includes(`<cfRule type="timePeriod" dxfId="3" priority="1" timePeriod="nextMonth" xr:uid="{I1}"><formula>${escaped(FORMULAS.nextMonth('D1'))}</formula></cfRule>`), xml);
  assert.ok(xml.includes(`<cfRule type="iconSet" priority="2"><iconSet iconSet="4Arrows">${cfvos(4)}</iconSet></cfRule>`), xml);
  await sameContent(saved, doc, 'read back');
  doc.history.undo();
  doc.history.undo();
  sameBytes(await save(doc), doc.bytes, 'after undo');
});

// ---------------------------------------------------------------------------
// Rows, columns and the clipboard

test('rows and columns that move take the formula of a period and the thresholds of an icon set with them, for a new rule and for a rule of the file', async () => {
  const doc = await loadBytes(await numbers(5, FILE_BLOCK));
  apply(doc, addConditionalRule(doc.wb, 0, range('G4:G6'), { type: 'timePeriod', timePeriod: 'thisMonth', format: RED }));
  assert.deepEqual(now(doc.wb), [`G4:G6|timePeriod|1|${FORMULAS.thisMonth('G4')}`, 'D1:D5|iconSet|2|', `D1:D5|timePeriod|3|${FORMULAS.lastWeek('$D1')}`]);
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 2 }]);
  assert.deepEqual(now(doc.wb), [`G6:G8|timePeriod|1|${FORMULAS.thisMonth('G6')}`, 'D3:D7|iconSet|2|', `D3:D7|timePeriod|3|${FORMULAS.lastWeek('$D3')}`]);
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'col', at: 1, count: 3 }]);
  assert.deepEqual(now(doc.wb), [`J6:J8|timePeriod|1|${FORMULAS.thisMonth('J6')}`, 'G3:G7|iconSet|2|', `G3:G7|timePeriod|3|${FORMULAS.lastWeek('$G3')}`]);
  // The threshold that is a formula names the cell that moved.
  assert.deepEqual(conditionalRules(doc.wb, 0)[1].rule.iconSet!.values.map((value) => value.value), ['0', '2', '$K$3'], 'the thresholds of the icon set');
  const saved = await save(doc);
  await checkPackage(saved);
  const xml = await partText(saved, SHEET1);
  assert.ok(xml.includes(`<formula>${escaped(FORMULAS.thisMonth('J6'))}</formula>`) && xml.includes(`<formula>${escaped(FORMULAS.lastWeek('$G3'))}</formula>`) && xml.includes('<cfvo type="formula" val="$K$3"/>'), xml);
  await sameContent(saved, doc, 'read back');
  // The rule of the file made live, then moved again.
  apply(doc, updateConditionalRule(doc.wb, 0, 2, { rule: { type: 'timePeriod', timePeriod: 'tomorrow' } }));
  apply(doc, [{ type: 'delete', sheet: 0, axis: 'row', at: 0, count: 1 }]);
  assert.equal(now(doc.wb)[2], 'G2:G6|timePeriod|3|FLOOR($G2,1)=TODAY()+1');
  // All its cells deleted: the rule is gone, and undo brings it back.
  apply(doc, [{ type: 'delete', sheet: 0, axis: 'row', at: 4, count: 3 }]);
  assert.deepEqual(now(doc.wb).map((line) => line.split('|').slice(0, 2).join('|')), ['G2:G4|iconSet', 'G2:G4|timePeriod']);
  doc.history.undo();
  assert.equal(now(doc.wb)[0], `J5:J7|timePeriod|1|${FORMULAS.thisMonth('J5')}`);
  while (doc.history.canUndo) doc.history.undo();
  await sameParts(doc, 'after undoing every step');
});

test('copied and cut cells take both kinds of rule along: the formula of a period is written for the cells it is pasted on, an icon set keeps its thresholds', async () => {
  const doc = await loadBytes(await numbers(5, FILE_BLOCK));
  apply(doc, addConditionalRule(doc.wb, 0, range('A1:A5'), { type: 'timePeriod', timePeriod: 'last7Days', format: RED }));
  apply(doc, addConditionalRule(doc.wb, 0, range('A1:A5'), { type: 'iconSet', iconSet: '5Arrows', reverse: true }));
  // A copy: the cells it is pasted on get the rules of the cells they are copies of.
  const copied = copyRange(doc.wb, 0, range('A2:A3'));
  apply(doc, pasteAt(doc.wb, 0, range('J10'), copied.text, copied).ops);
  const listed = () => conditionalRules(doc.wb, 0).map((info) => `${info.ranges.map((r) => `${String.fromCharCode(65 + r.left)}${r.top + 1}:${String.fromCharCode(65 + r.right)}${r.bottom + 1}`).join(' ')}|${info.rule.type}|${info.rule.formulas.join(';')}`);
  // The rules the sheet has get the cells; their formulas stay written for their first cell.
  assert.deepEqual(listed().filter((line) => line.includes('J10')), [`A1:A5 J10:J11|iconSet|`, `A1:A5 J10:J11|timePeriod|${FORMULAS.last7Days('A1')}`]);
  const pasted = new ConditionalFormats(doc.wb, doc.wb.sheets[0], { now: () => 3 });
  assert.equal(pasted.resolve(9, 9)!.icon!.set, '5Arrows', 'the icon set is drawn where it was pasted');
  // The rules of the file (D1:D5) are cut and pasted elsewhere: they go with their cells.
  const cut = copyRange(doc.wb, 0, range('D1:D5'));
  apply(doc, cutRange(doc.wb, 0, range('D1:D5')));
  apply(doc, pasteAt(doc.wb, 0, range('L3'), cut.text, cut).ops);
  assert.deepEqual(listed().filter((line) => /^[DL]/.test(line)), ['L3:L7|iconSet|', `L3:L7|timePeriod|${FORMULAS.lastWeek('$L3')}`]);
  assert.deepEqual(conditionalRules(doc.wb, 0).filter((info) => info.rule.type === 'iconSet' && info.ranges[0].left === 11)[0].rule.iconSet, { set: '3Flags', values: KEPT_VALUES, showValue: true, reverse: false }, 'the icon set that was moved');
  const saved = await save(doc);
  await checkPackage(saved);
  checkSheetOrder(await partText(saved, SHEET1));
  await sameContent(saved, doc, 'read back');
  while (doc.history.canUndo) doc.history.undo();
  await sameParts(doc, 'after undoing every step');
});

// ---------------------------------------------------------------------------
// A file built from the model

test('a file built from the model states both kinds as the model has them', async () => {
  const doc = await loadBytes(await numbers(9, FILE_BLOCK));
  for (const name of ICON_SET_NAMES) apply(doc, addConditionalRule(doc.wb, 0, range('A1:A9'), { type: 'iconSet', iconSet: name, reverse: name.charAt(0) === '4', showValue: name.charAt(0) !== '5' }));
  for (const period of TIME_PERIODS) apply(doc, addConditionalRule(doc.wb, 0, range('B2:C4'), { type: 'timePeriod', timePeriod: period, format: RED }));
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 }]);
  const built = await buildFromModel(doc.wb);
  await checkPackage(built);
  const xml = await partText(built, SHEET1);
  checkSheetOrder(xml);
  assert.ok(!xml.includes('extLst'), 'nothing goes into an extension list');
  const back = await (await open(built)).workbook();
  assert.deepEqual(now(back), now(doc.wb), 'the rules, where they apply');
  assert.equal(now(back)[0], `B3:C5|timePeriod|1|${FORMULAS.nextMonth('B3')}`);
  assert.equal(canonical(conditionalContent(back)), canonical(conditionalContent(doc.wb, true)), 'what each rule says and gives');
  assert.deepEqual(iconsOf(back, 10), iconsOf(doc.wb, 10), 'the icons of the cells');
});

// ---------------------------------------------------------------------------
// What is only kept

test('the sets of the 2010 extension are kept, not made: a rule only the extension list has stays byte for byte beside new rules', async () => {
  const doc = await loadBytes(await conditionalBook({ extensionOnly: true }));
  const rules = conditionalRules(doc.wb, 0);
  const stars = rules[rules.length - 1];
  assert.deepEqual([stars.rule.type, stars.rule.iconSet!.set, stars.rule.extended, stars.editable, stars.removable, stars.movable], ['iconSet', '3Stars', true, false, false, false], 'the rule of the extension list');
  assert.throws(() => updateConditionalRule(doc.wb, 0, stars.index, { rule: { type: 'iconSet', iconSet: '3Arrows' } }), /extension list/, 'changing it');
  assert.throws(() => updateConditionalRule(doc.wb, 0, stars.index, { rule: { type: 'iconSet', reverse: true } }), /extension list/, 'turning it round');
  assert.throws(() => removeConditionalRule(doc.wb, 0, stars.index), /extension list/, 'deleting it');
  // The icon set and the period of the main part are changed beside it, and new ones are added.
  apply(doc, updateConditionalRule(doc.wb, 0, 4, { rule: { type: 'iconSet', iconSet: '3Symbols2' } }));
  apply(doc, updateConditionalRule(doc.wb, 0, 5, { rule: { type: 'timePeriod', timePeriod: 'thisWeek' } }));
  apply(doc, addConditionalRule(doc.wb, 0, range('G1:G3'), { type: 'iconSet', iconSet: '5Rating' }));
  apply(doc, addConditionalRule(doc.wb, 0, range('G1:G3'), { type: 'timePeriod', timePeriod: 'today', format: RED }));
  const saved = await save(doc);
  await checkPackage(saved);
  const xml = await partText(saved, SHEET1);
  checkSheetOrder(xml);
  assert.ok(xml.includes(EXT_ONLY), 'the rule of the extension list, as the file wrote it');
  assert.ok(xml.includes('<cfRule type="iconSet" priority="5"><iconSet iconSet="3Symbols2"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule>'), xml);
  assert.ok(xml.includes(`<cfRule type="timePeriod" dxfId="2" priority="6" timePeriod="thisWeek"><formula>${escaped(FORMULAS.thisWeek('D1'))}</formula></cfRule>`), xml);
  // In a sheet with a rule whose number is kept, new rules come last.
  assert.ok(xml.includes(`<conditionalFormatting sqref="G1:G3"><cfRule type="iconSet" priority="8"><iconSet iconSet="5Rating">${cfvos(5)}</iconSet></cfRule></conditionalFormatting>`), xml);
  assert.ok(xml.includes(`<cfRule type="timePeriod" dxfId="3" priority="9" timePeriod="today"><formula>FLOOR(G1,1)=TODAY()</formula></cfRule>`), xml);
  await sameContent(saved, doc, 'read back');
  while (doc.history.canUndo) doc.history.undo();
  sameBytes(await save(doc), doc.bytes, 'after undoing every step');
});

test('a kind of rule that is not made here is still kept as it is, and a set or a period of a file that is not offered stays on its rule', async () => {
  const strange =
    cf('A1:A5', '<cfRule type="somethingNew" priority="1" future="1"/>') +
    cf('B1:B5', '<cfRule type="iconSet" priority="2"><iconSet iconSet="3Unheard"><cfvo type="percent" val="0"/><cfvo type="num" val="2"/><cfvo type="num" val="4"/></iconSet></cfRule>') +
    cf('C1:C5', '<cfRule type="timePeriod" dxfId="0" priority="3" timePeriod="someday"><formula>C1=TODAY()</formula></cfRule>');
  const doc = await loadBytes(await numbers(5, strange));
  const rules = conditionalRules(doc.wb, 0);
  assert.deepEqual(rules.map((info) => [info.rule.type, info.editable]), [['somethingNew', false], ['iconSet', true], ['timePeriod', true]]);
  assert.throws(() => updateConditionalRule(doc.wb, 0, 0, { rule: { type: 'iconSet' } }), /cannot be changed here/, 'a kind that is kept');
  // The set of the file stays while only its direction changes; so does the period while only the format does.
  apply(doc, updateConditionalRule(doc.wb, 0, 1, { rule: { type: 'iconSet', iconSet: '3Unheard' as '3Arrows', reverse: true } }));
  apply(doc, updateConditionalRule(doc.wb, 0, 2, { rule: { type: 'timePeriod', timePeriod: 'someday' as 'today', format: RED } }));
  const saved = await save(doc);
  await checkPackage(saved);
  const xml = await partText(saved, SHEET1);
  assert.ok(xml.includes('<cfRule type="somethingNew" priority="1" future="1"/>'), 'the kind that is kept, as the file wrote it');
  assert.ok(xml.includes('<iconSet iconSet="3Unheard" reverse="1"><cfvo type="percent" val="0"/><cfvo type="num" val="2"/><cfvo type="num" val="4"/></iconSet>'), xml);
  assert.ok(xml.includes('timePeriod="someday"><formula>C1=TODAY()</formula>'), xml);
  await sameContent(saved, doc, 'read back');
  // Any other name is refused.
  assert.throws(() => updateConditionalRule(doc.wb, 0, 1, { rule: { type: 'iconSet', iconSet: '3Stars' as '3Arrows' } }), /icon set is not known/, 'a set of the extension');
  assert.throws(() => updateConditionalRule(doc.wb, 0, 2, { rule: { type: 'timePeriod', timePeriod: 'sometime' as 'today' } }), /period is not known/, 'a period nobody knows');
});

// ---------------------------------------------------------------------------
// An independent reader

const PYTHON = `
import json, sys
import openpyxl
wb = openpyxl.load_workbook(sys.argv[1])
ws = wb.active
out = []
for cf in ws.conditional_formatting:
    for r in cf.rules:
        rule = {'sqref': str(cf.sqref), 'type': r.type, 'priority': r.priority, 'timePeriod': r.timePeriod, 'formulas': [str(f) for f in r.formula], 'dxf': r.dxf is not None and r.dxf.fill is not None}
        if r.iconSet is not None:
            s = r.iconSet
            rule['iconSet'] = {'set': s.iconSet, 'showValue': s.showValue is None or bool(s.showValue), 'reverse': bool(s.reverse), 'values': [[v.type, None if v.val is None else str(int(v.val)), v.gte is None or bool(v.gte)] for v in s.cfvo]}
        out.append(rule)
out.sort(key=lambda rule: rule['priority'])
print(json.dumps(out))
`;

test('openpyxl reads the rules that were made here as the model has them', async (t) => {
  const probe = spawnSync('python3', ['-c', 'import openpyxl'], { encoding: 'utf8' });
  if (probe.status !== 0) {
    t.skip('python3 with openpyxl is not available');
    return;
  }
  const doc = await loadBytes(await numbers(9));
  for (const name of ICON_SET_NAMES) apply(doc, addConditionalRule(doc.wb, 0, range('A1:A9'), { type: 'iconSet', iconSet: name, reverse: name.charAt(0) === '4', showValue: name.charAt(0) !== '5' }));
  for (const period of TIME_PERIODS) apply(doc, addConditionalRule(doc.wb, 0, range('B2:C4'), { type: 'timePeriod', timePeriod: period, format: RED }));
  const expected = conditionalRules(doc.wb, 0).map((info) => {
    const rule = info.rule;
    const view: Record<string, unknown> = { sqref: info.ranges.map((r) => `${String.fromCharCode(65 + r.left)}${r.top + 1}:${String.fromCharCode(65 + r.right)}${r.bottom + 1}`).join(' '), type: rule.type, priority: rule.priority, timePeriod: rule.timePeriod === undefined ? null : rule.timePeriod, formulas: rule.formulas, dxf: info.format !== undefined };
    if (rule.iconSet) view.iconSet = { set: rule.iconSet.set, showValue: rule.iconSet.showValue, reverse: rule.iconSet.reverse, values: rule.iconSet.values.map((value) => [value.type, value.value, value.gte]) };
    return view;
  });
  assert.equal(expected.length, ICON_SET_NAMES.length + TIME_PERIODS.length, 'every rule is listed');
  const folder = join(dirname(fileURLToPath(import.meta.url)), '..', `tmp-rules-icons-${process.pid}`);
  mkdirSync(folder, { recursive: true });
  try {
    for (const [label, bytes] of [['written into the file', await save(doc)], ['built from the model', await buildFromModel(doc.wb)]] as const) {
      const file = join(folder, 'rules.xlsx');
      writeFileSync(file, bytes);
      const run = spawnSync('python3', ['-c', PYTHON, file], { encoding: 'utf8' });
      assert.equal(run.status, 0, `${label}: ${run.stderr}`);
      assert.deepEqual(JSON.parse(run.stdout), expected, label);
    }
  } finally {
    rmSync(folder, { recursive: true, force: true });
  }
});
