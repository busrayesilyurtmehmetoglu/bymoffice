// Icon sets and "dates occurring" rules: what a review found. What the element of a rule of
// a file holds beside the model stays when the rule is changed (inside its thresholds too),
// a format record a file names on an icon set stays, the cell a period's formula names keeps
// its spelling whatever the spaces around it, and where the rule and its formula part ways.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine } from '../src/calc/engine';
import { ConditionalFormats } from '../src/format/conditional';
import { parseRange } from '../src/index';
import type { CellRange } from '../src/index';
import type { Op } from '../src/ops/ops';
import { addConditionalRule, conditionalRules, updateConditionalRule } from '../src/ops/rules';
import { STYLES, cf, conditionalContent, conditionalNow as now } from './conditional-fixtures';
import { buildXlsx, worksheetXml } from './fixtures';
import { checkPackage, checkSheetOrder, load, partText, reopen, save, sameBytes, type Loaded } from './rules-fixtures';
import { canonical } from './workbook-content';

const range = (ref: string): CellRange => parseRange(ref)!;
const SHEET1 = 'xl/worksheets/sheet1.xml';
const apply = (doc: Loaded, ops: Op[]) => doc.history.apply(ops, undefined, undefined);
const book = (inner: string, data = '<row r="1"><c r="A1"><v>1</v></c></row>') => buildXlsx({ sheets: [{ name: 'Data', xml: worksheetXml(`<sheetData>${data}</sheetData>${inner}`) }], styles: STYLES });
const rulesOf = (xml: string) => (/<conditionalFormatting[\s\S]*<\/conditionalFormatting>/.exec(xml) || [''])[0];
/** Saves, checks the package, and returns the rule elements of the first sheet. */
async function written(doc: Loaded, label: string): Promise<string> {
  const saved = await save(doc);
  await checkPackage(saved);
  const xml = await partText(saved, SHEET1);
  checkSheetOrder(xml, label);
  assert.equal(canonical(conditionalContent(await reopen(saved))), canonical(conditionalContent(doc.wb)), `${label}: read back`);
  return rulesOf(xml);
}

// ---------------------------------------------------------------------------
// 1. What a threshold and an icon set hold beside the model

const FIRST = '<cfvo type="percent" val="0" foo="1"/>';
const SECOND = '<cfvo type="num" val="2"><extLst><ext uri="{X}"/></extLst></cfvo>';
const THIRD = '<cfvo type="formula" val="$H$1" gte="0" bar="2"/>';
const FUTURE = '<future a="1"/>';
const icons = (attrs: string, inner = FIRST + SECOND + THIRD + FUTURE) => `<cfRule type="iconSet" priority="1"><iconSet ${attrs}>${inner}</iconSet></cfRule>`;

test('review: an icon set of a file that is changed keeps what its thresholds and its element hold beside the model', async () => {
  const doc = await load(await book(cf('A1:A5', icons('iconSet="3Flags"'))));
  // The direction alone.
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { rule: { type: 'iconSet', reverse: true } }));
  assert.equal(await written(doc, 'the direction'), cf('A1:A5', icons('iconSet="3Flags" reverse="1"')), 'the direction');
  // Another set of three, the icon alone: the thresholds are those of the file, with all they hold.
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { rule: { type: 'iconSet', iconSet: '3Arrows', showValue: false } }));
  assert.equal(await written(doc, 'the set'), cf('A1:A5', icons('iconSet="3Arrows" showValue="0" reverse="1"')), 'the set');
  // A threshold that is a formula follows its cell: its value alone is written anew.
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 2 }]);
  assert.equal(await written(doc, 'rows inserted'), cf('A3:A7', icons('iconSet="3Arrows" showValue="0" reverse="1"', FIRST + SECOND + THIRD.replace('$H$1', '$H$3') + FUTURE)), 'rows inserted');
  while (doc.history.canUndo) doc.history.undo();
  assert.equal(rulesOf(await partText(await save(doc), SHEET1)), cf('A1:A5', icons('iconSet="3Flags"')), 'after undoing every step');
});

test('review: the same holds when only the cells of the rule or "stop if true" change', async () => {
  const doc = await load(await book(cf('A1:A5', icons('iconSet="3Flags"'))));
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { ranges: range('B2:B6') }));
  assert.equal(await written(doc, 'the cells'), cf('B2:B6', icons('iconSet="3Flags"')), 'the cells');
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { stopIfTrue: true }));
  assert.equal(await written(doc, 'stop if true'), cf('B2:B6', icons('iconSet="3Flags"').replace('priority="1"', 'priority="1" stopIfTrue="1"')), 'stop if true');
});

test('review: unknown elements of an icon set stay where they are among its thresholds; a set with another number of icons has thresholds of its own', async () => {
  const between = FIRST + '<early b="1"/>' + SECOND + THIRD + FUTURE;
  const doc = await load(await book(cf('A1:A5', icons('iconSet="3Flags"', between))));
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { rule: { type: 'iconSet', reverse: true } }));
  assert.equal(await written(doc, 'the direction'), cf('A1:A5', icons('iconSet="3Flags" reverse="1"', between)), 'the direction');
  // Four icons: the thresholds are new, and what the old ones held goes with them. What the set itself holds stays.
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { rule: { type: 'iconSet', iconSet: '4Rating' } }));
  const four = '<cfvo type="percent" val="0"/><early b="1"/><cfvo type="percent" val="25"/><cfvo type="percent" val="50"/><cfvo type="percent" val="75"/>' + FUTURE;
  assert.equal(await written(doc, 'four icons'), cf('A1:A5', icons('iconSet="4Rating" reverse="1"', four)), 'four icons');
});

test('review: the thresholds of a colour scale and of a data bar keep what they hold beside the model as well', async () => {
  const scale = (colors: string) => `<cfRule type="colorScale" priority="1"><colorScale><cfvo type="min" foo="1"/><cfvo type="formula" val="$H$1"><extLst><ext uri="{X}"/></extLst></cfvo><cfvo type="max"/>${colors}</colorScale></cfRule>`;
  const bar = (color: string) => `<cfRule type="dataBar" priority="2"><dataBar><cfvo type="num" val="0" foo="1"/><cfvo type="percent" val="90" bar="2"/><color rgb="${color}"/></dataBar></cfRule>`;
  const doc = await load(await book(cf('A1:A5', scale('<color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/>')) + cf('B1:B5', bar('FF638EC6'))));
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { rule: { type: 'colorScale', colors: ['#FF0000', '#FFEB84', '#63BE7B'] } }));
  apply(doc, updateConditionalRule(doc.wb, 0, 1, { rule: { type: 'dataBar', colors: ['#00FF00'] } }));
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 }]);
  const xml = await written(doc, 'colours');
  assert.equal(xml, cf('A2:A6', scale('<color rgb="FFFF0000"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/>').replace('$H$1', '$H$2')) + cf('B2:B6', bar('FF00FF00')), 'colours');
});

// ---------------------------------------------------------------------------
// 2. A format record a file names on a rule that shows what it finds by itself

test('review: a format record a file names on an icon set, a colour scale or a data bar stays when the rule is changed, and stating the rule again is no change', async () => {
  const file =
    cf('A1:A5', '<cfRule type="iconSet" dxfId="0" priority="1"><iconSet iconSet="3Flags"><cfvo type="percent" val="0"/><cfvo type="num" val="2"/><cfvo type="num" val="4"/></iconSet></cfRule>') +
    cf('B1:B5', '<cfRule type="colorScale" dxfId="1" priority="2"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule>') +
    cf('C1:C5', '<cfRule type="dataBar" dxfId="2" priority="3"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule>');
  const doc = await load(await book(file));
  assert.deepEqual(updateConditionalRule(doc.wb, 0, 0, { rule: { type: 'iconSet', iconSet: '3Flags', reverse: false, showValue: true } }), [], 'an icon set as it is');
  assert.deepEqual(updateConditionalRule(doc.wb, 0, 1, { rule: { type: 'colorScale', colors: ['#F8696B', '#63BE7B'] } }), [], 'a colour scale as it is');
  assert.deepEqual(updateConditionalRule(doc.wb, 0, 2, { rule: { type: 'dataBar', colors: ['#638EC6'] } }), [], 'a data bar as it is');
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { rule: { type: 'iconSet', reverse: true } }));
  apply(doc, updateConditionalRule(doc.wb, 0, 1, { rule: { type: 'colorScale', colors: ['#FF0000', '#63BE7B'] } }));
  apply(doc, updateConditionalRule(doc.wb, 0, 2, { rule: { type: 'dataBar', colors: ['#00FF00'] } }));
  assert.deepEqual(conditionalRules(doc.wb, 0).map((info) => info.rule.dxfId), [0, 1, 2], 'the records the file names');
  const xml = await written(doc, 'changed');
  assert.ok(xml.includes('<cfRule type="iconSet" dxfId="0" priority="1"><iconSet iconSet="3Flags" reverse="1">'), xml);
  assert.ok(xml.includes('<cfRule type="colorScale" dxfId="1" priority="2">'), xml);
  assert.ok(xml.includes('<cfRule type="dataBar" dxfId="2" priority="3">'), xml);
  // A rule that becomes one of these kinds takes no format along, and a new one gets none.
  apply(doc, addConditionalRule(doc.wb, 0, range('E1'), { type: 'cellIs', operator: 'equal', formulas: ['1'], format: { fill: '#FFC7CE' } }));
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { rule: { type: 'iconSet' } }));
  assert.deepEqual([conditionalRules(doc.wb, 0)[0].rule.dxf, conditionalRules(doc.wb, 0)[0].rule.dxfId], [undefined, undefined], 'a rule that became an icon set');
  while (doc.history.canUndo) doc.history.undo();
  sameBytes(await save(doc), doc.bytes, 'after undoing every step');
});

// ---------------------------------------------------------------------------
// 3. The cell a period's formula names

test('review: when the period changes, the new formula names the cell the way the old one did, however the old one is spelled', async () => {
  const period = (formula: string, sqref = 'B2:B6') => cf(sqref, `<cfRule type="timePeriod" dxfId="0" priority="1" timePeriod="today"><formula>${formula}</formula></cfRule>`);
  const changed = async (formula: string, sqref?: string) => {
    const doc = await load(await book(period(formula, sqref)));
    apply(doc, updateConditionalRule(doc.wb, 0, 0, { rule: { type: 'timePeriod', timePeriod: 'tomorrow' } }));
    return now(doc.wb)[0].split('|')[3];
  };
  assert.equal(await changed('FLOOR($B2, 1)=TODAY()'), 'FLOOR($B2,1)=TODAY()+1', 'a space the standard formula has not');
  assert.equal(await changed('floor( B$2 ,1) = today()'), 'FLOOR(B$2,1)=TODAY()+1', 'small letters and spaces');
  assert.equal(await changed('TODAY()=INT($B$2)'), 'FLOOR($B$2,1)=TODAY()+1', 'another formula altogether');
  assert.equal(await changed('FLOOR(B2,1)=TODAY()'), 'FLOOR(B2,1)=TODAY()+1', 'the standard formula');
  // No reference to the first cell of the rule: the plain name of that cell.
  assert.equal(await changed('TODAY()=45000'), 'FLOOR(B2,1)=TODAY()+1', 'no reference');
  assert.equal(await changed('FLOOR($C$9,1)=TODAY()'), 'FLOOR(B2,1)=TODAY()+1', 'a reference to another cell');
  assert.equal(await changed('FLOOR(Data!$B2,1)=TODAY()'), 'FLOOR(B2,1)=TODAY()+1', 'a reference with a sheet name');
  assert.equal(await changed('SUM($B2:$B6)>TODAY()'), 'FLOOR(B2,1)=TODAY()+1', 'a range');
  assert.equal(await changed('"$B2"&FLOOR(B2,1)=TODAY()'), 'FLOOR(B2,1)=TODAY()+1', 'a text that looks like a reference');
});

// ---------------------------------------------------------------------------
// 6. Where the rule and its formula part ways, and which cell the formula is written for

test('review (UNSURE, not compared with Excel): a date that is kept as text is no date to the rule, though the formula it is written with takes it for one', async () => {
  // Serial 45000 is 15 March 2023. A1 a number, A2 the same day as a text of digits, A3 a text the locale reads as that date, A4 TRUE, A5 a number with a time.
  const data = '<row r="1"><c r="A1"><v>45000</v></c></row><row r="2"><c r="A2" t="str"><v>45000</v></c></row><row r="3"><c r="A3" t="str"><v>3/15/2023</v></c></row><row r="4"><c r="A4" t="b"><v>1</v></c></row><row r="5"><c r="A5"><v>45000.7</v></c></row>';
  const doc = await load(await book('', data));
  apply(doc, addConditionalRule(doc.wb, 0, range('A1:A5'), { type: 'timePeriod', timePeriod: 'today', format: { fill: '#FF0000' } }));
  const clock = () => 45000.4;
  const formats = new ConditionalFormats(doc.wb, doc.wb.sheets[0], { now: clock });
  const engine = new CalcEngine(doc.wb, { now: clock });
  const formula = conditionalRules(doc.wb, 0)[0].rule.formulas[0];
  const byRule: boolean[] = [];
  const byFormula: boolean[] = [];
  for (let row = 0; row < 5; row++) {
    const found = formats.resolve(row, 0);
    byRule.push(!!found && !!found.style && found.style.fill !== undefined);
    byFormula.push(engine.evaluateText(formula.replace(/A1/g, `A${row + 1}`), { sheet: 0, row, col: 0 }) === true);
  }
  // The rule is evaluated by its period and looks at numbers only; the formula coerces the texts.
  assert.deepEqual(byRule, [true, false, false, false, true], 'by the rule: numbers only');
  assert.deepEqual(byFormula, [true, true, true, false, true], 'by the formula: texts that read as the date as well');
});

test('review (UNSURE, not compared with Excel): the formula of a rule with several ranges is written for the first cell of the first of them', async () => {
  const doc = await load(await book(''));
  apply(doc, addConditionalRule(doc.wb, 0, [range('C3:C9'), range('A1:A5')], { type: 'timePeriod', timePeriod: 'today', format: { fill: '#FF0000' } }));
  assert.deepEqual(now(doc.wb), ['C3:C9 A1:A5|timePeriod|1|FLOOR(C3,1)=TODAY()']);
  // Evaluated for each cell from there: A1 is looked at through "two columns left, two rows up" of C3.
  const formats = new ConditionalFormats(doc.wb, doc.wb.sheets[0], { now: () => 1.5 });
  const filled = formats.resolve(0, 0);
  assert.ok(!!filled && !!filled.style && filled.style.fill !== undefined, 'A1 (the number 1) is today when today is day 1');
});
