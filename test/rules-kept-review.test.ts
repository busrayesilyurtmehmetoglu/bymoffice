// Review of what rules hold beside the fields of the model (RuleKept):
// - content of the markup-compatibility namespace (and any attribute whose value names
//   prefixes) is not carried: stated elsewhere, its prefixes would name nothing;
// - what is held is small in memory, and a sheet and a workbook hold only so much of it;
// - a rule whose content is not carried is a rule that is kept as the file has it: a cut
//   leaves it where it is, a copy does not take it, and both are counted;
// - what readers may pass over is said for the sheet the rule is in, not the one it came from;
// - only what a rule states now is counted as lost;
// - comments and processing instructions are content that is not carried; `uid` is an
//   identity only in the revision namespaces, <id> only in the one of the extension list;
// - what a data validation of the extension list holds is not written into the main part.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { parseRange, parseXml, type CellRange, type ConditionalFormat, type Workbook } from '../src/index';
import { writeXlsx } from '../src/formats/xlsx/write';
import { MAX_KEPT_BOOK, MAX_KEPT_SHEET, MAX_KEPT_SIZE } from '../src/model/sheet/types';
import { copyRange, cutRange, pasteAt } from '../src/ops/commands';
import { conditionalFormatNow, conditionalRules, dataValidationNow, ruleKeptNow, ruleSpace, updateConditionalRule, validationKeptNow } from '../src/ops/rules';
import { copySheet, copySheetLosses } from '../src/ops/sheets';
import { STYLES, cf } from './conditional-fixtures';
import { buildXlsx, worksheetXml } from './fixtures';
import { checkPackage, load, partText, reopen, save, sameBytes, type Loaded } from './rules-fixtures';
import { mceProblems } from './rules-random-fixtures';

const R = (text: string): CellRange => parseRange(text) as CellRange;
const apply = (doc: Loaded, ops: Parameters<Loaded['history']['apply']>[0]) => doc.history.apply(ops, undefined, undefined);
const formats = (xml: string) => xml.match(/<(?:\w+:)?conditionalFormatting[\s\S]*?<\/(?:\w+:)?conditionalFormatting>/g) || [];
const validations = (xml: string) => (xml.match(/<dataValidations[\s\S]*<\/dataValidations>/) || [''])[0];
const sheet = (root: string, inner: string) => worksheetXml(inner).replace('<worksheet ', `<worksheet ${root} `);
const book = (sheets: [name: string, xml: string][]) => buildXlsx({ sheets: sheets.map(([name, xml]) => ({ name, xml })), styles: STYLES });
const DATA = '<sheetData><row r="1"><c r="D1"><v>5</v></c></row></sheetData>';
const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const XR = 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision';

function keptOf(wb: Workbook, index: number): unknown[] {
  const space = ruleSpace(wb, wb.sheets[index]);
  const out: unknown[] = [];
  for (const format of wb.sheets[index].conditionalFormats || []) {
    const now = conditionalFormatNow(format, space);
    if (now) for (const rule of now.rules) out.push(ruleKeptNow(rule));
  }
  return out;
}
function validationKept(wb: Workbook, index: number): unknown[] {
  const space = ruleSpace(wb, wb.sheets[index]);
  const out: unknown[] = [];
  for (const rule of wb.sheets[index].dataValidations || []) {
    const now = dataValidationNow(rule, space);
    if (now) out.push(validationKeptNow(now));
  }
  return out;
}

test('B1: content of the markup-compatibility namespace, and an attribute that names prefixes, is not carried: the rule stays where it is, and a copy is counted', async () => {
  const choice = `<cfRule type="cellIs" dxfId="0" priority="1" operator="equal"><formula>1</formula><extLst><ext uri="{A}"><mc:AlternateContent><mc:Choice Requires="x14"><x14:thing/></mc:Choice><mc:Fallback><plain/></mc:Fallback></mc:AlternateContent></ext></extLst></cfRule>`;
  const ignorable = `<cfRule type="cellIs" dxfId="0" priority="2" operator="equal" mc:Ignorable="foo" xmlns:foo="urn:foo" foo:a="1"><formula>2</formula></cfRule>`;
  const typed = `<cfRule type="cellIs" dxfId="0" priority="3" operator="equal"><formula>3</formula><extLst><ext uri="{T}"><v xsi:type="x14:thing" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"/></ext></extLst></cfRule>`;
  const dv = `<dataValidations count="1"><dataValidation type="whole" mc:Ignorable="foo" xmlns:foo="urn:foo" foo:a="1" sqref="A1:A3"><formula1>1</formula1><formula2>9</formula2></dataValidation></dataValidations>`;
  const doc = await load(
    await book([
      ['Data', sheet(`${MC} xmlns:x14="${X14}"`, DATA + cf('D1:D5', choice) + cf('F1:F5', ignorable) + cf('H1:H5', typed) + dv)],
      ['Calc', worksheetXml('<sheetData/>')],
    ]),
  );
  assert.deepEqual(keptOf(doc.wb, 0), [{ over: true }, { over: true }, { over: true }], 'the model holds none of it, and says so');
  assert.deepEqual(validationKept(doc.wb, 0), [{ over: true }], 'the same for a data validation');
  sameBytes(await save(doc), doc.bytes, 'saved as it was opened');
  assert.deepEqual(copySheetLosses(doc.wb, 0), [{ what: 'ruleContent', count: 4 }], 'a copy of the sheet would be without these rules');
  const copied = copyRange(doc.wb, 0, R('A1:H5'));
  const result = pasteAt(doc.wb, 1, R('A1'), copied.text, copied);
  apply(doc, result.ops);
  assert.equal(result.skippedRules, 4, 'none of the four comes along, and each is counted');
  apply(doc, copySheet(doc.wb, 0, { name: 'Copy' }));
  const saved = await save(doc);
  await checkPackage(saved);
  const back = await reopen(saved);
  for (const item of back.sheets) {
    const text = await partText(saved, item.part.slice(1));
    assert.deepEqual(mceProblems(text), [], `${item.name}: every prefix a markup-compatibility attribute names is declared`);
    if (item.name !== 'Data') assert.ok(formats(text).length === 0 && validations(text) === '' && text.indexOf('Requires') < 0 && text.indexOf('Ignorable') < 0, `${item.name}: no such rule, in:\n${text}`);
  }
  const rulesOf = (text: string) => formats(text).join('') + validations(text);
  assert.equal(rulesOf(await partText(saved, 'xl/worksheets/sheet1.xml')), rulesOf(await partText(doc.bytes, 'xl/worksheets/sheet1.xml')), 'the rules of the sheet they are in are as they were');
});

test('B2: what rules hold beside the fields of the model takes memory in proportion to its text, and a sheet and a workbook hold only so much of it', async () => {
  const inner = `<extLst><ext uri="{A}"><q xmlns="">${'<a/>'.repeat(880)}</q></ext></extLst>`;
  const rules = (count: number) => {
    let out = '';
    for (let i = 0; i < count; i++) out += cf(`A${i + 1}`, `<cfRule type="cellIs" priority="${i + 1}" operator="equal"><formula>1</formula>${inner}</cfRule>`);
    return out;
  };
  const gc = (globalThis as { gc?: () => void }).gc;
  const one = await buildXlsx({ sheets: [{ name: 'S', xml: worksheetXml('<sheetData/>' + rules(1500)) }], styles: STYLES });
  if (gc) gc();
  const before = process.memoryUsage().heapUsed;
  const wb = await reopen(one);
  if (gc) gc();
  const grown = process.memoryUsage().heapUsed - before;
  const held = (list: ConditionalFormat[] | undefined) => (list || []).reduce((sum, format) => sum + format.rules.reduce((n, rule) => n + (rule.kept && rule.kept.packed ? rule.kept.packed.length : 0), 0), 0);
  const kept = (wb.sheets[0].conditionalFormats || []).filter((format) => format.rules[0].kept && !format.rules[0].kept.over).length;
  assert.ok(kept > 100 && kept < 1500, `some rules hold their content, the others say that there is more (${kept} of 1500 hold it)`);
  assert.ok(held(wb.sheets[0].conditionalFormats) <= MAX_KEPT_SHEET, `a sheet holds no more than ${MAX_KEPT_SHEET} characters (${held(wb.sheets[0].conditionalFormats)})`);
  // Generous: 1,500 rules of 3.6 KB of text each took 72 MB as trees of objects. Without a collection asked for, nothing is measured.
  if (gc) assert.ok(grown < 30e6, `the workbook takes ${Math.round(grown / 1e6)} MB, which is more than 30 MB`);
  // A rule whose content is over the budget is as any rule with more than is carried: in the file as it was.
  const doc = await load(one);
  sameBytes(await save(doc), doc.bytes, 'saved as it was opened');
  // Six sheets of that kind: the workbook holds no more than its own budget.
  const many = await reopen(await buildXlsx({ sheets: [0, 1, 2, 3, 4, 5].map((n) => ({ name: `S${n}`, xml: worksheetXml('<sheetData/>' + rules(1300)) })), styles: STYLES }));
  const total = many.sheets.reduce((sum, item) => sum + held(item.conditionalFormats), 0);
  assert.ok(total <= MAX_KEPT_BOOK && total > MAX_KEPT_BOOK / 2, `a workbook holds no more than ${MAX_KEPT_BOOK} characters (${total})`);
  for (const item of many.sheets) assert.ok(held(item.conditionalFormats) <= MAX_KEPT_SHEET, `${item.name}: no more than a sheet holds`);
});

test('B3: a rule whose content is not carried is left where it is by a cut and is not taken by a copy; both are counted', async () => {
  const big = `<cfRule type="iconSet" priority="5" note="${'n'.repeat(MAX_KEPT_SIZE + 10)}"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule>`;
  const bigV = `<dataValidations count="1"><dataValidation type="whole" note="${'v'.repeat(MAX_KEPT_SIZE + 10)}" sqref="A1:A3"><formula1>1</formula1><formula2>9</formula2></dataValidation></dataValidations>`;
  const source = worksheetXml(DATA + cf('D1:D5', big) + bigV);
  for (const [what, selection] of [
    ['all its cells', 'A1:D5'],
    ['some of its cells', 'A1:D2'],
  ] as const) {
    for (const target of [0, 1]) {
      const doc = await load(await book([['Data', source], ['Calc', worksheetXml('<sheetData/>')]]));
      const copied = copyRange(doc.wb, 0, R(selection));
      apply(doc, cutRange(doc.wb, 0, R(selection)));
      const result = pasteAt(doc.wb, target, R('H1'), copied.text, copied);
      apply(doc, result.ops);
      assert.equal(result.skippedRules, 2, `${what} cut, pasted into sheet ${target}: both rules are said not to have come`);
      const saved = await save(doc);
      await checkPackage(saved);
      const text = await partText(saved, 'xl/worksheets/sheet1.xml');
      assert.deepEqual(formats(text), [cf('D1:D5', big)], `${what} cut, pasted into sheet ${target}: the conditional rule is where it was, as it was`);
      assert.equal(validations(text), bigV, `${what} cut, pasted into sheet ${target}: the data validation as well`);
      if (target === 1) {
        const there = await partText(saved, 'xl/worksheets/sheet2.xml');
        assert.ok(formats(there).length === 0 && validations(there) === '', `${what}: no rule without its content in the other sheet:\n${there}`);
      }
    }
  }
  // A copy beside, in the sheet itself: the rule is not made a second time without what it holds.
  const doc = await load(await book([['Data', source]]));
  const copied = copyRange(doc.wb, 0, R('A1:D5'));
  const result = pasteAt(doc.wb, 0, R('H1'), copied.text, copied);
  apply(doc, result.ops);
  assert.equal(result.skippedRules, 2, 'copied beside: counted');
  sameBytes(await partTextBytes(await save(doc)), await partTextBytes(doc.bytes), 'copied beside: the rules of the sheet are as they were');
});

const partTextBytes = async (bytes: Uint8Array) => new TextEncoder().encode((formats(await partText(bytes, 'xl/worksheets/sheet1.xml')).join('') + validations(await partText(bytes, 'xl/worksheets/sheet1.xml'))).replace(/sqref="[^"]*"/g, ''));

test('B4: what readers may pass over is said for the sheet a rule is in: the model before the file is saved is the model after it is read back', async () => {
  const rule = '<cfRule type="cellIs" dxfId="0" priority="1" operator="equal" my:must="1"><formula>1</formula></cfRule>';
  const dv = '<dataValidations count="1"><dataValidation type="whole" my:must="1" sqref="A1:A3"><formula1>1</formula1><formula2>9</formula2></dataValidation></dataValidations>';
  const doc = await load(
    await book([
      ['Src', sheet('xmlns:my="urn:mine"', DATA + cf('D1:D5', rule) + dv)],
      ['Ign', sheet(`${MC} mc:Ignorable="q" xmlns:q="urn:mine"`, '<sheetData/>')],
      ['Third', worksheetXml('<sheetData/>')],
    ]),
  );
  let copied = copyRange(doc.wb, 0, R('A1:D5'));
  apply(doc, pasteAt(doc.wb, 1, R('A1'), copied.text, copied).ops);
  const expected = { attrs: [['urn:mine', 'must', '1']], ignorable: ['urn:mine'] };
  assert.deepEqual(keptOf(doc.wb, 1), [expected], 'the conditional rule, in the sheet that lets readers pass the namespace over');
  assert.deepEqual(validationKept(doc.wb, 1), [expected], 'the data validation');
  copied = copyRange(doc.wb, 1, R('A1:D5'));
  assert.equal(pasteAt(doc.wb, 2, R('A1'), copied.text, copied).skippedRules, 2, 'copied on into a sheet that does not: counted, before any save');
  const saved = await save(doc);
  const back = await load(saved);
  assert.deepEqual(keptOf(back.wb, 1), keptOf(doc.wb, 1), 'the file reads back as the model: the conditional rule');
  assert.deepEqual(validationKept(back.wb, 1), validationKept(doc.wb, 1), 'the file reads back as the model: the data validation');
  copied = copyRange(back.wb, 1, R('A1:D5'));
  assert.equal(pasteAt(back.wb, 2, R('A1'), copied.text, copied).skippedRules, 2, 'and the same is counted after it');
});

test('B5: only what a rule states now is counted as lost', async () => {
  const rule = '<cfRule type="iconSet" priority="1"><iconSet iconSet="3Arrows" my:h="1"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule>';
  const doc = await load(
    await book([
      ['Src', sheet(`${MC} mc:Ignorable="my" xmlns:my="urn:mine"`, DATA + cf('D1:D5', rule))],
      ['T', worksheetXml('<sheetData/>')],
    ]),
  );
  assert.deepEqual(copySheetLosses(doc.wb, 0), [{ what: 'ruleContent', count: 1 }], 'as an icon set, it holds what a copy cannot');
  // Another kind of rule now: what its icon set held is not written, and is nothing a copy lacks.
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { rule: { type: 'cellIs', operator: 'equal', formulas: ['1'] } }));
  assert.equal(ruleKeptNow(conditionalRules(doc.wb, 0)[0].rule), undefined, 'the rule states nothing beside its fields now');
  assert.deepEqual(copySheetLosses(doc.wb, 0), [], 'a copy of the sheet lacks nothing');
  const copied = copyRange(doc.wb, 0, R('D1'));
  const result = pasteAt(doc.wb, 1, R('A1'), copied.text, copied);
  assert.equal(result.skippedRules, undefined, 'a copy of the cells lacks nothing');
  apply(doc, result.ops);
  assert.equal(doc.wb.sheets[1].conditionalFormats![0].rules[0].kept, undefined, 'and the copy holds nothing of what the rule no longer states');
});

test('B6: a comment or a processing instruction is content that is not carried; uid is an identity in the revision namespaces only, <id> in the one of the extension list only', async () => {
  const noted = '<cfRule type="cellIs" dxfId="0" priority="1" operator="equal" my:a="1"><formula>1</formula><!-- a comment --><my:x>t</my:x></cfRule>';
  const inner = '<cfRule type="cellIs" dxfId="0" priority="2" operator="equal"><formula>2</formula><my:x>t<!-- inner --></my:x></cfRule>';
  const pi = '<cfRule type="cellIs" dxfId="0" priority="3" operator="equal"><formula>3</formula><?pi data?></cfRule>';
  const named = `<cfRule type="cellIs" dxfId="0" priority="4" operator="equal" my:uid="NOT-AN-IDENTITY" xr:uid="{IDENTITY}"><formula>4</formula><extLst><ext uri="{MINE}"><my:id>payload</my:id></ext><ext uri="{TIE}" xmlns:x14="${X14}"><x14:id>{TIED}</x14:id></ext></extLst></cfRule>`;
  const doc = await load(
    await book([
      ['Src', sheet(`xmlns:my="urn:mine" xmlns:xr="${XR}"`, DATA + cf('D1:D5', noted) + cf('F1:F5', inner) + cf('H1:H5', pi) + cf('J1:J5', named))],
      ['T', sheet('xmlns:my="urn:mine"', '<sheetData/>')],
    ]),
  );
  const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  assert.deepEqual(keptOf(doc.wb, 0), [
    { over: true },
    { over: true },
    { over: true },
    { attrs: [['urn:mine', 'uid', 'NOT-AN-IDENTITY']], children: [{ ns: MAIN, name: 'extLst', children: [{ ns: MAIN, name: 'ext', attrs: [['', 'uri', '{MINE}']], children: [{ ns: 'urn:mine', name: 'id', children: ['payload'] }] }] }] },
  ]);
  const copied = copyRange(doc.wb, 0, R('D1:J5'));
  const result = pasteAt(doc.wb, 1, R('A1'), copied.text, copied);
  apply(doc, result.ops);
  // Three that are not carried, and the tie of the fourth to the extension list.
  assert.equal(result.skippedRules, 4, 'what does not come is counted');
  const there = formats(await partText(await save(doc), 'xl/worksheets/sheet2.xml'));
  assert.deepEqual(there, ['<conditionalFormatting sqref="G1:G5"><cfRule type="cellIs" dxfId="0" priority="1" operator="equal" my:uid="NOT-AN-IDENTITY"><formula>4</formula><extLst><ext uri="{MINE}"><my:id>payload</my:id></ext></extLst></cfRule></conditionalFormatting>'], 'the one rule that is carried, with all but its identity and its tie');
  sameBytes(new TextEncoder().encode(await partText(await save(doc), 'xl/worksheets/sheet1.xml')), new TextEncoder().encode(await partText(doc.bytes, 'xl/worksheets/sheet1.xml')), 'the sheet they came from is as it was');
});

test('B7: what a data validation of the extension list holds beside the fields of the model is not written into the main part: its rule is not copied, and that is counted', async () => {
  const ext = (own: string) =>
    `<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="${X14}"><x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:dataValidation type="list" allowBlank="1" ${own} xr:uid="{X}" xmlns:xr="${XR}"><x14:formula1><xm:f>T!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>A1:A3</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>`;
  const own = await load(await book([['Src', worksheetXml(DATA + ext('x14:own="1" foo="2"'))], ['T', worksheetXml('<sheetData/>')]]));
  assert.deepEqual(validationKept(own.wb, 0), [{ over: true }], 'the model holds none of it');
  assert.deepEqual(copySheetLosses(own.wb, 0), [{ what: 'ruleContent', count: 1 }], 'which a copy of the sheet is said to lack');
  apply(own, copySheet(own.wb, 0, { name: 'Copy' }));
  let saved = await save(own);
  await checkPackage(saved);
  assert.equal(validations(await partText(saved, 'xl/worksheets/sheet3.xml')), '', 'the copy has no rule that says less than the one of the sheet');
  // A rule of that list that holds nothing but its identity is copied into the main part, as ever.
  const plain = await load(await book([['Src', worksheetXml(DATA + ext(''))], ['T', worksheetXml('<sheetData/>')]]));
  assert.deepEqual(copySheetLosses(plain.wb, 0), [], 'nothing is lacking');
  apply(plain, copySheet(plain.wb, 0, { name: 'Copy' }));
  saved = await save(plain);
  assert.equal(validations(await partText(saved, 'xl/worksheets/sheet3.xml')), '<dataValidations count="1"><dataValidation type="list" allowBlank="1" sqref="A1:A3"><formula1>T!$A$1:$A$3</formula1></dataValidation></dataValidations>', 'the rule in the main part of the copy');
});

test('a model put together by hand: the writer states only what the part can, also where nobody cut the rule down', async () => {
  const rule = '<cfRule type="cellIs" dxfId="0" priority="1" operator="equal" my:a="1" plain="p"><formula>1</formula></cfRule>';
  const dv = '<dataValidations count="1"><dataValidation type="whole" my:a="1" plain="p" sqref="A1:A3"><formula1>1</formula1><formula2>9</formula2></dataValidation></dataValidations>';
  const doc = await load(
    await book([
      ['Src', sheet(`${MC} mc:Ignorable="my" xmlns:my="urn:mine"`, DATA + cf('D1:D5', rule) + dv)],
      ['T', worksheetXml('<sheetData/>')],
    ]),
  );
  // The rules of the one sheet put into the other as they are, by no command.
  const from = doc.wb.sheets[0];
  const to = doc.wb.sheets[1];
  to.conditionalFormats = [{ ranges: [R('A1:A2')], rules: [{ ...from.conditionalFormats![0].rules[0], src: undefined }], live: true }];
  to.dataValidations = [{ ...from.dataValidations![0], src: undefined, ranges: [R('C1')], live: true }];
  to.conditionalModified = true;
  to.validationsModified = true;
  await writeXlsx(doc.file.package!, doc.wb);
  const saved = await doc.file.package!.save();
  await checkPackage(saved);
  const text = await partText(saved, 'xl/worksheets/sheet2.xml');
  assert.ok(text.indexOf('urn:mine') < 0 && !/\b(my|ns\d):/.test(text), `nothing of a namespace readers of this part are not let to pass over, in:\n${text}`);
  assert.ok((text.match(/plain="p"/g) || []).length === 2, `what the part can state is stated, in:\n${text}`);
  assert.equal(parseXml(text).root.name, 'worksheet', 'and the part is well-formed');
});
