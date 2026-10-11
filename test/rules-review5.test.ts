// Second review of the rules that go with cells:
// - a rule of a kind that looks at its own cell (a text rule, blanks, errors, a period of
//   dates) names that cell in its formula however it spells it ("$C1"): the formula follows
//   the first cell of the rule where the rule gets other cells, is copied, or loses the
//   column it stood in;
// - a rule that is pasted into another sheet takes along what its element holds beside the
//   fields of the model, but no identity of the element it was copied from;
// - of two rules the model holds alike, the one the cells were copied from gets the cells.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { parseRange, type CellRange } from '../src/index';
import { copyRange, cutRange, pasteAt } from '../src/ops/commands';
import { clearConditionalRules, conditionalRules, updateConditionalRule } from '../src/ops/rules';
import { deleteSheet, insertSheet } from '../src/ops/sheets';
import { buildXlsx as buildFromModel } from '../src/formats/xlsx/build';
import { open } from '../src/index';
import { BAR_ID, BLOCK_C, EXT_BAR, STYLES, cf, extLst } from './conditional-fixtures';
import { buildXlsx, worksheetXml } from './fixtures';
import { checkPackage, load, partText, reopen, save, sameBytes, type Loaded } from './rules-fixtures';

const R = (text: string): CellRange => parseRange(text) as CellRange;
const formats = (xml: string) => xml.match(/<conditionalFormatting[\s\S]*?<\/conditionalFormatting>/g) || [];
const unescaped = (text: string) => text.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&amp;/g, '&');
/** The formats of a sheet part as "sqref | formula; formula" lines, in the order of the part. */
const said = (xml: string) =>
  formats(xml).map((text) => `${(/sqref="([^"]*)"/.exec(text) || ['', ''])[1]} | ${(text.match(/<formula>[\s\S]*?<\/formula>/g) || []).map((f) => unescaped(f.slice(9, -10))).join('; ')}`);
const apply = (doc: Loaded, ops: Parameters<Loaded['history']['apply']>[0]) => doc.history.apply(ops, undefined, undefined);
const sheetXml = async (doc: Loaded, n = 1) => partText(await save(doc), `xl/worksheets/sheet${n}.xml`);
const modelSaid = (doc: Loaded, sheet = 0) => conditionalRules(doc.wb, sheet).map((info) => `${info.ranges.map((r) => `${r.top},${r.left}-${r.bottom},${r.right}`).join(' ')} | ${info.rule.formulas.join('; ')}`);

const rule = (attrs: string, formula: string) => `<cfRule ${attrs}><formula>${formula.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</formula></cfRule>`;
const oneSheet = (body: string, more: string[] = []) =>
  buildXlsx({
    sheets: [{ name: 'S', xml: worksheetXml('<sheetData><row r="1"><c r="C1"><v>1</v></c></row></sheetData>' + body) }].concat(more.map((name) => ({ name, xml: worksheetXml('<sheetData/>') }))),
    styles: STYLES,
  });

const TODAY = rule('type="timePeriod" dxfId="0" priority="1" timePeriod="today"', 'FLOOR($C1,1)=TODAY()');
const CONTAINS = rule('type="containsText" dxfId="1" priority="2" operator="containsText" text="x"', 'NOT(ISERROR(SEARCH("x",$E1)))');

test('a rule that looks at its own cell names the new first cell when it is given other cells, with the "$" it had', async () => {
  const doc = await load(await oneSheet(cf('C1:C9', TODAY) + cf('E1:E9', CONTAINS)));
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { ranges: R('G1:G9') }));
  apply(doc, updateConditionalRule(doc.wb, 0, 1, { ranges: R('H2:H5') }));
  assert.deepEqual(said(await sheetXml(doc)).sort(), ['G1:G9 | FLOOR($G1,1)=TODAY()', 'H2:H5 | NOT(ISERROR(SEARCH("x",$H2)))'], 'the formulas of the saved file');
  assert.deepEqual(modelSaid(await load(await save(doc))), modelSaid(doc), 'the file reads back as the model');
});

test('every kind of rule that looks at its own cell follows it, in each of the four spellings', async () => {
  const kinds: [attrs: string, formula: (cell: string) => string][] = [
    ['type="timePeriod" timePeriod="last7Days"', (c) => `AND(TODAY()-FLOOR(${c},1)<=6,FLOOR(${c},1)<=TODAY())`],
    ['type="timePeriod" timePeriod="thisMonth"', (c) => `AND(MONTH(${c})=MONTH(TODAY()),YEAR(${c})=YEAR(TODAY()))`],
    ['type="containsText" operator="containsText" text="a&quot;b"', (c) => `NOT(ISERROR(SEARCH("a""b",${c})))`],
    ['type="notContainsText" operator="notContains" text="x"', (c) => `ISERROR(SEARCH("x",${c}))`],
    ['type="beginsWith" operator="beginsWith" text="x"', (c) => `LEFT(${c},LEN("x"))="x"`],
    ['type="endsWith" operator="endsWith" text="x"', (c) => `RIGHT(${c},LEN("x"))="x"`],
    ['type="containsBlanks"', (c) => `LEN(TRIM(${c}))=0`],
    ['type="notContainsBlanks"', (c) => `LEN(TRIM(${c}))>0`],
    ['type="containsErrors"', (c) => `ISERROR(${c})`],
    ['type="notContainsErrors"', (c) => `NOT(ISERROR(${c}))`],
  ];
  const spellings: [from: string, to: string][] = [
    ['C2', 'F7'],
    ['$C2', '$F7'],
    ['C$2', 'F$7'],
    ['$C$2', '$F$7'],
  ];
  for (const [attrs, formula] of kinds) {
    for (const [from, to] of spellings) {
      const doc = await load(await oneSheet(cf('C2:C9', rule(`${attrs} dxfId="0" priority="1"`, formula(from)))));
      apply(doc, updateConditionalRule(doc.wb, 0, 0, { ranges: R('F7:F9') }));
      assert.deepEqual(said(await sheetXml(doc)), [`F7:F9 | ${formula(to)}`], `${attrs}, written with ${from}`);
    }
  }
});

test('a formula that is not the one of the kind, or that names another cell, moves as any formula does', async () => {
  // Written by hand: "$C1" is an anchored column there.
  let doc = await load(await oneSheet(cf('C1:C9', rule('type="timePeriod" dxfId="0" priority="1" timePeriod="today"', 'FLOOR($C1,1)=TODAY()+0'))));
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { ranges: R('G1:G9') }));
  assert.deepEqual(said(await sheetXml(doc)), ['G1:G9 | FLOOR($C1,1)=TODAY()+0'], 'a formula of its own');
  // The formula of the kind, for a cell that is not the first of the rule.
  doc = await load(await oneSheet(cf('C1:C9', rule('type="timePeriod" dxfId="0" priority="1" timePeriod="today"', 'FLOOR($B1,1)=TODAY()'))));
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { ranges: R('G1:G9') }));
  assert.deepEqual(said(await sheetXml(doc)), ['G1:G9 | FLOOR($B1,1)=TODAY()'], 'another cell');
  // An expression that happens to say what a text rule says is a formula.
  doc = await load(await oneSheet(cf('E1:E9', rule('type="expression" dxfId="0" priority="1"', 'NOT(ISERROR(SEARCH("x",$E1)))'))));
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { ranges: R('G1:G9') }));
  assert.deepEqual(said(await sheetXml(doc)), ['G1:G9 | NOT(ISERROR(SEARCH("x",$E1)))'], 'an expression');
});

test('cells that are cut and pasted elsewhere, or copied to another sheet, take such a rule along for the cell it is on there', async () => {
  let doc = await load(await oneSheet(cf('C1:C9', TODAY) + cf('E1:E9', CONTAINS), ['T']));
  let copied = copyRange(doc.wb, 0, R('C1:E9'));
  apply(doc, cutRange(doc.wb, 0, R('C1:E9')));
  apply(doc, pasteAt(doc.wb, 0, R('J3'), copied.text, copied).ops);
  assert.deepEqual(said(await sheetXml(doc)).sort(), ['J3:J11 | FLOOR($J3,1)=TODAY()', 'L3:L11 | NOT(ISERROR(SEARCH("x",$L3)))'], 'moved');

  doc = await load(await oneSheet(cf('C1:C9', TODAY) + cf('E1:E9', CONTAINS), ['T']));
  copied = copyRange(doc.wb, 0, R('C2:E3'));
  apply(doc, pasteAt(doc.wb, 1, R('J5'), copied.text, copied).ops);
  assert.deepEqual(said(await sheetXml(doc, 2)).sort(), ['J5:J6 | FLOOR($J5,1)=TODAY()', 'L5:L6 | NOT(ISERROR(SEARCH("x",$L5)))'], 'copied to another sheet');
  assert.deepEqual(said(await sheetXml(doc, 1)), ['C1:C9 | FLOOR($C1,1)=TODAY()', 'E1:E9 | NOT(ISERROR(SEARCH("x",$E1)))'], 'the sheet they were copied from');

  // In the sheet itself the rule that is there gets the cells, and goes on naming its first cell.
  doc = await load(await oneSheet(cf('C1:C9', TODAY) + cf('E1:E9', CONTAINS)));
  copied = copyRange(doc.wb, 0, R('C2:E3'));
  apply(doc, pasteAt(doc.wb, 0, R('J5'), copied.text, copied).ops);
  assert.deepEqual(said(await sheetXml(doc)).sort(), ['C1:C9 J5:J6 | FLOOR($C1,1)=TODAY()', 'E1:E9 L5:L6 | NOT(ISERROR(SEARCH("x",$E1)))'], 'copied beside');
});

test('a rule that loses the cells its formula named, by a cut or a clearing, names the first cell it has left', async () => {
  const doc = await load(await oneSheet(cf('C1:C9', rule('type="timePeriod" dxfId="0" priority="1" timePeriod="today"', 'FLOOR($C$1,1)=TODAY()'))));
  apply(doc, clearConditionalRules(doc.wb, 0, R('C1:C2')));
  assert.deepEqual(said(await sheetXml(doc)), ['C3:C9 | FLOOR($C$3,1)=TODAY()'], 'cleared');
  doc.history.undo();
  apply(doc, cutRange(doc.wb, 0, R('A1:D4')));
  assert.deepEqual(said(await sheetXml(doc)), ['C5:C9 | FLOOR($C$5,1)=TODAY()'], 'cut');
});

const TWO_RANGES = (cell: string) => cf('C3:C9 A1:A5', rule('type="timePeriod" dxfId="2" priority="6" timePeriod="last7Days"', `AND(TODAY()-FLOOR(${cell},1)<=6,FLOOR(${cell},1)<=TODAY())`));
const LAST7 = (cell: string) => `AND(TODAY()-FLOOR(${cell},1)<=6,FLOOR(${cell},1)<=TODAY())`;

test('the column of the first range deleted: the rule of the file names the first cell of the range that is left, not #REF!', async () => {
  for (const [from, to] of [['C3', 'A1'], ['$C3', '$A1'], ['C$3', 'A$1'], ['$C$3', '$A$1']]) {
    const doc = await load(await oneSheet(TWO_RANGES(from)));
    apply(doc, [{ type: 'delete', sheet: 0, axis: 'col', at: 2, count: 1 }]);
    const saved = await save(doc);
    assert.deepEqual(said(await partText(saved, 'xl/worksheets/sheet1.xml')), [`A1:A5 | ${LAST7(to)}`], `${from}: the saved file`);
    assert.deepEqual(conditionalRules(doc.wb, 0).map((info) => info.rule.formulas), [[LAST7(to)]], `${from}: the model`);
    assert.deepEqual(modelSaid(await load(saved)), modelSaid(doc), `${from}: the file reads back as the model`);
    // A file built from the model says the same.
    const built = await (await open(await buildFromModel(doc.wb))).workbook();
    assert.deepEqual(conditionalRules(built, 0).map((info) => info.rule.formulas), [[LAST7(to)]], `${from}: a file built from the model`);
    doc.history.undo();
    sameBytes(await save(doc), doc.bytes, `${from}: undone, the file is the one that was opened`);
  }
});

test('the same for a rule that was changed since it was read, and for rows', async () => {
  const doc = await load(await oneSheet(TWO_RANGES('$C3')));
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { stopIfTrue: true }));
  apply(doc, [{ type: 'delete', sheet: 0, axis: 'col', at: 2, count: 1 }]);
  assert.deepEqual(said(await sheetXml(doc)), [`A1:A5 | ${LAST7('$A1')}`], 'a live rule, its column deleted');
  doc.history.undo();
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 2 }]);
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'col', at: 1, count: 1 }]);
  assert.deepEqual(said(await sheetXml(doc)), [`D5:D11 A3:A7 | ${LAST7('$D5')}`], 'a live rule, rows and a column put in');

  // The rows the rule starts in deleted, in a rule written with an anchored row.
  const rows = await load(await oneSheet(cf('C1:C2 E5:E9', rule('type="containsBlanks" dxfId="0" priority="1"', 'LEN(TRIM(C$1))=0'))));
  apply(rows, [{ type: 'delete', sheet: 0, axis: 'row', at: 0, count: 2 }]);
  assert.deepEqual(said(await sheetXml(rows)), ['E3:E7 | LEN(TRIM(E$3))=0'], 'rows deleted');
  assert.deepEqual(conditionalRules(rows.wb, 0).map((info) => info.rule.formulas), [['LEN(TRIM(E$3))=0']], 'rows deleted, the model');
});

test('a formula of its own in a rule of such a kind still becomes #REF! when its column goes', async () => {
  const doc = await load(await oneSheet(cf('C3:C9 A1:A5', rule('type="timePeriod" dxfId="2" priority="6" timePeriod="today"', 'FLOOR($C3,1)=TODAY()+0'))));
  apply(doc, [{ type: 'delete', sheet: 0, axis: 'col', at: 2, count: 1 }]);
  assert.deepEqual(said(await sheetXml(doc)), ['A1:A5 | FLOOR(#REF!,1)=TODAY()+0'], 'the saved file');
});

// ---------------------------------------------------------------------------
// What a pasted rule takes along

const XR2 = ' xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"';
const ICON_RULE =
  '<cfRule type="iconSet" priority="5" xr2:uid="{ICON}" kept="yes"><iconSet iconSet="3Arrows" percent="0"><cfvo type="num" val="1" foo="bar"><extLst><ext uri="{KEEP-1}"/></extLst></cfvo><cfvo type="formula" val="$H$1" gte="0"/><cfvo type="percent" val="67"/></iconSet></cfRule>';
const SCALE_RULE = '<cfRule type="colorScale" priority="6"><colorScale><cfvo type="min" foo="low"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale><extLst><ext uri="{KEEP-2}"/></extLst></cfRule>';

const threeSheets = () =>
  buildXlsx({
    sheets: [
      { name: 'Data', xml: worksheetXml('<sheetData><row r="1"><c r="D1"><v>5</v></c><c r="H1"><v>1</v></c></row></sheetData>' + cf('D1:D5 G2:G4', ICON_RULE, XR2) + cf('F1:F3', SCALE_RULE)) },
      { name: 'Lists', xml: worksheetXml('<sheetData/>') },
      { name: 'Calc', xml: worksheetXml('<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>') },
    ],
    styles: STYLES,
  });

const KEPT_ICON = '<iconSet iconSet="3Arrows" percent="0"><cfvo type="num" val="1" foo="bar"><extLst><ext uri="{KEEP-1}"/></extLst></cfvo><cfvo type="formula" val="$H$1" gte="0"/><cfvo type="percent" val="67"/></iconSet>';

test('a rule copied into another sheet keeps what its element holds beside the fields of the model, without the identity of the element', async () => {
  const doc = await load(await threeSheets());
  const copied = copyRange(doc.wb, 0, R('D1:F1'));
  const result = pasteAt(doc.wb, 2, R('J1'), copied.text, copied);
  apply(doc, result.ops);
  assert.equal(result.skippedRules, undefined, 'nothing of the rules is left behind');
  const saved = await save(doc);
  await checkPackage(saved);
  const there = formats(await partText(saved, 'xl/worksheets/sheet3.xml')).join('\n');
  assert.ok(there.indexOf(KEPT_ICON) >= 0, `the icon set as the file states it, in:\n${there}`);
  assert.ok(/<cfRule type="iconSet"[^>]* kept="yes"/.test(there), `an attribute of the rule the model has no field for, in:\n${there}`);
  assert.ok(there.indexOf('<cfvo type="min" foo="low"/>') >= 0 && there.indexOf('<extLst><ext uri="{KEEP-2}"/></extLst></cfRule>') >= 0, `the scale as the file states it, in:\n${there}`);
  assert.ok(there.indexOf('uid') < 0 && there.indexOf('{ICON}') < 0 && there.indexOf('xr2') < 0, `no identity of the element it was copied from, in:\n${there}`);
  assert.deepEqual(formats(await partText(saved, 'xl/worksheets/sheet1.xml')), formats(await partText(doc.bytes, 'xl/worksheets/sheet1.xml')), 'the sheet it was copied from is as it was');
  // Read back, and saved again after another change: it stays.
  const back = await load(saved);
  assert.deepEqual(conditionalRules(back.wb, 2).map((info) => info.rule.type).sort(), ['colorScale', 'iconSet'], 'the rules of the file that was saved');
  apply(doc, updateConditionalRule(doc.wb, 2, 0, { stopIfTrue: true }));
  const again = formats(await sheetXml(doc, 3)).join('\n');
  assert.ok(again.indexOf(KEPT_ICON) >= 0 && again.indexOf('foo="low"') >= 0, `after a change of one of the rules, in:\n${again}`);
  while (doc.history.canUndo) doc.history.undo();
  sameBytes(await save(doc), doc.bytes, 'everything undone, the file is the one that was opened');
});

test('cells cut and pasted into another sheet move their rule with what its element holds; the sheet they came from may go', async () => {
  const doc = await load(await threeSheets());
  const copied = copyRange(doc.wb, 0, R('D1:G5'));
  apply(doc, cutRange(doc.wb, 0, R('D1:G5')));
  const result = pasteAt(doc.wb, 2, R('B2'), copied.text, copied);
  apply(doc, result.ops);
  assert.equal(result.skippedRules, undefined, 'nothing of the rules is left behind');
  let saved = await save(doc);
  await checkPackage(saved);
  assert.deepEqual(formats(await partText(saved, 'xl/worksheets/sheet1.xml')), [], 'the sheet the cells were cut from has no rule left');
  let there = formats(await partText(saved, 'xl/worksheets/sheet3.xml')).join('\n');
  assert.ok(/sqref="B2:B6 E3:E5"/.test(there) && there.indexOf(KEPT_ICON) >= 0, `the icon set where the cells went, in:\n${there}`);
  assert.ok(there.indexOf('foo="low"') >= 0 && there.indexOf('{KEEP-2}') >= 0, `the scale where the cells went, in:\n${there}`);
  assert.ok(there.indexOf('{ICON}') < 0, `no identity, in:\n${there}`);
  // The sheet the rule was read from is deleted: what was taken from it stays.
  apply(doc, deleteSheet(doc.wb, 0));
  saved = await save(doc);
  await checkPackage(saved);
  const calc = (await reopen(saved)).sheets.findIndex((sheet) => sheet.name === 'Calc');
  there = formats(await partText(saved, `xl/worksheets/sheet3.xml`)).join('\n');
  assert.ok(calc >= 0 && there.indexOf(KEPT_ICON) >= 0 && there.indexOf('foo="low"') >= 0, `after the sheet it came from was deleted, in:\n${there}`);
  // And in a second save, when the part it was read from is no part of the package any more.
  apply(doc, updateConditionalRule(doc.wb, 1, 0, { stopIfTrue: true }));
  saved = await save(doc);
  await checkPackage(saved);
  there = formats(await partText(saved, `xl/worksheets/sheet3.xml`)).join('\n');
  assert.ok(there.indexOf(KEPT_ICON) >= 0 && there.indexOf('foo="low"') >= 0 && there.indexOf('stopIfTrue="1"') >= 0, `saved again, in:\n${there}`);
  // Undone as far as the cut: the rules are those of the file again, where the file has them.
  while (doc.history.canUndo) doc.history.undo();
  sameBytes(await save(doc), doc.bytes, 'everything undone, the file is the one that was opened');
});

test('a rule pasted into a sheet that was added keeps the same, and a second paste of it there gives the rule the cells', async () => {
  const doc = await load(await threeSheets());
  apply(doc, insertSheet(doc.wb, 3, { name: 'New' }));
  const copied = copyRange(doc.wb, 0, R('D1'));
  apply(doc, pasteAt(doc.wb, 3, R('A1'), copied.text, copied).ops);
  const again = copyRange(doc.wb, 3, R('A1'));
  apply(doc, pasteAt(doc.wb, 3, R('C4'), again.text, again).ops);
  const saved = await save(doc);
  await checkPackage(saved);
  const back = await reopen(saved);
  const at = back.sheets.findIndex((sheet) => sheet.name === 'New');
  const there = formats(await partText(saved, back.sheets[at].part.slice(1)));
  assert.equal(there.length, 1, `one rule for both cells:\n${there.join('\n')}`);
  assert.ok(/sqref="A1 C4"/.test(there[0]) && there[0].indexOf(KEPT_ICON) >= 0 && there[0].indexOf('uid') < 0, `the rule as the file states it, in:\n${there[0]}`);
});

test('what a copied rule holds in a namespace of its own goes along when its element declared it, and when the other sheet declares it the same way', async () => {
  const MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
  const sheet = (rootAttrs: string, inner: string) => worksheetXml(inner).replace('<worksheet ', `<worksheet ${rootAttrs} `);
  const RULE = '<cfRule type="iconSet" priority="5" my:mark="1" plain="1"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0" my:x="1"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"><my:more/></cfvo></iconSet></cfRule>';
  const PLAIN = '<conditionalFormatting sqref="A1"><cfRule type="iconSet" priority="1" plain="1"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule></conditionalFormatting>';
  const FULL = '<conditionalFormatting sqref="A1"><cfRule type="iconSet" priority="1" my:mark="1" plain="1"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0" my:x="1"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"><my:more/></cfvo></iconSet></cfRule></conditionalFormatting>';
  const cases: [what: string, source: string, target: string, expected: string][] = [
    ['declared on the root of the sheet it came from only', `xmlns:my="urn:example:mine" ${MC} mc:Ignorable="my"`, '', PLAIN],
    ['declared for another namespace in the sheet it goes to', 'xmlns:my="urn:example:mine"', 'xmlns:my="urn:example:other"', PLAIN],
    ['passed over by readers of the one sheet, not of the other', `xmlns:my="urn:example:mine" ${MC} mc:Ignorable="my"`, 'xmlns:my="urn:example:mine"', PLAIN],
    ['declared the same way in both', `xmlns:my="urn:example:mine" ${MC} mc:Ignorable="my"`, `xmlns:my="urn:example:mine" ${MC} mc:Ignorable="x my"`, FULL],
  ];
  for (const [what, source, target, expected] of cases) {
    const doc = await load(
      await buildXlsx({
        sheets: [
          { name: 'Data', xml: sheet(source, '<sheetData/>' + cf('D1:D5', RULE)) },
          { name: 'Calc', xml: sheet(target, '<sheetData/>') },
        ],
        styles: STYLES,
      }),
    );
    const copied = copyRange(doc.wb, 0, R('D1'));
    apply(doc, pasteAt(doc.wb, 1, R('A1'), copied.text, copied).ops);
    const saved = await save(doc);
    await checkPackage(saved);
    assert.deepEqual(formats(await partText(saved, 'xl/worksheets/sheet2.xml')), [expected], what);
    // In the sheet it came from every namespace is at hand.
    const beside = await load(doc.bytes);
    const again = copyRange(beside.wb, 0, R('D1:D5'));
    apply(beside, cutRange(beside.wb, 0, R('D1:D5')));
    apply(beside, pasteAt(beside.wb, 0, R('F1'), again.text, again).ops);
    assert.deepEqual(formats(await sheetXml(beside)), [FULL.replace('sqref="A1"', 'sqref="F1:F5"')], `${what}: moved within the sheet`);
  }
});

test('a namespace declared on the element a copied rule stood in goes with the rule', async () => {
  const bytes = await buildXlsx({
    sheets: [
      { name: 'Data', xml: worksheetXml('<sheetData/>' + cf('D1:D5', '<cfRule type="iconSet" priority="5" my:mark="1" plain="1"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0" my:x="1"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"><my:more/></cfvo></iconSet></cfRule>', ' xmlns:my="urn:example:mine"')) },
      { name: 'Calc', xml: worksheetXml('<sheetData/>') },
    ],
    styles: STYLES,
  });
  const doc = await load(bytes);
  const copied = copyRange(doc.wb, 0, R('D1'));
  apply(doc, pasteAt(doc.wb, 1, R('A1'), copied.text, copied).ops);
  const saved = await save(doc);
  await checkPackage(saved);
  const there = formats(await partText(saved, 'xl/worksheets/sheet2.xml')).join('\n');
  // The declaration stands on the element the rule was in: it goes with the rule, and so does what uses it.
  assert.ok(/<conditionalFormatting xmlns:my="urn:example:mine" sqref="A1">/.test(there) && there.indexOf('my:mark="1"') >= 0 && there.indexOf('<my:more/>') >= 0 && there.indexOf('plain="1"') >= 0, `declared where the rule stood, in:\n${there}`);
  assert.equal((await reopen(saved)).sheets[1].conditionalFormats!.length, 1, 'the file reads back');
});

test('a data bar the extension list says more about is pasted into another sheet without the tie to its counterpart, and that is counted', async () => {
  const doc = await load(
    await buildXlsx({
      sheets: [
        { name: 'Data', xml: worksheetXml('<sheetData><row r="1"><c r="C1"><v>1</v></c></row></sheetData>' + BLOCK_C + extLst(EXT_BAR)) },
        { name: 'Calc', xml: worksheetXml('<sheetData/>') },
      ],
      styles: STYLES,
    }),
  );
  const copied = copyRange(doc.wb, 0, R('C1:C5'));
  const result = pasteAt(doc.wb, 1, R('E1'), copied.text, copied);
  apply(doc, result.ops);
  assert.equal(result.skippedRules, 1, 'what the extension list says of the bar does not come along');
  const saved = await save(doc);
  await checkPackage(saved);
  const there = await partText(saved, 'xl/worksheets/sheet2.xml');
  assert.deepEqual(formats(there), ['<conditionalFormatting sqref="E1:E5"><cfRule type="dataBar" priority="1"><dataBar minLength="0" maxLength="100"><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule></conditionalFormatting>'], 'the bar as the main part can state it (with the lengths the extension list gave it)');
  assert.ok(there.indexOf(BAR_ID) < 0 && there.indexOf('extLst') < 0, `no id of a counterpart the sheet does not have, in:\n${there}`);
  assert.ok((await partText(saved, 'xl/worksheets/sheet1.xml')).indexOf(`<x14:id>${BAR_ID}</x14:id>`) >= 0, 'the bar it was copied from keeps its counterpart');
  assert.equal(conditionalRules(await reopen(saved), 1).length, 1, 'the file reads back');
});

test('of two rules the model holds alike, the one the cells were copied from gets the cells', async () => {
  const A = '<cfRule type="iconSet" priority="2"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0" foo="FROM-A"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule>';
  const C = '<cfRule type="iconSet" priority="1" xr:uid="{FROM-C}" xmlns:xr="urn:x"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule>';
  const bytes = await buildXlsx({ sheets: [{ name: 'S', xml: worksheetXml('<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' + cf('A1:A5', A) + cf('C1:C5', C)) }], styles: STYLES });
  let doc = await load(bytes);
  let copied = copyRange(doc.wb, 0, R('A1'));
  apply(doc, pasteAt(doc.wb, 0, R('E1'), copied.text, copied).ops);
  let there = formats(await sheetXml(doc));
  assert.equal(there.length, 2, `two rules as before:\n${there.join('\n')}`);
  assert.ok(there.some((text) => /sqref="A1:A5 E1"/.test(text) && text.indexOf('FROM-A') >= 0), `the rule of A has the cell:\n${there.join('\n')}`);
  assert.ok(there.some((text) => /sqref="C1:C5"/.test(text) && text.indexOf('{FROM-C}') >= 0), `the rule of C is as it was:\n${there.join('\n')}`);

  doc = await load(bytes);
  copied = copyRange(doc.wb, 0, R('C1'));
  apply(doc, pasteAt(doc.wb, 0, R('E1'), copied.text, copied).ops);
  there = formats(await sheetXml(doc));
  assert.ok(there.some((text) => /sqref="C1:C5 E1"/.test(text) && text.indexOf('{FROM-C}') >= 0) && there.some((text) => /sqref="A1:A5"/.test(text) && text.indexOf('FROM-A') >= 0), `copied from C, the rule of C has the cell:\n${there.join('\n')}`);

  // Pasted over the cells of the other rule: those cells get the rule they are copies of, with what it holds.
  doc = await load(bytes);
  copied = copyRange(doc.wb, 0, R('A1:A5'));
  apply(doc, pasteAt(doc.wb, 0, R('C1'), copied.text, copied).ops);
  there = formats(await sheetXml(doc));
  assert.equal(there.length, 1, `one rule is left:\n${there.join('\n')}`);
  assert.ok(/sqref="A1:A5 C1:C5"/.test(there[0]) && there[0].indexOf('FROM-A') >= 0 && there[0].indexOf('FROM-C') < 0, `the rule of A on both columns:\n${there[0]}`);
});

// ---------------------------------------------------------------------------
// The press that asks for the context menu

test('Ctrl and the main button is the press for the context menu on a device of Apple, and a click anywhere else', async () => {
  const { isContextPress } = await import('../src/ui/rules-dialog');
  assert.equal(isContextPress({ button: 0, ctrlKey: true }, true), true, 'Ctrl and the main button, Apple');
  assert.equal(isContextPress({ button: 0, ctrlKey: true }, false), false, 'Ctrl and the main button, elsewhere');
  assert.equal(isContextPress({ button: 0, ctrlKey: false }, true), false, 'a plain click, Apple');
  assert.equal(isContextPress({ button: 2, ctrlKey: true }, true), false, 'another button is not this press (it is told by its button)');
  // Node has no navigator.platform of Apple's to go by unless it runs on one: the default follows the platform.
  const apple = typeof navigator !== 'undefined' && /Mac|iPhone|iPad|iPod/.test(navigator.platform || '');
  assert.equal(isContextPress({ button: 0, ctrlKey: true }), apple, 'by the platform of the page');
});
