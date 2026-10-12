// Data validations written into the package: a rule added to a sheet that had none, rules
// changed and deleted beside rules that stay byte for byte, the rules of the extension list,
// undo back to the file, sheets that are not in the file, and a file built from the model.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { buildXlsx as buildFromModel } from '../src/formats/xlsx/build';
import { open, parseRange, rangeName } from '../src/index';
import type { CellRange, DataValidation, Workbook } from '../src/index';
import type { Op } from '../src/ops/ops';
import { dataValidationNow, ruleSpace, setValidation } from '../src/ops/rules';
import { copySheet, copySheetLosses } from '../src/ops/sheets';
import { buildXlsx, worksheetXml } from './fixtures';
import { checkPackage, checkSheetOrder, load as loadBytes, partText, reopen, save, sameBytes, type Loaded } from './rules-fixtures';

const range = (ref: string): CellRange => parseRange(ref)!;
const SHEET1 = 'xl/worksheets/sheet1.xml';
const WHOLE = { type: 'whole', operator: 'between', formula1: '1', formula2: '5' } as const;
const apply = (doc: Loaded, ops: Op[]) => doc.history.apply(ops, undefined, undefined);
const now = (wb: Workbook, sheet = 0) => {
  const space = ruleSpace(wb, wb.sheets[sheet]);
  return (wb.sheets[sheet].dataValidations || [])
    .map((rule) => dataValidationNow(rule, space))
    .filter((rule): rule is DataValidation => !!rule)
    .map((rule) => `${rule.ranges.map(rangeName).join(' ')}|${rule.type}|${rule.formula1 ?? ''}|${rule.formula2 ?? ''}`);
};

/** A rule without where it is in a file and whether it was changed: what it says. */
const plain = (rule: DataValidation): DataValidation => {
  const copy = { ...rule };
  delete copy.src;
  delete copy.live;
  return copy;
};

const load = async (xml: string, more: { name: string; xml: string }[] = []) =>
  loadBytes(await buildXlsx({ sheets: [{ name: 'Data', xml: worksheetXml(xml) }, ...more.map((sheet) => ({ name: sheet.name, xml: worksheetXml(sheet.xml) }))] }));

// Written the way no writer of ours would: odd spacing, an attribute and a comment the model has no field for.
const FIRST = '<dataValidation type="whole"   showErrorMessage="1" xr:uid="{AAAA}" imeMode="off" sqref="A1:A10 C1:C3"><formula1>1</formula1><!-- kept --><formula2>9</formula2></dataValidation>';
const SECOND = "<dataValidation type='list' allowBlank='1' sqref='B1:B5'><formula1>$H$1:$H$2</formula1></dataValidation>";
const THIRD = '<dataValidation type="textLength" operator="lessThan" sqref="E1"><formula1>4</formula1></dataValidation>';
const RULES = `<dataValidations count="3" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">\n  ${FIRST}\n  ${SECOND}\n  ${THIRD}\n</dataValidations>`;
const BODY = '<sheetData><row r="1"><c r="H1"><v>1</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="J1:K1"/></mergeCells>';
const TAIL = '<hyperlinks><hyperlink ref="H1" location="Data!A1"/></hyperlinks><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>';

test('the first rule of a sheet: <dataValidations> goes where the schema has it, and nothing else of the file changes', async () => {
  const doc = await load(BODY + TAIL);
  const original = await partText(doc.bytes, SHEET1);
  apply(doc, setValidation(doc.wb, 0, range('B2:C4'), { ...WHOLE, promptTitle: 'Count', prompt: 'One to five\nplease', errorStyle: 'warning', errorTitle: 'No', error: 'Not <that> & "this"' }));
  const saved = await save(doc);
  await checkPackage(saved);
  const xml = await partText(saved, SHEET1);
  checkSheetOrder(xml);
  const added =
    '<dataValidations count="1"><dataValidation type="whole" errorStyle="warning" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="No" error="Not &lt;that> &amp; &quot;this&quot;" promptTitle="Count" prompt="One to five&#10;please" sqref="B2:C4"><formula1>1</formula1><formula2>5</formula2></dataValidation></dataValidations>';
  assert.equal(xml, original.replace('<hyperlinks>', `${added}<hyperlinks>`));
  // No cell changed: nothing is to be calculated again, and no other part is written.
  for (const name of ['xl/workbook.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels']) sameBytes(await doc.file.package!.readPart(`/${name}`), new TextEncoder().encode(await partText(doc.bytes, name)), name);
  const back = (await reopen(saved)).sheets[0].dataValidations!;
  assert.equal(back.length, 1);
  assert.deepEqual(plain(back[0]), plain(doc.wb.sheets[0].dataValidations![0]));
});

test('rules that were not touched keep their source text; the count follows; a rule that was cut keeps what the model has no field for', async () => {
  const doc = await load(BODY + RULES + TAIL);
  const original = await partText(doc.bytes, SHEET1);
  // The second rule loses a cell, the third goes, a new one comes; the first is not touched.
  apply(doc, setValidation(doc.wb, 0, range('B5'), undefined));
  apply(doc, setValidation(doc.wb, 0, range('E1'), undefined));
  apply(doc, setValidation(doc.wb, 0, range('F1:F2'), { type: 'list', formula1: '"a,b"' }));
  const saved = await save(doc);
  await checkPackage(saved);
  const xml = await partText(saved, SHEET1);
  checkSheetOrder(xml);
  const cut = '<dataValidation type="list" allowBlank="1" sqref="B1:B4"><formula1>$H$1:$H$2</formula1></dataValidation>';
  const added = '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="F1:F2"><formula1>"a,b"</formula1></dataValidation>';
  // The rule that was cut stands where its element was: the order of the rules decides which one a cell has.
  const expected = `<dataValidations count="3" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">\n  ${FIRST}\n  ${cut}\n  \n${added}</dataValidations>`;
  assert.equal(xml, original.replace(RULES, expected));
  assert.deepEqual(now(await reopen(saved)), now(doc.wb));

  // The first rule cut as well: its unknown attributes, its children and its comment are written with it.
  apply(doc, setValidation(doc.wb, 0, range('A1'), undefined));
  const again = await partText(await save(doc), SHEET1);
  assert.ok(again.includes('<dataValidation type="whole" showErrorMessage="1" xr:uid="{AAAA}" imeMode="off" sqref="A2:A10 C1:C3"><formula1>1</formula1><formula2>9</formula2><!-- kept --></dataValidation>'), again);
  assert.ok(!again.includes(FIRST));
});

test('save, undo everything, save: the archive is the file again, byte for byte', async () => {
  const doc = await load(BODY + RULES + TAIL);
  apply(doc, setValidation(doc.wb, 0, range('B5:E1'), WHOLE));
  apply(doc, setValidation(doc.wb, 0, range('M1'), { type: 'custom', formula1: 'M1>0' }));
  const changed = await save(doc);
  assert.notDeepEqual(changed, doc.bytes);
  doc.history.undo();
  doc.history.undo();
  sameBytes(await save(doc), doc.bytes, 'after undo');
  doc.history.redo();
  doc.history.redo();
  sameBytes(await save(doc), changed, 'after redo');
  // The same for a sheet that had no rule at all.
  const bare = await load(BODY + TAIL);
  apply(bare, setValidation(bare.wb, 0, range('A1'), WHOLE));
  await save(bare);
  bare.history.undo();
  sameBytes(await save(bare), bare.bytes, 'a sheet without rules, after undo');
});

test('the last rule deleted: the element goes with it', async () => {
  const doc = await load(BODY + RULES + TAIL);
  apply(doc, setValidation(doc.wb, 0, range('A1:Z100'), undefined));
  const xml = await partText(await save(doc), SHEET1);
  assert.ok(!/dataValidation/.test(xml), xml);
  checkSheetOrder(xml);
  assert.equal((await reopen(await save(doc))).sheets[0].dataValidations, undefined);
});

const X14 =
  '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
  '<x14:dataValidations xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" count="1">' +
  '<x14:dataValidation type="list" allowBlank="1" showErrorMessage="1"><x14:formula1><xm:f>Lists!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>M1:M5</xm:sqref></x14:dataValidation>' +
  '</x14:dataValidations></ext></extLst>';

test('the rules of the extension list stay as they are whatever happens to the others', async () => {
  const doc = await load(BODY + RULES + TAIL + X14, [{ name: 'Lists', xml: '<sheetData/>' }]);
  apply(doc, setValidation(doc.wb, 0, range('A1:F9'), undefined));
  apply(doc, setValidation(doc.wb, 0, range('A1'), WHOLE));
  const saved = await save(doc);
  const xml = await partText(saved, SHEET1);
  assert.ok(xml.includes(X14));
  assert.equal(xml.split('<dataValidations').length, 2, 'one list of the main part');
  assert.deepEqual(now(await reopen(saved)).sort(), now(doc.wb).sort());
  assert.equal((await reopen(saved)).sheets[0].dataValidations!.filter((rule) => rule.extended).length, 1);
});

test('rules of the file and rules changed since are both where the view has them, after rows, columns and sheets moved', async () => {
  const doc = await load(BODY + RULES + TAIL, [{ name: 'Lists', xml: '<sheetData/>' }]);
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 2 }]);
  apply(doc, setValidation(doc.wb, 0, range('A3'), { type: 'list', formula1: 'Lists!$A$1:$A$3' }));
  apply(doc, setValidation(doc.wb, 0, range('G1:G3'), { type: 'custom', formula1: 'G1<>$H$3' }));
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'col', at: 0, count: 1 }]);
  apply(doc, [{ type: 'insert', sheet: 1, axis: 'row', at: 0, count: 4 }]);
  apply(doc, [{ type: 'renameSheet', sheet: 1, name: 'Sources' }]);
  apply(doc, [{ type: 'delete', sheet: 0, axis: 'row', at: 0, count: 1 }]);
  const expected = ['B3:B11 D2:D4|whole|1|9', 'C2:C6|list|$I$2:$I$3|', 'F2|textLength|4|', 'B2|list|Sources!$A$5:$A$7|', 'H1:H2|custom|H1<>$I$2|'];
  assert.deepEqual(now(doc.wb), expected);
  const saved = await save(doc);
  await checkPackage(saved);
  checkSheetOrder(await partText(saved, SHEET1));
  assert.deepEqual(now(await reopen(saved)), expected);
  // Saved a second time from the same state: the same file.
  sameBytes(await save(doc), saved, 'second save');
});

test('a sheet that is not in the file is written with its rules, and so is a copy of a sheet', async () => {
  // (A rule with a comment in it is kept as the file has it, and is not copied: see rules-kept-review.test.ts.)
  const doc = await load(BODY + RULES.replace('<!-- kept -->', '') + TAIL, [{ name: 'Lists', xml: '<sheetData/>' }]);
  apply(doc, [{ type: 'insertSheet', sheet: 2, name: 'New' }]);
  apply(doc, setValidation(doc.wb, 2, range('A1:A3'), { type: 'list', formula1: 'Data!$H$1:$H$2' }));
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 }]);
  // The copy of a sheet has its rules, as they apply now; nothing of them is left behind.
  assert.deepEqual(copySheetLosses(doc.wb, 0).filter((loss) => loss.what === 'dataValidations'), []);
  apply(doc, copySheet(doc.wb, 0, { name: 'Copy' }));
  assert.deepEqual(now(doc.wb, 1), now(doc.wb, 0));
  assert.deepEqual(now(doc.wb, 1), ['A2:A11 C2:C4|whole|1|9', 'B2:B6|list|$H$2:$H$3|', 'E2|textLength|4|']);
  const saved = await save(doc);
  await checkPackage(saved);
  const back = await reopen(saved);
  assert.deepEqual(back.sheets.map((sheet) => sheet.name), ['Data', 'Copy', 'Lists', 'New']);
  for (let i = 0; i < 4; i++) assert.deepEqual(now(back, i), now(doc.wb, i), `sheet ${i}`);
  assert.deepEqual(now(back, 3), ['A1:A3|list|Data!$H$2:$H$3|']);
  // The copy goes on by itself: rows that move in it move its rules, not those of the sheet it was copied from.
  apply(doc, [{ type: 'insert', sheet: 1, axis: 'row', at: 0, count: 1 }]);
  assert.deepEqual(now(doc.wb, 1), ['A3:A12 C3:C5|whole|1|9', 'B3:B7|list|$H$3:$H$4|', 'E3|textLength|4|']);
  assert.deepEqual(now(doc.wb, 0), ['A2:A11 C2:C4|whole|1|9', 'B2:B6|list|$H$2:$H$3|', 'E2|textLength|4|']);
  const more = await reopen(await save(doc));
  for (let i = 0; i < 4; i++) assert.deepEqual(now(more, i), now(doc.wb, i), `sheet ${i}, second save`);
});

test('a file built from the model has the rules where they apply now', async () => {
  const doc = await load(BODY + RULES + TAIL, [{ name: 'Lists', xml: '<sheetData/>' }]);
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 2 }]);
  apply(doc, setValidation(doc.wb, 0, range('A3'), { type: 'list', formula1: 'Lists!$A$1:$A$3' }));
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'col', at: 0, count: 1 }]);
  apply(doc, [{ type: 'renameSheet', sheet: 1, name: 'Sources' }]);
  apply(doc, [{ type: 'insert', sheet: 1, axis: 'row', at: 0, count: 1 }]);
  const expected = ['B4:B12 D3:D5|whole|1|9', 'C3:C7|list|$I$3:$I$4|', 'F3|textLength|4|', 'B3|list|Sources!$A$2:$A$4|'];
  assert.deepEqual(now(doc.wb), expected);
  const built = await buildFromModel(doc.wb);
  await checkPackage(built);
  const xml = await partText(built, SHEET1);
  checkSheetOrder(xml);
  assert.equal(xml.split('<dataValidations').length, 2);
  assert.ok(xml.includes('<dataValidations count="4">'), xml);
  assert.deepEqual(now(await (await open(built)).workbook()), expected);
});

// ---------------------------------------------------------------------------
// Review: a rule that is edited where it is stays the element it was

test('a rule of the file that is edited on the very cells it has keeps its place, its source and what the model has no field for', async () => {
  const LIST = '<dataValidation type="list" imeMode="off" xr:uid="{BBBB}" allowBlank="1" sqref="B1:B5"><formula1>"a, b,c"</formula1></dataValidation>';
  const rules = `<dataValidations count="3" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">${FIRST}${LIST}${THIRD}</dataValidations>`;
  const doc = await load(BODY + rules + TAIL);
  const original = await partText(doc.bytes, SHEET1);
  const shown = doc.wb.sheets[0].dataValidations![1];
  // The same rule on the same cells is no step, however its list is spelled.
  assert.deepEqual(setValidation(doc.wb, 0, range('B1:B5'), { type: 'list', formula1: '"a, b,c"', allowBlank: true, showInputMessage: false, showErrorMessage: false }), []);
  apply(doc, setValidation(doc.wb, 0, range('B1:B5'), { type: 'list', formula1: '"a, b,c"', allowBlank: true, showInputMessage: false, showErrorMessage: true, errorTitle: 'No' }));
  const list = doc.wb.sheets[0].dataValidations!;
  assert.deepEqual(list.map((rule) => [rule.src, rule.live === true]), [[0, false], [1, true], [2, false]], 'in its place, with its source');
  assert.equal(list[1].formula1, shown.formula1);
  const saved = await save(doc);
  const changed = '<dataValidation type="list" allowBlank="1" showErrorMessage="1" errorTitle="No" imeMode="off" xr:uid="{BBBB}" sqref="B1:B5"><formula1>"a, b,c"</formula1></dataValidation>';
  assert.equal(await partText(saved, SHEET1), original.replace(LIST, changed));
  doc.history.undo();
  sameBytes(await save(doc), doc.bytes, 'after undo');
});
