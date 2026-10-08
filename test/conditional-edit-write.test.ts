// Conditional formats written into the package: a rule added to a sheet that had none,
// rules changed, deleted and reordered beside rules that stay byte for byte, the formats the
// rules give (appended to the style sheet, shared, taken back by undo), the extension list,
// sheets that are not in the file, and a file built from the model.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { buildXlsx as buildFromModel } from '../src/formats/xlsx/build';
import { open, parseRange, ZipArchive } from '../src/index';
import type { CellRange } from '../src/index';
import type { Op } from '../src/ops/ops';
import { addConditionalRule, clearConditionalRules, conditionalRules, moveConditionalRule, removeConditionalRule, setValidation, updateConditionalRule } from '../src/ops/rules';
import { copySheet, copySheetLosses } from '../src/ops/sheets';
import { buildXlsx, worksheetXml } from './fixtures';
import { BAR_ID, BLOCK_A, BLOCK_B, BLOCK_C, BLOCK_D, CELLS, DXFS, EXT_BAR, EXT_ONLY, FIRST_RULE, OTHER_EXT, SECOND_RULE, TAIL, cf, conditionalBook, conditionalContent, conditionalNow as now, extLst } from './conditional-fixtures';
import { canonical } from './workbook-content';
import { checkPackage, checkSheetOrder, load as loadBytes, partNames, partText, reopen, save, sameBytes, type Loaded } from './rules-fixtures';

const range = (ref: string): CellRange => parseRange(ref)!;
const SHEET1 = 'xl/worksheets/sheet1.xml';
const STYLES = 'xl/styles.xml';
const apply = (doc: Loaded, ops: Op[]) => doc.history.apply(ops, undefined, undefined);
const load = async (options?: Parameters<typeof conditionalBook>[0]) => loadBytes(await conditionalBook(options));
const RED = { fill: '#FFC7CE', color: '#9C0006' };
const RED_DXF = '<dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>';
/** What is read back says what the model says, rule for rule and format for format. */
const sameContent = async (bytes: Uint8Array, doc: Loaded, label: string) => assert.equal(canonical(conditionalContent(await reopen(bytes))), canonical(conditionalContent(doc.wb)), label);

test('the first rule of a sheet: the element goes where the schema has it, its format is added to the style sheet, and undo takes both back', async () => {
  const doc = await load({ bare: true });
  const original = await partText(doc.bytes, SHEET1);
  apply(doc, addConditionalRule(doc.wb, 0, range('G1:G3'), { type: 'cellIs', operator: 'lessThan', formulas: ['$H$1'], format: RED, stopIfTrue: true }));
  const saved = await save(doc);
  await checkPackage(saved);
  const xml = await partText(saved, SHEET1);
  checkSheetOrder(xml);
  const added = '<conditionalFormatting sqref="G1:G3"><cfRule type="cellIs" dxfId="3" priority="1" stopIfTrue="1" operator="lessThan"><formula>$H$1</formula></cfRule></conditionalFormatting>';
  assert.equal(xml, original.replace('<hyperlinks>', `${added}<hyperlinks>`));
  const styles = await partText(saved, STYLES);
  assert.ok(styles.includes(`<dxfs count="4">${DXFS}${RED_DXF}</dxfs>`), styles);
  sameBytes(await ZipArchive.read(saved).read('xl/workbook.xml'), await ZipArchive.read(doc.bytes).read('xl/workbook.xml'), 'the workbook part');
  await sameContent(saved, doc, 'read back');
  doc.history.undo();
  sameBytes(await save(doc), doc.bytes, 'after undo');
  doc.history.redo();
  sameBytes(await save(doc), saved, 'after redo');
});

test('rules that were not touched keep their source text; a rule whose number alone changed has only that number written anew', async () => {
  const doc = await load();
  const original = await partText(doc.bytes, SHEET1);
  apply(doc, addConditionalRule(doc.wb, 0, range('G1:G3'), { type: 'duplicateValues', format: RED }));
  const saved = await save(doc);
  await checkPackage(saved);
  const xml = await partText(saved, SHEET1);
  checkSheetOrder(xml);
  const bump = (text: string) => text.replace(/priority=(["'])(\d+)\1/g, (_, quote, n) => `priority=${quote}${Number(n) + 1}${quote}`);
  const added = '<conditionalFormatting sqref="G1:G3"><cfRule type="duplicateValues" dxfId="3" priority="1"/></conditionalFormatting>';
  assert.equal(xml, original.replace(BLOCK_A, bump(BLOCK_A)).replace(BLOCK_B, bump(BLOCK_B)).replace(BLOCK_C, bump(BLOCK_C)).replace(BLOCK_D, bump(BLOCK_D) + added));
  await sameContent(saved, doc, 'read back');
  doc.history.undo();
  sameBytes(await save(doc), doc.bytes, 'after undo');
});

test('a rule taken out of its element, a rule deleted, an element left empty; what the source of a changed rule holds beside the model is kept', async () => {
  const doc = await load();
  const original = await partText(doc.bytes, SHEET1);
  // The first rule of the element (number 2) gets other cells: it leaves the element, with its unknown attribute.
  apply(doc, updateConditionalRule(doc.wb, 0, 1, { ranges: range('A2:A4') }));
  let xml = await partText(await save(doc), SHEET1);
  const moved = '<conditionalFormatting sqref="A2:A4"><cfRule type="cellIs" dxfId="0" priority="2" operator="greaterThan" xr:uid="{R1}"><formula>$H$1</formula></cfRule></conditionalFormatting>';
  assert.equal(xml, original.replace(FIRST_RULE, '').replace(BLOCK_D, BLOCK_D + moved.replace('<conditionalFormatting ', '<conditionalFormatting xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" ')));
  await sameContent(await save(doc), doc, 'a rule moved out');
  // The other rule of the element is deleted: the element goes. The icon set goes; the time period stays in its element.
  apply(doc, removeConditionalRule(doc.wb, 0, 0));
  apply(doc, removeConditionalRule(doc.wb, 0, 3));
  const saved = await save(doc);
  await checkPackage(saved);
  xml = await partText(saved, SHEET1);
  checkSheetOrder(xml);
  assert.ok(!xml.includes(SECOND_RULE) && !xml.includes('type="iconSet"'));
  assert.ok(xml.includes('<conditionalFormatting sqref="D1:D5"><cfRule type="timePeriod" dxfId="2" priority="4" timePeriod="today"><formula>FLOOR(D1,1)=TODAY()</formula></cfRule></conditionalFormatting>'), xml);
  assert.deepEqual(now(await reopen(saved)), ['A2:A4|cellIs|1|$H$1|greaterThan', 'B1:B5|colorScale|2||min;$H$2;max', 'C1:C5|dataBar|3|', 'D1:D5|timePeriod|4|FLOOR(D1,1)=TODAY()']);
  await sameContent(saved, doc, 'rules deleted');
  // The data bar and its counterpart in the extension list are as they were.
  assert.ok(xml.includes(BLOCK_C.replace('priority="4"', 'priority="3"')) && xml.includes(extLst(EXT_BAR)));
  for (let i = 0; i < 3; i++) doc.history.undo();
  sameBytes(await save(doc), doc.bytes, 'everything undone');
});

test('formats are shared: the same format is one record, and a format the file has is used as it is', async () => {
  const doc = await load();
  apply(doc, addConditionalRule(doc.wb, 0, range('G1'), { type: 'duplicateValues', format: RED }));
  apply(doc, addConditionalRule(doc.wb, 0, range('G2'), { type: 'uniqueValues', format: RED }));
  // What the first record of the file says, and what the second says.
  apply(doc, addConditionalRule(doc.wb, 0, range('G3'), { type: 'containsBlanks', format: { fill: '#FF0000' } }));
  apply(doc, addConditionalRule(doc.wb, 1, range('G4'), { type: 'containsErrors', format: { bold: true } }));
  apply(doc, addConditionalRule(doc.wb, 1, range('G5'), { type: 'top10', format: { italic: true, underline: true, strike: true, bold: false } }));
  const saved = await save(doc);
  const styles = await partText(saved, STYLES);
  assert.ok(styles.includes(`<dxfs count="5">${DXFS}${RED_DXF}<dxf><font><b val="0"/><i/><strike/><u/></font></dxf></dxfs>`), styles);
  const sheet1 = await partText(saved, SHEET1);
  assert.ok(sheet1.includes('<cfRule type="containsBlanks" dxfId="0" priority="1">') && sheet1.includes('<cfRule type="uniqueValues" dxfId="3" priority="2"/>') && sheet1.includes('<cfRule type="duplicateValues" dxfId="3" priority="3"/>'), sheet1);
  const sheet2 = await partText(saved, 'xl/worksheets/sheet2.xml');
  assert.ok(sheet2.includes('<cfRule type="top10" dxfId="4" priority="1" rank="10"/>') && sheet2.includes('<cfRule type="containsErrors" dxfId="1" priority="2">'), sheet2);
  await sameContent(saved, doc, 'read back');
  // One of the two rules that share the record is taken back: the record stays for the other.
  doc.history.undo();
  doc.history.undo();
  doc.history.undo();
  doc.history.undo();
  assert.ok((await partText(await save(doc), STYLES)).includes(`<dxfs count="4">${DXFS}${RED_DXF}</dxfs>`));
  doc.history.undo();
  sameBytes(await save(doc), doc.bytes, 'everything undone');
});

test('a format that is changed keeps what its record holds beside font colours and fill', async () => {
  const bytes = await buildXlsx({ sheets: [{ name: 'Data', xml: worksheetXml(CELLS + cf('A1:A5', '<cfRule type="cellIs" dxfId="2" priority="1" operator="equal"><formula>5</formula></cfRule>')) }], styles: (await partText(await conditionalBook(), STYLES)) });
  const doc = await loadBytes(bytes);
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { rule: { type: 'cellIs', operator: 'equal', formulas: ['5'], format: { fill: '#FFC7CE', bold: true } } }));
  const saved = await save(doc);
  const styles = await partText(saved, STYLES);
  // Italic went (the dialog says so), bold came; the size, the alignment and the border are those of the record it started from.
  const changed = '<dxf><font><b/><sz val="14"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill><alignment horizontal="center"/><border><left style="thin"><color rgb="FF000000"/></left></border></dxf>';
  assert.ok(styles.includes(`<dxfs count="4">${DXFS}${changed}</dxfs>`), styles);
  assert.ok((await partText(saved, SHEET1)).includes('<cfRule type="cellIs" dxfId="3" priority="1" operator="equal"><formula>5</formula></cfRule>'));
  await sameContent(saved, doc, 'read back');
});

test('a file without a style sheet, and one without differential formats, get what the rules need', async () => {
  for (const styles of [undefined, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><tableStyles count="0"/></styleSheet>']) {
    const doc = await loadBytes(await buildXlsx({ sheets: [{ name: 'Data', xml: worksheetXml(CELLS) }], ...(styles === undefined ? {} : { styles }) }));
    apply(doc, addConditionalRule(doc.wb, 0, range('A1:A3'), { type: 'cellIs', operator: 'greaterThan', formulas: ['2'], format: RED }));
    const saved = await save(doc);
    await checkPackage(saved);
    const written = await partText(saved, STYLES);
    assert.ok(written.includes(`<dxfs count="1">${RED_DXF}</dxfs>`), written);
    // Where the schema has the table: behind the cell formats, before the table styles.
    if (styles !== undefined) assert.ok(written.indexOf('<dxfs') > written.indexOf('</cellXfs>') && written.indexOf('<dxfs') < written.indexOf('<tableStyles'), written);
    await sameContent(saved, doc, 'read back');
    doc.history.undo();
    const back = await save(doc);
    sameBytes(back, doc.bytes, 'after undo');
    assert.deepEqual(partNames(back), partNames(doc.bytes));
  }
});

test('rules of the file and rules changed since are both where the view has them, after rows, columns and sheets moved', async () => {
  const doc = await load();
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 2 }]);
  apply(doc, updateConditionalRule(doc.wb, 0, 1, { ranges: range('A3:A6') }));
  apply(doc, addConditionalRule(doc.wb, 0, range('G1:G3'), { type: 'expression', formulas: ['AND(G1<>$H$3,Lists!$A$1>0)'], format: RED }));
  apply(doc, moveConditionalRule(doc.wb, 0, 0, 3));
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'col', at: 0, count: 1 }]);
  apply(doc, [{ type: 'insert', sheet: 1, axis: 'row', at: 0, count: 4 }]);
  apply(doc, [{ type: 'renameSheet', sheet: 1, name: 'Sources' }]);
  apply(doc, [{ type: 'delete', sheet: 0, axis: 'row', at: 0, count: 1 }]);
  apply(doc, clearConditionalRules(doc.wb, 0, range('C2')));
  const expected = [
    'B2:B11 G2:G4|expression|1|AND(B2>C2,Sources!$A$5<>B2)',
    'B2:B5|cellIs|2|$I$2|greaterThan',
    'C3:C6|colorScale|3||min;$I$3;max',
    'H1:H2|expression|4|AND(H1<>$I$2,Sources!$A$5>0)',
    'D2:D6|dataBar|5|',
    'E2:E6|iconSet|6|',
    'E2:E6|timePeriod|7|FLOOR(E2,1)=TODAY()',
  ];
  assert.deepEqual(now(doc.wb), expected);
  const saved = await save(doc);
  await checkPackage(saved);
  checkSheetOrder(await partText(saved, SHEET1));
  assert.deepEqual(now(await reopen(saved)), expected);
  await sameContent(saved, doc, 'read back');
  sameBytes(await save(doc), saved, 'second save');
});

test('a sheet with a rule only the extension list has: that list is not touched', async () => {
  const doc = await load({ extensionOnly: true });
  apply(doc, addConditionalRule(doc.wb, 0, range('G1'), { type: 'duplicateValues', format: RED }));
  apply(doc, removeConditionalRule(doc.wb, 0, 1));
  const saved = await save(doc);
  const xml = await partText(saved, SHEET1);
  assert.ok(xml.includes(extLst(EXT_BAR + EXT_ONLY)));
  assert.ok(xml.includes(`\n  \n  ${SECOND_RULE}\n`), 'the rule that stays is the source text');
  await sameContent(saved, doc, 'read back');
});

test('a sheet that is not in the file is written with its rules, and so is a copy of a sheet', async () => {
  const doc = await load();
  apply(doc, [{ type: 'insertSheet', sheet: 3, name: 'New' }]);
  apply(doc, addConditionalRule(doc.wb, 3, range('A1:A3'), { type: 'cellIs', operator: 'equal', formulas: ['Data!$H$1'], format: RED }));
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 }]);
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { stopIfTrue: false }));
  // The copy has the rules, as they apply now. What the extension list says about the data bar is left behind, and said.
  assert.deepEqual(copySheetLosses(doc.wb, 0).filter((loss) => /conditional/i.test(loss.what)), [{ what: 'conditionalFormatExtensions', count: 1 }]);
  apply(doc, copySheet(doc.wb, 0, { name: 'Copy' }));
  assert.deepEqual(now(doc.wb, 1), now(doc.wb, 0));
  const saved = await save(doc);
  await checkPackage(saved);
  const back = await reopen(saved);
  assert.deepEqual(back.sheets.map((sheet) => sheet.name), ['Data', 'Copy', 'Lists', 'Other', 'New']);
  for (let i = 0; i < 5; i++) assert.deepEqual(now(back, i), now(doc.wb, i), `sheet ${i}`);
  assert.deepEqual(now(back, 4), ['A1:A3|cellIs|1|Data!$H$2|equal']);
  // A rule that names the copied sheet names the copy.
  assert.equal(now(back, 1)[0], 'A2:A11 F2:F4|expression|1|AND(A2>B2,Lists!$A$1<>A2)');
  // The copy goes on by itself.
  apply(doc, [{ type: 'insert', sheet: 1, axis: 'row', at: 0, count: 1 }]);
  apply(doc, removeConditionalRule(doc.wb, 1, 0));
  const more = await reopen(await save(doc));
  for (let i = 0; i < 5; i++) assert.deepEqual(now(more, i), now(doc.wb, i), `sheet ${i}, second save`);
  assert.equal(now(more, 1)[0], 'A3:A12 F3:F5|cellIs|1|$H$3|greaterThan');
  assert.equal(now(more, 0)[1], 'A2:A11 F2:F4|cellIs|2|$H$2|greaterThan');
});

test('a file built from the model has the rules where they apply now, with the formats they give', async () => {
  const doc = await load();
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 2 }]);
  apply(doc, addConditionalRule(doc.wb, 0, range('G1:G3'), { type: 'expression', formulas: ['Lists!$A$1>G1'], format: RED }));
  apply(doc, updateConditionalRule(doc.wb, 0, 2, { rule: { type: 'cellIs', operator: 'lessThan', formulas: ['$H$3'], format: { fill: '#00FF00' } } }));
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'col', at: 0, count: 1 }]);
  apply(doc, [{ type: 'renameSheet', sheet: 1, name: 'Sources' }]);
  const built = await buildFromModel(doc.wb);
  await checkPackage(built);
  checkSheetOrder(await partText(built, SHEET1));
  const back = await (await open(built)).workbook();
  assert.deepEqual(now(back), now(doc.wb));
  assert.equal(now(back)[0], 'H1:H3|expression|1|Sources!$A$1>H1');
  // What is read back gives the same formats (the bar without what the extension list said about it, as before).
  const plain = (content: unknown[]) => canonical(content).replace(/"extended":true,"gradient"/g, '"extended":false,"gradient"');
  assert.equal(conditionalRules(back, 0).length, conditionalRules(doc.wb, 0).length);
  assert.deepEqual(conditionalRules(back, 0).map((info) => canonical(info.format)), conditionalRules(doc.wb, 0).map((info) => canonical(info.format)));
  void plain;
});

test('a part that writes the main namespace with a prefix, and an empty list of validations, take new rules in the same spelling', async () => {
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const xml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<x:worksheet xmlns:x="${NS}">\r\n<x:sheetData><x:row r="1"><x:c r="A1"><x:v>5</x:v></x:c></x:row></x:sheetData>\r\n` +
    '<x:conditionalFormatting sqref="A1:A3"><x:cfRule type="cellIs" priority="1" operator="equal"><x:formula>5</x:formula></x:cfRule></x:conditionalFormatting>\r\n<x:dataValidations count="0"/>\r\n<x:pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></x:worksheet>';
  const doc = await loadBytes(await buildXlsx({ sheets: [{ name: 'Data', xml }] }));
  apply(doc, addConditionalRule(doc.wb, 0, range('B1:B3'), { type: 'top10', rank: 3, format: RED }));
  apply(doc, updateConditionalRule(doc.wb, 0, 1, { stopIfTrue: true }));
  apply(doc, setValidation(doc.wb, 0, range('C1'), { type: 'list', formula1: '"a,b"' }));
  const saved = await save(doc);
  await checkPackage(saved);
  const written = await partText(saved, SHEET1);
  assert.ok(written.includes('<x:conditionalFormatting sqref="B1:B3"><x:cfRule type="top10" dxfId="0" priority="1" rank="3"/></x:conditionalFormatting>'), written);
  assert.ok(written.includes('<x:conditionalFormatting sqref="A1:A3"><x:cfRule type="cellIs" priority="2" stopIfTrue="1" operator="equal"><x:formula>5</x:formula></x:cfRule></x:conditionalFormatting>'), written);
  assert.ok(written.includes('<x:dataValidations count="1"><x:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="C1"><x:formula1>"a,b"</x:formula1></x:dataValidation></x:dataValidations>'), written);
  checkSheetOrder(written);
  await sameContent(saved, doc, 'read back');
  assert.equal((await reopen(saved)).sheets[0].dataValidations!.length, 1);
  for (let i = 0; i < 3; i++) doc.history.undo();
  sameBytes(await save(doc), doc.bytes, 'everything undone');
});

// ---------------------------------------------------------------------------
// Review: what the reader does not take into the model is not the writer's to drop

test('elements the model never held (no rule, a range that is none, a rule that cannot be read) stay byte for byte when other rules change', async () => {
  const EMPTY = '<conditionalFormatting sqref="J1:J5" pivot="1"></conditionalFormatting>';
  const NOWHERE = '<conditionalFormatting sqref="XFE1:XFE5"><cfRule type="cellIs" priority="7" operator="equal"><formula>1</formula></cfRule></conditionalFormatting>';
  const UNREAD = '<dataValidation type="whole" sqref="XFE9"><formula1>1</formula1><formula2>2</formula2></dataValidation>';
  const GOOD = '<dataValidation type="whole" sqref="A1:A3"><formula1>1</formula1><formula2>2</formula2></dataValidation>';
  const xml = CELLS + EMPTY + cf('A1:A3', '<cfRule type="cellIs" priority="1" operator="equal"><formula>5</formula></cfRule>') + NOWHERE + `<dataValidations count="3">${UNREAD}${GOOD}${UNREAD.replace('XFE9', 'XFE8')}</dataValidations>` + TAIL;
  const doc = await loadBytes(await buildXlsx({ sheets: [{ name: 'Data', xml: worksheetXml(xml) }] }));
  assert.equal(doc.wb.sheets[0].conditionalFormats!.length, 1);
  assert.equal(doc.wb.sheets[0].dataValidations!.length, 1);
  apply(doc, removeConditionalRule(doc.wb, 0, 0));
  apply(doc, addConditionalRule(doc.wb, 0, range('B1'), { type: 'duplicateValues' }));
  apply(doc, setValidation(doc.wb, 0, range('A1:A3'), undefined));
  apply(doc, setValidation(doc.wb, 0, range('C1'), { type: 'list', formula1: '"a,b"' }));
  const saved = await save(doc);
  await checkPackage(saved);
  const written = await partText(saved, SHEET1);
  assert.ok(written.includes(EMPTY), written);
  assert.ok(written.includes(NOWHERE), written);
  assert.ok(written.includes(UNREAD) && written.includes(UNREAD.replace('XFE9', 'XFE8')), written);
  assert.ok(!written.includes(GOOD));
  // The count is the number of elements that are there.
  assert.ok(written.includes('<dataValidations count="3">'), written);
  checkSheetOrder(written);
  for (let i = 0; i < 4; i++) doc.history.undo();
  sameBytes(await save(doc), doc.bytes, 'everything undone');
});

test('a style sheet that writes the main namespace with a prefix gets its new records in the same spelling', async () => {
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<x:styleSheet xmlns:x="${NS}"><x:fonts count="1"><x:font><x:sz val="11"/><x:name val="Calibri"/></x:font></x:fonts><x:fills count="1"><x:fill><x:patternFill patternType="none"/></x:fill></x:fills><x:borders count="1"><x:border/></x:borders><x:cellStyleXfs count="1"><x:xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></x:cellStyleXfs><x:cellXfs count="1"><x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></x:cellXfs><x:dxfs count="1"><x:dxf><x:font><x:i/><x:sz val="14"/></x:font><x:border><x:left style="thin"/></x:border></x:dxf></x:dxfs></x:styleSheet>`;
  const doc = await loadBytes(await buildXlsx({ sheets: [{ name: 'Data', xml: worksheetXml(CELLS + cf('A1:A3', '<cfRule type="cellIs" dxfId="0" priority="1" operator="equal"><formula>5</formula></cfRule>')) }], styles }));
  apply(doc, addConditionalRule(doc.wb, 0, range('B1:B3'), { type: 'duplicateValues', format: RED }));
  // The record of the file, with a fill added: what it holds beside that stays, in its own spelling.
  apply(doc, updateConditionalRule(doc.wb, 0, 1, { format: { italic: true, fill: '#C6EFCE' } }));
  const saved = await save(doc);
  await checkPackage(saved);
  const written = await partText(saved, STYLES);
  assert.ok(
    written.includes(
      '<x:dxfs count="3"><x:dxf><x:font><x:i/><x:sz val="14"/></x:font><x:border><x:left style="thin"/></x:border></x:dxf>' +
        '<x:dxf><x:font><x:color rgb="FF9C0006"/></x:font><x:fill><x:patternFill><x:bgColor rgb="FFFFC7CE"/></x:patternFill></x:fill></x:dxf>' +
        '<x:dxf><x:font><x:i/><x:sz val="14"/></x:font><x:fill><x:patternFill><x:bgColor rgb="FFC6EFCE"/></x:patternFill></x:fill><x:border><x:left style="thin"/></x:border></x:dxf></x:dxfs>',
    ),
    written,
  );
  await sameContent(saved, doc, 'read back');
  doc.history.undo();
  doc.history.undo();
  sameBytes(await save(doc), doc.bytes, 'everything undone');
});

test('a style sheet has room for so many differential formats: a rule that would need one more is refused', async () => {
  const doc = await load({ bare: true });
  const record = doc.wb.styles.dxfs[0];
  while (doc.wb.styles.dxfs.length < 65536) doc.wb.styles.dxfs.push(record);
  assert.throws(() => addConditionalRule(doc.wb, 0, range('A1'), { type: 'duplicateValues', format: RED }), /too many formats/);
  // A rule that gives no format of its own needs no record.
  assert.equal(addConditionalRule(doc.wb, 0, range('A1'), { type: 'duplicateValues' }).length, 1);
  assert.equal(addConditionalRule(doc.wb, 0, range('A1'), { type: 'colorScale', colors: ['#FF0000', '#00FF00'] }).length, 1);
});

// ---------------------------------------------------------------------------
// Review: a data bar with a counterpart in the extension list

test('a data bar of Excel 2010 is deleted with its counterpart in the extension list; the rest of that list stays byte for byte', async () => {
  for (const options of [{}, { otherExt: true }, { extensionOnly: true }, { extensionOnly: true, otherExt: true }]) {
    const doc = await load(options);
    const original = await partText(doc.bytes, SHEET1);
    apply(doc, removeConditionalRule(doc.wb, 0, 3));
    const saved = await save(doc);
    await checkPackage(saved);
    const written = await partText(saved, SHEET1);
    checkSheetOrder(written);
    assert.ok(!written.includes(BAR_ID) && !written.includes('type="dataBar"'), written);
    // What is left of the list: the rule only it has, whatever else it holds - or nothing at all.
    const left = options.extensionOnly ? extLst(EXT_ONLY) : '';
    const expected = options.otherExt ? (left === '' ? `<extLst>${OTHER_EXT}</extLst>` : left.replace('</extLst>', `${OTHER_EXT}</extLst>`)) : left;
    assert.equal(written.slice(written.indexOf('<pageMargins')).replace(/^<pageMargins[^>]*\/>/, ''), `${expected}</worksheet>`);
    const bump = (text: string) => (options.extensionOnly ? text : text.replace(/priority="(\d+)"/g, (_, n) => `priority="${Number(n) - 1}"`));
    assert.ok(written.includes(bump(BLOCK_D)) && written.includes(BLOCK_A) && written.includes(BLOCK_B));
    await sameContent(saved, doc, 'read back');
    assert.ok(original.includes(BAR_ID));
    doc.history.undo();
    sameBytes(await save(doc), doc.bytes, 'after undo');
  }
});

test('a data bar of Excel 2010 that is put on other cells takes its counterpart along, and both follow the rows that move', async () => {
  const doc = await load({ extensionOnly: true, otherExt: true });
  apply(doc, updateConditionalRule(doc.wb, 0, 3, { ranges: [range('C2:C9'), range('G1')] }));
  let saved = await save(doc);
  await checkPackage(saved);
  let written = await partText(saved, SHEET1);
  checkSheetOrder(written);
  // The main rule in an element of its own, with the id that ties it to the other one; that one with the same cells.
  assert.ok(written.includes(`<conditionalFormatting sqref="C2:C9 G1"><cfRule type="dataBar" priority="4"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:id>${BAR_ID}</x14:id></ext></extLst></cfRule></conditionalFormatting>`), written);
  assert.ok(written.includes(EXT_BAR.replace('<xm:sqref>C1:C5</xm:sqref>', '<xm:sqref>C2:C9 G1</xm:sqref>')), written);
  assert.ok(written.includes(EXT_ONLY) && written.includes(OTHER_EXT), 'the rest of the list is the source text');
  await sameContent(saved, doc, 'read back');
  const bar = conditionalRules(await reopen(saved), 0)[3];
  assert.deepEqual([bar.rule.dataBar!.extended, bar.rule.dataBar!.minLength, bar.rule.dataBar!.maxLength], [true, 0, 100], 'still the bar the extension list describes');

  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 2 }]);
  apply(doc, clearConditionalRules(doc.wb, 0, range('C4')));
  saved = await save(doc);
  written = await partText(saved, SHEET1);
  assert.ok(written.includes('<xm:sqref>C5:C11 G3</xm:sqref>') && written.includes('<xm:sqref>E3:E7</xm:sqref>'), written);
  await sameContent(saved, doc, 'after rows moved');
  sameBytes(await save(doc), saved, 'second save');
  // Every cell of it deleted: both halves go; undone, both are back.
  apply(doc, [{ type: 'delete', sheet: 0, axis: 'col', at: 2, count: 1 }]);
  apply(doc, [{ type: 'delete', sheet: 0, axis: 'row', at: 2, count: 1 }]);
  saved = await save(doc);
  written = await partText(saved, SHEET1);
  // (The sparkline of the other extension stood in the deleted row and went with it, as it always did.)
  assert.ok(!written.includes(BAR_ID) && written.includes('{DA7ABA51-AAAA-BBBB-0001-000000000002}'), written);
  await sameContent(saved, doc, 'after its cells went');
  while (doc.history.canUndo) doc.history.undo();
  assert.deepEqual(now(await reopen(await save(doc))), now(doc.wb));
  assert.ok((await partText(await save(doc), SHEET1)).includes(extLst(EXT_BAR + EXT_ONLY).replace('</extLst>', `${OTHER_EXT}</extLst>`)));
});

test('the number of a rule that the extension list repeats is written there as well', async () => {
  const doc = await load({ barPriority: true });
  const original = await partText(doc.bytes, SHEET1);
  apply(doc, addConditionalRule(doc.wb, 0, range('G1'), { type: 'duplicateValues' }));
  const written = await partText(await save(doc), SHEET1);
  assert.ok(written.includes(`<x14:cfRule type="dataBar" priority="5" id="${BAR_ID}">`), written);
  assert.ok(written.includes(BLOCK_C.replace('priority="4"', 'priority="5"')));
  await sameContent(await save(doc), doc, 'read back');
  doc.history.undo();
  sameBytes(await save(doc), doc.bytes, 'after undo');
  // A counterpart that states no number gets none.
  const plain = await load();
  apply(plain, addConditionalRule(plain.wb, 0, range('G1'), { type: 'duplicateValues' }));
  assert.ok((await partText(await save(plain), SHEET1)).includes(extLst(EXT_BAR)));
  assert.ok(original.includes('priority="4" id='));
});
