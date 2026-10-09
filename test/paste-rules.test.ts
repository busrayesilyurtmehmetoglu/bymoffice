// Rules that travel with their cells: a copy inside the editor takes the data
// validation and the conditional formats of the copied cells along, a paste puts
// them on the cells it writes (in place of the rules those had), a cut takes
// them away. What the writer makes of it, and what undo gives back.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { parseRange, rangeName } from '../src/index';
import type { CellRange, Workbook } from '../src/index';
import { copyRange, cutRange, pasteAt } from '../src/ops/commands';
import type { Op } from '../src/ops/ops';
import { addConditionalRule, conditionalFormatNow, dataValidationNow, ruleSpace, setValidation } from '../src/ops/rules';
import { BLOCK_A, BLOCK_B, BLOCK_C, BLOCK_D, CELLS, EXT_BAR, EXT_ONLY, STYLES, cf, conditionalNow, extLst } from './conditional-fixtures';
import { buildXlsx, worksheetXml } from './fixtures';
import { checkPackage, checkSheetOrder, load as loadBytes, partText, reopen, save, sameBytes, type Loaded } from './rules-fixtures';

const range = (ref: string): CellRange => parseRange(ref)!;
const SHEET1 = 'xl/worksheets/sheet1.xml';
const SHEET2 = 'xl/worksheets/sheet2.xml';
const apply = (doc: Loaded, ops: Op[]) => doc.history.apply(ops, undefined, undefined);

/** The data validations of a sheet as they apply now: "ranges|type|formula1|formula2", with "|x14" for one of the extension list. */
const validations = (wb: Workbook, sheet = 0) => {
  const space = ruleSpace(wb, wb.sheets[sheet]);
  return (wb.sheets[sheet].dataValidations || [])
    .map((rule) => dataValidationNow(rule, space))
    .filter((rule) => !!rule)
    .map((rule) => `${rule!.ranges.map(rangeName).join(' ')}|${rule!.type}|${rule!.formula1 ?? ''}|${rule!.formula2 ?? ''}${rule!.extended ? '|x14' : ''}`);
};

/** The conditional formats of a sheet as they apply now, without the place of each rule in the order: "ranges|type|formulas". */
const formats = (wb: Workbook, sheet = 0) =>
  conditionalNow(wb, sheet).map((text) => {
    const fields = text.split('|');
    return `${fields[0]}|${fields[1]}|${fields[3]}`;
  });

const dv = (attrs: string, inner = '') => `<dataValidation ${attrs}>${inner}</dataValidation>`;
const WHOLE = dv('type="whole"   showErrorMessage="1" imeMode="off" sqref="A1:A3"', '<formula1>1</formula1><!-- kept --><formula2>B1*2</formula2>');
const LIST = dv('type="list" allowBlank="1" sqref="C1:C2"', '<formula1>$H$1:$H$2</formula1>');
const CUSTOM = dv('type="custom" sqref="E5"', "<formula1>AND(E5&gt;$H$1,Lists!$A1&lt;&gt;D5)</formula1>");
const X14_LIST =
  '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dataValidations xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" count="1">' +
  '<x14:dataValidation type="list" allowBlank="1"><x14:formula1><xm:f>Lists!$B$1:$B$3</xm:f></x14:formula1><xm:sqref>L1:L3</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>';
const DATA = '<sheetData><row r="1"><c r="A1"><v>3</v></c><c r="B1"><v>30</v></c><c r="H1"><v>1</v></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>10</v></c><c r="H2"><v>2</v></c></row><row r="3"><c r="A3"><v>2</v></c><c r="B3"><v>20</v></c></row></sheetData>';
const RULES = `<dataValidations count="3">${WHOLE}${LIST}${CUSTOM}</dataValidations>`;
const LISTS = { name: 'Lists', xml: worksheetXml('<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>4</v></c></row></sheetData>') };

const validationBook = (extension = false, lists = LISTS) => buildXlsx({ sheets: [{ name: 'Data', xml: worksheetXml(DATA + RULES + (extension ? X14_LIST : '')) }, lists] });

const conditionalBook = (blocks: string, ext = '') =>
  buildXlsx({ sheets: [{ name: 'Data', xml: worksheetXml(CELLS + blocks + (ext === '' ? '' : extLst(ext))) }, LISTS], styles: STYLES });

/** What the saved file says, read back: the rules of the model, where the model has them. */
async function savedAsModel(doc: Loaded, label: string): Promise<Uint8Array> {
  const saved = await save(doc);
  await checkPackage(saved);
  const back = await reopen(saved);
  for (let s = 0; s < doc.wb.sheets.length; s++) {
    checkSheetOrder(await partText(saved, `xl/worksheets/sheet${s + 1}.xml`), `${label}: sheet ${s + 1}`);
    assert.deepEqual(validations(back, s), validations(doc.wb, s), `${label}: validations of sheet ${s + 1} read back`);
    assert.deepEqual(conditionalNow(back, s), conditionalNow(doc.wb, s), `${label}: conditional formats of sheet ${s + 1} read back`);
  }
  return saved;
}

const rulesOf = (xml: string) => (xml.match(/<conditionalFormatting[\s\S]*?<\/conditionalFormatting>|<dataValidations[\s\S]*?<\/dataValidations>|<extLst>[\s\S]*<\/extLst>/g) || []).join('\n');

// ---------------------------------------------------------------------------
// Data validation

test('a copy takes the validation of its cells along, cut down to them, and a paste gives it to the cells it writes', async () => {
  const doc = await loadBytes(await validationBook());
  const copied = copyRange(doc.wb, 0, range('A2:C2'));
  assert.deepEqual(
    copied.rules!.validations.map((item) => [item.ranges.map(rangeName).join(' '), item.anchor, item.rule.type, item.rule.formula1, item.rule.formula2]),
    [
      // Written for A2 now: the relative reference moved with the first cell, the anchored ones did not.
      ['A1', { row: 1, col: 0 }, 'whole', '1', 'B2*2'],
      ['C1', { row: 1, col: 2 }, 'list', '$H$1:$H$2', undefined],
    ],
  );
  assert.deepEqual(copied.rules!.left, []);
  assert.equal((copied.rules!.validations[0].rule as { src?: number }).src, undefined, 'a copied rule is not the element of the file');

  const result = pasteAt(doc.wb, 0, range('E8'), copied.text, copied);
  assert.equal(result.skippedRules, undefined);
  assert.equal(result.ops.filter((op) => op.type === 'setDataValidations').length, 1);
  apply(doc, result.ops);
  // The rules the sheet has get the cells: no second rule says the same.
  assert.deepEqual(validations(doc.wb), ['A1:A3 E8|whole|1|B1*2', 'C1:C2 G8|list|$H$1:$H$2|', 'E5|custom|AND(E5>$H$1,Lists!$A1<>D5)|']);
  const saved = await savedAsModel(doc, 'pasted');
  const xml = await partText(saved, SHEET1);
  // The rule that got the cells is still the element it was read from: an attribute the model has no field for stays.
  assert.ok(xml.indexOf('<dataValidation type="whole" showErrorMessage="1" imeMode="off" sqref="A1:A3 E8"><formula1>1</formula1><formula2>B1*2</formula2></dataValidation>') > 0, rulesOf(xml));

  // One undo step, and the file is the file again.
  assert.equal(doc.history.undo() !== undefined, true);
  assert.equal(doc.history.canUndo, false);
  assert.equal(rulesOf(await partText(await save(doc), SHEET1)), rulesOf(await partText(doc.bytes, SHEET1)), 'undone');
  doc.history.redo();
  assert.deepEqual(validations(doc.wb), ['A1:A3 E8|whole|1|B1*2', 'C1:C2 G8|list|$H$1:$H$2|', 'E5|custom|AND(E5>$H$1,Lists!$A1<>D5)|']);
});

test('the cells that are written lose their own validation, also where the copy has none', async () => {
  const doc = await loadBytes(await validationBook());
  // B2 has no rule: pasted over A1:A2, it leaves them without one. The rule keeps A3, and is written for it.
  const bare = copyRange(doc.wb, 0, range('B2'));
  assert.equal(bare.rules, undefined);
  apply(doc, pasteAt(doc.wb, 0, range('A1:A2'), bare.text, bare).ops);
  assert.deepEqual(validations(doc.wb), ['A3|whole|1|B3*2', 'C1:C2|list|$H$1:$H$2|', 'E5|custom|AND(E5>$H$1,Lists!$A1<>D5)|']);
  // A rule of another kind takes the place of the one that was there.
  const custom = copyRange(doc.wb, 0, range('E5'));
  apply(doc, pasteAt(doc.wb, 0, range('C2'), custom.text, custom).ops);
  assert.deepEqual(validations(doc.wb), ['A3|whole|1|B3*2', 'C1|list|$H$1:$H$2|', 'E5 C2|custom|AND(E5>$H$1,Lists!$A1<>D5)|']);
  // Seen from C2 the rule reads the cells it read from E5, moved by the distance of the paste (and round the edge of the sheet).
  await savedAsModel(doc, 'replaced');
  // The same text from elsewhere (no copy of ours) changes no rule.
  const before = validations(doc.wb);
  apply(doc, pasteAt(doc.wb, 0, range('A3'), bare.text).ops);
  assert.deepEqual(validations(doc.wb), before);
  while (doc.history.canUndo) doc.history.undo();
  const rules = async (bytes: Uint8Array) => rulesOf(await partText(bytes, SHEET1));
  assert.equal(await rules(await save(doc)), await rules(doc.bytes));
});

test('relative references move by the distance of the paste, anchored ones stay, names of sheets stay', async () => {
  const doc = await loadBytes(await validationBook());
  const copied = copyRange(doc.wb, 0, range('E5'));
  apply(doc, pasteAt(doc.wb, 0, range('G7'), copied.text, copied).ops);
  // One rule for both cells: seen from G7 (two down, two right of E5) its formula is AND(G7>$H$1,Lists!$A3<>F7).
  assert.deepEqual(validations(doc.wb).slice(-1), ['E5 G7|custom|AND(E5>$H$1,Lists!$A1<>D5)|']);
  // Pasted where the rule cannot be the same one, the formula is written out for the new cell.
  apply(doc, setValidation(doc.wb, 0, range('E5'), undefined));
  assert.deepEqual(validations(doc.wb).slice(-1), ['G7|custom|AND(G7>$H$1,Lists!$A3<>F7)|']);
  // Towards the top-left a relative reference goes round the edge of the sheet, as it does when the rule is evaluated.
  apply(doc, pasteAt(doc.wb, 0, range('A1'), copied.text, copied).ops);
  assert.deepEqual(validations(doc.wb), ['A2:A3|whole|1|B2*2', 'C1:C2|list|$H$1:$H$2|', 'G7 A1|custom|AND(G7>$H$1,Lists!$A3<>F7)|']);
  const seen = dataValidationNow({ ...doc.wb.sheets[0].dataValidations![2] }, ruleSpace(doc.wb, doc.wb.sheets[0]))!;
  assert.equal(seen.ranges[1].top, 0);
  await savedAsModel(doc, 'moved');
});

test('a paste onto another sheet: the rule is a rule of that sheet, and the sheets it names are the ones it named', async () => {
  const doc = await loadBytes(await validationBook());
  const copied = copyRange(doc.wb, 0, range('A1:E5'));
  assert.equal(copied.rules!.validations.length, 3);
  const result = pasteAt(doc.wb, 1, range('B2'), copied.text, copied);
  assert.deepEqual(result.range, range('B2:F6'));
  apply(doc, result.ops);
  assert.deepEqual(validations(doc.wb, 1), ['B2:B4|whole|1|C2*2', 'D2:D3|list|$H$1:$H$2|', 'F6|custom|AND(F6>$H$1,Lists!$A2<>E6)|']);
  assert.deepEqual(validations(doc.wb, 0).length, 3, 'the sheet that was copied from keeps its rules');
  const saved = await savedAsModel(doc, 'other sheet');
  sameBytes(new TextEncoder().encode(await partText(saved, SHEET1)), new TextEncoder().encode(await partText(doc.bytes, SHEET1)), 'the sheet that was copied from');
  assert.ok((await partText(saved, SHEET2)).indexOf('<dataValidations count="3">') > 0, 'three rules in the other sheet');
  doc.history.undo();
  assert.equal(doc.history.canUndo, false);
  for (const part of [SHEET1, SHEET2]) assert.equal(rulesOf(await partText(await save(doc), part)), rulesOf(await partText(doc.bytes, part)), `undone: ${part}`);
});

test('a block repeated over a larger selection takes its rules to every repeat, in one rule', async () => {
  const doc = await loadBytes(await validationBook());
  const copied = copyRange(doc.wb, 0, range('A1:B1'));
  const result = pasteAt(doc.wb, 0, range('J5:M7'), copied.text, copied);
  assert.deepEqual(result.range, range('J5:M7'));
  apply(doc, result.ops);
  // Columns J and L of the three rows; where the repeats touch they are one range.
  assert.deepEqual(validations(doc.wb)[0], 'A1:A3 J5:J7 L5:L7|whole|1|B1*2');
  await savedAsModel(doc, 'tiled');
  // A block that is all one rule fills the selection with one range.
  const one = copyRange(doc.wb, 0, range('A2'));
  apply(doc, pasteAt(doc.wb, 1, range('B2:D40'), one.text, one).ops);
  assert.deepEqual(validations(doc.wb, 1), ['B2:D40|whole|1|C2*2']);
  await savedAsModel(doc, 'tiled, one range');
});

test('copy and paste in place changes no rule; a paste that changes none writes none', async () => {
  const doc = await loadBytes(await validationBook(true));
  for (const ref of ['A1:C3', 'A2', 'A1:L5', 'C1:C2']) {
    const copied = copyRange(doc.wb, 0, range(ref));
    const result = pasteAt(doc.wb, 0, range(ref.split(':')[0]), copied.text, copied);
    assert.deepEqual(result.ops.filter((op) => op.type !== 'setCells'), [], ref);
    apply(doc, result.ops);
  }
  const saved = await save(doc);
  assert.equal(rulesOf(await partText(saved, SHEET1)), rulesOf(await partText(doc.bytes, SHEET1)));
  // Cells without a rule pasted onto cells without one: nothing but the cells.
  const copied = copyRange(doc.wb, 0, range('B1:B3'));
  assert.deepEqual(pasteAt(doc.wb, 0, range('N1'), copied.text, copied).ops.map((op) => op.type), ['setCells']);
});

test('rules on empty cells beyond the used part of the sheet are copied and pasted with the selection', async () => {
  const doc = await loadBytes(await validationBook());
  apply(doc, setValidation(doc.wb, 0, range('P20:Q30'), { type: 'textLength', operator: 'lessThan', formula1: '4' }));
  const copied = copyRange(doc.wb, 0, range('P18:Q25'));
  assert.deepEqual(copied.rules!.validations.map((item) => item.ranges.map(rangeName).join(' ')), ['A3:B8']);
  const result = pasteAt(doc.wb, 1, range('A1'), copied.text, copied);
  assert.deepEqual(result.range, range('A1:B8'));
  apply(doc, result.ops);
  assert.deepEqual(validations(doc.wb, 1), ['A3:B8|textLength|4|']);
  // A whole column with a rule for all of it.
  apply(doc, setValidation(doc.wb, 0, range('S1:S1048576'), { type: 'whole', operator: 'greaterThan', formula1: 'T1' }));
  const column = copyRange(doc.wb, 0, range('S1:S1048576'));
  apply(doc, pasteAt(doc.wb, 1, range('D1'), column.text, column).ops);
  assert.deepEqual(validations(doc.wb, 1), ['A3:B8|textLength|4|', 'D1:D1048576|whole|E1|']);
  await savedAsModel(doc, 'beyond the used part');
});

test('a cut takes the rules of its cells away, and a paste of it puts them where the cells go', async () => {
  const doc = await loadBytes(await validationBook());
  const copied = copyRange(doc.wb, 0, range('A2:C3'));
  apply(doc, cutRange(doc.wb, 0, range('A2:C3')));
  assert.deepEqual(validations(doc.wb), ['A1|whole|1|B1*2', 'C1|list|$H$1:$H$2|', 'E5|custom|AND(E5>$H$1,Lists!$A1<>D5)|']);
  assert.equal(doc.wb.sheets[0].rows.get(1)!.cells[0], undefined);
  apply(doc, pasteAt(doc.wb, 0, range('J10'), copied.text, copied).ops);
  // What is left of each rule is the same rule: it gets the cells.
  assert.deepEqual(validations(doc.wb), ['A1 J10:J11|whole|1|B1*2', 'C1 L10|list|$H$1:$H$2|', 'E5|custom|AND(E5>$H$1,Lists!$A1<>D5)|']);
  await savedAsModel(doc, 'cut and pasted');
  // Each of the two is one step.
  doc.history.undo();
  doc.history.undo();
  assert.equal(doc.history.canUndo, false);
  assert.equal(rulesOf(await partText(await save(doc), SHEET1)), rulesOf(await partText(doc.bytes, SHEET1)), 'undone');
});

test('a validation of the extension list is neither copied nor taken away: its cells keep it, and the paste says so', async () => {
  const doc = await loadBytes(await validationBook(true));
  assert.deepEqual(validations(doc.wb).slice(-1), ['L1:L3|list|Lists!$B$1:$B$3||x14']);
  // Copied: it is left where it is, and counted.
  const copied = copyRange(doc.wb, 0, range('K1:L2'));
  assert.deepEqual([copied.rules!.validations, copied.rules!.conditional, copied.rules!.left.map((item) => `${item.kind} ${item.ranges.map(rangeName).join(' ')}`)], [[], [], ['validation B1:B2']]);
  const away = pasteAt(doc.wb, 0, range('N1'), copied.text, copied);
  assert.equal(away.skippedRules, 1);
  assert.deepEqual(away.ops.filter((op) => op.type !== 'setCells'), []);
  // Pasted over: the cells under it keep it, the others get the rule of the copy.
  const whole = copyRange(doc.wb, 0, range('A1'));
  const over = pasteAt(doc.wb, 0, range('K2:M3'), whole.text, whole);
  assert.equal(over.skippedRules, 1);
  apply(doc, over.ops);
  assert.deepEqual(validations(doc.wb), ['A1:A3 K2:K3 M2:M3|whole|1|B1*2', 'C1:C2|list|$H$1:$H$2|', 'E5|custom|AND(E5>$H$1,Lists!$A1<>D5)|', 'L1:L3|list|Lists!$B$1:$B$3||x14']);
  const saved = await savedAsModel(doc, 'beside the extension list');
  assert.ok((await partText(saved, SHEET1)).indexOf(X14_LIST) > 0, 'the extension list is as it was');
  // A cut leaves it as well.
  apply(doc, cutRange(doc.wb, 0, range('K1:M3')));
  assert.deepEqual(validations(doc.wb), ['A1:A3|whole|1|B1*2', 'C1:C2|list|$H$1:$H$2|', 'E5|custom|AND(E5>$H$1,Lists!$A1<>D5)|', 'L1:L3|list|Lists!$B$1:$B$3||x14']);
  while (doc.history.canUndo) doc.history.undo();
  assert.equal(rulesOf(await partText(await save(doc), SHEET1)), rulesOf(await partText(doc.bytes, SHEET1)), 'undone');
});

test('rows that a filter hides are not in the copy, and neither are their rules; a cut leaves them theirs', async () => {
  const doc = await loadBytes(await validationBook());
  apply(doc, setValidation(doc.wb, 0, range('B2'), { type: 'list', formula1: '"x,y"' }));
  apply(doc, [{ type: 'setAutoFilter', sheet: 0, autoFilter: { range: range('A1:B3'), columns: [{ colId: 0, kind: 'values', values: ['3', '2'] }] }, hide: [1] }]);
  const copied = copyRange(doc.wb, 0, range('A1:B3'));
  // Rows 1 and 3 stand next to each other in the copy: the rule of column A is one range of two rows, B2 is not there.
  assert.deepEqual(copied.rules!.validations.map((item) => [item.ranges.map(rangeName).join(' '), item.rule.type]), [['A1:A2', 'whole']]);
  apply(doc, pasteAt(doc.wb, 1, range('A1'), copied.text, copied).ops);
  assert.deepEqual(validations(doc.wb, 1), ['A1:A2|whole|1|B1*2']);
  apply(doc, cutRange(doc.wb, 0, range('A1:B3')));
  assert.deepEqual(validations(doc.wb), ['A2|whole|1|B2*2', 'C1:C2|list|$H$1:$H$2|', 'E5|custom|AND(E5>$H$1,Lists!$A1<>D5)|', 'B2|list|"x,y"|']);
});

// ---------------------------------------------------------------------------
// Conditional formats

test('a paste gives the cells the conditional formats of the cells they are copies of, and takes their own away', async () => {
  const doc = await loadBytes(await conditionalBook(BLOCK_A + BLOCK_B));
  assert.deepEqual(formats(doc.wb), ['A1:A10 F1:F3|expression|AND(A1>B1,Lists!$A$1<>A1)', 'A1:A10 F1:F3|cellIs|$H$1', 'B1:B5|colorScale|']);
  // B2:B3 has the scale; pasted over A9:A10 it takes those out of the two rules of A, and the scale gets them.
  const copied = copyRange(doc.wb, 0, range('B2:B3'));
  assert.deepEqual(copied.rules!.conditional.map((item) => [item.ranges.map(rangeName).join(' '), item.rules.map((rule) => rule.type)]), [['A1:A2', ['colorScale']]]);
  const result = pasteAt(doc.wb, 0, range('A9'), copied.text, copied);
  assert.equal(result.ops.filter((op) => op.type === 'setConditionalFormats').length, 1);
  apply(doc, result.ops);
  assert.deepEqual(formats(doc.wb), ['A1:A8 F1:F3|expression|AND(A1>B1,Lists!$A$1<>A1)', 'A1:A8 F1:F3|cellIs|$H$1', 'B1:B5 A9:A10|colorScale|']);
  await savedAsModel(doc, 'pasted');
  doc.history.undo();
  assert.equal(doc.history.canUndo, false);
  assert.equal(rulesOf(await partText(await save(doc), SHEET1)), rulesOf(await partText(doc.bytes, SHEET1)), 'undone');
});

test('rules that are new to the sheet are added with moved formulas, first in the order, and keep the format they give', async () => {
  const doc = await loadBytes(await conditionalBook(BLOCK_A + BLOCK_B));
  const copied = copyRange(doc.wb, 0, range('A3:A4'));
  const result = pasteAt(doc.wb, 1, range('C5:C8'), copied.text, copied);
  apply(doc, result.ops);
  // Two rules of one format: two rules of the other sheet, in the order they had, for both repeats of the block.
  assert.deepEqual(conditionalNow(doc.wb, 1), ['C5:C8|expression|1|AND(C5>D5,Lists!$A$1<>C5)', 'C5:C8|cellIs|2|$H$1|greaterThan']);
  const made = doc.wb.sheets[1].conditionalFormats!.map((format) => format.rules[0]);
  assert.deepEqual(made.map((rule) => [rule.dxfId, rule.src, rule.stopIfTrue]), [[1, undefined, true], [0, undefined, false]]);
  const saved = await savedAsModel(doc, 'new rules');
  const xml = await partText(saved, SHEET2);
  assert.ok(xml.indexOf('<conditionalFormatting sqref="C5:C8"><cfRule type="expression" dxfId="1" priority="1" stopIfTrue="1"><formula>AND(C5&gt;D5,Lists!$A$1&lt;&gt;C5)</formula></cfRule></conditionalFormatting>') > 0, xml);
  sameBytes(new TextEncoder().encode(await partText(saved, SHEET1)), new TextEncoder().encode(await partText(doc.bytes, SHEET1)), 'the sheet that was copied from');
  // Pasted a second time beside it, the rules that are there get the cells: no rule is made twice.
  apply(doc, pasteAt(doc.wb, 1, range('D5'), copied.text, copied).ops);
  assert.deepEqual(conditionalNow(doc.wb, 1), ['C5:C8 D5:D6|expression|1|AND(C5>D5,Lists!$A$1<>C5)', 'C5:C8 D5:D6|cellIs|2|$H$1|greaterThan']);
  await savedAsModel(doc, 'pasted again');
  doc.history.undo();
  doc.history.undo();
  for (const part of [SHEET1, SHEET2]) assert.equal(rulesOf(await partText(await save(doc), part)), rulesOf(await partText(doc.bytes, part)), `undone: ${part}`);
});

test('copy and paste in place leaves every conditional format as the file has it, whatever kind it is', async () => {
  const doc = await loadBytes(await conditionalBook(BLOCK_A + BLOCK_B + BLOCK_C + BLOCK_D, EXT_BAR + EXT_ONLY));
  for (const ref of ['A1:F10', 'A2:C3', 'C1:E5', 'D2']) {
    const copied = copyRange(doc.wb, 0, range(ref));
    const result = pasteAt(doc.wb, 0, range(ref.split(':')[0]), copied.text, copied);
    assert.deepEqual(result.ops.filter((op) => op.type !== 'setCells'), [], ref);
    apply(doc, result.ops);
  }
  assert.equal(rulesOf(await partText(await save(doc), SHEET1)), rulesOf(await partText(doc.bytes, SHEET1)));
});

test('what the extension list says: a rule it alone has stays where it is, one it says more about is pasted as the main part states it', async () => {
  const doc = await loadBytes(await conditionalBook(BLOCK_A + BLOCK_B + BLOCK_C + BLOCK_D, EXT_BAR + EXT_ONLY));
  const before = formats(doc.wb);
  assert.ok(before.indexOf('E1:E5|iconSet|') >= 0 && before.indexOf('C1:C5|dataBar|') >= 0, before.join('; '));
  // E has a rule of the extension list alone: it is not in the copy.
  const stars = copyRange(doc.wb, 0, range('E1:E2'));
  assert.deepEqual([stars.rules!.validations, stars.rules!.conditional, stars.rules!.left.map((item) => `${item.kind} ${item.ranges.map(rangeName).join(' ')}`)], [[], [], ['conditional A1:A2']]);
  const none = pasteAt(doc.wb, 0, range('J1'), stars.text, stars);
  assert.equal(none.skippedRules, 1);
  assert.deepEqual(none.ops.filter((op) => op.type !== 'setCells'), []);
  // Pasted over, it stays: the rules of the copy come beside it.
  const scale = copyRange(doc.wb, 0, range('B1'));
  apply(doc, pasteAt(doc.wb, 0, range('E2'), scale.text, scale).ops);
  assert.ok(formats(doc.wb).indexOf('E1:E5|iconSet|') >= 0, formats(doc.wb).join('; '));
  assert.ok(formats(doc.wb).indexOf('B1:B5 E2|colorScale|') >= 0, formats(doc.wb).join('; '));
  // The data bar of C, with a border the extension list states: on another sheet it is the bar of the main part, and the paste says so.
  const bar = copyRange(doc.wb, 0, range('C1:C2'));
  const pasted = pasteAt(doc.wb, 1, range('A1'), bar.text, bar);
  assert.equal(pasted.skippedRules, 1);
  apply(doc, pasted.ops);
  assert.deepEqual(formats(doc.wb, 1), ['A1:A2|dataBar|']);
  const rule = doc.wb.sheets[1].conditionalFormats![0].rules[0];
  assert.equal(rule.extId, undefined);
  assert.equal(rule.src, undefined);
  // In its own sheet the bar itself gets the cells, with everything it has.
  const beside = pasteAt(doc.wb, 0, range('C8'), bar.text, bar);
  assert.equal(beside.skippedRules, undefined);
  apply(doc, beside.ops);
  assert.ok(formats(doc.wb).indexOf('C1:C5 C8:C9|dataBar|') >= 0, formats(doc.wb).join('; '));
  const saved = await savedAsModel(doc, 'with the extension list');
  const xml = await partText(saved, SHEET1);
  assert.ok(xml.indexOf('<xm:sqref>C1:C5 C8:C9</xm:sqref>') > 0, 'the counterpart in the extension list follows the bar');
  assert.ok(xml.indexOf('<xm:sqref>E1:E5</xm:sqref>') > 0, 'the rule of the extension list is where it was');
  // A cut leaves the rule of the extension list, and takes the others.
  apply(doc, cutRange(doc.wb, 0, range('A1:F10')));
  assert.deepEqual(formats(doc.wb), ['E1:E5|iconSet|']);
  await savedAsModel(doc, 'cut');
  while (doc.history.canUndo) doc.history.undo();
  for (const part of [SHEET1, SHEET2]) assert.equal(rulesOf(await partText(await save(doc), part)), rulesOf(await partText(doc.bytes, part)), `undone: ${part}`);
});

test('a cut and a paste move the conditional formats of the cells; rules made here travel like those of a file', async () => {
  const doc = await loadBytes(await conditionalBook(''));
  apply(doc, addConditionalRule(doc.wb, 0, range('A1:A3'), { type: 'cellIs', operator: 'greaterThan', formulas: ['B1'], format: { fill: '#FFC7CE' } }));
  apply(doc, addConditionalRule(doc.wb, 0, range('A1:B3'), { type: 'top10', rank: 2 }));
  const copied = copyRange(doc.wb, 0, range('A2:B3'));
  apply(doc, cutRange(doc.wb, 0, range('A2:B3')));
  assert.deepEqual(conditionalNow(doc.wb), ['A1:B1|top10|1|', 'A1|cellIs|2|B1|greaterThan']);
  apply(doc, pasteAt(doc.wb, 0, range('D5'), copied.text, copied).ops);
  assert.deepEqual(conditionalNow(doc.wb), ['A1:B1 D5:E6|top10|1|', 'A1 D5:D6|cellIs|2|B1|greaterThan']);
  const saved = await savedAsModel(doc, 'cut and pasted');
  assert.ok((await partText(saved, SHEET1)).indexOf('sqref="A1 D5:D6"') > 0, 'the rule with its new cells');
  doc.history.undo();
  doc.history.undo();
  assert.deepEqual(conditionalNow(doc.wb), ['A1:B3|top10|1|', 'A1:A3|cellIs|2|B1|greaterThan']);
});

test('a chart sheet takes no rules, and a copy held over rows that were inserted is pasted as it was copied', async () => {
  const doc = await loadBytes(await validationBook());
  const copied = copyRange(doc.wb, 0, range('A1:A3'));
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 2 }]);
  assert.deepEqual(validations(doc.wb)[0], 'A3:A5|whole|1|B3*2');
  apply(doc, pasteAt(doc.wb, 0, range('N1'), copied.text, copied).ops);
  // The rule that moved down is the same rule (it reads the cell beside, doubled): it gets the cells.
  assert.deepEqual(validations(doc.wb)[0], 'A3:A5 N1:N3|whole|1|B3*2');
  await savedAsModel(doc, 'after rows were inserted');
});
