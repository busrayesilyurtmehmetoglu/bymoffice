// Data validations as the editor changes them: the rule of a range set, replaced and
// cleared, ranges cut out of the rules that were there, what is refused, undo and redo, and
// what rows, columns and sheets that move do to a rule that was changed.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine } from '../src/calc/engine';
import { open, parseCellRef, parseRange, rangeName } from '../src/index';
import type { CellRange, DataValidation, Workbook } from '../src/index';
import { DataValidations } from '../src/format/validation';
import { copyRange, pasteAt } from '../src/ops/commands';
import { setComment } from '../src/ops/notes';
import { History, type Op } from '../src/ops/ops';
import { sortRange } from '../src/ops/sort';
import { dataValidationNow, ruleSpace, setValidation, subtractRange, validationsIn } from '../src/ops/rules';
import { buildXlsx, worksheetXml } from './fixtures';

const dv = (attrs: string, inner = '') => `<dataValidation ${attrs}>${inner}</dataValidation>`;
const f1 = (text: string) => `<formula1>${text}</formula1>`;
const f2 = (text: string) => `<formula2>${text}</formula2>`;
const range = (ref: string): CellRange => parseRange(ref)!;
const refs = (rule: DataValidation | undefined) => (rule ? rule.ranges.map(rangeName).join(' ') : undefined);

const SHEET =
  '<sheetData><row r="1"><c r="H1"><v>1</v></c></row><row r="2"><c r="H2"><v>2</v></c></row></sheetData>' +
  '<dataValidations count="3">' +
  dv('type="whole" showErrorMessage="1" sqref="A1:A10 C1:C3"', f1('1') + f2('9')) +
  dv('sqref="ZZZZ1"') +
  dv('type="list" showErrorMessage="1" sqref="B1:B5"', f1('$H$1:$H$2')) +
  '</dataValidations>';

async function load(xml = SHEET, more: { name: string; xml: string }[] = []): Promise<{ wb: Workbook; history: History }> {
  const bytes = await buildXlsx({ sheets: [{ name: 'Data', xml: worksheetXml(xml) }, ...more.map((sheet) => ({ name: sheet.name, xml: worksheetXml(sheet.xml) }))] });
  const wb = await (await open(bytes)).workbook();
  return { wb, history: new History(wb) };
}

test('the reader notes which element of the part each rule is', async () => {
  const { wb } = await load();
  const rules = wb.sheets[0].dataValidations!;
  // The second element has no range inside the sheet and is no rule; the third keeps its own number.
  assert.deepEqual(rules.map((rule) => rule.src), [0, 2]);
  assert.equal(rules[0].live, undefined);
});

// ---------------------------------------------------------------------------
// Ranges

test('a range cut out of another leaves up to four rectangles, none of them overlapping', () => {
  const names = (ranges: CellRange[]) => ranges.map(rangeName);
  assert.deepEqual(names(subtractRange(range('A1:E5'), range('B2:C3'))), ['A1:E1', 'A4:E5', 'A2:A3', 'D2:E3']);
  assert.deepEqual(names(subtractRange(range('A1:E5'), range('A1:E5'))), []);
  assert.deepEqual(names(subtractRange(range('A1:E5'), range('G1:H2'))), ['A1:E5']);
  assert.deepEqual(names(subtractRange(range('A1:A10'), range('A3:C4'))), ['A1:A2', 'A5:A10']);
  assert.deepEqual(names(subtractRange(range('A1:E5'), range('C1:Z99'))), ['A1:B5']);
  // Every cell of the range is in exactly one piece or in the cut.
  const pieces = subtractRange(range('B2:H9'), range('D4:E20'));
  for (let row = 0; row < 12; row++) {
    for (let col = 0; col < 10; col++) {
      const inside = (r: CellRange) => row >= r.top && row <= r.bottom && col >= r.left && col <= r.right;
      const expected = inside(range('B2:H9')) && !inside(range('D4:E20')) ? 1 : 0;
      assert.equal(pieces.filter(inside).length, expected, `${row},${col}`);
    }
  }
});

// ---------------------------------------------------------------------------
// The command

const WHOLE = { type: 'whole', operator: 'between', formula1: '1', formula2: '5' } as const;
const apply = (doc: { wb: Workbook; history: History }, ops: Op[]) => doc.history.apply(ops, undefined, undefined);
/** The rules of the first sheet as they apply now: "ranges|type|formula1|formula2". */
const now = (wb: Workbook, sheet = 0) => {
  const space = ruleSpace(wb, wb.sheets[sheet]);
  return (wb.sheets[sheet].dataValidations || [])
    .map((rule) => dataValidationNow(rule, space))
    .filter((rule): rule is DataValidation => !!rule)
    .map((rule) => `${refs(rule)}|${rule.type}|${rule.formula1 ?? ''}|${rule.formula2 ?? ''}`);
};

test('a rule set on a range applies at once, is one undo step, and marks only the validations as changed', async () => {
  const doc = await load('<sheetData/>');
  const sheet = doc.wb.sheets[0];
  assert.deepEqual(apply(doc, setValidation(doc.wb, 0, range('B2:C4'), WHOLE)).length, 1);
  assert.deepEqual(now(doc.wb), ['B2:C4|whole|1|5']);
  const rule = sheet.dataValidations![0];
  assert.equal(rule.live, true);
  assert.equal(rule.src, undefined);
  // What a rule has when nothing else is said, as the dialog of Excel has it.
  assert.deepEqual([rule.allowBlank, rule.showInputMessage, rule.showErrorMessage, rule.errorStyle, rule.hideDropDown], [true, true, true, 'stop', false]);
  assert.equal(sheet.validationsModified, true);
  assert.equal(sheet.modified, undefined, 'no cell changed');
  const rules = new DataValidations(doc.wb, sheet);
  assert.equal(rules.check(2, 1, { kind: 'number', value: 9, style: 0 }, '9').ok, false);
  assert.equal(rules.check(2, 1, { kind: 'number', value: 3, style: 0 }, '3').ok, true);
  assert.equal(rules.check(9, 9, { kind: 'number', value: 9, style: 0 }, '9').ok, true);

  doc.history.undo();
  assert.equal(sheet.dataValidations, undefined);
  assert.equal(rules.check(2, 1, { kind: 'number', value: 9, style: 0 }, '9').ok, true, 'the rule went with the step');
  doc.history.redo();
  assert.deepEqual(now(doc.wb), ['B2:C4|whole|1|5']);
});

test('a rule set over part of another takes those cells out of it; the other ranges of that rule stay', async () => {
  const doc = await load();
  const before = doc.wb.sheets[0].dataValidations!;
  apply(doc, setValidation(doc.wb, 0, range('A3:B4'), { type: 'decimal', operator: 'greaterThan', formula1: '0' }));
  assert.deepEqual(now(doc.wb), ['A1:A2 A5:A10 C1:C3|whole|1|9', 'B1:B2 B5|list|$H$1:$H$2|', 'A3:B4|decimal|0|']);
  const rules = doc.wb.sheets[0].dataValidations!;
  // The rules that were cut are the rules of the file still, as far as the writer is concerned.
  assert.deepEqual(rules.map((rule) => [rule.src, rule.live]), [[0, true], [2, true], [undefined, true]]);
  doc.history.undo();
  assert.equal(doc.wb.sheets[0].dataValidations, before, 'the very list of the file is back');
});

test('a rule that another covers entirely is gone, and one that is not touched stays the object it was', async () => {
  const doc = await load();
  const list = doc.wb.sheets[0].dataValidations!;
  apply(doc, setValidation(doc.wb, 0, range('B1:B5'), WHOLE));
  assert.deepEqual(now(doc.wb), ['A1:A10 C1:C3|whole|1|9', 'B1:B5|whole|1|5']);
  assert.equal(doc.wb.sheets[0].dataValidations![0], list[0]);
});

test('clearing: no rule (or "any value" with nothing to say) takes the cells out of every rule', async () => {
  const doc = await load();
  apply(doc, setValidation(doc.wb, 0, range('A1:C3'), undefined));
  assert.deepEqual(now(doc.wb), ['A4:A10|whole|1|9', 'B4:B5|list|$H$1:$H$2|']);
  apply(doc, setValidation(doc.wb, 0, range('A1:Z100'), { type: 'none' }));
  assert.equal(doc.wb.sheets[0].dataValidations, undefined);
  // Nothing there, nothing to do: no step.
  assert.deepEqual(setValidation(doc.wb, 0, range('A1:Z100'), undefined), []);
  // "Any value" with a message to show is a rule: Excel keeps it for the message.
  apply(doc, setValidation(doc.wb, 0, range('D1'), { type: 'none', promptTitle: 'Hint', prompt: 'Type anything' }));
  assert.deepEqual(now(doc.wb), ['D1|none||']);
  assert.deepEqual(new DataValidations(doc.wb, doc.wb.sheets[0]).prompt(0, 3), { title: 'Hint', text: 'Type anything' });
});

test('the formulas of a rule that loses its first cell are written for its new first cell', async () => {
  const doc = await load('<sheetData/><dataValidations count="1">' + dv('type="custom" showErrorMessage="1" sqref="A1:A5"', f1('AND(A1&gt;B1,$C1&lt;A$2)')) + '</dataValidations>');
  apply(doc, setValidation(doc.wb, 0, range('A1:A2'), undefined));
  assert.deepEqual(now(doc.wb), ['A3:A5|custom|AND(A3>B3,$C3<A$2)|']);
});

test('the same rule set again on the same cells is no step', async () => {
  const doc = await load('<sheetData/>');
  apply(doc, setValidation(doc.wb, 0, range('B2:C4'), WHOLE));
  assert.deepEqual(setValidation(doc.wb, 0, range('B2:C4'), WHOLE), []);
  assert.equal(setValidation(doc.wb, 0, range('B2:C4'), { ...WHOLE, formula2: '6' }).length, 1);
});

test('what a rule may say: types, operators, formulas, and the limits Excel has', async () => {
  const doc = await load('<sheetData/>');
  const at = range('A1');
  const set = (rule: Parameters<typeof setValidation>[3]) => setValidation(doc.wb, 0, at, rule);
  const made = (rule: Parameters<typeof setValidation>[3]) => (set(rule)[0] as Extract<Op, { type: 'setDataValidations' }>).validations![0];
  assert.throws(() => set({ type: 'nonsense' as 'whole', formula1: '1' }), /kind of rule/);
  assert.throws(() => set({ type: 'whole', operator: 'sideways' as 'equal', formula1: '1' }), /operator/);
  assert.throws(() => set({ type: 'whole', operator: 'equal' }), /needs a value/);
  assert.throws(() => set({ type: 'whole', operator: 'between', formula1: '1' }), /needs two values/);
  assert.throws(() => set({ type: 'custom', formula1: 'SUM(A1' }), /cannot be read/);
  assert.throws(() => set({ type: 'list', formula1: `"${'x,'.repeat(140)}"` }), /255/);
  assert.throws(() => set({ type: 'whole', operator: 'equal', formula1: '1', promptTitle: 'x'.repeat(33) }), /32/);
  assert.throws(() => set({ type: 'whole', operator: 'equal', formula1: '1', errorTitle: 'x'.repeat(33) }), /32/);
  assert.throws(() => set({ type: 'whole', operator: 'equal', formula1: '1', prompt: 'x'.repeat(256) }), /255/);
  assert.throws(() => set({ type: 'whole', operator: 'equal', formula1: '1', error: 'x'.repeat(256) }), /255/);
  assert.throws(() => setValidation(doc.wb, 0, { top: 0, left: 0, bottom: 2000000, right: 0 }, WHOLE), /outside the sheet/);
  assert.throws(() => setValidation(doc.wb, 7, at, WHOLE), /No sheet/);
  // A leading "=" is how a formula is typed, not part of it; the second value goes where it says nothing.
  assert.deepEqual([made({ type: 'whole', operator: 'equal', formula1: '=$B$1', formula2: '9' }).formula1, made({ type: 'whole', operator: 'equal', formula1: '=$B$1', formula2: '9' }).formula2], ['$B$1', undefined]);
  // Each kind with each operator it has.
  for (const type of ['whole', 'decimal', 'date', 'time', 'textLength'] as const) {
    for (const operator of ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'] as const) {
      const two = operator === 'between' || operator === 'notBetween';
      const rule = made({ type, operator, formula1: '1', ...(two ? { formula2: '2' } : {}) });
      assert.deepEqual([rule.type, rule.operator, rule.formula1, rule.formula2], [type, operator, '1', two ? '2' : undefined]);
    }
  }
  assert.equal(made({ type: 'list', formula1: '"a,b,c"', hideDropDown: true }).hideDropDown, true);
  assert.equal(made({ type: 'list', formula1: '$H$1:$H$5', operator: 'equal' }).operator, 'between', 'a list has no operator');
  assert.equal(made({ type: 'custom', formula1: 'ISNUMBER(A1)', errorStyle: 'warning' }).errorStyle, 'warning');
});

test('cells under a rule of the extension list are not given another rule', async () => {
  const xml =
    '<sheetData/><dataValidations count="1">' + dv('type="whole" sqref="A1:A3"', f1('1') + f2('2')) + '</dataValidations>' +
    '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
    '<x14:dataValidations xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" count="1">' +
    '<x14:dataValidation type="list" allowBlank="1" showErrorMessage="1"><x14:formula1><xm:f>Lists!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>C1:C5</xm:sqref></x14:dataValidation>' +
    '</x14:dataValidations></ext></extLst>';
  const doc = await load(xml, [{ name: 'Lists', xml: '<sheetData/>' }]);
  assert.throws(() => setValidation(doc.wb, 0, range('B1:C2'), WHOLE), /extension list/);
  assert.throws(() => setValidation(doc.wb, 0, range('C5'), undefined), /extension list/);
  // Beside it everything works as ever.
  apply(doc, setValidation(doc.wb, 0, range('A2:B2'), WHOLE));
  // The new rule stands before the rules of the extension list, where a file that is read has it.
  assert.deepEqual(now(doc.wb), ['A1 A3|whole|1|2', 'A2:B2|whole|1|5', 'C1:C5|list|Lists!$A$1:$A$3|']);
});

test('the rules of a selection, for whoever asks what is there', async () => {
  const doc = await load();
  assert.deepEqual(validationsIn(doc.wb, 0, range('E1:F9')), []);
  assert.deepEqual(validationsIn(doc.wb, 0, range('A1')).map(refs), ['A1:A10 C1:C3']);
  assert.deepEqual(validationsIn(doc.wb, 0, range('A1:B1')).map((rule) => rule.type), ['whole', 'list']);
});

// ---------------------------------------------------------------------------
// A rule that was changed, and what moves afterwards

const insert = (sheet: number, axis: 'row' | 'col', at: number, count: number): Op[] => [{ type: 'insert', sheet, axis, at, count }];
const remove = (sheet: number, axis: 'row' | 'col', at: number, count: number): Op[] => [{ type: 'delete', sheet, axis, at, count }];

test('rows and columns that move take a changed rule with them, as they take a rule of the file', async () => {
  const doc = await load();
  // One rule of the file is cut (it is live from here on), one is added, one stays as the file has it.
  apply(doc, setValidation(doc.wb, 0, range('A9:A10'), { type: 'custom', formula1: 'AND(A9>$H$1,B9<>"")' }));
  assert.deepEqual(now(doc.wb), ['A1:A8 C1:C3|whole|1|9', 'B1:B5|list|$H$1:$H$2|', 'A9:A10|custom|AND(A9>$H$1,B9<>"")|']);
  apply(doc, insert(0, 'row', 0, 2));
  assert.deepEqual(now(doc.wb), ['A3:A10 C3:C5|whole|1|9', 'B3:B7|list|$H$3:$H$4|', 'A11:A12|custom|AND(A11>$H$3,B11<>"")|']);
  // A relative reference means "the cell this far from the one the rule is asked about": the rule did not
  // move, so it stays "one to the right", as it does in a rule of the file (see adjustAnchoredFormula()).
  apply(doc, insert(0, 'col', 1, 1));
  assert.deepEqual(now(doc.wb), ['A3:A10 D3:D5|whole|1|9', 'C3:C7|list|$I$3:$I$4|', 'A11:A12|custom|AND(A11>$I$3,B11<>"")|']);
  apply(doc, remove(0, 'row', 10, 1));
  assert.deepEqual(now(doc.wb), ['A3:A10 D3:D5|whole|1|9', 'C3:C7|list|$I$3:$I$4|', 'A11|custom|AND(A11>$I$3,B11<>"")|']);
  // All its cells deleted: the rule is gone, and undo brings it back.
  apply(doc, remove(0, 'row', 10, 1));
  assert.deepEqual(now(doc.wb), ['A3:A10 D3:D5|whole|1|9', 'C3:C7|list|$I$3:$I$4|']);
  doc.history.undo();
  assert.deepEqual(now(doc.wb), ['A3:A10 D3:D5|whole|1|9', 'C3:C7|list|$I$3:$I$4|', 'A11|custom|AND(A11>$I$3,B11<>"")|']);
  for (let i = 0; i < 3; i++) doc.history.undo();
  assert.deepEqual(now(doc.wb), ['A1:A8 C1:C3|whole|1|9', 'B1:B5|list|$H$1:$H$2|', 'A9:A10|custom|AND(A9>$H$1,B9<>"")|']);
  for (let i = 0; i < 4; i++) doc.history.redo();
  assert.deepEqual(now(doc.wb), ['A3:A10 D3:D5|whole|1|9', 'C3:C7|list|$I$3:$I$4|']);
});

test('a rule set after rows moved is where it was put, not where the earlier moves would put it', async () => {
  const doc = await load();
  apply(doc, insert(0, 'row', 0, 5));
  apply(doc, setValidation(doc.wb, 0, range('E1:E2'), WHOLE));
  assert.deepEqual(now(doc.wb), ['A6:A15 C6:C8|whole|1|9', 'B6:B10|list|$H$6:$H$7|', 'E1:E2|whole|1|5']);
  const rules = new DataValidations(doc.wb, doc.wb.sheets[0]);
  assert.equal(refs(rules.ruleAt(0, 4)), 'E1:E2');
  assert.equal(rules.ruleAt(5, 4), undefined);
  apply(doc, insert(0, 'row', 0, 1));
  assert.equal(refs(rules.ruleAt(1, 4)), 'E2:E3');
});

test('a changed rule that reads another sheet follows the rows of that sheet and its name', async () => {
  const doc = await load('<sheetData/>', [{ name: 'Lists', xml: '<sheetData/>' }, { name: 'Other', xml: '<sheetData/>' }]);
  apply(doc, setValidation(doc.wb, 0, range('A1:A5'), { type: 'list', formula1: 'Lists!$A$2:$A$4' }));
  apply(doc, setValidation(doc.wb, 0, range('B1'), { type: 'custom', formula1: 'Other!A1>Lists!$B$2' }));
  apply(doc, insert(1, 'row', 0, 3));
  assert.deepEqual(now(doc.wb), ['A1:A5|list|Lists!$A$5:$A$7|', 'B1|custom|Other!A1>Lists!$B$5|']);
  apply(doc, [{ type: 'renameSheet', sheet: 1, name: 'My Lists' }]);
  assert.deepEqual(now(doc.wb), ["A1:A5|list|'My Lists'!$A$5:$A$7|", "B1|custom|Other!A1>'My Lists'!$B$5|"]);
  // A sheet that takes the old name is not the sheet the rule reads.
  apply(doc, [{ type: 'renameSheet', sheet: 2, name: 'Lists' }]);
  assert.deepEqual(now(doc.wb), ["A1:A5|list|'My Lists'!$A$5:$A$7|", "B1|custom|Lists!A1>'My Lists'!$B$5|"]);
  apply(doc, [{ type: 'deleteSheet', sheet: 1 }]);
  assert.deepEqual(now(doc.wb), ['A1:A5|list|#REF!|', 'B1|custom|Lists!A1>#REF!|']);
  doc.history.undo();
  doc.history.undo();
  doc.history.undo();
  assert.deepEqual(now(doc.wb), ['A1:A5|list|Lists!$A$5:$A$7|', 'B1|custom|Other!A1>Lists!$B$5|']);
  doc.history.undo();
  assert.deepEqual(now(doc.wb), ['A1:A5|list|Lists!$A$2:$A$4|', 'B1|custom|Other!A1>Lists!$B$2|']);
});

test('a rule of the file that is changed after its sheet was renamed speaks of the sheets by the names they have', async () => {
  const xml = '<sheetData/><dataValidations count="1">' + dv('type="list" sqref="A1:A5"', f1('Lists!$A$1:$A$3')) + '</dataValidations>';
  const doc = await load(xml, [{ name: 'Lists', xml: '<sheetData/>' }]);
  apply(doc, [{ type: 'renameSheet', sheet: 1, name: 'Sources' }]);
  apply(doc, insert(1, 'row', 0, 1));
  assert.deepEqual(now(doc.wb), ['A1:A5|list|Sources!$A$2:$A$4|']);
  // Cut: the rule is live now, and holds what it said through the names and rows of the file.
  apply(doc, setValidation(doc.wb, 0, range('A1'), undefined));
  assert.equal(doc.wb.sheets[0].dataValidations![0].formula1, 'Sources!$A$2:$A$4');
  assert.deepEqual(now(doc.wb), ['A2:A5|list|Sources!$A$2:$A$4|']);
  apply(doc, [{ type: 'renameSheet', sheet: 1, name: 'Lists' }]);
  apply(doc, insert(1, 'row', 0, 1));
  assert.deepEqual(now(doc.wb), ['A2:A5|list|Lists!$A$3:$A$5|']);
});

// ---------------------------------------------------------------------------
// With the other things that happen to cells

test('sorting, pasting, merging and notes leave the rules where they are', async () => {
  const xml =
    '<sheetData><row r="1"><c r="A1"><v>3</v></c><c r="B1"><v>30</v></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>10</v></c></row><row r="3"><c r="A3"><v>2</v></c><c r="B3"><v>20</v></c></row></sheetData>';
  const doc = await load(xml);
  apply(doc, setValidation(doc.wb, 0, range('A1'), { type: 'whole', operator: 'equal', formula1: 'B1/10' }));
  apply(doc, setValidation(doc.wb, 0, range('B2:B3'), { type: 'list', formula1: '"10,20"' }));
  const before = now(doc.wb);
  // A sort moves cells, not rules (Excel does not move them either): A1 has its rule, whatever value stands there.
  apply(doc, sortRange(doc.wb, 0, range('A1:B3'), [{ col: 0, descending: false }], { header: false }));
  assert.equal(doc.wb.sheets[0].rows.get(0)!.cells[0].value, 1);
  assert.deepEqual(now(doc.wb), before);
  // A copy carries values and formats; the rules of the cells it lands on stay, and it brings none.
  const copied = copyRange(doc.wb, 0, range('A1:B1'));
  apply(doc, pasteAt(doc.wb, 0, range('D1'), copied.text, copied).ops);
  assert.deepEqual(now(doc.wb), before);
  apply(doc, pasteAt(doc.wb, 0, range('B2'), copied.text, copied).ops);
  assert.deepEqual(now(doc.wb), before);
  // A merged cell is the range it is; a rule set on part of it is set on that part.
  apply(doc, [{ type: 'setMerges', sheet: 0, merges: [range('F1:G2')] }]);
  apply(doc, setValidation(doc.wb, 0, range('F1'), WHOLE));
  assert.deepEqual(now(doc.wb), before.concat(['F1|whole|1|5']));
  apply(doc, setComment(doc.wb, 0, 0, 0, 'A note'));
  assert.deepEqual(now(doc.wb), before.concat(['F1|whole|1|5']));
  assert.equal(doc.wb.sheets[0].comments!.length, 1);
});

test('a rule of another sheet that reads a deleted sheet says #REF!, and undo brings the sheet and the rule back', async () => {
  const doc = await load('<sheetData/>', [{ name: 'Lists', xml: '<sheetData/><dataValidations count="1">' + dv('type="whole" sqref="A1"', f1('1') + f2('2')) + '</dataValidations>' }]);
  apply(doc, setValidation(doc.wb, 1, range('B1'), { type: 'custom', formula1: 'Data!A1>0' }));
  apply(doc, setValidation(doc.wb, 0, range('A1'), { type: 'list', formula1: 'Lists!$A$1:$A$3' }));
  apply(doc, [{ type: 'deleteSheet', sheet: 1 }]);
  assert.deepEqual(now(doc.wb), ['A1|list|#REF!|']);
  doc.history.undo();
  assert.deepEqual(now(doc.wb), ['A1|list|Lists!$A$1:$A$3|']);
  assert.deepEqual(now(doc.wb, 1), ['A1|whole|1|2', 'B1|custom|Data!A1>0|']);
});

test('a sheet takes no more rules than a file may hold', async () => {
  const doc = await load('<sheetData/>');
  const many: DataValidation[] = [];
  for (let i = 0; i < 65534; i++) many.push({ ranges: [{ top: i, left: 0, bottom: i, right: 0 }], type: 'whole', operator: 'equal', formula1: '1', allowBlank: true, hideDropDown: false, showInputMessage: false, showErrorMessage: true, errorStyle: 'stop', live: true });
  apply(doc, [{ type: 'setDataValidations', sheet: 0, validations: many }]);
  assert.throws(() => setValidation(doc.wb, 0, range('C1'), WHOLE), /too many data validations/);
  // Replacing one is no more than there were.
  assert.equal(setValidation(doc.wb, 0, range('A1'), WHOLE).length, 1);
});

// ---------------------------------------------------------------------------
// Review: a relative reference that leaves the sheet comes back in at the other edge

test('a rule whose formula goes round the edge of the sheet says the same for a cell before and after its range is cut', async () => {
  // Written for B3: "the last row of column B is 1". For B4 that is one row further: round the edge, B1.
  const xml =
    '<sheetData><row r="1"><c r="B1"><v>1</v></c><c r="D1"><v>7</v></c></row><row r="2"><c r="D2"><v>8</v></c></row></sheetData><dataValidations count="2">' +
    dv('type="custom" showErrorMessage="1" sqref="B3:B4"', f1('B1048576=1')) +
    dv('type="list" showErrorMessage="1" sqref="E3:E4"', f1('D1048576')) +
    '</dataValidations>';
  const doc = await load(xml);
  const engine = new CalcEngine(doc.wb);
  const rules = new DataValidations(doc.wb, doc.wb.sheets[0], { engine: () => engine });
  const ok = (ref: string, value: number) => {
    const { row, col } = parseCellRef(ref)!;
    return rules.check(row, col, { kind: 'number', value, style: 0 }, String(value)).ok;
  };
  // B3 reads the last row, which is empty; B4 reads B1, which is 1.
  assert.deepEqual([ok('B3', 5), ok('B4', 5)], [false, true]);
  // The list of E3 is the last cell of column D, which is empty; that of E4 is one row further: round the edge, D1.
  assert.deepEqual(rules.list(2, 4), { items: [], truncated: false });
  assert.deepEqual(rules.list(3, 4), { items: ['7'], truncated: false });
  assert.equal(rules.hasList(3, 4), true);
  assert.deepEqual([ok('E4', 7), ok('E4', 8)], [true, false]);
  // The rule itself is still stated for its first cell.
  assert.equal(rules.ruleAt(3, 4)!.formula1, 'D1048576');
  apply(doc, setValidation(doc.wb, 0, range('B3'), undefined));
  apply(doc, setValidation(doc.wb, 0, range('E3'), undefined));
  assert.deepEqual(now(doc.wb), ['B4|custom|B1=1|', 'E4|list|D1|']);
  assert.equal(ok('B4', 5), true, 'what the rule says for B4 is what it said');
  assert.deepEqual([ok('E4', 7), ok('E4', 8)], [true, false]);
});

test('a list whose corners change sides when the rule is read for another cell is the same list as when the rule is written for that cell', async () => {
  // Written for A3: C1 to (column E, the row above). For G3 the first corner is six columns on, right of the second.
  const xml =
    '<sheetData><row r="1"><c r="E1"><v>1</v></c><c r="F1"><v>2</v></c></row><row r="2"><c r="E2"><v>3</v></c><c r="I2"><v>4</v></c></row></sheetData><dataValidations count="1">' +
    dv('type="list" showErrorMessage="1" sqref="A3 G3"', f1('C$1:$E2')) +
    '</dataValidations>';
  const doc = await load(xml);
  const engine = new CalcEngine(doc.wb);
  const rules = new DataValidations(doc.wb, doc.wb.sheets[0], { engine: () => engine });
  assert.deepEqual(rules.list(2, 6), { items: ['1', '2', '3', '4'], truncated: false });
  apply(doc, setValidation(doc.wb, 0, range('A3'), undefined));
  assert.deepEqual(now(doc.wb), ['G3|list|I$1:$E2|']);
  assert.deepEqual(rules.list(2, 6), { items: ['1', '2', '3', '4'], truncated: false });
});
