// Conditional formats as the editor changes them: rules added, changed, deleted and
// reordered, ranges cleared, what is refused, undo and redo, the format a rule gives, and
// what rows, columns and sheets that move do to a rule that was changed.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine } from '../src/calc/engine';
import { ConditionalFormats } from '../src/format/conditional';
import { parseCellRef, parseRange, rangeName } from '../src/index';
import type { CellRange } from '../src/index';
import { copyRange, pasteAt } from '../src/ops/commands';
import { setComment } from '../src/ops/notes';
import type { Op } from '../src/ops/ops';
import { sortRange } from '../src/ops/sort';
import { addConditionalRule, clearConditionalRules, conditionalRules, moveConditionalRule, removeConditionalRule, setValidation, updateConditionalRule } from '../src/ops/rules';
import { copySheet } from '../src/ops/sheets';
import { conditionalBook, conditionalNow as now, reviewBook } from './conditional-fixtures';
import { load as loadBytes, type Loaded } from './rules-fixtures';

const range = (ref: string): CellRange => parseRange(ref)!;
const apply = (doc: Loaded, ops: Op[]) => doc.history.apply(ops, undefined, undefined);
const load = async (options?: Parameters<typeof conditionalBook>[0]) => loadBytes(await conditionalBook(options));

const FILE = [
  'A1:A10 F1:F3|expression|1|AND(A1>B1,Lists!$A$1<>A1)',
  'A1:A10 F1:F3|cellIs|2|$H$1|greaterThan',
  'B1:B5|colorScale|3||min;$H$2;max',
  'C1:C5|dataBar|4|',
  'D1:D5|iconSet|5|',
  'D1:D5|timePeriod|6|FLOOR(D1,1)=TODAY()',
];
const RED = { fill: '#FFC7CE', color: '#9C0006' };

function styleAt(doc: Loaded, ref: string, sheet = 0) {
  const engine = new CalcEngine(doc.wb);
  const formats = new ConditionalFormats(doc.wb, doc.wb.sheets[sheet], { engine: () => engine });
  const { row, col } = parseCellRef(ref)!;
  const found = formats.resolve(row, col);
  return found ? found.style : undefined;
}

test('the rules of a sheet are listed in order of priority, with what can be done to each', async () => {
  const doc = await load();
  assert.deepEqual(now(doc.wb), FILE);
  const rules = conditionalRules(doc.wb, 0);
  assert.deepEqual(rules.map((info) => [info.index, info.rule.type, info.ranges.map(rangeName).join(' '), info.editable, info.removable, info.movable]), [
    [0, 'expression', 'A1:A10 F1:F3', true, true, true],
    [1, 'cellIs', 'A1:A10 F1:F3', true, true, true],
    [2, 'colorScale', 'B1:B5', true, true, true],
    // A data bar that has more to say in the extension list: what it says stays; it can be moved, put elsewhere and deleted.
    [3, 'dataBar', 'C1:C5', false, true, true],
    // An icon set of the main part and a time period are made and changed here like the others (see rules-iconset-date.test.ts).
    [4, 'iconSet', 'D1:D5', true, true, true],
    [5, 'timePeriod', 'D1:D5', true, true, true],
  ]);
  assert.equal(rules[1].format!.fill!.type, 'pattern');
  assert.deepEqual(conditionalRules(doc.wb, 0, range('B2:C2')).map((info) => info.index), [2, 3]);
  assert.deepEqual(conditionalRules(doc.wb, 2), []);
});

test('a rule that is added comes first, applies at once, is one undo step, and marks only the formats as changed', async () => {
  const doc = await load();
  const sheet = doc.wb.sheets[0];
  const before = sheet.conditionalFormats;
  assert.equal(styleAt(doc, 'G2'), undefined);
  apply(doc, addConditionalRule(doc.wb, 0, range('G1:G3'), { type: 'cellIs', operator: 'lessThan', formulas: ['=$H$1'], format: RED }));
  assert.deepEqual(now(doc.wb), ['G1:G3|cellIs|1|$H$1|lessThan', ...FILE.map((line) => line.replace(/\|(\d)\|/, (_, n) => `|${Number(n) + 1}|`))]);
  assert.equal(sheet.conditionalModified, true);
  assert.equal(sheet.modified, undefined, 'no cell changed');
  // G2 is empty, which counts as less than 3.
  assert.deepEqual(styleAt(doc, 'G2'), { color: '#9C0006', fill: '#FFC7CE' });
  const added = sheet.conditionalFormats![sheet.conditionalFormats!.length - 1];
  assert.equal(added.live, true);
  assert.equal(added.rules[0].src, undefined);
  // The formats of the file are the formats of the file: only the number of each rule is another one.
  assert.equal(sheet.conditionalFormats![0].live, undefined);
  assert.deepEqual(sheet.conditionalFormats![0].rules.map((rule) => [rule.src, rule.priority]), [[[0, 0], 3], [[0, 1], 2]]);
  doc.history.undo();
  assert.equal(sheet.conditionalFormats, before, 'the very list of the file is back');
  assert.equal(styleAt(doc, 'G2'), undefined);
  doc.history.redo();
  assert.equal(now(doc.wb)[0], 'G1:G3|cellIs|1|$H$1|lessThan');
});

test('every kind of rule that can be made, with the formulas Excel writes for it', async () => {
  const doc = await load({ bare: true });
  const made = (settings: Parameters<typeof addConditionalRule>[3], at = 'C3:D9') => {
    const ops = addConditionalRule(doc.wb, 0, range(at), settings);
    const formats = (ops[0] as Extract<Op, { type: 'setConditionalFormats' }>).formats!;
    return formats[formats.length - 1].rules[0];
  };
  for (const operator of ['between', 'notBetween'] as const) assert.deepEqual([made({ type: 'cellIs', operator, formulas: ['1', '=A1*2'] }).operator, made({ type: 'cellIs', operator, formulas: ['1', '=A1*2'] }).formulas], [operator, ['1', 'A1*2']]);
  for (const operator of ['equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'] as const) assert.deepEqual(made({ type: 'cellIs', operator, formulas: ['5', 'ignored'] }).formulas, ['5']);
  assert.deepEqual([made({ type: 'containsText', text: 'a"b' }).formulas, made({ type: 'containsText', text: 'a"b' }).operator, made({ type: 'containsText', text: 'a"b' }).text], [['NOT(ISERROR(SEARCH("a""b",C3)))'], 'containsText', 'a"b']);
  assert.deepEqual([made({ type: 'notContainsText', text: 'x' }).formulas, made({ type: 'notContainsText', text: 'x' }).operator], [['ISERROR(SEARCH("x",C3))'], 'notContains']);
  assert.deepEqual([made({ type: 'beginsWith', text: 'x' }).formulas, made({ type: 'beginsWith', text: 'x' }).operator], [['LEFT(C3,LEN("x"))="x"'], 'beginsWith']);
  assert.deepEqual([made({ type: 'endsWith', text: 'x' }).formulas, made({ type: 'endsWith', text: 'x' }).operator], [['RIGHT(C3,LEN("x"))="x"'], 'endsWith']);
  assert.deepEqual(made({ type: 'duplicateValues' }).formulas, []);
  assert.equal(made({ type: 'uniqueValues' }).type, 'uniqueValues');
  const top = made({ type: 'top10', rank: 5, percent: true, bottom: true });
  assert.deepEqual([top.rank, top.percent, top.bottom], [5, true, true]);
  assert.deepEqual([made({ type: 'top10' }).rank, made({ type: 'top10' }).percent, made({ type: 'top10' }).bottom], [10, false, false]);
  assert.deepEqual([made({ type: 'aboveAverage', aboveAverage: false, equalAverage: true }).aboveAverage, made({ type: 'aboveAverage', aboveAverage: false, equalAverage: true }).equalAverage], [false, true]);
  assert.deepEqual(made({ type: 'containsBlanks' }).formulas, ['LEN(TRIM(C3))=0']);
  assert.deepEqual(made({ type: 'notContainsBlanks' }).formulas, ['LEN(TRIM(C3))>0']);
  assert.deepEqual(made({ type: 'containsErrors' }).formulas, ['ISERROR(C3)']);
  assert.deepEqual(made({ type: 'notContainsErrors' }).formulas, ['NOT(ISERROR(C3))']);
  assert.deepEqual(made({ type: 'expression', formulas: ['=MOD(ROW(),2)=0'], stopIfTrue: true }).formulas, ['MOD(ROW(),2)=0']);
  assert.equal(made({ type: 'expression', formulas: ['=MOD(ROW(),2)=0'], stopIfTrue: true }).stopIfTrue, true);
  const two = made({ type: 'colorScale', colors: ['#F8696B', '#63BE7B'] });
  assert.deepEqual([two.colorScale!.values.map((value) => value.type), two.colorScale!.colors], [['min', 'max'], [{ rgb: 'FFF8696B' }, { rgb: 'FF63BE7B' }]]);
  const three = made({ type: 'colorScale', colors: ['#F8696B', '#FFEB84', '#63BE7B'] });
  assert.deepEqual(three.colorScale!.values.map((value) => [value.type, value.value]), [['min', undefined], ['percentile', '50'], ['max', undefined]]);
  const bar = made({ type: 'dataBar', colors: ['#638EC6'] });
  assert.deepEqual([bar.dataBar!.color, bar.dataBar!.min.type, bar.dataBar!.max.type, bar.dataBar!.extended], [{ rgb: 'FF638EC6' }, 'min', 'max', false]);
  // The format: what is stated, and nothing else.
  const formatted = made({ type: 'duplicateValues', format: { color: '#9C0006', fill: '#FFC7CE', bold: true, italic: true, underline: true, strike: true } });
  assert.deepEqual(formatted.dxf, { font: { bold: true, italic: true, strike: true, underline: 'single', color: { rgb: 'FF9C0006' } }, fill: { type: 'pattern', pattern: 'solid', bg: { rgb: 'FFFFC7CE' } } });
  assert.equal(made({ type: 'duplicateValues', format: { bold: true } }).dxf!.fill, undefined);
  assert.equal(made({ type: 'duplicateValues' }).dxf, undefined);
});

test('what a rule cannot say is refused', async () => {
  const doc = await load({ bare: true });
  const add = (settings: Parameters<typeof addConditionalRule>[3]) => addConditionalRule(doc.wb, 0, range('A1'), settings);
  assert.throws(() => add({ type: 'somethingNew' as 'cellIs' }), /kind of rule/);
  assert.throws(() => add({ type: 'iconSet', iconSet: '3Stars' as '3Arrows' }), /icon set is not known/);
  assert.throws(() => add({ type: 'timePeriod' }), /needs a period/);
  assert.throws(() => add({ type: 'timePeriod', timePeriod: 'someday' as 'today' }), /period is not known/);
  assert.throws(() => add({ type: 'cellIs', operator: 'sideways' as 'equal', formulas: ['1'] }), /operator/);
  assert.throws(() => add({ type: 'cellIs', operator: 'equal' }), /needs a value/);
  assert.throws(() => add({ type: 'cellIs', operator: 'between', formulas: ['1'] }), /needs two values/);
  assert.throws(() => add({ type: 'expression', formulas: ['SUM(A1'] }), /cannot be read/);
  assert.throws(() => add({ type: 'containsText', text: '' }), /needs a text/);
  assert.throws(() => add({ type: 'top10', rank: 0 }), /between 1 and 1000/);
  assert.throws(() => add({ type: 'top10', rank: 101, percent: true }), /between 1 and 100/);
  assert.throws(() => add({ type: 'colorScale', colors: ['#FF0000'] }), /two or three colours/);
  assert.throws(() => add({ type: 'colorScale', colors: ['red', 'blue'] }), /#RRGGBB/);
  assert.throws(() => add({ type: 'dataBar' }), /colour/);
  assert.throws(() => add({ type: 'duplicateValues', format: { fill: 'pink' } }), /#RRGGBB/);
  assert.throws(() => addConditionalRule(doc.wb, 0, { top: 0, left: 0, bottom: 0, right: 99999 }, { type: 'duplicateValues' }), /outside the sheet/);
  assert.throws(() => addConditionalRule(doc.wb, 9, range('A1'), { type: 'duplicateValues' }), /No sheet/);
});

test('a rule is changed: where it applies, whether it stops the others, what it says', async () => {
  const doc = await load();
  // The second rule of the element of the file: it leaves the element and is a rule of its own from here on.
  apply(doc, updateConditionalRule(doc.wb, 0, 1, { ranges: [range('A3:A6'), range('H5')] }));
  assert.deepEqual(now(doc.wb).slice(0, 2), ['A1:A10 F1:F3|expression|1|AND(A1>B1,Lists!$A$1<>A1)', 'A3:A6 H5|cellIs|2|$H$1|greaterThan']);
  const formats = doc.wb.sheets[0].conditionalFormats!;
  assert.deepEqual(formats[0].rules.map((rule) => rule.src), [[0, 1]], 'the other rule of the element stays in it');
  assert.equal(formats[0].live, undefined);
  const moved = formats[formats.length - 1];
  assert.deepEqual([moved.live, moved.rules[0].src, moved.rules[0].dxfId], [true, [0, 0], 0]);
  // A relative reference keeps meaning the cell it meant for each cell of the rule.
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { ranges: range('C5:C6') }));
  assert.equal(now(doc.wb)[0], 'C5:C6|expression|1|AND(C5>D5,Lists!$A$1<>C5)');
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { stopIfTrue: false }));
  assert.equal(conditionalRules(doc.wb, 0)[0].rule.stopIfTrue, false);
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { rule: { type: 'cellIs', operator: 'between', formulas: ['1', '2'], format: { bold: false, fill: '#00FF00' } } }));
  assert.equal(now(doc.wb)[0], 'C5:C6|cellIs|1|1;2|between');
  // What the rule said before is not kept beside what it says now, its place in the file is.
  const changed = conditionalRules(doc.wb, 0)[0];
  assert.deepEqual([changed.rule.src, changed.rule.stopIfTrue], [[0, 1], false]);
  assert.deepEqual(changed.format, { font: { bold: false }, fill: { type: 'pattern', pattern: 'solid', bg: { rgb: 'FF00FF00' } } });
  // Nothing to change is no step.
  assert.deepEqual(updateConditionalRule(doc.wb, 0, 0, { stopIfTrue: false }), []);
  assert.deepEqual(updateConditionalRule(doc.wb, 0, 0, { ranges: range('C5:C6') }), []);
  for (let i = 0; i < 4; i++) doc.history.undo();
  assert.deepEqual(now(doc.wb), FILE);
});

test('a rule whose format had more than colours keeps the rest when the colours change', async () => {
  const doc = await load();
  // Format 2 (that of the time period rule) has italic, a size, a green fill, an alignment, a border: it is given to a "cell value" rule, by a file that says so.
  doc.wb.sheets[0].conditionalFormats![0].rules[0].dxfId = 2;
  apply(doc, updateConditionalRule(doc.wb, 0, 1, { rule: { type: 'cellIs', operator: 'greaterThan', formulas: ['$H$1'], format: { fill: '#FFC7CE', bold: true } } }));
  const info = conditionalRules(doc.wb, 0)[1];
  assert.deepEqual(info.format, { font: { bold: true }, fill: { type: 'pattern', pattern: 'solid', bg: { rgb: 'FFFFC7CE' } }, border: { left: { style: 'thin', color: { rgb: 'FF000000' } } } });
  assert.equal(info.rule.dxfId, 2, 'the record it started from is still named: the writer takes the rest from it');
});

test('rules are deleted and reordered; the numbers of the others follow', async () => {
  const doc = await load();
  apply(doc, removeConditionalRule(doc.wb, 0, 0));
  assert.deepEqual(now(doc.wb), ['A1:A10 F1:F3|cellIs|1|$H$1|greaterThan', 'B1:B5|colorScale|2||min;$H$2;max', 'C1:C5|dataBar|3|', 'D1:D5|iconSet|4|', 'D1:D5|timePeriod|5|FLOOR(D1,1)=TODAY()']);
  assert.equal(doc.wb.sheets[0].conditionalFormats![0].live, undefined, 'the element of the file lost a rule and is the element of the file');
  // The icon set, which is only kept, goes as well; the last rule moves to the top.
  apply(doc, removeConditionalRule(doc.wb, 0, 3));
  apply(doc, moveConditionalRule(doc.wb, 0, 3, 0));
  assert.deepEqual(now(doc.wb), ['D1:D5|timePeriod|1|FLOOR(D1,1)=TODAY()', 'A1:A10 F1:F3|cellIs|2|$H$1|greaterThan', 'B1:B5|colorScale|3||min;$H$2;max', 'C1:C5|dataBar|4|']);
  apply(doc, moveConditionalRule(doc.wb, 0, 0, 2));
  assert.deepEqual(now(doc.wb).map((line) => line.split('|')[1]), ['cellIs', 'colorScale', 'timePeriod', 'dataBar']);
  assert.deepEqual(moveConditionalRule(doc.wb, 0, 2, 2), []);
  assert.throws(() => moveConditionalRule(doc.wb, 0, 2, 9), /No rule/);
  assert.throws(() => removeConditionalRule(doc.wb, 0, 9), /No rule/);
  // The data bar has a counterpart in the extension list: what it says is kept as it is; where it applies is not.
  assert.throws(() => updateConditionalRule(doc.wb, 0, 3, { rule: { type: 'dataBar', colors: ['#FF0000'] } }), /extension list/);
  apply(doc, updateConditionalRule(doc.wb, 0, 3, { ranges: range('C1:C9') }));
  assert.equal(now(doc.wb)[3], 'C1:C9|dataBar|4|');
  const bar = conditionalRules(doc.wb, 0)[3].rule;
  assert.deepEqual([bar.extId !== undefined, bar.dataBar!.extended, bar.src], [true, true, [2, 0]], 'it is the bar of the file, with what the extension list says of it');
  apply(doc, removeConditionalRule(doc.wb, 0, 3));
  assert.deepEqual(now(doc.wb).map((line) => line.split('|')[1]), ['cellIs', 'colorScale', 'timePeriod']);
  for (let i = 0; i < 6; i++) doc.history.undo();
  assert.deepEqual(now(doc.wb), FILE);
});

test('clearing: the cells of a range leave every rule; the whole sheet loses them all', async () => {
  const doc = await load();
  apply(doc, clearConditionalRules(doc.wb, 0, range('A1:B2')));
  assert.deepEqual(now(doc.wb).slice(0, 3), ['A3:A10 F1:F3|expression|1|AND(A3>B3,Lists!$A$1<>A3)', 'A3:A10 F1:F3|cellIs|2|$H$1|greaterThan', 'B3:B5|colorScale|3||min;$H$2;max']);
  // The two rules of the element still share their cells: one element with both.
  const formats = doc.wb.sheets[0].conditionalFormats!;
  assert.equal(formats.filter((format) => format.live).length, 2);
  assert.deepEqual(clearConditionalRules(doc.wb, 0, range('J1:K9')), []);
  // The data bar that the extension list says more about loses cells like any other rule.
  apply(doc, clearConditionalRules(doc.wb, 0, range('C1')));
  assert.equal(now(doc.wb)[3], 'C2:C5|dataBar|4|');
  apply(doc, clearConditionalRules(doc.wb, 0));
  assert.equal(doc.wb.sheets[0].conditionalFormats, undefined);
  // Not so a rule that only the extension list has.
  const only = await load({ extensionOnly: true });
  assert.throws(() => clearConditionalRules(only.wb, 0, range('E1')), /extension list/);
  assert.throws(() => clearConditionalRules(only.wb, 0), /extension list/);
  apply(doc, clearConditionalRules(doc.wb, 1));
  assert.equal(doc.wb.sheets[1].conditionalFormats, undefined);
  assert.deepEqual(clearConditionalRules(doc.wb, 1), []);
});

test('a sheet with a rule only the extension list has: new rules come last, and the order stays as it is', async () => {
  const doc = await load({ extensionOnly: true });
  const rules = conditionalRules(doc.wb, 0);
  assert.deepEqual([rules[6].rule.type, rules[6].editable, rules[6].removable, rules[6].movable], ['iconSet', false, false, false]);
  assert.ok(rules.every((info) => !info.movable));
  apply(doc, addConditionalRule(doc.wb, 0, range('G1'), { type: 'duplicateValues' }));
  assert.deepEqual(now(doc.wb), [...FILE, 'E1:E5|iconSet|7|', 'G1|duplicateValues|8|']);
  assert.throws(() => moveConditionalRule(doc.wb, 0, 0, 1), /extension list/);
  assert.throws(() => removeConditionalRule(doc.wb, 0, 6), /extension list/);
  // A rule of the main part goes; the numbers of the others are left alone.
  apply(doc, removeConditionalRule(doc.wb, 0, 0));
  assert.deepEqual(now(doc.wb), [...FILE.slice(1), 'E1:E5|iconSet|7|', 'G1|duplicateValues|8|']);
});

// ---------------------------------------------------------------------------
// A rule that was changed, and what moves afterwards

const insert = (sheet: number, axis: 'row' | 'col', at: number, count: number): Op[] => [{ type: 'insert', sheet, axis, at, count }];
const remove = (sheet: number, axis: 'row' | 'col', at: number, count: number): Op[] => [{ type: 'delete', sheet, axis, at, count }];

test('rows and columns that move take a changed rule with them, as they take a rule of the file', async () => {
  const doc = await load();
  // One rule of the file is changed (it is live from here on), one is added, the others stay as the file has them.
  apply(doc, updateConditionalRule(doc.wb, 0, 2, { ranges: range('B2:B6') }));
  apply(doc, addConditionalRule(doc.wb, 0, range('G4:G6'), { type: 'expression', formulas: ['AND(G4>$H$1,F4<>"")'] }));
  apply(doc, addConditionalRule(doc.wb, 0, range('E8:E9'), { type: 'colorScale', colors: ['#FF0000', '#00FF00'] }));
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { ranges: range('E8:E9') }));
  assert.deepEqual(now(doc.wb), ['E8:E9|colorScale|1||min;max', 'G4:G6|expression|2|AND(G4>$H$1,F4<>"")', 'A1:A10 F1:F3|expression|3|AND(A1>B1,Lists!$A$1<>A1)', 'A1:A10 F1:F3|cellIs|4|$H$1|greaterThan', 'B2:B6|colorScale|5||min;$H$2;max', ...FILE.slice(3).map((line) => line.replace(/\|(\d)\|/, (_, n) => `|${Number(n) + 2}|`))]);
  apply(doc, insert(0, 'row', 0, 2));
  assert.deepEqual(now(doc.wb).slice(0, 5), ['E10:E11|colorScale|1||min;max', 'G6:G8|expression|2|AND(G6>$H$3,F6<>"")', 'A3:A12 F3:F5|expression|3|AND(A3>B3,Lists!$A$1<>A3)', 'A3:A12 F3:F5|cellIs|4|$H$3|greaterThan', 'B4:B8|colorScale|5||min;$H$4;max']);
  apply(doc, insert(0, 'col', 0, 1));
  assert.deepEqual(now(doc.wb).slice(0, 5), ['F10:F11|colorScale|1||min;max', 'H6:H8|expression|2|AND(H6>$I$3,G6<>"")', 'B3:B12 G3:G5|expression|3|AND(B3>C3,Lists!$A$1<>B3)', 'B3:B12 G3:G5|cellIs|4|$I$3|greaterThan', 'C4:C8|colorScale|5||min;$I$4;max']);
  // All its cells deleted: the rule is gone, and undo brings it back.
  apply(doc, remove(0, 'row', 5, 3));
  assert.deepEqual(now(doc.wb).slice(0, 4), ['F7:F8|colorScale|1||min;max', 'B3:B9 G3:G5|expression|3|AND(B3>C3,Lists!$A$1<>B3)', 'B3:B9 G3:G5|cellIs|4|$I$3|greaterThan', 'C4:C5|colorScale|5||min;$I$4;max']);
  doc.history.undo();
  assert.equal(now(doc.wb)[1], 'H6:H8|expression|2|AND(H6>$I$3,G6<>"")');
  for (let i = 0; i < 2; i++) doc.history.undo();
  assert.equal(now(doc.wb)[1], 'G4:G6|expression|2|AND(G4>$H$1,F4<>"")');
  for (let i = 0; i < 3; i++) doc.history.redo();
  assert.equal(now(doc.wb)[0], 'F7:F8|colorScale|1||min;max');
});

test('a changed rule that reads another sheet follows the rows of that sheet and its name', async () => {
  const doc = await load();
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { stopIfTrue: false }));
  apply(doc, addConditionalRule(doc.wb, 0, range('G1'), { type: 'cellIs', operator: 'between', formulas: ['Lists!$A$2', 'Other!A1'] }));
  apply(doc, insert(1, 'row', 0, 3));
  assert.deepEqual(now(doc.wb).slice(0, 2), ['G1|cellIs|1|Lists!$A$5;Other!A1|between', 'A1:A10 F1:F3|expression|2|AND(A1>B1,Lists!$A$4<>A1)']);
  apply(doc, [{ type: 'renameSheet', sheet: 1, name: 'My Lists' }]);
  assert.deepEqual(now(doc.wb).slice(0, 2), ["G1|cellIs|1|'My Lists'!$A$5;Other!A1|between", "A1:A10 F1:F3|expression|2|AND(A1>B1,'My Lists'!$A$4<>A1)"]);
  // A sheet that takes the old name is not the sheet the rules read.
  apply(doc, [{ type: 'renameSheet', sheet: 2, name: 'Lists' }]);
  assert.deepEqual(now(doc.wb).slice(0, 2), ["G1|cellIs|1|'My Lists'!$A$5;Lists!A1|between", "A1:A10 F1:F3|expression|2|AND(A1>B1,'My Lists'!$A$4<>A1)"]);
  apply(doc, [{ type: 'deleteSheet', sheet: 1 }]);
  assert.deepEqual(now(doc.wb).slice(0, 2), ['G1|cellIs|1|#REF!;Lists!A1|between', 'A1:A10 F1:F3|expression|2|AND(A1>B1,#REF!<>A1)']);
  for (let i = 0; i < 4; i++) doc.history.undo();
  assert.deepEqual(now(doc.wb).slice(0, 2), ['G1|cellIs|1|Lists!$A$2;Other!A1|between', 'A1:A10 F1:F3|expression|2|AND(A1>B1,Lists!$A$1<>A1)']);
});

test('a threshold that is a formula follows what it names, in a rule that was changed as in one of the file', async () => {
  const doc = await load();
  apply(doc, updateConditionalRule(doc.wb, 0, 2, { ranges: range('B1:B4') }));
  apply(doc, insert(0, 'row', 0, 1));
  assert.equal(now(doc.wb)[2], 'B2:B5|colorScale|3||min;$H$3;max');
  apply(doc, [{ type: 'renameSheet', sheet: 0, name: 'Numbers' }]);
  assert.equal(now(doc.wb)[2], 'B2:B5|colorScale|3||min;$H$3;max');
});

// ---------------------------------------------------------------------------
// With the other things that happen to cells

test('sorting, pasting, merging and notes leave the rules where they are; the format follows the values', async () => {
  const doc = await load({ bare: true });
  apply(doc, addConditionalRule(doc.wb, 0, range('A1:A3'), { type: 'cellIs', operator: 'greaterThan', formulas: ['4'], format: RED }));
  const before = now(doc.wb);
  // A is 5, 1, 9: the first and the third match.
  assert.deepEqual(['A1', 'A2', 'A3'].map((ref) => styleAt(doc, ref) !== undefined), [true, false, true]);
  // A sort moves cells, not rules (Excel does not move them either): 1, 5, 9 afterwards.
  apply(doc, sortRange(doc.wb, 0, range('A1:B3'), [{ col: 0, descending: false }], { header: false }));
  assert.deepEqual(now(doc.wb), before);
  assert.deepEqual(['A1', 'A2', 'A3'].map((ref) => styleAt(doc, ref) !== undefined), [false, true, true]);
  // Pasted text brings no rule, and the rules where it lands stay. (Copied cells bring theirs: see paste-rules.test.ts.)
  const copied = copyRange(doc.wb, 0, range('A3:B3'));
  apply(doc, pasteAt(doc.wb, 0, range('D1'), copied.text).ops);
  apply(doc, pasteAt(doc.wb, 0, range('A1'), copied.text).ops);
  assert.deepEqual(now(doc.wb), before);
  assert.equal(styleAt(doc, 'D1'), undefined);
  assert.notEqual(styleAt(doc, 'A1'), undefined, 'the pasted 9 is above 4');
  apply(doc, [{ type: 'setMerges', sheet: 0, merges: [range('A2:A3')] }]);
  apply(doc, setComment(doc.wb, 0, 0, 0, 'A note'));
  assert.deepEqual(now(doc.wb), before);
  // A value typed is judged at once: the caches of the rules are those of the view, which drops them after every step.
  apply(doc, [{ type: 'setCells', sheet: 0, cells: [{ row: 0, col: 0, cell: { kind: 'number', value: 1, style: 0 } }] }]);
  assert.equal(styleAt(doc, 'A1'), undefined);
});

// ---------------------------------------------------------------------------
// Review: names of sheets in thresholds and in copies

test('a threshold that names a sheet follows the name of that sheet, and names the copy in a copy of it', async () => {
  const doc = await load({ bare: true });
  // A scale whose midpoint is a cell of this very sheet, by name; a rule and a validation that name it as well.
  apply(doc, addConditionalRule(doc.wb, 0, range('A1:A3'), { type: 'colorScale', colors: ['#FF0000', '#00FF00', '#0000FF'] }));
  const formats = doc.wb.sheets[0].conditionalFormats!;
  const scale = formats[0].rules[0];
  apply(doc, [{ type: 'setConditionalFormats', sheet: 0, formats: [{ ...formats[0], rules: [{ ...scale, colorScale: { ...scale.colorScale!, values: [scale.colorScale!.values[0], { type: 'formula', value: 'Data!$H$2+Lists!$A$1', gte: true }, scale.colorScale!.values[2]] } }] }] }]);
  apply(doc, addConditionalRule(doc.wb, 0, range('B1:B3'), { type: 'expression', formulas: ['Data!$H$1>B1'] }));
  apply(doc, setValidation(doc.wb, 0, range('C1'), { type: 'list', formula1: 'Data!$H$1:$H$2' }));
  const midpoint = (sheet: number) => doc.wb.sheets[sheet].conditionalFormats!.map((format) => format.rules[0]).filter((rule) => rule.colorScale)[0].colorScale!.values[1].value;
  assert.equal(midpoint(0), 'Data!$H$2+Lists!$A$1');

  apply(doc, copySheet(doc.wb, 0, { name: 'Copy' }));
  assert.equal(midpoint(1), 'Copy!$H$2+Lists!$A$1', 'the threshold of the copy reads the copy');
  assert.equal(now(doc.wb, 1)[0], 'B1:B3|expression|1|Copy!$H$1>B1');
  assert.equal(doc.wb.sheets[1].dataValidations![0].formula1, 'Copy!$H$1:$H$2');
  assert.equal(midpoint(0), 'Data!$H$2+Lists!$A$1');

  apply(doc, [{ type: 'renameSheet', sheet: 2, name: 'Sources' }]);
  assert.equal(midpoint(0), 'Data!$H$2+Sources!$A$1');
  assert.equal(midpoint(1), 'Copy!$H$2+Sources!$A$1');
  apply(doc, [{ type: 'renameSheet', sheet: 0, name: 'Numbers' }]);
  assert.equal(midpoint(0), 'Numbers!$H$2+Sources!$A$1');
  doc.history.undo();
  doc.history.undo();
  assert.equal(midpoint(0), 'Data!$H$2+Lists!$A$1');
});

// ---------------------------------------------------------------------------
// Review: a rule that is restated keeps everything that was not restated

const loadReview = async () => loadBytes(await reviewBook());
const ruleAt = (doc: Loaded, index: number) => conditionalRules(doc.wb, 0)[index];

test('a rule stated again as it is listed is no step, whatever it says beside what a dialog can show', async () => {
  const doc = await loadReview();
  const settings: Parameters<typeof updateConditionalRule>[3]['rule'][] = [
    { type: 'colorScale', colors: ['#F8696B', '#4472C4', '#63BE7B'] },
    { type: 'dataBar', colors: ['#638EC6'] },
    { type: 'aboveAverage', aboveAverage: false, equalAverage: false },
    { type: 'containsText', text: 'ab' },
    { type: 'cellIs', operator: 'greaterThan', formulas: ['SUM( $H$1:$H$2 )'] },
    { type: 'cellIs', operator: 'between', formulas: ['1', 'SUM( $H$1:$H$2 )'] },
    { type: 'cellIs', operator: 'equal', formulas: ['"x"'] },
    { type: 'cellIs', operator: 'equal', formulas: ['5'], format: { fill: '#FFC7CE', color: '#000000' } },
    { type: 'beginsWith', text: 'q' },
  ];
  settings.forEach((rule, index) => {
    assert.deepEqual(updateConditionalRule(doc.wb, 0, index, { rule }), [], `rule ${index}`);
    assert.deepEqual(updateConditionalRule(doc.wb, 0, index, { rule: { type: rule!.type } }), [], `rule ${index}, by its kind alone`);
    assert.deepEqual(updateConditionalRule(doc.wb, 0, index, { rule, stopIfTrue: false, ranges: ruleAt(doc, index).ranges }), [], `rule ${index}, with its cells`);
  });
});

test('one thing changed is the one thing that changes: thresholds, lengths, deviations and the spelling of formulas stay', async () => {
  const doc = await loadReview();
  apply(doc, updateConditionalRule(doc.wb, 0, 0, { rule: { type: 'colorScale', colors: ['#FF0000', '#4472C4', '#63BE7B'] } }));
  const scale = ruleAt(doc, 0).rule.colorScale!;
  assert.deepEqual(scale.values.map((value) => [value.type, value.value]), [['num', '0'], ['formula', '$H$2'], ['percent', '90']]);
  assert.deepEqual(scale.colors, [{ rgb: 'FFFF0000' }, { theme: 4 }, { rgb: 'FF63BE7B' }], 'the colour that was not touched is the colour of the file');
  // Two colours instead of three is another scale: lowest to highest.
  assert.deepEqual(updateConditionalRule(doc.wb, 0, 0, { rule: { type: 'colorScale', colors: ['#FF0000', '#63BE7B'] } }).length, 1);

  apply(doc, updateConditionalRule(doc.wb, 0, 1, { rule: { type: 'dataBar', colors: ['#00B050'] } }));
  const bar = ruleAt(doc, 1).rule.dataBar!;
  assert.deepEqual([bar.minLength, bar.maxLength, bar.showValue, bar.min, bar.color], [5, 80, false, { type: 'num', value: '0', gte: true }, { rgb: 'FF00B050' }]);

  apply(doc, updateConditionalRule(doc.wb, 0, 2, { rule: { type: 'aboveAverage', equalAverage: true } }));
  const average = ruleAt(doc, 2).rule;
  assert.deepEqual([average.aboveAverage, average.equalAverage, average.stdDev], [false, true, 2]);

  // The text looked for is another one: the formula is written anew, for the cell the old one named.
  apply(doc, updateConditionalRule(doc.wb, 0, 3, { rule: { type: 'containsText', text: 'cd' } }));
  assert.deepEqual([ruleAt(doc, 3).rule.formulas, ruleAt(doc, 3).rule.text], [['NOT(ISERROR(SEARCH("cd",$A1)))'], 'cd']);
  // A formula of a text rule that is not the one Excel writes is replaced when the text is.
  apply(doc, updateConditionalRule(doc.wb, 0, 8, { rule: { type: 'beginsWith', text: 'z' } }));
  assert.deepEqual(ruleAt(doc, 8).rule.formulas, ['LEFT(A1,LEN("z"))="z"']);

  // One value of two: the other keeps its spelling.
  apply(doc, updateConditionalRule(doc.wb, 0, 5, { rule: { type: 'cellIs', operator: 'between', formulas: ['2', 'SUM( $H$1:$H$2 )'] } }));
  assert.deepEqual(ruleAt(doc, 5).rule.formulas, ['2', 'SUM( $H$1:$H$2 )']);
  apply(doc, updateConditionalRule(doc.wb, 0, 5, { rule: { type: 'cellIs', operator: 'notBetween' } }));
  assert.deepEqual([ruleAt(doc, 5).rule.operator, ruleAt(doc, 5).rule.formulas], ['notBetween', ['2', 'SUM( $H$1:$H$2 )']]);
});

test('a format the dialog cannot show is not a format to drop; what is changed of a format is what changes', async () => {
  const doc = await loadReview();
  // A border alone, a number format alone, a gradient alone: the rule says something else and gives what it gave.
  for (const [index, id] of [[4, 0], [5, 1], [6, 2]] as const) {
    apply(doc, updateConditionalRule(doc.wb, 0, index, { rule: { type: 'cellIs', operator: 'lessThan', formulas: ['3'] } }));
    const info = ruleAt(doc, index);
    assert.deepEqual([info.rule.operator, info.rule.dxfId, info.rule.dxf], ['lessThan', id, undefined], `rule ${index}`);
    assert.deepEqual(info.format, doc.wb.styles.dxfFormats![id]);
  }
  // A fill for the rule with the border: the border stays. For the rule with the gradient: the fill is the new one.
  apply(doc, updateConditionalRule(doc.wb, 0, 4, { format: { fill: '#00FF00' } }));
  assert.deepEqual(ruleAt(doc, 4).format, { fill: { type: 'pattern', pattern: 'solid', bg: { rgb: 'FF00FF00' } }, border: { left: { style: 'thin', color: { rgb: 'FF000000' } } } });
  // Bold stated as off, and a colour of the theme: a new fill changes neither.
  apply(doc, updateConditionalRule(doc.wb, 0, 7, { format: { fill: '#C6EFCE', color: '#000000' } }));
  assert.deepEqual(ruleAt(doc, 7).format!.font, { bold: false, color: { theme: 1 } });
  assert.deepEqual((ruleAt(doc, 7).format!.fill as { bg: unknown }).bg, { rgb: 'FFC6EFCE' });
  // Bold switched on is bold; switched off again it is not stated.
  apply(doc, updateConditionalRule(doc.wb, 0, 7, { format: { fill: '#C6EFCE', color: '#000000', bold: true } }));
  assert.equal(ruleAt(doc, 7).format!.font!.bold, true);
  apply(doc, updateConditionalRule(doc.wb, 0, 7, { format: { fill: '#C6EFCE', color: '#000000' } }));
  assert.equal(ruleAt(doc, 7).format!.font!.bold, undefined);
  // No format at all is said as that.
  apply(doc, updateConditionalRule(doc.wb, 0, 6, { format: null }));
  assert.deepEqual([ruleAt(doc, 6).format, ruleAt(doc, 6).rule.dxfId], [undefined, undefined]);
  assert.deepEqual(updateConditionalRule(doc.wb, 0, 6, { format: null }), []);
});
