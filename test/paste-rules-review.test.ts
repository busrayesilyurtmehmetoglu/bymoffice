// What a review of rules that travel with their cells found, each with the test
// that shows it: the order the rules of a pasted cell are applied in, rules that
// cannot be copied or changed (they stay, and nothing is refused), what
// `skippedRules` counts, ranges that are joined again, and a paste among
// thousands of rules.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { shiftFormula } from '../src/calc/refs';
import { CalcEngine } from '../src/calc/engine';
import { ConditionalFormats } from '../src/format/conditional';
import { parseCellRef, parseRange, rangeName } from '../src/index';
import type { CellRange, Workbook } from '../src/index';
import { copyRange, cutRange, pasteAt } from '../src/ops/commands';
import type { Op } from '../src/ops/ops';
import { conditionalRules, dataValidationNow, joinedRanges, ruleSpace } from '../src/ops/rules';
import { BLOCK_A, BLOCK_B, BLOCK_C, BLOCK_D, CELLS as BOOK_CELLS, EXT_BAR, EXT_ONLY, STYLES, cf, conditionalNow, extLst } from './conditional-fixtures';
import { buildXlsx, worksheetXml } from './fixtures';
import { checkPackage, load, partText, reopen, save, type Loaded } from './rules-fixtures';

const range = (ref: string): CellRange => parseRange(ref)!;
const SHEET1 = 'xl/worksheets/sheet1.xml';
const apply = (doc: Loaded, ops: Op[]) => {
  if (ops.length > 0) doc.history.apply(ops, undefined, undefined);
};
const ruleOps = (ops: Op[]) => ops.filter((op) => op.type !== 'setCells').map((op) => op.type);
const rulesOf = (xml: string) => (xml.match(/<conditionalFormatting[\s\S]*?<\/conditionalFormatting>|<dataValidations[\s\S]*?<\/dataValidations>|<extLst>[\s\S]*<\/extLst>/g) || []).join('\n');

const validations = (wb: Workbook, sheet = 0) => {
  const space = ruleSpace(wb, wb.sheets[sheet]);
  return (wb.sheets[sheet].dataValidations || [])
    .map((rule) => dataValidationNow(rule, space))
    .filter((rule) => !!rule)
    .map((rule) => `${rule!.ranges.map(rangeName).join(' ')}|${rule!.type}|${rule!.formula1 ?? ''}${rule!.extended ? '|x14' : ''}`);
};

/** The conditional rules that apply to a cell, in the order they are applied in, each as if the cell were A1. */
function applied(wb: Workbook, ref: string, sheet = 0): string[] {
  return conditionalRules(wb, sheet, range(ref)).map((info) => {
    const rule = info.rule;
    return `${rule.type}|${rule.operator ?? ''}|${rule.formulas.map((text) => shiftFormula(text, -info.ranges[0].top, -info.ranges[0].left, true)).join(';')}|${rule.dxfId ?? ''}|${rule.stopIfTrue}`;
  });
}

/** What the conditional formats give a cell. */
function styleAt(wb: Workbook, ref: string, sheet = 0): string | undefined {
  const engine = new CalcEngine(wb);
  const formats = new ConditionalFormats(wb, wb.sheets[sheet], { engine: () => engine });
  const { row, col } = parseCellRef(ref)!;
  const found = formats.resolve(row, col);
  return found ? JSON.stringify(found.style) : undefined;
}

const R = (priority: number) => `<cfRule type="cellIs" dxfId="0" priority="${priority}" operator="greaterThan"><formula>0</formula></cfRule>`;
const S = (priority: number) => `<cfRule type="cellIs" dxfId="2" priority="${priority}" operator="greaterThan"><formula>1</formula></cfRule>`;
const FIVES = '<sheetData><row r="1"><c r="A1"><v>5</v></c><c r="C1"><v>5</v></c></row></sheetData>';

// ---------------------------------------------------------------------------
// The order of the rules

test('a pasted cell has the rules of the cell it is a copy of in the order they are applied in there, also when the sheet has one of them already', async () => {
  const doc = await load(
    await buildXlsx({
      sheets: [
        { name: 'Data', xml: worksheetXml(FIVES + cf('A1', R(1)) + cf('A1', S(2))) },
        { name: 'Other', xml: worksheetXml('<sheetData/>' + cf('H1:H2', R(1))) },
      ],
      styles: STYLES,
    }),
  );
  const source = applied(doc.wb, 'A1');
  const style = styleAt(doc.wb, 'A1');
  const beside = applied(doc.wb, 'H1', 1);
  assert.equal(source.length, 2);
  const copied = copyRange(doc.wb, 0, range('A1'));
  apply(doc, pasteAt(doc.wb, 1, range('E1'), copied.text, copied).ops);
  assert.deepEqual(applied(doc.wb, 'E1', 1), source, 'the rules of the pasted cell, in order');
  assert.equal(styleAt(doc.wb, 'E1', 1), style, 'what the pasted cell looks like');
  assert.deepEqual(applied(doc.wb, 'H1', 1), beside, 'a cell that was not pasted over');
  const back = await reopen(await save(doc));
  assert.deepEqual(applied(back, 'E1', 1), source, 'read back');
  assert.equal(styleAt(back, 'E1', 1), style, 'read back, what it looks like');
});

test('the same within one sheet, where the same rule stands twice in the order', async () => {
  const doc = await load(await buildXlsx({ sheets: [{ name: 'Data', xml: worksheetXml(FIVES + cf('A1:A3', R(1)) + cf('C1', S(2)) + cf('C1', R(3))) }], styles: STYLES }));
  const source = applied(doc.wb, 'C1');
  const style = styleAt(doc.wb, 'C1');
  const others = ['A1', 'A2', 'C1'].map((ref) => applied(doc.wb, ref));
  const copied = copyRange(doc.wb, 0, range('C1'));
  apply(doc, pasteAt(doc.wb, 0, range('E1'), copied.text, copied).ops);
  assert.deepEqual(applied(doc.wb, 'E1'), source, 'the rules of the pasted cell, in order');
  assert.equal(styleAt(doc.wb, 'E1'), style, 'what the pasted cell looks like');
  assert.deepEqual(['A1', 'A2', 'C1'].map((ref) => applied(doc.wb, ref)), others, 'the cells that were not pasted over');
  const back = await reopen(await save(doc));
  assert.deepEqual(applied(back, 'E1'), source, 'read back');
});

test('rules the sheet has in the same order get the cells; no rule is made a second time', async () => {
  const doc = await load(await buildXlsx({ sheets: [{ name: 'Data', xml: worksheetXml(FIVES + cf('A1:A3', R(1)) + cf('A1:A3', S(2))) }], styles: STYLES }));
  const copied = copyRange(doc.wb, 0, range('A1'));
  apply(doc, pasteAt(doc.wb, 0, range('E1'), copied.text, copied).ops);
  assert.deepEqual(conditionalNow(doc.wb), ['A1:A3 E1|cellIs|1|0|greaterThan', 'A1:A3 E1|cellIs|2|1|greaterThan']);
  assert.deepEqual(applied(doc.wb, 'E1'), applied(doc.wb, 'A1'));
});

// ---------------------------------------------------------------------------
// Rules that cannot be copied or changed

const BAD = "'Lists!$A$1+1";
const THREES = '<sheetData><row r="1"><c r="A1"><v>3</v></c><c r="E1"><v>3</v></c></row></sheetData>';
const unreadBook = (cells: string, conditional: boolean, validation: boolean) =>
  buildXlsx({
    sheets: [
      {
        name: 'Data',
        xml: worksheetXml(
          THREES +
            (conditional ? cf(cells, `<cfRule type="expression" dxfId="0" priority="1"><formula>${BAD}</formula></cfRule>`) : '') +
            (validation ? `<dataValidations count="1"><dataValidation type="custom" sqref="${cells}"><formula1>${BAD}</formula1></dataValidation></dataValidations>` : ''),
        ),
      },
      { name: 'Lists', xml: worksheetXml('<sheetData/>') },
    ],
    styles: STYLES,
  });

test('a rule whose formula cannot be read stays where it is when its cell is copied and pasted in place: nothing is deleted', async () => {
  const doc = await load(await unreadBook('E1', true, true));
  const before = [validations(doc.wb), conditionalNow(doc.wb)];
  assert.equal(before[0].length + before[1].length, 2);
  const copied = copyRange(doc.wb, 0, range('E1'));
  const result = pasteAt(doc.wb, 0, range('E1'), copied.text, copied);
  assert.deepEqual(ruleOps(result.ops), []);
  assert.equal(result.skippedRules, undefined, 'nothing is missing where the cell was pasted');
  apply(doc, result.ops);
  assert.deepEqual([validations(doc.wb), conditionalNow(doc.wb)], before);
  assert.equal(rulesOf(await partText(await save(doc), SHEET1)), rulesOf(await partText(doc.bytes, SHEET1)));
});

test('pasted over, such a rule stays as it is and the paste is not refused; copied elsewhere it does not come, and both are counted', async () => {
  for (const kind of ['validation', 'conditional']) {
    const doc = await load(await unreadBook('E1:E2', kind === 'conditional', kind === 'validation'));
    const before = [validations(doc.wb), conditionalNow(doc.wb)];
    // A cell without rules over one of its two cells.
    const bare = copyRange(doc.wb, 0, range('A1'));
    const over = pasteAt(doc.wb, 0, range('E1'), bare.text, bare);
    assert.deepEqual(ruleOps(over.ops), [], kind);
    assert.equal(over.skippedRules, 1, `${kind}: the rule that stayed is counted`);
    apply(doc, over.ops);
    assert.deepEqual([validations(doc.wb), conditionalNow(doc.wb)], before, kind);
    // Its own cells in place.
    const own = copyRange(doc.wb, 0, range('E1:E2'));
    const place = pasteAt(doc.wb, 0, range('E1'), own.text, own);
    assert.deepEqual(ruleOps(place.ops), [], kind);
    assert.equal(place.skippedRules, undefined, kind);
    // Elsewhere: the cells are pasted, the rule is not.
    const away = pasteAt(doc.wb, 0, range('G5'), own.text, own);
    assert.deepEqual(ruleOps(away.ops), [], kind);
    assert.equal(away.skippedRules, 1, kind);
    // A cut leaves it.
    apply(doc, cutRange(doc.wb, 0, range('E1')));
    assert.deepEqual([validations(doc.wb), conditionalNow(doc.wb)], before, kind);
  }
});

test('a block repeated so often that a rule would have too many ranges is pasted without that rule, and says so', async () => {
  const doc = await load(
    await buildXlsx({
      sheets: [{ name: 'Data', xml: worksheetXml('<sheetData><row r="1"><c r="A1"><v>3</v></c></row><row r="2"><c r="B2"><v>4</v></c></row></sheetData>' + cf('A1', R(1)) + '<dataValidations count="1"><dataValidation type="whole" sqref="A1"><formula1>1</formula1><formula2>9</formula2></dataValidation></dataValidations>') }, { name: 'Other', xml: worksheetXml('<sheetData/>') }],
      styles: STYLES,
    }),
  );
  // A1 has the rules, the other three cells of the block have none: 100 by 100 repeats are 10,000 ranges.
  const copied = copyRange(doc.wb, 0, range('A1:B2'));
  const result = pasteAt(doc.wb, 1, range('A1:GR200'), copied.text, copied);
  assert.equal(result.skippedRules, 2);
  assert.deepEqual(ruleOps(result.ops), []);
  apply(doc, result.ops);
  assert.equal(doc.wb.sheets[1].rows.get(198)!.cells[198].value, 3);
});

// ---------------------------------------------------------------------------
// What is counted

test('skippedRules counts what the paste could not do as it is where the cells come from: a rule that did not come, a rule that stayed', async () => {
  const doc = await load(await buildXlsx({ sheets: [{ name: 'Data', xml: worksheetXml(BOOK_CELLS + BLOCK_A + BLOCK_B + BLOCK_C + BLOCK_D + extLst(EXT_BAR + EXT_ONLY)) }], styles: STYLES }));
  // In place over a rule only the extension list has: nothing is missing.
  let copied = copyRange(doc.wb, 0, range('E1:E2'));
  let result = pasteAt(doc.wb, 0, range('E1'), copied.text, copied);
  assert.deepEqual([ruleOps(result.ops), result.skippedRules], [[], undefined]);
  // A cell without rules over it: the rule stays though the cell it is a copy of has none.
  copied = copyRange(doc.wb, 0, range('H1'));
  assert.equal(copied.rules, undefined);
  result = pasteAt(doc.wb, 0, range('E1'), copied.text, copied);
  assert.equal(result.skippedRules, 1);
  // The same with a data validation of the extension list.
  const X14 =
    '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dataValidations xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" count="1">' +
    '<x14:dataValidation type="list" allowBlank="1"><x14:formula1><xm:f>Lists!$B$1:$B$3</xm:f></x14:formula1><xm:sqref>L1:L3</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>';
  const other = await load(await buildXlsx({ sheets: [{ name: 'Data', xml: worksheetXml('<sheetData><row r="1"><c r="A1"><v>3</v></c></row></sheetData>' + X14) }, { name: 'Lists', xml: worksheetXml('<sheetData/>') }] }));
  copied = copyRange(other.wb, 0, range('A1'));
  assert.equal(pasteAt(other.wb, 0, range('L1'), copied.text, copied).skippedRules, 1);
  copied = copyRange(other.wb, 0, range('L1:L2'));
  result = pasteAt(other.wb, 0, range('L1'), copied.text, copied);
  assert.deepEqual([ruleOps(result.ops), result.skippedRules], [[], undefined]);
  assert.equal(pasteAt(other.wb, 0, range('N1'), copied.text, copied).skippedRules, 1);
});

// ---------------------------------------------------------------------------
// Ranges that are joined again

test('joinedRanges joins rectangles that share a side, and each stands where the first of its parts stood', () => {
  const names = (refs: string) => joinedRanges(refs.split(' ').map(range)).map(rangeName).join(' ');
  assert.equal(names('A1 A4:A10 F1:F3 A2:A3'), 'A1:A10 F1:F3');
  assert.equal(names('B1:B2 D1:D2 C1:C2'), 'B1:D2');
  assert.equal(names('A1:B1 A2 B2 D4'), 'A1:B2 D4');
  assert.equal(names('A1 B2 C3'), 'A1 B2 C3');
  assert.equal(names('A1:A2 B1'), 'A1:A2 B1');
});

test('cut and pasted in place, the rules are the rules of the file again, byte for byte', async () => {
  const doc = await load(
    await buildXlsx({
      sheets: [
        {
          name: 'Data',
          xml: worksheetXml(
            BOOK_CELLS +
              BLOCK_A +
              BLOCK_B +
              '<dataValidations count="2"><dataValidation type="whole"   showErrorMessage="1" imeMode="off" sqref="A1:A10 C1:C3"><formula1>1</formula1><!-- kept --><formula2>B1*2</formula2></dataValidation><dataValidation type="list" sqref="B1:B5"><formula1>$H$1:$H$2</formula1></dataValidation></dataValidations>',
          ),
        },
        { name: 'Lists', xml: worksheetXml('<sheetData/>') },
      ],
      styles: STYLES,
    }),
  );
  const original = rulesOf(await partText(doc.bytes, SHEET1));
  const before = [validations(doc.wb), conditionalNow(doc.wb)];
  const copied = copyRange(doc.wb, 0, range('A2:C3'));
  apply(doc, cutRange(doc.wb, 0, range('A2:C3')));
  assert.notDeepEqual([validations(doc.wb), conditionalNow(doc.wb)], before);
  apply(doc, pasteAt(doc.wb, 0, range('A2'), copied.text, copied).ops);
  assert.deepEqual([validations(doc.wb), conditionalNow(doc.wb)], before, 'the ranges are joined again');
  const saved = await save(doc);
  await checkPackage(saved);
  assert.equal(rulesOf(await partText(saved, SHEET1)), original);
  // Pasted beside the cells of its rule, a cell makes the range longer, not the list.
  const one = copyRange(doc.wb, 0, range('B5'));
  apply(doc, pasteAt(doc.wb, 0, range('B6'), one.text, one).ops);
  assert.deepEqual(validations(doc.wb)[1], 'B1:B6|list|$H$1:$H$2');
  assert.ok(conditionalNow(doc.wb).indexOf('B1:B6|colorScale|3||min;$H$2;max') >= 0, conditionalNow(doc.wb).join('; '));
});

// ---------------------------------------------------------------------------
// Many rules

test('a paste among thousands of rules does not take a time that grows with rules times rules', async () => {
  const count = 4000;
  let here = '';
  let there = '';
  for (let i = 1; i <= count; i++) {
    here += cf(`A${i}`, `<cfRule type="cellIs" dxfId="0" priority="${i}" operator="greaterThan"><formula>${i}</formula></cfRule>`);
    there += cf(`A${i}`, `<cfRule type="cellIs" dxfId="0" priority="${i}" operator="greaterThan"><formula>${100000 + i}</formula></cfRule>`);
  }
  const doc = await load(await buildXlsx({ sheets: [{ name: 'Data', xml: worksheetXml(THREES + here) }, { name: 'Two', xml: worksheetXml('<sheetData/>' + there) }], styles: STYLES }));
  const copied = copyRange(doc.wb, 0, range('A1:A1000'));
  const started = Date.now();
  const result = pasteAt(doc.wb, 1, range('C1'), copied.text, copied);
  const took = Date.now() - started;
  assert.equal(ruleOps(result.ops).length, 1);
  // A guard against getting stuck, not a measurement: some ten times what an idle machine takes.
  assert.ok(took < 3000, `the paste took ${took} ms`);
});
