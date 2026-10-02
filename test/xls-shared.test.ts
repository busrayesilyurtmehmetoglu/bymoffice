// Shared formulas of an xls file: one text for a group, the cells after the
// first pointing to it as in an xlsx file, and the cells whose references
// wrap around the edge of the old sheet standing for themselves.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine } from '../src/calc/engine';
import { readXlsStream } from '../src/formats/xls/workbook';
import { getCell, type Sheet } from '../src/model/sheet/types';
import { editTextAt } from '../src/ops/commands';
import { explicitFormula, sharedMasters } from '../src/ops/ops';
import { cat, formula, number, ptg, shrFmla, tokens, u16, workbookStream } from './xls-fixtures';

const repeat = (part: number[], times: number): number[] => {
  const out: number[] = [];
  for (let i = 0; i < times; i++) for (const byte of part) out.push(byte);
  return out;
};

/** A sheet whose column `col` is one shared formula over the rows [first, first + count), with a number in column A of each row. */
function sharedColumn(count: number, shared: number[], first = 0, col = 2): Uint8Array {
  const parts: Uint8Array[] = [];
  for (let i = 0; i < count; i++) {
    const row = first + i;
    parts.push(number(row, 0, row + 1));
    parts.push(formula(row, col, ptg.exp(first, col), row + 0.5));
    if (i === 0) parts.push(shrFmla(first, first + count - 1, col, col, shared));
  }
  return workbookStream({ sheets: [{ name: 'Sheet1', records: cat(...parts) }] });
}

/** The formula of a cell written out for its own position. */
const textAt = (sheet: Sheet, row: number, col: number): string | undefined => explicitFormula(getCell(sheet, row, col)!, row, col, sharedMasters(sheet)).formula?.text;

test('a legitimate group of long formulas down thousands of rows is read in full', () => {
  // About a thousand characters, shared by 20,000 rows: more text than a reader may make for a file of this size.
  const word = 'y'.repeat(246);
  const shared = [...ptg.str(word), ...repeat([...ptg.str(word), ...ptg.concat], 3), ...ptg.refN(0, 0xfe), ...ptg.concat];
  const count = 20000;
  const { workbook, report } = readXlsStream(sharedColumn(count, shared));
  const sheet = workbook.sheets[0];
  assert.equal(report.untranslatedFormulas, 0);
  let without = 0;
  let chars = 0;
  for (let row = 0; row < count; row++) {
    const cell = getCell(sheet, row, 2);
    if (!cell || !cell.formula) without++;
    else chars += cell.formula.text.length;
  }
  assert.equal(without, 0);
  // One text for all of them.
  assert.equal(chars, 4 * 248 + 4 + 2);
  assert.equal(textAt(sheet, 0, 2), `"${word}"&"${word}"&"${word}"&"${word}"&A1`);
  assert.equal(textAt(sheet, count - 1, 2), `"${word}"&"${word}"&"${word}"&"${word}"&A${count}`);
});

test('the cells of a group are edited and calculated as the formula each stands for', () => {
  // C1:C3 = A1*2+$A$1, A2*2+$A$1, A3*2+$A$1.
  const shared = tokens(ptg.refN(0, 0xfe), ptg.int(2), ptg.mul, ptg.ref(0, 0, false, false), ptg.add);
  const { workbook } = readXlsStream(sharedColumn(3, shared));
  const sheet = workbook.sheets[0];
  assert.deepEqual(getCell(sheet, 0, 2)?.formula, { text: 'A1*2+$A$1', type: 'shared', si: 0, ref: 'C1:C3' });
  assert.deepEqual(getCell(sheet, 2, 2)?.formula, { text: '', type: 'shared', si: 0 });
  assert.equal(editTextAt(workbook, sheet, 0, 2), '=A1*2+$A$1');
  assert.equal(editTextAt(workbook, sheet, 1, 2), '=A2*2+$A$1');
  assert.equal(editTextAt(workbook, sheet, 2, 2), '=A3*2+$A$1');
  // The values in the file are not those of the formulas: calculating gives each cell its own.
  new CalcEngine(workbook).recalculate(undefined, { overwrite: true });
  assert.deepEqual([0, 1, 2].map((row) => getCell(sheet, row, 2)?.value), [3, 5, 7]);
});

test('a reference that wraps around the edge of the old sheet: the cell keeps the text the file means', () => {
  // "The cell above" in column A: for the first row that is the last row of the sheet.
  const shared = tokens(ptg.refN(0xffff, 0xfe), ptg.int(1), ptg.add);
  const { workbook, report } = readXlsStream(sharedColumn(3, shared));
  const sheet = workbook.sheets[0];
  assert.equal(report.untranslatedFormulas, 0);
  assert.deepEqual(getCell(sheet, 0, 2)?.formula, { text: 'A65536+1', type: 'shared', si: 0, ref: 'C1:C3' });
  // Moving the references of the first cell would leave the old sheet; the file wraps around.
  assert.deepEqual(getCell(sheet, 1, 2)?.formula, { text: 'A1+1', type: 'normal' });
  assert.deepEqual(getCell(sheet, 2, 2)?.formula, { text: 'A2+1', type: 'normal' });
  assert.deepEqual([0, 1, 2].map((row) => textAt(sheet, row, 2)), ['A65536+1', 'A1+1', 'A2+1']);

  // Columns wrap the same way: one column to the left of A is IV.
  const left = readXlsStream(
    workbookStream({
      sheets: [{ name: 'Sheet1', records: cat(formula(0, 0, ptg.exp(0, 0), 1), shrFmla(0, 0, 0, 1, tokens(ptg.refN(0, 0xff), ptg.int(1), ptg.add)), formula(0, 1, ptg.exp(0, 0), 2)) }],
    }),
  ).workbook.sheets[0];
  assert.deepEqual([textAt(left, 0, 0), textAt(left, 0, 1)], ['IV1+1', 'A1+1']);
});

test('areas with one end fixed, and areas over all rows, read the same whether a cell points to the first or not', () => {
  // SUM($A$1:A<row>): a running total.
  const running = [0x2d, ...u16(0), ...u16(0), ...u16(0), ...u16(0xc000 | 0xfe), 0x42, 1, ...u16(4)];
  const total = readXlsStream(sharedColumn(4, running, 2)).workbook.sheets[0];
  assert.deepEqual([2, 3, 5].map((row) => textAt(total, row, 2)), ['SUM($A$1:A3)', 'SUM($A$1:A4)', 'SUM($A$1:A6)']);
  assert.equal(getCell(total, 3, 2)?.formula?.text, '');

  // An area from the row of the cell to the last row of the sheet: all rows for the first cell, not for the others.
  const toEnd = [0x2d, ...u16(0), ...u16(0xffff), ...u16(0xc000 | 0xfe), ...u16(0x00), 0x42, 1, ...u16(4)];
  const sheet = readXlsStream(sharedColumn(3, toEnd)).workbook.sheets[0];
  assert.deepEqual([0, 1, 2].map((row) => textAt(sheet, row, 2)), ['SUM(A:$A)', 'SUM(A2:$A$65536)', 'SUM(A3:$A$65536)']);
  assert.equal(getCell(sheet, 1, 2)?.formula?.type, 'normal');
});

test('once reading the groups has cost too much, only cells whose references stay on the sheet point to the first unchecked', () => {
  // A long formula that refers to the cell above, in a group that starts in the first row: every cell but the
  // first has a text of its own, until the budget is used up; the cells after that are left out and counted.
  const word = 'y'.repeat(250);
  const shared = [...ptg.str(word), ...repeat([...ptg.str(word), ...ptg.concat], 29), ...ptg.refN(0xffff, 0xfe), ...ptg.concat];
  const count = 2500;
  const { workbook, report } = readXlsStream(sharedColumn(count, shared));
  const sheet = workbook.sheets[0];
  let withFormula = 0;
  for (let row = 0; row < count; row++) {
    const cell = getCell(sheet, row, 2)!;
    if (!cell.formula) continue;
    withFormula++;
    assert.ok(textAt(sheet, row, 2)!.endsWith(row === 0 ? '&A65536' : '&A' + row), `row ${row}`);
  }
  assert.ok(withFormula > 100 && withFormula < count, String(withFormula));
  assert.equal(report.untranslatedFormulas, count - withFormula);
  assert.equal(report.untranslatedReasons.get('shared formula: the texts of the workbook together are too large'), count - withFormula);

  // The same group further down the sheet wraps nowhere: every cell has its formula, and the right one.
  const lower = readXlsStream(sharedColumn(count, shared, 10));
  assert.equal(lower.report.untranslatedFormulas, 0);
  for (const row of [10, 11, 1000, count + 9]) assert.ok(textAt(lower.workbook.sheets[0], row, 2)!.endsWith('&A' + row), `row ${row}`);
});
