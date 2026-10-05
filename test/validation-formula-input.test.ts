// Data validation of a formula that is typed in: the rule judges what the
// formula gives at the moment it is entered, not the text of the formula.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine } from '../src/calc/engine';
import { LOCALE_TR } from '../src/format/locale';
import { DataValidations, tryOps, type ValidationVerdict } from '../src/format/validation';
import { getCell, open, parseCellRef } from '../src/index';
import type { Workbook } from '../src/index';
import { setCellInput } from '../src/ops/commands';
import { buildXlsx, worksheetXml } from './fixtures';

const num = (ref: string, value: number) => `<c r="${ref}"><v>${value}</v></c>`;
const str = (ref: string, text: string) => `<c r="${ref}" t="inlineStr"><is><t>${text}</t></is></c>`;
const dv = (attrs: string, inner = '') => `<dataValidation ${attrs}>${inner}</dataValidation>`;
const f1 = (text: string) => `<formula1>${text}</formula1>`;
const f2 = (text: string) => `<formula2>${text}</formula2>`;

/**
 * A1:A9 whole numbers 1 to 5; C1 decimal above 2 (warns); D1 text of at most 3 characters;
 * E1 a list written into the rule; F1 a list of cells (H1:H3); G1 a custom formula that
 * reads the cell and a formula that reads it (J1 = G1 * 2); K1 a date in 2024; L1 a whole
 * number whose rule ignores blanks. B2 = 9, B3 = 4, B4 is empty, B5 = "ab", B6 = "abcd".
 */
const workbook = () =>
  buildXlsx({
    sheets: [
      {
        name: 'Rules',
        xml: worksheetXml(
          `<sheetData><row r="1">${str('H1', 'Red')}<c r="J1"><f>G1*2</f><v>0</v></c></row>` +
            `<row r="2">${num('B2', 9)}${str('H2', 'Green')}</row><row r="3">${num('B3', 4)}${num('H3', 7)}</row>` +
            `<row r="5">${str('B5', 'ab')}</row><row r="6">${str('B6', 'abcd')}</row></sheetData>` +
            '<dataValidations count="8">' +
            dv('type="whole" showErrorMessage="1" error="1 to 5" sqref="A1:A9"', f1('1') + f2('5')) +
            dv('type="decimal" operator="greaterThan" errorStyle="warning" showErrorMessage="1" sqref="C1"', f1('2')) +
            dv('type="textLength" operator="lessThanOrEqual" showErrorMessage="1" sqref="D1"', f1('3')) +
            dv('type="list" showErrorMessage="1" sqref="E1"', f1('"Yes,No,4"')) +
            dv('type="list" showErrorMessage="1" sqref="F1"', f1('$H$1:$H$3')) +
            dv('type="custom" showErrorMessage="1" sqref="G1"', f1('AND(ISNUMBER(G1),$J$1&lt;=20)')) +
            dv('type="date" showErrorMessage="1" sqref="K1"', f1('45292') + f2('45657')) +
            dv('type="whole" allowBlank="1" showErrorMessage="1" sqref="L1"', f1('1') + f2('5')) +
            '</dataValidations>',
        ),
      },
    ],
  });

interface Loaded {
  wb: Workbook;
  engine: CalcEngine | undefined;
  type(ref: string, text: string, locale?: typeof LOCALE_TR): ValidationVerdict;
}

async function load(withEngine = true): Promise<Loaded> {
  const wb = await (await open(await workbook())).workbook();
  const engine = withEngine ? new CalcEngine(wb) : undefined;
  const rules = new DataValidations(wb, wb.sheets[0], { engine: () => engine });
  const type = (ref: string, text: string, locale?: typeof LOCALE_TR) => {
    const { row, col } = parseCellRef(ref);
    // What the view does: the operation of the input, and the cell it writes.
    const ops = setCellInput(wb, 0, row, col, text, { locale });
    let cell = getCell(wb.sheets[0], row, col);
    for (const op of ops) if (op.type === 'setCells') for (const change of op.cells) if (change.row === row && change.col === col) cell = change.cell;
    return rules.check(row, col, cell, text, (run) => tryOps(wb, engine, ops, run));
  };
  return { wb, engine, type };
}

test('validation of a typed formula: its result is judged by a number rule', async () => {
  const { type } = await load();
  // B2 is 9: outside 1 to 5.
  assert.deepEqual(type('A1', '=B2'), { ok: false, style: 'stop', message: '1 to 5' });
  assert.equal(type('A1', '=B3').ok, true);
  assert.equal(type('A1', '=B3+1').ok, true);
  assert.equal(type('A1', '=B3+2').ok, false);
  // Not a whole number.
  assert.equal(type('A1', '=B3/8').ok, false);
  // A reference to an empty cell is 0, as the cell would show.
  assert.equal(type('A1', '=B4').ok, false);
  // Text is not a number.
  assert.equal(type('A1', '=B5').ok, false);
  // TRUE is not a number either, as when it is typed.
  assert.equal(type('A1', '=1=1').ok, type('A1', 'TRUE').ok);
});

test('validation of a typed formula: a rule that only warns still only warns', async () => {
  const { type } = await load();
  assert.deepEqual(type('C1', '=B3-3'), { ok: false, style: 'warning' });
  assert.equal(type('C1', '=B3').ok, true);
});

test('validation of a typed formula: text length, lists and dates take the result', async () => {
  const { type } = await load();
  assert.equal(type('D1', '=B5').ok, true);
  assert.equal(type('D1', '=B6').ok, false);
  assert.equal(type('D1', '=B5&"c"').ok, true);
  // The length of a number is that of the number, not of the formula text.
  assert.equal(type('D1', '=100+23').ok, true);
  assert.equal(type('D1', '=1000+234').ok, false);
  // A list written into the rule: the text of the formula is not looked for in it.
  assert.equal(type('E1', '="Yes"').ok, true);
  assert.equal(type('E1', '="Perhaps"').ok, false);
  assert.equal(type('E1', '=B3').ok, true);
  assert.equal(type('E1', '=B2').ok, false);
  // A list of cells.
  assert.equal(type('F1', '=H2').ok, true);
  assert.equal(type('F1', '="red"').ok, true);
  assert.equal(type('F1', '=B3+3').ok, true);
  assert.equal(type('F1', '=B2').ok, false);
  // A date rule compares the serial number the formula gives.
  assert.equal(type('K1', '=DATE(2024,6,1)').ok, true);
  assert.equal(type('K1', '=DATE(2023,6,1)').ok, false);
});

test('validation of a typed formula: a custom rule reads the calculated cell and what depends on it', async () => {
  const { wb, type } = await load();
  const before = JSON.stringify(Array.from(wb.sheets[0].rows.entries()));
  // G1 = 4 makes J1 8.
  assert.equal(type('G1', '=B3').ok, true);
  // G1 = 18 makes J1 36.
  assert.equal(type('G1', '=B2*2').ok, false);
  // Text is not a number.
  assert.equal(type('G1', '=B5').ok, false);
  // Trying a formula out leaves the cells as they were.
  assert.equal(JSON.stringify(Array.from(wb.sheets[0].rows.entries())), before);
  assert.equal(wb.sheets[0].modified, undefined);
});

test('validation of a typed formula: what cannot be calculated is let through', async () => {
  const { type } = await load();
  // A function the engine does not have, a 3D reference, another workbook.
  assert.equal(type('A1', '=NOSUCHFUNCTION(B2)').ok, true);
  assert.equal(type('A1', '=SUM(Rules:Rules!B2)').ok, true);
  assert.equal(type('A1', '=[1]Sheet1!B2').ok, true);
  assert.equal(type('A1', '=NoSuchName').ok, true);
  // UNSURE: a formula whose result is an error value is let through (see DataValidations.check).
  assert.equal(type('A1', '=1/0').ok, true);
  assert.equal(type('A1', '=B5*2').ok, true);
  assert.equal(type('G1', '=1/0').ok, true);
  // An empty text is what an empty cell is: refused unless the rule ignores blanks.
  assert.equal(type('L1', '=""').ok, true);
});

test('validation of a typed formula: without an engine nothing is checked', async () => {
  const { type } = await load(false);
  assert.equal(type('A1', '=B2').ok, true);
  assert.equal(type('G1', '=B2*2').ok, true);
  // A value still is.
  assert.equal(type('A1', '9').ok, false);
});

test('validation of a typed formula: a formula typed the Turkish way is judged as the formula it is stored as', async () => {
  const { type } = await load();
  // TOPLA is SUM, ";" the separator: 4 + 1.
  assert.equal(type('A1', '=TOPLA(B3;1)', LOCALE_TR).ok, true);
  assert.equal(type('A1', '=TOPLA(B2;1)', LOCALE_TR).ok, false);
  // "1,5" is the number 1.5.
  assert.equal(type('A1', '=B3+1,5', LOCALE_TR).ok, false);
  assert.equal(type('C1', '=B3-1,5', LOCALE_TR).ok, true);
});

test('validation of a typed formula: one that reads its own cell is let through', async () => {
  const { type } = await load();
  // The cell is empty: with the value it has, each of these would give 9.
  assert.equal(type('A1', '=A1+9').ok, true);
  assert.equal(type('A1', '=$A$1+9').ok, true);
  assert.equal(type('A2', '=SUM(A1:A3)+9').ok, true);
  assert.equal(type('A2', '=SUM(A:A)+9').ok, true);
  assert.equal(type('A2', '=SUM(2:2)').ok, true);
  assert.equal(type('A1', '=Rules!A1+9').ok, true);
  assert.equal(type('A1', "=IF(B3>0,'Rules'!A1,1)+9").ok, true);
  // Another cell, a range beside the cell: judged as before.
  assert.equal(type('A1', '=A2+9').ok, false);
  assert.equal(type('A1', '=SUM(B1:C9)').ok, false);
  assert.equal(type('A1', '=SUM(2:2)').ok, false);
  assert.equal(type('A1', '=B3').ok, true);
});

test('validation of a typed formula: while a calculation is under way nothing is judged by stale values', async () => {
  // D1 = SUM of 200 numbers + 9 (20109), with 9 as the value of the file; 300 more formulas keep the engine busy.
  let rows = '';
  for (let r = 1; r <= 300; r++) {
    rows += `<row r="${r}">${r === 1 ? '<c r="D1"><f>SUM(B$1:B$200)+9</f><v>9</v></c>' : ''}${r <= 200 ? num(`B${r}`, r) : ''}<c r="C${r}"><f>SUM(B$1:B$200)+ROW()</f><v>0</v></c></row>`;
  }
  const bytes = await buildXlsx({
    sheets: [{ name: 'Busy', xml: worksheetXml(`<sheetData>${rows}</sheetData><dataValidations count="1">${dv('type="whole" showErrorMessage="1" sqref="A1"', f1('1') + f2('5'))}</dataValidations>`) }],
  });
  const wb = await (await open(bytes)).workbook();
  const engine = new CalcEngine(wb, { budget: { work: 5000, milliseconds: 0 } });
  const rules = new DataValidations(wb, wb.sheets[0], { engine: () => engine });
  const ops = setCellInput(wb, 0, 0, 0, '=D1');
  const cell = ops[0].type === 'setCells' ? ops[0].cells[0].cell : undefined;
  // Everything is asked to be calculated, and the budget runs out: D1 may still hold the 9 of the file.
  engine.recalculate(undefined, { overwrite: true });
  assert.equal(engine.incomplete, true);
  assert.equal(rules.check(0, 0, cell, '=D1').ok, true);
  for (let i = 0; engine.incomplete && i < 1000; i++) engine.recalculate([]);
  assert.equal(engine.incomplete, false);
  assert.equal(getCell(wb.sheets[0], 0, 3)!.value, 20109);
  // Calculated: D1 is far outside 1 to 5.
  assert.equal(rules.check(0, 0, cell, '=D1').ok, false);
});
