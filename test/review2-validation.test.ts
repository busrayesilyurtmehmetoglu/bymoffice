// Second review of data validation: sheet names are compared the way the engine compares
// them, and a rule whose limit or source gives an error value judges nothing.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine } from '../src/calc/engine';
import { DataValidations, tryOps, type ValidationVerdict } from '../src/format/validation';
import { getCell, open, parseCellRef } from '../src/index';
import type { Workbook } from '../src/index';
import { setCellInput } from '../src/ops/commands';
import { History } from '../src/ops/ops';
import { deleteSheet } from '../src/ops/sheets';
import { buildXlsx, worksheetXml } from './fixtures';

const num = (ref: string, value: number) => `<c r="${ref}"><v>${value}</v></c>`;
const dv = (attrs: string, inner = '') => `<dataValidation ${attrs}>${inner}</dataValidation>`;
const f1 = (text: string) => `<formula1>${text}</formula1>`;
const f2 = (text: string) => `<formula2>${text}</formula2>`;

interface Loaded {
  wb: Workbook;
  rules: DataValidations;
  type(ref: string, text: string): ValidationVerdict;
}

async function load(bytes: Uint8Array): Promise<Loaded> {
  const wb = await (await open(bytes)).workbook();
  const engine = new CalcEngine(wb);
  const rules = new DataValidations(wb, wb.sheets[0], { engine: () => engine });
  const type = (ref: string, text: string) => {
    const { row, col } = parseCellRef(ref)!;
    const ops = setCellInput(wb, 0, row, col, text);
    let cell = getCell(wb.sheets[0], row, col);
    for (const op of ops) if (op.type === 'setCells') for (const change of op.cells) if (change.row === row && change.col === col) cell = change.cell;
    return rules.check(row, col, cell, text, (run) => tryOps(wb, engine, ops, run));
  };
  return { wb, rules, type };
}

test('review 2: a rule and a typed formula name a sheet in whatever case, dotless and dotted letters included', async () => {
  const bytes = await buildXlsx({
    sheets: [
      {
        name: 'Kasım',
        xml: worksheetXml(
          `<sheetData><row r="1">${num('H1', 7)}</row><row r="2">${num('H2', 8)}</row></sheetData><dataValidations count="3">` +
            dv('type="whole" showErrorMessage="1" sqref="A1"', f1('1') + f2('5')) +
            dv('type="list" showErrorMessage="1" sqref="B1"', f1('KASIM!$H$1:$H$2')) +
            dv('type="list" showErrorMessage="1" sqref="C1"', f1('Yerel')) +
            '</dataValidations>',
        ),
      },
    ],
    workbookTail: '<definedNames><definedName name="Yerel" localSheetId="0">kasim!$H$1:$H$2</definedName></definedNames>',
  });
  const { type, rules } = await load(bytes);
  // The cell is empty: with the value it has, the formula would give 0, which the rule refuses.
  assert.equal(type('A1', '=A1+0').ok, true);
  assert.equal(type('A1', '=KASIM!A1+0').ok, true, 'a reference to the cell itself, under the name in capitals');
  assert.equal(type('A1', '=kasim!A1+0').ok, true);
  assert.equal(type('A1', '=KASIM!A2+0').ok, false);
  // The source of a list on the sheet, named in capitals: read as cells, with what they show.
  assert.deepEqual(rules.list(0, 1)?.items, ['7', '8']);
  assert.equal(rules.hasList(0, 1), true);
  assert.equal(type('B1', '8').ok, true);
  assert.equal(type('B1', '9').ok, false);
  assert.deepEqual(rules.list(0, 2)?.items, ['7', '8']);
});

test('review 2: a rule whose limit or source gives an error value judges nothing; a custom formula still refuses by its own errors', async () => {
  const bytes = await buildXlsx({
    sheets: [
      {
        name: 'Rules',
        xml: worksheetXml(
          `<sheetData><row r="1">${num('H1', 1)}</row><row r="2">${num('H2', 2)}<c r="G2" t="e"><f>1/0</f><v>#DIV/0!</v></c></row></sheetData><dataValidations count="7">` +
            dv('type="whole" showErrorMessage="1" sqref="A1"', f1('Gone!$A$1') + f2('5')) +
            dv('type="list" showErrorMessage="1" sqref="B1"', f1('Gone!$A$1:$A$3')) +
            dv('type="custom" showErrorMessage="1" sqref="C1"', f1('Gone!$A$1&gt;0')) +
            dv('type="custom" showErrorMessage="1" sqref="D1"', f1('MATCH(D1,$H$1:$H$2,0)')) +
            dv('type="decimal" operator="lessThan" showErrorMessage="1" sqref="E1"', f1('$G$2')) +
            dv('type="textLength" showErrorMessage="1" sqref="F1"', f1('1') + f2('1/0')) +
            dv('type="list" showErrorMessage="1" sqref="A3"', f1('INDIRECT("nowhere")')) +
            '</dataValidations>',
        ),
      },
      { name: 'Gone', xml: worksheetXml(`<sheetData><row r="1">${num('A1', 1)}</row><row r="2">${num('A2', 2)}</row></sheetData>`) },
    ],
  });
  const { wb, type, rules } = await load(bytes);
  assert.equal(type('A1', '9').ok, false);
  assert.equal(type('B1', '9').ok, false);
  assert.equal(type('C1', '9').ok, true);
  assert.deepEqual(rules.list(0, 1)?.items, ['1', '2']);
  // Limits that give an error whatever is typed.
  assert.equal(type('E1', '9').ok, true);
  assert.equal(type('F1', 'long text').ok, true);
  assert.deepEqual(rules.list(2, 0), { items: [], truncated: false });
  assert.equal(type('A3', 'anything').ok, true);

  new History(wb).apply(deleteSheet(wb, 1), undefined, undefined);
  // The limit, the source and the formula are #REF! now: the cells take what is typed.
  for (const ref of ['A1', 'B1', 'C1']) {
    assert.equal(type(ref, '9').ok, true, ref);
    assert.equal(type(ref, 'text').ok, ref !== 'A1', ref);
  }
  // What can be said without the limit still is: text is no whole number.
  assert.deepEqual(rules.list(0, 1), { items: [], truncated: false });
  assert.equal(rules.hasList(0, 1), false);
  // A formula whose error says "not among these" refuses, as the application does.
  assert.equal(type('D1', '2').ok, true);
  assert.equal(type('D1', '9').ok, false);
});
