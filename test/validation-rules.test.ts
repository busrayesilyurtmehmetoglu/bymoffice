// Data validation: what the reader makes of <dataValidations>, which values
// the rules let through, the lists they offer, and how they follow moved rows.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine } from '../src/calc/engine';
import { LOCALE_TR } from '../src/format/locale';
import { DataValidations, MAX_LIST_ITEMS, ValidationError } from '../src/format/validation';
import { open, rangeName } from '../src/index';
import type { Cell, Workbook } from '../src/index';
import { cellFromInput } from '../src/ops/commands';
import { applyOp } from '../src/ops/ops';
import { buildXlsx, worksheetXml } from './fixtures';

const num = (ref: string, value: number, style = 0) => `<c r="${ref}" s="${style}"><v>${value}</v></c>`;
const str = (ref: string, text: string) => `<c r="${ref}" t="inlineStr"><is><t>${text}</t></is></c>`;

const dv = (attrs: string, inner = '') => `<dataValidation ${attrs}>${inner}</dataValidation>`;
const f1 = (text: string) => `<formula1>${text}</formula1>`;
const f2 = (text: string) => `<formula2>${text}</formula2>`;

const RULES =
  '<dataValidations count="14">' +
  dv('type="whole" allowBlank="1" showErrorMessage="1" errorTitle="Range" error="1 to 10 only" sqref="A1:A10"', f1('1') + f2('10')) +
  dv('type="decimal" operator="greaterThan" errorStyle="warning" showErrorMessage="1" error="Too small" sqref="B1:B10"', f1('$H$1')) +
  dv('type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" promptTitle="Pick" prompt="Yes or no_x000a_or maybe" sqref="C1:C10 C20"', f1('"Yes,No, Maybe ,1.5,TRUE"')) +
  dv('type="list" errorStyle="information" showErrorMessage="1" sqref="D1:D10"', f1('$G$1:$G$5')) +
  dv('type="list" showErrorMessage="1" sqref="E1:E10"', f1('Colours')) +
  dv('type="custom" showErrorMessage="1" sqref="F1:F10"', f1('ISNUMBER(F1)')) +
  dv('type="textLength" operator="lessThanOrEqual" showErrorMessage="1" allowBlank="1" sqref="I1:I10"', f1('3')) +
  dv('type="date" showErrorMessage="1" sqref="J1:J10"', f1('45292') + f2('45657')) +
  dv('type="list" showDropDown="1" showErrorMessage="1" sqref="K1:K5"', f1('"a,b"')) +
  dv('type="list" showErrorMessage="1" sqref="L1"', f1('OFFSET($G$1,0,0,2,1)')) +
  dv('type="whole" operator="between" sqref="O1"', f1('1') + f2('2')) +
  dv('showInputMessage="1" prompt="Any value" sqref="P1"') +
  dv('type="time" operator="notBetween" showErrorMessage="1" sqref="Q1"', f1('0.25') + f2('0.75')) +
  dv('type="list" showErrorMessage="1" sqref="R1:R3"', f1('G1:G2')) +
  '</dataValidations>';

const EXTENSION =
  '<extLst><ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}">' +
  '<x14:dataValidations xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" count="1">' +
  '<x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1"><x14:formula1><xm:f>Lists!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>N1:N5</xm:sqref></x14:dataValidation>' +
  '</x14:dataValidations></ext></extLst>';

const STYLES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
  '</styleSheet>';

const rulesWorkbook = () =>
  buildXlsx({
    styles: STYLES,
    workbookTail: '<definedNames><definedName name="Colours">Lists!$A$1:$A$3</definedName></definedNames>',
    sheets: [
      {
        name: 'Rules',
        xml: worksheetXml(
          `<sheetData><row r="1">${str('G1', 'Red')}${num('H1', 2.5)}</row><row r="2">${str('G2', 'green')}</row><row r="3">${num('G3', 3, 1)}</row><row r="5"><c r="G5" t="b"><v>1</v></c></row></sheetData>` +
            RULES +
            EXTENSION,
        ),
      },
      { name: 'Lists', xml: worksheetXml(`<sheetData><row r="1">${str('A1', 'cyan')}</row><row r="2">${str('A2', 'magenta')}</row><row r="3">${str('A3', 'yellow')}</row></sheetData>`) },
    ],
  });

interface Loaded {
  wb: Workbook;
  rules: DataValidations;
  /** The verdict for text typed into a cell of the first sheet. */
  type(ref: string, text: string): { ok: boolean; style?: string; title?: string; message?: string };
}

async function load(engine = true): Promise<Loaded> {
  const wb = await (await open(await rulesWorkbook())).workbook();
  const calc = engine ? new CalcEngine(wb) : undefined;
  const rules = new DataValidations(wb, wb.sheets[0], { engine: () => calc });
  const type = (ref: string, text: string) => {
    const match = /^([A-Z]+)(\d+)$/.exec(ref)!;
    const col = match[1].split('').reduce((value, letter) => value * 26 + letter.charCodeAt(0) - 64, 0) - 1;
    const row = Number(match[2]) - 1;
    const cell: Cell | undefined = cellFromInput(wb, wb.sheets[0], row, col, text);
    return rules.check(row, col, cell, text, (run) => {
      const inverse = applyOp(wb, { type: 'setCells', sheet: 0, cells: [{ row, col, cell }] });
      try {
        return run();
      } finally {
        applyOp(wb, inverse);
      }
    });
  };
  return { wb, rules, type };
}

test('validation: the reader takes type, operator, formulas, flags and messages; showDropDown means the arrow is hidden', async () => {
  const { wb } = await load();
  const rules = wb.sheets[0].dataValidations!;
  assert.equal(rules.length, 15);
  assert.deepEqual(rules[0], {
    ranges: [{ top: 0, left: 0, bottom: 9, right: 0 }],
    type: 'whole',
    operator: 'between',
    allowBlank: true,
    hideDropDown: false,
    showInputMessage: false,
    showErrorMessage: true,
    errorStyle: 'stop',
    errorTitle: 'Range',
    error: '1 to 10 only',
    formula1: '1',
    formula2: '10',
    // Where the element is in the part.
    src: 0,
  });
  assert.equal(rules[1].errorStyle, 'warning');
  assert.equal(rules[1].operator, 'greaterThan');
  assert.equal(rules[1].formula1, '$H$1');
  assert.deepEqual(rules[2].ranges.map(rangeName), ['C1:C10', 'C20']);
  assert.equal(rules[2].promptTitle, 'Pick');
  // Line breaks in messages are written as _x000a_.
  assert.equal(rules[2].prompt, 'Yes or no\nor maybe');
  assert.equal(rules[3].errorStyle, 'information');
  assert.equal(rules[8].hideDropDown, true);
  assert.equal(rules[11].type, 'none');
  // The rule of the 2010 extension, whose list is on another sheet.
  assert.deepEqual(rules[14], {
    ranges: [{ top: 0, left: 13, bottom: 4, right: 13 }],
    type: 'list',
    operator: 'between',
    allowBlank: true,
    hideDropDown: false,
    showInputMessage: true,
    showErrorMessage: true,
    errorStyle: 'stop',
    formula1: 'Lists!$A$1:$A$3',
    extended: true,
  });
  assert.equal(wb.sheets[1].dataValidations, undefined);
});

test('validation: whole numbers, decimals, dates, times and text lengths', async () => {
  const { type } = await load();
  const failed = { ok: false, style: 'stop', title: 'Range', message: '1 to 10 only' };
  assert.deepEqual(type('A1', '5'), { ok: true });
  assert.deepEqual(type('A1', '1'), { ok: true });
  assert.deepEqual(type('A1', '10'), { ok: true });
  assert.deepEqual(type('A1', '11'), failed);
  assert.deepEqual(type('A1', '0'), failed);
  assert.deepEqual(type('A1', '5.5'), failed);
  assert.deepEqual(type('A1', 'five'), failed);
  assert.deepEqual(type('A1', 'TRUE'), failed);
  assert.deepEqual(type('A1', ''), { ok: true });
  // Outside the range of the rule nothing is asked.
  assert.deepEqual(type('A11', '999'), { ok: true });

  // The limit is a cell: H1 holds 2.5. The rule only warns.
  assert.deepEqual(type('B2', '2.6'), { ok: true });
  assert.deepEqual(type('B2', '2.5'), { ok: false, style: 'warning', message: 'Too small' });
  assert.deepEqual(type('B2', 'x'), { ok: false, style: 'warning', message: 'Too small' });
  // Blanks are not allowed there.
  assert.equal(type('B2', '').ok, false);

  assert.equal(type('I1', 'abc').ok, true);
  assert.equal(type('I1', 'abcd').ok, false);
  assert.equal(type('I1', '123').ok, true);
  assert.equal(type('I1', '1234').ok, false);

  assert.equal(type('J1', '1/1/2024').ok, true);
  assert.equal(type('J1', '12/31/2024').ok, true);
  assert.equal(type('J1', '1/1/2025').ok, false);
  assert.equal(type('J1', 'soon').ok, false);

  // Not between 06:00 and 18:00.
  assert.equal(type('Q1', '5:00').ok, true);
  assert.equal(type('Q1', '12:00').ok, false);
  assert.equal(type('Q1', '19:30').ok, true);
});

test('validation: a formula is checked by its result; a cell whose rule shows no error is not checked', async () => {
  const { type } = await load();
  assert.equal(type('A1', '=99+1').ok, false);
  assert.equal(type('A1', '=9+1').ok, true);
  assert.equal(type('O1', '77').ok, true);
  assert.equal(type('P1', 'anything').ok, true);
});

test('validation: lists - written into the rule, a range, a defined name, another sheet, a formula', async () => {
  const { rules, type } = await load();
  // UNSURE (see DataValidations): the items of a list written into the rule are compared by case.
  assert.equal(type('C1', 'Yes').ok, true);
  assert.equal(type('C1', 'yes').ok, false);
  assert.equal(type('C1', 'Maybe').ok, true);
  assert.equal(type('C1', '1.5').ok, true);
  assert.equal(type('C1', 'true').ok, true);
  assert.equal(type('C1', 'Perhaps').ok, false);
  assert.equal(type('C20', 'No').ok, true);
  assert.deepEqual(rules.list(10, 2), undefined);
  assert.deepEqual(rules.list(0, 2), { items: ['Yes', 'No', 'Maybe', '1.5', 'TRUE'], truncated: false });

  // A range: without regard to case; numbers by value; the list shows what the cells show.
  assert.equal(type('D1', 'RED').ok, true);
  assert.equal(type('D1', 'Green').ok, true);
  assert.equal(type('D1', '3').ok, true);
  assert.equal(type('D1', 'TRUE').ok, true);
  assert.deepEqual(type('D1', 'blue'), { ok: false, style: 'information' });
  assert.deepEqual(rules.list(0, 3), { items: ['Red', 'green', '3.00', 'TRUE'], truncated: false });

  assert.equal(type('E1', 'Magenta').ok, true);
  assert.equal(type('E1', 'Red').ok, false);
  assert.deepEqual(rules.list(0, 4), { items: ['cyan', 'magenta', 'yellow'], truncated: false });
  assert.deepEqual(rules.list(0, 13), { items: ['cyan', 'magenta', 'yellow'], truncated: false });
  assert.equal(type('N5', 'cyan').ok, true);
  assert.equal(type('N5', 'black').ok, false);

  // A formula: the engine gives its values.
  assert.deepEqual(rules.list(0, 11), { items: ['Red', 'green'], truncated: false });
  assert.equal(type('L1', 'green').ok, true);
  assert.equal(type('L1', '3').ok, false);

  // A relative source moves with the cell: R1 offers G1:G2, R3 offers G3:G4.
  assert.deepEqual(rules.list(0, 17), { items: ['Red', 'green'], truncated: false });
  assert.deepEqual(rules.list(2, 17), { items: ['3.00'], truncated: false });
  assert.equal(type('R3', '3').ok, true);
  assert.equal(type('R3', 'Red').ok, false);
});

test('validation: which cells show the button of a list, and their messages', async () => {
  const { rules } = await load();
  assert.equal(rules.hasList(0, 2), true);
  assert.equal(rules.hasList(19, 2), true);
  assert.equal(rules.hasList(10, 2), false);
  assert.equal(rules.hasList(0, 3), true);
  assert.equal(rules.hasList(0, 4), true);
  assert.equal(rules.hasList(0, 11), true);
  assert.equal(rules.hasList(0, 13), true);
  // showDropDown="1": no button, and the list still applies.
  assert.equal(rules.hasList(0, 10), false);
  assert.deepEqual(rules.list(0, 10), { items: ['a', 'b'], truncated: false });
  assert.equal(rules.hasList(0, 0), false);
  assert.deepEqual(rules.prompt(0, 2), { title: 'Pick', text: 'Yes or no\nor maybe' });
  assert.deepEqual(rules.prompt(0, 15), { title: '', text: 'Any value' });
  assert.equal(rules.prompt(0, 0), undefined);
  assert.equal(rules.prompt(0, 13), undefined);
  assert.equal(rules.ruleAt(0, 0)!.type, 'whole');
  assert.equal(rules.ruleAt(30, 30), undefined);
  assert.equal(rules.active, true);
});

test('validation: a custom formula reads the value that is being typed', async () => {
  const { wb, type } = await load();
  const before = JSON.stringify(wb.sheets[0].rows.get(2));
  assert.equal(type('F3', '12').ok, true);
  assert.deepEqual(type('F3', 'twelve'), { ok: false, style: 'stop' });
  assert.equal(type('F3', '').ok, false);
  // Trying a value out leaves the cells as they were.
  assert.equal(JSON.stringify(wb.sheets[0].rows.get(2)), before);
});

test('validation: without an engine, rules that need formulas are not applied; constants still are', async () => {
  const { rules, type } = await load(false);
  assert.equal(type('A1', '11').ok, false);
  assert.equal(type('B2', '-5').ok, true);
  assert.equal(type('F3', 'twelve').ok, true);
  assert.equal(type('L1', 'whatever').ok, true);
  assert.equal(rules.hasList(0, 11), false);
  assert.equal(rules.list(0, 11), undefined);
  // Ranges and names are read from the model, with or without an engine.
  assert.equal(type('D1', 'blue').ok, false);
  assert.equal(type('E1', 'cyan').ok, true);
  assert.deepEqual(rules.list(0, 4), { items: ['cyan', 'magenta', 'yellow'], truncated: false });
});

test('validation: lists are shown in the locale of the view', async () => {
  const { wb } = await load();
  const rules = new DataValidations(wb, wb.sheets[0], { locale: LOCALE_TR });
  assert.deepEqual(rules.list(0, 3), { items: ['Red', 'green', '3,00', 'DOĞRU'], truncated: false });
});

test('validation: the rules follow inserted and deleted rows and columns', async () => {
  const { wb, rules, type } = await load();
  applyOp(wb, { type: 'insert', sheet: 0, axis: 'row', at: 0, count: 2 });
  applyOp(wb, { type: 'insert', sheet: 0, axis: 'col', at: 0, count: 1 });
  // A1:A10 is B3:B12 now.
  assert.deepEqual(rules.ruleAt(2, 1)!.ranges.map(rangeName), ['B3:B12']);
  assert.equal(rules.ruleAt(1, 1), undefined);
  assert.equal(rules.ruleAt(2, 0), undefined);
  assert.equal(type('B3', '11').ok, false);
  // The limit of the decimal rule followed its cell (H1 -> I3), the list its range (G1:G5 -> H3:H7).
  assert.equal(rules.ruleAt(2, 2)!.formula1, '$I$3');
  assert.equal(type('C3', '2.6').ok, true);
  assert.equal(type('C3', '2.4').ok, false);
  assert.equal(rules.ruleAt(2, 4)!.formula1, '$H$3:$H$7');
  assert.deepEqual(rules.list(2, 4), { items: ['Red', 'green', '3.00', 'TRUE'], truncated: false });
  // The custom formula still speaks of the cell it is applied to.
  assert.equal(rules.ruleAt(2, 6)!.formula1, 'ISNUMBER(G3)');
  assert.equal(type('G5', 'x').ok, false);
  assert.equal(type('G5', '7').ok, true);
  // Rows deleted: a rule whose cells are all gone is gone.
  applyOp(wb, { type: 'delete', sheet: 0, axis: 'row', at: 0, count: 30 });
  assert.equal(rules.ruleAt(0, 1), undefined);
  assert.equal(rules.ruleAt(2, 1), undefined);
});

test('validation: a list source with more values than the list holds is cut, and every value still passes', async () => {
  let rows = '';
  for (let r = 1; r <= MAX_LIST_ITEMS + 500; r++) rows += `<row r="${r}"><c r="B${r}"><v>${r}</v></c></row>`;
  const bytes = await buildXlsx({
    sheets: [{ name: 'Long', xml: worksheetXml(`<sheetData>${rows}</sheetData><dataValidations count="1">${dv('type="list" showErrorMessage="1" sqref="A1"', f1('$B:$B'))}</dataValidations>`) }],
  });
  const wb = await (await open(bytes)).workbook();
  const rules = new DataValidations(wb, wb.sheets[0]);
  const started = Date.now();
  const list = rules.list(0, 0)!;
  assert.equal(list.items.length, MAX_LIST_ITEMS);
  assert.equal(list.truncated, true);
  const value = (n: number): Cell => ({ kind: 'number', value: n, style: 0 });
  assert.equal(rules.check(0, 0, value(MAX_LIST_ITEMS + 400), String(MAX_LIST_ITEMS + 400)).ok, true);
  assert.equal(rules.check(0, 0, value(MAX_LIST_ITEMS + 501), String(MAX_LIST_ITEMS + 501)).ok, false);
  assert.ok(Date.now() - started < 30000);
});

test('validation: rules that make no sense refuse nothing and throw nothing', async () => {
  const bytes = await buildXlsx({
    sheets: [
      {
        name: 'Odd',
        xml: worksheetXml(
          '<sheetData/><dataValidations>' +
            dv('type="whole" showErrorMessage="1" sqref="A1"') +
            dv('type="whole" operator="sideways" showErrorMessage="1" sqref="A2"', f1('1')) +
            dv('type="list" showErrorMessage="1" sqref="A3"', f1('"unterminated')) +
            dv('type="list" showErrorMessage="1" sqref="A4"', f1('NoSuchName')) +
            dv('type="list" showErrorMessage="1" sqref="A5"', f1('[1]Other!$A$1:$A$3')) +
            dv('type="custom" showErrorMessage="1" sqref="A6"', f1('((((')) +
            dv('type="mystery" showErrorMessage="1" sqref="A7"', f1('1')) +
            dv('type="decimal" showErrorMessage="1" sqref="A8"', f1('"text"') + f2('2')) +
            dv('type="list" showErrorMessage="1" sqref="nonsense"', f1('"a"')) +
            dv('type="list" showErrorMessage="1" sqref="A9"', f1('Sheet9!$A$1')) +
            '</dataValidations>',
        ),
      },
    ],
  });
  const wb = await (await open(bytes)).workbook();
  assert.equal(wb.sheets[0].dataValidations!.length, 9);
  const engine = new CalcEngine(wb);
  const rules = new DataValidations(wb, wb.sheets[0], { engine: () => engine });
  for (let row = 0; row < 9; row++) {
    const verdict = rules.check(row, 0, { kind: 'string', value: 'x', style: 0 }, 'x');
    // Only the whole-number rules can say that text is not a number.
    assert.equal(verdict.ok, row !== 0 && row !== 1 && row !== 7, `row ${row + 1}`);
    assert.equal(rules.hasList(row, 0), false, `row ${row + 1}`);
    assert.equal(rules.list(row, 0), undefined, `row ${row + 1}`);
  }
  assert.equal(rules.check(0, 0, { kind: 'number', value: 5, style: 0 }, '5').ok, true);
  assert.equal(rules.check(1, 0, { kind: 'number', value: 5, style: 0 }, '5').ok, true);
  assert.equal(rules.check(7, 0, { kind: 'number', value: 5, style: 0 }, '5').ok, true);
});

test('validation: the error a refused value is reported with carries the words of the rule', () => {
  const error = new ValidationError('Range: 1 to 10 only');
  assert.ok(error instanceof RangeError);
  assert.equal(error.display, 'Range: 1 to 10 only');
  assert.equal(error.name, 'ValidationError');
});
