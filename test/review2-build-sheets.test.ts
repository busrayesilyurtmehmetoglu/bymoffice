// Second review of buildXlsx(): content kept the way the file has it (validations, conditional
// formats) follows the sheets by the names of the file, as a save does: rows and columns are
// matched with the names the file has, and only then do the sheets get the names they have now.
// The same for a workbook read from xls, whose sheets have no parts.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine } from '../src/calc/engine';
import { DataValidations } from '../src/format/validation';
import { buildXlsx as buildFromModel } from '../src/formats/xlsx/build';
import { open, parseCellRef, sheetNameMap } from '../src/index';
import type { DataValidation, Workbook } from '../src/index';
import { fileSheetName } from '../src/model/sheet/types';
import { setCellInput } from '../src/ops/commands';
import { History, type Op } from '../src/ops/ops';
import { deleteSheet, insertSheet, renameSheet } from '../src/ops/sheets';
import { buildXlsx, worksheetXml } from './fixtures';
import { number, workbookStream, xlsFile } from './xls-fixtures';

const rule = (ref: string, type: string, formula1: string): DataValidation => {
  const { row, col } = parseCellRef(ref)!;
  return { ranges: [{ top: row, left: col, bottom: row, right: col }], type, operator: 'between', formula1, formula2: type === 'list' ? undefined : '99', allowBlank: false, hideDropDown: false, showInputMessage: false, showErrorMessage: true, errorStyle: 'stop' };
};

/** Sheets A, B and C; A has a validation and a formula that read C!$A$1. Read from xlsx. */
async function fromXlsx(): Promise<Workbook> {
  const bytes = await buildXlsx({
    sheets: [
      {
        name: 'A',
        xml: worksheetXml(
          '<sheetData><row r="1"><c r="B1"><f>C!$A$1</f><v>5</v></c></row></sheetData>' +
            '<conditionalFormatting sqref="D1"><cfRule type="expression" priority="1"><formula>C!$A$1&gt;0</formula></cfRule></conditionalFormatting>' +
            '<dataValidations count="2"><dataValidation type="whole" showErrorMessage="1" sqref="A1"><formula1>C!$A$1</formula1><formula2>99</formula2></dataValidation>' +
            '<dataValidation type="list" showErrorMessage="1" sqref="A2"><formula1>B!$A$1:$A$2</formula1></dataValidation></dataValidations>',
        ),
      },
      { name: 'B', xml: worksheetXml('<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>') },
      { name: 'C', xml: worksheetXml('<sheetData><row r="1"><c r="A1"><v>5</v></c></row></sheetData>') },
    ],
  });
  return (await open(bytes)).workbook();
}

/** The same sheets read from xls: the reader does not take validations from that format, so the model is given them as a reader would. */
async function fromXls(): Promise<Workbook> {
  const bytes = xlsFile(workbookStream({ sheets: [{ name: 'A' }, { name: 'B', records: number(0, 0, 1) }, { name: 'C', records: number(0, 0, 5) }] }));
  const wb = await (await open(bytes)).workbook();
  assert.deepEqual(wb.sheets.map((sheet) => sheet.part), ['', '', '']);
  wb.sheets[0].dataValidations = [rule('A1', 'whole', 'C!$A$1'), rule('A2', 'list', 'B!$A$1:$A$2')];
  new History(wb).apply(setCellInput(wb, 0, 0, 1, '=C!$A$1'), undefined, undefined);
  return wb;
}

async function built(wb: Workbook): Promise<{ rules: (string | undefined)[]; formula: string | undefined; names: string[] }> {
  const back = await (await open(await buildFromModel(wb))).workbook();
  const host = back.sheets.find((sheet) => sheet.name === 'A')!;
  return {
    rules: (host.dataValidations || []).map((entry) => entry.formula1),
    formula: host.rows.get(0)?.cells[1]?.formula?.text,
    names: back.sheets.map((sheet) => sheet.name),
  };
}

/** The limit of the rule of A1 as a view shows it. */
function shown(wb: Workbook): string | undefined {
  const engine = new CalcEngine(wb);
  const host = wb.sheets.find((sheet) => sheet.name === 'A')!;
  return new DataValidations(wb, host, { engine: () => engine }).ruleAt(0, 0)?.formula1;
}

const insertRow = (sheet: number): Op[] => [{ type: 'insert', sheet, axis: 'row', at: 0, count: 1 }];

for (const [label, make] of [
  ['xlsx', fromXlsx],
  ['xls', fromXls],
] as [string, () => Promise<Workbook>][]) {
  test(`review 2: buildXlsx after a sheet read from ${label} was renamed and rows moved in it: the rule names the sheet as it is called and the cell where it is`, async () => {
    const wb = await make();
    const history = new History(wb);
    assert.deepEqual(wb.sheets.map(fileSheetName), ['A', 'B', 'C']);
    history.apply(renameSheet(wb, 2, 'Yeni'), undefined, undefined);
    assert.equal(fileSheetName(wb.sheets[2]), 'C', 'the sheet goes on being the C of the file');
    assert.equal(sheetNameMap(wb)!('C'), 'Yeni');
    history.apply(insertRow(2), undefined, undefined);
    assert.equal(shown(wb), 'Yeni!$A$2', 'on screen');
    let out = await built(wb);
    assert.deepEqual(out.names, ['A', 'B', 'Yeni']);
    assert.equal(out.formula, 'Yeni!$A$2');
    assert.deepEqual(out.rules, ['Yeni!$A$2', 'B!$A$1:$A$2']);

    // A sheet that is not in the file takes the name the file had for the renamed one: nothing of the file means it.
    history.apply(insertSheet(wb, 3, { name: 'C' }), undefined, undefined);
    history.apply(insertRow(3), undefined, undefined);
    assert.equal(fileSheetName(wb.sheets[3]), undefined);
    assert.equal(shown(wb), 'Yeni!$A$2');
    out = await built(wb);
    assert.deepEqual(out.rules, ['Yeni!$A$2', 'B!$A$1:$A$2']);

    // Deleted: the rule and the formula are left with an error.
    history.apply(deleteSheet(wb, 2), undefined, undefined);
    assert.deepEqual(wb.removedSheetNames, ['C']);
    assert.equal(shown(wb), '#REF!');
    out = await built(wb);
    assert.equal(out.formula, '#REF!');
    assert.deepEqual(out.rules, ['#REF!', 'B!$A$1:$A$2']);

    // Undone to the start: as read.
    while (history.canUndo) history.undo();
    assert.equal(wb.sheets[2].nameInFile, undefined);
    assert.equal(wb.removedSheetNames, undefined);
    out = await built(wb);
    assert.deepEqual(out.rules, ['C!$A$1', 'B!$A$1:$A$2']);
  });
}

test('review 2: buildXlsx gives a conditional format the names the sheets have now', async () => {
  const wb = await fromXlsx();
  const history = new History(wb);
  history.apply(renameSheet(wb, 2, 'Yeni'), undefined, undefined);
  history.apply(insertRow(2), undefined, undefined);
  const back = await (await open(await buildFromModel(wb))).workbook();
  const formats = JSON.stringify(back.sheets[0].conditionalFormats);
  assert.ok(formats.includes('Yeni!$A$2') && !formats.includes('C!'), formats);
});
