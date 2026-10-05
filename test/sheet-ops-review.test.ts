// Findings of the review of the sheet operations (insert, delete, rename), each with the test
// that failed before it was fixed: what is left of a reference to a deleted sheet, 3D references
// that lose an end, workbooks whose content cannot follow a new name, names that differ only by
// case, and the bound on what the undo stack keeps alive.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { renameSheetsInFormula } from '../src/calc/refs';
import { DataValidations, createSheet, emptyStyleSheet, open, parseCellRef, type Workbook } from '../src/index';
import { currentSheetName, foldSheetName, sheetNameMap } from '../src/model/sheet/types';
import { History, type Op } from '../src/ops/ops';
import { canDeleteSheet, canInsertSheet, canRenameSheet, deleteSheet, insertSheet, renameSheet } from '../src/ops/sheets';
import { buildXlsx, worksheetXml } from './fixtures';

const num = (r: string, v: number) => `<c r="${r}"><v>${v}</v></c>`;
const formula = (r: string, f: string) => `<c r="${r}"><f>${f}</f><v>0</v></c>`;

async function workbook(names: string[], first: string): Promise<Workbook> {
  const file = await open(
    await buildXlsx({
      sheets: names.map((name, i) => ({
        name,
        xml: worksheetXml(i === 0 ? first : `<sheetData><row r="1">${num('A1', i)}</row><row r="2">${num('A2', 10 * i)}</row></sheetData>`),
      })),
    }),
  );
  return file.workbook();
}

function formulaAt(wb: Workbook, sheet: number, ref: string): string | undefined {
  const { row, col } = parseCellRef(ref)!;
  return wb.sheets[sheet].rows.get(row)?.cells[col]?.formula?.text;
}

const run = (history: History, ops: Op[]) => history.apply(ops, undefined, undefined);

// --- 1. what is left of a reference to a deleted sheet ---------------------------

test('review: a deleted sheet takes the error and the spill mark after its name with it', async () => {
  const gone = (text: string) => renameSheetsInFormula(text, (name) => (name === 'Sheet2' ? null : undefined));
  assert.equal(gone('Sheet2!#REF!'), '#REF!');
  assert.equal(gone('SUM(Sheet2!#REF!)+Sheet2!#REF!*2'), 'SUM(#REF!)+#REF!*2');
  assert.equal(gone("'Sheet2'!#REF!+1"), '#REF!+1');
  assert.equal(gone('Sheet2!A1#'), '#REF!');
  assert.equal(gone('SUM(Sheet2!A1:B2#)'), 'SUM(#REF!)');
  assert.equal(gone('Sheet2!#N/A'), '#REF!', 'whatever error stands for the cell');
  assert.equal(gone('Sheet2!A1+#REF!'), '#REF!+#REF!', 'an error that is not part of the reference stays');
  assert.equal(gone('Sheet3!#REF!+Sheet3!A1#'), 'Sheet3!#REF!+Sheet3!A1#', 'other sheets are not touched');
  // A new name leaves what follows the prefix as it is.
  const renamed = (text: string) => renameSheetsInFormula(text, (name) => (name === 'Sheet2' ? 'Yeni' : undefined));
  assert.equal(renamed('Sheet2!#REF!'), 'Yeni!#REF!');
  assert.equal(renamed('SUM(Sheet2!A1#)'), 'SUM(Yeni!A1#)');

  const wb = await workbook(['Sheet1', 'Sheet2', 'Sheet3'], `<sheetData><row r="1">${formula('A1', 'Sheet2!A2+1')}${formula('B1', 'SUM(Sheet2!#REF!)')}</row></sheetData>`);
  const history = new History(wb);
  run(history, [{ type: 'delete', sheet: 1, axis: 'row', at: 1, count: 1 }]);
  assert.equal(formulaAt(wb, 0, 'A1'), 'Sheet2!#REF!+1');
  run(history, renameSheet(wb, 1, 'Yeni'));
  assert.equal(formulaAt(wb, 0, 'A1'), 'Yeni!#REF!+1');
  assert.equal(formulaAt(wb, 0, 'B1'), 'SUM(Yeni!#REF!)');
  history.undo();
  run(history, deleteSheet(wb, 1));
  assert.equal(formulaAt(wb, 0, 'A1'), '#REF!+1');
  assert.equal(formulaAt(wb, 0, 'B1'), 'SUM(#REF!)');
  history.undo();
  assert.equal(formulaAt(wb, 0, 'A1'), 'Sheet2!#REF!+1');
});

test('review: a prefix none of whose names changes keeps its spelling', () => {
  // The map answers with the name as it was asked for a sheet that keeps its name.
  const map = (name: string) => (name.toLowerCase() === 'sheet3' ? 'X' : name.toLowerCase() === 'sheet1' ? name : undefined);
  assert.equal(renameSheetsInFormula("'sheet1'!A1+SHEET3!A1+'Sheet 9'!B2+sheet1!C3", map), "'sheet1'!A1+X!A1+'Sheet 9'!B2+sheet1!C3");
  assert.equal(renameSheetsInFormula("'sheet1:sheet3'!A1", map), 'sheet1:X!A1', 'a prefix that changes is written anew');
  assert.equal(renameSheetsInFormula("'Sheet1'!A1:'Sheet1'!B2", map), "'Sheet1'!A1:'Sheet1'!B2");
});

// --- 5. 3D references ------------------------------------------------------------

test('review: a 3D reference whose end is deleted ends at the sheet next to it', async () => {
  const wb = await workbook(
    ['Sheet1', 'Sheet2', 'Sheet3', 'Sheet4', 'Sheet5'],
    `<sheetData><row r="1">${formula('A1', 'SUM(Sheet2:Sheet4!A1)')}${formula('B1', 'SUM(Sheet2:Sheet2!A1)')}${formula('C1', "SUM('Sheet3:Sheet5'!A1:B2)+Sheet4!A1")}${formula('D1', 'SUM(Sheet2:Sheet3!A1)')}</row></sheetData>` +
      '<dataValidations count="2"><dataValidation type="whole" sqref="A5"><formula1>SUM(Sheet2:Sheet4!$A$1)</formula1><formula2>SUM(Sheet2:Sheet2!$A$1)</formula2></dataValidation>' +
      '<dataValidation type="whole" sqref="B5"><formula1>SUM(Sheet3:Sheet5!$A$1)</formula1><formula2>SUM(Sheet2:Sheet3!$A$1)</formula2></dataValidation></dataValidations>',
  );
  const history = new History(wb);
  const rules = new DataValidations(wb, wb.sheets[0], {});
  const raw = () => [rules.ruleAt(4, 0), rules.ruleAt(4, 1)].map((rule) => [rule?.formula1, rule?.formula2]);
  const live = () => ['A1', 'B1', 'C1', 'D1'].map((ref) => formulaAt(wb, 0, ref));
  const start = { live: live(), raw: raw() };

  // The first end goes: the run starts at the next sheet. A run of the one deleted sheet is gone.
  run(history, deleteSheet(wb, 1));
  assert.deepEqual(live(), ['SUM(Sheet3:Sheet4!A1)', 'SUM(#REF!)', "SUM('Sheet3:Sheet5'!A1:B2)+Sheet4!A1", 'SUM(Sheet3!A1)']);
  // What is kept as the file has it reads the same.
  assert.deepEqual(raw(), [
    ['SUM(Sheet3:Sheet4!$A$1)', 'SUM(#REF!)'],
    ['SUM(Sheet3:Sheet5!$A$1)', 'SUM(Sheet3!$A$1)'],
  ]);

  // The last end goes: the run ends at the sheet before it; a run of one sheet is a plain reference.
  run(history, deleteSheet(wb, 2));
  assert.deepEqual(live(), ['SUM(Sheet3!A1)', 'SUM(#REF!)', "SUM('Sheet3:Sheet5'!A1:B2)+#REF!", 'SUM(Sheet3!A1)']);
  assert.deepEqual(raw(), [
    ['SUM(Sheet3!$A$1)', 'SUM(#REF!)'],
    ['SUM(Sheet3:Sheet5!$A$1)', 'SUM(Sheet3!$A$1)'],
  ]);

  // A sheet put where the deleted one was, under its name, is not what the file meant.
  run(history, insertSheet(wb, 1, { name: 'Sheet2' }));
  assert.deepEqual(raw()[0], ['SUM(Sheet3!$A$1)', 'SUM(#REF!)']);
  history.undo();

  // The last sheet a run had goes: nothing is left.
  run(history, deleteSheet(wb, 1));
  assert.deepEqual(live(), ['SUM(#REF!)', 'SUM(#REF!)', 'SUM(Sheet5!A1:B2)+#REF!', 'SUM(#REF!)']);
  assert.deepEqual(raw(), [
    ['SUM(#REF!)', 'SUM(#REF!)'],
    ['SUM(Sheet5!$A$1)', 'SUM(#REF!)'],
  ]);

  // A renamed neighbour is named as it is called now.
  history.undo();
  history.undo();
  run(history, renameSheet(wb, 1, 'Üç'));
  run(history, deleteSheet(wb, 2));
  assert.equal(live()[0], "SUM('Üç'!A1)");
  assert.deepEqual(raw()[0], ["SUM('Üç'!$A$1)", 'SUM(#REF!)']);
  assert.deepEqual(raw()[1], ["SUM('Üç:Sheet5'!$A$1)", "SUM('Üç'!$A$1)"]);

  while (history.canUndo) history.undo();
  assert.deepEqual({ live: live(), raw: raw() }, start);
  assert.equal(sheetNameMap(wb), undefined);
});

// --- 3. content that cannot follow a new name --------------------------------------

test('review: a workbook with form controls refuses renaming and deleting sheets', async () => {
  for (const kind of ['formControl', 'ctrlProp', 'control']) {
    const wb = await workbook(['Sheet1', 'Sheet2', 'Sheet3'], '<sheetData/>');
    wb.sheets[2].anchored = [kind];
    assert.equal(canRenameSheet(wb, 0, 'Yeni'), 'Cannot rename a sheet in a workbook with form controls: not supported yet');
    assert.equal(canRenameSheet(wb, 0), 'Cannot rename a sheet in a workbook with form controls: not supported yet');
    assert.equal(canDeleteSheet(wb, 1), 'Cannot delete a sheet in a workbook with form controls: not supported yet');
    assert.throws(() => renameSheet(wb, 0, 'Yeni'), RangeError);
    assert.throws(() => deleteSheet(wb, 1), RangeError);
    assert.throws(() => new History(wb).apply([{ type: 'renameSheet', sheet: 0, name: 'Yeni' }], undefined, undefined), /form controls/);
    assert.throws(() => new History(wb).apply([{ type: 'deleteSheet', sheet: 1 }], undefined, undefined), /form controls/);
    // A new sheet changes no name.
    assert.equal(canInsertSheet(wb, 0), undefined);
  }
});

// --- 6. names that differ only by case ----------------------------------------------

test('review: sheet names are folded with the four Turkish letters as one', () => {
  assert.equal(foldSheetName('I'), foldSheetName('ı'));
  assert.equal(foldSheetName('İ'), foldSheetName('i'));
  assert.equal(foldSheetName('I'), foldSheetName('İ'));
  assert.equal(foldSheetName('IŞIK İLİ'), foldSheetName('ışık ili'));
  assert.notEqual(foldSheetName('IL'), foldSheetName('OL'));
  assert.equal(foldSheetName('ÇĞÖŞÜ'), foldSheetName('çğöşü'));
});

test('review: a name another sheet has in other Turkish letter case is refused', async () => {
  const wb = await workbook(['ILIK', 'Sheet2'], '<sheetData/>');
  assert.match(String(canRenameSheet(wb, 1, 'ılık')), /another sheet has that name/);
  assert.match(String(canRenameSheet(wb, 1, 'İLİK')), /another sheet has that name/);
  assert.match(String(canInsertSheet(wb, 0, 'ilik')), /another sheet has that name/);
  assert.equal(canRenameSheet(wb, 0, 'ılık'), undefined, 'its own name in other case');
  // References written in either case are the same sheet.
  const first = await workbook(['Sheet1', 'ILIK'], `<sheetData><row r="1">${formula('A1', 'ılık!A1+İLİK!A2')}</row></sheetData>`);
  new History(first).apply(renameSheet(first, 1, 'Yeni'), undefined, undefined);
  assert.equal(formulaAt(first, 0, 'A1'), 'Yeni!A1+Yeni!A2');
});

test('review: a workbook with two sheets whose names differ only by case refuses sheet operations', async () => {
  const wb = await workbook(['Sheet1', 'IL', 'İL'], `<sheetData><row r="1">${formula('A1', "IL!A1+'İL'!A1")}</row></sheetData>`);
  const message = 'Cannot change the sheets: "IL" and "İL" differ only by case';
  assert.equal(canRenameSheet(wb, 2, 'Yeni'), message);
  assert.equal(canRenameSheet(wb, 0), message);
  assert.equal(canDeleteSheet(wb, 1), message);
  assert.equal(canInsertSheet(wb, 0), message);
  assert.throws(() => renameSheet(wb, 2, 'Yeni'), RangeError);
  assert.throws(() => deleteSheet(wb, 0), RangeError);
  assert.throws(() => insertSheet(wb, 0), RangeError);
  for (const op of [{ type: 'renameSheet', sheet: 2, name: 'Yeni' }, { type: 'deleteSheet', sheet: 1 }, { type: 'insertSheet', sheet: 0, name: 'Yeni' }] as Op[]) {
    assert.throws(() => new History(wb).apply([op], undefined, undefined), /differ only by case/);
  }
  assert.equal(formulaAt(wb, 0, 'A1'), "IL!A1+'İL'!A1");
  assert.equal(wb.sheetsModified, undefined);
});

// --- 8. gaps ---------------------------------------------------------------------

test('review: a sheet renamed and then deleted is gone under the name the file has for it', async () => {
  const wb = await workbook(
    ['Sheet1', 'Sheet2', 'Sheet3'],
    '<sheetData/><dataValidations count="1"><dataValidation type="whole" sqref="A5"><formula1>Sheet2!$A$1</formula1><formula2>Sheet3!$A$1</formula2></dataValidation></dataValidations>',
  );
  const history = new History(wb);
  const rules = new DataValidations(wb, wb.sheets[0], {});
  run(history, renameSheet(wb, 1, 'Geçici'));
  assert.equal(rules.ruleAt(4, 0)?.formula1, "'Geçici'!$A$1");
  run(history, deleteSheet(wb, 1));
  assert.deepEqual(wb.removedSheetNames, ['Sheet2']);
  assert.equal(currentSheetName(wb, 'Sheet2'), null);
  assert.equal(currentSheetName(wb, 'Geçici'), undefined, 'no sheet of the file has that name');
  assert.deepEqual([rules.ruleAt(4, 0)?.formula1, rules.ruleAt(4, 0)?.formula2], ['#REF!', 'Sheet3!$A$1']);
  history.undo();
  assert.equal(rules.ruleAt(4, 0)?.formula1, "'Geçici'!$A$1");
  history.undo();
  assert.equal(rules.ruleAt(4, 0)?.formula1, 'Sheet2!$A$1');
});

test('review: deleted sheets count towards what the undo stack keeps alive', () => {
  const wb: Workbook = { sheets: [], styles: emptyStyleSheet(), sharedStrings: [], definedNames: [], date1904: false, activeSheet: 0 };
  for (const [i, name] of ['Small', 'Big1', 'Big2', 'Last'].entries()) {
    const sheet = createSheet({ name, sheetId: i + 1, state: 'visible', kind: 'worksheet', part: '' });
    // 200,000 rows each: two such sheets are more than the stack may hold.
    if (name.startsWith('Big')) for (let r = 0; r < 200000; r++) sheet.rows.set(r, { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [] });
    wb.sheets.push(sheet);
  }
  const history = new History(wb);
  run(history, renameSheet(wb, 0, 'Küçük'));
  run(history, deleteSheet(wb, 1));
  run(history, deleteSheet(wb, 1));
  assert.deepEqual(wb.sheets.map((sheet) => sheet.name), ['Küçük', 'Last']);
  // The newest step always stays; the two before it were dropped to make room.
  assert.ok(history.undo());
  assert.deepEqual(wb.sheets.map((sheet) => sheet.name), ['Küçük', 'Big2', 'Last']);
  assert.equal(history.canUndo, false);
  assert.equal(history.undo(), undefined);
});
