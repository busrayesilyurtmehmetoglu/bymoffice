// Operations on the sheets of a workbook: insert, delete, rename; what they do to formulas,
// defined names and links, and how they are undone.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine, getCell, open, type Sheet, type Workbook } from '../src/index';
import { formulaNamesSheetsUnread, quoteSheetName, renameSheetsInFormula } from '../src/calc/refs';
import { currentSheetName, fileSheetName, sheetNameMap } from '../src/model/sheet/types';
import { setCellInput } from '../src/ops/commands';
import { History, applyOp, type Op } from '../src/ops/ops';
import { canDeleteSheet, canInsertSheet, canRenameSheet, defaultSheetName, deleteSheet, insertSheet, renameSheet } from '../src/ops/sheets';
import { buildXlsx, worksheetXml } from './fixtures';

const num = (r: string, v: number) => `<c r="${r}"><v>${v}</v></c>`;
const formula = (r: string, f: string, v = 0, attrs = '') => `<c r="${r}"><f${attrs}>${f}</f><v>${v}</v></c>`;

/**
 * Three sheets. "Sheet1" and "My Data" hold numbers; "Sheet3" reads them in every way a
 * formula can name a sheet. Defined names and links name the sheets too.
 */
async function workbook(): Promise<Workbook> {
  const file = await open(
    await buildXlsx({
      workbookTail:
        '<definedNames>' +
        '<definedName name="Total">Sheet1!$A$1:$A$2</definedName>' +
        '<definedName name="Both">SUM(Sheet1!$A$1,\'My Data\'!$A$1)</definedName>' +
        '<definedName name="OnData" localSheetId="1">\'My Data\'!$B$1</definedName>' +
        '<definedName name="OnThree" localSheetId="2">Sheet3!$A$1</definedName>' +
        '</definedNames>',
      sheets: [
        { name: 'Sheet1', xml: worksheetXml(`<sheetData><row r="1">${num('A1', 1)}${num('B1', 10)}</row><row r="2">${num('A2', 2)}${formula('B2', 'Sheet1!A1+A2', 3)}</row><row r="3">${num('A3', 3)}</row></sheetData>`) },
        {
          name: 'My Data',
          xml: worksheetXml(`<sheetData><row r="1">${num('A1', 100)}${num('B1', 200)}</row><row r="2">${num('A2', 300)}${formula('B2', 'Sheet1!A1*2', 2)}</row></sheetData>`),
        },
        {
          name: 'Sheet3',
          xml: worksheetXml(
            '<sheetData>' +
              `<row r="1">${formula('A1', 'Sheet1!A1+1', 2)}${formula('B1', "SUM('My Data'!A1:A2)", 400)}${formula('C1', 'SUM(Sheet1:Sheet3!A9)', 0)}` +
              `${formula('D1', '[1]Sheet1!A1+1', 5)}${formula('E1', '"Sheet1!A1"&amp;Sheet1!B1', 0)}${formula('F1', 'SUM(sheet1!$A$1:SHEET1!$A$2)', 3)}${formula('G1', 'SUM(Total)+Sheet1!Total', 3)}</row>` +
              `<row r="2">${formula('A2', 'Sheet1!A1*2', 2, ' t="shared" ref="A2:A4" si="0"')}${formula('B2', "'My Data'!A1+'My Data'!B1", 300)}</row>` +
              '<row r="3"><c r="A3"><f t="shared" si="0"/><v>4</v></c></row>' +
              '<row r="4"><c r="A4"><f t="shared" si="0"/><v>6</v></c></row>' +
              `<row r="5">${formula('A5', 'A1+B1', 402)}</row>` +
              '</sheetData>' +
              '<hyperlinks><hyperlink ref="A6" location="Sheet1!A1"/><hyperlink ref="A7" location="\'My Data\'!B2"/><hyperlink ref="A8" location="Sheet3!A1"/><hyperlink ref="A9" location="Total"/></hyperlinks>',
          ),
        },
      ],
    }),
  );
  return file.workbook();
}

const formulaAt = (sheet: Sheet, ref: string) => {
  const m = /^([A-Z])(\d+)$/.exec(ref)!;
  return getCell(sheet, Number(m[2]) - 1, m[1].charCodeAt(0) - 65)?.formula?.text;
};
const valueAt = (sheet: Sheet, ref: string) => {
  const m = /^([A-Z])(\d+)$/.exec(ref)!;
  return getCell(sheet, Number(m[2]) - 1, m[1].charCodeAt(0) - 65)?.value;
};
const byName = (wb: Workbook, name: string) => wb.sheets.find((sheet) => sheet.name === name)!;

/**
 * Everything operations on sheets may touch, for comparing states. Flags that say "written
 * since read" are left out: undo does not clear them. So are the values of formulas, which
 * are the engine's to write.
 */
function picture(wb: Workbook): unknown {
  return {
    active: wb.activeSheet,
    removed: wb.removedSheetNames,
    names: wb.definedNames.map((n) => [n.name, n.value, n.localSheet, n.hidden]),
    sheets: wb.sheets.map((sheet) => ({
      name: sheet.name,
      nameInFile: sheet.nameInFile,
      part: sheet.part,
      sheetId: sheet.sheetId,
      rowCount: sheet.rowCount,
      colCount: sheet.colCount,
      links: sheet.hyperlinks.map((h) => [h.range, h.location]),
      edits: sheet.structuralEdits,
      // A row left without cells by a deleted column is dropped by the next change of its cells: not a difference.
      rows: [...sheet.rows.entries()]
        .filter(([, row]) => row.cells.some(() => true))
        .sort((a, b) => a[0] - b[0])
        .map(([index, row]) => [index, row.cells.map((cell, col) => (cell.formula ? [col, cell.style, cell.formula] : [col, cell.kind, cell.value, cell.style])).filter(Boolean)]),
    })),
  };
}

/** Calculates everything again, as the view does after a structural change. */
function recalculate(engine: CalcEngine): void {
  engine.invalidate();
  engine.recalculate(undefined, { overwrite: true });
  for (let i = 0; engine.incomplete && i < 1000; i++) engine.recalculate([]);
  assert.equal(engine.incomplete, false);
}

// ---------------------------------------------------------------------------
// The text of formulas

test('sheet names in formulas: one pass, quoting, 3D, what is never touched', () => {
  const to = (from: string, name: string | null) => (written: string) => (written.toLowerCase() === from.toLowerCase() ? name : undefined);
  const rename = (text: string, from: string, name: string | null) => renameSheetsInFormula(text, to(from, name));

  assert.equal(rename('Sheet1!A1+1', 'Sheet1', 'Data'), 'Data!A1+1');
  assert.equal(rename('SHEET1!$a$1+sheet1!B2', 'Sheet1', 'Data'), 'Data!$a$1+Data!B2', 'the reference keeps its own spelling');
  assert.equal(rename('Sheet1!A1', 'Sheet1', 'My Data'), "'My Data'!A1");
  assert.equal(rename("'My Data'!A1:B2", 'My Data', 'Data'), 'Data!A1:B2');
  assert.equal(rename("'My Data'!A1", 'My Data', "Bob's"), "'Bob''s'!A1");
  assert.equal(rename("'Bob''s'!A1", "Bob's", 'B'), 'B!A1');
  // Names that could be read as something else are quoted.
  for (const name of ['A1', 'R1C1', 'R', 'C', 'rc', 'TRUE', '2024', 'a-b', 'Özet', 'XFD1048576', 'a.b']) assert.equal(rename('S!A1', 'S', name), `'${name}'!A1`, name);
  for (const name of ['Data', '_x', 'Sheet_2', 'XFE1', 'A1048577']) assert.equal(rename('S!A1', 'S', name), `${name}!A1`, name);
  assert.equal(quoteSheetName("it's"), "'it''s'");
  assert.equal(quoteSheetName('Data'), 'Data');

  // Both ends of a 3D reference; the other end keeps its name and the pair is quoted as one.
  assert.equal(rename('SUM(Sheet1:Sheet3!A1)', 'Sheet1', 'First'), 'SUM(First:Sheet3!A1)');
  assert.equal(rename('SUM(Sheet1:Sheet3!A1)', 'Sheet3', 'The End'), "SUM('Sheet1:The End'!A1)");
  assert.equal(rename("SUM('My Data:Sheet3'!A1:B2)", 'My Data', 'D'), 'SUM(D:Sheet3!A1:B2)');
  // A range whose corners both carry the name.
  assert.equal(rename('SUM(Sheet1!A1:Sheet1!B2)', 'Sheet1', 'My Data'), "SUM('My Data'!A1:'My Data'!B2)");
  // A name defined on a sheet, whole columns and rows.
  assert.equal(rename('Sheet1!Rate*Sheet1!A:A+Sheet1!1:2', 'Sheet1', 'D'), 'D!Rate*D!A:A+D!1:2');

  // Other workbooks, text and structured references stay.
  for (const text of ['[1]Sheet1!A1', "'[book.xlsx]Sheet1'!A1", "'C:\\dir\\[b.xlsx]Sheet1'!A1", '"Sheet1!A1"', 'INDIRECT("Sheet1!A1")', 'Table1[[#This Row],[Sheet1]]', 'Sheet1+Sheet1A1', 'Sheet2!A1', 'Sheet11!A1']) {
    assert.equal(rename(text, 'Sheet1', 'Data'), text, text);
    assert.equal(rename(text, 'Sheet1', null), text, text);
  }

  // A sheet that is gone takes the whole reference with it.
  assert.equal(rename('Sheet1!A1+1', 'Sheet1', null), '#REF!+1');
  assert.equal(rename("SUM('My Data'!$A$1:$B$2,Sheet1!A1)", 'My Data', null), 'SUM(#REF!,Sheet1!A1)');
  assert.equal(rename('SUM(Sheet1!A1:Sheet1!B2)+Sheet1!A:A+Sheet1!Rate', 'Sheet1', null), 'SUM(#REF!)+#REF!+#REF!');
  assert.equal(rename('SUM(Sheet1:Sheet3!A1)', 'Sheet3', null), 'SUM(#REF!)');
  assert.equal(rename('#REF!A1+Sheet1!A1', 'Sheet1', 'D'), '#REF!A1+D!A1');

  // All names are looked up at once: a swap is a swap, not a chain.
  const swap = (written: string) => (written === 'A_' ? 'B_' : written === 'B_' ? 'A_' : undefined);
  assert.equal(renameSheetsInFormula('A_!A1+B_!A1+A_:B_!C3', swap), 'B_!A1+A_!A1+B_:A_!C3');
  const reuse = (written: string) => (written === 'Sheet2' ? null : written === 'Sheet3' ? 'Sheet2' : undefined);
  assert.equal(renameSheetsInFormula('Sheet2!A1+Sheet3!A1', reuse), '#REF!+Sheet2!A1');

  // Never throws; what cannot be read is left alone from there on, and can be asked about.
  assert.equal(rename('Sheet1!A1+"abc', 'Sheet1', 'D'), 'D!A1+"abc');
  assert.equal(formulaNamesSheetsUnread('Sheet1!A1+"abc'), false, 'the unread rest names no sheet');
  assert.equal(rename('Sheet1!A1+"abc"&"x+Sheet1!A2', 'Sheet1', 'D'), 'D!A1+"abc"&"x+Sheet1!A2');
  assert.equal(formulaNamesSheetsUnread('Sheet1!A1+"abc"&"x+Sheet1!A2'), true);
  assert.equal(formulaNamesSheetsUnread("A1+'Sheet 1!A2"), true);
  assert.equal(formulaNamesSheetsUnread('A1+[x'), false);
  assert.equal(formulaNamesSheetsUnread("'My Data'!A1+Sheet1!B2"), false);
  assert.equal(renameSheetsInFormula(undefined as unknown as string, swap), '');
  assert.equal(renameSheetsInFormula('', swap), '');
});

// ---------------------------------------------------------------------------
// Rename

test('renameSheet: rewrites formulas, defined names and links of every sheet; undo restores the texts exactly', async () => {
  const wb = await workbook();
  const before = picture(wb);
  const three = byName(wb, 'Sheet3');
  const history = new History(wb);
  history.apply(renameSheet(wb, 0, "Bob's data"), undefined, undefined);

  const q = "'Bob''s data'";
  assert.equal(wb.sheets[0].name, "Bob's data");
  assert.equal(wb.sheets[0].nameInFile, 'Sheet1');
  assert.equal(fileSheetName(wb.sheets[0]), 'Sheet1');
  assert.equal(wb.sheetsModified, true);
  assert.equal(formulaAt(three, 'A1'), `${q}!A1+1`);
  assert.equal(formulaAt(three, 'B1'), "SUM('My Data'!A1:A2)", 'another sheet is left alone');
  assert.equal(formulaAt(three, 'C1'), "SUM('Bob''s data:Sheet3'!A9)");
  assert.equal(formulaAt(three, 'D1'), '[1]Sheet1!A1+1');
  assert.equal(formulaAt(three, 'E1'), `"Sheet1!A1"&${q}!B1`);
  assert.equal(formulaAt(three, 'F1'), `SUM(${q}!$A$1:${q}!$A$2)`);
  assert.equal(formulaAt(three, 'G1'), `SUM(Total)+${q}!Total`);
  // The shared formula stays shared: its text changes in the one cell that holds it.
  assert.deepEqual(getCell(three, 1, 0)?.formula, { text: `${q}!A1*2`, type: 'shared', ref: 'A2:A4', si: 0 });
  assert.deepEqual(getCell(three, 2, 0)?.formula, { text: '', type: 'shared', si: 0 });
  // The sheet's own formulas that name it, and those of other sheets.
  assert.equal(formulaAt(wb.sheets[0], 'B2'), `${q}!A1+A2`);
  assert.equal(formulaAt(wb.sheets[1], 'B2'), `${q}!A1*2`);
  assert.deepEqual(
    wb.definedNames.map((n) => n.value),
    [`${q}!$A$1:$A$2`, `SUM(${q}!$A$1,'My Data'!$A$1)`, "'My Data'!$B$1", 'Sheet3!$A$1'],
  );
  assert.deepEqual(
    three.hyperlinks.map((h) => h.location),
    [`${q}!A1`, "'My Data'!B2", 'Sheet3!A1', 'Total'],
  );
  // Sheets whose formulas changed are marked for the writer, row by row; the others are not.
  assert.deepEqual([three.modified, [...three.changedRows!].sort()], [true, [0, 1]]);
  assert.equal(three.hyperlinksModified, true);
  assert.deepEqual([...wb.sheets[1].changedRows!], [1]);

  // The engine, built again, finds the sheet under its new name.
  const engine = new CalcEngine(wb);
  recalculate(engine);
  assert.deepEqual([valueAt(three, 'A1'), valueAt(three, 'A2'), valueAt(three, 'A3'), valueAt(three, 'A4'), valueAt(three, 'F1')], [2, 2, 4, 6, 3]);

  const after = picture(wb);
  history.undo();
  assert.deepEqual(picture(wb), before);
  assert.equal(wb.sheets[0].nameInFile, undefined);
  assert.equal('nameInFile' in wb.sheets[0], false);
  assert.equal(formulaAt(three, 'F1'), 'SUM(sheet1!$A$1:SHEET1!$A$2)', 'the spelling of the file comes back, not a rename back');
  history.redo();
  assert.deepEqual(picture(wb), after);
  history.undo();
  assert.deepEqual(picture(wb), before);
});

test('renameSheet: a second rename keeps the name of the file; case-only renames; the same name is nothing to do', async () => {
  const wb = await workbook();
  const history = new History(wb);
  history.apply(renameSheet(wb, 1, 'Data'), undefined, undefined);
  history.apply(renameSheet(wb, 1, 'Veri'), undefined, undefined);
  assert.deepEqual([wb.sheets[1].name, wb.sheets[1].nameInFile], ['Veri', 'My Data']);
  assert.equal(formulaAt(wb.sheets[2], 'B2'), 'Veri!A1+Veri!B1');
  // Back to the name of the file by a rename: the field stays, equal to the name, and the map has nothing to say.
  history.apply(renameSheet(wb, 1, 'My Data'), undefined, undefined);
  assert.deepEqual([wb.sheets[1].name, wb.sheets[1].nameInFile, sheetNameMap(wb)], ['My Data', 'My Data', undefined]);
  history.undo();
  history.undo();
  assert.deepEqual([wb.sheets[1].name, wb.sheets[1].nameInFile], ['Data', 'My Data']);
  assert.equal(currentSheetName(wb, 'my data'), 'Data');
  assert.equal(currentSheetName(wb, 'Data'), undefined);
  history.undo();
  assert.equal('nameInFile' in wb.sheets[1], false);

  assert.deepEqual(renameSheet(wb, 0, 'Sheet1'), []);
  assert.equal(history.apply(renameSheet(wb, 0, 'Sheet1'), undefined, undefined).length, 0);
  // Another letter case is a rename, and no clash with itself.
  history.apply(renameSheet(wb, 0, 'SHEET1'), undefined, undefined);
  assert.equal(wb.sheets[0].name, 'SHEET1');
  assert.equal(formulaAt(wb.sheets[2], 'A1'), 'SHEET1!A1+1');
  assert.equal(formulaAt(wb.sheets[2], 'F1'), 'SUM(SHEET1!$A$1:SHEET1!$A$2)');
});

// ---------------------------------------------------------------------------
// Delete

test('deleteSheet: #REF! elsewhere, local names, shifted indexes, the active sheet; undo and redo', async () => {
  const wb = await workbook();
  wb.activeSheet = 1;
  const before = picture(wb);
  const sheets = wb.sheets.slice();
  const three = sheets[2];
  const engine = new CalcEngine(wb);
  engine.recalculate();
  const history = new History(wb);
  history.apply(deleteSheet(wb, 1), undefined, undefined);

  assert.deepEqual(wb.sheets, [sheets[0], sheets[2]]);
  assert.deepEqual(wb.removedSheetNames, ['My Data']);
  assert.equal(wb.sheetsModified, true);
  assert.equal(wb.activeSheet, 1, 'the next visible sheet is shown');
  assert.equal(formulaAt(three, 'B1'), 'SUM(#REF!)');
  assert.equal(formulaAt(three, 'B2'), '#REF!+#REF!');
  assert.equal(formulaAt(three, 'A1'), 'Sheet1!A1+1');
  assert.deepEqual(
    wb.definedNames.map((n) => [n.name, n.value, n.localSheet]),
    [
      ['Total', 'Sheet1!$A$1:$A$2', undefined],
      ['Both', 'SUM(Sheet1!$A$1,#REF!)', undefined],
      ['OnThree', 'Sheet3!$A$1', 1],
    ],
  );
  assert.deepEqual(
    three.hyperlinks.map((h) => h.location),
    ['Sheet1!A1', '#REF!', 'Sheet3!A1', 'Total'],
  );
  // The engine, built again for one sheet less, reads no sheet that is not there.
  recalculate(engine);
  assert.deepEqual([valueAt(three, 'A1'), valueAt(three, 'B1'), valueAt(three, 'B2'), valueAt(three, 'A4')], [2, '#REF!', '#REF!', 6]);

  const after = picture(wb);
  history.undo();
  assert.deepEqual(picture(wb), before);
  assert.deepEqual(wb.sheets, sheets);
  assert.ok(wb.sheets.every((sheet, i) => sheet === sheets[i]), 'the very sheet objects come back');
  assert.equal(wb.removedSheetNames, undefined);
  assert.equal('removedSheetNames' in wb, false);
  recalculate(engine);
  assert.deepEqual([valueAt(three, 'B1'), valueAt(three, 'B2')], [400, 300]);
  history.redo();
  assert.deepEqual(picture(wb), after);
  history.undo();
  assert.deepEqual(picture(wb), before);
});

test('deleteSheet: the active sheet stays valid and visible; a 3D reference that ends at the sheet', async () => {
  const wb = await workbook();
  // Active sheet after the deleted one: it follows its sheet.
  wb.activeSheet = 2;
  const undoFirst = applyOp(wb, { type: 'deleteSheet', sheet: 0 });
  assert.equal(wb.activeSheet, 1);
  assert.equal(formulaAt(wb.sheets[1], 'C1'), "SUM('My Data:Sheet3'!A9)", 'the run starts at the sheet that stood next to the deleted one');
  assert.equal(formulaAt(wb.sheets[1], 'D1'), '[1]Sheet1!A1+1');
  assert.equal(formulaAt(wb.sheets[1], 'E1'), '"Sheet1!A1"&#REF!');
  assert.equal(formulaAt(wb.sheets[1], 'G1'), 'SUM(Total)+#REF!');
  assert.deepEqual(getCell(wb.sheets[1], 1, 0)?.formula, { text: '#REF!*2', type: 'shared', ref: 'A2:A4', si: 0 });
  assert.deepEqual(getCell(wb.sheets[1], 3, 0)?.formula, { text: '', type: 'shared', si: 0 });
  applyOp(wb, undoFirst);
  assert.equal(wb.activeSheet, 2);

  // The last sheet is deleted while shown: the one before it is, unless it is hidden.
  applyOp(wb, { type: 'deleteSheet', sheet: 2 });
  assert.equal(wb.activeSheet, 1);
  const hidden = await workbook();
  hidden.sheets[1].state = 'hidden';
  hidden.activeSheet = 2;
  applyOp(hidden, { type: 'deleteSheet', sheet: 2 });
  assert.equal(hidden.activeSheet, 0);
  // Before the active sheet nothing changes.
  const first = await workbook();
  first.activeSheet = 0;
  applyOp(first, { type: 'deleteSheet', sheet: 2 });
  assert.equal(first.activeSheet, 0);
});

// ---------------------------------------------------------------------------
// Insert

test('insertSheet: default names, shifted indexes, undo and redo keep the sheet object', async () => {
  const wb = await workbook();
  assert.equal(defaultSheetName(wb), 'Sheet2');
  assert.equal(defaultSheetName(wb, 'en-GB'), 'Sheet2');
  assert.equal(defaultSheetName(wb, 'tr'), 'Sayfa1');
  assert.equal(defaultSheetName(wb, 'tr-TR'), 'Sayfa1');
  assert.equal(defaultSheetName(wb, 'TR_tr'), 'Sayfa1');
  assert.equal(defaultSheetName(wb, 'tra'), 'Sheet2');
  wb.sheets[1].name = 'SHEET2';
  assert.equal(defaultSheetName(wb), 'Sheet4', 'names compare without regard to case');
  wb.sheets[1].name = 'SAYFA1';
  assert.equal(defaultSheetName(wb, 'tr'), 'Sayfa2');
  wb.sheets[1].name = 'sayfa\u01311';
  assert.equal(defaultSheetName(wb, 'tr'), 'Sayfa1', 'only what folds to the same name is taken');
  wb.sheets[1].name = 'SAYFA\u0130';
  wb.sheets[1].name = 'My Data';

  wb.activeSheet = 1;
  const before = picture(wb);
  const old = wb.sheets.slice();
  const history = new History(wb);
  history.apply(insertSheet(wb, 1, { locale: 'tr' }), undefined, undefined);
  const added = wb.sheets[1];
  assert.deepEqual(wb.sheets, [old[0], added, old[1], old[2]]);
  assert.deepEqual([added.name, added.part, added.kind, added.state, added.sheetId, added.modified, added.rows.size], ['Sayfa1', '', 'worksheet', 'visible', 4, true, 0]);
  assert.equal(fileSheetName(added), undefined);
  assert.equal(wb.activeSheet, 2, 'the sheet that was shown stays the one that is');
  assert.equal(wb.sheetsModified, true);
  assert.equal(wb.removedSheetNames, undefined);
  assert.deepEqual(
    wb.definedNames.map((n) => n.localSheet),
    [undefined, undefined, 2, 3],
  );
  // A formula on the new sheet reads the others, and they read it.
  history.apply(setCellInput(wb, 1, 0, 0, "=Sheet1!A1+'My Data'!A1"), undefined, undefined);
  history.apply(setCellInput(wb, 0, 5, 0, '=Sayfa1!A1*2'), undefined, undefined);
  const engine = new CalcEngine(wb);
  recalculate(engine);
  assert.deepEqual([valueAt(added, 'A1'), valueAt(wb.sheets[0], 'A6')], [101, 202]);

  // Renaming a sheet that is not in the file leaves no name of the file behind.
  history.apply(renameSheet(wb, 1, 'Yeni'), undefined, undefined);
  assert.deepEqual([added.name, added.nameInFile, formulaAt(wb.sheets[0], 'A6')], ['Yeni', undefined, 'Yeni!A1*2']);
  // Deleting it again does not list it among the sheets of the file that are gone.
  history.apply(deleteSheet(wb, 1), undefined, undefined);
  assert.equal(wb.removedSheetNames, undefined);
  assert.equal(formulaAt(wb.sheets[0], 'A6'), '#REF!*2');
  assert.deepEqual(wb.sheets, old);

  const after = picture(wb);
  for (let i = 0; i < 5; i++) history.undo();
  assert.equal(history.canUndo, false);
  assert.deepEqual(picture(wb), before);
  for (let i = 0; i < 2; i++) history.redo();
  assert.equal(wb.sheets[1], added, 'redo puts the same sheet back');
  assert.equal(valueAt(added, 'A1') !== undefined, true);
  for (let i = 0; i < 3; i++) history.redo();
  assert.deepEqual(picture(wb), after);

  // At the ends, with a name.
  history.apply(insertSheet(wb, 0, { name: 'First' }), undefined, undefined);
  history.apply(insertSheet(wb, wb.sheets.length, { name: 'Last' }), undefined, undefined);
  assert.deepEqual(
    wb.sheets.map((sheet) => sheet.name),
    ['First', 'Sheet1', 'My Data', 'Sheet3', 'Last'],
  );
  assert.deepEqual(
    wb.sheets.map((sheet) => sheet.sheetId),
    [4, 1, 2, 3, 5],
  );
  assert.equal(wb.activeSheet, 2);
});

test('insertSheet: undoing leaves formulas that named the sheet before it existed as they were', async () => {
  const wb = await workbook();
  const history = new History(wb);
  history.apply(setCellInput(wb, 0, 7, 0, '=Later!A1+1'), undefined, undefined);
  const before = picture(wb);
  history.apply(insertSheet(wb, 3, { name: 'Later' }), undefined, undefined);
  history.undo();
  assert.deepEqual(picture(wb), before);
  assert.equal(formulaAt(wb.sheets[0], 'A8'), 'Later!A1+1');
});

// ---------------------------------------------------------------------------
// Refusals

test('sheet operations that are refused leave the workbook as it was', async () => {
  const wb = await workbook();
  wb.activeSheet = 1;
  const before = picture(wb);
  const refused = (run: () => unknown, message: RegExp) => {
    assert.throws(run, (error: unknown) => error instanceof RangeError && message.test(error.message), String(message));
    assert.deepEqual(picture(wb), before);
    assert.equal(wb.sheetsModified, undefined);
  };
  const both = (op: Op, command: () => unknown, message: RegExp) => {
    refused(() => applyOp(wb, op), message);
    refused(command, message);
    refused(() => new History(wb).apply([op], undefined, undefined), message);
  };

  // Names a sheet cannot have.
  for (const name of ['', '   ', 'x'.repeat(32), 'a[b', 'a]b', 'a:b', 'a*b', 'a?b', 'a/b', 'a\\b', "'a", "a'", 'History', 'HISTORY', 'a\u0001b', 'a\ud800']) {
    both({ type: 'renameSheet', sheet: 0, name }, () => renameSheet(wb, 0, name), /^A sheet name has 1 to 31 characters/);
    both({ type: 'insertSheet', sheet: 0, name }, () => insertSheet(wb, 0, { name }), /^A sheet name has 1 to 31 characters/);
    assert.match(canRenameSheet(wb, 0, name)!, /^A sheet name/);
    assert.match(canInsertSheet(wb, 0, name)!, /^A sheet name/);
  }
  assert.equal(canRenameSheet(wb, 0, 'x'.repeat(31)), undefined);
  assert.equal(canRenameSheet(wb, 0, "it's"), undefined);
  // A name another sheet has, whatever the letter case.
  for (const name of ['My Data', 'my data', 'MY DATA', 'sheet3']) {
    both({ type: 'renameSheet', sheet: 0, name }, () => renameSheet(wb, 0, name), /another sheet has that name/);
    both({ type: 'insertSheet', sheet: 1, name }, () => insertSheet(wb, 1, { name }), /another sheet has that name/);
  }
  assert.equal(canRenameSheet(wb, 0, 'SHEET1'), undefined, 'its own name in other case is fine');
  assert.match(canInsertSheet(wb, 0, 'SHEET1')!, /another sheet has that name/);
  assert.equal(canInsertSheet(wb, 0), undefined);

  // Indexes.
  for (const index of [-1, 3, 1.5, NaN]) {
    both({ type: 'renameSheet', sheet: index, name: 'X' }, () => renameSheet(wb, index, 'X'), /^No sheet at index/);
    both({ type: 'deleteSheet', sheet: index }, () => deleteSheet(wb, index), /^No sheet at index/);
  }
  for (const index of [-1, 4, 1.5, NaN]) both({ type: 'insertSheet', sheet: index, name: 'X' }, () => insertSheet(wb, index, { name: 'X' }), /^Cannot insert a sheet at index/);

  // The last visible sheet.
  wb.sheets[0].state = 'hidden';
  wb.sheets[2].state = 'veryHidden';
  both({ type: 'deleteSheet', sheet: 1 }, () => deleteSheet(wb, 1), /^Cannot delete the last visible sheet$/);
  assert.equal(canDeleteSheet(wb, 1), 'Cannot delete the last visible sheet');
  assert.equal(canDeleteSheet(wb, 0), undefined, 'a hidden sheet may go while a visible one stays');
  wb.sheets[0].state = 'visible';
  wb.sheets[2].state = 'visible';

  // A sheet that is not a worksheet.
  wb.sheets[2].kind = 'chartsheet';
  both({ type: 'deleteSheet', sheet: 2 }, () => deleteSheet(wb, 2), /^Cannot delete a chartsheet: not supported yet$/);
  both({ type: 'renameSheet', sheet: 2, name: 'X' }, () => renameSheet(wb, 2, 'X'), /^Cannot rename a chartsheet: not supported yet$/);
  wb.sheets[2].kind = 'worksheet';

  // Charts and pivot tables name sheets in parts that are not rewritten: no sheet of the workbook goes or changes its name.
  for (const kind of ['chart', 'pivotTable']) {
    wb.sheets[2].anchored = ['drawing', kind];
    both({ type: 'deleteSheet', sheet: 0 }, () => deleteSheet(wb, 0), new RegExp(`^Cannot delete a sheet in a workbook with ${kind}: not supported yet$`));
    both({ type: 'renameSheet', sheet: 0, name: 'X' }, () => renameSheet(wb, 0, 'X'), new RegExp(`^Cannot rename a sheet in a workbook with ${kind}: not supported yet$`));
    assert.match(canDeleteSheet(wb, 0)!, /workbook with/);
    assert.match(canRenameSheet(wb, 0)!, /workbook with/);
    assert.equal(canInsertSheet(wb, 0, 'X'), undefined, 'a sheet may be added');
  }
  // A table is named by formulas elsewhere.
  wb.sheets[2].anchored = ['table'];
  both({ type: 'deleteSheet', sheet: 2 }, () => deleteSheet(wb, 2), /^Cannot delete a sheet with table: not supported yet$/);
  assert.equal(canDeleteSheet(wb, 0), undefined);
  assert.equal(canRenameSheet(wb, 2, 'X'), undefined);
  delete wb.sheets[2].anchored;
  // Pictures and comments are no obstacle.
  wb.sheets[2].anchored = ['comments', 'drawing', 'vmlDrawing'];
  assert.equal(canDeleteSheet(wb, 2), undefined);
  assert.equal(canRenameSheet(wb, 2, 'X'), undefined);
  delete wb.sheets[2].anchored;
  assert.deepEqual(picture(wb), before);
});

test('sheet operations are refused when a formula that may name the sheet cannot be read', async () => {
  const wb = await workbook();
  // What a file may hold; typed formulas are checked before they get here.
  const cell = getCell(wb.sheets[1], 1, 1)!;
  wb.sheets[1].rows.get(1)!.cells[1] = { ...cell, formula: { text: 'A1+"x"&"open+Sheet1!A1', type: 'normal' } };
  const before = picture(wb);
  for (const op of [{ type: 'renameSheet', sheet: 0, name: 'X' } as Op, { type: 'deleteSheet', sheet: 0 } as Op, { type: 'renameSheet', sheet: 2, name: 'X' } as Op]) {
    assert.throws(
      () => applyOp(wb, op),
      (error: unknown) => error instanceof RangeError && /^Cannot (rename|delete) the sheet: the formula of My Data!B2 cannot be read$/.test(error.message),
    );
    assert.deepEqual(picture(wb), before);
    assert.equal(wb.sheetsModified, undefined);
    const history = new History(wb);
    assert.throws(() => history.apply([op], undefined, undefined), RangeError);
    assert.equal(history.canUndo, false);
  }
  // The sheet that holds the formula can go: its formulas go with it.
  applyOp(wb, { type: 'deleteSheet', sheet: 1 });
  assert.equal(wb.sheets.length, 2);

  // A defined name is checked the same way.
  const named = await workbook();
  named.definedNames = named.definedNames.concat([{ name: 'Broken', value: "'Sheet1!$A$1", hidden: false }]);
  assert.throws(() => applyOp(named, { type: 'renameSheet', sheet: 0, name: 'X' }), /the formula of the name Broken cannot be read/);
  // An unclosed quote with no sheet named after it is no obstacle.
  const harmless = await workbook();
  harmless.sheets[1].rows.get(1)!.cells[1] = { ...cell, formula: { text: 'Sheet1!A1+"open', type: 'normal' } };
  applyOp(harmless, { type: 'renameSheet', sheet: 0, name: 'X' });
  assert.equal(formulaAt(harmless.sheets[1], 'B2'), 'X!A1+"open');
});

// ---------------------------------------------------------------------------
// History

test('sheet operations: a failing step in a group is rolled back; weights bound the undo stack', async () => {
  const wb = await workbook();
  const before = picture(wb);
  const history = new History(wb);
  assert.throws(() => history.apply([{ type: 'renameSheet', sheet: 0, name: 'One' }, { type: 'insertSheet', sheet: 1, name: 'Two' }, { type: 'deleteSheet', sheet: 9 }], undefined, undefined), RangeError);
  assert.deepEqual(picture(wb), before);
  assert.equal(history.canUndo, false);

  // Several operations as one step.
  history.apply([{ type: 'insertSheet', sheet: 3, name: 'Two' }, { type: 'renameSheet', sheet: 0, name: 'One' }, { type: 'deleteSheet', sheet: 1 }], undefined, undefined);
  assert.deepEqual(
    wb.sheets.map((sheet) => sheet.name),
    ['One', 'Sheet3', 'Two'],
  );
  history.undo();
  assert.deepEqual(picture(wb), before);
  history.redo();
  assert.deepEqual(
    wb.sheets.map((sheet) => sheet.name),
    ['One', 'Sheet3', 'Two'],
  );
  history.undo();

  // The inverse carries what it keeps alive: the cells it rewrote, the rows of a deleted sheet.
  const renamed = applyOp(wb, { type: 'renameSheet', sheet: 0, name: 'One' });
  assert.equal(renamed.type, 'restoreSheets');
  assert.equal(renamed.type === 'restoreSheets' && renamed.weight, 10, '8 cells and 2 names');
  applyOp(wb, renamed);
  const deleted = applyOp(wb, { type: 'deleteSheet', sheet: 2 });
  assert.equal(deleted.type === 'restoreSheets' && deleted.weight, 5, 'the 5 rows of the sheet');
  applyOp(wb, deleted);
  assert.deepEqual(picture(wb), before);
});

test('renameSheet: 100,000 formulas are rewritten and restored in reasonable time', async () => {
  const wb = await workbook();
  const sheet = wb.sheets[1];
  const rows = 100000;
  for (let r = 10; r < 10 + rows; r++) {
    sheet.rows.set(r, { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [{ kind: 'number', value: 0, style: 0, formula: { text: `Sheet1!A${r}+SUM(Sheet1!$A$1:$B$2)+A${r}`, type: 'normal' } }, { kind: 'number', value: 0, style: 0, formula: { text: `A${r}*2`, type: 'normal' } }] });
  }
  sheet.rowCount = 10 + rows;
  const history = new History(wb);
  const start = Date.now();
  history.apply(renameSheet(wb, 0, 'A longer name'), undefined, undefined);
  assert.equal(sheet.rows.get(5000)!.cells[0].formula!.text, "'A longer name'!A5000+SUM('A longer name'!$A$1:$B$2)+A5000");
  assert.equal(sheet.rows.get(5000)!.cells[1].formula!.text, 'A5000*2');
  history.undo();
  assert.equal(sheet.rows.get(5000)!.cells[0].formula!.text, 'Sheet1!A5000+SUM(Sheet1!$A$1:$B$2)+A5000');
  history.redo();
  history.apply(deleteSheet(wb, 0), undefined, undefined);
  assert.equal(sheet.rows.get(5000)!.cells[0].formula!.text, '#REF!+SUM(#REF!)+A5000');
  history.undo();
  history.undo();
  assert.equal(sheet.rows.get(100009)!.cells[0].formula!.text, 'Sheet1!A100009+SUM(Sheet1!$A$1:$B$2)+A100009');
  // Generous: a few hundred milliseconds are expected; this only catches work that is not linear.
  assert.ok(Date.now() - start < 30000, `took ${Date.now() - start} ms`);
});

// ---------------------------------------------------------------------------
// Random sequences

function mulberry(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

test('random sequences of sheet and cell operations: undo all gives the start, redo all the end', async () => {
  const pool = ['Sheet1', 'My Data', 'Sheet3', 'Data', "Bob's", 'A1', 'x y', 'Özet', 'sheet2', 'Sheet2', 'R1C1', 'Q_1', 'İ', 'i'];
  for (let seed = 1; seed <= 40; seed++) {
    const random = mulberry(seed);
    const pick = (n: number) => Math.floor(random() * n);
    const wb = await workbook();
    const identities = wb.sheets.slice();
    const before = picture(wb);
    const history = new History(wb);
    const engine = new CalcEngine(wb);
    const steps: unknown[] = [];
    let applied = 0;
    for (let step = 0; step < 60; step++) {
      const count = wb.sheets.length;
      const index = pick(count);
      const kind = pick(10);
      let ops: Op[];
      try {
        if (kind === 0) ops = insertSheet(wb, pick(count + 1), random() < 0.5 ? { locale: random() < 0.5 ? 'tr' : 'en' } : { name: pool[pick(pool.length)] });
        else if (kind === 1) ops = deleteSheet(wb, index);
        else if (kind <= 3) ops = renameSheet(wb, index, pool[pick(pool.length)]);
        else if (kind === 4) ops = [{ type: 'insert', sheet: index, axis: random() < 0.5 ? 'row' : 'col', at: pick(4), count: 1 + pick(2) }];
        else if (kind === 5) ops = [{ type: 'delete', sheet: index, axis: random() < 0.5 ? 'row' : 'col', at: pick(4), count: 1 }];
        else if (kind === 6) ops = setCellInput(wb, index, pick(6), pick(4), String(pick(100)));
        else {
          // A formula that names sheets that exist, existed or may come to exist, by the names of the pool or of the workbook.
          const name = random() < 0.7 ? wb.sheets[pick(count)].name : pool[pick(pool.length)];
          const other = wb.sheets[pick(count)].name;
          const text = kind === 7 ? `=${quoteSheetName(name)}!A${1 + pick(5)}+1` : kind === 8 ? `=SUM(${quoteSheetName(name)}!A1:B2)+${quoteSheetName(other)}!$B$${1 + pick(3)}&"${name}!A1"` : `=SUM('${name.replace(/'/g, "''")}:${other.replace(/'/g, "''")}'!A1)`;
          ops = setCellInput(wb, index, pick(6), pick(4), text);
        }
        if (ops.length === 0) continue;
        history.apply(ops, undefined, undefined);
      } catch (error) {
        // Refused (a name that is taken, the last visible sheet, an array cut): nothing changed.
        assert.ok(error instanceof RangeError, String(error));
        continue;
      }
      applied++;
      steps.push(picture(wb));
      // The model stays sound at every step.
      assert.ok(wb.activeSheet >= 0 && wb.activeSheet < wb.sheets.length && wb.sheets[wb.activeSheet].state === 'visible', `seed ${seed}: active sheet`);
      assert.equal(new Set(wb.sheets.map((sheet) => sheet.name.toLowerCase())).size, wb.sheets.length, `seed ${seed}: names are unique`);
      assert.equal(new Set(wb.sheets.map((sheet) => sheet.sheetId)).size, wb.sheets.length, `seed ${seed}: ids are unique`);
      for (const name of wb.definedNames) assert.ok(name.localSheet === undefined || (name.localSheet >= 0 && name.localSheet < wb.sheets.length), `seed ${seed}: local sheet`);
      for (const sheet of wb.sheets) assert.equal(sheet.nameInFile !== undefined && sheet.part === '', false, `seed ${seed}: only sheets of the file have a name in it`);
      // The engine, built again, copes with whatever the list of sheets is now.
      if (step % 7 === 0) recalculate(engine);
    }
    assert.ok(applied > 10, `seed ${seed}: ${applied} steps applied`);
    const after = picture(wb);
    for (let i = steps.length - 2; i >= 0; i--) {
      assert.ok(history.undo());
      assert.deepEqual(picture(wb), steps[i], `seed ${seed}: undo to step ${i}`);
    }
    assert.ok(history.undo());
    assert.equal(history.canUndo, false);
    assert.deepEqual(picture(wb), before, `seed ${seed}: undo all`);
    assert.ok(wb.sheets.length === identities.length && wb.sheets.every((sheet, i) => sheet === identities[i]), `seed ${seed}: the sheets are the same objects`);
    recalculate(engine);
    for (let i = 0; i < steps.length; i++) {
      assert.ok(history.redo());
      assert.deepEqual(picture(wb), steps[i], `seed ${seed}: redo to step ${i}`);
    }
    assert.equal(history.canRedo, false);
    assert.deepEqual(picture(wb), after, `seed ${seed}: redo all`);
    recalculate(engine);
  }
});
