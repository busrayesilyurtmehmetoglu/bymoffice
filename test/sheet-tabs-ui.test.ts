// What the sheet tabs of the workbook view decide without a DOM: the sheet shown once one
// is gone, and the words of every refusal of the operations on sheets, in both languages.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine, getCell, open, writeXlsx, type Op, type Sheet, type Workbook } from '../src/index';
import { setCellInput } from '../src/ops/commands';
import { canDeleteSheet, canRenameSheet, deleteSheet, insertSheet, renameSheet } from '../src/ops/sheets';
import { History } from '../src/ops/ops';
import { allSheets, recalculateForSheets, sheetErrorText, sheetToShow, type SheetCellAddress } from '../src/ui/sheet-tabs';
import { STRINGS_EN, STRINGS_TR } from '../src/ui/strings';
import { buildXlsx, worksheetXml } from './fixtures';

const visible = { state: 'visible' as const };
const hidden = { state: 'hidden' as const };

async function workbook(): Promise<Workbook> {
  const file = await open(
    await buildXlsx({
      sheets: [
        { name: 'One', xml: worksheetXml('<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>') },
        { name: 'Two', xml: worksheetXml('<sheetData><row r="1"><c r="A1"><f>One!A1</f><v>1</v></c></row></sheetData>') },
      ],
    }),
  );
  return file.workbook();
}

const messageOf = (run: () => unknown): string => {
  try {
    run();
  } catch (error) {
    assert.ok(error instanceof RangeError, 'a refusal is a RangeError');
    return (error as Error).message;
  }
  throw new assert.AssertionError({ message: 'the operation was not refused' });
};

test('sheetToShow: the next sheet with a tab, else the one before, else any', () => {
  assert.equal(sheetToShow([visible, visible, visible], 1), 1);
  assert.equal(sheetToShow([visible, hidden, visible], 1), 2);
  assert.equal(sheetToShow([visible, hidden, hidden], 1), 0);
  // Past the end: the last sheet was deleted.
  assert.equal(sheetToShow([visible, visible], 2), 1);
  assert.equal(sheetToShow([visible, hidden], 2), 0);
  assert.equal(sheetToShow([hidden, hidden], 1), 1);
  assert.equal(sheetToShow([hidden, hidden], 9), 1);
  assert.equal(sheetToShow([], 0), -1);
});

test('sheetToShow: an index that is none falls back to the first sheet with a tab', () => {
  assert.equal(sheetToShow([hidden, visible], -5), 1);
  assert.equal(sheetToShow([hidden, visible], Number.NaN), 1);
  assert.equal(sheetToShow([hidden, visible], 0.5), 1);
  assert.equal(sheetToShow([visible, visible], Number.POSITIVE_INFINITY), 0);
});

test('allSheets lists every index', () => {
  assert.deepEqual(allSheets(0), []);
  assert.deepEqual(allSheets(3), [0, 1, 2]);
});

test('every refusal of a sheet name has words of its own in both languages', async () => {
  const book = await workbook();
  const bad = ['', '   ', 'a/b', 'a[b]', "'quoted", 'History', 'x'.repeat(32), 'a\u0001b'];
  for (const name of bad) {
    const message = messageOf(() => renameSheet(book, 0, name));
    assert.equal(sheetErrorText(message, STRINGS_EN), STRINGS_EN.badSheetName, JSON.stringify(name));
    assert.equal(sheetErrorText(message, STRINGS_TR), STRINGS_TR.badSheetName, JSON.stringify(name));
    assert.equal(canRenameSheet(book, 0, name), message);
    assert.equal(sheetErrorText(messageOf(() => insertSheet(book, 0, { name })), STRINGS_EN), STRINGS_EN.badSheetName);
  }
  const taken = messageOf(() => renameSheet(book, 0, 'TWO'));
  assert.equal(sheetErrorText(taken, STRINGS_EN), 'Another sheet is already named "TWO"');
  assert.equal(sheetErrorText(taken, STRINGS_TR), '"TWO" adında başka bir sayfa var');
  assert.equal(sheetErrorText(messageOf(() => insertSheet(book, 2, { name: 'one' })), STRINGS_EN), 'Another sheet is already named "one"');
});

test('a name with a quote, a colon-free line of anything, is said as it was typed', () => {
  const message = 'Cannot name the sheet "a "b" c": another sheet has that name';
  assert.equal(sheetErrorText(message, STRINGS_EN), 'Another sheet is already named "a "b" c"');
});

test('refusals to delete and rename a sheet have words of their own in both languages', async () => {
  const book = await workbook();
  // The last visible sheet.
  book.sheets[1].state = 'hidden';
  const last = messageOf(() => deleteSheet(book, 0));
  assert.equal(canDeleteSheet(book, 0), last);
  assert.equal(sheetErrorText(last, STRINGS_EN), STRINGS_EN.cannotDeleteLastSheet);
  assert.equal(sheetErrorText(last, STRINGS_TR), STRINGS_TR.cannotDeleteLastSheet);
  book.sheets[1].state = 'visible';

  // A sheet with a table cannot go; a chart or a pivot table anywhere keeps every sheet and name.
  book.sheets[0].anchored = ['table'];
  const table = messageOf(() => deleteSheet(book, 0));
  assert.equal(sheetErrorText(table, STRINGS_EN), 'A sheet with a table cannot be deleted yet');
  assert.equal(sheetErrorText(table, STRINGS_TR), 'Tablo içeren sayfa henüz silinemiyor');
  assert.equal(canRenameSheet(book, 0), undefined);

  book.sheets[0].anchored = ['chart'];
  const chartDelete = messageOf(() => deleteSheet(book, 1));
  const chartRename = messageOf(() => renameSheet(book, 1, 'Other'));
  assert.equal(sheetErrorText(chartDelete, STRINGS_EN), 'Sheets cannot be deleted yet in a workbook with a chart');
  assert.equal(sheetErrorText(chartRename, STRINGS_EN), 'Sheets cannot be renamed yet in a workbook with a chart');
  assert.equal(sheetErrorText(chartRename, STRINGS_TR), 'Grafik içeren çalışma kitabında sayfalar henüz yeniden adlandırılamıyor');
  book.sheets[0].anchored = ['pivotTable'];
  assert.equal(sheetErrorText(messageOf(() => deleteSheet(book, 1)), STRINGS_TR), 'Pivot tablo içeren çalışma kitabında sayfalar henüz silinemiyor');
  delete book.sheets[0].anchored;

  // A sheet that is no worksheet.
  book.sheets[1].kind = 'chartsheet';
  const kindRename = messageOf(() => renameSheet(book, 1, 'Other'));
  const kindDelete = messageOf(() => deleteSheet(book, 1));
  assert.equal(sheetErrorText(kindRename, STRINGS_EN), 'This kind of sheet cannot be renamed yet');
  assert.equal(sheetErrorText(kindDelete, STRINGS_EN), 'This kind of sheet cannot be deleted yet');
  assert.equal(sheetErrorText(kindDelete, STRINGS_TR), 'Bu tür sayfalar henüz silinemiyor');
});

test('a formula that cannot be read refuses when the operation is applied, and is said with its cell', async () => {
  const book = await workbook();
  const cell = book.sheets[1].rows.get(0)!.cells[0];
  // A text that does not end: whether what follows names a sheet cannot be told.
  book.sheets[1].rows.get(0)!.cells[0] = { ...cell, formula: { ...cell.formula!, text: 'A1+"x"&"open+One!A1' } };
  const history = new History(book);
  for (const action of ['rename', 'delete'] as const) {
    const ops = action === 'rename' ? renameSheet(book, 0, 'Uno') : deleteSheet(book, 0);
    const message = messageOf(() => history.apply(ops, undefined, undefined));
    const en = sheetErrorText(message, STRINGS_EN);
    const tr = sheetErrorText(message, STRINGS_TR);
    assert.ok(en && en.indexOf('Two!A1') >= 0 && en.indexOf(action === 'rename' ? 'renamed' : 'deleted') >= 0, en);
    assert.ok(tr && tr.indexOf('Two!A1') >= 0, tr);
    // Nothing changed.
    assert.deepEqual(book.sheets.map((sheet) => sheet.name), ['One', 'Two']);
    assert.equal(history.canUndo, false);
  }
});

test('messages about anything else are not taken for one about sheets', () => {
  for (const message of ['', 'The formula cannot be read', 'Cannot insert rows in a sheet with tables: not supported yet', 'Cannot sort in a sheet with comments: not supported yet', 'No sheet at index 4']) {
    assert.equal(sheetErrorText(message, STRINGS_EN), undefined, message);
  }
});

// ---------------------------------------------------------------- formulas after the list of sheets changed

const CHAIN = 300;

/**
 * "Data": A1 = 1, B1 = A1*2 with 999 as the value of the file, C1 = RAND() with 0.5.
 * "Chain": A1 = Data!A1, A2 = A1+1 ... down to row CHAIN, each with the value it has for Data!A1 = 1.
 * "Empty": nothing.
 */
async function chainFile() {
  const rows: string[] = ['<row r="1"><c r="A1"><f>Data!A1</f><v>1</v></c></row>'];
  for (let r = 2; r <= CHAIN; r++) rows.push(`<row r="${r}"><c r="A${r}"><f>A${r - 1}+1</f><v>${r}</v></c></row>`);
  const bytes = await buildXlsx({
    sheets: [
      { name: 'Data', xml: worksheetXml('<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>999</v></c><c r="C1"><f>RAND()</f><v>0.5</v></c></row></sheetData>') },
      { name: 'Chain', xml: worksheetXml(`<sheetData>${rows.join('')}</sheetData>`) },
      { name: 'Empty', xml: worksheetXml('<sheetData/>') },
    ],
  });
  const file = await open(bytes);
  return { bytes, file, book: await file.workbook() };
}

const valueAt = (sheet: Sheet, row: number, col: number) => getCell(sheet, row, col)!.value;
const named = (book: Workbook, name: string) => book.sheets.find((sheet) => sheet.name === name)!;

/** The cells a deletion rewrote, from its inverse, by the indexes of a list of sheets. */
function rewritten(inverse: Op[], list: Sheet[]): SheetCellAddress[] {
  const out: SheetCellAddress[] = [];
  for (const op of inverse) {
    if (op.type !== 'restoreSheets') continue;
    for (const entry of op.state.cells) for (const change of entry.cells) out.push({ sheet: list.indexOf(entry.sheet), row: change.row, col: change.col });
  }
  return out;
}

function finish(engine: CalcEngine): void {
  for (let i = 0; engine.incomplete && i < 10000; i++) engine.recalculate([]);
  assert.equal(engine.incomplete, false);
}

test('a calculation under way when a sheet is renamed is not lost: everything is calculated', async () => {
  for (const careful of [false, true]) {
    const { book } = await chainFile();
    const history = new History(book);
    const engine = new CalcEngine(book, { budget: { work: 40, milliseconds: 0 } });
    history.apply(setCellInput(book, 0, 0, 0, '100', { locale: 'en-US' }), undefined, undefined);
    engine.recalculate([{ sheet: 0, row: 0, col: 0 }]);
    assert.equal(engine.incomplete, true, 'the budget is too small for the chain');
    history.apply(renameSheet(book, 0, 'Source'), undefined, undefined);
    if (careful) recalculateForSheets(engine, { kind: 'names' });
    else {
      // What the view did before: the graph is forgotten, and what was left to calculate with it.
      engine.invalidate();
      engine.recalculate([]);
    }
    finish(engine);
    const last = valueAt(named(book, 'Chain'), CHAIN - 1, 0);
    if (careful) assert.equal(last, 100 + CHAIN - 1);
    else assert.notEqual(last, 100 + CHAIN - 1, 'the old way leaves the end of the chain with the value of before');
  }
});

test('a sheet renamed, or an empty one added and taken away again, calculates nothing and marks no sheet', async () => {
  const { book, bytes, file } = await chainFile();
  const history = new History(book);
  const engine = new CalcEngine(book);
  // The graph is there, as after an edit.
  engine.recalculate([]);
  const random = valueAt(book.sheets[0], 0, 2);
  for (const sheet of book.sheets) delete sheet.modified;

  history.apply(insertSheet(book, 1, { name: 'New' }), undefined, undefined);
  recalculateForSheets(engine, { kind: 'names' });
  history.apply(renameSheet(book, 3, 'Blank'), undefined, undefined);
  recalculateForSheets(engine, { kind: 'names' });
  assert.equal(engine.incomplete, false);
  assert.equal(valueAt(named(book, 'Data'), 0, 1), 999, 'the value of the file stays');
  assert.equal(valueAt(named(book, 'Data'), 0, 2), random, 'RAND() is not calculated again');
  assert.deepEqual(book.sheets.filter((sheet) => sheet.modified).map((sheet) => sheet.name), ['New']);

  history.undo();
  recalculateForSheets(engine, { kind: 'names' });
  history.undo();
  recalculateForSheets(engine, { kind: 'names' });
  assert.equal(valueAt(named(book, 'Data'), 0, 1), 999);
  assert.deepEqual(book.sheets.filter((sheet) => sheet.modified), []);
  // Saved, the file is what it was.
  await writeXlsx(file.package!, book);
  assert.deepEqual(await file.package!.save(), bytes);

  // The engine works on: an edit is followed through the new graph.
  history.apply(setCellInput(book, 0, 0, 0, '5', { locale: 'en-US' }), undefined, undefined);
  engine.recalculate([{ sheet: 0, row: 0, col: 0 }]);
  assert.equal(valueAt(named(book, 'Data'), 0, 1), 10);
  assert.equal(valueAt(named(book, 'Chain'), 1, 0), 6);
});

test('a sheet deleted: what read it and what reads that is calculated, the rest keeps the values of the file', async () => {
  const { book } = await chainFile();
  const history = new History(book);
  const engine = new CalcEngine(book);
  engine.recalculate([]);

  // "Empty" has no readers: nothing to calculate.
  let before = book.sheets;
  let inverse = history.apply(deleteSheet(book, 2), undefined, undefined);
  assert.deepEqual(rewritten(inverse, before), []);
  recalculateForSheets(engine, { kind: 'gone', cells: [] });
  assert.equal(valueAt(named(book, 'Data'), 0, 1), 999);
  history.undo();
  recalculateForSheets(engine, { kind: 'back', cells: [] });

  // "Data" is read by Chain!A1, and that by the whole chain.
  before = book.sheets;
  inverse = history.apply(deleteSheet(book, 0), undefined, undefined);
  const cells = rewritten(inverse, before);
  assert.deepEqual(cells, [{ sheet: 1, row: 0, col: 0 }]);
  recalculateForSheets(engine, { kind: 'gone', cells });
  finish(engine);
  const chain = named(book, 'Chain');
  assert.equal(getCell(chain, 0, 0)!.formula!.text, '#REF!');
  assert.equal(valueAt(chain, 0, 0), '#REF!');
  assert.equal(valueAt(chain, CHAIN - 1, 0), '#REF!');

  // Back: by the indexes of now.
  history.undo();
  recalculateForSheets(engine, { kind: 'back', cells: rewritten(inverse, book.sheets) });
  finish(engine);
  assert.equal(valueAt(chain, 0, 0), 1);
  assert.equal(valueAt(chain, CHAIN - 1, 0), CHAIN);
  assert.equal(valueAt(named(book, 'Data'), 0, 1), 999, 'the value of the file is still there');
  // A formula that gives another value each time is calculated with every change of a cell, as this is one.
  assert.notEqual(valueAt(named(book, 'Data'), 0, 2), undefined);
});

test('a sheet deleted without a graph of before, or with cells unknown, calculates everything', async () => {
  for (const known of [true, false]) {
    const { book } = await chainFile();
    const history = new History(book);
    const engine = new CalcEngine(book);
    const before = book.sheets;
    const inverse = history.apply(deleteSheet(book, 0), undefined, undefined);
    if (known) assert.equal(engine.graphReady, false);
    else engine.recalculate([]);
    recalculateForSheets(engine, { kind: 'gone', cells: known ? rewritten(inverse, before) : undefined });
    finish(engine);
    assert.equal(valueAt(named(book, 'Chain'), CHAIN - 1, 0), '#REF!');
  }
});

test('the refusals for form controls and for names that differ only by case have words in both languages', () => {
  for (const action of ['rename', 'delete'] as const) {
    const message = `Cannot ${action} a sheet in a workbook with form controls: not supported yet`;
    assert.equal(sheetErrorText(message, STRINGS_EN), STRINGS_EN.sheetBlockedBy(action, 'form controls'));
    assert.match(sheetErrorText(message, STRINGS_TR) || '', /^Form denetimi içeren çalışma kitabında/);
  }
  const clash = 'Cannot change the sheets: "IL" and "İL" differ only by case';
  assert.equal(sheetErrorText(clash, STRINGS_EN), 'The sheets of this workbook cannot be changed: "IL" and "İL" differ only by case');
  assert.equal(sheetErrorText(clash, STRINGS_TR), STRINGS_TR.sheetNamesClash('IL', 'İL'));
});
