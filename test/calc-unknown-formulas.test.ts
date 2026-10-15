import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine } from '../src/calc/engine';
import { createBlank } from '../src/formats/blank';
import { open } from '../src/index';
import { getCell, type Cell, type Workbook } from '../src/model/sheet/types';
import { setCellInput } from '../src/ops/commands';
import { History } from '../src/ops/ops';
import { insertSheet, moveSheet, deleteSheet, renameSheet } from '../src/ops/sheets';

// A program that uses the operations and the engine without a view does not tell the engine
// which cells it wrote: a recalculation of everything has to find them by itself, and the
// engine has to notice by itself that the list of sheets is no longer the one it read.

async function blankBook(): Promise<{ workbook: Workbook; history: History; save: () => Promise<Uint8Array> }> {
  const file = await open(await createBlank('xlsx', { sheetName: 'Data' }));
  const workbook = await file.workbook();
  return { workbook, history: new History(workbook), save: () => file.save() };
}

function type(workbook: Workbook, history: History, sheet: number, row: number, col: number, text: string): void {
  history.apply(setCellInput(workbook, sheet, row, col, text), undefined, undefined);
}

const cellAt = (workbook: Workbook, sheet: number, row: number, col: number): Cell | undefined => getCell(workbook.sheets[sheet], row, col);
const valueAt = (workbook: Workbook, sheet: number, row: number, col: number) => cellAt(workbook, sheet, row, col)?.value;

test('a formula written after the engine calculated is found by a recalculation of everything', async () => {
  const { workbook, history } = await blankBook();
  type(workbook, history, 0, 0, 0, '2');
  type(workbook, history, 0, 1, 0, '=A1*10');
  const engine = new CalcEngine(workbook);
  engine.recalculate(undefined, { overwrite: true });
  assert.equal(valueAt(workbook, 0, 1, 0), 20);

  type(workbook, history, 0, 2, 0, '=A1+A2');
  assert.deepEqual(engine.uncalculated(), [{ sheet: 0, row: 2, col: 0 }], 'the cell without a value is listed before anything is calculated');
  engine.recalculate(undefined, { overwrite: true });
  assert.equal(valueAt(workbook, 0, 2, 0), 22);
  assert.deepEqual(engine.uncalculated(), []);
});

test('without overwrite as well: a formula that has no value is always calculated', async () => {
  const { workbook, history } = await blankBook();
  type(workbook, history, 0, 0, 0, '2');
  const engine = new CalcEngine(workbook);
  engine.recalculate();
  type(workbook, history, 0, 1, 0, '=A1*10');
  type(workbook, history, 0, 2, 0, '=A2+1');
  engine.recalculate();
  assert.equal(valueAt(workbook, 0, 1, 0), 20);
  assert.equal(valueAt(workbook, 0, 2, 0), 21);
});

test('a formula replaced behind the back of the engine is read again, and what reads it follows', async () => {
  const { workbook, history } = await blankBook();
  type(workbook, history, 0, 0, 0, '2');
  type(workbook, history, 0, 1, 0, '=A1*10');
  type(workbook, history, 0, 2, 0, '=A2+1');
  const engine = new CalcEngine(workbook);
  engine.recalculate(undefined, { overwrite: true });
  assert.equal(valueAt(workbook, 0, 2, 0), 21);

  type(workbook, history, 0, 1, 0, '=A1*100');
  engine.recalculate();
  assert.equal(valueAt(workbook, 0, 1, 0), 200);
  assert.equal(valueAt(workbook, 0, 2, 0), 201);

  // A formula that became a constant: its readers are calculated from the constant.
  type(workbook, history, 0, 1, 0, '7');
  engine.recalculate();
  assert.equal(valueAt(workbook, 0, 2, 0), 8);
});

test('a formula that cannot be parsed is not read again by every recalculation, and is listed', async () => {
  const { workbook, history } = await blankBook();
  type(workbook, history, 0, 0, 0, '1');
  const engine = new CalcEngine(workbook);
  engine.recalculate();
  // Placed directly: the input parser would not accept it.
  const row = workbook.sheets[0].rows.get(0)!;
  workbook.sheets[0].rows.set(0, { ...row, cells: [row.cells[0], { kind: 'blank', value: null, style: 0, formula: { text: 'SUM((', type: 'normal' } }] });
  engine.recalculate();
  assert.deepEqual(engine.unparsed, [{ sheet: 0, row: 0, col: 1 }]);
  engine.recalculate();
  assert.deepEqual(engine.unparsed, [{ sheet: 0, row: 0, col: 1 }]);
  assert.equal(valueAt(workbook, 0, 0, 1), null);
});

test('the values of the file stay when everything is recalculated and a new formula is found', async () => {
  const { workbook, history } = await blankBook();
  // A value of the file that this engine would not calculate: it is kept (see `mismatched`).
  const sheet = workbook.sheets[0];
  sheet.rows.set(0, { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [{ kind: 'number', value: 5, style: 0 }, { kind: 'number', value: 99, style: 0, formula: { text: 'A1+1', type: 'normal' } }] });
  sheet.rowCount = 1;
  sheet.colCount = 2;
  const engine = new CalcEngine(workbook);
  engine.recalculate();
  assert.equal(valueAt(workbook, 0, 0, 1), 99);
  type(workbook, history, 0, 1, 0, '=B1+1');
  engine.recalculate();
  assert.equal(valueAt(workbook, 0, 0, 1), 99, 'the value of the file is not replaced');
  assert.equal(valueAt(workbook, 0, 1, 0), 100);
  assert.deepEqual(engine.mismatched, [{ sheet: 0, row: 0, col: 1 }]);
});

test('after a sheet is inserted the engine reads the sheets again by itself', async () => {
  const { workbook, history } = await blankBook();
  type(workbook, history, 0, 0, 0, '3');
  const engine = new CalcEngine(workbook);
  engine.recalculate(undefined, { overwrite: true });

  history.apply(insertSheet(workbook, 0, { name: 'First' }), undefined, undefined);
  assert.equal(workbook.sheets[1].name, 'Data');
  type(workbook, history, 0, 0, 0, '=Data!A1*2');
  // Named, as a careful caller does: the engine must not look for it in the sheet that stood there.
  engine.recalculate([{ sheet: 0, row: 0, col: 0 }]);
  assert.equal(valueAt(workbook, 0, 0, 0), 6);

  history.apply(insertSheet(workbook, 2, { name: 'Last' }), undefined, undefined);
  type(workbook, history, 2, 4, 1, '=First!A1+Data!A1');
  engine.recalculate(undefined, { overwrite: true });
  assert.equal(valueAt(workbook, 2, 4, 1), 9);
  assert.deepEqual(engine.uncalculated(), []);
});

test('after sheets are moved, renamed or removed the engine reads the sheets again by itself', async () => {
  const { workbook, history } = await blankBook();
  history.apply(insertSheet(workbook, 1, { name: 'Two' }), undefined, undefined);
  history.apply(insertSheet(workbook, 2, { name: 'Three' }), undefined, undefined);
  type(workbook, history, 0, 0, 0, '4');
  type(workbook, history, 1, 0, 0, '=Data!A1+1');
  const engine = new CalcEngine(workbook);
  engine.recalculate(undefined, { overwrite: true });
  assert.equal(valueAt(workbook, 1, 0, 0), 5);

  history.apply(moveSheet(workbook, 1, 0), undefined, undefined);
  assert.equal(workbook.sheets[0].name, 'Two');
  type(workbook, history, 1, 0, 0, '10');
  engine.recalculate([{ sheet: 1, row: 0, col: 0 }]);
  assert.equal(valueAt(workbook, 0, 0, 0), 11);

  history.apply(renameSheet(workbook, 1, 'Numbers'), undefined, undefined);
  type(workbook, history, 2, 0, 0, '=Numbers!A1*3');
  engine.recalculate([{ sheet: 2, row: 0, col: 0 }]);
  assert.equal(valueAt(workbook, 2, 0, 0), 30);

  history.apply(deleteSheet(workbook, 0), undefined, undefined);
  assert.equal(workbook.sheets[1].name, 'Three');
  type(workbook, history, 0, 0, 0, '20');
  engine.recalculate([{ sheet: 0, row: 0, col: 0 }]);
  assert.equal(valueAt(workbook, 1, 0, 0), 60);
});

test('a sheet that formulas read is deleted: everything is calculated, whatever the engine is told', async () => {
  for (const told of ['nothing', 'another cell', 'all'] as const) {
    const { workbook, history } = await blankBook();
    history.apply(insertSheet(workbook, 1, { name: 'Two' }), undefined, undefined);
    history.apply(insertSheet(workbook, 2, { name: 'Three' }), undefined, undefined);
    type(workbook, history, 1, 0, 0, '5');
    type(workbook, history, 0, 0, 0, '=Two!A1+1');
    type(workbook, history, 0, 1, 0, '=A1*2');
    type(workbook, history, 2, 0, 0, '=Data!A2+1');
    const engine = new CalcEngine(workbook);
    engine.recalculate(undefined, { overwrite: true });
    assert.equal(valueAt(workbook, 2, 0, 0), 13);

    history.apply(deleteSheet(workbook, 1), undefined, undefined);
    assert.equal(cellAt(workbook, 0, 0, 0)!.formula!.text.indexOf('#REF!') >= 0, true, 'the formula that read the sheet was rewritten');
    if (told === 'nothing') {
      // A look at a formula (a preview) is enough for the engine to see the list: it says that there is work left.
      engine.evaluateText('=1+1', { sheet: 0, row: 5, col: 5 });
      assert.equal(engine.incomplete, true);
      engine.recalculate([]);
    } else if (told === 'another cell') engine.recalculate([{ sheet: 1, row: 7, col: 7 }]);
    else engine.recalculate();
    assert.equal(engine.incomplete, false);
    assert.equal(valueAt(workbook, 0, 0, 0), '#REF!', told);
    assert.equal(valueAt(workbook, 0, 1, 0), '#REF!', told);
    assert.equal(valueAt(workbook, 1, 0, 0), '#REF!', told);
  }
});

test('a sheet read through a defined name is deleted, while its own formulas read only it: everything is calculated', async () => {
  const { workbook, history } = await blankBook();
  history.apply(insertSheet(workbook, 1, { name: 'Two' }), undefined, undefined);
  type(workbook, history, 1, 0, 0, '5');
  // The formulas of the sheet that goes read that sheet only, a dynamic one (which watches its own cells) among them.
  type(workbook, history, 1, 1, 0, '=A1*2');
  type(workbook, history, 1, 2, 0, '=SEQUENCE(2)');
  workbook.definedNames.push({ name: 'Far', value: 'Two!$A$1' });
  type(workbook, history, 0, 0, 0, '=Far+1');
  const engine = new CalcEngine(workbook);
  engine.recalculate(undefined, { overwrite: true });
  assert.equal(valueAt(workbook, 0, 0, 0), 6);
  history.apply(deleteSheet(workbook, 1), undefined, undefined);
  engine.recalculate([]);
  assert.equal(engine.incomplete, false);
  assert.notEqual(valueAt(workbook, 0, 0, 0), 6, 'the formula that read the sheet through the name is calculated again');
});

test('a sheet nothing reads is deleted: nothing is calculated, and the values of the file stay', async () => {
  const { workbook, history } = await blankBook();
  history.apply(insertSheet(workbook, 1, { name: 'Two' }), undefined, undefined);
  const sheet = workbook.sheets[0];
  sheet.rows.set(0, { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [{ kind: 'number', value: 5, style: 0 }, { kind: 'number', value: 99, style: 0, formula: { text: 'A1+1', type: 'normal' } }] });
  sheet.rowCount = 1;
  sheet.colCount = 2;
  type(workbook, history, 1, 0, 0, '=Data!A1');
  const engine = new CalcEngine(workbook);
  engine.recalculate();
  history.apply(deleteSheet(workbook, 1), undefined, undefined);
  assert.deepEqual(engine.recalculate([]), []);
  assert.equal(engine.incomplete, false);
  assert.equal(valueAt(workbook, 0, 0, 1), 99);
});

test('a calculation left unfinished is not lost when the sheets change under it', async () => {
  const { workbook, history } = await blankBook();
  type(workbook, history, 0, 0, 0, '1');
  for (let r = 1; r < 400; r++) type(workbook, history, 0, r, 0, `=A${r}+1`);
  const engine = new CalcEngine(workbook, { budget: { work: 50 } });
  engine.recalculate(undefined, { overwrite: true });
  assert.equal(engine.incomplete, true);
  history.apply(insertSheet(workbook, 1, { name: 'Other' }), undefined, undefined);
  for (let k = 0; k < 10000 && (k === 0 || engine.incomplete); k++) engine.recalculate([]);
  assert.equal(engine.incomplete, false);
  assert.equal(valueAt(workbook, 0, 399, 0), 400);
  assert.deepEqual(engine.uncalculated(), []);
});

test('what is saved after a recalculation of everything holds a value for every formula', async () => {
  const { workbook, history, save } = await blankBook();
  type(workbook, history, 0, 0, 0, '2');
  const engine = new CalcEngine(workbook);
  engine.recalculate(undefined, { overwrite: true });
  type(workbook, history, 0, 1, 0, '=A1*4');
  history.apply(insertSheet(workbook, 1, { name: 'Sums' }), undefined, undefined);
  type(workbook, history, 1, 0, 0, '=Data!A2+1');
  engine.recalculate(undefined, { overwrite: true });
  const again = await (await open(await save())).workbook();
  assert.equal(valueAt(again, 0, 1, 0), 8);
  assert.equal(valueAt(again, 1, 0, 0), 9);
});
