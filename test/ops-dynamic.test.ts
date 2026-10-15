// Editing around the spill of a dynamic array formula: inserting and deleting rows and
// columns, sorting, copying and pasting, typing, undoing.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine } from '../src/calc/engine';
import { createBlank } from '../src/formats/blank';
import { getCell, open, type Workbook } from '../src/index';
import { copyRange, pasteAt, setCellInput } from '../src/ops/commands';
import { History, type Op } from '../src/ops/ops';
import { sortRange } from '../src/ops/sort';
import { formatRange } from '../src/ops/format';

async function sheetWith(inputs: [number, number, string][]): Promise<{ workbook: Workbook; history: History; engine: CalcEngine; apply(ops: Op[], structural?: boolean): void; column(col: number, rows: number): unknown[] }> {
  const workbook = await (await open(await createBlank('xlsx'))).workbook();
  const history = new History(workbook);
  const engine = new CalcEngine(workbook);
  const apply = (ops: Op[], structural = false) => {
    history.apply(ops, undefined, undefined);
    if (structural) {
      engine.invalidate();
      engine.recalculate(undefined, { overwrite: true });
    } else {
      const cells: { sheet: number; row: number; col: number }[] = [];
      for (const op of ops) if (op.type === 'setCells') for (const change of op.cells) cells.push({ sheet: op.sheet, row: change.row, col: change.col });
      engine.recalculate(cells);
    }
  };
  for (const [row, col, text] of inputs) apply(setCellInput(workbook, 0, row, col, text));
  const column = (col: number, rows: number) => Array.from({ length: rows }, (_, row) => getCell(workbook.sheets[0], row, col)?.value);
  return { workbook, history, engine, apply, column };
}

test('rows cannot be inserted or deleted through a spill, as through an array formula; around it they can', async () => {
  const { workbook, apply, column } = await sheetWith([[1, 1, '=SEQUENCE(4)'], [0, 3, '=SUM(B2#)']]);
  assert.deepEqual(column(1, 6), [undefined, 1, 2, 3, 4, undefined]);
  assert.throws(() => apply([{ type: 'insert', sheet: 0, axis: 'row', at: 3, count: 1 }], true), /through the array formula in B2:B5/);
  assert.throws(() => apply([{ type: 'delete', sheet: 0, axis: 'row', at: 2, count: 1 }], true), /through the array formula in B2:B5/);
  // Above it: the spill and what reads it move along.
  apply([{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 2 }], true);
  assert.deepEqual(column(1, 8), [undefined, undefined, undefined, 1, 2, 3, 4, undefined]);
  assert.equal(getCell(workbook.sheets[0], 3, 1)?.formula?.ref, 'B4:B7');
  assert.equal(getCell(workbook.sheets[0], 2, 3)?.formula?.text, 'SUM(_xlfn.ANCHORARRAY(B4))');
  assert.equal(getCell(workbook.sheets[0], 2, 3)?.value, 10);
  // A column inserted beside it changes nothing of it.
  apply([{ type: 'insert', sheet: 0, axis: 'col', at: 0, count: 1 }], true);
  assert.equal(getCell(workbook.sheets[0], 3, 2)?.formula?.ref, 'C4:C7');
  assert.deepEqual(column(2, 8), [undefined, undefined, undefined, 1, 2, 3, 4, undefined]);
});

test('a sort that takes in part of a spill is refused; one that takes in a formula of one value moves it', async () => {
  const { workbook, apply } = await sheetWith([[0, 0, '3'], [1, 0, '1'], [2, 0, '2'], [0, 1, '=SEQUENCE(2)'], [0, 2, '=A1*10']]);
  assert.throws(() => apply(sortRange(workbook, 0, { top: 0, left: 0, bottom: 2, right: 1 }, [{ col: 0 }])), /array formula in B1:B2/);
  apply(sortRange(workbook, 0, { top: 0, left: 0, bottom: 2, right: 0 }, [{ col: 0 }]));
  assert.deepEqual([0, 1, 2].map((row) => getCell(workbook.sheets[0], row, 0)?.value), [1, 2, 3]);
});

test('a copied formula spills where it is pasted; copied spilled cells paste as values', async () => {
  const { workbook, apply, column } = await sheetWith([[0, 0, '=SEQUENCE(3)']]);
  const copied = copyRange(workbook, 0, { top: 0, left: 0, bottom: 0, right: 0 });
  apply(pasteAt(workbook, 0, { top: 0, left: 2, bottom: 0, right: 2 }, copied.text, copied).ops);
  assert.equal(getCell(workbook.sheets[0], 0, 2)?.formula?.dynamic, true);
  assert.deepEqual(column(2, 3), [1, 2, 3]);
  const values = copyRange(workbook, 0, { top: 1, left: 0, bottom: 2, right: 0 });
  apply(pasteAt(workbook, 0, { top: 5, left: 0, bottom: 5, right: 0 }, values.text, values).ops);
  assert.deepEqual([getCell(workbook.sheets[0], 5, 0)?.value, getCell(workbook.sheets[0], 5, 0)?.formula], [2, undefined]);
  // Pasted onto the spill of another formula, the values are in its way.
  apply(pasteAt(workbook, 0, { top: 1, left: 2, bottom: 1, right: 2 }, values.text, values).ops);
  assert.deepEqual(column(2, 3), ['#SPILL!', 2, 3]);
  // And pasted into the spill they came from, one cell further down, too.
  const one = copyRange(workbook, 0, { top: 1, left: 0, bottom: 1, right: 0 });
  apply(pasteAt(workbook, 0, { top: 2, left: 0, bottom: 2, right: 0 }, one.text, one).ops);
  assert.deepEqual(column(0, 3), ['#SPILL!', undefined, 2]);
});

test('a spill copied with its formula is pasted empty, and the pasted formula fills it again (Ö5)', async () => {
  const { workbook, apply, column } = await sheetWith([[0, 0, '=SEQUENCE(3)']]);
  const copied = copyRange(workbook, 0, { top: 0, left: 0, bottom: 2, right: 0 });
  apply(pasteAt(workbook, 0, { top: 0, left: 2, bottom: 0, right: 2 }, copied.text, copied).ops);
  assert.deepEqual(column(2, 4), [1, 2, 3, undefined]);
  assert.equal(getCell(workbook.sheets[0], 0, 2)?.formula?.ref, 'C1:C3');
  assert.equal(getCell(workbook.sheets[0], 1, 2)?.spilled, true);
  // Without its formula, a spilled cell is a value where it is pasted.
  const part = copyRange(workbook, 0, { top: 1, left: 0, bottom: 2, right: 0 });
  apply(pasteAt(workbook, 0, { top: 0, left: 4, bottom: 0, right: 4 }, part.text, part).ops);
  assert.deepEqual([getCell(workbook.sheets[0], 0, 4)?.value, getCell(workbook.sheets[0], 0, 4)?.spilled], [2, undefined]);
});

test('typing into a spill makes its formula #SPILL!; undo brings the spill back', async () => {
  const { history, engine, apply, column } = await sheetWith([[0, 0, '=SEQUENCE(3)']]);
  apply((await import('../src/ops/commands')).setCellInput(history.workbook, 0, 1, 0, 'x'));
  assert.deepEqual(column(0, 3), ['#SPILL!', 'x', undefined]);
  history.undo();
  engine.recalculate([{ sheet: 0, row: 1, col: 0 }]);
  assert.deepEqual(column(0, 3), [1, 2, 3]);
  // Deleting a spilled cell changes nothing: it is the formula's.
  apply(setCellInput(history.workbook, 0, 2, 0, ''));
  assert.deepEqual(column(0, 3), [1, 2, 3]);
});

test('an ordinary formula of a file typed again as it was stays one; changed, it becomes a dynamic one', async () => {
  const { workbook, apply } = await sheetWith([[0, 0, '1'], [1, 0, '2']]);
  const sheet = workbook.sheets[0];
  const row = sheet.rows.get(1)!;
  sheet.rows.set(1, { ...row, cells: Object.assign(row.cells.slice(), { 2: { kind: 'number', value: 1, style: 0, formula: { text: 'A1:A2', type: 'normal' } } }) });
  apply(setCellInput(workbook, 0, 1, 2, '=A1:A2'));
  assert.equal(getCell(sheet, 1, 2)?.formula?.dynamic, undefined);
  assert.equal(getCell(sheet, 1, 2)?.value, 2, 'implicit intersection: the cell in its row');
  apply(setCellInput(workbook, 0, 1, 2, '=A1:A2+0'));
  assert.equal(getCell(sheet, 1, 2)?.formula?.dynamic, true);
  assert.deepEqual([getCell(sheet, 1, 2)?.value, getCell(sheet, 2, 2)?.value], [1, 2]);
});

// ---------------------------------------------------------------------------
// Which cells are a spill's: the model says so (Cell.spilled), not what one engine remembers.

test('a format put on a cell of a spill leaves it the spill\'s; undone, nothing changed', async () => {
  const { workbook, history, engine, apply, column } = await sheetWith([[0, 0, '=SEQUENCE(3)']]);
  apply(formatRange(workbook, 0, { top: 1, left: 0, bottom: 1, right: 0 }, { bold: true }));
  assert.deepEqual(column(0, 3), [1, 2, 3]);
  assert.equal(getCell(workbook.sheets[0], 1, 0)?.spilled, true);
  // The spill grows past it and shrinks back: the formatted cell follows, and keeps its format.
  apply(setCellInput(workbook, 0, 0, 0, '=SEQUENCE(4)'));
  assert.deepEqual(column(0, 4), [1, 2, 3, 4]);
  apply(setCellInput(workbook, 0, 0, 0, '=SEQUENCE(1)'));
  assert.deepEqual(column(0, 4), [1, null, undefined, undefined], 'the formatted cell stays, empty');
  assert.notEqual(getCell(workbook.sheets[0], 1, 0)?.style, 0, 'the format stays where the spill went');
  history.undo();
  engine.recalculate([{ sheet: 0, row: 0, col: 0 }]);
  assert.deepEqual(column(0, 4), [1, 2, 3, 4]);
});

test('typed into a spill, the formula deleted, both undone: the spill is back', async () => {
  const { history, engine, apply, column, workbook } = await sheetWith([[0, 0, '=SEQUENCE(3)']]);
  apply(setCellInput(workbook, 0, 1, 0, 'x'));
  apply(setCellInput(workbook, 0, 0, 0, ''));
  assert.deepEqual(column(0, 3), [undefined, 'x', undefined]);
  history.undo();
  engine.recalculate([{ sheet: 0, row: 0, col: 0 }]);
  assert.deepEqual(column(0, 3), ['#SPILL!', 'x', undefined]);
  history.undo();
  engine.recalculate([{ sheet: 0, row: 1, col: 0 }]);
  assert.deepEqual(column(0, 3), [1, 2, 3]);
  // An engine that reads the workbook afresh comes to the same.
  const again = new CalcEngine(workbook);
  again.recalculate(undefined, { overwrite: true });
  assert.deepEqual(column(0, 3), [1, 2, 3]);
});

test('a spilled value pasted somewhere is a value of its own there', async () => {
  const { workbook, apply } = await sheetWith([[0, 0, '=SEQUENCE(3)']]);
  const copied = copyRange(workbook, 0, { top: 1, left: 0, bottom: 1, right: 0 });
  apply(pasteAt(workbook, 0, { top: 0, left: 4, bottom: 0, right: 4 }, copied.text, copied).ops);
  assert.equal(getCell(workbook.sheets[0], 0, 4)?.value, 2);
  assert.equal(getCell(workbook.sheets[0], 0, 4)?.spilled, undefined);
});
