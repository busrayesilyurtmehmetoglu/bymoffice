// A range read as an array is read whole, empty cells included: what a formula gives never depends on
// how far its sheet was once written (Sheet.rowCount), which differs between two people who did the same
// edits in another order, and between one who edited and one who opened the saved file.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine } from '../src/calc/engine';
import { createBlank } from '../src/formats/blank';
import { getCell, open, writeXlsx, type Workbook } from '../src/index';
import { setCellInput } from '../src/ops/commands';
import { History, type Op } from '../src/ops/ops';

async function session(inputs: [string, string][]): Promise<{ workbook: Workbook; type(address: string, text: string): void; value(address: string): unknown; ref(address: string): string | undefined; save(): Promise<Uint8Array> }> {
  const file = await open(await createBlank('xlsx'));
  const workbook = await file.workbook();
  const history = new History(workbook);
  const engine = new CalcEngine(workbook);
  const at = (address: string) => {
    const m = /^([A-Z]+)([0-9]+)$/.exec(address) as RegExpExecArray;
    let col = 0;
    for (const ch of m[1]) col = col * 26 + ch.charCodeAt(0) - 64;
    return { row: Number(m[2]) - 1, col: col - 1 };
  };
  const type = (address: string, text: string) => {
    const { row, col } = at(address);
    const ops: Op[] = setCellInput(workbook, 0, row, col, text);
    history.apply(ops, undefined, undefined);
    engine.recalculate([{ sheet: 0, row, col }]);
  };
  for (const [address, text] of inputs) type(address, text);
  const cell = (address: string) => {
    const { row, col } = at(address);
    return getCell(workbook.sheets[0], row, col);
  };
  const value = (address: string) => {
    const c = cell(address);
    return c ? c.value : undefined;
  };
  const ref = (address: string) => cell(address)?.formula?.ref;
  const save = async () => {
    await writeXlsx(file.package!, workbook);
    return file.package!.save();
  };
  return { workbook, type, value, ref, save };
}

test('a range read as an array has its full size, whatever the sheet holds (E1)', async () => {
  const t = await session([['A1', '1'], ['A2', '2'], ['J10', 'x']]);
  t.type('B1', '=A1:A30*2');
  assert.equal(t.ref('B1'), 'B1:B30');
  assert.equal(t.value('B30'), 0);
  t.type('C1', '=FILTER(A1:A30,A1:A30<>"")');
  assert.equal(t.ref('C1'), 'C1:C2');
  assert.deepEqual([t.value('C1'), t.value('C2')], [1, 2]);
  t.type('D1', '=SUMPRODUCT(--(A1:A30=""))');
  assert.equal(t.value('D1'), 28);
  t.type('E1', '=ROWS(A1:A30*1)');
  assert.equal(t.value('E1'), 30);
  t.type('F1', '=COUNTBLANK(A1:A30)');
  assert.equal(t.value('F1'), 28);
});

test('a whole column is read as far as its columns hold something, whatever was written and cleared there (K-A)', async () => {
  // UNSURE: Excel reads all 1,048,576 rows of a whole column; here it is cut to the last row of its columns that holds
  // a value, a formula or a spill (not how far the sheet was once written, which differs between two people).
  const t = await session([['A1', '1'], ['A2', 'x'], ['A5', '7'], ['A30', 'y'], ['A30', '']]);
  t.type('C1', '=SUMPRODUCT(--(A:A=""))');
  // A1 to A5: A3 and A4 are empty. The cleared A30 counts for nothing.
  assert.equal(t.value('C1'), 2);
  t.type('C2', '=SUMPRODUCT((A:A="x")*1)');
  assert.equal(t.value('C2'), 1);
  t.type('C3', '=SUMPRODUCT(A:A,A:A)');
  assert.equal(t.value('C3'), 50);
  t.type('D1', '=A:A*2');
  assert.equal(t.ref('D1'), 'D1:D5');
  assert.deepEqual([t.value('D1'), `${t.value('D2')}`, t.value('D4'), t.value('D5')], [2, '#VALUE!', 0, 14]);
  // What is typed in other columns changes nothing of it.
  t.type('B9', '5');
  assert.deepEqual([t.value('C1'), t.ref('D1')], [2, 'D1:D5']);
  // Its own column does.
  t.type('A7', '1');
  assert.deepEqual([t.value('C1'), t.ref('D1')], [3, 'D1:D7']);
  t.type('E1', '=ROWS(A:A)');
  assert.equal(t.value('E1'), 1048576);
});

test('one who edited and one who opened the saved file get the same results: a cell written and cleared far away changes nothing (E1)', async () => {
  const member = await session([['A1', '1'], ['A2', '2'], ['F20', 'x'], ['F20', '']]);
  assert.equal(member.workbook.sheets[0].rowCount >= 20, true);
  const joinerFile = await open(await member.save());
  const joinerBook = await joinerFile.workbook();
  const joinerHistory = new History(joinerBook);
  const joinerEngine = new CalcEngine(joinerBook);
  joinerEngine.recalculate(joinerEngine.uncalculated());
  const steps: [number, number, string][] = [
    [0, 1, '=A1:A30*2'],
    [0, 2, '=SUMPRODUCT(--(A1:A30=""))'],
    [0, 3, '=FILTER(A1:A30,A1:A30<>"")'],
    [0, 4, '=MATCH(9,A1:A30,1)'],
    [0, 6, '=SUMPRODUCT(--(A:A=""))'],
  ];
  for (const [row, col, text] of steps) {
    member.type(String.fromCharCode(65 + col) + (row + 1), text);
    joinerHistory.apply(setCellInput(joinerBook, 0, row, col, text), undefined, undefined);
    joinerEngine.recalculate([{ sheet: 0, row, col }]);
  }
  for (const [row, col] of steps) {
    const a = getCell(member.workbook.sheets[0], row, col);
    const b = getCell(joinerBook.sheets[0], row, col);
    assert.deepEqual([a?.value, a?.formula?.ref], [b?.value, b?.formula?.ref], `${String.fromCharCode(65 + col)}${row + 1}`);
  }
  assert.equal(member.value('C1'), 28);
});

test('a binary search ends at the last cell that holds something: how far the sheet was written does not move it (E1)', async () => {
  // Unsorted on purpose: where a binary search looks depends on how many cells it searches.
  const inputs: [string, string][] = [['A1', '1'], ['A2', '5'], ['A3', '2'], ['A4', '8'], ['A5', '3']];
  const near = await session(inputs);
  const far = await session([...inputs, ['Z900', 'x'], ['Z900', '']]);
  for (const t of [near, far]) {
    t.type('B1', '=MATCH(4,A:A,1)');
    t.type('B2', '=VLOOKUP(4,A:A,1,TRUE)');
    t.type('B3', '=XMATCH(4,A:A,-1,2)');
  }
  assert.deepEqual([far.value('B1'), far.value('B2'), far.value('B3')], [near.value('B1'), near.value('B2'), near.value('B3')]);
});

test('XLOOKUP returns the whole row or column of what it returns, empty cells included (E1)', async () => {
  const t = await session([['A1', 'k'], ['B1', '7'], ['C1', '8']]);
  t.type('A3', '=XLOOKUP("k",A1:A2,B1:F2)');
  assert.equal(t.ref('A3'), 'A3:E3');
  assert.deepEqual([t.value('A3'), t.value('B3'), t.value('E3')], [7, 8, 0]);
});

test('database functions: the empty records below the rows the sheet has match a criterion that matches empty (E1)', async () => {
  // Criterion "=": a record whose Name is empty. Rows 3 to 30 are such records.
  const base: [string, string][] = [['A1', 'Name'], ['B1', 'Qty'], ['A2', 'a'], ['B2', '1'], ['D1', 'Name'], ['D2', '="="']];
  const near = await session(base);
  const far = await session([...base, ['Q99', 'x'], ['Q99', '']]);
  for (const t of [near, far]) {
    t.type('F1', '=DCOUNT(A1:B30,,D1:D2)');
    t.type('F2', '=DCOUNTA(A1:B30,"Name",D1:D2)');
    t.type('F3', '=DGET(A1:B30,"Qty",D1:D2)');
    t.type('F4', '=DGET(A1:B3,"Qty",D1:D2)');
  }
  // UNSURE: what Excel counts for records past the cells a sheet has was not checked; here they are records like any other.
  assert.equal(near.value('F1'), 28);
  assert.deepEqual(['F1', 'F2', 'F3', 'F4'].map((a) => far.value(a)), ['F1', 'F2', 'F3', 'F4'].map((a) => near.value(a)));
  assert.equal(`${near.value('F3')}`, '#NUM!');
});

test('formulas over whole columns take a bounded time (a guard against work per cell of the column, not a measure of speed)', async () => {
  const inputs: [string, string][] = [];
  for (let r = 1; r <= 2000; r++) inputs.push([`A${r}`, r % 3 ? 'x' : 'y'], [`B${r}`, String(r)]);
  const t = await session(inputs);
  const started = Date.now();
  for (let k = 1; k <= 40; k++) t.type(`D${k}`, `=SUMPRODUCT((A:A="x")*(B:B))+${k}`);
  const took = Date.now() - started;
  assert.equal(t.value("D1"), 1334667 + 1);
  assert.ok(took < 20000, `${took} ms`);
});
