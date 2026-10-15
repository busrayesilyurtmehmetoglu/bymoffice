// How far a whole column reads (the last row its column holds something in) follows every way cells change: a value
// tried and taken back, a recalculation of everything that was not told what changed, the cells an old array formula
// writes. What the session gives must be what an engine built afresh gives.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine } from '../src/calc/engine';
import { createBlank } from '../src/formats/blank';
import { tryOps } from '../src/format/validation';
import { getCell, open, type Workbook } from '../src/index';
import { clearRange, setCellInput } from '../src/ops/commands';
import { History, type Op } from '../src/ops/ops';
import { buildXlsx, worksheetXml } from './fixtures';

const value = (wb: Workbook, row: number, col: number) => getCell(wb.sheets[0], row, col)?.value;

function fresh(wb: Workbook, row: number, col: number): unknown {
  const copy: Workbook = { ...wb, sheets: wb.sheets.map((sheet) => ({ ...sheet, rows: new Map(Array.from(sheet.rows, ([k, r]) => [k, { ...r, cells: r.cells.slice() }])) })) };
  new CalcEngine(copy, {}).recalculate(undefined, { overwrite: true });
  return value(copy, row, col);
}

async function session(bytes?: Uint8Array) {
  const workbook = await (await open(bytes || (await createBlank('xlsx')))).workbook();
  const history = new History(workbook);
  const engine = new CalcEngine(workbook, {});
  engine.recalculate(engine.uncalculated());
  const tell = (ops: Op[]) => {
    const inverse = history.apply(ops, undefined, undefined);
    const cells: { sheet: number; row: number; col: number }[] = [];
    for (const op of inverse) if (op.type === 'setCells') for (const change of op.cells) cells.push({ sheet: op.sheet, row: change.row, col: change.col });
    engine.recalculate(cells);
  };
  const type = (row: number, col: number, text: string) => tell(setCellInput(workbook, 0, row, col, text));
  return { workbook, engine, history, tell, type };
}

test('a value tried and taken back (a rule of validation) leaves how far a column reads as it was (B1)', async () => {
  for (const [rows, row, text] of [[[0, 1, 2], 49, 'x'], [[0, 1, 2], 2, ''], [[0, 4], 4, '']] as [number[], number, string][]) {
    const t = await session();
    for (const r of rows) t.type(r, 0, String(r + 1));
    t.type(0, 2, '=SUMPRODUCT(--(A:A=""))+$E$1');
    t.type(0, 4, '0');
    tryOps(t.workbook, t.engine, setCellInput(t.workbook, 0, row, 0, text), () => value(t.workbook, 0, 2));
    // An edit C1 reads, not one of column A.
    t.type(0, 4, '0.5');
    assert.equal(value(t.workbook, 0, 2), fresh(t.workbook, 0, 2), `A${row + 1} = "${text}" tried`);
  }
});

test('a recalculation of everything finds how far columns hold something again: it was not told what changed (B2)', async () => {
  for (const overwrite of [false, true]) {
    const t = await session();
    t.type(0, 0, '1');
    t.type(1, 0, '2');
    t.type(0, 2, '=SUMPRODUCT(--(A:A=""))');
    t.type(0, 3, '=ROWS(A:A*1)');
    // Written without the engine being told, then all calculated.
    t.history.apply(setCellInput(t.workbook, 0, 9, 0, '10'), undefined, undefined);
    t.engine.recalculate(undefined, overwrite ? { overwrite: true } : undefined);
    assert.deepEqual([value(t.workbook, 0, 2), value(t.workbook, 0, 3)], [7, 10], `overwrite ${overwrite}`);
  }
});

test('the cells an old array formula writes hold something, as an engine built afresh sees them (B3)', async () => {
  const xml = '<row r="1"><c r="A1"><f t="array" ref="A1:A5">ROW(A1:A5)*D1</f><v>1</v></c><c r="D1"><v>1</v></c><c r="E1"><v>0</v></c></row>';
  const t = await session(await buildXlsx({ sheets: [{ name: 'S', xml: worksheetXml(`<sheetData>${xml}</sheetData>`) }] }));
  t.type(0, 5, '=ROWS(A:A*1)+E1');
  t.type(0, 3, '2');
  try {
    t.tell(clearRange(t.workbook, 0, { top: 0, left: 0, bottom: 0, right: 0 }));
  } catch {
    // Refused: the array stays.
  }
  t.type(0, 4, '0.5');
  assert.equal(value(t.workbook, 0, 5), fresh(t.workbook, 0, 5));
});
