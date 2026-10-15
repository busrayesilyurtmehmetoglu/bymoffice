// A property of the engine: whatever was typed, undone, inserted or saved in between, the engine of a
// session leaves the workbook as an engine built afresh from that workbook does. Values, which cells a
// spill fills and the `ref` of every formula are compared after every step, and after saving and opening.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine, type CalcEngineOptions, type CellAddress } from '../src/calc/engine';
import { createBlank } from '../src/formats/blank';
import { open, writeXlsx, type Workbook } from '../src/index';
import { clearRange, copyRange, pasteAt, setCellInput } from '../src/ops/commands';
import { replaceAllOps } from '../src/ops/find';
import { copySheet, deleteSheet, insertSheet, renameSheet } from '../src/ops/sheets';
import { sortRange } from '../src/ops/sort';
import { formatRange } from '../src/ops/format';
import { History, type Op } from '../src/ops/ops';
import type { Sheet } from '../src/model/sheet/types';

/** A small, repeatable random sequence (mulberry32). */
function random(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const ROWS = 9;
const COLS = 5;

// Inputs: values, and formulas that spill, read spills, compare empty cells and count them.
const INPUTS = [
  '1', '2', '3', '0', 'x', 'b', '', '', 'TRUE',
  '=SEQUENCE(2)', '=SEQUENCE(A1)', '=SEQUENCE(1,3)', '=SEQUENCE(COUNT(B2:B8)+1)', '=SEQUENCE(COUNT(A2:A8)+1)',
  '=A1:A4*2', '=A1:A12*1', '=B1:C2', '=FILTER(A1:A12,A1:A12<>"")', '=SORT(B1:B5)', '=UNIQUE(A1:A6)',
  '=A1#', '=SUM(B1#)', '=ROWS(C1#)', '=COUNT(A1:A9)', '=SUMPRODUCT(--(A1:A12=""))', '=ROWS(A1:A12*1)',
  '=COUNTBLANK(A1:A12)', '=A1+1', '=B2*2', '=IF(A1>1,SEQUENCE(2,2),0)', '=TRANSPOSE(A1:A3)', '=MATCH(2,A1:A12,0)',
  '=XLOOKUP(2,A1:A12,B1:B12)', '=LET(n,COUNT(A1:A9),SEQUENCE(n+1))', '=MAP(A1:A3,LAMBDA(v,v*10))',
  '=A:A*1', '=SUMPRODUCT(--(A:A=""))', '=COUNTA(B:B)', '=FILTER(A:A,A:A<>"")', '=SUM(C:C)',
];

interface Limits {
  spillCells?: number;
  workbookSpillCells?: number;
  /** The session's engine works in slices of a tiny budget (a clock of its own): the result must not depend on them. */
  sliced?: boolean;
}

/** What a sheet holds, cell by cell: value, kind, spill mark, formula and its `ref`. */
function snapshot(sheet: Sheet): string[] {
  const out: string[] = [];
  const keys = Array.from(sheet.rows.keys()).sort((a, b) => a - b);
  for (const r of keys) {
    const row = sheet.rows.get(r);
    if (!row) continue;
    row.cells.forEach((cell, c) => {
      if (!cell) return;
      const empty = !cell.formula && (cell.value === null || cell.value === undefined || cell.value === '');
      if (empty && !cell.spilled) return;
      const f = cell.formula ? ` f=${cell.formula.text}${cell.formula.ref ? ` ref=${cell.formula.ref}` : ''}` : '';
      out.push(`R${r + 1}C${c + 1} ${empty ? '' : `${cell.kind}:${String(cell.value)}`}${cell.spilled ? ' spilled' : ''}${f}`);
    });
  }
  return out;
}

/** A copy of a workbook an engine can write into without touching the original. */
function cloneBook(workbook: Workbook): Workbook {
  return {
    ...workbook,
    sheets: workbook.sheets.map((sheet) => ({ ...sheet, rows: new Map(Array.from(sheet.rows, ([k, row]) => [k, { ...row, cells: row.cells.slice() }])) })),
  };
}

function freshResult(workbook: Workbook, options: CalcEngineOptions): string[] {
  const copy = cloneBook(workbook);
  const engine = new CalcEngine(copy, options);
  engine.recalculate(undefined, { overwrite: true });
  assert.equal(engine.incomplete, false);
  return snapshot(copy.sheets[0]);
}

interface Step {
  cells: CellAddress[];
  structural: boolean;
}

async function run(seed: number, steps: number, limits: Limits): Promise<void> {
  const next = random(seed);
  const pick = <T>(list: T[]): T => list[Math.floor(next() * list.length)];
  const options: CalcEngineOptions = { random: () => 0.5, now: () => 45000, limits: { spillCells: limits.spillCells, workbookSpillCells: limits.workbookSpillCells } };
  let time = 0;
  // A budget of a few units of work and a clock that moves with every look: the recalculation is cut into many calls.
  const sessionOptions: CalcEngineOptions = limits.sliced
    ? { ...options, budget: { milliseconds: 1, work: 3 + Math.floor(next() * 40), hardMilliseconds: 1e9, hardWork: 1e12 }, clock: () => (time += 0.4) }
    : options;
  const finish = () => {
    for (let k = 0; k < 1000000 && engine.incomplete; k++) engine.recalculate([]);
  };
  let file = await open(await createBlank('xlsx'));
  let workbook = await file.workbook();
  let history = new History(workbook);
  let engine = new CalcEngine(workbook, sessionOptions);
  let undone: Step[] = [];
  let done: Step[] = [];
  const log: string[] = [];
  const apply = (ops: Op[], step: Step) => {
    if (ops.length === 0) return;
    const inverse = history.apply(ops, undefined, undefined);
    // As the view does: the cells to calculate again are the ones the inverse names (a format names them too).
    if (!step.structural) {
      step.cells = [];
      for (const op of inverse) if (op.type === 'setCells') for (const change of op.cells) step.cells.push({ sheet: op.sheet, row: change.row, col: change.col });
    }
    done.push(step);
    undone = [];
    recalc(step);
  };
  const recalc = (step: Step) => {
    if (step.structural) {
      engine.invalidate();
      engine.recalculate(undefined, { overwrite: true });
    } else engine.recalculate(step.cells);
    finish();
  };
  for (let k = 0; k < steps; k++) {
    const what = next();
    const row = Math.floor(next() * ROWS);
    const col = Math.floor(next() * COLS);
    try {
      if (what < 0.55) {
        const text = pick(INPUTS);
        log.push(`type R${row + 1}C${col + 1} ${text}`);
        apply(setCellInput(workbook, 0, row, col, text), { cells: [{ sheet: 0, row, col }], structural: false });
      } else if (what < 0.62) {
        log.push(`clear R${row + 1}C${col + 1}`);
        const range = { top: row, left: col, bottom: row, right: col };
        apply(clearRange(workbook, 0, range), { cells: [{ sheet: 0, row, col }], structural: false });
      } else if (what < 0.67) {
        log.push(`bold R${row + 1}C${col + 1}:R${row + 2}C${col + 1}`);
        const range = { top: row, left: col, bottom: row + 1, right: col };
        apply(formatRange(workbook, 0, range, { bold: true }), { cells: [], structural: false });
      } else if (what < 0.8) {
        const step = done.pop();
        if (!step) continue;
        log.push('undo');
        history.undo();
        undone.push(step);
        recalc(step);
      } else if (what < 0.85) {
        const step = undone.pop();
        if (!step) continue;
        log.push('redo');
        history.redo();
        done.push(step);
        recalc(step);
      } else if (what < 0.93) {
        const insert = next() < 0.5;
        log.push(`${insert ? 'insert' : 'delete'} row ${row + 1}`);
        const op: Op = { type: insert ? 'insert' : 'delete', sheet: 0, axis: 'row', at: row, count: 1 };
        apply([op], { cells: [], structural: true });
      } else {
        log.push('save and open');
        await writeXlsx(file.package!, workbook);
        file = await open(await file.package!.save());
        workbook = await file.workbook();
        history = new History(workbook);
        done = [];
        undone = [];
        engine = new CalcEngine(workbook, sessionOptions);
        engine.recalculate(engine.uncalculated());
        finish();
      }
    } catch (error) {
      // An edit the engine refuses (through a spill, a merged cell…) changes nothing.
      log.push(`  refused: ${(error as Error).message}`);
      continue;
    }
    assert.equal(engine.incomplete, false, `seed ${seed} does not finish, after:\n${log.join('\n')}`);
    const session = snapshot(workbook.sheets[0]);
    const fresh = freshResult(workbook, options);
    assert.deepEqual(session, fresh, `seed ${seed}, after:\n${log.join('\n')}`);
  }
}

test('the engine of a session leaves a workbook as one built afresh would, step after step', async () => {
  for (let seed = 1; seed <= 150; seed++) await run(seed, 50, {});
});

test('the same with small budgets for spills: who spills does not depend on the order formulas came in', async () => {
  for (let seed = 1001; seed <= 1150; seed++) await run(seed, 50, { spillCells: 6, workbookSpillCells: 10 });
});

test('the same with a recalculation cut into slices of a tiny budget: what it leaves does not depend on them (K-C)', async () => {
  for (let seed = 2001; seed <= 2100; seed++) await run(seed, 40, { sliced: true });
  for (let seed = 3001; seed <= 3050; seed++) await run(seed, 40, { sliced: true, spillCells: 6, workbookSpillCells: 10 });
});

test('seeds that once left a session apart from a workbook built afresh, in all four ways (a stale list of spills, a cell two spills were over, a circle the file was saved with)', async () => {
  const ways: Limits[] = [{}, { spillCells: 6, workbookSpillCells: 10 }, { sliced: true }, { sliced: true, spillCells: 6, workbookSpillCells: 10 }];
  for (const seed of [4092, 4335, 4409, 5409, 5775]) for (const limits of ways) await run(seed, 50, limits);
});

// A circular formula is calculated again only when what it reads changes (an unrelated edit and its undo leave the file
// as it was): then these seeds leave a session apart from a workbook built afresh again, a circle whose partner changed
// being kept by the session. Known and open (E3: conflicts of spills settled from no spill); reported, not hidden.
test('seeds a circle through spills still leaves apart when its partner changes (open)', { todo: 'E3 is open' }, async () => {
  const ways: Limits[] = [{}, { spillCells: 6, workbookSpillCells: 10 }, { sliced: true }, { sliced: true, spillCells: 6, workbookSpillCells: 10 }];
  for (const seed of [4196, 6167, 6225]) for (const limits of ways) await run(seed, 50, limits);
});

/** Every sheet of a workbook, cell by cell (see snapshot()). */
function snapshotAll(workbook: Workbook): string[] {
  const out: string[] = [];
  for (const sheet of workbook.sheets) for (const line of snapshot(sheet)) out.push(`${sheet.name}!${line}`);
  return out;
}

const MORE_INPUTS = ['1', '2', 'x', '', '=SEQUENCE(2)', '=SEQUENCE(A1)', '=SEQUENCE(1,3)', '=A1:A4*2', '=FILTER(A1:A9,A1:A9<>"")', '=SORT(B1:B5)',
  '=A1#', '=SUM(B1#)', '=COUNT(A1:A9)', '=A1+1', '=SEQUENCE(COUNT(B1:B8)+1)', '=XLOOKUP(2,A1:A9,B1:B9)', '=Sheet2!A1:A3', '=SUM(Sheet2!A:A)',
  '=SEQUENCE(Sheet2!A1)', '=A:A*1', '=TRANSPOSE(A1:A3)'];

/**
 * The same over two sheets, with what else edits cells: columns inserted and deleted, sorting, copying and pasting,
 * replacing, sheets inserted, renamed, copied and deleted; the session's engine works in slices or not.
 */
async function runMore(seed: number, steps: number, sliced: boolean): Promise<void> {
  const next = random(seed);
  const pick = <T>(list: T[]): T => list[Math.floor(next() * list.length)];
  const options: CalcEngineOptions = { random: () => 0.5, now: () => 45000 };
  let time = 0;
  const sessionOptions: CalcEngineOptions = sliced ? { ...options, budget: { milliseconds: 1, work: 3 + Math.floor(next() * 40), hardMilliseconds: 1e9, hardWork: 1e12 }, clock: () => (time += 0.4) } : options;
  const workbook = await (await open(await createBlank('xlsx'))).workbook();
  const history = new History(workbook);
  const engine = new CalcEngine(workbook, sessionOptions);
  const log: string[] = [];
  const recalc = (cells: CellAddress[] | undefined) => {
    if (!cells) {
      engine.invalidate();
      engine.recalculate(undefined, { overwrite: true });
    } else engine.recalculate(cells);
    for (let k = 0; k < 1000000 && engine.incomplete; k++) engine.recalculate([]);
  };
  const steps2: (CellAddress[] | undefined)[] = [];
  const apply = (ops: Op[]) => {
    if (ops.length === 0) return;
    const inverse = history.apply(ops, undefined, undefined);
    let cells: CellAddress[] | undefined = [];
    for (const op of inverse) {
      if (op.type !== 'setCells') {
        cells = undefined;
        break;
      }
      for (const change of op.cells) cells.push({ sheet: op.sheet, row: change.row, col: change.col });
    }
    steps2.push(cells);
    recalc(cells);
  };
  apply(insertSheet(workbook, 1, { name: 'Sheet2' }));
  for (let k = 0; k < steps; k++) {
    const what = next();
    const row = Math.floor(next() * 8);
    const col = Math.floor(next() * 5);
    const sh = workbook.sheets.length > 1 && next() < 0.25 ? 1 : 0;
    try {
      if (what < 0.5) {
        const text = pick(MORE_INPUTS);
        log.push(`type ${sh} R${row + 1}C${col + 1} ${text}`);
        apply(setCellInput(workbook, sh, row, col, text));
      } else if (what < 0.6) {
        const from = { top: row, left: col, bottom: row + 1, right: col + 1 };
        const copied = copyRange(workbook, sh, from);
        const r2 = Math.floor(next() * 8);
        const c2 = Math.floor(next() * 5);
        log.push(`copy ${sh} R${row + 1}C${col + 1} to R${r2 + 1}C${c2 + 1}`);
        apply(pasteAt(workbook, sh, { top: r2, left: c2, bottom: r2, right: c2 }, copied.text, copied).ops);
      } else if (what < 0.67) {
        log.push(`sort ${sh} column ${col + 1}`);
        apply(sortRange(workbook, sh, { top: 0, left: col, bottom: 7, right: Math.min(4, col + 1) }, [{ col, descending: next() < 0.5 }]));
      } else if (what < 0.72) {
        const a = pick(['1', '2', 'A1', 'B', 'SEQUENCE']);
        const b = pick(['3', '2', 'B1', 'C', 'SEQUENCE']);
        log.push(`replace ${a} with ${b}`);
        apply(replaceAllOps(workbook, a, b, {}).ops);
      } else if (what < 0.77) {
        if (workbook.sheets.length > 1) {
          const name = workbook.sheets[1].name === 'Sheet2' ? 'Other' : 'Sheet2';
          log.push(`rename to ${name}`);
          apply(renameSheet(workbook, 1, name));
        }
      } else if (what < 0.8) {
        if (workbook.sheets.length > 1 && next() < 0.5) {
          log.push('delete sheet 2');
          apply(deleteSheet(workbook, 1));
        } else {
          log.push('copy sheet 1');
          apply(copySheet(workbook, 0));
        }
      } else if (what < 0.88) {
        const cells = steps2.pop();
        if (history.undo() === undefined) continue;
        log.push('undo');
        recalc(cells);
      } else {
        const insert = next() < 0.5;
        log.push(`${insert ? 'insert' : 'delete'} column ${col + 1} of ${sh}`);
        apply([{ type: insert ? 'insert' : 'delete', sheet: sh, axis: 'col', at: col, count: 1 }]);
      }
    } catch (error) {
      log.push(`  refused: ${(error as Error).message}`);
      continue;
    }
    assert.equal(engine.incomplete, false);
    const copy = cloneBook(workbook);
    new CalcEngine(copy, options).recalculate(undefined, { overwrite: true });
    assert.deepEqual(snapshotAll(workbook), snapshotAll(copy), `seed ${seed}${sliced ? ' (sliced)' : ''}, after:\n${log.join('\n')}`);
  }
}

test('over two sheets, with columns, sorting, pasting, replacing and sheet edits, in slices or not (B7)', async () => {
  for (let seed = 1; seed <= 60; seed++) await runMore(seed, 40, false);
  for (let seed = 101; seed <= 140; seed++) await runMore(seed, 40, true);
});

test('an engine cut into slices of three units of work and one that is not leave the same, edit after edit (K-C, seed 107)', async () => {
  const inputs = ['1', '2', '3', '0', 'x', '', '', '=SEQUENCE(2)', '=SEQUENCE(A1)', '=SEQUENCE(1,3)', '=SEQUENCE(COUNT(B2:B8)+1)', '=SEQUENCE(COUNT(A2:A8)+1)',
    '=A1:A4*2', '=B1:C2', '=FILTER(A1:A12,A1:A12<>"")', '=SORT(B1:B5)', '=UNIQUE(A1:A6)', '=A1#', '=SUM(B1#)', '=ROWS(C1#)', '=COUNT(A1:A9)', '=A1+1',
    '=B2*2', '=IF(A1>1,SEQUENCE(2,2),0)', '=TRANSPOSE(A1:A3)', '=IF(ISNUMBER(A2),1,2)', '=IF(ISNUMBER(B3),2,3)', '=SEQUENCE(B1)', '=SEQUENCE(C1,2)'];
  for (const seed of [107, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20]) {
    const next = random(seed);
    const make = async (sliced: boolean) => {
      const workbook = await (await open(await createBlank('xlsx'))).workbook();
      let time = 0;
      const engine = new CalcEngine(workbook, sliced ? { budget: { milliseconds: 1, work: 3, hardMilliseconds: 1e6, hardWork: 1e9 }, clock: () => (time += 0.4) } : {});
      return { workbook, history: new History(workbook), engine };
    };
    const whole = await make(false);
    const sliced = await make(true);
    const log: string[] = [];
    for (let k = 0; k < 40; k++) {
      const row = Math.floor(next() * 6);
      const col = Math.floor(next() * 4);
      const clear = next() < 0.1;
      const text = inputs[Math.floor(next() * inputs.length)];
      log.push(clear ? `clear R${row + 1}C${col + 1}` : `R${row + 1}C${col + 1} ${text}`);
      for (const one of [whole, sliced]) {
        try {
          one.history.apply(clear ? clearRange(one.workbook, 0, { top: row, left: col, bottom: row, right: col }) : setCellInput(one.workbook, 0, row, col, text), undefined, undefined);
        } catch (error) {
          continue;
        }
        one.engine.recalculate([{ sheet: 0, row, col }]);
        for (let n = 0; n < 1000000 && one.engine.incomplete; n++) one.engine.recalculate([]);
      }
      assert.deepEqual(snapshot(sliced.workbook.sheets[0]), snapshot(whole.workbook.sheets[0]), `seed ${seed}, after:\n${log.join('\n')}`);
    }
  }
});
