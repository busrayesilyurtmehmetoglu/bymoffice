// The one who edits alone in a room - or beside someone who writes on another sheet only - has
// what the history of a view that is not edited together gives: after every change, undo and
// redo the workbook is the same, cell for cell, formula for formula, row for row; and nothing
// is said. Changes of every kind that can be undone, in any order.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { open } from '../src/index';
import { setCellInput } from '../src/ops/commands';
import { deleteColumns, deleteRows } from '../src/ops/filter';
import { formatRange, setColumnWidth, setRowHeight } from '../src/ops/format';
import { History, applyOp, type Op } from '../src/ops/ops';
import { insertSheet, moveSheet, renameSheet } from '../src/ops/sheets';
import { sortRange } from '../src/ops/sort';
import type { Workbook } from '../src/index';
import { Net, act, canon, firstDifference, int, join, pick, seeded, startFile } from './collab-fixtures';

const SEEDS = Math.max(1, Math.min(Number(process.env.COLLAB_SEEDS) || 150, 100000)) * 14;

const TEXTS = ['12', '3.5', 'text', '=A1+1', '=SUM(B2:B5)', '=$C$4+1', '=Other!A1*2', '=Data!C3', '=B3&"!"', '=SUM(A:A)', '=$B6*C$2', '=IF(B2>3,1,0)', ''];

/** One change that can be undone; `sheets` says whether the list of sheets may change. */
function make(wb: Workbook, random: () => number, sheets: boolean): () => Op[] {
  const s = sheets ? int(random, 0, wb.sheets.length - 1) : 0;
  const row = int(random, 0, 11);
  const col = int(random, 0, 5);
  const kinds: [number, () => Op[]][] = [
    [30, () => setCellInput(wb, s, row, col, pick(random, TEXTS))],
    [5, () => formatRange(wb, s, { top: row, left: col, bottom: row + 1, right: col }, { bold: true })],
    [3, () => setColumnWidth(wb, s, col, col, int(random, 4, 30))],
    [3, () => setRowHeight(s, row, row, int(random, 10, 40))],
    [9, () => [{ type: 'insert', sheet: s, axis: 'row', at: row, count: int(random, 1, 3) }]],
    [12, () => deleteRows(wb, s, row, int(random, 1, 3))],
    [6, () => [{ type: 'insert', sheet: s, axis: 'col', at: col, count: int(random, 1, 2) }]],
    [9, () => deleteColumns(wb, s, col, int(random, 1, 2))],
    [5, () => sortRange(wb, s, { top: 1, left: 0, bottom: int(random, 3, 9), right: 3 }, [{ col: int(random, 0, 3), descending: random() < 0.5 }])],
  ];
  if (sheets) {
    kinds.push([3, () => insertSheet(wb, int(random, 0, wb.sheets.length), { name: `New ${int(random, 0, 99)}` })]);
    kinds.push([3, () => renameSheet(wb, s, `Name ${int(random, 0, 99)}`)]);
    kinds.push([2, () => moveSheet(wb, s, int(random, 0, wb.sheets.length - 1))]);
  }
  let total = 0;
  for (const kind of kinds) total += kind[0];
  let roll = random() * total;
  for (const kind of kinds) {
    roll -= kind[0];
    if (roll < 0) return kind[1];
  }
  return kinds[0][1];
}

async function run(seed: number, bytes: Uint8Array, beside: boolean, sheets = false): Promise<boolean> {
  const net = new Net();
  const member = await join(net, 'p', bytes);
  const other = beside ? await join(net, 'q', bytes) : undefined;
  const plain = await (await open(bytes)).workbook();
  const history = new History(plain);
  const order = seeded(seed * 31);
  net.drain(order);
  const ours = seeded(seed);
  const theirs = seeded(seed);
  const dice = seeded(seed * 17);
  const steps = int(dice, 4, 40);
  const told: string[] = [];
  for (let i = 0; i < steps; i++) {
    const roll = dice();
    let what: string;
    if (roll < 0.55) {
      what = 'change';
      const made = member.session.stats.made;
      const took = act(member, make(member.wb, ours, sheets));
      const same = make(plain, theirs, sheets);
      let ops: Op[] = [];
      try {
        ops = same();
      } catch (error) {
        if (!(error instanceof RangeError)) throw error;
      }
      // What the workbook refuses, or what changes nothing, is no step on either side.
      if (took && member.session.stats.made > made) history.apply(ops, undefined, undefined);
    } else if (roll < 0.85) {
      what = 'undo';
      const a = member.session.undo();
      const b = history.undo();
      what += `(${a ? 1 : 0}${b ? 1 : 0})`;
    } else {
      what = 'redo';
      const a = member.session.redo();
      const b = history.redo();
      what += `(${a ? 1 : 0}${b ? 1 : 0})`;
    }
    told.push(what);
    if (other && dice() < 0.5) {
      // Someone else writes on the other sheet, which nothing here reads rows or columns of.
      const row = int(dice, 3, 9);
      const text = `q ${i}`;
      act(other, () => setCellInput(other.wb, 1, row, 5, text));
      for (const op of setCellInput(plain, 1, row, 5, text)) applyOp(plain, op);
    }
    net.drain(order);
    // The one thing that is not as the history has it, and is said: a sheet that was taken away (also by undoing its
    // insertion) cannot be brought back, and a name that a formula reads already cannot be taken back.
    if (sheets && member.notices.length > 0 && member.notices.every((notice) => notice.code === 'undo-none')) return false;
    {
      const a = canon(plain);
      const b = canon(member.wb);
      if (a !== b) assert.fail(`seed ${seed}${beside ? ' (beside someone)' : ''}: said ${member.notices.map((notice) => notice.code).join()}; after ${told.join(', ')} the workbook is not what the history leaves ${firstDifference(a, b)}`);
    }
  }
  net.drain(order);
  assert.equal(canon(member.wb), canon(plain), `seed ${seed}: the same at the end`);
  assert.deepEqual(member.notices.map((notice) => notice.code), [], `seed ${seed}${beside ? ' (beside someone)' : ''}: nothing is said after ${told.join(', ')}`);
  return true;
}

test(`alone in a room, every run of changes, undo and redo leaves what the history leaves, and nothing is said (${SEEDS} runs)`, async () => {
  const bytes = await startFile();
  for (let seed = 1; seed <= SEEDS; seed++) await run(seed, bytes, false);
});

test(`the same with sheets that are added, renamed and moved; all that is ever said is that a sheet cannot be brought back (${SEEDS} runs)`, async () => {
  const bytes = await startFile();
  let whole = 0;
  for (let seed = 1; seed <= SEEDS; seed++) if (await run(seed, bytes, false, true)) whole++;
  assert.ok(whole > SEEDS * 0.6, `most runs are as the history has them from end to end: ${whole} of ${SEEDS}`);
});

test(`the same beside someone who writes on another sheet (${SEEDS} runs)`, async () => {
  const bytes = await startFile();
  for (let seed = 1; seed <= SEEDS; seed++) await run(seed, bytes, true);
});

test('alone: two deletions of rows and two undos put the rows back in their order; a formula comes back as it was written', async () => {
  const bytes = await startFile();
  const net = new Net();
  const member = await join(net, 'p', bytes);
  const random = seeded(1);
  net.drain(random);
  const label = (): string[] => Array.from({ length: 10 }, (_, r) => String((member.wb.sheets[0].rows.get(r) || { cells: [] }).cells[0]?.value ?? ''));
  const before = label();
  for (const step of [() => deleteRows(member.wb, 0, 8, 1), () => deleteRows(member.wb, 0, 7, 1)]) {
    act(member, step);
    net.drain(random);
  }
  member.session.undo();
  net.drain(random);
  member.session.undo();
  net.drain(random);
  assert.deepEqual(label(), before, 'the rows are back where they were');
  for (const step of [() => setCellInput(member.wb, 0, 2, 4, '=$C$4+1'), () => setCellInput(member.wb, 0, 2, 4, 'text'), () => deleteColumns(member.wb, 0, 2, 1)]) {
    act(member, step);
    net.drain(random);
  }
  member.session.undo();
  net.drain(random);
  member.session.undo();
  net.drain(random);
  assert.equal(member.wb.sheets[0].rows.get(2)!.cells[4]!.formula!.text, '$C$4+1', 'the formula reads the cell it read');
  assert.deepEqual(member.notices, [], 'and nothing is said');
});
