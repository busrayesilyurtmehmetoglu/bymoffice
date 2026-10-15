// Dynamic array formulas edited together: the formula travels as a dynamic one, and every
// participant works out the same spill with its own engine.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine } from '../src/calc/engine';
import { decodeStep, encodeOp, encodeStep } from '../src/collab/wire';
import { contentVersion } from '../src/collab/protocol';
import { setCellInput } from '../src/ops/commands';
import { getCell, open } from '../src/index';
import { Net, act, assertSame, join, saveOf, seeded, settleAll, startFile, type Member } from './collab-fixtures';
import { formatRange } from '../src/ops/format';
import type { Op } from '../src/ops/ops';

test('a dynamic array formula spills alike for every participant', async () => {
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(11);
  const a = await join(net, 'a', bytes);
  const b = await join(net, 'b', bytes);
  net.drain(random);
  assert.equal(act(a, () => setCellInput(a.wb, 0, 0, 5, '=SEQUENCE(4)')), true);
  assert.equal(act(b, () => setCellInput(b.wb, 0, 0, 6, '=F1#*10')), true);
  assert.equal(act(b, () => setCellInput(b.wb, 0, 2, 7, 'x')), true);
  assert.equal(act(a, () => setCellInput(a.wb, 0, 0, 7, '=SEQUENCE(3)')), true);
  net.drain(random);
  assertSame([a, b], 'the formulas');
  assert.equal(getCell(b.wb.sheets[0], 0, 5)?.formula?.dynamic, true, 'b has the formula of a as a dynamic one');
  for (const member of [a, b]) {
    const engine = new CalcEngine(member.wb, { budget: { milliseconds: 0 } });
    engine.recalculate();
  }
  assertSame([a, b], 'the results');
  const column = (col: number) => [0, 1, 2, 3].map((row) => getCell(a.wb.sheets[0], row, col)?.value);
  assert.deepEqual(column(5), [1, 2, 3, 4]);
  assert.deepEqual(column(6), [10, 20, 30, 40]);
  assert.deepEqual(column(7), ['#SPILL!', undefined, 'x', undefined]);
});

test('the wire says whether a formula is a dynamic one, and takes nothing else for it', async () => {
  const file = await open(await startFile());
  const workbook = await file.workbook();
  const encoded = (text: string) => {
    const op = setCellInput(workbook, 0, 0, 0, text)[0];
    return JSON.parse(JSON.stringify(encodeStep([encodeOp(workbook, op).op!], workbook.styles)));
  };
  const typed = encoded('=SEQUENCE(2)');
  const formulaOf = (step: unknown[]) => JSON.stringify(step).match(/"f":\{[^}]*\}/)![0];
  assert.match(formulaOf(typed), /"d":1/);
  const decoded = decodeStep(typed, workbook.styles);
  assert.equal(JSON.stringify(decoded).indexOf('"d":1') >= 0, true);
  // Any other value of the flag makes the step one nobody takes.
  for (const bad of [2, true, '1']) {
    const step = JSON.parse(JSON.stringify(typed).replace('"d":1', `"d":${JSON.stringify(bad)}`));
    assert.throws(() => decodeStep(step, workbook.styles), undefined, `d: ${JSON.stringify(bad)}`);
  }
});

/** Each participant with an engine of its own, run after every change it hears of and every step it makes. */
function withEngines(members: Member[]) {
  const engines = new Map<Member, CalcEngine>();
  const heard = new Map<Member, number>();
  for (const member of members) {
    const engine = new CalcEngine(member.wb, { budget: { milliseconds: 0 } });
    engine.recalculate(engine.uncalculated());
    engines.set(member, engine);
    heard.set(member, member.changes.length);
  }
  const catchUp = () => {
    for (const member of members) {
      const engine = engines.get(member)!;
      for (let k = heard.get(member)!; k < member.changes.length; k++) {
        const change = member.changes[k] as { cells?: { sheet: number; row: number; col: number }[] };
        if (change.cells) engine.recalculate(change.cells);
        else {
          engine.invalidate();
          engine.recalculate(undefined, { overwrite: true });
        }
      }
      heard.set(member, member.changes.length);
    }
  };
  const local = (member: Member, make: () => Op[]) => {
    const ops = make();
    assert.equal(act(member, () => ops), true);
    const cells: { sheet: number; row: number; col: number }[] = [];
    for (const op of ops) if (op.type === 'setCells') for (const change of op.cells) cells.push({ sheet: op.sheet, row: change.row, col: change.col });
    engines.get(member)!.recalculate(cells);
    heard.set(member, member.changes.length);
  };
  return { engines, catchUp, local };
}

test('a format put on a cell of a spill: every participant, and one who joins later, has the spill', async () => {
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(3);
  const a = await join(net, 'a', bytes);
  const b = await join(net, 'b', bytes);
  net.drain(random);
  const { catchUp, local } = withEngines([a, b]);
  const values = (member: Member) => [10, 11, 12].map((row) => getCell(member.wb.sheets[0], row, 10)?.value);
  local(a, () => setCellInput(a.wb, 0, 10, 10, '=SEQUENCE(3)'));
  net.drain(random);
  catchUp();
  local(b, () => formatRange(b.wb, 0, { top: 11, left: 10, bottom: 11, right: 10 }, { bold: true }));
  net.drain(random);
  catchUp();
  assert.deepEqual([values(a), values(b)], [[1, 2, 3], [1, 2, 3]]);
  assertSame([a, b], 'after the format');
  const c = await join(net, 'c', bytes);
  net.drain(random);
  const late = withEngines([c]);
  late.catchUp();
  assert.deepEqual(values(c), [1, 2, 3]);
  assertSame([a, b, c], 'with the one who joined later');
});

test('a value typed into a spill and undone: every participant has the spill back', async () => {
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(5);
  const a = await join(net, 'a', bytes);
  const b = await join(net, 'b', bytes);
  net.drain(random);
  const { catchUp, local, engines } = withEngines([a, b]);
  const values = (member: Member) => [20, 21, 22].map((row) => getCell(member.wb.sheets[0], row, 12)?.value);
  local(a, () => setCellInput(a.wb, 0, 20, 12, '=SEQUENCE(3)'));
  net.drain(random);
  catchUp();
  local(b, () => setCellInput(b.wb, 0, 21, 12, 'x'));
  net.drain(random);
  catchUp();
  assert.deepEqual([values(a), values(b)], [['#SPILL!', 'x', undefined], ['#SPILL!', 'x', undefined]]);
  const undone = b.session.undo();
  // The undo is a step of its own: the one who undid calculates what it changed, as after any step.
  assert.ok(undone && undone.change.cells);
  engines.get(b)!.recalculate(undone.change.cells);
  net.drain(random);
  catchUp();
  assert.deepEqual([values(a), values(b)], [[1, 2, 3], [1, 2, 3]]);
  assertSame([a, b], 'after the undo');
});

/**
 * A random room (B7): three participants type dynamic array formulas and what they read, and undo; now and then one
 * saves the workbook and someone opens that file and joins, or joins with the file the room started from. Each
 * participant has an engine that calculates every change as it is told of it, as the view does (collabChanged()).
 * After every step every participant has the same values, spills (`spilled`, `ref`) and dynamic formulas.
 */
async function randomRoom(seed: number, steps: number): Promise<void> {
  const INPUTS = ['=SEQUENCE(3)', '=SEQUENCE(COUNT(A1:A6)+1)', '=A1#*2', '=SORT(B1:B4)', '=FILTER(A1:A6,A1:A6<>"")', '=A1:A3*1', '=TRANSPOSE(A1:A3)', '=UNIQUE(A1:A6)', '=MAP(A1:A3,LAMBDA(v,v*10))', '=SUM(C:C)', '=B2*2', '=ROWS(C1#)', '=SEQUENCE(1,3)', '=B1:C2', '5', '2', 'x', ''];
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(seed);
  const members: Member[] = [];
  const engines = new Map<Member, CalcEngine>();
  const calculate = (member: Member, change: { cells?: { sheet: number; row: number; col: number }[] }) => {
    const engine = engines.get(member)!;
    if (change.cells) engine.recalculate(change.cells);
    else {
      engine.invalidate();
      engine.recalculate(undefined, { overwrite: true });
    }
  };
  const add = async (name: string, file: Uint8Array) => {
    const member = await join(net, name, file);
    const engine = new CalcEngine(member.wb, { budget: { milliseconds: 0 } });
    engine.recalculate(engine.uncalculated());
    engines.set(member, engine);
    // Every change is calculated when it is told, the steps of the room that come in as one joins as well.
    const changes = member.changes;
    changes.push = (...items) => {
      for (const item of items) {
        Array.prototype.push.call(changes, item);
        calculate(member, item as { cells?: { sheet: number; row: number; col: number }[] });
      }
      return changes.length;
    };
    members.push(member);
    await settleAll(net, random, members);
  };
  for (const name of ['a', 'b', 'c']) await add(name, bytes);
  const log: string[] = [];
  let joined = 0;
  for (let k = 0; k < steps; k++) {
    const member = members[Math.floor(random() * members.length)];
    const what = random();
    const row = Math.floor(random() * 8);
    const col = Math.floor(random() * 5);
    const before = member.changes.length;
    if (what < 0.72) {
      const text = INPUTS[Math.floor(random() * INPUTS.length)];
      log.push(`${member.name}: R${row + 1}C${col + 1} ${text}`);
      let ops: Op[] = [];
      try {
        ops = setCellInput(member.wb, 0, row, col, text);
      } catch (error) {
        if (!(error instanceof RangeError)) throw error;
      }
      if (ops.length > 0 && act(member, () => ops) && member.changes.length === before) engines.get(member)!.recalculate([{ sheet: 0, row, col }]);
    } else if (what < 0.9) {
      log.push(`${member.name}: undo`);
      const undone = member.session.undo();
      if (undone && member.changes.length === before) calculate(member, undone.change);
    } else {
      const name = `j${++joined}`;
      if (random() < 0.5) {
        log.push(`${member.name} saves; ${name} opens that file and joins`);
        const saved = await saveOf(member);
        // As mount() does once the file is stored: the room knows the file, so that one who opens it may join.
        member.session.announceSaved([contentVersion(saved)], member.session.savableRev);
        await settleAll(net, random, members);
        await add(name, saved);
      } else {
        log.push(`${name} joins with the file the room started from`);
        await add(name, bytes);
      }
    }
    await settleAll(net, random, members);
    assert.deepEqual(members.map((m) => `${m.name} ${m.session.state}${m.session.state === 'online' ? '' : ' ' + JSON.stringify(m.notices)}`), members.map((m) => `${m.name} online`));
    assertSame(members, `seed ${seed}, after:\n${log.join('\n')}\n`);
  }
}

test('a random room with dynamic array formulas: every participant, and those who join later, has the same model (B7)', async () => {
  for (let seed = 1; seed <= 40; seed++) await randomRoom(seed, 30);
});
