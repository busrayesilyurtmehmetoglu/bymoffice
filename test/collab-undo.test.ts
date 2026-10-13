// Undo while a workbook is edited together, with nobody else doing anything in between: a
// step and its undoing leave the workbook the history of a view that is not edited together
// leaves it - or the participant is told that it is not so.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { open } from '../src/index';
import { History, type Op } from '../src/ops/ops';
import { Net, canon, firstDifference, int, join, seeded, startFile } from './collab-fixtures';
import { change } from './collab-random.test';

const SEEDS = Math.max(1, Math.min(Number(process.env.COLLAB_SEEDS) || 150, 100000)) * 4;

test(`a step and its undoing leave what the history leaves, or it is said (${SEEDS} runs)`, async () => {
  const bytes = await startFile();
  const kinds: { [name: string]: [exact: number, said: number] } = {};
  for (let seed = 1; seed <= SEEDS; seed++) {
    const net = new Net();
    const member = await join(net, 'p', bytes);
    const plain = await (await open(bytes)).workbook();
    const history = new History(plain);
    net.drain(seeded(seed));
    // A few steps first, the same on both, so that the step that is undone meets a workbook with something in it.
    const before = int(seeded(seed * 7), 0, 4);
    const ours = seeded(seed);
    const theirs = seeded(seed);
    const onHistory = (make: () => Op[]): boolean => {
      try {
        const ops = make();
        // What does not travel is left out on both sides.
        if (ops.length === 0 || ops.some((op) => op.type === 'copySheet' || ((op.type === 'setDataValidations' || op.type === 'setConditionalFormats' || op.type === 'setDrawings') && !op.by))) return false;
        history.apply(ops, undefined, undefined);
        return true;
      } catch (error) {
        if (error instanceof RangeError) return false;
        throw error;
      }
    };
    let last: string | undefined;
    let same = true;
    for (let i = 0; i <= before && same; i++) {
      last = change(member, ours);
      const other = change({ wb: plain, name: 'p' }, theirs, onHistory);
      // A paste whose rules are left out is another step than the one of the history: such a run says nothing.
      if (last !== other || canon(member.wb) !== canon(plain)) same = false;
    }
    if (!same || !last) continue;
    net.drain(seeded(seed));
    const notices = member.notices.length;
    const undone = member.session.undo();
    net.drain(seeded(seed));
    history.undo();
    const said = member.notices.slice(notices).map((notice) => notice.code);
    const exact = canon(member.wb) === canon(plain);
    const count = kinds[last] || (kinds[last] = [0, 0]);
    if (exact && undone) count[0]++;
    else {
      count[1]++;
      if (said.length === 0) assert.fail(`seed ${seed}: "${last}" was ${undone ? 'undone' : 'not undone'} and nothing was said, but the workbook is not what the history leaves ${firstDifference(canon(plain), canon(member.wb))}`);
      assert.ok(said.every((code) => code === 'undo-none' || code === 'undo-lossy' || code === 'undo-gone'), `seed ${seed}: "${last}": what was said is about the undo: ${said.join(', ')}`);
    }
  }
  // The kinds of step that can be undone are undone exactly.
  for (const name of ['type', 'format', 'insert rows', 'insert columns', 'sort', 'rename', 'sheet+', 'note', 'merge', 'width', 'height', 'filter', 'move sheet', 'tab']) {
    const count = kinds[name];
    if (count) assert.ok(count[0] > 0, `"${name}" is undone exactly at least once (${count[0]} exact, ${count[1]} said)`);
  }
  for (const name of ['sheet-', 'validation', 'rule+', 'chart+']) {
    const count = kinds[name];
    if (count) assert.equal(count[0], 0, `"${name}" cannot be undone, and says so every time`);
  }
});
