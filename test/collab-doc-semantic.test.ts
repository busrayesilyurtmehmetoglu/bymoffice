// A document that is edited together against the reference model (collab-doc-model.ts): two to
// five participants type, delete, replace, press Enter, colour text and indent paragraphs, with
// long queues, connections that go away and messages delivered in any order. After every
// message and every change each participant has to see exactly what the model says it sees;
// after every step the relay placed, a participant that only follows has to hold exactly the
// text of that step; nothing is dropped that the model does not expect to be dropped, and
// what is dropped is said.
//
// Every character that is typed is one that was never typed before, so that the text itself
// says which character is which.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { run, type Config } from './collab-doc-run';

const SEEDS = Math.max(1, Math.min(Number(process.env.COLLAB_SEEDS) || 60, 100000));
const TICKS = Number(process.env.COLLAB_TICKS) || 220;

const MODES: { [name: string]: Partial<Config> } = {
  'short queues': { deliver: 0.6, cuts: false },
  'long queues': { deliver: 0.25, cuts: false },
  'connections that go away': { deliver: 0.45, cuts: true },
};

for (const mode of Object.keys(MODES)) {
  test(`the reference model of a text, ${mode}: ${SEEDS} runs of ${TICKS} ticks`, async () => {
    const bad: string[] = [];
    let made = 0;
    let dropped = 0;
    for (let seed = 1; seed <= SEEDS && bad.length < 3; seed++) {
      const result = await run({ seed: seed * 7919 + mode.length, members: 2 + (seed % 4), ticks: TICKS, deliver: 0.5, cuts: false, ...MODES[mode] });
      made += result.made;
      dropped += result.dropped;
      if (result.violations.length > 0) bad.push(`seed ${seed}: ${result.violations.slice(0, 2).join(' | ').slice(0, 900)}`);
    }
    assert.deepEqual(bad, [], 'runs that ended with another text than the one that was meant');
    assert.ok(made > SEEDS * 10, 'the runs made changes');
    void dropped;
  });
}

test('two people who write: nothing they type or delete is dropped', async () => {
  let made = 0;
  let dropped = 0;
  const bad: string[] = [];
  const runs = Math.max(4, Math.floor(SEEDS / 3));
  for (let seed = 1; seed <= runs && bad.length < 3; seed++) {
    const result = await run({ seed: seed * 104729, members: 2, ticks: TICKS, deliver: 0.5, cuts: seed % 2 === 0, plain: true });
    made += result.made;
    dropped += result.dropped;
    if (result.violations.length > 0) bad.push(`seed ${seed}: ${result.violations.slice(0, 2).join(' | ').slice(0, 900)}`);
  }
  assert.deepEqual(bad, []);
  assert.ok(made > runs * 20);
  assert.equal(dropped, 0, 'typing, deleting and Enter are never dropped');
});
