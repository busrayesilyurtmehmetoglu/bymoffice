// The second reference model (collab-doc-model2.ts) against the session: four kinds of rooms -
// with much undoing (everybody waits for the relay before undoing), mixed, without undoing, and
// with much undoing while the caret of everybody is followed through what the others do.
// COLLAB_SEEDS sets how many rooms of each kind are run (the default is small); COLLAB_FROM where
// they start.
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { run, type Cfg } from './collab-doc-model2';

const SEEDS = Math.max(1, Math.min(Number(process.env.COLLAB_SEEDS) || 40, 100000));
const FROM = Math.max(1, Number(process.env.COLLAB_FROM) || 1);

for (const mode of ['undo', 'mix', 'noundo', 'carets'] as const) {
  test(`the second model of a text, ${mode === 'undo' ? 'with much undoing' : mode === 'mix' ? 'undoing among changes that wait' : mode === 'noundo' ? 'without undoing' : 'with much undoing, and the caret of everybody followed through what the others do'}: ${SEEDS} rooms`, async (t) => {
    // The model follows carets where CARET is set (as the view does: movedSelection) and fails one that does not stay behind its character.
    const caretsBefore = process.env.CARET;
    if (mode === 'carets') process.env.CARET = '1';
    else delete process.env.CARET;
    const bad: string[] = [];
    const sum = { made: 0, placed: 0, dropped: 0, quiet: 0, exact: 0, weak: 0, undoAsked: 0, undoDone: 0, caretMoves: 0, caretOk: 0, caretLost: 0 };
    const notices: Record<string, number> = {};
    for (let seed = FROM; seed < FROM + SEEDS; seed++) {
      const v = seed % 3;
      const cfg: Cfg = { seed: seed * 7907 + 13, members: 2 + (seed % 5), ticks: 220, deliver: v === 0 ? 0.6 : v === 1 ? 0.25 : 0.45, cuts: v === 2, undo: mode === 'noundo' ? 0 : mode === 'mix' ? 0.12 : 0.2, quiesce: mode === 'undo' || mode === 'carets', fmt: true };
      const result = await run(cfg);
      for (const key of Object.keys(sum) as (keyof typeof sum)[]) sum[key] += result[key];
      for (const code of Object.keys(result.notices)) notices[code] = (notices[code] || 0) + result.notices[code];
      if (result.bad.length > 0) bad.push(`seed ${seed}: ${result.bad.slice(0, 3).join(' | ')}`);
    }
    if (caretsBefore === undefined) delete process.env.CARET;
    else process.env.CARET = caretsBefore;
    if (mode === 'carets') assert.ok(sum.caretOk > 0, 'carets were followed');
    t.diagnostic(`${JSON.stringify(sum)} said: ${JSON.stringify(notices)}`);
    assert.deepEqual(bad.slice(0, 5), [], `${bad.length} rooms in which a change did something else than it meant`);
    assert.ok(sum.placed > 0 && sum.exact + sum.weak > 0, 'changes were made and checked');
    // (In a room or two of changes that wait, nothing may have been undone: undoing waits for them.)
    if (mode !== 'noundo' && SEEDS >= 10) assert.ok(sum.undoDone > 0, 'steps were undone');
  });
}
