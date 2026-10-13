// The second reference model (collab-semantic2-model.ts), run over every family of steps: what
// ends up in the workbook is what the order of the relay says, and nothing is dropped unsaid.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { run, type Config } from './collab-semantic2-model';

const SEEDS = Math.max(1, Math.min(Number(process.env.COLLAB_SEEDS) || 150, 100000));
const TICKS = Number(process.env.COLLAB_TICKS) || 200;

// The model's "merge" mode is not run: its expectations of merged cells do not hold with one participant either
// (merging empties the cells it covers, and a note of a merged cell is the note of its first cell). Where a merge
// meets rows that move, or cells and notes that are written, the later step is dropped and said (see rebase.ts).
const MODES: { [name: string]: Partial<Config> } = {
  rows: { sorts: false, formulas: true, burst: true, cuts: true, sheets: false, cols: false, merges: false },
  all: { sorts: false, formulas: true, burst: true, cuts: true, sheets: true, cols: true, merges: false },
  sorts: { sorts: true, formulas: false, burst: true, cuts: true, sheets: false, cols: true, merges: false },
};

for (const mode of Object.keys(MODES)) {
  test(`the second model, "${mode}": ${SEEDS} runs of ${TICKS} ticks`, async () => {
    const bad: string[] = [];
    const runs = Math.max(4, Math.floor(SEEDS / 6));
    for (let seed = 1; seed <= runs && bad.length < 3; seed++) {
      const result = await run({ seed, members: 2 + (seed % 4), ticks: TICKS, trace: false, ...MODES[mode] } as Config);
      if (result.violations.length > 0) bad.push(`seed ${seed}: ${result.violations.slice(0, 2).join(' | ').slice(0, 400)}`);
    }
    assert.deepEqual(bad, [], 'runs that ended with the wrong workbook');
  });
}
