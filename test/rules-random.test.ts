// Random runs over everything that touches rules; see rules-random-fixtures.ts. After every
// save the file is read back: the rules it has are the rules of the model, where the model
// has them, and are looked up the same. A second workbook, read from the last save, is given
// the same steps: there the rules are those of a file, here many were changed since, and the
// two must never differ. Everything undone gives the file back.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { buildXlsx as buildFromModel } from '../src/formats/xlsx/build';
import { open } from '../src/index';
import { History } from '../src/ops/ops';
import { canonical } from './workbook-content';
import { checkPackage, checkSheetOrder, load, partNames, partText, reopen, save, sameBytes } from './rules-fixtures';
import { conditionalCover, randomBook, randomStep, restatedProblem, rulesBuilt, rulesNow as validationsNow, rulesSeen as validationsSeen, seeded, validationBehaviour } from './rules-random-fixtures';


async function run(seed: number, steps: number, rulesOnly: boolean, conditional = false, extensionOnly = false): Promise<void> {
  const random = seeded(seed);
  const doc = await load(await randomBook({ conditional, extensionOnly }));
  const original = { now: canonical(validationsNow(doc.wb)), seen: canonical(validationsSeen(doc.wb)) };
  assert.equal(restatedProblem(doc.wb), undefined, `seed ${seed}: the file as it was opened`);
  // The same steps on a workbook that was read from the last save: its rules are those of a file.
  let twin: { wb: Awaited<ReturnType<typeof reopen>>; history: History } | undefined;
  const trail: string[] = [];
  let applied = 0;
  for (let i = 0; i < steps; i++) {
    const step = randomStep(doc.wb, doc.history, random, { rulesOnly, conditional });
    if (!step) continue;
    trail.push(step.what);
    const where = `seed ${seed}, step ${i} (${trail.slice(-6).join(', ')})`;
    if (step.undo) {
      doc.history.undo();
      applied--;
      twin = twin && twin.history.canUndo ? twin : undefined;
      if (twin) twin.history.undo();
    } else if (step.redo) {
      doc.history.redo();
      applied++;
      twin = twin && twin.history.canRedo ? twin : undefined;
      if (twin) twin.history.redo();
    } else if (step.ops && step.ops.length > 0) {
      // What a step that sets or clears the rules of cells may change: those cells, and no other.
      const cells = step.cells;
      const inside = (key: string) => {
        const [row, col] = key.split(',').map(Number);
        return !cells || !cells.ranges || cells.ranges.some((range) => row >= range.top && row <= range.bottom && col >= range.left && col <= range.right);
      };
      const before = cells && cells.kind === 'validation' ? validationBehaviour(doc.wb, cells.sheet) : undefined;
      const covered = cells && cells.kind === 'conditionalClear' ? conditionalCover(doc.wb, cells.sheet) : undefined;
      doc.history.apply(step.ops, undefined, undefined);
      applied++;
      if (cells && before) {
        const after = validationBehaviour(doc.wb, cells.sheet);
        for (const [key, said] of before) if (!inside(key)) assert.equal(after.get(key), said, `${where}: the validation of cell ${key}, which was not touched, says something else`);
        for (const key of after.keys()) assert.ok(before.has(key) || inside(key), `${where}: cell ${key} got a validation it was not given`);
        // Worked out apart from the ranges of the rules: every cell that was named has the rule, or has none.
        for (const range of cells.ranges || []) {
          for (let row = range.top; row <= Math.min(range.bottom, 23); row++) for (let col = range.left; col <= Math.min(range.right, 17); col++) assert.equal(after.has(`${row},${col}`), cells.set === true, `${where}: cell ${row},${col}`);
        }
      }
      if (cells && covered) {
        const after = conditionalCover(doc.wb, cells.sheet);
        for (const [key, count] of covered) assert.equal(after.get(key) || 0, inside(key) ? 0 : count, `${where}: the rules of cell ${key}`);
        for (const key of after.keys()) assert.ok(covered.has(key), `${where}: cell ${key} got a rule by clearing`);
      }
      // The command is made again for the other workbook: an operation holds lists and sheets of the workbook it was made for.
      if (twin) twin.history.apply(step.build!(twin.wb), undefined, undefined);
    } else continue;
    if (twin) {
      assert.equal(canonical(validationsNow(twin.wb)), canonical(validationsNow(doc.wb)), `${where}: rules of the file and rules changed since differ`);
      assert.equal(canonical(validationsSeen(twin.wb)), canonical(validationsSeen(doc.wb)), `${where}: looked up differently`);
    }
    if (random() < 0.2 || i === steps - 1) {
      const saved = await save(doc);
      await checkPackage(saved);
      for (const name of partNames(saved)) if (/^xl\/worksheets\/[^/]*\.xml$/.test(name)) checkSheetOrder(await partText(saved, name), `${where}: ${name}`);
      const back = await reopen(saved);
      assert.equal(canonical(validationsNow(back)), canonical(validationsNow(doc.wb)), `${where}: the saved file has other rules`);
      assert.equal(canonical(validationsSeen(back)), canonical(validationsSeen(doc.wb)), `${where}: the saved file is looked up differently`);
      // A file built from the model says the same, as far as the main part of a sheet can say it.
      const built = await (await open(await buildFromModel(doc.wb))).workbook();
      assert.equal(canonical(rulesBuilt(built)), canonical(rulesBuilt(doc.wb)), `${where}: the built file says something else`);
      assert.equal(restatedProblem(back), undefined, `${where}: the saved file`);
      assert.equal(restatedProblem(doc.wb), undefined, `${where}: the model`);
      twin = { wb: back, history: new History(back) };
    }
  }
  while (doc.history.canUndo) doc.history.undo();
  assert.equal(canonical(validationsNow(doc.wb)), original.now, `seed ${seed}: everything undone`);
  assert.equal(canonical(validationsSeen(doc.wb)), original.seen, `seed ${seed}: everything undone, looked up`);
  const last = await save(doc);
  await checkPackage(last);
  if (rulesOnly) sameBytes(last, doc.bytes, `seed ${seed}: everything undone, the file`);
  else {
    // Cells and sheets that changed are written again for good; the rules are the ones of the file all the same.
    const rules = (xml: string) => (/<conditionalFormatting[\s\S]*<\/conditionalFormatting>/.exec(xml) || [''])[0] + (/<dataValidations[\s\S]*<\/dataValidations>/.exec(xml) || [''])[0];
    for (const name of partNames(doc.bytes)) if (/^xl\/worksheets\/[^/]*\.xml$/.test(name)) assert.equal(rules(await partText(last, name)), rules(await partText(doc.bytes, name)), `seed ${seed}: ${name}`);
    assert.equal(canonical(validationsNow(await reopen(last))), original.now, `seed ${seed}: everything undone, read back`);
  }
}

test('random runs, rules only: every save reads back as the model, and everything undone is the file byte for byte', async () => {
  for (let seed = 1; seed <= 40; seed++) await run(seed, 40, true);
});

test('random runs with rows, columns, cells and sheets: every save reads back as the model', async () => {
  for (let seed = 101; seed <= 160; seed++) await run(seed, 60, false);
});

test('random runs with conditional formats, rules only: every save reads back as the model, and everything undone is the file byte for byte', async () => {
  for (let seed = 201; seed <= 240; seed++) await run(seed, 40, true, true, seed % 4 === 0);
});

test('random runs with conditional formats and with rows, columns, cells and sheets: every save reads back as the model', async () => {
  for (let seed = 301; seed <= 360; seed++) await run(seed, 60, false, true, seed % 5 === 0);
});
