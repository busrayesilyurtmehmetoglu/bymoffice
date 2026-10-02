// Property tests of the document editing commands: random documents, random
// commands, and what must hold whatever they are.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { DocHistory, resolvePosition, type DocRange } from '../src/ops/doc/index';
import { assertSameIdentities, checkInvariants, checkLayout, dice, docPicture, identities, randomDoc, randomEdit } from './doc-ops-fixtures';

const SEEDS = Number(process.env.DOC_OPS_SEEDS || 30);
const STEPS = Number(process.env.DOC_OPS_STEPS || 200);

test('generated documents are laid out with every source position resolving to the text it shows', async () => {
  let checked = 0;
  for (let seed = 1; seed <= 40; seed++) {
    const doc = await randomDoc(seed);
    checkInvariants(doc);
    checked += checkLayout(doc);
  }
  assert.ok(checked > 500, `only ${checked} text items were checked`);
});

test('random commands keep the model valid, undo to the very objects and redo to the same state', async () => {
  const counts = new Map<string, number>();
  let refused = 0;
  for (let seed = 1; seed <= SEEDS; seed++) {
    const doc = await randomDoc(seed);
    const d = dice(seed * 7919);
    const before = docPicture(doc);
    const objects = identities(doc);
    const history = new DocHistory<DocRange | undefined>(doc, 100000);
    const pictures: string[] = [before];
    let applied = 0;
    for (let step = 0; step < STEPS; step++) {
      let made;
      const snapshot = step % 10 === 0 ? docPicture(doc) : undefined;
      try {
        made = randomEdit(doc, d);
      } catch (error) {
        // A command may refuse (RangeError); it must not have changed anything.
        assert.ok(error instanceof RangeError, `seed ${seed} step ${step}: ${error instanceof Error ? error.stack : error}`);
        refused++;
        if (snapshot !== undefined) assert.equal(docPicture(doc), snapshot, `seed ${seed} step ${step}: a refused command changed the model`);
        continue;
      }
      if (!made) continue;
      // Building a command leaves the document as it was.
      if (snapshot !== undefined) assert.equal(docPicture(doc), snapshot, `seed ${seed} step ${step}: building ${made.name} changed the model`);
      if (made.edit.ops.length === 0) continue;
      history.applyEdit(made.edit, undefined);
      history.seal();
      applied++;
      counts.set(made.name, (counts.get(made.name) || 0) + 1);
      try {
        checkInvariants(doc);
        resolvePosition(doc, made.edit.selection.anchor);
        resolvePosition(doc, made.edit.selection.focus);
        if (step % 25 === 24) checkLayout(doc);
      } catch (error) {
        throw new Error(`seed ${seed} step ${step} after ${made.name}: ${error instanceof Error ? error.stack : error}`);
      }
      pictures.push(docPicture(doc));
    }
    checkLayout(doc);
    // Undo step by step: each state is the one that was there before, and the first is the document as read.
    for (let i = applied; i > 0; i--) {
      assert.ok(history.undo());
      assert.equal(docPicture(doc), pictures[i - 1], `seed ${seed}: undo of step ${i}`);
    }
    assert.equal(history.canUndo, false);
    assertSameIdentities(identities(doc), objects);
    checkInvariants(doc);
    for (let i = 1; i <= applied; i++) {
      assert.ok(history.redo());
      assert.equal(docPicture(doc), pictures[i], `seed ${seed}: redo of step ${i}`);
    }
    while (history.undo());
    assert.equal(docPicture(doc), before);
    assertSameIdentities(identities(doc), objects);
  }
  // Every command was exercised.
  assert.ok(counts.size >= 18, `commands applied: ${Array.from(counts.keys()).join(', ')}`);
  assert.ok(refused < SEEDS * STEPS * 0.2, `${refused} commands were refused`);
});
