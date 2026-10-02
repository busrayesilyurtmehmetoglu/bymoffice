// Property tests of the document editing commands: random documents, random
// commands, and what must hold whatever they are.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { DocHistory, insertText, resolvePosition, type DocEdit, type DocRange } from '../src/ops/doc/index';
import { assertSameIdentities, checkInvariants, checkLayout, checkWrappers, dice, docPicture, freezeBlocks, identities, randomDoc, randomEdit, undeletablesByStory } from './doc-ops-fixtures';

const SEEDS = Number(process.env.DOC_OPS_SEEDS || 30);
const STEPS = Number(process.env.DOC_OPS_STEPS || 200);

test('generated documents are laid out with every source position resolving to the text it shows', async () => {
  let checked = 0;
  for (let seed = 1; seed <= 40; seed++) {
    const doc = await randomDoc(seed);
    checkInvariants(doc);
    checkWrappers(doc);
    checked += checkLayout(doc);
  }
  assert.ok(checked > 500, `only ${checked} text items were checked`);
});

test('random commands keep the model valid, undo to the very objects and redo to the same state', async () => {
  const counts = new Map<string, number>();
  let refused = 0;
  let merged = 0;
  let controls = 0;
  for (let seed = 1; seed <= SEEDS; seed++) {
    const doc = await randomDoc(seed);
    const d = dice(seed * 7919);
    const before = docPicture(doc);
    const objects = identities(doc);
    // No command may change a block in place: everything the document holds is frozen, and what edits put in is
    // frozen as soon as it is in.
    freezeBlocks(doc);
    if (doc.body.blocks.some((block) => block.wrap !== undefined)) controls++;
    const history = new DocHistory<DocRange | undefined>(doc, 100000);
    // The picture of the document after each edit that was applied; undo steps may hold several of them.
    const pictures: string[] = [before];
    const apply = (name: string, edit: DocEdit, seal: boolean, where: string): void => {
      const kept = undeletablesByStory(doc);
      history.applyEdit(edit, undefined);
      if (seal) history.seal();
      counts.set(name, (counts.get(name) || 0) + 1);
      try {
        freezeBlocks(doc);
        checkInvariants(doc);
        checkWrappers(doc);
        // In every story that is still there: a note goes with its last reference, and what it held with it.
        undeletablesByStory(doc).forEach((now, story) => {
          if (kept.has(story)) assert.equal(now, kept.get(story), `${story}: something that is not content was removed or added`);
        });
        resolvePosition(doc, edit.selection.anchor);
        resolvePosition(doc, edit.selection.focus);
      } catch (error) {
        throw new Error(`${where} after ${name}: ${error instanceof Error ? error.stack : error}`);
      }
      pictures.push(docPicture(doc));
      assert.equal(history.canUndo, true);
    };
    for (let step = 0; step < STEPS; step++) {
      const where = `seed ${seed} step ${step}`;
      let made;
      // Building a command leaves the document as it was: the very blocks in every list, and (every tenth step,
      // for the time it takes) every value in it.
      const lists = identities(doc);
      const snapshot = step % 10 === 0 ? docPicture(doc) : undefined;
      try {
        made = randomEdit(doc, d);
      } catch (error) {
        // A command may refuse (RangeError); it must not have changed anything.
        assert.ok(error instanceof RangeError, `${where}: ${error instanceof Error ? error.stack : error}`);
        refused++;
        assertSameIdentities(identities(doc), lists);
        if (snapshot !== undefined) assert.equal(docPicture(doc), snapshot, `${where}: a refused command changed the model`);
        continue;
      }
      if (!made) continue;
      assertSameIdentities(identities(doc), lists);
      if (snapshot !== undefined) assert.equal(docPicture(doc), snapshot, `${where}: building ${made.name} changed the model`);
      if (made.edit.ops.length === 0) continue;
      // Most steps are steps of their own; some are left open, so that typing that follows merges into them.
      const typed = made.edit.mergeKey !== undefined;
      apply(made.name, made.edit, !typed && d.chance(0.8), where);
      if (typed) {
        // A few more characters at the caret: they join the step of the first one, up to a space.
        let caret = made.edit.selection.focus;
        const more = d.int(5);
        for (let i = 0; i < more; i++) {
          const edit = insertText(doc, caret, d.pick(['a', 'b', ' ', 'é', '😀']));
          apply('typing', edit, false, where);
          caret = edit.selection.focus;
          merged++;
        }
        if (d.chance(0.5)) history.seal();
      }
      if (step % 25 === 24) {
        try {
          checkLayout(doc);
        } catch (error) {
          throw new Error(`${where}: ${error instanceof Error ? error.stack : error}`);
        }
      }
    }
    checkLayout(doc);
    const last = pictures.length - 1;
    // Undo step by step: each state is one that was there after an edit, further back each time, and the last one
    // is the document as it was read.
    const visited: number[] = [];
    let at = last;
    while (history.canUndo) {
      assert.ok(history.undo());
      const picture = docPicture(doc);
      let found = at - 1;
      while (found >= 0 && pictures[found] !== picture) found--;
      assert.ok(found >= 0, `seed ${seed}: undo from state ${at} gave a state that never was`);
      at = found;
      visited.push(at);
    }
    assert.equal(docPicture(doc), before, `seed ${seed}: undoing everything`);
    assertSameIdentities(identities(doc), objects);
    checkInvariants(doc);
    checkWrappers(doc);
    // Steps were merged: fewer undo steps than edits, when anything was typed on.
    assert.ok(visited.length <= last);
    // Redo goes forward through the same states.
    for (let i = visited.length - 2; i >= -1; i--) {
      assert.ok(history.redo());
      assert.equal(docPicture(doc), pictures[i < 0 ? last : visited[i]], `seed ${seed}: redo`);
    }
    assert.equal(history.canRedo, false);
    while (history.undo());
    assert.equal(docPicture(doc), before);
    assertSameIdentities(identities(doc), objects);
  }
  // Every command was exercised.
  assert.ok(counts.size >= 19, `commands applied: ${Array.from(counts.keys()).join(', ')}`);
  assert.ok(refused < SEEDS * STEPS * 0.25, `${refused} commands were refused`);
  if (SEEDS >= 10) {
    assert.ok(merged > 20, `only ${merged} characters were typed on`);
    assert.ok(controls >= 3, `only ${controls} documents had content controls around blocks`);
  }
});
