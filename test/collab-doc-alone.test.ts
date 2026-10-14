// The one who edits a document alone in a room - or beside someone who only reads - has what
// the editor gives that is not in a room: after every change, undo and redo the document is
// the same, block for block and property for property, the file that is written holds the
// same document, and nothing is said. Changes of every kind that can be shared, in any order,
// with typing that is undone a word at a time.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CollabDocHistory, type DocHistoryLike } from '../src/collab/doc/history';
import { writeDocx } from '../src/formats/docx/write';
import { open } from '../src/index';
import { DocHistory, type DocEdit } from '../src/ops/doc/index';
import { aloneFile, make } from './collab-doc-changes';
import { DocNet, canonDoc, firstDifference, joinDoc, seeded } from './collab-doc-fixtures';

const SEEDS = Math.max(1, Math.min(Number(process.env.COLLAB_SEEDS) || 60, 100000));

async function run(seed: number, bytes: Uint8Array, beside: boolean, lazy: boolean): Promise<void> {
  const net = new DocNet();
  const member = await joinDoc(net, 'p', bytes);
  const other = beside ? await joinDoc(net, 'q', bytes, { readOnly: true }) : undefined;
  const plainFile = await open(bytes);
  const plain = await plainFile.document!();
  const order = seeded(seed * 31);
  net.drain(order);
  const shared: DocHistoryLike<string> = new CollabDocHistory<string>(member.session);
  const alone = new DocHistory<string>(plain);
  const random = seeded(seed);
  const label = (step: number): string => `seed ${seed}, step ${step}`;
  let made = 0;
  for (let step = 0; step < 60; step++) {
    const roll = random();
    let here: string | undefined;
    let there: string | undefined;
    const waited = member.session.pending > 0;
    const told = member.changes.length;
    if (roll < 0.17) {
      const a = alone.undo();
      const b = shared.undo();
      here = b ? b.meta : undefined;
      there = a ? a.meta : undefined;
    } else if (roll < 0.27) {
      const a = alone.redo();
      const b = shared.redo();
      here = b ? b.meta : undefined;
      there = a ? a.meta : undefined;
    } else {
      // Typing is one step of the history for a few characters: the same key, until something else is done.
      const key = roll < 0.5 ? 'typing' : undefined;
      const change = make(plain, seeded(seed * 1000 + step), true);
      const again = make(member.doc, seeded(seed * 1000 + step), true);
      let thrown: unknown;
      let edit: DocEdit | undefined;
      try {
        edit = change(plain);
      } catch (error) {
        if (!(error instanceof RangeError)) throw error;
        thrown = error;
      }
      let mine: DocEdit | undefined;
      try {
        mine = again(member.doc);
      } catch (error) {
        if (!(error instanceof RangeError)) throw error;
        assert.ok(thrown, `${label(step)}: only the document in the room refused the change: ${String(error)}`);
      }
      assert.equal(mine === undefined, edit === undefined, `${label(step)}: only the document outside the room refused the change`);
      if (edit && mine) {
        if (key === undefined) {
          alone.seal();
          shared.seal();
        }
        alone.apply(edit.ops, `before ${step}`, `after ${step}`, key);
        shared.apply(mine.ops, `before ${step}`, `after ${step}`, key);
        if (key === undefined) {
          alone.seal();
          shared.seal();
        }
        if (edit.ops.length > 0) made++;
      }
    }
    // Lazily: the relay answers when it likes, and undo has to wait for it.
    if (!lazy || order() < 0.5) net.drain(order);
    if (member.session.pending === 0) {
      // An undoing that had to wait for the relay hands back what it has when it is done.
      if (roll < 0.27 && waited && member.changes.length > told) here = member.changes[member.changes.length - 1].meta as string;
      assert.equal(here, there, `${label(step)}: what undo hands back (told: ${member.notices.map((notice) => notice.code).join()})`);
      const a = canonDoc(plain);
      const b = canonDoc(member.doc);
      if (a !== b) assert.fail(`${label(step)}: the document in the room is another one ${firstDifference(a, b)}`);
      assert.equal(shared.canUndo, alone.canUndo, `${label(step)}: whether there is something to undo`);
      assert.equal(shared.canRedo, alone.canRedo, `${label(step)}: whether there is something to make again`);
    } else if (roll < 0.27) {
      // An undoing that waits is done when the steps before it have their place: everything is delivered before the next step is compared.
      net.drain(order);
      const a = canonDoc(plain);
      const b = canonDoc(member.doc);
      if (a !== b) assert.fail(`${label(step)}: after an undoing that waited the document is another one ${firstDifference(a, b)}`);
    }
  }
  net.drain(order);
  assert.deepEqual(member.notices, [], `seed ${seed}: nothing is said to the one who edits alone`);
  assert.equal(member.session.stats.dropped, 0);
  if (other) {
    const a = canonDoc(member.doc);
    const b = canonDoc(other.doc);
    if (a !== b) assert.fail(`seed ${seed}: the one who reads holds another document ${firstDifference(a, b)}`);
    assert.deepEqual(other.notices, []);
  }
  // The files: the same parts with the same content, whichever history the document went through.
  await writeDocx(plainFile.package!, plain);
  await writeDocx(member.file.package!, member.doc);
  // (Not byte for byte: what an undoing put back is written from the model, not copied from the file that was read.)
  const one = canonDoc(await (await open(await plainFile.package!.save())).document!());
  const two = canonDoc(await (await open(await member.file.package!.save())).document!());
  if (one !== two) assert.fail(`seed ${seed}: the file that is written holds another document ${firstDifference(one, two)}`);
  assert.ok(made > 5, `seed ${seed}: changes were made`);
  member.session.leave();
  if (other) other.session.leave();
}

test(`alone in a room: every change, undo and redo leaves what the editor of one user leaves (${SEEDS} runs)`, async () => {
  const bytes = await aloneFile();
  for (let seed = 1; seed <= SEEDS; seed++) await run(seed, bytes, false, false);
});

test(`beside someone who reads, with a relay that answers when it likes (${SEEDS} runs)`, async () => {
  const bytes = await aloneFile();
  for (let seed = 1; seed <= SEEDS; seed++) await run(seed + 5000, bytes, true, true);
});

test('everything that was done is undone: the file that is written is the file that was read, as a document', async () => {
  const bytes = await aloneFile();
  for (let seed = 1; seed <= Math.max(4, Math.floor(SEEDS / 4)); seed++) {
    const net = new DocNet();
    const member = await joinDoc(net, 'p', bytes);
    net.drain();
    const history = new CollabDocHistory<string>(member.session);
    let made = 0;
    for (let step = 0; step < 25; step++) {
      try {
        const edit = make(member.doc, seeded(seed * 777 + step), true)(member.doc);
        history.seal();
        history.apply(edit.ops, '', '');
        if (edit.ops.length > 0) made++;
      } catch (error) {
        if (!(error instanceof RangeError)) throw error;
      }
      net.drain();
    }
    for (let i = 0; i < made; i++) {
      assert.ok(history.undo(), `seed ${seed}: step ${i} from the end is undone`);
      net.drain();
    }
    assert.equal(history.canUndo, false);
    await writeDocx(member.file.package!, member.doc);
    const fresh = canonDoc(await (await open(bytes)).document!());
    assert.equal(canonDoc(member.doc), fresh, `seed ${seed}: the document after everything was undone`);
    assert.equal(canonDoc(await (await open(await member.file.package!.save())).document!()), fresh, `seed ${seed}: the file after everything was undone`);
    assert.deepEqual(member.notices, []);
    member.session.leave();
  }
});
