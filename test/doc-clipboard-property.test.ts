// Property tests of the copy and paste of parts of documents: random
// documents, random ranges, random places, and what must hold whatever they
// are.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { DocHistory, deleteRange, expandRangeForTables, fragmentNeeds, fragmentOf, insertFragment, orderRange, resolvePosition, storyIds, type DocFragment, type DocRange } from '../src/ops/doc/index';
import { cutLoses, sliceOf } from '../src/ui/doc-clipboard';
import type { Paragraph } from '../src/model/doc/types';
import { assertSameIdentities, checkInvariants, checkLayout, checkWrappers, dice, docPicture, freezeBlocks, identities, paragraphAt, paragraphPaths, picture, randomDoc, randomPosition, texts, undeletablesByStory } from './doc-ops-fixtures';

/**
 * What a paragraph holds in its runs, item by item and character by
 * character, each with its formatting, its wrappers and whether it is field
 * code: equal for two paragraphs that differ only in how their text is cut
 * into runs. Only what a deletion can take and a caret can stand beside is
 * listed. What a deletion leaves where it is (a comment reference) and what
 * takes no offset and does not show (a soft hyphen that the range did not
 * reach over) may end up on the other side of the caret, as
 * it does for typing; the same goes for the marks and the code of a complex
 * field the end of whose result is cut: the caret that is left stands after
 * the field, and the text comes back there. (Fields that are cut whole and
 * pasted back are compared in doc-clipboard.test.ts.)
 */
const CONTENT = new Set<string>(['text', 'tab', 'break', 'noBreakHyphen', 'symbol', 'drawing', 'pict', 'object', 'positionalTab', 'pageNumber', 'footnoteReference', 'endnoteReference']);

function signature(paragraph: Paragraph): string[] {
  const out: string[] = [];
  for (const inline of paragraph.children) {
    if (inline.type !== 'run') continue;
    let wraps = '';
    for (let w = inline.wrap; w; w = w.parent) wraps += '@' + w.kind;
    const props = JSON.stringify(inline.props, (key, value) => (value && value.type === 'element' ? '<xml>' : value));
    inline.content.forEach((item, k) => {
      if (!CONTENT.has(item.type)) return;
      if (inline.fieldCode === true || (inline.fieldCode !== undefined && inline.fieldCode[k])) return;
      if (item.type === 'text') for (const ch of item.text) out.push(`${ch}${wraps}${props}`);
      else out.push(`<${item.type}>${wraps}${props}`);
    });
  }
  return out;
}

const SEEDS = Number(process.env.DOC_OPS_SEEDS || 30);
const STEPS = Number(process.env.DOC_CLIPBOARD_STEPS || 40);

/** Everything of a fragment frozen: a paste must not change what it is given. */
function freeze(value: unknown, depth = 0): void {
  if (value === null || typeof value !== 'object' || depth > 40 || Object.isFrozen(value) || value instanceof Uint8Array || value instanceof Map) return;
  Object.freeze(value);
  for (const key of Object.keys(value as object)) freeze((value as { [key: string]: unknown })[key], depth + 1);
}

test('random copies pasted at random places, in their document and in another: the model stays valid, nothing that is not content goes, undo brings the very objects back', async () => {
  let pasted = 0;
  let refused = 0;
  let abroad = 0;
  let tables = 0;
  let laid = 0;
  for (let seed = 1; seed <= SEEDS; seed++) {
    const doc = await randomDoc(seed);
    const other = await randomDoc(seed + 500);
    const d = dice(seed * 104729);
    const pictures = [docPicture(doc), docPicture(other)];
    const objects = [identities(doc), identities(other)];
    freezeBlocks(doc);
    freezeBlocks(other);
    const histories = [new DocHistory<DocRange | undefined>(doc, 100000), new DocHistory<DocRange | undefined>(other, 100000)];
    for (let step = 0; step < STEPS; step++) {
      const where = `seed ${seed} step ${step}`;
      const story = d.pick(storyIds(doc));
      const paths = paragraphPaths(doc, story);
      const range: DocRange = { anchor: randomPosition(doc, d, story, paths), focus: randomPosition(doc, d, story, paths) };
      let fragment: DocFragment;
      const copyBefore = docPicture(doc);
      try {
        fragment = fragmentOf(doc, expandRangeForTables(doc, range));
      } catch (error) {
        assert.ok(error instanceof RangeError, `${where}: ${error}`);
        continue;
      }
      assert.equal(docPicture(doc), copyBefore, `${where}: a copy changed the document`);
      freeze(fragment.blocks);
      if (fragment.tableDepth > 0) tables++;
      const toOther = d.chance(0.4);
      const target = toOther ? other : doc;
      const history = histories[toOther ? 1 : 0];
      const targetStory = d.pick(storyIds(target));
      const position = randomPosition(target, d, targetStory);
      const before = docPicture(target);
      const kept = undeletablesByStory(target);
      let changed = false;
      try {
        assert.deepEqual(fragmentNeeds(target, fragment), [], where);
        const edit = insertFragment(target, position, fragment);
        history.applyEdit(edit, undefined);
        history.seal();
        freezeBlocks(target);
        checkInvariants(target);
        checkWrappers(target);
        undeletablesByStory(target).forEach((now, id) => {
          if (kept.has(id)) assert.equal(now, kept.get(id), `${id}: something that is not content was removed or added`);
        });
        resolvePosition(target, edit.selection.anchor);
        changed = edit.ops.length > 0;
        if (edit.ops.length) pasted++;
        if (toOther && edit.ops.length) abroad++;
        if (step % 10 === 0) laid += checkLayout(target, false);
      } catch (error) {
        if (error instanceof RangeError) {
          // Refused: nothing changed.
          assert.equal(docPicture(target), before, `${where}: a refused paste changed the document`);
          refused++;
          continue;
        }
        throw new Error(`${where}: ${error instanceof Error ? error.stack : error}`);
      }
      // Most pastes are taken back, so that the documents do not only grow.
      if (changed && d.chance(0.7)) {
        history.undo();
        assert.equal(docPicture(target), before, `${where}: undo of a paste`);
      }
    }
    for (let i = 0; i < 2; i++) {
      while (histories[i].canUndo) histories[i].undo();
      const target = i === 0 ? doc : other;
      assert.equal(docPicture(target), pictures[i], `seed ${seed}: all undone`);
      assertSameIdentities(identities(target), objects[i]);
    }
  }
  assert.ok(pasted > SEEDS * STEPS * 0.5, `only ${pasted} pastes changed a document (${refused} refused)`);
  assert.ok(abroad > SEEDS, `only ${abroad} pastes went into another document`);
  assert.ok(tables > SEEDS / 3, `only ${tables} copies held a table`);
  assert.ok(laid > 100, `only ${laid} text items were checked`);
});

test('random cuts: one the copy covers is pasted back where it was and the text is what it was; undo brings the document back', async () => {
  let moved = 0;
  let kept = 0;
  let same = 0;
  let several = 0;
  for (let seed = 1; seed <= SEEDS; seed++) {
    const doc = await randomDoc(seed + 2000);
    const d = dice(seed * 15485863);
    const start = docPicture(doc);
    const objects = identities(doc);
    freezeBlocks(doc);
    const history = new DocHistory<DocRange | undefined>(doc, 100000);
    for (let step = 0; step < STEPS; step++) {
      const where = `seed ${seed} step ${step}`;
      const story = d.pick(storyIds(doc));
      const paths = paragraphPaths(doc, story);
      const a = randomPosition(doc, d, story, paths);
      // Mostly in one paragraph: there a cut and a paste give the paragraph back.
      const b = d.chance(0.6) ? randomPosition(doc, d, story, [a.path]) : randomPosition(doc, d, story, paths);
      const range: DocRange = { anchor: a, focus: b };
      const before = docPicture(doc);
      const text = texts(doc, story);
      const content = signature(paragraphAt(doc, a.path, story));
      // Of every paragraph of the story, inside tables too: what its runs hold and how it is formatted.
      const profile = (): string[] => paragraphPaths(doc, story).map((path) => picture(paragraphAt(doc, path, story).props) + ' ' + signature(paragraphAt(doc, path, story)).join(''));
      const paragraphs = profile();
      let loses: boolean;
      let fragment: DocFragment | undefined;
      try {
        const shown = orderRange(expandRangeForTables(doc, range));
        const slice = sliceOf(doc, shown.start, shown.end);
        fragment = slice ? slice.fragment : undefined;
        loses = cutLoses(doc, range, slice);
      } catch (error) {
        assert.ok(error instanceof RangeError, `${where}: ${error}`);
        assert.equal(docPicture(doc), before, where);
        continue;
      }
      assert.equal(docPicture(doc), before, `${where}: judging a cut changed the document`);
      if (loses || !fragment) {
        kept++;
        continue;
      }
      let steps = 0;
      try {
        const removal = deleteRange(doc, range);
        history.applyEdit(removal, undefined);
        history.seal();
        if (removal.ops.length) steps++;
        const paste = insertFragment(doc, removal.selection, fragment, { moved: true });
        history.applyEdit(paste, undefined);
        history.seal();
        if (paste.ops.length) steps++;
        freezeBlocks(doc);
        checkInvariants(doc);
        checkWrappers(doc);
      } catch (error) {
        throw new Error(`${where}: ${error instanceof Error ? error.stack : error}`);
      }
      moved++;
      const onePath = a.path.length === b.path.length && a.path.every((value, i) => value === b.path[i]);
      if (onePath) {
        assert.deepEqual(texts(doc, story), text, `${where}: the text after a cut and a paste at the same place`);
        // And everything else the runs hold: formatting, links, fields, breaks, pictures.
        assert.deepEqual(signature(paragraphAt(doc, a.path, story)), content, `${where}: the content after a cut and a paste at the same place`);
        same++;
      } else if (a.path.length === 1 && b.path.length === 1) {
        // From one paragraph of the story into another, whatever lies between them: every paragraph has its content
        // and its formatting again.
        assert.deepEqual(profile(), paragraphs, `${where}: the paragraphs after a cut over several of them and a paste at the same place`);
        several++;
      }
      // Both steps back: the document as it was.
      for (let i = 0; i < steps; i++) history.undo();
      assert.equal(docPicture(doc), before, `${where}: undo of a cut and a paste`);
      if (d.chance(0.3)) for (let i = 0; i < steps; i++) history.redo();
    }
    while (history.canUndo) history.undo();
    assert.equal(docPicture(doc), start, `seed ${seed}: all undone`);
    assertSameIdentities(identities(doc), objects);
  }
  assert.ok(moved > SEEDS * STEPS * 0.3, `only ${moved} cuts were pasted back (${kept} were refused)`);
  assert.ok(kept > SEEDS, `only ${kept} cuts were refused`);
  assert.ok(same > SEEDS * 3, `only ${same} cuts in one paragraph were compared`);
  assert.ok(several > SEEDS / 2, `only ${several} cuts over several paragraphs were compared`);
});
