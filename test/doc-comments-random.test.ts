// Random sequences of editing commands and commands of comments, with saves
// in between: comments are added (about ranges, at carets, in and across
// table cells), edited, answered, resolved and removed; text is typed, cut
// and pasted; tracking is turned on and off and tracked changes are accepted
// and rejected; now and then the file is saved, opened again and worked on.
//
// After every save the package is read the way a consumer that knows nothing
// of the engine reads it (see checkCommentedPackage and strict below): no
// mark without its comment, no comment without its reference, no range that
// ends before it begins, unique ids, a part of answers that names comments
// that exist, the ids that last in step with the comments; and reading it
// again gives the comments and the text the model holds. Undoing everything
// gives back the bytes that were opened.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { writeDocx } from '../src/formats/docx/write';
import type { DocModel } from '../src/model/doc/types';
import {
  DocHistory,
  acceptAll,
  addComment,
  deleteRange,
  editComment,
  fragmentOf,
  insertFragment,
  listComments,
  rejectAll,
  removeComment,
  replyToComment,
  resolveComment,
  setTrackRevisions,
  setTracking,
  type DocEdit,
  type DocFragment,
  type DocRange,
} from '../src/ops/doc/index';
import { OpcPackage } from '../src/opc';
import { cEnd, cRef, cStart, checkCommentedPackage, commentedSpec, sameCommentsAfterReading, type CommentedSpec } from './doc-comments-fixtures';
import { checkInvariants, dice, p, paragraphPaths, r, randomEdit, randomPosition, tbl, tc, texts, tr, type Dice } from './doc-ops-fixtures';
import { buildDocx as buildSpec } from './docx-fixtures';

const NOW = () => new Date('2026-10-12T08:09:10Z');
const decoder = new TextDecoder();
const hex = (n: number) => ('0000000' + n.toString(16).toUpperCase()).slice(-8);

/** Seven comments: a thread, a resolved one, one around a whole paragraph, two that overlap, one from a cell into the next. */
const BODY =
  p(r('One ') + cStart(1) + cStart(2) + r('two') + cEnd(1) + cRef(1) + cEnd(2) + cRef(2) + r(' three ') + cStart(3) + r('four') + cEnd(3) + cRef(3) + r(' five.')) +
  p(r('A second ') + cStart(5) + r('paragraph with') + cStart(6) + r(' several') + cEnd(5) + cRef(5) + r(' words') + cEnd(6) + cRef(6) + r(' in it.')) +
  tbl(tr(tc(p(r('Cell ') + cStart(7) + r('A'))) + tc(p(r('Cell B') + cEnd(7) + cRef(7)))) + tr(tc('Cell C') + tc('Cell D')), [2000, 2000]) +
  cStart(4) +
  p('A paragraph that a comment holds whole.') +
  cEnd(4) +
  p(cRef(4) + r('The last paragraph.'));
const IDS = [1, 2, 3, 4, 5, 6, 7];

/** Five kinds of files: no comments, comments alone, with the part of answers, with every part beside the comments, with tracking on. */
function specFor(kind: number): CommentedSpec {
  if (kind === 0) return { body: p('One two three four five.') + p('A second paragraph.') + tbl(tr(tc('Cell A') + tc('Cell B')) + tr(tc('Cell C') + tc('Cell D')), [2000, 2000]) + p('The last paragraph.') };
  const comments = IDS.map((id) => ({ id, author: 'A' + id, text: 'T' + id, paraId: kind === 1 ? undefined : hex(id + 9) }));
  if (kind === 1) return { body: BODY, comments };
  const extended = IDS.map((id): [string, string?, boolean?] => (id === 2 ? [hex(11), hex(10)] : id === 3 ? [hex(12), undefined, true] : [hex(id + 9)]));
  if (kind === 2) return { body: BODY, comments, extended };
  if (kind === 4) return { body: BODY, comments, extended, settings: '<w:trackRevisions/>' };
  return { body: BODY, comments, extended, ids: IDS.map((id): [string, string] => [hex(id + 9), hex(0x10000000 + id)]), extensible: [[hex(0x10000001), '2026-10-01T09:00:00Z'], [hex(0x10000003), '2026-10-01T09:00:00Z']] };
}

const TEXTS = ['Short.', 'Two\nlines', 'With <markup> & "quotes"', 'çığ 😀', '', 'tab\there'];

function commentEdit(doc: DocModel, d: Dice): { name: string; edit: DocEdit } {
  const list = listComments(doc);
  const name = list.length === 0 ? 'add' : d.pick(['add', 'add', 'addCaret', 'edit', 'reply', 'reply', 'resolve', 'resolve', 'remove']);
  const any = () => d.pick(list).id;
  // Paragraphs of the body and of the cells of its tables: a comment may begin in one cell and end in another.
  const paths = paragraphPaths(doc, 'body').filter((path) => path.indexOf(-1) < 0);
  switch (name) {
    case 'add':
      return { name, edit: addComment(doc, { anchor: randomPosition(doc, d, 'body', paths), focus: randomPosition(doc, d, 'body', paths) }, d.pick(TEXTS), d.chance(0.3) ? { author: 'Other One' } : undefined) };
    case 'addCaret':
      return { name, edit: addComment(doc, randomPosition(doc, d, 'body', paths), d.pick(TEXTS)) };
    case 'edit':
      return { name, edit: editComment(doc, any(), d.pick(TEXTS)) };
    case 'reply':
      return { name, edit: replyToComment(doc, any(), d.pick(TEXTS)) };
    case 'resolve':
      return { name, edit: resolveComment(doc, any(), d.chance(0.6)) };
    default:
      return { name, edit: removeComment(doc, any()) };
  }
}

/** Tracking on and off, every tracked change accepted or rejected, and the clipboard: a copy, a cut, a paste of what was copied. */
function otherEdit(doc: DocModel, d: Dice, clipboard: { fragment?: DocFragment }): { name: string; edit: DocEdit } | undefined {
  const paths = paragraphPaths(doc, 'body').filter((path) => path.indexOf(-1) < 0);
  const name = d.pick(['track', 'acceptAll', 'rejectAll', 'copy', 'cut', 'paste']);
  if (name === 'track') return { name, edit: setTrackRevisions(doc, !doc.settings.trackRevisions) };
  if (name === 'acceptAll') return { name, edit: acceptAll(doc) };
  if (name === 'rejectAll') return { name, edit: rejectAll(doc) };
  const selection = { anchor: randomPosition(doc, d, 'body', paths), focus: randomPosition(doc, d, 'body', paths) };
  if (name === 'copy') {
    clipboard.fragment = fragmentOf(doc, selection);
    return undefined;
  }
  if (name === 'cut') {
    clipboard.fragment = fragmentOf(doc, selection);
    return { name, edit: deleteRange(doc, selection) };
  }
  return clipboard.fragment ? { name, edit: insertFragment(doc, selection.anchor, clipboard.fragment) } : undefined;
}

/** The ids of the comments whose range ends, in the text of a part, before it begins. */
function endsBeforeStart(xml: string): string[] {
  const first = new Map<string, number>();
  const last = new Map<string, number>();
  const pattern = /<w:commentRange(Start|End) w:id="(\d+)"/g;
  for (let match = pattern.exec(xml); match; match = pattern.exec(xml)) {
    if (match[1] === 'Start') {
      if (!first.has(match[2])) first.set(match[2], match.index);
    } else last.set(match[2], match.index);
  }
  const bad: string[] = [];
  first.forEach((at, id) => {
    const end = last.get(id);
    if (end !== undefined && end < at) bad.push(id);
  });
  return bad;
}

/** What checkCommentedPackage does not look at: the order of the marks, and the parts of the ids that last. */
async function strict(bytes: Uint8Array): Promise<void> {
  await checkCommentedPackage(bytes);
  const pkg = await OpcPackage.open(bytes);
  const part = async (name: string) => (pkg.hasPart(name) ? decoder.decode(await pkg.readPart(name)) : undefined);
  const backwards = endsBeforeStart((await part('/word/document.xml')) as string);
  assert.deepEqual(backwards, [], 'the range of a comment ends before it begins');
  const comments = await part('/word/comments.xml');
  const lasts = new Set<string>();
  if (comments) {
    for (const one of comments.split('</w:comment>')) {
      const all = Array.from(one.matchAll(/w14:paraId="([0-9A-F]{8})"/g));
      if (all.length) lasts.add(all[all.length - 1][1]);
    }
  }
  const ids = await part('/word/commentsIds.xml');
  if (ids) {
    const durable = new Set<string>();
    const named = new Set<string>();
    for (const entry of ids.matchAll(/w16cid:paraId="([0-9A-Fa-f]{8})" w16cid:durableId="([0-9A-Fa-f]{8})"/g)) {
      assert.ok(!durable.has(entry[2]), `the durable id ${entry[2]} is given twice`);
      assert.ok(!named.has(entry[1]), `commentsIds.xml has two entries for ${entry[1]}`);
      assert.ok(lasts.has(entry[1]), `commentsIds.xml names ${entry[1]}, a comment that is gone`);
      durable.add(entry[2]);
      named.add(entry[1]);
    }
    lasts.forEach((id) => assert.ok(named.has(id), `the comment ${id} has no entry in commentsIds.xml`));
    const extensible = await part('/word/commentsExtensible.xml');
    if (extensible) for (const entry of extensible.matchAll(/w16cex:durableId="([0-9A-Fa-f]{8})"/g)) assert.ok(durable.has(entry[1]), `commentsExtensible.xml names the durable id ${entry[1]}, which is gone`);
  }
  const extended = await part('/word/commentsExtended.xml');
  if (extended && comments) lasts.forEach((id) => assert.ok(extended.indexOf(`paraId="${id}"`) >= 0, `the comment ${id} has no entry in commentsExtended.xml`));
}

interface Session {
  pkg: OpcPackage;
  doc: DocModel;
  bytes: Uint8Array;
  history: DocHistory<DocRange | undefined>;
}

async function openBytes(bytes: Uint8Array): Promise<Session> {
  const pkg = await OpcPackage.open(bytes);
  const doc = await readDocx(pkg);
  setTracking(doc, { author: 'Me', now: NOW });
  return { pkg, doc, bytes, history: new DocHistory<DocRange | undefined>(doc) };
}

for (let kind = 0; kind < 5; kind++) {
  test(`random edits and comments keep every saved package consistent (file kind ${kind})`, async () => {
    const counts: { [name: string]: number } = {};
    for (let seed = 1; seed <= 10; seed++) {
      const d = dice(seed * 977 + kind * 13);
      let o = await openBytes(await buildSpec(commentedSpec(specFor(kind))));
      const trail: string[] = [];
      const clipboard: { fragment?: DocFragment } = {};
      const where = () => `kind ${kind}, seed ${seed}, after ${trail.join(', ')}`;
      const save = async (): Promise<Uint8Array> => {
        await writeDocx(o.pkg, o.doc);
        const bytes = await o.pkg.save();
        try {
          await strict(bytes);
          const again = await sameCommentsAfterReading(bytes, o.doc);
          assert.deepEqual(texts(again), texts(o.doc), 'the text of the saved file is not the text of the model');
        } catch (error) {
          throw new Error(`${where()}: ${(error as Error).message}`);
        }
        return bytes;
      };
      /** Everything undone: the file that was opened (or opened again). */
      const undoAll = async (): Promise<void> => {
        while (o.history.canUndo) o.history.undo();
        await writeDocx(o.pkg, o.doc);
        assert.deepEqual(await o.pkg.save(), o.bytes, `${where()}: undoing everything did not give the file back`);
      };
      for (let step = 0; step < 60; step++) {
        const roll = d.next();
        if (roll < 0.1 && o.history.canUndo) {
          o.history.undo();
          trail.push('undo');
        } else if (roll < 0.15 && o.history.canRedo) {
          o.history.redo();
          trail.push('redo');
        } else if (roll < 0.19) {
          // Saved, everything undone (the file that was opened), and on with the file that was saved.
          const bytes = await save();
          await undoAll();
          o = await openBytes(bytes);
          trail.push('REOPEN');
        } else {
          let made: { name: string; edit: DocEdit } | undefined;
          try {
            made = roll < 0.55 ? commentEdit(o.doc, d) : roll < 0.7 ? otherEdit(o.doc, d, clipboard) : randomEdit(o.doc, d, ['body']);
          } catch (error) {
            // A place where the command does not apply: nothing was changed.
            if (!(error instanceof RangeError)) throw new Error(`${where()}: a command threw ${(error as Error).name}: ${(error as Error).message}`);
            continue;
          }
          if (!made || made.edit.ops.length === 0) continue;
          o.history.applyEdit(made.edit, undefined);
          o.history.seal();
          trail.push(made.name);
          counts[made.name] = (counts[made.name] || 0) + 1;
        }
        try {
          checkInvariants(o.doc, false);
        } catch (error) {
          throw new Error(`${where()}: ${(error as Error).message}`);
        }
        if (d.chance(0.3)) await save();
      }
      await save();
      await undoAll();
    }
    // Every kind of step was taken, so that the run says something about each.
    for (const name of ['add', 'addCaret', 'edit', 'reply', 'resolve', 'remove', 'track', 'cut', 'paste', 'insertText', 'deleteRange']) assert.ok((counts[name] || 0) > 0, `no "${name}" step was applied for file kind ${kind} (${JSON.stringify(counts)})`);
  });
}
