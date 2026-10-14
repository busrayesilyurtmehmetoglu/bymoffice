// A document in a room: the rules of two changes at one place, said one by one; undo beside
// someone else; whoever joins later; a participant that is not honest; and everything at once,
// at random, with the files that are saved read by a reader that is not this engine.

import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { join as joinPath } from 'node:path';
import { test } from 'node:test';
import { CollabDocHistory } from '../src/collab/doc/history';
import * as wireModule from '../src/collab/doc/wire';
import { MAX_LIST_BLOCKS, MAX_STORY_CELLS, MAX_STORY_CHARS_LIMIT, decodeStep, putDigest, putOf, runWire } from '../src/collab/doc/wire';
import { rebaseWire, type DocWire } from '../src/collab/doc/rebase';
import { paragraphText, type Block, type DocModel, type Paragraph } from '../src/model/doc/types';
import { contentVersion } from '../src/collab/protocol';
import { ZipArchive, open } from '../src/index';
import type { DocEdit } from '../src/ops/doc/index';
import { aloneFile, make, ops } from './collab-doc-changes';
import { buildDocx } from './docx-fixtures';
import { independent } from './one-go-fixtures';
import { CompositionHold, movedSelection, peerPlaces, placeNear } from '../src/ui/doc-collab';
import type { Mark } from '../src/collab/doc/rebase';
import { DocNet, actDoc, assertSameDoc, at, bodyText, canonDoc, docFile, firstDifference, joinDoc, lookOfDoc, saveDoc, seeded, type DocMember } from './collab-doc-fixtures';

const TBX = '<w:p><w:r><w:t>pre</w:t></w:r><w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="wps"><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="457200"/><wp:wrapNone/><wp:docPr id="1" name="T"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent><w:p w14:paraId="66666666"><w:bookmarkStart w:id="31" w:name="inbox"/><w:r><w:t>box</w:t></w:r><w:bookmarkEnd w:id="31"/></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:r></w:p>';

const SEEDS = Math.max(1, Math.min(Number(process.env.COLLAB_SEEDS) || 60, 100000));

async function pair(paragraphs: string[]): Promise<{ net: DocNet; a: DocMember; b: DocMember; bytes: Uint8Array }> {
  const bytes = await docFile(paragraphs);
  const net = new DocNet();
  const a = await joinDoc(net, 'a', bytes);
  const b = await joinDoc(net, 'b', bytes);
  net.drain();
  return { net, a, b, bytes };
}

const codes = (member: DocMember): string[] => member.notices.map((notice) => notice.code);

test('two who type at the same place: both texts stay, and the one that got its place later stands first', async () => {
  const { net, a, b } = await pair(['abcdef']);
  assert.ok(actDoc(a, [{ t: 'text', at: at(0, 3), x: 'XX' }]));
  assert.ok(actDoc(b, [{ t: 'text', at: at(0, 3), x: 'yy' }]));
  net.drain();
  assert.equal(bodyText(a.doc), bodyText(b.doc));
  // Whoever of the two the relay heard first: the other one's text stands before it.
  assert.ok(['abcXXyydef', 'abcyyXXdef'].indexOf(bodyText(a.doc)) >= 0, bodyText(a.doc));
  assertSameDoc([a, b], 'after both');
  assert.deepEqual(codes(a).concat(codes(b)), []);
});

test('a deletion deletes what its author saw: text someone put into the range stays', async () => {
  const { net, a, b } = await pair(['abcdef', 'second']);
  assert.ok(actDoc(a, [{ t: 'del', at: at(0, 1), to: at(1, 2) }]));
  assert.ok(actDoc(b, [{ t: 'text', at: at(0, 4), x: 'NEW' }]));
  net.drain();
  assert.equal(bodyText(a.doc), 'aNEWcond');
  assertSameDoc([a, b], 'after both');
  assert.deepEqual(codes(a).concat(codes(b)), []);
});

test('text that was typed into a range someone else deleted stays where the range was', async () => {
  const { net, a, b } = await pair(['abcdef']);
  assert.ok(actDoc(a, [{ t: 'text', at: at(0, 3), x: 'kept' }]));
  // The deletion is heard first.
  assert.ok(actDoc(b, [{ t: 'del', at: at(0, 1), to: at(0, 5) }]));
  net.drain();
  assert.equal(bodyText(a.doc), 'akeptf');
  assertSameDoc([a, b], 'after both');
});

test('Enter in a paragraph someone else types in: both halves are there, and the text is in the half it was typed in', async () => {
  const { net, a, b } = await pair(['abcdef']);
  assert.ok(actDoc(a, [{ t: 'split', at: at(0, 2) }]));
  assert.ok(actDoc(b, [{ t: 'text', at: at(0, 5), x: 'Z' }]));
  net.drain();
  assert.equal(bodyText(a.doc), 'ab\ncdeZf');
  assertSameDoc([a, b], 'after both');
});

test('two who delete ranges that overlap: what both deleted is deleted once, and the second one is no loss', async () => {
  const { net, a, b } = await pair(['abcdefgh']);
  assert.ok(actDoc(a, [{ t: 'del', at: at(0, 1), to: at(0, 5) }]));
  assert.ok(actDoc(b, [{ t: 'del', at: at(0, 3), to: at(0, 7) }]));
  net.drain();
  assert.equal(bodyText(a.doc), 'ah');
  const { net: net2, a: c, b: d } = await pair(['abcdefgh']);
  assert.ok(actDoc(c, [{ t: 'del', at: at(0, 1), to: at(0, 6) }]));
  assert.ok(actDoc(d, [{ t: 'del', at: at(0, 2), to: at(0, 4) }]));
  net2.drain();
  assert.equal(bodyText(c.doc), 'agh');
  assert.equal(c.session.stats.dropped + d.session.stats.dropped, 0);
  assert.equal(c.session.stats.quiet + d.session.stats.quiet, 1, 'one of the two had nothing left to delete');
  // Said all the same: should the other one take its deletion back, the text comes back whole.
  assert.deepEqual(codes(c).concat(codes(d)), ['partial']);
});

test('formatting covers what its author saw; a table that got a row and someone who typed in it do not go together', async () => {
  const bytes = await aloneFile();
  const net = new DocNet();
  const a = await joinDoc(net, 'a', bytes);
  const b = await joinDoc(net, 'b', bytes);
  net.drain();
  const history = new CollabDocHistory<string>(b.session);
  // The table is block 4 of the body: a row is put in, while someone types in one of its cells.
  history.apply(ops.insertRow(b.doc, { story: 'body', path: [4, 0, 0, 0], offset: 0 }, 'below').ops, '', '');
  assert.ok(actDoc(a, [{ t: 'text', at: { s: 'body', p: [4, 1, 0, 0], o: 0 }, x: 'typed' }]));
  // And someone types outside of the table, which goes together with anything in it.
  assert.ok(actDoc(a, [{ t: 'text', at: at(0, 0), x: 'outside ' }]));
  net.drain();
  assertSameDoc([a, b], 'after the row');
  const dropped = a.session.stats.dropped + b.session.stats.dropped;
  assert.equal(dropped, 1, 'the one that got its place second gave way');
  assert.equal(codes(a).concat(codes(b)).filter((code) => code === 'dropped').length, 1, 'which was said');
  assert.ok(bodyText(a.doc).indexOf('outside The first') === 0, 'what was typed outside of the table is there');
});

test('undo beside someone else: what nobody else touched is undone exactly, what someone changed since is not', async () => {
  const { net, a, b } = await pair(['first paragraph', 'second paragraph', 'third paragraph']);
  const before = canonDoc(a.doc);
  assert.ok(actDoc(a, [{ t: 'text', at: at(0, 5), x: ' typed' }], { before: 'a0', after: 'a1' }));
  net.drain();
  assert.ok(actDoc(b, [{ t: 'text', at: at(2, 0), x: 'B ' }]));
  assert.ok(actDoc(b, [{ t: 'split', at: at(2, 3) }]));
  net.drain();
  // Someone else changed another paragraph, and the number of paragraphs: the step is still undone as it was made.
  const undone = a.session.undo();
  assert.ok(undone, 'undone');
  assert.equal(undone.meta, 'a0');
  net.drain();
  assert.equal(bodyText(a.doc), 'first paragraph\nsecond paragraph\nB t\nhird paragraph');
  assertSameDoc([a, b], 'after undo');
  // And made again.
  assert.equal((a.session.redo() || { meta: undefined }).meta, 'a1');
  net.drain();
  assert.equal(bodyText(b.doc), 'first typed paragraph\nsecond paragraph\nB t\nhird paragraph');
  assertSameDoc([a, b], 'after redo');
  assert.ok(a.session.undo());
  net.drain();
  // Now someone else changes the very paragraph: the step that would be made again is of a paragraph that is no more.
  assert.ok(actDoc(b, [{ t: 'text', at: at(0, 0), x: '!' }]));
  net.drain();
  assert.equal(a.session.redo(), undefined, 'not made again');
  assert.deepEqual(codes(a), ['undo-gone'], 'and said');
  assert.equal(bodyText(a.doc), '!first paragraph\nsecond paragraph\nB t\nhird paragraph');
  assertSameDoc([a, b], 'at the end');
  void before;
});

test('undo is last in, first out: what someone else changed since is not taken back, and what lies under it still is when nothing moved', async () => {
  const { net, a, b } = await pair(['one', 'two', 'three']);
  assert.ok(actDoc(a, [{ t: 'text', at: at(1, 3), x: ' A1' }]));
  assert.ok(actDoc(a, [{ t: 'run', at: at(1, 0), to: at(1, 3), f: { bold: true } }]));
  net.drain();
  // Someone else types in the paragraph.
  assert.ok(actDoc(b, [{ t: 'text', at: at(1, 0), x: 'b' }]));
  net.drain();
  // The formatting would be taken back by putting the paragraph back as it was: not while it holds what someone else typed.
  assert.equal(a.session.undo(), undefined);
  assert.deepEqual(codes(a), ['undo-gone']);
  assert.equal(bodyText(a.doc), 'one\nbtwo A1\nthree');
  // The typing under it is taken back as "what was typed goes": that it can be, whoever typed beside it.
  assert.ok(a.session.undo(), 'the typing is undone');
  net.drain();
  assert.equal(bodyText(b.doc), 'one\nbtwo\nthree');
  assertSameDoc([a, b], 'after the undo');
  assert.ok(b.session.undo());
  net.drain();
  assert.equal(bodyText(a.doc), 'one\ntwo\nthree');
  assertSameDoc([a, b], 'at the end');
  assert.equal(a.session.canUndo, false);
});

test('a deletion that is undone comes back as it looked, for everybody, also for who joined with a file that was saved since', async () => {
  const bytes = await aloneFile();
  const net = new DocNet();
  const a = await joinDoc(net, 'a', bytes);
  const b = await joinDoc(net, 'b', bytes);
  net.drain();
  const before = canonDoc(a.doc);
  // The bold words and the plain ones after them, and the end of the paragraph before: two paragraphs become one.
  assert.ok(actDoc(a, [{ t: 'del', at: at(2, 0), to: at(3, 14) }]));
  net.drain();
  const stored = await saveDoc(b);
  b.session.announceSaved([contentVersion(stored)], b.session.rev);
  net.drain();
  const late = await joinDoc(net, 'late', stored);
  net.drain();
  assert.equal(late.session.rev, 1, 'it joined at the step of its file');
  assert.ok(a.session.undo(), 'undone');
  net.drain();
  assert.equal(canonDoc(a.doc), before, 'the document is what it was, block for block');
  assertSameDoc([a, b], 'the two who were there');
  assert.equal(late.session.state, 'online', 'who joined later takes the step like any other');
  assert.equal(lookOfDoc(late.doc), lookOfDoc(a.doc), 'and holds the same');
  assert.deepEqual(codes(a).concat(codes(b), codes(late)), []);
  // And made again, and undone again.
  assert.ok(a.session.redo());
  net.drain();
  assert.ok(a.session.undo());
  net.drain();
  assert.equal(canonDoc(b.doc), before);
  assert.equal(lookOfDoc(late.doc), lookOfDoc(a.doc));
});

test('what cannot be taken back by commands is said, and what lies under it can still be undone', async () => {
  const bytes = await aloneFile();
  const net = new DocNet();
  const a = await joinDoc(net, 'a', bytes);
  const b = await joinDoc(net, 'b', bytes);
  net.drain();
  const history = new CollabDocHistory<string>(a.session);
  const place = { story: 'body', path: [0], offset: 3 };
  history.apply(ops.insertText(a.doc, place, 'XY').ops, '', '');
  history.seal();
  // A list is made of the paragraph below, which moves nothing; a table is put in above, which moves everything below it.
  history.apply(ops.toggleList(a.doc, { anchor: { story: 'body', path: [1], offset: 0 }, focus: { story: 'body', path: [1], offset: 0 } }, 'bullet').ops, '', '');
  history.seal();
  net.drain();
  assert.equal(history.undo(), undefined, 'the list is not undone');
  assert.deepEqual(codes(a), ['undo-none']);
  assert.ok(history.undo(), 'the typing under it is');
  net.drain();
  assert.equal(bodyText(a.doc).split('\n')[0], 'The first paragraph of the file.');
  assertSameDoc([a, b], 'after the undo');
  // A table that was put in is taken out again by undo; it cannot be made again (a table does not travel in an undoing).
  history.apply(ops.insertTable(a.doc, { story: 'body', path: [0], offset: 3 }, 2, 2).ops, '', '');
  history.seal();
  history.apply(ops.insertText(a.doc, { story: 'body', path: [5], offset: 0 }, 'below ').ops, '', '');
  history.seal();
  net.drain();
  assert.ok(history.undo(), 'what was typed below the table');
  net.drain();
  assert.ok(history.undo(), 'the table');
  net.drain();
  assert.equal(bodyText(b.doc).split('\n')[0], 'The first paragraph of the file.');
  assertSameDoc([a, b], 'after the table was taken out');
  assert.equal(history.redo(), undefined);
  assert.deepEqual(codes(a), ['undo-none', 'undo-none']);
  assertSameDoc([a, b], 'at the end');
});

test('a step that waits behind one that is dropped is dropped with it, and both are said: nothing lands where it was not meant', async () => {
  // The paragraph is split and typed into while someone else takes back what stands in it.
  const { net, a, b } = await pair(['aaa', 'bbb']);
  assert.ok(actDoc(b, [{ t: 'text', at: at(1, 0), x: 'Z' }]));
  net.drain();
  a.transport.suspend();
  assert.ok(actDoc(a, [{ t: 'split', at: at(1, 2) }]));
  assert.ok(actDoc(a, [{ t: 'text', at: at(2, 1), x: 'X' }]));
  assert.ok(b.session.undo());
  net.drain();
  a.transport.resume();
  net.drain();
  net.drain();
  assert.equal(bodyText(a.doc), 'aaa\nb\nbXb', 'X stands between the two letters it was typed between');
  assertSameDoc([a, b], 'after all');
  // And with a step that does give way: a row is put into a table while someone types in it and, behind that, elsewhere.
  const bytes = await aloneFile();
  const net2 = new DocNet();
  const c = await joinDoc(net2, 'c', bytes);
  const d = await joinDoc(net2, 'd', bytes);
  net2.drain();
  c.transport.suspend();
  assert.ok(actDoc(c, [{ t: 'text', at: { s: 'body', p: [4, 1, 0, 0], o: 0 }, x: 'in the table' }]));
  assert.ok(actDoc(c, [{ t: 'text', at: at(5, 0), x: 'after it ' }]));
  assert.ok(actDoc(d, [{ t: 'cmd', name: 'insertRow', at: { s: 'body', p: [4, 0, 0, 0], o: 0 }, a: ['below'] }]));
  net2.drain();
  c.transport.resume();
  net2.drain();
  net2.drain();
  assert.equal(c.session.stats.dropped, 2, 'the typing in the table, and what was typed after it');
  assert.deepEqual(c.notices, [{ code: 'dropped', count: 2 }]);
  assert.ok(bodyText(c.doc).indexOf('after it') < 0 && bodyText(c.doc).indexOf('in the table') < 0);
  assertSameDoc([c, d], 'after the row');
});

test('a command about a place that someone else deleted is dropped and said; text typed there stays where the deletion was', async () => {
  const { net, a, b } = await pair(['one two three']);
  a.transport.suspend();
  assert.ok(actDoc(a, [{ t: 'cmd', name: 'insertPageBreak', at: at(0, 5), a: [] }]));
  assert.ok(actDoc(b, [{ t: 'del', at: at(0, 3), to: at(0, 8) }]));
  net.drain();
  a.transport.resume();
  net.drain();
  net.drain();
  assert.deepEqual(a.notices, [{ code: 'dropped', count: 1 }]);
  assert.equal(bodyText(a.doc), 'onethree');
  assertSameDoc([a, b], 'after the break was dropped');
});

test('an undoing that cannot be made leaves what lies under it in its stack, when its step moved the paragraphs below: none of it can be taken back any more', async () => {
  const { net, a, b } = await pair(['zero', 'one', 'two', 'three']);
  // First something typed in the last paragraph, then two paragraphs joined (which moves the last one up).
  assert.ok(actDoc(a, [{ t: 'text', at: at(3, 5), x: 'X' }]));
  a.session.seal();
  assert.ok(actDoc(a, [{ t: 'del', at: at(1, 3), to: at(2, 0) }]));
  net.drain();
  assert.equal(bodyText(a.doc), 'zero\nonetwo\nthreeX');
  // The join is taken back (the paragraph is put back as two) while someone else types in it: the undoing cannot be made.
  a.transport.suspend();
  assert.ok(a.session.undo());
  assert.ok(actDoc(b, [{ t: 'text', at: at(1, 0), x: 'b' }]));
  net.drain();
  a.transport.resume();
  net.drain();
  net.drain();
  assert.equal(bodyText(a.doc), 'zero\nbonetwo\nthreeX', 'the paragraphs stay joined');
  assert.deepEqual(codes(a), ['undo-gone'], 'an undoing that cannot be made is said as one');
  assertSameDoc([a, b], 'after the undoing was dropped');
  const before = canonDoc(a.doc);
  // What was typed in the last paragraph was kept for a document in which it stood a paragraph lower: it is not taken back from another place.
  assert.equal(a.session.undo(), undefined);
  net.drain();
  assert.equal(canonDoc(a.doc), before);
  assert.deepEqual(codes(a), ['undo-gone', 'undo-gone']);
});

test('while an undoing waits for the relay, what someone else does is read into the stack as the undoing leaves the document', async () => {
  const { net, a, b } = await pair(['abcdef', 'second']);
  // A letter deleted (taken back by putting the paragraph back), then the paragraph split.
  assert.ok(actDoc(a, [{ t: 'del', at: at(0, 4), to: at(0, 5) }]));
  a.session.seal();
  assert.ok(actDoc(a, [{ t: 'split', at: at(0, 1) }]));
  net.drain();
  assert.equal(bodyText(a.doc), 'a\nbcdf\nsecond');
  // The split is taken back while someone else types into the second half of it.
  a.transport.suspend();
  assert.ok(a.session.undo());
  assert.ok(actDoc(b, [{ t: 'text', at: at(1, 2), x: 'Z' }]));
  net.drain();
  a.transport.resume();
  net.drain();
  net.drain();
  assert.equal(bodyText(a.doc), 'abcZdf\nsecond');
  assertSameDoc([a, b], 'after the undoing');
  // The paragraph the deletion would be taken back in holds what the other one typed: it is not put back over it.
  assert.equal(a.session.undo(), undefined);
  net.drain();
  assert.equal(bodyText(b.doc), 'abcZdf\nsecond', 'what the other one typed stays');
  assert.deepEqual(codes(a), ['undo-gone']);
  assertSameDoc([a, b], 'at the end');
});

test('while an undoing waits, what lies under the step it makes in the other stack is read as the document was before it: nothing of someone else is written over', async () => {
  const { net, a, b } = await pair(['aaaa', 'bbbb', 'cccc', 'dddd']);
  assert.ok(actDoc(a, [{ t: 'split', at: at(0, 2) }]));
  a.session.seal();
  net.drain();
  assert.ok(actDoc(a, [{ t: 'text', at: at(3, 0), x: 'X' }]));
  a.session.seal();
  net.drain();
  assert.ok(a.session.undo());
  net.drain();
  // The Enter is taken back while someone else breaks the paragraph the X was typed in and types in it.
  a.transport.suspend();
  assert.ok(a.session.undo());
  assert.ok(actDoc(b, [{ t: 'text', at: at(3, 2), x: '\nZZZ' }]));
  net.drain();
  a.transport.resume();
  net.drain();
  net.drain();
  assert.equal(bodyText(a.doc), 'aaaa\nbbbb\ncc\nZZZcc\ndddd');
  assert.ok(a.session.redo(), 'the Enter is made again');
  net.drain();
  // The X would be made again by putting the paragraph back as it was: someone else changed it since.
  assert.equal(a.session.redo(), undefined);
  net.drain();
  assert.equal(bodyText(b.doc), 'aa\naa\nbbbb\ncc\nZZZcc\ndddd', 'what the other one typed stays, nothing is there twice');
  assert.deepEqual(codes(a), ['undo-gone']);
  assertSameDoc([a, b], 'at the end');
});

test('while a redoing waits, what lies under the step it makes in the other stack is read as the document was before it: the right letter is taken back', async () => {
  const { net, a, b } = await pair(['aaaa', 'bbbb', 'cccc', 'dddd']);
  assert.ok(actDoc(a, [{ t: 'text', at: at(2, 0), x: 'T' }]));
  a.session.seal();
  net.drain();
  assert.ok(actDoc(a, [{ t: 'split', at: at(0, 2) }]));
  a.session.seal();
  net.drain();
  assert.ok(a.session.undo());
  net.drain();
  a.transport.suspend();
  assert.ok(a.session.redo());
  assert.ok(actDoc(b, [{ t: 'text', at: at(1, 0), x: 'QQ' }]));
  net.drain();
  a.transport.resume();
  net.drain();
  net.drain();
  assert.equal(bodyText(a.doc), 'aa\naa\nQQbbbb\nTcccc\ndddd');
  assert.ok(a.session.undo());
  net.drain();
  assert.ok(a.session.undo());
  net.drain();
  assert.equal(bodyText(b.doc), 'aaaa\nQQbbbb\ncccc\ndddd', 'the T went, and nothing else');
  assert.deepEqual(codes(a), []);
  assertSameDoc([a, b], 'at the end');
});

test('alone: a step that is made again, something typed before the relay answered, and undo takes back what was typed', async () => {
  const net = new DocNet();
  const a = await joinDoc(net, 'a', await docFile(['aaa', 'bbb']));
  net.drain();
  assert.ok(actDoc(a, [{ t: 'text', at: at(0, 3), x: 'X' }]));
  a.session.seal();
  net.drain();
  assert.ok(a.session.undo());
  net.drain();
  assert.ok(a.session.redo());
  assert.ok(actDoc(a, [{ t: 'text', at: at(1, 3), x: 'Y' }]));
  net.drain();
  assert.equal(bodyText(a.doc), 'aaaX\nbbbY');
  assert.ok(a.session.undo());
  net.drain();
  assert.equal(bodyText(a.doc), 'aaaX\nbbb', 'as the editor of one user leaves it');
  assert.ok(a.session.undo());
  net.drain();
  assert.equal(bodyText(a.doc), 'aaa\nbbb');
  assert.deepEqual(a.notices, []);
});

test('a deletion that went with a deletion of someone else, who then takes that back: the text is back, and that the deletion was not made here was said', async () => {
  const { net, a, b } = await pair(['abcdef', 'xyz']);
  a.transport.suspend();
  assert.ok(actDoc(a, [{ t: 'del', at: at(0, 1), to: at(0, 3) }]));
  assert.ok(actDoc(b, [{ t: 'del', at: at(0, 0), to: at(0, 4) }]));
  net.drain();
  assert.ok(b.session.undo());
  net.drain();
  a.transport.resume();
  net.drain();
  net.drain();
  assert.equal(bodyText(a.doc), 'abcdef\nxyz');
  assert.deepEqual(a.notices, [{ code: 'partial', count: 1 }]);
  assertSameDoc([a, b], 'at the end');
});

test('a deletion or a colouring that went with a deletion of someone else is said: when the other one thinks again, the text is back without it', async () => {
  for (const kind of ['del', 'run'] as const) {
    const { net, a, b } = await pair(['one two three']);
    a.transport.suspend();
    assert.ok(actDoc(a, [kind === 'del' ? { t: 'del', at: at(0, 4), to: at(0, 8) } : { t: 'run', at: at(0, 4), to: at(0, 7), f: { color: 'FF0000' } }]));
    assert.ok(actDoc(b, [{ t: 'del', at: at(0, 3), to: at(0, 8) }]));
    net.drain();
    a.transport.resume();
    net.drain();
    net.drain();
    assert.equal(bodyText(a.doc), 'onethree');
    // Whatever of it could not be made is said: the other one may take its deletion back, and the text comes back without this change.
    assert.deepEqual(a.notices, [{ code: 'partial', count: 1 }], kind);
    assert.ok(b.session.undo());
    net.drain();
    assert.equal(bodyText(a.doc), 'one two three');
    assertSameDoc([a, b], kind);
  }
  // Part of it: the rest is done, and that part of it went is said as well.
  const { net, a, b } = await pair(['one two three']);
  a.transport.suspend();
  assert.ok(actDoc(a, [{ t: 'del', at: at(0, 2), to: at(0, 6) }]));
  assert.ok(actDoc(b, [{ t: 'del', at: at(0, 4), to: at(0, 8) }]));
  net.drain();
  a.transport.resume();
  net.drain();
  net.drain();
  assert.equal(bodyText(a.doc), 'onthree');
  assert.deepEqual(a.notices, [{ code: 'partial', count: 1 }]);
  assertSameDoc([a, b], 'partly');
});

test('a line break, a tab and a page break travel; what would make a document too large does not', async () => {
  const { net, a, b } = await pair(['abc', 'def']);
  const history = new CollabDocHistory<string>(a.session);
  history.apply(ops.insertText(a.doc, { story: 'body', path: [0], offset: 1 }, 'x\u000by\tz').ops, '', '');
  history.seal();
  history.apply(ops.insertPageBreak(a.doc, { story: 'body', path: [1], offset: 1 }).ops, '', '');
  net.drain();
  assertSameDoc([a, b], 'after the breaks');
  const items = (b.doc.body.blocks[0] as { children: { content?: { type: string; kind?: string }[] }[] }).children.reduce((all: string[], run) => all.concat((run.content || []).map((item) => item.type + (item.kind ? ':' + item.kind : ''))), []);
  assert.deepEqual(items.filter((item) => item !== 'text'), ['break:line', 'tab']);
  // Both are undone as they were made.
  assert.ok(history.undo());
  net.drain();
  assert.ok(history.undo());
  net.drain();
  assert.equal(bodyText(b.doc), 'abc\ndef');
  assertSameDoc([a, b], 'after undo');
  // A table of more cells than a command may make is no command; a document whose tables are full takes no more.
  assert.throws(() => decodeStep([{ t: 'cmd', name: 'insertTable', at: ['body', [0], 0], a: [500, 63] }]), { name: 'WireError' });
  assert.throws(() => decodeStep([{ t: 'text', at: ['body', [0], 0], x: '\n'.repeat(5001) }]), { name: 'WireError' });
  let made = 0;
  for (let i = 0; i < 40; i++) if (actDoc(a, [{ t: 'cmd', name: 'insertTable', at: at(a.doc.body.blocks.length - 1, 0), a: [100, 63] }])) made++;
  net.drain();
  assert.equal(made, 15, 'fifteen tables of 6,300 cells fit under the limit of a story; the sixteenth is refused');
  assertSameDoc([a, b], 'after the tables');
});
test('an undoing that was asked for while changes are on their way waits for them, and is then done', async () => {
  const { net, a, b } = await pair(['one', 'two']);
  assert.ok(actDoc(a, [{ t: 'text', at: at(0, 3), x: 'X' }], { before: 'x' }));
  assert.equal(a.session.undo(), undefined, 'not at once');
  net.drain();
  assert.equal(bodyText(a.doc), 'one\ntwo');
  assert.equal(a.changes.filter((change) => change.local).length, 1, 'the host was told when it was done');
  assert.equal(a.changes.filter((change) => change.local)[0].meta, 'x');
  assertSameDoc([a, b], 'after the undo');
  assert.deepEqual(codes(a), []);
});

/** A connection that says whatever the test tells it to. */
function intruder(net: DocNet, bytes: Uint8Array): { send(message: unknown): void; got: { t: string; rev?: number; id?: string }[] } {
  const rawTransport = net.connect();
  const got: { t: string; rev?: number; id?: string }[] = [];
  rawTransport.onMessage((text) => got.push(JSON.parse(text)));
  rawTransport.onState!((state) => {
    if (state === 'open') rawTransport.send(JSON.stringify({ t: 'join', v: [contentVersion(bytes)], key: 'evil', pv: 2 }));
  });
  return { send: (message) => rawTransport.send(JSON.stringify(message)), got };
}

test('what is no step of a document is refused when it is read', () => {
  let deep: unknown = 1;
  for (let i = 0; i < 20; i++) deep = { a: deep };
  const place = ['body', [0], 0];
  const bad: [string, unknown][] = [
    ['no list', { t: 'text' }],
    ['an empty step', []],
    ['an unknown operation', [{ t: 'cells', s: 0, c: [] }]],
    ['a place without a path', [{ t: 'text', at: ['body', [], 0], x: 'a' }]],
    ['a path that is no list of whole numbers', [{ t: 'text', at: ['body', [0.5], 0], x: 'a' }]],
    ['a path below zero', [{ t: 'text', at: ['body', [-1], 0], x: 'a' }]],
    ['a path that is too deep', [{ t: 'text', at: ['body', Array.from({ length: 200 }, () => 0), 0], x: 'a' }]],
    ['a story that is none', [{ t: 'text', at: ['__proto__', [0], 0], x: 'a' }]],
    ['an offset that is no number', [{ t: 'text', at: ['body', [0], '0'], x: 'a' }]],
    ['a text that is no text', [{ t: 'text', at: place, x: 5 }]],
    ['a text with a character no file holds', [{ t: 'text', at: place, x: 'a\u0000b' }]],
    ['a text with half a pair', [{ t: 'text', at: place, x: 'a\ud83d' }]],
    ['a format that is nested too deep', [{ t: 'run', at: place, to: place, f: deep }]],
    ['a format that is a list', [{ t: 'run', at: place, to: place, f: [] }]],
    ['a format with a name that reaches into the prototype', [{ t: 'run', at: place, to: place, f: JSON.parse('{"__proto__":{"x":1}}') }]],
    ['a deletion without an end', [{ t: 'del', at: place }]],
    ['a property that cannot be toggled', [{ t: 'unrun', at: place, to: place, k: 'color' }]],
    ['a command nobody knows', [{ t: 'cmd', name: 'insertImage', at: place, a: [] }]],
    ['a command of the prototype', [{ t: 'cmd', name: 'constructor', at: place, a: [] }]],
    ['a command with the wrong arguments', [{ t: 'cmd', name: 'insertTable', at: place, a: [100000, 2] }]],
    ['a link with a control character', [{ t: 'cmd', name: 'insertHyperlink', at: place, to: place, a: ['https://x/\u0001'] }]],
    ['a step of the kind that is no more', [{ t: 'rv', rev: 1 }]],
    ['paragraphs put at an offset', [{ t: 'put', at: ['body', [0], 2], n: 1, x: [{ r: [] }], h: '00000000' }]],
    ['no paragraph put', [{ t: 'put', at: place, n: 1, x: [], h: '00000000' }]],
    ['paragraphs put in the place of none', [{ t: 'put', at: place, n: 0, x: [{ r: [] }], h: '00000000' }]],
    ['a paragraph with a member nobody knows', [{ t: 'put', at: place, n: 1, x: [{ r: [], xml: '<w:p/>' }], h: '00000000' }]],
    ['a paragraph with a property the model keeps as XML', [{ t: 'put', at: place, n: 1, x: [{ r: [], p: { change: { type: 'element', name: 'w:pPrChange' } } }], h: '00000000' }]],
    ['a paragraph with a section', [{ t: 'put', at: place, n: 1, x: [{ r: [], p: { section: {} } }], h: '00000000' }]],
    ['a run that is tracked', [{ t: 'put', at: place, n: 1, x: [{ r: [{ c: ['a'], p: { inserted: { id: 1 } } }] }], h: '00000000' }]],
    ['a run with a property of the wrong kind', [{ t: 'put', at: place, n: 1, x: [{ r: [{ c: ['a'], p: { f: { bold: 'yes' } } }] }], h: '00000000' }]],
    ['a run with a colour that is none', [{ t: 'put', at: place, n: 1, x: [{ r: [{ c: ['a'], p: { f: { color: { rgb: '"><x' } } } }] }], h: '00000000' }]],
    ['a run with something that is no text', [{ t: 'put', at: place, n: 1, x: [{ r: [{ c: [{ type: 'drawing' }] }] }], h: '00000000' }]],
    ['a text with a line end in a run', [{ t: 'put', at: place, n: 1, x: [{ r: [{ c: ['a\nb'] }] }], h: '00000000' }]],
    ['an attribute that is none of those that are kept', [{ t: 'put', at: place, n: 1, x: [{ r: [], a: [['onclick', 'x']] }], h: '00000000' }]],
    ['an attribute that is kept, with a value that is none', [{ t: 'put', at: place, n: 1, x: [{ r: [], a: [['w:rsidR', '"/><w:x']] }], h: '00000000' }]],
    ['a paragraph from the prototype', [{ t: 'put', at: place, n: 1, x: [JSON.parse('{"r":[],"p":{"__proto__":{"bold":true}}}')], h: '00000000' }]],
    ['too many commands', Array.from({ length: 65 }, () => ({ t: 'split', at: place }))],
  ];
  for (const [what, step] of bad) assert.throws(() => decodeStep(step), { name: 'WireError' }, `refused: ${what}`);
  assert.equal(decodeStep([{ t: 'text', at: place, x: 'fine', f: { bold: true } }]).length, 1);
});

test('paragraphs that are put are read with the very rules of a change of formatting: what no command could write is written by none', async () => {
  const place = ['body', [1], 0];
  const put = (x: unknown, h = '00000000'): unknown[] => [{ t: 'put', at: place, n: 1, x: [x], h }];
  // Values no change of formatting takes.
  for (const [what, x] of [
    ['an alignment that is none', { p: { align: '_' }, r: [] }],
    ['a size below zero', { r: [{ c: ['x'], p: { f: { size: -5 } } }] }],
    ['a size in the mark of the paragraph below zero', { m: { f: { size: -5 } }, r: [] }],
    ['a size of complex scripts below zero', { r: [{ c: ['x'], p: { cs: { size: -5 } } }] }],
    ['a font slot that is empty', { r: [{ c: ['x'], p: { fonts: { eastAsia: '' } } }] }],
    ['a language that is none', { r: [{ c: ['x'], p: { lang: { val: '<x>' } } }] }],
    ['a property of a run no change of formatting sets', { r: [{ c: ['x'], p: { f: { border: {} } } }] }],
    ['a level of a list of a hundred million', { p: { numbering: { numId: 1, level: 100000000 } }, r: [] }],
    ['a highlight that is none', { r: [{ c: ['x'], p: { f: { highlight: 'zz' } } }] }],
    ['an underline that is none', { r: [{ c: ['x'], p: { f: { underline: 'zz' } } }] }],
    ['a colour with a tint and nothing to tint', { r: [{ c: ['x'], p: { f: { color: { tint: 30 } } } }] }],
    ['a mark of proofing that is none', { r: [{ e: 'foo' }] }],
    ['a property no change of formatting sets', { p: { cnfStyle: 'x' }, r: [] }],
    ['tab stops', { p: { tabs: [{ pos: 1, align: 'zz' }] }, r: [] }],
    ['an indent beyond the widest page', { p: { indent: { left: 100000000 } }, r: [] }],
    ['a line spacing below zero', { p: { spacing: { line: -100000000, lineRule: 'exact' } }, r: [] }],
    ['a language that is one letter', { r: [{ c: ['x'], p: { lang: { val: 'x' } } }] }],
    ['an rsid of nine digits', { a: [['w:rsidR', '123456789']], r: [] }],
  ] as [string, unknown][]) {
    assert.throws(() => decodeStep(put(x)), { name: 'WireError' }, what);
  }
  const two = (x: unknown, y: unknown, h = '00000000'): unknown[] => [{ t: 'put', at: place, n: 1, x: [x, y], h }];
  assert.doesNotThrow(() => decodeStep(two({ a: [['w14:paraId', '11111111'], ['w14:textId', '7FFFFFFF'], ['w:rsidR', '00AB12CD']], r: [{ b: [0, '5', 'm1'] }] }, { a: [['w14:paraId', '22222222']], r: [{ b: [1, '5'] }] })), 'a bookmark over two paragraphs');
  // What only the document can say: a style or a list it does not have, a bookmark whose number it has already.
  const bytes = await buildDocx({
    body: '<w:p><w:bookmarkStart w:id="0" w:name="mark"/><w:r><w:t>two</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p><w:p><w:r><w:t>a</w:t></w:r></w:p><w:p w14:paraId="0A0B0C0D"><w:r><w:t>b</w:t></w:r></w:p>',
    styles: '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/></w:style>',
  });
  const net = new DocNet();
  const a = await joinDoc(net, 'a', bytes);
  const b = await joinDoc(net, 'b', bytes);
  const evil = intruder(net, bytes);
  net.drain();
  let rev = a.session.rev;
  const say = (step: unknown[]): void => {
    evil.send({ t: 'op', base: rev, id: `evil:${rev}`, ops: step });
    net.drain();
    rev = a.session.rev;
  };
  const digest = (): string => putDigest([a.doc.body.blocks[1]]);
  say(put({ m: { f: { style: 'Nope' } }, r: [{ c: ['x'] }] }, digest()));
  say(put({ m: { f: { style: 'Normal' } }, r: [{ c: ['x'] }] }, digest()));
  say(put({ r: [{ c: ['x'], p: { f: { style: 'Nope' } } }] }, digest()));
  say(put({ p: { numbering: { numId: 4 } }, r: [{ c: ['x'] }] }, digest()));
  say(put({ r: [{ b: [0, '0', 'again'] }, { c: ['x'] }, { b: [1, '0'] }] }, digest()));
  // A bookmark of a name the document has; one that starts and never ends, or ends and never started.
  say(put({ r: [{ b: [0, '8', 'mark'] }, { c: ['x'] }, { b: [1, '8'] }] }, digest()));
  say(put({ r: [{ b: [0, '9', 'lonely'] }, { c: ['x'] }] }, digest()));
  say(put({ r: [{ c: ['x'] }, { b: [1, '9'] }] }, digest()));
  // A paragraph id the document has already.
  say(put({ a: [['w14:paraId', '0A0B0C0D']], r: [{ c: ['x'] }] }, digest()));
  // Values that are brought anew and are none of a file; names that the step itself has twice, over its paragraphs.
  say(put({ a: [['w:rsidR', '0']], r: [{ c: ['x'] }] }, digest()));
  say(put({ r: [{ c: ['x'], a: [['w:rsidRPr', 'ABC']] }] }, digest()));
  say(put({ a: [['w14:paraId', 'FFFFFFFF']], r: [{ c: ['x'] }] }, digest()));
  say(put({ a: [['w14:textId', '80000000']], r: [{ c: ['x'] }] }, digest()));
  say(put({ r: [{ b: [0, '5', 'x'] }, { b: [0, '5', 'y'] }, { b: [1, '5'] }] }, digest()));
  say(two({ r: [{ b: [0, '5', 'm1'] }, { c: ['x'] }, { b: [1, '5'] }] }, { r: [{ b: [0, '5', 'm2'] }, { c: ['y'] }, { b: [1, '5'] }] }, digest()));
  say(two({ r: [{ b: [0, '5', 'same'] }, { c: ['x'] }, { b: [1, '5'] }] }, { r: [{ b: [0, '6', 'SAME'] }, { c: ['y'] }, { b: [1, '6'] }] }, digest()));
  say(two({ a: [['w14:paraId', '11111111']], r: [{ c: ['x'] }] }, { a: [['w14:paraId', '11111111']], r: [{ c: ['y'] }] }, digest()));
  // Paragraphs that are fine, put over other paragraphs than the ones they were made for (the digest of what they replace is another one).
  say(put({ r: [{ c: ['x'] }] }, putDigest([a.doc.body.blocks[2]])));
  assert.equal(bodyText(a.doc), 'two\na\nb', 'nothing of it was applied');
  // And one that is fine: a bookmark of a number the document does not have, a character style it has.
  say(put({ r: [{ b: [0, '7', 'fresh'] }, { c: ['x'], p: { f: { style: 'Strong', bold: true } } }, { b: [1, '7'] }] }, digest()));
  assert.equal(bodyText(a.doc), 'two\nx\nb');
  assertSameDoc([a, b], 'after all');
  assert.deepEqual(independent({ a: await saveDoc(a) }), [], 'a reader that knows nothing of the engine finds the file whole');
});

test('a paragraph takes at most a million characters and a story twenty million: a step beyond that is skipped by everybody, and said', async () => {
  const { net, a, b, bytes } = await pair(['a', 'b']);
  const evil = intruder(net, bytes);
  net.drain();
  const say = (step: unknown[]): void => {
    evil.send({ t: 'op', base: a.session.rev, id: `evil:${a.session.rev}`, ops: step });
    net.drain();
  };
  const chunk = 'x'.repeat(400000);
  for (let i = 0; i < 4; i++) say([{ t: 'text', at: ['body', [0], 0], x: chunk }]);
  const first = a.doc.body.blocks[0] as Paragraph;
  assert.equal(paragraphText(first).length, 800001, 'two of the four fit');
  assert.ok(codes(a).indexOf('bad-step') >= 0);
  // Who writes there itself is refused as well.
  assert.throws(() => a.session.submit([{ t: 'text', at: at(0, 0), x: chunk }]), RangeError);
  // Many paragraphs of nearly a million each: the story stops at twenty million.
  const near = 'y'.repeat(399999);
  for (let i = 0; i < 60; i++) say([{ t: 'text', at: ['body', [1], 0], x: near + '\n' }]);
  const total = (a.doc.body.blocks as Paragraph[]).reduce((sum, block) => sum + paragraphText(block).length, 0);
  assert.ok(total <= 20000000 && total > 19000000, String(total));
  assertSameDoc([a, b], 'after all');
});

test('what a file holds that a put may not bring anew - a paragraph id twice, one of the upper half, a short rsid, bookmarks of one name or of one id, one that never ends - is put back where it stood: an undoing is not refused for it', async () => {
  const P = (attrs: string, text: string, extra = ''): string => `<w:p ${attrs}>${extra}<w:r><w:t>${text}</w:t></w:r></w:p>`;
  const cases: [string, string][] = [
    ['one paragraph id twice', P('w14:paraId="11111111"', 'one') + P('w14:paraId="11111111"', 'two')],
    ['a paragraph id of the upper half', P('w14:paraId="9ABCDEF0"', 'one') + P('', 'two')],
    ['a text id of the upper half', P('w14:paraId="11111111" w14:textId="FFFFFFFF"', 'one') + P('', 'two')],
    ['a short rsid', P('w:rsidR="ABC"', 'one') + P('', 'two')],
    ['a short rsid of a run', '<w:p><w:r w:rsidRPr="1"><w:t>one</w:t></w:r></w:p>' + P('', 'two')],
    ['bookmarks whose names differ in case', P('', 'one', '<w:bookmarkStart w:id="1" w:name="Mark"/><w:bookmarkEnd w:id="1"/>') + P('', 'two', '<w:bookmarkStart w:id="2" w:name="mark"/><w:bookmarkEnd w:id="2"/>')],
    ['two bookmarks of one id', P('', 'one', '<w:bookmarkStart w:id="1" w:name="A"/><w:bookmarkEnd w:id="1"/>') + P('', 'two', '<w:bookmarkStart w:id="1" w:name="B"/><w:bookmarkEnd w:id="1"/>')],
    ['a bookmark that never ends', P('', 'one', '<w:bookmarkStart w:id="1" w:name="A"/>') + P('', 'two')],
  ];
  for (const [what, body] of cases) {
    const bytes = await buildDocx({ body });
    const net = new DocNet();
    const a = await joinDoc(net, 'a', bytes);
    const b = await joinDoc(net, 'b', bytes);
    net.drain();
    const before = canonDoc(a.doc);
    assert.ok(actDoc(a, [{ t: 'del', at: at(0, 1), to: at(0, 2) }]), what);
    net.drain();
    assert.ok(a.session.undo(), what);
    net.drain();
    assert.equal(canonDoc(a.doc), before, `${what}: put back as it was`);
    assertSameDoc([a, b], what);
    assert.deepEqual(codes(a).concat(codes(b)), [], what);
  }
});

test('a paragraph id or a bookmark a put brings anew is one of its kind in every story of the document: headers, notes, comments', async () => {
  const body = '<w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p><w:p><w:r><w:t>three</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/></w:sectPr>';
  const header = '<w:p w14:paraId="11111111"><w:bookmarkStart w:id="7" w:name="hb"/><w:r><w:t>head</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>';
  const comments = '<w:comment w:id="0" w:author="x"><w:p w14:paraId="22222222"><w:r><w:t>c</w:t></w:r></w:p></w:comment>';
  for (const [what, x] of [
    ['the paragraph id of a header', { a: [['w14:paraId', '11111111']], r: [{ c: ['x'] }] }],
    ['the paragraph id of a comment', { a: [['w14:paraId', '22222222']], r: [{ c: ['x'] }] }],
    ['the bookmark id of a header', { r: [{ b: [0, '7', 'other'] }, { c: ['x'] }, { b: [1, '7'] }] }],
    ['the bookmark name of a header', { r: [{ b: [0, '8', 'HB'] }, { c: ['x'] }, { b: [1, '8'] }] }],
  ] as [string, unknown][]) {
    const bytes = await buildDocx({ body, headers: { rIdH1: header }, comments });
    const net = new DocNet();
    const a = await joinDoc(net, 'a', bytes);
    const evil = intruder(net, bytes);
    net.drain();
    evil.send({ t: 'op', base: a.session.rev, id: 'evil:1', ops: [{ t: 'put', at: ['body', [1], 0], n: 1, x: [x], h: putDigest(a.doc.body.blocks.slice(1, 2)) }] });
    net.drain();
    assert.equal(bodyText(a.doc), 'one\ntwo\nthree', what);
  }
});

test('a put brings what the blocks it replaces carried at most as often as they did: no paragraph id, bookmark or end of a bookmark twice, none left without its other end', async () => {
  const body =
    '<w:p w14:paraId="1A2B3C4D" w14:textId="77777777" w:rsidR="00AB12CD"><w:r><w:t>zero</w:t></w:r></w:p>' +
    '<w:p w14:paraId="11111111" w14:textId="22222222" w:rsidR="00AB12CD"><w:bookmarkStart w:id="1" w:name="one"/><w:r w:rsidRPr="00CD34EF"><w:t>one</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>' +
    '<w:p w14:paraId="33333333" w14:textId="44444444"><w:bookmarkStart w:id="2" w:name="over"/><w:r><w:t>two</w:t></w:r></w:p>' +
    '<w:p w14:paraId="55555555"><w:r><w:t>three</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p>';
  const P1 = { a: [['w14:paraId', '11111111'], ['w14:textId', '22222222']], r: [{ b: [0, '1', 'one'] }, { c: ['x'] }, { b: [1, '1'] }] };
  const cases: [string, number, unknown[], boolean][] = [
    ['what was there, put back once', 1, [P1], true],
    ['a paragraph id and a bookmark twice', 1, [P1, { ...P1, r: [{ b: [0, '1', 'one'] }, { c: ['y'] }, { b: [1, '1'] }] }], false],
    ['a paragraph id twice', 1, [{ a: P1.a, r: [{ c: ['x'] }] }, { a: P1.a, r: [{ c: ['y'] }] }], false],
    ['the end of a bookmark twice', 1, [{ r: [{ b: [0, '1', 'one'] }, { c: ['x'] }, { b: [1, '1'] }] }, { r: [{ c: ['y'] }, { b: [1, '1'] }] }], false],
    ['a second end of a bookmark that ends outside', 2, [{ r: [{ b: [0, '2', 'over'] }, { c: ['x'] }, { b: [1, '2'] }] }], false],
    ['the start of a bookmark twice', 2, [{ r: [{ b: [0, '2', 'over'] }, { c: ['x'] }] }, { r: [{ b: [0, '2', 'over'] }, { c: ['y'] }] }], false],
    ['the start dropped, the end left alone', 2, [{ r: [{ c: ['x'] }] }], false],
    ['the id of a paragraph that stays', 1, [{ a: [['w14:paraId', '33333333']], r: [{ c: ['x'] }] }], false],
  ];
  for (const [what, row, x, takes] of cases) {
    const bytes = await buildDocx({ body });
    const net = new DocNet();
    const a = await joinDoc(net, 'a', bytes);
    const b = await joinDoc(net, 'b', bytes);
    const evil = intruder(net, bytes);
    net.drain();
    evil.send({ t: 'op', base: a.session.rev, id: 'evil:1', ops: [{ t: 'put', at: ['body', [row], 0], n: 1, x, h: putDigest(a.doc.body.blocks.slice(row, row + 1)) }] });
    net.drain();
    assert.equal(JSON.stringify(a.doc.body.blocks).indexOf('"x"') >= 0, takes, what);
    assertSameDoc([a, b], what);
    assert.deepEqual(independent({ [what.replace(/[^a-z]+/g, '-')]: await saveDoc(a) }), [], what);
  }
});

test('the ids and bookmarks of every story are looked at: text boxes (also in mc:AlternateContent), the separators of the notes, bookmarks between blocks', async () => {
  const tbx = TBX;
  const body = '<w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p>' + tbx +
    '<w:bookmarkStart w:id="40" w:name="tablemark"/><w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:bookmarkEnd w:id="40"/><w:p/>';
  const footnotes = '<w:footnote w:type="separator" w:id="-1"><w:p w14:paraId="77777770"><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>';
  for (const [what, x] of [
    ['the paragraph id of a text box', { a: [['w14:paraId', '66666666']], r: [{ c: ['x'] }] }],
    ['the bookmark name of a text box', { r: [{ b: [0, '90', 'INBOX'] }, { c: ['x'] }, { b: [1, '90'] }] }],
    ['the bookmark id of a text box', { r: [{ b: [0, '31', 'n31'] }, { c: ['x'] }, { b: [1, '31'] }] }],
    ['the paragraph id of a separator', { a: [['w14:paraId', '77777770']], r: [{ c: ['x'] }] }],
    ['the name of a bookmark between blocks', { r: [{ b: [0, '91', 'tablemark'] }, { c: ['x'] }, { b: [1, '91'] }] }],
    ['the id of a bookmark between blocks', { r: [{ b: [0, '40', 'n40'] }, { c: ['x'] }, { b: [1, '40'] }] }],
  ] as [string, unknown][]) {
    const bytes = await buildDocx({ body, footnotes });
    const net = new DocNet();
    const a = await joinDoc(net, 'a', bytes);
    const evil = intruder(net, bytes);
    net.drain();
    evil.send({ t: 'op', base: a.session.rev, id: 'evil:1', ops: [{ t: 'put', at: ['body', [1], 0], n: 1, x: [x], h: putDigest(a.doc.body.blocks.slice(1, 2)) }] });
    net.drain();
    assert.equal(JSON.stringify(a.doc.body.blocks[1]).indexOf('"x"') >= 0, false, what);
  }
});

/** Sends a put of `x` over the blocks from `row` from a peer that does not care, and tells whether it was applied (by the marker "XX"). */
async function putBy(body: string, row: number, n: number, x: unknown[], spec: Record<string, unknown> = {}): Promise<{ applied: boolean; file: Uint8Array }> {
  const bytes = await buildDocx({ body, ...spec });
  const net = new DocNet();
  const a = await joinDoc(net, 'a', bytes);
  const b = await joinDoc(net, 'b', bytes);
  const evil = intruder(net, bytes);
  net.drain();
  evil.send({ t: 'op', base: a.session.rev, id: 'evil:1', ops: [{ t: 'put', at: ['body', [row], 0], n, x, h: putDigest(a.doc.body.blocks.slice(row, row + n)) }] });
  net.drain();
  assertSameDoc([a, b], 'the two who are honest');
  return { applied: JSON.stringify(a.doc.body.blocks).indexOf('"XX"') >= 0, file: await saveDoc(a) };
}

test('what names itself is found where the model keeps XML: bookmarks between rows and around them, in elements it does not know, in math, in both branches of mc:AlternateContent', async () => {
  const P0 = '<w:p w14:paraId="1A2B3C4D"><w:r><w:t>zero</w:t></w:r></w:p>';
  const mark = [{ r: [{ b: [0, '9', 'rowmark'] }, { c: ['XX'] }, { b: [1, '9'] }] }];
  const box = (id: string): string => TBX.replace('66666666', id).replace(/^<w:p><w:r><w:t>pre<\/w:t><\/w:r><w:r>/, '').replace(/<\/w:r><\/w:p>$/, '');
  const alt = (choice: string, fallback: string): string => `<w:p><w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="zzz">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent></w:r></w:p>`;
  for (const [what, tail, x] of [
    ['a bookmark between the cells of a row', '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:bookmarkStart w:id="9" w:name="rowmark"/><w:tc><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc><w:bookmarkEnd w:id="9"/></w:tr></w:tbl><w:p/>', mark],
    ['a bookmark between the rows of a table', '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:bookmarkStart w:id="9" w:name="rowmark"/><w:tr><w:tc><w:p><w:r><w:t>c</w:t></w:r></w:p></w:tc></w:tr><w:bookmarkEnd w:id="9"/></w:tbl><w:p/>', mark],
    ['a paragraph id and a bookmark in a block the model does not know', '<w:customXmlFoo><w:p w14:paraId="0ABCDEF0"><w:bookmarkStart w:id="9" w:name="rowmark"/><w:r><w:t>u</w:t></w:r><w:bookmarkEnd w:id="9"/></w:p></w:customXmlFoo><w:p/>', [{ a: [['w14:paraId', '0ABCDEF0']], r: [{ c: ['XX'] }] }]],
    ['a bookmark in math', '<w:p><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:bookmarkStart w:id="9" w:name="rowmark"/><m:r><m:t>x</m:t></m:r><w:bookmarkEnd w:id="9"/></m:oMath></w:p>', mark],
    ['a bookmark in an inline element the model does not know', '<w:p><w:foo><w:bookmarkStart w:id="9" w:name="rowmark"/></w:foo><w:r><w:t>x</w:t></w:r><w:bookmarkEnd w:id="9"/></w:p>', mark],
    ['the paragraph id of a text box in the branch that is not shown', alt(box('66666666'), box('77777777')), [{ a: [['w14:paraId', '77777777']], r: [{ c: ['XX'] }] }]],
  ] as [string, string, unknown[]][]) {
    const { applied, file } = await putBy(P0 + tail, 0, 1, x);
    assert.equal(applied, false, what);
    assert.deepEqual(independent({ [what.replace(/[^a-z]+/g, '-')]: file }), [], what);
  }
  // The same text box in both branches is one: its paragraph id is there once, and what names it once can be put back.
  const same = alt(box('66666666'), box('66666666'));
  assert.equal((await putBy(P0 + same, 0, 1, [{ a: [['w14:paraId', '1A2B3C4D']], r: [{ c: ['XX'] }] }])).applied, true, 'what the replaced paragraph had');
});

test('bookmarks a put brings: a start has a name, an id is a number written as numbers are and counted as one, the end comes after the start', async () => {
  const P = (id: string, inner = '<w:r><w:t>p</w:t></w:r>'): string => `<w:p w14:paraId="${id}">${inner}</w:p>`;
  const place = ['body', [0], 0];
  for (const [what, x] of [
    ['a start without a name', { r: [{ b: [0, '77'] }, { c: ['XX'] }, { b: [1, '77'] }] }],
    ['an id with a leading zero', { r: [{ b: [0, '09', 'other'] }, { c: ['XX'] }, { b: [1, '09'] }] }],
    ['an id beyond 2^31-1', { r: [{ b: [0, '2147483648', 'big'] }, { c: ['XX'] }, { b: [1, '2147483648'] }] }],
  ] as [string, unknown][]) {
    assert.throws(() => decodeStep([{ t: 'put', at: place, n: 1, x: [x], h: '00000000' }]), { name: 'WireError' }, what);
  }
  const orph = P('11111111', '<w:bookmarkStart w:id="7" w:name="o"/><w:r><w:t>p</w:t></w:r>') + P('22222222', '<w:r><w:t>q</w:t></w:r><w:bookmarkEnd w:id="7"/>');
  for (const [what, body, row, x, takes] of [
    ['the start dropped, its end left alone', orph, 0, [{ r: [{ c: ['XX'] }] }], false],
    ['the start put back', orph, 0, [{ r: [{ b: [0, '7', 'o'] }, { c: ['XX'] }] }], true],
    ['the id of a bookmark the file writes with a leading zero', P('11111111') + P('22222222', '<w:bookmarkStart w:id="09" w:name="nine"/><w:r><w:t>q</w:t></w:r><w:bookmarkEnd w:id="09"/>'), 0, [{ r: [{ b: [0, '9', 'other'] }, { c: ['XX'] }, { b: [1, '9'] }] }], false],
    ['an end before its start in a paragraph', P('11111111'), 0, [{ r: [{ b: [1, '50'] }, { c: ['XX'] }, { b: [0, '50', 'rev'] }] }], false],
    ['a start moved behind its end', P('11111111', '<w:r><w:t>p</w:t></w:r><w:bookmarkEnd w:id="7"/>') + P('22222222', '<w:bookmarkStart w:id="7" w:name="o"/><w:r><w:t>q</w:t></w:r>'), 1, [{ r: [{ c: ['XX'] }, { b: [0, '7', 'o'] }] }], true],
    ['a paragraph id the file has twice, a third time', P('11111111') + P('22222222') + P('22222222'), 0, [{ a: [['w14:paraId', '22222222']], r: [{ c: ['XX'] }] }], false],
  ] as [string, string, number, unknown[], boolean][]) {
    assert.equal((await putBy(body, row, 1, x)).applied, takes, what);
  }
});

test('what cannot be read whole for what it names is not taken to name nothing: a put there is refused', async () => {
  const names = (wireModule as unknown as { namesOfDocument?: (doc: DocModel, skip: Block[], visit: () => void) => void }).namesOfDocument;
  assert.ok(names, 'the reading of what a document names');
  const doc = await (await open(await docFile(['zero', 'one']))).document!();
  // A block whose bytes, as the writer would copy them, are no XML (made in the model itself: no file is read so).
  const at = doc.body.source.length;
  doc.body.source += '<w:p><w:r></w:p>';
  doc.body.blocks.push({ type: 'paragraph', props: {}, children: [], srcStart: at, srcEnd: doc.body.source.length } as unknown as Block);
  assert.throws(() => names(doc, [], () => undefined), RangeError);
  assert.throws(() => runWire(doc, { t: 'put', at: { s: 'body', p: [0], o: 0 }, n: 1, x: [{ r: [{ c: ['x'] }] }], h: putDigest(doc.body.blocks.slice(0, 1)) }), RangeError);
});

test('the largest document that is edited together is read whole for what it names', () => {
  const tokens = (wireModule as unknown as { NAME_TOKENS_PER_PART?: number }).NAME_TOKENS_PER_PART;
  // What a part may hold at most (characters, each an element; blocks; cells), four times over, stays under the tokens read in a part.
  assert.ok(tokens !== undefined && tokens >= 4 * (MAX_STORY_CHARS_LIMIT + MAX_LIST_BLOCKS + 4 * MAX_STORY_CELLS));
});

test('a composition while someone else changes the document: what nobody else touched is undone, what someone touched is not, and all agree', async () => {
  const run = async (mid: (a: DocMember, b: DocMember, i: number) => void, after?: (a: DocMember, b: DocMember) => void, enterAt = -1): Promise<{ a: DocMember; b: DocMember; undone: boolean; typed: string }> => {
    const { net, a, b } = await pair(['hello world', 'second', 'third']);
    let composed = '';
    let row = 0;
    let off = 5;
    for (let i = 0; i < 6; i++) {
      if (i === enterAt) {
        assert.ok(actDoc(a, [{ t: 'split', at: at(row, off + composed.length) }], { key: 'composition-1' }));
        net.drain();
        row++;
        off = 0;
        composed = '';
      }
      const next = composed + String.fromCharCode(0x3042 + i);
      assert.ok(actDoc(a, [composed.length ? { t: 'text', at: at(row, off), to: at(row, off + composed.length), x: next } : { t: 'text', at: at(row, off), x: next }], { key: 'composition-1' }));
      composed = next;
      net.drain();
      mid(a, b, i);
      net.drain();
    }
    if (after) {
      after(a, b);
      net.drain();
    }
    a.session.seal();
    const typed = bodyText(a.doc);
    const undone = a.session.undo() !== undefined;
    for (let i = 0; i < 5; i++) net.drain();
    assertSameDoc([a, b], 'after the undo');
    return { a, b, undone, typed };
  };
  // Someone types in another paragraph: the composition is undone, what they typed stays.
  let r = await run((_a, b, i) => void (i === 2 && actDoc(b, [{ t: 'text', at: at(1, 0), x: 'Z' }])));
  assert.equal(r.undone, true);
  assert.equal(bodyText(r.b.doc), 'hello world\nZsecond\nthird');
  assert.deepEqual(codes(r.a), []);
  // Enter in the composition: undone with it.
  r = await run(() => undefined, undefined, 3);
  assert.equal(r.undone, true);
  assert.equal(bodyText(r.b.doc), 'hello world\nsecond\nthird');
  // Someone types in the paragraph of the composition, at its end or at its start, or colours it, while it goes on
  // or after it: the paragraph is not put back over what they did; that is said, and nothing changes.
  for (const [what, mid, after] of [
    ['typed at the end of the paragraph', (_a: DocMember, b: DocMember, i: number) => void (i === 2 && actDoc(b, [{ t: 'text', at: at(0, bodyText(b.doc).split('\n')[0].length), x: 'Z' }])), undefined],
    ['typed at the start of the paragraph', (_a: DocMember, b: DocMember, i: number) => void (i === 2 && actDoc(b, [{ t: 'text', at: at(0, 0), x: 'Z' }])), undefined],
    ['coloured while it is composed', (_a: DocMember, b: DocMember, i: number) => void (i === 2 && actDoc(b, [{ t: 'run', at: at(0, 0), to: at(0, 3), f: { color: 'FF0000' } }])), undefined],
    ['coloured after it was composed', () => undefined, (_a: DocMember, b: DocMember) => void actDoc(b, [{ t: 'run', at: at(0, 0), to: at(0, 3), f: { color: 'FF0000' } }])],
  ] as [string, (a: DocMember, b: DocMember, i: number) => void, ((a: DocMember, b: DocMember) => void) | undefined][]) {
    r = await run(mid, after);
    assert.equal(r.undone, false, what);
    assert.equal(bodyText(r.a.doc), r.typed, `${what}: nothing changed`);
    assert.deepEqual(codes(r.a), ['undo-gone'], what);
  }
});

test('a composition of many updates, each one a step of its own, is undone and made again in one, for everybody', async () => {
  for (const updates of [9, 10, 40]) {
    const { net, a, b } = await pair(['hello world', 'second']);
    let composed = '';
    for (let i = 0; i < updates; i++) {
      const next = composed + String.fromCharCode(0x3042 + i);
      assert.ok(actDoc(a, [composed.length ? { t: 'text', at: at(0, 5), to: at(0, 5 + composed.length), x: next } : { t: 'text', at: at(0, 5), x: next }], { key: 'composition-1' }), `${updates}: update ${i}`);
      composed = next;
      net.drain();
    }
    a.session.seal();
    const typed = bodyText(a.doc);
    assert.ok(a.session.undo(), `${updates}: undone`);
    net.drain();
    assert.equal(bodyText(b.doc), 'hello world\nsecond', `${updates}: as it was`);
    assert.ok(a.session.redo(), `${updates}: made again`);
    net.drain();
    assert.equal(bodyText(b.doc), typed, `${updates}: as it was typed`);
    assert.deepEqual(codes(a).concat(codes(b)), [], `${updates}: nothing said`);
    assertSameDoc([a, b], `${updates}`);
  }
});

test('the puts of a step are sixteen at most and take out and put back four thousand paragraphs at most: what costs more is no step', () => {
  const put = (n: number, k: number): unknown => ({ t: 'put', at: ['body', [0], 0], n, x: Array.from({ length: k }, () => ({ r: [] })), h: '00000000' });
  assert.throws(() => decodeStep([put(2000, 1000), put(500, 501)]), { name: 'WireError' }, 'more than four thousand paragraphs');
  assert.doesNotThrow(() => decodeStep(Array.from({ length: 16 }, () => put(1, 1))), 'sixteen small ones');
  // Each put reads the whole document for what it names: a step holds sixteen at most.
  assert.throws(() => decodeStep(Array.from({ length: 17 }, () => put(1, 1))), { name: 'WireError' }, 'seventeen');
  assert.doesNotThrow(() => decodeStep(Array.from({ length: 8 }, () => put(250, 250))));
});

test('a put: what it may take out, the names it may carry, and the paragraphs that do not travel', async () => {
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  // A part that declares no w14: an attribute of that namespace would make it one no reader takes.
  const bare = await buildDocx({
    document: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${w}"><w:body><w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:hyperlink w:anchor="x"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p><w:p><w:r><w:t>three</w:t></w:r></w:p></w:body></w:document>`,
  });
  const net = new DocNet();
  const a = await joinDoc(net, 'a', bare);
  const b = await joinDoc(net, 'b', bare);
  const evil = intruder(net, bare);
  net.drain();
  const say = (row: number, x: unknown): void => {
    evil.send({ t: 'op', base: a.session.rev, id: `evil:${a.session.rev}`, ops: [{ t: 'put', at: ['body', [row], 0], n: 1, x: [x], h: putDigest([a.doc.body.blocks[row]]) }] });
    net.drain();
  };
  say(0, { r: [{ c: ['x'] }], a: [['w14:paraId', '1234ABCD']] });
  assert.equal(bodyText(a.doc), 'one\nlink\nthree', 'an attribute of a namespace the part does not declare');
  // A paragraph that holds a link does not travel, and is not taken out by a put either (even one that names it rightly).
  say(1, { r: [{ c: ['x'] }] });
  assert.equal(bodyText(a.doc), 'one\nlink\nthree', 'a paragraph that is not plain');
  say(0, { r: [{ c: ['x'] }], a: [['w:rsidR', '00AB12CD']] });
  assert.equal(bodyText(a.doc), 'x\nlink\nthree', 'an attribute of the namespace of the part');
  assertSameDoc([a, b], 'after all');
  // What travels is what is read: a paragraph whose formatting reads back as another one does not travel.
  const odd: Paragraph = { type: 'paragraph', props: {}, children: [{ type: 'run', props: { color: { rgb: 'ff00aa' } }, content: [{ type: 'text', text: 'x' }] }] };
  assert.equal(putOf(odd), undefined, 'a colour spelled in small letters reads back in capitals');
  const even: Paragraph = { type: 'paragraph', props: {}, children: [{ type: 'run', props: { color: { rgb: 'FF00AA' } }, content: [{ type: 'text', text: 'x' }] }] };
  assert.ok(putOf(even));
  assert.equal(putOf({ type: 'paragraph', props: {}, children: [{ type: 'run', props: { bold: true, underline: { style: 'single', color: { rgb: 'FF0000' } } }, content: [] }] }), undefined, 'an underline with a colour of its own');
});

test('a mark that says only that a story changed drops every waiting command of that story, of every story when it is the document', () => {
  const text: DocWire = { t: 'text', at: { s: 'body', p: [0], o: 1 }, x: 'a' };
  assert.equal(rebaseWire(text, [], [{ k: 'opq', s: 'body' }]).drop, true);
  assert.equal(rebaseWire(text, [], [{ k: 'opq', s: '*' }]).drop, true);
  assert.equal(rebaseWire(text, [], [{ k: 'opq', s: 'header:rId7' }]).drop, undefined, 'another story: it stays');
  // A command that itself changed its story in a way that is not known does not go with anything in it.
  const own: Mark[] = [{ k: 'opq', s: 'body' }];
  assert.equal(rebaseWire(text, own, [{ k: 'ins', s: 'body', c: [], row: 5, off: 0, rows: 0, last: 1 }]).drop, true);
  assert.equal(rebaseWire(text, own, [{ k: 'ins', s: 'header:rId7', c: [], row: 0, off: 0, rows: 0, last: 1 }]).drop, undefined);
});

test('a step that is no step of the document is skipped by everybody, also by who joins later, and the room goes on', async () => {
  const { net, a, b, bytes } = await pair(['one', 'two']);
  const evil = intruder(net, bytes);
  net.drain();
  assert.ok(actDoc(b, [{ t: 'text', at: at(1, 0), x: 'waits ' }]));
  net.drain();
  let rev = a.session.rev;
  const say = (step: unknown[]): void => {
    evil.send({ t: 'op', base: rev, id: `evil:${rev}`, ops: step });
    net.drain();
    rev = a.session.rev;
  };
  // One that cannot be read; ones that read and cannot be applied; one that takes back a step that is none, or not there yet.
  say([{ t: 'no such operation' }]);
  say([{ t: 'text', at: ['body', [7], 0], x: 'nowhere' }]);
  say([{ t: 'text', at: ['body', [0], 99], x: 'beyond the paragraph' }]);
  say([{ t: 'del', at: ['body', [0], 0], to: ['header:none', [0], 0] }]);
  say([{ t: 'run', at: ['body', [0], 0], to: ['body', [0], 2], f: { style: 'NoSuchStyle' } }]);
  say([{ t: 'run', at: ['body', [0], 0], to: ['body', [0], 2], f: { size: -4 } }]);
  say([{ t: 'cmd', name: 'insertRow', at: ['body', [0], 0], a: ['below'] }]);
  say([{ t: 'cmd', name: 'insertHyperlink', at: ['body', [0], 0], to: ['body', [0], 2], a: ['javascript:alert(1)'] }]);
  // Paragraphs that are put: where the document has no such paragraphs, with a style it does not have, with the mark of a namespace its part does not declare.
  say([{ t: 'put', at: ['body', [1], 0], n: 5, x: [{ r: [{ c: ['x'] }] }], h: putDigest([a.doc.body.blocks[0]]) }]);
  say([{ t: 'put', at: ['body', [0], 0], n: 1, x: [{ r: [{ c: ['x'] }], p: { style: 'NoSuchStyle' } }], h: putDigest([a.doc.body.blocks[0]]) }]);
  // A step whose first command is fine and whose second is not changes nothing.
  say([{ t: 'text', at: ['body', [0], 0], x: 'half ' }, { t: 'text', at: ['body', [9], 0], x: 'x' }]);
  assert.equal(rev, 12, 'the relay placed them all (and what waited)');
  for (const member of [a, b]) {
    assert.equal(member.session.state, 'online', `${member.name} stays in the room`);
    assert.deepEqual(codes(member), ['bad-step'], `${member.name} was told once`);
  }
  assert.equal(bodyText(a.doc), 'one\nwaits two');
  assert.ok(actDoc(a, [{ t: 'text', at: at(0, 3), x: ' goes on' }]));
  net.drain();
  const late = await joinDoc(net, 'late', bytes);
  net.drain();
  assert.equal(late.session.state, 'online', 'who replays the log stays as well');
  assertSameDoc([a, b, late], 'after the steps that were skipped');
  assert.equal(bodyText(late.doc), 'one goes on\nwaits two');
});

function python(): boolean {
  try {
    execFileSync('python3', ['-I', '-c', 'import xml.etree.ElementTree, zipfile'], { stdio: 'ignore', timeout: 60000 });
    return true;
  } catch {
    return false;
  }
}

const NASTY = ['"><w:x y="', ']]>', '&amp;<&>', "it's", ' lead and trail ', '<!--', '\u00e7\u011f\u0131\u00f6\u015f\u00fc\u0130', '\ud83d\ude00', 'a\tb', 'two\nlines', 'x'.repeat(300), '', '\u200f\u202e'];

test(`everything at once: all agree, what is dropped is said, and the files are whole (${SEEDS} runs)`, async (t) => {
  const bytes = await aloneFile();
  const withPython = python();
  if (!withPython) t.diagnostic('python3 is not here: the files are read by the engine only');
  const files: { [name: string]: string } = {};
  let made = 0;
  let dropped = 0;
  for (let seed = 1; seed <= SEEDS; seed++) {
    const random = seeded(seed * 613);
    const net = new DocNet();
    const members: DocMember[] = [];
    for (let i = 0; i < 2 + (seed % 3); i++) members.push(await joinDoc(net, `m${i}`, bytes));
    const histories = members.map((member) => new CollabDocHistory<string>(member.session));
    const evil = seed % 4 === 0 ? intruder(net, bytes) : undefined;
    let joined: DocMember | undefined;
    net.drain(random);
    for (let tick = 0; tick < 90; tick++) {
      if (tick === 45 && seed % 2 === 0) {
        // Half way: the file is saved, and someone opens it and joins. What it holds was read from a file, not edited into being.
        for (const member of members) member.transport.resume();
        net.drain(random);
        const stored = await saveDoc(members[0]);
        members[0].session.announceSaved([contentVersion(stored)], members[0].session.rev);
        net.drain(random);
        const late = await joinDoc(net, 'late', stored);
        net.drain(random);
        assert.equal(late.session.state, 'online', `seed ${seed}: who opens the saved file joins`);
        joined = late;
        members.push(late);
        histories.push(new CollabDocHistory<string>(late.session));
      }
      const roll = random();
      const who = Math.floor(random() * members.length);
      const member = members[who];
      if (roll < 0.4) net.step(random);
      else if (roll < 0.47) histories[who].undo();
      else if (roll < 0.51) histories[who].redo();
      else if (roll < 0.54) member.transport.suspend();
      else if (roll < 0.58) member.transport.resume();
      else if (evil && roll < 0.62) {
        const place = ['body', [Math.floor(random() * 4)], Math.floor(random() * 8)];
        const text = NASTY[Math.floor(random() * NASTY.length)];
        evil.send({ t: 'op', base: members[0].session.rev, id: `evil:${tick}`, ops: [random() < 0.5 ? { t: 'text', at: place, x: text, f: { font: text.slice(0, 20) || 'Arial' } } : { t: 'cmd', name: 'insertHyperlink', at: place, to: place, a: [`https://example.com/${encodeURIComponent(text.slice(0, 30))}`] }] });
      } else {
        let edit: DocEdit | undefined;
        try {
          edit = make(member.doc, seeded(seed * 1000 + tick))(member.doc);
          if (random() < 0.5) histories[who].seal();
          histories[who].apply(edit.ops, '', '', 'typing');
        } catch (error) {
          if (!(error instanceof RangeError)) throw error;
        }
      }
    }
    for (const member of members) member.transport.resume();
    for (let round = 0; round < 30; round++) {
      net.drain(random);
      await new Promise<void>((resolve) => setImmediate(resolve));
      if (net.waiting === 0) break;
    }
    // Those who read the same file hold the same document, block for block; the one who read the saved file holds what
    // they hold as text and as looks (a document read from a file is cut into other runs than the one that was edited).
    // An undoing of a step from before its file sends it back to the file, which is another test: here it stays or is kept apart.
    assertSameDoc(members.filter((member) => member !== joined), `seed ${seed}`);
    if (joined && joined.session.state === 'online') {
      const a = lookOfDoc(members[0].doc);
      const b = lookOfDoc(joined.doc);
      if (a !== b) assert.fail(`seed ${seed}: the one who joined with the saved file holds another document ${firstDifference(a, b)}`);
    } else if (joined) members.splice(members.indexOf(joined), 1);
    for (const member of members) {
      const stats = member.session.stats;
      assert.equal(member.session.state, 'online', `seed ${seed}: ${member.name} is in the room`);
      assert.equal(member.session.pending, 0, `seed ${seed}: nothing of ${member.name} waits`);
      assert.equal(stats.made, stats.placed + stats.dropped + stats.quiet, `seed ${seed}: every change of ${member.name} was placed, dropped or let go`);
      const said = member.notices.reduce((sum, notice) => sum + (notice.code === 'dropped' || notice.code === 'undo-gone' ? notice.count || 0 : 0), 0);
      assert.equal(said, stats.dropped, `seed ${seed}: ${member.name} was told of every change that was dropped`);
      made += stats.made;
      dropped += stats.dropped;
    }
    // The file of each: it opens, and holds the document.
    for (const member of members.slice(0, 2)) {
      const saved = await saveDoc(member);
      const again = await (await open(saved)).document!();
      assert.equal(bodyText(again), bodyText(member.doc), `seed ${seed}: the file of ${member.name} holds its text`);
      if (withPython && seed % 3 === 0) {
        const zip = ZipArchive.read(saved);
        for (const name of zip.names()) if (/\.(xml|rels)$/.test(name)) files[`seed ${seed} ${member.name}: ${name}`] = Buffer.from(await zip.read(name)).toString('base64');
      }
    }
    for (const member of members) member.session.leave();
  }
  assert.ok(made > SEEDS * 10, 'changes were made');
  t.diagnostic(`${made} changes, ${dropped} dropped and said`);
  if (withPython) {
    const dir = mkdtempSync(joinPath(process.cwd(), 'test', '.collab-doc-'));
    try {
      const file = joinPath(dir, 'parts.json');
      writeFileSync(file, JSON.stringify(files));
      const out = execFileSync('python3', ['-I', '-c', 'import sys,json,base64\nimport xml.etree.ElementTree as E\nbad=[]\nfor k,v in json.load(open(sys.argv[1])).items():\n  try:\n    E.fromstring(base64.b64decode(v))\n  except Exception as e:\n    bad.append(k+": "+str(e))\nprint(json.dumps(bad))', file], { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024, timeout: 120000, stdio: ['ignore', 'pipe', 'inherit'] });
      assert.deepEqual(JSON.parse(out), [], 'every part of every file is well formed, namespaces and all');
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  }
});

test('two people who write, as people do: what is dropped, and how often', async (t) => {
  const bytes = await docFile(['The first paragraph, which one of the two works on.', 'The second paragraph, for the other one.', 'A third one that both come to now and then.']);
  let made = 0;
  let dropped = 0;
  let undone = 0;
  let typed = 0;
  let changesDropped = 0;
  for (let seed = 1; seed <= Math.max(6, Math.floor(SEEDS / 3)); seed++) {
    const random = seeded(seed * 31337);
    const net = new DocNet();
    const members: DocMember[] = [await joinDoc(net, 'a', bytes), await joinDoc(net, 'b', bytes)];
    const histories = members.map((member) => new CollabDocHistory<string>(member.session));
    net.drain(random);
    // Each has a caret: a paragraph (its own, mostly) and an offset that moves on as it types.
    const carets = [
      { row: 0, offset: 10 },
      { row: 1, offset: 10 },
    ];
    for (let tick = 0; tick < 400; tick++) {
      // The relay is quick: most of the time everything has arrived before the next key.
      if (random() < 0.7) net.drain(random);
      else if (random() < 0.5) net.step(random);
      const who = random() < 0.5 ? 0 : 1;
      const member = members[who];
      const caret = carets[who];
      const blocks = member.doc.body.blocks;
      if (random() < 0.04) caret.row = random() < 0.5 ? 2 : who;
      caret.row = Math.min(caret.row, blocks.length - 1);
      const block = blocks[caret.row];
      if (block.type !== 'paragraph') continue;
      caret.offset = Math.min(caret.offset, ops.paragraphLength(block));
      const place = { story: 'body', path: [caret.row], offset: caret.offset };
      const roll = random();
      try {
        if (roll < 0.72) {
          const text = random() < 0.15 ? ' ' : 'abcdefgh'[Math.floor(random() * 8)];
          histories[who].apply(ops.insertText(member.doc, place, text).ops, '', '', 'typing');
          caret.offset++;
          typed++;
        } else if (roll < 0.84) {
          histories[who].apply(ops.deleteBackward(member.doc, place).ops, '', '', 'deleting');
          caret.offset = Math.max(0, caret.offset - 1);
        } else if (roll < 0.88) {
          histories[who].seal();
          histories[who].apply(ops.splitParagraph(member.doc, place).ops, '', '');
          caret.row++;
          caret.offset = 0;
        } else if (roll < 0.93) {
          histories[who].seal();
          histories[who].apply(ops.toggleRunFormat(member.doc, { anchor: { ...place, offset: Math.max(0, caret.offset - 5) }, focus: place }, 'bold').ops, '', '');
        } else if (roll < 0.96) {
          histories[who].seal();
          histories[who].apply(ops.setParaFormat(member.doc, place, { align: 'center' }).ops, '', '');
        } else {
          histories[who].undo();
          undone++;
        }
      } catch (error) {
        if (!(error instanceof RangeError)) throw error;
      }
    }
    net.drain(random);
    assertSameDoc(members, `seed ${seed}`);
    for (const member of members) {
      made += member.session.stats.made;
      dropped += member.session.stats.dropped;
      changesDropped += member.notices.reduce((sum, notice) => sum + (notice.code === 'dropped' ? notice.count || 0 : 0), 0);
      member.session.leave();
    }
  }
  t.diagnostic(`${made} changes (${typed} keys typed, ${undone} times undo), ${dropped} dropped and said (${changesDropped} changes, the rest undoings): ${((100 * dropped) / made).toFixed(2)} %`);
  assert.ok(made > 500);
});

test('a link, a list and bold laid over text someone else types into: the text stays where it was typed', async () => {
  const { net, a, b } = await pair(['abcdefgh', 'second']);
  // The link is heard first; the one who types has not heard of it.
  assert.ok(actDoc(a, [{ t: 'cmd', name: 'insertHyperlink', at: at(0, 2), to: at(0, 6), a: ['https://example.com/'] }]));
  assert.ok(actDoc(a, [{ t: 'cmd', name: 'toggleList', at: at(0, 0), to: at(1, 0), a: ['bullet'] }]));
  assert.ok(actDoc(a, [{ t: 'run', at: at(0, 0), to: at(0, 8), f: { bold: true } }]));
  assert.ok(actDoc(b, [{ t: 'text', at: at(0, 4), x: 'XY' }]));
  assert.ok(actDoc(b, [{ t: 'text', at: at(1, 6), x: '!' }]));
  net.drain();
  assert.equal(bodyText(a.doc), 'abcdXYefgh\nsecond!');
  assertSameDoc([a, b], 'after all');
  assert.equal(a.session.stats.dropped + b.session.stats.dropped, 0, 'nothing was dropped');
});

test('a caret whose paragraph someone else put back stays in that paragraph, as far in as the paragraph goes; without it, in a neighbour', async () => {
  const { net, a, b } = await pair(['first', 'second paragraph', 'third']);
  // a deletes in the second paragraph; b has its caret in it; a takes the deletion back (the paragraph is put back as a whole).
  assert.ok(actDoc(a, [{ t: 'run', at: at(1, 0), to: at(1, 6), f: { bold: true } }]));
  net.drain();
  const caret = { anchor: { story: 'body', path: [1], offset: 9 }, focus: { story: 'body', path: [1], offset: 9 } };
  const from = b.changes.length;
  assert.ok(a.session.undo());
  net.drain();
  const seen: Mark[][] = b.changes.slice(from).map((change) => change.marks);
  assert.ok(seen.length === 1 && seen[0].some((mark) => mark.k === 'blk'), 'the paragraph was replaced as a whole');
  const moved = movedSelection(caret, seen[0])!;
  assert.deepEqual(placeNear(b.doc, moved.focus), { story: 'body', path: [1], offset: 9 }, 'not at the start of the document, and where it was: the text is the same');
  // Text that comes back before the caret: the caret stays with the text after it. Text that comes back around it: it stands after it.
  const net2 = new DocNet();
  const bytes2 = await docFile(['abc0123456789', 'next']);
  const c = await joinDoc(net2, 'c', bytes2);
  const d = await joinDoc(net2, 'd', bytes2);
  net2.drain();
  assert.ok(actDoc(c, [{ t: 'del', at: at(0, 0), to: at(0, 3) }]));
  net2.drain();
  const after8 = { anchor: { story: 'body', path: [0], offset: 9 }, focus: { story: 'body', path: [0], offset: 9 } };
  let since = d.changes.length;
  assert.ok(c.session.undo());
  net2.drain();
  let place = d.changes.slice(since).reduce((sel: typeof after8 | undefined, change) => (sel ? movedSelection(sel, change.marks) : sel), after8)!;
  assert.deepEqual(place.focus, { story: 'body', path: [0], offset: 12 }, 'after the 8 still');
  assert.ok(actDoc(d, [{ t: 'text', at: at(0, place.focus.offset), x: 'Z' }]));
  net2.drain();
  assert.equal(bodyText(c.doc), 'abc012345678Z9\nnext');
  // Put back over two paragraphs that were one: the caret in the second half is in the second paragraph again.
  assert.ok(actDoc(c, [{ t: 'del', at: at(0, 5), to: at(1, 2) }]));
  net2.drain();
  assert.equal(bodyText(d.doc), 'abc01xt');
  const inTail = { anchor: { story: 'body', path: [0], offset: 6 }, focus: { story: 'body', path: [0], offset: 6 } };
  since = d.changes.length;
  assert.ok(c.session.undo());
  net2.drain();
  place = d.changes.slice(since).reduce((sel: typeof after8 | undefined, change) => (sel ? movedSelection(sel, change.marks) : sel), inTail)!;
  assert.deepEqual(place.focus, { story: 'body', path: [1], offset: 3 }, 'between x and t');
  // A paragraph that got shorter: the end of it. One that is no paragraph any more: the end of the one before, else the start of the one after.
  assert.deepEqual(placeNear(b.doc, { story: 'body', path: [1], offset: 400 }), { story: 'body', path: [1], offset: 16 });
  assert.deepEqual(placeNear(b.doc, { story: 'body', path: [7], offset: 2 }), { story: 'body', path: [2], offset: 2 });
  const table = await joinDoc(new DocNet(), 't', await aloneFile());
  const rows = table.doc.body.blocks;
  const at4 = rows.findIndex((block) => block.type === 'table');
  const near = placeNear(table.doc, { story: 'body', path: [at4], offset: 3 })!;
  assert.deepEqual(near.path, [at4 - 1], 'before a table: the paragraph above it');
  assert.equal(near.offset, bodyText(table.doc).split('\n')[at4 - 1].length);
  assert.deepEqual(placeNear(table.doc, { story: 'body', path: [at4, 9, 9, 0], offset: 1 }), { story: 'body', path: [0], offset: 0 }, 'a cell that is no more: the start of the story');
  assert.deepEqual(placeNear(table.doc, { story: 'nowhere', path: [0], offset: 1 }), { story: 'body', path: [0], offset: 0 });
});

test('what the others do is held back while text is composed, and not for longer than half a minute, nor past the keyboard going away', () => {
  const calls: string[] = [];
  const waits: { run: () => void; ms: number; live: boolean }[] = [];
  const later = (run: () => void, ms: number): (() => void) => {
    const wait = { run, ms, live: true };
    waits.push(wait);
    return () => (wait.live = false);
  };
  const tick = (ms: number): void => waits.splice(0).forEach((wait) => (wait.live && wait.ms <= ms ? wait.run() : wait.live && waits.push(wait)));
  const hold = new CompositionHold({ pause: () => calls.push('pause'), resume: () => calls.push('resume') }, later);
  // A composition that ends: what arrived is applied after the text of the composition is in (a turn later).
  hold.begin();
  hold.begin();
  assert.deepEqual(calls, ['pause']);
  hold.end(true);
  assert.deepEqual(calls, ['pause'], 'not in the same turn');
  tick(0);
  assert.deepEqual(calls, ['pause', 'resume']);
  tick(60000);
  assert.deepEqual(calls, ['pause', 'resume'], 'the limit of a composition that ended does nothing');
  // One whose end is never told.
  hold.begin();
  tick(29999);
  assert.equal(hold.holding, true);
  tick(30000);
  assert.deepEqual(calls, ['pause', 'resume', 'pause', 'resume']);
  assert.equal(hold.holding, false);
  hold.end(true);
  tick(0);
  assert.deepEqual(calls.length, 4, 'an end that comes after all is no second release');
  // A new composition that starts before the release of the last one was due: one hold, released once.
  hold.begin();
  hold.end(true);
  hold.begin();
  tick(0);
  assert.equal(hold.holding, true);
  assert.deepEqual(calls.slice(4), ['pause'], 'the session was never let go in between');
  hold.end(false);
  assert.deepEqual(calls.slice(4), ['pause', 'resume']);
});

test('letters typed one after the other while the relay is away wait as few steps, and are undone as they would be', async () => {
  const { net, a, b } = await pair(['start', 'other']);
  a.transport.suspend();
  for (let i = 0; i < 600; i++) a.session.submit([{ t: 'text', at: at(0, 5 + i), x: 'abcdefghij'[i % 10] }], { key: 'typing' });
  assert.ok(a.session.pending <= 3, `${a.session.pending} steps wait for 600 letters`);
  // A letter elsewhere, or after seal(), starts a step of its own.
  a.session.submit([{ t: 'text', at: at(1, 0), x: 'x' }], { key: 'typing' });
  a.session.seal();
  a.session.submit([{ t: 'text', at: at(1, 1), x: 'y' }], { key: 'typing' });
  const waiting = a.session.pending;
  assert.ok(waiting >= 3 && waiting <= 5, String(waiting));
  for (let i = 0; i < 50; i++) {
    assert.ok(actDoc(b, [{ t: 'text', at: at(1, 5 + i), x: 'b' }]));
    net.drain();
  }
  a.transport.resume();
  net.drain();
  net.drain();
  assert.equal(a.session.pending, 0);
  assert.equal(bodyText(a.doc).split('\n')[0].length, 605);
  assertSameDoc([a, b], 'after catching up');
  assert.deepEqual(codes(a).concat(codes(b)), []);
  assert.ok(a.session.undo());
  net.drain();
  assert.equal(bodyText(b.doc).split('\n')[1], 'xother' + 'b'.repeat(50), 'the last letter by itself');
  assert.ok(a.session.undo());
  net.drain();
  assert.equal(bodyText(b.doc), 'start\nother' + 'b'.repeat(50), 'and all that was typed before it together');
  assertSameDoc([a, b], 'after undo');
});

test('the caret of the view is read through what the others did, and the others are listed with where they are', () => {
  const caret = { anchor: { story: 'body', path: [2], offset: 5 }, focus: { story: 'body', path: [2], offset: 5 } };
  // Text before it in its paragraph, a paragraph above it, text at the very place (which goes behind it), text in another story.
  assert.deepEqual(movedSelection(caret, [{ k: 'ins', s: 'body', c: [], row: 2, off: 1, rows: 0, last: 3 }])!.focus, { story: 'body', path: [2], offset: 8 });
  assert.deepEqual(movedSelection(caret, [{ k: 'ins', s: 'body', c: [], row: 0, off: 0, rows: 2, last: 0 }])!.focus, { story: 'body', path: [4], offset: 5 });
  assert.deepEqual(movedSelection(caret, [{ k: 'ins', s: 'body', c: [], row: 2, off: 5, rows: 0, last: 3 }])!.focus, { story: 'body', path: [2], offset: 5 });
  assert.deepEqual(movedSelection(caret, [{ k: 'ins', s: 'header:rId7', c: [], row: 0, off: 0, rows: 3, last: 0 }])!.focus, caret.focus);
  // The text it stood in went: it is where that text was. The paragraph before it was joined with its own.
  assert.deepEqual(movedSelection(caret, [{ k: 'del', s: 'body', c: [], r1: 2, o1: 2, r2: 2, o2: 9 }])!.focus, { story: 'body', path: [2], offset: 2 });
  assert.deepEqual(movedSelection(caret, [{ k: 'del', s: 'body', c: [], r1: 1, o1: 7, r2: 2, o2: 0 }])!.focus, { story: 'body', path: [1], offset: 12 });
  // In a table cell: the table moves with the paragraphs above it.
  const inCell = { anchor: { story: 'body', path: [3, 1, 0, 0], offset: 2 }, focus: { story: 'body', path: [3, 1, 0, 0], offset: 2 } };
  assert.deepEqual(movedSelection(inCell, [{ k: 'ins', s: 'body', c: [], row: 1, off: 0, rows: 1, last: 0 }])!.focus.path, [4, 1, 0, 0]);
  // Its story changed in a way that is not known: no place is known.
  assert.equal(movedSelection(caret, [{ k: 'opq', s: 'body' }]), undefined);
  // The blocks it was in were replaced as a whole (someone took a step back in its paragraph): it stays in the block that stands there.
  assert.deepEqual(movedSelection(caret, [{ k: 'blk', s: 'body', c: [], at: 2, n: 1, m: 1 }])!.focus, { story: 'body', path: [2], offset: 5 });
  assert.deepEqual(movedSelection(caret, [{ k: 'blk', s: 'body', c: [], at: 1, n: 2, m: 1 }])!.focus, { story: 'body', path: [1], offset: 5 });
  assert.deepEqual(movedSelection(caret, [{ k: 'blk', s: 'body', c: [], at: 0, n: 2, m: 5 }])!.focus, { story: 'body', path: [5], offset: 5 });
  assert.deepEqual(movedSelection(inCell, [{ k: 'blk', s: 'body', c: [], at: 3, n: 1, m: 1 }])!.focus, { story: 'body', path: [3], offset: 0 });
  assert.deepEqual(movedSelection(caret, [{ k: 'tch', s: 'body', c: [], r1: 2, r2: 2 } as never])!.focus, caret.focus);
  // The others: who said where it is, with a name (a guest has a number) and a colour that is one.
  const places = peerPlaces(
    [
      { id: 'c1', name: 'Ada', color: '#D93025', guest: 3, readOnly: false, presence: { story: 'body', anchor: { path: [0], offset: 1 }, focus: { path: [0], offset: 4 } } },
      { id: 'c2', color: 'red; background:url(x)', guest: 7, readOnly: true, presence: { story: 'body', anchor: { path: [1], offset: 0 }, focus: { path: [1], offset: 0 } } },
      { id: 'c3', name: 'Nowhere', color: '#188038', guest: 9, readOnly: false },
    ],
    (n) => `Guest ${n}`,
  );
  assert.deepEqual(
    places.map((place) => [place.id, place.label, place.color, place.focus.offset]),
    [
      ['c1', 'Ada', '#D93025', 4],
      ['c2', 'Guest 7', '#5F6368', 0],
    ],
  );
});

test(`steps that read as steps and are made of whatever: nobody leaves the room over them, and all agree (${SEEDS} runs)`, async () => {
  const bytes = await aloneFile();
  const odd: unknown[] = [0, -1, 1.5, 1e9, '', 'x', 'body', null, true, [], {}, [0], { a: 1 }, 'left', 'center', 'FF0000', '#FF0000', 'Normal', { rgb: 'zz' }, { theme: 'accent1', tint: 300 }, { numId: 7 }, { left: 'x' }, { before: -5 }, 31681];
  const keys = ['bold', 'italic', 'underline', 'strike', 'font', 'size', 'color', 'highlight', 'vertAlign', 'style', 'align', 'indent', 'spacing', 'numbering', 'keepNext', 'outlineLevel', 'widowControl', 'nothing'];
  for (let seed = 1; seed <= SEEDS; seed++) {
    const random = seeded(seed * 4099);
    const any = <T>(list: readonly T[]): T => list[Math.floor(random() * list.length)];
    const net = new DocNet();
    const a = await joinDoc(net, 'a', bytes);
    const b = await joinDoc(net, 'b', bytes);
    const evil = intruder(net, bytes);
    net.drain(random);
    const place = (): unknown[] => {
      const path = random() < 0.7 ? [Math.floor(random() * 8)] : [4, Math.floor(random() * 3), Math.floor(random() * 3), Math.floor(random() * 2)];
      const story = random() < 0.85 ? 'body' : any(['header:rIdH', 'header:none', 'footnote:1', 'endnote:-1', 'footer:x']);
      return random() < 0.1 ? [story, path, Math.floor(random() * 40), 1] : [story, path, Math.floor(random() * 40)];
    };
    const patch = (): Record<string, unknown> => {
      const out: Record<string, unknown> = {};
      for (let i = 0; i < 1 + Math.floor(random() * 3); i++) out[any(keys)] = any(odd);
      return out;
    };
    for (let i = 0; i < 40; i++) {
      const step: unknown = any<() => unknown>([
        () => ({ t: 'text', at: place(), to: place(), x: any(['a', 'two\nlines', '\t', '', 'x'.repeat(50)]), f: patch() }),
        () => ({ t: 'text', at: place(), x: 'plain' }),
        () => ({ t: 'del', at: place(), to: place() }),
        () => ({ t: 'split', at: place() }),
        () => ({ t: 'run', at: place(), to: place(), f: patch() }),
        () => ({ t: 'unrun', at: place(), to: place(), k: any(['bold', 'italic', 'underline', 'strike']) }),
        () => ({ t: 'para', at: place(), to: place(), f: patch() }),
        () => ({ t: 'cmd', name: 'toggleList', at: place(), to: place(), a: [any(['bullet', 'number'])] }),
        () => ({ t: 'cmd', name: 'indentList', at: place(), to: place(), a: [any([-8, -1, 1, 8])] }),
        () => ({ t: 'cmd', name: 'insertTable', at: place(), a: [1 + Math.floor(random() * 3), 1 + Math.floor(random() * 3)] }),
        () => ({ t: 'cmd', name: any(['insertRow']), at: place(), a: [any(['above', 'below'])] }),
        () => ({ t: 'cmd', name: any(['deleteRow', 'deleteColumn', 'insertPageBreak']), at: place(), a: [] }),
        () => ({ t: 'cmd', name: 'insertColumn', at: place(), a: [any(['left', 'right'])] }),
        () => ({ t: 'cmd', name: any(['deleteBackward', 'deleteForward']), at: place(), a: [any(['char', 'word'])] }),
        () => ({ t: 'cmd', name: 'insertHyperlink', at: place(), to: place(), a: [any(['https://example.com/', '#mark', 'mailto:a@example.com', 'javascript:1', 'x'])] }),
        () => ({ t: 'put', at: [place()[0], [Math.floor(random() * 8)], 0], n: 1 + Math.floor(random() * 3), x: [{ r: [{ c: ['put', 0, 1, 2, ['a', ' a ']], p: any([{ f: patch() }, { cs: patch() }, { fonts: patch() }, { lang: patch() }, patch()]), a: any([undefined, [['w:rsidR', '00AB12CD']], [['w14:paraId', '1234ABCD']]]) }, { e: any(['spellStart', 'bogus']) }, { b: [0, any(['0', '1', '9']), 'm'] }], m: any([undefined, { f: patch() }]), p: patch() }], h: any(['00000000', putDigest(a.doc.body.blocks.slice(0, 1))]) }),
      ])();
      evil.send({ t: 'op', base: a.session.rev, id: `evil:${i}`, ops: [step] });
      net.drain(random);
      // And now and then an honest change in between.
      if (random() < 0.2) actDoc(b, [{ t: 'text', at: at(0, 0), x: 'b' }]);
      net.drain(random);
    }
    for (const member of [a, b]) assert.equal(member.session.state, 'online', `seed ${seed}: ${member.name} stays in the room (${member.notices.map((notice) => notice.code).join()})`);
    assertSameDoc([a, b], `seed ${seed}`);
    const saved = await saveDoc(a);
    assert.equal(bodyText(await (await open(saved)).document!()), bodyText(a.doc), `seed ${seed}: the file opens and holds the text`);
    a.session.leave();
    b.session.leave();
  }
});
