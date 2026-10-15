// Comments of a document that is edited together: adding, answering, editing, resolving and
// removing them travel as commands of their own; everybody - and whoever joins later from a saved
// file - gives a new comment the same number and the same paragraph id, and writes the same file.

import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { join as joinPath } from 'node:path';
import { test } from 'node:test';
import { CollabDocHistory } from '../src/collab/doc/history';
import type { CommentWire, DocWire } from '../src/collab/doc/rebase';
import { traced, wiresOf } from '../src/collab/doc/trace';
import { decodeStep, encodeStep, putDigest, runWire } from '../src/collab/doc/wire';
import { contentVersion } from '../src/collab/protocol';
import { samePackage } from '../src/formats/xlsx/collab-stamp';
import { ZipArchive, open } from '../src/index';
import type { DocModel } from '../src/model/doc/types';
import * as raw from '../src/ops/doc/index';
import { listComments, type DocComment } from '../src/ops/doc/comments';
import { MAX_ROOM_COMMENT_CHARS, MAX_ROOM_COMMENT_PARAGRAPHS, refOf } from '../src/collab/doc/comments';
import { stable } from '../src/collab/wire';
import { DocNet, actDoc, assertSameDoc, at, bodyText, docFile, joinDoc, saveDoc, seeded, type DocMember } from './collab-doc-fixtures';

const ops = traced(raw);

/** The comments of a document as everybody sees them, with the ids they are known by. */
function commentsOf(doc: DocModel): string {
  const list = listComments(doc).map((comment: DocComment) => ({
    id: comment.id,
    author: comment.author,
    initials: comment.initials,
    date: comment.date,
    text: comment.text,
    parent: comment.parent,
    done: comment.done,
    start: comment.start,
    end: comment.end,
  }));
  const paraIds = doc.comments ? Array.from(doc.comments.items.entries()).map(([id, comment]) => [id, comment.paraId || null]) : [];
  return stable({ list, paraIds });
}

function assertSameComments(members: DocMember[], label: string): void {
  const texts = members.map((member) => commentsOf(member.doc));
  for (let i = 1; i < texts.length; i++) assert.equal(texts[i], texts[0], `${label}: ${members[i].name} holds other comments than ${members[0].name}`);
}

/** Runs a command of comments as the view does in a room: the command made by the commands of the document, and sent as a step. Returns its edit; undefined when it was refused. */
function comment(member: DocMember, name: 'addComment' | 'replyToComment' | 'editComment' | 'resolveComment' | 'removeComment', ...args: unknown[]): { id?: number; wire: CommentWire } | undefined {
  let edit: raw.CommentEdit;
  try {
    edit = (ops[name] as (doc: DocModel, ...rest: unknown[]) => raw.CommentEdit)(member.doc, ...args);
  } catch (error) {
    if (error instanceof RangeError) return undefined;
    throw error;
  }
  const wires = wiresOf(edit.ops) as DocWire[] | undefined;
  assert.ok(wires && wires.length === 1 && wires[0].t === 'cmt', `${name} travels as a command of comments`);
  if (!actDoc(member, wires)) return undefined;
  return { id: edit.id, wire: wires[0] as CommentWire };
}

const range = (row: number, from: number, to: number) => ({ anchor: { story: 'body', path: [row], offset: from }, focus: { story: 'body', path: [row], offset: to } });
const who = (author: string, date = '2026-10-15T10:00:00Z') => ({ author, date });

async function room(paragraphs: string[], names = ['a', 'b'], spec = {}): Promise<{ net: DocNet; members: DocMember[]; bytes: Uint8Array }> {
  const bytes = await docFile(paragraphs, spec);
  const net = new DocNet();
  const members: DocMember[] = [];
  for (const name of names) members.push(await joinDoc(net, name, bytes));
  net.drain();
  return { net, members, bytes };
}

test('a comment that is added, answered, edited, resolved and removed in a room is the same comment for everybody', async () => {
  const { net, members } = await room(['one two three', 'four five']);
  const [a, b] = members;
  const made = comment(a, 'addComment', range(0, 4, 7), 'Why two?', who('Ayşe Kaya'));
  assert.ok(made && made.id !== undefined);
  net.drain();
  assertSameComments(members, 'added');
  const list = listComments(b.doc);
  assert.equal(list.length, 1);
  assert.deepEqual([list[0].author, list[0].initials, list[0].date, list[0].text], ['Ayşe Kaya', 'AK', '2026-10-15T10:00:00Z', 'Why two?']);
  assert.equal(list[0].id, made!.id, 'the id the view got is the one everybody gave');
  assert.ok(comment(b, 'replyToComment', list[0].id, 'Because.', who('Bora'))!);
  net.drain();
  assertSameComments(members, 'answered');
  assert.ok(comment(a, 'editComment', list[0].id, 'Why two, really?'));
  net.drain();
  assert.ok(comment(b, 'resolveComment', list[0].id, true));
  net.drain();
  assertSameComments(members, 'edited and resolved');
  assert.deepEqual(listComments(a.doc).map((one) => [one.text, one.done]), [['Why two, really?', true], ['Because.', true]]);
  assert.ok(comment(b, 'removeComment', list[0].id));
  net.drain();
  assertSameComments(members, 'removed');
  assert.equal(listComments(a.doc).length, 0);
  assertSameDoc(members, 'the text');
  assert.equal(bodyText(a.doc), 'one two three\nfour five');
  for (const member of members) assert.deepEqual(member.notices, [], `${member.name} was told nothing`);
});

test('two participants who add a comment at the same time both have it; the numbers are given by the order of the relay, for everybody', async () => {
  const { net, members } = await room(['one two three', 'four five']);
  const [a, b] = members;
  const one = comment(a, 'addComment', range(0, 0, 3), 'from a', who('A'));
  const two = comment(b, 'addComment', range(1, 0, 4), 'from b', who('B'));
  assert.ok(one && two);
  // Each gave its own the first free number.
  assert.equal(one!.id, two!.id);
  net.drain();
  assertSameComments(members, 'both');
  const list = listComments(a.doc);
  assert.deepEqual(list.map((c) => c.text).sort(), ['from a', 'from b']);
  assert.equal(new Set(list.map((c) => c.id)).size, 2, 'two numbers');
  // The one that waited answers its own comment, which took another number in the meantime: it is named by its paragraph id.
  for (const member of members) assert.deepEqual(member.notices, []);
});

test('a comment that waits, and an answer to it, keep together when a comment of someone else gets its place first', async () => {
  const { net, members } = await room(['one two three']);
  const [a, b] = members;
  a.transport.suspend();
  const mine = comment(a, 'addComment', range(0, 0, 3), 'mine', who('A'))!;
  const answer = comment(a, 'replyToComment', mine.id, 'and my answer', who('A'));
  assert.ok(answer);
  const theirs = comment(b, 'addComment', range(0, 8, 13), 'theirs', who('B'))!;
  assert.equal(theirs.id, mine.id, 'the same number on both sides, for now');
  // B's step gets its place first.
  net.drain();
  a.transport.resume();
  net.drain();
  net.drain();
  assertSameComments(members, 'all');
  assert.deepEqual(a.notices, []);
  assert.equal(listComments(a.doc).filter((c) => c.text === 'theirs')[0].id, theirs.id, 'the first in the order keeps its number');
  const list = listComments(a.doc);
  const root = list.filter((c) => c.text === 'mine')[0];
  const reply = list.filter((c) => c.text === 'and my answer')[0];
  assert.equal(reply.parent, root.id, 'the answer belongs to the comment it was given to');
});

test('two participants edit one comment at the same time: the later edit stays; an edit of a comment someone removed is dropped and said', async () => {
  const { net, members } = await room(['one two three']);
  const [a, b] = members;
  const made = comment(a, 'addComment', range(0, 0, 3), 'first', who('A'))!;
  net.drain();
  b.transport.suspend();
  comment(a, 'editComment', made.id, 'by a');
  comment(b, 'editComment', made.id, 'by b');
  net.drain();
  b.transport.resume();
  net.drain();
  net.drain();
  assertSameComments(members, 'edited twice');
  assert.equal(listComments(a.doc)[0].text, 'by b', 'the edit that got its place later');
  assert.deepEqual(b.notices, []);
  // Removed by one, edited by the other: the removal first.
  b.transport.suspend();
  comment(a, 'removeComment', made.id);
  comment(b, 'editComment', made.id, 'too late');
  net.drain();
  b.transport.resume();
  net.drain();
  net.drain();
  assertSameComments(members, 'removed and edited');
  assert.equal(listComments(a.doc).length, 0);
  assert.deepEqual(b.notices, [{ code: 'dropped', count: 1 }]);
  // Edited first, then removed: the removal takes the edited comment away.
  const again = comment(a, 'addComment', range(0, 0, 3), 'again', who('A'))!;
  net.drain();
  a.transport.suspend();
  comment(a, 'removeComment', again.id);
  comment(b, 'editComment', again.id, 'edited');
  net.drain();
  a.transport.resume();
  net.drain();
  net.drain();
  assertSameComments(members, 'edited and removed');
  assert.equal(listComments(a.doc).length, 0);
  assert.deepEqual(a.notices, []);
});

test('a comment whose whole range someone else deleted first is dropped and said; one that keeps part of its range is about that part', async () => {
  const { net, members } = await room(['one two three four']);
  const [a, b] = members;
  // B deletes "two " while A comments on "two".
  a.transport.suspend();
  assert.ok(actDoc(b, [{ t: 'del', at: at(0, 4), to: at(0, 8) }]));
  comment(a, 'addComment', range(0, 4, 7), 'on two', who('A'));
  net.drain();
  a.transport.resume();
  net.drain();
  net.drain();
  assertSameComments(members, 'gone');
  assert.equal(listComments(a.doc).length, 0, 'no comment about nothing');
  assert.ok(a.notices.some((notice) => notice.code === 'dropped'), 'said');
  // Now part of a range: B deletes "three" (with its space before), A comments on "one three".
  a.transport.suspend();
  assert.ok(actDoc(b, [{ t: 'del', at: at(0, 3), to: at(0, 9) }]));
  comment(a, 'addComment', range(0, 0, 9), 'on one three', who('A'));
  net.drain();
  a.transport.resume();
  net.drain();
  net.drain();
  assert.ok(a.notices.some((notice) => notice.code === 'partial'), 'said');
  assertSameComments(members, 'part');
  const list = listComments(a.doc);
  assert.equal(list.length, 1);
  assert.equal(bodyText(a.doc), 'one four');
  assert.deepEqual([list[0].start!.offset, list[0].end!.offset], [0, 3], 'about what is left of its range');
});

test('a comment that was added first keeps its marks when someone deletes its text: it is about the place where the text was, for everybody', async () => {
  const { net, members } = await room(['one two three']);
  const [a, b] = members;
  b.transport.suspend();
  comment(a, 'addComment', range(0, 4, 7), 'on two', who('A'));
  assert.ok(actDoc(b, [{ t: 'del', at: at(0, 3), to: at(0, 7) }]));
  net.drain();
  b.transport.resume();
  net.drain();
  net.drain();
  assertSameComments(members, 'all');
  assertSameDoc(members, 'text');
  assert.equal(bodyText(a.doc), 'one three');
  const list = listComments(a.doc);
  assert.equal(list.length, 1, 'the comment stays');
  assert.deepEqual([list[0].start!.offset, list[0].end!.offset], [3, 3], 'about the place where its text was');
  assert.deepEqual(b.notices, []);
});

test('undo and redo of commands of comments are steps like any other', async () => {
  const { net, members } = await room(['one two three']);
  const [a, b] = members;
  const history = new CollabDocHistory<string>(a.session);
  const made = comment(a, 'addComment', range(0, 0, 3), 'note', who('A'))!;
  net.drain();
  assert.equal(listComments(b.doc).length, 1);
  assert.ok(history.undo() !== undefined || a.notices.length === 0);
  net.drain();
  assertSameComments(members, 'undone');
  assert.equal(listComments(b.doc).length, 0, 'the undoing took the comment away for everybody');
  // An edit is undone into the text it had.
  const again = comment(a, 'addComment', range(0, 0, 3), 'before', who('A'))!;
  net.drain();
  comment(a, 'editComment', again.id, 'after');
  net.drain();
  history.undo();
  net.drain();
  assertSameComments(members, 'edit undone');
  assert.equal(listComments(b.doc)[0].text, 'before');
  // Resolving is undone into an open thread.
  comment(a, 'resolveComment', again.id, true);
  net.drain();
  assert.equal(listComments(b.doc)[0].done, true);
  history.undo();
  net.drain();
  assert.equal(listComments(b.doc)[0].done, false);
  assertSameComments(members, 'resolving undone');
  // A comment that was removed is not brought back: that is said.
  comment(a, 'removeComment', again.id);
  net.drain();
  const before = a.notices.length;
  assert.equal(history.undo(), undefined);
  assert.ok(a.notices.slice(before).some((notice) => notice.code === 'undo-none'));
  void made;
});

test('who joins from a saved file gives a new comment the number and the id everybody gives it, and writes the same file', async () => {
  const { net, members } = await room(['one two three', 'four five six']);
  const [a, b] = members;
  comment(a, 'addComment', range(0, 0, 3), 'first', who('A'));
  net.drain();
  const first = listComments(a.doc)[0];
  comment(b, 'replyToComment', first.id, 'answer', who('B'));
  net.drain();
  comment(a, 'resolveComment', first.id, true);
  net.drain();
  const stored = await saveDoc(a);
  a.session.announceSaved([contentVersion(stored)], a.session.rev);
  net.drain();
  const late = await joinDoc(net, 'late', stored);
  net.drain();
  assert.equal(late.session.state, 'online');
  assertSameComments([a, late], 'read from the file');
  // New comments by who joined late and by who was there; an answer, an edit and a removal named by the file's ids.
  comment(late, 'addComment', range(1, 0, 4), 'late one', who('L'));
  net.drain();
  comment(a, 'addComment', range(1, 5, 9), 'a again', who('A'));
  net.drain();
  const all = listComments(late.doc);
  comment(late, 'replyToComment', all.filter((c) => c.text === 'a again')[0].id, 'late answer', who('L'));
  net.drain();
  comment(b, 'editComment', first.id, 'first, edited');
  net.drain();
  comment(late, 'resolveComment', first.id, false);
  net.drain();
  assertSameComments([a, b, late], 'after');
  const one = await saveDoc(a);
  const two = await saveDoc(late);
  const three = await saveDoc(b);
  assert.ok(await samePackage(one, three), 'who were there write the same file');
  if (!(await samePackage(one, two))) {
    const zips = [ZipArchive.read(one), ZipArchive.read(two)];
    const differing: string[] = [];
    for (const name of zips[0].names()) {
      const x = Buffer.from(await zips[0].read(name)).toString('utf8');
      const y = zips[1].names().indexOf(name) >= 0 ? Buffer.from(await zips[1].read(name)).toString('utf8') : '<missing>';
      if (x !== y) {
        let i = 0;
        while (i < x.length && x[i] === y[i]) i++;
        differing.push(`${name} at ${i}: …${x.slice(Math.max(0, i - 80), i + 80)}… against …${y.slice(Math.max(0, i - 80), i + 80)}…`);
      }
    }
    assert.fail(`who joined from the saved file writes another file:\n${differing.join('\n')}`);
  }
});

test('the file a room writes with comments is one an independent reader takes: comments, their marks, the parts beside them', async (t) => {
  try {
    execFileSync('python3', ['-I', '-c', 'import zipfile, xml.etree.ElementTree'], { stdio: 'ignore', timeout: 60000 });
  } catch {
    return t.skip('python3 is not there');
  }
  const { net, members } = await room(['one two three', 'four five six']);
  const [a, b] = members;
  comment(a, 'addComment', range(0, 0, 3), 'first', who('A'));
  comment(b, 'addComment', { story: 'body', path: [1], offset: 6 }, 'at a caret', who('B'));
  net.drain();
  const first = listComments(a.doc).filter((c) => c.text === 'first')[0];
  comment(b, 'replyToComment', first.id, 'answer', who('B'));
  net.drain();
  comment(a, 'resolveComment', first.id, true);
  net.drain();
  const dir = mkdtempSync(joinPath(process.cwd(), 'test', '.comments-'));
  try {
    writeFileSync(joinPath(dir, 'file.docx'), await saveDoc(a));
    const found = JSON.parse(execFileSync('python3', ['-I', joinPath(process.cwd(), 'test', 'tools', 'package_check.py'), dir], { encoding: 'utf8', timeout: 120000 }));
    assert.deepEqual(found, []);
    const zip = ZipArchive.read(await saveDoc(a));
    const names = zip.names();
    assert.ok(names.indexOf('word/comments.xml') >= 0 && names.indexOf('word/commentsExtended.xml') >= 0, names.join(' '));
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});

test('commands of comments that arrive are not trusted: what is no command is refused as it is read, what the document does not take is passed over by everybody', async () => {
  const good = (wire: Record<string, unknown>): unknown => decodeStep([wire]);
  const add = { t: 'cmt', k: 'add', at: ['body', [0], 0], to: ['body', [0], 3], x: 'x', u: 'A', i: 'A', d: '2026-10-15T10:00:00Z', q: '1A2B3C4D' };
  assert.doesNotThrow(() => good(add));
  for (const [what, bad] of [
    ['no kind', { ...add, k: 'nope' }],
    ['a text too long', { ...add, x: 'x'.repeat(32768) }],
    ['a text with a character no file holds', { ...add, x: 'a\u0001b' }],
    ['a carriage return', { ...add, x: 'a\rb' }],
    ['an author with a line', { ...add, u: 'a\nb' }],
    ['no author', { ...add, u: '' }],
    ['an author too long', { ...add, u: 'a'.repeat(256) }],
    ['no date', { ...add, d: 'yesterday' }],
    ['a paragraph id in lower case', { ...add, q: '1a2b3c4d' }],
    ['a paragraph id above 7FFFFFFF', { ...add, q: '80000000' }],
    ['a paragraph id 0', { ...add, q: '00000000' }],
    ['a place that is none', { ...add, at: ['body', [-1], 0] }],
    ['a comment named by nothing', { t: 'cmt', k: 'edit', c: 7, x: 'x' }],
    ['a comment named by a bad digest', { t: 'cmt', k: 'edit', c: [7, 'XYZ'], x: 'x' }],
    ['a resolution that is no yes or no', { t: 'cmt', k: 'done', c: '1A2B3C4D', v: true }],
    ['a text of more than 1.000 paragraphs', { ...add, x: '\n'.repeat(1000) }],
  ] as [string, Record<string, unknown>][]) {
    assert.throws(() => good(bad), { name: 'WireError' }, what);
  }
  // Many commands of comments in one step: each reads the whole document.
  assert.throws(() => decodeStep(Array.from({ length: 9 }, () => add)), { name: 'WireError' });
  // A thousand paragraphs are taken.
  assert.doesNotThrow(() => good({ ...add, x: '\n'.repeat(999) }));

  // What the document does not take: a comment it has not, an id it holds, a place it has not. Passed over by
  // everybody; the room goes on.
  const { net, members } = await room(['one two three'], ['a', 'b'], { body: '<w:p w14:paraId="1A2B3C4D"><w:r><w:t>held</w:t></w:r></w:p>' });
  const [a, b] = members;
  const send = (wire: Record<string, unknown>): void => {
    (a.transport as unknown as { send(text: string): void }).send(JSON.stringify({ t: 'op', base: a.session.rev, id: 'evil:' + Math.random(), ops: [wire] }));
    net.drain();
  };
  send({ t: 'cmt', k: 'edit', c: '0BADF00D', x: 'x' });
  send({ t: 'cmt', k: 'del', c: [99, '00000000'] });
  send({ ...add, q: '1A2B3C4D' });
  send({ ...add, at: ['body', [5], 0], to: ['body', [5], 1] });
  assert.equal(listComments(b.doc).length, 0);
  assert.ok(b.notices.some((notice) => notice.code === 'bad-step'));
  assert.ok(comment(b, 'addComment', range(0, 0, 3), 'still works', who('B')));
  net.drain();
  assertSameComments(members, 'after the evil ones');
  assert.equal(listComments(a.doc).length, 1);
});

test('a command of comments that arrives is run with what it says: who wrote it and when, not the clock or the name of whoever runs it', async () => {
  const bytes = await docFile(['one two']);
  const net = new DocNet();
  const a = await joinDoc(net, 'a', bytes);
  const b = await joinDoc(net, 'b', bytes);
  net.drain();
  raw.setTracking(b.doc, { author: 'Somebody Else', now: () => new Date(Date.UTC(2001, 0, 1)) });
  comment(a, 'addComment', range(0, 0, 3), 'x', who('Ayşe', '2026-10-15T11:12:13Z'));
  net.drain();
  const list = listComments(b.doc);
  assert.deepEqual([list[0].author, list[0].date], ['Ayşe', '2026-10-15T11:12:13Z']);
});

test('a step of comments goes through its own text: what is sent is what the others read', async () => {
  const { members } = await room(['one two']);
  const [a] = members;
  const edit = ops.addComment(a.doc, range(0, 0, 3), 'x\r\ny', who('A'));
  const wires = wiresOf(edit.ops) as DocWire[];
  const back = decodeStep(JSON.parse(JSON.stringify(encodeStep(wires))));
  assert.deepEqual(back, wires);
  assert.equal((wires[0] as Extract<CommentWire, { k: 'add' }>).x, 'x\ny', 'the text is cleaned as the command cleans it');
  // And run on the document it was made on, it is what the view got.
  assert.equal(runWire(a.doc, wires[0]).ops.length, edit.ops.length);
});

test('a comment at a caret is about the word at its place when it runs; one whose place someone deleted is dropped and said; text put into a range makes it grow', async () => {
  const { net, members } = await room(['one two three']);
  const [a, b] = members;
  // The caret in "three", while "two " goes before it: the comment is about "three" still.
  a.transport.suspend();
  assert.ok(actDoc(b, [{ t: 'del', at: at(0, 4), to: at(0, 8) }]));
  comment(a, 'addComment', { story: 'body', path: [0], offset: 10 }, 'on three', who('A'));
  net.drain();
  a.transport.resume();
  net.drain();
  net.drain();
  assertSameComments(members, 'caret');
  let list = listComments(a.doc);
  assert.equal(bodyText(a.doc), 'one three');
  assert.deepEqual([list[0].start!.offset, list[0].end!.offset], [4, 9], 'the word the caret was in');
  // The caret in a word that someone deleted around it: dropped.
  a.transport.suspend();
  assert.ok(actDoc(b, [{ t: 'del', at: at(0, 0), to: at(0, 4) }]));
  comment(a, 'addComment', { story: 'body', path: [0], offset: 2 }, 'on one', who('A'));
  net.drain();
  a.transport.resume();
  net.drain();
  net.drain();
  assertSameComments(members, 'caret gone');
  assert.equal(listComments(a.doc).length, 1);
  assert.ok(a.notices.some((notice) => notice.code === 'dropped'));
  // Text typed into the range by someone else: the comment covers it too.
  a.transport.suspend();
  assert.ok(actDoc(b, [{ t: 'text', at: at(0, 2), x: 'XX' }]));
  comment(a, 'addComment', range(0, 0, 5), 'whole', who('A'));
  net.drain();
  a.transport.resume();
  net.drain();
  net.drain();
  assertSameComments(members, 'grown');
  list = listComments(a.doc).filter((c) => c.text === 'whole');
  assert.equal(bodyText(a.doc), 'thXXree');
  assert.deepEqual([list[0].start!.offset, list[0].end!.offset], [0, 7]);
});

test('undoing a comment someone else answered since would take their answer along: the undoing is not done, and that is said', async () => {
  const { net, members } = await room(['one two three']);
  const [a, b] = members;
  const history = new CollabDocHistory<string>(a.session);
  const made = comment(a, 'addComment', range(0, 0, 3), 'mine', who('A'))!;
  net.drain();
  comment(b, 'replyToComment', made.id, 'theirs', who('B'));
  net.drain();
  assert.equal(history.undo(), undefined);
  net.drain();
  assert.ok(a.notices.some((notice) => notice.code === 'undo-gone'), 'said');
  assert.equal(listComments(b.doc).length, 2, 'both comments stay');
  assertSameComments(members, 'after');
  // An answer of one's own is undone alone, whatever else the thread holds.
  assert.ok(comment(a, 'replyToComment', made.id, 'my answer', who('A')));
  net.drain();
  history.undo();
  net.drain();
  assertSameComments(members, 'answer undone');
  assert.deepEqual(listComments(b.doc).map((c) => c.text), ['mine', 'theirs']);
});

test('a deletion, a format and an Enter in a paragraph that holds the marks of a comment are undone in a room, with the marks where they were', async () => {
  const { net, members } = await room(['one two three', 'four five']);
  const [a, b] = members;
  const history = new CollabDocHistory<string>(a.session);
  comment(b, 'addComment', range(0, 4, 7), 'note', who('B'));
  net.drain();
  const before = commentsOf(a.doc);
  const text = bodyText(a.doc);
  // A deletion over the commented word, a format over it, an Enter inside it: each undone.
  assert.ok(actDoc(a, [{ t: 'del', at: at(0, 2), to: at(0, 9) }]));
  net.drain();
  assert.ok(history.undo() !== undefined, 'the deletion is undone');
  net.drain();
  assert.equal(bodyText(b.doc), text);
  assert.equal(commentsOf(b.doc), before, 'the comment is about its text again');
  assert.ok(actDoc(a, [{ t: 'run', at: at(0, 0), to: at(0, 9), f: { bold: true } }]));
  net.drain();
  assert.ok(history.undo() !== undefined, 'the format is undone');
  net.drain();
  assert.ok(actDoc(a, [{ t: 'split', at: at(0, 5) }]));
  net.drain();
  assert.ok(history.undo() !== undefined, 'the Enter is undone');
  net.drain();
  assertSameDoc(members, 'after');
  assertSameComments(members, 'after');
  assert.equal(commentsOf(b.doc), before);
  assert.deepEqual(a.notices, []);
});

test('an undoing that would bring back the marks of a comment someone removed, or put back a paragraph someone commented on since, is not done, and that is said', async () => {
  const { net, members } = await room(['one two three']);
  const [a, b] = members;
  const history = new CollabDocHistory<string>(a.session);
  const made = comment(b, 'addComment', range(0, 4, 7), 'note', who('B'))!;
  net.drain();
  assert.ok(actDoc(a, [{ t: 'del', at: at(0, 0), to: at(0, 4) }]));
  net.drain();
  comment(b, 'removeComment', made.id);
  net.drain();
  assert.equal(history.undo(), undefined);
  assert.ok(a.notices.some((notice) => notice.code === 'undo-gone'));
  assert.equal(listComments(a.doc).length, 0, 'no mark of the removed comment came back');
  // Someone comments on the paragraph after a deletion: the deletion is not undone over the new comment.
  assert.ok(actDoc(a, [{ t: 'del', at: at(0, 0), to: at(0, 2) }]));
  net.drain();
  comment(b, 'addComment', range(0, 0, 3), 'new', who('B'));
  net.drain();
  const said = a.notices.length;
  assert.equal(history.undo(), undefined);
  assert.ok(a.notices.slice(said).some((notice) => notice.code === 'undo-gone'));
  assertSameDoc(members, 'after');
  assertSameComments(members, 'after');
  assert.equal(listComments(b.doc).length, 1);
});

test('a put that brings marks of a comment other than the ones it replaces is refused by everybody', async () => {
  const { net, members } = await room(['one two three', 'four']);
  const [a] = members;
  comment(a, 'addComment', range(0, 4, 7), 'note', who('A'));
  net.drain();
  const id = String(listComments(a.doc)[0].id);
  // A second start of the comment in another paragraph, or the mark of a comment that is none.
  for (const r of [[{ k: [0, id] }, { c: ['four'] }], [{ c: ['four'] }, { k: [1, '77'] }], [{ c: ['four', { r: Number(id) }] }]]) {
    const wire = { t: 'put', at: at(1, 0), n: 1, x: [{ r }], h: putDigest(a.doc.body.blocks.slice(1, 2)) } as unknown as DocWire;
    assert.throws(() => runWire(a.doc, decodeStep([encodeStep([wire])[0]])[0]), RangeError, JSON.stringify(r));
  }
});

test('a put that brings the marks of comments in another order than the blocks it replaces hold them is refused by everybody: reversed, numbers swapped, crossing, across paragraphs', async () => {
  const FCOM = (id: number): string => `<w:comment w:id="${id}" w:author="F${id}" w:date="2025-01-01T00:00:00Z"><w:p><w:r><w:t>c${id}</w:t></w:r></w:p></w:comment>`;
  const body =
    '<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>aaa</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r><w:commentRangeStart w:id="2"/><w:r><w:t>bbb</w:t></w:r><w:commentRangeEnd w:id="2"/><w:r><w:commentReference w:id="2"/></w:r></w:p>' +
    '<w:p><w:commentRangeStart w:id="3"/><w:r><w:t>ccc</w:t></w:r></w:p><w:p><w:r><w:t>ddd</w:t></w:r><w:commentRangeEnd w:id="3"/><w:r><w:commentReference w:id="3"/></w:r></w:p>';
  const { members } = await room(['x'], ['a'], { body, comments: FCOM(1) + FCOM(2) + FCOM(3) });
  const [a] = members;
  const k = (edge: 0 | 1, n: string) => ({ k: [edge, n] });
  const t = (text: string) => ({ c: [text] });
  const ref = (n: number) => ({ c: [{ r: n }] });
  const cases: [string, number, number, unknown[][]][] = [
    ['reversed in one paragraph', 1, 1, [[k(1, '1'), ref(1), t('aaa'), k(0, '1'), k(0, '2'), t('bbb'), k(1, '2'), ref(2)]]],
    ['the numbers of two comments swapped', 1, 1, [[k(0, '2'), t('aaa'), k(1, '2'), ref(2), k(0, '1'), t('bbb'), k(1, '1'), ref(1)]]],
    ['crossing ranges', 1, 1, [[k(0, '1'), t('aaa'), k(0, '2'), k(1, '1'), ref(1), t('bbb'), k(1, '2'), ref(2)]]],
    ['reversed across paragraphs', 2, 2, [[t('ccc'), k(1, '3'), ref(3)], [k(0, '3'), t('ddd')]]],
  ];
  for (const [what, row, n, paragraphs] of cases) {
    const wire = { t: 'put', at: at(row, 0), n, x: paragraphs.map((r) => ({ r })), h: putDigest(a.doc.body.blocks.slice(row, row + n)) } as unknown as DocWire;
    assert.throws(() => runWire(a.doc, decodeStep([encodeStep([wire])[0]])[0]), RangeError, what);
  }
  // The same marks in their order, in other paragraphs (what undoing an Enter brings back): taken.
  const joined = { t: 'put', at: at(2, 0), n: 2, x: [{ r: [k(0, '3'), t('cccddd'), k(1, '3'), ref(3)] }], h: putDigest(a.doc.body.blocks.slice(2, 4)) } as unknown as DocWire;
  assert.doesNotThrow(() => runWire(a.doc, decodeStep([encodeStep([joined])[0]])[0]));
});

test('an edit of a comment is held to the budget of the comments of the room: growing beyond it is refused, shrinking and growing within it are not; an empty comment is one paragraph, as it is counted', async () => {
  const C = (id: number, text: string): string => `<w:comment w:id="${id}" w:author="F" w:date="2026-01-01T00:00:00Z"><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:comment>`;
  const M = (id: number): string => `<w:commentRangeStart w:id="${id}"/><w:r><w:t>a</w:t></w:r><w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`;
  const big = await (await open(await docFile(['one'], { body: '<w:p>' + M(0) + M(1) + '</w:p>', comments: C(0, 'x'.repeat(MAX_ROOM_COMMENT_CHARS - 50)) + C(1, 's') }))).document!();
  const edit = (doc: DocModel, id: number, x: string): DocWire => ({ t: 'cmt', k: 'edit', c: refOf(doc, id), x });
  assert.throws(() => runWire(big, edit(big, 1, 'y'.repeat(100))), /too large to be edited together/, 'beyond the characters');
  assert.doesNotThrow(() => runWire(big, edit(big, 1, 'y'.repeat(40))), 'within them');
  assert.doesNotThrow(() => runWire(big, edit(big, 0, 'short')), 'a comment that shrinks');
  // Paragraphs: a comment of MAX - 1 paragraphs and an empty one make the budget; the empty one may say one line.
  const lines = Array.from({ length: MAX_ROOM_COMMENT_PARAGRAPHS - 1 }, () => '<w:p/>').join('');
  const full = await (await open(await docFile(['one'], { body: '<w:p>' + M(0) + M(1) + '</w:p>', comments: `<w:comment w:id="0" w:author="F" w:date="2026-01-01T00:00:00Z">${lines}</w:comment><w:comment w:id="1" w:author="F" w:date="2026-01-01T00:00:00Z"><w:p/></w:comment>` }))).document!();
  assert.doesNotThrow(() => runWire(full, edit(full, 1, 'one line')), 'an empty comment says one line: no paragraph more');
  assert.throws(() => runWire(full, edit(full, 1, 'two\nlines')), /too large to be edited together/, 'one paragraph more');
});

test('the paragraph id a comment of the room is known by is counted where the file will hold it: a paragraph that would bring it a second time is refused by everybody, also by who joined from the saved file', async () => {
  const { net, members } = await room(['one two three', 'four']);
  const [a] = members;
  const made = comment(a, 'addComment', range(0, 0, 3), 'mine', who('A'))!;
  net.drain();
  const paraId = (made.wire as Extract<CommentWire, { k: 'add' }>).q;
  const stored = await saveDoc(a);
  a.session.announceSaved([contentVersion(stored)], a.session.rev);
  net.drain();
  const late = await joinDoc(net, 'late', stored);
  net.drain();
  for (const member of [a, late]) {
    const wire: DocWire = { t: 'put', at: at(1, 0), n: 1, x: [{ a: [['w14:paraId', paraId]], r: [{ c: ['four'] }] }], h: putDigest(member.doc.body.blocks.slice(1, 2)) };
    assert.throws(() => runWire(member.doc, wire), RangeError, member.name);
    // Another id goes in.
    const other: DocWire = { ...wire, x: [{ a: [['w14:paraId', '0A0B0C0D']], r: [{ c: ['four'] }] }] } as DocWire;
    assert.doesNotThrow(() => runWire(member.doc, other), member.name);
  }
});

const FILE_COMMENT = '<w:comment w:id="0" w:author="F" w:date="2026-01-01T00:00:00Z"><w:p><w:r><w:t>file</w:t></w:r></w:p></w:comment>';
const FILE_BODY = '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>abc</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>';

test('a comment of the file that has no paragraph id gets the same one from whoever answers it first: what waits behind an answer finds it', async () => {
  const { net, members } = await room(['one two'], ['a', 'b'], { body: FILE_BODY, comments: FILE_COMMENT });
  const [a, b] = members;
  const root = listComments(a.doc)[0].id;
  assert.equal(a.doc.comments!.items.get(root)!.paraId, undefined, 'the file gives it none');
  a.transport.suspend();
  assert.ok(comment(a, 'replyToComment', root, 'A1', who('A')));
  assert.ok(comment(a, 'editComment', root, 'edited by A'));
  assert.ok(comment(b, 'replyToComment', root, 'B1', who('B')));
  net.drain();
  a.transport.resume();
  net.drain();
  net.drain();
  assert.deepEqual(a.notices, [], 'nothing of A was dropped');
  assertSameComments(members, 'answered at once');
  assert.equal(listComments(b.doc).filter((c) => c.id === root)[0].text, 'edited by A');
});

test('the thread of a file comment without ids is resolved at once by two participants: its comments get the same ids, and what waits behind finds them', async () => {
  const two = FILE_COMMENT + '<w:comment w:id="1" w:author="G" w:date="2026-01-02T00:00:00Z"><w:p><w:r><w:t>other</w:t></w:r></w:p></w:comment>';
  const body = FILE_BODY + '<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>def</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>';
  const { net, members } = await room(['one two'], ['a', 'b'], { body, comments: two });
  const [a, b] = members;
  a.transport.suspend();
  assert.ok(comment(a, 'resolveComment', 1, true));
  assert.ok(comment(a, 'editComment', 1, 'edited'));
  assert.ok(comment(b, 'resolveComment', 1, true));
  net.drain();
  a.transport.resume();
  net.drain();
  net.drain();
  assert.deepEqual(a.notices, []);
  assertSameComments(members, 'resolved at once');
  assert.equal(listComments(b.doc).filter((c) => c.id === 1)[0].text, 'edited');
});

test('the comments of a document in a room have a budget: a step that would go beyond it is passed over by everybody, and the room goes on', async () => {
  const { net, members } = await room(['one two three four']);
  const [a, b] = members;
  const send = (wires: Record<string, unknown>[]): void => {
    (a.transport as unknown as { send(text: string): void }).send(JSON.stringify({ t: 'op', base: a.session.rev, id: 'evil:' + Math.random(), ops: wires }));
    net.drain();
  };
  // 8 comments of 1.000 paragraphs a step: the budget of 50.000 paragraphs is met within seven steps, then nothing more comes in.
  let n = 0;
  const x = '\n'.repeat(999) + 'a';
  for (let step = 0; step < 9; step++) send(Array.from({ length: 8 }, () => ({ t: 'cmt', k: 'add', at: ['body', [0], 0], to: ['body', [0], 3], x, u: 'E', i: '', d: '2026-10-15T10:00:00Z', q: (0x10000000 + n++).toString(16).toUpperCase() })));
  let paragraphs = 0;
  for (const c of listComments(b.doc)) paragraphs += c.text.split('\n').length;
  assert.ok(paragraphs <= MAX_ROOM_COMMENT_PARAGRAPHS, `${paragraphs} paragraphs of comments`);
  assert.ok(b.notices.some((notice) => notice.code === 'bad-step'));
  assertSameComments(members, 'after the budget');
  // An edit that would grow a comment beyond it is refused here as well; a short one goes in.
  const first = listComments(a.doc)[0].id;
  assert.equal(comment(a, 'editComment', first, x + '\n' + x), undefined);
  assert.ok(comment(a, 'editComment', first, 'short'));
  net.drain();
  assertSameComments(members, 'after the edits');
});

test('a document whose comments have the largest number there is takes no new comment in a room, as it takes none alone: said, and nothing changes', async () => {
  for (const id of ['2147483647', '2147483646']) {
    const bytes = await docFile(['one two'], { body: `<w:p><w:commentRangeStart w:id="${id}"/><w:r><w:t>x</w:t></w:r><w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r></w:p>`, comments: `<w:comment w:id="${id}" w:author="F"><w:p><w:r><w:t>c</w:t></w:r></w:p></w:comment>` });
    const doc = await (await open(bytes)).document!();
    const wire: DocWire = { t: 'cmt', k: 'add', at: at(0, 0), to: at(0, 3), x: 'y', u: 'A', i: '', d: '2026-10-15T10:00:00Z', q: '1A2B3C4D' };
    if (id === '2147483647') {
      assert.throws(() => runWire(doc, wire), /no room for another comment/);
      assert.throws(() => raw.addComment(doc, range(0, 0, 3), 'alone', who('A')), RangeError, 'nor alone');
    } else assert.equal((runWire(doc, wire) as raw.CommentEdit).id, 2147483647);
  }
});
