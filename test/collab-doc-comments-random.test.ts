// Comments of a document that is edited together, at random: three participants add, answer,
// edit, resolve and remove comments while they type, delete, undo and lose their connection;
// then one joins from a saved file and goes on. Everybody ends with the same document, the same
// comments under the same numbers and ids, and writes the same file - one an independent reader
// takes.

import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { join as joinPath } from 'node:path';
import { test } from 'node:test';
import { CollabDocHistory } from '../src/collab/doc/history';
import { contentVersion } from '../src/collab/protocol';
import { stable } from '../src/collab/wire';
import { samePackage } from '../src/formats/xlsx/collab-stamp';
import { ZipArchive } from '../src/index';
import type { Block, DocModel } from '../src/model/doc/types';
import { listComments } from '../src/ops/doc/comments';
import * as raw from '../src/ops/doc/index';
import { aloneFile, make, ops } from './collab-doc-changes';
import { DocNet, assertSameDoc, docFile, int, joinDoc, pick, saveDoc, seeded, type DocMember } from './collab-doc-fixtures';

const SEEDS = Math.max(1, Math.min(Number(process.env.COLLAB_SEEDS) || 40, 100000));

function commentsOf(doc: DocModel): string {
  const list = listComments(doc).map((c) => [c.id, c.author, c.date, c.text, c.parent, c.done, c.start, c.end]);
  const paraIds = doc.comments ? Array.from(doc.comments.items.entries()).map(([id, c]) => [id, c.paraId || null]) : [];
  return stable({ list, paraIds });
}

function paragraphs(blocks: Block[], base: number[], out: { path: number[]; length: number }[]): void {
  blocks.forEach((block, index) => {
    if (block.type === 'paragraph') out.push({ path: base.concat([index]), length: raw.paragraphLength(block) });
    else if (block.type === 'table') block.rows.forEach((row, r) => row.cells.forEach((cell, c) => paragraphs(cell.blocks, base.concat([index, r, c]), out)));
  });
}

/** A command of comments made from random numbers on what a participant holds now. */
function commentChange(member: DocMember, random: () => number): raw.DocEdit | undefined {
  const doc = member.doc;
  const list = listComments(doc);
  const roll = random();
  const author = { author: pick(random, ['Ayşe', 'Bora', 'Can Demir']), date: `2026-10-15T1${int(random, 0, 9)}:00:0${int(random, 0, 9)}Z` };
  const text = pick(random, ['note', 'İki satır\nikinci', 'x', 'a longer remark']);
  if (roll < 0.4 || list.length === 0) {
    const all: { path: number[]; length: number }[] = [];
    paragraphs(doc.body.blocks, [], all);
    const at = pick(random, all);
    const from = int(random, 0, at.length);
    const to = random() < 0.3 ? from : int(random, from, at.length);
    return ops.addComment(doc, { anchor: { story: 'body', path: at.path, offset: from }, focus: { story: 'body', path: at.path, offset: to } }, text, author);
  }
  const target = pick(random, list).id;
  if (roll < 0.6) return ops.replyToComment(doc, target, text, author);
  if (roll < 0.75) return ops.editComment(doc, target, text + ' (edited)');
  if (roll < 0.9) return ops.resolveComment(doc, target, random() < 0.6);
  return ops.removeComment(doc, target);
}

/** The parts two files hold differently, with where they first differ. */
async function differingParts(one: Uint8Array, two: Uint8Array): Promise<{ name: string; said: string }[]> {
  const zips = [ZipArchive.read(one), ZipArchive.read(two)];
  const out: { name: string; said: string }[] = [];
  const names = new Set(zips[0].names().concat(zips[1].names()));
  for (const name of names) {
    const x = zips[0].names().indexOf(name) >= 0 ? Buffer.from(await zips[0].read(name)).toString('utf8') : '<missing>';
    const y = zips[1].names().indexOf(name) >= 0 ? Buffer.from(await zips[1].read(name)).toString('utf8') : '<missing>';
    if (x === y) continue;
    let i = 0;
    while (i < x.length && x[i] === y[i]) i++;
    out.push({ name, said: `${name} at ${i}: …${x.slice(Math.max(0, i - 80), i + 80)}… against …${y.slice(Math.max(0, i - 80), i + 80)}…` });
  }
  return out;
}

/** The relationships of the body of two files: the same types and targets, whatever their ids (and, with `extended`, but for one to commentsExtended.xml). */
async function sameRelations(one: Uint8Array, two: Uint8Array, extended: boolean): Promise<boolean> {
  const of = async (file: Uint8Array): Promise<string> => {
    const text = Buffer.from(await ZipArchive.read(file).read('word/_rels/document.xml.rels')).toString('utf8');
    const rels = (text.match(/<Relationship [^>]*>/g) || []).map((rel) => rel.replace(/ Id="[^"]*"/, ''));
    return rels.filter((rel) => !(extended && /commentsExtended/.test(rel))).sort().join('|');
  };
  return (await of(one)) === (await of(two));
}

/** The first file has no commentsExtended.xml, the second one that says of no comment that it answers another or is resolved. */
async function silentExtended(one: Uint8Array, two: Uint8Array): Promise<boolean> {
  const names = [ZipArchive.read(one).names(), ZipArchive.read(two).names()];
  if (names[0].indexOf('word/commentsExtended.xml') >= 0 || names[1].indexOf('word/commentsExtended.xml') < 0) return false;
  const text = Buffer.from(await ZipArchive.read(two).read('word/commentsExtended.xml')).toString('utf8');
  return !/paraIdParent|done="1"/.test(text);
}

let relsOnly = 0;
let orderOnly = 0;

const FC = (id: number, author: string, date: string, text: string): string => `<w:comment w:id="${id}" w:author="${author}" w:date="${date}"><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:comment>`;

/** A file with comments that have no paragraph ids, one about nothing, and a mark of no comment. */
function commentedFile(): Promise<Uint8Array> {
  return docFile(['Alpha beta gamma delta.', 'Second para here with words.', 'Third one.'], {
    body:
      '<w:p><w:commentRangeStart w:id="3"/><w:r><w:t>File commented text</w:t></w:r><w:commentRangeEnd w:id="3"/><w:r><w:commentReference w:id="3"/></w:r>' +
      '<w:commentRangeStart w:id="7"/><w:r><w:t> more</w:t></w:r><w:commentRangeEnd w:id="7"/><w:r><w:commentReference w:id="7"/></w:r></w:p>' +
      '<w:p><w:commentRangeStart w:id="12"/><w:r><w:t>orphan start only</w:t></w:r></w:p>',
    comments: FC(3, 'F1', '2025-01-01T00:00:00Z', 'file one') + FC(7, 'F2', '2025-01-02T00:00:00Z', 'file two') + FC(9, 'F3', '2025-01-03T00:00:00Z', 'unanchored'),
  });
}

/** Two files whose commentsExtended.xml hold the same entries, in another order. */
async function sameEntries(one: Uint8Array, two: Uint8Array): Promise<boolean> {
  const of = async (file: Uint8Array): Promise<string | undefined> => {
    const zip = ZipArchive.read(file);
    if (zip.names().indexOf('word/commentsExtended.xml') < 0) return undefined;
    const text = Buffer.from(await zip.read('word/commentsExtended.xml')).toString('utf8');
    return (text.match(/<w15:commentEx [^>]*>/g) || []).sort().join('|');
  };
  const [a, b] = [await of(one), await of(two)];
  return a !== undefined && a === b;
}

async function run(seed: number): Promise<{ comments: number; late: boolean; files: Uint8Array[] }> {
  const random = seeded(seed * 7919);
  // Every other room starts from a file with comments of its own: without paragraph ids (as old Word writes them), one with
  // no text it is about, and a mark whose comment is none.
  const bytes = seed % 2 === 0 ? await commentedFile() : await aloneFile();
  const net = new DocNet();
  const members: DocMember[] = [await joinDoc(net, 'a', bytes), await joinDoc(net, 'b', bytes), await joinDoc(net, 'c', bytes)];
  const histories = members.map((member) => new CollabDocHistory<string>(member.session));
  const suspended = [false, false, false];
  net.drain(random);
  const act = (who: number, tick: number): void => {
    const member = members[who];
    try {
      const edit = random() < 0.5 ? commentChange(member, random) : make(member.doc, seeded(seed * 1000 + tick))(member.doc);
      if (edit && edit.ops.length > 0) histories[who].apply(edit.ops, '', '');
    } catch (error) {
      if (!(error instanceof RangeError)) throw error;
    }
  };
  for (let tick = 0; tick < 70; tick++) {
    const who = int(random, 0, 2);
    const roll = random();
    if (roll < 0.35) net.step(random);
    else if (roll < 0.43) histories[who].undo();
    else if (roll < 0.46) histories[who].redo();
    else if (roll < 0.5) {
      if (suspended[who]) members[who].transport.resume();
      else members[who].transport.suspend();
      suspended[who] = !suspended[who];
    } else act(who, tick);
  }
  members.forEach((member, i) => {
    if (suspended[i]) member.transport.resume();
  });
  net.drain(random);
  net.drain(random);
  assertSameDoc(members, `seed ${seed}`);
  const texts = members.map((member) => commentsOf(member.doc));
  for (let i = 1; i < texts.length; i++) assert.equal(texts[i], texts[0], `seed ${seed}: ${members[i].name} holds other comments`);
  const files = [await saveDoc(members[0]), await saveDoc(members[1]), await saveDoc(members[2])];
  assert.ok(await samePackage(files[0], files[1]), `seed ${seed}: a and b write the same file`);
  assert.ok(await samePackage(files[0], files[2]), `seed ${seed}: a and c write the same file`);

  // One who joins with the saved file, and comments with the others.
  members[0].session.announceSaved([contentVersion(files[0])], members[0].session.rev);
  net.drain(random);
  const late = await joinDoc(net, 'late', files[0]);
  net.drain(random);
  let joined = false;
  if (late.session.state === 'online') {
    joined = true;
    assert.equal(commentsOf(late.doc), texts[0], `seed ${seed}: what the saved file holds`);
    const lateHistory = new CollabDocHistory<string>(late.session);
    for (let tick = 0; tick < 12; tick++) {
      const one = random() < 0.5;
      try {
        const member = one ? late : members[0];
        const edit = commentChange(member, random);
        if (edit && edit.ops.length > 0) (one ? lateHistory : histories[0]).apply(edit.ops, '', '');
      } catch (error) {
        if (!(error instanceof RangeError)) throw error;
      }
      if (random() < 0.5) net.drain(random);
    }
    net.drain(random);
    net.drain(random);
    assert.equal(commentsOf(late.doc), commentsOf(members[0].doc), `seed ${seed}: who joined late holds the same comments`);
    const mine = await saveDoc(members[0]);
    const theirs = await saveDoc(late);
    const differing = await differingParts(mine, theirs);
    // KNOWN LIMIT: the writer gives a part it makes (comments, the parts beside them, lists) the first free id of a
    // relationship as it writes; a part that came to be needed after the save may take, at who was there, an id that
    // the saved file gives a part made before it. The same relationships, under other ids: only that may differ.
    // And a commentsExtended.xml that the saved file has (a thread was answered or resolved then) stays in the file of
    // who joined from it, with an entry for each comment, while who was there writes none once no comment is an
    // answer or resolved any more: it says nothing more than its absence does.
    // And where the file the room started from had comments without paragraph ids: the entries of commentsExtended.xml
    // stand in the order the comments got their ids at who was there, in the order of the comments at who read the file.
    const silent = await silentExtended(mine, theirs);
    const ordered = !silent && differing.some((one) => one.name === 'word/commentsExtended.xml') && (await sameEntries(mine, theirs));
    if (ordered) orderOnly++;
    const known = differing.every((one) => one.name === 'word/_rels/document.xml.rels' || (ordered && one.name === 'word/commentsExtended.xml') || (silent && (one.name === 'word/commentsExtended.xml' || one.name === '[Content_Types].xml')));
    if (differing.length > 0 && known && (differing.every((one) => one.name !== 'word/_rels/document.xml.rels') || (await sameRelations(mine, theirs, silent)))) relsOnly++;
    else assert.deepEqual(differing.map((one) => one.said), [], `seed ${seed}: who joined from the saved file writes another file`);
    files.push(mine);
  }
  for (const member of members.concat([late])) member.session.leave();
  return { comments: listComments(members[0].doc).length, late: joined, files };
}

test('comments at random, with typing, undo and lost connections: everybody holds and writes the same, also who joins from a saved file', async (t) => {
  let comments = 0;
  let late = 0;
  let checkable = true;
  try {
    execFileSync('python3', ['-I', '-c', 'import zipfile, xml.etree.ElementTree'], { stdio: 'ignore', timeout: 60000 });
  } catch {
    checkable = false;
  }
  const dir = checkable ? mkdtempSync(joinPath(process.cwd(), 'test', '.comments-random-')) : '';
  try {
    for (let seed = 1; seed <= SEEDS; seed++) {
      const done = await run(seed);
      comments += done.comments;
      if (done.late) late++;
      if (checkable) done.files.forEach((file, i) => writeFileSync(joinPath(dir, `s${seed}-${i}.docx`), file));
    }
    if (checkable) {
      const found = JSON.parse(execFileSync('python3', ['-I', joinPath(process.cwd(), 'test', 'tools', 'package_check.py'), dir], { encoding: 'utf8', timeout: 600000, maxBuffer: 64 * 1024 * 1024 }));
      // The mark of no comment that the commented file starts with stays, as it was read (the reader says so; nothing else).
      assert.deepEqual((found as string[]).filter((said) => !/comment 12 \(commentRangeStart\)/.test(said)), [], 'an independent reader takes every file');
    } else t.diagnostic('python3 is not there: the files were not read by an independent reader');
  } finally {
    if (dir) rmSync(dir, { recursive: true, force: true });
  }
  t.diagnostic(`${SEEDS} rooms, ${comments} comments at the end, ${late} with one who joined from the saved file (${relsOnly} of them with the known differences: relationships under other ids, a silent commentsExtended.xml kept, ${orderOnly} with its entries in another order)`);
  assert.ok(comments > 0);
});
