// Text boxes of a document that is edited together: a command in a box travels with its places
// in the box and everybody runs it there, as the editor of one does (editInTextBox). Typing,
// deleting, Enter and formatting in a box reach everybody; two who type in one box both keep
// their text; what moves the box itself (a paragraph end before the drawing, text that went from
// the paragraph that anchors it) makes a waiting change of the box give way, and that is said; a
// box that may not be changed is refused by everybody alike; undo in a box is a step like any
// other; who joins from a saved file holds and writes the same.

import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { join as joinPath } from 'node:path';
import { test } from 'node:test';
import { CollabDocHistory } from '../src/collab/doc/history';
import { boxShaken, rebaseWire, type DocWire, type Mark, type WirePos } from '../src/collab/doc/rebase';
import { sharedInTextBox, wiresOf } from '../src/collab/doc/trace';
import { decodeStep, encodeStep, namesOfDocument, putDigest, putOf, runWire } from '../src/collab/doc/wire';
import { contentVersion } from '../src/collab/protocol';
import { samePackage } from '../src/formats/xlsx/collab-stamp';
import { ZipArchive, open } from '../src/index';
import { paragraphText, type Block, type DocModel } from '../src/model/doc/types';
import * as raw from '../src/ops/doc/index';
import type { DocEdit, DocPosition, DocRange } from '../src/ops/doc/index';
import { movedSelection } from '../src/ui/doc-collab';
import { ops } from './collab-doc-changes';
import { DocNet, actDoc, assertSameDoc, canonDoc, firstDifference, joinDoc, lookOfDoc, int, pick, saveDoc, seeded, type DocMember } from './collab-doc-fixtures';
import { buildDocx } from './docx-fixtures';

// ------------------------------------------------------------------- fixtures

const ANCHOR_HEAD =
  '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="1828800" cy="457200"/><wp:wrapNone/>';

const p = (text: string): string => (text === '' ? '<w:p/>' : `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`);

/** A DrawingML text box (wps) without alternate content; `locks`: the shape's own locks. */
const wpsBox = (inner: string, id: number, locks = ''): string =>
  `<w:r><w:drawing>${ANCHOR_HEAD}<wp:docPr id="${id}" name="Box ${id}"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>${locks ? `<wps:cNvSpPr txBox="1">${locks}</wps:cNvSpPr>` : ''}<wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>` +
  `<wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;

/** A text box the way Word writes it: the DrawingML shape in a Choice, the same text in a VML Fallback (or another one). */
const wordBox = (inner: string, id: number, fallback = inner): string =>
  `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>${ANCHOR_HEAD}<wp:docPr id="${id}" name="Box ${id}"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>` +
  `<wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>` +
  `<mc:Fallback><w:pict><v:shape style="width:144pt;height:36pt"><v:textbox><w:txbxContent>${fallback}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;

const SECT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';

/**
 * Paragraph 0: plain. 1: text, a box Word writes (two branches) holding two paragraphs, text. 2: plain. 3: a DrawingML box,
 * then a box whose shape locks its text. 4: a box whose other branch holds another text (it cannot take a change). 5: a box
 * with a box in it. 6: plain.
 */
function boxFile(ids = false): Promise<Uint8Array> {
  // `ids`: the two branches of the first box hold the same text under other paragraph ids and rsids, as Word may write them.
  const one = (text: string, id: string) => (ids ? `<w:p w14:paraId="${id}" w:rsidR="00${id.slice(0, 6)}"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>` : p(text));
  return buildDocx({
    body:
      p('Before the boxes.') +
      `<w:p><w:r><w:t xml:space="preserve">Anchor </w:t></w:r>${wordBox(one('Box one text', '11110001') + one('Second line', '11110002'), 1, one('Box one text', '22220001') + one('Second line', '22220002'))}<w:r><w:t xml:space="preserve"> tail</w:t></w:r></w:p>` +
      p('Middle paragraph.') +
      `<w:p><w:r><w:t xml:space="preserve">Two boxes </w:t></w:r>${wpsBox(p('Plain box'), 2)}${wpsBox(p('Locked box'), 3, '<a:spLocks noTextEdit="1"/>')}</w:p>` +
      `<w:p>${wordBox(p('Choice text'), 4, p('Other fallback'))}</w:p>` +
      `<w:p><w:r><w:t xml:space="preserve">Outer </w:t></w:r>${wpsBox(p('Outer box') + `<w:p><w:r><w:t xml:space="preserve">in </w:t></w:r>${wpsBox(p('Inner box'), 6)}</w:p>`, 5)}</w:p>` +
      p('After the boxes.') +
      SECT,
  });
}

const BOX1 = [1, -1, 0, 0];
const PLAIN = [3, -1, 0, 0];
const LOCKED = [3, -1, 1, 0];
const UNWRITABLE = [4, -1, 0, 0];
const INNER = [5, -1, 0, 0, 1, -1, 0, 0];

const wp = (path: number[], o: number, b = false): WirePos => (b ? { s: 'body', p: path, o, b: true } : { s: 'body', p: path, o });

/** The text of a list of blocks of the body (a box), a line a paragraph. */
function boxText(doc: DocModel, list: number[]): string {
  return raw
    .containerBlocks(doc, 'body', list)
    .map((block) => (block.type === 'paragraph' ? paragraphText(block) : `[${block.type}]`))
    .join('\n');
}

async function room(names: string[] = ['a', 'b'], ids = false): Promise<{ net: DocNet; members: DocMember[]; bytes: Uint8Array }> {
  const bytes = await boxFile(ids);
  const net = new DocNet();
  const members: DocMember[] = [];
  for (const name of names) members.push(await joinDoc(net, name, bytes));
  net.drain();
  return { net, members, bytes };
}

const codes = (member: DocMember): string[] => member.notices.map((notice) => notice.code);

let python: boolean | undefined;
/** Python's zipfile and ElementTree on saved files (test/tools/package_check.py): what is wrong, [] when nothing is or without python3. */
function packageProblems(files: Uint8Array[]): unknown[] {
  if (python === undefined) {
    try {
      execFileSync('python3', ['-I', '-c', 'import zipfile, xml.etree.ElementTree'], { stdio: 'ignore', timeout: 60000 });
      python = true;
    } catch {
      python = false;
    }
  }
  if (!python) return [];
  const dir = mkdtempSync(joinPath(process.cwd(), 'test', '.collab-box-'));
  try {
    files.forEach((file, i) => writeFileSync(joinPath(dir, `f${i}.docx`), file));
    return JSON.parse(execFileSync('python3', ['-I', joinPath(process.cwd(), 'test', 'tools', 'package_check.py'), dir], { encoding: 'utf8', timeout: 600000, maxBuffer: 64 * 1024 * 1024 })) as unknown[];
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

/** The text of every box of the fixture, by name. */
function boxes(doc: DocModel): Record<string, string> {
  return { one: boxText(doc, BOX1), plain: boxText(doc, PLAIN), locked: boxText(doc, LOCKED), unwritable: boxText(doc, UNWRITABLE), inner: boxText(doc, INNER) };
}

/** Where the document parts of two files first differ. */
async function firstPartDifference(one: Uint8Array, two: Uint8Array): Promise<string> {
  const [x, y] = await Promise.all([one, two].map(async (file) => Buffer.from(await ZipArchive.read(file).read('word/document.xml')).toString('utf8')));
  return x === y ? 'elsewhere than in word/document.xml' : firstDifference(x, y);
}

// ------------------------------------------------------------------------- wire

test('places in text boxes travel: a path through a box is read, a path that is no path of a box is no place', () => {
  const step = (p: unknown[]): unknown[] => [{ t: 'text', at: ['body', p, 0], x: 'x' }];
  for (const good of [BOX1.concat([0]), INNER.concat([0]), [2, 0, 1, 0, -1, 3, 1, 0, 2]]) assert.deepEqual((decodeStep(step(good))[0] as { at: WirePos }).at.p, good);
  const deep: number[] = [];
  for (let i = 0; i < 17; i++) deep.push(0, -1, 0, 0);
  deep.push(0);
  for (const bad of [
    [-1, 0, 0, 0],
    [1, -1, 0],
    [1, -1, 0, 0],
    [1, -1, -1, 0, 0],
    [1, -1, 0, -1, 0],
    [1, -1, 0, 0, -1, 0, 0, 0],
    [1, -2, 0, 0, 0],
    [1, -1, 0.5, 0, 0],
    deep,
  ])
    assert.throws(() => decodeStep(step(bad)), /Invalid path/, JSON.stringify(bad));
  // A wire written is read back as it was.
  const wire: DocWire = { t: 'del', at: wp(BOX1.concat([0]), 1), to: wp(BOX1.concat([1]), 2) };
  assert.deepEqual(decodeStep(JSON.parse(JSON.stringify(encodeStep([wire])))), [wire]);
});

test('a command whose two ends are in different boxes, or one in a box and one outside, is refused', async () => {
  const { members } = await room(['a']);
  const doc = members[0].doc;
  assert.throws(() => runWire(doc, { t: 'del', at: wp(BOX1.concat([0]), 1), to: wp(PLAIN.concat([0]), 1) }), RangeError);
  assert.throws(() => runWire(doc, { t: 'del', at: wp(BOX1.concat([0]), 1), to: wp([2], 1) }), RangeError);
  members[0].session.leave();
});

// ------------------------------------------------------------------ one at a time

test('typing, Enter, deleting and formatting in a box reach everybody, and everybody writes the same file', async () => {
  const { net, members } = await room(['a', 'b']);
  const [a, b] = members;
  assert.ok(actDoc(a, [{ t: 'text', at: wp(BOX1.concat([0]), 3), x: ' NEW' }]));
  assert.ok(actDoc(a, [{ t: 'split', at: wp(BOX1.concat([1]), 6) }]));
  assert.ok(actDoc(b, [{ t: 'del', at: wp(PLAIN.concat([0]), 0), to: wp(PLAIN.concat([0]), 6) }]));
  assert.ok(actDoc(b, [{ t: 'run', at: wp(INNER.concat([0]), 0), to: wp(INNER.concat([0]), 5), f: { bold: true } }]));
  net.drain();
  assert.deepEqual(boxes(a.doc), { one: 'Box NEW one text\nSecond\n line', plain: 'box', locked: 'Locked box', unwritable: 'Choice text', inner: 'Inner box' });
  assert.equal(raw.rangeFormat(a.doc, { anchor: { story: 'body', path: INNER.concat([0]), offset: 0 }, focus: { story: 'body', path: INNER.concat([0]), offset: 5 } }).bold, true);
  // The paragraphs that anchor the boxes keep their text.
  assert.equal(paragraphText(a.doc.body.blocks[1] as never), 'Anchor  tail');
  assertSameDoc([a, b], 'after the steps');
  assert.deepEqual(codes(a).concat(codes(b)), []);
  const files = [await saveDoc(a), await saveDoc(b)];
  assert.ok(await samePackage(files[0], files[1]), 'the same file');
  // Both branches of the box Word writes got the new text.
  const xml = Buffer.from(await members[0].file.package!.readPart('/word/document.xml')).toString('utf8');
  assert.equal(xml.split('Box NEW one text').length - 1, 2, 'the Choice and the Fallback');
  assert.deepEqual(packageProblems(files), []);
  for (const member of members) member.session.leave();
});

test('the editor runs a command in a box as it travels: sharedInTextBox notes the command with its places in the box', async () => {
  const { net, members } = await room(['a', 'b']);
  const [a, b] = members;
  const history = new CollabDocHistory<DocRange>(a.session);
  const caret: DocPosition = { story: 'body', path: BOX1.concat([0]), offset: 12 };
  const edit = sharedInTextBox(a.doc, caret, (doc, range) => ops.insertText(doc, range, '!'));
  assert.deepEqual(wiresOf(edit.ops), [{ t: 'text', at: wp(BOX1.concat([0]), 12), x: '!' }]);
  history.applyEdit(edit, { anchor: caret, focus: caret });
  const back = sharedInTextBox(a.doc, { story: 'body', path: INNER.concat([0]), offset: 5 }, (doc, range) => ops.deleteBackward(doc, range.focus, 'char'));
  assert.deepEqual(wiresOf(back.ops), [{ t: 'del', at: wp(INNER.concat([0]), 4), to: wp(INNER.concat([0]), 5) }]);
  history.applyEdit(back, { anchor: caret, focus: caret });
  net.drain();
  assert.equal(boxText(b.doc, BOX1), 'Box one text!\nSecond line');
  assert.equal(boxText(b.doc, INNER), 'Inne box');
  // Outside a box it is the command itself.
  const plain = sharedInTextBox(a.doc, { story: 'body', path: [0], offset: 0 }, (doc, range) => ops.insertText(doc, range, 'x'));
  assert.ok(wiresOf(plain.ops));
  // A command that would change more than the box is refused there (a list that is made).
  assert.throws(() => sharedInTextBox(a.doc, caret, (doc, range) => ops.toggleList(doc, range, 'bullet')), (error: unknown) => error instanceof RangeError && error.message === raw.TEXT_BOX_OUTSIDE);
  assertSameDoc([a, b], 'after the editor');
  for (const member of members) member.session.leave();
});

// ------------------------------------------------------------------- together

test('two who type in one box at once: both texts stay, in both orders of the relay', async () => {
  for (const first of [0, 1]) {
    const { net, members } = await room(['a', 'b']);
    const [a, b] = members;
    members[1 - first].transport.suspend();
    assert.ok(actDoc(a, [{ t: 'text', at: wp(BOX1.concat([0]), 3), x: 'AA' }]));
    assert.ok(actDoc(b, [{ t: 'text', at: wp(BOX1.concat([0]), 7), x: 'BB' }]));
    assert.ok(actDoc(b, [{ t: 'del', at: wp(BOX1.concat([1]), 0), to: wp(BOX1.concat([1]), 7) }]));
    net.drain();
    members[1 - first].transport.resume();
    net.drain();
    net.drain();
    assert.equal(boxText(a.doc, BOX1), 'BoxAA oneBB text\nline');
    assertSameDoc([a, b], `first ${first}`);
    assert.deepEqual(codes(a).concat(codes(b)), []);
    for (const member of members) member.session.leave();
  }
});

test('a box whose paragraph moves (a paragraph end before it, a paragraph deleted before it): the change in the box follows it', async () => {
  const { net, members } = await room(['a', 'b']);
  const [a, b] = members;
  b.transport.suspend();
  assert.ok(actDoc(a, [{ t: 'split', at: wp([0], 6) }]));
  assert.ok(actDoc(a, [{ t: 'text', at: wp([0], 0), x: 'typed ' }]));
  assert.ok(actDoc(b, [{ t: 'text', at: wp(BOX1.concat([0]), 0), x: 'B: ' }]));
  assert.ok(actDoc(b, [{ t: 'text', at: wp(INNER.concat([0]), 0), x: 'deep ' }]));
  net.drain();
  b.transport.resume();
  net.drain();
  net.drain();
  assert.equal(boxText(a.doc, [2, -1, 0, 0]), 'B: Box one text\nSecond line');
  assert.equal(boxText(a.doc, [6, -1, 0, 0, 1, -1, 0, 0]), 'deep Inner box');
  assertSameDoc([a, b], 'after the move');
  assert.deepEqual(codes(a).concat(codes(b)), []);
  // And back up: the first two paragraphs joined again.
  b.transport.suspend();
  assert.ok(actDoc(a, [{ t: 'del', at: wp([0], 12), to: wp([1], 0) }]));
  assert.ok(actDoc(b, [{ t: 'text', at: wp([2, -1, 0, 0, 1], 0), x: '> ' }]));
  net.drain();
  b.transport.resume();
  net.drain();
  net.drain();
  assert.equal(boxText(a.doc, BOX1), 'B: Box one text\n> Second line');
  assertSameDoc([a, b], 'after the join');
  for (const member of members) member.session.leave();
});

test('typing in the paragraph that anchors a box, and formatting it, moves no box: the change in the box stays', async () => {
  const { net, members } = await room(['a', 'b']);
  const [a, b] = members;
  b.transport.suspend();
  assert.ok(actDoc(a, [{ t: 'text', at: wp([1], 0), x: 'Typed ' }]));
  assert.ok(actDoc(a, [{ t: 'run', at: wp([1], 0), to: wp([1], 12), f: { italic: true } }]));
  assert.ok(actDoc(a, [{ t: 'para', at: wp([1], 0), to: wp([1], 0), f: { align: 'center' } }]));
  assert.ok(actDoc(b, [{ t: 'text', at: wp(BOX1.concat([0]), 12), x: '.' }]));
  net.drain();
  b.transport.resume();
  net.drain();
  net.drain();
  assert.equal(boxText(a.doc, BOX1), 'Box one text.\nSecond line');
  assert.equal(paragraphText(a.doc.body.blocks[1] as never), 'Typed Anchor  tail');
  assertSameDoc([a, b], 'after both');
  assert.deepEqual(codes(a).concat(codes(b)), []);
  for (const member of members) member.session.leave();
});

test('what may move the box itself - a paragraph end before its drawing, a drawing that went, its paragraph joined to the one before - makes a waiting change of the box give way, and that is said', async () => {
  // "Anchor " is 7 characters; the drawing stands at 7, " tail" follows.
  const changes: DocWire[] = [
    { t: 'split', at: wp([1], 3) },
    { t: 'del', at: wp([1], 6), to: wp([1], 8) },
    { t: 'del', at: wp([0], 3), to: wp([1], 2) },
    { t: 'text', at: wp([1], 5), to: wp([2], 3), x: 'over' },
    { t: 'text', at: wp([1], 2), to: wp([1], 4), x: 'NEW\nPARA' },
  ];
  for (const change of changes) {
    const { net, members } = await room(['a', 'b']);
    const [a, b] = members;
    b.transport.suspend();
    assert.ok(actDoc(a, [change]));
    assert.ok(actDoc(b, [{ t: 'text', at: wp(BOX1.concat([0]), 0), x: 'B ' }]));
    // What waits behind it was made on a document that had it: it goes with it (the rule of every step that moved text).
    assert.ok(actDoc(b, [{ t: 'text', at: wp([6], 0), x: 'end ' }]));
    net.drain();
    b.transport.resume();
    net.drain();
    net.drain();
    assert.deepEqual(b.notices, [{ code: 'dropped', count: 2 }], JSON.stringify(change));
    assert.ok(canonDoc(a.doc).indexOf('B Box') < 0, JSON.stringify(change));
    assertSameDoc([a, b], JSON.stringify(change));
    for (const member of members) member.session.leave();
  }
});

test('what leaves the drawing of the box where it was - text that went before it or behind it, a paragraph end behind it - moves no box: the change in the box stays', async () => {
  const changes: DocWire[] = [
    { t: 'del', at: wp([1], 0), to: wp([1], 3) },
    // Backspace right before the drawing, as it travels (the deletion of a range).
    { t: 'del', at: wp([1], 5), to: wp([1], 6) },
    { t: 'del', at: wp([1], 8), to: wp([1], 12) },
    { t: 'split', at: wp([1], 9) },
    { t: 'del', at: wp([1], 9), to: wp([2], 3) },
    { t: 'text', at: wp([1], 2), to: wp([1], 4), x: 'NEW' },
    { t: 'text', at: wp([1], 9), x: 'NEW\nPARA' },
  ];
  for (const change of changes) {
    const { net, members } = await room(['a', 'b']);
    const [a, b] = members;
    b.transport.suspend();
    assert.ok(actDoc(a, [change]), JSON.stringify(change));
    assert.ok(actDoc(b, [{ t: 'text', at: wp(BOX1.concat([0]), 0), x: 'B ' }]));
    net.drain();
    b.transport.resume();
    net.drain();
    net.drain();
    assert.deepEqual(codes(b), [], JSON.stringify(change));
    assert.ok(allBoxes(a.doc).indexOf('B Box one text\nSecond line') >= 0, JSON.stringify(change));
    assertSameDoc([a, b], JSON.stringify(change));
    for (const member of members) member.session.leave();
  }
});

/** Paragraph 1: "AA ", a box (zero), " BB ", a box (one), " CC": the drawings stand at offsets 3 and 7. */
function twoBoxFile(): Promise<Uint8Array> {
  const r = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
  return buildDocx({ body: p('Zero') + `<w:p>${r('AA ')}${wpsBox(p('box zero'), 1)}${r(' BB ')}${wpsBox(p('box one'), 2)}${r(' CC')}</w:p>` + p('End') + SECT });
}

test('a drawing that goes without a character (deleted by a key, by a range around it alone, or written over) moves the boxes behind it: what waits for them gives way', async () => {
  const before = (o: number): WirePos => ({ s: 'body', p: [1], o, b: true });
  const changes: DocWire[] = [
    { t: 'cmd', name: 'deleteBackward', at: wp([1], 3), a: ['char'] },
    { t: 'del', at: before(3), to: wp([1], 3) },
    { t: 'text', at: before(3), to: wp([1], 3), x: 'Z' },
    { t: 'cmd', name: 'deleteForward', at: before(3), a: ['char'] },
  ];
  for (const change of changes) {
    for (const box of [0, 1]) {
      const bytes = await twoBoxFile();
      const net = new DocNet();
      const a = await joinDoc(net, 'a', bytes);
      const b = await joinDoc(net, 'b', bytes);
      net.drain();
      a.transport.suspend();
      assert.ok(actDoc(a, [{ t: 'text', at: wp([1, -1, box, 0, 0], 0), x: 'X' }]));
      assert.ok(actDoc(b, [change]), JSON.stringify(change));
      net.drain();
      a.transport.resume();
      net.drain();
      net.drain();
      const label = `${JSON.stringify(change)} box ${box}`;
      assert.equal(allBoxes(b.doc).length, 1, label + ': the drawing went');
      assert.deepEqual(a.notices, [{ code: 'dropped', count: 1 }], label);
      assert.deepEqual(allBoxes(a.doc), ['box one'], label + ': nothing landed in the other box');
      assertSameDoc([a, b], label);
      for (const member of [a, b]) member.session.leave();
    }
  }
});

test('an undo in a box after its drawing went is not made, and said: it would land in the box behind it', async () => {
  for (const change of [{ t: 'cmd', name: 'deleteBackward', at: wp([1], 3), a: ['char'] } as DocWire, { t: 'del', at: { s: 'body', p: [1], o: 3, b: true }, to: wp([1], 3) } as DocWire]) {
    const bytes = await twoBoxFile();
    const net = new DocNet();
    const a = await joinDoc(net, 'a', bytes);
    const b = await joinDoc(net, 'b', bytes);
    net.drain();
    assert.ok(actDoc(a, [{ t: 'text', at: wp([1, -1, 0, 0, 0], 0), x: 'X' }]));
    a.session.seal();
    net.drain();
    assert.ok(actDoc(b, [change]));
    net.drain();
    assert.deepEqual(allBoxes(a.doc), ['box one']);
    assert.equal(a.session.undo(), undefined);
    assert.deepEqual(codes(a), ['undo-gone']);
    assert.deepEqual(allBoxes(b.doc), ['box one'], 'nothing was taken from the box behind it');
    for (const member of [a, b]) member.session.leave();
  }
});

test('a caret in a box behind a drawing that went leaves the box', () => {
  const caret: DocRange = { anchor: { story: 'body', path: [1, -1, 1, 0, 0], offset: 2 }, focus: { story: 'body', path: [1, -1, 1, 0, 0], offset: 2 } };
  // What deleting the drawing alone says: nothing moved in the text, the paragraph is another one.
  const moved = movedSelection(caret, [{ k: 'tch', s: 'body', c: [], r1: 1, r2: 1 }]);
  assert.deepEqual(moved && moved.focus, { story: 'body', path: [1], offset: 0 });
  // Formatting the paragraph that is known to keep its drawings: the caret stays.
  const kept = movedSelection(caret, [{ k: 'tch', s: 'body', c: [], r1: 1, r2: 1, gks: [2] }]);
  assert.deepEqual(kept && kept.focus, caret.focus);
});

test('a box in a box: the box around it moving makes a waiting change of the inner box give way as well', async () => {
  const { net, members } = await room(['a', 'b']);
  const [a, b] = members;
  b.transport.suspend();
  // A paragraph end in the outer box, before the paragraph that anchors the inner one: that paragraph moves down.
  assert.ok(actDoc(a, [{ t: 'split', at: wp([5, -1, 0, 0, 0], 2) }]));
  assert.ok(actDoc(b, [{ t: 'text', at: wp(INNER.concat([0]), 0), x: 'B ' }]));
  net.drain();
  b.transport.resume();
  net.drain();
  net.drain();
  assert.equal(boxText(a.doc, [5, -1, 0, 0, 2, -1, 0, 0]), 'B Inner box', 'the inner box moved with its paragraph: the change followed');
  assertSameDoc([a, b], 'after the outer split');
  b.transport.suspend();
  // A paragraph end in the paragraph of the outer box that anchors the inner box: the inner change gives way.
  assert.ok(actDoc(a, [{ t: 'split', at: wp([5, -1, 0, 0, 2], 0) }]));
  assert.ok(actDoc(b, [{ t: 'text', at: wp([5, -1, 0, 0, 2, -1, 0, 0, 0], 0), x: 'C ' }]));
  net.drain();
  b.transport.resume();
  net.drain();
  net.drain();
  assert.deepEqual(codes(b), ['dropped']);
  assertSameDoc([a, b], 'after the inner split');
  for (const member of members) member.session.leave();
});

test('the rule of a box that moved: boxShaken', () => {
  const box = [1, -1, 0, 0, 0];
  const marks: [Mark, boolean][] = [
    // Text that came into the paragraph is known to leave its drawings only where that was read off the document (gk).
    [{ k: 'ins', s: 'body', c: [], row: 1, off: 3, rows: 0, last: 2, gk: 1 }, false],
    [{ k: 'ins', s: 'body', c: [], row: 1, off: 3, rows: 0, last: 2 }, true],
    [{ k: 'ins', s: 'body', c: [], row: 1, off: 3, rows: 1, last: 0 }, true],
    [{ k: 'ins', s: 'body', c: [], row: 1, off: 3, rows: 1, last: 0, gk: 1 }, false],
    [{ k: 'ins', s: 'body', c: [], row: 1, off: 0, rows: 1, last: 0, gk: 0 }, true],
    [{ k: 'ins', s: 'body', c: [], row: 0, off: 3, rows: 2, last: 0 }, false],
    [{ k: 'del', s: 'body', c: [], r1: 1, o1: 0, r2: 1, o2: 1 }, true],
    [{ k: 'del', s: 'body', c: [], r1: 0, o1: 0, r2: 1, o2: 0 }, true],
    [{ k: 'del', s: 'body', c: [], r1: 2, o1: 0, r2: 3, o2: 0 }, false],
    // Nothing moved in the text, but a drawing may have gone (it has no width): only where its drawings are known to stay.
    [{ k: 'tch', s: 'body', c: [], r1: 1, r2: 1 }, true],
    [{ k: 'tch', s: 'body', c: [], r1: 0, r2: 2, gks: [0, 1, 0] }, false],
    [{ k: 'tch', s: 'body', c: [], r1: 0, r2: 2, gks: [1, 0, 1] }, true],
    [{ k: 'del', s: 'body', c: [], r1: 1, o1: 0, r2: 1, o2: 1, gk: 1 }, false],
    [{ k: 'tch', s: 'body', c: [], r1: 2, r2: 3 }, false],
    [{ k: 'blk', s: 'body', c: [], at: 1, n: 1, m: 1 }, true],
    [{ k: 'blk', s: 'body', c: [], at: 1, n: 0, m: 1 }, false],
    [{ k: 'opq', s: 'body' }, true],
    [{ k: 'opq', s: 'header:rId1' }, false],
    // In the box itself, and in another box of the same paragraph: the box is where it was.
    [{ k: 'del', s: 'body', c: [1, -1, 0, 0], r1: 0, o1: 0, r2: 1, o2: 0 }, false],
    [{ k: 'ins', s: 'body', c: [1, -1, 1, 0], row: 0, off: 0, rows: 3, last: 0 }, false],
  ];
  for (const [mark, shaken] of marks) assert.equal(boxShaken('body', box, mark), shaken, JSON.stringify(mark));
  assert.equal(boxShaken('body', [2], { k: 'opq', s: 'body' }), false, 'no box');
  // A box in a table cell: the table that went.
  assert.equal(boxShaken('body', [2, 0, 0, 1, -1, 0, 0, 0], { k: 'del', s: 'body', c: [], r1: 1, o1: 0, r2: 3, o2: 0 }), true);
  assert.equal(boxShaken('body', [2, 0, 0, 1, -1, 0, 0, 0], { k: 'ins', s: 'body', c: [], row: 1, off: 0, rows: 4, last: 0 }), false);
  assert.equal(boxShaken('body', [2, 0, 0, 1, -1, 0, 0, 0], { k: 'ins', s: 'body', c: [2, 0, 0], row: 1, off: 0, rows: 1, last: 0 }), true);
  // Rebasing: a waiting change of the box through what moves it is dropped; through typing beside it, it stays.
  assert.equal(rebaseWire({ t: 'text', at: wp(box, 0), x: 'x' }, [], [{ k: 'ins', s: 'body', c: [], row: 1, off: 0, rows: 1, last: 0 }]).drop, true);
  assert.deepEqual(rebaseWire({ t: 'text', at: wp(box, 0), x: 'x' }, [], [{ k: 'ins', s: 'body', c: [], row: 0, off: 0, rows: 1, last: 0 }]).wires, [{ t: 'text', at: wp([2, -1, 0, 0, 0], 0), x: 'x' }]);
});

test('a box in a table cell and a box in a header: typed in at once by two, held and written alike, by who joins from the saved file too', async () => {
  const cell = (inner: string) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;
  const bytes = await buildDocx({
    body:
      p('Above the table.') +
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
      `<w:tr>${cell(p('plain cell'))}${cell(`<w:p><w:r><w:t xml:space="preserve">cell </w:t></w:r>${wordBox(p('Cell box'), 8)}</w:p>`)}</w:tr></w:tbl>` +
      p('Below.') +
      '<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>',
    headers: { rIdH1: `<w:p><w:r><w:t xml:space="preserve">Head </w:t></w:r>${wpsBox(p('Header box'), 9)}</w:p>` },
  });
  const net = new DocNet();
  const a = await joinDoc(net, 'a', bytes);
  const b = await joinDoc(net, 'b', bytes);
  net.drain();
  const cellBox = [1, 0, 1, 0, -1, 0, 0];
  const header = 'header:rIdH1';
  const inHeader = (path: number[], o: number): WirePos => ({ s: header, p: path, o });
  b.transport.suspend();
  assert.ok(actDoc(a, [{ t: 'text', at: wp(cellBox.concat([0]), 4), x: ' A' }]));
  assert.ok(actDoc(a, [{ t: 'split', at: wp([0], 5) }]));
  assert.ok(actDoc(a, [{ t: 'text', at: inHeader([0, -1, 0, 0, 0], 0), x: 'A ' }]));
  assert.ok(actDoc(b, [{ t: 'text', at: wp(cellBox.concat([0]), 0), x: 'B ' }]));
  assert.ok(actDoc(b, [{ t: 'split', at: inHeader([0, -1, 0, 0, 0], 6) }]));
  net.drain();
  b.transport.resume();
  net.drain();
  net.drain();
  assert.deepEqual(codes(a).concat(codes(b)), []);
  assert.equal(boxText(a.doc, [2, 0, 1, 0, -1, 0, 0]), 'B Cell A box');
  assert.equal(raw.containerBlocks(a.doc, header, [0, -1, 0, 0]).map((block) => paragraphText(block as never)).join('\n'), 'A Header\n box');
  assertSameDoc([a, b], 'after both');
  const file = await saveDoc(a);
  assert.ok(await samePackage(file, await saveDoc(b)));
  a.session.announceSaved([contentVersion(file)], a.session.rev);
  net.drain();
  const late = await joinDoc(net, 'late', file);
  net.drain();
  assert.equal(late.session.state, 'online');
  assert.ok(actDoc(late, [{ t: 'text', at: inHeader([0, -1, 0, 0, 1], 0), x: 'L' }]));
  assert.ok(actDoc(b, [{ t: 'del', at: wp([2, 0, 1, 0, -1, 0, 0, 0], 0), to: wp([2, 0, 1, 0, -1, 0, 0, 0], 2) }]));
  net.drain();
  net.drain();
  assert.equal(lookOfDoc(late.doc), lookOfDoc(a.doc));
  assert.deepEqual(allBoxes(late.doc), allBoxes(a.doc));
  const files = [await saveDoc(a), await saveDoc(late)];
  assert.ok(await samePackage(files[0], files[1]), 'who joined late writes the same');
  assert.deepEqual(packageProblems(files), []);
  for (const member of [a, b, late]) member.session.leave();
});

test('a table in a box: a row and a column come and go, cells are typed in, a table is made there; everybody holds and writes the same', async () => {
  const table =
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="1200"/></w:tblGrid>' +
    `<w:tr><w:tc>${p('c11')}</w:tc><w:tc>${p('c12')}</w:tc></w:tr><w:tr><w:tc>${p('c21')}</w:tc><w:tc>${p('c22')}</w:tc></w:tr></w:tbl>`;
  const bytes = await buildDocx({ body: p('Before.') + `<w:p><w:r><w:t xml:space="preserve">Anchor </w:t></w:r>${wordBox(p('Above the table') + table + p('Below the table'), 1)}</w:p>` + SECT });
  const net = new DocNet();
  const a = await joinDoc(net, 'a', bytes);
  const b = await joinDoc(net, 'b', bytes);
  net.drain();
  const box = [1, -1, 0, 0];
  assert.ok(actDoc(a, [{ t: 'cmd', name: 'insertRow', at: wp(box.concat([1, 0, 0, 0]), 0), a: ['below'] }]));
  net.drain();
  assert.ok(actDoc(b, [{ t: 'text', at: wp(box.concat([1, 1, 0, 0]), 0), x: 'new row' }]));
  net.drain();
  // A column of the table and a table after it, at once: they are about other blocks of the box.
  assert.ok(actDoc(a, [{ t: 'cmd', name: 'insertColumn', at: wp(box.concat([1, 2, 1, 0]), 0), a: ['right'] }]));
  assert.ok(actDoc(b, [{ t: 'cmd', name: 'insertTable', at: wp(box.concat([2]), 5), a: [2, 2] }]));
  net.drain();
  net.drain();
  assertSameDoc([a, b], 'after the tables');
  const inner = raw.containerBlocks(a.doc, 'body', box);
  assert.equal(inner.filter((block) => block.type === 'table').length, 2, 'the table and the new one ');
  assert.equal((inner[1] as { rows: { cells: unknown[] }[] }).rows.length, 3);
  assert.equal((inner[1] as { rows: { cells: unknown[] }[] }).rows[0].cells.length, 3);
  assert.deepEqual(codes(a).concat(codes(b)), []);
  assert.equal(paragraphText(raw.containerBlocks(a.doc, 'body', box.concat([1, 1, 0]))[0] as never), 'new row');
  // The undo of a row is none (the structure of a table is not taken back while the document is edited together), and said.
  assert.equal(a.session.undo(), undefined);
  assert.deepEqual(codes(a), ['undo-none']);
  const files = [await saveDoc(a), await saveDoc(b)];
  assert.ok(await samePackage(files[0], files[1]));
  assert.deepEqual(packageProblems(files), []);
  a.session.leave();
  b.session.leave();
});

test('who joins from a saved file holds the model of the boxes the others hold; only the tree a drawing keeps as its file had it differs', async () => {
  const { net, members } = await room(['a'], true);
  const [a] = members;
  assert.ok(actDoc(a, [{ t: 'text', at: wp(BOX1.concat([0]), 3), x: 'XY' }]));
  a.session.seal();
  net.drain();
  const file = await saveDoc(a);
  a.session.announceSaved([contentVersion(file)], a.session.rev);
  net.drain();
  const late = await joinDoc(net, 'late', file);
  net.drain();
  // The drawing keeps the tree of its file (with the box as that file had it) beside the blocks of its box: who was there
  // keeps that of the file the room started from, who joined that of the saved file. That tree is what the writer starts
  // from, so it cannot be made the same without writing the file the room started from again; it is all that differs.
  assert.notEqual(canonDoc(late.doc), canonDoc(a.doc));
  assert.equal(canonDoc(withoutKeptTrees(late.doc)), canonDoc(withoutKeptTrees(a.doc)));
  assert.deepEqual(namesOf(late.doc), namesOf(a.doc));
  // Taken back, at both: the same content and the same names. The XML of the box may be written otherwise (who was there
  // writes it with the bytes of the file the room started from), and the files are read back the same.
  assert.ok(a.session.undo());
  net.drain();
  assert.equal(canonDoc(withoutKeptTrees(late.doc)), canonDoc(withoutKeptTrees(a.doc)));
  assert.deepEqual(namesOf(late.doc), namesOf(a.doc));
  const [x, y] = await Promise.all([saveDoc(a), saveDoc(late)].map(async (made) => (await open(await made)).document!()));
  assert.equal(lookOfDoc(x), lookOfDoc(y));
  assert.deepEqual(allBoxes(x), allBoxes(y));
  assert.deepEqual(namesOf(x), namesOf(y));
  for (const member of [a, late]) member.session.leave();
});

test('a key that replaces the paragraph that anchors boxes but leaves its drawings as they were moves no box', async () => {
  const changes: DocWire[] = [
    // The space behind the first drawing, by a key: known only as "the paragraph was replaced".
    { t: 'cmd', name: 'deleteBackward', at: wp([1], 4), a: ['char'] },
    { t: 'cmd', name: 'deleteForward', at: wp([1], 4), a: ['char'] },
  ];
  for (const change of changes) {
    for (const box of [0, 1]) {
      const bytes = await twoBoxFile();
      const net = new DocNet();
      const a = await joinDoc(net, 'a', bytes);
      const b = await joinDoc(net, 'b', bytes);
      net.drain();
      a.transport.suspend();
      assert.ok(actDoc(a, [{ t: 'text', at: wp([1, -1, box, 0, 0], 0), x: 'X' }]));
      assert.ok(actDoc(b, [change]));
      net.drain();
      a.transport.resume();
      net.drain();
      net.drain();
      const label = `${JSON.stringify(change)} box ${box}`;
      assert.deepEqual(codes(a), [], label);
      assert.deepEqual(allBoxes(a.doc), box === 0 ? ['Xbox zero', 'box one'] : ['box zero', 'Xbox one'], label);
      assertSameDoc([a, b], label);
      for (const member of [a, b]) member.session.leave();
    }
  }
});

test('a drawing that went moved the boxes behind it for undo too: alone in a room, the undo of a box behind it is not made', async () => {
  const r = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
  const bytes = await buildDocx({ body: p('Zero') + `<w:p>${r('AA ')}${wpsBox(p('box zero'), 1)}${r(' BB ')}${wpsBox(p('box one'), 2)}${r(' CC ')}${wpsBox(p('box two'), 3)}${r(' DD')}</w:p>` + p('End') + SECT });
  for (const alone of [true, false]) {
    const net = new DocNet();
    const a = await joinDoc(net, 'a', bytes);
    const b = alone ? undefined : await joinDoc(net, 'b', bytes);
    net.drain();
    assert.ok(actDoc(a, [{ t: 'text', at: wp([1, -1, 1, 0, 0], 0), x: 'X' }]));
    a.session.seal();
    net.drain();
    // The first drawing goes (by a key): box one is the first box behind it now, box two the second.
    assert.ok(actDoc(a, [{ t: 'cmd', name: 'deleteBackward', at: wp([1], 3), a: ['char'] }]));
    a.session.seal();
    net.drain();
    assert.equal(a.session.undo(), undefined, 'the deletion of a drawing is not undone');
    assert.equal(a.session.undo(), undefined, 'and what lies under it was about the boxes as they stood');
    assert.deepEqual(allBoxes(a.doc), ['Xbox one', 'box two'], `alone ${alone}: nothing was taken from another box`);
    if (b) assertSameDoc([a, b], 'after the undo');
    for (const member of b ? [a, b] : [a]) member.session.leave();
  }
});

test('a deletion of a drawing that waits and gives way takes what waits behind it with it: the box it was typed in is not the one there', async () => {
  const bytes = await twoBoxFile();
  for (const deletion of [{ t: 'cmd', name: 'deleteBackward', at: wp([1], 3), a: ['char'] } as DocWire, { t: 'del', at: { s: 'body', p: [1], o: 3, b: true }, to: wp([1], 3) } as DocWire]) {
    const net = new DocNet();
    const a = await joinDoc(net, 'a', bytes);
    const b = await joinDoc(net, 'b', bytes);
    net.drain();
    a.transport.suspend();
    // a deletes the first drawing and types in the box that is the first one now (box one); b's key deletes in that paragraph.
    assert.ok(actDoc(a, [deletion]));
    assert.ok(actDoc(a, [{ t: 'text', at: wp([1, -1, 0, 0, 0], 0), x: 'X' }]));
    assert.ok(actDoc(b, [{ t: 'cmd', name: 'deleteForward', at: wp([1], 4), a: ['char'] }]));
    net.drain();
    a.transport.resume();
    net.drain();
    net.drain();
    assertSameDoc([a, b], JSON.stringify(deletion));
    assert.deepEqual(a.notices, [{ code: 'dropped', count: 2 }], JSON.stringify(deletion));
    assert.deepEqual(allBoxes(a.doc), ['box zero', 'box one'], 'nothing typed in the box the deletion would have taken away');
    for (const member of [a, b]) member.session.leave();
  }
});

test('typing that arrives in the paragraph of a drawing that went here meanwhile takes no drawing: the box typed in behind it stays where it is', async () => {
  const bytes = await twoBoxFile();
  const net = new DocNet();
  const a = await joinDoc(net, 'a', bytes);
  const b = await joinDoc(net, 'b', bytes);
  net.drain();
  a.transport.suspend();
  assert.ok(actDoc(a, [{ t: 'cmd', name: 'deleteBackward', at: wp([1], 3), a: ['char'] }]));
  assert.ok(actDoc(a, [{ t: 'text', at: wp([1, -1, 0, 0, 0], 0), x: 'X' }]));
  assert.ok(actDoc(b, [{ t: 'text', at: wp([1], 0), x: 'q' }]));
  net.drain();
  a.transport.resume();
  net.drain();
  net.drain();
  assert.deepEqual(codes(a), []);
  assert.deepEqual(allBoxes(a.doc), ['Xbox one']);
  assertSameDoc([a, b], 'after both');
  for (const member of [a, b]) member.session.leave();
});

test('ids only the other branch of a box holds: everybody counts the names alike before an edit, after it, after a save and after the undo', async () => {
  const bytes = await buildDocx({ body: p('Before') + `<w:p><w:r><w:t xml:space="preserve">Anchor </w:t></w:r>${wordBox(p('Box text') + p('Two'), 1, `<w:p w14:paraId="22222221"><w:r w:rsidR="00BB0002"><w:t xml:space="preserve">Box text</w:t></w:r></w:p>` + `<w:p w14:paraId="22222222"><w:r><w:t xml:space="preserve">Two</w:t></w:r></w:p>`)}</w:p>` + p('After') + SECT });
  const net = new DocNet();
  const a = await joinDoc(net, 'a', bytes);
  const b = await joinDoc(net, 'b', bytes);
  net.drain();
  const alike = (members: DocMember[], label: string): void => {
    for (const member of members) assert.deepEqual(namesOf(member.doc), namesOf(members[0].doc), `${label}: ${member.name}`);
    for (const id of ['22222221', '22222222', '33333333']) {
      const answers = members.map((member) => accepts(member.doc, idPut(member.doc, id)));
      assert.deepEqual(answers, answers.map(() => answers[0]), `${label}: everybody decides alike on ${id}`);
    }
  };
  assert.equal(accepts(a.doc, idPut(a.doc, '22222221')), 'refused', 'the file names it');
  alike([a, b], 'as read');
  assert.ok(actDoc(a, [{ t: 'text', at: wp([1, -1, 0, 0, 0], 3), x: 'XY' }]));
  assert.ok(actDoc(a, [{ t: 'del', at: wp([1, -1, 0, 0, 0], 8), to: wp([1, -1, 0, 0, 1], 3) }]));
  a.session.seal();
  net.drain();
  alike([a, b], 'edited');
  const file = await saveDoc(a);
  a.session.announceSaved([contentVersion(file)], a.session.rev);
  net.drain();
  const late = await joinDoc(net, 'late', file);
  net.drain();
  alike([a, b, late], 'saved');
  assert.ok(a.session.undo());
  net.drain();
  assert.ok(a.session.undo());
  net.drain();
  alike([a, b, late], 'taken back');
  assert.deepEqual(allBoxes(late.doc), ['Box text\nTwo']);
  const files = [await saveDoc(a), await saveDoc(late)];
  assert.deepEqual(packageProblems(files), []);
  for (const member of [a, b, late]) member.session.leave();
});

test('a box whose other branch does not stand element for element with the branch read is not changed: ids have no place to go', async () => {
  const cases: [string, string][] = [
    // The same text in two runs, and in one.
    ['<w:p><w:r><w:t xml:space="preserve">Box </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>text</w:t></w:r></w:p>', '<w:p><w:r><w:t>Box text</w:t></w:r></w:p>'],
    // One paragraph more.
    [p('Box text'), p('Box text') + p('')],
  ];
  for (const [choice, fallback] of cases) {
    const bytes = await buildDocx({ body: `<w:p><w:r><w:t xml:space="preserve">Anchor </w:t></w:r>${wordBox(choice, 1, fallback)}</w:p>` + SECT });
    const doc = await (await open(bytes)).document!();
    assert.equal(raw.textBoxProblem(doc, { story: 'body', path: [0, -1, 0, 0, 0], offset: 0 }), 'unwritable', fallback);
  }
});

test('at random, alone and in a pair: typing in a box behind drawings, a drawing that goes, undo twice - no other box changes', async () => {
  const r = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
  const bytes = await buildDocx({ body: p('Zero') + `<w:p>${r('AA ')}${wpsBox(p('box zero'), 1)}${r(' BB ')}${wpsBox(p('box one'), 2)}${r(' CC ')}${wpsBox(p('box two'), 3)}${r(' DD')}</w:p>` + p('End') + SECT });
  const at = [3, 8, 13];
  for (let seed = 1; seed <= Math.min(SEEDS, 60); seed++) {
    const random = seeded(seed * 7907);
    const alone = seed % 2 === 1;
    const net = new DocNet();
    const a = await joinDoc(net, 'a', bytes);
    const b = alone ? undefined : await joinDoc(net, 'b', bytes);
    net.drain();
    const typed = int(random, 0, 2);
    const gone = int(random, 0, 2);
    const offline = !alone && random() < 0.5;
    if (offline) a.transport.suspend();
    assert.ok(actDoc(a, [{ t: 'text', at: wp([1, -1, typed, 0, 0], 0), x: 'X' }]));
    a.session.seal();
    if (!offline) net.drain();
    const deletion: DocWire = random() < 0.5 ? { t: 'cmd', name: 'deleteBackward', at: wp([1], at[gone]), a: ['char'] } : { t: 'del', at: { s: 'body', p: [1], o: at[gone], b: true }, to: wp([1], at[gone]) };
    assert.ok(actDoc(a, [deletion]));
    a.session.seal();
    // Someone else meanwhile: a key beside the drawings, or typing in the paragraph.
    if (b) assert.ok(actDoc(b, [random() < 0.5 ? { t: 'cmd', name: 'deleteForward', at: wp([1], 4), a: ['char'] } : { t: 'text', at: wp([1], 0), x: 'q' }]));
    if (offline) {
      net.drain();
      a.transport.resume();
    }
    net.drain();
    net.drain();
    a.session.undo();
    net.drain();
    a.session.undo();
    net.drain();
    // The boxes that stand: each holds its own text, or its own text with the X where it was typed.
    const names = ['box zero', 'box one', 'box two'];
    for (const text of allBoxes(a.doc)) assert.ok(names.indexOf(text) >= 0 || text === 'X' + names[typed], `seed ${seed}: "${text}" in ${JSON.stringify(allBoxes(a.doc))}`);
    if (b) assertSameDoc([a, b], `seed ${seed}`);
    for (const member of b ? [a, b] : [a]) member.session.leave();
  }
});

// ------------------------------------------------------------------- refused

test('a box that may not be changed is refused by everybody alike: its shape locks its text, its other branch cannot take the change, it is not there', async () => {
  const { net, members } = await room(['a', 'b']);
  const [a, b] = members;
  // Its author is refused at once.
  assert.equal(actDoc(a, [{ t: 'text', at: wp(LOCKED.concat([0]), 0), x: 'x' }]), false);
  assert.equal(actDoc(a, [{ t: 'text', at: wp(UNWRITABLE.concat([0]), 0), x: 'x' }]), false);
  // One who is not honest sends them anyway: everybody passes over them, and says so once.
  const bad: unknown[][] = [
    encodeStep([{ t: 'text', at: wp(LOCKED.concat([0]), 0), x: 'x' }]),
    encodeStep([{ t: 'text', at: wp(UNWRITABLE.concat([0]), 0), x: 'x' }]),
    encodeStep([{ t: 'text', at: wp([1, -1, 0, 5, 0], 0), x: 'x' }]),
    encodeStep([{ t: 'text', at: wp([1, -1, 9, 0, 0], 0), x: 'x' }]),
    encodeStep([{ t: 'text', at: wp([2, -1, 0, 0, 0], 0), x: 'x' }]),
    encodeStep([{ t: 'text', at: wp(BOX1.concat([7]), 0), x: 'x' }]),
    encodeStep([{ t: 'text', at: wp(BOX1.concat([0]), 99), x: 'x' }]),
    // What the editor does not put into a box, a step does not either: a page break, a link.
    encodeStep([{ t: 'cmd', name: 'insertPageBreak', at: wp(BOX1.concat([0]), 2), a: [] }]),
    encodeStep([{ t: 'cmd', name: 'insertHyperlink', at: wp(BOX1.concat([0]), 0), to: wp(BOX1.concat([0]), 3), a: ['https://example.com/', 'rIdQ1'] }]),
    // A paragraph of a box that would grow beyond what a paragraph may hold, in steps a relay takes.
    encodeStep([{ t: 'del', at: wp(BOX1.concat([0]), 0), to: wp(PLAIN.concat([0]), 0) }]),
  ];
  for (const step of bad) {
    const before = boxes(b.doc);
    (b.session as unknown as { remote(raw: unknown[]): boolean }).remote(step);
    assert.deepEqual(boxes(b.doc), before, JSON.stringify(step));
  }
  assert.equal(codes(b).filter((code) => code === 'bad-step').length, 1, 'said once');
  net.drain();
  assertSameDoc([a, b], 'nothing changed');
  for (const member of members) member.session.leave();
});

test('a text too long for a box is refused by everybody alike: the text of the boxes counts with the text of the story', async () => {
  const { members } = await room(['a']);
  const doc = members[0].doc;
  // A paragraph of a box takes no more than a paragraph of the body.
  const first = runWire(doc, { t: 'text', at: wp(BOX1.concat([0]), 0), x: 'x'.repeat(700000) });
  for (const op of first.ops) raw.applyDocOp(doc, op);
  assert.throws(() => runWire(doc, { t: 'text', at: wp(BOX1.concat([0]), 0), x: 'x'.repeat(400000) }), /A paragraph would be too long/);
  const long = 'y'.repeat(700000);
  // Many steps of a large text, each one a relay takes: the box grows to its limit, and no further.
  let refused = false;
  for (let i = 0; i < 40 && !refused; i++) {
    try {
      const edit = runWire(doc, { t: 'text', at: wp(BOX1.concat([1]), 0), x: long + '\n' });
      for (const op of edit.ops) raw.applyDocOp(doc, op);
    } catch (error) {
      assert.ok(error instanceof RangeError);
      assert.match((error as Error).message, /too long/);
      refused = true;
    }
  }
  assert.ok(refused, 'the story with its boxes has a limit');
  members[0].session.leave();
});

test('what paragraphs put back into a box name is counted in the whole document: a bookmark or a paragraph id the body holds is not brought twice', async () => {
  const bytes = await buildDocx({
    body:
      '<w:p w14:paraId="1111AAAA"><w:bookmarkStart w:id="5" w:name="Mark"/><w:r><w:t>Body with a bookmark</w:t></w:r><w:bookmarkEnd w:id="5"/></w:p>' +
      `<w:p><w:r><w:t xml:space="preserve">Anchor </w:t></w:r>${wpsBox(p('In the box'), 1)}</w:p>` +
      SECT,
  });
  const net = new DocNet();
  const a = await joinDoc(net, 'a', bytes);
  const b = await joinDoc(net, 'b', bytes);
  net.drain();
  const box = [1, -1, 0, 0];
  const h = putDigest(raw.containerBlocks(a.doc, 'body', box).slice(0, 1));
  const put = (r: unknown[], attrs?: [string, string][]): DocWire => ({ t: 'put', at: wp(box.concat([0]), 0), n: 1, x: [attrs ? { a: attrs, r: r as never } : { r: r as never }], h });
  const refused = [put([{ b: [0, '9', 'mark'] }, { c: ['x'] }, { b: [1, '9'] }]), put([{ b: [0, '5', 'Other'] }, { c: ['x'] }, { b: [1, '5'] }]), put([{ c: ['x'] }], [['w14:paraId', '1111AAAA']])];
  for (const wire of refused) assert.throws(() => runWire(a.doc, wire), /The document has a (bookmark|paragraph) of this/, JSON.stringify(wire));
  // One that names nothing the document holds is taken, by everybody.
  (b.session as unknown as { remote(raw: unknown[]): boolean }).remote(encodeStep([put([{ b: [0, '12', 'Fresh'] }, { c: ['put back'] }, { b: [1, '12'] }], [['w14:paraId', '2222BBBB']])]));
  assert.equal(boxText(b.doc, box), 'put back');
  assert.deepEqual(codes(b), []);
  // And then it is a name of the document: a put into the box cannot bring it again.
  const again = putDigest(raw.containerBlocks(b.doc, 'body', box).slice(0, 1));
  assert.throws(() => runWire(b.doc, { t: 'put', at: wp(box.concat([0]), 0), n: 1, x: [{ r: [{ b: [0, '12', 'Fresh'] }, { c: ['y'] }, { b: [1, '12'] }, { b: [0, '13', 'fresh'] }, { b: [1, '13'] }] as never }], h: again }), /The document has a bookmark of this name/);
  // Nor in the body: the name stands in the box now.
  const body = putDigest(b.doc.body.blocks.slice(0, 1));
  assert.throws(() => runWire(b.doc, { t: 'put', at: wp([0], 0), n: 1, x: [{ r: [{ b: [0, '14', 'FRESH'] }, { c: ['z'] }, { b: [1, '14'] }] as never }], h: body }), /The document has a bookmark of this name/);
  a.session.leave();
  b.session.leave();
});

/** A box Word writes whose two branches hold the same text under other ids (the paragraph ids, the rsids). */
function idsFile(): Promise<Uint8Array> {
  const para = (text: string, attrs: string) => `<w:p ${attrs}><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  return buildDocx({
    body:
      p('Before') +
      `<w:p><w:r><w:t xml:space="preserve">Anchor </w:t></w:r>${wordBox(para('Box text', 'w14:paraId="11111111" w:rsidR="00AA0001"') + para('Second', 'w14:paraId="11111112"'), 1, para('Box text', 'w14:paraId="22222221" w:rsidR="00BB0002"') + para('Second', 'w14:paraId="22222222"'))}</w:p>` +
      p('After') +
      SECT,
  });
}

/** A put of the first paragraph of the body that gives it a paragraph id. */
function idPut(doc: DocModel, id: string): DocWire {
  const block = doc.body.blocks[0];
  const x = putOf(block as never)!;
  x.a = [['w14:paraId', id]];
  return { t: 'put', at: wp([0], 0), n: 1, x: [x], h: putDigest([block]) };
}

const accepts = (doc: DocModel, wire: DocWire): string => {
  try {
    runWire(doc, wire);
    return 'accepted';
  } catch (error) {
    return 'refused';
  }
};

test('an edited box Word writes gets the ids of the branch read in its other branch too, as Word writes them: the file names each id once, at everybody', async () => {
  const net = new DocNet();
  const bytes = await idsFile();
  const a = await joinDoc(net, 'a', bytes);
  const b = await joinDoc(net, 'b', bytes);
  net.drain();
  assert.ok(actDoc(a, [{ t: 'text', at: wp([1, -1, 0, 0, 0], 3), x: 'XY' }]));
  assert.ok(actDoc(a, [{ t: 'split', at: wp([1, -1, 0, 0, 1], 3) }]));
  a.session.seal();
  net.drain();
  const file = await saveDoc(a);
  const xml = Buffer.from(await ZipArchive.read(file).read('word/document.xml')).toString('utf8');
  const fallback = xml.slice(xml.indexOf('<mc:Fallback>'));
  // The other branch is made of the model alone: the ids of the branch read (both branches name them; counted once).
  for (const id of ['11111111', '11111112', '00AA0001']) assert.ok(fallback.indexOf(`"${id}"`) >= 0, `the Fallback has ${id}`);
  for (const id of ['22222221', '22222222', '00BB0002']) assert.ok(fallback.indexOf(`"${id}"`) < 0, `the Fallback no longer has ${id}`);
  a.session.announceSaved([contentVersion(file)], a.session.rev);
  net.drain();
  const late = await joinDoc(net, 'late', file);
  net.drain();
  assert.equal(late.session.state, 'online');
  const alike = (label: string): void => {
    for (const member of [a, b, late]) assert.deepEqual(namesOf(member.doc), namesOf(a.doc), `${label}: ${member.name}`);
    for (const id of ['22222221', '22222222', '11111111', '33333333']) {
      const answers = [a, b, late].map((member) => accepts(member.doc, idPut(member.doc, id)));
      assert.deepEqual(answers, [answers[0], answers[0], answers[0]], `${label}: everybody decides alike on ${id}`);
    }
  };
  alike('saved');
  // The edit taken back (as a command): at everybody the box is written from the model, in both branches.
  assert.ok(a.session.undo());
  net.drain();
  alike('taken back');
  // An id the file no longer names may come into the body now (one who speaks the protocol itself sends it), at everybody;
  // it is not named twice afterwards, at anybody.
  assert.deepEqual([a, b, late].map((member) => accepts(member.doc, idPut(member.doc, '22222221'))), ['accepted', 'accepted', 'accepted']);
  late.transport.send(JSON.stringify({ t: 'op', base: late.session.rev, id: 'forged-1', ops: encodeStep([idPut(late.doc, '22222221')]) }));
  net.drain();
  const files = [await saveDoc(a), await saveDoc(b), await saveDoc(late)];
  for (const one of files) {
    const text = Buffer.from(await ZipArchive.read(one).read('word/document.xml')).toString('utf8');
    const body = text.slice(0, text.indexOf('<mc:AlternateContent'));
    assert.equal(body.split('w14:paraId="22222221"').length - 1, 1, 'the id brought into the body');
    assert.equal(text.slice(text.indexOf('<mc:AlternateContent')).indexOf('22222221'), -1, 'and in no branch of the box');
  }
  assert.deepEqual(packageProblems(files), []);
  for (const member of [a, b, late]) member.session.leave();
});

// ------------------------------------------------------------------------ undo

test('undo in a box is a step like any other: typing, deleting and formatting there are taken back, also at the others', async () => {
  const { net, members } = await room(['a', 'b']);
  const [a, b] = members;
  const initial = boxes(b.doc);
  assert.ok(actDoc(a, [{ t: 'text', at: wp(BOX1.concat([0]), 3), x: ' typed' }], { key: 't' }));
  a.session.seal();
  assert.ok(actDoc(a, [{ t: 'del', at: wp(BOX1.concat([0]), 0), to: wp(BOX1.concat([1]), 3) }]));
  a.session.seal();
  assert.ok(actDoc(a, [{ t: 'run', at: wp(INNER.concat([0]), 0), to: wp(INNER.concat([0]), 3), f: { underline: 'single' } }]));
  a.session.seal();
  assert.ok(actDoc(a, [{ t: 'split', at: wp(PLAIN.concat([0]), 2) }]));
  net.drain();
  assert.equal(boxText(b.doc, BOX1), 'ond line');
  for (let i = 0; i < 4; i++) {
    assert.ok(a.session.undo(), `undo ${i}`);
    net.drain();
  }
  assert.deepEqual(boxes(b.doc), initial);
  assert.deepEqual(codes(a), []);
  assertSameDoc([a, b], 'after undo');
  // And made again.
  for (let i = 0; i < 4; i++) {
    assert.ok(a.session.redo(), `redo ${i}`);
    net.drain();
  }
  assert.equal(boxText(b.doc, BOX1), 'ond line');
  assertSameDoc([a, b], 'after redo');
  for (const member of members) member.session.leave();
});

test('an undo in a box after someone else changed the same paragraph of the box is not made, and said; one in another box is made', async () => {
  const { net, members } = await room(['a', 'b']);
  const [a, b] = members;
  assert.ok(actDoc(a, [{ t: 'del', at: wp(BOX1.concat([0]), 0), to: wp(BOX1.concat([0]), 4) }]));
  a.session.seal();
  assert.ok(actDoc(a, [{ t: 'del', at: wp(PLAIN.concat([0]), 0), to: wp(PLAIN.concat([0]), 6) }]));
  a.session.seal();
  net.drain();
  assert.ok(actDoc(b, [{ t: 'text', at: wp(BOX1.concat([0]), 0), x: 'b ' }]));
  net.drain();
  // The last one, in the other box, is taken back; the one before it stays (b typed in its paragraph).
  assert.ok(a.session.undo());
  net.drain();
  assert.equal(boxText(b.doc, PLAIN), 'Plain box');
  assert.equal(a.session.undo(), undefined);
  assert.deepEqual(codes(a), ['undo-gone']);
  assert.equal(boxText(b.doc, BOX1), 'b one text\nSecond line');
  assertSameDoc([a, b], 'after undo');
  for (const member of members) member.session.leave();
});

test('an undo in a box whose paragraph someone else moved meanwhile: the box moves with it, and the undo with the box; one whose box may have moved is not made', async () => {
  const { net, members } = await room(['a', 'b']);
  const [a, b] = members;
  assert.ok(actDoc(a, [{ t: 'text', at: wp(BOX1.concat([0]), 0), x: 'one ' }]));
  a.session.seal();
  net.drain();
  // A paragraph end before the paragraph that anchors the box: the box is one paragraph further on.
  assert.ok(actDoc(b, [{ t: 'split', at: wp([0], 0) }]));
  net.drain();
  assert.ok(a.session.undo());
  net.drain();
  assert.equal(boxText(b.doc, [2, -1, 0, 0]), 'Box one text\nSecond line');
  assertSameDoc([a, b], 'after the undo');
  // Again, but the paragraph that anchors the box gets a paragraph end: whether the box is the one is not known.
  assert.ok(actDoc(a, [{ t: 'text', at: wp([2, -1, 0, 0, 0], 0), x: 'two ' }]));
  a.session.seal();
  net.drain();
  assert.ok(actDoc(b, [{ t: 'split', at: wp([2], 3) }]));
  net.drain();
  assert.equal(a.session.undo(), undefined);
  assert.deepEqual(codes(a), ['undo-gone']);
  assertSameDoc([a, b], 'after the undo that was not made');
  for (const member of members) member.session.leave();
});

// -------------------------------------------------------------------- the caret

test('a caret in a box follows what someone else did in the box, and leaves a box that may have moved', () => {
  const caret: DocRange = { anchor: { story: 'body', path: BOX1.concat([0]), offset: 8 }, focus: { story: 'body', path: BOX1.concat([0]), offset: 8 } };
  const deleted = movedSelection(caret, [{ k: 'del', s: 'body', c: BOX1, r1: 0, o1: 0, r2: 0, o2: 4 }]);
  assert.deepEqual(deleted && deleted.focus, { story: 'body', path: BOX1.concat([0]), offset: 4 });
  const typed = movedSelection(caret, [{ k: 'ins', s: 'body', c: BOX1, row: 0, off: 0, rows: 1, last: 2 }]);
  assert.deepEqual(typed && typed.focus, { story: 'body', path: BOX1.concat([1]), offset: 10 });
  const moved = movedSelection(caret, [{ k: 'ins', s: 'body', c: [], row: 0, off: 0, rows: 1, last: 0 }]);
  assert.deepEqual(moved && moved.focus, { story: 'body', path: [2, -1, 0, 0, 0], offset: 8 });
  const shaken = movedSelection(caret, [{ k: 'ins', s: 'body', c: [], row: 1, off: 2, rows: 1, last: 0 }]);
  assert.deepEqual(shaken && shaken.focus, { story: 'body', path: [1], offset: 0 }, 'at the paragraph that anchored the box');
});

// --------------------------------------------------------------------- joining

test('who joins from a saved file holds the boxes as the others do and writes the same file', async () => {
  const { net, members } = await room(['a', 'b']);
  const [a, b] = members;
  assert.ok(actDoc(a, [{ t: 'text', at: wp(BOX1.concat([0]), 3), x: ' saved' }]));
  assert.ok(actDoc(b, [{ t: 'split', at: wp(INNER.concat([0]), 5) }]));
  net.drain();
  const file = await saveDoc(a);
  a.session.announceSaved([contentVersion(file)], a.session.rev);
  net.drain();
  const late = await joinDoc(net, 'late', file);
  net.drain();
  assert.equal(late.session.state, 'online');
  assert.deepEqual(boxes(late.doc), boxes(a.doc));
  assert.equal(boxText(late.doc, INNER), 'Inner\n box');
  // Everybody goes on in the boxes.
  assert.ok(actDoc(late, [{ t: 'text', at: wp(BOX1.concat([1]), 0), x: 'L ' }]));
  assert.ok(actDoc(a, [{ t: 'text', at: wp(BOX1.concat([0]), 0), x: 'A ' }]));
  assert.ok(actDoc(b, [{ t: 'del', at: wp(INNER.concat([1]), 0), to: wp(INNER.concat([1]), 1) }]));
  net.drain();
  net.drain();
  assert.deepEqual(boxes(late.doc), boxes(a.doc));
  assert.equal(boxText(late.doc, BOX1), 'A Box saved one text\nL Second line');
  const files = [await saveDoc(a), await saveDoc(b), await saveDoc(late)];
  assert.ok(await samePackage(files[0], files[1]), 'a and b write the same');
  assert.ok(await samePackage(files[0], files[2]), 'who joined late writes the same');
  assert.deepEqual(packageProblems(files), []);
  for (const member of members.concat([late])) member.session.leave();
});

// ---------------------------------------------------------------------- random

/** Every paragraph id and bookmark the document names as its file will hold them (what decides a put; see namesOfDocument). */
function namesOf(doc: DocModel): string[] {
  const out: string[] = [];
  namesOfDocument(doc, [], (thing) => out.push(JSON.stringify(thing)));
  return out.sort();
}

/** The document without the trees its drawings keep as their file had them (the blocks of their boxes stay). */
function withoutKeptTrees(doc: DocModel): DocModel {
  const strip = (value: unknown, depth: number): unknown => {
    if (depth > 200 || value === null || typeof value !== 'object') return value;
    if (Array.isArray(value)) return value.map((item) => strip(item, depth + 1));
    if (value instanceof Map || value instanceof Uint8Array) return value;
    const out: Record<string, unknown> = {};
    const record = value as Record<string, unknown>;
    const graphic = record.type === 'drawing' || record.type === 'pict' || record.type === 'object' || record.type === 'alternate';
    for (const key of Object.keys(record)) if (!(graphic && key === 'xml') && key !== 'parent') out[key] = strip(record[key], depth + 1);
    return out;
  };
  return { ...doc, body: { ...doc.body, blocks: strip(doc.body.blocks, 0) as Block[] } };
}

/** The text of every box of a document, wherever it stands, in the order of the document (boxes in boxes after their box). */
function allBoxes(doc: DocModel): string[] {
  const out: string[] = [];
  const walk = (blocks: readonly Block[], depth: number): void => {
    for (const block of blocks) {
      if (block.type === 'table') for (const row of block.rows) for (const cell of row.cells) walk(cell.blocks, depth);
      if (block.type !== 'paragraph' || depth > 16) continue;
      raw.eachGraphic(block, (item) => {
        for (const box of item.textBoxes || []) {
          out.push(box.map((one) => (one.type === 'paragraph' ? paragraphText(one) : `[${one.type}]`)).join('\n'));
          walk(box, depth + 1);
        }
      });
    }
  };
  walk(doc.body.blocks, 0);
  doc.headers.forEach((story) => walk(story.blocks, 0));
  return out;
}

/** Every paragraph of the boxes that can be changed, by its path. */
function boxParagraphs(doc: DocModel): { path: number[]; length: number }[] {
  const out: { path: number[]; length: number }[] = [];
  const lists: number[][] = [];
  // The boxes of the fixture, wherever their paragraphs are now (found again: Enter and joins move them).
  doc.body.blocks.forEach((block, row) => {
    if (block.type !== 'paragraph') return;
    raw.eachGraphic(block, (item, g, place) => {
      (item.textBoxes || []).forEach((box, b) => {
        const list = [row, -1, g, b];
        if (raw.textBoxProblem(doc, { story: 'body', path: list.concat([0]), offset: 0 }) === undefined) lists.push(list);
        box.forEach((inner, r) => {
          if (inner.type !== 'paragraph') return;
          raw.eachGraphic(inner, (deep, g2) => (deep.textBoxes || []).forEach((_, b2) => lists.push(list.concat([r, -1, g2, b2]))));
        });
      });
      void place;
    });
  });
  for (const list of lists) {
    let blocks;
    try {
      blocks = raw.containerBlocks(doc, 'body', list);
    } catch {
      continue;
    }
    blocks.forEach((block, r) => {
      if (block.type === 'paragraph') out.push({ path: list.concat([r]), length: raw.paragraphLength(block) });
    });
  }
  return out;
}

/** A change of a box (or, now and then, of the body around it), made from random numbers on what a participant holds now. */
function change(member: DocMember, random: () => number): DocEdit | undefined {
  const doc = member.doc;
  const roll = random();
  if (roll < 0.25) {
    const rows: { path: number[]; length: number }[] = [];
    doc.body.blocks.forEach((block, row) => block.type === 'paragraph' && rows.push({ path: [row], length: raw.paragraphLength(block) }));
    const at = pick(random, rows);
    const offset = int(random, 0, at.length);
    const place: DocPosition = { story: 'body', path: at.path, offset };
    const kind = random();
    // A drawing of the paragraph, deleted by a key or written over (it has no width: nothing moves in the text).
    const atoms: number[] = [];
    const paragraph = doc.body.blocks[at.path[0]];
    if (paragraph.type === 'paragraph') for (let o = 0; o <= at.length; o++) if (raw.hasAtomAt(paragraph, o)) atoms.push(o);
    if (atoms.length > 0 && kind < 0.15) {
      const o = pick(random, atoms);
      const before: DocPosition = { story: 'body', path: at.path, offset: o, before: true };
      const which = random();
      if (which < 0.35) return ops.deleteBackward(doc, { story: 'body', path: at.path, offset: o }, 'char');
      if (which < 0.7) return ops.deleteForward(doc, before, 'char');
      return ops.insertText(doc, { anchor: before, focus: { story: 'body', path: at.path, offset: o } }, 'Z');
    }
    if (kind < 0.5) return ops.insertText(doc, place, pick(random, ['b', ' body ', 'x']));
    if (kind < 0.75) return ops.splitParagraph(doc, place);
    return ops.deleteRange(doc, { anchor: place, focus: { story: 'body', path: at.path, offset: Math.min(at.length, offset + 2) } });
  }
  const list = boxParagraphs(doc);
  if (list.length === 0) return undefined;
  const at = pick(random, list);
  const offset = int(random, 0, at.length);
  const place: DocPosition = { story: 'body', path: at.path, offset };
  const near = list.filter((one) => one.path.slice(0, -1).join() === at.path.slice(0, -1).join());
  const other = pick(random, near);
  const range: DocRange = { anchor: place, focus: { story: 'body', path: other.path, offset: int(random, 0, other.length) } };
  const kind = random();
  const run = (command: (doc: DocModel, range: DocRange) => DocEdit, where: DocRange = { anchor: place, focus: place }): DocEdit => sharedInTextBox(doc, where, command);
  if (kind < 0.4) return run((d, r) => ops.insertText(d, r, pick(random, ['a', 'word', ' ', 'İş', 'x\ny'])));
  if (kind < 0.5) return run((d, r) => ops.splitParagraph(d, r.focus));
  if (kind < 0.6) return run((d, r) => ops.deleteRange(d, r), range);
  if (kind < 0.7) return run((d, r) => ops.deleteBackward(d, r.focus, 'char'));
  if (kind < 0.8) return run((d, r) => ops.setRunFormat(d, r, { bold: true, color: '00A000' }), range);
  if (kind < 0.9) return run((d, r) => ops.toggleRunFormat(d, r, 'italic'), range);
  return run((d, r) => ops.setParaFormat(d, r, { align: pick(random, ['left', 'center']) }), range);
}

const SEEDS = Math.max(1, Math.min(Number(process.env.COLLAB_SEEDS) || 30, 100000));

test('boxes at random: three type, delete, format, undo and lose their connection in boxes and around them; everybody holds and writes the same, also who joins from a saved file', async (t) => {
  let late = 0;
  let steps = 0;
  let inBoxes = 0;
  let dropped = 0;
  let spelled = 0;
  const files: Uint8Array[] = [];
  for (let seed = 1; seed <= SEEDS; seed++) {
    const random = seeded(seed * 104729);
    // Every other room: the first box holds other ids in its two branches.
    const { net, members } = await room(['a', 'b', 'c'], seed % 2 === 0);
    const histories = members.map((member) => new CollabDocHistory<string>(member.session));
    const suspended = [false, false, false];
    for (let tick = 0; tick < 60; tick++) {
      const who = int(random, 0, 2);
      const roll = random();
      if (roll < 0.3) net.step(random);
      else if (roll < 0.38) histories[who].undo();
      else if (roll < 0.41) histories[who].redo();
      else if (roll < 0.45) {
        if (suspended[who]) members[who].transport.resume();
        else members[who].transport.suspend();
        suspended[who] = !suspended[who];
      } else {
        try {
          const edit = change(members[who], random);
          if (edit && edit.ops.length > 0) {
            histories[who].apply(edit.ops, '', '', random() < 0.5 ? 'k' : undefined);
            steps++;
            const wires = wiresOf(edit.ops) || [];
            if (wires.some((wire) => wire.t !== 'cmt' && wire.at.p.indexOf(-1) >= 0)) inBoxes++;
          }
        } catch (error) {
          if (!(error instanceof RangeError)) throw error;
        }
      }
    }
    members.forEach((member, i) => suspended[i] && member.transport.resume());
    net.drain(random);
    net.drain(random);
    assertSameDoc(members, `seed ${seed}`);
    for (const member of members) dropped += member.session.stats.dropped;
    const saved = [await saveDoc(members[0]), await saveDoc(members[1]), await saveDoc(members[2])];
    assert.ok(await samePackage(saved[0], saved[1]), `seed ${seed}: a and b write the same file`);
    assert.ok(await samePackage(saved[0], saved[2]), `seed ${seed}: a and c write the same file`);
    files.push(saved[0]);
    // One who joins from the saved file, and goes on with the others.
    members[0].session.announceSaved([contentVersion(saved[0])], members[0].session.rev);
    net.drain(random);
    const joined = await joinDoc(net, 'late', saved[0]);
    net.drain(random);
    if (joined.session.state === 'online') {
      late++;
      const lateHistory = new CollabDocHistory<string>(joined.session);
      for (let tick = 0; tick < 16; tick++) {
        const mine = random() < 0.5;
        const history = mine ? lateHistory : histories[0];
        const roll = random();
        // The member undoes what it did before the save too: the boxes go back to what the file it started from held.
        if (roll < 0.2) history.undo();
        else if (roll < 0.25) history.redo();
        else {
          try {
            const edit = change(mine ? joined : members[0], random);
            if (edit && edit.ops.length > 0) history.apply(edit.ops, '', '');
          } catch (error) {
            if (!(error instanceof RangeError)) throw error;
          }
        }
        if (random() < 0.5) net.drain(random);
      }
      net.drain(random);
      net.drain(random);
      // KNOWN LIMIT (see the test "who joins from a saved file holds the model of the boxes the others hold"): the model keeps the tree of a drawing as its file had
      // it, boxes and all, beside the blocks of its boxes - who was there keeps that of the file the room started from,
      // who joined that of the saved one - so canonDoc differs there, and only there.
      assert.equal(canonDoc(withoutKeptTrees(joined.doc)), canonDoc(withoutKeptTrees(members[0].doc)), `seed ${seed}: who joined late`);
      assert.deepEqual(allBoxes(joined.doc), allBoxes(members[0].doc), `seed ${seed}: the boxes of who joined late`);
      // What decides a put is the same at both: what the document names, read as the file will hold it.
      assert.deepEqual(namesOf(joined.doc), namesOf(members[0].doc), `seed ${seed}: who joined late counts the same names`);
      const mine = await saveDoc(members[0]);
      const theirs = await saveDoc(joined);
      // KNOWN LIMIT: a box that is back to what the file of one of them held is written with the bytes of that file (the
      // writer keeps what did not change as it was read), the other writes it from the model: the XML of that box may be
      // written otherwise (spaces, attributes, their order). What the document holds and what it names are the same
      // (asserted above); the bytes are not asked for here.
      if (!(await samePackage(mine, theirs))) spelled++;
      files.push(theirs);
    }
    for (const member of members.concat([joined])) member.session.leave();
  }
  assert.deepEqual(packageProblems(files), [], 'an independent reader takes every file');
  t.diagnostic(`${SEEDS} rooms, ${steps} changes (${inBoxes} in boxes, ${dropped} given way), ${late} with one who joined from the saved file (${spelled} of whom wrote a box in other XML)`);
  assert.ok(steps > SEEDS * 5);
  assert.ok(inBoxes > steps / 3, 'most of them in boxes');
});
